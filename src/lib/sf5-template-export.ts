import JSZip from "jszip";

export type SF5TemplateRow = {
  lrn: string;
  name: string;
  average: number | null;
  action: string;
  failedAreas: string;
};

export type SF5TemplateOptions = {
  region: string;
  division: string;
  schoolId: string;
  schoolYear: string;
  curriculum: string;
  schoolName: string;
  gradeLevel: string;
  section: string;
  maleRows: SF5TemplateRow[];
  femaleRows: SF5TemplateRow[];
  summary: {
    promotedMale: number;
    promotedFemale: number;
    conditionalMale: number;
    conditionalFemale: number;
    retainedMale: number;
    retainedFemale: number;
  };
  progress: {
    label: string;
    male: number;
    female: number;
  }[];
  adviser: string;
  schoolHead: string;
};

const SF5_EXCEL_TEMPLATE_URL = "/templates/SF5-class-adviser.xlsx";

const SF5_MAX_MALE_LEARNERS = 19;
const SF5_MAX_FEMALE_LEARNERS = 20;

// The supplied template uses merged blocks of different heights so its learner
// table can align with the summary and signature sections on the right. The
// export reuses these blocks from top to bottom, then removes the unused blocks
// in A:L. This makes TOTAL MALE, FEMALE, TOTAL FEMALE, and COMBINED follow the
// last real learner without deleting any rows used by the form on the right.
const SF5_TEMPLATE_BLOCK_START_ROWS = [
  14, 16, 17, 20, 22, 23, 25, 26, 28, 29, 31, 32, 34, 35, 37, 38, 39, 42, 44, 45, 46, 47, 48, 50,
  52, 54, 56, 59, 61, 64, 65, 68, 71, 73, 75, 77, 79, 81, 83, 85, 87, 89, 92,
] as const;

const SF5_TEMPLATE_BLOCKS = SF5_TEMPLATE_BLOCK_START_ROWS.map((start, index) => ({
  start,
  end: (SF5_TEMPLATE_BLOCK_START_ROWS[index + 1] ?? 93) - 1,
}));

const SF5_LEFT_COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"] as const;

const SPREADSHEET_XML_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

function getElementChildren(parent: Element) {
  return Array.from(parent.childNodes).filter((child): child is Element => child.nodeType === 1);
}

function parseSF5TemplateXml(xml: string, partName: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`The SF5 Excel template has invalid ${partName}.`);
  }
  return document;
}

function getSF5TemplateCell(document: XMLDocument, address: string) {
  const cell = Array.from(document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c")).find(
    (candidate) => candidate.getAttribute("r") === address,
  );

  if (!cell) {
    throw new Error(`The SF5 Excel template is missing cell ${address}.`);
  }
  return cell;
}

function setSF5TemplateCell(document: XMLDocument, address: string, value: string | number | null) {
  const cell = getSF5TemplateCell(document, address);

  // Remove only the old value/formula nodes. The cell's style attribute stays
  // intact, so the supplied font, border, fill, and alignment are preserved.
  getElementChildren(cell).forEach((child) => {
    if (["v", "is", "f"].includes(child.localName)) child.remove();
  });

  if (value == null || value === "") {
    cell.removeAttribute("t");
    return;
  }

  if (typeof value === "number") {
    cell.setAttribute("t", "n");
    const valueNode = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "v");
    valueNode.textContent = String(value);
    cell.appendChild(valueNode);
    return;
  }

  // Inline strings let us update the provided template without rebuilding its
  // shared-string table. This also keeps 12-digit LRNs as exact text values.
  cell.setAttribute("t", "inlineStr");
  const inlineString = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "is");
  const textNode = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "t");
  textNode.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
  textNode.textContent = value;
  inlineString.appendChild(textNode);
  cell.appendChild(inlineString);
}

function setSF5TemplateLearner(document: XMLDocument, rowNumber: number, learner: SF5TemplateRow) {
  setSF5TemplateCell(document, `A${rowNumber}`, learner.lrn || null);
  setSF5TemplateCell(document, `C${rowNumber}`, learner.name || null);
  setSF5TemplateCell(document, `G${rowNumber}`, learner.average ?? null);
  setSF5TemplateCell(document, `H${rowNumber}`, learner.action || null);
  setSF5TemplateCell(document, `J${rowNumber}`, learner.failedAreas || null);
}

function getSF5CellColumnNumber(address: string) {
  const columnLetters = address.match(/^[A-Z]+/)?.[0] ?? "";
  return Array.from(columnLetters).reduce(
    (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
    0,
  );
}

function getSF5TemplateStyleProfile(document: XMLDocument, sourceRow: number) {
  return Object.fromEntries(
    SF5_LEFT_COLUMNS.map((column) => [
      column,
      getSF5TemplateCell(document, `${column}${sourceRow}`).getAttribute("s"),
    ]),
  ) as Record<(typeof SF5_LEFT_COLUMNS)[number], string | null>;
}

function applySF5TemplateStyleProfile(
  document: XMLDocument,
  block: { start: number; end: number },
  profile: Record<(typeof SF5_LEFT_COLUMNS)[number], string | null>,
) {
  for (let rowNumber = block.start; rowNumber <= block.end; rowNumber += 1) {
    SF5_LEFT_COLUMNS.forEach((column) => {
      const cell = getSF5TemplateCell(document, `${column}${rowNumber}`);
      const style = profile[column];
      if (style == null) cell.removeAttribute("s");
      else cell.setAttribute("s", style);
    });
  }
}

function clearSF5TemplateBlockValues(document: XMLDocument, block: { start: number; end: number }) {
  Array.from(document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c")).forEach((cell) => {
    const address = cell.getAttribute("r") ?? "";
    const rowNumber = Number(address.match(/\d+$/)?.[0] ?? 0);
    if (rowNumber < block.start || rowNumber > block.end || getSF5CellColumnNumber(address) > 12) {
      return;
    }

    getElementChildren(cell).forEach((child) => {
      if (["v", "is", "f"].includes(child.localName)) child.remove();
    });
    cell.removeAttribute("t");
  });
}

function removeSF5TemplateBlock(document: XMLDocument, block: { start: number; end: number }) {
  const mergeCells = document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "mergeCells")[0];

  if (mergeCells) {
    getElementChildren(mergeCells).forEach((mergeCell) => {
      const reference = mergeCell.getAttribute("ref") ?? "";
      const [from = "", to = from] = reference.split(":");
      const fromRow = Number(from.match(/\d+$/)?.[0] ?? 0);
      const toRow = Number(to.match(/\d+$/)?.[0] ?? 0);
      const fromColumn = getSF5CellColumnNumber(from);
      const toColumn = getSF5CellColumnNumber(to);

      if (fromColumn >= 1 && toColumn <= 12 && fromRow >= block.start && toRow <= block.end) {
        mergeCell.remove();
      }
    });
    mergeCells.setAttribute("count", String(getElementChildren(mergeCells).length));
  }

  Array.from(document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c")).forEach((cell) => {
    const address = cell.getAttribute("r") ?? "";
    const rowNumber = Number(address.match(/\d+$/)?.[0] ?? 0);
    if (
      rowNumber >= block.start &&
      rowNumber <= block.end &&
      getSF5CellColumnNumber(address) <= 12
    ) {
      cell.remove();
    }
  });
}

function compactSF5TemplateLearnerTable(document: XMLDocument, options: SF5TemplateOptions) {
  const styles = {
    learner: getSF5TemplateStyleProfile(document, 14),
    totalMale: getSF5TemplateStyleProfile(document, 45),
    femaleHeader: getSF5TemplateStyleProfile(document, 46),
    totalFemale: getSF5TemplateStyleProfile(document, 89),
    combined: getSF5TemplateStyleProfile(document, 92),
  };
  const entries: Array<
    | { kind: "learner"; learner: SF5TemplateRow }
    | {
        kind: "label";
        style: keyof typeof styles;
        count: number | null;
        label: string;
      }
  > = [
    ...options.maleRows.map((learner) => ({
      kind: "learner" as const,
      learner,
    })),
    {
      kind: "label",
      style: "totalMale",
      count: options.maleRows.length,
      label: "<=== TOTAL MALE",
    },
    {
      kind: "label",
      style: "femaleHeader",
      count: null,
      label: "FEMALE",
    },
    ...options.femaleRows.map((learner) => ({
      kind: "learner" as const,
      learner,
    })),
    {
      kind: "label",
      style: "totalFemale",
      count: options.femaleRows.length,
      label: "<=== TOTAL FEMALE",
    },
    {
      kind: "label",
      style: "combined",
      count: options.maleRows.length + options.femaleRows.length,
      label: "<=== COMBINED",
    },
  ];

  SF5_TEMPLATE_BLOCKS.forEach((block) => {
    clearSF5TemplateBlockValues(document, block);
  });

  entries.forEach((entry, index) => {
    const block = SF5_TEMPLATE_BLOCKS[index];
    if (!block) throw new Error("The SF5 template does not have enough learner rows.");

    if (entry.kind === "learner") {
      applySF5TemplateStyleProfile(document, block, styles.learner);
      setSF5TemplateLearner(document, block.start, entry.learner);
      return;
    }

    applySF5TemplateStyleProfile(document, block, styles[entry.style]);
    setSF5TemplateCell(document, `A${block.start}`, entry.count);
    setSF5TemplateCell(document, `C${block.start}`, entry.label);
  });

  SF5_TEMPLATE_BLOCKS.slice(entries.length).forEach((block) => {
    removeSF5TemplateBlock(document, block);
  });
}

function configureSF5TemplatePageSetup(document: XMLDocument) {
  const worksheet = document.documentElement;
  let sheetProperties = getElementChildren(worksheet).find(
    (child) => child.localName === "sheetPr",
  );
  if (!sheetProperties) {
    sheetProperties = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "x:sheetPr");
    worksheet.insertBefore(sheetProperties, worksheet.firstChild);
  }

  let pageSetupProperties = getElementChildren(sheetProperties).find(
    (child) => child.localName === "pageSetUpPr",
  );
  if (!pageSetupProperties) {
    pageSetupProperties = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "x:pageSetUpPr");
    sheetProperties.appendChild(pageSetupProperties);
  }
  pageSetupProperties.setAttribute("fitToPage", "1");

  let pageMargins = getElementChildren(worksheet).find(
    (child) => child.localName === "pageMargins",
  );
  if (!pageMargins) {
    pageMargins = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "x:pageMargins");
    worksheet.appendChild(pageMargins);
  }
  pageMargins.setAttribute("left", "0.25");
  pageMargins.setAttribute("right", "0.25");
  pageMargins.setAttribute("top", "0.25");
  pageMargins.setAttribute("bottom", "0.25");
  pageMargins.setAttribute("header", "0.1");
  pageMargins.setAttribute("footer", "0.1");

  getElementChildren(worksheet)
    .filter((child) => child.localName === "pageSetup")
    .forEach((child) => child.remove());
  const pageSetup = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "x:pageSetup");
  // OOXML paper size 14 is Folio/long bond (8.5 × 13 inches).
  pageSetup.setAttribute("paperSize", "14");
  pageSetup.setAttribute("orientation", "landscape");
  pageSetup.setAttribute("fitToWidth", "1");
  pageSetup.setAttribute("fitToHeight", "1");
  worksheet.insertBefore(pageSetup, pageMargins.nextSibling);
}

async function buildSF5TemplateBuffer(options: SF5TemplateOptions) {
  if (options.maleRows.length > SF5_MAX_MALE_LEARNERS) {
    throw new Error(`The SF5 template supports up to ${SF5_MAX_MALE_LEARNERS} male learners.`);
  }
  if (options.femaleRows.length > SF5_MAX_FEMALE_LEARNERS) {
    throw new Error(`The SF5 template supports up to ${SF5_MAX_FEMALE_LEARNERS} female learners.`);
  }

  const response = await fetch(SF5_EXCEL_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(
      `SF5 Excel template not found (${response.status}). Put SF5-class-adviser.xlsx in public/templates/.`,
    );
  }

  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const worksheetPart = zip.file("xl/worksheets/sheet1.xml");
  if (!worksheetPart) {
    throw new Error("The SF5 Excel template does not contain its worksheet.");
  }
  const worksheet = parseSF5TemplateXml(await worksheetPart.async("string"), "worksheet XML");

  // School and class information.
  setSF5TemplateCell(worksheet, "E4", options.region || null);
  setSF5TemplateCell(worksheet, "G4", options.division || null);
  setSF5TemplateCell(worksheet, "E5", options.schoolId || null);
  setSF5TemplateCell(worksheet, "I5", options.schoolYear || null);
  setSF5TemplateCell(worksheet, "L5", options.curriculum || null);
  setSF5TemplateCell(worksheet, "E6", options.schoolName || null);
  setSF5TemplateCell(worksheet, "L6", options.gradeLevel || null);
  setSF5TemplateCell(worksheet, "S6", options.section || null);
  setSF5TemplateCell(
    worksheet,
    "H8",
    "ACTION TAKEN: RETAINED, PROMOTED or PROMOTED WITH ACADEMIC EXCELLENCE AWARD",
  );

  compactSF5TemplateLearnerTable(worksheet, options);

  // Promotion summary.
  setSF5TemplateCell(worksheet, "Q11", options.summary.promotedMale);
  setSF5TemplateCell(worksheet, "T11", options.summary.promotedFemale);
  setSF5TemplateCell(
    worksheet,
    "U11",
    options.summary.promotedMale + options.summary.promotedFemale,
  );
  setSF5TemplateCell(worksheet, "Q13", options.summary.conditionalMale);
  setSF5TemplateCell(worksheet, "T13", options.summary.conditionalFemale);
  setSF5TemplateCell(
    worksheet,
    "U13",
    options.summary.conditionalMale + options.summary.conditionalFemale,
  );
  setSF5TemplateCell(worksheet, "Q15", options.summary.retainedMale);
  setSF5TemplateCell(worksheet, "T15", options.summary.retainedFemale);
  setSF5TemplateCell(
    worksheet,
    "U15",
    options.summary.retainedMale + options.summary.retainedFemale,
  );

  // Level of progress and achievement, ordered from 74-and-below to 90-100.
  const progressRows = [24, 27, 30, 33, 36] as const;
  progressRows.forEach((rowNumber, index) => {
    const progress = options.progress[index];
    const male = progress?.male ?? 0;
    const female = progress?.female ?? 0;
    setSF5TemplateCell(worksheet, `Q${rowNumber}`, male);
    setSF5TemplateCell(worksheet, `T${rowNumber}`, female);
    setSF5TemplateCell(worksheet, `U${rowNumber}`, male + female);
  });

  // Signature fields from the live class/profile data.
  setSF5TemplateCell(worksheet, "N55", options.adviser || null);
  setSF5TemplateCell(worksheet, "N63", options.schoolHead || null);

  configureSF5TemplatePageSetup(worksheet);

  zip.file("xl/worksheets/sheet1.xml", new XMLSerializer().serializeToString(worksheet));
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

export async function downloadSF5ExcelTemplate(fileName: string, options: SF5TemplateOptions) {
  const buffer = await buildSF5TemplateBuffer(options);
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${fileName}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const SF5_WORD_TEMPLATE_URL = "/templates/SF5-class-adviser.docx";
const WORDPROCESSING_XML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_DOCUMENT_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function parseSF5WordTemplateXml(xml: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw new Error("The SF5 Word template contains invalid document XML.");
  }
  return document;
}

function getDirectWordChildren(parent: Element, localName: string) {
  return getElementChildren(parent).filter(
    (child) => child.namespaceURI === WORDPROCESSING_XML_NAMESPACE && child.localName === localName,
  );
}

function getDirectWordRows(table: Element) {
  return getDirectWordChildren(table, "tr");
}

function getDirectWordCells(row: Element) {
  return getDirectWordChildren(row, "tc");
}

function getWordNodeText(node: Element) {
  return Array.from(node.getElementsByTagNameNS(WORDPROCESSING_XML_NAMESPACE, "t"))
    .map((textNode) => textNode.textContent || "")
    .join("")
    .trim();
}

function setSF5WordParagraphText(paragraph: Element, value: string | number | null) {
  getElementChildren(paragraph).forEach((child) => {
    if (child.localName !== "pPr") child.remove();
  });

  const run = paragraph.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:r");
  const lines = String(value ?? "").split("\n");

  lines.forEach((line, index) => {
    if (index > 0) {
      run.appendChild(
        paragraph.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:br"),
      );
    }
    const text = paragraph.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:t");
    text.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
    text.textContent = line;
    run.appendChild(text);
  });

  paragraph.appendChild(run);
}

function setSF5WordCellText(cell: Element, value: string | number | null) {
  let paragraph = getDirectWordChildren(cell, "p")[0];
  if (!paragraph) {
    paragraph = cell.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:p");
    cell.appendChild(paragraph);
  }
  setSF5WordParagraphText(paragraph, value);
}

function setSF5WordTableCell(
  table: Element,
  rowIndex: number,
  cellIndex: number,
  value: string | number | null,
) {
  const row = getDirectWordRows(table)[rowIndex];
  const cell = row ? getDirectWordCells(row)[cellIndex] : undefined;
  if (!cell) {
    throw new Error(
      `The SF5 Word template is missing table cell ${rowIndex + 1}:${cellIndex + 1}.`,
    );
  }
  setSF5WordCellText(cell, value);
}

function populateSF5WordLearnerRow(row: Element, learner: SF5TemplateRow) {
  const cells = getDirectWordCells(row);
  if (cells.length < 5) {
    throw new Error("The SF5 Word template has an invalid learner row.");
  }
  setSF5WordCellText(cells[0], learner.lrn || null);
  setSF5WordCellText(cells[1], learner.name || null);
  setSF5WordCellText(cells[2], learner.average ?? null);
  setSF5WordCellText(cells[3], learner.action || null);
  setSF5WordCellText(cells[4], learner.failedAreas || null);
}

function findSF5WordRow(rows: Element[], text: string) {
  const row = rows.find((candidate) => getWordNodeText(candidate) === text);
  if (!row) throw new Error(`The SF5 Word template is missing the ${text} row.`);
  return row;
}

function populateSF5WordLearnerTable(
  table: Element,
  maleRows: SF5TemplateRow[],
  femaleRows: SF5TemplateRow[],
) {
  const rows = getDirectWordRows(table);
  const maleHeader = findSF5WordRow(rows, "MALE");
  const totalMale = findSF5WordRow(rows, "<=== TOTAL MALE");
  const femaleHeader = findSF5WordRow(rows, "FEMALE");
  const totalFemale = findSF5WordRow(rows, "<=== TOTAL FEMALE");
  const combined = findSF5WordRow(rows, "<=== COMBINED");
  const maleHeaderIndex = rows.indexOf(maleHeader);
  const totalMaleIndex = rows.indexOf(totalMale);
  const femaleHeaderIndex = rows.indexOf(femaleHeader);
  const totalFemaleIndex = rows.indexOf(totalFemale);
  const maleTemplate = rows[maleHeaderIndex + 1]?.cloneNode(true) as Element;
  const femaleTemplate = rows[femaleHeaderIndex + 1]?.cloneNode(true) as Element;

  if (!maleTemplate || !femaleTemplate) {
    throw new Error("The SF5 Word template has no reusable learner row.");
  }

  rows.slice(maleHeaderIndex + 1, totalMaleIndex).forEach((row) => row.remove());
  rows.slice(femaleHeaderIndex + 1, totalFemaleIndex).forEach((row) => row.remove());

  maleRows.forEach((learner) => {
    const row = maleTemplate.cloneNode(true) as Element;
    populateSF5WordLearnerRow(row, learner);
    table.insertBefore(row, totalMale);
  });

  femaleRows.forEach((learner) => {
    const row = femaleTemplate.cloneNode(true) as Element;
    populateSF5WordLearnerRow(row, learner);
    table.insertBefore(row, totalFemale);
  });

  setSF5WordCellText(getDirectWordCells(totalMale)[0], maleRows.length);
  setSF5WordCellText(getDirectWordCells(totalFemale)[0], femaleRows.length);
  setSF5WordCellText(getDirectWordCells(combined)[0], maleRows.length + femaleRows.length);
}

function populateSF5WordCountRow(table: Element, rowIndex: number, male: number, female: number) {
  setSF5WordTableCell(table, rowIndex, 1, male);
  setSF5WordTableCell(table, rowIndex, 2, female);
  setSF5WordTableCell(table, rowIndex, 3, male + female);
}

function countSF5WordRows(rows: SF5TemplateRow[], lower: number, upper: number | null) {
  return rows.filter((row) => {
    if (row.average == null) return false;
    const average = Math.round(row.average);
    return average >= lower && (upper == null || average <= upper);
  }).length;
}

function setSF5WordParagraphBefore(paragraph: Element, before: number) {
  let properties = getDirectWordChildren(paragraph, "pPr")[0];
  if (!properties) {
    properties = paragraph.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:pPr");
    paragraph.insertBefore(properties, paragraph.firstChild);
  }

  let spacing = getDirectWordChildren(properties, "spacing")[0];
  if (!spacing) {
    spacing = paragraph.ownerDocument.createElementNS(WORDPROCESSING_XML_NAMESPACE, "w:spacing");
    properties.appendChild(spacing);
  }
  spacing.setAttributeNS(WORDPROCESSING_XML_NAMESPACE, "w:before", String(before));
}

function populateSF5WordSignatures(rightCell: Element, adviser: string, schoolHead: string) {
  const paragraphs = Array.from(
    rightCell.getElementsByTagNameNS(WORDPROCESSING_XML_NAMESPACE, "p"),
  );
  const setAfterLabel = (label: string, value: string) => {
    const index = paragraphs.findIndex((paragraph) => getWordNodeText(paragraph) === label);
    const target = index >= 0 ? paragraphs[index + 1] : undefined;
    if (!target) {
      throw new Error(`The SF5 Word template is missing the ${label} field.`);
    }
    setSF5WordParagraphText(target, value.toUpperCase());
  };

  setAfterLabel("PREPARED BY:", adviser);
  setAfterLabel("CERTIFIED CORRECT & SUBMITTED BY:", schoolHead);

  // Keep the complete signature block on the same long-bond landscape page.
  // The original template's generous paragraph spacing can push only the last
  // SCC signature onto a second page after learner rows are populated.
  const compactLabels = new Set([
    "Instructions:",
    "PREPARED BY:",
    "CERTIFIED CORRECT & SUBMITTED BY:",
    "REVIEWED BY: SCC Members",
    "Generated thru LIS (SCC CO-Chair)",
  ]);
  const reviewedIndex = paragraphs.findIndex(
    (paragraph) => getWordNodeText(paragraph) === "REVIEWED BY: SCC Members",
  );
  const generatedIndex = paragraphs.findIndex(
    (paragraph) => getWordNodeText(paragraph) === "Generated thru LIS (SCC CO-Chair)",
  );

  paragraphs.forEach((paragraph, index) => {
    const text = getWordNodeText(paragraph);
    if (compactLabels.has(text)) {
      setSF5WordParagraphBefore(paragraph, text.startsWith("Generated thru LIS") ? 0 : 40);
    } else if (text === adviser.toUpperCase() || text === schoolHead.toUpperCase()) {
      setSF5WordParagraphBefore(paragraph, 220);
    } else if (
      !text &&
      reviewedIndex >= 0 &&
      generatedIndex >= 0 &&
      index > reviewedIndex &&
      index < generatedIndex
    ) {
      setSF5WordParagraphBefore(paragraph, 140);
    }
  });
}

async function buildSF5WordTemplateBuffer(options: SF5TemplateOptions) {
  const response = await fetch(SF5_WORD_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(
      `SF5 Word template not found (${response.status}). Put SF5-class-adviser.docx in public/templates/.`,
    );
  }

  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentPart = zip.file("word/document.xml");
  if (!documentPart) {
    throw new Error("The SF5 Word template does not contain word/document.xml.");
  }

  const documentXml = parseSF5WordTemplateXml(await documentPart.async("string"));
  const tables = Array.from(
    documentXml.getElementsByTagNameNS(WORDPROCESSING_XML_NAMESPACE, "tbl"),
  );
  const [metadataTable, outerTable, learnerTable, summaryTable, progressTable] = tables;

  if (!metadataTable || !outerTable || !learnerTable || !summaryTable || !progressTable) {
    throw new Error("The SF5 Word template does not contain the required tables.");
  }

  // School and class information. Only the empty value cells are replaced;
  // the template's fonts, borders, widths, and alignment remain unchanged.
  setSF5WordTableCell(metadataTable, 0, 1, options.region || null);
  setSF5WordTableCell(metadataTable, 0, 3, options.division || null);
  setSF5WordTableCell(metadataTable, 1, 1, options.schoolId || null);
  setSF5WordTableCell(metadataTable, 1, 3, options.schoolYear || null);
  setSF5WordTableCell(metadataTable, 1, 5, options.curriculum || "K to 12");
  setSF5WordTableCell(metadataTable, 2, 1, options.schoolName || null);
  setSF5WordTableCell(metadataTable, 2, 3, options.gradeLevel || null);
  setSF5WordTableCell(metadataTable, 2, 5, options.section || null);

  populateSF5WordLearnerTable(learnerTable, options.maleRows, options.femaleRows);

  // The supplied Word template has separate rows for regular promotion and
  // the Academic Excellence Award, so the 90+ learners are counted separately.
  const promotedMale = countSF5WordRows(options.maleRows, 75, 89);
  const promotedFemale = countSF5WordRows(options.femaleRows, 75, 89);
  const awardMale = countSF5WordRows(options.maleRows, 90, null);
  const awardFemale = countSF5WordRows(options.femaleRows, 90, null);
  const retainedMale = countSF5WordRows(options.maleRows, 0, 74);
  const retainedFemale = countSF5WordRows(options.femaleRows, 0, 74);
  populateSF5WordCountRow(summaryTable, 2, promotedMale, promotedFemale);
  populateSF5WordCountRow(summaryTable, 3, awardMale, awardFemale);
  populateSF5WordCountRow(summaryTable, 4, retainedMale, retainedFemale);

  options.progress.slice(0, 5).forEach((progress, index) => {
    populateSF5WordCountRow(progressTable, index + 2, progress.male, progress.female);
  });

  const outerRows = getDirectWordRows(outerTable);
  const rightCell = outerRows[0] ? getDirectWordCells(outerRows[0])[1] : undefined;
  if (!rightCell) {
    throw new Error("The SF5 Word template is missing its signature section.");
  }
  populateSF5WordSignatures(rightCell, options.adviser, options.schoolHead);

  zip.file("word/document.xml", new XMLSerializer().serializeToString(documentXml));
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

export async function downloadSF5WordTemplate(fileName: string, options: SF5TemplateOptions) {
  const buffer = await buildSF5WordTemplateBuffer(options);
  const blob = new Blob([buffer], { type: WORD_DOCUMENT_MIME });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${fileName}.docx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return true;
}
