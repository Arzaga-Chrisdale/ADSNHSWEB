import { Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useEffect, useMemo, useState } from "react";
import { ArrowLeft, GraduationCap, FileText, Users } from "lucide-react";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  computeAverage,
  type ActivityScore,
  type ClassRow,
  type GradeActivity,
  type GradeComponent,
  type GradeRow,
  type StudentRow,
} from "@/lib/data";
import { PdfPreviewShell, type PdfPaper } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import depedLogo from "@/assets/deped_logo.png";
import schoolLogo from "@/assets/ASNSHS Logo.png";
import { toast } from "sonner";
import html2canvas from "html2canvas";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  AlignmentType,
  Document as WordDocument,
  ImageRun,
  Packer,
  PageBreak,
  PageOrientation,
  Paragraph,
} from "docx";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

type Variant = "classic" | "matatag";

type SF9ClassRow = ClassRow & {
  track_shs?: string | null;
  units?: number | null;
};

// Use the user-provided SF9 spreadsheet as the single export template.
// Keep this file in public/templates/ so Excel export uses the original
// workbook formatting without rebuilding or restyling the sheet.
const SF9_EXCEL_TEMPLATE_URL = "/templates/sf9-class-adviser.xlsx";
// Grade 11 SHS uses its own SF9 spreadsheet layout.
// Keep the uploaded Grade 11 workbook in public/templates/ with this name.
const SF9_GRADE11_EXCEL_TEMPLATE_URL = "/templates/sf9-grade-11-shs.xlsx";

// Grade 12 SHS uses the exact uploaded Grade-12-SHS.xlsx workbook.
// Put the accompanying file at:
//   public/templates/sf9-grade-12-shs.xlsx
const SF9_GRADE12_EXCEL_TEMPLATE_URL = "/templates/sf9-grade-12-shs.xlsx";

const SF9_GRADE12_SUBJECTS = [
  // Core Subjects
  { label: "Media and Information Literacy", term: "1", excelRow: 33, group: "core" },
  { label: "PE and Health 3", term: "1", excelRow: 34, group: "core" },
  { label: "Introduction to Human Philosophy", term: "2", excelRow: 35, group: "core" },
  { label: "Disaster Readiness and Risk Reduction", term: "2", excelRow: 36, group: "core" },
  { label: "Contemporary Philippine Arts", term: "3", excelRow: 37, group: "core" },
  { label: "PE and Health 4", term: "3", excelRow: 38, group: "core" },

  // Applied and Specialized Subjects
  { label: "Filipino sa Piling Larang", term: "1", excelRow: 40, group: "applied" },
  { label: "Practical Research 2", term: "1", excelRow: 41, group: "applied" },
  { label: "General Biology 1", term: "1", excelRow: 42, group: "applied" },
  { label: "General Physics 1", term: "1", excelRow: 43, group: "applied" },
  { label: "English for Academic & Professional Purposes", term: "2", excelRow: 44, group: "applied" },
  { label: "Entrepreneurship", term: "2", excelRow: 45, group: "applied" },
  { label: "General Physics 2", term: "2", excelRow: 46, group: "applied" },
  { label: "General Biology 2", term: "3", excelRow: 47, group: "applied" },
  { label: "Inquiries, Investigation and Immersion", term: "3", excelRow: 49, group: "applied" },
  { label: "Capstone Project", term: "3", excelRow: 51, group: "applied" },
] as const;

type Sf9Grade12SubjectTerm = "1" | "2" | "3";

type Sf9Grade12PreviewRow = {
  label: string;
  term: Sf9Grade12SubjectTerm;
  excelRow: number;
  group: "core" | "applied";
  grade: number | null;
  units: number | null;
};

// Grade 11 Mabisang Komunikasyon is a composite class-record subject.
// Each term is the rounded average of Effective Communication and
// Mabisang Komunikasyon for that same term.
const SF9_COMMUNICATION_SCOPES = [
  { key: "EC_T1", storageTerm: "EC_T1" },
  { key: "MK_T1", storageTerm: "MK_T1" },
  { key: "EC_T2", storageTerm: "EC_T2" },
  { key: "MK_T2", storageTerm: "MK_T2" },
  { key: "EC_T3", storageTerm: "EC_T3" },
  { key: "MK_T3", storageTerm: "MK_T3" },
] as const;

type Sf9GradeTerm = "1" | "2" | "3" | "final";

const SF9_NEW_LEARNING_AREAS = [
  { label: "Filipino", aliases: ["Filipino"], indented: false },
  { label: "English", aliases: ["English"], indented: false },
  { label: "Mathematics", aliases: ["Mathematics", "Math"], indented: false },
  { label: "Science", aliases: ["Science"], indented: false },
  {
    label: "Araling Panlipunan (AP)",
    aliases: ["Araling Panlipunan", "Araling Panlipunan (AP)", "AP"],
    indented: false,
  },
  {
    label: "Values Education",
    aliases: ["Values Education", "Values Ed", "GMRC", "GMRC / Values Education"],
    indented: false,
  },
  {
    label: "Creative Technologies",
    aliases: ["Creative Technologies", "Creative Tech", "Technology"],
    indented: false,
  },
  { label: "MAPEH", aliases: ["MAPEH"], indented: false },
  {
    label: "Music and Arts",
    aliases: ["Music and Arts", "Music", "Arts"],
    indented: true,
  },
  {
    label: "Physical Education and Health",
    aliases: ["Physical Education and Health", "PE and Health", "PE & Health"],
    indented: true,
  },
  {
    label: "Research",
    aliases: ["Research"],
    indented: false,
  },
] as const;

const SF9_PERFORMANCE_DESCRIPTORS = [
  { scale: "90 - 100", description: "Advancing", remarks: "Passed" },
  { scale: "80 - 89", description: "Benchmarking", remarks: "Passed" },
  { scale: "75 - 79", description: "Connecting", remarks: "Passed" },
  { scale: "65 - 74", description: "Developing", remarks: "Failed" },
  { scale: "0 - 64", description: "Emerging", remarks: "Failed" },
] as const;

const SF9_ATTENDANCE_MONTHS = [
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
  "January",
  "February",
  "March",
  "April",
] as const;


function normalizeSubjectName(value: unknown) {
  return String(value ?? "").trim().toLowerCase();
}

function normalizeSf9SubjectKey(value: unknown) {
  return normalizeSubjectName(value)
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isSf9CommunicationCompositeSubject(value: unknown) {
  const normalized = normalizeSubjectName(value).replace(/\s+/g, " ");
  return (
    normalized === "mabisang komunikasyon" ||
    normalized === "mabisang komuniksyon"
  );
}

function isSf9AcademicElectiveSubject(value: unknown) {
  return normalizeSubjectName(value).replace(/\s+/g, " ") === "elective subject";
}

function roundSf9TermGrade(value: number | null | undefined): number | null {
  if (typeof value !== "number" || Number.isNaN(value)) return null;
  return Math.max(0, Math.min(100, Math.floor(value + 0.5)));
}

function resolveSf9TermGradeBase(
  computedBase: number | null,
  storedBase: number | null | undefined,
): number | null {
  if (computedBase == null) return null;

  const numericBase = storedBase == null ? computedBase : Number(storedBase);
  const resolvedBase = Number.isFinite(numericBase) ? numericBase : computedBase;
  return Math.max(0, Math.min(100, Math.floor(resolvedBase + 0.5)));
}

function averageSf9Pair(first: number | null, second: number | null): number | null {
  if (first == null || second == null) return null;
  return Math.round((first + second) / 2);
}

function learningAreaForSubject(subject: string) {
  const normalized = normalizeSubjectName(subject);

  const known = SF9_NEW_LEARNING_AREAS.find((area) =>
    area.aliases.some((alias) => normalizeSubjectName(alias) === normalized),
  );

  if (known) return known;

  return {
    label: subject.trim(),
    aliases: [subject.trim()] as readonly string[],
    indented: false,
  };
}

function sanitizeExcelFileName(value: string) {
  return value
    .replace(/[\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function downloadExcelBuffer(buffer: ArrayBuffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${sanitizeExcelFileName(fileName)}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}


const XLSX_MAIN_NAMESPACE =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XLSX_OFFICE_REL_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

function getXlsxCell(sheetDocument: Document, reference: string) {
  return sheetDocument.querySelector(`c[r="${reference}"]`);
}

function setXlsxCellText(
  sheetDocument: Document,
  reference: string,
  value: string | number | null | undefined,
) {
  const cell = getXlsxCell(sheetDocument, reference);
  if (!cell) {
    throw new Error(`The Grade 11 Excel template is missing cell ${reference}.`);
  }

  while (cell.firstChild) cell.removeChild(cell.firstChild);

  const text = String(value ?? "");
  if (!text) {
    cell.removeAttribute("t");
    return;
  }

  cell.setAttribute("t", "inlineStr");
  const inlineString = sheetDocument.createElementNS(XLSX_MAIN_NAMESPACE, "is");
  const textNode = sheetDocument.createElementNS(XLSX_MAIN_NAMESPACE, "t");
  textNode.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
  textNode.textContent = text;
  inlineString.appendChild(textNode);
  cell.appendChild(inlineString);
}

function setXlsxCellNumber(
  sheetDocument: Document,
  reference: string,
  value: number | null | undefined,
) {
  const cell = getXlsxCell(sheetDocument, reference);
  if (!cell) {
    throw new Error(`The Grade 11 Excel template is missing cell ${reference}.`);
  }

  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.removeAttribute("t");

  if (value == null || !Number.isFinite(Number(value))) return;

  const valueNode = sheetDocument.createElementNS(XLSX_MAIN_NAMESPACE, "v");
  valueNode.textContent = String(Number(value));
  cell.appendChild(valueNode);
}

function setXlsxCellValue(
  sheetDocument: Document,
  reference: string,
  value: string | number | null | undefined,
) {
  if (typeof value === "number" && Number.isFinite(value)) {
    setXlsxCellNumber(sheetDocument, reference, value);
  } else {
    setXlsxCellText(sheetDocument, reference, value);
  }
}

/**
 * Clone the template cell style and change only the alignment/number format.
 * This preserves its font, borders, and gray/white fill.
 */
function applyCenteredXlsxStyle(
  sheetDocument: Document,
  stylesDocument: Document,
  reference: string,
  options?: { integer?: boolean },
) {
  const cell = getXlsxCell(sheetDocument, reference);
  if (!cell) {
    throw new Error(`The Grade 11 Excel template is missing cell ${reference}.`);
  }

  const cellXfs = stylesDocument.getElementsByTagNameNS(
    XLSX_MAIN_NAMESPACE,
    "cellXfs",
  )[0];

  if (!cellXfs) {
    throw new Error("The Grade 11 Excel template is missing cell styles.");
  }

  const currentStyleIndex = Number(cell.getAttribute("s") || "0");
  const currentStyle = cellXfs.children.item(currentStyleIndex);

  if (!currentStyle) {
    throw new Error(
      `The Grade 11 Excel template is missing style ${currentStyleIndex}.`,
    );
  }

  const clonedStyle = currentStyle.cloneNode(true) as Element;

  if (options?.integer) {
    // Built-in Excel number format 1 = "0".
    // This makes a selected 3 Units display as 3 instead of 3.00.
    clonedStyle.setAttribute("numFmtId", "1");
    clonedStyle.setAttribute("applyNumberFormat", "1");
  }

  let alignment = clonedStyle.getElementsByTagNameNS(
    XLSX_MAIN_NAMESPACE,
    "alignment",
  )[0];

  if (!alignment) {
    alignment = stylesDocument.createElementNS(
      XLSX_MAIN_NAMESPACE,
      "alignment",
    );
    clonedStyle.appendChild(alignment);
  }

  alignment.setAttribute("horizontal", "center");
  alignment.setAttribute("vertical", "center");
  clonedStyle.setAttribute("applyAlignment", "1");

  const newStyleIndex = cellXfs.children.length;
  cellXfs.appendChild(clonedStyle);
  cellXfs.setAttribute("count", String(cellXfs.children.length));
  cell.setAttribute("s", String(newStyleIndex));
}

function setXlsxRowHidden(
  sheetDocument: Document,
  rowNumber: number,
  hidden: boolean,
) {
  const row = sheetDocument.querySelector(`row[r="${rowNumber}"]`);
  if (!row) return;

  if (hidden) {
    row.setAttribute("hidden", "1");
  } else {
    row.removeAttribute("hidden");
  }
}

async function resolveXlsxWorksheet(
  zip: JSZip,
  preferredSheetNames: string[],
) {
  const workbookFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");

  if (!workbookFile || !relsFile) {
    throw new Error("The Grade 11 Excel template is missing workbook metadata.");
  }

  const parser = new DOMParser();
  const workbookDocument = parser.parseFromString(
    await workbookFile.async("text"),
    "application/xml",
  );
  const relationshipsDocument = parser.parseFromString(
    await relsFile.async("text"),
    "application/xml",
  );

  const sheets = Array.from(
    workbookDocument.getElementsByTagNameNS("*", "sheet"),
  );
  const preferred = preferredSheetNames.map((name) => name.trim().toLowerCase());
  const sheet =
    sheets.find((item) =>
      preferred.includes(String(item.getAttribute("name") ?? "").trim().toLowerCase()),
    ) ?? sheets[0];

  if (!sheet) {
    throw new Error("The Grade 11 worksheet was not found in the template.");
  }

  const relationshipId =
    sheet.getAttributeNS(XLSX_OFFICE_REL_NAMESPACE, "id") ||
    sheet.getAttribute("r:id");

  const relationship = Array.from(
    relationshipsDocument.getElementsByTagNameNS("*", "Relationship"),
  ).find((item) => item.getAttribute("Id") === relationshipId);

  const target = relationship?.getAttribute("Target");
  if (!target) {
    throw new Error("The Grade 11 worksheet relationship is missing.");
  }

  const cleanTarget = target.replace(/^\/+/, "").replace(/^\.\//, "");
  const worksheetPath = cleanTarget.startsWith("xl/")
    ? cleanTarget
    : `xl/${cleanTarget}`;
  const worksheetFile = zip.file(worksheetPath);

  if (!worksheetFile) {
    throw new Error("The Grade 11 worksheet XML could not be loaded.");
  }

  const worksheetDocument = parser.parseFromString(
    await worksheetFile.async("text"),
    "application/xml",
  );

  return { worksheetPath, worksheetDocument };
}

async function removeStaleXlsxCalculationChain(zip: JSZip) {
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  // A copied Excel template can contain calcChain.xml entries that still
  // point to formulas/cells from the original workbook. After SF9 Grade 12
  // values are replaced, that chain is stale and Excel may display:
  // "We found a problem with some content..."
  zip.remove("xl/calcChain.xml");

  const workbookRelationshipsFile = zip.file(
    "xl/_rels/workbook.xml.rels",
  );

  if (workbookRelationshipsFile) {
    const relationshipsDocument = parser.parseFromString(
      await workbookRelationshipsFile.async("string"),
      "application/xml",
    );

    Array.from(
      relationshipsDocument.getElementsByTagNameNS("*", "Relationship"),
    ).forEach((relationship) => {
      const type = relationship.getAttribute("Type") || "";
      const target = relationship.getAttribute("Target") || "";

      if (
        type.endsWith("/calcChain") ||
        target.endsWith("calcChain.xml")
      ) {
        relationship.parentNode?.removeChild(relationship);
      }
    });

    zip.file(
      "xl/_rels/workbook.xml.rels",
      serializer.serializeToString(relationshipsDocument),
    );
  }

  const contentTypesFile = zip.file("[Content_Types].xml");

  if (contentTypesFile) {
    const contentTypesDocument = parser.parseFromString(
      await contentTypesFile.async("string"),
      "application/xml",
    );

    Array.from(
      contentTypesDocument.getElementsByTagNameNS("*", "Override"),
    ).forEach((override) => {
      if (
        (override.getAttribute("PartName") || "") ===
        "/xl/calcChain.xml"
      ) {
        override.parentNode?.removeChild(override);
      }
    });

    zip.file(
      "[Content_Types].xml",
      serializer.serializeToString(contentTypesDocument),
    );
  }

  // Force Excel to recalculate formulas from the workbook itself rather
  // than relying on the removed/stale calculation chain.
  const workbookFile = zip.file("xl/workbook.xml");

  if (workbookFile) {
    const workbookDocument = parser.parseFromString(
      await workbookFile.async("string"),
      "application/xml",
    );

    const workbookElement =
      workbookDocument.getElementsByTagNameNS(
        XLSX_MAIN_NAMESPACE,
        "workbook",
      )[0];

    if (workbookElement) {
      let calcPr = workbookDocument.getElementsByTagNameNS(
        XLSX_MAIN_NAMESPACE,
        "calcPr",
      )[0];

      if (!calcPr) {
        calcPr = workbookDocument.createElementNS(
          XLSX_MAIN_NAMESPACE,
          "calcPr",
        );

        const extLst = workbookDocument.getElementsByTagNameNS(
          XLSX_MAIN_NAMESPACE,
          "extLst",
        )[0];

        if (extLst?.parentNode === workbookElement) {
          workbookElement.insertBefore(calcPr, extLst);
        } else {
          workbookElement.appendChild(calcPr);
        }
      }

      calcPr.setAttribute("calcMode", "auto");
      calcPr.setAttribute("fullCalcOnLoad", "1");
      calcPr.setAttribute("forceFullCalc", "1");
    }

    zip.file(
      "xl/workbook.xml",
      serializer.serializeToString(workbookDocument),
    );
  }
}

async function writePreservedXlsxSheet(
  zip: JSZip,
  worksheetPath: string,
  worksheetDocument: Document,
  stylesDocument?: Document,
) {
  const serializer = new XMLSerializer();
  zip.file(worksheetPath, serializer.serializeToString(worksheetDocument));

  if (stylesDocument) {
    zip.file(
      "xl/styles.xml",
      serializer.serializeToString(stylesDocument),
    );
  }

  // Preserve every original workbook part (images, formatting, merged cells,
  // print settings, relationships, shared strings, etc.). Grade 11 edits only
  // the worksheet values and the few cloned cell styles needed for alignment.
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}


function excelColumnNumber(columnLetters: string) {
  return columnLetters
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function excelColumnLetters(columnNumber: number) {
  let value = columnNumber;
  let result = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}

function parseExcelRangeAddress(address: string) {
  const match = address.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return null;

  return {
    startColumn: excelColumnNumber(match[1]),
    startRow: Number(match[2]),
    endColumn: excelColumnNumber(match[3]),
    endRow: Number(match[4]),
  };
}

function cloneExcelStyle<T>(value: T): T {
  if (value == null) return value;
  return JSON.parse(JSON.stringify(value)) as T;
}

function copyExcelBlock(
  worksheet: ExcelJS.Worksheet,
  sourceStartRow: number,
  sourceEndRow: number,
  targetStartRow: number,
  startColumn: number,
  endColumn: number,
) {
  const rowOffset = targetStartRow - sourceStartRow;

  const worksheetModel = (
    worksheet as unknown as {
      model?: { merges?: string[] };
    }
  ).model;

  const sourceMerges = (worksheetModel?.merges ?? [])
    .map((address) => {
      const parsed = parseExcelRangeAddress(address);
      return parsed ? { address, parsed } : null;
    })
    .filter(
      (
        item,
      ): item is {
        address: string;
        parsed: {
          startColumn: number;
          startRow: number;
          endColumn: number;
          endRow: number;
        };
      } => item !== null,
    )
    .filter(
      ({ parsed }) =>
        parsed.startRow >= sourceStartRow &&
        parsed.endRow <= sourceEndRow &&
        parsed.startColumn >= startColumn &&
        parsed.endColumn <= endColumn,
    );

  for (let row = sourceStartRow; row <= sourceEndRow; row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      const sourceCell = worksheet.getCell(row, column);
      const targetCell = worksheet.getCell(row + rowOffset, column);

      targetCell.value = sourceCell.value;
      targetCell.style = cloneExcelStyle(sourceCell.style);

      if (sourceCell.numFmt) {
        targetCell.numFmt = sourceCell.numFmt;
      }

      targetCell.alignment = cloneExcelStyle(sourceCell.alignment);
      targetCell.font = cloneExcelStyle(sourceCell.font);
      targetCell.fill = cloneExcelStyle(sourceCell.fill);
      targetCell.border = cloneExcelStyle(sourceCell.border);
      targetCell.protection = cloneExcelStyle(sourceCell.protection);
    }
  }

  sourceMerges.forEach(({ parsed }) => {
    const targetAddress =
      `${excelColumnLetters(parsed.startColumn)}${parsed.startRow + rowOffset}:` +
      `${excelColumnLetters(parsed.endColumn)}${parsed.endRow + rowOffset}`;

    try {
      worksheet.mergeCells(targetAddress);
    } catch {
      // Ignore duplicate/overlapping merge ranges.
    }
  });
}

function clearExcelBlock(
  worksheet: ExcelJS.Worksheet,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
) {
  const worksheetModel = (
    worksheet as unknown as {
      model?: { merges?: string[] };
    }
  ).model;

  const mergesToRemove = (worksheetModel?.merges ?? [])
    .map((address) => {
      const parsed = parseExcelRangeAddress(address);
      return parsed ? { address, parsed } : null;
    })
    .filter(
      (
        item,
      ): item is {
        address: string;
        parsed: {
          startColumn: number;
          startRow: number;
          endColumn: number;
          endRow: number;
        };
      } => item !== null,
    )
    .filter(
      ({ parsed }) =>
        parsed.startRow >= startRow &&
        parsed.endRow <= endRow &&
        parsed.startColumn >= startColumn &&
        parsed.endColumn <= endColumn,
    );

  mergesToRemove.forEach(({ address }) => {
    try {
      worksheet.unMergeCells(address);
    } catch {
      // Ignore already-unmerged ranges.
    }
  });

  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      const cell = worksheet.getCell(row, column);
      cell.value = null;
      cell.style = {};
    }
  }
}

function getGeneralAverageRemark(
  generalAverage: number | null | undefined,
) {
  if (generalAverage == null) return "";
  return generalAverage >= 75 && generalAverage <= 100
    ? "Promoted"
    : "Failed";
}

function calculateLearnerAge(birthdate: string | null | undefined) {
  if (!birthdate) return "";

  const date = new Date(birthdate);
  if (Number.isNaN(date.getTime())) return "";

  const today = new Date();
  let age = today.getFullYear() - date.getFullYear();
  const monthDifference = today.getMonth() - date.getMonth();

  if (
    monthDifference < 0 ||
    (monthDifference === 0 && today.getDate() < date.getDate())
  ) {
    age -= 1;
  }

  return age >= 0 ? String(age) : "";
}



function oklchToRgba(match: string) {
  const parsed = match.match(
    /oklch\(\s*([+-]?(?:\d+\.?\d*|\.\d+)%?)\s+([+-]?(?:\d+\.?\d*|\.\d+))\s+([+-]?(?:\d+\.?\d*|\.\d+))(?:deg)?(?:\s*\/\s*([+-]?(?:\d+\.?\d*|\.\d+)%?))?\s*\)/i,
  );

  if (!parsed) return match;

  const lightness = parsed[1].endsWith("%")
    ? Number.parseFloat(parsed[1]) / 100
    : Number.parseFloat(parsed[1]);
  const chroma = Number.parseFloat(parsed[2]);
  const hueRadians = (Number.parseFloat(parsed[3]) * Math.PI) / 180;
  const alphaText = parsed[4];
  const alpha = alphaText
    ? alphaText.endsWith("%")
      ? Number.parseFloat(alphaText) / 100
      : Number.parseFloat(alphaText)
    : 1;

  const labA = chroma * Math.cos(hueRadians);
  const labB = chroma * Math.sin(hueRadians);

  const lPrime = lightness + 0.3963377774 * labA + 0.2158037573 * labB;
  const mPrime = lightness - 0.1055613458 * labA - 0.0638541728 * labB;
  const sPrime = lightness - 0.0894841775 * labA - 1.291485548 * labB;

  const l = lPrime ** 3;
  const m = mPrime ** 3;
  const s = sPrime ** 3;

  const linearRed =
    4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s;
  const linearGreen =
    -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s;
  const linearBlue =
    -0.0041960863 * l - 0.7034186147 * m + 1.707614701 * s;

  const encode = (value: number) => {
    const encoded =
      value <= 0.0031308
        ? 12.92 * value
        : 1.055 * Math.pow(Math.max(value, 0), 1 / 2.4) - 0.055;

    return Math.round(Math.min(1, Math.max(0, encoded)) * 255);
  };

  return `rgba(${encode(linearRed)}, ${encode(linearGreen)}, ${encode(
    linearBlue,
  )}, ${Math.min(1, Math.max(0, alpha))})`;
}

function replaceUnsupportedColors(value: string) {
  if (!value.includes("oklch(")) return value;

  return value.replace(
    /oklch\(\s*[+-]?(?:\d+\.?\d*|\.\d+)%?\s+[+-]?(?:\d+\.?\d*|\.\d+)\s+[+-]?(?:\d+\.?\d*|\.\d+)(?:deg)?(?:\s*\/\s*[+-]?(?:\d+\.?\d*|\.\d+)%?)?\s*\)/gi,
    (match) => oklchToRgba(match),
  );
}

async function imageSourceToDataUrl(source: string) {
  if (!source || source.startsWith("data:")) return source;

  const response = await fetch(source, { credentials: "same-origin" });
  if (!response.ok) {
    throw new Error(`Unable to load image: ${source}`);
  }

  const blob = await response.blob();

  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () =>
      reject(reader.error || new Error("Unable to read image."));
    reader.readAsDataURL(blob);
  });
}

const HTML2CANVAS_COLOR_PROPERTIES = [
  "color",
  "background-color",
  "border-top-color",
  "border-right-color",
  "border-bottom-color",
  "border-left-color",
  "outline-color",
  "text-decoration-color",
  "column-rule-color",
  "caret-color",
  "fill",
  "stroke",
  "box-shadow",
  "text-shadow",
] as const;

function applyHtml2CanvasSafeColors(source: Element, target: Element) {
  if (source instanceof HTMLElement && target instanceof HTMLElement) {
    const computedStyle = window.getComputedStyle(source);

    HTML2CANVAS_COLOR_PROPERTIES.forEach((property) => {
      const originalValue = computedStyle.getPropertyValue(property);
      if (!originalValue) return;

      try {
        target.style.setProperty(
          property,
          replaceUnsupportedColors(originalValue),
          computedStyle.getPropertyPriority(property),
        );
      } catch {
        // Ignore browser-only properties that cannot be assigned inline.
      }
    });

    const backgroundImage = computedStyle.getPropertyValue("background-image");
    if (backgroundImage.includes("oklch(")) {
      target.style.backgroundImage = "none";
    }
  }

  const sourceChildren = Array.from(source.children);
  const targetChildren = Array.from(target.children);

  sourceChildren.forEach((sourceChild, index) => {
    const targetChild = targetChildren[index];
    if (targetChild) {
      applyHtml2CanvasSafeColors(sourceChild, targetChild);
    }
  });
}



async function waitForSf9Images(pageElement: HTMLElement) {
  const images = Array.from(
    pageElement.querySelectorAll("img"),
  ) as HTMLImageElement[];

  await Promise.all(
    images.map(async (image) => {
      if (image.complete && image.naturalWidth > 0) return;

      await new Promise<void>((resolve) => {
        const finish = () => resolve();
        image.addEventListener("load", finish, { once: true });
        image.addEventListener("error", finish, { once: true });
      });
    }),
  );
}

const SF9_EXPORT_WIDTH = 1224;
const SF9_EXPORT_HEIGHT = 800;

// Grade 11 SHS uses the uploaded Excel template's A4 landscape page setup.
// PdfPreviewShell renders A4 landscape at approximately 1124 × 795 px.
const SF9_GRADE11_A4_WIDTH = 1124;
const SF9_GRADE11_A4_HEIGHT = 795;

async function renderSf9PageToPng(
  pageElement: HTMLElement,
  targetPageWidth: number,
  targetPageHeight: number,
) {
  await document.fonts?.ready;
  await waitForSf9Images(pageElement);
  await new Promise<void>((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
  );

  const exportHost = document.createElement("div");
  exportHost.setAttribute("aria-hidden", "true");
  Object.assign(exportHost.style, {
    position: "fixed",
    left: "-20000px",
    top: "0",
    width: `${SF9_EXPORT_WIDTH}px`,
    height: `${SF9_EXPORT_HEIGHT}px`,
    overflow: "hidden",
    background: "#ffffff",
    pointerEvents: "none",
    zIndex: "-1",
  });

  const exportPage = pageElement.cloneNode(true) as HTMLElement;
  exportPage.dataset.sf9ExportPage = "true";
  Object.assign(exportPage.style, {
    width: `${SF9_EXPORT_WIDTH}px`,
    height: `${SF9_EXPORT_HEIGHT}px`,
    minWidth: `${SF9_EXPORT_WIDTH}px`,
    minHeight: `${SF9_EXPORT_HEIGHT}px`,
    maxWidth: "none",
    maxHeight: "none",
    margin: "0",
    transform: "none",
    transformOrigin: "top left",
    background: "#ffffff",
    fontFamily: '"Times New Roman", Times, serif',
    letterSpacing: "normal",
    wordSpacing: "normal",
    fontKerning: "normal",
    textRendering: "geometricPrecision",
  });

  exportHost.appendChild(exportPage);
  document.body.appendChild(exportHost);

  try {
    const exportImages = Array.from(
      exportPage.querySelectorAll("img"),
    ) as HTMLImageElement[];

    await Promise.all(
      exportImages.map(async (image) => {
        const source = image.currentSrc || image.src;
        if (!source) return;

        try {
          image.src = await imageSourceToDataUrl(source);
          image.removeAttribute("srcset");
          image.removeAttribute("sizes");
          image.removeAttribute("crossorigin");

          if ("decode" in image) {
            try {
              await image.decode();
            } catch {
              // The image can still be drawn by html2canvas.
            }
          }
        } catch (error) {
          console.warn("Unable to embed an SF9 image.", error);
        }
      }),
    );

    await waitForSf9Images(exportPage);
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );

    const captureScale = 3;
    const sourceCanvas = await html2canvas(exportPage, {
      backgroundColor: "#ffffff",
      scale: captureScale,
      useCORS: true,
      allowTaint: false,
      logging: false,
      foreignObjectRendering: false,
      imageTimeout: 15000,
      removeContainer: true,
      width: SF9_EXPORT_WIDTH,
      height: SF9_EXPORT_HEIGHT,
      windowWidth: SF9_EXPORT_WIDTH,
      windowHeight: SF9_EXPORT_HEIGHT,
      scrollX: 0,
      scrollY: 0,
      onclone: (clonedDocument) => {
        const clonedPage = clonedDocument.body.querySelector(
          '[data-sf9-export-page="true"]',
        ) as HTMLElement | null;

        if (!clonedPage) {
          throw new Error("Unable to prepare the SF9 bondpaper for Word.");
        }

        applyHtml2CanvasSafeColors(exportPage, clonedPage);

        Object.assign(clonedPage.style, {
          width: `${SF9_EXPORT_WIDTH}px`,
          height: `${SF9_EXPORT_HEIGHT}px`,
          minWidth: `${SF9_EXPORT_WIDTH}px`,
          minHeight: `${SF9_EXPORT_HEIGHT}px`,
          maxWidth: "none",
          maxHeight: "none",
          margin: "0",
          transform: "none",
          transformOrigin: "top left",
          backgroundColor: "#ffffff",
          fontFamily: '"Times New Roman", Times, serif',
          letterSpacing: "normal",
          wordSpacing: "normal",
          fontKerning: "normal",
          textRendering: "geometricPrecision",
        });

        clonedPage.querySelectorAll("img").forEach((image) => {
          image.removeAttribute("srcset");
          image.removeAttribute("sizes");
          image.removeAttribute("crossorigin");
        });

        clonedDocument.documentElement.style.backgroundColor = "#ffffff";
        clonedDocument.body.style.backgroundColor = "#ffffff";
        clonedDocument.body.style.margin = "0";
      },
    });

    if (sourceCanvas.width === 0 || sourceCanvas.height === 0) {
      throw new Error("The SF9 bondpaper could not be captured for Word.");
    }

    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = Math.round(targetPageWidth * captureScale);
    pageCanvas.height = Math.round(targetPageHeight * captureScale);

    const context = pageCanvas.getContext("2d");
    if (!context) throw new Error("Unable to create the SF9 page canvas.");

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = "high";

    const safeMargin = 3 * captureScale;
    const availableWidth = pageCanvas.width - safeMargin * 2;
    const availableHeight = pageCanvas.height - safeMargin * 2;
    const fitScale = Math.min(
      availableWidth / sourceCanvas.width,
      availableHeight / sourceCanvas.height,
    );

    const drawWidth = sourceCanvas.width * fitScale;
    const drawHeight = sourceCanvas.height * fitScale;
    const drawX = (pageCanvas.width - drawWidth) / 2;
    const drawY = (pageCanvas.height - drawHeight) / 2;

    context.drawImage(sourceCanvas, drawX, drawY, drawWidth, drawHeight);

    const pngBlob = await new Promise<Blob>((resolve, reject) => {
      pageCanvas.toBlob(
        (blob) =>
          blob
            ? resolve(blob)
            : reject(new Error("Unable to create the SF9 Word image.")),
        "image/png",
        1,
      );
    });

    return {
      data: new Uint8Array(await pngBlob.arrayBuffer()),
      width: targetPageWidth,
      height: targetPageHeight,
    };
  } finally {
    exportHost.remove();
  }
}


type Sf9PreviewLearningRow = {
  label: string;
  indented: boolean;
  term1: number | null;
  term2: number | null;
  term3: number | null;
  final: number | null;
};

type Sf9AcademicElectiveGrades = {
  term1: number | null;
  term2: number | null;
  term3: number | null;
};

function Grade11ShsSf9Preview({
  learner,
  klass,
  learnerRows,
  academicElectiveGrades,
  academicElectiveUnits,
  generalAverage,
  profile,
  regionText,
  divisionText,
  districtText,
  municipalityText,
  schoolNameText,
}: {
  learner: StudentRow;
  klass: SF9ClassRow | undefined;
  learnerRows: Sf9PreviewLearningRow[];
  academicElectiveGrades: Sf9AcademicElectiveGrades;
  academicElectiveUnits: number | null;
  generalAverage: number | null;
  profile: any;
  regionText: string;
  divisionText: string;
  districtText: string;
  municipalityText: string;
  schoolNameText: string;
}) {
  // Grade 11 SF9 should show only real subjects.
  // Do not render placeholder/empty learning-area rows.
  const visibleSubjects = learnerRows
    .filter((row) => Boolean(row.label?.trim()))
    .slice(0, 7);

  const unitValue =
    klass?.units == null || !Number.isFinite(Number(klass.units))
      ? ""
      : String(klass.units);

  const learnerName = [
    learner.last_name,
    learner.first_name,
    learner.middle_name,
  ]
    .filter(Boolean)
    .join(", ")
    .toUpperCase();

  const rowHeight = 20;

  return (
    <div
      className="grid h-full w-full gap-[22px] text-black"
      style={{
        gridTemplateColumns: "1fr 1fr",
        fontFamily: '"Bookman Old Style", "Times New Roman", serif',
        fontSize: "7.1px",
        lineHeight: 1.05,
      }}
    >
      {/* LEFT HALF — follows Grade-11-SHS.xlsx rows A2:O53 */}
      <section className="flex h-full min-h-0 flex-col border-2 border-black px-[9px] py-[7px]">
        <div className="grid grid-cols-[70px_1fr_74px] items-center">
          <img
            src={depedLogo}
            alt="Department of Education seal"
            className="mx-auto h-[61px] w-[61px] object-contain"
          />

          <div className="text-center leading-[1.03]">
            <div>Republic of the Philippines</div>
            <div>Department of Education</div>
            <div>{regionText}</div>
            <div className="font-bold uppercase">
              SCHOOLS DIVISION OFFICE OF {divisionText}
            </div>
            <div>{districtText}</div>
            <div>{municipalityText}</div>
          </div>

          <img
            src={schoolLogo}
            alt="School seal"
            className="mx-auto h-[66px] w-[66px] object-contain"
          />
        </div>

        <div className="mt-[2px] text-center text-[8.4px] font-bold uppercase">
          {schoolNameText}
        </div>

        <div className="mt-[5px] text-center text-[10.5px] font-bold">
          LEARNER&apos;S PERFORMANCE REPORT
        </div>

        <div className="mt-[2px] text-center text-[7.4px]">
          School Year {klass?.school_year || ""}
        </div>

        {/* Match the Excel layout: Name/LRN/Track stay on the left,
            while Age+Sex and Grade+Section appear as two short paired lines
            on the right. */}
        <div className="mt-[7px] space-y-[2px] text-[7.3px]">
          <div className="grid grid-cols-[56px_220px_1fr] items-end gap-x-[4px]">
            <div>Name:</div>
            <div className="flex h-[14px] items-end truncate border-b border-black px-[2px] pb-[1px] font-semibold leading-none">
              {learnerName}
            </div>
            <div className="grid grid-cols-[42px_58px_48px_58px] items-end gap-x-[4px]">
              <div>Age:</div>
              <div className="flex h-[14px] items-end justify-center border-b border-black pb-[1px] text-center leading-none">
                {calculateLearnerAge(learner.birthdate)}
              </div>
              <div>Sex:</div>
              <div className="flex h-[14px] items-end justify-center border-b border-black pb-[1px] text-center capitalize leading-none">
                {learner.sex || ""}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-[56px_220px_1fr] items-end gap-x-[4px]">
            <div>LRN:</div>
            <div className="flex h-[14px] items-end border-b border-black px-[2px] pb-[1px] leading-none">
              {learner.lrn || ""}
            </div>
            <div className="grid grid-cols-[42px_58px_48px_58px] items-end gap-x-[4px]">
              <div>Grade:</div>
              <div className="flex h-[14px] items-end justify-center border-b border-black pb-[1px] text-center leading-none">
                {klass?.grade_level || ""}
              </div>
              <div>Section:</div>
              <div className="flex h-[14px] items-end justify-center border-b border-black pb-[1px] text-center leading-none">
                {klass?.section || ""}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-[88px_142px] items-end gap-x-[4px]">
            <div className="whitespace-nowrap">Track (SHS only):</div>
            <div className="flex h-[14px] items-end border-b border-black px-[2px] pb-[1px] leading-none">
              {klass?.track_shs || ""}
            </div>
          </div>
        </div>

        <div className="mt-[7px] leading-[1.12]">
          <div>Dear Parents,</div>
          <div className="pl-[34px]">
            This Performance Report presents your child&apos;s progress and achievement
          </div>
          <div>in the different learning areas.</div>
          <div className="pl-[34px]">
            The school welcomes you to reach out should you wish to know more about your
          </div>
          <div>child&apos;s learning and performance.</div>
        </div>

        <div className="mt-[8px] grid grid-cols-2 gap-[36px] px-[8px] text-center">
          <div>
            <div className="h-[13px] border-b border-black">
              {profile?.principal || ""}
            </div>
            <div className="mt-[1px]">School Head</div>
          </div>
          <div>
            <div className="h-[13px] border-b border-black">
              {klass?.teacher_name || profile?.full_name || ""}
            </div>
            <div className="mt-[1px]">Adviser</div>
          </div>
        </div>

        <div className="mt-[7px] text-center text-[8.3px] font-bold">
          LEARNING PROGRESS AND ACHIEVEMENT
        </div>

        <table
          className="mt-[2px] w-full table-fixed border-collapse border border-black"
          style={{ fontSize: "6.7px", lineHeight: 1 }}
        >
          <colgroup>
            <col style={{ width: "36%" }} />
            <col style={{ width: "8.5%" }} />
            <col style={{ width: "8.5%" }} />
            <col style={{ width: "8.5%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "9.5%" }} />
            <col style={{ width: "20%" }} />
          </colgroup>
          <thead>
            <tr style={{ height: `${rowHeight}px` }}>
              <th
                rowSpan={2}
                className="border border-black px-[2px] text-center font-bold"
              >
                Learning Areas
              </th>
              <th
                colSpan={3}
                className="border border-black px-[1px] text-center font-bold"
              >
                TERM
              </th>
              <th
                rowSpan={2}
                className="border border-black px-[1px] text-center font-bold"
              >
                Units
              </th>
              <th
                rowSpan={2}
                className="border border-black px-[1px] text-center font-bold"
              >
                Final
                <br />
                Grade
              </th>
              <th
                rowSpan={2}
                className="border border-black px-[1px] text-center font-bold"
              >
                Remarks
              </th>
            </tr>
            <tr style={{ height: `${rowHeight}px` }}>
              <th className="border border-black text-center">1</th>
              <th className="border border-black text-center">2</th>
              <th className="border border-black text-center">3</th>
            </tr>
          </thead>

          <tbody>
            <tr style={{ height: `${rowHeight}px` }}>
              <td className="border border-black px-[4px] font-bold italic">
                Elective Subjects
              </td>
              <td className="border border-black" />
              <td className="border border-black" />
              <td className="border border-black" />
              <td className="border border-black" />
              <td className="border border-black" />
              <td className="border border-black" />
            </tr>

            {visibleSubjects.map((row, index) => (
              <tr key={`grade11-subject-${row.label}-${index}`} style={{ height: `${rowHeight}px` }}>
                <td className="border border-black px-[3px]">
                  {row.label}
                </td>
                <td className="border border-black text-center">
                  {row.term1 ?? ""}
                </td>
                <td className="border border-black text-center">
                  {row.term2 ?? ""}
                </td>
                <td className="border border-black text-center">
                  {row.term3 ?? ""}
                </td>
                <td className="border border-black text-center">
                  {index === 0 ? unitValue : ""}
                </td>
                <td className="border border-black text-center font-semibold">
                  {row.final ?? ""}
                </td>
                <td className="border border-black text-center">
                  {row.final == null
                    ? ""
                    : row.final >= 75
                      ? "Passed"
                      : "Failed"}
                </td>
              </tr>
            ))}

            {[
              {
                label: "Academic Elective 1",
                shade: [false, true, true],
                grade: academicElectiveGrades.term1,
              },
              {
                label: "Academic Elective 2",
                shade: [true, false, true],
                grade: academicElectiveGrades.term2,
              },
              {
                label: "Academic Elective 3",
                shade: [true, true, false],
                grade: academicElectiveGrades.term3,
              },
            ].map((row) => (
              <tr key={row.label} style={{ height: `${rowHeight}px` }}>
                <td className="border border-black px-[3px]">{row.label}</td>
                {row.shade.map((isGray, termIndex) => (
                  <td
                    key={`${row.label}-term-${termIndex + 1}`}
                    className={
                      isGray
                        ? "border border-black bg-black/[0.18]"
                        : "border border-black text-center font-semibold"
                    }
                  >
                    {isGray ? "" : (row.grade ?? "")}
                  </td>
                ))}
                <td className="border border-black text-center">
                  {row.grade == null ? "" : (academicElectiveUnits ?? "")}
                </td>
                <td className="border border-black" />
                <td className="border border-black" />
              </tr>
            ))}

            <tr style={{ height: `${rowHeight}px` }} className="font-bold italic">
              <td colSpan={4} className="border border-black px-[3px] text-center">
                General Average
              </td>
              <td className="border border-black" />
              <td className="border border-black text-center">
                {generalAverage ?? ""}
              </td>
              <td className="border border-black text-center">
                {getGeneralAverageRemark(generalAverage)}
              </td>
            </tr>
          </tbody>
        </table>

        {/* Grade 11 Excel places the Performance Descriptors only a few
            spreadsheet rows below General Average. Do not push this section
            to the bottom of the bond paper with mt-auto. */}
        <div className="mt-[18px]">
          <div className="text-[7.4px] font-bold uppercase">
            PERFORMANCE DESCRIPTORS
          </div>

          <div className="mt-[3px] grid grid-cols-[116px_1fr] items-start gap-[7px]">
            <div />
            <table
              className="w-full border-collapse"
              style={{ fontSize: "6.2px", lineHeight: 1 }}
            >
              <thead>
                <tr>
                  <th className="px-[2px] py-[1px]">
                    Grading Scale
                  </th>
                  <th className="px-[2px] py-[1px]">
                    Descriptors
                  </th>
                  <th className="px-[2px] py-[1px]">
                    Remarks
                  </th>
                </tr>
              </thead>
              <tbody>
                {SF9_PERFORMANCE_DESCRIPTORS.map((descriptor) => (
                  <tr key={descriptor.scale}>
                    <td className="px-[2px] py-[1px] text-center">
                      {descriptor.scale.replace(/\s/g, "")}
                    </td>
                    <td className="px-[2px] py-[1px] text-center">
                      {descriptor.description}
                    </td>
                    <td className="px-[2px] py-[1px] text-center">
                      {descriptor.remarks}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {/* RIGHT HALF — follows Grade-11-SHS.xlsx rows Q2:AD53 */}
      <section className="flex h-full min-h-0 flex-col border-2 border-black px-[9px] py-[7px]">
        <div className="text-center text-[9px] font-bold">ATTENDANCE RECORD</div>

        {/* Grade 11 attendance table follows the uploaded Excel format:
            simple straight grid lines, horizontal month labels, and no
            diagonal header separators. */}
        <table
          className="mt-[3px] w-full table-fixed border-collapse border border-black"
          style={{ fontSize: "6.2px", lineHeight: 1.05 }}
        >
          <colgroup>
            <col style={{ width: "20%" }} />
            {SF9_ATTENDANCE_MONTHS.map((month) => (
              <col key={month} style={{ width: "6.3%" }} />
            ))}
            <col style={{ width: "10.7%" }} />
          </colgroup>

          <thead>
            <tr style={{ height: "22px" }}>
              <th className="border border-black px-[2px] text-center text-[6.5px] font-bold">
                Month
              </th>

              {SF9_ATTENDANCE_MONTHS.map((month) => (
                <th
                  key={month}
                  className="border border-black px-[1px] text-center text-[6.2px] font-bold"
                >
                  {month.slice(0, 3)}
                </th>
              ))}

              <th className="border border-black px-[1px] text-center text-[6.2px] font-bold">
                Total
              </th>
            </tr>
          </thead>

          <tbody>
            {[
              "No. of Class Days",
              "No. of Days Present",
              "No. of Days Absent",
            ].map((label) => (
              <tr key={label} style={{ height: "31px" }}>
                <td className="border border-black px-[3px] text-left text-[6.1px] font-bold leading-[1.05]">
                  {label}
                </td>

                {SF9_ATTENDANCE_MONTHS.map((month) => (
                  <td key={month} className="border border-black" />
                ))}

                <td className="border border-black" />
              </tr>
            ))}
          </tbody>
        </table>

        <div className="mt-[8px] text-center text-[8.4px] font-bold">
          TEACHER&apos;S COMMENTS / REMARKS
        </div>

        <div className="mt-[3px] grid grid-cols-[44px_1fr]">
          <div className="grid grid-rows-3 text-[7.5px] font-bold">
            {["Term 1", "Term 2", "Term 3"].map((term) => (
              <div key={term} className="flex items-start pt-[5px]">
                {term}
              </div>
            ))}
          </div>
          <div className="border border-black">
            <div className="h-[70px] border-b border-black" />
            <div className="h-[70px] border-b border-black" />
            <div className="h-[70px]" />
          </div>
        </div>

        <div className="mt-[9px] text-center text-[8px] font-bold">
          PARENT/S GUARDIAN&apos;S SIGNATURE
        </div>

        <div className="mt-[6px] space-y-[6px] px-[45px] text-[7px]">
          {["Term 1", "Term 2", "Term 3"].map((term) => (
            <div key={term} className="grid grid-cols-[43px_1fr] items-end gap-[5px]">
              <div>{term}</div>
              <div className="h-[10px] border-b border-black" />
            </div>
          ))}
        </div>

        <div className="mt-[11px] text-center text-[8.5px] font-bold">
          CERTIFICATE OF TRANSFER
        </div>

        <div className="mt-[5px] text-[6.5px] italic leading-[1.3]">
          <div>
            This is to certify that the above-named learner has satisfactorily completed
            the requirements
          </div>
          <div>for the grade level indicated.</div>
        </div>

        <div className="mt-[13px] space-y-[9px] text-[6.8px]">
          <div className="flex items-end gap-[6px]">
            <div className="w-[100px] shrink-0">Admitted to Grade:</div>
            <div className="h-[8px] w-[105px] shrink-0 border-b border-black" />
          </div>

          <div className="flex items-end gap-[6px]">
            <div className="w-[100px] shrink-0 leading-[1.05]">
              Eligible for Admission to Grade:
            </div>
            <div className="h-[8px] w-[105px] shrink-0 border-b border-black" />
          </div>
        </div>

        <div className="mt-[10px] grid grid-cols-[48px_1fr_1fr] items-end gap-[8px] text-[6.8px]">
          <div>Approved:</div>
          <div className="border-b border-black" />
          <div>
            <div className="h-[12px] border-b border-black text-center">
              {klass?.teacher_name || profile?.full_name || ""}
            </div>
            <div className="mt-[1px] text-center">Adviser</div>
          </div>
        </div>

        <div className="mt-[10px] mx-auto w-[54%] text-center text-[6.8px]">
          <div className="h-[12px] border-b border-black">
            {profile?.principal || ""}
          </div>
          <div className="mt-[1px]">School Head</div>
        </div>

        <div className="mt-[12px]">
          <div className="text-center text-[7.8px] font-bold">
            CANCELLATION OF ELIGIBILITY TO TRANSFER
          </div>

          <div className="mt-[8px] grid grid-cols-[57px_1fr_29px_1fr] items-end gap-[5px] text-[6.6px]">
            <div>Admitted in:</div>
            <div className="border-b border-black" />
            <div>Date:</div>
            <div className="border-b border-black" />
          </div>

          <div className="mx-auto mt-[12px] w-[48%] text-center text-[6.6px]">
            <div className="h-[12px] border-b border-black">
              {profile?.principal || ""}
            </div>
            <div className="mt-[1px]">School Head</div>
          </div>
        </div>
      </section>
    </div>
  );
}


function Grade12ShsSf9Preview({
  learner,
  klass,
  subjectRows,
  generalAverage,
  profile,
  regionText,
  divisionText,
  districtText,
  municipalityText,
  schoolNameText,
}: {
  learner: StudentRow;
  klass: SF9ClassRow | undefined;
  subjectRows: Sf9Grade12PreviewRow[];
  generalAverage: number | null;
  profile: any;
  regionText: string;
  divisionText: string;
  districtText: string;
  municipalityText: string;
  schoolNameText: string;
}) {
  const learnerName = [
    learner.last_name,
    learner.first_name,
    learner.middle_name,
  ]
    .filter(Boolean)
    .join(", ")
    .toUpperCase();

  const coreRows = subjectRows.filter((row) => row.group === "core");
  const appliedRows = subjectRows.filter((row) => row.group === "applied");

  const renderSubjectRow = (row: Sf9Grade12PreviewRow) => (
    <tr key={row.label} className="h-[16px]">
      <td className="border border-black px-[3px] text-left">
        {row.label}
      </td>

      {(["1", "2", "3"] as const).map((termValue) => {
        const isAvailable = row.term === termValue;

        return (
          <td
            key={`${row.label}-${termValue}`}
            className={
              isAvailable
                ? "border border-black text-center font-semibold"
                : "border border-black bg-black/[0.22]"
            }
          >
            {isAvailable ? (row.grade ?? "") : ""}
          </td>
        );
      })}

      <td className="border border-black text-center">
        {row.units ?? ""}
      </td>

      <td className="border border-black text-center font-bold">
        {row.grade ?? ""}
      </td>

      <td className="border border-black text-center">
        {row.grade == null ? "" : row.grade >= 75 ? "Passed" : "Failed"}
      </td>
    </tr>
  );

  return (
    <div
      className="grid h-full w-full gap-[20px] text-black"
      style={{
        gridTemplateColumns: "1fr 1fr",
        fontFamily: '"Arial", "Bookman Old Style", "Times New Roman", serif',
        fontSize: "6.65px",
        lineHeight: 1.03,
      }}
    >
      {/* LEFT HALF — Grade-12-SHS.xlsx SF9 layout */}
      <section className="flex h-full min-h-0 flex-col border-2 border-black px-[8px] py-[6px]">
        <div className="grid grid-cols-[68px_1fr_70px] items-center">
          <img
            src={depedLogo}
            alt="Department of Education seal"
            className="mx-auto h-[58px] w-[58px] object-contain"
          />

          <div className="text-center leading-[1.04]">
            <div>Republic of the Philippines</div>
            <div>Department of Education</div>
            <div>{regionText}</div>
            <div className="font-bold uppercase">
              SCHOOLS DIVISION OFFICE OF {divisionText}
            </div>
            <div>{districtText}</div>
            <div>{municipalityText}</div>
            <div className="mt-[2px] font-bold uppercase">
              {schoolNameText}
            </div>
          </div>

          <img
            src={schoolLogo}
            alt="School seal"
            className="mx-auto h-[62px] w-[62px] object-contain"
          />
        </div>

        <div className="mt-[2px] text-center text-[9.8px] font-bold">
          LEARNER&apos;S PERFORMANCE REPORT
        </div>

        <div className="mt-[1px] text-center text-[7px]">
          School Year {klass?.school_year || ""}
        </div>

        <div className="mt-[5px] space-y-[2px] text-[6.8px]">
          <div className="grid grid-cols-[42px_200px_1fr] items-end gap-x-[4px]">
            <div>Name:</div>
            <div className="flex h-[12px] items-end truncate border-b border-black px-[2px] pb-[1px] font-semibold leading-none">
              {learnerName}
            </div>
            <div className="grid grid-cols-[30px_48px_34px_48px] items-end gap-x-[3px]">
              <div>Age:</div>
              <div className="flex h-[12px] items-end justify-center border-b border-black pb-[1px]">
                {calculateLearnerAge(learner.birthdate)}
              </div>
              <div>Sex:</div>
              <div className="flex h-[12px] items-end justify-center border-b border-black pb-[1px] capitalize">
                {learner.sex || ""}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-[42px_200px_1fr] items-end gap-x-[4px]">
            <div>LRN:</div>
            <div className="flex h-[12px] items-end border-b border-black px-[2px] pb-[1px]">
              {learner.lrn || ""}
            </div>
            <div className="grid grid-cols-[30px_48px_34px_48px] items-end gap-x-[3px]">
              <div>Grade:</div>
              <div className="flex h-[12px] items-end justify-center border-b border-black pb-[1px]">
                {klass?.grade_level || ""}
              </div>
              <div>Section:</div>
              <div className="flex h-[12px] items-end justify-center border-b border-black pb-[1px]">
                {klass?.section || ""}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-[76px_130px] items-end gap-x-[4px]">
            <div className="whitespace-nowrap">Track (SHS only):</div>
            <div className="flex h-[12px] items-end border-b border-black px-[2px] pb-[1px]">
              {klass?.track_shs || ""}
            </div>
          </div>
        </div>

        <div className="mt-[5px] leading-[1.08]">
          <div>Dear Parents,</div>
          <div className="pl-[28px]">
            This Performance Report presents your child&apos;s progress and achievement
          </div>
          <div>in the different learning areas.</div>
          <div className="pl-[28px]">
            The school welcomes you to reach out should you wish to know more about your
          </div>
          <div>child&apos;s learning and performance.</div>
        </div>

        <div className="mt-[5px] grid grid-cols-2 gap-[28px] px-[6px] text-center">
          <div>
            <div className="h-[11px] border-b border-black">
              {profile?.principal || ""}
            </div>
            <div>School Head</div>
          </div>
          <div>
            <div className="h-[11px] border-b border-black">
              {klass?.teacher_name || profile?.full_name || ""}
            </div>
            <div>Adviser</div>
          </div>
        </div>

        <div className="mt-[4px] text-center text-[7.8px] font-bold">
          LEARNING PROGRESS AND ACHIEVEMENT
        </div>

        <table
          className="mt-[2px] w-full table-fixed border-collapse border border-black"
          style={{ fontSize: "5.9px", lineHeight: 1 }}
        >
          <colgroup>
            <col style={{ width: "43%" }} />
            <col style={{ width: "8.3%" }} />
            <col style={{ width: "8.3%" }} />
            <col style={{ width: "8.3%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "10%" }} />
            <col style={{ width: "13%" }} />
          </colgroup>

          <thead>
            <tr className="h-[16px]">
              <th rowSpan={2} className="border border-black px-[2px] text-center font-bold">
                Learning Areas
              </th>
              <th colSpan={3} className="border border-black text-center font-bold">
                TERM
              </th>
              <th rowSpan={2} className="border border-black text-center font-bold">
                Units
              </th>
              <th rowSpan={2} className="border border-black text-center font-bold">
                Final
                <br />
                Grade
              </th>
              <th rowSpan={2} className="border border-black text-center font-bold">
                Remarks
              </th>
            </tr>
            <tr className="h-[14px]">
              <th className="border border-black text-center">T1</th>
              <th className="border border-black text-center">T2</th>
              <th className="border border-black text-center">T3</th>
            </tr>
          </thead>

          <tbody>
            <tr className="h-[15px]">
              <td colSpan={7} className="border border-black bg-black/[0.16] px-[3px] font-bold italic">
                Core Subjects
              </td>
            </tr>

            {coreRows.map(renderSubjectRow)}

            <tr className="h-[15px]">
              <td colSpan={7} className="border border-black bg-black/[0.16] px-[3px] font-bold">
                Applied and Specialized Subjects
              </td>
            </tr>

            {appliedRows.map(renderSubjectRow)}

            <tr className="h-[16px] font-bold italic">
              <td colSpan={4} className="border border-black px-[3px] text-center">
                General Average
              </td>
              <td className="border border-black" />
              <td className="border border-black text-center">
                {generalAverage ?? ""}
              </td>
              <td className="border border-black text-center">
                {getGeneralAverageRemark(generalAverage)}
              </td>
            </tr>
          </tbody>
        </table>

        <div className="mt-[8px]">
          <div className="text-[6.8px] font-bold uppercase">
            PERFORMANCE DESCRIPTORS
          </div>

          <table
            className="mt-[2px] ml-[70px] w-[72%] border-collapse"
            style={{ fontSize: "5.7px", lineHeight: 1 }}
          >
            <thead>
              <tr>
                <th className="px-[2px] py-[1px]">Grading Scale</th>
                <th className="px-[2px] py-[1px]">Descriptors</th>
                <th className="px-[2px] py-[1px]">Remarks</th>
              </tr>
            </thead>
            <tbody>
              {SF9_PERFORMANCE_DESCRIPTORS.map((descriptor) => (
                <tr key={descriptor.scale}>
                  <td className="px-[2px] py-[1px] text-center">
                    {descriptor.scale.replace(/\s/g, "")}
                  </td>
                  <td className="px-[2px] py-[1px] text-center">
                    {descriptor.description}
                  </td>
                  <td className="px-[2px] py-[1px] text-center">
                    {descriptor.remarks}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* RIGHT HALF — follows the uploaded Grade 12 SF9 sheet */}
      <section className="flex h-full min-h-0 flex-col border-2 border-black px-[8px] py-[6px]">
        <div className="text-center text-[8.2px] font-bold">
          ATTENDANCE RECORD
        </div>

        <table
          className="mt-[2px] w-full table-fixed border-collapse border border-black"
          style={{ fontSize: "5.8px", lineHeight: 1.02 }}
        >
          <colgroup>
            <col style={{ width: "18%" }} />
            {SF9_ATTENDANCE_MONTHS.map((month) => (
              <col key={month} style={{ width: "6.4%" }} />
            ))}
            <col style={{ width: "11.6%" }} />
          </colgroup>

          <thead>
            <tr className="h-[20px]">
              <th className="border border-black px-[1px] text-center font-bold">
                Month
              </th>
              {SF9_ATTENDANCE_MONTHS.map((month) => (
                <th key={month} className="border border-black px-[1px] text-center font-bold">
                  {month.slice(0, 3)}
                </th>
              ))}
              <th className="border border-black px-[1px] text-center font-bold">
                Total
              </th>
            </tr>
          </thead>

          <tbody>
            {[
              "No. of Class Days",
              "No. of Days Present",
              "No. of Days Absent",
            ].map((label) => (
              <tr key={label} className="h-[28px]">
                <td className="border border-black px-[2px] text-left font-bold">
                  {label}
                </td>
                {SF9_ATTENDANCE_MONTHS.map((month) => (
                  <td key={month} className="border border-black" />
                ))}
                <td className="border border-black" />
              </tr>
            ))}
          </tbody>
        </table>

        <div className="mt-[7px] text-center text-[7.8px] font-bold">
          TEACHER&apos;S COMMENTS / REMARKS
        </div>

        <div className="mt-[2px] grid grid-cols-[40px_1fr]">
          <div className="grid grid-rows-3 text-[6.7px] font-bold">
            {["Term 1", "Term 2", "Term 3"].map((term) => (
              <div key={term} className="flex items-start pt-[4px]">
                {term}
              </div>
            ))}
          </div>
          <div className="border border-black">
            <div className="h-[62px] border-b border-black" />
            <div className="h-[62px] border-b border-black" />
            <div className="h-[62px]" />
          </div>
        </div>

        <div className="mt-[7px] text-center text-[7.5px] font-bold">
          PARENT&apos;S / GUARDIAN&apos;S SIGNATURE
        </div>

        <div className="mt-[5px] space-y-[5px] px-[38px] text-[6.5px]">
          {["Term 1", "Term 2", "Term 3"].map((term) => (
            <div key={term} className="grid grid-cols-[38px_1fr] items-end gap-[5px]">
              <div>{term}</div>
              <div className="h-[9px] border-b border-black" />
            </div>
          ))}
        </div>

        <div className="mt-[9px] text-center text-[8px] font-bold">
          CERTIFICATE OF TRANSFER
        </div>

        <div className="mt-[4px] text-[6px] italic leading-[1.25]">
          <div>
            This is to certify that the above-named learner has satisfactorily completed the requirements
          </div>
          <div>for the grade level indicated.</div>
        </div>

        <div className="mt-[10px] space-y-[8px] text-[6.2px]">
          <div className="flex items-end gap-[5px]">
            <div className="w-[92px] shrink-0">Admitted to Grade:</div>
            <div className="h-[8px] w-[100px] shrink-0 border-b border-black" />
          </div>

          <div className="flex items-end gap-[5px]">
            <div className="w-[92px] shrink-0">
              Eligible for Admission to Grade:
            </div>
            <div className="h-[8px] w-[100px] shrink-0 border-b border-black" />
          </div>
        </div>

        <div className="mt-[9px] grid grid-cols-[43px_1fr_1fr] items-end gap-[7px] text-[6.2px]">
          <div>Approved:</div>
          <div className="border-b border-black" />
          <div>
            <div className="h-[11px] border-b border-black text-center">
              {klass?.teacher_name || profile?.full_name || ""}
            </div>
            <div className="text-center">Adviser</div>
          </div>
        </div>

        <div className="mx-auto mt-[9px] w-[52%] text-center text-[6.2px]">
          <div className="h-[11px] border-b border-black">
            {profile?.principal || ""}
          </div>
          <div>School Head</div>
        </div>

        <div className="mt-[10px]">
          <div className="text-center text-[7.2px] font-bold">
            CANCELLATION OF ELIGIBILITY TO TRANSFER
          </div>

          <div className="mt-[7px] grid grid-cols-[52px_1fr_26px_1fr] items-end gap-[4px] text-[6px]">
            <div>Admitted in:</div>
            <div className="border-b border-black" />
            <div>Date:</div>
            <div className="border-b border-black" />
          </div>

          <div className="mx-auto mt-[11px] w-[48%] text-center text-[6px]">
            <div className="h-[11px] border-b border-black">
              {profile?.principal || ""}
            </div>
            <div>School Head</div>
          </div>
        </div>
      </section>
    </div>
  );
}


function makeSF9Component(variant: Variant) {
  return function SF9Page() {
    const [classId] = useState<string>(readSchoolFormsClassId);
    const [length, setLength] = useState<"short" | "full">("short");
    const [paper, setPaper] = useState<PdfPaper>("long");
    // SF9 is an individual learner report card, so only ONE learner can be
    // selected at a time for Grades 7–12.
    const [selectedLearnerId, setSelectedLearnerId] = useState<string>("");
    const [summarySubjects, setSummarySubjects] = useState<string[]>([]);

    const { data: profile } = useQuery({
      queryKey: ["profile"],
      queryFn: async () => (await supabase.from("profiles").select("*").maybeSingle()).data as any,
    });
    // Use an SF9-specific cache and refetch the selected class directly so
    // Track (SHS only) always comes from the latest database row.
    const { data: classes = [] } = useQuery({
      queryKey: ["sf9-classes"],
      queryFn: async () => {
        const { data, error } = await (supabase as any)
          .from("classes")
          .select("*")
          .order("created_at", { ascending: false });

        if (error) throw error;
        return (data ?? []) as SF9ClassRow[];
      },
      staleTime: 0,
      refetchOnMount: "always",
    });

    const selectedSchoolFormsClass = classes.find((item) => item.id === classId);
    const active = selectedSchoolFormsClass?.id || classes[0]?.id;

    const { data: activeClass } = useQuery({
      enabled: !!active,
      queryKey: ["sf9-class", active],
      queryFn: async () => {
        const { data, error } = await (supabase as any)
          .from("classes")
          .select("*")
          .eq("id", active!)
          .single();

        if (error) throw error;
        return data as SF9ClassRow;
      },
      staleTime: 0,
      refetchOnMount: "always",
    });

    const klass =
      activeClass ?? classes.find((c) => c.id === active);

    const { data: students = [] } = useQuery({
      enabled: !!active,
      queryKey: ["students", active],
      queryFn: async () => (await supabase.from("students").select("*").eq("class_id", active!).order("last_name")).data as StudentRow[],
    });

    // A learner selected from one class must never remain selected after
    // switching to another class.
    useEffect(() => {
      setSelectedLearnerId("");
    }, [active]);

    // Also clear the selection if that learner is no longer part of the
    // currently loaded class roster.
    useEffect(() => {
      if (
        selectedLearnerId &&
        !students.some((student) => student.id === selectedLearnerId)
      ) {
        setSelectedLearnerId("");
      }
    }, [students, selectedLearnerId]);
    const { data: grades = [] } = useQuery({
      enabled: !!active,
      queryKey: ["grades-all", active],
      queryFn: async () => (await supabase.from("grades").select("*").eq("class_id", active!)).data as GradeRow[],
    });

    // Grade 12 SF9 combines the learner's grades from every Grade 12 subject
    // class in the same section and school year.
    const matchingGrade12Classes = useMemo(() => {
      if (
        normalizeSubjectName(klass?.grade_level) !== "grade 12" ||
        !klass?.section ||
        !klass?.school_year
      ) {
        return [] as SF9ClassRow[];
      }

      return classes.filter(
        (candidate) =>
          normalizeSubjectName(candidate.grade_level) === "grade 12" &&
          normalizeSubjectName(candidate.section) ===
            normalizeSubjectName(klass.section) &&
          String(candidate.school_year || "") ===
            String(klass.school_year || ""),
      ) as SF9ClassRow[];
    }, [classes, klass?.grade_level, klass?.section, klass?.school_year]);

    const grade12ClassIds = matchingGrade12Classes.map(
      (candidate) => candidate.id,
    );

    const { data: grade12StudentsData = [] } = useQuery<StudentRow[]>({
      enabled: grade12ClassIds.length > 0,
      queryKey: ["sf9-grade12-students", grade12ClassIds.join(",")],
      queryFn: async () => {
        if (grade12ClassIds.length === 0) return [];
        const { data, error } = await (supabase as any)
          .from("students")
          .select("*")
          .in("class_id", grade12ClassIds);

        if (error) throw error;
        return Array.isArray(data) ? (data as StudentRow[]) : [];
      },
      placeholderData: [],
    });

    const grade12Students = Array.isArray(grade12StudentsData)
      ? grade12StudentsData
      : [];

    const { data: grade12GradesData = [] } = useQuery<GradeRow[]>({
      enabled: grade12ClassIds.length > 0,
      queryKey: ["sf9-grade12-grades", grade12ClassIds.join(",")],
      queryFn: async () => {
        if (grade12ClassIds.length === 0) return [];
        const { data, error } = await (supabase as any)
          .from("grades")
          .select("*")
          .in("class_id", grade12ClassIds);

        if (error) throw error;
        return Array.isArray(data) ? (data as GradeRow[]) : [];
      },
      placeholderData: [],
    });

    const grade12Grades = Array.isArray(grade12GradesData)
      ? grade12GradesData
      : [];

    const grade12StudentIdsByLearnerId = useMemo(() => {
      const result = new Map<string, string[]>();

      const nameKey = (student: StudentRow) =>
        [
          normalizeSf9SubjectKey(student.last_name),
          normalizeSf9SubjectKey(student.first_name),
          normalizeSf9SubjectKey(student.middle_name),
        ].join("|");

      students.forEach((learner) => {
        const learnerLrn = String(learner.lrn || "").trim();
        const learnerName = nameKey(learner);

        const matches = grade12Students
          .filter((candidate) => {
            const candidateLrn = String(candidate.lrn || "").trim();

            if (learnerLrn && candidateLrn) {
              return learnerLrn === candidateLrn;
            }

            return nameKey(candidate) === learnerName;
          })
          .map((candidate) => candidate.id);

        if (!matches.includes(learner.id)) {
          matches.push(learner.id);
        }

        result.set(learner.id, matches);
      });

      return result;
    }, [students, grade12Students, grade12ClassIds.join(",")]);

    const grade12ClassById = useMemo(() => {
      const map = new Map<string, SF9ClassRow>();
      matchingGrade12Classes.forEach((candidate) => {
        map.set(candidate.id, candidate);
      });
      return map;
    }, [matchingGrade12Classes]);

    // Grade 11 Academic Electives are encoded in a dedicated "Elective Subject"
    // class. Academic Elective 1 uses Term 1 only, Academic Elective 2 uses
    // Term 2 only, and Academic Elective 3 uses Term 3 only.
    const matchingAcademicElectiveClasses = useMemo(() => {
      if (
        normalizeSubjectName(klass?.grade_level) !== "grade 11" ||
        !klass?.section ||
        !klass?.school_year
      ) {
        return [] as SF9ClassRow[];
      }

      return classes.filter(
        (candidate) =>
          normalizeSubjectName(candidate.grade_level) === "grade 11" &&
          normalizeSubjectName(candidate.section) ===
            normalizeSubjectName(klass.section) &&
          String(candidate.school_year || "") === String(klass.school_year || "") &&
          isSf9AcademicElectiveSubject(candidate.subject),
      ) as SF9ClassRow[];
    }, [classes, klass?.grade_level, klass?.section, klass?.school_year]);

    const academicElectiveClassIds = matchingAcademicElectiveClasses.map(
      (candidate) => candidate.id,
    );

    // The Grade 11 Elective Subject has one configured unit value
    // (for example 3). That unit is displayed only on an Academic Elective
    // row that actually has a grade. Empty elective rows stay blank.
    const academicElectiveUnits = useMemo(() => {
      const electiveClass =
        matchingAcademicElectiveClasses[0] ??
        (isSf9AcademicElectiveSubject(klass?.subject)
          ? (klass as SF9ClassRow | undefined)
          : undefined);

      if (electiveClass?.units == null) return null;

      const numericUnits = Number(electiveClass.units);
      return Number.isFinite(numericUnits) ? numericUnits : null;
    }, [matchingAcademicElectiveClasses, klass]);

    const { data: academicElectiveStudentsData = [] } = useQuery<StudentRow[]>({
      enabled: academicElectiveClassIds.length > 0,
      queryKey: [
        "sf9-grade11-academic-elective-students",
        academicElectiveClassIds.join(","),
      ],
      queryFn: async () => {
        if (academicElectiveClassIds.length === 0) return [];
        const { data, error } = await (supabase as any)
          .from("students")
          .select("*")
          .in("class_id", academicElectiveClassIds);
        if (error) throw error;
        return Array.isArray(data) ? (data as StudentRow[]) : [];
      },
      placeholderData: [],
    });

    const academicElectiveStudents = Array.isArray(academicElectiveStudentsData)
      ? academicElectiveStudentsData
      : [];

    const { data: academicElectiveGradesData = [] } = useQuery<GradeRow[]>({
      enabled: academicElectiveClassIds.length > 0,
      queryKey: [
        "sf9-grade11-academic-elective-grades",
        academicElectiveClassIds.join(","),
      ],
      queryFn: async () => {
        if (academicElectiveClassIds.length === 0) return [];
        const { data, error } = await (supabase as any)
          .from("grades")
          .select("*")
          .in("class_id", academicElectiveClassIds);
        if (error) throw error;
        return Array.isArray(data) ? (data as GradeRow[]) : [];
      },
      placeholderData: [],
    });

    const academicElectiveGrades = Array.isArray(academicElectiveGradesData)
      ? academicElectiveGradesData
      : [];

    const academicElectiveStudentIdsByLearnerId = useMemo(() => {
      const result = new Map<string, string[]>();

      const nameKey = (student: StudentRow) =>
        [
          normalizeSubjectName(student.last_name),
          normalizeSubjectName(student.first_name),
          normalizeSubjectName(student.middle_name),
        ].join("|");

      students.forEach((learner) => {
        const learnerLrn = String(learner.lrn || "").trim();
        const learnerName = nameKey(learner);

        const matches = academicElectiveStudents
          .filter((candidate) => {
            const candidateLrn = String(candidate.lrn || "").trim();
            if (learnerLrn && candidateLrn) {
              return learnerLrn === candidateLrn;
            }
            return nameKey(candidate) === learnerName;
          })
          .map((candidate) => candidate.id);

        // If the selected SF9 class itself is the Elective Subject class,
        // student IDs already match and this fallback keeps the lookup direct.
        if (matches.length === 0 && academicElectiveClassIds.includes(
          String((learner as any).class_id || ""),
        )) {
          matches.push(learner.id);
        }

        result.set(learner.id, matches);
      });

      return result;
    }, [students, academicElectiveStudents, academicElectiveClassIds.join(",")]);

    const academicElectiveGradeForStudentTerm = (
      studentId: string,
      termValue: "1" | "2" | "3",
    ): number | null => {
      const matchingStudentIds =
        academicElectiveStudentIdsByLearnerId.get(studentId) ?? [studentId];

      const scores = academicElectiveGrades
        .filter(
          (grade) =>
            matchingStudentIds.includes(grade.student_id) &&
            isSf9AcademicElectiveSubject(grade.subject) &&
            grade.term === termValue,
        )
        .map((grade) => grade.score)
        .filter((score): score is number => typeof score === "number");

      return scores.length ? computeAverage(scores) : null;
    };


    // Grade 11 Mabisang Komunikasyon is stored in six separate component
    // scopes, not as plain subject/term rows. Read those same class-record
    // inputs so SF9 shows exactly the same term grades as Summary of Grades.
    const isCommunicationClass =
      String(klass?.grade_level ?? "").trim().toLowerCase() === "grade 11" &&
      isSf9CommunicationCompositeSubject(klass?.subject);
    const communicationSubjectName =
      String(klass?.subject || "Mabisang Komunikasyon").trim() ||
      "Mabisang Komunikasyon";
    const communicationStorageTerms = SF9_COMMUNICATION_SCOPES.map(
      (scope) => scope.storageTerm,
    );
    const communicationSummarySubjects = SF9_COMMUNICATION_SCOPES.map(
      (scope) => `${communicationSubjectName}_${scope.key}_ALL`,
    );

    const { data: communicationComponentsData = [] } = useQuery<GradeComponent[]>({
      enabled: Boolean(active && isCommunicationClass),
      queryKey: ["sf9-communication-components", active],
      queryFn: async () => {
        const { data, error } = await supabase
          .from("grade_components")
          .select("*")
          .eq("class_id", active!)
          .in("term", communicationStorageTerms);
        if (error) throw error;
        return Array.isArray(data) ? (data as GradeComponent[]) : [];
      },
      placeholderData: [],
    });
    const communicationComponents = Array.isArray(communicationComponentsData)
      ? communicationComponentsData
      : [];

    const { data: communicationActivitiesData = [] } = useQuery<GradeActivity[]>({
      enabled: Boolean(active && isCommunicationClass),
      queryKey: ["sf9-communication-activities", active],
      queryFn: async () => {
        const { data, error } = await supabase
          .from("grade_activities")
          .select("*")
          .eq("class_id", active!)
          .in("term", communicationStorageTerms)
          .order("position");
        if (error) throw error;
        return Array.isArray(data) ? (data as GradeActivity[]) : [];
      },
      placeholderData: [],
    });
    const communicationActivities = Array.isArray(communicationActivitiesData)
      ? communicationActivitiesData
      : [];
    const communicationActivityIds = communicationActivities.map(
      (activity) => activity.id,
    );

    const { data: communicationScoresData = [] } = useQuery<ActivityScore[]>({
      enabled: Boolean(
        active && isCommunicationClass && communicationActivityIds.length > 0,
      ),
      queryKey: [
        "sf9-communication-activity-scores",
        active,
        communicationActivityIds.join(","),
      ],
      queryFn: async () => {
        if (communicationActivityIds.length === 0) return [];
        const { data, error } = await supabase
          .from("activity_scores")
          .select("*")
          .in("activity_id", communicationActivityIds);
        if (error) throw error;
        return Array.isArray(data) ? (data as ActivityScore[]) : [];
      },
      placeholderData: [],
    });
    const communicationScores = Array.isArray(communicationScoresData)
      ? communicationScoresData
      : [];

    const { data: communicationSavedBasesData = [] } = useQuery<
      Array<{
        student_id: string;
        subject: string;
        term_grade_base: number | null;
      }>
    >({
      enabled: Boolean(active && isCommunicationClass),
      queryKey: [
        "sf9-communication-term-grade-bases",
        active,
        communicationSubjectName,
      ],
      queryFn: async () => {
        const { data, error } = await (supabase as any)
          .from("grades")
          .select("student_id, subject, term_grade_base")
          .eq("class_id", active!)
          .eq("term", "1")
          .in("subject", communicationSummarySubjects);
        if (error) throw error;
        return Array.isArray(data)
          ? (data as Array<{
              student_id: string;
              subject: string;
              term_grade_base: number | null;
            }>)
          : [];
      },
      placeholderData: [],
    });

    const communicationScoreMap = useMemo(() => {
      const map = new Map<string, number | null>();
      communicationScores.forEach((score) => {
        map.set(`${score.activity_id}|${score.student_id}`, score.score);
      });
      return map;
    }, [communicationScores]);

    const communicationSavedBaseMap = useMemo(() => {
      const map = new Map<string, number>();
      communicationSavedBasesData.forEach((row) => {
        if (row.term_grade_base == null) return;
        const value = Number(row.term_grade_base);
        if (Number.isFinite(value)) {
          map.set(`${row.student_id}|${row.subject}`, value);
        }
      });
      return map;
    }, [communicationSavedBasesData]);

    // Use the exact subject columns currently shown in Summary of Grades.
    useEffect(() => {
      if (!active) {
        setSummarySubjects([]);
        return;
      }

      const fallbackSubject = klass?.subject?.trim();
      const storageKey = `sog-subjects-${active}`;

      try {
        const storedValue = localStorage.getItem(storageKey);
        const parsedSubjects: unknown = storedValue
          ? JSON.parse(storedValue)
          : [];

        const uniqueSubjects = Array.isArray(parsedSubjects)
          ? Array.from(
              new Map(
                parsedSubjects
                  .filter(
                    (subject): subject is string =>
                      typeof subject === "string" && Boolean(subject.trim()),
                  )
                  .map((subject) => [
                    subject.trim().toLocaleLowerCase(),
                    subject.trim(),
                  ]),
              ).values(),
            )
          : [];

        setSummarySubjects(
          uniqueSubjects.length > 0
            ? uniqueSubjects
            : fallbackSubject
              ? [fallbackSubject]
              : [],
        );
      } catch {
        localStorage.removeItem(storageKey);
        setSummarySubjects(fallbackSubject ? [fallbackSubject] : []);
      }
    }, [active, klass?.subject]);

    const isGrade11Shs =
      String(klass?.grade_level ?? "").trim().toLowerCase() === "grade 11";
    const isGrade12Shs =
      String(klass?.grade_level ?? "").trim().toLowerCase() === "grade 12";
    const isShsTemplatePreview = isGrade11Shs || isGrade12Shs;

    // Grades 7–10 use the JHS layout. Grades 11 and 12 each use their
    // dedicated SHS SF9 bond-paper layouts.
    const isJuniorHighSf9 = [
      "grade 7",
      "grade 8",
      "grade 9",
      "grade 10",
    ].includes(
      String(klass?.grade_level ?? "").trim().toLowerCase(),
    );

    const sf9LearningAreas = useMemo(
      () =>
        summarySubjects
          .filter(
            (subject) =>
              !(isGrade11Shs && isSf9AcademicElectiveSubject(subject)),
          )
          .map(learningAreaForSubject),
      [summarySubjects, isGrade11Shs],
    );

    const isMatatag = variant === "matatag";

    const previewPageWidth = isShsTemplatePreview
      ? SF9_GRADE11_A4_WIDTH
      : SF9_EXPORT_WIDTH;
    const previewPageHeight = isShsTemplatePreview
      ? SF9_GRADE11_A4_HEIGHT
      : SF9_EXPORT_HEIGHT;

    const HEADER_BG = isMatatag ? "#166534" : DEPED_BLUE;
    const FILE_PREFIX = isMatatag ? "SF9_MATATAG" : "SF9_ReportCard";
    const Icon = isMatatag ? FileText : GraduationCap;
    const PageTitle = isMatatag ? "School Form 9 (New MATATAG) — Learner's Progress Report" : "School Form 9 — Report Card";

    const communicationComponentGrade = (
      studentId: string,
      scopeKey: string,
      storageTerm: string,
    ): number | null => {
      const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
        const componentActivities = communicationActivities
          .filter(
            (activity) =>
              activity.term === storageTerm && activity.component === component,
          )
          .sort((first, second) => first.position - second.position)
          .slice(0, component === "QA" ? 3 : undefined);

        const componentWeight =
          communicationComponents.find(
            (item) =>
              item.term === storageTerm && item.component === component,
          )?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);

        let raw = 0;
        let hps = 0;
        let hasScore = false;
        let qaPercentageScore = 0;
        const qaPartWeights = [30, 30, 40] as const;

        componentActivities.forEach((activity, index) => {
          const value = communicationScoreMap.get(`${activity.id}|${studentId}`);
          if (typeof value !== "number") return;

          raw += value;
          hps += activity.hps;
          hasScore = true;

          if (component === "QA" && activity.hps > 0) {
            qaPercentageScore +=
              (value / activity.hps) *
              100 *
              ((qaPartWeights[index] ?? 0) / 100);
          }
        });

        if (!hasScore || hps <= 0) return null;

        const percentageScore =
          component === "QA" ? qaPercentageScore : (raw / hps) * 100;
        return percentageScore * (componentWeight / 100);
      });

      if (weightedParts.some((part) => part == null)) return null;

      const initialGrade =
        Math.round(
          (weightedParts as number[]).reduce((sum, part) => sum + part, 0) * 100,
        ) / 100;
      const computedBase = roundSf9TermGrade(initialGrade);
      if (computedBase == null) return null;

      return resolveSf9TermGradeBase(
        computedBase,
        communicationSavedBaseMap.get(
          `${studentId}|${communicationSubjectName}_${scopeKey}_ALL`,
        ),
      );
    };

    const communicationGradeForStudentTerm = (
      studentId: string,
      termValue: Sf9GradeTerm,
    ): number | null => {
      const gradeForPair = (firstScopeKey: string, secondScopeKey: string) =>
        averageSf9Pair(
          communicationComponentGrade(studentId, firstScopeKey, firstScopeKey),
          communicationComponentGrade(studentId, secondScopeKey, secondScopeKey),
        );

      if (termValue === "1") return gradeForPair("EC_T1", "MK_T1");
      if (termValue === "2") return gradeForPair("EC_T2", "MK_T2");
      if (termValue === "3") return gradeForPair("EC_T3", "MK_T3");

      const term1 = gradeForPair("EC_T1", "MK_T1");
      const term2 = gradeForPair("EC_T2", "MK_T2");
      const term3 = gradeForPair("EC_T3", "MK_T3");
      const termValues = [term1, term2, term3];

      return termValues.every(
        (value): value is number => typeof value === "number",
      )
        ? Math.round(
            termValues.reduce((sum, value) => sum + value, 0) / 3,
          )
        : null;
    };

    const gradeForAny = (
      studentId: string,
      subjectAliases: readonly string[],
      term: string,
    ) => {
      const normalizedAliases = new Set(
        subjectAliases.map((subject) => normalizeSubjectName(subject)),
      );

      if (
        isCommunicationClass &&
        (normalizedAliases.has(normalizeSubjectName(communicationSubjectName)) ||
          Array.from(normalizedAliases).some(isSf9CommunicationCompositeSubject))
      ) {
        const termValue: Sf9GradeTerm =
          term === "1" || term === "2" || term === "3" ? term : "final";
        return communicationGradeForStudentTerm(studentId, termValue);
      }

      const matchingGrades = grades.filter(
        (grade) =>
          grade.student_id === studentId &&
          normalizedAliases.has(normalizeSubjectName(grade.subject)),
      );

      if (term === "final") {
        const termAverages = ["1", "2", "3"].map((termNumber) => {
          const scores = matchingGrades
            .filter((grade) => grade.term === termNumber)
            .map((grade) => grade.score)
            .filter((score): score is number => typeof score === "number");

          return scores.length ? computeAverage(scores) : null;
        });

        return computeAverage(termAverages);
      }

      const scores = matchingGrades
        .filter((grade) => grade.term === term)
        .map((grade) => grade.score)
        .filter((score): score is number => typeof score === "number");

      return scores.length ? computeAverage(scores) : null;
    };

    const grade12GradeForSubjectTerm = (
      learnerId: string,
      subject: string,
      term: Sf9Grade12SubjectTerm,
    ): number | null => {
      const matchingStudentIds =
        grade12StudentIdsByLearnerId.get(learnerId) ?? [learnerId];
      const targetSubject = normalizeSf9SubjectKey(subject);

      const scores = grade12Grades
        .filter((grade) => {
          if (!matchingStudentIds.includes(grade.student_id)) return false;
          if (String(grade.term) !== term) return false;

          const gradeClass = grade12ClassById.get(
            String((grade as any).class_id || ""),
          );
          const rowSubject =
            String(grade.subject || "").trim() ||
            String(gradeClass?.subject || "").trim();

          return normalizeSf9SubjectKey(rowSubject) === targetSubject;
        })
        .map((grade) => grade.score)
        .filter((score): score is number => typeof score === "number");

      return scores.length ? computeAverage(scores) : null;
    };

    // Grade 12 SF9 keeps the official Units column in the template,
    // but the system no longer assigns or displays units for Grade 12.
    const buildGrade12RowsForLearner = (
      learnerId: string,
    ): Sf9Grade12PreviewRow[] =>
      SF9_GRADE12_SUBJECTS.map((subject) => ({
        ...subject,
        grade: grade12GradeForSubjectTerm(
          learnerId,
          subject.label,
          subject.term,
        ),
        units: null,
      }));

    // SF9 is generated for one learner only. Nothing is previewed or
    // exported until the adviser explicitly selects a learner.
    const toRender = selectedLearnerId
      ? students.filter((student) => student.id === selectedLearnerId).slice(0, 1)
      : [];

    const exportSf9ToExcel = async () => {
      if (toRender.length === 0) {
        toast.error("Select a learner before exporting to Excel.");
        return;
      }

      try {
        // Re-read the selected class at export time so Excel always gets the
        // latest Track (SHS only) value.
        let exportClass = klass;

        if (active) {
          const { data: latestClass, error: latestClassError } = await (supabase as any)
            .from("classes")
            .select("*")
            .eq("id", active)
            .single();

          if (latestClassError) throw latestClassError;
          exportClass = latestClass as SF9ClassRow;
        }

        const isGrade11Shs =
          normalizeSubjectName(exportClass?.grade_level) === "grade 11";
        const isGrade12Shs =
          normalizeSubjectName(exportClass?.grade_level) === "grade 12";

        const excelTemplateUrl = isGrade11Shs
          ? SF9_GRADE11_EXCEL_TEMPLATE_URL
          : isGrade12Shs
            ? SF9_GRADE12_EXCEL_TEMPLATE_URL
            : SF9_EXCEL_TEMPLATE_URL;

        const templateResponse = await fetch(excelTemplateUrl, { cache: "no-store" });
        if (!templateResponse.ok) {
          throw new Error(`Unable to load ${excelTemplateUrl}.`);
        }

        const templateBytes = await templateResponse.arrayBuffer();
        const regionText = profile?.region || "CARAGA Region";
        const divisionText = String(
          profile?.division || "AGUSAN DEL SUR",
        ).replace(/^SCHOOLS DIVISION OFFICE OF\s*/i, "");
        const districtText = profile?.district || "Prosperidad District";
        const municipalityText =
          profile?.city_municipality_province ||
          profile?.municipality ||
          "Prosperidad, Agusan del Sur";
        const schoolNameText =
          profile?.school_name ||
          "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL";

        for (const learner of toRender) {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(templateBytes.slice(0) as any);

          const worksheet = isGrade11Shs
            ? workbook.getWorksheet("SF9 - GRADE 11 SHS") ||
              workbook.getWorksheet("SF9 - GRADE 12 ACADEMIC") ||
              workbook.worksheets[0]
            : isGrade12Shs
              ? workbook.getWorksheet("SF9 - GRADE 11 ACADEMIC") ||
                workbook.getWorksheet("SF9 - GRADE 12 ACADEMIC") ||
                workbook.getWorksheet("SF9 - GRADE 12 SHS") ||
                workbook.worksheets[0]
              : workbook.getWorksheet("School Form 9") || workbook.worksheets[0];
          if (!worksheet) {
            throw new Error("The SF9 worksheet was not found in the template.");
          }

          if (isGrade11Shs && worksheet.name !== "SF9 - GRADE 11 SHS") {
            worksheet.name = "SF9 - GRADE 11 SHS";
          }

          workbook.creator = schoolNameText;
          workbook.lastModifiedBy = profile?.full_name || "School Forms System";
          workbook.created = new Date();
          workbook.modified = new Date();
          workbook.title = "School Form 9 (New MATATAG)";
          workbook.subject = "Learner's Performance Report";

          // Grade 11 SHS has a different SF9 spreadsheet layout.
          // This branch changes ONLY the Excel export. The on-screen SF9/PDF
          // and Word output keep their existing behavior.
          if (isGrade11Shs) {
            // IMPORTANT: Do not save the Grade 11 template with ExcelJS.
            // ExcelJS can rewrite/drop unsupported parts from a complex template,
            // which is why Microsoft Excel may show "We found a problem with some
            // content" and the exported workbook may no longer look exactly like
            // the original uploaded Excel file.
            //
            // Instead, keep the original .xlsx ZIP package intact and edit only
            // the worksheet XML values. This preserves the original workbook's
            // formatting, logos/images, merged cells, borders, fonts, row heights,
            // column widths, print setup, and all other workbook parts.
            const grade11Zip = await JSZip.loadAsync(templateBytes.slice(0));
            const { worksheetPath, worksheetDocument } =
              await resolveXlsxWorksheet(grade11Zip, [
                "SF9 - GRADE 11 SHS",
                "SF9 - GRADE 12 ACADEMIC",
              ]);

            const stylesFile = grade11Zip.file("xl/styles.xml");
            if (!stylesFile) {
              throw new Error(
                "The Grade 11 Excel template is missing xl/styles.xml.",
              );
            }

            const stylesDocument = new DOMParser().parseFromString(
              await stylesFile.async("string"),
              "application/xml",
            );

            if (stylesDocument.querySelector("parsererror")) {
              throw new Error(
                "The Grade 11 Excel template styles could not be parsed.",
              );
            }

            // Header / school information.
            setXlsxCellText(worksheetDocument, "B7", regionText);
            setXlsxCellText(
              worksheetDocument,
              "B8",
              `SCHOOLS DIVISION OFFICE OF ${divisionText}`,
            );
            setXlsxCellText(worksheetDocument, "B9", districtText);
            setXlsxCellText(worksheetDocument, "B10", municipalityText);
            setXlsxCellText(worksheetDocument, "B12", schoolNameText);
            setXlsxCellText(
              worksheetDocument,
              "B15",
              `School Year ${exportClass?.school_year || ""}`.trim(),
            );

            // Learner information.
            setXlsxCellText(
              worksheetDocument,
              "C17",
              [learner.last_name, learner.first_name, learner.middle_name]
                .filter(Boolean)
                .join(", ")
                .toUpperCase(),
            );
            setXlsxCellValue(
              worksheetDocument,
              "I17",
              calculateLearnerAge(learner.birthdate),
            );
            setXlsxCellText(worksheetDocument, "K17", learner.sex || "");
            setXlsxCellText(worksheetDocument, "C18", learner.lrn || "");
            setXlsxCellText(
              worksheetDocument,
              "I18",
              exportClass?.grade_level || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "K18",
              exportClass?.section || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "E19",
              exportClass?.track_shs || "",
            );

            // Signature lines.
            setXlsxCellText(
              worksheetDocument,
              "B27",
              profile?.principal || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "H27",
              exportClass?.teacher_name || profile?.full_name || "",
            );

            const grade11Rows = sf9LearningAreas.map((learningArea) => {
              const term1 = gradeForAny(learner.id, learningArea.aliases, "1");
              const term2 = gradeForAny(learner.id, learningArea.aliases, "2");
              const term3 = gradeForAny(learner.id, learningArea.aliases, "3");
              const final = gradeForAny(
                learner.id,
                learningArea.aliases,
                "final",
              );

              const normalizedAliases = new Set(
                learningArea.aliases.map(normalizeSubjectName),
              );
              const matchingClass = classes.find(
                (candidate) =>
                  normalizeSubjectName(candidate.grade_level) === "grade 11" &&
                  normalizeSubjectName(candidate.section) ===
                    normalizeSubjectName(exportClass?.section) &&
                  String(candidate.school_year || "") ===
                    String(exportClass?.school_year || "") &&
                  normalizedAliases.has(normalizeSubjectName(candidate.subject)),
              ) as SF9ClassRow | undefined;

              const numericUnits = Number(matchingClass?.units);
              const units =
                matchingClass?.units != null && Number.isFinite(numericUnits)
                  ? numericUnits
                  : null;

              return {
                ...learningArea,
                term1,
                term2,
                term3,
                final,
                units,
              };
            });

            // Rows 33:39 are the real Grade 11 subject slots in the original
            // workbook. Keep only actual subjects visible. Rows 40:42 are the
            // template's Academic Elective rows, and row 43 is General Average.
            const firstShsSubjectRow = 33;
            const lastShsSubjectRow = 39;
            const maxShsSubjectRows =
              lastShsSubjectRow - firstShsSubjectRow + 1;
            const displayedRows = grade11Rows
              .filter((row) => Boolean(row.label?.trim()))
              .slice(0, maxShsSubjectRows);

            for (let index = 0; index < maxShsSubjectRows; index += 1) {
              const excelRow = firstShsSubjectRow + index;
              const row = displayedRows[index];

              if (!row) {
                setXlsxRowHidden(worksheetDocument, excelRow, true);
                ["B", "F", "G", "H", "I", "J", "K"].forEach((column) => {
                  setXlsxCellText(
                    worksheetDocument,
                    `${column}${excelRow}`,
                    "",
                  );
                });
                continue;
              }

              setXlsxRowHidden(worksheetDocument, excelRow, false);
              setXlsxCellText(
                worksheetDocument,
                `B${excelRow}`,
                row.label,
              );
              setXlsxCellValue(
                worksheetDocument,
                `F${excelRow}`,
                row.term1,
              );
              setXlsxCellValue(
                worksheetDocument,
                `G${excelRow}`,
                row.term2,
              );
              setXlsxCellValue(
                worksheetDocument,
                `H${excelRow}`,
                row.term3,
              );
              // Grade 12 does not use Units. Keep the official Units column
              // but always leave its cell blank in the exported SF9.
              setXlsxCellText(
                worksheetDocument,
                `I${excelRow}`,
                "",
              );
              setXlsxCellValue(
                worksheetDocument,
                `J${excelRow}`,
                row.final,
              );
              setXlsxCellText(
                worksheetDocument,
                `K${excelRow}`,
                row.final == null
                  ? ""
                  : row.final >= 75
                    ? "Passed"
                    : "Failed",
              );
            }

            const academicElectiveExportRows = [
              {
                row: 40,
                label: "Academic Elective 1",
                termColumn: "F",
                grade: academicElectiveGradeForStudentTerm(learner.id, "1"),
                units: academicElectiveUnits,
              },
              {
                row: 41,
                label: "Academic Elective 2",
                termColumn: "G",
                grade: academicElectiveGradeForStudentTerm(learner.id, "2"),
                units: academicElectiveUnits,
              },
              {
                row: 42,
                label: "Academic Elective 3",
                termColumn: "H",
                grade: academicElectiveGradeForStudentTerm(learner.id, "3"),
                units: academicElectiveUnits,
              },
            ] as const;

            academicElectiveExportRows.forEach(
              ({ row, label, termColumn, grade, units }) => {
                setXlsxRowHidden(worksheetDocument, row, false);
                setXlsxCellText(worksheetDocument, `B${row}`, label);

                // Clear every term/final field first. Only the one white cell
                // for this Academic Elective receives a grade. The gray cells
                // remain intentionally unused.
                ["F", "G", "H", "I", "J", "K"].forEach((column) => {
                  setXlsxCellText(worksheetDocument, `${column}${row}`, "");
                });

                setXlsxCellValue(
                  worksheetDocument,
                  `${termColumn}${row}`,
                  grade,
                );

                // Units belong only to an Academic Elective that actually
                // has a grade for its assigned term. Empty elective rows must
                // remain empty and must never inherit the main subject's unit.
                setXlsxCellValue(
                  worksheetDocument,
                  `I${row}`,
                  grade == null ? null : units,
                );

                // Export-to-Excel formatting:
                // - all Academic Elective Term 1/2/3 values are centered
                // - all Unit values are centered
                // - whole values display as 3 / 80, never 3.00 / 80.00
                ["F", "G", "H"].forEach((column) => {
                  applyCenteredXlsxStyle(
                    worksheetDocument,
                    stylesDocument,
                    `${column}${row}`,
                    { integer: true },
                  );
                });

                applyCenteredXlsxStyle(
                  worksheetDocument,
                  stylesDocument,
                  `I${row}`,
                  { integer: true },
                );
              },
            );

            setXlsxRowHidden(worksheetDocument, 43, false);

            const generalAverage = computeAverage(
              displayedRows
                .filter((row) => !row.indented)
                .map((row) => row.final),
            );

            // Row 43 is the General Average row in SF9 Grade 11.
            // Clear any leftover Academic Elective text/units from the
            // original workbook template before writing the average.
            ["B", "F", "G", "H", "I", "J", "K"].forEach((column) => {
              setXlsxCellText(
                worksheetDocument,
                `${column}43`,
                "",
              );
            });

            setXlsxCellText(
              worksheetDocument,
              "B43",
              "General Average",
            );
            setXlsxCellValue(
              worksheetDocument,
              "J43",
              generalAverage,
            );

            // Center the General Average value in the Final Grade cell.
            // Keep it as an integer-style Excel value such as 82, not 82.00.
            applyCenteredXlsxStyle(
              worksheetDocument,
              stylesDocument,
              "J43",
              { integer: true },
            );

            setXlsxCellText(
              worksheetDocument,
              "K43",
              getGeneralAverageRemark(generalAverage),
            );

            const fileName = `SF9_GRADE11_SHS_${
              exportClass?.section || "class"
            }_${learner.last_name || learner.id}`;
            const buffer = await writePreservedXlsxSheet(
              grade11Zip,
              worksheetPath,
              worksheetDocument,
              stylesDocument,
            );
            downloadExcelBuffer(buffer, fileName);

            if (toRender.length > 1) {
              await new Promise((resolve) =>
                window.setTimeout(resolve, 250),
              );
            }

            continue;
          }

          if (isGrade12Shs) {
            // Preserve the uploaded Grade-12-SHS.xlsx exactly. Only worksheet
            // values are changed; styles, gray cells, merged cells, images,
            // widths, heights, formulas on untouched sheets, and print setup
            // stay from the original workbook.
            const grade12Zip = await JSZip.loadAsync(templateBytes.slice(0));
            const { worksheetPath, worksheetDocument } =
              await resolveXlsxWorksheet(grade12Zip, [
                "SF9 - GRADE 11 ACADEMIC",
                "SF9 - GRADE 12 ACADEMIC",
                "SF9 - GRADE 12 SHS",
              ]);

            // School information — same cells as the uploaded Grade 12 sheet.
            setXlsxCellText(worksheetDocument, "B7", regionText);
            setXlsxCellText(
              worksheetDocument,
              "B8",
              `SCHOOLS DIVISION OFFICE OF ${divisionText}`,
            );
            setXlsxCellText(worksheetDocument, "B9", districtText);
            setXlsxCellText(worksheetDocument, "B10", municipalityText);
            setXlsxCellText(worksheetDocument, "B12", schoolNameText);
            setXlsxCellText(
              worksheetDocument,
              "B15",
              `School Year ${exportClass?.school_year || ""}`.trim(),
            );

            // Learner information.
            setXlsxCellText(
              worksheetDocument,
              "C17",
              [learner.last_name, learner.first_name, learner.middle_name]
                .filter(Boolean)
                .join(", ")
                .toUpperCase(),
            );
            setXlsxCellValue(
              worksheetDocument,
              "I17",
              calculateLearnerAge(learner.birthdate),
            );
            setXlsxCellText(worksheetDocument, "K17", learner.sex || "");
            setXlsxCellText(worksheetDocument, "C18", learner.lrn || "");
            setXlsxCellText(
              worksheetDocument,
              "I18",
              exportClass?.grade_level || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "K18",
              exportClass?.section || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "E19",
              exportClass?.track_shs || "",
            );

            setXlsxCellText(
              worksheetDocument,
              "B27",
              profile?.principal || "",
            );
            setXlsxCellText(
              worksheetDocument,
              "H27",
              exportClass?.teacher_name || profile?.full_name || "",
            );

            const grade12Rows = buildGrade12RowsForLearner(learner.id);

            grade12Rows.forEach((row) => {
              const excelRow = row.excelRow;

              // Keep the exact labels from the supplied template.
              setXlsxCellText(worksheetDocument, `B${excelRow}`, row.label);

              // Clear all three term cells first. The workbook's original gray
              // and white fills remain untouched because cell styles are not
              // replaced by setXlsxCellValue/setXlsxCellText.
              ["F", "G", "H"].forEach((column) => {
                setXlsxCellText(
                  worksheetDocument,
                  `${column}${excelRow}`,
                  "",
                );
              });

              const termColumn =
                row.term === "1" ? "F" : row.term === "2" ? "G" : "H";

              setXlsxCellValue(
                worksheetDocument,
                `${termColumn}${excelRow}`,
                row.grade,
              );
              setXlsxCellValue(
                worksheetDocument,
                `I${excelRow}`,
                row.units,
              );
              setXlsxCellValue(
                worksheetDocument,
                `J${excelRow}`,
                row.grade,
              );
              setXlsxCellText(
                worksheetDocument,
                `K${excelRow}`,
                row.grade == null
                  ? ""
                  : row.grade >= 75
                    ? "Passed"
                    : "Failed",
              );

            });

            const grade12GeneralAverage = computeAverage(
              grade12Rows.map((row) => row.grade),
            );

            // The uploaded template contains a broken original formula in J53.
            // Replace it with the actual system-computed Grade 12 average.
            setXlsxCellValue(
              worksheetDocument,
              "J53",
              grade12GeneralAverage,
            );
            setXlsxCellText(
              worksheetDocument,
              "K53",
              getGeneralAverageRemark(grade12GeneralAverage),
            );

            const fileName = `SF9_GRADE12_SHS_${
              exportClass?.section || "class"
            }_${learner.last_name || learner.id}`;

            await removeStaleXlsxCalculationChain(grade12Zip);

            const buffer = await writePreservedXlsxSheet(
              grade12Zip,
              worksheetPath,
              worksheetDocument,
            );
            downloadExcelBuffer(buffer, fileName);

            if (toRender.length > 1) {
              await new Promise((resolve) =>
                window.setTimeout(resolve, 250),
              );
            }

            continue;
          }

          // Do not rebuild, resize, restyle, or replace the worksheet.
          // The original sf9-class-adviser.xlsx formatting, merged cells,
          // borders, fonts, column widths, row heights, print settings, and
          // embedded objects are preserved. We only replace the data cells.

          // School information cells from the provided SF9 template.
          worksheet.getCell("B4").value = regionText;
          worksheet.getCell("B5").value = `SCHOOLS DIVISION OFFICE OF ${divisionText}`;
          worksheet.getCell("B6").value = districtText;
          worksheet.getCell("B7").value = municipalityText;
          worksheet.getCell("B8").value = schoolNameText;

          // Learner information cells from the provided SF9 template.
          worksheet.getCell("D11").value = exportClass?.school_year || "";
          worksheet.getCell("C13").value = [
            learner.last_name,
            learner.first_name,
            learner.middle_name,
          ]
            .filter(Boolean)
            .join(", ")
            .toUpperCase();
          worksheet.getCell("G13").value = calculateLearnerAge(learner.birthdate);
          worksheet.getCell("C14").value = learner.lrn || "";
          worksheet.getCell("G14").value = learner.sex || "";
          worksheet.getCell("C15").value = exportClass?.track_shs || "";
          worksheet.getCell("G15").value = exportClass?.grade_level || "";
          worksheet.getCell("G16").value = exportClass?.section || "";

          const learnerRows = sf9LearningAreas.map((learningArea) => {
            const term1 = gradeForAny(learner.id, learningArea.aliases, "1");
            const term2 = gradeForAny(learner.id, learningArea.aliases, "2");
            const term3 = gradeForAny(learner.id, learningArea.aliases, "3");
            const final = gradeForAny(learner.id, learningArea.aliases, "final");

            return {
              ...learningArea,
              term1,
              term2,
              term3,
              final,
            };
          });

          // Show ONLY the subjects that actually exist in Summary of Grades.
          // No extra blank Learning Area row is kept.
          //
          // Example with 2 subjects:
          //   row 22 = Filipino
          //   row 23 = English
          //   row 24 = GENERAL AVERAGE
          //   row 25 = blank spacing
          //   row 26-31 = Performance Descriptors
          //
          // Only the LEFT side (B:G) is compacted. The right-side Attendance,
          // Comments, Transfer and signature areas stay exactly in place.
          const firstLearningAreaRow = 22;
          const lastLearningAreaRow = 33;
          const maxLearningAreaRows =
            lastLearningAreaRow - firstLearningAreaRow + 1;
          const usedLearningAreaCount = Math.min(
            learnerRows.length,
            maxLearningAreaRows,
          );

          // Fill the real subjects first.
          for (let index = 0; index < usedLearningAreaCount; index += 1) {
            const excelRow = firstLearningAreaRow + index;
            const row = learnerRows[index];

            worksheet.getCell(`B${excelRow}`).value = row?.label ?? "";
            worksheet.getCell(`C${excelRow}`).value = row?.term1 ?? "";
            worksheet.getCell(`D${excelRow}`).value = row?.term2 ?? "";
            worksheet.getCell(`E${excelRow}`).value = row?.term3 ?? "";
            worksheet.getCell(`F${excelRow}`).value = row?.final ?? "";
            worksheet.getCell(`G${excelRow}`).value =
              row?.final == null ? "" : row.final >= 75 ? "Passed" : "Failed";
          }

          // Remove every unused subject row visually from B:G.
          const firstUnusedLearningAreaRow =
            firstLearningAreaRow + usedLearningAreaCount;

          if (firstUnusedLearningAreaRow <= lastLearningAreaRow) {
            clearExcelBlock(
              worksheet,
              firstUnusedLearningAreaRow,
              lastLearningAreaRow,
              2,
              7,
            );
          }

          const generalAverage = computeAverage(
            learnerRows
              .filter((row) => !row.indented)
              .map((row) => row.final),
          );

          const originalGeneralAverageRow = 34;
          const originalDescriptorStartRow = 36;
          const originalDescriptorEndRow = 41;

          // GENERAL AVERAGE goes immediately after the final real subject.
          const targetGeneralAverageRow =
            firstLearningAreaRow + usedLearningAreaCount;

          // Keep one normal blank row between GENERAL AVERAGE and descriptors,
          // matching the requested example image.
          const targetDescriptorStartRow = targetGeneralAverageRow + 2;
          const targetDescriptorEndRow =
            targetDescriptorStartRow +
            (originalDescriptorEndRow - originalDescriptorStartRow);

          if (targetGeneralAverageRow !== originalGeneralAverageRow) {
            // Move the template-formatted GENERAL AVERAGE row and descriptor
            // block upward without deleting any full Excel worksheet rows.
            copyExcelBlock(
              worksheet,
              originalGeneralAverageRow,
              originalGeneralAverageRow,
              targetGeneralAverageRow,
              2,
              7,
            );

            copyExcelBlock(
              worksheet,
              originalDescriptorStartRow,
              originalDescriptorEndRow,
              targetDescriptorStartRow,
              2,
              7,
            );

            // The row between General Average and the descriptor table must be
            // completely clean.
            clearExcelBlock(
              worksheet,
              targetGeneralAverageRow + 1,
              targetGeneralAverageRow + 1,
              2,
              7,
            );

            // Remove the old left-side template copies below the compact block.
            const clearOldFrom = targetDescriptorEndRow + 1;
            if (clearOldFrom <= originalDescriptorEndRow) {
              clearExcelBlock(
                worksheet,
                clearOldFrom,
                originalDescriptorEndRow,
                2,
                7,
              );
            }
          }

          worksheet.getCell(`F${targetGeneralAverageRow}`).value =
            generalAverage ?? "";
          worksheet.getCell(`G${targetGeneralAverageRow}`).value =
            getGeneralAverageRemark(generalAverage);

          const fileName = `${FILE_PREFIX}_${klass?.section || "class"}_${
            learner.last_name || learner.id
          }`;
          const buffer = await workbook.xlsx.writeBuffer();
          downloadExcelBuffer(buffer as ArrayBuffer, fileName);

          if (toRender.length > 1) {
            await new Promise((resolve) => window.setTimeout(resolve, 250));
          }
        }

        toast.success(
          `Exported ${toRender.length} SF9 Excel file${
            toRender.length === 1 ? "" : "s"
          } using the spreadsheet template.`,
        );
      } catch (error) {
        console.error("Unable to export SF9 to Excel.", error);
        toast.error(
          error instanceof Error
            ? error.message
            : "Unable to export SF9 to Excel.",
        );
      }
    };

    const copySf9ToWord = async () => {
      const documentElement = document.getElementById("sf9-doc");
      if (!documentElement || toRender.length === 0) {
        toast.error("Select a learner before creating the Word document.");
        return;
      }

      try {
        const pageElements = Array.from(documentElement.children).filter(
          (child): child is HTMLElement => child instanceof HTMLElement,
        );

        if (pageElements.length === 0) {
          throw new Error("No SF9 bondpaper is available to copy.");
        }

        // The Word export follows the supplied SF9 spreadsheet template:
        // long bond paper (8.5 x 13 in), landscape. This is intentionally
        // independent of the preview paper dropdown so the downloaded Word
        // file keeps the same page shape as sf9-class-adviser.xlsx.
        const paperInches = { width: 8.5, height: 13 };

        const landscapeWidthInches = paperInches.height;
        const landscapeHeightInches = paperInches.width;
        const marginInches = 0.2;

        // A Word paragraph always keeps a small paragraph mark below an
        // image. If the image uses the complete printable height, Word pushes
        // it to page 2 and page 1 appears blank. Keep a small safety area.
        const wordSafetyWidthPx = 12;
        const wordSafetyHeightPx = 24;
        const targetPageWidth = Math.max(
          1,
          Math.round((landscapeWidthInches - marginInches * 2) * 96) -
            wordSafetyWidthPx,
        );
        const targetPageHeight = Math.max(
          1,
          Math.round((landscapeHeightInches - marginInches * 2) * 96) -
            wordSafetyHeightPx,
        );

        const renderedPages = [];
        for (const pageElement of pageElements) {
          renderedPages.push(
            await renderSf9PageToPng(
              pageElement,
              targetPageWidth,
              targetPageHeight,
            ),
          );
        }

        const wordChildren: Paragraph[] = [];

        renderedPages.forEach((page, index) => {
          wordChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 0,
                after: 0,
              },
              indent: {
                left: 0,
                right: 0,
                firstLine: 0,
              },
              children: [
                new ImageRun({
                  data: page.data,
                  type: "png",
                  transformation: {
                    width: page.width,
                    height: page.height,
                  },
                }),
              ],
            }),
          );

          if (index < renderedPages.length - 1) {
            wordChildren.push(
              new Paragraph({ children: [new PageBreak()] }),
            );
          }
        });

        const wordDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: { font: "Times New Roman", size: 20 },
                paragraph: { spacing: { before: 0, after: 0 } },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    // The docx library applies the landscape rotation itself.
                    // Keep the original portrait paper dimensions here;
                    // passing the already-swapped dimensions rotates twice and
                    // causes the right half of the SF9 image to be clipped.
                    width: Math.round(paperInches.width * 1440),
                    height: Math.round(paperInches.height * 1440),
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: Math.round(marginInches * 1440),
                    bottom: Math.round(marginInches * 1440),
                    left: Math.round(marginInches * 1440),
                    right: Math.round(marginInches * 1440),
                  },
                },
              },
              children: wordChildren,
            },
          ],
        });

        const blob = await Packer.toBlob(wordDocument);
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${FILE_PREFIX}_${klass?.section || "class"}.docx`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);

        toast.success("SF9 Word document downloaded in the same long-bond landscape format as the spreadsheet template.");
      } catch (error) {
        console.error("Unable to create the SF9 Word document.", error);
        toast.error(
          error instanceof Error
            ? error.message
            : "Unable to create the SF9 Word document.",
        );
      }
    };

    return (
      <div className="space-y-4 pb-8">
        <div className="flex flex-wrap items-center gap-3">
          <Link
            to="/school-forms"
            className="inline-flex items-center gap-1.5 rounded-lg border-2 px-3 py-1.5 text-sm font-medium text-white shadow-sm hover:opacity-90"
            style={{ backgroundColor: HEADER_BG, borderColor: HEADER_BG }}
          >
            <ArrowLeft className="size-4" /> Back to Forms
          </Link>
          <div className="flex items-center gap-2 text-lg font-semibold">
            <Icon className="size-5" style={{ color: HEADER_BG }} />
            {PageTitle}
          </div>
        </div>

        <div className="rounded-2xl border-2 p-4 shadow-sm" style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}>
          <div className="mb-2 text-sm font-semibold text-amber-900">Form Configuration</div>
          <div className="mx-auto mt-1 flex max-w-2xl flex-col items-center">
            <div className="flex w-full items-center justify-center gap-3 sm:gap-5">
              <div className="hidden h-px w-20 bg-amber-400 sm:block" />
              <div
                className="grid size-9 shrink-0 place-items-center rounded-full"
                style={{ backgroundColor: "#FEF0B6" }}
              >
                <GraduationCap className="size-5 text-amber-950" />
              </div>
              <div className="hidden h-px w-20 bg-amber-400 sm:block" />
            </div>

            <div className="mt-1 text-center text-xs font-bold uppercase tracking-[0.32em] text-amber-950">
              Class
            </div>

            <div className="mt-1 flex w-full max-w-[300px] items-center gap-1.5 rounded-lg border border-amber-300 bg-white/75 px-2 py-2 shadow-sm">
              <div className="grid size-5 shrink-0 place-items-center rounded-md bg-amber-50 text-amber-700">
                <Users className="size-3" />
              </div>

              <div className="min-w-0 flex-1 text-center">
                <div className="truncate text-[11px] font-bold text-amber-950">
                  {klass
                    ? `${klass.grade_level || "—"} · ${klass.subject || "—"} · ${klass.section || "—"}`
                    : "No class selected"}
                </div>
              </div>
            </div>

            <p className="mt-0.5 text-center text-[9px] text-amber-900/60">
              Selected class for this form
            </p>
          </div>

          {/* SF9 learner selection — one learner only for Grades 7–12 */}
          <div
            className="mt-4 rounded-xl border-2 bg-white p-3"
            style={{ borderColor: HEADER_BG }}
          >
            <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-amber-900">
              Learner
            </label>

            {students.length > 0 ? (
              <Select
                value={selectedLearnerId || undefined}
                onValueChange={setSelectedLearnerId}
              >
                <SelectTrigger className="bg-white">
                  <SelectValue placeholder="Select one learner for SF9" />
                </SelectTrigger>

                <SelectContent>
                  {students.map((student) => {
                    const learnerName = [
                      student.last_name,
                      student.first_name,
                      student.middle_name,
                    ]
                      .filter(Boolean)
                      .join(", ");

                    return (
                      <SelectItem key={student.id} value={student.id}>
                        {learnerName || student.lrn || "Unnamed learner"}
                        {student.lrn ? ` · LRN ${student.lrn}` : ""}
                      </SelectItem>
                    );
                  })}
                </SelectContent>
              </Select>
            ) : (
              <div className="text-xs text-muted-foreground">
                No students yet.
              </div>
            )}

            {selectedLearnerId && (
              <div className="mt-2 text-xs font-medium" style={{ color: HEADER_BG }}>
                1 learner selected
              </div>
            )}

            <div className="mt-1 text-xs text-muted-foreground">
              SF9 is an individual learner report. Only one learner can be
              selected at a time for Grades 7, 8, 9, 10, 11, and 12.
            </div>
          </div>
        </div>

        <PdfPreviewShell
          fileName={`${FILE_PREFIX}_${klass?.section || "class"}.pdf`}
          docBaseName={`${FILE_PREFIX}_${klass?.section || "class"}`}
          printTargetId="sf9-doc"
          length={length}
          onLengthChange={setLength}
          paper={isShsTemplatePreview ? "a4" : paper}
          onPaperChange={isShsTemplatePreview ? undefined : setPaper}
          printMargin={
            isShsTemplatePreview
              ? "0.2507in 0.3617in 0.144in 0.4725in"
              : "0.4in"
          }
          fitPrintToPage={isShsTemplatePreview}
          onCopyToWord={copySf9ToWord}
          hideCopyToWord={
            isGrade11Shs ||
            String(klass?.grade_level ?? "")
              .trim()
              .toLowerCase() === "grade 12"
          }
          onExportToExcel={exportSf9ToExcel}
          pageLabel={toRender.length === 1 ? "1 learner" : "No learner selected"}
        >
          <div
            id="sf9-doc"
            className="text-black"
            style={{
              fontFamily: '"Bookman Old Style", "Times New Roman", serif',
            }}
          >
            {toRender.map((s, idx) => {
              const learnerRows = sf9LearningAreas.map((learningArea) => {
                const term1 = gradeForAny(s.id, learningArea.aliases, "1");
                const term2 = gradeForAny(s.id, learningArea.aliases, "2");
                const term3 = gradeForAny(s.id, learningArea.aliases, "3");
                const final = gradeForAny(s.id, learningArea.aliases, "final");

                return {
                  ...learningArea,
                  term1,
                  term2,
                  term3,
                  final,
                };
              });

              const academicElectiveGradesForLearner: Sf9AcademicElectiveGrades = {
                term1: academicElectiveGradeForStudentTerm(s.id, "1"),
                term2: academicElectiveGradeForStudentTerm(s.id, "2"),
                term3: academicElectiveGradeForStudentTerm(s.id, "3"),
              };

              const generalAverage = computeAverage(
                learnerRows
                  .filter((row) => !row.indented)
                  .map((row) => row.final),
              );

              const regionText = profile?.region || "CARAGA Region";
              const divisionText = String(
                profile?.division || "AGUSAN DEL SUR",
              ).replace(/^SCHOOLS DIVISION OFFICE OF\s*/i, "");
              const districtText = profile?.district || "Prosperidad District";
              const municipalityText =
                profile?.city_municipality_province ||
                profile?.municipality ||
                "Prosperidad, Agusan del Sur";
              const schoolNameText =
                profile?.school_name ||
                "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL";

              return (
                <div
                  key={s.id}
                  data-sf9-page={s.id}
                  className="mx-auto box-border bg-white text-[9.5px] text-black"
                  style={{
                    width: `${previewPageWidth}px`,
                    minWidth: `${previewPageWidth}px`,
                    maxWidth: `${previewPageWidth}px`,
                    minHeight: `${previewPageHeight}px`,
                    height: `${previewPageHeight}px`,
                    maxHeight: `${previewPageHeight}px`,
                    padding: isShsTemplatePreview ? "12px" : "14px",
                    overflow: "hidden",
                    pageBreakAfter:
                      idx < toRender.length - 1 ? "always" : "auto",
                    fontFamily:
                      '"Bookman Old Style", "Times New Roman", serif',
                  }}
                >
                  {isGrade11Shs ? (
                    <Grade11ShsSf9Preview
                      learner={s}
                      klass={klass}
                      learnerRows={learnerRows}
                      academicElectiveGrades={academicElectiveGradesForLearner}
                      academicElectiveUnits={academicElectiveUnits}
                      generalAverage={generalAverage}
                      profile={profile}
                      regionText={regionText}
                      divisionText={divisionText}
                      districtText={districtText}
                      municipalityText={municipalityText}
                      schoolNameText={schoolNameText}
                    />
                  ) : isGrade12Shs ? (
                    <Grade12ShsSf9Preview
                      learner={s}
                      klass={klass}
                      subjectRows={buildGrade12RowsForLearner(s.id)}
                      generalAverage={computeAverage(
                        buildGrade12RowsForLearner(s.id).map(
                          (row) => row.grade,
                        ),
                      )}
                      profile={profile}
                      regionText={regionText}
                      divisionText={divisionText}
                      districtText={districtText}
                      municipalityText={municipalityText}
                      schoolNameText={schoolNameText}
                    />
                  ) : (
                    <div
                      className="h-full"
                      style={{ padding: "12px 14px" }}
                    >
                      <div
                        className="grid h-full min-h-0 gap-3 overflow-hidden"
                        style={{ gridTemplateColumns: "48.5fr 51.5fr" }}
                      >
                      {/* LEFT SIDE — spreadsheet layout */}
                      <section className="flex min-w-0 flex-col">
                        <div className="grid grid-cols-[82px_1fr_82px] items-start gap-2">
                          <img
                            src={depedLogo}
                            alt="DepEd logo"
                            className="h-[76px] w-[76px] object-contain"
                          />

                          <div className="text-center text-[9.5px] leading-[1.18]">
                            <div>Republic of the Philippines</div>
                            <div>Department of Education</div>
                            <div>{regionText}</div>
                            <div className="font-bold uppercase">
                              SCHOOLS DIVISION OFFICE OF {divisionText}
                            </div>
                            <div>{districtText}</div>
                            <div>{municipalityText}</div>
                            <div className="mt-1 font-semibold uppercase">
                              {schoolNameText}
                            </div>
                          </div>

                          <img
                            src={schoolLogo}
                            alt="School logo"
                            className="h-[80px] w-[80px] object-contain"
                          />
                        </div>

                        <div className="mt-3 text-center text-[16px] font-bold">
                          LEARNER&apos;S PERFORMANCE REPORT
                        </div>

                        {isJuniorHighSf9 ? (
                          <>
                            {/* Grades 7–10 bondpaper-only learner information.
                                Match the Excel reference exactly:
                                left side = Name / LRN / Track,
                                right side = Age / Sex / Grade / Section,
                                with Age, Sex, Grade, and Section stacked vertically. */}
                            <div className="mt-1 grid grid-cols-[92px_260px] items-end gap-x-2 text-[11px]">
                              <div className="font-bold">School Year</div>
                              <div className="flex min-h-[19px] items-end justify-center border-b border-black px-2 pb-[1px] text-center leading-none">
                                {klass?.school_year || ""}
                              </div>
                            </div>

                            {/* Grades 7–10 only:
                                move the Age / Sex / Grade / Section block
                                farther to the RIGHT to match the Excel SF9. */}
                            <div className="mt-3 grid grid-cols-[48px_272px_48px_92px] items-end gap-x-[7px] gap-y-[3px] text-[11px]">
                              <div className="font-bold">Name:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] font-semibold uppercase leading-none">
                                {[s.last_name, s.first_name, s.middle_name]
                                  .filter(Boolean)
                                  .join(", ")}
                              </div>
                              <div className="font-bold">Age:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                {calculateLearnerAge(s.birthdate)}
                              </div>

                              <div className="font-bold">LRN:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                {s.lrn || ""}
                              </div>
                              <div className="font-bold">Sex:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] capitalize leading-none">
                                {s.sex || ""}
                              </div>

                              {/* Track needs a wider label column than Name/LRN.
                                  Keep a clear fill-out line AFTER the full
                                  "Track (SHS only):" label. */}
                              <div className="col-span-2 grid grid-cols-[102px_218px] items-end gap-x-[6px]">
                                <div className="whitespace-nowrap font-bold">
                                  Track (SHS only):
                                </div>
                                <div className="min-h-[19px] border-b border-black px-1 pb-[1px]">
                                  {klass?.track_shs || ""}
                                </div>
                              </div>
                              <div className="font-bold">Grade:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                {klass?.grade_level || ""}
                              </div>

                              <div />
                              <div />
                              <div className="font-bold">Section:</div>
                              <div className="flex min-h-[19px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                {klass?.section || ""}
                              </div>
                            </div>
                          </>
                        ) : (
                          <>
                            <div className="mt-1 grid grid-cols-[96px_1fr] items-end gap-2 text-[10.5px]">
                              <div className="font-bold">School Year</div>
                              <div className="min-h-[18px] border-b border-black px-2 text-center">
                                {klass?.school_year || ""}
                              </div>
                            </div>

                            <div className="mt-3 space-y-1 text-[10.5px]">
                              <div className="grid grid-cols-[92px_280px_1fr] items-end gap-x-2">
                                <div className="font-bold">Name:</div>
                                <div className="flex min-h-[18px] items-end border-b border-black px-1 pb-[1px] font-semibold uppercase leading-none">
                                  {[s.last_name, s.first_name, s.middle_name]
                                    .filter(Boolean)
                                    .join(", ")}
                                </div>
                                <div className="grid grid-cols-[52px_74px_64px_74px] items-end gap-x-2">
                                  <div className="font-bold">Age:</div>
                                  <div className="flex min-h-[18px] items-end justify-center border-b border-black px-1 pb-[1px] text-center leading-none">
                                    {calculateLearnerAge(s.birthdate)}
                                  </div>
                                  <div className="font-bold">Sex:</div>
                                  <div className="flex min-h-[18px] items-end justify-center border-b border-black px-1 pb-[1px] text-center capitalize leading-none">
                                    {s.sex || ""}
                                  </div>
                                </div>
                              </div>

                              <div className="grid grid-cols-[92px_280px_1fr] items-end gap-x-2">
                                <div className="font-bold">LRN:</div>
                                <div className="flex min-h-[18px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                  {s.lrn || ""}
                                </div>
                                <div className="grid grid-cols-[52px_74px_64px_74px] items-end gap-x-2">
                                  <div className="font-bold">Grade:</div>
                                  <div className="flex min-h-[18px] items-end justify-center border-b border-black px-1 pb-[1px] text-center leading-none">
                                    {klass?.grade_level || ""}
                                  </div>
                                  <div className="font-bold">Section:</div>
                                  <div className="flex min-h-[18px] items-end justify-center border-b border-black px-1 pb-[1px] text-center leading-none">
                                    {klass?.section || ""}
                                  </div>
                                </div>
                              </div>

                              <div className="grid grid-cols-[122px_160px] items-end gap-x-2">
                                <div className="font-bold">Track (SHS only):</div>
                                <div className="flex min-h-[18px] items-end border-b border-black px-1 pb-[1px] leading-none">
                                  {klass?.track_shs || ""}
                                </div>
                              </div>
                            </div>
                          </>
                        )}

                        <div className="mt-3 text-[9.5px] italic leading-[1.18]">
                          <div className="font-bold">Dear Parents,</div>
                          <div className="indent-8 text-justify">
                            This Performance Report shows the ability and progress your child
                            has made in the different learning areas as well as his / her core
                            values. The school welcomes you should you desire to know more
                            about your child&apos;s progress.
                          </div>
                        </div>

                        <div className="mt-3 text-center text-[11.5px] font-bold">
                          LEARNING PROGRESS AND ACHIEVEMENT
                        </div>

                        <table className="mt-1 w-full table-fixed border-collapse border border-black text-[9.2px]">
                          <thead>
                            <tr>
                              <th rowSpan={2} className="w-[31%] border border-black px-1 py-1 text-center">
                                Learning Areas
                              </th>
                              <th colSpan={3} className="border border-black px-1 py-[1px] text-center">
                                TERM
                              </th>
                              <th rowSpan={2} className="w-[14%] border border-black px-1 py-1 text-center">
                                Final Grade
                              </th>
                              <th rowSpan={2} className="w-[14%] border border-black px-1 py-1 text-center">
                                Remarks
                              </th>
                            </tr>
                            <tr>
                              <th className="border border-black py-0.5 text-center">1</th>
                              <th className="border border-black py-0.5 text-center">2</th>
                              <th className="border border-black py-0.5 text-center">3</th>
                            </tr>
                          </thead>
                          <tbody>
                            {learnerRows
                              .filter((row) => Boolean(row.label?.trim()))
                              .map((row) => (
                                <tr key={row.label}>
                                  <td
                                    className={`border border-black px-1 py-[1.5px] ${
                                      row.indented ? "pl-5" : ""
                                    }`}
                                  >
                                    {row.label}
                                  </td>
                                  <td className="border border-black px-1 py-[1.5px] text-center">
                                    {row.term1 ?? ""}
                                  </td>
                                  <td className="border border-black px-1 py-[1.5px] text-center">
                                    {row.term2 ?? ""}
                                  </td>
                                  <td className="border border-black px-1 py-[1.5px] text-center">
                                    {row.term3 ?? ""}
                                  </td>
                                  <td className="border border-black px-1 py-[1.5px] text-center font-semibold">
                                    {row.final ?? ""}
                                  </td>
                                  <td className="border border-black px-1 py-[1.5px] text-center">
                                    {row.final != null
                                      ? row.final >= 75
                                        ? "Passed"
                                        : "Failed"
                                      : ""}
                                  </td>
                                </tr>
                              ))}

                            <tr className="font-bold uppercase">
                              <td colSpan={4} className="border border-black px-2 py-1 text-center">
                                GENERAL AVERAGE
                              </td>
                              <td className="border border-black px-1 py-1 text-center">
                                {generalAverage ?? ""}
                              </td>
                              <td className="border border-black px-1 py-1 text-center">
                                {getGeneralAverageRemark(generalAverage)}
                              </td>
                            </tr>
                          </tbody>
                        </table>

                        <div className="mt-3 grid grid-cols-[145px_1fr] items-center gap-3">
                          <div className="text-center text-[10.5px]">
                            Performance Descriptors
                          </div>

                          <table className="w-full border-collapse border border-black text-[8.5px]">
                            <thead>
                              <tr>
                                <th className="border border-black px-1 py-[1px]">Grading Scale</th>
                                <th className="border border-black px-1 py-[1px]">Description</th>
                                <th className="border border-black px-1 py-[1px]">Remarks</th>
                              </tr>
                            </thead>
                            <tbody>
                              {SF9_PERFORMANCE_DESCRIPTORS.map((descriptor) => (
                                <tr key={descriptor.scale}>
                                  <td className="border border-black px-1 py-[1px] text-center">
                                    {descriptor.scale}
                                  </td>
                                  <td className="border border-black px-1 py-[1px] text-center">
                                    {descriptor.description}
                                  </td>
                                  <td className="border border-black px-1 py-[1px] text-center">
                                    {descriptor.remarks}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </section>

                      {/* RIGHT SIDE — spreadsheet layout */}
                      <section className="flex min-w-0 max-w-full flex-col overflow-hidden px-0.5 [&>*]:max-w-full">
                        <div className="text-center text-[13px] font-bold">
                          ATTENDANCE RECORD
                        </div>

                        <table className="mt-1 w-full table-fixed border-collapse border border-black text-[6px]">
                          <colgroup>
                            <col style={{ width: "15%" }} />
                            {SF9_ATTENDANCE_MONTHS.map((month) => (
                              <col key={month} style={{ width: "6.35%" }} />
                            ))}
                            <col style={{ width: "5%" }} />
                          </colgroup>
                          <thead>
                            <tr className="h-[58px]">
                              <th className="border-t border-l border-b border-r border-black px-1 text-[9px] font-bold">
                                Month
                              </th>

                              {SF9_ATTENDANCE_MONTHS.map((month) => (
                                <th
                                  key={month}
                                  className="relative overflow-visible border-y border-black p-0"
                                >
                                  {/* Excel-style diagonal separator. The month text stays real/editable HTML text. */}
                                  <span
                                    aria-hidden="true"
                                    className="pointer-events-none absolute bottom-0 left-0 h-[76px] w-px origin-bottom-left rotate-[43deg] bg-black"
                                  />
                                  <span className="absolute inset-0 flex items-center justify-center overflow-visible">
                                    <span className="-rotate-45 whitespace-nowrap text-[10px] font-normal leading-none">
                                      {month}
                                    </span>
                                  </span>
                                </th>
                              ))}

                              <th className="relative overflow-visible border-y border-r border-black p-0 text-[7px] font-bold">
                                <span
                                
                                />
                                <span className="absolute bottom-[3px] right-[4px] whitespace-nowrap">
                                  TOTAL
                                </span>
                              </th>
                            </tr>
                          </thead>
                          
                           <tbody>
                          {[
                            "No. of Class Days",
                            "No. of Days Present",
                            "No. of Days Absent",
                          ].map((label) => (
                            <tr key={label}>
                              <td className="overflow-hidden border border-black px-[2px] py-[5px] text-left text-[5.8px] leading-tight">
                                {label}
                              </td>

                              {SF9_ATTENDANCE_MONTHS.map((month, index) => (
                              <td
                                key={month}
                                className={
                                  index === 0
                                    ? "border-y border-r border-black"
                                    : index === SF9_ATTENDANCE_MONTHS.length - 1
                                      ? "border-y border-l border-black"
                                      : "border border-black"
                                }
                              >
                                &nbsp;
                              </td>
                            ))}

                              <td className="border-y border-r border-black">
                                &nbsp;
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        </table>

                        <div className="mt-2 text-center text-[12px] font-bold">
                          TEACHER&apos;S COMMENTS / REMARKS
                        </div>

                        <div className="mt-1 grid min-w-0 grid-cols-[52px_minmax(0,1fr)] overflow-hidden">
                          <div className="grid grid-rows-3 text-[12px] font-bold">
                            <div className="flex items-center">Term 1</div>
                            <div className="flex items-center">Term 2</div>
                            <div className="flex items-center">Term 3</div>
                          </div>
                          <div className="border border-black">
                            <div className="h-[64px] border-b border-black" />
                            <div className="h-[64px] border-b border-black" />
                            <div className="h-[64px]" />
                          </div>
                        </div>

                        <div className="mt-2 text-center text-[12px] font-normal">
                          PARENTS / GUARDIAN&apos;S SIGNATURE
                        </div>
                        <div className="mt-2 space-y-2 px-1 text-[10.5px]">
                          {["TERM 1", "TERM 2", "TERM 3"].map((term) => (
                            <div key={term} className="grid min-w-0 grid-cols-[52px_minmax(0,1fr)] items-end gap-3">
                              <div>{term}</div>
                              <div className="min-w-0 w-full border-b border-black">&nbsp;</div>
                            </div>
                          ))}
                        </div>

                        {isJuniorHighSf9 ? (
                          <>
                            {/* Grades 7–10 only: match the Excel SF9 certificate
                                and cancellation line layout exactly. */}
                            <div className="mt-3 text-center text-[12px] font-bold">
                              CERTIFICATE OF TRANSFER
                            </div>

                            <div className="mt-[7px] px-1 text-[10px] italic leading-[1.35]">
                              <div>
                                This is to certify that the above-named learner has satisfactorily completed the requirements
                              </div>
                              <div>for the grade level indicated.</div>
                            </div>

                            {/* Excel layout: both certificate fields are on the
                                same row, with a separate underline after each label. */}
                            <div className="mt-5 grid min-w-0 grid-cols-[118px_138px_190px_minmax(0,1fr)] items-end gap-x-2 px-1 text-[10.5px] font-bold">
                              <div className="whitespace-nowrap">Admitted to Grade</div>
                              <div className="h-[13px] min-w-0 border-b border-black" />
                              <div className="whitespace-nowrap">Eligible for Admission to Grade</div>
                              <div className="h-[13px] min-w-0 border-b border-black" />
                            </div>

                            {/* Approved/signature area follows the spreadsheet:
                                Adviser line on the right first, School Head line
                                below on the left. */}
                            <div className="mt-4 grid min-w-0 grid-cols-2 gap-x-10 px-1 text-[9.5px]">
                              <div className="font-bold">Approved:</div>
                              <div className="text-center">
                                <div className="min-h-[18px] w-full border-b border-black px-1">
                                  {klass?.teacher_name || profile?.full_name || ""}
                                </div>
                                <div>Adviser</div>
                              </div>

                              <div className="mt-2 text-center">
                                <div className="min-h-[18px] w-full border-b border-black px-1">
                                  {profile?.principal || ""}
                                </div>
                                <div>School Head</div>
                              </div>
                              <div />
                            </div>

                            <div className="mt-3 text-center text-[11.5px] font-bold">
                              CANCELLATION OF ELIGIBILITY TO TRANSFER
                            </div>

                            <div className="mt-2 grid min-w-0 grid-cols-[78px_minmax(0,1fr)_36px_215px] items-end gap-x-2 px-1 text-[9.5px] font-bold">
                              <div className="whitespace-nowrap">Admitted in:</div>
                              <div className="h-[13px] min-w-0 border-b border-black" />
                              <div>Date</div>
                              <div className="h-[13px] border-b border-black" />
                            </div>

                            <div className="mx-auto mt-3 w-[58%] max-w-full text-center text-[9.5px]">
                              <div className="min-h-[18px] border-b border-black">
                                {profile?.principal || ""}
                              </div>
                              <div>School Head</div>
                            </div>
                          </>
                        ) : (
                          <>
                            <div className="mt-3 text-center text-[12px] font-normal">
                              CERTIFICATE OF TRANSFER
                            </div>

                            <div className="mt-[7px] px-1 text-[10px] leading-[1.35]">
                              <div>
                                This is to certify that the above-named learner has satisfactorily completed the requirements
                              </div>
                              <div>for the grade level indicated.</div>
                            </div>

                            <div className="mt-5 space-y-[8px] px-1 text-[10.5px]">
                              <div className="flex items-end gap-2">
                                <div className="whitespace-nowrap">Admitted to Grade:</div>
                                <div className="w-[105px] shrink-0 border-b border-black">&nbsp;</div>
                              </div>

                              <div className="flex items-end gap-2">
                                <div className="whitespace-nowrap">Eligible for Admission to Grade:</div>
                                <div className="w-[105px] shrink-0 border-b border-black">&nbsp;</div>
                              </div>
                            </div>

                            <div className="mt-3 grid min-w-0 grid-cols-[62px_minmax(0,1fr)] items-end gap-3 px-1 text-[10px] font-bold">
                              <div>Approved:</div>
                              <div className="min-w-0 w-full border-b border-black">&nbsp;</div>
                            </div>

                            <div className="mt-4 grid min-w-0 grid-cols-2 gap-10 px-1 text-center text-[9.5px]">
                              <div>
                                <div className="min-h-[18px] min-w-0 w-full border-b border-black px-1">
                                  {profile?.principal || ""}
                                </div>
                                <div>School Head</div>
                              </div>
                              <div>
                                <div className="min-h-[18px] min-w-0 w-full border-b border-black px-1">
                                  {klass?.teacher_name || profile?.full_name || ""}
                                </div>
                                <div>Adviser</div>
                              </div>
                            </div>

                            <div className="mt-2 text-center text-[11.5px] font-normal">
                              CANCELLATION OF ELIGIBILITY TO TRANSFER
                            </div>
                            <div className="mt-1 grid min-w-0 grid-cols-[72px_minmax(0,1fr)_32px_minmax(0,1fr)] items-end gap-2 overflow-hidden px-1 text-[9.5px] font-bold">
                              <div>Admitted in:</div>
                              <div className="min-w-0 w-full border-b border-black">&nbsp;</div>
                              <div>Date:</div>
                              <div className="min-w-0 w-full border-b border-black">&nbsp;</div>
                            </div>

                            <div className="mx-auto mt-3 w-[58%] max-w-full text-center text-[9.5px]">
                              <div className="min-h-[18px] border-b border-black">
                                {profile?.principal || ""}
                              </div>
                              <div>School Head</div>
                            </div>
                          </>
                        )}
                      </section>
                    </div>
                  </div>
                  )}
                </div>
              );
            })}

            {toRender.length === 0 && (
              <div className="p-16 text-center text-slate-500">
                Select a learner above to preview the report card.
              </div>
            )}
          </div>
        </PdfPreviewShell>
      </div>
    );
  };
}

// Export only stable React components from this module so Vite Fast Refresh
// can reload the SF9 pages without the incompatible export warning.
export const SF9ClassicPage = makeSF9Component("classic");
export const SF9MatatagPage = makeSF9Component("matatag");
