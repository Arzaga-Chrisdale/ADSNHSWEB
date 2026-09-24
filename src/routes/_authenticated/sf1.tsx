import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useMemo, useState } from "react";
import { ArrowLeft, BookOpen, GraduationCap, Users } from "lucide-react";
import type { ClassRow, StudentRow } from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import { toast } from "sonner";
import CFB from "cfb";
import JSZip from "jszip";
import {
  type Alignment,
  type Borders,
  type Cell,
  type Fill,
  type Font,
  type PaperSize,
  type Worksheet,
} from "exceljs";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/sf1")({
  component: SF1Page,
});

type SF1StudentRow = StudentRow & {
  mother_tongue?: string | null;
  ip_ethnic_group?: string | null;
  ip_ethnic?: string | null;
  religion?: string | null;
  house_street?: string | null;
  barangay?: string | null;
  municipality_city?: string | null;
  municipality?: string | null;
  province?: string | null;
  guardian_relationship?: string | null;
  learning_modality?: string | null;
  remarks?: string | null;

  // Transfer In / Transfer Out state used by the SF1 BoSY/EoSY summary.
  enrollment_status?: "active" | "transferred_in" | "transferred_out" | string | null;
  is_active?: boolean | null;
  transfer_date?: string | null;
  transfer_effective_term?: string | null;
  previous_school?: string | null;
  destination_school?: string | null;
  transfer_reason?: string | null;
  transferred_at?: string | null;
};

type SF1ExcelOptions = {
  fileName: string;
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  adviser: string;
  schoolHead: string;
  startDate: string;
  endDate: string;
  learners: SF1StudentRow[];
};

const BLACK = { argb: "FF000000" };
const THIN_BORDER = { style: "thin" as const, color: BLACK };
const MEDIUM_BORDER = { style: "medium" as const, color: BLACK };
const NO_FILL: Fill = { type: "pattern", pattern: "none" };
const BODY_FONT: Partial<Font> = { name: "SansSerif", size: 7 };
const CENTERED: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

// Exact A:AU column proportions from the supplied SF1 spreadsheet.
const SF1_COLUMN_WIDTHS = [
  10.07, 1.67, 1.67, 3.68, 9.73, 5.03, 2.51, 5.87, 3.35, 1.67, 2.51, 2.51, 3.35,
  5.87, 5.87, 3.35, 4.19, 7.22, 0.33, 0.83, 0.83, 7.55, 0.83, 2.34, 0.16, 2.51,
  3.35, 1.67, 1.51, 0.16, 6.87, 3.18, 0.16, 1.51, 2.51, 2.68, 6.54, 0.83, 0.83,
  1.51, 7.38, 3.18, 3.02, 7.88, 1.84, 7.22, 0.16,
] as const;

const SF1_FIELD_RANGES = [
  "A:B",
  "C:F",
  "G",
  "H:I",
  "J:K",
  "L:M",
  "N",
  "O",
  "P:Q",
  "R:T",
  "U:V",
  "W:AA",
  "AB:AE",
  "AF:AJ",
  "AK:AN",
  "AO",
  "AP:AQ",
  "AR",
  "AS:AT",
] as const;

function tableBorder(strongBottom = false): Partial<Borders> {
  return {
    top: THIN_BORDER,
    bottom: strongBottom ? MEDIUM_BORDER : THIN_BORDER,
    left: THIN_BORDER,
    right: THIN_BORDER,
  };
}

function styleCell(
  cell: Cell,
  options: {
    font?: Partial<Font>;
    alignment?: Partial<Alignment>;
    border?: Partial<Borders>;
    fill?: Fill;
    numberFormat?: string;
  } = {},
) {
  cell.font = { ...BODY_FONT, ...options.font };
  cell.alignment = options.alignment || CENTERED;
  cell.border = options.border || {};
  cell.fill = options.fill || NO_FILL;
  if (options.numberFormat) cell.numFmt = options.numberFormat;
}

function mergeAndSet(
  worksheet: Worksheet,
  range: string,
  value: string | number,
  options: Parameters<typeof styleCell>[1] = {},
) {
  const [firstCell, lastCell] = range.split(":");

  if (lastCell && firstCell !== lastCell) {
    worksheet.mergeCells(range);
  }

  const cell = worksheet.getCell(firstCell);
  cell.value = value === "" ? null : value;
  styleCell(cell, options);
}

function rowRange(range: string, row: number) {
  const [start, end] = range.split(":");
  return end ? `${start}${row}:${end}${row}` : `${start}${row}`;
}

function excelBirthdate(value?: string | null) {
  if (!value) return "";

  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(value);
  if (iso) return `${iso[2]}-${iso[3]}-${iso[1]}`;

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  })
    .format(date)
    .replaceAll("/", "-");
}

function formatSF1SchoolDate(value?: string | null) {
  if (!value) return "";

  // Class dates are stored as YYYY-MM-DD. Read the date parts directly so
  // the displayed SF1 date cannot shift by one day because of timezone
  // conversion. SF1 shows dates in MM/DD/YYYY format.
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(value);
  if (iso) return `${iso[2]}/${iso[3]}/${iso[1]}`;

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }).format(date);
}

function sf1DateLine(label: "BoSY Date:" | "EoSY Date:", value: string) {
  return value ? `${label} ${value}` : label;
}

type SF1RegistrationCount = {
  male: number;
  female: number;
  total: number;
};

type SF1RegistrationSummary = {
  bosy: SF1RegistrationCount;
  eosy: SF1RegistrationCount;
  transferIn: SF1RegistrationCount;
  transferOut: SF1RegistrationCount;
};

function sf1TransferStatus(
  student: Pick<
    SF1StudentRow,
    "enrollment_status" | "previous_school"
  >,
): "in" | "out" | "none" {
  const status = String(student.enrollment_status ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, "_");

  if (status === "transferred_out" || status === "transfer_out") {
    return "out";
  }

  if (
    status === "transferred_in" ||
    status === "transfer_in" ||
    Boolean(String(student.previous_school ?? "").trim())
  ) {
    return "in";
  }

  return "none";
}

function sf1CountLearners(
  learners: readonly SF1StudentRow[],
): SF1RegistrationCount {
  const male = learners.filter(
    (student) => String(student.sex ?? "").trim().toLowerCase() === "male",
  ).length;
  const female = learners.filter(
    (student) => String(student.sex ?? "").trim().toLowerCase() === "female",
  ).length;

  return {
    male,
    female,
    total: learners.length,
  };
}

function buildSF1RegistrationSummary(
  learners: readonly SF1StudentRow[],
): SF1RegistrationSummary {
  const transferInLearners = learners.filter(
    (student) => sf1TransferStatus(student) === "in",
  );
  const transferOutLearners = learners.filter(
    (student) => sf1TransferStatus(student) === "out",
  );

  // A learner who has previous_school was a Transfer In learner even if the
  // current status later becomes transferred_out. Such a learner was never
  // part of this class's BoSY count.
  const bosyLearners = learners.filter((student) => {
    const status = sf1TransferStatus(student);
    const cameFromAnotherSchool = Boolean(
      String(student.previous_school ?? "").trim(),
    );

    return status !== "in" && !cameFromAnotherSchool;
  });

  // EoSY is the final roster: original learners + Transfer In - Transfer Out.
  const eosyLearners = learners.filter(
    (student) => sf1TransferStatus(student) !== "out",
  );

  return {
    bosy: sf1CountLearners(bosyLearners),
    eosy: sf1CountLearners(eosyLearners),
    transferIn: sf1CountLearners(transferInLearners),
    transferOut: sf1CountLearners(transferOutLearners),
  };
}

function setTableRow(
  worksheet: Worksheet,
  rowNumber: number,
  values: readonly (string | number)[],
  strong = false,
) {
  SF1_FIELD_RANGES.forEach((range, index) => {
    const centered = [2, 3, 4, 16, 17].includes(index);
    mergeAndSet(worksheet, rowRange(range, rowNumber), values[index] ?? "", {
      font: { size: 7, bold: strong },
      alignment: {
        horizontal: centered ? "center" : "left",
        vertical: "middle",
        wrapText: true,
        shrinkToFit: index === 0,
      },
      border: tableBorder(strong),
      numberFormat: index === 0 ? "@" : undefined,
    });
  });
}

function setLearnerExcelRow(
  worksheet: Worksheet,
  rowNumber: number,
  student: SF1StudentRow,
  schoolYear: string,
) {
  const sex =
    student.sex?.toLowerCase() === "male"
      ? "M"
      : student.sex?.toLowerCase() === "female"
        ? "F"
        : "—";

  setTableRow(worksheet, rowNumber, [
    student.lrn || "",
    learnerName(student),
    sex,
    excelBirthdate(student.birthdate),
    ageOnFirstFridayOfJune(student.birthdate, schoolYear),
    student.mother_tongue || "",
    student.ip_ethnic_group || student.ip_ethnic || "",
    student.religion || "",
    student.house_street || student.address || "",
    student.barangay || "",
    student.municipality_city || student.municipality || "",
    student.province || "",
    student.father_name || "",
    student.mother_name || "",
    student.guardian || "",
    student.guardian_relationship || "",
    student.contact_number || "",
    student.learning_modality || "",
    student.remarks || "",
  ]);

  // Keep all LRN digits visible instead of allowing scientific notation.
  if (student.lrn) {
    worksheet.getCell(`A${rowNumber}`).value = {
      richText: [{ text: student.lrn }],
    };
  }

  worksheet.getRow(rowNumber).height = 47;
}

function setTotalExcelRow(
  worksheet: Worksheet,
  rowNumber: number,
  count: number,
  label: string,
  strongBottom = false,
) {
  setTableRow(
    worksheet,
    rowNumber,
    [
      count,
      `<=== ${label}`,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ],
    strongBottom,
  );
  worksheet.getCell(`A${rowNumber}`).alignment = {
    horizontal: "right",
    vertical: "middle",
  };
  worksheet.getRow(rowNumber).height = 20;
}

function setExcelHeader(
  worksheet: Worksheet,
  range: string,
  value: string,
  options: { size?: number; bold?: boolean } = {},
) {
  mergeAndSet(worksheet, range, value, {
    font: {
      size: options.size ?? 6,
      bold: options.bold ?? true,
    },
    alignment: CENTERED,
    border: tableBorder(),
  });
}

function addWorkbookHeader(worksheet: Worksheet, options: SF1ExcelOptions) {
  mergeAndSet(worksheet, "A1:AT1", "School Form 1 (SF 1) School Register", {
    font: { size: 21, bold: true },
    alignment: CENTERED,
  });
  mergeAndSet(
    worksheet,
    "A2:AT2",
    "(This replaces Form 1, Master List & STS Form 2-Family Background and Profile)",
    {
      font: { size: 7, italic: true },
      alignment: CENTERED,
    },
  );

  const metadata = [
    ["A3:E3", "School ID", true],
    ["F3:I3", options.schoolId, false],
    ["J3", "Region", true],
    ["K3:O3", options.region, false],
    ["Q3:R3", "Division", true],
    ["T3:AF3", options.division, false],
    ["A4:E4", "School Name", true],
    ["F4:O4", options.schoolName, false],
    ["Q4:S4", "School Year", true],
    ["T4:X4", options.schoolYear, false],
    ["Z4:AC4", "Grade Level", true],
    ["AE4:AG4", options.gradeLevel, false],
    ["AJ4:AL4", "Section", true],
    ["AM4:AU4", options.section, false],
  ] as const;

  metadata.forEach(([range, value, isLabel]) => {
    mergeAndSet(worksheet, range, value, {
      font: { size: 8 },
      alignment: {
        horizontal: isLabel ? "right" : "center",
        vertical: "middle",
        wrapText: false,
        shrinkToFit: !isLabel,
      },
      border: isLabel ? {} : tableBorder(),
    });
  });

  setExcelHeader(worksheet, "A5:B6", "LRN");
  setExcelHeader(
    worksheet,
    "C5:F6",
    "NAME\n(Last Name, First Name, Middle Name)",
  );
  setExcelHeader(worksheet, "G5:G6", "Sex\n(M/F)");
  setExcelHeader(worksheet, "H5:I6", "BIRTH DATE\n(mm/dd/yyyy)");
  setExcelHeader(worksheet, "J5:K6", "AGE as of\n1st Friday June", { size: 7 });
  setExcelHeader(worksheet, "L5:M6", "MOTHER TONGUE\n(Grade 1 to 3 Only)");
  setExcelHeader(worksheet, "N5:N6", "IP\n(Ethnic Group)");
  setExcelHeader(worksheet, "O5:O6", "RELIGION");
  setExcelHeader(worksheet, "P5:AA5", "ADDRESS");
  setExcelHeader(worksheet, "AB5:AJ5", "PARENTS");
  setExcelHeader(worksheet, "AK5:AO5", "GUARDIAN\n(if Not Parent)");
  setExcelHeader(worksheet, "AP5:AQ6", "Contact Number of\nParent or Guardian");
  setExcelHeader(worksheet, "AR5:AR6", "Learning\nModality");
  setExcelHeader(worksheet, "AS5:AT5", "REMARKS");
  setExcelHeader(worksheet, "P6:Q6", "House #/ Street/\nSitio/ Purok");
  setExcelHeader(worksheet, "R6:T6", "Barangay");
  setExcelHeader(worksheet, "U6:V6", "Municipality/\nCity");
  setExcelHeader(worksheet, "W6:AA6", "Province");
  setExcelHeader(
    worksheet,
    "AB6:AE6",
    "Father's Name\n(Last Name, First Name,\nMiddle Name)",
  );
  setExcelHeader(
    worksheet,
    "AF6:AJ6",
    "Mother's Maiden Name\n(Last Name, First Name,\nMiddle Name)",
  );
  setExcelHeader(worksheet, "AK6:AN6", "Name");
  setExcelHeader(worksheet, "AO6", "Relationship");
  setExcelHeader(
    worksheet,
    "AS6:AT6",
    "(Please refer to\nthe legend on last\npage)",
  );

  worksheet.getRow(1).height = 30;
  worksheet.getRow(2).height = 20;
  worksheet.getRow(3).height = 20;
  worksheet.getRow(4).height = 20;
  worksheet.getRow(5).height = 25;
  worksheet.getRow(6).height = 40;
}

function addLegendAndSignatures(
  worksheet: Worksheet,
  startRow: number,
  options: SF1ExcelOptions,
  counts: SF1RegistrationSummary,
) {
  const headingRow = startRow;
  const columnHeadingRow = startRow + 1;
  const contentStartRow = startRow + 2;
  const contentEndRow = startRow + 8;
  const generatedRow = startRow + 11;

  mergeAndSet(
    worksheet,
    `A${headingRow}:S${headingRow}`,
    "List and Code of Indicators under REMARKS column",
    {
      font: { size: 10, bold: true },
      alignment: CENTERED,
      border: { bottom: THIN_BORDER },
    },
  );

  const legendHeaders = [
    [`A${columnHeadingRow}`, "Indicator"],
    [`B${columnHeadingRow}:C${columnHeadingRow}`, "Code"],
    [`D${columnHeadingRow}:H${columnHeadingRow}`, "Required Information"],
    [`I${columnHeadingRow}:L${columnHeadingRow}`, "Indicator"],
    [`M${columnHeadingRow}`, "Code"],
    [`N${columnHeadingRow}:S${columnHeadingRow}`, "Required Information"],
  ] as const;

  legendHeaders.forEach(([range, value]) => {
    mergeAndSet(worksheet, range, value, {
      font: { size: 6, bold: true },
      alignment: { horizontal: "left", vertical: "middle", wrapText: true },
      border: tableBorder(),
    });
  });

  const legendBlocks = [
    [
      `A${contentStartRow}:A${contentEndRow}`,
      "Transferred Out\n\nTransferred In\n\nDropped\nLate Enrollment",
    ],
    [`B${contentStartRow}:C${contentEndRow}`, "T/O\n\nT/I\n\nDRP\nLE"],
    [
      `D${contentStartRow}:H${contentEndRow}`,
      "Name of Public (P) or Private (PR) School & Effectivity Date\n\nName of Public (P) or Private (PR) School & Effectivity Date\n\nReason and Effectivity Date\nReason (Enrollment beyond 1st Friday of SY)",
    ],
    [
      `I${contentStartRow}:L${contentEndRow}`,
      "CCT Recipient\n\nBalik Aral\n\nSpecial Needs Education\nAccelerated",
    ],
    [`M${contentStartRow}:M${contentEndRow}`, "CCT\n\nB/A\n\nSNED\nACL"],
    [
      `N${contentStartRow}:S${contentEndRow}`,
      "CCT Control/reference number & Effectivity Date\n\nName of school last attended & Year\n\nSpecify\nSpecify Level & Effectivity Date",
    ],
  ] as const;

  legendBlocks.forEach(([range, value], index) => {
    mergeAndSet(worksheet, range, value, {
      font: {
        size: 6,
        bold: index === 0 || index === 1 || index === 3 || index === 4,
      },
      alignment: {
        horizontal: "left",
        vertical: "top",
        wrapText: true,
      },
      border: tableBorder(),
    });
  });

  const registrationHeaders = [
    [`V${columnHeadingRow}:W${columnHeadingRow}`, "REGISTERED"],
    [`X${columnHeadingRow}:Z${columnHeadingRow}`, "BoSY"],
    [`AA${columnHeadingRow}:AB${columnHeadingRow}`, "EoSY"],
  ] as const;

  registrationHeaders.forEach(([range, value]) => {
    mergeAndSet(worksheet, range, value, {
      font: { size: 6, bold: true },
      alignment: CENTERED,
      border: tableBorder(),
    });
  });

  const registrationRows = [
    [
      contentStartRow,
      contentStartRow + 2,
      "MALE",
      counts.bosy.male,
      counts.eosy.male,
    ],
    [
      contentStartRow + 3,
      contentStartRow + 4,
      "FEMALE",
      counts.bosy.female,
      counts.eosy.female,
    ],
    [
      contentStartRow + 5,
      contentEndRow,
      "TOTAL",
      counts.bosy.total,
      counts.eosy.total,
    ],
  ] as const;

  registrationRows.forEach(
    ([fromRow, toRow, label, bosyValue, eosyValue]) => {
      mergeAndSet(worksheet, `V${fromRow}:W${toRow}`, label, {
        font: { size: 7, bold: true },
        alignment: CENTERED,
        border: tableBorder(),
      });
      mergeAndSet(worksheet, `X${fromRow}:Z${toRow}`, bosyValue, {
        font: { size: 7, bold: true },
        alignment: CENTERED,
        border: tableBorder(),
        numberFormat: "0",
      });
      mergeAndSet(worksheet, `AA${fromRow}:AB${toRow}`, eosyValue, {
        font: { size: 7, bold: true },
        alignment: CENTERED,
        border: tableBorder(),
        numberFormat: "0",
      });
    },
  );

  mergeAndSet(
    worksheet,
    `AE${columnHeadingRow}:AK${columnHeadingRow}`,
    "Prepared by:",
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );
  mergeAndSet(
    worksheet,
    `AE${contentStartRow}:AK${contentStartRow + 2}`,
    options.adviser.toUpperCase(),
    {
      font: { size: 8 },
      alignment: {
        horizontal: "center",
        vertical: "bottom",
        shrinkToFit: true,
      },
      border: { bottom: MEDIUM_BORDER },
    },
  );
  mergeAndSet(
    worksheet,
    `AE${contentStartRow + 3}:AK${contentStartRow + 4}`,
    "(Signature of Adviser over Printed Name)",
    {
      font: { size: 7, bold: true },
      alignment: CENTERED,
    },
  );
  mergeAndSet(
    worksheet,
    `AE${contentStartRow + 5}:AH${contentEndRow}`,
    sf1DateLine("BoSY Date:", options.startDate),
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "bottom" },
      border: { bottom: MEDIUM_BORDER },
    },
  );
  mergeAndSet(
    worksheet,
    `AI${contentStartRow + 5}:AK${contentEndRow}`,
    sf1DateLine("EoSY Date:", options.endDate),
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "bottom" },
      border: { bottom: MEDIUM_BORDER },
    },
  );

  mergeAndSet(
    worksheet,
    `AN${columnHeadingRow}:AS${columnHeadingRow}`,
    "Certified Correct:",
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );
  mergeAndSet(
    worksheet,
    `AN${contentStartRow}:AS${contentStartRow}`,
    options.schoolHead.toUpperCase(),
    {
      font: { size: 8 },
      alignment: {
        horizontal: "center",
        vertical: "bottom",
        shrinkToFit: true,
      },
    },
  );
  mergeAndSet(
    worksheet,
    `AN${contentStartRow + 2}:AS${contentStartRow + 3}`,
    "(Signature of School Head over Printed Name)",
    {
      font: { size: 7, bold: true },
      alignment: CENTERED,
    },
  );
  mergeAndSet(
    worksheet,
    `AN${contentStartRow + 4}:AP${contentStartRow + 5}`,
    sf1DateLine("BoSY Date:", options.startDate),
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "bottom" },
      border: { bottom: MEDIUM_BORDER },
    },
  );
  mergeAndSet(
    worksheet,
    `AQ${contentStartRow + 4}:AS${contentStartRow + 5}`,
    sf1DateLine("EoSY Date:", options.endDate),
    {
      font: { size: 7, bold: true },
      alignment: { horizontal: "left", vertical: "bottom" },
      border: { bottom: MEDIUM_BORDER },
    },
  );
  mergeAndSet(
    worksheet,
    `AN${contentStartRow + 7}:AS${contentEndRow + 2}`,
    "Generated thru LIS",
    {
      font: { size: 7, bold: true },
      alignment: CENTERED,
      border: { top: MEDIUM_BORDER },
    },
  );

  mergeAndSet(
    worksheet,
    `A${generatedRow}:AK${generatedRow}`,
    `Generated on: ${new Intl.DateTimeFormat("en-US", {
      weekday: "long",
      month: "long",
      day: "numeric",
      year: "numeric",
    }).format(new Date())}`,
    {
      font: { size: 7 },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );

  worksheet.getRow(headingRow).height = 15;
  worksheet.getRow(columnHeadingRow).height = 15;
  for (let row = contentStartRow; row <= contentEndRow; row += 1) {
    worksheet.getRow(row).height = row === contentStartRow ? 14 : 10;
  }
  worksheet.getRow(generatedRow).height = 20;

  return generatedRow;
}

const TEMPLATE_LEARNER_COLUMNS = [
  "A",
  "C",
  "G",
  "H",
  "J",
  "L",
  "N",
  "O",
  "P",
  "R",
  "U",
  "W",
  "AB",
  "AF",
  "AK",
  "AO",
  "AP",
  "AR",
  "AS",
] as const;

const SF1_LEGACY_TEMPLATE_URL = "/templates/SF1-class-adviser.xls";
const SF1_WORD_TEMPLATE_URL = "/templates/SF1-class-adviser.docx";
const SF1_EXCEL_MALE_ROWS = 20;
const SF1_EXCEL_FEMALE_ROWS = 18;
const SF1_WORD_MALE_ROWS = 20;
const SF1_WORD_FEMALE_ROWS = 18;
const SF1_WORD_TABLE_WIDTH_SCALE = 0.974;
const SF1_WORD_DATE_LINES_SCALE = 0.65;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function directWordChildren(parent: Element, localName: string) {
  return Array.from(parent.childNodes).filter(
    (node): node is Element =>
      node.nodeType === Node.ELEMENT_NODE &&
      (node as Element).localName === localName,
  );
}

function scaleWordDxaAttribute(element: Element, scale: number) {
  const width = element.getAttributeNS(WORD_NS, "w");
  const parsedWidth = width ? Number(width) : Number.NaN;
  if (!Number.isFinite(parsedWidth) || parsedWidth <= 0) return;

  element.setAttributeNS(
    WORD_NS,
    "w:w",
    String(Math.max(1, Math.round(parsedWidth * scale))),
  );
}

function scaleWordTableGeometry(table: Element, scale: number) {
  const tableProperties = directWordChildren(table, "tblPr")[0];
  const tableWidth = tableProperties
    ? directWordChildren(tableProperties, "tblW").find(
        (width) => width.getAttributeNS(WORD_NS, "type") === "dxa",
      )
    : undefined;
  if (tableWidth) scaleWordDxaAttribute(tableWidth, scale);

  const tableGrid = directWordChildren(table, "tblGrid")[0];
  if (tableGrid) {
    directWordChildren(tableGrid, "gridCol").forEach((column) => {
      scaleWordDxaAttribute(column, scale);
    });
  }

  wordRows(table).forEach((row) => {
    wordCells(row).forEach((cell) => {
      const cellProperties = directWordChildren(cell, "tcPr")[0];
      const cellWidth = cellProperties
        ? directWordChildren(cellProperties, "tcW").find(
            (width) => width.getAttributeNS(WORD_NS, "type") === "dxa",
          )
        : undefined;
      if (cellWidth) scaleWordDxaAttribute(cellWidth, scale);
    });
  });

  directWordChildren(table, "tr").forEach((row) => {
    directWordChildren(row, "tc").forEach((cell) => {
      directWordChildren(cell, "tbl").forEach((nestedTable) => {
        scaleWordTableGeometry(nestedTable, scale);
      });
    });
  });
}

function resizeWordDateTable(table: Element, scale: number) {
  const tableProperties = directWordChildren(table, "tblPr")[0];
  const tableWidth = tableProperties
    ? directWordChildren(tableProperties, "tblW").find(
        (width) => width.getAttributeNS(WORD_NS, "type") === "dxa",
      )
    : undefined;
  const originalWidth = Number(tableWidth?.getAttributeNS(WORD_NS, "w"));
  if (!tableWidth || !Number.isFinite(originalWidth) || originalWidth <= 0)
    return;

  // Use one exact width for both cells so the BoSY and EoSY lines are equal.
  const cellWidth = Math.max(1, Math.round((originalWidth * scale) / 2));
  tableWidth.setAttributeNS(WORD_NS, "w:w", String(cellWidth * 2));

  const tableGrid = directWordChildren(table, "tblGrid")[0];
  if (tableGrid) {
    directWordChildren(tableGrid, "gridCol").forEach((column) => {
      column.setAttributeNS(WORD_NS, "w:w", String(cellWidth));
    });
  }

  wordRows(table).forEach((row) => {
    wordCells(row).forEach((cell) => {
      const cellProperties = directWordChildren(cell, "tcPr")[0];
      const width = cellProperties
        ? directWordChildren(cellProperties, "tcW").find(
            (item) => item.getAttributeNS(WORD_NS, "type") === "dxa",
          )
        : undefined;
      width?.setAttributeNS(WORD_NS, "w:w", String(cellWidth));
    });
  });
}

function setWordCellText(cell: Element, value: string | number) {
  const textNodes = Array.from(cell.getElementsByTagNameNS(WORD_NS, "t"));
  const text = String(value ?? "");

  if (textNodes.length > 0) {
    textNodes[0].textContent = text;
    textNodes[0].setAttributeNS(
      "http://www.w3.org/XML/1998/namespace",
      "xml:space",
      "preserve",
    );
    textNodes.slice(1).forEach((node) => {
      node.textContent = "";
    });
    return;
  }

  const run = cell.getElementsByTagNameNS(WORD_NS, "r")[0];
  const paragraph = cell.getElementsByTagNameNS(WORD_NS, "p")[0];
  if (!paragraph) {
    throw new Error(
      "The SF1 Word template contains a cell without a paragraph.",
    );
  }

  const targetRun = run || cell.ownerDocument.createElementNS(WORD_NS, "w:r");
  const textNode = cell.ownerDocument.createElementNS(WORD_NS, "w:t");
  textNode.setAttributeNS(
    "http://www.w3.org/XML/1998/namespace",
    "xml:space",
    "preserve",
  );
  textNode.textContent = text;
  targetRun.appendChild(textNode);
  if (!run) paragraph.appendChild(targetRun);
}

function replaceWordUnderline(container: Element, value: string) {
  const textNodes = Array.from(container.getElementsByTagNameNS(WORD_NS, "t"));
  const underlineNode = textNodes.find((node) =>
    /_{4,}/.test(node.textContent || ""),
  );
  if (!underlineNode) {
    throw new Error(
      "The SF1 Word template is missing a required signature/data line.",
    );
  }

  if (!value.trim()) {
    // Keep the original template's name/signature line unchanged when blank.
    return;
  }

  underlineNode.textContent = value;
  underlineNode.setAttributeNS(
    "http://www.w3.org/XML/1998/namespace",
    "xml:space",
    "preserve",
  );
}

function setWordSF1Dates(
  container: Element,
  startDate: string,
  endDate: string,
) {
  const dateTable = Array.from(
    container.getElementsByTagNameNS(WORD_NS, "tbl"),
  ).find(
    (table) =>
      table.textContent?.includes("BoSY Date:") &&
      table.textContent?.includes("EoSY Date:"),
  );

  if (!dateTable) {
    throw new Error(
      "The original SF1 Word template is missing its BoSY/EoSY date fields.",
    );
  }

  resizeWordDateTable(dateTable, SF1_WORD_DATE_LINES_SCALE);

  const dateRow = wordRows(dateTable)[0];
  const dateCells = dateRow ? wordCells(dateRow) : [];
  if (dateCells.length < 2) {
    throw new Error(
      "The original SF1 Word template has an invalid BoSY/EoSY date row.",
    );
  }

  setWordCellText(
    dateCells[0],
    sf1DateLine("BoSY Date:", startDate),
  );
  setWordCellText(
    dateCells[1],
    sf1DateLine("EoSY Date:", endDate),
  );
}

function wordRows(table: Element) {
  return directWordChildren(table, "tr");
}

function wordCells(row: Element) {
  return directWordChildren(row, "tc");
}

function requireWordCell(table: Element, rowIndex: number, cellIndex: number) {
  const row = wordRows(table)[rowIndex];
  const cell = row ? wordCells(row)[cellIndex] : undefined;
  if (!cell) {
    throw new Error(
      `The SF1 Word template is missing table cell ${rowIndex}:${cellIndex}.`,
    );
  }
  return cell;
}

function findWordRowByText(table: Element, text: string) {
  return wordRows(table).find((row) => row.textContent?.includes(text));
}

function resizeWordLearnerRows(
  table: Element,
  options: {
    firstRowIndex?: number;
    afterRowText?: string;
    beforeRowText: string;
    desiredCount: number;
  },
) {
  const rows = wordRows(table);
  const beforeRow = findWordRowByText(table, options.beforeRowText);
  if (!beforeRow) {
    throw new Error(
      `The original SF1 Word template is missing "${options.beforeRowText}".`,
    );
  }

  const beforeIndex = rows.indexOf(beforeRow);
  let firstIndex = options.firstRowIndex ?? 0;

  if (options.afterRowText) {
    const afterRow = findWordRowByText(table, options.afterRowText);
    if (!afterRow) {
      throw new Error(
        `The original SF1 Word template is missing "${options.afterRowText}".`,
      );
    }
    firstIndex = rows.indexOf(afterRow) + 1;
  }

  const learnerRows = rows.slice(firstIndex, beforeIndex);
  if (learnerRows.length === 0 && options.desiredCount > 0) {
    throw new Error(
      `The original SF1 Word template has no learner row before "${options.beforeRowText}".`,
    );
  }

  const templateRow = learnerRows[learnerRows.length - 1];
  while (learnerRows.length < options.desiredCount) {
    if (!templateRow || !beforeRow.parentNode) break;
    const clone = templateRow.cloneNode(true) as Element;
    wordCells(clone).forEach((cell) => setWordCellText(cell, ""));
    beforeRow.parentNode.insertBefore(clone, beforeRow);
    learnerRows.push(clone);
  }

  while (learnerRows.length > options.desiredCount) {
    const row = learnerRows.pop();
    row?.parentNode?.removeChild(row);
  }

  return learnerRows;
}

type SF1LegacyCellUpdate = {
  row: number;
  column: number;
  value: string | number;
};

function sf1LearnerExportValues(
  student: SF1StudentRow | undefined,
  schoolYear: string,
): Array<string | number> {
  if (!student) return TEMPLATE_LEARNER_COLUMNS.map(() => "");

  return [
    student.lrn || "",
    learnerName(student),
    student.sex?.toLowerCase() === "male"
      ? "M"
      : student.sex?.toLowerCase() === "female"
        ? "F"
        : "",
    excelBirthdate(student.birthdate),
    ageOnFirstFridayOfJune(student.birthdate, schoolYear),
    student.mother_tongue || "",
    student.ip_ethnic_group || student.ip_ethnic || "",
    student.religion || "",
    student.house_street || student.address || "",
    student.barangay || "",
    student.municipality_city || student.municipality || "",
    student.province || "",
    student.father_name || "",
    student.mother_name || "",
    student.guardian || "",
    student.guardian_relationship || "",
    student.contact_number || "",
    student.learning_modality || "",
    student.remarks || "",
  ];
}

function excelColumnIndex(column: string) {
  return (
    column
      .toUpperCase()
      .split("")
      .reduce(
        (value, character) => value * 26 + character.charCodeAt(0) - 64,
        0,
      ) - 1
  );
}

function legacyCellUpdate(
  address: string,
  value: string | number,
): SF1LegacyCellUpdate {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) throw new Error(`Invalid SF1 template cell address: ${address}.`);

  return {
    row: Number(match[2]) - 1,
    column: excelColumnIndex(match[1]),
    value,
  };
}

function buildSF1LegacyCellUpdates(options: SF1ExcelOptions) {
  const registration = buildSF1RegistrationSummary(options.learners);

  const male = options.learners.filter(
    (student) => student.sex?.toLowerCase() === "male",
  );
  const female = options.learners.filter(
    (student) => student.sex?.toLowerCase() === "female",
  );
  const unspecified = options.learners.filter(
    (student) =>
      student.sex?.toLowerCase() !== "male" &&
      student.sex?.toLowerCase() !== "female",
  );
  const femaleAndUnspecified = [...female, ...unspecified];

  if (
    male.length > SF1_EXCEL_MALE_ROWS ||
    femaleAndUnspecified.length > SF1_EXCEL_FEMALE_ROWS
  ) {
    throw new Error(
      `Your original SF1 .xls template supports up to ${SF1_EXCEL_MALE_ROWS} male rows and ${SF1_EXCEL_FEMALE_ROWS} female/unspecified rows.`,
    );
  }

  const updates: SF1LegacyCellUpdate[] = [
    legacyCellUpdate("F3", options.schoolId),
    legacyCellUpdate("K3", options.region),
    legacyCellUpdate("T3", options.division),
    legacyCellUpdate("F4", options.schoolName),
    legacyCellUpdate("T4", options.schoolYear),
    legacyCellUpdate("AE4", options.gradeLevel),
    legacyCellUpdate("AM4", options.section),
  ];

  for (let index = 0; index < SF1_EXCEL_MALE_ROWS; index += 1) {
    const values = sf1LearnerExportValues(male[index], options.schoolYear);
    TEMPLATE_LEARNER_COLUMNS.forEach((column, columnIndex) => {
      updates.push(
        legacyCellUpdate(`${column}${index + 7}`, values[columnIndex] ?? ""),
      );
    });
  }
  for (let index = 0; index < SF1_EXCEL_FEMALE_ROWS; index += 1) {
    const values = sf1LearnerExportValues(
      femaleAndUnspecified[index],
      options.schoolYear,
    );
    TEMPLATE_LEARNER_COLUMNS.forEach((column, columnIndex) => {
      updates.push(
        legacyCellUpdate(`${column}${index + 28}`, values[columnIndex] ?? ""),
      );
    });
  }

  updates.push(
    legacyCellUpdate("A27", male.length),
    legacyCellUpdate("A46", femaleAndUnspecified.length),
    legacyCellUpdate("A47", options.learners.length),
    // REGISTERED summary: BoSY and EoSY.
    legacyCellUpdate("X50", registration.bosy.male),
    legacyCellUpdate("X53", registration.bosy.female),
    legacyCellUpdate("X55", registration.bosy.total),
    legacyCellUpdate("AA50", registration.eosy.male),
    legacyCellUpdate("AA53", registration.eosy.female),
    legacyCellUpdate("AA55", registration.eosy.total),
    legacyCellUpdate("AE50", options.adviser.toUpperCase()),
    legacyCellUpdate("AN50", options.schoolHead.toUpperCase()),
    // Use the E-Class Record class dates in the original SF1 footer.
    // Prepared by: BoSY / EoSY date lines.
    legacyCellUpdate(
      "AE55",
      sf1DateLine("BoSY Date:", options.startDate),
    ),
    legacyCellUpdate(
      "AI55",
      sf1DateLine("EoSY Date:", options.endDate),
    ),
    // Certified Correct: BoSY / EoSY date lines.
    legacyCellUpdate(
      "AN54",
      sf1DateLine("BoSY Date:", options.startDate),
    ),
    legacyCellUpdate(
      "AQ54",
      sf1DateLine("EoSY Date:", options.endDate),
    ),
    legacyCellUpdate(
      "A59",
      `Generated on: ${new Intl.DateTimeFormat("en-US", {
        weekday: "long",
        month: "long",
        day: "numeric",
        year: "numeric",
      }).format(new Date())}`,
    ),
  );

  return updates;
}

function buildSF1LegacyHiddenRows(options: SF1ExcelOptions) {
  const maleCount = options.learners.filter(
    (student) => student.sex?.toLowerCase() === "male",
  ).length;
  const femaleAndUnspecifiedCount = options.learners.length - maleCount;
  const hiddenRows = new Set<number>();

  // BIFF row indexes are zero-based. Keep populated learner rows visible and
  // hide only the unused template rows so each total moves directly below the
  // last learner without changing the user's original .xls template.
  for (let row = 6 + maleCount; row <= 25; row += 1) hiddenRows.add(row);
  for (let row = 27 + femaleAndUnspecifiedCount; row <= 44; row += 1)
    hiddenRows.add(row);

  return hiddenRows;
}

function readUInt16LE(bytes: Uint8Array, offset: number) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function uint16LE(value: number) {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function concatenateBytes(chunks: Uint8Array[]) {
  const result = new Uint8Array(
    chunks.reduce((length, chunk) => length + chunk.length, 0),
  );
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

function makeBiffRecord(id: number, body: Uint8Array) {
  return concatenateBytes([uint16LE(id), uint16LE(body.length), body]);
}

function makeBiffBlankRecord(row: number, column: number, xf: number) {
  return makeBiffRecord(
    0x0201,
    concatenateBytes([uint16LE(row), uint16LE(column), uint16LE(xf)]),
  );
}

function makeBiffNumberRecord(
  row: number,
  column: number,
  xf: number,
  value: number,
) {
  const body = new Uint8Array(14);
  body.set(uint16LE(row), 0);
  body.set(uint16LE(column), 2);
  body.set(uint16LE(xf), 4);
  new DataView(body.buffer).setFloat64(6, value, true);
  return makeBiffRecord(0x0203, body);
}

function makeBiffLabelRecord(
  row: number,
  column: number,
  xf: number,
  value: string,
) {
  const text = value.normalize("NFC").slice(0, 255);
  let wide = false;
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) > 0xff) {
      wide = true;
      break;
    }
  }

  const characters = new Uint8Array(text.length * (wide ? 2 : 1));
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    characters[index * (wide ? 2 : 1)] = code & 0xff;
    if (wide) characters[index * 2 + 1] = (code >>> 8) & 0xff;
  }

  return makeBiffRecord(
    0x0204,
    concatenateBytes([
      uint16LE(row),
      uint16LE(column),
      uint16LE(xf),
      uint16LE(text.length),
      Uint8Array.of(wide ? 1 : 0),
      characters,
    ]),
  );
}

function makeBiffValueRecord(
  row: number,
  column: number,
  xf: number,
  value: string | number,
) {
  if (value === "") return makeBiffBlankRecord(row, column, xf);
  if (typeof value === "number" && Number.isFinite(value)) {
    return makeBiffNumberRecord(row, column, xf, value);
  }
  return makeBiffLabelRecord(row, column, xf, String(value));
}

function makeBiffMulBlankRecord(
  row: number,
  firstColumn: number,
  xfs: number[],
) {
  if (xfs.length === 0) return undefined;
  return makeBiffRecord(
    0x00be,
    concatenateBytes([
      uint16LE(row),
      uint16LE(firstColumn),
      ...xfs.map((xf) => uint16LE(xf)),
      uint16LE(firstColumn + xfs.length - 1),
    ]),
  );
}

function patchOriginalSF1Workbook(
  workbookBytes: Uint8Array,
  cellUpdates: SF1LegacyCellUpdate[],
  hiddenRows: ReadonlySet<number>,
) {
  const pending = new Map<string, SF1LegacyCellUpdate>(
    cellUpdates.map(
      (update) => [`${update.row}:${update.column}`, update] as const,
    ),
  );
  const chunks: Uint8Array[] = [];
  const singleCellRecordIds = new Set([
    0x0006, 0x00fd, 0x0201, 0x0203, 0x0204, 0x0205, 0x027e,
  ]);
  let offset = 0;
  let worksheetIndex = -1;
  let inPrimaryWorksheet = false;

  while (offset + 4 <= workbookBytes.length) {
    const id = readUInt16LE(workbookBytes, offset);
    const length = readUInt16LE(workbookBytes, offset + 2);
    const start = offset + 4;
    const end = start + length;
    if (end > workbookBytes.length) {
      throw new Error("The original SF1 .xls workbook stream is incomplete.");
    }

    if (id === 0x0809 && length >= 4) {
      const substreamType = readUInt16LE(workbookBytes, start + 2);
      if (substreamType === 0x0010) {
        worksheetIndex += 1;
        inPrimaryWorksheet = worksheetIndex === 0;
      }
    }

    // INDEX and DBCELL contain byte offsets that become stale when text is inserted.
    // They are optional performance records and Excel/LibreOffice rebuild them on save.
    if (id === 0x020b || id === 0x00d7) {
      offset = end;
      continue;
    }

    if (inPrimaryWorksheet && id === 0x0208 && length >= 16) {
      const row = readUInt16LE(workbookBytes, start);
      if (hiddenRows.has(row)) {
        const hiddenRowRecord = workbookBytes.slice(offset, end);
        // ROW.fDyZero is bit 5 of the first option byte. Preserve miyRw so
        // Excel can restore the original row height if the row is unhidden.
        hiddenRowRecord[16] |= 0x20;
        chunks.push(hiddenRowRecord);
      } else {
        chunks.push(workbookBytes.slice(offset, end));
      }
    } else if (inPrimaryWorksheet && id === 0x00be && length >= 8) {
      const row = readUInt16LE(workbookBytes, start);
      const firstColumn = readUInt16LE(workbookBytes, start + 2);
      const lastColumn = readUInt16LE(workbookBytes, end - 2);
      const xfs: number[] = [];
      for (let cursor = start + 4; cursor < end - 2; cursor += 2) {
        xfs.push(readUInt16LE(workbookBytes, cursor));
      }

      const rowUpdates = Array.from(pending.values())
        .filter(
          (update) =>
            update.row === row &&
            update.column >= firstColumn &&
            update.column <= lastColumn,
        )
        .sort((left, right) => left.column - right.column);

      if (rowUpdates.length > 0) {
        let cursorColumn = firstColumn;
        rowUpdates.forEach((update) => {
          const before = makeBiffMulBlankRecord(
            row,
            cursorColumn,
            xfs.slice(cursorColumn - firstColumn, update.column - firstColumn),
          );
          if (before) chunks.push(before);
          chunks.push(
            makeBiffValueRecord(
              row,
              update.column,
              xfs[update.column - firstColumn],
              update.value,
            ),
          );
          pending.delete(`${update.row}:${update.column}`);
          cursorColumn = update.column + 1;
        });

        const after = makeBiffMulBlankRecord(
          row,
          cursorColumn,
          xfs.slice(cursorColumn - firstColumn),
        );
        if (after) chunks.push(after);
      } else {
        chunks.push(workbookBytes.slice(offset, end));
      }
    } else if (
      inPrimaryWorksheet &&
      singleCellRecordIds.has(id) &&
      length >= 6
    ) {
      const row = readUInt16LE(workbookBytes, start);
      const column = readUInt16LE(workbookBytes, start + 2);
      const key = `${row}:${column}`;
      const update = pending.get(key);
      if (update) {
        chunks.push(
          makeBiffValueRecord(
            row,
            column,
            readUInt16LE(workbookBytes, start + 4),
            update.value,
          ),
        );
        pending.delete(key);
      } else {
        chunks.push(workbookBytes.slice(offset, end));
      }
    } else {
      chunks.push(workbookBytes.slice(offset, end));
    }

    if (id === 0x000a && inPrimaryWorksheet) inPrimaryWorksheet = false;
    offset = end;
  }

  if (offset < workbookBytes.length) chunks.push(workbookBytes.slice(offset));
  if (pending.size > 0) {
    const first = pending.values().next().value as
      SF1LegacyCellUpdate | undefined;
    throw new Error(
      first
        ? `The original SF1 .xls template is missing cell row ${first.row + 1}, column ${first.column + 1}.`
        : "The original SF1 .xls template is missing an expected cell.",
    );
  }

  return concatenateBytes(chunks);
}

function setTemplateCell(
  worksheet: Worksheet,
  address: string,
  value: string | number,
) {
  worksheet.getCell(address).value = value === "" ? null : value;
}

function setTemplateLearnerRow(
  worksheet: Worksheet,
  row: number,
  student: SF1StudentRow | undefined,
  schoolYear: string,
) {
  const values: Array<string | number> = student
    ? [
        student.lrn || "",
        learnerName(student),
        student.sex?.toLowerCase() === "male"
          ? "M"
          : student.sex?.toLowerCase() === "female"
            ? "F"
            : "",
        excelBirthdate(student.birthdate),
        ageOnFirstFridayOfJune(student.birthdate, schoolYear),
        student.mother_tongue || "",
        student.ip_ethnic_group || student.ip_ethnic || "",
        student.religion || "",
        student.house_street || student.address || "",
        student.barangay || "",
        student.municipality_city || student.municipality || "",
        student.province || "",
        student.father_name || "",
        student.mother_name || "",
        student.guardian || "",
        student.guardian_relationship || "",
        student.contact_number || "",
        student.learning_modality || "",
        student.remarks || "",
      ]
    : TEMPLATE_LEARNER_COLUMNS.map(() => "");

  TEMPLATE_LEARNER_COLUMNS.forEach((column, index) => {
    setTemplateCell(worksheet, `${column}${row}`, values[index] ?? "");
  });
}

async function downloadSF1Excel(options: SF1ExcelOptions) {
  const response = await fetch(SF1_LEGACY_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(
      `SF1 .xls template not found (${response.status}). Put SF1-class-adviser.xls in public/templates/.`,
    );
  }

  const templateBytes = new Uint8Array(await response.arrayBuffer());
  const compoundFile = CFB.read(templateBytes, { type: "array" });
  const workbookEntry =
    CFB.find(compoundFile, "Workbook") || CFB.find(compoundFile, "Book");

  if (!workbookEntry?.content) {
    throw new Error("The SF1 .xls template has no BIFF workbook stream.");
  }

  const workbookBytes = patchOriginalSF1Workbook(
    new Uint8Array(workbookEntry.content),
    buildSF1LegacyCellUpdates(options),
    buildSF1LegacyHiddenRows(options),
  );

  workbookEntry.content = workbookBytes;
  workbookEntry.size = workbookBytes.length;

  const output = CFB.write(compoundFile, { type: "array" }) as Uint8Array;
  const outputBuffer = Uint8Array.from(output).buffer as ArrayBuffer;
  const blob = new Blob([outputBuffer], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${options.fileName}.xls`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadSF1Word(options: SF1ExcelOptions) {
  const response = await fetch(SF1_WORD_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(
      `SF1 Word template not found (${response.status}). Put SF1-class-adviser.docx in public/templates/.`,
    );
  }

  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("The SF1 Word template has no word/document.xml file.");
  }

  const parser = new DOMParser();
  const wordDocument = parser.parseFromString(
    await documentFile.async("string"),
    "application/xml",
  );
  if (wordDocument.getElementsByTagName("parsererror").length > 0) {
    throw new Error("The SF1 Word template document XML is invalid.");
  }

  const body = wordDocument.getElementsByTagNameNS(WORD_NS, "body")[0];
  const topLevelTables = body ? directWordChildren(body, "tbl") : [];
  if (topLevelTables.length < 1) {
    throw new Error(
      "The original SF1 Word template does not contain its main SF1 table.",
    );
  }

  // Use the user's original SF1 Word table exactly as the layout source.
  // The supplied template has 16 outer rows (not 28), so identify learner,
  // total and footer rows by their labels instead of hard-coded row numbers.
  const sf1Table = topLevelTables[0];
  const initialRows = wordRows(sf1Table);
  if (initialRows.length < 16) {
    throw new Error(
      "The original SF1 Word template is missing required SF1 rows.",
    );
  }

  setWordCellText(requireWordCell(sf1Table, 0, 1), options.schoolId);
  setWordCellText(requireWordCell(sf1Table, 0, 3), options.region);
  setWordCellText(requireWordCell(sf1Table, 0, 5), options.division);
  setWordCellText(requireWordCell(sf1Table, 1, 1), options.schoolName);
  setWordCellText(requireWordCell(sf1Table, 1, 3), options.schoolYear);
  setWordCellText(requireWordCell(sf1Table, 1, 5), options.gradeLevel);
  setWordCellText(requireWordCell(sf1Table, 1, 7), options.section);

  const male = options.learners.filter(
    (student) => student.sex?.toLowerCase() === "male",
  );
  const femaleAndUnspecified = options.learners.filter(
    (student) => student.sex?.toLowerCase() !== "male",
  );

  if (
    male.length > SF1_WORD_MALE_ROWS ||
    femaleAndUnspecified.length > SF1_WORD_FEMALE_ROWS
  ) {
    throw new Error(
      `The SF1 Word export supports up to ${SF1_WORD_MALE_ROWS} male rows and ${SF1_WORD_FEMALE_ROWS} female/unspecified rows. Select Short or use Excel for the complete register.`,
    );
  }

  // Keep the original row formatting. Only clone a template learner row when
  // more rows are needed, and remove unused learner rows in the downloaded copy.
  // The source .docx in public/templates is never modified.
  const maleRows = resizeWordLearnerRows(sf1Table, {
    firstRowIndex: 4,
    beforeRowText: "<=== TOTAL MALE",
    // Keep the same fixed learner-row capacity as the Excel template.
    desiredCount: SF1_WORD_MALE_ROWS,
  });

  const femaleRows = resizeWordLearnerRows(sf1Table, {
    afterRowText: "<=== TOTAL MALE",
    beforeRowText: "<=== TOTAL FEMALE",
    // Keep the same fixed learner-row capacity as the Excel template.
    desiredCount: SF1_WORD_FEMALE_ROWS,
  });

  const fillRows = (
    rows: Element[],
    learnersForRows: SF1StudentRow[],
  ) => {
    rows.forEach((row, index) => {
      const cells = wordCells(row);
      if (cells.length < TEMPLATE_LEARNER_COLUMNS.length) {
        throw new Error(
          `The original SF1 Word template is missing learner cells in row ${index + 1}.`,
        );
      }

      const values = sf1LearnerExportValues(
        learnersForRows[index],
        options.schoolYear,
      );
      cells
        .slice(0, TEMPLATE_LEARNER_COLUMNS.length)
        .forEach((cell, cellIndex) => {
          setWordCellText(cell, values[cellIndex] ?? "");
        });
    });
  };

  fillRows(maleRows, male);
  fillRows(femaleRows, femaleAndUnspecified);

  const maleTotalRow = findWordRowByText(sf1Table, "<=== TOTAL MALE");
  const femaleTotalRow = findWordRowByText(sf1Table, "<=== TOTAL FEMALE");
  const combinedRow = findWordRowByText(sf1Table, "<=== COMBINED");

  if (!maleTotalRow || !femaleTotalRow || !combinedRow) {
    throw new Error(
      "The original SF1 Word template is missing one or more total rows.",
    );
  }

  setWordCellText(wordCells(maleTotalRow)[0], male.length);
  setWordCellText(wordCells(femaleTotalRow)[0], femaleAndUnspecified.length);
  setWordCellText(wordCells(combinedRow)[0], options.learners.length);

  const currentRows = wordRows(sf1Table);
  const combinedIndex = currentRows.indexOf(combinedRow);
  const footerRow = currentRows[combinedIndex + 1];
  const footerCells = footerRow ? wordCells(footerRow) : [];
  if (footerCells.length < 4) {
    throw new Error(
      "The original SF1 Word template is missing its footer fields.",
    );
  }

  // Preserve the original template's widths and lines. Only replace the text
  // that belongs on the existing adviser and school-head signature lines.
  replaceWordUnderline(footerCells[2], options.adviser.toUpperCase());
  replaceWordUnderline(footerCells[3], options.schoolHead.toUpperCase());

  // SF1 BoSY/EoSY dates come from the same Start Date and End Date used by
  // the selected class in the E-Class Record. Apply them to both the
  // adviser and school-head date lines in the original Word template.
  setWordSF1Dates(footerCells[2], options.startDate, options.endDate);
  setWordSF1Dates(footerCells[3], options.startDate, options.endDate);

  const nestedTables = Array.from(
    sf1Table.getElementsByTagNameNS(WORD_NS, "tbl"),
  );
  const registrationTable = nestedTables.find((table) => {
    const firstCell = wordRows(table)[0]
      ? wordCells(wordRows(table)[0])[0]
      : undefined;
    return firstCell?.textContent?.includes("REGISTERED");
  });
  if (!registrationTable) {
    throw new Error(
      "The original SF1 Word template is missing the registration summary.",
    );
  }

  const registration = buildSF1RegistrationSummary(options.learners);
  const counts = [
    [registration.bosy.male, registration.eosy.male],
    [registration.bosy.female, registration.eosy.female],
    [registration.bosy.total, registration.eosy.total],
  ] as const;

  counts.forEach(([bosyCount, eosyCount], index) => {
    setWordCellText(
      requireWordCell(registrationTable, index + 1, 1),
      bosyCount,
    );
    setWordCellText(
      requireWordCell(registrationTable, index + 1, 2),
      eosyCount,
    );
  });

  const generatedText = `Generated on: ${new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date())}`;

  const generatedContainer = Array.from(
    body?.getElementsByTagNameNS(WORD_NS, "p") || [],
  ).find((paragraph) => paragraph.textContent?.includes("Generated on:"));

  if (generatedContainer) {
    const texts = Array.from(
      generatedContainer.getElementsByTagNameNS(WORD_NS, "t"),
    );
    const generatedLabel = texts.find((node) =>
      node.textContent?.includes("Generated on:"),
    );
    const underline = texts.find((node) => /_{4,}/.test(node.textContent || ""));

    if (generatedLabel) generatedLabel.textContent = generatedText;
    if (underline) underline.textContent = "";
  }

  const serialized = new XMLSerializer().serializeToString(wordDocument);
  zip.file("word/document.xml", serialized);
  const blob = await zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${options.fileName}.docx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const borderCell =
  "border border-black px-[2px] py-[1px] align-middle leading-[1.05]";
const headerCell = `${borderCell} text-center font-bold`;

function formatBirthdate(value?: string | null) {
  if (!value) return "";

  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(value);
  if (iso) return `${iso[2]}/${iso[3]}/${iso[1]}`;

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }).format(date);
}

function firstFridayOfJune(schoolYear?: string | null) {
  const matchedYear = schoolYear?.match(/\d{4}/)?.[0];
  const year = matchedYear ? Number(matchedYear) : new Date().getFullYear();
  const date = new Date(Date.UTC(year, 5, 1));
  const daysUntilFriday = (5 - date.getUTCDay() + 7) % 7;
  date.setUTCDate(date.getUTCDate() + daysUntilFriday);
  return date;
}

function ageOnFirstFridayOfJune(
  birthdate?: string | null,
  schoolYear?: string | null,
) {
  if (!birthdate) return "";

  const birth = new Date(`${birthdate.slice(0, 10)}T00:00:00Z`);
  if (Number.isNaN(birth.getTime())) return "";

  const reference = firstFridayOfJune(schoolYear);
  let age = reference.getUTCFullYear() - birth.getUTCFullYear();
  const beforeBirthday =
    reference.getUTCMonth() < birth.getUTCMonth() ||
    (reference.getUTCMonth() === birth.getUTCMonth() &&
      reference.getUTCDate() < birth.getUTCDate());

  if (beforeBirthday) age -= 1;
  return String(Math.max(0, age));
}

function learnerName(student: SF1StudentRow) {
  return [student.last_name, student.first_name, student.middle_name]
    .filter(Boolean)
    .join(", ")
    .toUpperCase();
}

function generatedDate() {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date());
}

function SF1Page() {
  const [classId] = useState(readSchoolFormsClassId);
  const [length, setLength] = useState<"short" | "full">("full");

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () =>
      (await supabase.from("profiles").select("*").maybeSingle()).data,
  });

  const { data: classes = [] } = useQuery({
    queryKey: ["classes"],
    queryFn: async () =>
      (
        await supabase
          .from("classes")
          .select("*")
          .order("created_at", { ascending: false })
      ).data as ClassRow[],
  });

  const selectedSchoolFormsClass = classes.find((item) => item.id === classId);
  const scopedClasses = selectedSchoolFormsClass ? [selectedSchoolFormsClass] : classes;
  const active = selectedSchoolFormsClass?.id || scopedClasses[0]?.id || "";
  const klass = classes.find((item) => item.id === active);

  const { data: students = [] } = useQuery({
    enabled: Boolean(active),
    queryKey: ["students", active],
    queryFn: async () =>
      (
        await supabase
          .from("students")
          .select("*")
          .eq("class_id", active)
          .order("last_name")
          .order("first_name")
      ).data as SF1StudentRow[],
  });

  const learners = useMemo(
    () =>
      [...students].sort((left, right) => {
        const lastName = left.last_name.localeCompare(right.last_name);
        return lastName || left.first_name.localeCompare(right.first_name);
      }),
    [students],
  );

  const male = useMemo(
    () => learners.filter((student) => student.sex === "male"),
    [learners],
  );
  const female = useMemo(
    () => learners.filter((student) => student.sex === "female"),
    [learners],
  );
  const unspecified = useMemo(
    () =>
      learners.filter(
        (student) => student.sex !== "male" && student.sex !== "female",
      ),
    [learners],
  );

  const registrationSummary = useMemo(
    () => buildSF1RegistrationSummary(learners),
    [learners],
  );

  const displayedMale = length === "full" ? male : male.slice(0, 10);
  const displayedFemale = length === "full" ? female : female.slice(0, 10);
  const displayedUnspecified =
    length === "full" ? unspecified : unspecified.slice(0, 4);
  const displayedCount =
    displayedMale.length + displayedFemale.length + displayedUnspecified.length;
  const learnerRowHeight = Math.max(
    8,
    Math.min(16, Math.floor(430 / Math.max(1, displayedCount + 3))),
  );
  const learnerFontSize =
    displayedCount > 46 ? 3.6 : displayedCount > 34 ? 4 : 4.5;

  const schoolName =
    profile?.school_name ||
    klass?.school_name ||
    "Agusan del Sur National Science High School";
  const schoolId = profile?.school_id || klass?.school_id || "";
  const region = profile?.region || klass?.region || "";
  const division = profile?.division || klass?.division || "";
  const adviser = klass?.teacher_name || profile?.full_name || "";
  const schoolHead = profile?.principal || "";
  const bosyDate = formatSF1SchoolDate(klass?.start_date);
  const eosyDate = formatSF1SchoolDate(klass?.end_date);

  const exportToExcel = async () => {
    if (!klass) {
      toast.error("Please select a class before exporting");
      return;
    }

    try {
      await downloadSF1Excel({
        fileName: `SF1_Register_${klass.section || "class"}`,
        schoolName,
        schoolId,
        region,
        division,
        schoolYear: klass.school_year || "",
        gradeLevel: klass.grade_level || "",
        section: klass.section || "",
        adviser,
        schoolHead,
        startDate: bosyDate,
        endDate: eosyDate,
        learners,
      });
      toast.success("SF1 downloaded as an Excel 97-2003 (.xls) workbook");
    } catch (error) {
      console.error("Unable to export SF1 to Excel", error);
      toast.error(
        error instanceof Error
          ? error.message
          : "Unable to export SF1 to Excel",
      );
    }
  };

  const copySF1ToWord = async () => {
    if (!klass) {
      toast.error("Please select a class before creating the Word document");
      return;
    }

    try {
      await downloadSF1Word({
        fileName: `SF1_Register_${klass.section || "class"}`,
        schoolName,
        schoolId,
        region,
        division,
        schoolYear: klass.school_year || "",
        gradeLevel: klass.grade_level || "",
        section: klass.section || "",
        adviser,
        schoolHead,
        startDate: bosyDate,
        endDate: eosyDate,
        learners:
          length === "full"
            ? learners
            : [...displayedMale, ...displayedFemale, ...displayedUnspecified],
      });
      toast.success("SF1 Word template downloaded with the class data");
    } catch (error) {
      console.error("Unable to create SF1 Word document", error);
      toast.error(
        error instanceof Error
          ? error.message
          : "Unable to create SF1 Word document",
      );
    }
  };

  const renderLearnerRow = (
    student: SF1StudentRow,
    sexLabel: "M" | "F" | "—",
  ) => (
    <tr key={student.id} style={{ height: learnerRowHeight }}>
      <td className={borderCell}>{student.lrn || ""}</td>
      <td className={`${borderCell} font-semibold`}>{learnerName(student)}</td>
      <td className={`${borderCell} text-center`}>{sexLabel}</td>
      <td className={`${borderCell} whitespace-nowrap text-center`}>
        {formatBirthdate(student.birthdate)}
      </td>
      <td className={`${borderCell} text-center`}>
        {ageOnFirstFridayOfJune(student.birthdate, klass?.school_year)}
      </td>
      <td className={borderCell}>{student.mother_tongue || ""}</td>
      <td className={borderCell}>
        {student.ip_ethnic_group || student.ip_ethnic || ""}
      </td>
      <td className={borderCell}>{student.religion || ""}</td>
      <td className={borderCell}>
        {student.house_street || student.address || ""}
      </td>
      <td className={borderCell}>{student.barangay || ""}</td>
      <td className={borderCell}>
        {student.municipality_city || student.municipality || ""}
      </td>
      <td className={borderCell}>{student.province || ""}</td>
      <td className={borderCell}>{student.father_name || ""}</td>
      <td className={borderCell}>{student.mother_name || ""}</td>
      <td className={borderCell}>{student.guardian || ""}</td>
      <td className={borderCell}>{student.guardian_relationship || ""}</td>
      <td className={borderCell}>{student.contact_number || ""}</td>
      <td className={`${borderCell} text-center`}>
        {student.learning_modality || ""}
      </td>
      <td className={borderCell}>{student.remarks || ""}</td>
    </tr>
  );

  const renderTotalRow = (count: number, label: string) => (
    <tr className="font-semibold" style={{ height: 12 }}>
      <td className={`${borderCell} text-right`}>{count}</td>
      <td className={borderCell}>{`<=== ${label}`}</td>
      <td className={borderCell} colSpan={17} />
    </tr>
  );

  return (
    <div className="space-y-4 pb-8">
      <div className="flex flex-wrap items-center gap-3">
        <Link
          to="/school-forms"
          className="inline-flex items-center gap-1.5 rounded-lg border-2 px-3 py-1.5 text-sm font-medium text-white shadow-sm hover:opacity-90"
          style={{ backgroundColor: DEPED_BLUE, borderColor: DEPED_BLUE }}
        >
          <ArrowLeft className="size-4" /> Back to Forms
        </Link>

        <div className="flex items-center gap-2 text-lg font-semibold">
          <BookOpen className="size-5" style={{ color: DEPED_BLUE }} />
          School Form 1 — School Register
        </div>
      </div>

      <div
        className="relative overflow-hidden rounded-2xl border-2 px-5 py-3 shadow-sm sm:px-6 sm:py-0.5"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="text-sm font-semibold text-amber-900">
          Form Configuration
        </div>

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
      </div>

      <PdfPreviewShell
        fileName={`SF1_Register_${klass?.section || "class"}.pdf`}
        docBaseName={`SF1_Register_${klass?.section || "class"}`}
        printTargetId="sf1-doc"
        length={length}
        onLengthChange={setLength}
        orientation="landscape"
        pageLabel="Page 1 of 1"
        onCopyToWord={copySF1ToWord}
        hideCopyToWord={
          String(klass?.grade_level ?? "")
            .trim()
            .toLowerCase() === "grade 12"
        }
        onExportToExcel={exportToExcel}
      >
        <div
          id="sf1-doc"
          className="box-border overflow-hidden bg-white p-[18px] text-black"
          style={{
            width: 1250,
            height: 817,
            fontFamily: "Arial, Helvetica, sans-serif",
          }}
        >
          <div className="text-center leading-tight">
            <div className="text-[17px] font-bold">
              School Form 1 (SF 1) School Register
            </div>
            <div className="mt-1 text-[6px] italic">
              (This replaces Form 1, Master List &amp; STS Form 2-Family
              Background and Profile)
            </div>
          </div>

          <table className="mt-3 w-full table-fixed border-collapse text-[6px]">
            <tbody>
              <tr>
                <td className="w-[8%] py-[2px] text-right font-semibold">
                  School ID
                </td>
                <td className="w-[18%] border border-black px-2 text-center">
                  {schoolId}
                </td>
                <td className="w-[7%] py-[2px] text-right font-semibold">
                  Region
                </td>
                <td className="w-[15%] border border-black px-2 text-center">
                  {region}
                </td>
                <td className="w-[8%] py-[2px] text-right font-semibold">
                  Division
                </td>
                <td
                  className="border border-black px-2 text-center"
                  colSpan={3}
                >
                  {division}
                </td>
              </tr>
              <tr>
                <td className="py-[2px] text-right font-semibold">
                  School Name
                </td>
                <td
                  className="border border-black px-2 text-center"
                  colSpan={3}
                >
                  {schoolName}
                </td>
                <td className="py-[2px] text-right font-semibold">
                  School Year
                </td>
                <td className="w-[12%] border border-black px-2 text-center">
                  {klass?.school_year || ""}
                </td>
                <td className="w-[8%] py-[2px] text-right font-semibold">
                  Grade Level
                </td>
                <td className="w-[13%] border border-black px-2 text-center">
                  {klass?.grade_level || ""}
                </td>
                <td className="w-[6%] py-[2px] text-right font-semibold">
                  Section
                </td>
                <td className="w-[13%] border border-black px-2 text-center">
                  {klass?.section || ""}
                </td>
              </tr>
            </tbody>
          </table>

          <table
            className="mt-[1px] w-full table-fixed border-collapse"
            style={{ fontSize: learnerFontSize }}
          >
            <colgroup>
              {[
                7, 12, 2.2, 5.5, 2.8, 4.8, 3.6, 4, 5.5, 5.5, 6, 6, 6, 6, 5, 4,
                4.5, 5.5, 4.1,
              ].map((width, index) => (
                <col
                  key={`sf1-column-${index}`}
                  style={{ width: `${width}%` }}
                />
              ))}
            </colgroup>

            <thead className="text-[4.5px]">
              <tr style={{ height: 22 }}>
                <th className={headerCell} rowSpan={2}>
                  LRN
                </th>
                <th className={headerCell} rowSpan={2}>
                  NAME
                  <br />
                  <span className="font-normal">
                    (Last Name, First Name, Middle Name)
                  </span>
                </th>
                <th className={headerCell} rowSpan={2}>
                  Sex
                  <br />
                  (M/F)
                </th>
                <th className={headerCell} rowSpan={2}>
                  BIRTH DATE
                  <br />
                  (mm/dd/yyyy)
                </th>
                <th className={headerCell} rowSpan={2}>
                  AGE as of
                  <br />
                  1st Friday June
                </th>
                <th className={headerCell} rowSpan={2}>
                  MOTHER TONGUE
                  <br />
                  (Grade 1 to 3 Only)
                </th>
                <th className={headerCell} rowSpan={2}>
                  IP
                  <br />
                  (Ethnic Group)
                </th>
                <th className={headerCell} rowSpan={2}>
                  RELIGION
                </th>
                <th className={headerCell} colSpan={4}>
                  ADDRESS
                </th>
                <th className={headerCell} colSpan={2}>
                  PARENTS
                </th>
                <th className={headerCell} colSpan={2}>
                  GUARDIAN
                  <br />
                  (if Not Parent)
                </th>
                <th className={headerCell} rowSpan={2}>
                  Contact Number of Parent or Guardian
                </th>
                <th className={headerCell} rowSpan={2}>
                  Learning Modality
                </th>
                <th className={headerCell} rowSpan={2}>
                  REMARKS
                  <br />
                  <span className="font-normal">
                    (Please refer to the legend below)
                  </span>
                </th>
              </tr>
              <tr style={{ height: 25 }}>
                <th className={headerCell}>
                  House #/Street/
                  <br />
                  Sitio/Purok
                </th>
                <th className={headerCell}>Barangay</th>
                <th className={headerCell}>Municipality/City</th>
                <th className={headerCell}>Province</th>
                <th className={headerCell}>
                  Father&apos;s Name
                  <br />
                  (Last, First, Middle)
                </th>
                <th className={headerCell}>
                  Mother&apos;s Maiden Name
                  <br />
                  (Last, First, Middle)
                </th>
                <th className={headerCell}>Name</th>
                <th className={headerCell}>Relationship</th>
              </tr>
            </thead>

            <tbody>
              {displayedMale.map((student) => renderLearnerRow(student, "M"))}
              {renderTotalRow(male.length, "TOTAL MALE")}

              {displayedFemale.map((student) => renderLearnerRow(student, "F"))}
              {renderTotalRow(female.length, "TOTAL FEMALE")}

              {displayedUnspecified.map((student) =>
                renderLearnerRow(student, "—"),
              )}
              {renderTotalRow(learners.length, "COMBINED")}
            </tbody>
          </table>

          <div className="mt-[2px] grid grid-cols-[52%_11%_18%_19%] gap-[10px] text-[4.5px] leading-[1.08]">
            <div>
              <div className="border-b border-black pb-[2px] text-center text-[6px] font-bold">
                List and Code of Indicators under REMARKS column
              </div>
              <div className="grid grid-cols-2 gap-[7px] pt-[2px]">
                <table className="w-full table-fixed border-collapse">
                  <colgroup>
                    <col style={{ width: "25%" }} />
                    <col style={{ width: "9%" }} />
                    <col style={{ width: "66%" }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className={`${borderCell} text-left`}>Indicator</th>
                      <th className={headerCell}>Code</th>
                      <th className={`${borderCell} text-left`}>
                        Required Information
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        Transferred Out
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        T/O
                      </td>
                      <td className={borderCell}>
                        Name of Public (P) or Private (PR) School &amp;
                        Effectivity Date
                      </td>
                    </tr>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        Transferred In
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        T/I
                      </td>
                      <td className={borderCell}>
                        Name of Public (P) or Private (PR) School &amp;
                        Effectivity Date
                      </td>
                    </tr>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        Dropped
                        <br />
                        Late Enrollment
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        DRP
                        <br />
                        LE
                      </td>
                      <td className={borderCell}>
                        Reason and Effectivity Date
                        <br />
                        Reason (Enrollment beyond 1st Friday of SY)
                      </td>
                    </tr>
                  </tbody>
                </table>

                <table className="w-full table-fixed border-collapse">
                  <colgroup>
                    <col style={{ width: "25%" }} />
                    <col style={{ width: "9%" }} />
                    <col style={{ width: "66%" }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className={`${borderCell} text-left`}>Indicator</th>
                      <th className={headerCell}>Code</th>
                      <th className={`${borderCell} text-left`}>
                        Required Information
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        CCT Recipient
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        CCT
                      </td>
                      <td className={borderCell}>
                        CCT Control/reference number &amp; Effectivity Date
                      </td>
                    </tr>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        Balik Aral
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        B/A
                      </td>
                      <td className={borderCell}>
                        Name of school last attended &amp; Year
                      </td>
                    </tr>
                    <tr>
                      <td className={`${borderCell} font-semibold`}>
                        Special Needs Education
                        <br />
                        Accelerated
                      </td>
                      <td className={`${borderCell} text-center font-bold`}>
                        SNED
                        <br />
                        ACL
                      </td>
                      <td className={borderCell}>
                        Specify
                        <br />
                        Specify Level &amp; Effectivity Date
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            <table className="h-fit w-full table-fixed border-collapse text-center">
              <thead>
                <tr>
                  <th className={headerCell}>REGISTERED</th>
                  <th className={headerCell}>BoSY</th>
                  <th className={headerCell}>EoSY</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <th className={headerCell}>MALE</th>
                  <td className={borderCell}>
                    {registrationSummary.bosy.male}
                  </td>
                  <td className={borderCell}>
                    {registrationSummary.eosy.male}
                  </td>
                </tr>
                <tr>
                  <th className={headerCell}>FEMALE</th>
                  <td className={borderCell}>
                    {registrationSummary.bosy.female}
                  </td>
                  <td className={borderCell}>
                    {registrationSummary.eosy.female}
                  </td>
                </tr>
                <tr>
                  <th className={headerCell}>TOTAL</th>
                  <td className={borderCell}>
                    {registrationSummary.bosy.total}
                  </td>
                  <td className={borderCell}>
                    {registrationSummary.eosy.total}
                  </td>
                </tr>
              </tbody>
            </table>

            <div className="flex min-h-[74px] flex-col justify-between">
              <div className="font-bold">Prepared by:</div>
              <div className="text-center">
                <div className="border-b border-black px-1 pb-[2px] text-[5px]">
                  {adviser.toUpperCase()}
                </div>
                <div className="pt-[2px] font-bold">
                  (Signature of Adviser over Printed Name)
                </div>
              </div>
              <div className="grid w-[65%] grid-cols-2 gap-1 font-bold">
                <div className="border-b border-black pb-[2px]">
                  {sf1DateLine("BoSY Date:", bosyDate)}
                </div>
                <div className="border-b border-black pb-[2px]">
                  {sf1DateLine("EoSY Date:", eosyDate)}
                </div>
              </div>
            </div>

            <div className="flex min-h-[74px] flex-col justify-between">
              <div className="font-bold">Certified Correct:</div>
              <div className="text-center">
                <div className="grid w-[65%] border-b border-black px-1 pb-[2px] text-[5px]">
                  {schoolHead.toUpperCase()}
                </div>
                <div className=" grid w-[65%] pt-[2px] font-bold">
                  (Signature of School Head over Printed Name)
                </div>
              </div>
              <div className="grid w-[65%] grid-cols-2 gap-1 font-bold">
                <div className="border-b border-black pb-[2px]">
                  {sf1DateLine("BoSY Date:", bosyDate)}
                </div>
                <div className="border-b border-black pb-[2px]">
                  {sf1DateLine("EoSY Date:", eosyDate)}
                </div>
              </div>
              <div className=" grid w-[65%] border-t border-black pt-[2px] text-center font-bold">
                Generated thru LIS
              </div>
            </div>
          </div>

          <div className="mt-[4px] text-[5px]">
            Generated on: {generatedDate()}
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
