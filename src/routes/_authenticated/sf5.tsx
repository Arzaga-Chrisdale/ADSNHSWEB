import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useEffect, useMemo, useState } from "react";
import { ArrowLeft, GraduationCap, Users } from "lucide-react";
import {
  computeAverage,
  getUserId,
  type ClassRow,
  type GradeRow,
  type StudentRow,
} from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import { toast } from "sonner";
import JSZip from "jszip";
import ExcelJS, {
  type Alignment,
  type Borders,
  type Cell,
  type Fill,
  type Font,
  type PaperSize,
  type Worksheet,
} from "exceljs";
import {
  AlignmentType,
  BorderStyle,
  Document as WordDocument,
  HeightRule,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/sf5")({
  component: SF5Page,
});

type SF5ExcelRow = {
  lrn: string;
  name: string;
  average: number | null;
  action: string;
  failedAreas: string;
};

type SF5Profile = {
  region?: string | null;
  division?: string | null;
  district?: string | null;
  school_id?: string | null;
  school_year?: string | null;
  school_name?: string | null;
  full_name?: string | null;
  principal?: string | null;
};

type SF5ExcelOptions = {
  region: string;
  division: string;
  schoolId: string;
  schoolYear: string;
  curriculum: string;
  schoolName: string;
  gradeLevel: string;
  section: string;
  maleRows: SF5ExcelRow[];
  femaleRows: SF5ExcelRow[];
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

function getSF5Action(average: number | null): string {
  if (average == null) return "";

  const roundedAverage = Math.round(average);

  if (roundedAverage <= 74) return "RETAINED";
  if (roundedAverage <= 89) return "PROMOTED";
  return "PROMOTED WITH ACADEMIC EXCELLENCE AWARD";
}

const BLACK = { argb: "FF000000" };
const THIN_BORDER = { style: "thin" as const, color: BLACK };
const MEDIUM_BORDER = { style: "medium" as const, color: BLACK };
const NO_FILL: Fill = { type: "pattern", pattern: "none" };
const TOTAL_FILL: Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFC7C7FF" },
};
const EXCEL_FONT: Partial<Font> = { name: "Arial", size: 9 };
const CENTERED: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

// Column geometry copied from the supplied SF5 spreadsheet (A:Z).
const SF5_COLUMN_WIDTHS = [
  2.51, 10.91, 6.21, 10.07, 8.89, 9.22, 9.9, 6.04, 12.42, 1, 11.75, 5.03, 1,
  0.16, 5.53, 6.04, 3.18, 0.83, 4.52, 9.9, 3.18, 3.35, 1.33, 3.18, 0.16, 0.16,
] as const;

function sf5TableBorder({
  top = false,
  bottom = false,
  left = false,
  right = false,
}: {
  top?: boolean;
  bottom?: boolean;
  left?: boolean;
  right?: boolean;
} = {}): Partial<Borders> {
  return {
    top: top ? MEDIUM_BORDER : THIN_BORDER,
    bottom: bottom ? MEDIUM_BORDER : THIN_BORDER,
    left: left ? MEDIUM_BORDER : THIN_BORDER,
    right: right ? MEDIUM_BORDER : THIN_BORDER,
  };
}

function styleSF5Cell(
  cell: Cell,
  options: {
    font?: Partial<Font>;
    alignment?: Partial<Alignment>;
    border?: Partial<Borders>;
    fill?: Fill;
    numberFormat?: string;
  } = {},
) {
  cell.font = { ...EXCEL_FONT, ...options.font };
  cell.alignment = options.alignment || CENTERED;
  cell.border = options.border || {};
  cell.fill = options.fill || NO_FILL;
  if (options.numberFormat) cell.numFmt = options.numberFormat;
}

function mergeAndSetSF5(
  worksheet: Worksheet,
  range: string,
  value: string | number | null,
  options: Parameters<typeof styleSF5Cell>[1] = {},
) {
  const [firstCell, lastCell] = range.split(":");
  if (lastCell && firstCell !== lastCell) worksheet.mergeCells(range);
  const cell = worksheet.getCell(firstCell);
  cell.value = value === "" ? null : value;
  styleSF5Cell(cell, options);
}

function setSF5MainRow(
  worksheet: Worksheet,
  rowNumber: number,
  row: SF5ExcelRow,
  options: { top?: boolean; bottom?: boolean } = {},
) {
  const commonBorder = {
    top: options.top,
    bottom: options.bottom,
  };
  const numericLrn = /^\d{1,15}$/.test(row.lrn) ? Number(row.lrn) : null;

  mergeAndSetSF5(
    worksheet,
    `A${rowNumber}:B${rowNumber}`,
    numericLrn ?? row.lrn,
    {
      font: { name: "Arial", size: 8 },
      alignment: {
        horizontal: "right",
        vertical: "middle",
        wrapText: false,
        shrinkToFit: true,
      },
      border: sf5TableBorder({ ...commonBorder, left: true }),
      numberFormat: numericLrn == null ? "@" : "0",
    },
  );
  mergeAndSetSF5(worksheet, `C${rowNumber}:F${rowNumber}`, row.name, {
    font: { name: "Arial", size: 8 },
    alignment: {
      horizontal: "left",
      vertical: "middle",
      wrapText: true,
      shrinkToFit: true,
    },
    border: sf5TableBorder(commonBorder),
  });
  mergeAndSetSF5(
    worksheet,
    `G${rowNumber}`,
    row.average == null ? "" : row.average,
    {
      font: { name: "Arial", size: 8 },
      alignment: { horizontal: "right", vertical: "middle" },
      border: sf5TableBorder(commonBorder),
      numberFormat: "0",
    },
  );
  mergeAndSetSF5(worksheet, `H${rowNumber}:I${rowNumber}`, row.action, {
    font: { name: "Arial", size: 8 },
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
    border: sf5TableBorder(commonBorder),
  });
  mergeAndSetSF5(worksheet, `J${rowNumber}:L${rowNumber}`, row.failedAreas, {
    font: { name: "Arial", size: 8 },
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
    border: sf5TableBorder({ ...commonBorder, right: true }),
  });
  worksheet.getRow(rowNumber).height = 20;
}

function setSF5GroupRow(
  worksheet: Worksheet,
  rowNumber: number,
  label: string,
  options: { top?: boolean; bottom?: boolean } = {},
) {
  mergeAndSetSF5(worksheet, `A${rowNumber}:B${rowNumber}`, label, {
    font: { size: 9, bold: true },
    border: sf5TableBorder({ ...options, left: true }),
  });
  mergeAndSetSF5(worksheet, `C${rowNumber}:F${rowNumber}`, "", {
    border: sf5TableBorder(options),
  });
  mergeAndSetSF5(worksheet, `G${rowNumber}`, "", {
    border: sf5TableBorder(options),
  });
  mergeAndSetSF5(worksheet, `H${rowNumber}:I${rowNumber}`, "", {
    border: sf5TableBorder(options),
  });
  mergeAndSetSF5(worksheet, `J${rowNumber}:L${rowNumber}`, "", {
    border: sf5TableBorder({ ...options, right: true }),
  });
  worksheet.getRow(rowNumber).height = 20;
}

function setSF5TotalRow(
  worksheet: Worksheet,
  rowNumber: number,
  count: number,
  label: string,
  options: { bottom?: boolean } = {},
) {
  mergeAndSetSF5(worksheet, `A${rowNumber}:B${rowNumber}`, count, {
    font: { size: 9, bold: true },
    alignment: { horizontal: "right", vertical: "middle" },
    border: sf5TableBorder({ top: true, bottom: options.bottom, left: true }),
  });
  mergeAndSetSF5(worksheet, `C${rowNumber}:F${rowNumber}`, `<=== ${label}`, {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: sf5TableBorder({ top: true, bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `G${rowNumber}`, "", {
    fill: TOTAL_FILL,
    border: sf5TableBorder({ top: true, bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `H${rowNumber}:I${rowNumber}`, "", {
    fill: TOTAL_FILL,
    border: sf5TableBorder({ top: true, bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `J${rowNumber}:L${rowNumber}`, "", {
    fill: TOTAL_FILL,
    border: sf5TableBorder({
      top: true,
      bottom: options.bottom,
      right: true,
    }),
  });
  worksheet.getRow(rowNumber).height = 20;
}

function setSF5SummaryValue(
  worksheet: Worksheet,
  rowRange: string,
  label: string,
  male: number,
  female: number,
  options: { italic?: boolean; bottom?: boolean } = {},
) {
  const [startRow, endRow] = rowRange.split(":").map(Number);
  mergeAndSetSF5(worksheet, `N${startRow}:P${endRow}`, label, {
    font: {
      size: options.italic ? 6 : 8,
      bold: true,
      italic: options.italic,
    },
    border: sf5TableBorder({ left: true, bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `Q${startRow}:S${endRow}`, male, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `T${startRow}:T${endRow}`, female, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `U${startRow}:Z${endRow}`, male + female, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom, right: true }),
  });
}

function setSF5ProgressValue(
  worksheet: Worksheet,
  rowRange: string,
  label: string,
  male: number,
  female: number,
  options: { bold?: boolean; bottom?: boolean } = {},
) {
  const [startRow, endRow] = rowRange.split(":").map(Number);
  mergeAndSetSF5(worksheet, `N${startRow}:P${endRow}`, label, {
    font: { size: 7, bold: options.bold },
    border: sf5TableBorder({ left: true, bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `Q${startRow}:S${endRow}`, male, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `T${startRow}:T${endRow}`, female, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom }),
  });
  mergeAndSetSF5(worksheet, `U${startRow}:Z${endRow}`, male + female, {
    font: { size: 8 },
    numberFormat: "0",
    border: sf5TableBorder({ bottom: options.bottom, right: true }),
  });
}

async function buildSF5Workbook(options: SF5ExcelOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "School Form 5 (SF5)";
  workbook.subject =
    "Report on Promotion and Level of Proficiency & Achievement";

  const worksheet = workbook.addWorksheet("school_form_5", {
    views: [
      {
        state: "normal",
        showGridLines: true,
        zoomScale: 100,
        activeCell: "A1",
      },
    ],
    pageSetup: {
      orientation: "landscape",
      paperSize: 14 as PaperSize,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: false,
      verticalCentered: false,
      pageOrder: "overThenDown",
      blackAndWhite: false,
      margins: {
        left: 0.19,
        right: 0.19,
        top: 0.25,
        bottom: 0.19,
        header: 0.15,
        footer: 0.15,
      },
    },
  });

  worksheet.properties.defaultRowHeight = 12.75;
  SF5_COLUMN_WIDTHS.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  mergeAndSetSF5(
    worksheet,
    "A1:Z2",
    "School Form 5 (SF 5) Report on Promotion and Level of Proficiency & Achievement",
    {
      font: { size: 17, bold: true },
      border: {},
    },
  );
  mergeAndSetSF5(
    worksheet,
    "A3:Z3",
    "(This replaces Forms 18-E1, 18-E2, 18A and List of Graduates)",
    {
      font: { size: 7, italic: true },
      border: {},
    },
  );

  const metadata = [
    ["A4:D4", "Region", true],
    ["E4", options.region, false],
    ["F4", "Division", true],
    ["G4:J4", options.division, false],
    ["A5:D5", "School ID", true],
    ["E5:F5", options.schoolId, false],
    ["G5:H5", "School Year", true],
    ["I5:J5", options.schoolYear, false],
    ["K5", "Curriculum", true],
    ["L5:Q5", options.curriculum, false],
    ["A6:D7", "School Name", true],
    ["E6:J7", options.schoolName, false],
    ["K6:K7", "Grade Level", true],
    ["L6:O7", options.gradeLevel, false],
    ["P6:Q7", "Section", true],
    ["S6:Z7", options.section, false],
  ] as const;

  metadata.forEach(([range, value, isLabel]) => {
    const isSchoolName = range === "E6:J7";
    mergeAndSetSF5(worksheet, range, value, {
      font: { size: 9 },
      alignment: {
        horizontal: isLabel ? "right" : "center",
        vertical: "middle",
        wrapText: !isSchoolName,
        shrinkToFit: isSchoolName,
      },
      border: isLabel ? {} : sf5TableBorder(),
    });
  });

  mergeAndSetSF5(worksheet, "A8:B11", "LRN", {
    font: { size: 9, bold: true },
    border: sf5TableBorder({ top: true, left: true }),
  });
  mergeAndSetSF5(
    worksheet,
    "C8:F11",
    "LEARNER'S NAME\n(Last Name, First Name, Middle Name)",
    {
      font: { size: 9, bold: true },
      border: sf5TableBorder({ top: true }),
    },
  );
  mergeAndSetSF5(worksheet, "G8:G11", "GENERAL\nAVERAGE", {
    font: { size: 9, bold: true },
    border: sf5TableBorder({ top: true }),
  });
  mergeAndSetSF5(
    worksheet,
    "H8:I11",
    "ACTION TAKEN:\nRETAINED, PROMOTED\nor PROMOTED WITH ACADEMIC EXCELLENCE AWARD",
    {
      font: { size: 8, bold: true },
      border: sf5TableBorder({ top: true }),
    },
  );
  mergeAndSetSF5(
    worksheet,
    "J8:L11",
    "Did Not Meet Expectations of the ff.\nLearning Area/s as of end of current\nSchool Year",
    {
      font: { size: 8 },
      border: sf5TableBorder({ top: true, right: true }),
    },
  );

  mergeAndSetSF5(worksheet, "N8:Z9", "SUMMARY TABLE", {
    font: { size: 8, bold: true },
    border: sf5TableBorder({ top: true, left: true, right: true }),
  });
  mergeAndSetSF5(worksheet, "N10:P10", "STATUS", {
    font: { size: 8, bold: true },
    border: sf5TableBorder({ left: true }),
  });
  mergeAndSetSF5(worksheet, "Q10:S10", "MALE", {
    font: { size: 8, bold: true },
    border: sf5TableBorder(),
  });
  mergeAndSetSF5(worksheet, "T10", "FEMALE", {
    font: { size: 8, bold: true },
    border: sf5TableBorder(),
  });
  mergeAndSetSF5(worksheet, "U10:Z10", "TOTAL", {
    font: { size: 8, bold: true },
    border: sf5TableBorder({ right: true }),
  });
  setSF5SummaryValue(
    worksheet,
    "11:12",
    "PROMOTED",
    options.summary.promotedMale,
    options.summary.promotedFemale,
  );
  setSF5SummaryValue(
    worksheet,
    "13:14",
    "PROMOTED\nWITH ACADEMIC EXCELLENCE AWARD",
    options.summary.conditionalMale,
    options.summary.conditionalFemale,
    { italic: true },
  );
  setSF5SummaryValue(
    worksheet,
    "15:17",
    "RETAINED",
    options.summary.retainedMale,
    options.summary.retainedFemale,
    { bottom: true },
  );

  mergeAndSetSF5(worksheet, "N19:Z20", "LEVEL OF PROGRESS AND ACHIEVEMENT", {
    font: { size: 8, bold: true },
    border: sf5TableBorder({ top: true, left: true, right: true }),
  });
  mergeAndSetSF5(worksheet, "N21:P23", "Descriptor &\nGrading", {
    font: { size: 7, bold: true },
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
    border: sf5TableBorder({ left: true }),
  });
  mergeAndSetSF5(worksheet, "Q21:S23", "MALE", {
    font: { size: 8, bold: true },
    border: sf5TableBorder(),
  });
  mergeAndSetSF5(worksheet, "T21:T23", "FEMALE", {
    font: { size: 8, bold: true },
    border: sf5TableBorder(),
  });
  mergeAndSetSF5(worksheet, "U21:Z23", "TOTAL", {
    font: { size: 8, bold: true },
    border: sf5TableBorder({ right: true }),
  });

  const progressRanges = ["24:26", "27:29", "30:32", "33:35", "36:39"];
  options.progress.forEach((progress, index) => {
    setSF5ProgressValue(
      worksheet,
      progressRanges[index],
      progress.label,
      progress.male,
      progress.female,
      {
        bold: index > 0,
        bottom: index === options.progress.length - 1,
      },
    );
  });

  mergeAndSetSF5(worksheet, "N42:Z42", "Instructions:", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  mergeAndSetSF5(
    worksheet,
    "N43:Z48",
    [
      "1. The SCC shall conduct checking in their own school, no swapping of SCC from one school to another is permitted.",
      "2. The name of SCC members shall be printed and put their signature on top (additional space may be added).",
      "3. The school head is accountable and liable for any wrongful entry on the forms; therefore, the DCC is not required to put their names and signatures in SF 5.",
      "4. Only LIS generated SF5 shall be recognized.",
      "5. This form shall be submitted to the DCC together with the accomplished SFCR1.",
    ].join("\n"),
    {
      font: { size: 7 },
      alignment: {
        horizontal: "left",
        vertical: "top",
        wrapText: true,
      },
      border: {},
    },
  );

  mergeAndSetSF5(worksheet, "N51:Z52", "PREPARED BY:", {
    font: { size: 9 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  mergeAndSetSF5(worksheet, "N55:Z56", options.adviser.toUpperCase(), {
    font: { size: 9 },
    alignment: { horizontal: "center", vertical: "bottom", shrinkToFit: true },
    border: { bottom: THIN_BORDER },
  });
  mergeAndSetSF5(worksheet, "N57:Z57", "Class Adviser", {
    font: { size: 8, italic: true },
    border: {},
  });
  mergeAndSetSF5(worksheet, "N58:Z58", "(Name and Signature)", {
    font: { size: 8, italic: true },
    border: {},
  });

  mergeAndSetSF5(worksheet, "N60:Z61", "CERTIFIED CORRECT & SUBMITTED BY:", {
    font: { size: 9 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  mergeAndSetSF5(worksheet, "N64:Z65", options.schoolHead.toUpperCase(), {
    font: { size: 9 },
    alignment: { horizontal: "center", vertical: "bottom", shrinkToFit: true },
    border: { bottom: THIN_BORDER },
  });
  mergeAndSetSF5(worksheet, "N66:Z66", "School Head & SCC Chair", {
    font: { size: 8, italic: true },
    border: {},
  });
  mergeAndSetSF5(worksheet, "N67:Z67", "(Name and Signature)", {
    font: { size: 8, italic: true },
    border: {},
  });

  mergeAndSetSF5(worksheet, "N69:Z70", "REVIEWED BY: SCC Members", {
    font: { size: 9 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  [
    ["N74:Z75", "N76:Z76"],
    ["N79:Z80", "N81:Z81"],
    ["N84:Z85", "N86:Z86"],
  ].forEach(([lineRange, labelRange]) => {
    mergeAndSetSF5(worksheet, lineRange, "", {
      border: { bottom: THIN_BORDER },
    });
    mergeAndSetSF5(worksheet, labelRange, "(Signature Over Printed Name)", {
      font: { size: 7, italic: true },
      border: {},
    });
  });
  mergeAndSetSF5(worksheet, "N88:Z89", "Generated thru LIS (SCC CO-Chair)", {
    font: { size: 7 },
    border: { top: THIN_BORDER },
  });

  setSF5GroupRow(worksheet, 12, "MALE", { top: true });
  let currentRow = 13;
  options.maleRows.forEach((row) => {
    setSF5MainRow(worksheet, currentRow, row);
    currentRow += 1;
  });
  setSF5TotalRow(worksheet, currentRow, options.maleRows.length, "TOTAL MALE");
  currentRow += 1;

  setSF5GroupRow(worksheet, currentRow, "FEMALE");
  currentRow += 1;
  options.femaleRows.forEach((row) => {
    setSF5MainRow(worksheet, currentRow, row);
    currentRow += 1;
  });
  setSF5TotalRow(
    worksheet,
    currentRow,
    options.femaleRows.length,
    "TOTAL FEMALE",
  );
  currentRow += 1;
  setSF5TotalRow(
    worksheet,
    currentRow,
    options.maleRows.length + options.femaleRows.length,
    "COMBINED",
    { bottom: true },
  );

  [
    15, 18, 21, 21, 21, 13, 8, 25, 5, 30, 23, 20, 13, 13, 13, 13, 13, 6, 11, 11,
    19, 1, 9, 11, 20, 13, 7, 20, 13, 7, 24, 19, 1, 24, 15, 5, 24, 20, 1, 21, 2,
    14, 20, 20, 20, 20, 20, 22, 2, 21, 20, 20, 12, 9, 20, 20, 9, 10, 15, 17, 17,
    17, 17, 20, 20, 15, 15, 10, 20, 20, 10, 10, 21, 20, 15, 20, 20, 15, 20, 20,
    15, 20, 20, 15, 20, 10, 20, 20, 10,
  ].forEach((height, index) => {
    if (worksheet.getRow(index + 1).height == null) {
      worksheet.getRow(index + 1).height = height;
    }
  });

  const finalRow = Math.max(currentRow, 89);
  worksheet.pageSetup.printArea = `A1:Z${finalRow}`;

  // Enforce Arial only inside the exported workbook. The website preview is unchanged.
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.value !== null) {
        cell.font = { ...cell.font, name: "Arial" };
      }
    });
  });

  return workbook;
}

const SF5_EXCEL_TEMPLATE_URL = "/templates/SF5-class-adviser.xlsx";

const SF5_MAX_MALE_LEARNERS = 19;
const SF5_MAX_FEMALE_LEARNERS = 20;

// The supplied template uses merged blocks of different heights so its learner
// table can align with the summary and signature sections on the right. The
// export reuses these blocks from top to bottom, then removes the unused blocks
// in A:L. This makes TOTAL MALE, FEMALE, TOTAL FEMALE, and COMBINED follow the
// last real learner without deleting any rows used by the form on the right.
const SF5_TEMPLATE_BLOCK_START_ROWS = [
  14, 16, 17, 20, 22, 23, 25, 26, 28, 29, 31, 32, 34, 35, 37, 38, 39, 42, 44,
  45, 46, 47, 48, 50, 52, 54, 56, 59, 61, 64, 65, 68, 71, 73, 75, 77, 79, 81,
  83, 85, 87, 89, 92,
] as const;

const SF5_TEMPLATE_BLOCKS = SF5_TEMPLATE_BLOCK_START_ROWS.map(
  (start, index) => ({
    start,
    end: (SF5_TEMPLATE_BLOCK_START_ROWS[index + 1] ?? 93) - 1,
  }),
);

const SF5_LEFT_COLUMNS = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
] as const;

const SPREADSHEET_XML_NAMESPACE =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

function parseSF5TemplateXml(xml: string, partName: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`The SF5 Excel template has invalid ${partName}.`);
  }
  return document;
}

function getSF5TemplateCell(document: XMLDocument, address: string) {
  const cell = Array.from(
    document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c"),
  ).find((candidate) => candidate.getAttribute("r") === address);

  if (!cell) {
    throw new Error(`The SF5 Excel template is missing cell ${address}.`);
  }
  return cell;
}

function setSF5TemplateCell(
  document: XMLDocument,
  address: string,
  value: string | number | null,
) {
  const cell = getSF5TemplateCell(document, address);

  // Remove only the old value/formula nodes. The cell's style attribute stays
  // intact, so the supplied font, border, fill, and alignment are preserved.
  Array.from(cell.children).forEach((child) => {
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
  const inlineString = document.createElementNS(
    SPREADSHEET_XML_NAMESPACE,
    "is",
  );
  const textNode = document.createElementNS(SPREADSHEET_XML_NAMESPACE, "t");
  textNode.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
  textNode.textContent = value;
  inlineString.appendChild(textNode);
  cell.appendChild(inlineString);
}

function setSF5TemplateLearner(
  document: XMLDocument,
  rowNumber: number,
  learner: SF5ExcelRow,
) {
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

function clearSF5TemplateBlockValues(
  document: XMLDocument,
  block: { start: number; end: number },
) {
  Array.from(
    document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c"),
  ).forEach((cell) => {
    const address = cell.getAttribute("r") ?? "";
    const rowNumber = Number(address.match(/\d+$/)?.[0] ?? 0);
    if (
      rowNumber < block.start ||
      rowNumber > block.end ||
      getSF5CellColumnNumber(address) > 12
    ) {
      return;
    }

    Array.from(cell.children).forEach((child) => {
      if (["v", "is", "f"].includes(child.localName)) child.remove();
    });
    cell.removeAttribute("t");
  });
}

function removeSF5TemplateBlock(
  document: XMLDocument,
  block: { start: number; end: number },
) {
  const mergeCells = document.getElementsByTagNameNS(
    SPREADSHEET_XML_NAMESPACE,
    "mergeCells",
  )[0];

  if (mergeCells) {
    Array.from(mergeCells.children).forEach((mergeCell) => {
      const reference = mergeCell.getAttribute("ref") ?? "";
      const [from = "", to = from] = reference.split(":");
      const fromRow = Number(from.match(/\d+$/)?.[0] ?? 0);
      const toRow = Number(to.match(/\d+$/)?.[0] ?? 0);
      const fromColumn = getSF5CellColumnNumber(from);
      const toColumn = getSF5CellColumnNumber(to);

      if (
        fromColumn >= 1 &&
        toColumn <= 12 &&
        fromRow >= block.start &&
        toRow <= block.end
      ) {
        mergeCell.remove();
      }
    });
    mergeCells.setAttribute("count", String(mergeCells.children.length));
  }

  Array.from(
    document.getElementsByTagNameNS(SPREADSHEET_XML_NAMESPACE, "c"),
  ).forEach((cell) => {
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

function compactSF5TemplateLearnerTable(
  document: XMLDocument,
  options: SF5ExcelOptions,
) {
  const styles = {
    learner: getSF5TemplateStyleProfile(document, 14),
    totalMale: getSF5TemplateStyleProfile(document, 45),
    femaleHeader: getSF5TemplateStyleProfile(document, 46),
    totalFemale: getSF5TemplateStyleProfile(document, 89),
    combined: getSF5TemplateStyleProfile(document, 92),
  };
  const entries: Array<
    | { kind: "learner"; learner: SF5ExcelRow }
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
    if (!block)
      throw new Error("The SF5 template does not have enough learner rows.");

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

async function buildSF5TemplateBuffer(options: SF5ExcelOptions) {
  if (options.maleRows.length > SF5_MAX_MALE_LEARNERS) {
    throw new Error(
      `The SF5 template supports up to ${SF5_MAX_MALE_LEARNERS} male learners.`,
    );
  }
  if (options.femaleRows.length > SF5_MAX_FEMALE_LEARNERS) {
    throw new Error(
      `The SF5 template supports up to ${SF5_MAX_FEMALE_LEARNERS} female learners.`,
    );
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
  const worksheet = parseSF5TemplateXml(
    await worksheetPart.async("string"),
    "worksheet XML",
  );

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

  zip.file(
    "xl/worksheets/sheet1.xml",
    new XMLSerializer().serializeToString(worksheet),
  );
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

async function downloadSF5Excel(fileName: string, options: SF5ExcelOptions) {
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
const WORDPROCESSING_XML_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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
  return Array.from(parent.children).filter(
    (child) =>
      child.namespaceURI === WORDPROCESSING_XML_NAMESPACE &&
      child.localName === localName,
  );
}

function getDirectWordRows(table: Element) {
  return getDirectWordChildren(table, "tr");
}

function getDirectWordCells(row: Element) {
  return getDirectWordChildren(row, "tc");
}

function getWordNodeText(node: Element) {
  return Array.from(
    node.getElementsByTagNameNS(WORDPROCESSING_XML_NAMESPACE, "t"),
  )
    .map((textNode) => textNode.textContent || "")
    .join("")
    .trim();
}

function setSF5WordParagraphText(
  paragraph: Element,
  value: string | number | null,
) {
  Array.from(paragraph.children).forEach((child) => {
    if (child.localName !== "pPr") child.remove();
  });

  const run = paragraph.ownerDocument.createElementNS(
    WORDPROCESSING_XML_NAMESPACE,
    "w:r",
  );
  const lines = String(value ?? "").split("\n");

  lines.forEach((line, index) => {
    if (index > 0) {
      run.appendChild(
        paragraph.ownerDocument.createElementNS(
          WORDPROCESSING_XML_NAMESPACE,
          "w:br",
        ),
      );
    }
    const text = paragraph.ownerDocument.createElementNS(
      WORDPROCESSING_XML_NAMESPACE,
      "w:t",
    );
    text.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
    text.textContent = line;
    run.appendChild(text);
  });

  paragraph.appendChild(run);
}

function setSF5WordCellText(cell: Element, value: string | number | null) {
  let paragraph = getDirectWordChildren(cell, "p")[0];
  if (!paragraph) {
    paragraph = cell.ownerDocument.createElementNS(
      WORDPROCESSING_XML_NAMESPACE,
      "w:p",
    );
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

function populateSF5WordLearnerRow(row: Element, learner: SF5ExcelRow) {
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
  if (!row)
    throw new Error(`The SF5 Word template is missing the ${text} row.`);
  return row;
}

function populateSF5WordLearnerTable(
  table: Element,
  maleRows: SF5ExcelRow[],
  femaleRows: SF5ExcelRow[],
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
  const femaleTemplate = rows[femaleHeaderIndex + 1]?.cloneNode(
    true,
  ) as Element;

  if (!maleTemplate || !femaleTemplate) {
    throw new Error("The SF5 Word template has no reusable learner row.");
  }

  rows
    .slice(maleHeaderIndex + 1, totalMaleIndex)
    .forEach((row) => row.remove());
  rows
    .slice(femaleHeaderIndex + 1, totalFemaleIndex)
    .forEach((row) => row.remove());

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
  setSF5WordCellText(
    getDirectWordCells(combined)[0],
    maleRows.length + femaleRows.length,
  );
}

function populateSF5WordCountRow(
  table: Element,
  rowIndex: number,
  male: number,
  female: number,
) {
  setSF5WordTableCell(table, rowIndex, 1, male);
  setSF5WordTableCell(table, rowIndex, 2, female);
  setSF5WordTableCell(table, rowIndex, 3, male + female);
}

function countSF5WordRows(
  rows: SF5ExcelRow[],
  lower: number,
  upper: number | null,
) {
  return rows.filter((row) => {
    if (row.average == null) return false;
    const average = Math.round(row.average);
    return average >= lower && (upper == null || average <= upper);
  }).length;
}

function populateSF5WordSignatures(
  rightCell: Element,
  adviser: string,
  schoolHead: string,
) {
  const paragraphs = getDirectWordChildren(rightCell, "p");
  const setAfterLabel = (label: string, value: string) => {
    const index = paragraphs.findIndex(
      (paragraph) => getWordNodeText(paragraph) === label,
    );
    const target = index >= 0 ? paragraphs[index + 1] : undefined;
    if (!target) {
      throw new Error(`The SF5 Word template is missing the ${label} field.`);
    }
    setSF5WordParagraphText(target, value.toUpperCase());
  };

  setAfterLabel("PREPARED BY:", adviser);
  setAfterLabel("CERTIFIED CORRECT & SUBMITTED BY:", schoolHead);
}

async function buildSF5WordTemplateBuffer(options: SF5ExcelOptions) {
  const response = await fetch(SF5_WORD_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(
      `SF5 Word template not found (${response.status}). Put SF5-class-adviser.docx in public/templates/.`,
    );
  }

  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentPart = zip.file("word/document.xml");
  if (!documentPart) {
    throw new Error(
      "The SF5 Word template does not contain word/document.xml.",
    );
  }

  const documentXml = parseSF5WordTemplateXml(
    await documentPart.async("string"),
  );
  const tables = Array.from(
    documentXml.getElementsByTagNameNS(WORDPROCESSING_XML_NAMESPACE, "tbl"),
  );
  const [metadataTable, outerTable, learnerTable, summaryTable, progressTable] =
    tables;

  if (
    !metadataTable ||
    !outerTable ||
    !learnerTable ||
    !summaryTable ||
    !progressTable
  ) {
    throw new Error(
      "The SF5 Word template does not contain the required tables.",
    );
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

  populateSF5WordLearnerTable(
    learnerTable,
    options.maleRows,
    options.femaleRows,
  );

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
    populateSF5WordCountRow(
      progressTable,
      index + 2,
      progress.male,
      progress.female,
    );
  });

  const outerRows = getDirectWordRows(outerTable);
  const rightCell = outerRows[0]
    ? getDirectWordCells(outerRows[0])[1]
    : undefined;
  if (!rightCell) {
    throw new Error("The SF5 Word template is missing its signature section.");
  }
  populateSF5WordSignatures(rightCell, options.adviser, options.schoolHead);

  zip.file(
    "word/document.xml",
    new XMLSerializer().serializeToString(documentXml),
  );
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

async function downloadSF5Word(fileName: string, options: SF5ExcelOptions) {
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

function formatGradeLevel(value?: string | null) {
  const trimmed = value?.trim() || "";
  if (!trimmed) return "";
  return /^grade\b/i.test(trimmed) ? trimmed : `Grade ${trimmed}`;
}

function SF5Page() {
  const [classId] = useState<string>(readSchoolFormsClassId);
  const [length, setLength] = useState<"short" | "full">("full");
  const [summarySubjects, setSummarySubjects] = useState<string[]>([]);

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () => {
      const id = await getUserId();
      const { data, error } = await supabase
        .from("profiles")
        .select("*")
        .eq("id", id)
        .maybeSingle();

      if (error) throw error;
      return data as SF5Profile | null;
    },
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
  const active = selectedSchoolFormsClass?.id || classes[0]?.id;
  const klass = classes.find((c) => c.id === active);
  const schoolHeadName = profile?.principal?.trim() || "";

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

  const { data: students = [] } = useQuery({
    enabled: !!active,
    queryKey: ["students", active],
    queryFn: async () =>
      (
        await supabase
          .from("students")
          .select("*")
          .eq("class_id", active!)
          .order("last_name")
      ).data as StudentRow[],
  });
  const { data: grades = [] } = useQuery({
    enabled: !!active,
    queryKey: ["grades-all", active],
    queryFn: async () =>
      (await supabase.from("grades").select("*").eq("class_id", active!))
        .data as GradeRow[],
  });

  const rows = useMemo(
    () =>
      students.map((st) => {
        const finalSubjectGrades = summarySubjects.map((subjectName) => {
          const storedFinal = grades.find(
            (grade) =>
              grade.student_id === st.id &&
              grade.subject === subjectName &&
              grade.term === "final",
          )?.score;

          // Match the FINAL-GRADES-AND-GENERAL-AVERAGE export: use the saved
          // Final grade first and average Terms 1-3 only as a fallback.
          if (
            typeof storedFinal === "number" &&
            !Number.isNaN(storedFinal)
          ) {
            return storedFinal;
          }

          return computeAverage(
            ["1", "2", "3"].map(
              (termValue) =>
                grades.find(
                  (grade) =>
                    grade.student_id === st.id &&
                    grade.subject === subjectName &&
                    grade.term === termValue,
                )?.score ?? null,
            ),
          );
        });

        const avg = computeAverage(finalSubjectGrades);
        const fail = finalSubjectGrades
          .map((grade, index) =>
            grade != null && Math.round(grade) < 75
              ? summarySubjects[index]
              : null,
          )
          .filter(Boolean)
          .join(", ");
        return { st, avg, fail };
      }),
    [students, grades, summarySubjects],
  );

  const male = rows.filter((r) => r.st.sex === "male");
  const female = rows.filter((r) => r.st.sex === "female");
  const summary = (list: typeof rows) => ({
    promoted: list.filter(
      (r) => r.avg != null && Math.round(r.avg) >= 75,
    ).length,
    retained: list.filter(
      (r) => r.avg != null && Math.round(r.avg) < 75,
    ).length,
    conditional: 0,
  });
  const sm = summary(male),
    sf = summary(female);
  const bandCount = (list: typeof rows, lo: number, hi: number | null) =>
    list.filter(
      (r) =>
        r.avg != null &&
        Math.round(r.avg) >= lo &&
        (hi == null || Math.round(r.avg) <= hi),
    ).length;
  const bands = [
    { label: "Outstanding (90-100)", lo: 90, hi: null as number | null },
    { label: "Very Satisfactory (85-89)", lo: 85, hi: 89 },
    { label: "Satisfactory (80-84)", lo: 80, hi: 84 },
    { label: "Fairly Satisfactory (75-79)", lo: 75, hi: 79 },
    { label: "Did Not Meet Expectations (74 and below)", lo: 0, hi: 74 },
  ];

  const copySF5ToWord = async () => {
    if (!klass) {
      toast.error("Please select a class before creating the Word document");
      return;
    }

    try {
      const toWordRow = (row: (typeof rows)[number]): SF5ExcelRow => ({
        lrn: row.st.lrn || "",
        name: [row.st.last_name, row.st.first_name, row.st.middle_name]
          .filter(Boolean)
          .join(", ")
          .toUpperCase(),
        average: row.avg == null ? null : Math.round(row.avg),
        action: getSF5Action(row.avg),
        failedAreas: row.fail.toUpperCase(),
      });

      const usedWordTemplate = await downloadSF5Word(
        `SF5_Promotion_${klass.section || "class"}`,
        {
          region: profile?.region || klass.region || "",
          division: profile?.division || klass.division || "",
          schoolId: profile?.school_id || klass.school_id || "",
          schoolYear: klass.school_year || profile?.school_year || "",
          curriculum: "K to 12",
          schoolName: profile?.school_name || klass.school_name || "",
          gradeLevel: formatGradeLevel(klass.grade_level),
          section: klass.section || "",
          maleRows: male.map(toWordRow),
          femaleRows: female.map(toWordRow),
          summary: {
            promotedMale: sm.promoted,
            promotedFemale: sf.promoted,
            conditionalMale: sm.conditional,
            conditionalFemale: sf.conditional,
            retainedMale: sm.retained,
            retainedFemale: sf.retained,
          },
          progress: [...bands].reverse().map((band) => ({
            label: band.label,
            male: bandCount(male, band.lo, band.hi),
            female: bandCount(female, band.lo, band.hi),
          })),
          adviser: klass.teacher_name || profile?.full_name || "",
          schoolHead: schoolHeadName,
        },
      );

      if (usedWordTemplate) {
        toast.success("SF5 Word document downloaded from your template");
        return;
      }

      const inchesToTwip = (inches: number) => Math.round(inches * 1440);

      // Use portrait dimensions here. PageOrientation.LANDSCAPE swaps them
      // internally so the exported page becomes Long bond paper landscape.
      const pageWidth = inchesToTwip(8.5);
      const pageHeight = inchesToTwip(13);
      const margin = inchesToTwip(0.25);
      const contentWidth = pageHeight - margin * 2;

      const leftWidth = Math.floor(contentWidth * 0.69);
      const rightWidth = contentWidth - leftWidth;
      const wordSansFont = "SansSerif";
      const wordMonospacedFont = "Monospaced";

      const noBorder = {
        style: BorderStyle.NONE,
        size: 0,
        color: "FFFFFF",
      };
      const thinBorder = {
        style: BorderStyle.SINGLE,
        size: 4,
        color: "000000",
        space: 0,
      };
      const mediumBorder = {
        style: BorderStyle.SINGLE,
        size: 8,
        color: "000000",
        space: 0,
      };
      const allThinBorders = {
        top: thinBorder,
        bottom: thinBorder,
        left: thinBorder,
        right: thinBorder,
      };
      // Apply borders at table level as well as at cell level. Word can
      // otherwise collapse a shared cell edge and hide a vertical column
      // line, especially in nested tables or at reduced zoom levels.
      const fullGridBorders = {
        top: mediumBorder,
        bottom: mediumBorder,
        left: mediumBorder,
        right: mediumBorder,
        insideHorizontal: mediumBorder,
        insideVertical: mediumBorder,
      };
      const allMediumBorders = {
        top: mediumBorder,
        bottom: mediumBorder,
        left: mediumBorder,
        right: mediumBorder,
      };
      const noBorders = {
        top: noBorder,
        bottom: noBorder,
        left: noBorder,
        right: noBorder,
        insideHorizontal: noBorder,
        insideVertical: noBorder,
      };

      const paragraph = (
        text: string,
        options: {
          bold?: boolean;
          italic?: boolean;
          size?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          before?: number;
          after?: number;
          font?: string;
        } = {},
      ) =>
        new Paragraph({
          alignment: options.alignment ?? AlignmentType.CENTER,
          spacing: {
            before: options.before ?? 0,
            after: options.after ?? 0,
            line: 190,
          },
          children: text.split("\n").flatMap((line, index) => [
            new TextRun({
              text: line,
              break: index === 0 ? undefined : 1,
              bold: options.bold,
              italics: options.italic,
              // docx uses half-points. Keep the numeric options below in
              // normal point sizes so they match the supplied spreadsheet.
              size: (options.size ?? 7) * 2,
              font: options.font ?? wordSansFont,
              color: "000000",
            }),
          ]),
        });

      const cell = ({
        text,
        width,
        bold = false,
        italic = false,
        size = 7,
        font = wordSansFont,
        alignment = AlignmentType.CENTER,
        fill,
        columnSpan,
        rowSpan,
        borders = allMediumBorders,
        marginX = 35,
        marginY = 25,
      }: {
        text: string;
        width: number;
        bold?: boolean;
        italic?: boolean;
        size?: number;
        font?: string;
        alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
        fill?: string;
        columnSpan?: number;
        rowSpan?: number;
        borders?:
          typeof allThinBorders | typeof allMediumBorders | typeof noBorders;
        marginX?: number;
        marginY?: number;
      }) =>
        new TableCell({
          width: { size: width, type: WidthType.DXA },
          columnSpan,
          rowSpan,
          verticalAlign: VerticalAlign.CENTER,
          borders,
          shading: fill ? { fill } : undefined,
          margins: {
            top: marginY,
            bottom: marginY,
            left: marginX,
            right: marginX,
          },
          children: [
            paragraph(text, {
              bold,
              italic,
              size,
              font,
              alignment,
            }),
          ],
        });

      const metadataRow = (
        items: Array<{ text: string; width: number; value?: boolean }>,
      ) =>
        new Table({
          width: { size: contentWidth, type: WidthType.DXA },
          columnWidths: items.map((item) => item.width),
          layout: TableLayoutType.FIXED,
          borders: noBorders,
          rows: [
            new TableRow({
              cantSplit: true,
              children: items.map((item) =>
                cell({
                  text: item.text,
                  width: item.width,
                  size: 8,
                  alignment: item.value
                    ? AlignmentType.CENTER
                    : AlignmentType.RIGHT,
                  borders: item.value ? allThinBorders : noBorders,
                  marginX: item.value ? 30 : 45,
                  marginY: 12,
                }),
              ),
            }),
          ],
        });

      // Each metadata line uses its own column geometry because the original
      // SF5 spreadsheet does not align every label/value pair to one grid.
      const metadataTables = [
        metadataRow([
          { text: "Region", width: 3600 },
          {
            text: profile?.region || klass.region || "",
            width: 1100,
            value: true,
          },
          { text: "Division", width: 1200 },
          {
            text: profile?.division || klass.division || "",
            width: 4100,
            value: true,
          },
          { text: "", width: contentWidth - 10000 },
        ]),
        metadataRow([
          { text: "School ID", width: 3600 },
          {
            text: profile?.school_id || klass.school_id || "",
            width: 2300,
            value: true,
          },
          { text: "School Year", width: 1900 },
          {
            text: klass.school_year || profile?.school_year || "",
            width: 1700,
            value: true,
          },
          { text: "Curriculum", width: 1800 },
          { text: "K to 12", width: contentWidth - 11300, value: true },
        ]),
        metadataRow([
          { text: "School Name", width: 3600 },
          {
            text: profile?.school_name || klass.school_name || "",
            width: 6200,
            value: true,
          },
          { text: "Grade Level", width: 1700 },
          {
            text: formatGradeLevel(klass.grade_level),
            width: 2300,
            value: true,
          },
          { text: "Section", width: 1100 },
          {
            text: klass.section || "",
            width: contentWidth - 14900,
            value: true,
          },
        ]),
      ];

      const leftColumnWidths = [
        Math.floor(leftWidth * 0.143),
        Math.floor(leftWidth * 0.366),
        Math.floor(leftWidth * 0.105),
        Math.floor(leftWidth * 0.196),
      ];
      leftColumnWidths.push(
        leftWidth - leftColumnWidths.reduce((sum, width) => sum + width, 0),
      );

      const learnerHeader = new TableRow({
        tableHeader: true,
        cantSplit: true,
        height: { value: 820, rule: HeightRule.ATLEAST },
        children: [
          cell({
            text: "LRN",
            width: leftColumnWidths[0],
            bold: true,
            size: 9,
            borders: allMediumBorders,
          }),
          cell({
            text: "LEARNER'S NAME\n(Last Name, First Name, Middle Name)",
            width: leftColumnWidths[1],
            bold: true,
            size: 8,
            borders: allMediumBorders,
          }),
          cell({
            text: "GENERAL\nAVERAGE",
            width: leftColumnWidths[2],
            bold: true,
            size: 8,
            borders: allMediumBorders,
          }),
          cell({
            text: "ACTION TAKEN:\nRETAINED, PROMOTED\nor PROMOTED WITH ACADEMIC EXCELLENCE AWARD",
            width: leftColumnWidths[3],
            bold: true,
            size: 8,
            borders: allMediumBorders,
          }),
          cell({
            text: "Did Not Meet Expectations of the ff.\nLearning Area/s as of end of current\nSchool Year",
            width: leftColumnWidths[4],
            size: 7,
            borders: allMediumBorders,
          }),
        ],
      });

      const groupRow = (label: string) =>
        new TableRow({
          cantSplit: true,
          children: leftColumnWidths.map((width, index) =>
            cell({
              text: index === 0 ? label : "",
              width,
              bold: index === 0,
              size: 8,
              alignment: AlignmentType.LEFT,
            }),
          ),
        });

      const learnerRow = (row: (typeof rows)[number]) =>
        new TableRow({
          cantSplit: true,
          height: { value: 245, rule: HeightRule.ATLEAST },
          children: [
            cell({
              text: row.st.lrn || "",
              width: leftColumnWidths[0],
              size: 9,
              font: wordMonospacedFont,
              alignment: AlignmentType.RIGHT,
            }),
            cell({
              text: [row.st.last_name, row.st.first_name, row.st.middle_name]
                .filter(Boolean)
                .join(", ")
                .toUpperCase(),
              width: leftColumnWidths[1],
              size: 7,
              alignment: AlignmentType.LEFT,
            }),
            cell({
              text: row.avg == null ? "" : String(Math.round(row.avg)),
              width: leftColumnWidths[2],
              size: 9,
              font: wordMonospacedFont,
              alignment: AlignmentType.RIGHT,
            }),
            cell({
              text: getSF5Action(row.avg),
              width: leftColumnWidths[3],
              size: 7,
              alignment: AlignmentType.LEFT,
            }),
            cell({
              text: row.fail.toUpperCase(),
              width: leftColumnWidths[4],
              size: 7,
              alignment: AlignmentType.LEFT,
            }),
          ],
        });

      const totalRow = (count: number, label: string) =>
        new TableRow({
          cantSplit: true,
          children: [
            cell({
              text: String(count),
              width: leftColumnWidths[0],
              bold: true,
              size: 8,
              alignment: AlignmentType.RIGHT,
              borders: allMediumBorders,
            }),
            cell({
              text: `<=== ${label}`,
              width: leftColumnWidths[1],
              bold: true,
              size: 8,
              alignment: AlignmentType.LEFT,
              borders: allMediumBorders,
            }),
            ...leftColumnWidths.slice(2).map((width) =>
              cell({
                text: "",
                width,
                fill: "C7C7FF",
                borders: allMediumBorders,
              }),
            ),
          ],
        });

      const learnerTable = new Table({
        width: { size: leftWidth, type: WidthType.DXA },
        columnWidths: leftColumnWidths,
        layout: TableLayoutType.FIXED,
        borders: fullGridBorders,
        rows: [
          learnerHeader,
          groupRow("MALE"),
          ...male.map(learnerRow),
          totalRow(male.length, "TOTAL MALE"),
          groupRow("FEMALE"),
          ...female.map(learnerRow),
          totalRow(female.length, "TOTAL FEMALE"),
          totalRow(male.length + female.length, "COMBINED"),
        ],
      });

      const rightColumnWidths = [
        Math.floor(rightWidth * 0.34),
        Math.floor(rightWidth * 0.21),
        Math.floor(rightWidth * 0.22),
      ];
      rightColumnWidths.push(
        rightWidth - rightColumnWidths.reduce((sum, width) => sum + width, 0),
      );

      const summaryRows = [
        ["PROMOTED", sm.promoted, sf.promoted],
        ["PROMOTED\n WITH ACADEMIC EXCELLENCE AWARD", sm.conditional, sf.conditional],
        ["RETAINED", sm.retained, sf.retained],
      ] as const;

      const summaryTable = new Table({
        width: { size: rightWidth, type: WidthType.DXA },
        columnWidths: rightColumnWidths,
        layout: TableLayoutType.FIXED,
        borders: fullGridBorders,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              cell({
                text: "SUMMARY TABLE",
                width: rightWidth,
                columnSpan: 4,
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              cell({
                text: "STATUS",
                width: rightColumnWidths[0],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
              cell({
                text: "MALE",
                width: rightColumnWidths[1],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
              cell({
                text: "FEMALE",
                width: rightColumnWidths[2],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
              cell({
                text: "TOTAL",
                width: rightColumnWidths[3],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
            ],
          }),
          ...summaryRows.map(
            ([label, maleCount, femaleCount]) =>
              new TableRow({
                cantSplit: true,
                children: [
                  cell({
                    text: label,
                    width: rightColumnWidths[0],
                    bold: true,
                    size: 7,
                    alignment: AlignmentType.LEFT,
                    borders: allMediumBorders,
                  }),
                  cell({
                    text: String(maleCount),
                    width: rightColumnWidths[1],
                    size: 8,
                    borders: allMediumBorders,
                  }),
                  cell({
                    text: String(femaleCount),
                    width: rightColumnWidths[2],
                    size: 8,
                    borders: allMediumBorders,
                  }),
                  cell({
                    text: String(maleCount + femaleCount),
                    width: rightColumnWidths[3],
                    size: 8,
                    borders: allMediumBorders,
                  }),
                ],
              }),
          ),
        ],
      });

      const progressRows = [...bands].reverse().map((band) => {
        const maleCount = bandCount(male, band.lo, band.hi);
        const femaleCount = bandCount(female, band.lo, band.hi);
        return new TableRow({
          cantSplit: true,
          height: { value: 360, rule: HeightRule.ATLEAST },
          children: [
            cell({
              text: band.label,
              width: rightColumnWidths[0],
              bold: true,
              size: 7,
              alignment: AlignmentType.LEFT,
              borders: allMediumBorders,
            }),
            cell({
              text: String(maleCount),
              width: rightColumnWidths[1],
              size: 8,
              borders: allMediumBorders,
            }),
            cell({
              text: String(femaleCount),
              width: rightColumnWidths[2],
              size: 8,
              borders: allMediumBorders,
            }),
            cell({
              text: String(maleCount + femaleCount),
              width: rightColumnWidths[3],
              size: 8,
              borders: allMediumBorders,
            }),
          ],
        });
      });

      const progressTable = new Table({
        width: { size: rightWidth, type: WidthType.DXA },
        columnWidths: rightColumnWidths,
        layout: TableLayoutType.FIXED,
        borders: fullGridBorders,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              cell({
                text: "LEVEL OF PROGRESS AND ACHIEVEMENT",
                width: rightWidth,
                columnSpan: 4,
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              cell({
                text: "Descriptor &\nGrading",
                width: rightColumnWidths[0],
                bold: true,
                size: 7,
                alignment: AlignmentType.LEFT,
                borders: allMediumBorders,
              }),
              cell({
                text: "MALE",
                width: rightColumnWidths[1],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
              cell({
                text: "FEMALE",
                width: rightColumnWidths[2],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
              cell({
                text: "TOTAL",
                width: rightColumnWidths[3],
                bold: true,
                size: 8,
                borders: allMediumBorders,
              }),
            ],
          }),
          ...progressRows,
        ],
      });

      const signatureBlock = (heading: string, name: string, role: string) => [
        paragraph(heading, {
          bold: true,
          size: 8,
          alignment: AlignmentType.LEFT,
          before: 120,
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 300, after: 0 },
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              size: 5,
              color: "000000",
              space: 1,
            },
          },
          children: [
            new TextRun({
              text: name.toUpperCase(),
              size: 16,
              font: wordSansFont,
            }),
          ],
        }),
        paragraph(`${role}\n(Name and Signature)`, {
          italic: true,
          size: 7,
        }),
      ];

      const rightContent: Array<Paragraph | Table> = [
        summaryTable,
        new Paragraph({ spacing: { after: 35 } }),
        progressTable,
      ];

      if (length === "full") {
        rightContent.push(
          paragraph("Instructions:", {
            bold: true,
            size: 8,
            alignment: AlignmentType.LEFT,
            before: 120,
          }),
          paragraph(
            [
              "1. The SCC shall conduct checking in their own school; no swapping of SCC from one school to another is permitted.",
              "2. The name of SCC members shall be printed and they shall put their signature on top. Additional space may be added.",
              "3. The school head is accountable and liable for any wrongful entry on the forms; therefore, the DCC is not required to put their names and signatures in SF 5.",
              "4. Only LIS-generated SF5 shall be recognized.",
              "5. This form shall be submitted to the DCC together with the accomplished SFCR1.",
            ].join("\n"),
            {
              size: 6,
              alignment: AlignmentType.LEFT,
              before: 40,
            },
          ),
        );
      }

      rightContent.push(
        ...signatureBlock(
          "PREPARED BY:",
          klass.teacher_name || profile?.full_name || "",
          "Class Adviser",
        ),
        ...signatureBlock(
          "CERTIFIED CORRECT & SUBMITTED BY:",
          schoolHeadName,
          "School Head & SCC Chair",
        ),
      );

      if (length === "full") {
        rightContent.push(
          paragraph("REVIEWED BY: SCC Members", {
            bold: true,
            size: 8,
            alignment: AlignmentType.LEFT,
            before: 120,
          }),
          ...[0, 1, 2].flatMap(() => [
            new Paragraph({
              spacing: { before: 140, after: 0 },
              border: {
                bottom: {
                  style: BorderStyle.SINGLE,
                  size: 5,
                  color: "000000",
                  space: 1,
                },
              },
            }),
            paragraph("(Signature Over Printed Name)", {
              italic: true,
              size: 6,
            }),
          ]),
          new Paragraph({
            spacing: { before: 40, after: 0 },
            border: {
              top: {
                style: BorderStyle.SINGLE,
                size: 5,
                color: "000000",
                space: 1,
              },
            },
            children: [
              new TextRun({
                text: "Generated thru LIS (SCC CO-Chair)",
                size: 14,
                font: wordSansFont,
              }),
            ],
          }),
        );
      }

      const bodyTable = new Table({
        width: { size: contentWidth, type: WidthType.DXA },
        columnWidths: [leftWidth, rightWidth],
        layout: TableLayoutType.FIXED,
        borders: noBorders,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: leftWidth, type: WidthType.DXA },
                borders: noBorders,
                verticalAlign: VerticalAlign.TOP,
                margins: { top: 0, bottom: 0, left: 0, right: 45 },
                children: [learnerTable],
              }),
              new TableCell({
                width: { size: rightWidth, type: WidthType.DXA },
                borders: noBorders,
                verticalAlign: VerticalAlign.TOP,
                margins: { top: 0, bottom: 0, left: 45, right: 0 },
                children: rightContent,
              }),
            ],
          }),
        ],
      });

      const wordDocument = new WordDocument({
        styles: {
          default: {
            document: {
              run: { font: wordSansFont, size: 14, color: "000000" },
              paragraph: {
                spacing: { before: 0, after: 0, line: 190 },
              },
            },
          },
        },
        sections: [
          {
            properties: {
              page: {
                size: {
                  width: pageWidth,
                  height: pageHeight,
                  orientation: PageOrientation.LANDSCAPE,
                },
                margin: {
                  top: margin,
                  bottom: margin,
                  left: margin,
                  right: margin,
                },
              },
            },
            children: [
              paragraph(
                "School Form 5 (SF 5) Report on Promotion and Level of Proficiency & Achievement",
                { bold: true, size: 17, after: 20 },
              ),
              paragraph(
                "(This replaces Forms 18-E1, 18-E2, 18A and List of Graduates)",
                { italic: true, size: 7, after: 80 },
              ),
              ...metadataTables,
              new Paragraph({ spacing: { after: 45 } }),
              bodyTable,
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(wordDocument);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");

      anchor.href = url;
      anchor.download = `SF5_Promotion_${klass.section || "class"}.docx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);

      toast.success("SF5 Word document downloaded");
    } catch (error) {
      console.error("Unable to export SF5 to Word", error);
      toast.error(
        error instanceof Error ? error.message : "Unable to export SF5 to Word",
      );
    }
  };

  const exportToExcel = async () => {
    if (!klass) {
      toast.error("Please select a class before exporting");
      return;
    }

    const toExcelRow = (row: (typeof rows)[number]): SF5ExcelRow => ({
      lrn: row.st.lrn || "",
      name: [row.st.last_name, row.st.first_name, row.st.middle_name]
        .filter(Boolean)
        .join(", ")
        .toUpperCase(),
      average: row.avg == null ? null : Math.round(row.avg),
      action: getSF5Action(row.avg),
      failedAreas: row.fail.toUpperCase(),
    });

    try {
      await downloadSF5Excel(`SF5_Promotion_${klass.section || "class"}`, {
        region: profile?.region || klass.region || "",
        division: profile?.division || klass.division || "",
        schoolId: profile?.school_id || klass.school_id || "",
        schoolYear: klass.school_year || profile?.school_year || "",
        curriculum: "K to 12",
        schoolName: profile?.school_name || klass.school_name || "",
        gradeLevel: formatGradeLevel(klass.grade_level),
        section: klass.section || "",
        maleRows: male.map(toExcelRow),
        femaleRows: female.map(toExcelRow),
        summary: {
          promotedMale: sm.promoted,
          promotedFemale: sf.promoted,
          conditionalMale: sm.conditional,
          conditionalFemale: sf.conditional,
          retainedMale: sm.retained,
          retainedFemale: sf.retained,
        },
        progress: [...bands].reverse().map((band) => ({
          label: band.label,
          male: bandCount(male, band.lo, band.hi),
          female: bandCount(female, band.lo, band.hi),
        })),
        adviser: klass.teacher_name || profile?.full_name || "",
        schoolHead: schoolHeadName,
      });
      toast.success("SF5 Excel spreadsheet downloaded");
    } catch (error) {
      console.error("Unable to export SF5 to Excel", error);
      toast.error(
        error instanceof Error
          ? error.message
          : "Unable to export SF5 to Excel",
      );
    }
  };

  // Header/config row (like ILS "Letter Configuration")
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
          <GraduationCap className="size-5" style={{ color: DEPED_BLUE }} />
          School Form 5 — Report on Promotion
        </div>
      </div>

      {/* Yellow configuration card (matches ILS visual language) */}
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
        fileName={`SF5_Promotion_${klass?.section || "class"}.pdf`}
        docBaseName={`SF5_Promotion_${klass?.section || "class"}`}
        printTargetId="sf5-doc"
        orientation="landscape"
        printMargin="0in"
        fitPrintToPage
        length={length}
        onLengthChange={setLength}
        onCopyToWord={copySF5ToWord}
        hideCopyToWord={
          String(klass?.grade_level ?? "")
            .trim()
            .toLowerCase() === "grade 12"
        }
        onExportToExcel={exportToExcel}
        pageLabel="Page 1 of 1"
      >
        <div
          id="sf5-doc"
          className="overflow-hidden bg-white px-6 py-5 text-[8px] leading-tight text-black"
          style={{ width: 1200, fontFamily: "Arial, Helvetica, sans-serif" }}
        >
          <div className="text-center text-[17px] font-bold leading-none">
            School Form 5 (SF 5) Report on Promotion and Level of Proficiency
            &amp; Achievement
          </div>
          <div className="mt-3 text-center text-[7px] italic">
            (This replaces Forms 18-E1, 18-E2, 18A and List of Graduates)
          </div>

          <div
            className="mt-2 grid items-stretch text-[8px]"
            style={{
              gridTemplateColumns: SF5_COLUMN_WIDTHS.map(
                (width) => `${width}fr`,
              ).join(" "),
              gridTemplateRows: "18px 18px 20px",
            }}
          >
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "1 / span 4", gridRow: 1 }}
            >
              Region
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "5", gridRow: 1 }}
            >
              {profile?.region || klass?.region || ""}
            </div>
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "6", gridRow: 1 }}
            >
              Division
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "7 / span 4", gridRow: 1 }}
            >
              {profile?.division || klass?.division || ""}
            </div>

            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "1 / span 4", gridRow: 2 }}
            >
              School ID
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "5 / span 2", gridRow: 2 }}
            >
              {profile?.school_id || klass?.school_id || ""}
            </div>
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "7 / span 2", gridRow: 2 }}
            >
              School Year
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "9 / span 2", gridRow: 2 }}
            >
              {klass?.school_year || profile?.school_year || ""}
            </div>
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "11", gridRow: 2 }}
            >
              Curriculum
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "12 / span 6", gridRow: 2 }}
            >
              K to 12
            </div>

            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "1 / span 4", gridRow: 3 }}
            >
              School Name
            </div>
            <div
              className="flex items-center justify-center whitespace-nowrap border border-black px-1"
              style={{ gridColumn: "5 / span 6", gridRow: 3 }}
            >
              {profile?.school_name || klass?.school_name || ""}
            </div>
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "11", gridRow: 3 }}
            >
              Grade Level
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "12 / span 4", gridRow: 3 }}
            >
              {formatGradeLevel(klass?.grade_level)}
            </div>
            <div
              className="flex items-center justify-end pr-1"
              style={{ gridColumn: "16 / span 2", gridRow: 3 }}
            >
              Section
            </div>
            <div
              className="flex items-center justify-center border border-black px-1"
              style={{ gridColumn: "19 / span 8", gridRow: 3 }}
            >
              {klass?.section || ""}
            </div>
          </div>

          <div
            className="mt-1 grid items-start"
            style={{
              gridTemplateColumns: SF5_COLUMN_WIDTHS.map(
                (width) => `${width}fr`,
              ).join(" "),
            }}
          >
            <table
              className="w-full table-fixed border-collapse text-[7.5px]"
              style={{ gridColumn: "1 / span 12" }}
            >
              <colgroup>
                <col className="w-[14.3%]" />
                <col className="w-[36.6%]" />
                <col className="w-[10.5%]" />
                <col className="w-[19.6%]" />
                <col className="w-[19%]" />
              </colgroup>
              <thead>
                <tr className="h-[70px]">
                  <th className="border-[1.5px] border-black px-1">LRN</th>
                  <th className="border-[1.5px] border-black px-1">
                    LEARNER&apos;S NAME
                    <br />
                    <span className="font-normal">
                      (Last Name, First Name, Middle Name)
                    </span>
                  </th>
                  <th className="border-[1.5px] border-black px-1">
                    GENERAL
                    <br />
                    AVERAGE
                  </th>
                  <th className="border-[1.5px] border-black px-1">
                    ACTION TAKEN:
                    <br />
                    RETAINED, PROMOTED
                    <br />
                    or PROMOTED WITH ACADEMIC EXCELLENCE AWARD
                  </th>
                  <th className="border-[1.5px] border-black px-1 font-normal">
                    Did Not Meet Expectations of the ff.
                    <br />
                    Learning Area/s as of end of current
                    <br />
                    School Year
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr className="h-[16px] font-bold">
                  <td className="border-[1.5px] border-black px-1">MALE</td>
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                </tr>
                {male.map((r) => (
                  <tr key={r.st.id} className="h-[15px]">
                    <td className="whitespace-nowrap border border-black px-1 text-right">
                      {r.st.lrn || ""}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {[r.st.last_name, r.st.first_name, r.st.middle_name]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                    <td className="border border-black px-1 text-right">
                      {r.avg != null ? Math.round(r.avg) : ""}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {getSF5Action(r.avg)}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {r.fail}
                    </td>
                  </tr>
                ))}
                <tr className="h-[16px] font-bold">
                  <td className="border-[1.5px] border-black px-1 text-right">
                    {male.length}
                  </td>
                  <td className="border-[1.5px] border-black px-1">
                    &lt;=== TOTAL MALE
                  </td>
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                </tr>
                <tr className="h-[16px] font-bold">
                  <td className="border-[1.5px] border-black px-1">FEMALE</td>
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                  <td className="border-[1.5px] border-black" />
                </tr>
                {female.map((r) => (
                  <tr key={r.st.id} className="h-[15px]">
                    <td className="whitespace-nowrap border border-black px-1 text-right">
                      {r.st.lrn || ""}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {[r.st.last_name, r.st.first_name, r.st.middle_name]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                    <td className="border border-black px-1 text-right">
                      {r.avg != null ? Math.round(r.avg) : ""}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {getSF5Action(r.avg)}
                    </td>
                    <td className="border border-black px-1 uppercase">
                      {r.fail}
                    </td>
                  </tr>
                ))}
                <tr className="h-[16px] font-bold">
                  <td className="border-[1.5px] border-black px-1 text-right">
                    {female.length}
                  </td>
                  <td className="border-[1.5px] border-black px-1">
                    &lt;=== TOTAL FEMALE
                  </td>
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                </tr>
                <tr className="h-[16px] font-bold">
                  <td className="border-[1.5px] border-black px-1 text-right">
                    {male.length + female.length}
                  </td>
                  <td className="border-[1.5px] border-black px-1">
                    &lt;=== COMBINED
                  </td>
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                  <td className="border-[1.5px] border-black bg-[#c7c7ff]" />
                </tr>
              </tbody>
            </table>

            <div
              className="text-[7.5px]"
              style={{ gridColumn: "14 / span 13" }}
            >
              <table className="w-full table-fixed border-collapse">
                <colgroup>
                  <col className="w-[28%]" />
                  <col className="w-[21%]" />
                  <col className="w-[24%]" />
                  <col className="w-[27%]" />
                </colgroup>
                <thead>
                  <tr>
                    <th
                      className="h-[26px] border-[1.5px] border-black text-center font-bold"
                      colSpan={4}
                    >
                      SUMMARY TABLE
                    </th>
                  </tr>
                  <tr className="h-[22px]">
                    <th className="border-[1.5px] border-black">STATUS</th>
                    <th className="border-[1.5px] border-black">MALE</th>
                    <th className="border-[1.5px] border-black">FEMALE</th>
                    <th className="border-[1.5px] border-black">TOTAL</th>
                  </tr>
                </thead>
                <tbody className="text-center">
                  <tr className="h-[23px]">
                    <td className="border-[1.5px] border-black px-1 text-left font-bold">
                      PROMOTED
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.promoted}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sf.promoted}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.promoted + sf.promoted}
                    </td>
                  </tr>
                  <tr className="h-[30px]">
                    <td className="border-[1.5px] border-black px-1 text-left text-[6.5px] font-bold italic">
                      PROMOTED
                      <br />
                      WITH ACADEMIC EXCELLENCE AWARD
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.conditional}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sf.conditional}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.conditional + sf.conditional}
                    </td>
                  </tr>
                  <tr className="h-[30px]">
                    <td className="border-[1.5px] border-black px-1 text-left font-bold">
                      RETAINED
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.retained}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sf.retained}
                    </td>
                    <td className="border-[1.5px] border-black">
                      {sm.retained + sf.retained}
                    </td>
                  </tr>
                </tbody>
              </table>

              <table className="mt-1 w-full table-fixed border-collapse">
                <colgroup>
                  <col className="w-[28%]" />
                  <col className="w-[21%]" />
                  <col className="w-[24%]" />
                  <col className="w-[27%]" />
                </colgroup>
                <thead>
                  <tr>
                    <th
                      className="h-[28px] border-[1.5px] border-black text-center font-bold"
                      colSpan={4}
                    >
                      LEVEL OF PROGRESS AND ACHIEVEMENT
                    </th>
                  </tr>
                  <tr className="h-[30px]">
                    <th className="border-[1.5px] border-black px-1 text-left">
                      Descriptor &amp;
                      <br />
                      Grading
                    </th>
                    <th className="border-[1.5px] border-black">MALE</th>
                    <th className="border-[1.5px] border-black">FEMALE</th>
                    <th className="border-[1.5px] border-black">TOTAL</th>
                  </tr>
                </thead>
                <tbody className="text-center">
                  {[...bands].reverse().map((band) => {
                    const maleCount = bandCount(male, band.lo, band.hi);
                    const femaleCount = bandCount(female, band.lo, band.hi);

                    return (
                      <tr key={band.label} className="h-[35px]">
                        <td className="border-[1.5px] border-black px-1 text-left font-bold">
                          {band.label}
                        </td>
                        <td className="border-[1.5px] border-black">
                          {maleCount}
                        </td>
                        <td className="border-[1.5px] border-black">
                          {femaleCount}
                        </td>
                        <td className="border-[1.5px] border-black">
                          {maleCount + femaleCount}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>

              {length === "full" && (
                <div className="mt-4">
                  <div className="font-bold">Instructions:</div>
                  <ol className="ml-3 mt-2 list-decimal space-y-1 text-[6.5px]">
                    <li>
                      The SCC shall conduct checking in their own school; no
                      swapping of SCC from one school to another is permitted.
                    </li>
                    <li>
                      The name of SCC members shall be printed and they shall
                      put their signature on top. Additional space may be added.
                    </li>
                    <li>
                      The school head is accountable and liable for any wrongful
                      entry on the forms; therefore, the DCC is not required to
                      put their names and signatures in SF 5.
                    </li>
                    <li>Only LIS-generated SF5 shall be recognized.</li>
                    <li>
                      This form shall be submitted to the DCC together with the
                      accomplished SFCR1.
                    </li>
                  </ol>
                </div>
              )}

              <div className="mt-4">
                <div className="font-bold">PREPARED BY:</div>
                <div className="mt-9 border-b border-black pb-0.5 text-center text-[8px]">
                  {(
                    klass?.teacher_name ||
                    profile?.full_name ||
                    ""
                  ).toUpperCase()}
                </div>
                <div className="text-center text-[7px] italic">
                  Class Adviser
                  <br />
                  (Name and Signature)
                </div>

                <div className="mt-4 font-bold">
                  CERTIFIED CORRECT &amp; SUBMITTED BY:
                </div>
                <div className="mt-9 border-b border-black pb-0.5 text-center text-[8px]">
                  {schoolHeadName.toUpperCase()}
                </div>
                <div className="text-center text-[7px] italic">
                  School Head &amp; SCC Chair
                  <br />
                  (Name and Signature)
                </div>

                {length === "full" && (
                  <>
                    <div className="mt-4 font-bold">
                      REVIEWED BY: SCC Members
                    </div>
                    {[0, 1, 2].map((index) => (
                      <div key={index} className="mt-4">
                        <div className="border-b border-black">&nbsp;</div>
                        <div className="text-center text-[6.5px] italic">
                          (Signature Over Printed Name)
                        </div>
                      </div>
                    ))}
                    <div className="mt-4 border-t border-black pt-1 text-center text-[6.5px]">
                      Generated thru LIS (SCC CO-Chair)
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
