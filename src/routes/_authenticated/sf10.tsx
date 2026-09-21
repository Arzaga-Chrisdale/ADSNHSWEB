import { createFileRoute, Link } from "@tanstack/react-router";

import { useQuery } from "@tanstack/react-query";
import { Fragment, useEffect, useMemo, useState } from "react";
import { ArrowLeft, BookOpen } from "lucide-react";
import { toast } from "sonner";
import ExcelJS, {
  type Alignment,
  type Borders,
  type Fill,
  type Font,
  type PaperSize,
  type Worksheet,
} from "exceljs";
import { supabase } from "@/integrations/supabase/client";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { computeAverage, type ClassRow, type GradeRow, type StudentRow } from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import depedLogo from "@/assets/deped.jpg";
import schoolLogo from "@/assets/ASNSHS Logo.png";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/sf10")({
  component: SF10Page,
});

type ProfileRow = {
  school_name?: string | null;
  school_id?: string | null;
  region?: string | null;
  division?: string | null;
  district?: string | null;
  school_year?: string | null;
  principal?: string | null;
  full_name?: string | null;
};

type SF10LearningAreaRow = {
  learningArea: string;
  quarterlyRatings: Array<number | null>;
  finalRating: number | null;
  remarks: string;
};

type SF10Record = {
  schoolName: string;
  schoolId: string;
  district: string;
  division: string;
  region: string;
  gradeLevel: string;
  section: string;
  schoolYear: string;
  adviser: string;
  rows: SF10LearningAreaRow[];
  generalAverage: number | null;
};

type SF10LearnerExport = {
  lastName: string;
  firstName: string;
  middleName: string;
  nameExtension: string;
  lrn: string;
  birthdate: string;
  sex: string;
  record: SF10Record;
};

type SF10ExportOptions = {
  schoolHead: string;
  learners: SF10LearnerExport[];
};

type ExcelCellStyle = {
  font?: Partial<Font>;
  alignment?: Partial<Alignment>;
  fill?: Fill;
  border?: Partial<Borders>;
  numberFormat?: string;
};

const SF10_LEARNING_AREAS = [
  "Filipino",
  "English",
  "Mathematics",
  "Science",
  "Araling Panlipunan (AP)",
  "Values Education",
  "MAPEH",
  "Music & Arts",
  "Physical Education & Health",
  "Research",
] as const;

const GENERAL_AVERAGE_EXCLUSIONS = new Set([
  "Music & Arts",
  "Physical Education & Health",
]);

const SF10_FRONT_COLUMN_WIDTHS = [
  6, 6, 6, 6, 6, 6, 5, 5, 5, 5, 7.5, 10, 1.5, 6, 6, 6, 6, 6, 6, 5, 5, 5, 5, 7.5, 10,
] as const;

const GREY_FILL: Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFD9D9D9" },
};

const LIGHT_GREY_FILL: Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFF2F2F2" },
};

const THIN = { style: "thin" as const, color: { argb: "FF000000" } };
const MEDIUM = { style: "medium" as const, color: { argb: "FF000000" } };

function formatBirthdate(value?: string | null) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }).format(date);
}

function normalizeSubject(value: string) {
  return value
    .toLowerCase()
    .replaceAll("&", "and")
    .replaceAll(/[^a-z0-9]+/g, " ")
    .trim();
}

function subjectMatches(learningArea: string, gradeSubject: string) {
  const area = normalizeSubject(learningArea.replace(/^\*/, ""));
  const subject = normalizeSubject(gradeSubject);

  if (area === subject) return true;
  if (learningArea === "Araling Panlipunan (AP)") {
    return subject === "araling panlipunan" || subject === "ap" || subject.includes("araling panlipunan");
  }
  if (learningArea === "Values Education") {
    return subject.includes("values education") || subject.includes("pagpapakatao");
  }
  if (learningArea === "Music & Arts") {
    return subject.includes("music") || subject.includes("arts");
  }
  if (learningArea === "Physical Education & Health") {
    return subject.includes("physical education") || subject === "pe" || subject.includes("health");
  }

  return false;
}

function scoreForArea(grades: GradeRow[], studentId: string, learningArea: string, term: string) {
  const terms = term === "4" ? ["4", "final"] : [term];

  return (
    grades.find(
      (grade) =>
        grade.student_id === studentId &&
        subjectMatches(learningArea, grade.subject) &&
        terms.includes(grade.term),
    )?.score ?? null
  );
}

function buildSF10Record(
  student: StudentRow,
  grades: GradeRow[],
  klass: ClassRow,
  profile?: ProfileRow | null,
): SF10Record {
  const rows = SF10_LEARNING_AREAS.map((learningArea) => {
    const quarterlyRatings = ["1", "2", "3", "4"].map((term) =>
      scoreForArea(grades, student.id, learningArea, term),
    );
    const finalRating = computeAverage(quarterlyRatings);

    return {
      learningArea,
      quarterlyRatings,
      finalRating,
      remarks: finalRating == null ? "" : finalRating >= 75 ? "PASSED" : "FAILED",
    };
  });

  const generalAverage = computeAverage(
    rows
      .filter((row) => !GENERAL_AVERAGE_EXCLUSIONS.has(row.learningArea))
      .map((row) => row.finalRating),
  );

  return {
    schoolName:
      profile?.school_name || klass.school_name || "Agusan del Sur National Science High School",
    schoolId: profile?.school_id || klass.school_id || "",
    district: profile?.district || klass.district || "",
    division: profile?.division || klass.division || "",
    region: profile?.region || klass.region || "",
    gradeLevel: klass.grade_level || "",
    section: klass.section || "",
    schoolYear: klass.school_year || profile?.school_year || "",
    adviser: klass.teacher_name || profile?.full_name || "",
    rows,
    generalAverage,
  };
}

function toLearnerExport(
  student: StudentRow,
  grades: GradeRow[],
  klass: ClassRow,
  profile?: ProfileRow | null,
): SF10LearnerExport {
  return {
    lastName: student.last_name || "",
    firstName: student.first_name || "",
    middleName: student.middle_name || "",
    nameExtension: "",
    lrn: student.lrn || "",
    birthdate: formatBirthdate(student.birthdate),
    sex: student.sex === "male" ? "M" : student.sex === "female" ? "F" : "",
    record: buildSF10Record(student, grades, klass, profile),
  };
}

function columnName(column: number) {
  let result = "";
  let current = column;

  while (current > 0) {
    current -= 1;
    result = String.fromCharCode(65 + (current % 26)) + result;
    current = Math.floor(current / 26);
  }

  return result;
}

function cellAddress(column: number, row: number) {
  return `${columnName(column)}${row}`;
}

function excelRange(startColumn: number, startRow: number, endColumn: number, endRow: number) {
  return `${cellAddress(startColumn, startRow)}:${cellAddress(endColumn, endRow)}`;
}

function sf10Border({
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
    top: top ? MEDIUM : THIN,
    bottom: bottom ? MEDIUM : THIN,
    left: left ? MEDIUM : THIN,
    right: right ? MEDIUM : THIN,
  };
}

function mergeAndSetSF10(
  worksheet: Worksheet,
  range: string,
  value: string | number | null,
  style: ExcelCellStyle = {},
) {
  const [topLeft] = range.split(":");

  if (range.includes(":")) {
    worksheet.mergeCells(range);
  }

  const cell = worksheet.getCell(topLeft);
  cell.value = value === "" || value == null ? null : value;
  cell.font = {
    name: "Arial Narrow",
    size: 9,
    ...style.font,
  };
  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
    ...style.alignment,
  };

  if (style.fill) cell.fill = style.fill;
  if (style.border) cell.border = style.border;
  if (style.numberFormat) cell.numFmt = style.numberFormat;
}

function styleSF10Range(
  worksheet: Worksheet,
  range: string,
  style: ExcelCellStyle,
  includeEmpty = true,
) {
  const [from, to = from] = range.split(":");
  const fromMatch = from.match(/^([A-Z]+)(\d+)$/);
  const toMatch = to.match(/^([A-Z]+)(\d+)$/);

  if (!fromMatch || !toMatch) return;

  const columnNumber = (letters: string) =>
    [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
  const startColumn = columnNumber(fromMatch[1]);
  const endColumn = columnNumber(toMatch[1]);
  const startRow = Number(fromMatch[2]);
  const endRow = Number(toMatch[2]);

  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      const cell = worksheet.getCell(row, column);
      if (!includeEmpty && cell.value == null) continue;

      if (style.font) cell.font = { ...cell.font, ...style.font };
      if (style.alignment) cell.alignment = { ...cell.alignment, ...style.alignment };
      if (style.fill) cell.fill = style.fill;
      if (style.border) cell.border = { ...cell.border, ...style.border };
      if (style.numberFormat) cell.numFmt = style.numberFormat;
    }
  }
}

function setUnderlinedValue(
  worksheet: Worksheet,
  range: string,
  value: string,
  alignment: "left" | "center" = "left",
) {
  mergeAndSetSF10(worksheet, range, value, {
    font: { size: 8 },
    alignment: { horizontal: alignment, vertical: "bottom", wrapText: false, shrinkToFit: true },
    border: { bottom: THIN },
    numberFormat: "@",
  });
}

function setRecordMetadata(
  worksheet: Worksheet,
  startColumn: number,
  startRow: number,
  record?: SF10Record,
) {
  const c = startColumn;
  const r = startRow;

  mergeAndSetSF10(worksheet, excelRange(c, r, c + 1, r), "School:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 2, r, c + 7, r), record?.schoolName || "");
  mergeAndSetSF10(worksheet, excelRange(c + 8, r, c + 9, r), "School ID:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 10, r, c + 11, r), record?.schoolId || "");

  mergeAndSetSF10(worksheet, cellAddress(c, r + 1), "District:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 1, r + 1, c + 3, r + 1), record?.district || "");
  mergeAndSetSF10(worksheet, cellAddress(c + 4, r + 1), "Division:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 5, r + 1, c + 7, r + 1), record?.division || "");
  mergeAndSetSF10(worksheet, cellAddress(c + 8, r + 1), "Region:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 9, r + 1, c + 11, r + 1), record?.region || "");

  mergeAndSetSF10(worksheet, excelRange(c, r + 2, c + 1, r + 2), "Classified as Grade:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, cellAddress(c + 2, r + 2), record?.gradeLevel || "", "center");
  mergeAndSetSF10(worksheet, cellAddress(c + 3, r + 2), "Section:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 4, r + 2, c + 7, r + 2), record?.section || "");
  mergeAndSetSF10(worksheet, excelRange(c + 8, r + 2, c + 9, r + 2), "School Year:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(
    worksheet,
    excelRange(c + 10, r + 2, c + 11, r + 2),
    record?.schoolYear || "",
    "center",
  );

  mergeAndSetSF10(worksheet, excelRange(c, r + 3, c + 2, r + 3), "Name of Adviser/Teacher:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 3, r + 3, c + 7, r + 3), record?.adviser || "");
  mergeAndSetSF10(worksheet, excelRange(c + 8, r + 3, c + 9, r + 3), "Signature:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, excelRange(c + 10, r + 3, c + 11, r + 3), "");

  styleSF10Range(worksheet, excelRange(c, r, c + 11, r + 3), {
    font: { name: "Arial Narrow", size: 8 },
  });
}

function buildSF10RecordBlock(
  worksheet: Worksheet,
  startColumn: number,
  startRow: number,
  record?: SF10Record,
) {
  const c = startColumn;
  const r = startRow;

  setRecordMetadata(worksheet, c, r, record);

  mergeAndSetSF10(worksheet, excelRange(c, r + 4, c + 5, r + 5), "LEARNING AREAS", {
    font: { size: 9, bold: true },
  });
  mergeAndSetSF10(worksheet, excelRange(c + 6, r + 4, c + 9, r + 4), "Quarterly Rating", {
    font: { size: 9, bold: true },
  });
  [1, 2, 3, 4].forEach((quarter, index) => {
    mergeAndSetSF10(worksheet, cellAddress(c + 6 + index, r + 5), quarter, {
      font: { size: 8, bold: true },
    });
  });
  mergeAndSetSF10(worksheet, excelRange(c + 10, r + 4, c + 10, r + 5), "Final\nRating", {
    font: { size: 8, bold: true },
  });
  mergeAndSetSF10(worksheet, excelRange(c + 11, r + 4, c + 11, r + 5), "Remarks", {
    font: { size: 8, bold: true },
  });

  SF10_LEARNING_AREAS.forEach((learningArea, index) => {
    const row = r + 6 + index;
    const recordRow = record?.rows[index];
    const isComponent = ["Music", "Arts", "Physical Education", "Health"].includes(learningArea);
    const isOptional = learningArea.startsWith("*");

    mergeAndSetSF10(worksheet, excelRange(c, row, c + 5, row), learningArea, {
      font: {
        size: 8,
        bold: !isComponent && !isOptional,
        italic: isComponent,
      },
      alignment: { horizontal: "left", vertical: "middle", wrapText: false },
    });

    [0, 1, 2, 3].forEach((quarterIndex) => {
      mergeAndSetSF10(
        worksheet,
        cellAddress(c + 6 + quarterIndex, row),
        recordRow?.quarterlyRatings[quarterIndex] == null
          ? ""
          : Math.round(recordRow.quarterlyRatings[quarterIndex]!),
        {
          font: { size: 8 },
          numberFormat: "0",
        },
      );
    });

    mergeAndSetSF10(
      worksheet,
      cellAddress(c + 10, row),
      recordRow?.finalRating == null ? "" : Math.round(recordRow.finalRating),
      { font: { size: 8 }, numberFormat: "0" },
    );
    mergeAndSetSF10(worksheet, cellAddress(c + 11, row), recordRow?.remarks || "", {
      font: { size: 7 },
    });
  });

  const generalAverageRow = r + 21;
  mergeAndSetSF10(
    worksheet,
    excelRange(c, generalAverageRow, c + 9, generalAverageRow),
    "General Average",
    {
      font: { size: 8, bold: true },
      alignment: { horizontal: "left" },
    },
  );
  mergeAndSetSF10(
    worksheet,
    cellAddress(c + 10, generalAverageRow),
    record?.generalAverage == null ? "" : Math.round(record.generalAverage),
    { font: { size: 8, bold: true }, numberFormat: "0" },
  );
  mergeAndSetSF10(worksheet, cellAddress(c + 11, generalAverageRow), "", {
    font: { size: 8 },
  });

  mergeAndSetSF10(worksheet, excelRange(c, r + 22, c + 3, r + 22), "Remedial Classes", {
    font: { size: 8, bold: true },
  });
  mergeAndSetSF10(
    worksheet,
    excelRange(c + 4, r + 22, c + 11, r + 22),
    "Date Conducted:                              to",
    {
      font: { size: 8, bold: true },
      alignment: { horizontal: "left" },
    },
  );

  const remedialHeaders = [
    [c, c + 3, "Learning Areas"],
    [c + 4, c + 5, "Final Rating"],
    [c + 6, c + 7, "Remedial Class\nMark"],
    [c + 8, c + 9, "Recomputed\nFinal Grade"],
    [c + 10, c + 11, "Remarks"],
  ] as const;

  remedialHeaders.forEach(([from, to, label]) => {
    mergeAndSetSF10(worksheet, excelRange(from, r + 23, to, r + 23), label, {
      font: { size: 7, bold: true },
    });
  });

  [r + 24, r + 25].forEach((row) => {
    remedialHeaders.forEach(([from, to]) => {
      mergeAndSetSF10(worksheet, excelRange(from, row, to, row), "", {
        font: { size: 7 },
      });
    });
  });

  const tableRange = excelRange(c, r + 4, c + 11, r + 25);
  styleSF10Range(worksheet, tableRange, {
    border: sf10Border(),
  });
  styleSF10Range(worksheet, excelRange(c, r + 4, c + 11, r + 5), {
    fill: LIGHT_GREY_FILL,
  });

  for (let row = r; row <= r + 3; row += 1) worksheet.getRow(row).height = 16;
  worksheet.getRow(r + 4).height = 17;
  worksheet.getRow(r + 5).height = 15;
  for (let row = r + 6; row <= r + 21; row += 1) worksheet.getRow(row).height = 18;
  worksheet.getRow(r + 22).height = 15;
  worksheet.getRow(r + 23).height = 24;
  worksheet.getRow(r + 24).height = 18;
  worksheet.getRow(r + 25).height = 18;
}

async function imageAssetToDataUrl(assetUrl: string) {
  if (assetUrl.startsWith("data:")) return assetUrl;

  const response = await fetch(assetUrl);
  if (!response.ok) throw new Error(`Unable to load image (${response.status})`);
  const blob = await response.blob();

  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("Unable to read image"));
    reader.readAsDataURL(blob);
  });
}

async function addSF10Images(workbook: ExcelJS.Workbook) {
  let depedImageId: number | null = null;
  let schoolImageId: number | null = null;

  try {
    const depedData = await imageAssetToDataUrl(depedLogo);
    depedImageId = workbook.addImage({
      base64: depedData,
      extension: depedData.startsWith("data:image/png") ? "png" : "jpeg",
    });
  } catch (error) {
    console.warn("Unable to add the DepEd logo to SF10 Excel", error);
  }

  try {
    const schoolData = await imageAssetToDataUrl(schoolLogo);
    schoolImageId = workbook.addImage({
      base64: schoolData,
      extension: "png",
    });
  } catch (error) {
    console.warn("Unable to add the school logo to SF10 Excel", error);
  }

  return { depedImageId, schoolImageId };
}

function configureSF10Worksheet(worksheet: Worksheet, printArea: string) {
  worksheet.views = [
    {
      state: "normal",
      showGridLines: false,
      zoomScale: 90,
      activeCell: "A1",
    },
  ];
  worksheet.pageSetup = {
    orientation: "landscape",
    paperSize: 14 as PaperSize,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    verticalCentered: false,
    pageOrder: "overThenDown",
    printArea,
    margins: {
      left: 0.18,
      right: 0.18,
      top: 0.2,
      bottom: 0.2,
      header: 0.1,
      footer: 0.1,
    },
  };
  worksheet.properties.defaultRowHeight = 10;

  SF10_FRONT_COLUMN_WIDTHS.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });
}

function setPersonalInformation(worksheet: Worksheet, learner: SF10LearnerExport) {
  mergeAndSetSF10(worksheet, "A6:Y6", "LEARNER'S PERSONAL INFORMATION", {
    font: { size: 10, bold: true },
    fill: GREY_FILL,
    border: sf10Border(),
  });

  mergeAndSetSF10(worksheet, "A7:B7", "LAST NAME:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "C7:G7", learner.lastName.toUpperCase());
  mergeAndSetSF10(worksheet, "H7:I7", "FIRST NAME:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "J7:N7", learner.firstName.toUpperCase());
  mergeAndSetSF10(worksheet, "O7:Q7", "NAME EXTN. (Jr,I,II)", {
    font: { size: 7 },
  });
  setUnderlinedValue(worksheet, "R7:S7", learner.nameExtension.toUpperCase(), "center");
  mergeAndSetSF10(worksheet, "T7:U7", "MIDDLE NAME:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "V7:Y7", learner.middleName.toUpperCase());

  mergeAndSetSF10(worksheet, "A8:E8", "Learner Reference Number (LRN):", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "F8:M8", learner.lrn, "center");
  worksheet.getCell("F8").value = {
    richText: [
      {
        font: { name: "Arial Narrow", size: 8 },
        text: learner.lrn,
      },
    ],
  };
  mergeAndSetSF10(worksheet, "N8:R8", "Birthdate (mm/dd/yyyy):", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "S8:V8", learner.birthdate, "center");
  mergeAndSetSF10(worksheet, "W8:X8", "Sex:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "Y8", learner.sex, "center");

  worksheet.getRow(6).height = 18;
  worksheet.getRow(7).height = 20;
  worksheet.getRow(8).height = 20;
}

function setEligibilitySection(worksheet: Worksheet) {
  mergeAndSetSF10(worksheet, "A9:Y9", "ELIGIBILITY FOR JHS ENROLMENT", {
    font: { size: 11, bold: true },
    fill: GREY_FILL,
    border: sf10Border(),
  });

  mergeAndSetSF10(worksheet, "A10:F10", "Elementary School Completer:", {
    font: { size: 8, italic: true },
    alignment: { horizontal: "left" },
  });
  mergeAndSetSF10(worksheet, "G10:K10", "Kinder Progress Report", {
    font: { size: 8, italic: true },
  });
  mergeAndSetSF10(worksheet, "L10:P10", "ECCD Checklist", {
    font: { size: 8, italic: true },
  });
  mergeAndSetSF10(worksheet, "Q10:Y10", "Kindergarten Certificate of Completion", {
    font: { size: 8, italic: true },
  });

  mergeAndSetSF10(worksheet, "A11:C11", "Name of School:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "D11:J11", "");
  mergeAndSetSF10(worksheet, "K11:L11", "School ID:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "M11:O11", "");
  mergeAndSetSF10(worksheet, "P11:R11", "Address of School:", {
    alignment: { horizontal: "left" },
  });
  setUnderlinedValue(worksheet, "S11:Y11", "");

  mergeAndSetSF10(worksheet, "A12:Y12", "Other Credential Presented", {
    alignment: { horizontal: "left" },
  });
  mergeAndSetSF10(worksheet, "A13:C13", "PEPT Passer  Rating:", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "D13:F13", "");
  mergeAndSetSF10(worksheet, "G13:M13", "Date of Examination/Assessment (mm/dd/yyyy):", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "N13:Q13", "");
  mergeAndSetSF10(worksheet, "R13:U13", "Others (Pls. Specify):", {
    alignment: { horizontal: "right" },
  });
  setUnderlinedValue(worksheet, "V13:Y13", "");

  [9, 10, 11, 12, 13].forEach((row) => {
    worksheet.getRow(row).height = row === 9 ? 18 : 17;
  });
}

function setFrontHeader(
  worksheet: Worksheet,
  imageIds: { depedImageId: number | null; schoolImageId: number | null },
) {
  mergeAndSetSF10(worksheet, "D1:V1", "Republic of the Philippines", {
    font: { size: 12 },
    border: {},
  });
  mergeAndSetSF10(worksheet, "D2:V2", "Department of Education", {
    font: { size: 12 },
    border: {},
  });
  mergeAndSetSF10(
    worksheet,
    "D3:V4",
    "Learner Permanent Academic Record for Junior High School (SF10-JHS)\n(Formerly Form 137)",
    {
      font: { size: 15, bold: true },
      border: {},
    },
  );
  mergeAndSetSF10(worksheet, "A5:C5", "SF10-JHS", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left" },
    border: {},
  });

  if (imageIds.depedImageId != null) {
    worksheet.addImage(imageIds.depedImageId, {
      tl: { col: 0.15, row: 0.15 },
      ext: { width: 105, height: 52 },
      editAs: "oneCell",
    });
  }
  if (imageIds.schoolImageId != null) {
    worksheet.addImage(imageIds.schoolImageId, {
      tl: { col: 23.05, row: 0.05 },
      ext: { width: 50, height: 50 },
      editAs: "oneCell",
    });
  }

  worksheet.getRow(1).height = 18;
  worksheet.getRow(2).height = 18;
  worksheet.getRow(3).height = 24;
  worksheet.getRow(4).height = 24;
  worksheet.getRow(5).height = 12;
}

function setCertificationBlock(
  worksheet: Worksheet,
  startRow: number,
  learner: SF10LearnerExport,
  schoolHead: string,
) {
  mergeAndSetSF10(worksheet, excelRange(1, startRow, 25, startRow), "CERTIFICATION", {
    font: { size: 11, bold: true },
    fill: LIGHT_GREY_FILL,
    border: sf10Border({ top: true, left: true, right: true }),
  });
  mergeAndSetSF10(
    worksheet,
    excelRange(1, startRow + 1, 25, startRow + 1),
    `I CERTIFY that this is a true record of ${[
      learner.lastName,
      learner.firstName,
      learner.middleName,
    ]
      .filter(Boolean)
      .join(
        ", ",
      )} with LRN ${learner.lrn} and that he/she is eligible for admission to Grade ______.`,
    {
      font: { size: 8, bold: true },
      alignment: { horizontal: "center", wrapText: false, shrinkToFit: true },
      border: { left: MEDIUM, right: MEDIUM },
    },
  );
  mergeAndSetSF10(
    worksheet,
    excelRange(1, startRow + 2, 25, startRow + 2),
    `School Name: ${learner.record.schoolName}    School ID: ${learner.record.schoolId}    Division: ${learner.record.division}    Last School Year Attended: ${learner.record.schoolYear}`,
    {
      font: { size: 8, bold: true },
      alignment: { horizontal: "center", wrapText: false, shrinkToFit: true },
      border: { left: MEDIUM, right: MEDIUM },
    },
  );
  mergeAndSetSF10(worksheet, excelRange(3, startRow + 4, 7, startRow + 4), "", {
    border: { bottom: THIN },
  });
  mergeAndSetSF10(
    worksheet,
    excelRange(9, startRow + 4, 17, startRow + 4),
    schoolHead.toUpperCase(),
    {
      font: { size: 8 },
      alignment: { vertical: "bottom" },
      border: { bottom: THIN },
    },
  );
  mergeAndSetSF10(worksheet, excelRange(19, startRow + 3, 24, startRow + 4), "", {
    border: {},
  });
  mergeAndSetSF10(worksheet, excelRange(3, startRow + 5, 7, startRow + 5), "Date", {
    font: { size: 8 },
    border: {},
  });
  mergeAndSetSF10(
    worksheet,
    excelRange(9, startRow + 5, 17, startRow + 5),
    "Signature of Principal/School Head over Printed Name",
    {
      font: { size: 8 },
      border: {},
    },
  );
  mergeAndSetSF10(
    worksheet,
    excelRange(19, startRow + 5, 24, startRow + 5),
    "(Affix School Seal here)",
    {
      font: { size: 7 },
      border: {},
    },
  );
  for (let row = startRow + 1; row <= startRow + 5; row += 1) {
    worksheet.getCell(row, 1).border = {
      ...worksheet.getCell(row, 1).border,
      left: MEDIUM,
    };
    worksheet.getCell(row, 25).border = {
      ...worksheet.getCell(row, 25).border,
      right: MEDIUM,
    };
  }
  styleSF10Range(worksheet, excelRange(1, startRow + 5, 25, startRow + 5), {
    border: { bottom: MEDIUM },
  });

  worksheet.getRow(startRow).height = 18;
  worksheet.getRow(startRow + 1).height = 20;
  worksheet.getRow(startRow + 2).height = 20;
  worksheet.getRow(startRow + 3).height = 8;
  worksheet.getRow(startRow + 4).height = 18;
  worksheet.getRow(startRow + 5).height = 16;
}

function buildSF10FrontSheet(
  workbook: ExcelJS.Workbook,
  learner: SF10LearnerExport,
  index: number,
  imageIds: { depedImageId: number | null; schoolImageId: number | null },
) {
  const worksheet = workbook.addWorksheet(index === 0 ? "Front" : `Front ${index + 1}`);
  configureSF10Worksheet(worksheet, "A1:Y68");
  setFrontHeader(worksheet, imageIds);
  setPersonalInformation(worksheet, learner);
  setEligibilitySection(worksheet);

  mergeAndSetSF10(worksheet, "A14:Y14", "SCHOLASTIC RECORD", {
    font: { size: 10, bold: true },
    fill: GREY_FILL,
    border: sf10Border({ top: true, bottom: true, left: true, right: true }),
  });
  worksheet.getRow(14).height = 18;

  buildSF10RecordBlock(worksheet, 1, 15, learner.record);
  buildSF10RecordBlock(worksheet, 14, 15);
  worksheet.getRow(41).height = 4;
  buildSF10RecordBlock(worksheet, 1, 42);
  buildSF10RecordBlock(worksheet, 14, 42);

  mergeAndSetSF10(worksheet, "W68:Y68", "SFRT 2017", {
    font: { size: 8, italic: true },
    alignment: { horizontal: "right" },
    border: {},
  });
  worksheet.getRow(68).height = 12;
}

function buildSF10BackSheet(
  workbook: ExcelJS.Workbook,
  learner: SF10LearnerExport,
  index: number,
  schoolHead: string,
) {
  const worksheet = workbook.addWorksheet(index === 0 ? "Back" : `Back ${index + 1}`);
  configureSF10Worksheet(worksheet, "A1:Y77");

  mergeAndSetSF10(worksheet, "A1:C1", "SF10-JHS", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left" },
    border: {},
  });
  mergeAndSetSF10(worksheet, "V1:Y1", "Page 2 of ________", {
    font: { size: 8 },
    alignment: { horizontal: "right" },
    border: {},
  });
  mergeAndSetSF10(worksheet, "A2:Y2", "SCHOLASTIC RECORD", {
    font: { size: 10, bold: true },
    fill: GREY_FILL,
    border: sf10Border({ top: true, bottom: true, left: true, right: true }),
  });
  worksheet.getRow(1).height = 14;
  worksheet.getRow(2).height = 18;

  buildSF10RecordBlock(worksheet, 1, 3);
  buildSF10RecordBlock(worksheet, 14, 3);
  worksheet.getRow(29).height = 4;
  buildSF10RecordBlock(worksheet, 1, 30);
  buildSF10RecordBlock(worksheet, 14, 30);

  mergeAndSetSF10(worksheet, "A56:Y56", "For Transfer Out / Elementary School Completer Only", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left" },
    border: {},
  });
  setCertificationBlock(worksheet, 57, learner, schoolHead);
  setCertificationBlock(worksheet, 64, learner, schoolHead);
  setCertificationBlock(worksheet, 71, learner, schoolHead);

  mergeAndSetSF10(worksheet, "A77:M77", "May add Certification Box if needed", {
    font: { size: 7 },
    alignment: { horizontal: "left" },
    border: {},
  });
  mergeAndSetSF10(worksheet, "W77:Y77", "SFRT Revised 2017", {
    font: { size: 7, italic: true },
    alignment: { horizontal: "right" },
    border: {},
  });
}


function setSF10TemplateCell(
  worksheet: Worksheet,
  address: string,
  value: string | number | null,
  numberFormat?: string,
) {
  const cell = worksheet.getCell(address);
  cell.value = value === "" || value == null ? null : value;
  if (numberFormat) cell.numFmt = numberFormat;
}

function clearSF10TemplateCell(worksheet: Worksheet, address: string) {
  worksheet.getCell(address).value = null;
}

function fillSF10TemplateRecord(
  worksheet: Worksheet,
  headerRow: number,
  firstSubjectRow: number,
  averageRow: number,
  record?: SF10Record,
) {
  const values = record
    ? {
        schoolName: record.schoolName,
        schoolId: record.schoolId,
        district: record.district,
        division: record.division,
        region: record.region,
        gradeLevel: record.gradeLevel,
        section: record.section,
        schoolYear: record.schoolYear,
        adviser: record.adviser,
      }
    : {
        schoolName: "",
        schoolId: "",
        district: "",
        division: "",
        region: "",
        gradeLevel: "",
        section: "",
        schoolYear: "",
        adviser: "",
      };

  setSF10TemplateCell(worksheet, `E${headerRow}`, values.schoolName);
  setSF10TemplateCell(worksheet, `U${headerRow}`, values.schoolId, "@");
  setSF10TemplateCell(worksheet, `AD${headerRow}`, values.district);
  setSF10TemplateCell(worksheet, `AP${headerRow}`, values.division);
  setSF10TemplateCell(worksheet, `BB${headerRow}`, values.region);
  setSF10TemplateCell(worksheet, `I${headerRow + 1}`, values.gradeLevel);
  setSF10TemplateCell(worksheet, `N${headerRow + 1}`, values.section);
  setSF10TemplateCell(worksheet, `V${headerRow + 1}`, values.schoolYear);
  setSF10TemplateCell(worksheet, `AI${headerRow + 1}`, values.adviser);
  setSF10TemplateCell(worksheet, `AW${headerRow + 1}`, "");

  for (let index = 0; index < SF10_LEARNING_AREAS.length; index += 1) {
    const row = firstSubjectRow + index;
    const recordRow = record?.rows[index];
    ["U", "Y", "AC", "AG", "AJ", "AP"].forEach((column) => {
      clearSF10TemplateCell(worksheet, `${column}${row}`);
    });

    if (!recordRow) continue;

    setSF10TemplateCell(
      worksheet,
      `U${row}`,
      recordRow.quarterlyRatings[0] == null ? null : Math.round(recordRow.quarterlyRatings[0]!),
      "0",
    );
    setSF10TemplateCell(
      worksheet,
      `Y${row}`,
      recordRow.quarterlyRatings[1] == null ? null : Math.round(recordRow.quarterlyRatings[1]!),
      "0",
    );
    setSF10TemplateCell(
      worksheet,
      `AC${row}`,
      recordRow.quarterlyRatings[2] == null ? null : Math.round(recordRow.quarterlyRatings[2]!),
      "0",
    );
    setSF10TemplateCell(
      worksheet,
      `AG${row}`,
      recordRow.quarterlyRatings[3] == null ? null : Math.round(recordRow.quarterlyRatings[3]!),
      "0",
    );
    setSF10TemplateCell(
      worksheet,
      `AJ${row}`,
      recordRow.finalRating == null ? null : Math.round(recordRow.finalRating),
      "0",
    );
    setSF10TemplateCell(worksheet, `AP${row}`, recordRow.remarks);
  }

  setSF10TemplateCell(
    worksheet,
    `AJ${averageRow}`,
    record?.generalAverage == null ? null : Math.round(record.generalAverage),
    "0",
  );
  setSF10TemplateCell(
    worksheet,
    `AP${averageRow}`,
    record?.generalAverage == null ? "" : record.generalAverage >= 75 ? "PASSED" : "FAILED",
  );
}

function fillSF10TemplateCertification(
  worksheet: Worksheet,
  startRow: number,
  learner: SF10LearnerExport,
  schoolHead: string,
) {
  const learnerName = [learner.lastName, learner.firstName, learner.middleName]
    .filter(Boolean)
    .join(", ");

  setSF10TemplateCell(worksheet, `O${startRow}`, learnerName);
  setSF10TemplateCell(worksheet, `AF${startRow}`, learner.lrn, "@");
  setSF10TemplateCell(worksheet, `BB${startRow}`, "");
  setSF10TemplateCell(worksheet, `H${startRow + 1}`, learner.record.schoolName);
  setSF10TemplateCell(worksheet, `AC${startRow + 1}`, learner.record.schoolId, "@");
  setSF10TemplateCell(worksheet, `AS${startRow + 1}`, learner.record.schoolYear);
  setSF10TemplateCell(worksheet, `S${startRow + 4}`, schoolHead.toUpperCase());
}

async function buildSF10TemplateWorkbook(options: SF10ExportOptions) {
  const response = await fetch("/templates/SF10.xlsx");
  if (!response.ok) {
    throw new Error(`Unable to load SF10 Excel template (${response.status}).`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const front = workbook.getWorksheet("Front") || workbook.worksheets[0];
  const back = workbook.getWorksheet("Back") || workbook.worksheets[2];

  if (!front || !back) {
    throw new Error("The SF10 Excel template must contain Front and Back sheets.");
  }

  const learner = options.learners[0];
  if (!learner) return workbook;

  setSF10TemplateCell(front, "G7", learner.lastName.toUpperCase());
  setSF10TemplateCell(front, "W7", learner.firstName.toUpperCase());
  setSF10TemplateCell(front, "AN7", learner.nameExtension.toUpperCase());
  setSF10TemplateCell(front, "AX7", learner.middleName.toUpperCase());
  setSF10TemplateCell(front, "M8", learner.lrn, "@");
  setSF10TemplateCell(front, "AH8", learner.birthdate);
  setSF10TemplateCell(front, "AV8", learner.sex);

  fillSF10TemplateRecord(front, 21, 26, 40, learner.record);
  fillSF10TemplateRecord(front, 49, 54, 68);
  fillSF10TemplateRecord(back, 3, 8, 22);
  fillSF10TemplateRecord(back, 31, 36, 50);
  fillSF10TemplateRecord(back, 59, 64, 78);

  fillSF10TemplateCertification(front, 79, learner, options.schoolHead);
  fillSF10TemplateCertification(back, 89, learner, options.schoolHead);

  workbook.creator = learner.record.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "School Form 10 (SF10-JHS)";
  workbook.subject = "Learner Permanent Academic Record for Junior High School";

  return workbook;
}

async function buildSF10Workbook(options: SF10ExportOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.learners[0]?.record.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "School Form 10 (SF10-JHS)";
  workbook.subject = "Learner Permanent Academic Record for Elementary School";

  const imageIds = await addSF10Images(workbook);

  options.learners.forEach((learner, index) => {
    buildSF10FrontSheet(workbook, learner, index, imageIds);
    buildSF10BackSheet(workbook, learner, index, options.schoolHead);
  });

  workbook.worksheets.forEach((worksheet) => {
    worksheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.value !== null) {
          cell.font = { ...cell.font, name: "Arial Narrow" };
        }
      });
    });
  });

  return workbook;
}

async function downloadSF10Excel(fileName: string, options: SF10ExportOptions) {
  const workbook = await buildSF10TemplateWorkbook(options);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as ArrayBuffer], {
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

function SF10RecordPanel({ record }: { record?: SF10Record }) {
  const emptyRows = SF10_LEARNING_AREAS.map((learningArea) => ({
    learningArea,
    quarterlyRatings: [null, null, null, null],
    finalRating: null,
    remarks: "",
  }));
  const rows = record?.rows || emptyRows;

  return (
    <div className="border border-black">
      <div className="grid grid-cols-[auto_1fr_auto_80px] items-end gap-x-1 px-1 leading-[10px]">
        <span>School:</span>
        <span className="truncate border-b border-black px-1">{record?.schoolName || ""}</span>
        <span>School ID:</span>
        <span className="border-b border-black text-center">{record?.schoolId || ""}</span>
        <span>District:</span>
        <span className="truncate border-b border-black px-1">{record?.district || ""}</span>
        <span>Region:</span>
        <span className="truncate border-b border-black px-1">{record?.region || ""}</span>
        <span>Grade / Section:</span>
        <span className="truncate border-b border-black px-1">
          {[record?.gradeLevel, record?.section].filter(Boolean).join(" - ")}
        </span>
        <span>School Year:</span>
        <span className="border-b border-black text-center">{record?.schoolYear || ""}</span>
        <span>Adviser/Teacher:</span>
        <span className="truncate border-b border-black px-1">{record?.adviser || ""}</span>
        <span>Signature:</span>
        <span className="border-b border-black">&nbsp;</span>
      </div>

      <table className="mt-1 w-full table-fixed border-collapse text-[7px] leading-[8px]">
        <colgroup>
          <col style={{ width: "45%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "11%" }} />
          <col style={{ width: "20%" }} />
        </colgroup>
        <thead>
          <tr className="bg-[#f2f2f2]">
            <th rowSpan={2} className="border border-black px-1 font-bold">
              LEARNING AREAS
            </th>
            <th colSpan={4} className="border border-black font-bold">
              Quarterly Rating
            </th>
            <th rowSpan={2} className="border border-black font-bold">
              Final Rating
            </th>
            <th rowSpan={2} className="border border-black font-bold">
              Remarks
            </th>
          </tr>
          <tr className="bg-[#f2f2f2]">
            {[1, 2, 3, 4].map((quarter) => (
              <th key={quarter} className="border border-black">
                {quarter}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const isComponent = ["Music", "Arts", "Physical Education", "Health"].includes(
              row.learningArea,
            );
            const isOptional = row.learningArea.startsWith("*");

            return (
              <tr key={row.learningArea} className="h-[9px]">
                <td
                  className={`border border-black px-1 ${
                    isComponent ? "italic" : isOptional ? "" : "font-bold"
                  }`}
                >
                  {row.learningArea}
                </td>
                {row.quarterlyRatings.map((rating, index) => (
                  <td key={index} className="border border-black text-center">
                    {rating == null ? "" : Math.round(rating)}
                  </td>
                ))}
                <td className="border border-black text-center font-bold">
                  {row.finalRating == null ? "" : Math.round(row.finalRating)}
                </td>
                <td className="border border-black text-center text-[6px]">{row.remarks}</td>
              </tr>
            );
          })}
          <tr className="h-[9px] font-bold">
            <td colSpan={5} className="border border-black px-1">
              General Average
            </td>
            <td className="border border-black text-center">
              {record?.generalAverage == null ? "" : Math.round(record.generalAverage)}
            </td>
            <td className="border border-black"></td>
          </tr>
          <tr className="h-[9px] font-bold">
            <td colSpan={2} className="border border-black text-center">
              Remedial Classes
            </td>
            <td colSpan={5} className="border border-black px-1 text-left">
              Date Conducted: __________________ to __________________
            </td>
          </tr>
          <tr className="h-[15px] font-bold">
            <td className="border border-black text-center">Learning Areas</td>
            <td colSpan={2} className="border border-black text-center">
              Final Rating
            </td>
            <td className="border border-black text-center">Remedial Mark</td>
            <td colSpan={2} className="border border-black text-center">
              Recomputed Final Grade
            </td>
            <td className="border border-black text-center">Remarks</td>
          </tr>
          {[0, 1].map((row) => (
            <tr key={row} className="h-[9px]">
              <td className="border border-black">&nbsp;</td>
              <td colSpan={2} className="border border-black"></td>
              <td className="border border-black"></td>
              <td colSpan={2} className="border border-black"></td>
              <td className="border border-black"></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function SF10FrontPage({ learner }: { learner: SF10LearnerExport }) {
  return (
    <section
      className="h-full w-full bg-white p-[18px] text-[8px] text-black"
      style={{
        fontFamily: '"Arial Narrow", Arial, sans-serif',
        pageBreakAfter: "always",
      }}
    >
      <div className="grid grid-cols-[105px_1fr_58px] items-center">
        <img src={depedLogo} alt="DepEd" className="h-[42px] w-[100px] object-contain" />
        <div className="text-center">
          <div className="text-[11px]">Republic of the Philippines</div>
          <div className="text-[11px]">Department of Education</div>
          <div className="mt-1 text-[15px] font-bold">
            Learner Permanent Academic Record for Junior High School (SF10-JHS)
          </div>
          <div className="text-[9px] italic">(Formerly Form 137)</div>
        </div>
        <img src={schoolLogo} alt="School" className="h-[50px] w-[50px] object-contain" />
      </div>
      <div className="-mt-1 font-bold">SF10-JHS</div>

      <div className="mt-1 border border-black bg-[#d9d9d9] py-[2px] text-center text-[10px] font-bold">
        LEARNER&apos;S PERSONAL INFORMATION
      </div>
      <div className="grid grid-cols-[auto_1fr_auto_1fr_auto_60px_auto_1fr] items-end gap-x-1 px-1 leading-[15px]">
        <span>LAST NAME:</span>
        <span className="border-b border-black px-1 font-bold">{learner.lastName}</span>
        <span>FIRST NAME:</span>
        <span className="border-b border-black px-1 font-bold">{learner.firstName}</span>
        <span>NAME EXTN.:</span>
        <span className="border-b border-black text-center">{learner.nameExtension}</span>
        <span>MIDDLE NAME:</span>
        <span className="border-b border-black px-1 font-bold">{learner.middleName}</span>
      </div>
      <div className="grid grid-cols-[auto_1fr_auto_150px_auto_45px] items-end gap-x-1 px-1 leading-[15px]">
        <span>Learner Reference Number (LRN):</span>
        <span className="border-b border-black px-1 text-center font-bold">{learner.lrn}</span>
        <span>Birthdate (mm/dd/yyyy):</span>
        <span className="border-b border-black text-center">{learner.birthdate}</span>
        <span>Sex:</span>
        <span className="border-b border-black text-center">{learner.sex}</span>
      </div>

      <div className="mt-[2px] border border-black bg-[#d9d9d9] py-[2px] text-center text-[10px] font-bold">
        ELIGIBILITY FOR JHS ENROLMENT
      </div>
      <div className="grid grid-cols-4 px-1 italic leading-[14px]">
        <span>Elementary School Completer:</span>
        <span className="text-center">Kinder Progress Report</span>
        <span className="text-center">ECCD Checklist</span>
        <span className="text-center">Kindergarten Certificate of Completion</span>
      </div>
      <div className="grid grid-cols-[auto_1fr_auto_90px_auto_1fr] items-end gap-x-1 px-1 leading-[13px]">
        <span>Name of School:</span>
        <span className="border-b border-black">&nbsp;</span>
        <span>School ID:</span>
        <span className="border-b border-black">&nbsp;</span>
        <span>Address of School:</span>
        <span className="border-b border-black">&nbsp;</span>
      </div>
      <div className="px-1 leading-[13px]">Other Credential Presented</div>
      <div className="grid grid-cols-[auto_100px_auto_150px_auto_1fr] items-end gap-x-1 px-1 leading-[13px]">
        <span>PEPT Passer Rating:</span>
        <span className="border-b border-black">&nbsp;</span>
        <span>Date of Examination/Assessment:</span>
        <span className="border-b border-black">&nbsp;</span>
        <span className="border-b border-black">Others (Please Specify):</span>
      </div>

      <div className="mt-[2px] border border-black bg-[#d9d9d9] py-[2px] text-center text-[10px] font-bold">
        SCHOLASTIC RECORD
      </div>
      <div className="grid grid-cols-2 gap-[3px]">
        <SF10RecordPanel record={learner.record} />
        <SF10RecordPanel />
        <SF10RecordPanel />
        <SF10RecordPanel />
      </div>
      <div className="mt-[1px] text-right text-[7px] italic">SFRT 2017</div>
    </section>
  );
}

function SF10Page() {
  const [classId, setClassId] = useState<string>(readSchoolFormsClassId);

  // SF10 is an individual permanent academic record, so only ONE learner
  // can be selected at a time.
  const [selectedLearnerId, setSelectedLearnerId] = useState<string>("");

  const [length, setLength] = useState<"short" | "full">("full");

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () =>
      (await supabase.from("profiles").select("*").maybeSingle()).data as ProfileRow | null,
  });
  const { data: classes = [] } = useQuery({
    queryKey: ["classes"],
    queryFn: async () =>
      (await supabase.from("classes").select("*").order("created_at", { ascending: false }))
        .data as ClassRow[],
  });
  const selectedSchoolFormsClass = classes.find((item) => item.id === classId);
  const scopedClasses = selectedSchoolFormsClass ? [selectedSchoolFormsClass] : classes;
  const active = selectedSchoolFormsClass?.id || scopedClasses[0]?.id;
  const klass = classes.find((item) => item.id === active);
  const matchingClassIds = useMemo(() => {
    if (!klass) return [] as string[];

    const selectedGrade = normalizeSubject(klass.grade_level || "");
    const selectedSection = normalizeSubject(klass.section || "");
    const selectedSchoolYear = normalizeSubject(klass.school_year || "");

    return classes
      .filter(
        (item) =>
          normalizeSubject(item.grade_level || "") === selectedGrade &&
          normalizeSubject(item.section || "") === selectedSection &&
          normalizeSubject(item.school_year || "") === selectedSchoolYear,
      )
      .map((item) => item.id);
  }, [classes, klass]);
  const { data: students = [] } = useQuery({
    enabled: !!active,
    queryKey: ["students", active],
    queryFn: async () =>
      (await supabase.from("students").select("*").eq("class_id", active!).order("last_name"))
        .data as StudentRow[],
  });

  // Never keep a learner selected after moving to a different class.
  useEffect(() => {
    setSelectedLearnerId("");
  }, [active]);

  // Also clear the selected learner if that learner is no longer in the
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
    enabled: matchingClassIds.length > 0,
    queryKey: ["grades-all-sf10", matchingClassIds.join(",")],
    queryFn: async () =>
      (await supabase.from("grades").select("*").in("class_id", matchingClassIds)).data as GradeRow[],
  });

  const shown = useMemo(
    () =>
      selectedLearnerId
        ? students.filter((student) => student.id === selectedLearnerId).slice(0, 1)
        : [],
    [students, selectedLearnerId],
  );

  const learnerExports = useMemo(
    () =>
      klass
        ? shown.map((student) => toLearnerExport(student, grades, klass, profile))
        : ([] as SF10LearnerExport[]),
    [shown, grades, klass, profile],
  );

  const baseName = `SF10_${klass?.section || "class"}`.replace(/[<>:"/\\|?*]+/g, "_");

  const exportToExcel = async () => {
    if (!klass) {
      toast.error("Please select a class before exporting");
      return;
    }
    if (learnerExports.length === 0) {
      toast.error("Please select one learner");
      return;
    }

    try {
      await downloadSF10Excel(baseName, {
        schoolHead: profile?.principal || "",
        learners: learnerExports,
      });
      toast.success("SF10 Excel spreadsheet downloaded");
    } catch (error) {
      console.error("Unable to export SF10 to Excel", error);
      toast.error("Unable to export SF10 to Excel");
    }
  };

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
          SF10 — Learner&apos;s Permanent Academic Record
        </div>
      </div>

      <div
        className="rounded-2xl border-2 p-4 shadow-sm"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="mb-2 text-sm font-semibold text-amber-900">Form Configuration</div>
        <div className="grid gap-3 md:grid-cols-3">
          <div>
            <label className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-amber-900">
              Class
            </label>
            <Select
              value={active}
              onValueChange={(value) => {
                setClassId(value);
                setSelectedLearnerId("");
              }}
              disabled={Boolean(selectedSchoolFormsClass)}
            >
              <SelectTrigger className="bg-white">
                <SelectValue placeholder="Pick class" />
              </SelectTrigger>
              <SelectContent>
                {scopedClasses.map((item) => (
                  <SelectItem key={item.id} value={item.id}>
                    {item.grade_level} · {item.subject} · {item.section || "—"}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="md:col-span-2">
            <label className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-amber-900">
              Learner
            </label>

            {students.length > 0 ? (
              <Select
                value={selectedLearnerId || undefined}
                onValueChange={setSelectedLearnerId}
              >
                <SelectTrigger className="bg-white">
                  <SelectValue placeholder="Select one learner for SF10" />
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
              <div className="rounded-md border bg-white px-3 py-2 text-sm text-muted-foreground">
                No students yet.
              </div>
            )}

            {selectedLearnerId && (
              <div className="mt-2 text-xs font-medium" style={{ color: DEPED_BLUE }}>
                1 learner selected
              </div>
            )}

            <div className="mt-1 text-xs text-muted-foreground">
              SF10 is an individual learner permanent academic record. Only one
              learner can be selected at a time.
            </div>
          </div>
        </div>
      </div>

      {learnerExports.length > 0 ? (
        <PdfPreviewShell
          fileName={`${baseName}.pdf`}
          docBaseName={baseName}
          printTargetId="sf10-doc"
          length={length}
          onLengthChange={setLength}
          orientation="portrait"
          paper="long"
          hideCopyToWord={
            String(klass?.grade_level ?? "")
            .trim()
            .toLowerCase() === "grade 12"
          }
          onExportToExcel={exportToExcel}
        >
          <div id="sf10-doc" className="bg-white">
            {learnerExports.map((learner, index) => (
              <Fragment key={`${learner.lrn}-${index}`}>
                <SF10FrontPage learner={learner} />
              </Fragment>
            ))}
          </div>
        </PdfPreviewShell>
      ) : (
        <div className="rounded-2xl border border-dashed border-slate-300 bg-white px-6 py-10 text-center text-sm text-slate-500">
          Select one learner to show the SF10 bond-paper format.
        </div>
      )}
    </div>
  );
}
