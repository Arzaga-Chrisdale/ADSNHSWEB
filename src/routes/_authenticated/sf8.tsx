import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useMemo, useState } from "react";
import { ArrowLeft, HeartPulse } from "lucide-react";
import ExcelJS, { type PaperSize } from "exceljs";
import { toast } from "sonner";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import type { ClassRow, StudentRow } from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import schoolLogo from "@/assets/ASNSHS Logo.png";
import depedLogo from "@/assets/deped_logo.png";
const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/sf8")({
  component: SF8Page,
});

type HealthStudent = StudentRow & Record<string, unknown>;

type StudentHealthData = {
  weight: number | null;
  height: number | null;
  heightSquared: number | null;
  bmi: number | null;
  bmiCategory: string;
  heightForAge: string;
  remarks: string;
};

type SummaryCategory = {
  key: string;
  label: string;
};

const BMI_SUMMARY_CATEGORIES: SummaryCategory[] = [
  { key: "severely wasted", label: "Severely Wasted" },
  { key: "wasted", label: "Wasted" },
  { key: "normal", label: "Normal" },
  { key: "overweight", label: "Overweight" },
  { key: "obese", label: "Obese" },
];

const HFA_SUMMARY_CATEGORIES: SummaryCategory[] = [
  { key: "severely stunted", label: "Severely Stunted" },
  { key: "stunted", label: "Stunted" },
  { key: "normal", label: "Normal" },
  { key: "tall", label: "Tall" },
];

function asDate(value?: string | null) {
  if (!value) return null;

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDate(value?: string | null) {
  const date = asDate(value);

  if (!date) return "";

  return date.toLocaleDateString("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  });
}

function ageFrom(birthdate?: string | null, referenceDate?: Date) {
  const birth = asDate(birthdate);

  if (!birth) return "";

  const reference = referenceDate ?? new Date();
  let years = reference.getFullYear() - birth.getFullYear();
  let months = reference.getMonth() - birth.getMonth();

  if (reference.getDate() < birth.getDate()) {
    months -= 1;
  }

  if (months < 0) {
    years -= 1;
    months += 12;
  }

  return months > 0 ? `${years}.${String(months).padStart(2, "0")}` : String(years);
}

function firstText(...values: unknown[]) {
  for (const value of values) {
    const text = String(value ?? "").trim();
    if (text) return text;
  }

  return "";
}

function firstNumber(...values: unknown[]) {
  for (const value of values) {
    if (value === null || value === undefined || value === "") continue;

    const number = Number(value);
    if (Number.isFinite(number)) return number;
  }

  return null;
}

function studentSex(student: StudentRow) {
  const value = String((student as HealthStudent).sex ?? "")
    .trim()
    .toLowerCase();

  if (value === "male" || value === "m") return "male";
  if (value === "female" || value === "f") return "female";
  return "unknown";
}

function studentFullName(student: StudentRow) {
  const record = student as HealthStudent;
  const middleName = firstText(record.middle_name);
  const extension = firstText(
    record.name_extension,
    record.extension_name,
    record.suffix,
  );

  const givenNames = [student.first_name, middleName, extension]
    .filter(Boolean)
    .join(" ");

  return `${student.last_name}, ${givenNames}`.trim();
}

function getStudentHealthData(student: StudentRow): StudentHealthData {
  const record = student as HealthStudent;

  const rawHeight = firstNumber(
    record.height_m,
    record.height,
    record.current_height,
    record.height_cm,
  );
  const height =
    rawHeight !== null && rawHeight > 3 ? rawHeight / 100 : rawHeight;
  const weight = firstNumber(
    record.weight_kg,
    record.weight,
    record.current_weight,
  );
  const heightSquared =
    height !== null && height > 0 ? height * height : null;
  const calculatedBmi =
    weight !== null && heightSquared !== null && heightSquared > 0
      ? weight / heightSquared
      : null;
  const savedBmi = firstNumber(record.bmi, record.body_mass_index);

  return {
    weight,
    height,
    heightSquared,
    bmi: calculatedBmi ?? savedBmi,
    bmiCategory: firstText(
      record.bmi_category,
      record.nutritional_status,
      record.nutrition_status,
    ),
    heightForAge: firstText(
      record.height_for_age,
      record.hfa,
      record.height_age_status,
    ),
    remarks: firstText(record.health_remarks, record.remarks),
  };
}

function numberText(value: number | null, decimalPlaces = 2) {
  if (value === null || !Number.isFinite(value)) return "";
  return value.toFixed(decimalPlaces).replace(/\.00$/, "");
}

function normalizeCategory(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

function countCategory(
  students: StudentRow[],
  category: string,
  field: "bmiCategory" | "heightForAge",
) {
  return students.filter((student) => {
    const health = getStudentHealthData(student);
    return normalizeCategory(health[field]) === normalizeCategory(category);
  }).length;
}

type SF8ExcelOptions = {
  fileName: string;
  schoolName: string;
  district: string;
  division: string;
  region: string;
  schoolId: string;
  gradeLevel: string;
  section: string;
  trackStrand: string;
  schoolYear: string;
  measurementDate: Date;
  maleStudents: StudentRow[];
  femaleStudents: StudentRow[];
  conductedBy: string;
  certifiedBy: string;
  reviewedBy: string;
};

async function downloadSF8Excel(options: SF8ExcelOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "School Form 8";
  workbook.created = new Date();

  const sheet = workbook.addWorksheet("Nutritional Status", {
    pageSetup: {
      paperSize: 9 as PaperSize,
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: {
        left: 0.2,
        right: 0.2,
        top: 0.25,
        bottom: 0.25,
        header: 0,
        footer: 0,
      },
    },
    views: [{ showGridLines: false }],
  });

  sheet.properties.defaultRowHeight = 13;
  sheet.pageSetup.printArea = "B1:P121";
  sheet.pageSetup.printTitlesRow = "9:10";

  const widths = [2, 5, 15, 16, 16, 16, 9, 10, 7, 9, 9, 10, 10, 17, 15, 18];
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });

  const thin = { style: "thin" as const, color: { argb: "FF000000" } };
  const medium = { style: "medium" as const, color: { argb: "FF000000" } };
  const center = { horizontal: "center" as const, vertical: "middle" as const };
  const font = { name: "Arial Narrow", size: 8, color: { argb: "FF000000" } };

  sheet.getCell("B1").value = "SF 8";
  sheet.getCell("B1").font = { ...font, bold: true };

  sheet.mergeCells("B2:P2");
  sheet.getCell("B2").value = "Department of Education";
  sheet.getCell("B2").font = { ...font, size: 11, bold: true };
  sheet.getCell("B2").alignment = center;

  sheet.mergeCells("B3:P3");
  sheet.getCell("B3").value =
    "School Form 8 Learner's Basic Health and Nutrition Report (SF8)";
  sheet.getCell("B3").font = { ...font, size: 11, bold: true };
  sheet.getCell("B3").alignment = center;

  sheet.mergeCells("B4:P4");
  sheet.getCell("B4").value = "(For All Grade Levels)";
  sheet.getCell("B4").font = { ...font, italic: true };
  sheet.getCell("B4").alignment = center;

  const field = (labelRange: string, valueRange: string, label: string, value: string) => {
    sheet.mergeCells(labelRange);
    sheet.mergeCells(valueRange);
    const labelCell = sheet.getCell(labelRange.split(":")[0]);
    const valueCell = sheet.getCell(valueRange.split(":")[0]);
    labelCell.value = label;
    labelCell.font = font;
    labelCell.alignment = { horizontal: "right", vertical: "middle" };
    valueCell.value = value;
    valueCell.font = { ...font, bold: true };
    valueCell.alignment = center;
    valueCell.border = { top: thin, left: thin, bottom: thin, right: thin };
  };

  field("D5:E5", "F5:H5", "School Name", options.schoolName);
  field("I5:I5", "J5:K5", "District", options.district);
  field("L5:L5", "M5:N5", "Division", options.division);
  field("O5:O5", "P5:P5", "Region", options.region);
  field("B7:C7", "D7:E7", "School ID", options.schoolId);
  field("F7:F7", "G7:G7", "Grade", options.gradeLevel);
  field("H7:H7", "I7:J7", "Section", options.section);
  field("K7:L7", "M7:N7", "Track/Strand (SHS)", options.trackStrand);
  field("O7:O7", "P7:P7", "School Year", options.schoolYear);

  sheet.getCell("R7").value = "Date of Measurement:";
  sheet.getCell("R8").value = "(mm/dd/yyyy)";
  sheet.getCell("S7").value = options.measurementDate;
  sheet.getCell("S7").numFmt = "mm/dd/yyyy";

  const headerCells = [
    ["B9:B10", "No."],
    ["C9:C10", "LRN"],
    ["D9:G10", "Learner's Name\n(Last Name, First Name, Name Extension, Middle Name)"],
    ["H9:H10", "Birthdate\n(MM/DD/YYYY)"],
    ["I9:I10", "Age"],
    ["J9:J10", "Weight\n(kg)"],
    ["K9:K10", "Height\n(m)"],
    ["L9:L10", "Height²\n(m²)"],
    ["M9:N9", "Nutritional Status"],
    ["O9:O10", "Height for Age (HFA)"],
    ["P9:P10", "Remarks"],
  ] as const;

  headerCells.forEach(([range, label]) => {
    sheet.mergeCells(range);
    const cell = sheet.getCell(range.split(":")[0]);
    cell.value = label;
    cell.font = { ...font, bold: true };
    cell.alignment = { ...center, wrapText: true };
  });
  sheet.getCell("M10").value = "BMI\n(kg/m²)";
  sheet.getCell("N10").value = "BMI Category";

  for (let row = 9; row <= 10; row += 1) {
    for (let col = 2; col <= 16; col += 1) {
      const cell = sheet.getCell(row, col);
      cell.font = { ...font, bold: true };
      cell.alignment = { ...center, wrapText: true };
      cell.border = { top: medium, left: medium, bottom: medium, right: medium };
    }
  }
  sheet.getRow(9).height = 25;
  sheet.getRow(10).height = 24;

  const writeSection = (
    label: string,
    sectionRow: number,
    startRow: number,
    endRow: number,
    learners: StudentRow[],
  ) => {
    sheet.mergeCells(`B${sectionRow}:C${sectionRow}`);
    sheet.mergeCells(`D${sectionRow}:G${sectionRow}`);
    sheet.getCell(`B${sectionRow}`).value = label;
    sheet.getCell(`D${sectionRow}`).value = "-";

    for (let col = 2; col <= 16; col += 1) {
      const cell = sheet.getCell(sectionRow, col);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7E6E6" } };
      cell.font = { ...font, bold: true, italic: true };
      cell.alignment = col === 2 ? { horizontal: "left", vertical: "middle" } : center;
      cell.border = { top: thin, left: thin, bottom: thin, right: thin };
    }

    for (let row = startRow; row <= endRow; row += 1) {
      const student = learners[row - startRow];
      const health = student ? getStudentHealthData(student) : null;

      if (student) {
        const birthdate = asDate(student.birthdate);
        const values: unknown[] = [
          row - startRow + 1,
          firstText((student as HealthStudent).lrn),
          studentFullName(student),
          birthdate,
          ageFrom(student.birthdate, options.measurementDate),
          health?.weight ?? null,
          health?.height ?? null,
          health?.heightSquared ?? null,
          health?.bmi ?? null,
          health?.bmiCategory ?? "",
          health?.heightForAge ?? "",
          health?.remarks ?? "",
        ];

        sheet.getCell(`B${row}`).value = values[0] as number;
        sheet.getCell(`C${row}`).value = values[1] as string;
        sheet.mergeCells(`D${row}:G${row}`);
        sheet.getCell(`D${row}`).value = values[2] as string;
        sheet.getCell(`H${row}`).value = values[3] as Date | null;
        sheet.getCell(`I${row}`).value = values[4] as string;
        sheet.getCell(`J${row}`).value = values[5] as number | null;
        sheet.getCell(`K${row}`).value = values[6] as number | null;
        sheet.getCell(`L${row}`).value = values[7] as number | null;
        sheet.getCell(`M${row}`).value = values[8] as number | null;
        sheet.getCell(`N${row}`).value = values[9] as string;
        sheet.getCell(`O${row}`).value = values[10] as string;
        sheet.getCell(`P${row}`).value = values[11] as string;
      } else {
        sheet.mergeCells(`D${row}:G${row}`);
      }

      for (let col = 2; col <= 16; col += 1) {
        const cell = sheet.getCell(row, col);
        cell.font = font;
        cell.alignment =
          col === 4 || col === 16
            ? { horizontal: "left", vertical: "middle" }
            : center;
        cell.border = { top: thin, left: thin, bottom: thin, right: thin };
      }

      sheet.getCell(`C${row}`).numFmt = "@";
      sheet.getCell(`H${row}`).numFmt = "mm/dd/yyyy";
      sheet.getCell(`J${row}`).numFmt = "0.00";
      sheet.getCell(`K${row}`).numFmt = "0.000";
      sheet.getCell(`L${row}`).numFmt = "0.0000";
      sheet.getCell(`M${row}`).numFmt = "0.00";
    }
  };

  writeSection("MALE", 11, 12, 63, options.maleStudents);
  writeSection("FEMALE", 64, 65, 108, options.femaleStudents);

  sheet.mergeCells("B110:P110");
  sheet.getCell("B110").value = "SUMMARY TABLE";
  sheet.getCell("B110").font = { ...font, bold: true };
  sheet.getCell("B110").alignment = center;

  sheet.mergeCells("B112:C113");
  sheet.mergeCells("D112:J112");
  sheet.mergeCells("K112:O112");
  sheet.getCell("B112").value = "SEX";
  sheet.getCell("D112").value = "Nutritional Status\nSummary Table";
  sheet.getCell("K112").value = "Height for Age (HFA)\nSummary Table";

  const summaryHeaders = [
    "Severely Wasted",
    "Wasted",
    "Normal",
    "Overweight",
    "Obese",
    "TOTAL",
    "Severely Stunted",
    "Stunted",
    "Normal",
    "Tall",
    "TOTAL",
  ];
  summaryHeaders.forEach((label, index) => {
    sheet.getCell(113, index + 4).value = label;
  });

  const writeSummary = (row: number, label: string, group: StudentRow[]) => {
    sheet.mergeCells(`B${row}:C${row}`);
    sheet.getCell(`B${row}`).value = label;
    const bmiCounts = BMI_SUMMARY_CATEGORIES.map((category) =>
      countCategory(group, category.key, "bmiCategory"),
    );
    const hfaCounts = HFA_SUMMARY_CATEGORIES.map((category) =>
      countCategory(group, category.key, "heightForAge"),
    );
    const values = [
      ...bmiCounts,
      bmiCounts.reduce((sum, value) => sum + value, 0),
      ...hfaCounts,
      hfaCounts.reduce((sum, value) => sum + value, 0),
    ];
    values.forEach((value, index) => {
      sheet.getCell(row, index + 4).value = value;
    });
  };

  writeSummary(114, "MALE", options.maleStudents);
  writeSummary(115, "FEMALE", options.femaleStudents);
  writeSummary(116, "TOTAL", [...options.maleStudents, ...options.femaleStudents]);

  for (let row = 112; row <= 116; row += 1) {
    for (let col = 2; col <= 15; col += 1) {
      const cell = sheet.getCell(row, col);
      cell.font = { ...font, bold: row <= 113 || col <= 3 };
      cell.alignment = { ...center, wrapText: true };
      cell.border = { top: thin, left: thin, bottom: thin, right: thin };
    }
  }

  const signature = (labelCell: string, valueRange: string, label: string, value: string) => {
    sheet.getCell(labelCell).value = label;
    sheet.getCell(labelCell).font = { ...font, bold: true };
    sheet.mergeCells(valueRange);
    const valueCell = sheet.getCell(valueRange.split(":")[0]);
    valueCell.value = value;
    valueCell.font = { ...font, bold: true };
    valueCell.alignment = center;
    valueCell.border = { bottom: thin };
  };
  signature("B118", "B119:E119", "Date of Assessment:", formatDate(options.measurementDate.toISOString()));
  signature("F118", "F119:I119", "Conducted/Assessed By:", options.conductedBy);
  signature("J118", "J119:M119", "Certified Correct By:", options.certifiedBy);
  signature("N118", "N119:P119", "Reviewed By:", options.reviewedBy);

  sheet.getRow(112).height = 22;
  sheet.getRow(113).height = 25;
  sheet.autoFilter = "D11:G108";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${options.fileName}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadSF8TemplateExcel(options: SF8ExcelOptions) {
  const response = await fetch("/templates/sf8-template.xlsx", {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(
      "SF8 template was not found. Put sf8-template.xlsx in public/templates/.",
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());

  const sheet =
    workbook.getWorksheet("Nutritional Status") ?? workbook.worksheets[0];

  if (!sheet) {
    throw new Error("The SF8 template does not contain a worksheet.");
  }

  const setValue = (address: string, value: ExcelJS.CellValue) => {
    sheet.getCell(address).value = value;
  };

  // School and class information. These are the value cells in the official
  // SF8 "Nutritional Status" worksheet.
  setValue("F5", options.schoolName);
  setValue("J5", options.district);
  setValue("M5", options.division);
  setValue("P5", options.region);
  setValue("D7", options.schoolId);
  setValue("G7", options.gradeLevel);
  setValue("I7", options.section);
  setValue("M7", options.trackStrand);
  setValue("P7", options.schoolYear);
  setValue("S8", options.measurementDate);
  sheet.getCell("S8").numFmt = "mm/dd/yyyy";

  const clearLearnerRows = (startRow: number, endRow: number) => {
    for (let row = startRow; row <= endRow; row += 1) {
      // Only clear the editable fields. Columns B, I and L:O contain the
      // template's numbering, age, BMI and classification formulas.
      for (const column of ["C", "D", "H", "J", "K", "P"]) {
        sheet.getCell(`${column}${row}`).value = null;
      }
    }
  };

  const writeLearners = (
    learners: StudentRow[],
    startRow: number,
    endRow: number,
  ) => {
    learners.slice(0, endRow - startRow + 1).forEach((student, index) => {
      const row = startRow + index;
      const health = getStudentHealthData(student);
      const birthdate = asDate(student.birthdate);

      setValue(`C${row}`, firstText((student as HealthStudent).lrn));
      setValue(`D${row}`, studentFullName(student));
      setValue(`H${row}`, birthdate);
      setValue(`J${row}`, health.weight);
      setValue(`K${row}`, health.height);
      setValue(`P${row}`, health.remarks);

      sheet.getCell(`C${row}`).numFmt = "@";
      sheet.getCell(`H${row}`).numFmt = "mm/dd/yyyy";
      sheet.getCell(`J${row}`).numFmt = "0.00";
      sheet.getCell(`K${row}`).numFmt = "0.000";
    });
  };

  clearLearnerRows(12, 63);
  clearLearnerRows(65, 108);
  writeLearners(options.maleStudents, 12, 63);
  writeLearners(options.femaleStudents, 65, 108);

  // Keep rows 114:116 untouched. Their shared COUNTIF/SUM formulas calculate
  // the Male, Female and Total summaries from the learner rows.

  setValue("B119", formatDate(options.measurementDate.toISOString()));
  setValue("F119", options.conductedBy);
  setValue("J119", options.certifiedBy);
  setValue("N119", options.reviewedBy);

  workbook.creator = "School Form 8";
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${options.fileName}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function SF8Page() {
  const [classId, setClassId] = useState<string>(readSchoolFormsClassId);
  const [length, setLength] = useState<"short" | "full">("full");

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () =>
      (await supabase.from("profiles").select("*").maybeSingle()).data as any,
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
  const active = selectedSchoolFormsClass?.id || scopedClasses[0]?.id;
  const klass = classes.find((classRow) => classRow.id === active);

  const { data: students = [] } = useQuery({
    enabled: Boolean(active),
    queryKey: ["students", active],
    queryFn: async () =>
      (
        await supabase
          .from("students")
          .select("*")
          .eq("class_id", active!)
          .order("last_name")
          .order("first_name")
      ).data as StudentRow[],
  });

  const classRecord = (klass ?? {}) as ClassRow & Record<string, unknown>;
  const measurementDateValue = firstText(
    classRecord.measurement_date,
    classRecord.date_of_measurement,
    new Date().toISOString(),
  );
  const measurementDate = asDate(measurementDateValue) ?? new Date();

  const maleStudents = useMemo(
    () => students.filter((student) => studentSex(student) === "male"),
    [students],
  );
  const femaleStudents = useMemo(
    () => students.filter((student) => studentSex(student) === "female"),
    [students],
  );
  const unspecifiedStudents = useMemo(
    () => students.filter((student) => studentSex(student) === "unknown"),
    [students],
  );

  const schoolName =
    profile?.school_name || "Agusan del Sur National Science High School";
  const district = firstText(profile?.district, profile?.school_district);
  const division = firstText(profile?.division, profile?.schools_division);
  const region = firstText(profile?.region);
  const schoolId = firstText(profile?.school_id, profile?.school_code);
  // Fitness Test in classes.$classId.tsx saves weight_kg and height_m
  // on the same student records used by SF8. Track also supports track_shs.
  const trackStrand = firstText(
    classRecord.track_shs,
    classRecord.track_strand,
    classRecord.track,
    classRecord.strand,
  );
  const conductedBy = firstText(profile?.full_name, profile?.name);
  const certifiedBy = firstText(
    profile?.principal_name,
    profile?.school_head,
  );
  const reviewedBy = firstText(
    profile?.reviewed_by,
    profile?.school_nurse,
  );

  const renderStudentRow = (
    student: StudentRow | null,
    index: number,
    keyPrefix: string,
  ) => {
    const health = student ? getStudentHealthData(student) : null;

    return (
      <tr
        key={`${keyPrefix}-${student?.id ?? `blank-${index}`}`}
        className="h-[13px]"
      >
        <td className="border border-black px-0.5 text-center">
          {student ? index + 1 : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {student ? firstText((student as HealthStudent).lrn) : ""}
        </td>
        <td className="border border-black px-1 text-left font-medium">
          {student ? studentFullName(student) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {student ? formatDate(student.birthdate) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {student ? ageFrom(student.birthdate, measurementDate) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health ? numberText(health.weight, 2) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health ? numberText(health.height, 3) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health ? numberText(health.heightSquared, 4) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health ? numberText(health.bmi, 2) : ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health?.bmiCategory || ""}
        </td>
        <td className="border border-black px-0.5 text-center">
          {health?.heightForAge || ""}
        </td>
        <td className="border border-black px-0.5 text-left">
          {health?.remarks || ""}
        </td>
      </tr>
    );
  };

  const renderStudentSection = (
    label: string,
    sectionStudents: StudentRow[],
    keyPrefix: string,
  ) => {
    // Do not print a section or placeholder rows when no learner exists.
    if (sectionStudents.length === 0) return null;

    return (
      <>
        <tr className="h-[15px] bg-[#e7e6e6] font-bold italic">
          <td className="border border-black px-1 text-left">{label}</td>
          <td className="border border-black" />
          <td className="border border-black px-1 text-center">-</td>
          {Array.from({ length: 9 }).map((_, index) => (
            <td key={`${keyPrefix}-label-${index}`} className="border border-black" />
          ))}
        </tr>

        {sectionStudents.map((student, index) =>
          renderStudentRow(student, index, keyPrefix),
        )}
      </>
    );
  };

  const summaryRow = (label: string, group: StudentRow[]) => {
    const bmiCounts = BMI_SUMMARY_CATEGORIES.map((category) =>
      countCategory(group, category.key, "bmiCategory"),
    );
    const hfaCounts = HFA_SUMMARY_CATEGORIES.map((category) =>
      countCategory(group, category.key, "heightForAge"),
    );

    return (
      <tr>
        <td className="border border-black px-1 py-1 font-bold">{label}</td>
        {bmiCounts.map((count, index) => (
          <td
            key={`${label}-bmi-${BMI_SUMMARY_CATEGORIES[index].key}`}
            className="border border-black px-1 py-1 text-center"
          >
            {count}
          </td>
        ))}
        <td className="border border-black px-1 py-1 text-center font-bold">
          {bmiCounts.reduce((sum, count) => sum + count, 0)}
        </td>
        {hfaCounts.map((count, index) => (
          <td
            key={`${label}-hfa-${HFA_SUMMARY_CATEGORIES[index].key}`}
            className="border border-black px-1 py-1 text-center"
          >
            {count}
          </td>
        ))}
        <td className="border border-black px-1 py-1 text-center font-bold">
          {hfaCounts.reduce((sum, count) => sum + count, 0)}
        </td>
      </tr>
    );
  };

  const combinedStudents = [...maleStudents, ...femaleStudents, ...unspecifiedStudents];

  const exportToExcel = async () => {
    if (!klass) {
      toast.error("Please select a class first");
      return;
    }

    try {
      const baseName = `SF8_Health_${klass.section || "class"}`.replace(
        /[<>:"/\\|?*]+/g,
        "_",
      );

      await downloadSF8TemplateExcel({
        fileName: baseName,
        schoolName,
        district,
        division,
        region,
        schoolId,
        gradeLevel: firstText(klass.grade_level),
        section: firstText(klass.section),
        trackStrand,
        schoolYear: firstText(klass.school_year),
        measurementDate,
        maleStudents,
        femaleStudents,
        conductedBy,
        certifiedBy,
        reviewedBy,
      });

      toast.success("SF8 exported as a valid .xlsx workbook");
    } catch (error) {
      console.error("SF8 Excel export failed", error);
      toast.error("Unable to export SF8 to Excel");
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
          <HeartPulse className="size-5" style={{ color: DEPED_BLUE }} />
          School Form 8 — Learner&apos;s Health &amp; Nutrition Profile
        </div>
      </div>

      <div
        className="rounded-2xl border-2 p-4 shadow-sm"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="mb-2 text-sm font-semibold text-amber-900">
          Form Configuration
        </div>

        <div className="grid gap-3 md:grid-cols-3">
          <div>
            <label className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-amber-900">
              Class
            </label>

            <Select
              value={active}
              onValueChange={setClassId}
              disabled={Boolean(selectedSchoolFormsClass)}
            >
              <SelectTrigger className="bg-white">
                <SelectValue placeholder="Pick class" />
              </SelectTrigger>
              <SelectContent>
                {scopedClasses.map((classRow) => (
                  <SelectItem key={classRow.id} value={classRow.id}>
                    {classRow.grade_level} · {classRow.subject} · {classRow.section || "—"}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>
      </div>

      <PdfPreviewShell
        fileName={`SF8_Health_${klass?.section || "class"}.pdf`}
        docBaseName={`SF8_Health_${klass?.section || "class"}`}
        printTargetId="sf8-doc"
        length={length}
        onLengthChange={setLength}
        onExportToExcel={exportToExcel}
      >
        <div
          id="sf8-doc"
          className="bg-white p-[10px] text-black"
          style={{
            width: 1020,
            minHeight: length === "full" ? 1560 : 1320,
            fontFamily: "Arial Narrow, Arial, Helvetica, sans-serif",
            fontSize: 7,
            lineHeight: 1.05,
          }}
        >
          <div className="text-left text-[7px] font-bold">SF 8</div>

         <div className="mt-1 grid grid-cols-[110px_1fr_110px] items-center">
  <div className="flex justify-center">
    <img
      src={schoolLogo}
      alt="School Logo"
      className="h-[78px] w-[78px] object-contain"
    />
  </div>

  <div className="text-center">
    <div className="text-[14px] font-bold">
      Department of Education
    </div>

    <div className="mt-0.5 text-[14px] font-bold">
      School Form 8 Learner&apos;s Basic Health and Nutrition Report
      (SF8)
    </div>

    <div className="mt-0.5 text-[10px] italic">
      (For All Grade Levels)
    </div>
  </div>

  <div className="flex justify-center">
    <img
      src={depedLogo}
      alt="DepEd Logo"
      className="h-[62px] w-[100px] object-contain"
    />
  </div>
</div>

          <div className="mt-3 grid grid-cols-[95px_1.35fr_75px_0.85fr_75px_0.85fr_55px_0.45fr] items-end gap-x-1 text-[9px]">
            <div className="text-right">School Name</div>
            <div className="border border-black px-2 py-1 text-center font-bold">
              {schoolName}
            </div>
            <div className="text-right">District</div>
            <div className="border border-black px-2 py-1 text-center font-bold">
              {district}
            </div>
            <div className="text-right">Division</div>
            <div className="border border-black px-2 py-1 text-center font-bold">
              {division}
            </div>
            <div className="text-right">Region</div>
            <div className="border border-black px-2 py-1 text-center font-bold">
              {region}
            </div>
          </div>

          <div className="mt-2 grid grid-cols-[70px_85px_45px_75px_50px_130px_95px_105px_105px_95px] items-center gap-x-1 text-[9px]">
            <div className="text-right">School ID</div>
            <div className="border border-black px-1 py-1 text-center font-bold">
              {schoolId}
            </div>
            <div className="text-right">Grade</div>
            <div className="border border-black px-1 py-1 text-center font-bold">
              {klass?.grade_level || ""}
            </div>
            <div className="text-right">Section</div>
            <div className="border border-black px-1 py-1 text-center font-bold">
              {klass?.section || ""}
            </div>
            <div className="text-right">Track/Strand (SHS)</div>
            <div className="border border-black px-1 py-1 text-center font-bold">
              {trackStrand}
            </div>
            <div className="text-right">School Year</div>
            <div className="border border-black px-1 py-1 text-center font-bold">
              {klass?.school_year || ""}
            </div>
          </div>

          <div className="mt-2 flex items-center justify-end gap-2 text-[9px]">
            <div className="text-right leading-tight">
              Date of Measurement:
              <br />
              (mm/dd/yyyy)
            </div>
            <div className="min-w-[110px] border border-black px-2 py-1 text-center font-bold">
              {formatDate(measurementDateValue)}
            </div>
          </div>

          <table className="mt-2 w-full table-fixed border-collapse text-[6.8px]">
            <colgroup>
              <col style={{ width: "3%" }} />
              <col style={{ width: "11%" }} />
              <col style={{ width: "25%" }} />
              <col style={{ width: "9%" }} />
              <col style={{ width: "5%" }} />
              <col style={{ width: "6%" }} />
              <col style={{ width: "6%" }} />
              <col style={{ width: "6%" }} />
              <col style={{ width: "6%" }} />
              <col style={{ width: "9%" }} />
              <col style={{ width: "8%" }} />
              <col style={{ width: "6%" }} />
            </colgroup>

            <thead className="font-bold">
              <tr className="h-[26px]">
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  No.
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  LRN
                </th>
                <th rowSpan={2} className="border-2 border-black px-1 text-center">
                  Learner&apos;s Name
                  <br />
                  <span className="font-normal">
                    (Last Name, First Name, Name Extension, Middle Name)
                  </span>
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Birthdate
                  <br />
                  (MM/DD/YYYY)
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Age
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Weight
                  <br />
                  (kg)
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Height
                  <br />
                  (m)
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Height²
                  <br />
                  (m²)
                </th>
                <th colSpan={2} className="border-2 border-black px-0.5 text-center">
                  Nutritional Status
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Height for Age
                  <br />
                  (HFA)
                </th>
                <th rowSpan={2} className="border-2 border-black px-0.5 text-center">
                  Remarks
                </th>
              </tr>
              <tr className="h-[22px]">
                <th className="border-2 border-black px-0.5 text-center">
                  BMI
                  <br />
                  (kg/m²)
                </th>
                <th className="border-2 border-black px-0.5 text-center">
                  BMI Category
                </th>
              </tr>
            </thead>

            <tbody>
              {renderStudentSection("MALE", maleStudents, "male")}
              {renderStudentSection("FEMALE", femaleStudents, "female")}
              {unspecifiedStudents.length > 0 &&
                renderStudentSection(
                  "UNSPECIFIED",
                  unspecifiedStudents,
                  "unspecified",
                )}
            </tbody>
          </table>

          <div className="mt-3 text-center text-[9px] font-bold">SUMMARY TABLE</div>

          <table className="mt-1 w-full table-fixed border-collapse text-[6.8px]">
            <thead className="font-bold">
              <tr>
                <th rowSpan={2} className="border-2 border-black px-1 py-1 text-center">
                  SEX
                </th>
                <th colSpan={6} className="border-2 border-black px-1 py-1 text-center">
                  Nutritional Status Summary Table
                </th>
                <th colSpan={5} className="border-2 border-black px-1 py-1 text-center">
                  Height for Age (HFA) Summary Table
                </th>
              </tr>
              <tr>
                {BMI_SUMMARY_CATEGORIES.map((category) => (
                  <th key={category.key} className="border border-black px-1 py-1 text-center">
                    {category.label}
                  </th>
                ))}
                <th className="border border-black px-1 py-1 text-center">TOTAL</th>
                {HFA_SUMMARY_CATEGORIES.map((category) => (
                  <th key={category.key} className="border border-black px-1 py-1 text-center">
                    {category.label}
                  </th>
                ))}
                <th className="border border-black px-1 py-1 text-center">TOTAL</th>
              </tr>
            </thead>
            <tbody>
              {summaryRow("MALE", maleStudents)}
              {summaryRow("FEMALE", femaleStudents)}
              {unspecifiedStudents.length > 0 &&
                summaryRow("UNSPECIFIED", unspecifiedStudents)}
              {summaryRow("TOTAL", combinedStudents)}
            </tbody>
          </table>

          <div className="mt-4 grid grid-cols-4 gap-6 text-[8px]">
            <SignatureBlock
              label="Date of Assessment:"
              value={formatDate(measurementDateValue)}
              helper="Date"
            />
            <SignatureBlock
              label="Conducted/Assessed By:"
              value={conductedBy}
              helper="Name and Signature"
            />
            <SignatureBlock
              label="Certified Correct By:"
              value={certifiedBy}
              helper="School Head"
            />
            <SignatureBlock
              label="Reviewed By:"
              value={reviewedBy}
              helper="School Health Personnel"
            />
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}

function SignatureBlock({
  label,
  value,
  helper,
}: {
  label: string;
  value: string;
  helper: string;
}) {
  return (
    <div>
      <div className="font-bold">{label}</div>
      <div className="mt-5 min-h-[18px] border-b border-black px-1 text-center font-bold">
        {value}
      </div>
      <div className="mt-1 text-center text-[7px] italic">{helper}</div>
    </div>
  );
}
