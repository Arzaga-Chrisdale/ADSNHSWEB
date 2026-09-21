import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { Fragment, useMemo, useState } from "react";
import { ArrowLeft, BadgeCheck } from "lucide-react";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  SUBJECTS,
  computeAverage,
  type ClassRow,
  type GradeRow,
  type StudentRow,
} from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import depedLogo from "@/assets/deped_logo.png";
import { toast } from "sonner";
import ExcelJS, {
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

export const Route = createFileRoute("/_authenticated/sf6")({
  component: SF6Page,
});

type SexCounts = {
  male: number;
  female: number;
  total: number;
};

type LearnerSummary = {
  st: StudentRow;
  avg: number | null;
};

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

const GRADE_LEVELS = [
  { value: "7", label: "GRADE 7" },
  { value: "8", label: "GRADE 8" },
  { value: "9", label: "GRADE 9" },
  { value: "10", label: "GRADE 10" },
  { value: "11", label: "GRADE 11" },
  { value: "12", label: "GRADE 12" },
] as const;

const PROGRESS_BANDS = [
  {
    key: "did-not-meet",
    label: "Did Not Meet Expectations",
    detail: "(74% and below)",
    matches: (average: number) => average < 75,
  },
  {
    key: "fairly-satisfactory",
    label: "Fairly Satisfactory",
    detail: "(75%–79%)",
    matches: (average: number) => average >= 75 && average <= 79,
  },
  {
    key: "satisfactory",
    label: "Satisfactory",
    detail: "(80%–84%)",
    matches: (average: number) => average >= 80 && average <= 84,
  },
  {
    key: "very-satisfactory",
    label: "Very Satisfactory",
    detail: "(85%–89%)",
    matches: (average: number) => average >= 85 && average <= 89,
  },
  {
    key: "outstanding",
    label: "Outstanding",
    detail: "(90%–100%)",
    matches: (average: number) => average >= 90,
  },
] as const;

const EMPTY_COUNTS: SexCounts = { male: 0, female: 0, total: 0 };

const SF6_EXCEL_LOGO_DATA_URL =
  "data:image/gif;base64,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";

type SF6ExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeSection: string;
  statusRows: {
    label: string;
    values: SexCounts[];
    total: SexCounts;
  }[];
  bandValues: SexCounts[][];
  enrolledByGrade: SexCounts[];
  schoolHead: string;
};

function getGradeNumber(gradeLevel?: string | null) {
  return gradeLevel?.match(/\d+/)?.[0] || "";
}

async function imageAssetToDataUrl(assetUrl: string) {
  if (assetUrl.startsWith("data:")) return assetUrl;

  const response = await fetch(assetUrl);
  if (!response.ok) {
    throw new Error(`Unable to load the Excel logo (${response.status})`);
  }

  const blob = await response.blob();

  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("Unable to read the Excel logo"));
    reader.readAsDataURL(blob);
  });
}

function addCounts(values: SexCounts[]): SexCounts {
  return values.reduce(
    (total, value) => ({
      male: total.male + value.male,
      female: total.female + value.female,
      total: total.total + value.total,
    }),
    { ...EMPTY_COUNTS },
  );
}

const BLACK = { argb: "FF000000" };
const THIN_BORDER = { style: "thin" as const, color: BLACK };
const MEDIUM_BORDER = { style: "medium" as const, color: BLACK };
const BODY_FONT: Partial<Font> = { name: "SansSerif", size: 7 };
const CENTERED: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};
const NO_FILL: Fill = { type: "pattern", pattern: "none" };
const SF6_COLUMN_WIDTHS = [
  0.16, 13.09, 3.18, 0.33, 5.71, 5.71, 4.69, 1, 0.16, 5.53, 3.02, 2.68, 5.37, 0.33, 0.33, 5.37,
  5.71, 5.71, 1.33, 4.36, 3.02, 2.68, 5.71, 5.71, 5.71, 5.71, 1.33, 4.02, 0.33, 5.71, 5.71, 5.03,
  5.03, 5.03,
] as const;

const SF6_GRADE_GROUPS = [
  {
    label: "GRADE 7",
    headerRange: "E5:H5",
    countRanges: ["E", "F", "G:H"],
  },
  {
    label: "GRADE 8",
    headerRange: "I5:N5",
    countRanges: ["I:J", "K:L", "M:N"],
  },
  {
    label: "GRADE 9",
    headerRange: "O5:R5",
    countRanges: ["O:P", "Q", "R"],
  },
  {
    label: "GRADE 10",
    headerRange: "S5:W5",
    countRanges: ["S:T", "U:V", "W"],
  },
  {
    label: "GRADE 11",
    headerRange: "X5:Z5",
    countRanges: ["X", "Y", "Z"],
  },
  {
    label: "GRADE 12",
    headerRange: "AA5:AE5",
    countRanges: ["AA:AC", "AD", "AE"],
  },
] as const;

const SF6_TOTAL_GROUP = {
  label: "TOTAL",
  headerRange: "AF5:AH5",
  countRanges: ["AF", "AG", "AH"],
} as const;

const SF6_ALL_GROUPS = [...SF6_GRADE_GROUPS, SF6_TOTAL_GROUP] as const;

function tableBorder({
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

function setCountRange(
  worksheet: Worksheet,
  row: number,
  range: string,
  value: number,
  strong = false,
  options: {
    top?: boolean;
    bottom?: boolean;
    right?: boolean;
  } = {},
) {
  mergeAndSet(worksheet, rowRange(range, row), value, {
    font: { size: 7, bold: strong },
    numberFormat: "0",
    border: tableBorder(options),
  });
}

function setCountGroup(
  worksheet: Worksheet,
  row: number,
  ranges: readonly string[],
  value: SexCounts,
  strong = false,
  options: {
    top?: boolean;
    bottom?: boolean;
    right?: boolean;
  } = {},
) {
  [value.male, value.female, value.total].forEach((count, index) => {
    setCountRange(worksheet, row, ranges[index], count, strong, {
      ...options,
      right: options.right && index === 2,
    });
  });
}

async function buildSF6Workbook(options: SF6ExportOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject = "Summarized Report on Promotion and Learning Progress & Achievement";
  workbook.title = "School Form 6 (SF6)";

  const worksheet = workbook.addWorksheet("SF6", {
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
      horizontalCentered: true,
      verticalCentered: false,
      pageOrder: "overThenDown",
      blackAndWhite: false,
      printArea: "A1:AH20",
      margins: {
        left: 0.28,
        right: 0.28,
        top: 0.28,
        bottom: 0.28,
        header: 0.15,
        footer: 0.15,
      },
    },
  });

  worksheet.properties.defaultRowHeight = 20;
  SF6_COLUMN_WIDTHS.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  mergeAndSet(worksheet, "A1:B3", "", { border: {} });
  let excelLogoDataUrl = SF6_EXCEL_LOGO_DATA_URL;

  try {
    excelLogoDataUrl = await imageAssetToDataUrl(depedLogo);
  } catch (error) {
    console.warn("Using the fallback SF6 Excel logo", error);
  }

  const logoId = workbook.addImage({
    base64: excelLogoDataUrl,
    extension: excelLogoDataUrl.startsWith("data:image/png") ? "png" : "gif",
  });
  worksheet.addImage(logoId, {
    tl: { col: 0.08, row: 0.04 },
    ext: { width: 112, height: 112 },
    editAs: "oneCell",
  });

  mergeAndSet(
    worksheet,
    "D1:AH1",
    "School Form 6 (SF6)\nSummarized Report on Promotion and Learning Progress & Achievement",
    {
      font: { size: 15, bold: true },
      alignment: {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      },
      border: {},
    },
  );

  const metadata = [
    ["E2:H2", "School ID", true],
    ["I2:M2", options.schoolId, false],
    ["P2:R2", options.region, false],
    ["S2:W2", "Division", true],
    ["X2:Z2", options.division, false],
    ["E4:H4", "School Name", true],
    ["I4:R4", options.schoolName, false],
    ["S4:W4", "District", true],
    ["X4:Z4", options.district, false],
    ["AA4:AE4", "School Year", true],
    ["AF4:AH4", options.schoolYear, false],
  ] as const;

  metadata.forEach(([range, value, isLabel]) => {
  const isSchoolName = range === "I4:R4";

  mergeAndSet(worksheet, range, value, {
    font: { size: 9 },
    alignment: {
      horizontal: isLabel ? "right" : "left",
      vertical: "middle",
      wrapText: !isSchoolName,
      shrinkToFit: isSchoolName,
    },
    border: isLabel ? {} : tableBorder(),
  });
});

  mergeAndSet(worksheet, "A5:D6", "SUMMARY TABLE", {
    font: { size: 9 },
    border: tableBorder({ top: true, left: true }),
  });

  SF6_ALL_GROUPS.forEach((group, groupIndex) => {
    mergeAndSet(worksheet, group.headerRange, group.label, {
      font: {
        size: groupIndex === SF6_ALL_GROUPS.length - 1 ? 11 : 9,
      },
      border: tableBorder({
        top: true,
        right: groupIndex === SF6_ALL_GROUPS.length - 1,
      }),
    });

    ["Male", "Female", "Total"].forEach((sexLabel, sexIndex) => {
      mergeAndSet(worksheet, rowRange(group.countRanges[sexIndex], 6), sexLabel, {
        font: { size: 7 },
        border: tableBorder({
          right: groupIndex === SF6_ALL_GROUPS.length - 1 && sexIndex === 2,
        }),
      });
    });
  });

  options.statusRows.forEach((status, statusIndex) => {
    const rowNumber = 7 + statusIndex;
    mergeAndSet(worksheet, `A${rowNumber}:D${rowNumber}`, status.label, {
      font: { size: 9 },
      alignment: { horizontal: "left", vertical: "middle" },
      border: tableBorder({ left: true }),
    });

    status.values.forEach((value, gradeIndex) => {
      setCountGroup(worksheet, rowNumber, SF6_GRADE_GROUPS[gradeIndex].countRanges, value);
    });
    setCountGroup(worksheet, rowNumber, SF6_TOTAL_GROUP.countRanges, status.total, true, {
      right: true,
    });
  });

  mergeAndSet(worksheet, "A10:D10", "LEARNING PROGRESS AND\nACHIEVEMENT", {
    font: { size: 5, bold: false },
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
    border: tableBorder({ left: true }),
  });

  SF6_ALL_GROUPS.forEach((group, groupIndex) => {
    ["Male", "Female", "Total"].forEach((sexLabel, sexIndex) => {
      mergeAndSet(worksheet, rowRange(group.countRanges[sexIndex], 10), sexLabel, {
        font: { size: 7 },
        border: tableBorder({
          right: groupIndex === SF6_ALL_GROUPS.length - 1 && sexIndex === 2,
        }),
      });
    });
  });

  PROGRESS_BANDS.forEach((band, bandIndex) => {
    const rowNumber = 11 + bandIndex;
    const values = options.bandValues[bandIndex];
    const separator = bandIndex === 2 || bandIndex === 4 ? "\n" : " ";
    mergeAndSet(
      worksheet,
      `A${rowNumber}:D${rowNumber}`,
      `${band.label}${separator}${band.detail}`,
      {
        font: { size: 5 },
        alignment: { horizontal: "left", vertical: "middle", wrapText: true },
        border: tableBorder({ left: true }),
      },
    );

    values.forEach((value, gradeIndex) => {
      setCountGroup(worksheet, rowNumber, SF6_GRADE_GROUPS[gradeIndex].countRanges, value);
    });
    setCountGroup(worksheet, rowNumber, SF6_TOTAL_GROUP.countRanges, addCounts(values), true, {
      right: true,
    });
  });

  mergeAndSet(worksheet, "A16:D16", "TOTAL", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: tableBorder({ top: true, bottom: true, left: true }),
  });
  options.enrolledByGrade.forEach((value, gradeIndex) => {
    setCountGroup(worksheet, 16, SF6_GRADE_GROUPS[gradeIndex].countRanges, value, true, {
      top: true,
      bottom: true,
    });
  });
  setCountGroup(
    worksheet,
    16,
    SF6_TOTAL_GROUP.countRanges,
    addCounts(options.enrolledByGrade),
    true,
    { top: true, bottom: true, right: true },
  );

  mergeAndSet(worksheet, "A17:E17", "Prepared and Submitted by:", {
    font: { size: 7 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  const signatures = [
    ["A18:G18", "GERMAN CORDERO TALAGAON"],
    ["L18:S18", ""],
    ["V18:AA18", ""],
    ["AC18:AH18", ""],
  ] as const;
  signatures.forEach(([range, value]) => {
    mergeAndSet(worksheet, range, value, {
      font: { size: range === "A18:G18" ? 9 : 8 },
      alignment: { horizontal: "center", vertical: "bottom" },
      border: { bottom: MEDIUM_BORDER },
    });
  });

  const signatureLabels = [
    ["A19:G19", "(Signature of School Head/SCC Chair)"],
    ["L19:S19", "SCC-Vice Chair (Curriculum)"],
    ["V19:AA19", "SCC Member"],
    ["AC19:AH19", "SCC-Vice Chair (Generated thru LIS)"],
  ] as const;
  signatureLabels.forEach(([range, value]) => {
    mergeAndSet(worksheet, range, value, {
      font: { size: 7 },
      alignment: { horizontal: "center", vertical: "top", wrapText: true },
      border: {},
    });
  });

  mergeAndSet(worksheet, "B20:I20", "(Additional slots may be added for SCC members.)", {
    font: { size: 7 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });

  [50, 20, 4, 20, 25, 25, 25, 25, 25, 30, 30, 30, 30, 30, 30, 25, 20, 20, 20, 20].forEach(
    (height, index) => {
      worksheet.getRow(index + 1).height = height;
    },
  );

  // Keep every populated Excel cell in SansSerif while preserving its size,
  // weight, italic, underline, and other existing font settings.
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.value !== null) {
        cell.font = {
          ...cell.font,
          name: "SansSerif",
        };
      }
    });
  });

  return workbook;
}

async function downloadExcelWorkbook(fileName: string, options: SF6ExportOptions) {
  const workbook = await buildSF6Workbook(options);
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

function CountCells({ value, strong = false }: { value: SexCounts; strong?: boolean }) {
  return (
    <>
      <td
        className={`border border-black px-1 py-2 text-center tabular-nums ${strong ? "font-bold" : ""}`}
        style={{ borderLeftWidth: 2 }}
      >
        {value.male}
      </td>
      <td
        className={`border border-black px-1 py-2 text-center tabular-nums ${strong ? "font-bold" : ""}`}
      >
        {value.female}
      </td>
      <td
        className={`border border-black px-1 py-2 text-center tabular-nums ${strong ? "font-bold" : ""}`}
      >
        {value.total}
      </td>
    </>
  );
}

function MetadataField({
  label,
  value,
  wide = false,
}: {
  label: string;
  value: string;
  wide?: boolean;
}) {
  return (
    <div className={`flex items-end gap-1 ${wide ? "col-span-2" : ""}`}>
      <span className="shrink-0 text-[10px]">{label}</span>
      <span className="min-h-6 flex-1 border border-black px-1.5 py-1 text-[10px]">{value}</span>
    </div>
  );
}

function SF6Page() {
  const [classId, setClassId] = useState(readSchoolFormsClassId);

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
  const active = selectedSchoolFormsClass?.id || scopedClasses[0]?.id || "";
  const klass = classes.find((item) => item.id === active);

  const { data: students = [] } = useQuery({
    enabled: !!active,
    queryKey: ["students", active],
    queryFn: async () =>
      (await supabase.from("students").select("*").eq("class_id", active).order("last_name"))
        .data as StudentRow[],
  });

  const { data: grades = [] } = useQuery({
    enabled: !!active,
    queryKey: ["grades-all", active],
    queryFn: async () =>
      (await supabase.from("grades").select("*").eq("class_id", active)).data as GradeRow[],
  });

  const learnerRows = useMemo<LearnerSummary[]>(
    () =>
      students.map((student) => {
        const subjectAverages = SUBJECTS.map((subject) => {
          const finalGrade = grades.find(
            (grade) =>
              grade.student_id === student.id &&
              grade.subject === subject &&
              grade.term === "final",
          )?.score;

          if (typeof finalGrade === "number") {
            return finalGrade;
          }

          return computeAverage(
            ["1", "2", "3"].map(
              (term) =>
                grades.find(
                  (grade) =>
                    grade.student_id === student.id &&
                    grade.subject === subject &&
                    grade.term === term,
                )?.score ?? null,
            ),
          );
        });

        return {
          st: student,
          avg: computeAverage(subjectAverages),
        };
      }),
    [students, grades],
  );

  const selectedGrade = getGradeNumber(klass?.grade_level);

  const report = useMemo(() => {
    const count = (
      list: LearnerSummary[],
      predicate: (row: LearnerSummary) => boolean,
    ): SexCounts => {
      const matching = list.filter(predicate);
      const male = matching.filter((row) => row.st.sex === "male").length;
      const female = matching.filter((row) => row.st.sex === "female").length;

      return { male, female, total: male + female };
    };

    return GRADE_LEVELS.map((grade) => {
      if (grade.value !== selectedGrade) {
        return {
          grade,
          promoted: { ...EMPTY_COUNTS },
          conditional: { ...EMPTY_COUNTS },
          retained: { ...EMPTY_COUNTS },
          enrolled: { ...EMPTY_COUNTS },
          bands: PROGRESS_BANDS.map(() => ({ ...EMPTY_COUNTS })),
        };
      }

      return {
        grade,
        promoted: count(learnerRows, (row) => row.avg != null && row.avg >= 75),
        conditional: { ...EMPTY_COUNTS },
        retained: count(learnerRows, (row) => row.avg != null && row.avg < 75),
        enrolled: count(learnerRows, () => true),
        bands: PROGRESS_BANDS.map((band) =>
          count(learnerRows, (row) => row.avg != null && band.matches(row.avg)),
        ),
      };
    });
  }, [learnerRows, selectedGrade]);

  const promotedByGrade = report.map((item) => item.promoted);
  const conditionalByGrade = report.map((item) => item.conditional);
  const retainedByGrade = report.map((item) => item.retained);
  const enrolledByGrade = report.map((item) => item.enrolled);
  const bandValues = PROGRESS_BANDS.map((_, bandIndex) =>
    report.map((item) => item.bands[bandIndex]),
  );

  const schoolName =
    profile?.school_name || klass?.school_name || "Agusan del Sur National Science High School";
  const schoolId = profile?.school_id || klass?.school_id || "";
  const region = klass?.region || profile?.region || "";
  const division = klass?.division || profile?.division || "";
  const district = klass?.district || profile?.district || "";
  const schoolYear = klass?.school_year || profile?.school_year || "";
  const schoolHead = profile?.principal || profile?.full_name || "";

  const statusRows = [
    {
      label: "PROMOTED",
      values: promotedByGrade,
      total: addCounts(promotedByGrade),
    },
    {
      label: "CONDITIONAL",
      values: conditionalByGrade,
      total: addCounts(conditionalByGrade),
    },
    {
      label: "RETAINED",
      values: retainedByGrade,
      total: addCounts(retainedByGrade),
    },
  ];

  const exportToExcel = async () => {
    if (!klass) {
      toast.error("Please select a class first");
      return;
    }

    try {
      const baseName = `SF6_Promotion_${klass.section || "school"}`.replace(/[<>:"/\\|?*]+/g, "_");
      await downloadExcelWorkbook(baseName, {
        schoolName,
        schoolId,
        region,
        division,
        district,
        schoolYear,
        gradeSection: [klass.grade_level, klass.section].filter(Boolean).join(" - "),
        statusRows,
        bandValues,
        enrolledByGrade,
        schoolHead,
      });

      toast.success("SF6 exported as a valid .xlsx workbook");
    } catch (error) {
      console.error("SF6 Excel export failed", error);
      toast.error("Unable to export SF6 to Excel");
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
          <ArrowLeft className="size-4" />
          Back to Forms
        </Link>

        <div className="flex items-center gap-2 text-lg font-semibold">
          <BadgeCheck className="size-5" style={{ color: DEPED_BLUE }} />
          School Form 6 — Summarized Report on Promotion
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
              onValueChange={setClassId}
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
        </div>
      </div>

      <PdfPreviewShell
        fileName={`SF6_Promotion_${klass?.section || "school"}.pdf`}
        docBaseName={`SF6_Promotion_${klass?.section || "school"}`}
        printTargetId="sf6-doc"
        orientation="landscape"
        pageLabel="Long Bond • 8.5 × 13 in • Landscape"
        onExportToExcel={exportToExcel}
      >
        <div
          id="sf6-doc"
          className="h-full overflow-hidden bg-white p-[18px] text-black"
          style={{
            width: 1250,
            fontFamily: "Arial, Helvetica, sans-serif",
          }}
        >
          <div className="relative min-h-[112px]">
            <img
              src={depedLogo}
              alt="Department of Education"
              className="absolute left-3 top-0 h-[98px] w-[98px] object-contain"
            />

            <div className="px-[150px] pt-1 text-center">
              <div className="text-[22px] font-bold leading-tight">School Form 6 (SF6)</div>
              <div className="text-[20px] font-bold leading-tight">
                Summarized Report on Promotion and Learning Progress &amp; Achievement
              </div>
            </div>

            <div className="mx-auto mt-3 grid max-w-[820px] grid-cols-4 gap-x-5 gap-y-2">
              <MetadataField label="School ID" value={schoolId} />
              <MetadataField label="Region" value={region} />
              <MetadataField label="Division" value={division} />
              <MetadataField label="District" value={district} />
              <MetadataField label="School Name" value={schoolName} wide />
              <MetadataField label="School Year" value={schoolYear} />
              <div className="flex items-end gap-1">
                <span className="shrink-0 text-[10px]">Grade &amp; Section</span>
                <span className="min-h-6 flex-1 border border-black px-1.5 py-1 text-[10px]">
                  {[klass?.grade_level, klass?.section].filter(Boolean).join(" - ")}
                </span>
              </div>
            </div>
          </div>

          <table className="mt-2 w-full table-fixed border-2 border-black border-collapse text-[8.5px] leading-tight">
            <colgroup>
              <col style={{ width: 148 }} />
              {Array.from({ length: 21 }).map((_, index) => (
                <col key={index} />
              ))}
            </colgroup>

            <thead>
              <tr>
                <th
                  rowSpan={2}
                  className="border border-black px-1 py-2 text-left text-[11px] font-normal"
                >
                  SUMMARY TABLE
                </th>
                {GRADE_LEVELS.map((grade) => (
                  <th
                    key={grade.value}
                    colSpan={3}
                    className="border border-black px-1 py-2 text-center text-[11px] font-normal"
                    style={{ borderLeftWidth: 2 }}
                  >
                    {grade.label}
                  </th>
                ))}
                <th
                  colSpan={3}
                  className="border border-black px-1 py-2 text-center text-[12px] font-normal"
                  style={{ borderLeftWidth: 2 }}
                >
                  TOTAL
                </th>
              </tr>

              <tr>
                {[...GRADE_LEVELS, { value: "total", label: "TOTAL" }].map((grade) => (
                  <Fragment key={grade.value}>
                    <th
                      className="border border-black px-0.5 py-2 text-center font-normal"
                      style={{ borderLeftWidth: 2 }}
                    >
                      Male
                    </th>
                    <th className="border border-black px-0.5 py-2 text-center font-normal">
                      Female
                    </th>
                    <th className="border border-black px-0.5 py-2 text-center font-normal">
                      Total
                    </th>
                  </Fragment>
                ))}
              </tr>
            </thead>

            <tbody>
              {statusRows.map((row) => (
                <tr key={row.label}>
                  <td className="border border-black px-1 py-2.5 text-[10px]">{row.label}</td>
                  {row.values.map((value, index) => (
                    <CountCells key={GRADE_LEVELS[index].value} value={value} />
                  ))}
                  <CountCells value={row.total} strong />
                </tr>
              ))}

              <tr>
                <td className="border border-black px-1 py-2 text-[7px] uppercase">
                  Learning Progress and
                  <br />
                  Achievement
                </td>
                {[...GRADE_LEVELS, { value: "total", label: "TOTAL" }].map((grade) => (
                  <Fragment key={grade.value}>
                    <td
                      className="border border-black px-0.5 py-2 text-center"
                      style={{ borderLeftWidth: 2 }}
                    >
                      Male
                    </td>
                    <td className="border border-black px-0.5 py-2 text-center">Female</td>
                    <td className="border border-black px-0.5 py-2 text-center">Total</td>
                  </Fragment>
                ))}
              </tr>

              {PROGRESS_BANDS.map((band, bandIndex) => {
                const values = bandValues[bandIndex];

                return (
                  <tr key={band.key}>
                    <td className="border border-black px-1 py-2 text-[7px]">
                      <div>{band.label}</div>
                      <div>{band.detail}</div>
                    </td>
                    {values.map((value, index) => (
                      <CountCells key={GRADE_LEVELS[index].value} value={value} />
                    ))}
                    <CountCells value={addCounts(values)} strong />
                  </tr>
                );
              })}

              <tr className="border-t-2 border-black">
                <td className="border border-black px-1 py-2.5 text-[11px]">TOTAL</td>
                {enrolledByGrade.map((value, index) => (
                  <CountCells key={GRADE_LEVELS[index].value} value={value} strong />
                ))}
                <CountCells value={addCounts(enrolledByGrade)} strong />
              </tr>
            </tbody>
          </table>

          <div className="mt-1 text-[8px]">Prepared and Submitted by:</div>

          <div className="mt-7 grid grid-cols-4 gap-10 px-1 text-center text-[8px]">
            <div>
              <div className="min-h-[13px] text-[10px] uppercase">{schoolHead}</div>
              <div className="border-t-2 border-black pt-1">
                (Signature of School Head/SCC Chair)
              </div>
            </div>

            <div>
              <div className="min-h-[13px]" />
              <div className="border-t-2 border-black pt-1">SCC-Vice Chair (Curriculum)</div>
            </div>

            <div>
              <div className="min-h-[13px]" />
              <div className="border-t-2 border-black pt-1">SCC Member</div>
            </div>

            <div>
              <div className="min-h-[13px]" />
              <div className="border-t-2 border-black pt-1">
                SCC-Vice Chair (Generated thru LIS)
              </div>
            </div>
          </div>

          <div className="mt-5 text-[8px]">(Additional slots may be added for SCC members.)</div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
