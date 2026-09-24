import { createFileRoute, Link, useRouterState } from "@tanstack/react-router";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import {
  computeAverage,
  getUserId,
  type ActivityScore,
  type ClassRow,
  type GradeActivity,
  type GradeComponent,
  type GradeRow,
  type StudentRow,
} from "@/lib/data";
import { Fragment, useEffect, useMemo, useState } from "react";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import {
  ArrowLeft,
  BarChart3,
  CheckCircle2,
  FileSpreadsheet,
  GripVertical,
  Hourglass,
  Inbox,
  Lightbulb,
  LoaderCircle,
  Mail,
  Plus,
  RefreshCw,
  Send,
  Trophy,
  UsersRound,
  X,
} from "lucide-react";
import { toast } from "sonner";
import JSZip from "jszip";
import { AnalyticsInsightsPanel } from "@/components/AnalyticsInsightsPanel";

export const Route = createFileRoute("/_authenticated/summary-of-grades")({
  component: SOGPage,
});

const TERMS = [
  { value: "1", label: "Term 1" },
  { value: "2", label: "Term 2" },
  { value: "3", label: "Term 3" },
  { value: "final", label: "FINAL" },
] as const;

const GRADE11_ACADEMIC_ELECTIVE_SUBJECTS = [
  "Academic Elective 1",
  "Academic Elective 2",
  "Academic Elective 3",
] as const;

type Grade11AcademicElectiveSubject =
  (typeof GRADE11_ACADEMIC_ELECTIVE_SUBJECTS)[number];

const MAPEH_SUMMARY_SCOPES = [
  { key: "MA_T1", storageTerm: "1" },
  { key: "PEH_T1", storageTerm: "PEH_T1" },
  { key: "MA_T2", storageTerm: "MA_T2" },
  { key: "PEH_T2", storageTerm: "PEH_T2" },
  { key: "MA_T3", storageTerm: "MA_T3" },
  { key: "PEH_T3", storageTerm: "PEH_T3" },
] as const;

const COMMUNICATION_SUMMARY_SCOPES = [
  { key: "EC_T1", storageTerm: "EC_T1" },
  { key: "MK_T1", storageTerm: "MK_T1" },
  { key: "EC_T2", storageTerm: "EC_T2" },
  { key: "MK_T2", storageTerm: "MK_T2" },
  { key: "EC_T3", storageTerm: "EC_T3" },
  { key: "MK_T3", storageTerm: "MK_T3" },
] as const;

type SummaryGradeTerm = "1" | "2" | "3" | "final";

function roundMapehInitialGrade(
  initialGrade: number | null | undefined,
): number | null {
  if (typeof initialGrade !== "number" || Number.isNaN(initialGrade)) {
    return null;
  }

  const grade = Math.max(0, Math.min(100, initialGrade));
  return Math.floor(grade + 0.5);
}

function resolveMapehTermGradeBase(
  computedBase: number | null,
  storedBase: number | null | undefined,
): number | null {
  if (computedBase == null) return null;

  const numericBase = storedBase == null ? computedBase : Number(storedBase);
  const resolvedBase = Number.isFinite(numericBase) ? numericBase : computedBase;

  return Math.max(0, Math.min(100, Math.floor(resolvedBase + 0.5)));
}

function averageMapehPair(
  first: number | null,
  second: number | null,
): number | null {
  if (first == null || second == null) return null;
  return Math.round((first + second) / 2);
}

const REQUEST_SUBJECTS = [
  "English",
  "Filipino",
  "Mathematics",
  "Science",
  "Araling Panlipunan",
] as const;

const SEND_REQUEST_ROWS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] as const;

const SUMMARY_EXCEL_TEMPLATE_URL = "/templates/Summary-of-grades.xlsm";
const SUMMARY_PER_SUBJECT_TEMPLATE_URL =
  "/templates/Summary-of-Final-Grades-per-Subject.xlsm";
const FINAL_GRADES_AND_GENERAL_AVERAGE_TEMPLATE_URL =
  "/templates/FINAL-GRADES-AND-GENERAL-AVERAGE.xlsm";

type SummaryExportProfile = {
  full_name?: string | null;
  school_name?: string | null;
  school_id?: string | null;
  region?: string | null;
  division?: string | null;
  district?: string | null;
  principal?: string | null;
};


type TeacherType = "class_adviser" | "subject_teacher";

type ClassWithTeacherName = ClassRow & {
  teacher_name?: string | null;
  school_name?: string | null;
  school_id?: string | null;
  region?: string | null;
  division?: string | null;
  district?: string | null;
  principal?: string | null;
  track_shs?: string | null;
};

function formatGradeLevel(value: unknown) {
  const text = String(value ?? "").trim();

  if (!text) return "Grade —";

  return /^grade\s+/i.test(text) ? text : `Grade ${text}`;
}

type GradeRequest = {
  id: string;
  requester_id: string;
  subject_teacher_id: string | null;
  subject: string;
  term: string;
  message: string;
  priority: "low" | "medium" | "high";
  status: "pending" | "in_progress" | "completed" | "rejected";
  teacher_response: string | null;
  created_at: string;
};

type ParsedReceivedForm = GradeRequest & {
  source: "receive_form" | "finalized_request";
  subjectRequestId?: string;
  advisoryClassId?: string;
  gradeLevel: string;
  section: string;
  schoolYear: string;
  learners: Array<{
    studentId?: string;
    name: string;
    score: number | null;
  }>;
};

type AddSubjectGradeSource = "request_grades" | "send_grades";

type FinalizedWorkflowRow = {
  batch_id: string;
  advisory_class_id: string;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  grading_period: "1" | "2" | "3" | "final";
  overall_status: "pending" | "completed";
  is_finalized: boolean;
  finalized_at: string | null;
  subject_request_id: string;
  subject: string;
  subject_status: "pending" | "submitted";
  submitted_at: string | null;
  teacher_note: string | null;
};

type FinalizedWorkflowLearner = {
  id: string;
  first_name: string;
  middle_name: string | null;
  last_name: string;
  lrn: string | null;
  sex: string | null;
  score: number;
};

type RecipientProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type?: TeacherType | null;
};

function normalizeText(value: unknown) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function isCommunicationCompositeSubject(value: unknown) {
  const normalized = normalizeText(value);
  return (
    normalized === "mabisang komunikasyon" ||
    normalized === "mabisang komuniksyon"
  );
}

function isGrade11ElectiveSubject(
  gradeLevel: unknown,
  subject: unknown,
) {
  return (
    normalizeText(gradeLevel) === "grade 11" &&
    normalizeText(subject) === "elective subject"
  );
}

function academicElectiveTerm(
  subject: unknown,
): "1" | "2" | "3" | null {
  const normalized = normalizeText(subject);

  if (normalized === "academic elective 1") return "1";
  if (normalized === "academic elective 2") return "2";
  if (normalized === "academic elective 3") return "3";

  return null;
}

function normalizeLearnerName(value: unknown) {
  return normalizeText(value)
    .replace(/[.'’]/g, "")
    .replace(/\s*,\s*/g, ",")
    .replace(/\s+/g, " ");
}

function gradeLevelKey(value: unknown) {
  return normalizeText(value).replace(/^grade\s+/, "");
}

type Grade12SummaryAssignedTerm = "1" | "2" | "3";

function normalizeGrade12SummarySubject(value: unknown) {
  return normalizeText(value)
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Grade 12 SF9 subject placement.
// Only the subject assigned to the currently selected term appears in
// Summary of Grades. The Term 1 / Term 2 / Term 3 / FINAL tabs remain.
//
// PE and Health 3 / PE and Health 4 are intentionally not included here,
// matching the subject list requested for this Summary of Grades change.
const GRADE12_SUMMARY_SUBJECT_TERM: Record<
  string,
  Grade12SummaryAssignedTerm
> = {
  // 1st Term
  "media and information literacy": "1",
  "filipino sa piling larang": "1",
  "practical research 2": "1",
  "general biology 1": "1",
  "general physics 1": "1",

  // 2nd Term
  "introduction to human philosophy": "2",
  "disaster readiness and risk reduction": "2",
  "english for academic and professional purposes": "2",
  entrepreneurship: "2",
  "general physics 2": "2",

  // 3rd Term
  "contemporary philippine arts": "3",
  "general biology 2": "3",
  "inquiries investigation and immersion": "3",
  "capstone project": "3",
};

function grade12SummarySubjectTerm(
  subject: unknown,
): Grade12SummaryAssignedTerm | null {
  return (
    GRADE12_SUMMARY_SUBJECT_TERM[
      normalizeGrade12SummarySubject(subject)
    ] ?? null
  );
}

function parseReceivedForm(request: GradeRequest): ParsedReceivedForm {
  const lines = String(request.message ?? "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const classDetailsLine =
    lines.find(
      (line) => /grade\s+\S+/i.test(line) && /section\s*:/i.test(line),
    ) ?? "";
  const gradeLevelMatch = classDetailsLine.match(/\b(Grade\s+[^·|]+)/i);
  const sectionMatch = classDetailsLine.match(/Section:\s*([^·|]+)/i);
  const schoolYearMatch = classDetailsLine.match(/School Year:\s*([^·|]+)/i);

  const learners = lines.flatMap((line) => {
    const match = line.match(/^\d+\.\s+(.+?)\s+[—–-]\s+[^:]+:\s*(.+)$/);
    if (!match) return [];

    const scoreText = match[2].trim();
    const parsedScore = Number(scoreText);

    return [
      {
        name: match[1].trim(),
        score:
          Number.isFinite(parsedScore) && parsedScore >= 0 && parsedScore <= 100
            ? parsedScore
            : null,
      },
    ];
  });

  return {
    ...request,
    source: "receive_form",
    gradeLevel: formatGradeLevel(gradeLevelMatch?.[1] ?? ""),
    section: sectionMatch?.[1]?.trim() ?? "",
    schoolYear: schoolYearMatch?.[1]?.trim() ?? "",
    learners,
  };
}

function roundedGrade(value: number | null | undefined) {
  if (typeof value !== "number" || Number.isNaN(value)) return null;
  return Math.round(value);
}

function descriptor(avg: number | null) {
  if (avg == null) return "—";
  if (avg >= 90) return "ADVANCING";
  if (avg >= 80) return "BENCHMARKING";
  if (avg >= 75) return "CONNECTING";
  if (avg >= 65) return "DEVELOPING";
  return "EMERGING";
}

function award(avg: number | null) {
  if (avg == null) return null;
  if (avg >= 98) {
    return {
      label: "PROMOTED WITH HIGH HONORS",
      tone: "text-amber-600",
    };
  }
  if (avg >= 90) {
    return {
      label: "PROMOTED WITH HONORS",
      tone: "text-amber-600",
    };
  }
  if (avg >= 75) {
    return {
      label: "PROMOTED",
      tone: "text-emerald-600",
    };
  }
  return null;
}


function passFailRemark(avg: number | null) {
  if (avg == null) return "";
  return avg >= 75 ? "PASSED" : "FAILED";
}


function xmlCell(sheetDocument: Document, reference: string) {
  return sheetDocument.querySelector(`c[r="${reference}"]`);
}

function clearXmlCell(sheetDocument: Document, reference: string) {
  const cell = xmlCell(sheetDocument, reference);
  if (!cell) return;

  cell.removeAttribute("t");
  Array.from(cell.children).forEach((child) => child.remove());
}


function copyXmlCellStyle(
  sheetDocument: Document,
  sourceReference: string,
  targetReference: string,
) {
  const sourceCell = xmlCell(sheetDocument, sourceReference);
  const targetCell = xmlCell(sheetDocument, targetReference);
  if (!sourceCell || !targetCell) return;

  const style = sourceCell.getAttribute("s");
  if (style == null) {
    targetCell.removeAttribute("s");
  } else {
    targetCell.setAttribute("s", style);
  }
}

function clearXmlCellVisual(
  sheetDocument: Document,
  reference: string,
) {
  const cell = xmlCell(sheetDocument, reference);
  if (!cell) return;

  clearXmlCell(sheetDocument, reference);
  cell.removeAttribute("s");
}


function updateSummaryConditionalFormatting(
  sheetDocument: Document,
  visibleSubjectCount: number,
) {
  const firstSubjectColumn = 3; // C
  const maxSubjectColumns = 15; // C:Q
  const safeVisibleCount = Math.max(
    1,
    Math.min(visibleSubjectCount || 1, maxSubjectColumns),
  );
  const lastSubjectColumnName = excelColumnName(
    firstSubjectColumn + safeVisibleCount - 1,
  );

  const conditionalFormattingNodes = Array.from(
    sheetDocument.querySelectorAll('conditionalFormatting'),
  );

  for (const node of conditionalFormattingNodes) {
    const sqref = node.getAttribute('sqref') ?? '';
    if (sqref === 'C9:Q19') {
      node.setAttribute('sqref', `C9:${lastSubjectColumnName}19`);
    }
    if (sqref === 'C21:Q44') {
      node.setAttribute('sqref', `C21:${lastSubjectColumnName}44`);
    }
  }
}


function removeSummaryStaticYellowRow(sheetDocument: Document) {
  // The template's first learner row (row 9) also has a built-in yellow style.
  // Copy only the cell styles from the normal learner row below (row 10)
  // so row 9 keeps its values but loses the yellow fill.
  for (let column = 1; column <= 20; column += 1) {
    copyXmlCellStyle(
      sheetDocument,
      `${excelColumnName(column)}10`,
      `${excelColumnName(column)}9`,
    );
  }
}

function removeSummaryActiveRowHighlight(sheetDocument: Document) {
  // The original XLSM template contains a conditional-format rule:
  //   ROW()=CELL("ROW")
  // over A9:T44. That rule paints the currently selected learner row yellow.
  // Remove only that rule so exported Summary of Grades files no longer show
  // the long yellow highlight across the worksheet.
  const conditionalFormattingNodes = Array.from(
    sheetDocument.querySelectorAll("conditionalFormatting"),
  );

  conditionalFormattingNodes.forEach((node) => {
    const rules = Array.from(node.querySelectorAll("cfRule"));

    rules.forEach((rule) => {
      const formula = rule.querySelector("formula")?.textContent?.trim() ?? "";
      if (formula === 'ROW()=CELL("ROW")') {
        rule.remove();
      }
    });

    if (node.querySelectorAll("cfRule").length === 0) {
      node.remove();
    }
  });
}

function addSummaryMergeRange(
  sheetDocument: Document,
  startColumn: number,
  endColumn: number,
  row: number,
) {
  if (endColumn <= startColumn) return;

  const mainNs =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  let mergeCells = sheetDocument.querySelector("mergeCells");

  if (!mergeCells) {
    mergeCells = sheetDocument.createElementNS(mainNs, "mergeCells");
    const pageMargins = sheetDocument.querySelector("pageMargins");
    const sheetData = sheetDocument.querySelector("sheetData");

    if (pageMargins?.parentNode) {
      pageMargins.parentNode.insertBefore(mergeCells, pageMargins);
    } else if (sheetData?.parentNode) {
      sheetData.parentNode.insertBefore(mergeCells, sheetData.nextSibling);
    } else {
      sheetDocument.documentElement.appendChild(mergeCells);
    }
  }

  const ref = `${excelColumnName(startColumn)}${row}:${excelColumnName(
    endColumn,
  )}${row}`;

  const alreadyExists = Array.from(
    mergeCells.querySelectorAll("mergeCell"),
  ).some((mergeCell) => mergeCell.getAttribute("ref") === ref);

  if (!alreadyExists) {
    const mergeCell = sheetDocument.createElementNS(mainNs, "mergeCell");
    mergeCell.setAttribute("ref", ref);
    mergeCells.appendChild(mergeCell);
  }

  mergeCells.setAttribute(
    "count",
    String(mergeCells.querySelectorAll("mergeCell").length),
  );
}

function expandSummaryRemarksArea(
  sheetDocument: Document,
  remarksColumn: number,
) {
  // Keep Remarks compact: use only two worksheet columns.
  // Example with one subject: F:G is the Remarks area.
  const lastWorksheetTableColumn = 20; // T
  const remarksEndColumn = Math.min(
    lastWorksheetTableColumn,
    remarksColumn + 1,
  );

  if (remarksEndColumn <= remarksColumn) return remarksEndColumn;

  // Merge the Remarks header and learner rows only.
  // MALE/FEMALE separator rows are handled separately so their horizontal
  // lines stay continuous across the whole table.
  const remarksRows = [
    7,
    ...Array.from({ length: 11 }, (_, index) => 9 + index),
    ...Array.from({ length: 24 }, (_, index) => 21 + index),
  ];

  for (const row of remarksRows) {
    for (
      let column = remarksColumn + 1;
      column <= remarksEndColumn;
      column += 1
    ) {
      copyXmlCellStyle(
        sheetDocument,
        `${excelColumnName(remarksColumn)}${row}`,
        `${excelColumnName(column)}${row}`,
      );
      clearXmlCell(sheetDocument, `${excelColumnName(column)}${row}`);
    }

    addSummaryMergeRange(
      sheetDocument,
      remarksColumn,
      remarksEndColumn,
      row,
    );
  }

  // Make the MALE and FEMALE separator rows one continuous cell from
  // STUDENT NAME through the end of Remarks. This removes the broken/
  // disconnected Remarks line visible in Excel.
  for (const row of [8, 20]) {
    for (let column = 3; column <= remarksEndColumn; column += 1) {
      copyXmlCellStyle(
        sheetDocument,
        `B${row}`,
        `${excelColumnName(column)}${row}`,
      );
      clearXmlCell(sheetDocument, `${excelColumnName(column)}${row}`);
    }

    addSummaryMergeRange(sheetDocument, 2, remarksEndColumn, row);
  }

  return remarksEndColumn;
}

function compactSummaryGradeTable(
  sheetDocument: Document,
  visibleSubjectCount: number,
) {
  // Remove the template's built-in yellow style before copying table styles.
  removeSummaryStaticYellowRow(sheetDocument);

  const firstSubjectColumn = 3; // C
  const lastSubjectColumn = 17; // Q
  const safeVisibleCount = Math.max(
    0,
    Math.min(visibleSubjectCount, lastSubjectColumn - firstSubjectColumn + 1),
  );

  // Keep rows 1:6 untouched so the original workbook header never changes.
  // Only the grade table beginning on row 7 is compacted.
  const averageColumn = firstSubjectColumn + safeVisibleCount;
  const roundedAverageColumn = averageColumn + 1;
  const remarksColumn = averageColumn + 2;

  const sourceAverageColumn = 18; // R
  const sourceRoundedAverageColumn = 19; // S
  const sourceRemarksColumn = 20; // T

  // Copy the original table presentation from R:S:T to the columns directly
  // after the last real subject. This changes only the table area, not the
  // REGION / DIVISION / SCHOOL YEAR / SCHOOL NAME / SCHOOL ID header above.
  for (let row = 7; row <= 44; row += 1) {
    copyXmlCellStyle(
      sheetDocument,
      `${excelColumnName(sourceAverageColumn)}${row}`,
      `${excelColumnName(averageColumn)}${row}`,
    );
    copyXmlCellStyle(
      sheetDocument,
      `${excelColumnName(sourceRoundedAverageColumn)}${row}`,
      `${excelColumnName(roundedAverageColumn)}${row}`,
    );
    copyXmlCellStyle(
      sheetDocument,
      `${excelColumnName(sourceRemarksColumn)}${row}`,
      `${excelColumnName(remarksColumn)}${row}`,
    );
  }

  // Remove the visual table cells for unused subject slots only.
  // Their worksheet columns remain present, so the top header layout does not
  // stretch or move.
  const firstUnusedSubjectColumn = firstSubjectColumn + safeVisibleCount;
  for (
    let column = firstUnusedSubjectColumn;
    column <= lastSubjectColumn;
    column += 1
  ) {
    // Do not clear the compact Average / helper / Remarks destination cells.
    if (
      column === averageColumn ||
      column === roundedAverageColumn ||
      column === remarksColumn
    ) {
      continue;
    }

    const columnName = excelColumnName(column);
    for (let row = 7; row <= 44; row += 1) {
      clearXmlCellVisual(sheetDocument, `${columnName}${row}`);
    }
  }

  // Remove the old Average / rounded / Remarks area at R:S:T when those
  // columns are no longer used by the compact table. This removes the extra
  // empty boxes shown on the far right side of the worksheet.
  for (let row = 7; row <= 44; row += 1) {
    for (let column = sourceAverageColumn; column <= sourceRemarksColumn; column += 1) {
      if (
        column === averageColumn ||
        column === roundedAverageColumn ||
        column === remarksColumn
      ) {
        continue;
      }
      clearXmlCellVisual(sheetDocument, `${excelColumnName(column)}${row}`);
    }
  }

  // Limit the pink conditional-format area to only the real subject columns.
  // This keeps the unused empty subject table area white.
  updateSummaryConditionalFormatting(sheetDocument, safeVisibleCount);

  // Remove the template's selected-row yellow highlight.
  removeSummaryActiveRowHighlight(sheetDocument);

  // Keep Remarks compact but wide enough for normal/long text,
  // without changing the worksheet header.
  const remarksEndColumn = expandSummaryRemarksArea(
    sheetDocument,
    remarksColumn,
  );

  // Remove any remaining visual cells to the right of the expanded Remarks
  // area so no empty boxes or leftover colors remain.
  for (let column = remarksEndColumn + 1; column <= sourceRemarksColumn; column += 1) {
    const columnName = excelColumnName(column);
    for (let row = 7; row <= 44; row += 1) {
      clearXmlCellVisual(sheetDocument, `${columnName}${row}`);
    }
  }

  return {
    averageColumn,
    roundedAverageColumn,
    remarksColumn,
    remarksEndColumn,
  };
}


function setXmlRowHidden(
  sheetDocument: Document,
  rowNumber: number,
  hidden: boolean,
) {
  const row = sheetDocument.querySelector(`row[r="${rowNumber}"]`);
  if (!row) return;

  if (hidden) {
    row.setAttribute("hidden", "1");
    row.setAttribute("ht", "0");
    row.setAttribute("customHeight", "1");
  } else {
    row.removeAttribute("hidden");

    // Restore normal visibility. We do not force a new height so the
    // original XLSM template formatting remains in control.
    if (row.getAttribute("ht") === "0") {
      row.removeAttribute("ht");
      row.removeAttribute("customHeight");
    }
  }
}

function setXmlText(
  sheetDocument: Document,
  reference: string,
  value: string | null | undefined,
) {
  const cell = xmlCell(sheetDocument, reference);
  if (!cell) return;

  Array.from(cell.children).forEach((child) => child.remove());

  const mainNs =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  cell.setAttribute("t", "inlineStr");

  const inlineString = sheetDocument.createElementNS(mainNs, "is");
  const textNode = sheetDocument.createElementNS(mainNs, "t");
  textNode.setAttribute(
    "xml:space",
    "preserve",
  );
  textNode.textContent = String(value ?? "");

  inlineString.appendChild(textNode);
  cell.appendChild(inlineString);
}

function setXmlNumber(
  sheetDocument: Document,
  reference: string,
  value: number | null | undefined,
) {
  const cell = xmlCell(sheetDocument, reference);
  if (!cell) return;

  Array.from(cell.children).forEach((child) => child.remove());
  cell.removeAttribute("t");

  if (value == null || Number.isNaN(value)) return;

  const mainNs =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const valueNode = sheetDocument.createElementNS(mainNs, "v");
  valueNode.textContent = String(value);
  cell.appendChild(valueNode);
}

function configureSummarySubjectColumns(
  sheetDocument: Document,
  visibleSubjectCount: number,
) {
  // Subject columns in the original template are C:Q = columns 3..17.
  // We keep only the subject columns currently shown in Summary of Grades
  // visible. The remaining subject columns are hidden, so Average/Remarks
  // appear immediately after the last real subject without changing the
  // original XLSM structure or VBA project.
  const firstSubjectColumn = 3;
  const lastSubjectColumn = 17;
  const safeVisibleCount = Math.max(
    0,
    Math.min(visibleSubjectCount, lastSubjectColumn - firstSubjectColumn + 1),
  );
  const lastVisibleSubjectColumn =
    firstSubjectColumn + safeVisibleCount - 1;

  const mainNs =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  let cols = sheetDocument.getElementsByTagName("cols")[0];

  if (!cols) {
    cols = sheetDocument.createElementNS(mainNs, "cols");
    const sheetData = sheetDocument.getElementsByTagName("sheetData")[0];

    if (sheetData?.parentNode) {
      sheetData.parentNode.insertBefore(cols, sheetData);
    } else {
      sheetDocument.documentElement.appendChild(cols);
    }
  }

  // Remove only the original C:Q column definition. The provided template
  // defines C:Q as one range, so we replace it with visible/hidden ranges.
  Array.from(cols.getElementsByTagName("col")).forEach((columnNode) => {
    const min = Number(columnNode.getAttribute("min") ?? "0");
    const max = Number(columnNode.getAttribute("max") ?? "0");

    if (min === firstSubjectColumn && max === lastSubjectColumn) {
      columnNode.remove();
    }
  });

  const addColumnRange = (
    min: number,
    max: number,
    hidden: boolean,
  ) => {
    if (min > max) return;

    const column = sheetDocument.createElementNS(mainNs, "col");
    column.setAttribute("min", String(min));
    column.setAttribute("max", String(max));
    column.setAttribute("width", "8.7109375");
    column.setAttribute("customWidth", "1");

    if (hidden) {
      column.setAttribute("hidden", "1");
    }

    cols!.appendChild(column);
  };

  if (safeVisibleCount > 0) {
    addColumnRange(
      firstSubjectColumn,
      lastVisibleSubjectColumn,
      false,
    );
  }

  if (lastVisibleSubjectColumn < lastSubjectColumn) {
    addColumnRange(
      Math.max(firstSubjectColumn, lastVisibleSubjectColumn + 1),
      lastSubjectColumn,
      true,
    );
  }
}

function excelColumnName(columnNumber: number) {
  let current = columnNumber;
  let result = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }

  return result;
}

function widenFinalGradesTermColumns(sheetDocument: Document) {
  // The provided template uses 14-point grade text inside very narrow
  // Term 1/2/3 columns. Excel therefore shows valid two-digit grades as #.
  // Widen only the three term columns in every subject block. Final-grade,
  // learner-name, average and remarks columns keep their original widths.
  const termColumns = new Set<number>();

  for (let subjectIndex = 0; subjectIndex < 15; subjectIndex += 1) {
    const firstColumn = 3 + subjectIndex * 4;
    termColumns.add(firstColumn);
    termColumns.add(firstColumn + 1);
    termColumns.add(firstColumn + 2);
  }

  Array.from(sheetDocument.getElementsByTagName("col")).forEach(
    (columnNode) => {
      const min = Number(columnNode.getAttribute("min") ?? "0");
      const max = Number(columnNode.getAttribute("max") ?? "0");

      if (!min || !max || min > max) return;

      let containsTermColumn = false;
      let containsOnlyTermColumns = true;

      for (let column = min; column <= max; column += 1) {
        if (termColumns.has(column)) {
          containsTermColumn = true;
        } else {
          containsOnlyTermColumns = false;
        }
      }

      if (containsTermColumn && containsOnlyTermColumns) {
        columnNode.setAttribute("width", "5.7109375");
        columnNode.setAttribute("customWidth", "1");
      }
    },
  );
}

function safeExcelFilePart(value: string) {
  return value
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 70);
}

async function removeCalcChainFromXlsm(zip: JSZip) {
  zip.remove("xl/calcChain.xml");

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (relsFile) {
    const relsXml = await relsFile.async("text");
    const relsDocument = parser.parseFromString(
      relsXml,
      "application/xml",
    );

    Array.from(relsDocument.getElementsByTagName("Relationship")).forEach(
      (relationship) => {
        const type = relationship.getAttribute("Type") ?? "";
        if (type.endsWith("/calcChain")) {
          relationship.remove();
        }
      },
    );

    zip.file(
      "xl/_rels/workbook.xml.rels",
      serializer.serializeToString(relsDocument),
    );
  }

  const contentTypesFile = zip.file("[Content_Types].xml");
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async("text");
    const contentTypesDocument = parser.parseFromString(
      contentTypesXml,
      "application/xml",
    );

    Array.from(contentTypesDocument.getElementsByTagName("Override")).forEach(
      (override) => {
        if (override.getAttribute("PartName") === "/xl/calcChain.xml") {
          override.remove();
        }
      },
    );

    zip.file(
      "[Content_Types].xml",
      serializer.serializeToString(contentTypesDocument),
    );
  }

  const workbookFile = zip.file("xl/workbook.xml");
  if (workbookFile) {
    const workbookXml = await workbookFile.async("text");
    const workbookDocument = parser.parseFromString(
      workbookXml,
      "application/xml",
    );

    const mainNs =
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let calcPr = workbookDocument.getElementsByTagName("calcPr")[0];

    if (!calcPr) {
      calcPr = workbookDocument.createElementNS(mainNs, "calcPr");
      workbookDocument.documentElement.appendChild(calcPr);
    }

    calcPr.setAttribute("calcMode", "auto");
    calcPr.setAttribute("fullCalcOnLoad", "1");
    calcPr.setAttribute("forceFullCalc", "1");

    zip.file(
      "xl/workbook.xml",
      serializer.serializeToString(workbookDocument),
    );
  }
}

function downloadXlsmBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function SummaryOfGradesSkeleton() {
  const learnerRows = Array.from({ length: 8 }, (_, index) => index);
  const requestRows = Array.from({ length: 5 }, (_, index) => index);

  return (
    <div
      className="space-y-4 pb-24 md:pb-6"
      role="status"
      aria-busy="true"
      aria-label="Loading Summary of Grades"
    >
      <span className="sr-only">Loading Summary of Grades...</span>

      <div className="relative flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card py-3 pl-4 pr-4 shadow-sm sm:pr-28">
        <div className="flex min-w-0 flex-1 items-center gap-3">
          <Skeleton className="h-9 w-20 shrink-0 rounded-lg" />
          <Skeleton className="size-6 shrink-0 rounded-full" />
          <div className="min-w-0 space-y-2">
            <div className="flex flex-wrap items-center gap-3">
              <Skeleton className="h-6 w-44 rounded-md" />
              <Skeleton className="h-8 w-40 rounded-lg" />
            </div>
            <Skeleton className="h-3 w-60 max-w-full rounded" />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <Skeleton className="h-8 w-56 rounded-md" />
          <Skeleton className="hidden h-8 w-52 rounded-md md:block" />
          <Skeleton className="hidden h-8 w-40 rounded-md lg:block" />
        </div>

        <Skeleton className="absolute right-4 top-3 hidden h-7 w-20 rounded-full sm:block" />
      </div>

      <div className="flex min-w-0 flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex max-w-full flex-wrap gap-1 rounded-xl border bg-card p-1">
          {TERMS.map((termItem) => (
            <Skeleton
              key={termItem.value}
              className="h-8 w-[4.5rem] rounded-lg"
            />
          ))}
        </div>
        <div className="flex w-full gap-2 sm:w-auto">
          <Skeleton className="h-9 flex-1 rounded-md sm:w-40 sm:flex-none" />
          <Skeleton className="h-9 flex-1 rounded-md sm:w-28 sm:flex-none" />
        </div>
      </div>

      <div className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="grid min-w-[880px] grid-cols-[2.5rem_minmax(14rem,2fr)_repeat(3,minmax(7rem,1fr))_5rem_5rem_8rem_10rem] border-b bg-muted/50 px-2 py-2">
          {Array.from({ length: 9 }, (_, index) => (
            <div key={index} className="flex items-center px-2">
              <Skeleton
                className={`h-3 rounded ${
                  index === 1 ? "w-32" : index >= 7 ? "w-20" : "w-12"
                }`}
              />
            </div>
          ))}
        </div>

        <div className="min-w-[880px]">
          <div className="border-b bg-muted/30 px-3 py-2">
            <Skeleton className="h-3 w-12 rounded" />
          </div>

          {learnerRows.map((row) => (
            <div
              key={row}
              className="grid grid-cols-[2.5rem_minmax(14rem,2fr)_repeat(3,minmax(7rem,1fr))_5rem_5rem_8rem_10rem] items-center border-b px-2 py-2"
            >
              <Skeleton className="mx-2 h-3 w-3 rounded" />
              <Skeleton
                className={`mx-2 h-3 rounded ${
                  row % 3 === 0 ? "w-40" : row % 3 === 1 ? "w-32" : "w-36"
                }`}
              />
              {Array.from({ length: 7 }, (_, column) => (
                <Skeleton
                  key={column}
                  className={`mx-auto h-3 rounded ${
                    column >= 5 ? "w-20" : "w-8"
                  }`}
                />
              ))}
            </div>
          ))}
        </div>
      </div>

      <div className="overflow-hidden rounded-2xl border-2 border-primary/20 bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b bg-primary/5 px-4 py-3">
          <div className="flex items-center gap-2">
            <Skeleton className="size-4 rounded" />
            <Skeleton className="h-4 w-52 rounded" />
            <Skeleton className="h-6 w-12 rounded-full" />
          </div>
          <Skeleton className="h-8 w-32 rounded-md" />
        </div>

        <div className="divide-y">
          {requestRows.map((row) => (
            <div key={row} className="flex items-center gap-2 p-3">
              <Skeleton className="h-9 w-24 shrink-0 rounded-md" />
              <Skeleton className="h-9 flex-1 rounded-md" />
              <Skeleton className="h-9 w-32 shrink-0 rounded-md" />
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function SOGPage() {
  const queryClient = useQueryClient();
  const locationHash = useRouterState({
    select: (state) => state.location.hash,
  });
  const routeSearch = useRouterState({
    select: (state) =>
      state.location.search as {
        addSubject?: boolean | string;
        classId?: string;
        subjectRequestId?: string;
      },
  });

  const [term, setTerm] = useState<string>("1");
  const [classId, setClassId] = useState<string>("");
  const [subjects, setSubjects] = useState<string[]>([]);
  const [draggedSubject, setDraggedSubject] = useState<string | null>(null);
  const [dragOverSubject, setDragOverSubject] = useState<string | null>(null);
  const [hydrated, setHydrated] = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [addGradeSource, setAddGradeSource] =
    useState<AddSubjectGradeSource | "">("");
  const [addPick, setAddPick] = useState<string>("");
  const [addGradeLevel, setAddGradeLevel] = useState<string>("");
  const [addRequestId, setAddRequestId] = useState<string>("");
  const [isImportingSubject, setIsImportingSubject] = useState(false);
  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [isExportingFinalGrades, setIsExportingFinalGrades] = useState(false);
  const [isExportingPerSubject, setIsExportingPerSubject] = useState(false);
  const [perSubjectOpen, setPerSubjectOpen] = useState(false);
  const [perSubjectPick, setPerSubjectPick] = useState<string>("");
  const [analyticsOpen, setAnalyticsOpen] = useState(false);
  const [selectedAnalyticsSubject, setSelectedAnalyticsSubject] = useState("");

  useEffect(() => {
    if (!analyticsOpen) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setAnalyticsOpen(false);
      }
    };

    window.addEventListener("keydown", closeOnEscape);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener("keydown", closeOnEscape);
    };
  }, [analyticsOpen]);

  const { data: teacherType = "class_adviser", isLoading: teacherTypeLoading } =
    useQuery<TeacherType>({
    queryKey: ["current-user-teacher-type"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return "class_adviser";

      const metadataTeacherType = user.user_metadata?.teacher_type;

      const { data: profile } = await (supabase as any)
        .from("profiles")
        .select("teacher_type")
        .eq("id", user.id)
        .maybeSingle();

      const profileTeacherType = profile?.teacher_type;

      if (
        profileTeacherType === "class_adviser" ||
        profileTeacherType === "subject_teacher"
      ) {
        return profileTeacherType;
      }

      if (
        metadataTeacherType === "class_adviser" ||
        metadataTeacherType === "subject_teacher"
      ) {
        return metadataTeacherType;
      }

      return "class_adviser";
    },
    staleTime: 5 * 60 * 1000,
  });

  const dashboardPath =
    teacherType === "subject_teacher"
      ? "/subject-teacher-dashboard"
      : "/dashboard";

  // Both Subject Teachers and Class Advisers can submit completed grades
  // to another Class Adviser.
  const canSendGradeRequests =
    teacherType === "subject_teacher" || teacherType === "class_adviser";

  // Analytics & Insights is available to both supported teacher roles.
  const canOpenAnalytics =
    teacherType === "subject_teacher" || teacherType === "class_adviser";

  const {
    data: classesData,
    isLoading: classesLoading,
    isPlaceholderData: classesPlaceholder,
  } = useQuery<ClassRow[]>({
    queryKey: ["classes"],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("classes")
        .select("*")
        .order("created_at", { ascending: false });

      if (error) throw error;
      return Array.isArray(data) ? (data as ClassRow[]) : [];
    },
    placeholderData: [],
  });

  const classes = Array.isArray(classesData) ? classesData : [];

  useEffect(() => {
    const requestedClassId = String(routeSearch?.classId ?? "").trim();

    if (
      requestedClassId &&
      requestedClassId !== classId &&
      classes.some((item) => item.id === requestedClassId)
    ) {
      setClassId(requestedClassId);
    }
  }, [classId, classes, routeSearch?.classId]);

  const active =
    classId || String(routeSearch?.classId ?? "") || classes[0]?.id || "";
  const klass = classes.find((item) => item.id === active) as
    ClassWithTeacherName | undefined;

  // If the E-Class Record changes this class's subject, refresh the class
  // immediately so Summary of Grades follows the new subject name.
  useEffect(() => {
    if (!active) return;

    const channel = supabase
      .channel(`summary-class-subject:${active}`)
      .on(
        "postgres_changes",
        {
          event: "UPDATE",
          schema: "public",
          table: "classes",
          filter: `id=eq.${active}`,
        },
        () => {
          void queryClient.invalidateQueries({
            queryKey: ["classes"],
          });
        },
      )
      .subscribe();

    return () => {
      void supabase.removeChannel(channel);
    };
  }, [active, queryClient]);


  const { data: exportProfile } = useQuery<SummaryExportProfile>({
    queryKey: ["summary-export-profile"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return {};

      const { data, error } = await (supabase as any)
        .from("profiles")
        .select(
          "full_name, school_name, school_id, region, division, district, principal",
        )
        .eq("id", user.id)
        .maybeSingle();

      if (error) throw error;
      return (data ?? {}) as SummaryExportProfile;
    },
    staleTime: 5 * 60 * 1000,
  });

  const {
    data: studentsData,
    isLoading: studentsLoading,
    isPlaceholderData: studentsPlaceholder,
  } = useQuery<StudentRow[]>({
    enabled: Boolean(active),
    queryKey: ["students", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("students")
        .select("*")
        .eq("class_id", active)
        .order("last_name");

      if (error) throw error;
      return Array.isArray(data) ? (data as StudentRow[]) : [];
    },
    placeholderData: [],
  });

  const students = Array.isArray(studentsData) ? studentsData : [];

  const {
    data: gradesData,
    isLoading: gradesLoading,
    isPlaceholderData: gradesPlaceholder,
  } = useQuery<GradeRow[]>({
    enabled: Boolean(active),
    queryKey: ["grades-all", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grades")
        .select("*")
        .eq("class_id", active);

      if (error) throw error;
      return Array.isArray(data) ? (data as GradeRow[]) : [];
    },
    placeholderData: [],
  });

  const grades = Array.isArray(gradesData) ? gradesData : [];

  // MAPEH is different from ordinary subjects. Its displayed term grade is
  // the rounded average of the MA and PEH term grades, and each MA/PEH grade
  // is computed from the MAPEH E-Class Record component/activity data.
  // Do this only for a class whose own E-Class Record subject is MAPEH.
  const isMapehClass = normalizeText(klass?.subject) === "mapeh";
  const mapehSubjectName = String(klass?.subject || "MAPEH").trim() || "MAPEH";
  const mapehStorageTerms = MAPEH_SUMMARY_SCOPES.map(
    (scope) => scope.storageTerm,
  );
  const mapehSummarySubjects = MAPEH_SUMMARY_SCOPES.map(
    (scope) => `${mapehSubjectName}_${scope.key}_ALL`,
  );

  const {
    data: mapehComponentsData,
    isLoading: mapehComponentsLoading,
    isPlaceholderData: mapehComponentsPlaceholder,
  } = useQuery<GradeComponent[]>({
    enabled: Boolean(active && isMapehClass),
    queryKey: ["summary-mapeh-components", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_components")
        .select("*")
        .eq("class_id", active)
        .in("term", mapehStorageTerms);

      if (error) throw error;
      return Array.isArray(data) ? (data as GradeComponent[]) : [];
    },
    placeholderData: [],
  });

  const mapehComponents = Array.isArray(mapehComponentsData)
    ? mapehComponentsData
    : [];

  const {
    data: mapehActivitiesData,
    isLoading: mapehActivitiesLoading,
    isPlaceholderData: mapehActivitiesPlaceholder,
  } = useQuery<GradeActivity[]>({
    enabled: Boolean(active && isMapehClass),
    queryKey: ["summary-mapeh-activities", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_activities")
        .select("*")
        .eq("class_id", active)
        .in("term", mapehStorageTerms)
        .order("position");

      if (error) throw error;
      return Array.isArray(data) ? (data as GradeActivity[]) : [];
    },
    placeholderData: [],
  });

  const mapehActivities = Array.isArray(mapehActivitiesData)
    ? mapehActivitiesData
    : [];
  const mapehActivityIds = mapehActivities.map((activity) => activity.id);

  const {
    data: mapehScoresData,
    isLoading: mapehScoresLoading,
    isPlaceholderData: mapehScoresPlaceholder,
  } = useQuery<ActivityScore[]>({
    enabled: Boolean(
      active && isMapehClass && mapehActivityIds.length > 0,
    ),
    queryKey: [
      "summary-mapeh-activity-scores",
      active,
      mapehActivityIds.join(","),
    ],
    queryFn: async () => {
      if (mapehActivityIds.length === 0) return [];

      const { data, error } = await supabase
        .from("activity_scores")
        .select("*")
        .in("activity_id", mapehActivityIds);

      if (error) throw error;
      return Array.isArray(data) ? (data as ActivityScore[]) : [];
    },
    placeholderData: [],
  });

  const mapehScores = Array.isArray(mapehScoresData) ? mapehScoresData : [];

  const {
    data: mapehSavedBasesData,
    isLoading: mapehSavedBasesLoading,
    isPlaceholderData: mapehSavedBasesPlaceholder,
  } = useQuery<
    Array<{
      student_id: string;
      subject: string;
      term_grade_base: number | null;
    }>
  >({
    enabled: Boolean(active && isMapehClass),
    queryKey: ["summary-mapeh-term-grade-bases", active, mapehSubjectName],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, subject, term_grade_base")
        .eq("class_id", active)
        .eq("term", "1")
        .in("subject", mapehSummarySubjects);

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

  const mapehSavedBases = Array.isArray(mapehSavedBasesData)
    ? mapehSavedBasesData
    : [];

  const mapehScoreMap = useMemo(() => {
    const map = new Map<string, number | null>();
    mapehScores.forEach((score) => {
      map.set(`${score.activity_id}|${score.student_id}`, score.score);
    });
    return map;
  }, [mapehScores]);

  const mapehSavedBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    mapehSavedBases.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) {
        map.set(`${row.student_id}|${row.subject}`, value);
      }
    });
    return map;
  }, [mapehSavedBases]);

  // Grade 11 Mabisang Komunikasyon is also a composite subject. Its displayed
  // term grade is the rounded average of Effective Communication and
  // Mabisang Komunikasyon for the same term.
  const isCommunicationClass = isCommunicationCompositeSubject(klass?.subject);
  const communicationSubjectName =
    String(klass?.subject || "Mabisang Komunikasyon").trim() ||
    "Mabisang Komunikasyon";

  // Grade 11 "Elective Subject" is represented in Summary of Grades as
  // three separate academic elective subjects:
  //   Term 1 -> Academic Elective 1
  //   Term 2 -> Academic Elective 2
  //   Term 3 -> Academic Elective 3
  // On FINAL, all three are shown as separate subject columns and each
  // keeps/copies its own term grade. The AVG column then averages the
  // three separate final subject grades.
  const isGrade11ElectiveClass = isGrade11ElectiveSubject(
    klass?.grade_level,
    klass?.subject,
  );
  const communicationStorageTerms = COMMUNICATION_SUMMARY_SCOPES.map(
    (scope) => scope.storageTerm,
  );
  const communicationSummarySubjects = COMMUNICATION_SUMMARY_SCOPES.map(
    (scope) => `${communicationSubjectName}_${scope.key}_ALL`,
  );

  const {
    data: communicationComponentsData,
    isLoading: communicationComponentsLoading,
    isPlaceholderData: communicationComponentsPlaceholder,
  } = useQuery<GradeComponent[]>({
    enabled: Boolean(active && isCommunicationClass),
    queryKey: ["summary-communication-components", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_components")
        .select("*")
        .eq("class_id", active)
        .in("term", communicationStorageTerms);

      if (error) throw error;
      return Array.isArray(data) ? (data as GradeComponent[]) : [];
    },
    placeholderData: [],
  });

  const communicationComponents = Array.isArray(communicationComponentsData)
    ? communicationComponentsData
    : [];

  const {
    data: communicationActivitiesData,
    isLoading: communicationActivitiesLoading,
    isPlaceholderData: communicationActivitiesPlaceholder,
  } = useQuery<GradeActivity[]>({
    enabled: Boolean(active && isCommunicationClass),
    queryKey: ["summary-communication-activities", active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_activities")
        .select("*")
        .eq("class_id", active)
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

  const {
    data: communicationScoresData,
    isLoading: communicationScoresLoading,
    isPlaceholderData: communicationScoresPlaceholder,
  } = useQuery<ActivityScore[]>({
    enabled: Boolean(
      active && isCommunicationClass && communicationActivityIds.length > 0,
    ),
    queryKey: [
      "summary-communication-activity-scores",
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

  const {
    data: communicationSavedBasesData,
    isLoading: communicationSavedBasesLoading,
    isPlaceholderData: communicationSavedBasesPlaceholder,
  } = useQuery<
    Array<{
      student_id: string;
      subject: string;
      term_grade_base: number | null;
    }>
  >({
    enabled: Boolean(active && isCommunicationClass),
    queryKey: [
      "summary-communication-term-grade-bases",
      active,
      communicationSubjectName,
    ],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, subject, term_grade_base")
        .eq("class_id", active)
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

  const communicationSavedBases = Array.isArray(communicationSavedBasesData)
    ? communicationSavedBasesData
    : [];

  const communicationScoreMap = useMemo(() => {
    const map = new Map<string, number | null>();
    communicationScores.forEach((score) => {
      map.set(`${score.activity_id}|${score.student_id}`, score.score);
    });
    return map;
  }, [communicationScores]);

  const communicationSavedBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    communicationSavedBases.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) {
        map.set(`${row.student_id}|${row.subject}`, value);
      }
    });
    return map;
  }, [communicationSavedBases]);

  const {
    data: requestsData,
    isLoading: requestsLoading,
    isPlaceholderData: requestsPlaceholder,
  } = useQuery<GradeRequest[]>({
    enabled: Boolean(active),
    queryKey: ["grade-requests", active],
    queryFn: async () => {
      const requesterId = await getUserId();
      const { data, error } = await (supabase as any)
        .from("grade_requests")
        .select(
          "id, requester_id, subject_teacher_id, subject, term, message, priority, status, teacher_response, created_at",
        )
        .eq("requester_id", requesterId)
        .order("created_at", { ascending: false });

      if (error) throw error;
      return Array.isArray(data) ? (data as GradeRequest[]) : [];
    },
    placeholderData: [],
  });

  const requests = Array.isArray(requestsData) ? requestsData : [];

  const {
    data: receivedRequestsData,
    isLoading: receivedRequestsLoading,
    isPlaceholderData: receivedRequestsPlaceholder,
  } = useQuery<GradeRequest[]>({
      enabled: teacherType === "class_adviser",
      queryKey: ["completed-received-grade-requests"],
      queryFn: async () => {
        const adviserId = await getUserId();
        const { data, error } = await (supabase as any)
          .from("grade_requests")
          .select(
            "id, requester_id, subject_teacher_id, subject, term, message, priority, status, teacher_response, created_at",
          )
          .eq("subject_teacher_id", adviserId)
          .eq("status", "completed")
          .order("created_at", { ascending: false });

        if (error) throw error;
        return Array.isArray(data) ? (data as GradeRequest[]) : [];
      },
      placeholderData: [],
    });

  const receivedForms = useMemo(
    () =>
      (Array.isArray(receivedRequestsData) ? receivedRequestsData : [])
        .map(parseReceivedForm)
        .filter(
          (request) =>
            request.gradeLevel !== "Grade —" && request.learners.length > 0,
        ),
    [receivedRequestsData],
  );

  const {
    data: finalizedRequestFormsData,
    isLoading: finalizedRequestFormsLoading,
    isPlaceholderData: finalizedRequestFormsPlaceholder,
  } = useQuery<ParsedReceivedForm[]>({
    enabled: teacherType === "class_adviser",
    queryKey: ["summary-finalized-grade-request-forms"],
    queryFn: async () => {
      const { data: workflowData, error: workflowError } = await (
        supabase as any
      ).rpc("list_my_sent_grade_request_workflow");

      if (workflowError) throw workflowError;

      const finalizedRows = (Array.isArray(workflowData) ? workflowData : [])
        .map((row: any) => row as FinalizedWorkflowRow)
        .filter(
          (row) =>
            Boolean(row.is_finalized) &&
            row.overall_status === "completed" &&
            row.subject_status === "submitted" &&
            Boolean(row.subject_request_id),
        );

      const forms = await Promise.all(
        finalizedRows.map(async (row) => {
          const { data: learnerData, error: learnerError } = await (
            supabase as any
          ).rpc("list_my_sent_grade_request_review_by_sex", {
            p_subject_request_id: row.subject_request_id,
          });

          if (learnerError) throw learnerError;

          const learnerRows = (Array.isArray(learnerData) ? learnerData : [])
            .map((learner: any) => learner as FinalizedWorkflowLearner)
            .filter((learner) => Number.isFinite(Number(learner.score)));

          const createdAt =
            row.finalized_at || row.submitted_at || new Date().toISOString();

          return {
            id: `finalized:${row.subject_request_id}`,
            requester_id: "",
            subject_teacher_id: null,
            subject: row.subject,
            term: row.grading_period,
            message: "",
            priority: "low" as const,
            status: "completed" as const,
            teacher_response: row.teacher_note,
            created_at: createdAt,
            source: "finalized_request" as const,
            subjectRequestId: row.subject_request_id,
            advisoryClassId: row.advisory_class_id,
            gradeLevel: formatGradeLevel(row.grade_level ?? ""),
            section: row.section?.trim() ?? "",
            schoolYear: row.school_year?.trim() ?? "",
            learners: learnerRows.map((learner) => ({
              studentId: learner.id,
              name: `${learner.last_name}, ${learner.first_name}${
                learner.middle_name ? ` ${learner.middle_name}` : ""
              }`.trim(),
              score: Number(learner.score),
            })),
          } satisfies ParsedReceivedForm;
        }),
      );

      return forms.filter(
        (form) => form.gradeLevel !== "Grade —" && form.learners.length > 0,
      );
    },
    placeholderData: [],
  });

  const finalizedRequestForms = Array.isArray(finalizedRequestFormsData)
    ? finalizedRequestFormsData
    : [];

  const allAddSubjectForms = useMemo(() => {
    const combined = [...receivedForms, ...finalizedRequestForms];
    const unique = new Map<string, ParsedReceivedForm>();

    combined.forEach((form) => {
      const key =
        form.source === "finalized_request"
          ? `finalized:${form.subjectRequestId || form.id}`
          : `received:${form.id}`;

      if (!unique.has(key)) unique.set(key, form);
    });

    return Array.from(unique.values()).sort(
      (first, second) =>
        new Date(second.created_at).getTime() -
        new Date(first.created_at).getTime(),
    );
  }, [finalizedRequestForms, receivedForms]);

  // Both teacher roles can send completed grades to another Class Adviser.
  const recipientTeacherType: TeacherType = "class_adviser";

  const {
    data: recipientProfilesData,
    isLoading: recipientProfilesLoading,
    isPlaceholderData: recipientProfilesPlaceholder,
  } = useQuery<RecipientProfile[]>({
    enabled: canSendGradeRequests,
    queryKey: ["grade-request-recipients", recipientTeacherType],
    queryFn: async () => {
      const senderId = await getUserId();
      const { data, error } = await (supabase as any)
        .from("profiles")
        .select("id, full_name, email, teacher_type")
        .eq("teacher_type", recipientTeacherType)
        .neq("id", senderId)
        .order("full_name", { ascending: true });

      if (error) throw error;

      return Array.isArray(data) ? (data as unknown as RecipientProfile[]) : [];
    },
    placeholderData: [],
  });

  const recipientProfiles = Array.isArray(recipientProfilesData)
    ? recipientProfilesData
    : [];

  const storageKey = active ? `sog-subjects-${active}` : "";
  const ownSubjectStorageKey = active
    ? `sog-own-subject-${active}`
    : "";

  useEffect(() => {
    if (!active || !klass) {
      setSubjects([]);
      setHydrated(false);
      return;
    }

    const raw = localStorage.getItem(storageKey);
    let savedSubjects: string[] = [];

    if (raw) {
      try {
        const parsed: unknown = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          savedSubjects = parsed.filter(
            (value): value is string => typeof value === "string",
          );
        }
      } catch {
        localStorage.removeItem(storageKey);
      }
    }

    const currentOwnSubject = String(klass.subject ?? "").trim();
    const previousOwnSubject =
      localStorage.getItem(ownSubjectStorageKey)?.trim() ?? "";

    let nextSubjects = [...savedSubjects];

    if (currentOwnSubject) {
      if (previousOwnSubject) {
        // Normal case after this fix has been used once:
        // replace only the previous class-owned subject with the current one.
        let replacedOwnSubject = false;

        nextSubjects = nextSubjects.map((subjectName) => {
          if (
            !replacedOwnSubject &&
            normalizeText(subjectName) === normalizeText(previousOwnSubject)
          ) {
            replacedOwnSubject = true;
            return currentOwnSubject;
          }

          return subjectName;
        });

        if (
          !nextSubjects.some(
            (subjectName) =>
              normalizeText(subjectName) === normalizeText(currentOwnSubject),
          )
        ) {
          nextSubjects.unshift(currentOwnSubject);
        }
      } else if (nextSubjects.length > 0) {
        // Compatibility with your OLD saved Summary data.
        // The old code stored the class's own subject as the FIRST subject.
        // So if Filipino was saved before this fix and the E-Class Record is
        // now MAPEH, replace that first main subject instead of adding MAPEH.
        if (
          normalizeText(nextSubjects[0]) !== normalizeText(currentOwnSubject)
        ) {
          nextSubjects = [currentOwnSubject, ...nextSubjects.slice(1)];
        }
      } else {
        nextSubjects = [currentOwnSubject];
      }

      localStorage.setItem(ownSubjectStorageKey, currentOwnSubject);
    } else {
      // If the class no longer has a subject, remove only its old main subject.
      if (previousOwnSubject) {
        nextSubjects = nextSubjects.filter(
          (subjectName) =>
            normalizeText(subjectName) !== normalizeText(previousOwnSubject),
        );
      }
      localStorage.removeItem(ownSubjectStorageKey);
    }

    // Prevent duplicate subject columns.
    const seenSubjects = new Set<string>();
    nextSubjects = nextSubjects.filter((subjectName) => {
      const key = normalizeText(subjectName);

      if (!key || seenSubjects.has(key)) {
        return false;
      }

      seenSubjects.add(key);
      return true;
    });

    setSubjects(nextSubjects);
    localStorage.setItem(storageKey, JSON.stringify(nextSubjects));
    setHydrated(true);
  }, [active, klass, storageKey, ownSubjectStorageKey]);

  useEffect(() => {
    if (!active || !hydrated || !storageKey) return;
    localStorage.setItem(storageKey, JSON.stringify(subjects));
  }, [subjects, active, hydrated, storageKey]);

  const safeSubjects = Array.isArray(subjects) ? subjects : [];
  const isGrade12SummaryClass = gradeLevelKey(klass?.grade_level) === "12";

  const displayedSubjects = useMemo(() => {
    if (isGrade12SummaryClass) {
      const mappedGrade12Subjects = safeSubjects.filter(
        (subjectName) => grade12SummarySubjectTerm(subjectName) != null,
      );

      // FINAL keeps all Grade 12 SF9 subjects visible.
      if (term === "final") {
        return mappedGrade12Subjects;
      }

      if (term === "1" || term === "2" || term === "3") {
        return mappedGrade12Subjects.filter(
          (subjectName) =>
            grade12SummarySubjectTerm(subjectName) === term,
        );
      }

      return [];
    }

    if (!isGrade11ElectiveClass) {
      return safeSubjects;
    }

    const ownElectiveSubject = String(klass?.subject || "Elective Subject");
    const otherSubjects = safeSubjects.filter(
      (subjectName) =>
        normalizeText(subjectName) !== normalizeText(ownElectiveSubject),
    );

    const hasOwnElective =
      safeSubjects.some(
        (subjectName) =>
          normalizeText(subjectName) === normalizeText(ownElectiveSubject),
      ) || safeSubjects.length === 0;

    if (!hasOwnElective) {
      return safeSubjects;
    }

    if (term === "1") {
      return ["Academic Elective 1", ...otherSubjects];
    }

    if (term === "2") {
      return ["Academic Elective 2", ...otherSubjects];
    }

    if (term === "3") {
      return ["Academic Elective 3", ...otherSubjects];
    }

    // FINAL: show all three Academic Electives as three separate subjects.
    return [
      "Academic Elective 1",
      "Academic Elective 2",
      "Academic Elective 3",
      ...otherSubjects,
    ];
  }, [
    safeSubjects,
    isGrade12SummaryClass,
    isGrade11ElectiveClass,
    klass?.subject,
    term,
  ]);

  // The teacher's Analytics dropdown follows the subject columns currently
  // displayed in Summary of Grades, including imported subject grades and
  // Grade 11 virtual elective labels. Admin's class analytics is unchanged.
  const analyticsSubjectOptions = useMemo(() => {
    const seen = new Set<string>();
    return displayedSubjects.filter((name) => {
      const key = normalizeText(name);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }, [displayedSubjects]);

  const defaultAnalyticsSubject =
    analyticsSubjectOptions.find(
      (name) => normalizeText(name) === normalizeText(klass?.subject),
    ) ?? analyticsSubjectOptions[0] ?? "";

  useEffect(() => {
    setSelectedAnalyticsSubject((current) =>
      analyticsSubjectOptions.includes(current) ? current : defaultAnalyticsSubject,
    );
  }, [analyticsSubjectOptions, defaultAnalyticsSubject]);

  const activeAnalyticsSubject = analyticsSubjectOptions.includes(
    selectedAnalyticsSubject,
  )
    ? selectedAnalyticsSubject
    : defaultAnalyticsSubject;

  const analyticsIsOwnSubject =
    normalizeText(activeAnalyticsSubject) === normalizeText(klass?.subject);

  // Other subjects on this class's SOG may be imported final grades only.
  // Do not use the current class's component scores to forecast another subject.
  const analyticsCanUseOwnComponents =
    analyticsIsOwnSubject &&
    !isCommunicationClass &&
    !isGrade11ElectiveClass &&
    !isGrade12SummaryClass;
  const analyticsUsesSummaryGrades = !analyticsCanUseOwnComponents;

  const reorderSubjects = (source: string, target: string) => {
    if (!source || !target || source === target) return;

    setSubjects((previous) => {
      const safePrevious = Array.isArray(previous) ? [...previous] : [];
      const sourceIndex = safePrevious.indexOf(source);
      const targetIndex = safePrevious.indexOf(target);

      if (sourceIndex === -1 || targetIndex === -1) return safePrevious;

      const [movedSubject] = safePrevious.splice(sourceIndex, 1);
      safePrevious.splice(targetIndex, 0, movedSubject);

      return safePrevious;
    });
  };

  useEffect(() => {
    if (!perSubjectOpen) return;

    if (!perSubjectPick || !safeSubjects.includes(perSubjectPick)) {
      setPerSubjectPick(safeSubjects[0] ?? "");
    }
  }, [perSubjectOpen, perSubjectPick, safeSubjects]);

  useEffect(() => {
    if (!perSubjectOpen) return;

    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setPerSubjectOpen(false);
      }
    };

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [perSubjectOpen]);

  const selectedGradeLevel = klass
    ? formatGradeLevel(klass.grade_level)
    : "Grade —";
  const selectedSchoolYear = klass?.school_year || "—";
  const sourceFilteredAddSubjectForms = useMemo(() => {
    if (addGradeSource === "request_grades") {
      return allAddSubjectForms.filter(
        (form) => form.source === "finalized_request",
      );
    }

    if (addGradeSource === "send_grades") {
      return allAddSubjectForms.filter((form) => form.source === "receive_form");
    }

    return [];
  }, [addGradeSource, allAddSubjectForms]);
  const availableReceivedSubjects = useMemo(
    () =>
      [
        ...new Set(
          sourceFilteredAddSubjectForms.map((request) => request.subject),
        ),
      ].sort(),
    [sourceFilteredAddSubjectForms],
  );
  const availableGradeLevels = useMemo(
    () =>
      [
        ...new Set(
          sourceFilteredAddSubjectForms
            .filter((request) => request.subject === addPick)
            .map((request) => request.gradeLevel),
        ),
      ].sort(),
    [sourceFilteredAddSubjectForms, addPick],
  );
  const matchingReceivedForms = useMemo(
    () =>
      sourceFilteredAddSubjectForms.filter(
        (request) =>
          request.subject === addPick &&
          gradeLevelKey(request.gradeLevel) === gradeLevelKey(addGradeLevel),
      ),
    [sourceFilteredAddSubjectForms, addPick, addGradeLevel],
  );
  const selectedReceivedForm =
    matchingReceivedForms.find((request) => request.id === addRequestId) ??
    null;

  useEffect(() => {
    const shouldOpen =
      routeSearch?.addSubject === true ||
      String(routeSearch?.addSubject ?? "").toLowerCase() === "true";
    const subjectRequestId = String(routeSearch?.subjectRequestId ?? "").trim();

    if (!shouldOpen || !subjectRequestId || finalizedRequestFormsLoading) return;

    const target = allAddSubjectForms.find(
      (form) =>
        form.source === "finalized_request" &&
        form.subjectRequestId === subjectRequestId,
    );

    if (!target) return;

    setAddGradeSource("request_grades");
    setAddPick(target.subject);
    setAddGradeLevel(target.gradeLevel);
    setAddRequestId(target.id);
    setTerm(target.term);
    setAddOpen(true);
  }, [
    allAddSubjectForms,
    finalizedRequestFormsLoading,
    routeSearch?.addSubject,
    routeSearch?.subjectRequestId,
  ]);

  useEffect(() => {
    if (!addOpen) return;

    if (!addPick || !availableReceivedSubjects.includes(addPick)) {
      setAddPick(availableReceivedSubjects[0] ?? "");
    }
  }, [addOpen, addPick, availableReceivedSubjects]);

  useEffect(() => {
    if (!addOpen || !addPick) return;

    const currentGradeLevel = availableGradeLevels.find(
      (gradeLevel) =>
        gradeLevelKey(gradeLevel) === gradeLevelKey(selectedGradeLevel),
    );

    if (!addGradeLevel || !availableGradeLevels.includes(addGradeLevel)) {
      setAddGradeLevel(currentGradeLevel ?? availableGradeLevels[0] ?? "");
    }
  }, [
    addOpen,
    addPick,
    addGradeLevel,
    availableGradeLevels,
    selectedGradeLevel,
  ]);

  useEffect(() => {
    if (!addOpen || !addGradeLevel) return;

    if (!matchingReceivedForms.some((request) => request.id === addRequestId)) {
      setAddRequestId(matchingReceivedForms[0]?.id ?? "");
    }
  }, [addOpen, addGradeLevel, addRequestId, matchingReceivedForms]);

  useEffect(() => {
    if (!addOpen) return;

    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setAddOpen(false);
    };

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [addOpen]);

  const closeAddSubject = () => {
    setAddOpen(false);
    setAddGradeSource("");
    setAddPick("");
    setAddGradeLevel("");
    setAddRequestId("");
  };

  const addSubject = async () => {
    if (!active || !klass || !selectedReceivedForm) return;

    if (
      gradeLevelKey(selectedReceivedForm.gradeLevel) !==
      gradeLevelKey(klass.grade_level)
    ) {
      toast.error(
        `Select a ${selectedGradeLevel} received form for the currently selected class.`,
      );
      return;
    }

    if (
      selectedReceivedForm.section &&
      klass.section &&
      normalizeText(selectedReceivedForm.section) !==
        normalizeText(klass.section)
    ) {
      toast.error(
        `This form is for Section ${selectedReceivedForm.section}, not ${klass.section}.`,
      );
      return;
    }

    if (
      selectedReceivedForm.schoolYear &&
      klass.school_year &&
      normalizeText(selectedReceivedForm.schoolYear) !==
        normalizeText(klass.school_year)
    ) {
      toast.error(
        `This form is for School Year ${selectedReceivedForm.schoolYear}, not ${klass.school_year}.`,
      );
      return;
    }

    const studentById = new Map(
      students.map((student) => [student.id, student]),
    );
    const studentByName = new Map(
      students.map((student) => [
        normalizeLearnerName(`${student.last_name}, ${student.first_name}`),
        student,
      ]),
    );

    const matchedLearners = selectedReceivedForm.learners.flatMap((learner) => {
      const student =
        (learner.studentId ? studentById.get(learner.studentId) : undefined) ||
        studentByName.get(normalizeLearnerName(learner.name));

      return student ? [{ student, score: learner.score }] : [];
    });
    const learnersWithScores = matchedLearners.filter(
      (learner): learner is typeof learner & { score: number } =>
        typeof learner.score === "number",
    );

    if (matchedLearners.length === 0) {
      toast.error(
        "No learner names in this form match the selected class roster.",
      );
      return;
    }

    setIsImportingSubject(true);

    try {
      if (learnersWithScores.length > 0) {
        const teacherId = await getUserId();
        const rowsToImport = learnersWithScores.map(({ student, score }) => ({
          student_id: student.id,
          class_id: active,
          teacher_id: teacherId,
          subject: selectedReceivedForm.subject,
          term: selectedReceivedForm.term,
          score,
        }));

        const { error } = await (supabase as any)
          .from("grades")
          .upsert(rowsToImport, { onConflict: "student_id,subject,term" });

        if (error) throw error;
      }

      setSubjects((previous) => {
        const safePrevious = Array.isArray(previous) ? previous : [];
        return safePrevious.includes(selectedReceivedForm.subject)
          ? safePrevious
          : [...safePrevious, selectedReceivedForm.subject];
      });

      await queryClient.invalidateQueries({
        queryKey: ["grades-all", active],
      });

      const unmatchedCount =
        selectedReceivedForm.learners.length - matchedLearners.length;
      const missingGradeCount =
        matchedLearners.length - learnersWithScores.length;
      const noteParts = [
        unmatchedCount > 0 ? `${unmatchedCount} unmatched learner(s)` : "",
        missingGradeCount > 0
          ? `${missingGradeCount} without a submitted grade`
          : "",
      ].filter(Boolean);

      toast.success(
        `${selectedReceivedForm.subject} ${
          selectedReceivedForm.term === "final"
            ? "Final Grade"
            : `Term ${selectedReceivedForm.term}`
        } was added${noteParts.length > 0 ? ` (${noteParts.join(", ")})` : ""}.`,
      );
      setTerm(selectedReceivedForm.term);
      closeAddSubject();
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : "The received form could not be added.",
      );
    } finally {
      setIsImportingSubject(false);
    }
  };

  const isOwnMapehSubject = (subjectName: string) =>
    isMapehClass &&
    normalizeText(subjectName) === normalizeText(mapehSubjectName);

  const isOwnCommunicationSubject = (subjectName: string) =>
    isCommunicationClass &&
    normalizeText(subjectName) === normalizeText(communicationSubjectName);

  const mapehComponentGrade = (
    studentId: string,
    scopeKey: string,
    storageTerm: string,
  ): number | null => {
    const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
      const componentActivities = mapehActivities
        .filter(
          (activity) =>
            activity.term === storageTerm && activity.component === component,
        )
        .sort((first, second) => first.position - second.position)
        .slice(0, component === "QA" ? 3 : undefined);

      const componentWeight =
        mapehComponents.find(
          (item) =>
            item.term === storageTerm && item.component === component,
        )?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);

      let raw = 0;
      let hps = 0;
      let hasScore = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;

      componentActivities.forEach((activity, index) => {
        const value = mapehScoreMap.get(`${activity.id}|${studentId}`);
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

    const computedBase = roundMapehInitialGrade(initialGrade);
    if (computedBase == null) return null;

    return resolveMapehTermGradeBase(
      computedBase,
      mapehSavedBaseMap.get(
        `${studentId}|${mapehSubjectName}_${scopeKey}_ALL`,
      ),
    );
  };

  const mapehGradeForStudentTerm = (
    studentId: string,
    termValue: SummaryGradeTerm,
  ): number | null => {
    const gradeForPair = (
      firstScopeKey: string,
      firstStorageTerm: string,
      secondScopeKey: string,
      secondStorageTerm: string,
    ) =>
      averageMapehPair(
        mapehComponentGrade(studentId, firstScopeKey, firstStorageTerm),
        mapehComponentGrade(studentId, secondScopeKey, secondStorageTerm),
      );

    if (termValue === "1") {
      return gradeForPair("MA_T1", "1", "PEH_T1", "PEH_T1");
    }

    if (termValue === "2") {
      return gradeForPair("MA_T2", "MA_T2", "PEH_T2", "PEH_T2");
    }

    if (termValue === "3") {
      return gradeForPair("MA_T3", "MA_T3", "PEH_T3", "PEH_T3");
    }

    const term1 = gradeForPair("MA_T1", "1", "PEH_T1", "PEH_T1");
    const term2 = gradeForPair("MA_T2", "MA_T2", "PEH_T2", "PEH_T2");
    const term3 = gradeForPair("MA_T3", "MA_T3", "PEH_T3", "PEH_T3");
    const termValues = [term1, term2, term3];

    return termValues.every(
      (value): value is number => typeof value === "number",
    )
      ? Math.round(
          termValues.reduce((sum, value) => sum + value, 0) / 3,
        )
      : null;
  };

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
        const value = communicationScoreMap.get(
          `${activity.id}|${studentId}`,
        );
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

    const computedBase = roundMapehInitialGrade(initialGrade);
    if (computedBase == null) return null;

    return resolveMapehTermGradeBase(
      computedBase,
      communicationSavedBaseMap.get(
        `${studentId}|${communicationSubjectName}_${scopeKey}_ALL`,
      ),
    );
  };

  const communicationGradeForStudentTerm = (
    studentId: string,
    termValue: SummaryGradeTerm,
  ): number | null => {
    const gradeForPair = (
      firstScopeKey: string,
      secondScopeKey: string,
    ) =>
      averageMapehPair(
        communicationComponentGrade(studentId, firstScopeKey, firstScopeKey),
        communicationComponentGrade(studentId, secondScopeKey, secondScopeKey),
      );

    if (termValue === "1") {
      return gradeForPair("EC_T1", "MK_T1");
    }

    if (termValue === "2") {
      return gradeForPair("EC_T2", "MK_T2");
    }

    if (termValue === "3") {
      return gradeForPair("EC_T3", "MK_T3");
    }

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

  const scoreForTerm = (
    studentId: string,
    subjectName: string,
    termValue: SummaryGradeTerm,
  ): number | null => {
    if (isOwnMapehSubject(subjectName)) {
      return mapehGradeForStudentTerm(studentId, termValue);
    }

    if (isOwnCommunicationSubject(subjectName)) {
      return communicationGradeForStudentTerm(studentId, termValue);
    }

    if (isGrade12SummaryClass) {
      const assignedTerm = grade12SummarySubjectTerm(subjectName);

      if (assignedTerm) {
        if (termValue !== "final" && termValue !== assignedTerm) {
          return null;
        }

        return (
          grades.find(
            (grade) =>
              grade.student_id === studentId &&
              normalizeGrade12SummarySubject(grade.subject) ===
                normalizeGrade12SummarySubject(subjectName) &&
              grade.term === assignedTerm,
          )?.score ?? null
        );
      }
    }

    if (isGrade11ElectiveClass) {
      const mappedTerm = academicElectiveTerm(subjectName);

      if (mappedTerm) {
        // Each Academic Elective is a different one-term subject.
        // In FINAL, copy its own term grade instead of averaging the
        // three Academic Electives into one combined subject.
        return (
          grades.find(
            (grade) =>
              grade.student_id === studentId &&
              normalizeText(grade.subject) ===
                normalizeText(klass?.subject || "Elective Subject") &&
              grade.term === mappedTerm,
          )?.score ?? null
        );
      }
    }

    if (termValue === "final") {
      const values = ["1", "2", "3"].map(
        (value) =>
          grades.find(
            (grade) =>
              grade.student_id === studentId &&
              grade.subject === subjectName &&
              grade.term === value,
          )?.score ?? null,
      );
      return computeAverage(values);
    }

    return (
      grades.find(
        (grade) =>
          grade.student_id === studentId &&
          grade.subject === subjectName &&
          grade.term === termValue,
      )?.score ?? null
    );
  };

  // Use this exact subject's own saved Summary grades when its matching
  // E-Class Record components are not in the selected class.
  const analyticsSummarySubjectGrades = useMemo(() => {
    if (!analyticsUsesSummaryGrades || !activeAnalyticsSubject) {
      return undefined;
    }

    return Object.fromEntries(
      students.map((student) => {
        const computedFinal = scoreForTerm(
          student.id,
          activeAnalyticsSubject,
          "final",
        );
        // An imported subject can have only a submitted FINAL row rather
        // than separate T1/T2/T3 rows. Use its own saved final score as an
        // analytics-only fallback. Never borrow another subject's scores.
        const savedFinal = grades.find(
          (row) =>
            row.student_id === student.id &&
            normalizeText(row.subject) === normalizeText(activeAnalyticsSubject) &&
            row.term === "final",
        )?.score;
        const finalGrade =
          computedFinal != null
            ? computedFinal
            : savedFinal != null && Number.isFinite(Number(savedFinal))
              ? Number(savedFinal)
              : null;

        return [
          student.id,
          {
            "1": scoreForTerm(student.id, activeAnalyticsSubject, "1"),
            "2": scoreForTerm(student.id, activeAnalyticsSubject, "2"),
            "3": scoreForTerm(student.id, activeAnalyticsSubject, "3"),
            final: finalGrade,
          },
        ];
      }),
    ) as Record<
      string,
      Record<SummaryGradeTerm, number | null>
    >;
  }, [
    students,
    activeAnalyticsSubject,
    analyticsUsesSummaryGrades,
    grades,
    klass?.subject,
    isMapehClass,
    mapehSubjectName,
    mapehActivities,
    mapehComponents,
    mapehScoreMap,
    mapehSavedBaseMap,
    isCommunicationClass,
    communicationSubjectName,
    communicationActivities,
    communicationComponents,
    communicationScoreMap,
    communicationSavedBaseMap,
    isGrade11ElectiveClass,
    isGrade12SummaryClass,
  ]);

  const scoreFor = (studentId: string, subjectName: string) =>
    scoreForTerm(studentId, subjectName, term as SummaryGradeTerm);

  const rows = useMemo(() => {
    return students.map((student) => {
      const scores = displayedSubjects.map((subjectName) =>
        scoreFor(student.id, subjectName),
      );
      // The displayed subject grades are whole numbers. Compute AVG from those
      // same rounded grades so, for example, 75 and 88 produce 81.5 -> 82.
      const roundedScores = scores.map((score) => roundedGrade(score));
      return { st: student, scores, avg: computeAverage(roundedScores) };
    });
  }, [
    students,
    displayedSubjects,
    grades,
    term,
    isMapehClass,
    mapehSubjectName,
    mapehActivities,
    mapehComponents,
    mapehScoreMap,
    mapehSavedBaseMap,
    isCommunicationClass,
    communicationSubjectName,
    communicationActivities,
    communicationComponents,
    communicationScoreMap,
    communicationSavedBaseMap,
  ]);

  const ranked = useMemo(() => {
    const sorted = [...rows]
      .filter((row) => row.avg != null)
      .sort((first, second) => second.avg! - first.avg!);
    const rankMap = new Map<string, number>();
    sorted.forEach((row, index) => rankMap.set(row.st.id, index + 1));
    return rankMap;
  }, [rows]);

  const male = rows.filter((row) => row.st.sex === "male");
  const female = rows.filter((row) => row.st.sex === "female");
  const other = rows.filter(
    (row) => row.st.sex !== "male" && row.st.sex !== "female",
  );


  const exportSummaryPerSubject = async () => {
    if (!active || !klass) {
      toast.error("Select a class first.");
      return;
    }

    if (!perSubjectPick || !safeSubjects.includes(perSubjectPick)) {
      toast.error("Select a subject first.");
      return;
    }

    if (isExportingPerSubject) return;

    setIsExportingPerSubject(true);

    try {
      const { data: latestClassData, error: latestClassError } = await (
        supabase as any
      )
        .from("classes")
        .select("*")
        .eq("id", active)
        .single();

      if (latestClassError) throw latestClassError;

      const exportClass =
        (latestClassData as ClassWithTeacherName | null) ?? klass;
      const isGrade12PerSubjectExport =
        gradeLevelKey(exportClass.grade_level) === "12";

      const response = await fetch(SUMMARY_PER_SUBJECT_TEMPLATE_URL, {
        cache: "no-store",
      });

      if (!response.ok) {
        throw new Error(
          "Could not load /templates/Summary-of-Final-Grades-per-Subject.xlsm. Put the original .xlsm file inside public/templates.",
        );
      }

      const templateBytes = await response.arrayBuffer();

      const scoreForSubjectTerm = (
        studentId: string,
        subjectName: string,
        termValue: "1" | "2" | "3" | "final",
      ) => {
        if (isOwnMapehSubject(subjectName)) {
          return mapehGradeForStudentTerm(studentId, termValue);
        }

        if (isOwnCommunicationSubject(subjectName)) {
          return communicationGradeForStudentTerm(studentId, termValue);
        }

        if (isGrade12PerSubjectExport) {
          const assignedTerm = grade12SummarySubjectTerm(subjectName);

          if (assignedTerm) {
            if (termValue !== "final" && termValue !== assignedTerm) {
              return null;
            }

            return (
              grades.find(
                (grade) =>
                  grade.student_id === studentId &&
                  normalizeGrade12SummarySubject(grade.subject) ===
                    normalizeGrade12SummarySubject(subjectName) &&
                  grade.term === assignedTerm,
              )?.score ?? null
            );
          }
        }

        if (termValue === "final") {
          const values = ["1", "2", "3"].map(
            (value) =>
              grades.find(
                (grade) =>
                  grade.student_id === studentId &&
                  grade.subject === subjectName &&
                  grade.term === value,
              )?.score ?? null,
          );

          return computeAverage(values);
        }

        return (
          grades.find(
            (grade) =>
              grade.student_id === studentId &&
              grade.subject === subjectName &&
              grade.term === termValue,
          )?.score ?? null
        );
      };

      const subjectName = perSubjectPick;
      {
        const zip = await JSZip.loadAsync(templateBytes.slice(0));

        const sheetFile = zip.file("xl/worksheets/sheet1.xml");
        if (!sheetFile) {
          throw new Error(
            "The Summary of Final Grades per Subject template sheet was not found.",
          );
        }

        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetXml = await sheetFile.async("text");
        const sheetDocument = parser.parseFromString(
          sheetXml,
          "application/xml",
        );

        if (sheetDocument.querySelector("parsererror")) {
          throw new Error(
            "The Summary of Final Grades per Subject template XML is invalid.",
          );
        }

        // Keep this exporter separate from the existing Summary of Grades export.
        // It uses the provided per-subject XLSM template.
        setXmlText(
          sheetDocument,
          "E5",
          exportClass.region || exportProfile?.region || "",
        );
        setXmlText(
          sheetDocument,
          "U5",
          exportClass.division || exportProfile?.division || "",
        );
        setXmlText(
          sheetDocument,
          "E7",
          exportClass.school_name || exportProfile?.school_name || "",
        );
        setXmlText(
          sheetDocument,
          "U7",
          exportClass.school_id || exportProfile?.school_id || "",
        );
        setXmlText(
          sheetDocument,
          "F9",
          `${formatGradeLevel(exportClass.grade_level)} - ${
            exportClass.section || ""
          }`.trim(),
        );
        setXmlText(sheetDocument, "S9", exportClass.school_year || "");
        setXmlText(
          sheetDocument,
          "F10",
          exportClass.teacher_name || exportProfile?.full_name || "",
        );
        setXmlText(sheetDocument, "S10", subjectName);

        const maleStudents = students.filter(
          (student) => student.sex === "male",
        );
        const femaleStudents = students.filter(
          (student) => student.sex === "female",
        );

        const maleRows = Array.from({ length: 15 }, (_, index) => 14 + index);
        const femaleRows = Array.from({ length: 25 }, (_, index) => 30 + index);
        const learnerRows = [...maleRows, ...femaleRows];

        learnerRows.forEach((excelRow) => {
          clearXmlCell(sheetDocument, `B${excelRow}`);
          clearXmlCell(sheetDocument, `F${excelRow}`);
          clearXmlCell(sheetDocument, `J${excelRow}`);
          clearXmlCell(sheetDocument, `N${excelRow}`);
          clearXmlCell(sheetDocument, `R${excelRow}`);
          clearXmlCell(sheetDocument, `V${excelRow}`);
          clearXmlCell(sheetDocument, `Z${excelRow}`);
        });

        const writeGroup = (
          group: StudentRow[],
          excelRows: number[],
        ) => {
          group.slice(0, excelRows.length).forEach((student, index) => {
            const excelRow = excelRows[index];
            const learnerName =
              `${student.last_name}, ${student.first_name}${
                student.middle_name ? ` ${student.middle_name}` : ""
              }`.trim();

            const term1 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "1"),
            );
            const term2 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "2"),
            );
            const term3 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "3"),
            );
            const finalGrade = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "final"),
            );
            setXmlText(sheetDocument, `B${excelRow}`, learnerName);
            setXmlNumber(sheetDocument, `F${excelRow}`, term1);
            setXmlNumber(sheetDocument, `J${excelRow}`, term2);
            setXmlNumber(sheetDocument, `N${excelRow}`, term3);
            setXmlNumber(sheetDocument, `R${excelRow}`, finalGrade);
            setXmlText(
              sheetDocument,
              `V${excelRow}`,
              descriptor(finalGrade),
            );
            setXmlText(
              sheetDocument,
              `Z${excelRow}`,
              passFailRemark(finalGrade),
            );
          });
        };

        writeGroup(maleStudents, maleRows);
        writeGroup(femaleStudents, femaleRows);

        // Footer signatures in the provided XLSM template:
        //   B57 = "Prepared by:"
        //   B60 = adviser name
        //   B61 = "Class Adviser"
        //   R57 = "Noted:"
        //   T60 = principal name
        //   T61 = "School Principal"
        setXmlText(
          sheetDocument,
          "B60",
          exportClass.teacher_name || exportProfile?.full_name || "",
        );
        setXmlText(
          sheetDocument,
          "T60",
          exportClass.principal || exportProfile?.principal || "",
        );

        // Remove empty learner rows from the exported per-subject Excel.
        //
        // We hide the unused template rows instead of physically deleting
        // worksheet XML rows. In Excel they disappear/collapse completely,
        // while the original XLSM layout, formulas, merged cells and VBA
        // remain safe.
        maleRows.forEach((excelRow, index) => {
          setXmlRowHidden(
            sheetDocument,
            excelRow,
            index >= maleStudents.length,
          );
        });

        femaleRows.forEach((excelRow, index) => {
          setXmlRowHidden(
            sheetDocument,
            excelRow,
            index >= femaleStudents.length,
          );
        });

        zip.file(
          "xl/worksheets/sheet1.xml",
          serializer.serializeToString(sheetDocument),
        );

        await removeCalcChainFromXlsm(zip);

        const blob = await zip.generateAsync({
          type: "blob",
          mimeType:
            "application/vnd.ms-excel.sheet.macroEnabled.12",
          compression: "DEFLATE",
          compressionOptions: { level: 6 },
        });

        const fileName = [
          "Summary-of-Grades-per-Subject",
          subjectName,
          formatGradeLevel(exportClass.grade_level),
          exportClass.section || "Class",
        ]
          .map((part) => safeExcelFilePart(part))
          .filter(Boolean)
          .join("_");

        downloadXlsmBlob(blob, `${fileName}.xlsm`);
      }

      toast.success(`${perSubjectPick} Summary of Grades exported.`);
      setPerSubjectOpen(false);
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : "Could not export Summary of Grades per Subject.",
      );
    } finally {
      setIsExportingPerSubject(false);
    }
  };

  const exportFinalGradesAndGeneralAverage = async () => {
    if (!active || !klass) {
      toast.error("Select a class first.");
      return;
    }

    if (safeSubjects.length === 0) {
      toast.error("Add at least one subject before exporting.");
      return;
    }

    if (isExportingFinalGrades) return;

    setIsExportingFinalGrades(true);

    try {
      const { data: latestClassData, error: latestClassError } = await (
        supabase as any
      )
        .from("classes")
        .select("*")
        .eq("id", active)
        .single();

      if (latestClassError) throw latestClassError;

      const exportClass =
        (latestClassData as ClassWithTeacherName | null) ?? klass;

      const response = await fetch(
        FINAL_GRADES_AND_GENERAL_AVERAGE_TEMPLATE_URL,
        { cache: "no-store" },
      );

      if (!response.ok) {
        throw new Error(
          "Could not load /templates/FINAL-GRADES-AND-GENERAL-AVERAGE.xlsm. Put the original .xlsm file inside public/templates.",
        );
      }

      const templateBytes = await response.arrayBuffer();
      const zip = await JSZip.loadAsync(templateBytes);
      const sheetFile = zip.file("xl/worksheets/sheet1.xml");

      if (!sheetFile) {
        throw new Error(
          "The FINAL-GRADES-AND-GENERAL-AVERAGE template sheet was not found.",
        );
      }

      const parser = new DOMParser();
      const serializer = new XMLSerializer();
      const sheetXml = await sheetFile.async("text");
      const sheetDocument = parser.parseFromString(
        sheetXml,
        "application/xml",
      );

      if (sheetDocument.querySelector("parsererror")) {
        throw new Error(
          "The FINAL-GRADES-AND-GENERAL-AVERAGE template XML is invalid.",
        );
      }

      // School and class information in the provided XLSM template.
      setXmlText(
        sheetDocument,
        "D3",
        exportClass.region || exportProfile?.region || "",
      );
      setXmlText(
        sheetDocument,
        "AH3",
        exportClass.division || exportProfile?.division || "",
      );
      setXmlText(sheetDocument, "BF3", exportClass.school_year || "");
      setXmlText(
        sheetDocument,
        "D5",
        exportClass.school_name || exportProfile?.school_name || "",
      );
      setXmlText(
        sheetDocument,
        "AH5",
        exportClass.school_id || exportProfile?.school_id || "",
      );
      setXmlText(
        sheetDocument,
        "BF5",
        `${formatGradeLevel(exportClass.grade_level)} - ${
          exportClass.section || ""
        }`.trim(),
      );

      const isGrade12FinalExport =
        gradeLevelKey(exportClass.grade_level) === "12";

      const adviserName =
        exportClass.teacher_name || exportProfile?.full_name || "";
      setXmlText(
        sheetDocument,
        "A7",
        adviserName ? `ADVISER: ${adviserName}` : "ADVISER:",
      );
      setXmlText(sheetDocument, "B51", adviserName);
      setXmlText(
        sheetDocument,
        "BJ51",
        exportClass.principal || exportProfile?.principal || "",
      );

      // Each subject occupies four columns: Term 1, Term 2, Term 3 and Final.
      const maxSubjects = 15;
      const exportedSubjects = isGrade11ElectiveClass
        ? [
            "Academic Elective 1",
            "Academic Elective 2",
            "Academic Elective 3",
            ...safeSubjects.filter(
              (subjectName) =>
                normalizeText(subjectName) !==
                normalizeText(klass?.subject || "Elective Subject"),
            ),
          ].slice(0, maxSubjects)
        : isGrade12FinalExport
          ? safeSubjects
              .filter(
                (subjectName) =>
                  grade12SummarySubjectTerm(subjectName) != null,
              )
              .slice(0, maxSubjects)
          : safeSubjects.slice(0, maxSubjects);
      const subjectStartColumns = Array.from(
        { length: maxSubjects },
        (_, index) => 3 + index * 4,
      );

      widenFinalGradesTermColumns(sheetDocument);

      const maleStudents = students.filter(
        (student) => student.sex === "male",
      );
      const femaleStudents = students.filter(
        (student) => student.sex === "female",
      );
      const maleRows = Array.from({ length: 11 }, (_, index) => 11 + index);
      const femaleRows = Array.from({ length: 24 }, (_, index) => 23 + index);
      const learnerRows = [...maleRows, ...femaleRows];

      const scoreForSubjectTerm = (
        studentId: string,
        subjectName: string,
        termValue: "1" | "2" | "3",
      ) => {
        if (isOwnMapehSubject(subjectName)) {
          return mapehGradeForStudentTerm(studentId, termValue);
        }

        if (isOwnCommunicationSubject(subjectName)) {
          return communicationGradeForStudentTerm(studentId, termValue);
        }

        if (isGrade12FinalExport) {
          const assignedTerm = grade12SummarySubjectTerm(subjectName);

          if (assignedTerm) {
            if (termValue !== assignedTerm) {
              return null;
            }

            return (
              grades.find(
                (grade) =>
                  grade.student_id === studentId &&
                  normalizeGrade12SummarySubject(grade.subject) ===
                    normalizeGrade12SummarySubject(subjectName) &&
                  grade.term === assignedTerm,
              )?.score ?? null
            );
          }
        }

        if (isGrade11ElectiveClass) {
          const mappedTerm = academicElectiveTerm(subjectName);
          if (mappedTerm) {
            // This subject exists in one term only.
            return termValue === mappedTerm
              ? grades.find(
                  (grade) =>
                    grade.student_id === studentId &&
                    normalizeText(grade.subject) ===
                      normalizeText(klass?.subject || "Elective Subject") &&
                    grade.term === mappedTerm,
                )?.score ?? null
              : null;
          }
        }

        return (
          grades.find(
            (grade) =>
              grade.student_id === studentId &&
              grade.subject === subjectName &&
              grade.term === termValue,
          )?.score ?? null
        );
      };

      const finalScoreForSubject = (
        studentId: string,
        subjectName: string,
      ) => {
        if (isOwnMapehSubject(subjectName)) {
          return mapehGradeForStudentTerm(studentId, "final");
        }

        if (isOwnCommunicationSubject(subjectName)) {
          return communicationGradeForStudentTerm(studentId, "final");
        }

        if (isGrade12FinalExport) {
          const assignedTerm = grade12SummarySubjectTerm(subjectName);

          if (assignedTerm) {
            return (
              grades.find(
                (grade) =>
                  grade.student_id === studentId &&
                  normalizeGrade12SummarySubject(grade.subject) ===
                    normalizeGrade12SummarySubject(subjectName) &&
                  grade.term === assignedTerm,
              )?.score ?? null
            );
          }
        }

        if (isGrade11ElectiveClass) {
          const mappedTerm = academicElectiveTerm(subjectName);
          if (mappedTerm) {
            // FINAL for each Academic Elective is simply its own one-term grade.
            return grades.find(
              (grade) =>
                grade.student_id === studentId &&
                normalizeText(grade.subject) ===
                  normalizeText(klass?.subject || "Elective Subject") &&
                grade.term === mappedTerm,
            )?.score ?? null;
          }
        }

        // IMPORTANT:
        // The Excel FINAL grade must match the FINAL value shown in
        // Summary of Grades. Always recalculate it from Term 1, Term 2,
        // and Term 3 instead of using an older stored "final" grade row.
        //
        // Example:
        // Term 1 = 64, Term 2 = 90, Term 3 = 88
        // (64 + 90 + 88) / 3 = 80.67 -> rounded FINAL = 81
        return computeAverage(
          (["1", "2", "3"] as const).map((termValue) =>
            scoreForSubjectTerm(studentId, subjectName, termValue),
          ),
        );
      };

      // Clear all template learner, subject and grade cells first.
      subjectStartColumns.forEach((startColumn) => {
        setXmlText(sheetDocument, `${excelColumnName(startColumn)}7`, "");
      });

      learnerRows.forEach((excelRow) => {
        clearXmlCell(sheetDocument, `B${excelRow}`);
        clearXmlCell(sheetDocument, `BK${excelRow}`);
        clearXmlCell(sheetDocument, `BL${excelRow}`);
        clearXmlCell(sheetDocument, `BM${excelRow}`);

        subjectStartColumns.forEach((startColumn) => {
          for (let offset = 0; offset < 4; offset += 1) {
            clearXmlCell(
              sheetDocument,
              `${excelColumnName(startColumn + offset)}${excelRow}`,
            );
          }
        });
      });

      exportedSubjects.forEach((subjectName, subjectIndex) => {
        const startColumn = subjectStartColumns[subjectIndex];
        setXmlText(
          sheetDocument,
          `${excelColumnName(startColumn)}7`,
          subjectName,
        );
      });

      const writeLearnerGroup = (
        group: StudentRow[],
        excelRows: number[],
      ) => {
        group.slice(0, excelRows.length).forEach((student, index) => {
          const excelRow = excelRows[index];
          const learnerName =
            `${student.last_name}, ${student.first_name}${
              student.middle_name ? ` ${student.middle_name}` : ""
            }`.trim();

          setXmlText(sheetDocument, `B${excelRow}`, learnerName);

          const finalGrades: Array<number | null> = [];

          exportedSubjects.forEach((subjectName, subjectIndex) => {
            const startColumn = subjectStartColumns[subjectIndex];
            const term1 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "1"),
            );
            const term2 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "2"),
            );
            const term3 = roundedGrade(
              scoreForSubjectTerm(student.id, subjectName, "3"),
            );
            const finalGrade = roundedGrade(
              finalScoreForSubject(student.id, subjectName),
            );

            setXmlNumber(
              sheetDocument,
              `${excelColumnName(startColumn)}${excelRow}`,
              term1,
            );
            setXmlNumber(
              sheetDocument,
              `${excelColumnName(startColumn + 1)}${excelRow}`,
              term2,
            );
            setXmlNumber(
              sheetDocument,
              `${excelColumnName(startColumn + 2)}${excelRow}`,
              term3,
            );
            setXmlNumber(
              sheetDocument,
              `${excelColumnName(startColumn + 3)}${excelRow}`,
              finalGrade,
            );

            finalGrades.push(finalGrade);
          });

          const generalAverage = computeAverage(finalGrades);
          const roundedGeneralAverage = roundedGrade(generalAverage);

          setXmlNumber(sheetDocument, `BK${excelRow}`, generalAverage);
          setXmlNumber(
            sheetDocument,
            `BL${excelRow}`,
            roundedGeneralAverage,
          );
          setXmlText(
            sheetDocument,
            `BM${excelRow}`,
            roundedGeneralAverage == null
              ? ""
              : roundedGeneralAverage >= 75
                ? "PROMOTED"
                : "RETAINED",
          );
        });
      };

      writeLearnerGroup(maleStudents, maleRows);
      writeLearnerGroup(femaleStudents, femaleRows);

      maleRows.forEach((excelRow, index) => {
        setXmlRowHidden(
          sheetDocument,
          excelRow,
          index >= maleStudents.length,
        );
      });
      femaleRows.forEach((excelRow, index) => {
        setXmlRowHidden(
          sheetDocument,
          excelRow,
          index >= femaleStudents.length,
        );
      });

      zip.file(
        "xl/worksheets/sheet1.xml",
        serializer.serializeToString(sheetDocument),
      );
      await removeCalcChainFromXlsm(zip);

      const blob = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });

      const fileName = [
        "Final-Grades-and-General-Average",
        formatGradeLevel(exportClass.grade_level),
        exportClass.section || "Class",
      ]
        .map((part) => safeExcelFilePart(part))
        .filter(Boolean)
        .join("_");

      downloadXlsmBlob(blob, `${fileName}.xlsm`);

      const warnings: string[] = [];
      if (safeSubjects.length > maxSubjects) {
        warnings.push(
          `Only the first ${maxSubjects} subjects were exported because of the template limit.`,
        );
      }
      if (maleStudents.length > maleRows.length) {
        warnings.push(
          `Only the first ${maleRows.length} male learners were exported because of the template limit.`,
        );
      }
      if (femaleStudents.length > femaleRows.length) {
        warnings.push(
          `Only the first ${femaleRows.length} female learners were exported because of the template limit.`,
        );
      }

      if (warnings.length > 0) {
        toast.warning(warnings.join(" "));
      } else {
        toast.success(
          "Final grades, general averages and learner names exported.",
        );
      }
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : "Could not export FINAL-GRADES-AND-GENERAL-AVERAGE.",
      );
    } finally {
      setIsExportingFinalGrades(false);
    }
  };

  const exportSummaryToExcel = async () => {
    if (!active || !klass) {
      toast.error("Select a class first.");
      return;
    }

    if (isExportingExcel) return;

    setIsExportingExcel(true);

    try {
      // Fetch the selected class again at export time so the workbook receives
      // the latest school/class information shown on the Class page.
      const { data: latestClassData, error: latestClassError } = await (
        supabase as any
      )
        .from("classes")
        .select("*")
        .eq("id", active)
        .single();

      if (latestClassError) throw latestClassError;

      const exportClass =
        (latestClassData as ClassWithTeacherName | null) ?? klass;

      const response = await fetch(SUMMARY_EXCEL_TEMPLATE_URL, {
        cache: "no-store",
      });

      if (!response.ok) {
        throw new Error(
          "Could not load /templates/Summary-of-grades.xlsm. Put the original .xlsm file inside public/templates.",
        );
      }

      const templateBytes = await response.arrayBuffer();
      const zip = await JSZip.loadAsync(templateBytes);

      // The workbook is edited directly as XML instead of being rewritten
      // through ExcelJS. This preserves the original .xlsm VBA project.
      const sheetFile = zip.file("xl/worksheets/sheet1.xml");
      if (!sheetFile) {
        throw new Error("The Summary of Grades template sheet was not found.");
      }

      const parser = new DOMParser();
      const serializer = new XMLSerializer();
      const sheetXml = await sheetFile.async("text");
      const sheetDocument = parser.parseFromString(
        sheetXml,
        "application/xml",
      );

      if (sheetDocument.querySelector("parsererror")) {
        throw new Error("The Summary of Grades template XML is invalid.");
      }

      const termTitle =
        term === "final"
          ? "SUMMARY OF FINAL GRADES"
          : `SUMMARY OF TERM ${term} GRADES`;

      setXmlText(sheetDocument, "A1", termTitle);

      // School/class information.
      // Use the latest selected class values first because these are the same
      // fields displayed on the Class page. Fall back to profile values only
      // when a class field is empty.
      setXmlText(
        sheetDocument,
        "C3",
        exportClass.region || exportProfile?.region || "",
      );
      setXmlText(
        sheetDocument,
        "J3",
        exportClass.division || exportProfile?.division || "",
      );
      setXmlText(
        sheetDocument,
        "P3",
        exportClass.school_year || "",
      );
      setXmlText(
        sheetDocument,
        "C5",
        exportClass.school_name || exportProfile?.school_name || "",
      );
      setXmlText(
        sheetDocument,
        "J5",
        exportClass.school_id || exportProfile?.school_id || "",
      );
      setXmlText(
        sheetDocument,
        "P5",
        `${formatGradeLevel(exportClass.grade_level)} - ${
          exportClass.section || ""
        }`.trim(),
      );

      // The provided macro-enabled template supports up to 15 subject columns
      // from C through Q.
      const maxSubjectColumns = 15;
      const exportedSubjects = displayedSubjects.slice(0, maxSubjectColumns);

      const compactTableColumns = compactSummaryGradeTable(
        sheetDocument,
        exportedSubjects.length,
      );
      const averageColumn = excelColumnName(compactTableColumns.averageColumn);
      const roundedAverageColumn = excelColumnName(
        compactTableColumns.roundedAverageColumn,
      );
      const remarksColumn = excelColumnName(compactTableColumns.remarksColumn);

      // Keep the original template column widths for the worksheet header.
      // Do NOT hide C:Q here because column hiding affects the entire sheet,
      // including REGION / DIVISION / SCHOOL YEAR / SCHOOL NAME / SCHOOL ID /
      // GRADE and SECTION. The table cells are still populated only for the
      // actual subjects below.
      for (let index = 0; index < maxSubjectColumns; index += 1) {
        const column = excelColumnName(3 + index);
        const subjectName = exportedSubjects[index] ?? "";

        if (index < exportedSubjects.length) {
          setXmlText(sheetDocument, `${column}7`, subjectName);
        } else if (
          column !== averageColumn &&
          column !== roundedAverageColumn &&
          column !== remarksColumn
        ) {
          clearXmlCellVisual(sheetDocument, `${column}7`);
        }
      }

      // Move only the grade-table labels beside the last real subject.
      // Workbook header rows 1:6 stay exactly where the template put them.
      setXmlText(sheetDocument, `${averageColumn}7`, "Average");
      clearXmlCell(sheetDocument, `${roundedAverageColumn}7`);
      setXmlText(sheetDocument, `${remarksColumn}7`, "Remarks");

      // Remove the old table labels at R:S:T only when they are outside the
      // compact Average / helper / expanded Remarks area.
      const compactUsedColumns = new Set<number>([
        compactTableColumns.averageColumn,
        compactTableColumns.roundedAverageColumn,
        ...Array.from(
          {
            length:
              compactTableColumns.remarksEndColumn -
              compactTableColumns.remarksColumn +
              1,
          },
          (_, index) => compactTableColumns.remarksColumn + index,
        ),
      ]);

      [18, 19, 20].forEach((columnNumber) => {
        if (!compactUsedColumns.has(columnNumber)) {
          clearXmlCellVisual(
            sheetDocument,
            `${excelColumnName(columnNumber)}7`,
          );
        }
      });

      // Clear all learner rows first so old template values/formulas do not
      // appear when the current class contains fewer learners.
      const maleRows = Array.from({ length: 11 }, (_, index) => 9 + index);
      const femaleRows = Array.from({ length: 24 }, (_, index) => 21 + index);
      const allLearnerRows = [...maleRows, ...femaleRows];

      allLearnerRows.forEach((excelRow) => {
        clearXmlCell(sheetDocument, `B${excelRow}`);

        for (let subjectIndex = 0; subjectIndex < maxSubjectColumns; subjectIndex += 1) {
          const column = excelColumnName(3 + subjectIndex);
          clearXmlCell(sheetDocument, `${column}${excelRow}`);
        }

        clearXmlCell(sheetDocument, `${averageColumn}${excelRow}`);
        clearXmlCell(sheetDocument, `${roundedAverageColumn}${excelRow}`);
        clearXmlCell(sheetDocument, `${remarksColumn}${excelRow}`);

        [18, 19, 20].forEach((columnNumber) => {
          const isUsedByCompactTable =
            columnNumber === compactTableColumns.averageColumn ||
            columnNumber === compactTableColumns.roundedAverageColumn ||
            (columnNumber >= compactTableColumns.remarksColumn &&
              columnNumber <= compactTableColumns.remarksEndColumn);

          if (!isUsedByCompactTable) {
            clearXmlCellVisual(
              sheetDocument,
              `${excelColumnName(columnNumber)}${excelRow}`,
            );
          }
        });
      });

      const writeLearnerGroup = (
        groupRows: typeof rows,
        excelRows: number[],
      ) => {
        groupRows.slice(0, excelRows.length).forEach((row, index) => {
          const excelRow = excelRows[index];
          const learnerName =
            `${row.st.last_name}, ${row.st.first_name}${
              row.st.middle_name ? ` ${row.st.middle_name}` : ""
            }`.trim();

          setXmlText(sheetDocument, `B${excelRow}`, learnerName);

          exportedSubjects.forEach((subjectName, subjectIndex) => {
            const column = excelColumnName(3 + subjectIndex);
            const score = scoreFor(row.st.id, subjectName);
            setXmlNumber(
              sheetDocument,
              `${column}${excelRow}`,
              roundedGrade(score),
            );
          });

          const rawAverage = row.avg;
          const roundedAverage = roundedGrade(rawAverage);

          setXmlNumber(
            sheetDocument,
            `${averageColumn}${excelRow}`,
            rawAverage,
          );
          setXmlNumber(
            sheetDocument,
            `${roundedAverageColumn}${excelRow}`,
            roundedAverage,
          );

          const awardInfo = award(roundedAverage);
          setXmlText(
            sheetDocument,
            `${remarksColumn}${excelRow}`,
            awardInfo?.label ?? "",
          );
        });
      };

      writeLearnerGroup(male, maleRows);
      writeLearnerGroup(female, femaleRows);

      // Remove empty learner rows from the MAIN Summary of Grades Excel export.
      // Keep only rows that actually contain students.
      // The MALE/FEMALE separator rows remain visible.
      maleRows.forEach((excelRow, index) => {
        setXmlRowHidden(
          sheetDocument,
          excelRow,
          index >= male.length,
        );
      });

      femaleRows.forEach((excelRow, index) => {
        setXmlRowHidden(
          sheetDocument,
          excelRow,
          index >= female.length,
        );
      });

      // Signature names in the original template.
      setXmlText(
        sheetDocument,
        "B49",
        exportClass.teacher_name || exportProfile?.full_name || "",
      );
      setXmlText(
        sheetDocument,
        "R49",
        exportClass.principal || exportProfile?.principal || "",
      );

      zip.file(
        "xl/worksheets/sheet1.xml",
        serializer.serializeToString(sheetDocument),
      );

      // Remove the old formula calculation chain because learner cells are
      // now written as values. VBA and all other .xlsm parts remain intact.
      await removeCalcChainFromXlsm(zip);

      const blob = await zip.generateAsync({
        type: "blob",
        mimeType:
          "application/vnd.ms-excel.sheet.macroEnabled.12",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });

      const fileName = [
        "Summary-of-Grades",
        formatGradeLevel(exportClass.grade_level),
        exportClass.section || "Class",
        term === "final" ? "Final" : `Term-${term}`,
      ]
        .map((part) => safeExcelFilePart(part))
        .filter(Boolean)
        .join("_");

      downloadXlsmBlob(blob, `${fileName}.xlsm`);

      if (safeSubjects.length > maxSubjectColumns) {
        toast.warning(
          `The template has room for ${maxSubjectColumns} subjects. Only the first ${maxSubjectColumns} were exported.`,
        );
      } else {
        toast.success("Summary of Grades Excel exported.");
      }
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : "Could not export the Summary of Grades Excel file.",
      );
    } finally {
      setIsExportingExcel(false);
    }
  };

  useEffect(() => {
    const normalizedHash = String(locationHash ?? "").replace(/^#/, "");

    if (
      normalizedHash !== "send-grade-request" ||
      !canSendGradeRequests ||
      !active
    ) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      document.getElementById("send-grade-request")?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    }, 80);

    return () => window.clearTimeout(timeoutId);
  }, [locationHash, canSendGradeRequests, active, requests.length]);

  const selectedClassDataLoading =
    Boolean(active && klass) &&
    (!hydrated ||
      studentsLoading ||
      studentsPlaceholder ||
      gradesLoading ||
      gradesPlaceholder ||
      requestsLoading ||
      requestsPlaceholder ||
      recipientProfilesLoading ||
      recipientProfilesPlaceholder ||
      (isMapehClass &&
        (mapehComponentsLoading ||
          mapehComponentsPlaceholder ||
          mapehActivitiesLoading ||
          mapehActivitiesPlaceholder ||
          mapehScoresLoading ||
          mapehScoresPlaceholder ||
          mapehSavedBasesLoading ||
          mapehSavedBasesPlaceholder)) ||
      (isCommunicationClass &&
        (communicationComponentsLoading ||
          communicationComponentsPlaceholder ||
          communicationActivitiesLoading ||
          communicationActivitiesPlaceholder ||
          communicationScoresLoading ||
          communicationScoresPlaceholder ||
          communicationSavedBasesLoading ||
          communicationSavedBasesPlaceholder)) ||
      (teacherType === "class_adviser" &&
        (receivedRequestsLoading ||
          receivedRequestsPlaceholder ||
          finalizedRequestFormsLoading ||
          finalizedRequestFormsPlaceholder)));

  if (
    teacherTypeLoading ||
    classesLoading ||
    classesPlaceholder ||
    selectedClassDataLoading
  ) {
    return <SummaryOfGradesSkeleton />;
  }

  return (
    <div className="min-w-0 space-y-4 pb-24 md:pb-6">
      <div className="min-w-0 rounded-2xl border bg-card p-4 shadow-sm sm:p-5">
        <div className="flex min-w-0 flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
          <div className="flex min-w-0 flex-1 flex-col gap-3 sm:flex-row sm:items-center">
            <Link
              to={dashboardPath}
              className="inline-flex w-fit shrink-0 items-center gap-1 rounded-lg border px-3 py-1.5 text-sm hover:bg-muted"
            >
              <ArrowLeft className="size-4" />
              Back
            </Link>

            <div className="min-w-0 flex-1">
              <div className="flex min-w-0 flex-col gap-2 sm:flex-row sm:flex-wrap sm:items-center">
                <div className="flex min-w-0 items-center gap-2 text-lg font-semibold">
                  <Trophy className="size-5 shrink-0 text-primary" />
                  <span className="min-w-0">Summary of Grades</span>
                </div>

                <Select
                  value={active || undefined}
                  onValueChange={setClassId}
                  disabled={classes.length === 0}
                >
                  <SelectTrigger className="h-9 w-full min-w-0 sm:w-48">
                    <SelectValue placeholder="Pick class" />
                  </SelectTrigger>
                  <SelectContent>
                    {classes.map((classItem) => (
                      <SelectItem key={classItem.id} value={classItem.id}>
                        {formatGradeLevel(classItem.grade_level)} ·{" "}
                        {classItem.section || "—"}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="mt-1 break-words text-xs text-muted-foreground">
                Adviser: {klass?.teacher_name || "—"} · Subject: (auto-pulled)
              </div>
            </div>
          </div>

          <span className="chip w-fit shrink-0 bg-[color:var(--success-bg)] text-[color:var(--success)]">
            Autosaved
          </span>
        </div>

        {teacherType === "class_adviser" && (
          <div className="mt-4 grid min-w-0 grid-cols-1 gap-2 sm:grid-cols-2 xl:grid-cols-3">
            <Button
              type="button"
              size="sm"
              onClick={exportFinalGradesAndGeneralAverage}
              disabled={!active || safeSubjects.length === 0 || isExportingFinalGrades}
              className="h-auto min-h-10 w-full min-w-0 whitespace-normal bg-emerald-600 px-3 py-2 text-center text-white hover:bg-emerald-700"
            >
              {isExportingFinalGrades ? (
                <LoaderCircle className="mr-1.5 size-4 shrink-0 animate-spin" />
              ) : (
                <FileSpreadsheet className="mr-1.5 size-4 shrink-0" />
              )}
              <span className="min-w-0 break-words">
                {isExportingFinalGrades ? "Exporting..." : "Final Grades & General Average"}
              </span>
            </Button>

            <Button
              type="button"
              size="sm"
              onClick={() => {
                setPerSubjectPick((current) =>
                  current && safeSubjects.includes(current)
                    ? current
                    : safeSubjects[0] ?? "",
                );
                setPerSubjectOpen(true);
              }}
              disabled={!active || safeSubjects.length === 0}
              className="h-auto min-h-10 w-full min-w-0 whitespace-normal bg-emerald-600 px-3 py-2 text-center text-white hover:bg-emerald-700"
            >
              <FileSpreadsheet className="mr-1.5 size-4 shrink-0" />
              <span className="min-w-0 break-words">Summary of Grades per Subject</span>
            </Button>

            <Button
              type="button"
              size="sm"
              onClick={exportSummaryToExcel}
              disabled={!active || isExportingExcel}
              className="h-auto min-h-10 w-full min-w-0 whitespace-normal bg-emerald-600 px-3 py-2 text-center text-white hover:bg-emerald-700 sm:col-span-2 xl:col-span-1"
            >
              {isExportingExcel ? (
                <LoaderCircle className="mr-1.5 size-4 shrink-0 animate-spin" />
              ) : (
                <FileSpreadsheet className="mr-1.5 size-4 shrink-0" />
              )}
              <span className="min-w-0 break-words">
                {isExportingExcel ? "Exporting..." : "Summary of Grades"}
              </span>
            </Button>
          </div>
        )}
      </div>

      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex gap-1 rounded-xl border bg-card p-1">
          {TERMS.map((termItem) => (
            <button
              key={termItem.value}
              type="button"
              onClick={() => setTerm(termItem.value)}
              className={`rounded-lg px-4 py-1.5 text-sm font-medium ${
                term === termItem.value
                  ? "bg-primary text-primary-foreground"
                  : "hover:bg-muted"
              }`}
            >
              {termItem.label}
            </button>
          ))}
        </div>

        <div className="flex w-full flex-col gap-2 sm:w-auto sm:flex-row sm:items-center">
          {canOpenAnalytics && (
            <Button
              type="button"
              variant={analyticsOpen ? "default" : "outline"}
              size="sm"
              className={`w-full sm:w-auto ${
                analyticsOpen
                  ? "bg-emerald-600 text-white hover:bg-emerald-700"
                  : ""
              }`}
              disabled={!active}
              onClick={() => setAnalyticsOpen((current) => !current)}
              aria-expanded={analyticsOpen}
              aria-controls="summary-analytics-insights"
            >
              <BarChart3 className="mr-1 size-4" />
              Analytics & Insights
            </Button>
          )}

          {teacherType === "class_adviser" && (
            <Button
              variant="outline"
              size="sm"
              className="w-full sm:w-auto"
              disabled={
                !active ||
                receivedRequestsLoading ||
                finalizedRequestFormsLoading ||
                allAddSubjectForms.length === 0
              }
              onClick={() => setAddOpen(true)}
              aria-haspopup="dialog"
              aria-expanded={addOpen}
            >
              <Plus className="mr-1 size-4" />
              Add Subject
            </Button>
          )}
        </div>
      </div>

      {analyticsOpen && canOpenAnalytics && active && klass && (
        <>
          <button
            type="button"
            aria-label="Close Analytics & Insights"
            className="fixed inset-0 z-40 cursor-default bg-black/70 backdrop-blur-[1px]"
            onClick={() => setAnalyticsOpen(false)}
          />

          <div
            id="summary-analytics-insights"
            role="dialog"
            aria-modal="true"
            aria-labelledby="analytics-insights-title"
            className="fixed left-1/2 top-1/2 z-50 flex max-h-[92vh] w-[calc(100%-1rem)] max-w-[1240px] -translate-x-1/2 -translate-y-1/2 flex-col overflow-hidden rounded-2xl border border-[#e5d4c5] bg-[#fffaf3] shadow-2xl sm:w-[calc(100%-2rem)]"
          >
            <div className="flex shrink-0 items-center justify-between gap-3 border-b border-[#eadfd4] bg-[#fffdf9] px-4 py-3 sm:px-5">
              <div className="min-w-0">
                <div className="flex items-center gap-2">
                  <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-emerald-100 text-emerald-700">
                    <BarChart3 className="size-4" />
                  </div>

                  <div className="min-w-0">
                    <h2
                      id="analytics-insights-title"
                      className="truncate text-base font-semibold text-[#33241f] sm:text-lg"
                    >
                      Analytics & Insights
                    </h2>
                    <p className="truncate text-xs text-[#8a756b]">
                      {formatGradeLevel(klass.grade_level)} · {klass.section || "—"} ·{" "}
                      {activeAnalyticsSubject || "—"}
                    </p>
                    <p className="mt-0.5 hidden text-[10px] text-emerald-700 sm:block">
                      Term 1 + Term 2 → Gaussian Naive Bayes → Term 3 proficiency forecast
                    </p>
                  </div>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setAnalyticsOpen(false)}
                className="grid size-9 shrink-0 place-items-center rounded-lg border border-[#e5d4c5] bg-white text-[#6f554a] shadow-sm transition hover:bg-[#f7eee6] hover:text-[#33241f]"
                aria-label="Close Analytics & Insights"
                title="Close"
              >
                <X className="size-4" />
              </button>
            </div>

            <div className="min-h-0 flex-1 overflow-y-auto bg-[#fffaf3] p-3 sm:p-4">
              <AnalyticsInsightsPanel
                key={active}
                classId={active}
                students={students}
                subject={activeAnalyticsSubject}
                subjectOptions={analyticsSubjectOptions}
                onSubjectChange={setSelectedAnalyticsSubject}
                summarySubjectGrades={analyticsSummarySubjectGrades}
                allowComponentForecast={analyticsCanUseOwnComponents}
                initialTerm={
                  term === "1" || term === "2" || term === "3" || term === "final"
                    ? term
                    : "1"
                }
              />
            </div>
          </div>
        </>
      )}

      <div className="overflow-x-auto rounded-2xl border bg-card">
        <table className="w-full min-w-[760px] text-xs">
          <thead className="bg-muted/50">
            <tr className="border-b">
              <th className="w-8 px-2 py-2 text-left">#</th>
              <th className="px-2 py-2 text-left">LEARNER'S NAME</th>
              {displayedSubjects.map((subjectName) => {
                const academicElectiveVirtualTerm =
                  isGrade11ElectiveClass
                    ? academicElectiveTerm(subjectName)
                    : null;
                const storageSubjectName = academicElectiveVirtualTerm
                  ? String(klass?.subject || "Elective Subject")
                  : subjectName;
                const canDragSubject = academicElectiveVirtualTerm == null;

                const isDragging =
                  canDragSubject && draggedSubject === subjectName;
                const isDropTarget =
                  canDragSubject &&
                  dragOverSubject === subjectName &&
                  draggedSubject !== subjectName;

                return (
                  <th
                    key={subjectName}
                    draggable={canDragSubject}
                    onDragStart={(event) => {
                      if (!canDragSubject) {
                        event.preventDefault();
                        return;
                      }

                      setDraggedSubject(subjectName);
                      setDragOverSubject(null);
                      event.dataTransfer.effectAllowed = "move";
                      event.dataTransfer.setData("text/plain", subjectName);
                    }}
                    onDragEnter={(event) => {
                      event.preventDefault();
                      if (draggedSubject && draggedSubject !== subjectName) {
                        setDragOverSubject(subjectName);
                      }
                    }}
                    onDragOver={(event) => {
                      event.preventDefault();
                      event.dataTransfer.dropEffect = "move";
                      if (draggedSubject && draggedSubject !== subjectName) {
                        setDragOverSubject(subjectName);
                      }
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      const source =
                        draggedSubject ||
                        event.dataTransfer.getData("text/plain");

                      reorderSubjects(source, subjectName);
                      setDraggedSubject(null);
                      setDragOverSubject(null);
                    }}
                    onDragEnd={() => {
                      setDraggedSubject(null);
                      setDragOverSubject(null);
                    }}
                    className={`border-l px-2 py-2 text-center transition ${
                      isDragging
                        ? "cursor-grabbing bg-emerald-50/80 opacity-60"
                        : canDragSubject
                          ? "cursor-grab"
                          : "cursor-default"
                    } ${
                      isDropTarget
                        ? "bg-emerald-50 ring-2 ring-inset ring-emerald-400"
                        : ""
                    }`}
                    title="Drag left or right to reorder this subject"
                  >
                    <div className="flex items-center justify-center gap-1.5">
                      {teacherType !== "class_adviser" && (
                        <GripVertical
                          className={`size-3.5 shrink-0 ${
                            isDragging
                              ? "text-emerald-600"
                              : "text-muted-foreground"
                          }`}
                          aria-hidden="true"
                        />
                      )}

                      <span>{subjectName}</span>

                      <button
                        type="button"
                        draggable={false}
                        onMouseDown={(event) => event.stopPropagation()}
                        onClick={(event) => {
                          event.stopPropagation();
                          setSubjects((previous) =>
                            (Array.isArray(previous) ? previous : []).filter(
                              (value) =>
                                normalizeText(value) !==
                                normalizeText(storageSubjectName),
                            ),
                          );
                        }}
                        className="text-muted-foreground hover:text-destructive"
                        aria-label={`Remove ${subjectName}`}
                        title={`Remove ${subjectName}`}
                      >
                        <X className="size-3" />
                      </button>
                    </div>
                  </th>
                );
              })}
              <th className="border-l px-2 py-2 text-center">AVG</th>
              <th className="border-l px-2 py-2 text-center">RANK</th>
              <th className="border-l px-2 py-2 text-center">DESCRIPTOR</th>
              <th className="border-l px-2 py-2 text-center">AWARD</th>
            </tr>
          </thead>
          <tbody>
            {[
              { label: "MALE", list: male },
              { label: "FEMALE", list: female },
              ...(other.length > 0 ? [{ label: "OTHER", list: other }] : []),
            ].map((group) => (
              <Fragment key={group.label}>
                <tr className="border-b bg-muted/30">
                  <td
                    colSpan={displayedSubjects.length + 6}
                    className="px-2 py-1 text-[11px] font-semibold tracking-wide"
                  >
                    {group.label}
                  </td>
                </tr>
                {group.list.map((row, index) => {
                  const roundedAverage = roundedGrade(row.avg);
                  const awardInfo = award(roundedAverage);
                  return (
                    <tr key={row.st.id} className="border-b">
                      <td className="px-2 py-1 text-muted-foreground">
                        {index + 1}
                      </td>
                      <td className="whitespace-nowrap px-2 py-1 font-medium">
                        {row.st.last_name}, {row.st.first_name}
                      </td>
                      {displayedSubjects.map((subjectName, subjectIndex) => (
                        <td
                          key={`${subjectName}-${term}`}
                          className="border-l px-1 py-1 text-center font-medium"
                        >
                          {roundedGrade(row.scores[subjectIndex]) ?? "—"}
                        </td>
                      ))}
                      <td className="border-l px-2 py-1 text-center font-semibold">
                        {roundedAverage ?? "—"}
                      </td>
                      <td className="border-l px-2 py-1 text-center">
                        {ranked.get(row.st.id) ?? "—"}
                      </td>
                      <td className="border-l px-2 py-1 text-center text-[11px]">
                        {descriptor(roundedAverage)}
                      </td>
                      <td
                        className={`border-l px-2 py-1 text-center text-[11px] font-medium ${
                          awardInfo?.tone || ""
                        }`}
                      >
                        {awardInfo ? (
                          <span className="inline-flex items-center gap-1">
                            <Trophy className="size-3" />
                            {awardInfo.label}
                          </span>
                        ) : (
                          "—"
                        )}
                      </td>
                    </tr>
                  );
                })}
              </Fragment>
            ))}
            {rows.length === 0 && (
              <tr>
                <td
                  colSpan={displayedSubjects.length + 6}
                  className="p-8 text-center text-muted-foreground"
                >
                  No learners yet. Add students in the class roster.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {active && canSendGradeRequests && (
        <RequestPanel
          classInfo={klass}
          students={students}
          requests={requests}
          recipientProfiles={recipientProfiles}
          resolveGrade={scoreForTerm}
          onChange={() =>
            queryClient.invalidateQueries({
              queryKey: ["grade-requests", active],
            })
          }
        />
      )}

      {perSubjectOpen && (
        <>
          <button
            type="button"
            aria-label="Close Summary of Grades per Subject"
            className="fixed inset-0 z-40 cursor-default bg-black/55 backdrop-blur-[1px]"
            onClick={() => setPerSubjectOpen(false)}
          />

          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="per-subject-export-title"
            className="fixed left-1/2 top-1/2 z-50 w-[calc(100%-2rem)] max-w-2xl -translate-x-1/2 -translate-y-1/2 overflow-hidden rounded-2xl border border-[#eadfd4] bg-[#fffdf9] shadow-2xl"
          >
            <div className="flex items-center justify-between border-b border-[#eadfd4] px-5 py-4">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="size-5 text-emerald-600" />
                <h2
                  id="per-subject-export-title"
                  className="text-lg font-semibold text-[#33241f]"
                >
                  Summary of Grades per Subject
                </h2>
              </div>

              <button
                type="button"
                onClick={() => setPerSubjectOpen(false)}
                className="rounded-lg p-2 text-[#6f554a] transition hover:bg-[#f5ece4] hover:text-[#33241f]"
                aria-label="Close Summary of Grades per Subject"
              >
                <X className="size-5" />
              </button>
            </div>

            <div className="space-y-4 p-5">
              <div className="grid grid-cols-2 gap-4 rounded-xl border border-[#eadfd4] bg-[#fffaf3] p-4 sm:grid-cols-4">
                <div>
                  <div className="text-xs text-[#8a756b]">Grade Level</div>
                  <div className="mt-1 text-sm font-semibold text-[#33241f]">
                    {klass ? formatGradeLevel(klass.grade_level) : "—"}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-[#8a756b]">Section</div>
                  <div className="mt-1 text-sm font-semibold text-[#33241f]">
                    {klass?.section || "—"}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-[#8a756b]">School Year</div>
                  <div className="mt-1 text-sm font-semibold text-[#33241f]">
                    {klass?.school_year || "—"}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-[#8a756b]">Students</div>
                  <div className="mt-1 text-sm font-semibold text-[#33241f]">
                    {students.length}
                  </div>
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  Select Subject
                </label>
                <Select
                  value={perSubjectPick}
                  onValueChange={setPerSubjectPick}
                >
                  <SelectTrigger className="h-11 border-[#dfcfc1] bg-white">
                    <SelectValue placeholder="Select a subject" />
                  </SelectTrigger>
                  <SelectContent>
                    {safeSubjects.map((subjectName) => (
                      <SelectItem key={subjectName} value={subjectName}>
                        {subjectName}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="rounded-xl border border-[#eadfd4] bg-white">
                <div className="flex items-center justify-between border-b border-[#eadfd4] px-4 py-3">
                  <div>
                    <div className="font-semibold text-[#33241f]">
                      Subject Export Preview
                    </div>
                    <div className="text-xs text-[#8a756b]">
                      Only the selected subject will be exported.
                    </div>
                  </div>
                  <span className="rounded-full bg-[#f5ece4] px-3 py-1 text-xs font-medium text-[#6f554a]">
                    {perSubjectPick ? "1 subject selected" : "No subject selected"}
                  </span>
                </div>

                <div className="overflow-x-auto p-3">
                  <table className="w-full min-w-[560px] text-xs">
                    <thead>
                      <tr className="border-b text-left text-[#6f554a]">
                        <th className="px-2 py-2">Subject</th>
                        <th className="px-2 py-2">Terms Included</th>
                        <th className="px-2 py-2">Export File</th>
                        <th className="px-2 py-2">Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="px-2 py-3 font-medium text-[#33241f]">
                          {perSubjectPick || "—"}
                        </td>
                        <td className="px-2 py-3 text-[#6f554a]">
                          {isGrade12SummaryClass &&
                          grade12SummarySubjectTerm(perSubjectPick)
                            ? `Term ${grade12SummarySubjectTerm(
                                perSubjectPick,
                              )}, Final`
                            : "Term 1, Term 2, Term 3, Final"}
                        </td>
                        <td className="px-2 py-3 text-[#6f554a]">
                          {perSubjectPick
                            ? `Summary-of-Grades-per-Subject_${safeExcelFilePart(
                                perSubjectPick,
                              )}_${safeExcelFilePart(
                                klass
                                  ? formatGradeLevel(klass.grade_level)
                                  : "Grade",
                              )}_${safeExcelFilePart(
                                klass?.section || "Class",
                              )}.xlsm`
                            : "—"}
                        </td>
                        <td className="px-2 py-3">
                          <span className="rounded-full bg-emerald-100 px-2 py-1 text-[11px] font-medium text-emerald-700">
                            {perSubjectPick
                              ? "Ready to Export"
                              : "Select Subject"}
                          </span>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="flex gap-3 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-xs text-amber-800">
                <Lightbulb className="mt-0.5 size-4 shrink-0" />
                <p>
                  Only the selected subject will be exported to Excel. Choose
                  another subject from the dropdown to generate a different
                  subject file.
                </p>
              </div>

              <div className="flex justify-end gap-2 pt-1">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setPerSubjectOpen(false)}
                  disabled={isExportingPerSubject}
                >
                  Cancel
                </Button>
                <Button
                  type="button"
                  onClick={exportSummaryPerSubject}
                  disabled={!perSubjectPick || isExportingPerSubject}
                  className="bg-emerald-600 text-white hover:bg-emerald-700"
                >
                  {isExportingPerSubject ? (
                    <LoaderCircle className="mr-2 size-4 animate-spin" />
                  ) : (
                    <FileSpreadsheet className="mr-2 size-4" />
                  )}
                  {isExportingPerSubject
                    ? "Exporting..."
                    : "Export Subject Excel"}
                </Button>
              </div>
            </div>
          </div>
        </>
      )}

      {addOpen && (
        <>
          <button
            type="button"
            aria-label="Close Add Subject panel"
            className="fixed inset-0 z-40 cursor-default bg-black/10 backdrop-blur-[1px]"
            onClick={closeAddSubject}
          />

          <aside
            role="dialog"
            aria-modal="true"
            aria-labelledby="add-subject-title"
            className="fixed inset-y-0 right-0 z-50 flex w-full max-w-[440px] flex-col overflow-y-auto border-l border-[#eadfd4] bg-[#fffdf9] shadow-2xl"
          >
            <div className="flex items-center justify-between border-b border-[#eadfd4] px-6 py-5">
              <h2
                id="add-subject-title"
                className="text-xl font-semibold text-[#33241f]"
              >
                Add Subject
              </h2>
              <button
                type="button"
                onClick={closeAddSubject}
                className="rounded-lg p-2 text-[#6f554a] transition hover:bg-[#f5ece4] hover:text-[#33241f]"
                aria-label="Close Add Subject panel"
              >
                <X className="size-5" />
              </button>
            </div>

            <div className="flex flex-1 flex-col gap-6 px-6 py-5">
              <div className="flex gap-3 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">
                <Lightbulb className="mt-0.5 size-4 shrink-0" />
                <p>
                  Choose Request Grades or Send Grades, then select one
                  completed or finalized form. Only that selected form's
                  subject and grades will be added to this summary.
                </p>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  Grade Source
                </label>
                <Select
                  value={addGradeSource}
                  onValueChange={(value) => {
                    setAddGradeSource(value as AddSubjectGradeSource);
                    setAddPick("");
                    setAddGradeLevel("");
                    setAddRequestId("");
                  }}
                >
                  <SelectTrigger className="h-11 border-[#dfcfc1] bg-white">
                    <SelectValue placeholder="Select Request Grades or Send Grades" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="request_grades">
                      Request Grades
                    </SelectItem>
                    <SelectItem value="send_grades">Send Grades</SelectItem>
                  </SelectContent>
                </Select>
                <p className="text-xs text-[#8a756b]">
                  {addGradeSource === "request_grades"
                    ? "Shows finalized grades returned through Request Form."
                    : addGradeSource === "send_grades"
                      ? "Shows completed grades received through Send Grades."
                      : "Select which grade workflow you want to use."}
                </p>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  Select Subject
                </label>
                <Select
                  value={addPick}
                  disabled={!addGradeSource}
                  onValueChange={(value) => {
                    setAddPick(value);
                    setAddGradeLevel("");
                    setAddRequestId("");
                  }}
                >
                  <SelectTrigger className="h-11 border-[#dfcfc1] bg-white">
                    <SelectValue placeholder="Select a subject" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableReceivedSubjects.map((subjectName) => (
                      <SelectItem key={subjectName} value={subjectName}>
                        {subjectName}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  Grade Level
                </label>
                <Select
                  value={addGradeLevel}
                  onValueChange={(value) => {
                    setAddGradeLevel(value);
                    setAddRequestId("");
                  }}
                  disabled={!addPick}
                >
                  <SelectTrigger className="h-11 border-[#dfcfc1] bg-white">
                    <SelectValue placeholder="Select a grade level" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableGradeLevels.map((gradeLevel) => (
                      <SelectItem key={gradeLevel} value={gradeLevel}>
                        {gradeLevel}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  Select Completed / Finalized Form
                </label>
                <Select
                  value={addRequestId}
                  onValueChange={setAddRequestId}
                  disabled={!addPick || !addGradeLevel}
                >
                  <SelectTrigger className="h-11 border-[#dfcfc1] bg-white">
                    <SelectValue placeholder="Choose one completed or finalized form" />
                  </SelectTrigger>
                  <SelectContent>
                    {matchingReceivedForms.map((request) => (
                      <SelectItem key={request.id} value={request.id}>
                        {request.term === "final"
                          ? "Final Grade"
                          : `Term ${request.term}`}{" "}
                        ·{" "}
                        {request.section
                          ? `Section ${request.section}`
                          : "No section"}{" "}
                        · {new Date(request.created_at).toLocaleDateString()}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {addPick &&
                  addGradeLevel &&
                  matchingReceivedForms.length === 0 && (
                    <p className="text-xs text-[#9b1c28]">
                      No {addGradeSource === "request_grades"
                        ? "finalized Request Grades form"
                        : "completed Send Grades form"} matches this subject
                      and grade level.
                    </p>
                  )}
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-[#33241f]">
                  School Year
                </label>
                <div className="flex h-11 items-center rounded-md border border-[#dfcfc1] bg-white px-3 text-sm text-[#33241f]">
                  {selectedReceivedForm?.schoolYear || selectedSchoolYear}
                </div>
              </div>

              <div className="space-y-3">
                <div className="text-sm font-semibold text-[#33241f]">
                  Student Names
                </div>

                <div className="flex gap-3 rounded-xl border border-emerald-300 bg-emerald-50/40 p-4">
                  <CheckCircle2 className="mt-0.5 size-4 shrink-0 text-emerald-600" />
                  <UsersRound className="mt-0.5 size-5 shrink-0 text-emerald-600" />
                  <span>
                    <span className="block text-sm font-semibold text-[#33241f]">
                      Use the existing {selectedGradeLevel} class list
                    </span>
                    <span className="mt-1 block text-xs text-[#8a756b]">
                      Every added subject uses the same learner names. Grades
                      are matched by name from the selected form.
                    </span>
                  </span>
                </div>
              </div>

              {selectedReceivedForm && (
                <div className="rounded-xl border border-[#dfcfc1] bg-white p-4">
                  <div className="flex items-start gap-3">
                    <Inbox className="mt-0.5 size-5 shrink-0 text-[#941c2f]" />
                    <div className="min-w-0 text-sm">
                      <p className="font-semibold text-[#33241f]">
                        {selectedReceivedForm.subject} ·{" "}
                        {selectedReceivedForm.term === "final"
                          ? "Final Grade"
                          : `Term ${selectedReceivedForm.term}`}
                      </p>
                      <p className="mt-1 text-xs font-medium text-[#941c2f]">
                        {selectedReceivedForm.source === "finalized_request"
                          ? "Finalized Request Form"
                          : "Completed Receive Form"}
                      </p>
                      <p className="mt-1 text-xs text-[#8a756b]">
                        {selectedReceivedForm.gradeLevel}
                        {selectedReceivedForm.section
                          ? ` · Section ${selectedReceivedForm.section}`
                          : ""}
                        {selectedReceivedForm.schoolYear
                          ? ` · ${selectedReceivedForm.schoolYear}`
                          : ""}
                      </p>
                      <p className="mt-2 text-xs font-medium text-emerald-700">
                        {selectedReceivedForm.learners.length} learner record(s)
                        ready to match
                      </p>
                    </div>
                  </div>
                </div>
              )}

              <Button
                className="mt-auto h-12 w-full bg-[#941c2f] text-white hover:bg-[#7f1728]"
                disabled={!selectedReceivedForm || isImportingSubject}
                onClick={addSubject}
              >
                {isImportingSubject ? (
                  <LoaderCircle className="mr-2 size-4 animate-spin" />
                ) : (
                  <Plus className="mr-2 size-4" />
                )}
                {isImportingSubject ? "Adding Subject..." : "Add Selected Form"}
              </Button>
            </div>
          </aside>
        </>
      )}
    </div>
  );
}

function RequestPanel({
  classInfo,
  students,
  requests,
  recipientProfiles,
  resolveGrade,
  onChange,
}: {
  classInfo?: ClassWithTeacherName;
  students: StudentRow[];
  requests: GradeRequest[];
  recipientProfiles: RecipientProfile[];
  resolveGrade: (
    studentId: string,
    subjectName: string,
    termValue: SummaryGradeTerm,
  ) => number | null;
  onChange: () => void;
}) {
  const safeRequests = Array.isArray(requests) ? requests : [];
  const visibleRequests = safeRequests.filter(
    (request) => request.status !== "completed",
  );
  const safeRecipients = Array.isArray(recipientProfiles)
    ? recipientProfiles
    : [];
  const sent = visibleRequests.length;

  const recipientRoleLabel = "Class Adviser" as const;
  const requestSubject = classInfo?.subject || REQUEST_SUBJECTS[0];

  return (
    <div
      id="send-grade-request"
      className="min-w-0 scroll-mt-32 overflow-hidden rounded-2xl border-2 border-primary/30 bg-card"
    >
      <div className="flex flex-col gap-3 border-b bg-primary/5 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex min-w-0 flex-wrap items-center gap-2 text-sm font-semibold text-primary">
          <Mail className="size-4" />
          Grade Submission to Class Adviser
          <span className="chip bg-primary/10 text-primary">
            {sent}/{SEND_REQUEST_ROWS.length}
          </span>
        </div>

        <Button variant="outline" size="sm" onClick={onChange} className="w-full sm:w-auto">
          <RefreshCw className="mr-1 size-3" />
          Refresh Status
        </Button>
      </div>

      <div className="divide-y">
        {SEND_REQUEST_ROWS.map((rowNumber) => (
          <SubjectRequestRow
            key={rowNumber}
            subject={requestSubject}
            rowNumber={rowNumber}
            classInfo={classInfo}
            students={students}
            requests={
              rowNumber === 1
                ? visibleRequests.filter(
                    (request) => request.subject === requestSubject,
                  )
                : []
            }
            recipientRoleLabel={recipientRoleLabel}
            recipientProfiles={safeRecipients}
            resolveGrade={resolveGrade}
            onChange={onChange}
          />
        ))}
      </div>
    </div>
  );
}

function SubjectRequestRow({
  subject,
  rowNumber,
  classInfo,
  students,
  requests,
  recipientRoleLabel,
  recipientProfiles,
  resolveGrade,
  onChange,
}: {
  subject: string;
  rowNumber: number;
  classInfo?: ClassWithTeacherName;
  students: StudentRow[];
  requests: GradeRequest[];
  recipientRoleLabel: "Class Adviser" | "Subject Teacher";
  recipientProfiles: RecipientProfile[];
  resolveGrade: (
    studentId: string,
    subjectName: string,
    termValue: SummaryGradeTerm,
  ) => number | null;
  onChange: () => void;
}) {
  const [requestTerm, setRequestTerm] = useState("1");
  const [email, setEmail] = useState("");
  const [sending, setSending] = useState(false);

  const safeStudents = Array.isArray(students) ? students : [];

  const recipientById = new Map(
    recipientProfiles.map((profile) => [profile.id, profile]),
  );

  const termLabel =
    requestTerm === "final" ? "Final Grade" : `Term ${requestTerm}`;

  const gradeForStudent = (studentId: string) =>
    resolveGrade(studentId, subject, requestTerm as SummaryGradeTerm);

  const send = async () => {
    const cleanEmail = email.trim().toLowerCase();

    if (!/^\S+@\S+\.\S+$/.test(cleanEmail)) {
      toast.error(`Enter a valid ${recipientRoleLabel} email.`);
      return;
    }

    const recipient = recipientProfiles.find(
      (profile) => profile.email?.trim().toLowerCase() === cleanEmail,
    );

    if (!recipient) {
      toast.error(
        `No other ${recipientRoleLabel} account matches this email.`,
      );
      return;
    }

    if (safeStudents.length === 0) {
      toast.error("There are no students in the selected class.");
      return;
    }

    setSending(true);

    try {
      const requesterId = await getUserId();
      const classLabel = [
        classInfo?.grade_level ? formatGradeLevel(classInfo.grade_level) : null,
        classInfo?.subject ? `Class Subject: ${classInfo.subject}` : null,
        classInfo?.section ? `Section: ${classInfo.section}` : null,
        classInfo?.school_year ? `School Year: ${classInfo.school_year}` : null,
      ]
        .filter(Boolean)
        .join(" · ");

      const learnerLines = safeStudents.map((student, index) => {
        const score = gradeForStudent(student.id);
        const learnerName = `${student.last_name}, ${student.first_name}`;

        return `${index + 1}. ${learnerName} — ${termLabel}: ${score ?? "Not yet submitted"}`;
      });

      const requestMessage = [
        `Grade Submission ${rowNumber} for ${subject} ${termLabel}.`,
        classLabel,
        `Students (${safeStudents.length}):`,
        ...learnerLines,
      ]
        .filter(Boolean)
        .join("\n");

      const { error } = await (supabase as any).from("grade_requests").insert({
        requester_id: requesterId,
        // This existing column stores the assigned receiving Class Adviser.
        subject_teacher_id: recipient.id,
        subject,
        term: requestTerm,
        message: requestMessage,
        priority: "low",
        status: "pending",
      });

      if (error) throw error;

      toast.success("Grades sent to the Class Adviser for review.");
      setEmail("");
      onChange();
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Request failed.");
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="space-y-2 p-3">
      {requests.map((request) => {
        const recipient = request.subject_teacher_id
          ? recipientById.get(request.subject_teacher_id)
          : undefined;

        return (
          <div
            key={request.id}
            className="flex min-w-0 flex-wrap items-center gap-2 rounded-lg bg-muted/50 px-3 py-2 text-xs"
          >
            <span className="chip bg-primary/10 text-primary">
              {request.term === "final"
                ? "Final Grade"
                : `Term ${request.term}`}
            </span>
            <span className="min-w-0 flex-1 break-all text-muted-foreground">
              {recipient?.email || recipient?.full_name || recipientRoleLabel}
            </span>
            <span className="chip inline-flex items-center gap-1 border border-amber-200 bg-amber-100 text-amber-700 capitalize">
              <Hourglass className="size-3" />
              {request.status === "pending"
                ? "Pending Review"
                : request.status === "completed"
                  ? "Accepted"
                  : request.status.replace("_", " ")}
            </span>
          </div>
        );
      })}

      <div className="grid min-w-0 grid-cols-1 gap-2 md:grid-cols-[7rem_minmax(0,1fr)_9.5rem]">
        <Select value={requestTerm} onValueChange={setRequestTerm}>
          <SelectTrigger className="h-10 w-full">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="1">Term 1</SelectItem>
            <SelectItem value="2">Term 2</SelectItem>
            <SelectItem value="3">Term 3</SelectItem>
            <SelectItem value="final">Final Grade</SelectItem>
          </SelectContent>
        </Select>

        <Input
          value={email}
          onChange={(event) => setEmail(event.target.value)}
          placeholder={`Another ${recipientRoleLabel} Email (@gmail.com or @deped.gov.ph)`}
          aria-label={`${subject} ${recipientRoleLabel} Email row ${rowNumber}`}
          className="h-10 min-w-0 w-full"
        />

        <Button onClick={send} size="sm" disabled={sending} className="h-10 w-full whitespace-nowrap">
          {sending ? (
            <RefreshCw className="mr-1 size-4 animate-spin" />
          ) : (
            <Send className="mr-1 size-4" />
          )}
          {sending ? "Submitting..." : "Submit Grades"}
        </Button>
      </div>
    </div>
  );
}