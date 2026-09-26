import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import { Skeleton } from "@/components/ui/skeleton";
import {
  ArrowLeft,
  ArrowRight,
  AlertTriangle,
  Atom,
  BookMarked,
  BookOpen,
  BriefcaseBusiness,
  Calculator,
  CalendarDays,
  Dumbbell,
  FileSearch,
  FlaskConical,
  Globe2,
  HeartHandshake,
  Languages,
  Landmark,
  LogIn,
  LogOut,
  Music2,
  Palette,
  School,
  Printer,
  Download,
  ClipboardList,
  HeartPulse,
  Loader2,
  Minus,
  Pencil,
  Plus,
  Save,
  Send,
  Target,
  Trash2,
  Users,
  UserPlus2,
  X,
  type LucideIcon,
} from "lucide-react";
import { Fragment, useMemo, useState, useEffect, useRef, type ReactNode, type InputHTMLAttributes } from "react";
import { toast } from "sonner";
import {
  getUserId,
  descriptorFor,
  type ClassRow,
  type StudentRow,
  type GradeComponent,
  type GradeActivity,
  type ActivityScore,
} from "@/lib/data";
import JSZip from "jszip";

type SchoolYearSetting = {
  id: number;
  school_year: string | null;
  is_locked: boolean;
  updated_by: string | null;
  updated_at: string;
};

async function getSchoolYearSetting(): Promise<SchoolYearSetting> {
  const { data, error } = await (supabase as any)
    .from("school_year_settings")
    .select("id, school_year, is_locked, updated_by, updated_at")
    .eq("id", 1)
    .maybeSingle();

  if (error) throw error;

  return {
    id: Number(data?.id ?? 1),
    school_year: data?.school_year ?? null,
    is_locked: Boolean(data?.is_locked),
    updated_by: data?.updated_by ?? null,
    updated_at: data?.updated_at ?? new Date(0).toISOString(),
  };
}

type SchoolYearLibraryRow = {
  id: number;
  school_year: string;
  is_locked: boolean;
};

async function getSchoolYearLibrary(): Promise<SchoolYearLibraryRow[]> {
  const { data, error } = await (supabase as any)
    .from("school_year_library")
    .select("id, school_year, is_locked")
    .order("school_year", { ascending: false });

  if (error) throw error;

  return (data ?? []).map((row: any) => ({
    id: Number(row.id),
    school_year: String(row.school_year ?? ""),
    is_locked: Boolean(row.is_locked),
  }));
}

export const Route = createFileRoute("/_authenticated/classes/$classId")({
  component: ClassDetail,
  errorComponent: ({ error }) => (
    <div className="p-6 text-sm text-destructive">{error.message}</div>
  ),
  notFoundComponent: () => <div className="p-6 text-sm">Class not found.</div>,
});

type Tab = "grades" | "fitness";

type ClassWithTrack = ClassRow & {
  track_shs?: string | null;
};

type SubjectTeacherOption = {
  id: string;
  full_name: string | null;
  email: string | null;
};

type StudentListAssignment = {
  id: string;
  subject_teacher_id: string;
  subject: string;
  status: "pending" | "accepted" | "declined";
  decline_reason: string | null;
  sent_at: string;
  responded_at: string | null;
};

const JUNIOR_HIGH_SUBJECTS = [
  "Filipino",
  "Science",
  "Math",
  "MAPEH",
  "English",
  "Aral Pan",
  "Research",
  "Creative Teach",
  "Val Ed",
];

const SUBJECTS_BY_GRADE: Record<string, readonly string[]> = {
  "Grade 7": JUNIOR_HIGH_SUBJECTS,
  "Grade 8": JUNIOR_HIGH_SUBJECTS,
  "Grade 9": JUNIOR_HIGH_SUBJECTS,
  "Grade 10": JUNIOR_HIGH_SUBJECTS,
  "Grade 11": [
    "General Mathematics",
    "Mabisang Komunikasyon",
    "General Science",
    "Life & Career Skills",
    "Pag-aaral sa Kasaysayan at Lipunang Pilipino",
    "Finite Mathematics",
    "Chemistry 1",
    "Practical Research",
    "Elective Subject",
  ],
  "Grade 12": [
    "General Biology 1",
    "General Physics 1",
    "Media & Information Literacy",
    "Practical Research 2",
    "P.E. & Health",
    "Filipino sa Piling Larang",
  ],
};

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

function normalizeLearnerPart(value: unknown) {
  return normalize(value)
    .replace(/[.'’]/g, "")
    .replace(/\s+/g, " ");
}

function learnerNameKey(value: {
  last_name?: string | null;
  first_name?: string | null;
  middle_name?: string | null;
}) {
  return [
    normalizeLearnerPart(value.last_name),
    normalizeLearnerPart(value.first_name),
    normalizeLearnerPart(value.middle_name),
  ].join("|");
}

function sameLearner(
  existing: {
    last_name?: string | null;
    first_name?: string | null;
    middle_name?: string | null;
    lrn?: string | null;
  },
  candidate: {
    last_name?: string | null;
    first_name?: string | null;
    middle_name?: string | null;
    lrn?: string | null;
  },
) {
  const existingLrn = String(existing.lrn ?? "").replace(/\D/g, "").trim();
  const candidateLrn = String(candidate.lrn ?? "").replace(/\D/g, "").trim();

  // If both records have an LRN, the LRN is the strongest identity check.
  if (existingLrn && candidateLrn) {
    return existingLrn === candidateLrn;
  }

  // If one/both LRNs are missing, prevent duplicate full names in the same class.
  return learnerNameKey(existing) === learnerNameKey(candidate);
}

function deduplicateStudentRows(rows: StudentRow[]) {
  const unique: StudentRow[] = [];

  for (const student of rows) {
    if (unique.some((existing) => sameLearner(existing, student))) {
      continue;
    }
    unique.push(student);
  }

  return unique;
}

function isCommunicationCompositeSubject(subject: unknown) {
  const normalizedSubject = normalize(subject).replace(/\s+/g, " ");
  return (
    normalizedSubject === "mabisang komunikasyon" ||
    normalizedSubject === "mabisang komuniksyon"
  );
}

function isGrade11ElectiveSubject(
  gradeLevel: unknown,
  subject: unknown,
) {
  return (
    normalize(gradeLevel).replace(/\s+/g, " ") === "grade 11" &&
    normalize(subject).replace(/\s+/g, " ") === "elective subject"
  );
}

type SubjectVisualKind =
  | "mathematics"
  | "filipino"
  | "mapeh"
  | "science"
  | "social-studies"
  | "research"
  | "language"
  | "career"
  | "values"
  | "general";

type SubjectTheme = {
  kind: SubjectVisualKind;
  category: string;
  Icon: LucideIcon;
  cardClass: string;
  iconClass: string;
  iconBackgroundClass: string;
  badgeClass: string;
  artworkClass: string;
};

function getSubjectTheme(subject: string): SubjectTheme {
  const normalizedSubject = subject.toLowerCase().trim();

  if (
    normalizedSubject.includes("math") ||
    normalizedSubject.includes("calculus") ||
    normalizedSubject.includes("algebra")
  ) {
    return {
      kind: "mathematics",
      category: "Mathematics",
      Icon: Calculator,
      cardClass:
        "border-sky-200/90 bg-gradient-to-br from-sky-50 via-cyan-50 to-blue-100/90",
      iconClass: "text-sky-700",
      iconBackgroundClass: "bg-white/80 ring-sky-200",
      badgeClass: "border-sky-200 bg-sky-100/80 text-sky-800",
      artworkClass: "text-sky-700/15",
    };
  }

  if (
    normalizedSubject.includes("filipino") ||
    normalizedSubject.includes("mabisang komunikasyon")
  ) {
    return {
      kind: "filipino",
      category: "Wika at Panitikan",
      Icon: Languages,
      cardClass:
        "border-amber-200/90 bg-gradient-to-br from-amber-50 via-yellow-50 to-rose-100/80",
      iconClass: "text-rose-700",
      iconBackgroundClass: "bg-white/80 ring-rose-200",
      badgeClass: "border-rose-200 bg-rose-100/80 text-rose-800",
      artworkClass: "text-rose-700/15",
    };
  }

  if (
    normalizedSubject.includes("mapeh") ||
    normalizedSubject.includes("p.e") ||
    normalizedSubject.includes("physical education") ||
    normalizedSubject.includes("health") ||
    normalizedSubject.includes("music") ||
    normalizedSubject.includes("arts")
  ) {
    return {
      kind: "mapeh",
      category: "MAPEH",
      Icon: Music2,
      cardClass:
        "border-fuchsia-200/90 bg-gradient-to-br from-fuchsia-50 via-pink-50 to-orange-100/80",
      iconClass: "text-fuchsia-700",
      iconBackgroundClass: "bg-white/80 ring-fuchsia-200",
      badgeClass: "border-fuchsia-200 bg-fuchsia-100/80 text-fuchsia-800",
      artworkClass: "text-fuchsia-700/15",
    };
  }

  if (
    normalizedSubject.includes("science") ||
    normalizedSubject.includes("biology") ||
    normalizedSubject.includes("physics") ||
    normalizedSubject.includes("chemistry")
  ) {
    return {
      kind: "science",
      category: "Science",
      Icon: Atom,
      cardClass:
        "border-emerald-200/90 bg-gradient-to-br from-emerald-50 via-teal-50 to-cyan-100/80",
      iconClass: "text-emerald-700",
      iconBackgroundClass: "bg-white/80 ring-emerald-200",
      badgeClass: "border-emerald-200 bg-emerald-100/80 text-emerald-800",
      artworkClass: "text-emerald-700/15",
    };
  }

  if (
    normalizedSubject.includes("aral pan") ||
    normalizedSubject.includes("kasaysayan") ||
    normalizedSubject.includes("lipunan") ||
    normalizedSubject.includes("history")
  ) {
    return {
      kind: "social-studies",
      category: "Araling Panlipunan",
      Icon: Globe2,
      cardClass:
        "border-orange-200/90 bg-gradient-to-br from-orange-50 via-amber-50 to-yellow-100/80",
      iconClass: "text-orange-700",
      iconBackgroundClass: "bg-white/80 ring-orange-200",
      badgeClass: "border-orange-200 bg-orange-100/80 text-orange-800",
      artworkClass: "text-orange-700/15",
    };
  }

  if (normalizedSubject.includes("research")) {
    return {
      kind: "research",
      category: "Research",
      Icon: FileSearch,
      cardClass:
        "border-indigo-200/90 bg-gradient-to-br from-indigo-50 via-violet-50 to-purple-100/80",
      iconClass: "text-indigo-700",
      iconBackgroundClass: "bg-white/80 ring-indigo-200",
      badgeClass: "border-indigo-200 bg-indigo-100/80 text-indigo-800",
      artworkClass: "text-indigo-700/15",
    };
  }

  if (
    normalizedSubject.includes("english") ||
    normalizedSubject.includes("communication") ||
    normalizedSubject.includes("media")
  ) {
    return {
      kind: "language",
      category: "Language & Media",
      Icon: BookMarked,
      cardClass:
        "border-violet-200/90 bg-gradient-to-br from-violet-50 via-purple-50 to-fuchsia-100/70",
      iconClass: "text-violet-700",
      iconBackgroundClass: "bg-white/80 ring-violet-200",
      badgeClass: "border-violet-200 bg-violet-100/80 text-violet-800",
      artworkClass: "text-violet-700/15",
    };
  }

  if (
    normalizedSubject.includes("career") ||
    normalizedSubject.includes("entrepreneur")
  ) {
    return {
      kind: "career",
      category: "Life & Career",
      Icon: BriefcaseBusiness,
      cardClass:
        "border-cyan-200/90 bg-gradient-to-br from-cyan-50 via-sky-50 to-indigo-100/70",
      iconClass: "text-cyan-700",
      iconBackgroundClass: "bg-white/80 ring-cyan-200",
      badgeClass: "border-cyan-200 bg-cyan-100/80 text-cyan-800",
      artworkClass: "text-cyan-700/15",
    };
  }

  if (
    normalizedSubject.includes("val ed") ||
    normalizedSubject.includes("values")
  ) {
    return {
      kind: "values",
      category: "Values Education",
      Icon: HeartHandshake,
      cardClass:
        "border-lime-200/90 bg-gradient-to-br from-lime-50 via-green-50 to-emerald-100/70",
      iconClass: "text-green-700",
      iconBackgroundClass: "bg-white/80 ring-green-200",
      badgeClass: "border-green-200 bg-green-100/80 text-green-800",
      artworkClass: "text-green-700/15",
    };
  }

  return {
    kind: "general",
    category: "General Subject",
    Icon: BookOpen,
    cardClass:
      "border-stone-200/90 bg-gradient-to-br from-stone-50 via-orange-50/60 to-amber-100/70",
    iconClass: "text-primary",
    iconBackgroundClass: "bg-white/80 ring-primary/15",
    badgeClass: "border-primary/15 bg-primary/10 text-primary",
    artworkClass: "text-primary/10",
  };
}

function SubjectHeaderArtwork({ theme }: { theme: SubjectTheme }) {
  const commonClass = `pointer-events-none absolute inset-y-0 right-0 w-[55%] overflow-hidden ${theme.artworkClass}`;

  if (theme.kind === "mathematics") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute -right-3 -top-2 rotate-[-8deg] text-right font-serif text-base font-bold leading-7">
          <div>∑ x² + y²</div>
          <div>f(x) = mx + b</div>
          <div>πr² · √144</div>
        </div>
        <Calculator className="absolute -bottom-8 right-24 size-24 rotate-12 opacity-40" />
      </div>
    );
  }

  if (theme.kind === "filipino") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute -right-8 -top-10 size-28 rounded-full border-[14px] border-amber-500/20" />
        <div className="absolute right-4 top-3 rotate-[-7deg] text-right font-serif font-bold tracking-widest">
          <div className="text-2xl">WIKA</div>
          <div className="mt-1 text-[9px] tracking-[0.25em]">PANITIKAN · KULTURA</div>
        </div>
        <Languages className="absolute -bottom-8 right-20 size-24 -rotate-6 opacity-45" />
      </div>
    );
  }

  if (theme.kind === "mapeh") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Music2 className="absolute right-8 top-2 size-12 -rotate-12" />
        <Palette className="absolute right-24 top-9 size-10 rotate-12" />
        <Dumbbell className="absolute bottom-1 right-3 size-12 rotate-6" />
        <HeartPulse className="absolute bottom-1 right-28 size-10 -rotate-6" />
      </div>
    );
  }

  if (theme.kind === "science") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Atom className="absolute -right-5 -top-6 size-28 rotate-12 opacity-65" />
        <FlaskConical className="absolute bottom-0 right-24 size-14 -rotate-6" />
        <div className="absolute bottom-2 right-2 font-mono text-xs font-bold">
          H₂O · F=ma · DNA
        </div>
      </div>
    );
  }

  if (theme.kind === "social-studies") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Globe2 className="absolute -right-4 -top-5 size-28 opacity-60" />
        <Landmark className="absolute bottom-0 right-24 size-16" />
      </div>
    );
  }

  if (theme.kind === "research") {
    return (
      <div className={commonClass} aria-hidden="true">
        <FileSearch className="absolute -right-3 -top-3 size-28 rotate-6 opacity-60" />
        <div className="absolute bottom-2 right-5 font-mono text-[9px] font-bold tracking-widest">
          QUESTION · DATA · EVIDENCE
        </div>
      </div>
    );
  }

  if (theme.kind === "language") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute right-5 -top-5 font-serif text-7xl font-bold leading-none opacity-60">
          “
        </div>
        <BookMarked className="absolute -bottom-3 right-20 size-20 -rotate-6" />
      </div>
    );
  }

  if (theme.kind === "career") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Target className="absolute -right-4 -top-5 size-28 opacity-60" />
        <BriefcaseBusiness className="absolute bottom-0 right-24 size-16 -rotate-6" />
      </div>
    );
  }

  if (theme.kind === "values") {
    return (
      <div className={commonClass} aria-hidden="true">
        <HeartHandshake className="absolute -right-2 -top-5 size-28 rotate-6 opacity-65" />
      </div>
    );
  }

  return (
    <div className={commonClass} aria-hidden="true">
      <BookOpen className="absolute -bottom-8 -right-3 size-28 -rotate-6 opacity-60" />
    </div>
  );
}

/**
 * Keep the fixed DepEd Excel template layout intact, but show only enrolled
 * learners in JHS exported worksheets. Physically deleting rows would break
 * the template's formulas, section headers and achievement-count tables.
 * Do not hide INPUT DATA rows: its school information shares those rows.
 */
function hideUnusedJhsExcelRows(
  worksheetXml: Document,
  layout: "term" | "summary",
  maleCount: number,
  femaleCount: number,
) {
  const maleFirst = layout === "term" ? 12 : 13;
  const femaleFirst = layout === "term" ? 63 : 64;
  const maleHeader = maleFirst - 1;
  const femaleHeader = femaleFirst - 1;
  const rowElements = Array.from(worksheetXml.getElementsByTagName("row"));

  rowElements.forEach((row) => {
    const rowNumber = Number(row.getAttribute("r"));
    const isUnusedMale =
      rowNumber >= maleFirst + maleCount && rowNumber < maleFirst + 50;
    const isUnusedFemale =
      rowNumber >= femaleFirst + femaleCount && rowNumber < femaleFirst + 50;
    const isEmptyGroupHeader =
      (rowNumber === maleHeader && maleCount === 0) ||
      (rowNumber === femaleHeader && femaleCount === 0);

    if (isUnusedMale || isUnusedFemale || isEmptyGroupHeader) {
      row.setAttribute("hidden", "1");
    } else if (
      (rowNumber >= maleFirst && rowNumber < maleFirst + maleCount) ||
      (rowNumber >= femaleFirst && rowNumber < femaleFirst + femaleCount) ||
      (rowNumber === maleHeader && maleCount > 0) ||
      (rowNumber === femaleHeader && femaleCount > 0)
    ) {
      row.removeAttribute("hidden");
    }
  });
}

const MAPEH_TEMPLATE_SHEETS = [
  { sheet: "MA_T1", storageTerm: "1" },
  { sheet: "PEH_T1", storageTerm: "PEH_T1" },
  { sheet: "MA_T2", storageTerm: "MA_T2" },
  { sheet: "PEH_T2", storageTerm: "PEH_T2" },
  { sheet: "MA_T3", storageTerm: "MA_T3" },
  { sheet: "PEH_T3", storageTerm: "PEH_T3" },
] as const;

async function exportMapehClassRecord({
  classId,
  students,
  klass,
  principal,
}: {
  classId: string;
  students: StudentRow[];
  klass: ClassRow;
  principal: string;
}) {
  const isJuniorHighMapehExport = /^(?:grade\s*)?(?:7|8|9|10)$/i.test(
    String(klass.grade_level ?? "").trim(),
  );
  const templateResponse = await fetch("/templates/Mapeh-class.xlsx");
  if (!templateResponse.ok) {
    throw new Error("MAPEH template not found. Put Mapeh-class.xlsx in public/templates.");
  }

  const [
    { data: components, error: componentsError },
    { data: activities, error: activitiesError },
    { data: savedMapehBases, error: savedMapehBasesError },
  ] = await Promise.all([
    supabase.from("grade_components").select("*").eq("class_id", classId),
    supabase.from("grade_activities").select("*").eq("class_id", classId).order("position"),
    (supabase as any)
      .from("grades")
      .select("student_id, subject, term_grade_base")
      .eq("class_id", classId)
      .eq("term", "1"),
  ]);

  if (componentsError) throw componentsError;
  if (activitiesError) throw activitiesError;
  if (savedMapehBasesError) throw savedMapehBasesError;

  const allComponents = (components ?? []) as GradeComponent[];
  const allActivities = (activities ?? []) as GradeActivity[];
  const allSavedMapehBases = (savedMapehBases ?? []) as Array<{
    student_id: string;
    subject: string;
    term_grade_base: number | null;
  }>;
  const activityIds = allActivities.map((activity) => activity.id);
  let allScores: ActivityScore[] = [];

  if (activityIds.length) {
    const { data: scores, error: scoresError } = await supabase
      .from("activity_scores")
      .select("*")
      .in("activity_id", activityIds);
    if (scoresError) throw scoresError;
    allScores = (scores ?? []) as ActivityScore[];
  }

  const zip = await JSZip.loadAsync(await templateResponse.arrayBuffer());
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const workbookFile = zip.file("xl/workbook.xml");
  const relationshipsFile = zip.file("xl/_rels/workbook.xml.rels");

  if (!workbookFile || !relationshipsFile) {
    throw new Error("The MAPEH Excel template has an invalid structure.");
  }

  const workbookXml = parser.parseFromString(await workbookFile.async("text"), "application/xml");
  const relationshipsXml = parser.parseFromString(
    await relationshipsFile.async("text"),
    "application/xml",
  );
  const relationshipTargets = new Map<string, string>();

  Array.from(relationshipsXml.getElementsByTagName("Relationship")).forEach((relationship) => {
    relationshipTargets.set(
      relationship.getAttribute("Id") ?? "",
      relationship.getAttribute("Target") ?? "",
    );
  });

  const sheetPaths = new Map<string, string>();
  Array.from(workbookXml.getElementsByTagName("sheet")).forEach((sheet) => {
    const name = sheet.getAttribute("name");
    const relationshipId =
      sheet.getAttribute("r:id") ??
      sheet.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      );
    const target = relationshipId ? relationshipTargets.get(relationshipId) : undefined;

    if (name && target) {
      sheetPaths.set(name, `xl/${target.replace(/^\/?xl\//, "")}`);
    }
  });

  const requiredSheets = [
    "INPUT DATA",
    ...MAPEH_TEMPLATE_SHEETS.map(({ sheet }) => sheet),
    "MAPEH_T1",
    "MAPEH_T2",
    "MAPEH_T3",
    "SUMMARY OF GRADES",
  ];
  if (requiredSheets.some((sheet) => !sheetPaths.has(sheet))) {
    throw new Error("The selected file is not the expected MAPEH template.");
  }

  const pendingValues = new Map<string, Map<string, string | number>>();
  const setCellValue = (
    sheetName: string,
    address: string,
    value: string | number | null | undefined,
  ) => {
    const sheetValues = pendingValues.get(sheetName) ?? new Map<string, string | number>();
    sheetValues.set(address, value ?? "");
    pendingValues.set(sheetName, sheetValues);
  };

  const fullName = (student: StudentRow) =>
    [student.last_name, student.first_name].filter(Boolean).join(", ");
  const maleStudents = students.filter((student) => normalize(student.sex) === "male").slice(0, 50);
  const femaleStudents = students
    .filter((student) => normalize(student.sex) === "female")
    .slice(0, 50);

  setCellValue("INPUT DATA", "E8", klass.region);
  setCellValue("INPUT DATA", "E9", klass.division);
  setCellValue("INPUT DATA", "E10", klass.school_name);
  setCellValue("INPUT DATA", "E11", klass.school_id);
  setCellValue("INPUT DATA", "E12", klass.school_year);
  setCellValue("INPUT DATA", "E14", [klass.grade_level, klass.section].filter(Boolean).join(" - "));
  setCellValue("INPUT DATA", "E15", "MAPEH");
  setCellValue("INPUT DATA", "E17", klass.teacher_name);
  setCellValue("INPUT DATA", "E18", "");
  setCellValue("INPUT DATA", "E20", principal);
  setCellValue("INPUT DATA", "E21", "");

  for (let index = 0; index < 50; index += 1) {
    setCellValue(
      "INPUT DATA",
      `J${index + 9}`,
      maleStudents[index] ? fullName(maleStudents[index]) : "",
    );
    setCellValue(
      "INPUT DATA",
      `P${index + 9}`,
      femaleStudents[index] ? fullName(femaleStudents[index]) : "",
    );
  }

  // Keep the downloaded MAPEH workbook synchronized with the current
  // E-Class Record UI. The original Excel template links many of these
  // cells to INPUT DATA using formulas, but Protected View can keep old
  // cached values until Excel recalculates. Writing the current values
  // directly makes the learner names and class information appear
  // immediately after Export to Excel.
  const gradeSection = [klass.grade_level, klass.section]
    .filter(Boolean)
    .join(" - ");

  // MA / PEH detailed class-record sheets.
  MAPEH_TEMPLATE_SHEETS.forEach(({ sheet }) => {
    setCellValue(sheet, "G4", klass.region);
    setCellValue(sheet, "L4", klass.division);
    setCellValue(sheet, "G5", klass.school_name);
    setCellValue(sheet, "S5", klass.school_id);
    setCellValue(sheet, "AB5", klass.school_year);
    setCellValue(sheet, "J7", gradeSection);
    setCellValue(sheet, "Q7", klass.teacher_name);

    for (let index = 0; index < 50; index += 1) {
      setCellValue(
        sheet,
        `B${index + 12}`,
        maleStudents[index] ? fullName(maleStudents[index]) : "",
      );
      setCellValue(
        sheet,
        `B${index + 63}`,
        femaleStudents[index] ? fullName(femaleStudents[index]) : "",
      );
    }
  });

  // MAPEH term summaries and final Summary of Grades.
  const mapehSummarySheets = [
    "MAPEH_T1",
    "MAPEH_T2",
    "MAPEH_T3",
    "SUMMARY OF GRADES",
  ] as const;

  mapehSummarySheets.forEach((sheetName) => {
    setCellValue(sheetName, "G5", klass.region);
    setCellValue(sheetName, "O5", klass.division);
    setCellValue(sheetName, "G6", klass.school_name);
    setCellValue(sheetName, "W6", klass.school_id);
    setCellValue(sheetName, "K8", gradeSection);
    setCellValue(sheetName, "W8", klass.school_year);
    setCellValue(sheetName, "K9", klass.teacher_name);
    setCellValue(sheetName, "W9", "MAPEH");

    for (let index = 0; index < 50; index += 1) {
      setCellValue(
        sheetName,
        `B${index + 13}`,
        maleStudents[index] ? fullName(maleStudents[index]) : "",
      );
      setCellValue(
        sheetName,
        `B${index + 64}`,
        femaleStudents[index] ? fullName(femaleStudents[index]) : "",
      );
    }
  });

  const scoreMap = new Map<string, number | null>();
  allScores.forEach((score) =>
    scoreMap.set(`${score.activity_id}|${score.student_id}`, score.score),
  );

  const componentConfig = {
    WW: {
      columns: ["F", "G", "H", "I", "J"],
      weightCell: "M10",
      headerCell: "F8",
      label: "WRITTEN / ORAL WORKS",
      defaultWeight: 20,
    },
    PT: {
      columns: ["N", "O", "P"],
      weightCell: "S10",
      headerCell: "N8",
      label: "PRODUCT / PERFORMANCE TASKS",
      defaultWeight: 50,
    },
    QA: {
      columns: ["T", "U", "V"],
      weightCell: "AA10",
      headerCell: "T8",
      label: "SUMMATIVE TESTS AND TERM EXAMINATIONS",
      defaultWeight: 30,
    },
  } as const;

  // Same MAPEH grading logic used by the E-Class Record UI.
  // MA/PEH component grades are calculated first, then each MAPEH term
  // is the rounded average of MA + PEH. The final MAPEH grade is the
  // rounded average of Term 1 + Term 2 + Term 3.
  const savedMapehBaseMap = new Map<string, number>();
  allSavedMapehBases.forEach((row) => {
    if (row.term_grade_base == null) return;
    const numericValue = Number(row.term_grade_base);
    if (Number.isFinite(numericValue)) {
      savedMapehBaseMap.set(`${row.student_id}|${row.subject}`, numericValue);
    }
  });

  const exportedComponentGrade = (
    studentId: string,
    scopeKey: string,
    storageTerm: string,
  ): number | null => {
    const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
      const componentActivities = allActivities
        .filter(
          (activity) =>
            activity.term === storageTerm &&
            activity.component === component,
        )
        .sort((a, b) => a.position - b.position)
        .slice(0, component === "QA" ? 3 : undefined);

      const componentWeight =
        allComponents.find(
          (item) =>
            item.term === storageTerm &&
            item.component === component,
        )?.weight ??
        (component === "WW" ? 20 : component === "PT" ? 50 : 30);

      let raw = 0;
      let hps = 0;
      let hasScore = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;

      componentActivities.forEach((activity, index) => {
        const value = scoreMap.get(`${activity.id}|${studentId}`);
        if (typeof value !== "number") return;

        // JHS exported Term Grades use the same score limits as the UI.
        // Leave all non-JHS export behavior unchanged.
        const effectiveScore = isJuniorHighMapehExport
          ? Math.max(0, Math.min(Math.max(0, Number(activity.hps) || 0), value))
          : value;
        raw += effectiveScore;
        hps += activity.hps;
        hasScore = true;

        if (component === "QA" && activity.hps > 0) {
          qaPercentageScore +=
            (effectiveScore / activity.hps) *
            100 *
            ((qaPartWeights[index] ?? 0) / 100);
        }
      });

      if (!hasScore || hps <= 0) return null;

      const percentageScore =
        component === "QA"
          ? qaPercentageScore
          : (raw / hps) * 100;

      if (isJuniorHighMapehExport) {
        // Exactly mirror the MAPEH TermGradesTable PS/WS rounding sequence.
        const roundedPercentage = Math.round(percentageScore * 100) / 100;
        return Math.round(roundedPercentage * (componentWeight / 100) * 100) / 100;
      }
      return percentageScore * (componentWeight / 100);
    });

    if (weightedParts.some((part) => part == null)) return null;

    const initialGrade =
      Math.round(
        (weightedParts as number[]).reduce((sum, part) => sum + part, 0) * 100,
      ) / 100;

    const computedBase = roundInitialGrade(initialGrade);
    if (computedBase == null) return null;

    const savedSubject = `${klass.subject || "MAPEH"}_${scopeKey}_ALL`;
    return resolveTermGradeBase(
      computedBase,
      savedMapehBaseMap.get(`${studentId}|${savedSubject}`),
    );
  };

  const averageMapehPair = (
    first: number | null,
    second: number | null,
  ): number | null => {
    if (first == null || second == null) return null;
    return Math.round((first + second) / 2);
  };

  const mapehSummaryForStudent = (studentId: string) => {
    const maT1 = exportedComponentGrade(studentId, "MA_T1", "1");
    const pehT1 = exportedComponentGrade(studentId, "PEH_T1", "PEH_T1");
    const maT2 = exportedComponentGrade(studentId, "MA_T2", "MA_T2");
    const pehT2 = exportedComponentGrade(studentId, "PEH_T2", "PEH_T2");
    const maT3 = exportedComponentGrade(studentId, "MA_T3", "MA_T3");
    const pehT3 = exportedComponentGrade(studentId, "PEH_T3", "PEH_T3");

    const term1 = averageMapehPair(maT1, pehT1);
    const term2 = averageMapehPair(maT2, pehT2);
    const term3 = averageMapehPair(maT3, pehT3);
    const terms = [term1, term2, term3];

    const final =
      terms.every((value): value is number => typeof value === "number")
        ? Math.round(
            (terms as number[]).reduce((sum, value) => sum + value, 0) / 3,
          )
        : null;

    return {
      maT1,
      pehT1,
      maT2,
      pehT2,
      maT3,
      pehT3,
      term1,
      term2,
      term3,
      final,
    };
  };

  const writeMapehSummaryStudent = (
    student: StudentRow,
    detailRow: number,
    summaryRow: number,
  ) => {
    const grades = mapehSummaryForStudent(student.id);

    const termData = [
      {
        sheet: "MAPEH_T1",
        ma: grades.maT1,
        peh: grades.pehT1,
        termGrade: grades.term1,
      },
      {
        sheet: "MAPEH_T2",
        ma: grades.maT2,
        peh: grades.pehT2,
        termGrade: grades.term2,
      },
      {
        sheet: "MAPEH_T3",
        ma: grades.maT3,
        peh: grades.pehT3,
        termGrade: grades.term3,
      },
    ] as const;

    termData.forEach(({ sheet, ma, peh, termGrade }) => {
      setCellValue(sheet, `F${summaryRow}`, ma ?? "");
      setCellValue(sheet, `J${summaryRow}`, peh ?? "");
      setCellValue(sheet, `N${summaryRow}`, termGrade ?? "");
      setCellValue(
        sheet,
        `V${summaryRow}`,
        termGrade == null ? "" : descriptorFor(termGrade).label,
      );
    });

    // Refresh the detailed MA/PEH Term Grade cells too, so the class-record
    // sheets and the summary sheets show the same grades as the web UI.
    setCellValue("MA_T1", `AC${detailRow}`, grades.maT1 ?? "");
    setCellValue("PEH_T1", `AC${detailRow}`, grades.pehT1 ?? "");
    setCellValue("MA_T2", `AC${detailRow}`, grades.maT2 ?? "");
    setCellValue("PEH_T2", `AC${detailRow}`, grades.pehT2 ?? "");
    setCellValue("MA_T3", `AC${detailRow}`, grades.maT3 ?? "");
    setCellValue("PEH_T3", `AC${detailRow}`, grades.pehT3 ?? "");

    setCellValue("SUMMARY OF GRADES", `F${summaryRow}`, grades.term1 ?? "");
    setCellValue("SUMMARY OF GRADES", `J${summaryRow}`, grades.term2 ?? "");
    setCellValue("SUMMARY OF GRADES", `N${summaryRow}`, grades.term3 ?? "");
    setCellValue("SUMMARY OF GRADES", `R${summaryRow}`, grades.final ?? "");
    setCellValue(
      "SUMMARY OF GRADES",
      `V${summaryRow}`,
      grades.final == null ? "" : descriptorFor(grades.final).label,
    );
    setCellValue(
      "SUMMARY OF GRADES",
      `Z${summaryRow}`,
      grades.final == null ? "" : grades.final >= 75 ? "PASSED" : "FAILED",
    );
  };

  MAPEH_TEMPLATE_SHEETS.forEach(({ sheet, storageTerm }) => {
    (Object.keys(componentConfig) as Array<keyof typeof componentConfig>).forEach((component) => {
      const config = componentConfig[component];
      const sheetActivities = allActivities
        .filter((activity) => activity.term === storageTerm && activity.component === component)
        .sort((a, b) => a.position - b.position)
        .slice(0, config.columns.length);
      const weight =
        allComponents.find((item) => item.term === storageTerm && item.component === component)
          ?.weight ?? config.defaultWeight;

      setCellValue(sheet, config.weightCell, weight / 100);
      setCellValue(sheet, config.headerCell, `${config.label} (${weight}%)`);

      config.columns.forEach((column, activityIndex) => {
        const activity = sheetActivities[activityIndex];
        setCellValue(sheet, `${column}10`, activity?.hps);

        maleStudents.forEach((student, studentIndex) => {
          setCellValue(
            sheet,
            `${column}${studentIndex + 12}`,
            activity ? (scoreMap.get(`${activity.id}|${student.id}`) ?? "") : "",
          );
        });
        femaleStudents.forEach((student, studentIndex) => {
          setCellValue(
            sheet,
            `${column}${studentIndex + 63}`,
            activity ? (scoreMap.get(`${activity.id}|${student.id}`) ?? "") : "",
          );
        });
      });
    });
  });

  // Excel summary rows are offset by one row from the MA/PEH detail sheets:
  // detail male row 12 -> summary male row 13
  // detail female row 63 -> summary female row 64
  maleStudents.forEach((student, index) => {
    writeMapehSummaryStudent(student, index + 12, index + 13);
  });
  femaleStudents.forEach((student, index) => {
    writeMapehSummaryStudent(student, index + 63, index + 64);
  });

  // JHS MAPEH: clear nonexistent learners' component/summary results as
  // literal empty cells. Template VLOOKUPs otherwise label blank learners
  // "Emerging" and count them in the achievement totals.
  if (isJuniorHighMapehExport) {
    MAPEH_TEMPLATE_SHEETS.forEach(({ sheet }) => {
      for (let index = maleStudents.length; index < 50; index += 1) {
        setCellValue(sheet, `AC${index + 12}`, "");
      }
      for (let index = femaleStudents.length; index < 50; index += 1) {
        setCellValue(sheet, `AC${index + 63}`, "");
      }
    });

    const clearUnusedSummaryRows = (
      sheetName: string,
      columns: readonly string[],
    ) => {
      for (let index = maleStudents.length; index < 50; index += 1) {
        columns.forEach((column) =>
          setCellValue(sheetName, `${column}${index + 13}`, ""),
        );
      }
      for (let index = femaleStudents.length; index < 50; index += 1) {
        columns.forEach((column) =>
          setCellValue(sheetName, `${column}${index + 64}`, ""),
        );
      }
    };

    ["MAPEH_T1", "MAPEH_T2", "MAPEH_T3"].forEach((sheetName) =>
      clearUnusedSummaryRows(sheetName, ["F", "J", "N", "V"]),
    );
    clearUnusedSummaryRows(
      "SUMMARY OF GRADES",
      ["F", "J", "N", "R", "V", "Z"],
    );
  }

  for (const [sheetName, values] of pendingValues) {
    const sheetPath = sheetPaths.get(sheetName);
    const sheetFile = sheetPath ? zip.file(sheetPath) : null;
    if (!sheetPath || !sheetFile) {
      throw new Error(`Unable to read the ${sheetName} worksheet.`);
    }

    const sheetXml = parser.parseFromString(await sheetFile.async("text"), "application/xml");
    const cells = new Map(
      Array.from(sheetXml.getElementsByTagName("c")).map((cell) => [
        cell.getAttribute("r") ?? "",
        cell,
      ]),
    );

    for (const [address, value] of values) {
      const cell = cells.get(address);
      if (!cell) {
        throw new Error(`The template cell ${sheetName}!${address} is missing.`);
      }

      // Preserve an existing Excel formula and only refresh its cached
      // displayed value. This keeps linked template cells valid while
      // making the exported workbook show the current UI data immediately.
      const formulaNode = Array.from(cell.childNodes).find(
        (child) => child.nodeName === "f",
      );

      const replaceJhsMapehTermGrade =
        isJuniorHighMapehExport &&
        MAPEH_TEMPLATE_SHEETS.some(({ sheet }) => sheet === sheetName) &&
        /^AC\d+$/.test(address);
      const replaceJhsMapehSummaryResult =
        isJuniorHighMapehExport &&
        ((["MAPEH_T1", "MAPEH_T2", "MAPEH_T3"].includes(sheetName) &&
          /^(?:F|J|N|V)\d+$/.test(address)) ||
          (sheetName === "SUMMARY OF GRADES" &&
            /^(?:F|J|N|R|V|Z)\d+$/.test(address)));
      if (formulaNode && !replaceJhsMapehTermGrade && !replaceJhsMapehSummaryResult) {
        Array.from(cell.childNodes)
          .filter((child) => child.nodeName === "v")
          .forEach((child) => cell.removeChild(child));

        if (value === "") {
          cell.removeAttribute("t");
          continue;
        }

        const cachedValueNode = sheetXml.createElementNS(
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "v",
        );

        if (typeof value === "number" && Number.isFinite(value)) {
          cell.removeAttribute("t");
          cachedValueNode.textContent = String(value);
        } else {
          cell.setAttribute("t", "str");
          cachedValueNode.textContent = String(value);
        }

        cell.appendChild(cachedValueNode);
        continue;
      }

      Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
      cell.removeAttribute("t");

      if (value === "") continue;

      if (typeof value === "number" && Number.isFinite(value)) {
        const numberNode = sheetXml.createElementNS(
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "v",
        );
        numberNode.textContent = String(value);
        cell.appendChild(numberNode);
      } else {
        cell.setAttribute("t", "inlineStr");
        const inlineString = sheetXml.createElementNS(
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "is",
        );
        const textNode = sheetXml.createElementNS(
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "t",
        );
        textNode.textContent = String(value);
        inlineString.appendChild(textNode);
        cell.appendChild(inlineString);
      }
    }

    if (isJuniorHighMapehExport && sheetName !== "INPUT DATA") {
      hideUnusedJhsExcelRows(
        sheetXml,
        MAPEH_TEMPLATE_SHEETS.some(({ sheet }) => sheet === sheetName)
          ? "term"
          : "summary",
        maleStudents.length,
        femaleStudents.length,
      );
    }
    zip.file(sheetPath, serializer.serializeToString(sheetXml));
  }

  if (isJuniorHighMapehExport) {
    zip.remove("xl/calcChain.xml");
    Array.from(relationshipsXml.getElementsByTagName("Relationship"))
      .filter((relationship) =>
        (relationship.getAttribute("Target") || "").endsWith("calcChain.xml") ||
        (relationship.getAttribute("Type") || "").endsWith("/calcChain"),
      )
      .forEach((relationship) => relationship.parentNode?.removeChild(relationship));
    zip.file("xl/_rels/workbook.xml.rels", serializer.serializeToString(relationshipsXml));
    const contentTypesFile = zip.file("[Content_Types].xml");
    if (contentTypesFile) {
      const contentTypesXml = parser.parseFromString(
        await contentTypesFile.async("text"),
        "application/xml",
      );
      Array.from(contentTypesXml.getElementsByTagName("Override"))
        .filter((entry) => entry.getAttribute("PartName") === "/xl/calcChain.xml")
        .forEach((entry) => entry.parentNode?.removeChild(entry));
      zip.file("[Content_Types].xml", serializer.serializeToString(contentTypesXml));
    }
  }

  const calculationProperties = workbookXml.getElementsByTagName("calcPr")[0];
  if (calculationProperties) {
    calculationProperties.setAttribute("fullCalcOnLoad", "1");
    calculationProperties.setAttribute("forceFullCalc", "1");
  }
  zip.file("xl/workbook.xml", serializer.serializeToString(workbookXml));

  const output = await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  const blob = new Blob([output], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  const safeClassName = [klass.grade_level, klass.section, "MAPEH-Class-Record"]
    .filter(Boolean)
    .join("-")
    .replace(/[^a-z0-9-_]+/gi, "-")
    .replace(/^-+|-+$/g, "");
  anchor.href = url;
  anchor.download = `${safeClassName || "MAPEH-Class-Record"}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);

  return students.length > maleStudents.length + femaleStudents.length;
}

function cleanLrnInput(value: string) {
  return value.replace(/\D/g, "").slice(0, 12);
}

function isValidLrn(value: string) {
  return /^\d{12}$/.test(cleanLrnInput(value));
}

type TransferAwareStudent = StudentRow & {
  enrollment_status?: "active" | "transferred_in" | "transferred_out" | null;
  is_active?: boolean | null;
  transfer_date?: string | null;
  transfer_effective_term?: string | null;
  previous_school?: string | null;
  destination_school?: string | null;
  transfer_reason?: string | null;
  transferred_at?: string | null;
  remarks?: string | null;
};

function localDateInputValue() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function transferStatus(
  student: StudentRow,
): "active" | "transferred_in" | "transferred_out" {
  const transferStudent = student as TransferAwareStudent;

  if (
    transferStudent.enrollment_status === "transferred_in" ||
    transferStudent.enrollment_status === "transferred_out"
  ) {
    return transferStudent.enrollment_status;
  }

  // Backward compatibility for transfer records created before the
  // student-transfer migration was installed.
  const remarks = String(transferStudent.remarks ?? "").trim();
  if (remarks.startsWith("[TRANSFER_OUT]")) return "transferred_out";
  if (remarks.startsWith("[TRANSFER_IN]")) return "transferred_in";

  return "active";
}

function roundInitialGrade(initialGrade: number | null | undefined): number | null {
  if (typeof initialGrade !== "number" || Number.isNaN(initialGrade)) {
    return null;
  }

  const grade = Math.max(0, Math.min(100, initialGrade));

  // Term Grade Base uses ordinary half-up rounding:
  // 66.49 -> 66, 66.50 -> 67, and 66.77 -> 67.
  return Math.floor(grade + 0.5);
}

function resolveTermGradeBase(
  computedBase: number | null,
  storedBase: number | null | undefined,
): number | null {
  if (computedBase == null) return null;

  const numericBase = storedBase == null ? computedBase : Number(storedBase);
  const resolvedBase = Number.isFinite(numericBase) ? numericBase : computedBase;

  return Math.max(0, Math.min(100, Math.floor(resolvedBase + 0.5)));
}

function EClassRecordSkeleton() {
  return (
    <div
      className="space-y-4 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading E-Class Record...</span>

      {/* Subject title bar */}
      <div
        className="relative overflow-hidden rounded-2xl border border-amber-200/80 bg-gradient-to-r from-[#fffdf8] via-[#fffbf2] to-[#fff6e8] px-4 py-3 shadow-sm"
        aria-hidden="true"
      >
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-3">
            <Skeleton className="size-11 shrink-0 rounded-xl" />
            <div className="space-y-2">
              <div className="flex items-center gap-2">
                <Skeleton className="h-5 w-32" />
                <Skeleton className="h-5 w-28 rounded-full" />
              </div>
              <Skeleton className="h-3.5 w-48" />
            </div>
          </div>
          <div className="flex items-center gap-2">
            <Skeleton className="h-7 w-24 rounded-full" />
            <Skeleton className="h-9 w-40 rounded-lg" />
          </div>
        </div>
      </div>

      {/* School information form */}
      <div className="rounded-2xl border bg-card p-4 shadow-sm" aria-hidden="true">
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {Array.from({ length: 12 }).map((_, index) => (
            <div key={index} className="space-y-2">
              <Skeleton className={`h-3 ${index % 3 === 0 ? "w-24" : "w-20"}`} />
              <Skeleton className="h-9 w-full rounded-lg" />
              {index === 2 && <Skeleton className="h-3 w-4/5" />}
            </div>
          ))}
        </div>
      </div>

      {/* E-Class Record toolbar */}
      <div className="rounded-2xl border bg-card p-2 shadow-sm" aria-hidden="true">
        <div className="flex flex-wrap items-center gap-2">
          <Skeleton className="h-9 w-44 rounded-lg" />
          <Skeleton className="h-9 w-20 rounded-lg" />
          <Skeleton className="h-9 w-44 rounded-lg" />
        </div>
        <div className="mt-2 flex flex-wrap justify-end gap-2">
          {Array.from({ length: 7 }).map((_, index) => (
            <Skeleton
              key={index}
              className={`h-9 rounded-lg ${index < 3 ? "w-32" : "w-28"}`}
            />
          ))}
        </div>
      </div>

      {/* Term controls */}
      <div
        className="flex flex-wrap items-center gap-2 rounded-2xl border bg-card p-2 shadow-sm"
        aria-hidden="true"
      >
        {Array.from({ length: 4 }).map((_, index) => (
          <Skeleton key={index} className="h-9 w-24 rounded-lg" />
        ))}
        <div className="flex shrink-0 items-center gap-2">
          <Skeleton className="h-8 w-28 rounded-full" />
          <Skeleton className="h-9 w-24 rounded-lg" />
          <Skeleton className="h-9 w-32 rounded-lg" />
        </div>
      </div>

      {/* Grade table */}
      <div className="overflow-hidden rounded-2xl border bg-card shadow-sm" aria-hidden="true">
        <div className="overflow-x-auto">
          <div className="min-w-[1180px]">
            <div className="grid grid-cols-[190px_repeat(11,minmax(72px,1fr))] border-b bg-muted/20">
              <div className="row-span-2 border-r p-3">
                <Skeleton className="h-4 w-28" />
              </div>
              <div className="col-span-5 border-r p-3">
                <Skeleton className="mx-auto h-4 w-40" />
              </div>
              <div className="col-span-3 border-r p-3">
                <Skeleton className="mx-auto h-4 w-44" />
              </div>
              <div className="col-span-3 p-3">
                <Skeleton className="mx-auto h-4 w-44" />
              </div>
            </div>

            <div className="grid grid-cols-[190px_repeat(11,minmax(72px,1fr))] border-b bg-muted/10">
              <div className="border-r p-2" />
              {Array.from({ length: 11 }).map((_, index) => (
                <div key={index} className="border-r p-2 last:border-r-0">
                  <Skeleton className="mx-auto h-7 w-11 rounded-md" />
                </div>
              ))}
            </div>

            <div className="border-b bg-blue-50/60 px-3 py-2">
              <Skeleton className="h-4 w-16" />
            </div>
            {Array.from({ length: 4 }).map((_, rowIndex) => (
              <div
                key={`male-${rowIndex}`}
                className="grid grid-cols-[190px_repeat(11,minmax(72px,1fr))] border-b"
              >
                <div className="border-r p-3">
                  <Skeleton className={`h-4 ${rowIndex % 2 === 0 ? "w-36" : "w-28"}`} />
                </div>
                {Array.from({ length: 11 }).map((_, cellIndex) => (
                  <div key={cellIndex} className="border-r p-2 last:border-r-0">
                    <Skeleton className="mx-auto h-7 w-11 rounded-md" />
                  </div>
                ))}
              </div>
            ))}

            <div className="border-b bg-pink-50/60 px-3 py-2">
              <Skeleton className="h-4 w-20" />
            </div>
            {Array.from({ length: 3 }).map((_, rowIndex) => (
              <div
                key={`female-${rowIndex}`}
                className="grid grid-cols-[190px_repeat(11,minmax(72px,1fr))] border-b last:border-b-0"
              >
                <div className="border-r p-3">
                  <Skeleton className={`h-4 ${rowIndex % 2 === 0 ? "w-32" : "w-40"}`} />
                </div>
                {Array.from({ length: 11 }).map((_, cellIndex) => (
                  <div key={cellIndex} className="border-r p-2 last:border-r-0">
                    <Skeleton className="mx-auto h-7 w-11 rounded-md" />
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="space-y-2" aria-hidden="true">
        <Skeleton className="h-3 w-full max-w-4xl" />
        <Skeleton className="h-3 w-full max-w-3xl" />
      </div>
    </div>
  );
}

function ClassDetail() {
  const { classId } = Route.useParams();
  const qc = useQueryClient();
  const [tab, setTab] = useState<Tab>("grades");

  const [duplicateSubjectNotice, setDuplicateSubjectNotice] = useState<{
    subject: string;
    existingClassId: string | null;
  } | null>(null);

  const { data: klass, isLoading: klassLoading } = useQuery({
    queryKey: ["class", classId],
    queryFn: async () => {
      const { data, error } = await supabase.from("classes").select("*").eq("id", classId).single();
      if (error) throw error;
      return data as ClassWithTrack;
    },
  });
  const { data: profile, isLoading: profileLoading } = useQuery({
    queryKey: ["profile"],
    queryFn: async () => {
      const id = await getUserId();
      const { data, error } = await supabase
        .from("profiles")
        .select("*")
        .eq("id", id)
        .maybeSingle();
      if (error) throw error;
      return data;
    },
  });
  const { data: schoolYearSetting, isLoading: schoolYearLoading } = useQuery<SchoolYearSetting>({
    queryKey: ["school-year-setting"],
    queryFn: getSchoolYearSetting,
    refetchInterval: 15000,
    refetchOnWindowFocus: true,
  });

  const activeSchoolYear =
    schoolYearSetting?.is_locked && schoolYearSetting.school_year
      ? schoolYearSetting.school_year
      : null;

  const { data: schoolYearLibrary = [], isLoading: schoolYearLibraryLoading } = useQuery<SchoolYearLibraryRow[]>({
    queryKey: ["school-year-library"],
    queryFn: getSchoolYearLibrary,
    refetchInterval: 15000,
    refetchOnWindowFocus: true,
  });

  const { data: students = [], isLoading: studentsLoading } = useQuery({
    queryKey: ["students", classId],
    queryFn: async () => {
      // Use the untyped client locally because the generated Supabase types
      // may not include the new transfer columns yet.
      const { data, error } = await (supabase as any)
        .from("students")
        .select("*")
        .eq("class_id", classId)
        .order("last_name");
      if (error) throw error;

      // The E-Class Record is the active roster. A transferred-out learner
      // remains in the database for historical grades/records but is hidden
      // from the current editable class roster.
      const activeRows = (data ?? []).filter((row: any) => {
        if (row.is_active === false) return false;

        // Backward compatibility for old transfer-out rows that were stored
        // only in remarks before the transfer migration was installed.
        const remarks = String(row.remarks ?? "").trim();
        return !remarks.startsWith("[TRANSFER_OUT]");
      });

      // Protect the E-Class Record from old duplicate rows already stored
      // in the database. Only one copy of the same learner is displayed.
      return deduplicateStudentRows(activeRows as StudentRow[]);
    },
  });

  const { data: subjectTeachers = [], isLoading: subjectTeachersLoading } = useQuery({
    queryKey: ["subject-teacher-directory"],
    queryFn: async () => {
      // `teacher_type` exists in the database migration, but older generated
      // Supabase types do not include it yet. Keep this query locally untyped
      // until `src/integrations/supabase/types.ts` is regenerated.
      const { data, error } = await (supabase as any)
        .from("profiles")
        .select("id, full_name, email")
        .eq("teacher_type", "subject_teacher")
        .order("full_name");
      if (error) throw error;
      return (data ?? []) as SubjectTeacherOption[];
    },
  });

  const { data: studentAssignments = [], isLoading: studentAssignmentsLoading } = useQuery({
    queryKey: ["student-list-assignments", classId],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("student_list_assignments")
        .select("id, subject_teacher_id, subject, status, decline_reason, sent_at, responded_at")
        .eq("class_id", classId)
        .order("sent_at", { ascending: false });
      if (error) throw error;
      return (data ?? []) as StudentListAssignment[];
    },
  });

  const updateClass = useMutation({
    mutationFn: async (patch: Partial<ClassWithTrack>) => {
      const { data, error } = await (supabase as any)
        .from("classes")
        .update(patch)
        .eq("id", classId)
        .select("*")
        .single();

      if (error) throw error;
      return data as ClassWithTrack;
    },
    onSuccess: async (savedClass) => {
      // Keep both query caches synchronized. SF9 reads ["classes"], while
      // this page reads ["class", classId].
      qc.setQueryData(["class", classId], savedClass);
      qc.setQueryData<ClassWithTrack[]>(["classes"], (current = []) =>
        current.map((item) => (item.id === savedClass.id ? savedClass : item)),
      );

      await Promise.all([
        qc.invalidateQueries({ queryKey: ["class", classId] }),
        qc.invalidateQueries({ queryKey: ["classes"] }),
      ]);
    },
    onError: (error: Error) => {
      toast.error(`Unable to save class details: ${error.message}`);
    },
  });
  const updateAcceptedSubjectClass = useMutation({
    mutationFn: async (details: {
      schoolName: string;
      schoolId: string;
      schoolYear: string;
      region: string;
      division: string;
      district: string;
      gradeLevel: string;
      section: string;
      startDate: string | null;
      endDate: string | null;
      subject: string;
    }) => {
      const { error } = await (supabase as any).rpc("update_accepted_subject_class_details", {
        p_class_id: classId,
        p_school_name: details.schoolName,
        p_school_id: details.schoolId,
        p_school_year: details.schoolYear,
        p_region: details.region,
        p_division: details.division,
        p_district: details.district,
        p_grade_level: details.gradeLevel,
        p_section: details.section,
        p_start_date: details.startDate,
        p_end_date: details.endDate,
        p_subject: details.subject,
      });
      if (error) throw error;
    },
    onSuccess: async () => {
      await Promise.all([
        qc.invalidateQueries({ queryKey: ["class", classId] }),
        qc.invalidateQueries({ queryKey: ["classes"] }),
        qc.invalidateQueries({ queryKey: ["student-list-assignments", classId] }),
      ]);
      toast.success("Assigned class details updated");
    },
    onError: async (error: unknown, details) => {
      const databaseError = error as {
        code?: string;
        message?: string;
        details?: string;
      };

      const errorMessage =
        databaseError?.message ||
        databaseError?.details ||
        (error instanceof Error ? error.message : String(error));

      const duplicateSubject =
        databaseError?.code === "23505" ||
        errorMessage.includes(
          "student_list_assignments_class_id_subject_teacher_id_subject_key",
        ) ||
        errorMessage.toLowerCase().includes(
          "duplicate key value violates unique constraint",
        );

      // Reload the saved class after any failed update so the Subject dropdown
      // returns to the actual database value instead of looking changed.
      await Promise.all([
        qc.invalidateQueries({ queryKey: ["class", classId] }),
        qc.invalidateQueries({ queryKey: ["classes"] }),
        qc.invalidateQueries({ queryKey: ["student-list-assignments", classId] }),
      ]);

      if (duplicateSubject) {
        let existingClassId: string | null = null;

        // Try to locate the already-existing subject class so the message can
        // offer a direct "Open Existing Subject Class" action.
        if (profile?.id) {
          const { data: existingClasses } = await (supabase as any)
            .from("classes")
            .select("id, subject, grade_level, section, teacher_id")
            .eq("teacher_id", profile.id)
            .eq("grade_level", details.gradeLevel)
            .eq("subject", details.subject)
            .neq("id", classId)
            .limit(10);

          const matchedClass = (existingClasses ?? []).find(
            (item: any) =>
              normalize(item.section) === normalize(details.section) &&
              normalize(item.subject) === normalize(details.subject),
          );

          existingClassId = matchedClass?.id ?? null;
        }

        setDuplicateSubjectNotice({
          subject: details.subject,
          existingClassId,
        });
        return;
      }

      toast.error("Unable to change the subject. Please try again.");
    },
  });
  const updatePrincipal = useMutation({
    mutationFn: async (principal: string) => {
      const id = await getUserId();
      const { error } = await supabase
        .from("profiles")
        .upsert({ id, principal, updated_at: new Date().toISOString() });
      if (error) throw error;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ["profile"] }),
  });

  const [addOpen, setAddOpen] = useState<false | "male" | "female">(false);
  const [sendStudentsOpen, setSendStudentsOpen] = useState(false);
  const [editStudentsOpen, setEditStudentsOpen] = useState(false);
  const [deleteStudentsOpen, setDeleteStudentsOpen] = useState(false);
  const [transferInOpen, setTransferInOpen] = useState(false);
  const [transferOutOpen, setTransferOutOpen] = useState(false);
  const [pendingSubject, setPendingSubject] = useState<string | null>(null);

  const pageLoading =
    klassLoading ||
    profileLoading ||
    schoolYearLoading ||
    schoolYearLibraryLoading ||
    studentsLoading ||
    subjectTeachersLoading ||
    studentAssignmentsLoading;

  if (pageLoading || !klass) {
    return <EClassRecordSkeleton />;
  }

  const subjectTheme = getSubjectTheme(klass.subject ?? "");
  const SubjectHeaderIcon = subjectTheme.Icon;

  const canSendStudentList =
    normalize((profile as any)?.teacher_type) === "class_adviser" &&
    klass.teacher_id === profile?.id;
  const isAcceptedSubjectClass = Boolean(
    (klass as ClassRow & { student_list_assignment_id?: string | null }).student_list_assignment_id,
  );

  const saveAssignedClassDetails = (patch: {
    school_name?: string;
    school_id?: string;
    school_year?: string;
    region?: string;
    division?: string;
    district?: string;
    grade_level?: string;
    section?: string;
    start_date?: string | null;
    end_date?: string | null;
    subject?: string;
    track_shs?: string;
  }) => {
    // Admin School Year rules:
    // - If one School Year is locked, only that locked value may be saved.
    // - If none is locked, teachers may choose only from the Admin library.
    // - Existing historical classes keep their saved value until explicitly
    //   changed while the library is unlocked.
    if (Object.prototype.hasOwnProperty.call(patch, "school_year")) {
      const requestedSchoolYear = String(patch.school_year ?? "").trim();

      if (activeSchoolYear) {
        if (requestedSchoolYear !== activeSchoolYear) {
          toast.error(
            `Only the active School Year ${activeSchoolYear} is allowed. The School Year was not saved.`,
          );
          return;
        }

        patch = {
          ...patch,
          school_year: activeSchoolYear,
        };
      } else {
        const existsInLibrary = schoolYearLibrary.some(
          (year) => year.school_year === requestedSchoolYear,
        );

        if (!requestedSchoolYear || !existsInLibrary) {
          toast.error(
            "Select a School Year from the Administrator's School Year Library.",
          );
          return;
        }

        patch = {
          ...patch,
          school_year: requestedSchoolYear,
        };
      }
    }

    // Track (SHS only) is stored directly on public.classes so both
    // Class Adviser classes and accepted Subject Teacher class copies can edit it.
    if (Object.prototype.hasOwnProperty.call(patch, "track_shs")) {
      updateClass.mutate({ track_shs: patch.track_shs ?? "" });
      return;
    }

    if (!isAcceptedSubjectClass) {
      updateClass.mutate(patch);
      return;
    }

    const nextGradeLevel = patch.grade_level ?? klass.grade_level ?? "";
    const availableSubjects = SUBJECTS_BY_GRADE[nextGradeLevel] ?? [];
    const currentSubject = patch.subject ?? klass.subject ?? "";
    const nextSubject = availableSubjects.includes(currentSubject)
      ? currentSubject
      : (availableSubjects[0] ?? "");

    updateAcceptedSubjectClass.mutate({
      schoolName: patch.school_name ?? klass.school_name ?? "",
      schoolId: patch.school_id ?? klass.school_id ?? "",
      schoolYear: patch.school_year ?? klass.school_year ?? "",
      region: patch.region ?? klass.region ?? "",
      division: patch.division ?? klass.division ?? "",
      district: patch.district ?? klass.district ?? "",
      gradeLevel: nextGradeLevel,
      section: patch.section ?? klass.section ?? "",
      startDate: patch.start_date ?? klass.start_date ?? null,
      endDate: patch.end_date ?? klass.end_date ?? null,
      subject: nextSubject,
    });
  };

  const requestSubjectChange = (nextSubject: string) => {
    if (nextSubject === (klass.subject ?? "")) return;
    setPendingSubject(nextSubject);
  };

  const cancelSubjectChange = () => {
    setPendingSubject(null);
  };

  const confirmSubjectChange = () => {
    if (!pendingSubject) return;
    saveAssignedClassDetails({ subject: pendingSubject });
    setPendingSubject(null);
  };

  const subjectChangePending = updateClass.isPending || updateAcceptedSubjectClass.isPending;

  return (
    <div className="space-y-4 pb-24 md:pb-6">
      {/* Title bar - same subject theme as the dashboard class card */}
      <div
        className={`relative isolate overflow-hidden rounded-2xl border px-4 py-3 shadow-sm ${subjectTheme.cardClass}`}
      >
        <SubjectHeaderArtwork theme={subjectTheme} />
        <div className="pointer-events-none absolute inset-0 bg-gradient-to-r from-white/90 via-white/70 to-white/5" />

        <div className="relative z-10 flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-3">
            <div
              className={`grid size-11 shrink-0 place-items-center rounded-xl shadow-sm ring-1 backdrop-blur-sm ${subjectTheme.iconBackgroundClass}`}
            >
              <SubjectHeaderIcon className={`size-5 ${subjectTheme.iconClass}`} />
            </div>

            <div className="min-w-0">
              <div className="flex flex-wrap items-center gap-2">
                <div className="truncate text-lg font-semibold text-foreground">
                  {klass.subject || "Untitled Class"}
                </div>
                <span
                  className={`rounded-full border px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wide backdrop-blur-sm ${subjectTheme.badgeClass}`}
                >
                  {subjectTheme.category}
                </span>
              </div>
              <div className="text-xs text-foreground/65">
                School Learner Grade Monitoring
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <span className="chip bg-[color:var(--success-bg)] text-[color:var(--success)]">
              Autosaved
            </span>
            <Link
              to="/dashboard"
              className="inline-flex items-center gap-1 rounded-lg border border-white/70 bg-white/70 px-3 py-1.5 text-sm text-foreground shadow-sm backdrop-blur-sm hover:bg-white/90"
            >
              <ArrowLeft className="size-4" />
              Back to My Classes
            </Link>
          </div>
        </div>
      </div>

      {/* School Header Form */}
      <div className="rounded-2xl border bg-card p-4 shadow-sm">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <HeaderField
            label="School Name"
            value={klass.school_name ?? ""}
            onSave={(v) => saveAssignedClassDetails({ school_name: v })}
            inputMode="text"
            sanitize={(value) =>
              value
                .replace(/[^\p{L}\s]/gu, "")
                .replace(/\s{2,}/g, " ")
                .replace(/^\s+/, "")
            }
            invalidMessage="School Name must contain letters and spaces only."
          />
          <HeaderField
            label="School ID"
            value={klass.school_id ?? ""}
            onSave={(v) => saveAssignedClassDetails({ school_id: v })}
            inputMode="numeric"
            pattern="[0-9]*"
            sanitize={(value) => value.replace(/\D/g, "")}
            invalidMessage="School ID must contain numbers only."
          />
          <div>
            <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
              School Year
            </div>

            {activeSchoolYear ? (
              <Input
                value={klass.school_year ?? activeSchoolYear}
                readOnly
                aria-readonly="true"
                className={
                  klass.school_year === activeSchoolYear
                    ? "h-9 cursor-not-allowed bg-emerald-50/50 text-foreground"
                    : "h-9 cursor-not-allowed border-amber-300 bg-amber-50/60 text-foreground"
                }
                aria-label="School Year controlled by administrator"
              />
            ) : (
              <Select
                value={
                  klass.school_year &&
                  schoolYearLibrary.some(
                    (year) => year.school_year === klass.school_year,
                  )
                    ? klass.school_year
                    : undefined
                }
                onValueChange={(value) =>
                  saveAssignedClassDetails({ school_year: value })
                }
                disabled={
                  schoolYearLibrary.length === 0 ||
                  updateClass.isPending ||
                  updateAcceptedSubjectClass.isPending
                }
              >
                <SelectTrigger
                  className={
                    schoolYearLibrary.length === 0
                      ? "h-9 border-amber-300 bg-amber-50/60"
                      : "h-9"
                  }
                >
                  <SelectValue placeholder="-- Select School Year --" />
                </SelectTrigger>
                <SelectContent>
                  {schoolYearLibrary.map((year) => (
                    <SelectItem key={year.id} value={year.school_year}>
                      {year.school_year}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}

            <p
              className={`mt-1 text-[10px] ${
                activeSchoolYear && activeSchoolYear === klass.school_year
                  ? "text-emerald-700"
                  : "text-amber-700"
              }`}
            >
              {activeSchoolYear
                ? activeSchoolYear === klass.school_year
                  ? `Locked by Admin: ${activeSchoolYear}. Only this School Year is allowed.`
                  : `Historical class: ${klass.school_year || "No School Year"}. The current locked School Year is ${activeSchoolYear}, so this historical value cannot be changed while the lock is active.`
                : schoolYearLibrary.length > 0
                  ? klass.school_year
                    ? `Current class School Year: ${klass.school_year}. No year is locked, so you may change it using the Admin School Year Library.`
                    : "No year is locked. Select the correct School Year from the Admin School Year Library."
                  : "No School Years are available. Ask the Administrator to add one to the School Year Library."}
            </p>
          </div>
          <HeaderField
            label="Region"
            value={klass.region ?? ""}
            onSave={(v) => saveAssignedClassDetails({ region: v })}
            inputMode="text"
            sanitize={(value) =>
              value
                .replace(/[^\p{L}\p{N}\s]/gu, "")
                .replace(/\s{2,}/g, " ")
                .replace(/^\s+/, "")
            }
            invalidMessage="Region can contain letters, numbers, and spaces only. Special characters are not allowed."
          />
          <HeaderField
            label="Division"
            value={klass.division ?? ""}
            onSave={(v) => saveAssignedClassDetails({ division: v })}
            inputMode="text"
            sanitize={(value) =>
              value
                .replace(/[^\p{L}\s.-]/gu, "")
                .replace(/\s{2,}/g, " ")
                .replace(/^\s+/, "")
            }
            invalidMessage="Division must contain text only. Numbers are not allowed."
          />
          <HeaderField
            label="District"
            value={klass.district ?? ""}
            onSave={(v) => saveAssignedClassDetails({ district: v })}
            inputMode="text"
            sanitize={(value) =>
              value
                .replace(/[^\p{L}\s.-]/gu, "")
                .replace(/\s{2,}/g, " ")
                .replace(/^\s+/, "")
            }
            invalidMessage="District must contain text only. Numbers are not allowed."
          />
          <div className="grid grid-cols-2 gap-3">
            <div>
              <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
                Grade Level
              </div>
              <Select
                value={klass.grade_level || undefined}
                onValueChange={(v) => saveAssignedClassDetails({ grade_level: v })}
                disabled={updateAcceptedSubjectClass.isPending}
              >
                <SelectTrigger className="h-9">
                  <SelectValue placeholder="-- Select Grade --" />
                </SelectTrigger>
                <SelectContent>
                  {["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"].map(
                    (g) => (
                      <SelectItem key={g} value={g}>
                        {g}
                      </SelectItem>
                    ),
                  )}
                </SelectContent>
              </Select>
            </div>
            <HeaderField
              label="Section"
              value={klass.section ?? ""}
              onSave={(v) => saveAssignedClassDetails({ section: v })}
              inputMode="text"
              sanitize={(value) =>
                value
                  .replace(/[^\p{L}\s.-]/gu, "")
                  .replace(/\s{2,}/g, " ")
                  .replace(/^\s+/, "")
              }
              invalidMessage="Section must contain text only. Numbers are not allowed."
            />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
                Start Date
              </div>
              <Input
                type="date"
                defaultValue={klass.start_date ?? ""}
                onBlur={(e) =>
                  e.target.value !== (klass.start_date ?? "") &&
                  saveAssignedClassDetails({
                    start_date: e.target.value || null,
                  })
                }
                className="h-9"
              />
            </div>
            <div>
              <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
                End Date
              </div>
              <Input
                type="date"
                defaultValue={klass.end_date ?? ""}
                onBlur={(e) =>
                  e.target.value !== (klass.end_date ?? "") &&
                  saveAssignedClassDetails({
                    end_date: e.target.value || null,
                  })
                }
                className="h-9"
              />
            </div>
          </div>
          <HeaderField
            label="Teacher"
            value={
              isAcceptedSubjectClass
                ? profile?.full_name || profile?.email || klass.teacher_name || ""
                : klass.teacher_name || ""
            }
            onSave={
              isAcceptedSubjectClass ? undefined : (v) => updateClass.mutate({ teacher_name: v })
            }
          />
          <div>
            <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
              Subject
            </div>
            <Select
              value={klass.subject ?? undefined}
              onValueChange={requestSubjectChange}
              disabled={!klass.grade_level || subjectChangePending}
            >
              <SelectTrigger className="h-9">
                <SelectValue
                  placeholder={
                    klass.grade_level ? "-- Select Subject --" : "-- Select Grade First --"
                  }
                />
              </SelectTrigger>
              <SelectContent>
                {(SUBJECTS_BY_GRADE[klass.grade_level] ?? []).map((s) => (
                  <SelectItem key={s} value={s}>
                    {s}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            {/* Senior High School only: Grade 11 requires 2/3/6 Units,
                while Grade 12 permits 3 Units or no units (SQL NULL). */}
            {(klass.grade_level === "Grade 11" || klass.grade_level === "Grade 12") && (
              <div className="mt-3">
                <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
                  Units{" "}
                  <span className="font-normal normal-case">
                    {klass.grade_level === "Grade 11" ? "(Required)" : "(Optional)"}
                  </span>
                </div>
                <Select
                  value={
                    klass.grade_level === "Grade 11"
                      ? klass.units != null && [2, 3, 6].includes(Number(klass.units))
                        ? String(klass.units)
                        : undefined
                      : klass.units === 3 ? "3" : "none"
                  }
                  onValueChange={(value) => {
                    if (klass.grade_level === "Grade 11" && !["2", "3", "6"].includes(value)) {
                      return;
                    }
                    updateClass.mutate({ units: value === "none" ? null : Number(value) });
                  }}
                  disabled={updateClass.isPending}
                >
                  <SelectTrigger
                    className="h-9"
                    aria-label={
                      klass.grade_level === "Grade 11"
                        ? "Grade 11 required units"
                        : "Grade 12 optional units"
                    }
                  >
                    <SelectValue placeholder="-- Select Units --" />
                  </SelectTrigger>
                  <SelectContent>
                    {klass.grade_level === "Grade 12" ? (
                      <>
                        <SelectItem value="none">None (No Units)</SelectItem>
                        <SelectItem value="3">3 Units</SelectItem>
                      </>
                    ) : (
                      <>
                        <SelectItem value="2">2 Units</SelectItem>
                        <SelectItem value="3">3 Units</SelectItem>
                        <SelectItem value="6">6 Units</SelectItem>
                      </>
                    )}
                  </SelectContent>
                </Select>
                <p className="mt-1 text-[10px] text-muted-foreground">
                  {klass.grade_level === "Grade 11"
                    ? "Required for Grade 11. Select 2, 3, or 6 Units; changes are saved automatically."
                    : "Optional for Grade 12. Select 3 Units or None; changes are saved automatically."}
                </p>
              </div>
            )}
          </div>
          <HeaderField
            key={profile?.principal ?? ""}
            label="Principal"
            value={profile?.principal ?? ""}
            onSave={(v) => updatePrincipal.mutate(v)}
          />
          <TrackShsField
            value={klass.track_shs ?? ""}
            onSave={(value) => updateClass.mutate({ track_shs: value })}
          />
        </div>
      </div>

      {/* Subject Change Confirmation */}
      <Dialog
        open={pendingSubject !== null}
        onOpenChange={(open) => {
          if (!open && !subjectChangePending) {
            cancelSubjectChange();
          }
        }}
      >
        <DialogContent className="overflow-hidden border-0 p-0 shadow-2xl sm:max-w-[520px]">
          <div className="px-6 pb-5 pt-7 text-center sm:px-8">
            <div className="mx-auto grid size-16 place-items-center rounded-full border-2 border-amber-300 bg-amber-50 text-amber-600">
              <AlertTriangle className="size-8" />
            </div>

            <DialogHeader className="mt-4 space-y-2 text-center">
              <DialogTitle className="text-center text-2xl font-bold tracking-tight">
                Change Subject?
              </DialogTitle>
              <p className="text-center text-sm text-muted-foreground">
                Are you sure you want to change the subject?
              </p>
            </DialogHeader>

            <div className="mt-6 grid items-stretch gap-3 sm:grid-cols-[1fr_auto_1fr]">
              <div className="rounded-2xl border bg-amber-50/70 p-4">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                  Current Subject
                </div>
                <div className="mt-3">
                  <span className="inline-flex min-w-24 justify-center rounded-full bg-orange-100 px-4 py-1.5 text-sm font-bold text-orange-800">
                    {klass.subject || "Not selected"}
                  </span>
                </div>
              </div>

              <div className="flex items-center justify-center text-2xl font-semibold text-muted-foreground">
                →
              </div>

              <div className="rounded-2xl border bg-blue-50/70 p-4">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                  New Subject
                </div>
                <div className="mt-3">
                  <span className="inline-flex min-w-24 justify-center rounded-full bg-blue-100 px-4 py-1.5 text-sm font-bold text-blue-800">
                    {pendingSubject || "—"}
                  </span>
                </div>
              </div>
            </div>

            <div className="mt-5 rounded-xl bg-muted/40 px-4 py-3 text-sm leading-6 text-muted-foreground">
              Changing the subject will refresh the E-Class Record activities, grades, and computed
              summaries for the selected subject.
            </div>
          </div>

          <DialogFooter className="border-t bg-muted/20 px-6 py-4 sm:px-8">
            <Button
              type="button"
              variant="outline"
              onClick={cancelSubjectChange}
              disabled={subjectChangePending}
              className="min-w-28"
            >
              Cancel
            </Button>
            <Button
              type="button"
              onClick={confirmSubjectChange}
              disabled={!pendingSubject || subjectChangePending}
              className="min-w-44 bg-red-700 text-white hover:bg-red-800"
            >
              {subjectChangePending ? (
                <Loader2 className="mr-2 size-4 animate-spin" />
              ) : (
                <AlertTriangle className="mr-2 size-4" />
              )}
              Yes, Change Subject
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Duplicate Subject Notice - friendly UI for both Class Adviser and Subject Teacher */}
      <Dialog
        open={duplicateSubjectNotice !== null}
        onOpenChange={(open) => {
          if (!open) {
            setDuplicateSubjectNotice(null);
          }
        }}
      >
        <DialogContent
          className="overflow-hidden border-0 p-0 shadow-2xl sm:max-w-[430px]"
          onEscapeKeyDown={() => setDuplicateSubjectNotice(null)}
        >
          <div className="px-7 pb-7 pt-8 text-center sm:px-9">
            <div className="relative mx-auto grid size-20 place-items-center">
              <div className="absolute inset-0 rounded-full bg-red-100/80" />
              <div className="relative grid size-16 place-items-center rounded-full border border-red-200 bg-red-50 text-red-500 shadow-sm">
                <AlertTriangle className="size-9" strokeWidth={2.2} />
              </div>
            </div>

            <DialogHeader className="mt-4 space-y-3 text-center">
              <DialogTitle className="text-center text-2xl font-bold tracking-tight">
                Subject Already Exists
              </DialogTitle>

              <div className="flex justify-center">
                <span className="inline-flex items-center gap-2 rounded-full border border-red-200 bg-red-50 px-3 py-1 text-xs font-semibold text-red-600">
                  <Users className="size-3.5" />
                  For Class Adviser &amp; Subject Teacher
                </span>
              </div>
            </DialogHeader>

            <div className="mt-5 space-y-3 text-sm leading-6 text-muted-foreground">
              <p>
                <span className="font-semibold text-foreground">
                  {duplicateSubjectNotice?.subject ?? "This subject"}
                </span>{" "}
                is already assigned to this class for the selected teacher.
                No duplicate class or student list was created.
              </p>

              <p>
                Please open the existing subject class or choose a different subject.
              </p>
            </div>

            <div className="mt-7 grid gap-3">
              <Button
                type="button"
                className="h-11 w-full bg-red-600 text-white hover:bg-red-700"
                onClick={() => {
                  const existingClassId =
                    duplicateSubjectNotice?.existingClassId ?? null;

                  setDuplicateSubjectNotice(null);

                  if (existingClassId) {
                    window.location.assign(`/classes/${existingClassId}`);
                    return;
                  }

                  // If the exact class row cannot be resolved from the current
                  // database response, return to My Classes so the teacher can
                  // open the existing subject safely.
                  window.location.assign("/dashboard");
                }}
              >
                <ArrowRight className="mr-2 size-4" />
                Open Existing Subject Class
              </Button>

              <Button
                type="button"
                variant="outline"
                className="h-11 w-full border-red-300 text-red-600 hover:bg-red-50 hover:text-red-700"
                onClick={() => setDuplicateSubjectNotice(null)}
              >
                Choose Another Subject
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      {/* Toolbar */}
      <div className="rounded-2xl border bg-card p-2 shadow-sm">
        {/* First line: E-Class Record navigation */}
        <div className="flex flex-wrap items-center gap-2">
          <ToolTab
          active={tab === "grades"}
          onClick={() => setTab("grades")}
          icon={<ClipboardList className="size-4" />}
          label="E-Class Record: Grades"
          tone="violet"
        />
        <ToolTab
          active={tab === "fitness"}
          onClick={() => setTab("fitness")}
          icon={<HeartPulse className="size-4" />}
          label="BMI"
        />
        <div className="mx-1 h-6 w-px bg-border" />
        <button
          onClick={() => window.print()}
          className="inline-flex shrink-0 items-center gap-1 whitespace-nowrap rounded-lg bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground hover:opacity-90"
        >
          <Download className="size-4" />
          Preview & Download
        </button>
        </div>

        {/* Second line: student management actions, aligned to the left */}
        <div className="mt-2 flex flex-wrap items-center justify-start gap-2">
          {canSendStudentList && (
            <button
              onClick={() => setSendStudentsOpen(true)}
              className="inline-flex items-center gap-1 rounded-lg border border-primary px-3 py-1.5 text-sm font-medium text-primary hover:bg-primary/10"
            >
              <Send className="size-4" />
              Send Students
            </button>
          )}
          <button
            onClick={() => setEditStudentsOpen(true)}
            className="inline-flex items-center gap-1 rounded-lg border border-amber-400 px-3 py-1.5 text-sm font-medium text-amber-700 hover:bg-amber-50"
          >
            <Pencil className="size-4" />
            Edit Students
          </button>
          <button
            onClick={() => setDeleteStudentsOpen(true)}
            className="inline-flex items-center gap-1 rounded-lg border border-red-400 px-3 py-1.5 text-sm font-medium text-red-700 hover:bg-red-50"
          >
            <Trash2 className="size-4" />
            Delete Students
          </button>
          <button
            onClick={() => setAddOpen("male")}
            className="inline-flex items-center gap-1 rounded-lg border border-blue-400 px-3 py-1.5 text-sm font-medium text-blue-600 hover:bg-blue-50"
          >
            <UserPlus2 className="size-4" />
            Add Male
          </button>
          <button
            onClick={() => setAddOpen("female")}
            className="inline-flex items-center gap-1 rounded-lg border border-pink-400 px-3 py-1.5 text-sm font-medium text-pink-600 hover:bg-pink-50"
          >
            <UserPlus2 className="size-4" />
            Add Female
          </button>
          {canSendStudentList && (
            <>
              <button
                onClick={() => setTransferInOpen(true)}
                className="inline-flex items-center gap-1 rounded-lg border border-emerald-500 px-3 py-1.5 text-sm font-medium text-emerald-700 hover:bg-emerald-50"
              >
                <LogIn className="size-4" />
                Transfer In
              </button>
              <button
                onClick={() => setTransferOutOpen(true)}
                className="inline-flex items-center gap-1 rounded-lg border border-red-500 px-3 py-1.5 text-sm font-medium text-red-700 hover:bg-red-50"
              >
                <LogOut className="size-4" />
                Transfer Out
              </button>
            </>
          )}
        </div>
      </div>
      {tab === "grades" &&
        (klass.subject === "MAPEH" ? (
          <MapehGradesPanel
            classId={classId}
            students={students}
            subject={klass.subject}
            klass={klass}
            principal={profile?.principal ?? ""}
          />
        ) : isCommunicationCompositeSubject(klass.subject) ? (
          <CommunicationGradesPanel
            classId={classId}
            students={students}
            subject={klass.subject}
            klass={klass}
            principal={profile?.principal ?? ""}
          />
        ) : (
          <GradesPanel
            classId={classId}
            students={students}
            subject={klass?.subject ?? ""}
            klass={klass}
            principal={profile?.principal ?? ""}
          />
        ))}
      {tab === "fitness" && <FitnessTestPanel classId={classId} students={students} />}

      <AddStudentDialog
        open={!!addOpen}
        sex={addOpen || "male"}
        classId={classId}
        onClose={() => setAddOpen(false)}
      />

      {canSendStudentList && (
        <>
          <TransferInDialog
            open={transferInOpen}
            classId={classId}
            klass={klass}
            onClose={() => setTransferInOpen(false)}
          />
          <TransferOutDialog
            open={transferOutOpen}
            classId={classId}
            klass={klass}
            students={students}
            onClose={() => setTransferOutOpen(false)}
          />
        </>
      )}

      {canSendStudentList && (
        <SendStudentsDialog
          open={sendStudentsOpen}
          klass={klass}
          classId={classId}
          students={students}
          subjectTeachers={subjectTeachers}
          assignments={studentAssignments}
          onClose={() => setSendStudentsOpen(false)}
        />
      )}

      <EditStudentsDialog
        open={editStudentsOpen}
        classId={classId}
        students={students}
        onClose={() => setEditStudentsOpen(false)}
      />

      <DeleteStudentsDialog
        open={deleteStudentsOpen}
        classId={classId}
        students={students}
        onClose={() => setDeleteStudentsOpen(false)}
      />
    </div>
  );
}

/* ---------------- Small pieces ---------------- */
function TrackShsField({ value, onSave }: { value: string; onSave: (value: string) => void }) {
  const [localValue, setLocalValue] = useState(value);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  useEffect(() => {
    return () => {
      if (saveTimerRef.current) {
        clearTimeout(saveTimerRef.current);
      }
    };
  }, []);

  const queueSave = (nextValue: string) => {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
    }

    saveTimerRef.current = setTimeout(() => {
      onSave(nextValue.trim());
      saveTimerRef.current = null;
    }, 500);
  };

  const saveNow = () => {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }

    const cleaned = localValue.trim();

    if (cleaned !== (value ?? "")) {
      onSave(cleaned);
    }
  };

  return (
    <div>
      <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
        Track (SHS only)
      </div>

      <Input
        value={localValue}
        placeholder=""
        onChange={(e) => {
          const nextValue = e.target.value;
          setLocalValue(nextValue);
          queueSave(nextValue);
        }}
        onBlur={saveNow}
        className="h-9"
      />
    </div>
  );
}

function HeaderField({
  label,
  value,
  onSave,
  inputMode,
  pattern,
  sanitize,
  invalidMessage,
}: {
  label: string;
  value: string;
  onSave?: (v: string) => void;
  inputMode?: InputHTMLAttributes<HTMLInputElement>["inputMode"];
  pattern?: string;
  sanitize?: (value: string) => string;
  invalidMessage?: string;
}) {
  const [hasInvalidInput, setHasInvalidInput] = useState(false);

  return (
    <div>
      <div
        className={`mb-1 text-[10px] font-semibold uppercase tracking-wide ${
          hasInvalidInput ? "text-red-600" : "text-muted-foreground"
        }`}
      >
        {label}
      </div>

      <Input
        defaultValue={value}
        placeholder={label}
        readOnly={!onSave}
        inputMode={inputMode}
        pattern={pattern}
        aria-invalid={hasInvalidInput}
        onInput={(e) => {
          if (!sanitize) return;

          const input = e.currentTarget;
          const original = input.value;
          const cleaned = sanitize(original);
          const invalid = original !== cleaned;

          setHasInvalidInput(invalid);

          // Keep the valid part of the value, but immediately warn the user.
          if (invalid) {
            input.value = cleaned;
          }
        }}
        onBlur={(e) => {
          const cleaned = sanitize
            ? sanitize(e.target.value).trim()
            : e.target.value.trim();

          if (e.target.value !== cleaned) {
            e.target.value = cleaned;
          }

          // Clear the red warning after the user leaves a corrected field.
          setHasInvalidInput(false);

          if (cleaned !== value) {
            onSave?.(cleaned);
          }
        }}
        className={`h-9 read-only:bg-muted read-only:text-muted-foreground ${
          hasInvalidInput
            ? "border-red-500 bg-red-50/50 text-red-900 ring-2 ring-red-500/20 focus-visible:ring-red-500/30"
            : ""
        }`}
      />

      {hasInvalidInput && invalidMessage && (
        <p className="mt-1 text-[11px] font-medium text-red-600" role="alert">
          {invalidMessage}
        </p>
      )}
    </div>
  );
}

function TransferInDialog({
  open,
  classId,
  klass,
  onClose,
}: {
  open: boolean;
  classId: string;
  klass: ClassRow;
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const [form, setForm] = useState({
    last_name: "",
    first_name: "",
    middle_name: "",
    sex: "male" as "male" | "female",
    lrn: "",
    previous_school: "",
    transfer_date: localDateInputValue(),
    effective_term: "Term 1",
    reason: "",
  });

  useEffect(() => {
    if (!open) return;

    setForm({
      last_name: "",
      first_name: "",
      middle_name: "",
      sex: "male",
      lrn: "",
      previous_school: "",
      transfer_date: localDateInputValue(),
      effective_term: "Term 1",
      reason: "",
    });
  }, [open]);

  const transferIn = useMutation({
    mutationFn: async () => {
      const lastName = form.last_name.trim();
      const firstName = form.first_name.trim();
      const middleName = form.middle_name.trim();
      const previousSchool = form.previous_school.trim();
      const reason = form.reason.trim();
      const lrn = cleanLrnInput(form.lrn);

      if (!lastName || !firstName) {
        throw new Error("Last name and first name are required.");
      }

      if (!isValidLrn(lrn)) {
        throw new Error("LRN must contain exactly 12 digits.");
      }

      if (!previousSchool) {
        throw new Error("Previous school is required for a transferred-in learner.");
      }

      if (!form.transfer_date) {
        throw new Error("Transfer date is required.");
      }

      const teacherId = await getUserId();

      // Reuse an existing inactive learner with the same LRN when possible.
      // This prevents a second learner identity from being created after a
      // previous Transfer Out.
      const { data: existingRows, error: existingError } = await (supabase as any)
        .from("students")
        .select(
          "id, class_id, teacher_id, last_name, first_name, middle_name, sex, lrn, enrollment_status, is_active",
        )
        .eq("lrn", lrn)
        .limit(10);

      if (existingError) throw existingError;

      const existingStudent = (existingRows ?? [])[0] as
        | {
            id: string;
            class_id: string | null;
            teacher_id: string | null;
            last_name: string | null;
            first_name: string | null;
            middle_name: string | null;
            sex: string | null;
            lrn: string | null;
            enrollment_status:
              | "active"
              | "transferred_in"
              | "transferred_out"
              | null;
            is_active: boolean | null;
          }
        | undefined;

      let studentId: string;
      let createdTemporaryStudent = false;

      if (existingStudent) {
        const isActive = existingStudent.is_active !== false;

        if (isActive && existingStudent.class_id === classId) {
          throw new Error("This learner is already active in this class.");
        }

        if (
          isActive &&
          existingStudent.class_id !== classId &&
          existingStudent.enrollment_status !== "transferred_out"
        ) {
          throw new Error(
            "This learner is still active in another class. Transfer the learner out first before transferring in.",
          );
        }

        studentId = existingStudent.id;

        const { error: updateIdentityError } = await (supabase as any)
          .from("students")
          .update({
            teacher_id: teacherId,
            last_name: lastName,
            first_name: firstName,
            middle_name: middleName || null,
            sex: form.sex,
            lrn,
          })
          .eq("id", studentId);

        if (updateIdentityError) throw updateIdentityError;
      } else {
        // transfer_in_student() operates on an existing student row.
        // Create the incoming learner as temporarily inactive, then let the
        // RPC activate the learner and write student_transfers history.
        const { data: createdStudent, error: insertError } = await (supabase as any)
          .from("students")
          .insert({
            class_id: classId,
            teacher_id: teacherId,
            last_name: lastName,
            first_name: firstName,
            middle_name: middleName || null,
            sex: form.sex,
            lrn,
            enrollment_status: "active",
            is_active: false,
          })
          .select("id")
          .single();

        if (insertError) throw insertError;
        if (!createdStudent?.id) {
          throw new Error("Unable to create the incoming learner record.");
        }

        studentId = String(createdStudent.id);
        createdTemporaryStudent = true;
      }

      const { error: transferError } = await (supabase as any).rpc(
        "transfer_in_student",
        {
          p_student_id: studentId,
          p_to_class_id: classId,
          p_transfer_date: form.transfer_date,
          p_previous_school: previousSchool,
          p_effective_term: form.effective_term,
          p_reason: reason || null,
          p_remarks: null,
        },
      );

      if (transferError) {
        // Avoid leaving a hidden temporary learner behind if the RPC fails.
        if (createdTemporaryStudent) {
          await (supabase as any).from("students").delete().eq("id", studentId);
        }
        throw transferError;
      }

      return studentId;
    },
    onSuccess: async () => {
      toast.success("Learner transferred in successfully");

      await Promise.all([
        qc.invalidateQueries({ queryKey: ["students", classId] }),
        qc.invalidateQueries({ queryKey: ["students-count"] }),
        qc.invalidateQueries({ queryKey: ["class-adviser-dashboard-students"] }),
        qc.invalidateQueries({ queryKey: ["student-transfers"] }),
      ]);

      onClose();
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const learnerName = [form.last_name.trim(), form.first_name.trim()]
    .filter(Boolean)
    .join(", ");

  const gradeSection =
    [klass.grade_level, klass.section].filter(Boolean).join(" - ") || "—";

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) =>
        !nextOpen && !transferIn.isPending && onClose()
      }
    >
      <DialogContent className="max-h-[92vh] overflow-y-auto border-0 p-0 shadow-2xl sm:max-w-[680px]">
        <div className="px-6 pb-5 pt-7 sm:px-8">
          <div className="text-center">
            <div className="mx-auto grid size-16 place-items-center rounded-full bg-emerald-100 text-emerald-700 ring-8 ring-emerald-50">
              <LogIn className="size-8" />
            </div>

            <DialogHeader className="mt-5 space-y-2 text-center">
              <DialogTitle className="text-center text-2xl font-bold tracking-tight">
                Transfer Student In
              </DialogTitle>
              <p className="text-center text-sm text-muted-foreground">
                Add an incoming learner to this class and save the transfer in
                the transfer history.
              </p>
            </DialogHeader>
          </div>

          <div className="mt-6 grid gap-4 sm:grid-cols-2">
            <div>
              <Label>Last name</Label>
              <Input
                value={form.last_name}
                onChange={(event) =>
                  setForm({ ...form, last_name: event.target.value })
                }
                className="mt-1"
              />
            </div>

            <div>
              <Label>First name</Label>
              <Input
                value={form.first_name}
                onChange={(event) =>
                  setForm({ ...form, first_name: event.target.value })
                }
                className="mt-1"
              />
            </div>

            <div>
              <Label>Middle name</Label>
              <Input
                value={form.middle_name}
                onChange={(event) =>
                  setForm({ ...form, middle_name: event.target.value })
                }
                className="mt-1"
              />
            </div>

            <div>
              <Label>Sex</Label>
              <Select
                value={form.sex}
                onValueChange={(value) =>
                  setForm({
                    ...form,
                    sex: value === "female" ? "female" : "male",
                  })
                }
              >
                <SelectTrigger className="mt-1">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="male">Male</SelectItem>
                  <SelectItem value="female">Female</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div>
              <Label>LRN</Label>
              <Input
                value={form.lrn}
                inputMode="numeric"
                maxLength={12}
                pattern="[0-9]{12}"
                placeholder="12-digit LRN"
                aria-invalid={form.lrn.length > 0 && !isValidLrn(form.lrn)}
                onChange={(event) =>
                  setForm({
                    ...form,
                    lrn: cleanLrnInput(event.target.value),
                  })
                }
                className={`mt-1 ${
                  form.lrn.length > 0 && !isValidLrn(form.lrn)
                    ? "border-destructive focus-visible:ring-destructive"
                    : ""
                }`}
              />

              {form.lrn.length > 0 && !isValidLrn(form.lrn) && (
                <p className="mt-1 text-xs text-destructive">
                  LRN must be exactly 12 digits. {form.lrn.length}/12 entered.
                </p>
              )}
            </div>

            <div>
              <Label>Transfer date</Label>
              <div className="relative mt-1">
                <CalendarDays className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  type="date"
                  value={form.transfer_date}
                  onChange={(event) =>
                    setForm({ ...form, transfer_date: event.target.value })
                  }
                  className="pl-9"
                />
              </div>
            </div>

            <div>
              <Label>Effective term</Label>
              <Select
                value={form.effective_term}
                onValueChange={(value) =>
                  setForm({ ...form, effective_term: value })
                }
              >
                <SelectTrigger className="mt-1">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="Term 1">Term 1</SelectItem>
                  <SelectItem value="Term 2">Term 2</SelectItem>
                  <SelectItem value="Term 3">Term 3</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div>
              <Label>School year</Label>
              <Input
                value={klass.school_year ?? ""}
                readOnly
                className="mt-1 bg-muted"
              />
            </div>

            <div className="sm:col-span-2">
              <Label>Previous school</Label>
              <div className="relative mt-1">
                <School className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  value={form.previous_school}
                  placeholder="Enter previous school"
                  onChange={(event) =>
                    setForm({ ...form, previous_school: event.target.value })
                  }
                  className="pl-9"
                />
              </div>
            </div>

            <div className="sm:col-span-2">
              <Label>Reason / remarks</Label>
              <Input
                value={form.reason}
                placeholder="Optional"
                onChange={(event) =>
                  setForm({ ...form, reason: event.target.value })
                }
                className="mt-1"
              />
            </div>
          </div>

          <div className="mt-5 grid gap-3 sm:grid-cols-2">
            <div className="rounded-2xl border bg-muted/20 p-4">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                Learner
              </div>
              <div className="mt-1 font-semibold">
                {learnerName || "Enter learner name"}
              </div>
            </div>

            <div className="rounded-2xl border bg-muted/20 p-4">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                Grade &amp; Section
              </div>
              <div className="mt-1 font-semibold">{gradeSection}</div>
            </div>
          </div>

          <div className="mt-5 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">
            The learner becomes active in this class. The transfer is recorded
            in <span className="font-semibold">student_transfers</span> so it
            also appears in the Transfer In &amp; Transfer Out list.
          </div>
        </div>

        <DialogFooter className="border-t bg-muted/20 px-6 py-4 sm:px-8">
          <Button
            type="button"
            variant="outline"
            onClick={onClose}
            disabled={transferIn.isPending}
          >
            Cancel
          </Button>

          <Button
            type="button"
            onClick={() => transferIn.mutate()}
            disabled={
              transferIn.isPending ||
              !form.last_name.trim() ||
              !form.first_name.trim() ||
              !isValidLrn(form.lrn) ||
              !form.previous_school.trim() ||
              !form.transfer_date
            }
            className="bg-emerald-600 text-white hover:bg-emerald-700"
          >
            {transferIn.isPending ? (
              <Loader2 className="mr-2 size-4 animate-spin" />
            ) : (
              <LogIn className="mr-2 size-4" />
            )}
            Yes, Transfer In
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function TransferOutDialog({
  open,
  classId,
  klass,
  students,
  onClose,
}: {
  open: boolean;
  classId: string;
  klass: ClassRow;
  students: StudentRow[];
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const eligibleStudents = useMemo(() => students, [students]);
  const [studentId, setStudentId] = useState("");
  const [receivingSchool, setReceivingSchool] = useState("");
  const [transferDate, setTransferDate] = useState(localDateInputValue());
  const [effectiveTerm, setEffectiveTerm] = useState("Term 1");
  const [reason, setReason] = useState("");

  useEffect(() => {
    if (!open) return;

    setStudentId(eligibleStudents[0]?.id ?? "");
    setReceivingSchool("");
    setTransferDate(localDateInputValue());
    setEffectiveTerm("Term 1");
    setReason("");
  }, [open, eligibleStudents]);

  const selectedStudent =
    eligibleStudents.find((student) => student.id === studentId) ?? null;

  const currentStatus = selectedStudent
    ? transferStatus(selectedStudent) === "transferred_in"
      ? "Transferred In"
      : "Active"
    : "—";

  const gradeSection =
    [klass.grade_level, klass.section].filter(Boolean).join(" - ") || "—";

  const transferOut = useMutation({
    mutationFn: async () => {
      if (!selectedStudent) {
        throw new Error("Select a learner to transfer out.");
      }

      if (!transferDate) {
        throw new Error("Transfer date is required.");
      }

      const { error } = await (supabase as any).rpc("transfer_out_student", {
        p_student_id: selectedStudent.id,
        p_transfer_date: transferDate,
        p_destination_school: receivingSchool.trim() || null,
        p_reason: reason.trim() || null,
        p_effective_term: effectiveTerm,
        p_remarks: null,
      });

      if (error) throw error;
    },
    onSuccess: async () => {
      toast.success("Learner transferred out successfully");

      await Promise.all([
        qc.invalidateQueries({ queryKey: ["students", classId] }),
        qc.invalidateQueries({ queryKey: ["students-count"] }),
        qc.invalidateQueries({ queryKey: ["class-adviser-dashboard-students"] }),
        qc.invalidateQueries({ queryKey: ["student-transfers"] }),
      ]);

      onClose();
    },
    onError: (error: Error) => toast.error(error.message),
  });

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) =>
        !nextOpen && !transferOut.isPending && onClose()
      }
    >
      <DialogContent className="max-h-[92vh] overflow-y-auto border-0 p-0 shadow-2xl sm:max-w-[680px]">
        <div className="px-6 pb-5 pt-7 sm:px-8">
          <div className="text-center">
            <div className="mx-auto grid size-16 place-items-center rounded-full bg-red-100 text-red-600 ring-8 ring-red-50">
              <LogOut className="size-8" />
            </div>

            <DialogHeader className="mt-5 space-y-2 text-center">
              <DialogTitle className="text-center text-2xl font-bold tracking-tight">
                Transfer Student Out
              </DialogTitle>
              <p className="text-center text-sm text-muted-foreground">
                Mark the selected learner as transferred out while preserving
                all existing grades and historical records.
              </p>
            </DialogHeader>
          </div>

          <div className="mt-6">
            <Label>Select learner</Label>
            <Select value={studentId || undefined} onValueChange={setStudentId}>
              <SelectTrigger className="mt-1">
                <SelectValue placeholder="Select learner" />
              </SelectTrigger>
              <SelectContent>
                {eligibleStudents.map((student) => (
                  <SelectItem key={student.id} value={student.id}>
                    {student.last_name}, {student.first_name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          {selectedStudent ? (
            <>
              <div className="mt-5 grid items-stretch gap-3 sm:grid-cols-[1fr_auto_1fr]">
                <div className="rounded-2xl border bg-emerald-50/70 p-4 text-center">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                    Current Status
                  </div>
                  <div className="mt-3 inline-flex items-center gap-2 rounded-full bg-emerald-100 px-4 py-1.5 text-sm font-bold text-emerald-800">
                    <span className="size-2 rounded-full bg-emerald-500" />
                    {currentStatus}
                  </div>
                </div>

                <div className="flex items-center justify-center text-muted-foreground">
                  <ArrowRight className="size-6" />
                </div>

                <div className="rounded-2xl border bg-red-50/70 p-4 text-center">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                    New Status
                  </div>
                  <div className="mt-3 inline-flex items-center gap-2 rounded-full bg-red-100 px-4 py-1.5 text-sm font-bold text-red-700">
                    <span className="size-2 rounded-full bg-red-500" />
                    Transferred Out
                  </div>
                </div>
              </div>

              <div className="mt-5 grid gap-4 sm:grid-cols-2">
                <div>
                  <Label>Receiving school</Label>
                  <div className="relative mt-1">
                    <School className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
                    <Input
                      value={receivingSchool}
                      placeholder="Enter receiving school"
                      onChange={(event) =>
                        setReceivingSchool(event.target.value)
                      }
                      className="pl-9"
                    />
                  </div>
                </div>

                <div>
                  <Label>Transfer date</Label>
                  <div className="relative mt-1">
                    <CalendarDays className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
                    <Input
                      type="date"
                      value={transferDate}
                      onChange={(event) =>
                        setTransferDate(event.target.value)
                      }
                      className="pl-9"
                    />
                  </div>
                </div>

                <div>
                  <Label>Effective term</Label>
                  <Select value={effectiveTerm} onValueChange={setEffectiveTerm}>
                    <SelectTrigger className="mt-1">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Term 1">Term 1</SelectItem>
                      <SelectItem value="Term 2">Term 2</SelectItem>
                      <SelectItem value="Term 3">Term 3</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                <div>
                  <Label>School year</Label>
                  <Input
                    value={klass.school_year ?? ""}
                    readOnly
                    className="mt-1 bg-muted"
                  />
                </div>

                <div className="sm:col-span-2">
                  <Label>Reason / remarks</Label>
                  <Input
                    value={reason}
                    placeholder="Optional"
                    onChange={(event) => setReason(event.target.value)}
                    className="mt-1"
                  />
                </div>
              </div>

              <div className="mt-5 grid gap-3 sm:grid-cols-2">
                <div className="rounded-2xl border bg-muted/20 p-4">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                    Learner
                  </div>
                  <div className="mt-1 font-semibold">
                    {selectedStudent.last_name}, {selectedStudent.first_name}
                  </div>
                </div>

                <div className="rounded-2xl border bg-muted/20 p-4">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                    Grade &amp; Section
                  </div>
                  <div className="mt-1 font-semibold">{gradeSection}</div>
                </div>
              </div>

              <div className="mt-5 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">
                The learner remains in the database. Existing grades and
                historical records are preserved, but the learner is removed
                from the current active E-Class Record roster.
              </div>
            </>
          ) : (
            <div className="mt-5 rounded-xl border border-dashed p-8 text-center text-sm text-muted-foreground">
              No active learners are available to transfer out.
            </div>
          )}
        </div>

        <DialogFooter className="border-t bg-muted/20 px-6 py-4 sm:px-8">
          <Button
            type="button"
            variant="outline"
            onClick={onClose}
            disabled={transferOut.isPending}
          >
            Cancel
          </Button>

          <Button
            type="button"
            onClick={() => transferOut.mutate()}
            disabled={!selectedStudent || !transferDate || transferOut.isPending}
            className="bg-red-600 text-white hover:bg-red-700"
          >
            {transferOut.isPending ? (
              <Loader2 className="mr-2 size-4 animate-spin" />
            ) : (
              <LogOut className="mr-2 size-4" />
            )}
            Yes, Transfer Out
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function SendStudentsDialog({
  open,
  klass,
  classId,
  students,
  subjectTeachers,
  assignments,
  onClose,
}: {
  open: boolean;
  klass: ClassRow;
  classId: string;
  students: StudentRow[];
  subjectTeachers: SubjectTeacherOption[];
  assignments: StudentListAssignment[];
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const [subject, setSubject] = useState(klass.subject ?? "");
  const [teacherId, setTeacherId] = useState("");
  const [assignmentToDelete, setAssignmentToDelete] = useState<StudentListAssignment | null>(null);
  const maleStudents = useMemo(
    () => students.filter((student) => normalize(student.sex) === "male"),
    [students],
  );
  const femaleStudents = useMemo(
    () => students.filter((student) => normalize(student.sex) === "female"),
    [students],
  );

  useEffect(() => {
    if (open) {
      setSubject(klass.subject ?? "");
      setTeacherId("");
    }
  }, [open, klass.subject]);

  const sendAssignment = useMutation({
    mutationFn: async () => {
      if (!subject || !teacherId) throw new Error("Select a subject and Subject Teacher first.");
      if (!students.length) throw new Error("Add at least one student before sending the list.");
      const { data, error } = await (supabase as any).rpc("send_student_list_assignment", {
        p_class_id: classId,
        p_subject_teacher_id: teacherId,
        p_subject: subject,
      });
      if (error) throw error;
      return data;
    },
    onSuccess: () => {
      toast.success("Student list sent to the Subject Teacher");
      qc.invalidateQueries({ queryKey: ["student-list-assignments", classId] });
      setTeacherId("");
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const deleteAssignment = useMutation({
    mutationFn: async (assignmentId: string) => {
      const { error } = await (supabase as any).rpc("delete_student_list_assignment", {
        p_assignment_id: assignmentId,
      });
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success("Assignment permanently deleted. You can send it again.");
      setAssignmentToDelete(null);
      qc.invalidateQueries({ queryKey: ["student-list-assignments", classId] });
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const teacherLabel = (id: string) => {
    const teacher = subjectTeachers.find((item) => item.id === id);
    return teacher?.full_name || teacher?.email || "Subject Teacher";
  };

  return (
    <Dialog open={open} onOpenChange={(next) => !next && onClose()}>
      <DialogContent className="max-h-[90vh] max-w-3xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Users className="size-5 text-primary" /> Send Student List
          </DialogTitle>
        </DialogHeader>

        <div className="grid gap-4">
          <div className="grid gap-3 rounded-xl border bg-muted/30 p-4 text-sm sm:grid-cols-4">
            {[
              ["Grade Level", klass.grade_level || "—"],
              ["Section", klass.section || "—"],
              ["School Year", klass.school_year || "—"],
              ["Students", String(students.length)],
            ].map(([label, value]) => (
              <div key={label}>
                <div className="text-xs text-muted-foreground">{label}</div>
                <div className="font-medium">{value}</div>
              </div>
            ))}
          </div>

          <div className="grid gap-3 sm:grid-cols-2">
            <div>
              <Label>Assigned Subject</Label>
              <Select value={subject || undefined} onValueChange={setSubject}>
                <SelectTrigger className="mt-1">
                  <SelectValue placeholder="Select subject" />
                </SelectTrigger>
                <SelectContent>
                  {(SUBJECTS_BY_GRADE[klass.grade_level] ?? []).map((item) => (
                    <SelectItem key={item} value={item}>
                      {item}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label>Subject Teacher</Label>
              <Select value={teacherId || undefined} onValueChange={setTeacherId}>
                <SelectTrigger className="mt-1">
                  <SelectValue placeholder="Select Subject Teacher" />
                </SelectTrigger>
                <SelectContent>
                  {subjectTeachers.map((teacher) => (
                    <SelectItem key={teacher.id} value={teacher.id}>
                      {teacher.full_name || teacher.email || "Unnamed teacher"}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </div>

          <div className="rounded-xl border">
            <div className="flex items-center justify-between border-b px-4 py-3">
              <div>
                <div className="font-medium">Student List Preview</div>
                <div className="text-xs text-muted-foreground">
                  The Subject Teacher must accept this list before managing the assigned subject.
                </div>
              </div>
              <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-semibold text-primary">
                {students.length} learners
              </span>
            </div>
            <div className="max-h-72 space-y-3 overflow-y-auto p-3">
              {!students.length ? (
                <div className="p-4 text-center text-sm text-muted-foreground">
                  No students have been added to this class.
                </div>
              ) : (
                <>
                  <div className="overflow-hidden rounded-lg border border-blue-200">
                    <div className="flex items-center justify-between bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-700">
                      <span>MALE</span>
                      <span>
                        {maleStudents.length} {maleStudents.length === 1 ? "learner" : "learners"}
                      </span>
                    </div>
                    <div className="p-1">
                      {!maleStudents.length ? (
                        <div className="px-3 py-3 text-sm text-muted-foreground">
                          No male learners.
                        </div>
                      ) : (
                        maleStudents.map((student, index) => (
                          <div
                            key={student.id}
                            className="flex gap-3 rounded-md px-3 py-2 text-sm odd:bg-blue-50/50"
                          >
                            <span className="w-6 text-xs text-muted-foreground">{index + 1}.</span>
                            <span className="font-medium">
                              {[student.last_name, student.first_name, student.middle_name]
                                .filter(Boolean)
                                .join(", ")}
                            </span>
                            <span className="ml-auto text-xs text-muted-foreground">
                              {student.lrn || "No LRN"}
                            </span>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  <div className="overflow-hidden rounded-lg border border-pink-200">
                    <div className="flex items-center justify-between bg-pink-50 px-3 py-2 text-xs font-semibold text-pink-700">
                      <span>FEMALE</span>
                      <span>
                        {femaleStudents.length}{" "}
                        {femaleStudents.length === 1 ? "learner" : "learners"}
                      </span>
                    </div>
                    <div className="p-1">
                      {!femaleStudents.length ? (
                        <div className="px-3 py-3 text-sm text-muted-foreground">
                          No female learners.
                        </div>
                      ) : (
                        femaleStudents.map((student, index) => (
                          <div
                            key={student.id}
                            className="flex gap-3 rounded-md px-3 py-2 text-sm odd:bg-pink-50/50"
                          >
                            <span className="w-6 text-xs text-muted-foreground">{index + 1}.</span>
                            <span className="font-medium">
                              {[student.last_name, student.first_name, student.middle_name]
                                .filter(Boolean)
                                .join(", ")}
                            </span>
                            <span className="ml-auto text-xs text-muted-foreground">
                              {student.lrn || "No LRN"}
                            </span>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>

          {!!assignments.length && (
            <div className="rounded-xl border p-4">
              <div className="mb-3 font-medium">Sent Assigned to Subject Teachers</div>
              <div className="space-y-2">
                {assignments.map((assignment) => (
                  <div
                    key={assignment.id}
                    className="flex flex-wrap items-center gap-2 rounded-lg bg-muted/40 px-3 py-2 text-sm"
                  >
                    <span className="font-medium">{assignment.subject}</span>
                    <span className="text-muted-foreground">
                      • {teacherLabel(assignment.subject_teacher_id)}
                    </span>
                    <span
                      className={`ml-auto rounded-full px-2.5 py-1 text-xs font-semibold capitalize ${
                        assignment.status === "accepted"
                          ? "bg-emerald-100 text-emerald-700"
                          : assignment.status === "declined"
                            ? "bg-red-100 text-red-700"
                            : "bg-amber-100 text-amber-700"
                      }`}
                    >
                      {assignment.status === "pending" ? "Pending Acceptance" : assignment.status}
                    </span>
                    <Button
                      type="button"
                      variant="ghost"
                      size="icon"
                      className="size-8 text-destructive hover:bg-destructive/10 hover:text-destructive"
                      title="Permanently delete assignment"
                      disabled={deleteAssignment.isPending}
                      onClick={() => setAssignmentToDelete(assignment)}
                    >
                      <Trash2 className="size-4" />
                      <span className="sr-only">Delete assignment</span>
                    </Button>
                    {assignment.status === "declined" && assignment.decline_reason && (
                      <div className="w-full text-xs text-red-600">
                        Reason: {assignment.decline_reason}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          <div className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800">
            After acceptance, the Subject Teacher can view these students and manage only the
            assigned subject. Student information remains controlled by the Class Adviser.
          </div>
        </div>

        {assignmentToDelete && (
          <div
            className="fixed inset-0 z-[60] flex items-center justify-center bg-black/55 p-4"
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="delete-assignment-title"
            aria-describedby="delete-assignment-description"
            onMouseDown={(event) => {
              if (event.target === event.currentTarget && !deleteAssignment.isPending) {
                setAssignmentToDelete(null);
              }
            }}
          >
            <div className="w-full max-w-lg overflow-hidden rounded-2xl border border-red-100 bg-background shadow-2xl">
              <div className="flex gap-4 border-b border-red-100 bg-red-50/70 p-6">
                <div className="flex size-12 shrink-0 items-center justify-center rounded-full bg-red-100 text-red-600">
                  <AlertTriangle className="size-6" />
                </div>
                <div>
                  <h2
                    id="delete-assignment-title"
                    className="text-xl font-semibold text-foreground"
                  >
                    Permanently delete this assigned?
                  </h2>
                  <p className="mt-1 text-sm text-muted-foreground">
                    This action cannot be undone.
                  </p>
                </div>
              </div>

              <div className="space-y-4 p-6">
                <div className="rounded-xl border bg-muted/30 p-4 text-sm">
                  <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    Assignment
                  </div>
                  <div className="mt-1 font-semibold">{assignmentToDelete.subject}</div>
                  <div className="mt-1 text-muted-foreground">
                    Subject Teacher: {teacherLabel(assignmentToDelete.subject_teacher_id)}
                  </div>
                </div>

                <p id="delete-assignment-description" className="text-sm leading-6 text-foreground">
                  The Subject Teacher&apos;s copied class, students, and grades will be permanently
                  deleted.
                </p>

                <div className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">
                  Your original Class Adviser student list will remain safe, and you can send this
                  subject again afterward.
                </div>
              </div>

              <div className="flex flex-col-reverse gap-2 border-t bg-muted/20 p-4 sm:flex-row sm:justify-end">
                <Button
                  type="button"
                  variant="outline"
                  disabled={deleteAssignment.isPending}
                  onClick={() => setAssignmentToDelete(null)}
                >
                  Cancel
                </Button>
                <Button
                  type="button"
                  variant="destructive"
                  disabled={deleteAssignment.isPending}
                  onClick={() => deleteAssignment.mutate(assignmentToDelete.id)}
                >
                  {deleteAssignment.isPending ? (
                    <Loader2 className="mr-2 size-4 animate-spin" />
                  ) : (
                    <Trash2 className="mr-2 size-4" />
                  )}
                  Permanently Delete
                </Button>
              </div>
            </div>
          </div>
        )}

        <DialogFooter>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            onClick={() => sendAssignment.mutate()}
            disabled={!teacherId || !subject || !students.length || sendAssignment.isPending}
          >
            {sendAssignment.isPending ? (
              <Loader2 className="mr-2 size-4 animate-spin" />
            ) : (
              <Send className="mr-2 size-4" />
            )}
            Send Student List
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function ToolTab({
  active,
  onClick,
  icon,
  label,
  tone,
}: {
  active: boolean;
  onClick: () => void;
  icon: ReactNode;
  label: string;
  tone?: "primary" | "violet";
}) {
  const activeCls =
    tone === "violet"
      ? "bg-violet-600 text-white"
      : tone === "primary"
        ? "bg-primary text-primary-foreground"
        : "bg-foreground text-background";
  return (
    <button
      onClick={onClick}
      className={`inline-flex shrink-0 items-center gap-1.5 whitespace-nowrap rounded-lg px-3 py-1.5 text-sm font-medium transition ${active ? activeCls : "border bg-background hover:bg-muted"}`}
    >
      {icon}
      {label}
    </button>
  );
}

function AddStudentDialog({
  open,
  sex,
  classId,
  onClose,
}: {
  open: boolean;
  sex: "male" | "female";
  classId: string;
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const [form, setForm] = useState<{
    last_name: string;
    first_name: string;
    middle_name: string;
    lrn: string;
  }>({ last_name: "", first_name: "", middle_name: "", lrn: "" });
  const add = useMutation({
    mutationFn: async () => {
      const lastName = form.last_name.trim();
      const firstName = form.first_name.trim();
      const middleName = form.middle_name.trim();
      const lrn = cleanLrnInput(form.lrn);

      if (!lastName || !firstName) {
        throw new Error("Last name and first name are required.");
      }

      if (!isValidLrn(lrn)) {
        throw new Error("LRN must contain exactly 12 digits.");
      }

      // Re-check the database immediately before insert. This prevents:
      // 1) repeated clicking on Add,
      // 2) adding the same learner through Add Male then Add Female,
      // 3) adding a duplicate learner with the same LRN/full name.
      const { data: existingRows, error: existingError } = await supabase
        .from("students")
        .select("id, last_name, first_name, middle_name, lrn")
        .eq("class_id", classId);

      if (existingError) throw existingError;

      const duplicate = ((existingRows ?? []) as Array<{
        id: string;
        last_name: string | null;
        first_name: string | null;
        middle_name: string | null;
        lrn: string | null;
      }>).find((student) =>
        sameLearner(student, {
          last_name: lastName,
          first_name: firstName,
          middle_name: middleName || null,
          lrn: lrn || null,
        }),
      );

      if (duplicate) {
        throw new Error("This learner is already in this class.");
      }

      const teacher_id = await getUserId();
      const { error } = await supabase.from("students").insert({
        class_id: classId,
        teacher_id,
        sex,
        last_name: lastName,
        first_name: firstName,
        middle_name: middleName || null,
        lrn: lrn || null,
      });
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success("Learner added");
      qc.invalidateQueries({ queryKey: ["students", classId] });
      setForm({ last_name: "", first_name: "", middle_name: "", lrn: "" });
      onClose();
    },
    onError: (e: Error) => toast.error(e.message),
  });
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Add {sex === "male" ? "Male" : "Female"} Learner</DialogTitle>
        </DialogHeader>
        <div className="grid gap-3 sm:grid-cols-2">
          <div>
            <Label>Last name</Label>
            <Input
              value={form.last_name}
              onChange={(e) => setForm({ ...form, last_name: e.target.value })}
            />
          </div>
          <div>
            <Label>First name</Label>
            <Input
              value={form.first_name}
              onChange={(e) => setForm({ ...form, first_name: e.target.value })}
            />
          </div>
          <div>
            <Label>Middle name</Label>
            <Input
              value={form.middle_name}
              onChange={(e) => setForm({ ...form, middle_name: e.target.value })}
            />
          </div>
          <div>
            <Label>LRN</Label>
            <Input
              value={form.lrn}
              inputMode="numeric"
              maxLength={12}
              pattern="[0-9]{12}"
              placeholder="12-digit LRN"
              aria-invalid={form.lrn.length > 0 && !isValidLrn(form.lrn)}
              className={
                form.lrn.length > 0 && !isValidLrn(form.lrn)
                  ? "border-destructive focus-visible:ring-destructive"
                  : undefined
              }
              onChange={(e) =>
                setForm({
                  ...form,
                  lrn: cleanLrnInput(e.target.value),
                })
              }
            />
            {!isValidLrn(form.lrn) && (
              <p className="mt-1 text-xs text-destructive">
                LRN must be exactly 12 digits. {form.lrn.length}/12 entered.
              </p>
            )}
          </div>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            disabled={
              !form.last_name.trim() ||
              !form.first_name.trim() ||
              !isValidLrn(form.lrn) ||
              add.isPending
            }
            onClick={() => add.mutate()}
          >
            {add.isPending ? "Adding..." : "Add"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function EditStudentsDialog({
  open,
  classId,
  students,
  onClose,
}: {
  open: boolean;
  classId: string;
  students: StudentRow[];
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const [selectedId, setSelectedId] = useState("");
  const [form, setForm] = useState<{
    last_name: string;
    first_name: string;
    middle_name: string;
    sex: "male" | "female";
  }>({
    last_name: "",
    first_name: "",
    middle_name: "",
    sex: "male",
  });

  const selectedStudent = students.find((student) => student.id === selectedId);

  useEffect(() => {
    if (!open) return;
    setSelectedId(students[0]?.id ?? "");
  }, [open, students]);

  useEffect(() => {
    if (!selectedStudent) {
      setForm({
        last_name: "",
        first_name: "",
        middle_name: "",
        sex: "male",
      });
      return;
    }

    setForm({
      last_name: selectedStudent.last_name ?? "",
      first_name: selectedStudent.first_name ?? "",
      middle_name: selectedStudent.middle_name ?? "",
      sex: normalize(selectedStudent.sex) === "female" ? "female" : "male",
    });
  }, [selectedStudent]);

  const refreshStudents = async () => {
    await Promise.all([
      qc.invalidateQueries({ queryKey: ["students", classId] }),
      qc.invalidateQueries({ queryKey: ["scores-all", classId] }),
      qc.invalidateQueries({ queryKey: ["scores-all-a", classId] }),
      qc.invalidateQueries({ queryKey: ["students-count"] }),
      qc.invalidateQueries({ queryKey: ["class-adviser-dashboard-students"] }),
    ]);
  };

  const updateStudent = useMutation({
    mutationFn: async () => {
      if (!selectedStudent) throw new Error("Please select a learner.");

      const lastName = form.last_name.trim();
      const firstName = form.first_name.trim();
      const middleName = form.middle_name.trim();

      if (!lastName || !firstName) {
        throw new Error("Last name and first name are required.");
      }

      const { error } = await supabase
        .from("students")
        .update({
          last_name: lastName,
          first_name: firstName,
          middle_name: middleName || null,
          sex: form.sex,
        })
        .eq("id", selectedStudent.id);

      if (error) throw error;
    },
    onSuccess: async () => {
      toast.success("Learner updated");
      await refreshStudents();
      onClose();
    },
    onError: (e: Error) => toast.error(e.message),
  });

  return (
    <Dialog open={open} onOpenChange={(nextOpen) => !nextOpen && onClose()}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Edit Student Name / Sex</DialogTitle>
        </DialogHeader>

        {students.length === 0 ? (
          <div className="rounded-xl border border-dashed p-6 text-center text-sm text-muted-foreground">
            No learners yet.
          </div>
        ) : (
          <div className="grid gap-3">
            <div>
              <Label>Select student</Label>
              <Select value={selectedId || undefined} onValueChange={setSelectedId}>
                <SelectTrigger>
                  <SelectValue placeholder="Select student" />
                </SelectTrigger>
                <SelectContent>
                  {students.map((student) => (
                    <SelectItem key={student.id} value={student.id}>
                      {student.last_name}, {student.first_name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="grid gap-3 sm:grid-cols-2">
              <div>
                <Label>Last name</Label>
                <Input
                  value={form.last_name}
                  onChange={(e) => setForm({ ...form, last_name: e.target.value })}
                />
              </div>
              <div>
                <Label>First name</Label>
                <Input
                  value={form.first_name}
                  onChange={(e) => setForm({ ...form, first_name: e.target.value })}
                />
              </div>
              <div>
                <Label>Middle name</Label>
                <Input
                  value={form.middle_name}
                  onChange={(e) => setForm({ ...form, middle_name: e.target.value })}
                />
              </div>
              <div>
                <Label>Sex</Label>
                <Select
                  value={form.sex}
                  onValueChange={(value) =>
                    setForm({
                      ...form,
                      sex: value === "female" ? "female" : "male",
                    })
                  }
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Select sex" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="male">Male</SelectItem>
                    <SelectItem value="female">Female</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>
          </div>
        )}

        <DialogFooter>
          <Button variant="outline" onClick={onClose} disabled={updateStudent.isPending}>
            Cancel
          </Button>
          <Button
            onClick={() => updateStudent.mutate()}
            disabled={
              updateStudent.isPending ||
              students.length === 0 ||
              !form.last_name.trim() ||
              !form.first_name.trim()
            }
          >
            {updateStudent.isPending ? (
              <Loader2 className="mr-2 size-4 animate-spin" />
            ) : (
              <Save className="mr-2 size-4" />
            )}
            Save Changes
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function DeleteStudentsDialog({
  open,
  classId,
  students,
  onClose,
}: {
  open: boolean;
  classId: string;
  students: StudentRow[];
  onClose: () => void;
}) {
  const qc = useQueryClient();
  const [studentToDelete, setStudentToDelete] = useState<StudentRow | null>(null);

  const refreshStudents = async () => {
    await Promise.all([
      qc.invalidateQueries({ queryKey: ["students", classId] }),
      qc.invalidateQueries({ queryKey: ["scores", classId] }),
      qc.invalidateQueries({ queryKey: ["scores-all", classId] }),
      qc.invalidateQueries({ queryKey: ["scores-all-a", classId] }),
      qc.invalidateQueries({ queryKey: ["students-count"] }),
      qc.invalidateQueries({ queryKey: ["class-adviser-dashboard-students"] }),
    ]);
  };

  const remove = useMutation({
    mutationFn: async (student: StudentRow) => {
      await supabase.from("activity_scores").delete().eq("student_id", student.id);

      await supabase.from("grades").delete().eq("student_id", student.id).eq("class_id", classId);

      const { error } = await supabase.from("students").delete().eq("id", student.id);

      if (error) throw error;
    },
    onSuccess: async () => {
      toast.success("Learner deleted");
      setStudentToDelete(null);
      await refreshStudents();
    },
    onError: (e: Error) => toast.error(e.message),
  });

  const deleteStudent = (student: StudentRow) => {
    setStudentToDelete(student);
  };

  const deleteStudentName = studentToDelete
    ? `${studentToDelete.last_name || ""}, ${studentToDelete.first_name || ""}`.trim()
    : "";

  return (
    <Dialog open={open} onOpenChange={(nextOpen) => !nextOpen && onClose()}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Delete Students</DialogTitle>
        </DialogHeader>

        <div className="max-h-[420px] space-y-2 overflow-y-auto pr-1">
          {students.map((student) => (
            <div
              key={student.id}
              className="flex items-center justify-between gap-3 rounded-xl border p-3"
            >
              <div>
                <div className="font-medium">
                  {student.last_name}, {student.first_name} {student.middle_name ?? ""}
                </div>
                <div className="text-xs capitalize text-muted-foreground">
                  {student.sex || "No sex"} {student.lrn ? `· LRN ${student.lrn}` : ""}
                </div>
              </div>

              <Button
                type="button"
                variant="outline"
                size="sm"
                disabled={remove.isPending}
                onClick={() => deleteStudent(student)}
                className="border-red-200 text-red-700 hover:bg-red-50"
              >
                {remove.isPending ? (
                  <Loader2 className="mr-2 size-4 animate-spin" />
                ) : (
                  <Trash2 className="mr-2 size-4" />
                )}
                Delete
              </Button>
            </div>
          ))}

          {students.length === 0 && (
            <div className="rounded-xl border border-dashed p-6 text-center text-sm text-muted-foreground">
              No learners yet.
            </div>
          )}
        </div>

        {studentToDelete && (
          <div className="rounded-2xl border border-red-200 bg-red-50 p-4">
            <div className="flex items-start gap-3">
              <div className="grid size-10 shrink-0 place-items-center rounded-full bg-red-100 text-red-700">
                <Trash2 className="size-5" />
              </div>

              <div className="min-w-0 flex-1">
                <h3 className="font-semibold text-red-900">Delete this student?</h3>
                <p className="mt-1 text-sm text-red-700">
                  Are you sure you want to delete{" "}
                  <span className="font-semibold">{deleteStudentName || "this learner"}</span>? This
                  will remove the student from the class list and E-Class Record.
                </p>

                <div className="mt-4 flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
                  <Button
                    type="button"
                    variant="outline"
                    disabled={remove.isPending}
                    onClick={() => setStudentToDelete(null)}
                  >
                    Cancel
                  </Button>

                  <Button
                    type="button"
                    disabled={remove.isPending}
                    onClick={() => remove.mutate(studentToDelete)}
                    className="bg-red-700 text-white hover:bg-red-800"
                  >
                    {remove.isPending ? (
                      <Loader2 className="mr-2 size-4 animate-spin" />
                    ) : (
                      <Trash2 className="mr-2 size-4" />
                    )}
                    Delete Student
                  </Button>
                </div>
              </div>
            </div>
          </div>
        )}

        <DialogFooter>
          <Button
            variant="outline"
            onClick={() => {
              setStudentToDelete(null);
              onClose();
            }}
          >
            Close
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

/* ---------------- MAPEH SUMMATIVE HEADER ---------------- */
// MAPEH summative table uses ST1, ST2, TE, WS ST1(30%), WS ST2(30%), WS TE(40%), PS ST&TE, WS ST&TE

/* ---------------- MAPEH Grades Panel ---------------- */
function MapehGradesPanel({
  classId,
  students,
  subject,
  klass,
  principal,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  klass: ClassRow;
  principal: string;
}) {
  const [component, setComponent] = useState("MA_T1");
  const [assessment] = useState("ALL");

  const components = ["MA_T1", "PEH_T1", "MA_T2", "PEH_T2", "MA_T3", "PEH_T3", "SUMMARY OF GRADES"];

  // Older MAPEH records were stored under term "1", so keep that scope for
  // MA_T1 to preserve existing scores. Every other MAPEH tab gets its own
  // activity scope so switching tabs cannot reuse the same activity IDs and
  // scores.
  const activityTerm = component === "MA_T1" ? "1" : component;

  const isSummary = component === "SUMMARY OF GRADES";

  return (
    <div className="space-y-3">
      <GradesPanel
        key={component}
        classId={classId}
        students={students}
        subject={isSummary ? subject : `${subject}_${component}_${assessment}`}
        klass={klass}
        principal={principal}
        hideTermSelector={true}
        activityTerm={isSummary ? undefined : activityTerm}
        mapehSummary={isSummary}
        headerContent={components.map((item) => (
          <button
            key={item}
            onClick={() => setComponent(item)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${
              component === item ? "bg-primary text-primary-foreground" : "border hover:bg-muted"
            }`}
          >
            {item}
          </button>
        ))}
      />
    </div>
  );
}

/* ---------------- Grade 11 Communication Composite Panel ---------------- */
// The Grade 11 Mabisang Komunikasyon class contains two related subjects:
// Effective Communication and Mabisang Komunikasyon. Each subject keeps its
// own Term 1, Term 2, and Term 3 class-record data, while SUMMARY OF GRADES
// averages the two subject grades per term (the same composite pattern used
// by MAPEH, but with different subject names).
function CommunicationGradesPanel({
  classId,
  students,
  subject,
  klass,
  principal,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  klass: ClassRow;
  principal: string;
}) {
  const communicationTabs = [
    { key: "EC_T1", label: "Effective Communication T1" },
    { key: "MK_T1", label: "Mabisang Komunikasyon T1" },
    { key: "EC_T2", label: "Effective Communication T2" },
    { key: "MK_T2", label: "Mabisang Komunikasyon T2" },
    { key: "EC_T3", label: "Effective Communication T3" },
    { key: "MK_T3", label: "Mabisang Komunikasyon T3" },
    { key: "SUMMARY OF GRADES", label: "SUMMARY OF GRADES" },
  ] as const;

  const [component, setComponent] = useState<string>("EC_T1");
  const isSummary = component === "SUMMARY OF GRADES";

  return (
    <div className="space-y-3">
      <GradesPanel
        key={component}
        classId={classId}
        students={students}
        subject={isSummary ? subject : `${subject}_${component}_ALL`}
        klass={klass}
        principal={principal}
        hideTermSelector={true}
        activityTerm={isSummary ? undefined : component}
        communicationSummary={isSummary}
        hideExcelExport={true}
        headerContent={communicationTabs.map((item) => (
          <button
            key={item.key}
            onClick={() => setComponent(item.key)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${
              component === item.key
                ? "bg-primary text-primary-foreground"
                : "border hover:bg-muted"
            }`}
          >
            {item.label}
          </button>
        ))}
      />
    </div>
  );
}

/* ---------------- E-Class Record: Grades ---------------- */
const TERM_TABS = [
  { value: "1", label: "1st Term" },
  { value: "2", label: "2nd Term" },
  { value: "3", label: "3rd Term" },
  { value: "final", label: "Final Grade" },
] as const;

type Grade12SingleTerm = "1" | "2" | "3";

function normalizeGrade12Subject(value: string | null | undefined) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Grade 12 SF9 one-term subject mapping.
// Gray cells in the Grade 12 SF9 template are unavailable terms.
// Only the one non-gray term is exposed in the E-Class Record,
// together with Final Grade.
//
// PE and Health 3 / PE and Health 4 are intentionally excluded so
// their current E-Class Record UI remains unchanged.
const GRADE12_SINGLE_TERM_SUBJECTS: Record<string, Grade12SingleTerm> = {
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

function getGrade12SingleTerm(
  gradeLevel: string | null | undefined,
  subject: string | null | undefined,
): Grade12SingleTerm | null {
  if (String(gradeLevel ?? "").trim().toLowerCase() !== "grade 12") {
    return null;
  }

  return (
    GRADE12_SINGLE_TERM_SUBJECTS[normalizeGrade12Subject(subject)] ?? null
  );
}

function getGrade12SingleTermTabs(term: Grade12SingleTerm) {
  return TERM_TABS.filter(
    (tab) => tab.value === term || tab.value === "final",
  );
}

const GRADE11_ELECTIVE_TABS = [
  { value: "1", label: "Academic Elective 1" },
  { value: "2", label: "Academic Elective 2" },
  { value: "3", label: "Academic Elective 3" },
] as const;

function GradesPanel({
  classId,
  students,
  subject,
  klass,
  principal,
  hideTermSelector = false,
  activityTerm,
  headerContent,
  mapehSummary = false,
  communicationSummary = false,
  hideExcelExport = false,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  klass: ClassRow;
  principal: string;
  hideTermSelector?: boolean;
  activityTerm?: string;
  headerContent?: ReactNode;
  mapehSummary?: boolean;
  communicationSummary?: boolean;
  hideExcelExport?: boolean;
}) {
  const [term, setTerm] = useState<string>("1");
  const [exporting, setExporting] = useState(false);

  const isGrade11ElectiveClass = isGrade11ElectiveSubject(
    klass?.grade_level,
    subject,
  );

  const grade12SingleTerm = getGrade12SingleTerm(
    klass?.grade_level,
    subject,
  );

  const visibleTermTabs = isGrade11ElectiveClass
    ? GRADE11_ELECTIVE_TABS
    : grade12SingleTerm
      ? getGrade12SingleTermTabs(grade12SingleTerm)
      : TERM_TABS;

  useEffect(() => {
    // Academic Elective 1/2/3 are three different one-term electives.
    // Do not expose or calculate a combined Final Grade for this special class.
    if (isGrade11ElectiveClass && term === "final") {
      setTerm("1");
      return;
    }

    // Mapped Grade 12 subjects expose only their one non-gray SF9 term
    // plus Final Grade. Move automatically to the allowed term if needed.
    if (
      grade12SingleTerm &&
      term !== grade12SingleTerm &&
      term !== "final"
    ) {
      setTerm(grade12SingleTerm);
    }
  }, [isGrade11ElectiveClass, grade12SingleTerm, term]);

  const exportToExcel = async () => {
    if (exporting) return;
    setExporting(true);

    try {
      if (hideTermSelector) {
        const truncated = await exportMapehClassRecord({
          classId,
          students,
          klass,
          principal,
        });

        if (truncated) {
          toast.warning(
            "The MAPEH template supports up to 50 male and 50 female learners. Extra learners were not exported.",
          );
        } else {
          toast.success("MAPEH class record exported to Excel.");
        }
        return;
      }

      const templateResponse = await fetch("/templates/summary-of-grades.xlsx");
      if (!templateResponse.ok) {
        throw new Error("Template not found. Put summary-of-grades.xlsx in public/templates.");
      }

      const [
        { data: components, error: componentsError },
        { data: activities, error: activitiesError },
      ] = await Promise.all([
        supabase.from("grade_components").select("*").eq("class_id", classId),
        supabase.from("grade_activities").select("*").eq("class_id", classId).order("position"),
      ]);

      if (componentsError) throw componentsError;
      if (activitiesError) throw activitiesError;

      const allComponents = (components ?? []) as GradeComponent[];
      const allActivities = (activities ?? []) as GradeActivity[];
      const activityIds = allActivities.map((activity) => activity.id);
      let allScores: ActivityScore[] = [];

      if (activityIds.length) {
        const { data: scores, error: scoresError } = await supabase
          .from("activity_scores")
          .select("*")
          .in("activity_id", activityIds);
        if (scoresError) throw scoresError;
        allScores = (scores ?? []) as ActivityScore[];
      }

      const zip = await JSZip.loadAsync(await templateResponse.arrayBuffer());
      const parser = new DOMParser();
      const serializer = new XMLSerializer();
      const workbookFile = zip.file("xl/workbook.xml");
      const relationshipsFile = zip.file("xl/_rels/workbook.xml.rels");

      if (!workbookFile || !relationshipsFile) {
        throw new Error("The Excel template has an invalid workbook structure.");
      }

      const workbookXml = parser.parseFromString(
        await workbookFile.async("text"),
        "application/xml",
      );
      const relationshipsXml = parser.parseFromString(
        await relationshipsFile.async("text"),
        "application/xml",
      );
      const relationshipTargets = new Map<string, string>();

      Array.from(relationshipsXml.getElementsByTagName("Relationship")).forEach((relationship) => {
        relationshipTargets.set(
          relationship.getAttribute("Id") ?? "",
          relationship.getAttribute("Target") ?? "",
        );
      });

      const sheetPaths = new Map<string, string>();
      Array.from(workbookXml.getElementsByTagName("sheet")).forEach((sheet) => {
        const name = sheet.getAttribute("name");
        const relationshipId =
          sheet.getAttribute("r:id") ??
          sheet.getAttributeNS(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "id",
          );
        const target = relationshipId ? relationshipTargets.get(relationshipId) : undefined;

        if (name && target) {
          const normalizedTarget = target.replace(/^\/?xl\//, "");
          sheetPaths.set(name, `xl/${normalizedTarget}`);
        }
      });

      const requiredSheets = ["INPUT DATA", "TERM1", "TERM2", "TERM3", "SUMMARY OF GRADES"];
      if (requiredSheets.some((name) => !sheetPaths.has(name))) {
        throw new Error(
          "The Excel template is missing INPUT DATA, TERM1, TERM2, TERM3, or SUMMARY OF GRADES.",
        );
      }

      const pendingValues = new Map<string, Map<string, string | number>>();
      const pendingFormulas = new Map<string, Map<string, string>>();
      const pendingFormulaCaches = new Map<
        string,
        Map<string, string | number>
      >();
      const setCellValue = (
        sheetName: string,
        address: string,
        value: string | number | null | undefined,
      ) => {
        const sheetValues = pendingValues.get(sheetName) ?? new Map<string, string | number>();
        sheetValues.set(address, value ?? "");
        pendingValues.set(sheetName, sheetValues);
      };
      const setCellFormula = (
        sheetName: string,
        address: string,
        formula: string,
        cachedValue?: string | number,
      ) => {
        const sheetFormulas = pendingFormulas.get(sheetName) ?? new Map<string, string>();
        sheetFormulas.set(address, formula);
        pendingFormulas.set(sheetName, sheetFormulas);

        // Excel Protected View can display cached values before recalculating.
        if (cachedValue !== undefined) {
          const cachedCells = pendingFormulaCaches.get(sheetName) ??
            new Map<string, string | number>();
          cachedCells.set(address, cachedValue);
          pendingFormulaCaches.set(sheetName, cachedCells);
        }
      };

      const fullName = (student: StudentRow) => {
        const givenNames = [student.first_name, student.middle_name].filter(Boolean).join(" ");
        return [student.last_name, givenNames].filter(Boolean).join(", ");
      };

      const maleStudents = students
        .filter((student) => normalize(student.sex) === "male")
        .slice(0, 50);
      const femaleStudents = students
        .filter((student) => normalize(student.sex) === "female")
        .slice(0, 50);

      setCellValue("INPUT DATA", "E8", klass.region);
      setCellValue("INPUT DATA", "E9", klass.division);
      setCellValue("INPUT DATA", "E10", klass.school_name);
      setCellValue("INPUT DATA", "E11", klass.school_id);
      setCellValue("INPUT DATA", "E12", klass.school_year);
      setCellValue(
        "INPUT DATA",
        "E14",
        [klass.grade_level, klass.section].filter(Boolean).join(" - "),
      );
      setCellValue("INPUT DATA", "E15", subject);
      setCellValue("INPUT DATA", "E17", klass.teacher_name);
      setCellValue("INPUT DATA", "E20", principal);

      for (let index = 0; index < 50; index += 1) {
        const maleName = maleStudents[index] ? fullName(maleStudents[index]) : "";
        const femaleName = femaleStudents[index] ? fullName(femaleStudents[index]) : "";

        setCellValue("INPUT DATA", `J${index + 9}`, maleName);
        setCellValue("INPUT DATA", `P${index + 9}`, femaleName);

        ["TERM1", "TERM2", "TERM3"].forEach((sheetName) => {
          setCellValue(sheetName, `B${index + 12}`, maleName);
          setCellValue(sheetName, `B${index + 63}`, femaleName);
        });

        setCellValue("SUMMARY OF GRADES", `B${index + 13}`, maleName);
        setCellValue("SUMMARY OF GRADES", `B${index + 64}`, femaleName);
      }

      const scoreMap = new Map<string, number | null>();
      allScores.forEach((score) =>
        scoreMap.set(`${score.activity_id}|${score.student_id}`, score.score),
      );

      // JHS only (Grades 7-10): the downloaded TERM1/2/3 Term Grade must
      // match the E-Class Record, including any teacher-edited Base/Final.
      // Other grade levels and the separate MAPEH export are unchanged.
      const isJuniorHighExcelExport = /^(?:grade\s*)?(?:7|8|9|10)$/i.test(
        String(klass.grade_level ?? "").trim(),
      );
      const savedJhsBaseMap = new Map<string, number>();

      if (isJuniorHighExcelExport) {
        const { data: savedGrades, error: savedGradesError } = await (supabase as any)
          .from("grades")
          .select("student_id, term, term_grade_base")
          .eq("class_id", classId)
          .eq("subject", subject)
          .in("term", ["1", "2", "3"]);

        if (savedGradesError) throw savedGradesError;
        (savedGrades ?? []).forEach((row: {
          student_id: string;
          term: string;
          term_grade_base: number | null;
        }) => {
          if (row.term_grade_base == null) return;
          const savedBase = Number(row.term_grade_base);
          if (Number.isFinite(savedBase)) {
            savedJhsBaseMap.set(`${row.student_id}|${row.term}`, savedBase);
          }
        });
      }

      // Mirror TermGradesTable.computeRow(): include every configured WW/PT
      // activity, the first three QA activities, the actual component weights,
      // clamped scores, and the same per-component rounding. Do not rely on
      // the template's different ST1/ST2/TE transmutation formula.
      const jhsSystemTermGrade = (
        studentId: string,
        termValue: string,
      ): number | null => {
        const weightedScores = (["WW", "PT", "QA"] as const).map((component) => {
          const componentActivities = allActivities
            .filter((activity) => activity.term === termValue && activity.component === component)
            .sort((left, right) => left.position - right.position)
            .slice(0, component === "QA" ? 3 : undefined);

          let raw = 0;
          let hps = 0;
          let hasScore = false;
          componentActivities.forEach((activity) => {
            const storedScore = scoreMap.get(`${activity.id}|${studentId}`);
            if (typeof storedScore !== "number") return;
            const activityHps = Math.max(0, Number(activity.hps) || 0);
            raw += Math.max(0, Math.min(activityHps, storedScore));
            hps += Number(activity.hps);
            hasScore = true;
          });
          if (!hasScore || hps <= 0) return null;

          const percentageScore = Math.round(((raw / hps) * 100) * 100) / 100;
          const weight = allComponents.find(
            (item) => item.term === termValue && item.component === component,
          )?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);
          return Math.round(percentageScore * (weight / 100) * 100) / 100;
        });

        if (weightedScores.some((value) => value == null)) return null;
        const initialGrade = Math.round(
          (weightedScores as number[]).reduce((sum, value) => sum + value, 0) * 100,
        ) / 100;
        const computedBase = roundInitialGrade(initialGrade);
        return resolveTermGradeBase(
          computedBase,
          savedJhsBaseMap.get(`${studentId}|${termValue}`),
        );
      };

      const componentColumns: Record<"WW" | "PT" | "QA", { columns: string[]; limit: number }> = {
        WW: { columns: ["F", "G", "H", "I", "J"], limit: 5 },
        PT: { columns: ["N", "O", "P"], limit: 3 },
        QA: { columns: ["T", "U", "V"], limit: 3 },
      };

      ["TERM1", "TERM2", "TERM3"].forEach((sheetName, termIndex) => {
        const termValue = String(termIndex + 1);

        if (isJuniorHighExcelExport) {
          // Write the exact Term Grade currently shown in the JHS UI for
          // each learner. This is an export snapshot, not an Excel formula:
          // later edits to Excel scores require exporting again from SIGLA.
          // AD descriptors and Summary of Grades still reference AC.
          maleStudents.forEach((student, studentIndex) => {
            setCellValue(
              sheetName,
              `AC${studentIndex + 12}`,
              jhsSystemTermGrade(student.id, termValue),
            );
          });
          femaleStudents.forEach((student, studentIndex) => {
            setCellValue(
              sheetName,
              `AC${studentIndex + 63}`,
              jhsSystemTermGrade(student.id, termValue),
            );
          });
          // Clear template formulas for unused learner rows too, avoiding
          // transmuted values appearing in otherwise empty JHS rows.
          for (let index = maleStudents.length; index < 50; index += 1) {
            setCellValue(sheetName, `AC${index + 12}`, "");
          }
          for (let index = femaleStudents.length; index < 50; index += 1) {
            setCellValue(sheetName, `AC${index + 63}`, "");
          }
        } else {
          // SHS and any other grade levels retain the existing template rule.
          [
            ...Array.from({ length: 50 }, (_, index) => index + 12),
            ...Array.from({ length: 50 }, (_, index) => index + 63),
          ].forEach((row) => {
            setCellFormula(
              sheetName,
              `AC${row}`,
              `IF(AB${row}="","",VLOOKUP(AB${row},'Helper (Do Not Delete)'!$H$3:$K$43,4,TRUE))`,
            );
          });
        }

        (["WW", "PT", "QA"] as const).forEach((component) => {
          const config = componentColumns[component];
          const termActivities = allActivities
            .filter((activity) => activity.term === termValue && activity.component === component)
            .sort((a, b) => a.position - b.position)
            .slice(0, config.limit);

          const componentWeight =
            allComponents.find((item) => item.term === termValue && item.component === component)
              ?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);

          const weightCell = component === "WW" ? "M10" : component === "PT" ? "S10" : "AA10";
          setCellValue(sheetName, weightCell, componentWeight / 100);

          config.columns.forEach((column, activityIndex) => {
            const activity = termActivities[activityIndex];
            setCellValue(sheetName, `${column}10`, activity?.hps);

            maleStudents.forEach((student, studentIndex) => {
              setCellValue(
                sheetName,
                `${column}${studentIndex + 12}`,
                activity ? (scoreMap.get(`${activity.id}|${student.id}`) ?? "") : "",
              );
            });
            femaleStudents.forEach((student, studentIndex) => {
              setCellValue(
                sheetName,
                `${column}${studentIndex + 63}`,
                activity ? (scoreMap.get(`${activity.id}|${student.id}`) ?? "") : "",
              );
            });
          });
        });
      });

      if (isJuniorHighExcelExport) {
        const achievementLevels = [
          "Advancing", "Benchmarking", "Connecting", "Developing", "Emerging",
        ] as const;
        const maleAchievement = new Map<string, number>();
        const femaleAchievement = new Map<string, number>();

        const writeJhsSummaryRow = (
          student: StudentRow | undefined,
          summaryRow: number,
          detailRow: number,
          genderStats: Map<string, number>,
        ) => {
          const grades = (["1", "2", "3"] as const).map((termValue) =>
            student ? jhsSystemTermGrade(student.id, termValue) : null,
          );
          const isComplete = student !== undefined &&
            grades.every((grade): grade is number => typeof grade === "number");
          const final = isComplete
            ? Math.round((grades[0]! + grades[1]! + grades[2]!) / 3)
            : null;
          const descriptor = final == null ? "" : descriptorFor(final).label;
          const remark = final == null ? "" : final >= 75 ? "PASSED" : "FAILED";
          if (descriptor) {
            genderStats.set(descriptor, (genderStats.get(descriptor) ?? 0) + 1);
          }

          (["F", "J", "N"] as const).forEach((column, index) => {
            const termSheet = `TERM${index + 1}`;
            setCellFormula(
              "SUMMARY OF GRADES",
              `${column}${summaryRow}`,
              `IF(OR($B${summaryRow}="",NOT(ISNUMBER(${termSheet}!AC${detailRow}))),"",${termSheet}!AC${detailRow})`,
              grades[index] ?? "",
            );
          });
          setCellFormula(
            "SUMMARY OF GRADES",
            `R${summaryRow}`,
            `IF(OR($B${summaryRow}="",COUNT(F${summaryRow},J${summaryRow},N${summaryRow})<3),"",ROUND(AVERAGE(F${summaryRow},J${summaryRow},N${summaryRow}),0))`,
            final ?? "",
          );
          setCellFormula(
            "SUMMARY OF GRADES",
            `V${summaryRow}`,
            `IF(OR($B${summaryRow}="",NOT(ISNUMBER(R${summaryRow}))),"",IFERROR(VLOOKUP(R${summaryRow},DESCRIPTORS,4,TRUE),""))`,
            descriptor,
          );
          setCellFormula(
            "SUMMARY OF GRADES",
            `Z${summaryRow}`,
            `IF(OR($B${summaryRow}="",NOT(ISNUMBER(R${summaryRow}))),"",IF(R${summaryRow}>=75,"PASSED","FAILED"))`,
            remark,
          );
        };

        for (let index = 0; index < 50; index += 1) {
          writeJhsSummaryRow(
            maleStudents[index], index + 13, index + 12, maleAchievement,
          );
          writeJhsSummaryRow(
            femaleStudents[index], index + 64, index + 63, femaleAchievement,
          );
        }

        // Count only completed grades belonging to real learners. Fix the
        // original template's off-by-one achievement COUNTIF ranges too.
        achievementLevels.forEach((level, index) => {
          const targetRow = index + 117;
          const maleCount = maleAchievement.get(level) ?? 0;
          const femaleCount = femaleAchievement.get(level) ?? 0;
          setCellFormula(
            "SUMMARY OF GRADES", `F${targetRow}`,
            `COUNTIF($V$13:$V$62,"${level}")`, maleCount,
          );
          setCellFormula(
            "SUMMARY OF GRADES", `J${targetRow}`,
            `COUNTIF($V$64:$V$113,"${level}")`, femaleCount,
          );
          setCellFormula(
            "SUMMARY OF GRADES", `M${targetRow}`,
            `F${targetRow}+J${targetRow}`, maleCount + femaleCount,
          );
        });
        const maleTotal = Array.from(maleAchievement.values()).reduce((a, b) => a + b, 0);
        const femaleTotal = Array.from(femaleAchievement.values()).reduce((a, b) => a + b, 0);
        setCellFormula("SUMMARY OF GRADES", "F122", "SUM(F117:F121)", maleTotal);
        setCellFormula("SUMMARY OF GRADES", "J122", "SUM(J117:J121)", femaleTotal);
        setCellFormula(
          "SUMMARY OF GRADES", "M122", "F122+J122", maleTotal + femaleTotal,
        );
      }

      for (const sheetName of requiredSheets) {
        const values = pendingValues.get(sheetName) ?? new Map();
        const formulas = pendingFormulas.get(sheetName) ?? new Map();
        const sheetPath = sheetPaths.get(sheetName);
        if (!sheetPath) {
          throw new Error(`Unable to locate the ${sheetName} worksheet.`);
        }
        const sheetFile = zip.file(sheetPath);
        if (!sheetFile) {
          throw new Error(`Unable to read the ${sheetName} worksheet.`);
        }

        const sheetXml = parser.parseFromString(await sheetFile.async("text"), "application/xml");

        for (const [address, formula] of formulas) {
          const cell = Array.from(sheetXml.getElementsByTagName("c")).find(
            (item) => item.getAttribute("r") === address,
          );
          if (!cell) {
            throw new Error(`The template cell ${sheetName}!${address} is missing.`);
          }

          // Replace every Term Grade formula in the original shared-formula
          // range with a normal standalone formula. Updating the whole range
          // avoids leaving broken shared-formula references in the workbook.
          Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
          cell.removeAttribute("t");

          const formulaNode = sheetXml.createElementNS(
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "f",
          );
          formulaNode.textContent = formula;
          cell.appendChild(formulaNode);

          const cachedValue = pendingFormulaCaches.get(sheetName)?.get(address);
          if (cachedValue !== undefined) {
            if (typeof cachedValue === "string") {
              cell.setAttribute("t", "str");
            }
            const valueNode = sheetXml.createElementNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "v",
            );
            valueNode.textContent = String(cachedValue);
            cell.appendChild(valueNode);
          }
        }

        for (const [address, value] of values) {
          const cell = Array.from(sheetXml.getElementsByTagName("c")).find(
            (item) => item.getAttribute("r") === address,
          );
          if (!cell) {
            throw new Error(`The template cell ${sheetName}!${address} is missing.`);
          }

          // The learner-name cells in the Term and Summary sheets contain
          // shared formulas that point to INPUT DATA. Removing their <f>
          // nodes breaks the shared-formula group and makes Excel repair the
          // downloaded workbook. Keep every existing formula and update only
          // its cached displayed result.
          const formulaNode = Array.from(cell.childNodes).find((child) => child.nodeName === "f");
          // For JHS AC Term Grade only, discard the old shared VLOOKUP
          // formula so Excel uses the actual system grade written above.
          const replaceJhsTermGradeFormula =
            isJuniorHighExcelExport &&
            (sheetName === "TERM1" || sheetName === "TERM2" || sheetName === "TERM3") &&
            /^AC\d+$/.test(address);
          if (formulaNode && !replaceJhsTermGradeFormula) {
            Array.from(cell.childNodes)
              .filter((child) => child.nodeName === "v")
              .forEach((child) => cell.removeChild(child));

            cell.setAttribute("t", "str");
            const cachedValueNode = sheetXml.createElementNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "v",
            );
            cachedValueNode.textContent = String(value ?? "");
            cell.appendChild(cachedValueNode);
            continue;
          }

          Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
          cell.removeAttribute("t");

          if (value === "") continue;

          if (typeof value === "number" && Number.isFinite(value)) {
            const numberNode = sheetXml.createElementNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "v",
            );
            numberNode.textContent = String(value);
            cell.appendChild(numberNode);
          } else {
            cell.setAttribute("t", "inlineStr");
            const inlineString = sheetXml.createElementNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "is",
            );
            const textNode = sheetXml.createElementNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "t",
            );
            textNode.textContent = String(value);
            inlineString.appendChild(textNode);
            cell.appendChild(inlineString);
          }
        }

        if (isJuniorHighExcelExport && sheetName !== "INPUT DATA") {
          hideUnusedJhsExcelRows(
            sheetXml,
            sheetName === "SUMMARY OF GRADES" ? "summary" : "term",
            maleStudents.length,
            femaleStudents.length,
          );
        }
        zip.file(sheetPath, serializer.serializeToString(sheetXml));
      }

      if (isJuniorHighExcelExport) {
        zip.remove("xl/calcChain.xml");
        Array.from(relationshipsXml.getElementsByTagName("Relationship"))
          .filter((relationship) => {
            const target = relationship.getAttribute("Target") || "";
            const type = relationship.getAttribute("Type") || "";
            return target.endsWith("calcChain.xml") || type.endsWith("/calcChain");
          })
          .forEach((relationship) => relationship.parentNode?.removeChild(relationship));
        zip.file("xl/_rels/workbook.xml.rels", serializer.serializeToString(relationshipsXml));

        const contentTypesFile = zip.file("[Content_Types].xml");
        if (contentTypesFile) {
          const contentTypesXml = parser.parseFromString(
            await contentTypesFile.async("text"),
            "application/xml",
          );
          Array.from(contentTypesXml.getElementsByTagName("Override"))
            .filter((entry) => entry.getAttribute("PartName") === "/xl/calcChain.xml")
            .forEach((entry) => entry.parentNode?.removeChild(entry));
          zip.file("[Content_Types].xml", serializer.serializeToString(contentTypesXml));
        }
      }

      const calculationProperties = workbookXml.getElementsByTagName("calcPr")[0];
      if (calculationProperties) {
        calculationProperties.setAttribute("fullCalcOnLoad", "1");
        calculationProperties.setAttribute("forceFullCalc", "1");
      }
      zip.file("xl/workbook.xml", serializer.serializeToString(workbookXml));

      const output = await zip.generateAsync({
        type: "arraybuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });
      const blob = new Blob([output], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      const safeClassName = [klass.grade_level, klass.section, subject]
        .filter(Boolean)
        .join("-")
        .replace(/[^a-z0-9-_]+/gi, "-")
        .replace(/^-+|-+$/g, "");
      anchor.href = url;
      anchor.download = `${safeClassName || "summary-of-grades"}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);

      if (students.length > 100) {
        toast.warning(
          "The template supports up to 50 male and 50 female learners. Only the first 100 learners were exported.",
        );
      } else {
        toast.success("Class grades exported to Excel.");
      }
    } catch (error) {
      console.error("Unable to export class grades:", error);
      toast.error(
        error instanceof Error ? error.message : "Unable to export class grades to Excel.",
      );
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2 rounded-2xl border bg-card p-2">
        {headerContent}
        {!hideTermSelector &&
          visibleTermTabs.map((t) => (
            <button
              key={t.value}
              onClick={() => setTerm(t.value)}
              className={`rounded-lg px-4 py-1.5 text-sm font-semibold transition ${term === t.value ? "bg-primary text-primary-foreground" : "border hover:bg-muted"}`}
            >
              {t.label}
            </button>
          ))}
        <div className="ml-auto flex flex-wrap items-center gap-2">
          <span className="chip bg-[color:var(--success-bg)] text-[color:var(--success)]">
            ✓ Weights: 100%
          </span>
          <button
            onClick={() => window.print()}
            className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 text-sm hover:bg-muted"
          >
            <Printer className="size-4" />
            Print A4
          </button>
          {!hideExcelExport && (
            <button
              onClick={exportToExcel}
              disabled={exporting}
              className="inline-flex items-center gap-1 rounded-lg border border-[color:var(--success)] px-3 py-1.5 text-sm font-medium text-[color:var(--success)] hover:bg-[color:var(--success-bg)] disabled:cursor-not-allowed disabled:opacity-60"
            >
              {exporting ? (
                <Loader2 className="size-4 animate-spin" />
              ) : (
                <Download className="size-4" />
              )}
              {exporting ? "Exporting..." : "Export to Excel"}
            </button>
          )}
        </div>
      </div>

      {mapehSummary ? (
        <MapehSummaryGradesTable classId={classId} students={students} subject={subject} />
      ) : communicationSummary ? (
        <CommunicationSummaryGradesTable
          classId={classId}
          students={students}
          subject={subject}
        />
      ) : term === "final" ? (
        <FinalGradesTable
          classId={classId}
          students={students}
          subject={subject}
          klass={klass}
        />
      ) : (
        <TermGradesTable
          classId={classId}
          students={students}
          term={term}
          subject={subject}
          activityTerm={activityTerm}
        />
      )}
    </div>
  );
}

function TermGradesTable({
  classId,
  students,
  term,
  subject,
  activityTerm,
}: {
  classId: string;
  students: StudentRow[];
  term: string;
  subject: string;
  activityTerm?: string;
}) {
  const qc = useQueryClient();
  const storageTerm = activityTerm ?? term;
  const isMapehSubject = normalize(subject).startsWith("mapeh");
  const { data: components = [] } = useQuery({
    queryKey: ["components", classId, storageTerm],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_components")
        .select("*")
        .eq("class_id", classId)
        .eq("term", storageTerm);
      if (error) throw error;
      return data as GradeComponent[];
    },
  });
  const { data: activities = [] } = useQuery({
    queryKey: ["activities", classId, storageTerm],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_activities")
        .select("*")
        .eq("class_id", classId)
        .eq("term", storageTerm)
        .order("position");
      if (error) throw error;
      return data as GradeActivity[];
    },
  });
  const activityIds = activities.map((a) => a.id);
  const { data: scores = [] } = useQuery({
    queryKey: ["scores", classId, storageTerm, activityIds.join(",")],
    queryFn: async () => {
      if (!activityIds.length) return [];
      const { data, error } = await supabase
        .from("activity_scores")
        .select("*")
        .in("activity_id", activityIds);
      if (error) throw error;
      return data as ActivityScore[];
    },
  });

  const { data: termGradeRows = [] } = useQuery({
    queryKey: ["term-grade-bases", classId, subject, term],
    enabled: Boolean(subject),
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, score, term_grade_base")
        .eq("class_id", classId)
        .eq("subject", subject)
        .eq("term", term);
      if (error) throw error;
      return (data ?? []) as Array<{
        student_id: string;
        score: number | null;
        term_grade_base: number | null;
      }>;
    },
  });

  const termGradeBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    termGradeRows.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) map.set(row.student_id, value);
    });
    return map;
  }, [termGradeRows]);

  const resolvedTermGrade = (studentId: string, baseTermGrade: number | null): number | null => {
    return resolveTermGradeBase(baseTermGrade, termGradeBaseMap.get(studentId));
  };

  const compBy = (c: "WW" | "PT" | "QA") => components.find((x) => x.component === c);
  const weight = (c: "WW" | "PT" | "QA") =>
    compBy(c)?.weight ?? (c === "WW" ? 20 : c === "PT" ? 50 : 30);
  const actsFor = (c: "WW" | "PT" | "QA") =>
    activities
      .filter((a) => a.component === c)
      .sort((a, b) => a.position - b.position)
      .slice(0, c === "QA" ? 3 : undefined);

  const upsertComponent = useMutation({
    mutationFn: async ({
      component,
      weight,
    }: {
      component: "WW" | "PT" | "QA";
      weight: number;
    }) => {
      const teacher_id = await getUserId();
      const { error } = await supabase.from("grade_components").upsert(
        {
          class_id: classId,
          teacher_id,
          term: storageTerm,
          component,
          weight,
        },
        { onConflict: "class_id,term,component" },
      );
      if (error) throw error;
    },
    onSuccess: () =>
      qc.invalidateQueries({
        queryKey: ["components", classId, storageTerm],
      }),
  });

  const requiredSlotsBeingCreated = useRef<string | null>(null);
  useEffect(() => {
    const requiredSlotCounts = {
      WW: 5,
      PT: 3,
      QA: 3,
    } as const;
    const missingSlots = (
      Object.entries(requiredSlotCounts) as Array<["WW" | "PT" | "QA", number]>
    ).flatMap(([component, count]) => {
      const existingPositions = new Set(
        activities
          .filter((activity) => activity.component === component)
          .map((activity) => activity.position),
      );
      return Array.from({ length: count }, (_, index) => index + 1)
        .filter((position) => !existingPositions.has(position))
        .map((position) => ({ component, position }));
    });
    const creationKey = `${classId}|${storageTerm}|${missingSlots
      .map(({ component, position }) => `${component}${position}`)
      .join(",")}`;

    if (!missingSlots.length || requiredSlotsBeingCreated.current === creationKey) return;

    requiredSlotsBeingCreated.current = creationKey;
    void (async () => {
      try {
        const teacher_id = await getUserId();
        const { error } = await supabase.from("grade_activities").insert(
          missingSlots.map(({ component, position }) => ({
            class_id: classId,
            teacher_id,
            term: storageTerm,
            component,
            position,
            hps: 10,
          })),
        );
        if (error) throw error;
        await qc.invalidateQueries({
          queryKey: ["activities", classId, storageTerm],
        });
      } finally {
        requiredSlotsBeingCreated.current = null;
      }
    })();
  }, [activities, classId, qc, storageTerm]);

  const addActivity = useMutation({
    mutationFn: async (component: "WW" | "PT") => {
      const teacher_id = await getUserId();
      const position = (actsFor(component).at(-1)?.position ?? 0) + 1;
      const { error } = await supabase.from("grade_activities").insert({
        class_id: classId,
        teacher_id,
        term: storageTerm,
        component,
        position,
        hps: 10,
      });
      if (error) throw error;
    },
    onSuccess: () =>
      qc.invalidateQueries({
        queryKey: ["activities", classId, storageTerm],
      }),
  });

  const removeActivity = useMutation({
    mutationFn: async (component: "WW" | "PT") => {
      const minimum = component === "WW" ? 5 : 3;
      const componentActivities = actsFor(component);
      if (componentActivities.length <= minimum) return;

      const last = componentActivities.at(-1);
      if (!last) return;

      const { error } = await supabase.from("grade_activities").delete().eq("id", last.id);
      if (error) throw error;
    },
    onSuccess: () => {
      qc.invalidateQueries({
        queryKey: ["activities", classId, storageTerm],
      });
      qc.invalidateQueries({
        queryKey: ["scores", classId, storageTerm],
      });
    },
  });

  const updateHps = useMutation({
    mutationFn: async ({ id, hps }: { id: string; hps: number }) => {
      // Highest Possible Score may be zero or any positive whole number.
      // Zero is allowed when the teacher intentionally wants an activity
      // to have no points. There is also no artificial upper limit.
      const safeHps = Math.max(
        0,
        Math.round(Number.isFinite(hps) ? hps : 10),
      );

      const { error } = await supabase
        .from("grade_activities")
        .update({ hps: safeHps })
        .eq("id", id);

      if (error) throw error;

      // If the teacher lowers the Highest Possible Score, make sure existing
      // learner scores can never remain above the new maximum.
      const { error: highScoreError } = await supabase
        .from("activity_scores")
        .update({ score: safeHps })
        .eq("activity_id", id)
        .gt("score", safeHps);

      if (highScoreError) throw highScoreError;

      const { error: negativeScoreError } = await supabase
        .from("activity_scores")
        .update({ score: 0 })
        .eq("activity_id", id)
        .lt("score", 0);

      if (negativeScoreError) throw negativeScoreError;
    },
    onSuccess: async () => {
      await Promise.all([
        qc.invalidateQueries({
          queryKey: ["activities", classId, storageTerm],
        }),
        qc.invalidateQueries({
          queryKey: ["scores", classId, storageTerm],
        }),
        qc.invalidateQueries({ queryKey: ["scores-all", classId] }),
        qc.invalidateQueries({ queryKey: ["scores-all-a", classId] }),
      ]);
    },
  });

  const upsertScore = useMutation({
    mutationFn: async ({
      activity_id,
      student_id,
      score,
      hps,
    }: {
      activity_id: string;
      student_id: string;
      score: number | null;
      hps: number;
    }) => {
      const teacher_id = await getUserId();
      const safeHps = Math.max(0, Number(hps) || 0);
      const safeScore =
        score == null
          ? null
          : Math.max(0, Math.min(safeHps, Number.isFinite(score) ? score : 0));

      const { error } = await supabase
        .from("activity_scores")
        .upsert(
          {
            activity_id,
            student_id,
            class_id: classId,
            teacher_id,
            score: safeScore,
          },
          { onConflict: "activity_id,student_id" },
        );

      if (error) throw error;
    },
    onSuccess: () =>
      qc.invalidateQueries({
        queryKey: ["scores", classId, storageTerm],
      }),
  });

  const lastSyncedTerm = useRef<Map<string, number>>(new Map());
  const syncTermGrade = useMutation({
    mutationFn: async ({ student_id, score }: { student_id: string; score: number }) => {
      const teacher_id = await getUserId();
      const { error } = await (supabase as any)
        .from("grades")
        .upsert(
          { student_id, class_id: classId, teacher_id, subject, term, score },
          { onConflict: "student_id,subject,term" },
        );
      if (error) throw error;
    },
  });

  const saveTermGradeBase = useMutation({
    mutationFn: async ({
      student_id,
      baseTermGrade,
    }: {
      student_id: string;
      baseTermGrade: number;
    }) => {
      const teacher_id = await getUserId();
      const finalScore = Math.max(0, Math.min(100, Math.floor(baseTermGrade + 0.5)));

      const { error } = await (supabase as any).from("grades").upsert(
        {
          student_id,
          class_id: classId,
          teacher_id,
          subject,
          term,
          score: finalScore,
          term_grade_base: finalScore,
        },
        { onConflict: "student_id,subject,term" },
      );

      if (error) throw error;
    },
    onSuccess: async () => {
      await Promise.all([
        qc.invalidateQueries({
          queryKey: ["term-grade-bases", classId, subject, term],
        }),
        qc.invalidateQueries({ queryKey: ["grades", classId] }),
        qc.invalidateQueries({ queryKey: ["grades-all", classId] }),
      ]);
    },
    onError: (error: Error) => {
      toast.error(`Unable to save Term Grade Base: ${error.message}`);
    },
  });

  const activityHpsMap = useMemo(() => {
    const map = new Map<string, number>();
    activities.forEach((activity) => {
      map.set(activity.id, Math.max(0, Number(activity.hps) || 0));
    });
    return map;
  }, [activities]);

  const scoreMap = useMemo(() => {
    const m = new Map<string, number | null>();

    scores.forEach((s) => {
      const hps = activityHpsMap.get(s.activity_id);
      const rawScore = s.score;

      if (typeof rawScore !== "number" || hps == null) {
        m.set(`${s.activity_id}|${s.student_id}`, rawScore);
        return;
      }

      m.set(
        `${s.activity_id}|${s.student_id}`,
        Math.max(0, Math.min(hps, rawScore)),
      );
    });

    return m;
  }, [scores, activityHpsMap]);

  // Repair legacy invalid values already stored in the database, such as
  // a learner score of 321 when the Highest Possible Score is only 10.
  useEffect(() => {
    const invalidScores = scores
      .map((row) => {
        const hps = activityHpsMap.get(row.activity_id);
        if (hps == null || typeof row.score !== "number") return null;

        const safeScore = Math.max(0, Math.min(hps, row.score));
        if (safeScore === row.score) return null;

        return {
          activity_id: row.activity_id,
          student_id: row.student_id,
          class_id: classId,
          score: safeScore,
        };
      })
      .filter(Boolean) as Array<{
        activity_id: string;
        student_id: string;
        class_id: string;
        score: number;
      }>;

    if (invalidScores.length === 0) return;

    let cancelled = false;

    void (async () => {
      const teacher_id = await getUserId();

      const rows = invalidScores.map((row) => ({
        ...row,
        teacher_id,
      }));

      const { error } = await supabase
        .from("activity_scores")
        .upsert(rows, { onConflict: "activity_id,student_id" });

      if (!cancelled && !error) {
        await Promise.all([
          qc.invalidateQueries({
            queryKey: ["scores", classId, storageTerm],
          }),
          qc.invalidateQueries({ queryKey: ["scores-all", classId] }),
          qc.invalidateQueries({ queryKey: ["scores-all-a", classId] }),
        ]);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [scores, activityHpsMap, classId, storageTerm, qc]);

  const computeRow = (sid: string) => {
    const results: Record<
      "WW" | "PT" | "QA",
      { raw: number; hps: number; ps: number | null; ws: number | null }
    > = {
      WW: { raw: 0, hps: 0, ps: null, ws: null },
      PT: { raw: 0, hps: 0, ps: null, ws: null },
      QA: { raw: 0, hps: 0, ps: null, ws: null },
    };
    (["WW", "PT", "QA"] as const).forEach((c) => {
      const acts = actsFor(c);
      let raw = 0,
        hps = 0,
        any = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;
      acts.forEach((a) => {
        const v = scoreMap.get(`${a.id}|${sid}`);
        if (typeof v === "number") {
          raw += v;
          hps += a.hps;
          any = true;
          if (c === "QA" && isMapehSubject) {
            const activityIndex = acts.indexOf(a);
            const partWeight = qaPartWeights[activityIndex] ?? 0;
            if (a.hps > 0) {
              qaPercentageScore += (v / a.hps) * 100 * (partWeight / 100);
            }
          }
        }
      });
      const ps =
        any && hps > 0
          ? Math.round(
              (c === "QA" && isMapehSubject ? qaPercentageScore : (raw / hps) * 100) * 100,
            ) / 100
          : null;
      const ws = ps == null ? null : Math.round(ps * (weight(c) / 100) * 100) / 100;
      results[c] = { raw, hps, ps, ws };
    });
    const initial = (["WW", "PT", "QA"] as const).every((c) => results[c].ws != null)
      ? Math.round((results.WW.ws! + results.PT.ws! + results.QA.ws!) * 100) / 100
      : null;
    const term_grade = roundInitialGrade(initial);
    return { results, initial, term_grade };
  };
  useEffect(() => {
    if (!subject) return;
    students.forEach((s) => {
      const { term_grade } = computeRow(s.id);
      const finalTermGrade = resolvedTermGrade(s.id, term_grade);
      if (finalTermGrade == null) return;
      const prev = lastSyncedTerm.current.get(s.id);
      if (prev !== finalTermGrade) {
        lastSyncedTerm.current.set(s.id, finalTermGrade);
        syncTermGrade.mutate({ student_id: s.id, score: finalTermGrade });
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scores, activities, components, subject, students, termGradeRows]);

  const totalWeight = weight("WW") + weight("PT") + weight("QA");

  const wwActs = actsFor("WW"),
    ptActs = actsFor("PT"),
    qaActs = actsFor("QA");
  const male = students.filter((s) => normalize(s.sex) === "male");
  const female = students.filter((s) => normalize(s.sex) === "female");

  const componentHeader = (c: "WW" | "PT" | "QA", label: string, color: string) => {
    const acts = actsFor(c);
    return (
      <th
        colSpan={
          c === "QA" && isMapehSubject ? 8 : Math.max(1, acts.length) + 2
        }
        className={`border-x border-b px-2 py-1 text-center text-xs ${color}`}
      >
        <div className="flex items-center justify-center gap-2">
          <span className="font-semibold">{label}</span>
          <input
            type="number"
            min={0}
            max={100}
            defaultValue={weight(c)}
            onBlur={(e) =>
              upsertComponent.mutate({
                component: c,
                weight: Number(e.target.value || 0),
              })
            }
            className="h-6 w-12 rounded border bg-background text-center text-xs"
          />
          <span className="text-[10px]">%</span>
          {c !== "QA" && (
            <>
              <button
                type="button"
                onClick={() => removeActivity.mutate(c)}
                disabled={acts.length <= (c === "WW" ? 5 : 3) || removeActivity.isPending}
                className="rounded p-0.5 hover:bg-muted disabled:cursor-not-allowed disabled:opacity-40"
                title={`Remove last ${c === "WW" ? "Written / Oral Works" : "Product / Performance Tasks"} slot`}
                aria-label={`Remove last ${label} slot`}
              >
                <Minus className="size-3" />
              </button>
              <button
                type="button"
                onClick={() => addActivity.mutate(c)}
                disabled={addActivity.isPending}
                className="rounded p-0.5 hover:bg-muted disabled:cursor-not-allowed disabled:opacity-40"
                title={`Add ${c === "WW" ? "Written / Oral Works" : "Product / Performance Tasks"} slot`}
                aria-label={`Add ${label} slot`}
              >
                <Plus className="size-3" />
              </button>
            </>
          )}
        </div>
      </th>
    );
  };

  const enforceScoreInputLimit = (
    input: HTMLInputElement,
    hps: number,
  ) => {
    if (input.value === "") return;

    const maximum = Math.max(0, Number(hps) || 0);
    const entered = Number(input.value);

    if (!Number.isFinite(entered)) {
      input.value = "";
      return;
    }

    if (entered > maximum) {
      input.value = String(maximum);
      toast.error(`Maximum score is ${maximum}.`);
      return;
    }

    if (entered < 0) {
      input.value = "0";
      toast.error("Score cannot be below 0.");
    }
  };

  const safeScoreFromInput = (
    value: string,
    hps: number,
  ): number | null => {
    if (value.trim() === "") return null;

    const maximum = Math.max(0, Number(hps) || 0);
    const entered = Number(value);

    if (!Number.isFinite(entered)) return null;
    return Math.max(0, Math.min(maximum, entered));
  };

  const renderScoreCells = (sid: string, acts: GradeActivity[], comp: "WW" | "PT" | "QA") => {
    const row = computeRow(sid).results[comp];

    if (comp === "QA" && isMapehSubject) {
      const qaWeights = [30, 30, 40] as const;
      const qaWeighted = acts.map((activity, index) => {
        const score = scoreMap.get(`${activity.id}|${sid}`);
        if (typeof score !== "number" || !activity.hps) return "—";
        return Math.round((score / activity.hps) * 100 * (qaWeights[index] / 100) * 100) / 100;
      });

      return (
        <Fragment>
          {acts.map((a) => (
            <td key={a.id} className="p-0.5">
              <input
                type="number"
                min={0}
                max={a.hps}
                step="any"
                inputMode="decimal"
                defaultValue={scoreMap.get(`${a.id}|${sid}`) ?? ""}
                onKeyDown={(event) => {
                  if (["e", "E", "+", "-"].includes(event.key)) {
                    event.preventDefault();
                  }
                }}
                onInput={(event) =>
                  enforceScoreInputLimit(event.currentTarget, a.hps)
                }
                onBlur={(e) => {
                  const v = safeScoreFromInput(e.target.value, a.hps);
                  e.target.value = v == null ? "" : String(v);

                  if (v !== (scoreMap.get(`${a.id}|${sid}`) ?? null)) {
                    upsertScore.mutate({
                      activity_id: a.id,
                      student_id: sid,
                      score: v,
                      hps: a.hps,
                    });
                  }
                }}
                className="h-8 w-12 rounded border text-center text-xs"
                aria-label={`Student score. Maximum ${a.hps}`}
                title={`Maximum score: ${a.hps}`}
              />
            </td>
          ))}
          {qaWeighted.map((value, index) => (
            <td
              key={`qa-ws-${index}`}
              className="bg-muted/20 px-2 text-center text-xs font-semibold text-blue-700"
            >
              {value}
            </td>
          ))}
          <td className="border-l bg-muted/20 px-2 text-center text-xs font-semibold text-green-700">
            {row.ps ?? "—"}
          </td>
          <td className="bg-muted/20 px-2 text-center text-xs font-semibold text-blue-700">
            {row.ws == null ? "—" : row.ws.toFixed(2)}
          </td>
        </Fragment>
      );
    }

    return (
      <Fragment>
        {(acts.length ? acts : [null]).map((a, i) =>
          a ? (
            <td key={a.id} className="p-0.5">
              <input
                type="number"
                min={0}
                max={a.hps}
                step="any"
                inputMode="decimal"
                defaultValue={scoreMap.get(`${a.id}|${sid}`) ?? ""}
                onKeyDown={(event) => {
                  if (["e", "E", "+", "-"].includes(event.key)) {
                    event.preventDefault();
                  }
                }}
                onInput={(event) =>
                  enforceScoreInputLimit(event.currentTarget, a.hps)
                }
                onBlur={(e) => {
                  const v = safeScoreFromInput(e.target.value, a.hps);
                  e.target.value = v == null ? "" : String(v);

                  if (v !== (scoreMap.get(`${a.id}|${sid}`) ?? null)) {
                    upsertScore.mutate({
                      activity_id: a.id,
                      student_id: sid,
                      score: v,
                      hps: a.hps,
                    });
                  }
                }}
                className="h-8 w-12 rounded border text-center text-xs"
                aria-label={`Student score. Maximum ${a.hps}`}
                title={`Maximum score: ${a.hps}`}
              />
            </td>
          ) : (
            <td
              key={`${comp}-empty-${i}`}
              className="p-0.5 text-center text-xs text-muted-foreground"
            >
              —
            </td>
          ),
        )}
        <td className="border-l bg-muted/20 px-2 text-center text-xs font-semibold text-green-700">
          {row.ps ?? "—"}
        </td>
        <td className="bg-muted/20 px-2 text-center text-xs font-semibold text-blue-700">
          {row.ws == null
            ? "—"
            : comp === "QA"
              ? row.ws.toFixed(2)
              : row.ws}
        </td>
      </Fragment>
    );
  };

  const renderStudentRows = (rows: StudentRow[], groupLabel: string, headBg: string) => (
    <Fragment>
      <tr className={headBg}>
        <td
          className={`sticky left-0 z-30 w-[280px] min-w-[280px] max-w-[280px] border-r px-3 py-1.5 text-xs font-bold uppercase shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)] ${headBg}`}
        >
          {groupLabel}
        </td>
        <td
          colSpan={
            isMapehSubject
              ? wwActs.length + ptActs.length + qaActs.length + 13
              : wwActs.length + ptActs.length + qaActs.length + 7
          }
        />
      </tr>
      {rows.map((s, i) => {
        const c = computeRow(s.id);
        const finalTermGrade = resolvedTermGrade(s.id, c.term_grade);
        const d = descriptorFor(finalTermGrade);
        return (
          <tr key={s.id} className="border-t">
            <td className="sticky left-0 z-30 w-[280px] min-w-[280px] max-w-[280px] border-r bg-[#fffdf8] px-3 py-1.5 text-sm whitespace-nowrap shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
              <div className="flex min-w-0 items-center">
                <span className="inline-block w-7 shrink-0 text-xs text-muted-foreground">
                  {i + 1}.
                </span>
                <span className="min-w-0 truncate font-medium">
                  {s.last_name}, {s.first_name}
                </span>
              </div>
            </td>
            {renderScoreCells(s.id, wwActs, "WW")}
            {renderScoreCells(s.id, ptActs, "PT")}
            {renderScoreCells(s.id, qaActs, "QA")}
            <td className="border-l bg-green-50/60 px-2 text-center text-sm font-bold text-green-700">
              {c.initial ?? "—"}
            </td>
            <td className="bg-primary/5 px-1 text-center">
              {c.term_grade == null ? (
                "—"
              ) : (
                <div className="flex items-center justify-center">
                  <Input
                    key={`${s.id}-${term}-base-${finalTermGrade}`}
                    type="number"
                    min={0}
                    max={100}
                    step={1}
                    defaultValue={finalTermGrade ?? ""}
                    onBlur={(event) => {
                      const nextBase =
                        event.target.value.trim() === ""
                          ? c.term_grade!
                          : Number(event.target.value);
                      if (!Number.isFinite(nextBase)) {
                        event.target.value = String(finalTermGrade);
                        return;
                      }

                      const safeBase = Math.max(0, Math.min(100, Math.floor(nextBase + 0.5)));
                      event.target.value = String(safeBase);

                      if (safeBase !== finalTermGrade) {
                        saveTermGradeBase.mutate({
                          student_id: s.id,
                          baseTermGrade: safeBase,
                        });
                      }
                    }}
                    className="h-7 w-14 px-1 text-center text-xs font-semibold"
                    aria-label={`Editable Term Grade Base for ${s.last_name}, ${s.first_name}`}
                  />
                </div>
              )}
            </td>
            <td className="bg-primary/10 px-2 text-center text-base font-extrabold text-primary">
              {finalTermGrade ?? "—"}
            </td>
            <td className="px-2 text-center text-xs">
              <span
                className={`chip ${d.tone === "success" ? "bg-[color:var(--success-bg)] text-[color:var(--success)]" : d.tone === "danger" ? "bg-destructive/10 text-destructive" : "bg-muted text-muted-foreground"}`}
              >
                {d.label}
              </span>
            </td>
          </tr>
        );
      })}
    </Fragment>
  );

  return (
    <div className="space-y-2">
      {totalWeight !== 100 && (
        <div className="rounded-lg border border-yellow-300 bg-yellow-50 px-3 py-2 text-xs text-yellow-800">
          Component weights total {totalWeight}%. DepEd standard should equal 100%.
        </div>
      )}

      {/* The learner column stays fixed while score columns scroll behind it.
          Opaque sticky backgrounds prevent score inputs from showing through. */}
      <div className="relative overflow-x-auto rounded-2xl border bg-card shadow-sm">
        <table className="min-w-[1800px] border-separate border-spacing-0 text-sm">
          <thead>
            <tr>
              <th
                rowSpan={3}
                className="sticky left-0 z-40 w-[280px] min-w-[280px] max-w-[280px] border-r bg-[#fffaf2] px-3 py-2 text-left text-xs uppercase text-muted-foreground shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]"
              >
                Learner's Name
              </th>
              {componentHeader("WW", "Written / Oral Works", "bg-blue-50 text-blue-800")}
              {componentHeader("PT", "Product / Performance Tasks", "bg-orange-50 text-orange-800")}
              {
                /* SUMMATIVE UI: ST1 | ST2 | TE | WS ST1(30%) | WS ST2(30%) | WS TE(40%) | PS ST&TE | WS ST&TE */
                componentHeader(
                  "QA",
                  "Summative Tests & Term Exams",
                  "bg-violet-50 text-violet-800",
                )
              }
              <th
                rowSpan={3}
                className="border-l bg-green-50 px-2 text-center text-[11px] font-semibold text-green-700"
              >
                Initial
                <br />
                Grade
              </th>
              <th
                colSpan={2}
                className="bg-primary/10 px-2 py-1 text-center text-[11px] font-semibold text-primary"
              >
                Term Grade
                <div className="text-[9px] font-normal text-muted-foreground">
                  Editable Base = Final
                </div>
              </th>
              <th
                rowSpan={3}
                className="px-2 text-center text-[11px] font-semibold text-muted-foreground"
              >
                Descriptor
              </th>
            </tr>
            <tr className="bg-muted/20">
              {[wwActs, ptActs].map((acts, gi) => (
                <Fragment key={gi}>
                  {(acts.length ? acts : [null]).map((a, i) =>
                    a ? (
                      <th
                        key={a.id}
                        className="border-x px-1 py-1 text-center text-[10px] font-normal text-muted-foreground"
                      >
                        {a.position}
                      </th>
                    ) : (
                      <th
                        key={`empty-${gi}-${i}`}
                        className="border-x px-1 py-1 text-center text-[10px] text-muted-foreground"
                      >
                        —
                      </th>
                    ),
                  )}
                  <th className="border-l bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    PS
                  </th>
                  <th className="bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    WS
                  </th>
                </Fragment>
              ))}

              {isMapehSubject ? (
                <Fragment>
                  {["ST1", "ST2", "TE"].map((label) => (
                    <th
                      key={label}
                      className="border-x px-1 py-1 text-center text-[10px] font-normal text-muted-foreground"
                    >
                      {label}
                    </th>
                  ))}
                  {["WS ST1 (30%)", "WS ST2 (30%)", "WS TE (40%)"].map((label) => (
                    <th
                      key={label}
                      className="border-x bg-muted/20 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground"
                    >
                      {label}
                    </th>
                  ))}
                  <th className="border-l bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    PS ST&amp;TE
                  </th>
                  <th className="bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    WS ST&amp;TE
                  </th>
                </Fragment>
              ) : (
                <Fragment>
                  {(qaActs.length ? qaActs : [null]).map((a, i) =>
                    a ? (
                      <th
                        key={a.id}
                        className="border-x px-1 py-1 text-center text-[10px] font-normal text-muted-foreground"
                      >
                        {a.position}
                      </th>
                    ) : (
                      <th
                        key={`qa-empty-${i}`}
                        className="border-x px-1 py-1 text-center text-[10px] text-muted-foreground"
                      >
                        —
                      </th>
                    ),
                  )}
                  <th className="border-l bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    PS
                  </th>
                  <th className="bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    WS
                  </th>
                </Fragment>
              )}
              <th
                rowSpan={2}
                className="border-l bg-primary/5 px-2 py-1 text-center text-[10px] font-semibold text-primary"
              >
                Base
              </th>
              <th
                rowSpan={2}
                className="bg-primary/10 px-2 py-1 text-center text-[10px] font-semibold text-primary"
              >
                Final
              </th>
            </tr>
            <tr className="bg-muted/10">
              {[wwActs, ptActs].map((acts, gi) => (
                <Fragment key={gi}>
                  {(acts.length ? acts : [null]).map((a, i) =>
                    a ? (
                      <th key={a.id} className="border-x px-1 py-1">
                        <input
                          type="text"
                          inputMode="numeric"
                          pattern="[0-9]*"
                          defaultValue={a.hps}
                          onInput={(e) => {
                            // Allow the field to be temporarily blank while
                            // typing a new multi-digit HPS such as 10, 25,
                            // 50, 100, 150, 200, etc.
                            const digitsOnly = e.currentTarget.value.replace(
                              /[^0-9]/g,
                              "",
                            );

                            if (e.currentTarget.value !== digitsOnly) {
                              e.currentTarget.value = digitsOnly;
                            }
                          }}
                          onBlur={(e) => {
                            const raw = e.target.value.trim();
                            const entered = Number(raw);

                            const v =
                              raw !== "" &&
                              Number.isFinite(entered) &&
                              entered >= 0
                                ? Math.max(0, Math.round(entered))
                                : 10;

                            e.target.value = String(v);

                            if (v !== a.hps) {
                              updateHps.mutate({ id: a.id, hps: v });
                            }
                          }}
                          className="h-6 w-12 rounded border text-center text-[10px]"
                          aria-label="Highest possible score"
                          title="Enter 0 or any positive whole number"
                        />
                      </th>
                    ) : (
                      <th
                        key={`hps-empty-${gi}-${i}`}
                        className="border-x px-1 py-1 text-[10px] text-muted-foreground"
                      >
                        —
                      </th>
                    ),
                  )}
                  <th className="border-l bg-muted/30 text-[10px] text-muted-foreground">100</th>
                  <th className="bg-muted/30 text-[10px] text-muted-foreground">
                    {weight(gi === 0 ? "WW" : "PT")}
                  </th>
                </Fragment>
              ))}

              {isMapehSubject ? (
                <Fragment>
                  {(qaActs.length ? qaActs : [null]).map((a, i) =>
                    a ? (
                      <th key={a.id} className="border-x px-1 py-1">
                        <input
                          type="text"
                          inputMode="numeric"
                          pattern="[0-9]*"
                          defaultValue={a.hps}
                          onInput={(e) => {
                            // Allow the field to be temporarily blank while
                            // typing a new multi-digit HPS such as 10, 25,
                            // 50, 100, 150, 200, etc.
                            const digitsOnly = e.currentTarget.value.replace(
                              /[^0-9]/g,
                              "",
                            );

                            if (e.currentTarget.value !== digitsOnly) {
                              e.currentTarget.value = digitsOnly;
                            }
                          }}
                          onBlur={(e) => {
                            const raw = e.target.value.trim();
                            const entered = Number(raw);

                            const v =
                              raw !== "" &&
                              Number.isFinite(entered) &&
                              entered >= 0
                                ? Math.max(0, Math.round(entered))
                                : 10;

                            e.target.value = String(v);

                            if (v !== a.hps) {
                              updateHps.mutate({ id: a.id, hps: v });
                            }
                          }}
                          className="h-6 w-12 rounded border text-center text-[10px]"
                          aria-label="Highest possible score"
                          title="Enter 0 or any positive whole number"
                        />
                      </th>
                    ) : (
                      <th
                        key={`mapeh-hps-empty-${i}`}
                        className="border-x px-1 py-1 text-[10px] text-muted-foreground"
                      >
                        —
                      </th>
                    ),
                  )}
                  <th className="border-x bg-muted/20 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    30
                  </th>
                  <th className="border-x bg-muted/20 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    30
                  </th>
                  <th className="border-x bg-muted/20 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    40
                  </th>
                  <th className="border-l bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    100.00
                  </th>
                  <th className="bg-muted/30 px-1 py-1 text-center text-[10px] font-semibold text-muted-foreground">
                    {weight("QA")}%
                  </th>
                </Fragment>
              ) : (
                <Fragment>
                  {(qaActs.length ? qaActs : [null]).map((a, i) =>
                    a ? (
                      <th key={a.id} className="border-x px-1 py-1">
                        <input
                          type="text"
                          inputMode="numeric"
                          pattern="[0-9]*"
                          defaultValue={a.hps}
                          onInput={(e) => {
                            // Allow the field to be temporarily blank while
                            // typing a new multi-digit HPS such as 10, 25,
                            // 50, 100, 150, 200, etc.
                            const digitsOnly = e.currentTarget.value.replace(
                              /[^0-9]/g,
                              "",
                            );

                            if (e.currentTarget.value !== digitsOnly) {
                              e.currentTarget.value = digitsOnly;
                            }
                          }}
                          onBlur={(e) => {
                            const raw = e.target.value.trim();
                            const entered = Number(raw);

                            const v =
                              raw !== "" &&
                              Number.isFinite(entered) &&
                              entered >= 0
                                ? Math.max(0, Math.round(entered))
                                : 10;

                            e.target.value = String(v);

                            if (v !== a.hps) {
                              updateHps.mutate({ id: a.id, hps: v });
                            }
                          }}
                          className="h-6 w-12 rounded border text-center text-[10px]"
                          aria-label="Highest possible score"
                          title="Enter 0 or any positive whole number"
                        />
                      </th>
                    ) : (
                      <th
                        key={`qa-hps-empty-${i}`}
                        className="border-x px-1 py-1 text-[10px] text-muted-foreground"
                      >
                        —
                      </th>
                    ),
                  )}
                  <th className="border-l bg-muted/30 text-[10px] text-muted-foreground">100</th>
                  <th className="bg-muted/30 text-[10px] text-muted-foreground">
                    {weight("QA")}
                  </th>
                </Fragment>
              )}
            </tr>
          </thead>
          <tbody>
            {renderStudentRows(male, "Male", "bg-blue-50")}
            {renderStudentRows(female, "Female", "bg-pink-50")}
            {students.length === 0 && (
              <tr>
                <td colSpan={99} className="p-6 text-center text-muted-foreground">
                  Add learners first.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-muted-foreground">
        Each term starts with 5 Written / Oral Works slots and 3 Product / Performance Tasks slots.
        Use <Plus className="inline size-3" /> / <Minus className="inline size-3" /> to add or
        remove optional slots. ST1, ST2, and TE remain fixed. Highest possible scores are editable
        in the row below and may be 0 or any positive whole number. Blank slots are not included
        in the grade calculation. PS = percentage score, WS = weighted score.
      </p>
    </div>
  );
}

function MapehSummaryGradesTable({
  classId,
  students,
  subject,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
}) {
  const summaryScopes = useMemo(
    () => [
      { key: "MA_T1", storageTerm: "1" },
      { key: "PEH_T1", storageTerm: "PEH_T1" },
      { key: "MA_T2", storageTerm: "MA_T2" },
      { key: "PEH_T2", storageTerm: "PEH_T2" },
      { key: "MA_T3", storageTerm: "MA_T3" },
      { key: "PEH_T3", storageTerm: "PEH_T3" },
    ],
    [],
  );
  const storageTerms = summaryScopes.map((scope) => scope.storageTerm);
  const summarySubjects = summaryScopes.map((scope) => `${subject}_${scope.key}_ALL`);

  const { data: components = [] } = useQuery({
    queryKey: ["mapeh-summary-components", classId],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_components")
        .select("*")
        .eq("class_id", classId)
        .in("term", storageTerms);
      if (error) throw error;
      return (data ?? []) as GradeComponent[];
    },
  });

  const { data: activities = [] } = useQuery({
    queryKey: ["mapeh-summary-activities", classId],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_activities")
        .select("*")
        .eq("class_id", classId)
        .in("term", storageTerms)
        .order("position");
      if (error) throw error;
      return (data ?? []) as GradeActivity[];
    },
  });

  const activityIds = activities.map((activity) => activity.id);
  const { data: scores = [] } = useQuery({
    queryKey: ["mapeh-summary-scores", classId, activityIds.join(",")],
    queryFn: async () => {
      if (!activityIds.length) return [];
      const { data, error } = await supabase
        .from("activity_scores")
        .select("*")
        .in("activity_id", activityIds);
      if (error) throw error;
      return (data ?? []) as ActivityScore[];
    },
  });

  const { data: savedBases = [] } = useQuery({
    queryKey: ["mapeh-summary-term-grade-bases", classId, subject],
    enabled: Boolean(subject),
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, subject, term_grade_base")
        .eq("class_id", classId)
        .eq("term", "1")
        .in("subject", summarySubjects);
      if (error) throw error;
      return (data ?? []) as Array<{
        student_id: string;
        subject: string;
        term_grade_base: number | null;
      }>;
    },
  });

  const scoreMap = useMemo(() => {
    const map = new Map<string, number | null>();
    scores.forEach((score) => map.set(`${score.activity_id}|${score.student_id}`, score.score));
    return map;
  }, [scores]);

  const savedBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    savedBases.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) {
        map.set(`${row.student_id}|${row.subject}`, value);
      }
    });
    return map;
  }, [savedBases]);

  const componentGrade = (studentId: string, scopeKey: string, storageTerm: string) => {
    const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
      const componentActivities = activities
        .filter((activity) => activity.term === storageTerm && activity.component === component)
        .sort((a, b) => a.position - b.position)
        .slice(0, component === "QA" ? 3 : undefined);

      const componentWeight =
        components.find((item) => item.term === storageTerm && item.component === component)
          ?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);

      let raw = 0;
      let hps = 0;
      let hasScore = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;

      componentActivities.forEach((activity, index) => {
        const value = scoreMap.get(`${activity.id}|${studentId}`);
        if (typeof value !== "number") return;

        raw += value;
        hps += activity.hps;
        hasScore = true;

        if (component === "QA" && activity.hps > 0) {
          qaPercentageScore += (value / activity.hps) * 100 * ((qaPartWeights[index] ?? 0) / 100);
        }
      });

      if (!hasScore || hps <= 0) return null;

      const percentageScore = component === "QA" ? qaPercentageScore : (raw / hps) * 100;
      return percentageScore * (componentWeight / 100);
    });

    if (weightedParts.some((part) => part == null)) return null;

    const initialGrade =
      Math.round((weightedParts as number[]).reduce((sum, part) => sum + part, 0) * 100) / 100;
    const computedBase = roundInitialGrade(initialGrade);
    if (computedBase == null) return null;

    return resolveTermGradeBase(
      computedBase,
      savedBaseMap.get(`${studentId}|${subject}_${scopeKey}_ALL`),
    );
  };

  const averagePair = (first: number | null, second: number | null) => {
    if (first == null || second == null) return null;
    return Math.round((first + second) / 2);
  };

  return (
    <div className="space-y-2">
      <div className="relative overflow-x-auto rounded-2xl border bg-card shadow-sm">
        <table className="w-full min-w-[980px] border-separate border-spacing-0 text-sm">
          <thead className="bg-[#f7f3ed] text-xs uppercase text-muted-foreground">
            <tr>
              <th className="sticky left-0 z-40 w-[320px] min-w-[320px] max-w-[320px] border-r bg-[#f7f3ed] px-4 py-3 text-left shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
                Learner
              </th>
              <th className="min-w-[150px] px-4 py-3 text-center">1st Term</th>
              <th className="min-w-[150px] px-4 py-3 text-center">2nd Term</th>
              <th className="min-w-[150px] px-4 py-3 text-center">3rd Term</th>
              <th className="min-w-[130px] px-4 py-3 text-center">Final</th>
              <th className="min-w-[180px] px-4 py-3 text-left">Descriptor</th>
            </tr>
          </thead>
          <tbody>
            {students.map((student) => {
              const maT1 = componentGrade(student.id, "MA_T1", "1");
              const pehT1 = componentGrade(student.id, "PEH_T1", "PEH_T1");
              const maT2 = componentGrade(student.id, "MA_T2", "MA_T2");
              const pehT2 = componentGrade(student.id, "PEH_T2", "PEH_T2");
              const maT3 = componentGrade(student.id, "MA_T3", "MA_T3");
              const pehT3 = componentGrade(student.id, "PEH_T3", "PEH_T3");

              const term1 = averagePair(maT1, pehT1);
              const term2 = averagePair(maT2, pehT2);
              const term3 = averagePair(maT3, pehT3);
              const termValues = [term1, term2, term3];
              const final = termValues.every((value): value is number => typeof value === "number")
                ? Math.round(termValues.reduce((sum, value) => sum + value, 0) / 3)
                : null;
              const descriptor = descriptorFor(final);

              return (
                <tr key={student.id} className="border-t">
                  <td className="sticky left-0 z-30 w-[320px] min-w-[320px] max-w-[320px] border-r bg-[#fffdf8] px-4 py-3 font-medium whitespace-nowrap shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
                    <span className="block truncate">
                      {student.last_name}, {student.first_name}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-center text-base">{term1 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-base">{term2 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-base">{term3 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-lg font-extrabold text-primary">
                    {final ?? "—"}
                  </td>
                  <td className="px-4 py-3">
                    <span
                      className={`chip ${
                        descriptor.tone === "success"
                          ? "bg-[color:var(--success-bg)] text-[color:var(--success)]"
                          : descriptor.tone === "danger"
                            ? "bg-destructive/10 text-destructive"
                            : "bg-muted text-muted-foreground"
                      }`}
                    >
                      {descriptor.label}
                    </span>
                  </td>
                </tr>
              );
            })}
            {students.length === 0 && (
              <tr>
                <td colSpan={6} className="p-8 text-center text-muted-foreground">
                  No learners yet.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-muted-foreground">
        Each term grade is the rounded average of its MA and PEH grades. Final grade is the rounded
        average of the 1st, 2nd, and 3rd term grades.
      </p>
    </div>
  );
}

function CommunicationSummaryGradesTable({
  classId,
  students,
  subject,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
}) {
  const summaryScopes = useMemo(
    () => [
      { key: "EC_T1", storageTerm: "EC_T1" },
      { key: "MK_T1", storageTerm: "MK_T1" },
      { key: "EC_T2", storageTerm: "EC_T2" },
      { key: "MK_T2", storageTerm: "MK_T2" },
      { key: "EC_T3", storageTerm: "EC_T3" },
      { key: "MK_T3", storageTerm: "MK_T3" },
    ],
    [],
  );

  const storageTerms = summaryScopes.map((scope) => scope.storageTerm);
  const summarySubjects = summaryScopes.map(
    (scope) => `${subject}_${scope.key}_ALL`,
  );

  const { data: components = [] } = useQuery({
    queryKey: ["communication-summary-components", classId],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_components")
        .select("*")
        .eq("class_id", classId)
        .in("term", storageTerms);
      if (error) throw error;
      return (data ?? []) as GradeComponent[];
    },
  });

  const { data: activities = [] } = useQuery({
    queryKey: ["communication-summary-activities", classId],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("grade_activities")
        .select("*")
        .eq("class_id", classId)
        .in("term", storageTerms)
        .order("position");
      if (error) throw error;
      return (data ?? []) as GradeActivity[];
    },
  });

  const activityIds = activities.map((activity) => activity.id);
  const { data: scores = [] } = useQuery({
    queryKey: [
      "communication-summary-scores",
      classId,
      activityIds.join(","),
    ],
    queryFn: async () => {
      if (!activityIds.length) return [];
      const { data, error } = await supabase
        .from("activity_scores")
        .select("*")
        .in("activity_id", activityIds);
      if (error) throw error;
      return (data ?? []) as ActivityScore[];
    },
  });

  const { data: savedBases = [] } = useQuery({
    queryKey: ["communication-summary-term-grade-bases", classId, subject],
    enabled: Boolean(subject),
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, subject, term_grade_base")
        .eq("class_id", classId)
        .eq("term", "1")
        .in("subject", summarySubjects);
      if (error) throw error;
      return (data ?? []) as Array<{
        student_id: string;
        subject: string;
        term_grade_base: number | null;
      }>;
    },
  });

  const scoreMap = useMemo(() => {
    const map = new Map<string, number | null>();
    scores.forEach((score) =>
      map.set(`${score.activity_id}|${score.student_id}`, score.score),
    );
    return map;
  }, [scores]);

  const savedBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    savedBases.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) {
        map.set(`${row.student_id}|${row.subject}`, value);
      }
    });
    return map;
  }, [savedBases]);

  const componentGrade = (
    studentId: string,
    scopeKey: string,
    storageTerm: string,
  ) => {
    const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
      const componentActivities = activities
        .filter(
          (activity) =>
            activity.term === storageTerm && activity.component === component,
        )
        .sort((a, b) => a.position - b.position)
        .slice(0, component === "QA" ? 3 : undefined);

      const componentWeight =
        components.find(
          (item) =>
            item.term === storageTerm && item.component === component,
        )?.weight ?? (component === "WW" ? 20 : component === "PT" ? 50 : 30);

      let raw = 0;
      let hps = 0;
      let hasScore = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;

      componentActivities.forEach((activity, index) => {
        const value = scoreMap.get(`${activity.id}|${studentId}`);
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
    const computedBase = roundInitialGrade(initialGrade);
    if (computedBase == null) return null;

    return resolveTermGradeBase(
      computedBase,
      savedBaseMap.get(`${studentId}|${subject}_${scopeKey}_ALL`),
    );
  };

  const averagePair = (first: number | null, second: number | null) => {
    if (first == null || second == null) return null;
    return Math.round((first + second) / 2);
  };

  return (
    <div className="space-y-2">
      <div className="relative overflow-x-auto rounded-2xl border bg-card shadow-sm">
        <table className="w-full min-w-[980px] border-separate border-spacing-0 text-sm">
          <thead className="bg-[#f7f3ed] text-xs uppercase text-muted-foreground">
            <tr>
              <th className="sticky left-0 z-40 w-[320px] min-w-[320px] max-w-[320px] border-r bg-[#f7f3ed] px-4 py-3 text-left shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
                Learner
              </th>
              <th className="min-w-[150px] px-4 py-3 text-center">1st Term</th>
              <th className="min-w-[150px] px-4 py-3 text-center">2nd Term</th>
              <th className="min-w-[150px] px-4 py-3 text-center">3rd Term</th>
              <th className="min-w-[130px] px-4 py-3 text-center">Final</th>
              <th className="min-w-[180px] px-4 py-3 text-left">Descriptor</th>
            </tr>
          </thead>
          <tbody>
            {students.map((student) => {
              const ecT1 = componentGrade(student.id, "EC_T1", "EC_T1");
              const mkT1 = componentGrade(student.id, "MK_T1", "MK_T1");
              const ecT2 = componentGrade(student.id, "EC_T2", "EC_T2");
              const mkT2 = componentGrade(student.id, "MK_T2", "MK_T2");
              const ecT3 = componentGrade(student.id, "EC_T3", "EC_T3");
              const mkT3 = componentGrade(student.id, "MK_T3", "MK_T3");

              const term1 = averagePair(ecT1, mkT1);
              const term2 = averagePair(ecT2, mkT2);
              const term3 = averagePair(ecT3, mkT3);
              const termValues = [term1, term2, term3];
              const final = termValues.every(
                (value): value is number => typeof value === "number",
              )
                ? Math.round(
                    termValues.reduce((sum, value) => sum + value, 0) / 3,
                  )
                : null;
              const descriptor = descriptorFor(final);

              return (
                <tr key={student.id} className="border-t">
                  <td className="sticky left-0 z-30 w-[320px] min-w-[320px] max-w-[320px] border-r bg-[#fffdf8] px-4 py-3 font-medium whitespace-nowrap shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
                    <span className="block truncate">
                      {student.last_name}, {student.first_name}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-center text-base">{term1 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-base">{term2 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-base">{term3 ?? "—"}</td>
                  <td className="px-4 py-3 text-center text-lg font-extrabold text-primary">
                    {final ?? "—"}
                  </td>
                  <td className="px-4 py-3">
                    <span
                      className={`chip ${
                        descriptor.tone === "success"
                          ? "bg-[color:var(--success-bg)] text-[color:var(--success)]"
                          : descriptor.tone === "danger"
                            ? "bg-destructive/10 text-destructive"
                            : "bg-muted text-muted-foreground"
                      }`}
                    >
                      {descriptor.label}
                    </span>
                  </td>
                </tr>
              );
            })}
            {students.length === 0 && (
              <tr>
                <td colSpan={6} className="p-8 text-center text-muted-foreground">
                  No learners yet.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-muted-foreground">
        Each term grade is the rounded average of Effective Communication and
        Mabisang Komunikasyon. The final grade is the rounded average of the
        1st, 2nd, and 3rd term grades.
      </p>
    </div>
  );
}

function FinalGradesTable({
  classId,
  students,
  subject,
  klass,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  klass: ClassRow;
}) {
  const grade12SingleTerm = getGrade12SingleTerm(
    klass?.grade_level,
    subject,
  );
  const isSingleTermGrade12Subject = Boolean(grade12SingleTerm);

  // Read term grades. Most subjects use all three terms. Mapped Grade 12
  // SF9 subjects use only their assigned non-gray term.
  const { data: components = [] } = useQuery({
    queryKey: ["components-all", classId],
    queryFn: async () =>
      ((await supabase.from("grade_components").select("*").eq("class_id", classId))
        .data as GradeComponent[]) ?? [],
  });
  const { data: activities = [] } = useQuery({
    queryKey: ["activities-all", classId],
    queryFn: async () =>
      ((await supabase.from("grade_activities").select("*").eq("class_id", classId))
        .data as GradeActivity[]) ?? [],
  });
  const activityIds = activities.map((a) => a.id);
  const { data: scores = [] } = useQuery({
    queryKey: ["scores-all", classId, activityIds.length],
    queryFn: async () => {
      if (!activityIds.length) return [];
      return (
        ((await supabase.from("activity_scores").select("*").in("activity_id", activityIds))
          .data as ActivityScore[]) ?? []
      );
    },
  });
  const { data: savedTermGradeBases = [] } = useQuery({
    queryKey: ["term-grade-bases-all", classId, subject],
    enabled: Boolean(subject),
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, term, term_grade_base")
        .eq("class_id", classId)
        .eq("subject", subject)
        .in("term", ["1", "2", "3"]);
      if (error) throw error;
      return (data ?? []) as Array<{
        student_id: string;
        term: string;
        term_grade_base: number | null;
      }>;
    },
  });
  const savedTermGradeBaseMap = new Map<string, number>();
  savedTermGradeBases.forEach((row) => {
    if (row.term_grade_base == null) return;
    const value = Number(row.term_grade_base);
    if (Number.isFinite(value)) {
      savedTermGradeBaseMap.set(`${row.student_id}|${row.term}`, value);
    }
  });
  const sMap = new Map<string, number | null>();
  scores.forEach((s) => sMap.set(`${s.activity_id}|${s.student_id}`, s.score));

  const termGrade = (sid: string, term: string) => {
    const results = (["WW", "PT", "QA"] as const).map((c) => {
      const acts = activities
        .filter((a) => a.term === term && a.component === c)
        .sort((a, b) => a.position - b.position);
      const w =
        components.find((x) => x.term === term && x.component === c)?.weight ??
        (c === "WW" ? 20 : c === "PT" ? 50 : 30);
      let raw = 0,
        hps = 0,
        any = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;
      acts.forEach((a) => {
        const v = sMap.get(`${a.id}|${sid}`);
        if (typeof v === "number") {
          raw += v;
          hps += a.hps;
          any = true;
          if (c === "QA") {
            const activityIndex = acts.indexOf(a);
            const partWeight = qaPartWeights[activityIndex] ?? 0;
            if (a.hps > 0) {
              qaPercentageScore += (v / a.hps) * 100 * (partWeight / 100);
            }
          }
        }
      });
      if (!any || hps === 0) return null;
      const ps = c === "QA" ? qaPercentageScore : (raw / hps) * 100;
      return ps * (w / 100);
    });
    if (results.some((r) => r == null)) return null;
    const initial = results.reduce((a, b) => a! + b!, 0)!;
    const base = roundInitialGrade(Math.round(initial * 100) / 100);
    if (base == null) return null;
    return resolveTermGradeBase(base, savedTermGradeBaseMap.get(`${sid}|${term}`));
  };
  const lastSyncedFinal = useRef<Map<string, number>>(new Map());
  const syncFinalGrade = useMutation({
    mutationFn: async ({ student_id, score }: { student_id: string; score: number }) => {
      const teacher_id = await getUserId();
      const { error } = await supabase.from("grades").upsert(
        {
          student_id,
          class_id: classId,
          teacher_id,
          subject,
          term: "final",
          score,
        },
        { onConflict: "student_id,subject,term" },
      );
      if (error) throw error;
    },
  });
  useEffect(() => {
    if (!subject) return;
    students.forEach((s) => {
      const t1 = termGrade(s.id, "1"),
        t2 = termGrade(s.id, "2"),
        t3 = termGrade(s.id, "3");

      const assignedGrade =
        grade12SingleTerm === "1"
          ? t1
          : grade12SingleTerm === "2"
            ? t2
            : grade12SingleTerm === "3"
              ? t3
              : null;

      const final = grade12SingleTerm
        ? assignedGrade
        : [t1, t2, t3].every(
              (value): value is number => typeof value === "number",
            )
          ? Math.round(((t1 as number) + (t2 as number) + (t3 as number)) / 3)
          : null;

      if (final == null) return;
      const prev = lastSyncedFinal.current.get(s.id);
      if (prev !== final) {
        lastSyncedFinal.current.set(s.id, final);
        syncFinalGrade.mutate({ student_id: s.id, score: final });
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    components,
    activities,
    scores,
    subject,
    students,
    grade12SingleTerm,
  ]);

  return (
    <div className="relative overflow-x-auto rounded-2xl border bg-card shadow-sm">
      <table className="min-w-[1800px] border-separate border-spacing-0 text-sm">
        <thead className="bg-[#f7f3ed] text-xs uppercase text-muted-foreground">
          <tr>
            <th className="sticky left-0 z-40 w-[280px] min-w-[280px] max-w-[280px] border-r bg-[#f7f3ed] px-3 py-2 text-left shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">Learner</th>
            {grade12SingleTerm ? (
              <th className="px-3 py-2 text-center">
                {TERM_TABS.find((tab) => tab.value === grade12SingleTerm)?.label}
              </th>
            ) : (
              <>
                <th className="px-3 py-2 text-center">1st Term</th>
                <th className="px-3 py-2 text-center">2nd Term</th>
                <th className="px-3 py-2 text-center">3rd Term</th>
              </>
            )}
            <th className="px-3 py-2 text-center">Final</th>
            <th className="px-3 py-2 text-left">Descriptor</th>
          </tr>
        </thead>
        <tbody>
          {students.map((s) => {
            const t1 = termGrade(s.id, "1"),
              t2 = termGrade(s.id, "2"),
              t3 = termGrade(s.id, "3");

            const assignedGrade =
              grade12SingleTerm === "1"
                ? t1
                : grade12SingleTerm === "2"
                  ? t2
                  : grade12SingleTerm === "3"
                    ? t3
                    : null;

            const final = grade12SingleTerm
              ? assignedGrade
              : [t1, t2, t3].every(
                    (value): value is number => typeof value === "number",
                  )
                ? Math.round(
                    ((t1 as number) + (t2 as number) + (t3 as number)) / 3,
                  )
                : null;

            const d = descriptorFor(final);
            return (
              <tr key={s.id} className="border-t">
                <td className="sticky left-0 z-30 w-[280px] min-w-[280px] max-w-[280px] border-r bg-[#fffdf8] px-3 py-1.5 shadow-[5px_0_10px_-8px_rgba(0,0,0,0.55)]">
                  <span className="block truncate font-medium">
                    {s.last_name}, {s.first_name}
                  </span>
                </td>
                {grade12SingleTerm ? (
                  <td className="px-3 py-1.5 text-center">
                    {grade12SingleTerm === "1"
                      ? (t1 ?? "—")
                      : grade12SingleTerm === "2"
                        ? (t2 ?? "—")
                        : (t3 ?? "—")}
                  </td>
                ) : (
                  <>
                    <td className="px-3 py-1.5 text-center">{t1 ?? "—"}</td>
                    <td className="px-3 py-1.5 text-center">{t2 ?? "—"}</td>
                    <td className="px-3 py-1.5 text-center">{t3 ?? "—"}</td>
                  </>
                )}
                <td className="px-3 py-1.5 text-center text-base font-extrabold text-primary">
                  {final ?? "—"}
                </td>
                <td className="px-3 py-1.5">
                  <span
                    className={`chip ${d.tone === "success" ? "bg-[color:var(--success-bg)] text-[color:var(--success)]" : d.tone === "danger" ? "bg-destructive/10 text-destructive" : "bg-muted text-muted-foreground"}`}
                  >
                    {d.label}
                  </span>
                </td>
              </tr>
            );
          })}
          {students.length === 0 && (
            <tr>
              <td
                colSpan={isSingleTermGrade12Subject ? 4 : 6}
                className="p-6 text-center text-muted-foreground"
              >
                No learners yet.
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

/* ---------------- Fitness Test ---------------- */
type FitnessStudent = StudentRow & {
  weight_kg?: number | null;
  height_m?: number | null;
  weight?: number | null;
  height?: number | null;
  current_weight?: number | null;
  current_height?: number | null;
  height_cm?: number | null;
};

function fitnessNumber(...values: unknown[]) {
  for (const value of values) {
    if (value === null || value === undefined || value === "") continue;
    const number = Number(value);
    if (Number.isFinite(number)) return number;
  }
  return null;
}

function fitnessWeight(student: StudentRow) {
  const record = student as FitnessStudent;
  return fitnessNumber(record.weight_kg, record.weight, record.current_weight);
}

function fitnessHeight(student: StudentRow) {
  const record = student as FitnessStudent;
  const raw = fitnessNumber(
    record.height_m,
    record.height,
    record.current_height,
    record.height_cm,
  );

  if (raw === null) return null;
  return raw > 3 ? raw / 100 : raw;
}

function fitnessBmi(weight: number | null, height: number | null) {
  if (
    weight === null ||
    height === null ||
    !Number.isFinite(weight) ||
    !Number.isFinite(height) ||
    weight <= 0 ||
    height <= 0
  ) {
    return null;
  }

  return Math.round((weight / (height * height)) * 10) / 10;
}

function FitnessTestPanel({ classId, students }: { classId: string; students: StudentRow[] }) {
  const qc = useQueryClient();

  const updateFitness = useMutation({
    mutationFn: async ({
      studentId,
      weight,
      height,
    }: {
      studentId: string;
      weight: number | null;
      height: number | null;
    }) => {
      const { error } = await (supabase as any)
        .from("students")
        .update({
          weight_kg: weight,
          height_m: height,
        })
        .eq("id", studentId)
        .eq("class_id", classId);

      if (error) throw error;
    },
    onSuccess: async () => {
      await qc.invalidateQueries({ queryKey: ["students", classId] });
      await qc.invalidateQueries({ queryKey: ["students"] });
    },
    onError: (error: Error) => {
      toast.error(`Unable to save fitness data: ${error.message}`);
    },
  });

  const saveFitnessValue = (student: StudentRow, field: "weight" | "height", rawValue: string) => {
    const currentWeight = fitnessWeight(student);
    const currentHeight = fitnessHeight(student);
    const parsed = rawValue.trim() === "" ? null : Number(rawValue);

    if (parsed !== null && (!Number.isFinite(parsed) || parsed < 0)) {
      toast.error("Enter a valid positive number.");
      return;
    }

    const nextWeight = field === "weight" ? parsed : currentWeight;
    const nextHeight = field === "height" ? parsed : currentHeight;

    if (nextWeight === currentWeight && nextHeight === currentHeight) {
      return;
    }

    updateFitness.mutate({
      studentId: student.id,
      weight: nextWeight,
      height: nextHeight,
    });
  };

  return (
    <div className="rounded-2xl border bg-card p-4 shadow-sm">
      <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
        <div className="flex items-start gap-3">
          <div className="grid size-10 shrink-0 place-items-center rounded-xl bg-primary/10 text-primary">
            <HeartPulse className="size-5" />
          </div>
          <div>
            <div className="text-base font-semibold">BMI</div>
            <div className="text-xs text-muted-foreground">
              Manage learner height and weight for SF8 health and nutrition data.
            </div>
          </div>
        </div>

        <div className="rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-medium text-emerald-700">
          Saved values sync automatically to SF8.
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border">
        <table className="w-full min-w-[760px] text-sm">
          <thead className="bg-muted/30 text-xs text-muted-foreground">
            <tr className="border-b">
              <th className="px-3 py-3 text-left font-semibold">Learner Name</th>
              <th className="px-3 py-3 text-center font-semibold">Sex</th>
              <th className="px-3 py-3 text-center font-semibold">Weight (kg)</th>
              <th className="px-3 py-3 text-center font-semibold">Height (m)</th>
              <th className="px-3 py-3 text-center font-semibold">BMI</th>
            </tr>
          </thead>

          <tbody>
            {students.map((student, index) => {
              const weight = fitnessWeight(student);
              const height = fitnessHeight(student);
              const bmi = fitnessBmi(weight, height);

              return (
                <tr key={student.id} className="border-b last:border-b-0">
                  <td className="px-3 py-2.5">
                    <span className="mr-2 text-xs text-muted-foreground">{index + 1}.</span>
                    <span className="font-medium">
                      {student.last_name}, {student.first_name} {student.middle_name ?? ""}
                    </span>
                  </td>

                  <td className="px-3 py-2.5 text-center capitalize">{student.sex || "—"}</td>

                  <td className="px-3 py-2.5 text-center">
                    <Input
                      key={`${student.id}-weight-${weight ?? "blank"}`}
                      type="number"
                      min={0}
                      step="0.1"
                      defaultValue={weight ?? ""}
                      onBlur={(event) => saveFitnessValue(student, "weight", event.target.value)}
                      className="mx-auto h-8 w-28 text-center"
                      aria-label={`Weight for ${student.last_name}, ${student.first_name}`}
                    />
                  </td>

                  <td className="px-3 py-2.5 text-center">
                    <Input
                      key={`${student.id}-height-${height ?? "blank"}`}
                      type="number"
                      min={0}
                      step="0.01"
                      defaultValue={height ?? ""}
                      onBlur={(event) => saveFitnessValue(student, "height", event.target.value)}
                      className="mx-auto h-8 w-28 text-center"
                      aria-label={`Height for ${student.last_name}, ${student.first_name}`}
                    />
                  </td>

                  <td className="px-3 py-2.5 text-center font-semibold text-emerald-700">
                    {bmi !== null ? bmi.toFixed(1) : "—"}
                  </td>
                </tr>
              );
            })}

            {students.length === 0 && (
              <tr>
                <td colSpan={5} className="px-4 py-10 text-center text-sm text-muted-foreground">
                  No learners yet.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <p className="mt-3 text-xs text-muted-foreground">
        BMI is calculated automatically. Formula: BMI = weight (kg) / height (m)².
      </p>
    </div>
  );
}
