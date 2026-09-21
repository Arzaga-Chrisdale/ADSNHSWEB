import { createFileRoute, useNavigate } from "@tanstack/react-router";
import {
  Fragment,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type ComponentType,
  type ReactNode,
} from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  AlertTriangle,
  BarChart3,
  BookOpen,
  CalendarDays,
  CalendarRange,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleCheck,
  ClipboardCheck,
  Copy,
  Crown,
  FileSpreadsheet,
  FileStack,
  FileText,
  Files,
  FolderOpen,
  GraduationCap,
  HeartHandshake,
  IdCard,
  LayoutDashboard,
  Loader2,
  LockKeyhole,
  LogOut,
  Mars,
  Maximize2,
  Menu,
  Minus,
  Pencil,
  Plus,
  Printer,
  RefreshCcw,
  Save,
  Search,
  ShieldCheck,
  Trash2,
  UserCheck,
  UserRound,
  Users,
  UsersRound,
  UnlockKeyhole,
  Venus,
  X,
} from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import { useAdminOnlineStatus } from "@/hooks/useAdminOnlineStatus";
import { releaseAdminAccess } from "@/lib/admin-session";
import { NotificationBell } from "@/components/NotificationBell";
import { AnalyticsInsightsPanel } from "@/components/AnalyticsInsightsPanel";
import type { StudentRow as AnalyticsStudentRow } from "@/lib/data";
import { getSF5Action } from "@/lib/sf5-action";
import {
  downloadSF5ExcelTemplate,
  downloadSF5WordTemplate,
  type SF5TemplateOptions,
  type SF5TemplateRow,
} from "../../lib/sf5-template-export";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  AlignmentType,
  BorderStyle,
  Document as WordDocument,
  HeightRule,
  ImageRun,
  Packer,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import CFB from "cfb";
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
import logo from "@/assets/ASNSHS Logo.png";
import logos from "@/assets/deped_logo.png";
import deped from "@/assets/deped_logo.png";

// Admin color values are defined in src/styles.css under the ADMIN UI COLOR TOKENS section.

export const Route = createFileRoute("/_authenticated/admin")({
  component: AdminDashboardPage,
});

type AdminSection =
  | "dashboard"
  | "users"
  | "students"
  | "credential_requests"
  | "forms"
  | "analytics"
  | "notifications"
  | "profile";

type ProfileRow = {
  id: string;
  full_name: string | null;
  email: string | null;
  school_name: string | null;
  school_id: string | null;
  division: string | null;
  district: string | null;
  region: string | null;
  principal?: string | null;
  municipality?: string | null;
  city_municipality_province?: string | null;
  avatar_url?: string | null;
  teacher_type?: string | null;
  created_at: string;
};

type RoleRow = {
  user_id: string;
  role: "admin" | "teacher" | string;
};

type ClassRow = {
  id: string;
  teacher_id: string;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  teacher_name?: string | null;
  track_shs?: string | null;
  units?: number | null;
  created_at?: string | null;
};

type StudentRow = {
  id: string;
  class_id: string;
  teacher_id: string;
  lrn: string | null;
  first_name: string | null;
  middle_name: string | null;
  last_name: string | null;
  sex: string | null;
  address: string | null;
  guardian: string | null;
  mother_name: string | null;
  father_name: string | null;
  contact_number: string | null;
  birthdate: string | null;
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
  remarks?: string | null;
  disability_condition?: string | null;
  disability?: string | null;
  learning_modality?: string | null;
  created_at?: string | null;

  // Transfer In / Transfer Out state.
  enrollment_status?: "active" | "transferred_in" | "transferred_out" | string | null;
  is_active?: boolean | null;
  transfer_date?: string | null;
  transfer_effective_term?: string | null;
  previous_school?: string | null;
  destination_school?: string | null;
  transfer_reason?: string | null;
  transferred_at?: string | null;
};

type AttendanceRow = {
  id: string;
  student_id: string;
  class_id: string;
  teacher_id: string;
  status: string | null;
  date: string | null;
};

type GradeRow = {
  id: string;
  student_id: string;
  class_id: string;
  teacher_id: string;
  subject: string | null;
  term: string | null;
  score: number | null;
};

type StudentCredentialRequestRow = {
  id: string;
  student_id: string | null;
  requested_student_name?: string | null;
  requester_id: string;
  requesting_class_id: string | null;
  previous_class_id: string | null;
  reason: string | null;
  status: "pending" | "completed" | string;
  release_notes: string | null;
  admin_remarks?: string | null;
  reviewed_by?: string | null;
  reviewed_at?: string | null;
  released_at?: string | null;
  processed_by: string | null;
  processed_at: string | null;
  request_date: string;
  created_at: string;
};

type SchoolFormSubmissionRow = {
  id: string;
  class_id: string;
  form_code: string;
  adviser_id: string;
  adviser_name: string | null;
  grade_level: string;
  section: string | null;
  school_year: string | null;
  status: "submitted" | "pending_review" | "approved" | "returned";
  admin_remarks: string | null;
  submitted_at: string;
  reviewed_at: string | null;
  reviewed_by: string | null;
  snapshot: Record<string, unknown> | null;
};

type AdminData = {
  isAdmin: boolean;
  currentUserId: string | null;
  profiles: ProfileRow[];
  roles: RoleRow[];
  classes: ClassRow[];
  students: StudentRow[];
  attendance: AttendanceRow[];
  grades: GradeRow[];
  credentialRequests: StudentCredentialRequestRow[];
};

type AdminSchoolYearLibraryRow = {
  id: number;
  school_year: string;
  is_locked: boolean;
};

type AdminSchoolYearSetting = {
  id: number;
  school_year: string | null;
  is_locked: boolean;
  updated_by: string | null;
  updated_at: string | null;
};

type Sf6Count = {
  male: number;
  female: number;
  total: number;
};

type Sf6MetricsData = {
  gradeReports: Array<{
    grade: string;
    label: string;
    promoted: Sf6Count;
    conditional: Sf6Count;
    retained: Sf6Count;
    enrolled: Sf6Count;
    bands: Sf6Count[];
  }>;
  promotedByGrade: Sf6Count[];
  conditionalByGrade: Sf6Count[];
  retainedByGrade: Sf6Count[];
  enrolledByGrade: Sf6Count[];
  bandValues: Sf6Count[][];
};

type IlsMetricsData = {
  learner: StudentRow | null;
  parentGuardian: string;
  subjects: Array<{
    subject: string;
    term1: number | null;
    term2: number | null;
    term3: number | null;
    average: number | null;
    trend: string;
    remarks: string;
  }>;
  generalAverage: number | null;
  classRank: number | null;
  classSize: number;
  classAverage: number | null;
  failingSubjectCount: number;
  absences: number;
  lateHalfDays: number;
  missingActivities: number;
  riskLevel: "LOW" | "MODERATE" | "HIGH";
  componentSubject: string;
  componentRows: Array<{
    termLabel: string;
    writtenWorks: string;
    performanceTasks: string;
    summativeExam: string;
  }>;
  recommendations: string[];
};

type FormTypeDefinition = {
  id: string;
  title: string;
  description: string;
  icon: ComponentType<{ className?: string }>;
  iconBackground: string;
  iconColor: string;
};

const RAW_FORM_TYPE_DEFINITIONS: FormTypeDefinition[] = [
  {
    id: "GSA",
    title: "GSA",
    description: "General Scholastic Aptitude Report",
    icon: FileText,
    iconBackground: "bg-rose-50",
    iconColor: "text-rose-600",
  },
  {
    id: "SF5",
    title: "SF5",
    description: "Report on Promotion & Level of Proficiency",
    icon: GraduationCap,
    iconBackground: "bg-pink-50",
    iconColor: "text-pink-600",
  },
  {
    id: "SOG",
    title: "SOG Report",
    description: "Printable Summary of Grades Per Term",
    icon: FileStack,
    iconBackground: "bg-indigo-50",
    iconColor: "text-indigo-600",
  },
  {
    id: "SF1",
    title: "SF1",
    description: "School Register",
    icon: BookOpen,
    iconBackground: "bg-emerald-50",
    iconColor: "text-emerald-600",
  },
  {
    id: "SF9_NEW",
    title: "SF9 (New)",
    description: "Learner's Progress/Performance Report",
    icon: FileText,
    iconBackground: "bg-orange-50",
    iconColor: "text-orange-600",
  },
  {
    id: "SF6",
    title: "SF6",
    description: "Summarized Report on Promotion",
    icon: ShieldCheck,
    iconBackground: "bg-cyan-50",
    iconColor: "text-cyan-600",
  },
  {
    id: "SF8",
    title: "SF8",
    description: "Learner Health & Nutrition Profile",
    icon: HeartHandshake,
    iconBackground: "bg-fuchsia-50",
    iconColor: "text-fuchsia-600",
  },
  {
    id: "SF10",
    title: "SF10",
    description: "Learner's Permanent Academic Record",
    icon: BookOpen,
    iconBackground: "bg-blue-50",
    iconColor: "text-blue-600",
  },
  {
    id: "ANECDOTAL",
    title: "Anecdotal",
    description: "Anecdotal Record",
    icon: FileText,
    iconBackground: "bg-orange-50",
    iconColor: "text-orange-600",
  },
  {
    id: "ITEM_ANALYSIS",
    title: "Item Analysis",
    description: "Item Analysis Report",
    icon: BarChart3,
    iconBackground: "bg-green-50",
    iconColor: "text-green-600",
  },
];

// Forms available in the Admin School Forms sidebar.
const ADMIN_VISIBLE_FORM_TYPE_IDS = [
  "GSA",
  "SOG",
  "SF5",
  "SF1",
  "SF9_NEW",
  "SF6",
  "SF8",
  "SF10",
  "ANECDOTAL",
] as const;

const FORM_TYPE_DEFINITIONS: FormTypeDefinition[] =
  ADMIN_VISIBLE_FORM_TYPE_IDS.map((formId) =>
    RAW_FORM_TYPE_DEFINITIONS.find((form) => form.id === formId),
  )
    .filter((form): form is FormTypeDefinition => Boolean(form))
    .map((form) =>
      form.id === "SF9_NEW"
        ? {
            ...form,
            title: "SF9 New Matatag",
            description: "Learner Progress and Performance Report",
          }
        : form,
    );

const FORM_TYPES = FORM_TYPE_DEFINITIONS.map((form) => form.id);

// Grade-level School Forms visibility for the Admin page.
// Grades 7-10 follow the Junior High School form set shown in the teacher UI.
// Grades 11-12 follow the Senior High School form set shown in the teacher UI.
const ADMIN_JHS_FORM_TYPE_IDS = [
  "GSA",
  "SF5",
  "SOG",
  "SF1",
  "SF9_NEW",
  "SF8",
  "SF10",
  "ANECDOTAL",
] as const;

const ADMIN_SHS_FORM_TYPE_IDS = [
  "SF5",
  "SF1",
  "SF9_NEW",
  "SF10",
] as const;

function adminFormIdsForGradeLevel(gradeLevel: string) {
  if (["7", "8", "9", "10"].includes(gradeLevel)) {
    return ADMIN_JHS_FORM_TYPE_IDS as readonly string[];
  }

  if (["11", "12"].includes(gradeLevel)) {
    return ADMIN_SHS_FORM_TYPE_IDS as readonly string[];
  }

  // "All Grade Levels" keeps the complete Admin form library available.
  return ADMIN_VISIBLE_FORM_TYPE_IDS as readonly string[];
}

function adminGradeLevelFormsLabel(gradeLevel: string) {
  if (["7", "8", "9", "10"].includes(gradeLevel)) {
    return `Junior High School · Grade ${gradeLevel}`;
  }

  if (["11", "12"].includes(gradeLevel)) {
    return `Senior High School · Grade ${gradeLevel}`;
  }

  return "All Grade Levels";
}

type PaperSizeKey = "long" | "letter" | "a4";
type Sf5ReportLength = "short" | "full";
type Sf1ReportLength = "short" | "full";
type SogTerm = "1" | "2" | "3";

const PAPER_SIZE_PRESETS: Record<
  PaperSizeKey,
  { label: string; width: number; height: number }
> = {
  long: { label: "Long (8.5×13)", width: 816, height: 1248 },
  letter: { label: "Short / Letter (8.5×11)", width: 816, height: 1056 },
  a4: { label: "A4 (8.27×11.69)", width: 794, height: 1123 },
};
// Admin SF9 Senior High School templates. Grade 11 and Grade 12 use the
// same A4 landscape bond-paper structure and spreadsheet templates as the
// Class Adviser SF9 page.
const ADMIN_SF9_GRADE11_TEMPLATE_URL = "/templates/sf9-grade-11-shs.xlsx";
const ADMIN_SF9_GRADE12_TEMPLATE_URL = "/templates/sf9-grade-12-shs.xlsx";

type AdminSf9AssignedTerm = "1" | "2" | "3";

const ADMIN_SF9_GRADE12_SUBJECTS = [
  { label: "Media and Information Literacy", term: "1", excelRow: 33, group: "core" },
  { label: "PE and Health 3", term: "1", excelRow: 34, group: "core" },
  { label: "Introduction to Human Philosophy", term: "2", excelRow: 35, group: "core" },
  { label: "Disaster Readiness and Risk Reduction", term: "2", excelRow: 36, group: "core" },
  { label: "Contemporary Philippine Arts", term: "3", excelRow: 37, group: "core" },
  { label: "PE and Health 4", term: "3", excelRow: 38, group: "core" },
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

function normalizeAdminSf9Subject(value: unknown) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function adminSf9GradeKey(value: unknown) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

// Keep the Admin SF5 bond-paper preview at the same compact visual size as the
// Class Adviser SF5 preview. The logical page remains 1248 × 816 so printing
// still produces an exact 13 × 8.5 inch landscape page.
const ADMIN_SF5_PREVIEW_SCALE = 5 / 6;
// Original SF5 spreadsheet column geometry (A:Z). Admin uses the same grid as
// the Class Adviser preview so the title and metadata align with the template.
const ADMIN_SF5_COLUMN_WIDTHS = [
  2.51, 10.91, 6.21, 10.07, 8.89, 9.22, 9.9, 6.04, 12.42, 1, 11.75, 5.03, 1,
  0.16, 5.53, 6.04, 3.18, 0.83, 4.52, 9.9, 3.18, 3.35, 1.33, 3.18, 0.16, 0.16,
] as const;
const PAGE_SIZE = 8;
const EMPTY_PROFILES: ProfileRow[] = [];
const EMPTY_ROLES: RoleRow[] = [];
const EMPTY_CLASSES: ClassRow[] = [];
const EMPTY_STUDENTS: StudentRow[] = [];
const EMPTY_ATTENDANCE: AttendanceRow[] = [];
const EMPTY_GRADES: GradeRow[] = [];
const EMPTY_CREDENTIAL_REQUESTS: StudentCredentialRequestRow[] = [];
const EMPTY_SF6_COUNT: Sf6Count = { male: 0, female: 0, total: 0 };

const SF6_GRADE_LEVELS = [
  { value: "7", label: "GRADE 7" },
  { value: "8", label: "GRADE 8" },
  { value: "9", label: "GRADE 9" },
  { value: "10", label: "GRADE 10" },
  { value: "11", label: "GRADE 11" },
  { value: "12", label: "GRADE 12" },
] as const;

const SF6_PROGRESS_BANDS = [
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
    matches: (average: number) => average >= 75 && average < 80,
  },
  {
    key: "satisfactory",
    label: "Satisfactory",
    detail: "(80%–84%)",
    matches: (average: number) => average >= 80 && average < 85,
  },
  {
    key: "very-satisfactory",
    label: "Very Satisfactory",
    detail: "(85%–89%)",
    matches: (average: number) => average >= 85 && average < 90,
  },
  {
    key: "outstanding",
    label: "Outstanding",
    detail: "(90%–100%)",
    matches: (average: number) => average >= 90,
  },
] as const;

const SF6_DEPED_SEAL_URL =
  "/__l5e/assets-v1/83e36017-55a9-49b9-a821-fb04e48056ba/deped-logo.png";

type Sf6ExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  statusRows: Array<{
    label: string;
    values: Sf6Count[];
    total: Sf6Count;
  }>;
  bandValues: Sf6Count[][];
  enrolledByGrade: Sf6Count[];
  schoolHead: string;
};

const SF6_BLACK = { argb: "FF000000" };
const SF6_THIN_BORDER = {
  style: "thin" as const,
  color: SF6_BLACK,
};
const SF6_MEDIUM_BORDER = {
  style: "medium" as const,
  color: SF6_BLACK,
};
const SF6_BODY_FONT: Partial<Font> = { name: "Arial", size: 7 };
const SF6_CENTERED: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};
const SF6_NO_FILL: Fill = { type: "pattern", pattern: "none" };
const SF6_COLUMN_WIDTHS = [
  0.16, 13.09, 3.18, 0.33, 5.71, 5.71, 4.69, 1, 0.16, 5.53, 3.02, 2.68, 5.37,
  0.33, 0.33, 5.37, 5.71, 5.71, 1.33, 4.36, 3.02, 2.68, 5.71, 5.71, 5.71, 5.71,
  1.33, 4.02, 0.33, 5.71, 5.71, 5.03, 5.03, 5.03,
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

const NAV_ITEMS: Array<{
  id: AdminSection;
  label: string;
  icon: ComponentType<{ className?: string }>;
}> = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "users", label: "All Teachers / Users", icon: UsersRound },
  { id: "students", label: "All Learners / Students", icon: GraduationCap },
  {
    id: "credential_requests",
    label: "Credential Requests",
    icon: ClipboardCheck,
  },
  { id: "forms", label: "School Forms", icon: Files },
  { id: "analytics", label: "Analytics & Insights", icon: BarChart3 },
];

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

// Admin Analytics & Insights is defined locally in this file so admin.tsx
// does not depend on a separate AdminAnalyticsInsights.tsx module.
type AdminAnalyticsClass = {
  id: string;
  teacher_id: string;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  created_at?: string | null;
};

type AdminAnalyticsProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type?: string | null;
};

type TeacherTypeFilter = "all" | "class_adviser" | "subject_teacher";

function classLabel(row: AdminAnalyticsClass) {
  const grade = row.grade_level?.trim() || "Grade";
  const section = row.section?.trim();
  return section ? `${grade} · ${section}` : grade;
}

function teacherTypeLabel(value: string | null | undefined) {
  const normalized = normalize(value);

  if (normalized === "class_adviser") return "Class Adviser";
  if (normalized === "subject_teacher") return "Subject Teacher";
  return "Teacher";
}

function sortGradeLevels(values: string[]) {
  return [...values].sort((left, right) => {
    const leftNumber = Number(left.match(/\d+/)?.[0] ?? Number.MAX_SAFE_INTEGER);
    const rightNumber = Number(
      right.match(/\d+/)?.[0] ?? Number.MAX_SAFE_INTEGER,
    );

    if (leftNumber !== rightNumber) return leftNumber - rightNumber;
    return left.localeCompare(right);
  });
}

function AdminAnalyticsInsights() {
  const [schoolYear, setSchoolYear] = useState("all");
  const [gradeLevel, setGradeLevel] = useState("all");
  const [section, setSection] = useState("all");
  const [teacherType, setTeacherType] =
    useState<TeacherTypeFilter>("all");
  const [subject, setSubject] = useState("all");
  const [search, setSearch] = useState("");
  const [selectedClassId, setSelectedClassId] = useState<string | null>(null);

  const {
    data: directory,
    isLoading: directoryLoading,
    error: directoryError,
  } = useQuery({
    queryKey: ["admin-analytics-directory"],
    queryFn: async () => {
      const [classesResponse, profilesResponse] = await Promise.all([
        supabase
          .from("classes")
          .select(
            "id, teacher_id, subject, grade_level, section, school_year, created_at",
          )
          .order("school_year", { ascending: false })
          .order("grade_level", { ascending: true })
          .order("section", { ascending: true }),
        supabase
          .from("profiles")
          .select("id, full_name, email, teacher_type"),
      ]);

      if (classesResponse.error) throw classesResponse.error;
      if (profilesResponse.error) throw profilesResponse.error;

      return {
        classes: (classesResponse.data ?? []) as AdminAnalyticsClass[],
        profiles: (profilesResponse.data ?? []) as AdminAnalyticsProfile[],
      };
    },
  });

  const classes = directory?.classes ?? [];
  const profiles = directory?.profiles ?? [];

  const profileById = useMemo(
    () => new Map(profiles.map((profile) => [profile.id, profile])),
    [profiles],
  );

  const schoolYears = useMemo(
    () =>
      [...new Set(classes.map((row) => row.school_year).filter(Boolean) as string[])]
        .sort()
        .reverse(),
    [classes],
  );

  const gradeLevels = useMemo(
    () =>
      sortGradeLevels(
        [...new Set(classes.map((row) => row.grade_level).filter(Boolean) as string[])],
      ),
    [classes],
  );

  const sections = useMemo(
    () =>
      [...new Set(classes.map((row) => row.section).filter(Boolean) as string[])].sort(
        (a, b) => a.localeCompare(b),
      ),
    [classes],
  );

  const subjects = useMemo(
    () =>
      [...new Set(classes.map((row) => row.subject).filter(Boolean) as string[])].sort(
        (a, b) => a.localeCompare(b),
      ),
    [classes],
  );

  const filteredClasses = useMemo(() => {
    const searchTerm = normalize(search);

    return classes.filter((row) => {
      const teacher = profileById.get(row.teacher_id);
      const rowTeacherType = normalize(teacher?.teacher_type);

      if (schoolYear !== "all" && row.school_year !== schoolYear) return false;
      if (gradeLevel !== "all" && row.grade_level !== gradeLevel) return false;
      if (section !== "all" && row.section !== section) return false;
      if (subject !== "all" && row.subject !== subject) return false;
      if (teacherType !== "all" && rowTeacherType !== teacherType) return false;

      if (!searchTerm) return true;

      return [
        row.grade_level,
        row.section,
        row.subject,
        row.school_year,
        teacher?.full_name,
        teacher?.email,
        teacherTypeLabel(teacher?.teacher_type),
      ].some((value) => normalize(value).includes(searchTerm));
    });
  }, [
    classes,
    gradeLevel,
    profileById,
    schoolYear,
    search,
    section,
    subject,
    teacherType,
  ]);

  useEffect(() => {
    if (filteredClasses.length === 0) {
      setSelectedClassId(null);
      return;
    }

    if (
      !selectedClassId ||
      !filteredClasses.some((row) => row.id === selectedClassId)
    ) {
      setSelectedClassId(filteredClasses[0].id);
    }
  }, [filteredClasses, selectedClassId]);

  const selectedClass =
    filteredClasses.find((row) => row.id === selectedClassId) ??
    classes.find((row) => row.id === selectedClassId) ??
    null;

  const selectedTeacher = selectedClass
    ? profileById.get(selectedClass.teacher_id) ?? null
    : null;

  const {
    data: selectedStudents = [],
    isLoading: studentsLoading,
    error: studentsError,
  } = useQuery<StudentRow[]>({
    queryKey: ["admin-analytics-students", selectedClass?.id],
    enabled: Boolean(selectedClass?.id),
    queryFn: async () => {
      if (!selectedClass?.id) return [];

      const { data, error } = await supabase
        .from("students")
        .select("*")
        .eq("class_id", selectedClass.id)
        .order("sex", { ascending: false })
        .order("last_name", { ascending: true })
        .order("first_name", { ascending: true });

      if (error) throw error;
      return (data ?? []) as StudentRow[];
    },
  });

  const analyticsStudents = useMemo<AnalyticsStudentRow[]>(
    () =>
      selectedStudents.map((student) => {
        const normalizedSex = normalize(student.sex);

        return {
          id: student.id,
          class_id: student.class_id,
          teacher_id: student.teacher_id,
          last_name: student.last_name ?? "",
          first_name: student.first_name ?? "",
          middle_name: student.middle_name ?? null,
          sex:
            normalizedSex === "male"
              ? "male"
              : normalizedSex === "female"
                ? "female"
                : null,
          lrn: student.lrn ?? null,
          birthdate: student.birthdate ?? null,
          address: student.address ?? null,
          mother_name: student.mother_name ?? null,
          father_name: student.father_name ?? null,
          guardian: student.guardian ?? null,
          contact_number: student.contact_number ?? null,
        };
      }),
    [selectedStudents],
  );

  const clearFilters = () => {
    setSchoolYear("all");
    setGradeLevel("all");
    setSection("all");
    setTeacherType("all");
    setSubject("all");
    setSearch("");
  };

  if (directoryLoading) {
    return <AdminAnalyticsInsightsSkeleton />;
  }

  if (directoryError) {
    return (
      <section className="rounded-2xl border border-rose-200 bg-rose-50 p-6 text-sm text-rose-700">
        {directoryError instanceof Error
          ? directoryError.message
          : "Unable to load Admin Analytics & Insights."}
      </section>
    );
  }

  return (
    <section className="space-y-5">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-12 shrink-0 place-items-center rounded-2xl bg-rose-50 text-[var(--admin-primary)]">
            <BarChart3 className="size-6" />
          </div>
          <div className="min-w-0">
            <h1 className="text-2xl font-bold text-[var(--admin-text-heading)] sm:text-3xl">
              Analytics &amp; Insights
            </h1>
            <p className="mt-1 max-w-3xl text-sm text-[var(--admin-text-muted)]">
              View academic analytics across all classes created by Class
              Advisers and Subject Teachers.
            </p>
          </div>
        </div>

        <div className="flex max-w-sm items-start gap-3 rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-emerald-800">
          <Crown className="mt-0.5 size-5 shrink-0" />
          <div>
            <div className="text-sm font-bold">Admin Access: All Classes</div>
            <div className="mt-0.5 text-xs leading-5 text-emerald-700">
              You can view analytics for all classes across all grade levels and
              sections.
            </div>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-4 shadow-[var(--admin-shadow-panel)]">
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-6">
          <FilterField label="School Year">
            <select
              value={schoolYear}
              onChange={(event) => setSchoolYear(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All School Years</option>
              {schoolYears.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Grade Level">
            <select
              value={gradeLevel}
              onChange={(event) => setGradeLevel(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Grade Levels</option>
              {gradeLevels.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Section">
            <select
              value={section}
              onChange={(event) => setSection(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Sections</option>
              {sections.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Teacher Type">
            <select
              value={teacherType}
              onChange={(event) =>
                setTeacherType(event.target.value as TeacherTypeFilter)
              }
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Teachers</option>
              <option value="class_adviser">Class Adviser</option>
              <option value="subject_teacher">Subject Teacher</option>
            </select>
          </FilterField>

          <FilterField label="Subject">
            <select
              value={subject}
              onChange={(event) => setSubject(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Subjects</option>
              {subjects.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Search">
            <div className="relative">
              <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
              <Input
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                placeholder="Search class or teacher..."
                className="h-10 rounded-xl pl-9 text-xs"
              />
            </div>
          </FilterField>
        </div>

        {(schoolYear !== "all" ||
          gradeLevel !== "all" ||
          section !== "all" ||
          teacherType !== "all" ||
          subject !== "all" ||
          search.trim()) && (
          <div className="mt-3 flex justify-end">
            <Button
              type="button"
              variant="outline"
              onClick={clearFilters}
              className="h-8 rounded-lg px-3 text-xs"
            >
              Clear filters
            </Button>
          </div>
        )}
      </div>

      <div className="grid gap-4 2xl:grid-cols-[minmax(420px,0.9fr)_minmax(0,1.35fr)]">
        <div className="overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-[var(--admin-shadow-panel)]">
          <div className="flex items-start justify-between gap-3 border-b border-[var(--admin-divider)] px-4 py-4">
            <div className="flex items-start gap-3">
              <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-rose-50 text-[var(--admin-primary)]">
                <BookOpen className="size-4" />
              </div>
              <div>
                <h2 className="font-bold text-[var(--admin-text-heading)]">
                  Classes (All Classes)
                </h2>
                <p className="mt-0.5 text-xs text-[var(--admin-text-muted)]">
                  List of all classes created by Class Advisers and Subject
                  Teachers.
                </p>
              </div>
            </div>
            <span className="shrink-0 rounded-full bg-rose-50 px-2.5 py-1 text-[10px] font-bold text-[var(--admin-primary)]">
              {filteredClasses.length} class
              {filteredClasses.length === 1 ? "" : "es"}
            </span>
          </div>

          {filteredClasses.length === 0 ? (
            <div className="px-6 py-16 text-center">
              <GraduationCap className="mx-auto size-8 text-[var(--admin-text-muted)]" />
              <p className="mt-3 text-sm font-semibold text-[var(--admin-text-heading)]">
                No classes found
              </p>
              <p className="mt-1 text-xs text-[var(--admin-text-muted)]">
                Change the filters or search term.
              </p>
            </div>
          ) : (
            <div className="max-h-[860px] overflow-auto">
              <table className="w-full min-w-[680px] border-collapse text-left">
                <thead className="sticky top-0 z-10 bg-[var(--admin-surface-soft)]">
                  <tr className="border-b border-[var(--admin-divider)] text-[9px] font-bold uppercase tracking-wide text-[var(--admin-text-muted)]">
                    <th className="px-3 py-3">#</th>
                    <th className="px-3 py-3">Class</th>
                    <th className="px-3 py-3">Teacher</th>
                    <th className="px-3 py-3">Subject</th>
                    <th className="px-3 py-3">Type</th>
                    <th className="px-3 py-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredClasses.map((row, index) => {
                    const teacher = profileById.get(row.teacher_id);
                    const selected = row.id === selectedClassId;
                    const type = teacherTypeLabel(teacher?.teacher_type);

                    return (
                      <tr
                        key={row.id}
                        className={`border-b border-[var(--admin-divider)] text-xs transition last:border-b-0 ${
                          selected
                            ? "bg-rose-50/60"
                            : "hover:bg-[var(--admin-nav-hover)]"
                        }`}
                      >
                        <td className="px-3 py-3 text-[var(--admin-text-muted)]">
                          {index + 1}
                        </td>
                        <td className="px-3 py-3">
                          <div className="font-bold text-[var(--admin-text-heading)]">
                            {classLabel(row)}
                          </div>
                          <div className="mt-0.5 text-[10px] text-[var(--admin-text-muted)]">
                            SY {row.school_year || "—"}
                          </div>
                        </td>
                        <td className="px-3 py-3">
                          <div className="max-w-[150px] truncate font-medium text-[var(--admin-text-heading)]">
                            {teacher?.full_name || "Unknown Teacher"}
                          </div>
                        </td>
                        <td className="px-3 py-3 text-[var(--admin-color-735e55)]">
                          {row.subject || "—"}
                        </td>
                        <td className="px-3 py-3">
                          <span
                            className={`inline-flex rounded-full px-2.5 py-1 text-[9px] font-semibold ${
                              normalize(teacher?.teacher_type) === "class_adviser"
                                ? "bg-emerald-100 text-emerald-700"
                                : normalize(teacher?.teacher_type) ===
                                    "subject_teacher"
                                  ? "bg-sky-100 text-sky-700"
                                  : "bg-slate-100 text-slate-700"
                            }`}
                          >
                            {type}
                          </span>
                        </td>
                        <td className="px-3 py-3 text-right">
                          <Button
                            type="button"
                            size="sm"
                            onClick={() => setSelectedClassId(row.id)}
                            className="h-8 gap-1.5 rounded-lg bg-[var(--admin-primary)] px-3 text-[10px] hover:bg-[var(--admin-primary-hover)]"
                          >
                            <BarChart3 className="size-3.5" />
                            {selected ? "Viewing" : "Open Analytics"}
                          </Button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div className="min-w-0 overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-[var(--admin-shadow-panel)]">
          <div className="flex flex-col gap-3 border-b border-[var(--admin-divider)] px-4 py-4 sm:flex-row sm:items-start sm:justify-between">
            <div className="flex min-w-0 items-start gap-3">
              <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-rose-50 text-[var(--admin-primary)]">
                <BarChart3 className="size-4" />
              </div>
              <div className="min-w-0">
                <h2 className="font-bold text-[var(--admin-text-heading)]">
                  Class Analytics Preview
                </h2>
                <p className="mt-0.5 text-xs text-[var(--admin-text-muted)]">
                  Select a class and review the same analytics available to its
                  Class Adviser or Subject Teacher.
                </p>
              </div>
            </div>

            {selectedClass ? (
              <div className="shrink-0 rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2 text-right">
                <div className="text-xs font-bold text-emerald-800">
                  {classLabel(selectedClass)}
                </div>
                <div className="mt-0.5 text-[10px] text-emerald-700">
                  {selectedClass.subject || "No Subject"}
                </div>
              </div>
            ) : null}
          </div>

          <div className="p-3 sm:p-4">
            {!selectedClass ? (
              <div className="grid min-h-[420px] place-items-center rounded-xl border border-dashed border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-8 text-center">
                <div>
                  <ShieldCheck className="mx-auto size-9 text-[var(--admin-text-muted)]" />
                  <p className="mt-3 text-sm font-semibold text-[var(--admin-text-heading)]">
                    Select a class
                  </p>
                  <p className="mt-1 max-w-sm text-xs leading-5 text-[var(--admin-text-muted)]">
                    Choose any class from the list to view its term analytics,
                    proficiency distribution, and forecast panel.
                  </p>
                </div>
              </div>
            ) : studentsError ? (
              <div className="rounded-xl border border-rose-200 bg-rose-50 p-5 text-sm text-rose-700">
                {studentsError instanceof Error
                  ? studentsError.message
                  : "Unable to load learners for this class."}
              </div>
            ) : studentsLoading ? (
              <div className="space-y-3">
                <Skeleton className="h-16 w-full rounded-xl" />
                <Skeleton className="h-40 w-full rounded-xl" />
                <Skeleton className="h-72 w-full rounded-xl" />
              </div>
            ) : (
              <AnalyticsInsightsPanel
                classId={selectedClass.id}
                students={analyticsStudents}
                subject={selectedClass.subject || ""}
                initialTerm="1"
                forecastData={null}
                adminView
                classLabel={classLabel(selectedClass)}
                teacherLabel={
                  selectedTeacher?.full_name ||
                  selectedTeacher?.email ||
                  "Unknown Teacher"
                }
              />
            )}
          </div>
        </div>
      </div>

      <div className="flex items-start gap-2 rounded-xl border border-sky-200 bg-sky-50 px-4 py-3 text-xs leading-5 text-sky-800">
        <UserRound className="mt-0.5 size-4 shrink-0" />
        <span>
          Admin Analytics is read-only. It reuses the existing class analytics
          calculations so the Admin sees the same actual grades, proficiency
          distribution, and model forecast output as the teacher.
        </span>
      </div>
    </section>
  );
}

function FilterField({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}) {
  return (
    <label className="block">
      <span className="mb-1.5 block text-[10px] font-bold uppercase tracking-wide text-[var(--admin-text-muted)]">
        {label}
      </span>
      {children}
    </label>
  );
}

function AdminAnalyticsInsightsSkeleton() {
  return (
    <section className="space-y-5">
      <div className="flex items-start gap-3">
        <Skeleton className="size-12 rounded-2xl" />
        <div className="space-y-2">
          <Skeleton className="h-7 w-64" />
          <Skeleton className="h-4 w-[34rem] max-w-full" />
        </div>
      </div>

      <div className="grid gap-3 rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-4 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-6">
        {Array.from({ length: 6 }).map((_, index) => (
          <Skeleton key={index} className="h-10 rounded-xl" />
        ))}
      </div>

      <div className="grid gap-4 2xl:grid-cols-[minmax(420px,0.9fr)_minmax(0,1.35fr)]">
        <Skeleton className="h-[620px] rounded-2xl" />
        <Skeleton className="h-[620px] rounded-2xl" />
      </div>
    </section>
  );
}


function adminUserRoleLabel(
  profile: ProfileRow,
  roleByUserId: Map<string, string>,
) {
  const accessRole = normalize(roleByUserId.get(profile.id));
  const teacherType = normalize(profile.teacher_type);

  if (accessRole === "admin") return "Admin";
  if (teacherType === "class_adviser") return "Class Adviser";
  if (teacherType === "subject_teacher") return "Subject Teacher";

  return "Teacher";
}

function sf6GradeNumber(value: string | null | undefined) {
  return value?.match(/\d+/)?.[0] || "";
}

function addSf6Counts(values: Sf6Count[]): Sf6Count {
  return values.reduce(
    (total, value) => ({
      male: total.male + value.male,
      female: total.female + value.female,
      total: total.total + value.total,
    }),
    { ...EMPTY_SF6_COUNT },
  );
}

function sf6TableBorder({
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
    top: top ? SF6_MEDIUM_BORDER : SF6_THIN_BORDER,
    bottom: bottom ? SF6_MEDIUM_BORDER : SF6_THIN_BORDER,
    left: left ? SF6_MEDIUM_BORDER : SF6_THIN_BORDER,
    right: right ? SF6_MEDIUM_BORDER : SF6_THIN_BORDER,
  };
}

function styleSf6ExcelCell(
  cell: Cell,
  options: {
    font?: Partial<Font>;
    alignment?: Partial<Alignment>;
    border?: Partial<Borders>;
    fill?: Fill;
    numberFormat?: string;
  } = {},
) {
  cell.font = { ...SF6_BODY_FONT, ...options.font };
  cell.alignment = options.alignment || SF6_CENTERED;
  cell.border = options.border || {};
  cell.fill = options.fill || SF6_NO_FILL;
  if (options.numberFormat) cell.numFmt = options.numberFormat;
}

function mergeAndSetSf6ExcelCell(
  worksheet: Worksheet,
  range: string,
  value: string | number,
  options: Parameters<typeof styleSf6ExcelCell>[1] = {},
) {
  const [firstCell, lastCell] = range.split(":");
  if (lastCell && firstCell !== lastCell) {
    worksheet.mergeCells(range);
  }

  const cell = worksheet.getCell(firstCell);
  cell.value = value === "" ? null : value;
  styleSf6ExcelCell(cell, options);
}

function sf6RowRange(range: string, row: number) {
  const [start, end] = range.split(":");
  return end ? `${start}${row}:${end}${row}` : `${start}${row}`;
}

function setSf6ExcelCountGroup(
  worksheet: Worksheet,
  row: number,
  ranges: readonly string[],
  value: Sf6Count,
  strong = false,
  options: {
    top?: boolean;
    bottom?: boolean;
    right?: boolean;
  } = {},
) {
  [value.male, value.female, value.total].forEach((count, index) => {
    mergeAndSetSf6ExcelCell(worksheet, sf6RowRange(ranges[index], row), count, {
      font: { size: 7, bold: strong },
      numberFormat: "0",
      border: sf6TableBorder({
        ...options,
        right: options.right && index === 2,
      }),
    });
  });
}

async function imageSourceToDataUrl(source: string) {
  const response = await fetch(source);
  if (!response.ok) {
    throw new Error(`Unable to load image: ${response.status}`);
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

async function getSf6ExcelSeal() {
  for (const source of [SF6_DEPED_SEAL_URL, logo]) {
    try {
      const dataUrl = await imageSourceToDataUrl(source);
      if (dataUrl) return dataUrl;
    } catch {
      // Continue with the next available logo source.
    }
  }

  return "";
}

async function buildSf6ExcelWorkbook(options: Sf6ExportOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject =
    "Summarized Report on Promotion and Learning Progress & Achievement";
  workbook.title = "School Form 6 (SF6)";

  const worksheet = workbook.addWorksheet("report1", {
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
      paperSize: 9 as PaperSize,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: false,
      verticalCentered: false,
      pageOrder: "overThenDown",
      blackAndWhite: false,
      printArea: "A1:AH20",
      margins: {
        left: 0.28,
        right: 0.28,
        top: 0.28,
        bottom: 0.28,
        header: 0.51,
        footer: 0.51,
      },
    },
  });

  worksheet.properties.defaultRowHeight = 20;
  SF6_COLUMN_WIDTHS.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  mergeAndSetSf6ExcelCell(worksheet, "A1:B3", "", { border: {} });
  const sealDataUrl = await getSf6ExcelSeal();
  if (sealDataUrl) {
    const logoId = workbook.addImage({
      base64: sealDataUrl,
      extension: "png",
    });
    worksheet.addImage(logoId, {
      tl: { col: 0.08, row: 0.04 },
      ext: { width: 112, height: 112 },
      editAs: "oneCell",
    });
  }

  mergeAndSetSf6ExcelCell(
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
    mergeAndSetSf6ExcelCell(worksheet, range, value, {
      font: { size: 9 },
      alignment: {
        horizontal: isLabel ? "right" : "left",
        vertical: "middle",
        wrapText: true,
      },
      border: isLabel ? {} : sf6TableBorder(),
    });
  });

  mergeAndSetSf6ExcelCell(worksheet, "A5:D6", "SUMMARY TABLE", {
    font: { size: 9 },
    border: sf6TableBorder({ top: true, left: true }),
  });

  SF6_ALL_GROUPS.forEach((group, groupIndex) => {
    mergeAndSetSf6ExcelCell(worksheet, group.headerRange, group.label, {
      font: {
        size: groupIndex === SF6_ALL_GROUPS.length - 1 ? 11 : 9,
      },
      border: sf6TableBorder({
        top: true,
        right: groupIndex === SF6_ALL_GROUPS.length - 1,
      }),
    });

    ["Male", "Female", "Total"].forEach((sexLabel, sexIndex) => {
      mergeAndSetSf6ExcelCell(
        worksheet,
        sf6RowRange(group.countRanges[sexIndex], 6),
        sexLabel,
        {
          font: { size: 7 },
          border: sf6TableBorder({
            right: groupIndex === SF6_ALL_GROUPS.length - 1 && sexIndex === 2,
          }),
        },
      );
    });
  });

  options.statusRows.forEach((status, statusIndex) => {
    const rowNumber = 7 + statusIndex;
    mergeAndSetSf6ExcelCell(
      worksheet,
      `A${rowNumber}:D${rowNumber}`,
      status.label,
      {
        font: { size: 9 },
        alignment: { horizontal: "left", vertical: "middle" },
        border: sf6TableBorder({ left: true }),
      },
    );

    status.values.forEach((value, gradeIndex) => {
      setSf6ExcelCountGroup(
        worksheet,
        rowNumber,
        SF6_GRADE_GROUPS[gradeIndex].countRanges,
        value,
      );
    });
    setSf6ExcelCountGroup(
      worksheet,
      rowNumber,
      SF6_TOTAL_GROUP.countRanges,
      status.total,
      true,
      {
        right: true,
      },
    );
  });

  mergeAndSetSf6ExcelCell(
    worksheet,
    "A10:D10",
    "LEARNING PROGRESS AND\nACHIEVEMENT",
    {
      font: { size: 5, bold: true },
      alignment: {
        horizontal: "left",
        vertical: "middle",
        wrapText: true,
      },
      border: sf6TableBorder({ left: true }),
    },
  );

  SF6_ALL_GROUPS.forEach((group, groupIndex) => {
    ["Male", "Female", "Total"].forEach((sexLabel, sexIndex) => {
      mergeAndSetSf6ExcelCell(
        worksheet,
        sf6RowRange(group.countRanges[sexIndex], 10),
        sexLabel,
        {
          font: { size: 7 },
          border: sf6TableBorder({
            right: groupIndex === SF6_ALL_GROUPS.length - 1 && sexIndex === 2,
          }),
        },
      );
    });
  });

  SF6_PROGRESS_BANDS.forEach((band, bandIndex) => {
    const rowNumber = 11 + bandIndex;
    const values = options.bandValues[bandIndex];
    const separator = bandIndex === 2 || bandIndex === 4 ? "\n" : " ";
    mergeAndSetSf6ExcelCell(
      worksheet,
      `A${rowNumber}:D${rowNumber}`,
      `${band.label}${separator}${band.detail}`,
      {
        font: { size: 5 },
        alignment: {
          horizontal: "left",
          vertical: "middle",
          wrapText: true,
        },
        border: sf6TableBorder({ left: true }),
      },
    );

    values.forEach((value, gradeIndex) => {
      setSf6ExcelCountGroup(
        worksheet,
        rowNumber,
        SF6_GRADE_GROUPS[gradeIndex].countRanges,
        value,
      );
    });
    setSf6ExcelCountGroup(
      worksheet,
      rowNumber,
      SF6_TOTAL_GROUP.countRanges,
      addSf6Counts(values),
      true,
      { right: true },
    );
  });

  mergeAndSetSf6ExcelCell(worksheet, "A16:D16", "TOTAL", {
    font: { size: 9, bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: sf6TableBorder({
      top: true,
      bottom: true,
      left: true,
    }),
  });
  options.enrolledByGrade.forEach((value, gradeIndex) => {
    setSf6ExcelCountGroup(
      worksheet,
      16,
      SF6_GRADE_GROUPS[gradeIndex].countRanges,
      value,
      true,
      {
        top: true,
        bottom: true,
      },
    );
  });
  setSf6ExcelCountGroup(
    worksheet,
    16,
    SF6_TOTAL_GROUP.countRanges,
    addSf6Counts(options.enrolledByGrade),
    true,
    { top: true, bottom: true, right: true },
  );

  mergeAndSetSf6ExcelCell(worksheet, "A17:E17", "Prepared and Submitted by:", {
    font: { size: 7 },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });

  const signatures = [
    ["A18:G18", options.schoolHead.toUpperCase()],
    ["L18:S18", ""],
    ["V18:AA18", ""],
    ["AC18:AH18", ""],
  ] as const;
  signatures.forEach(([range, value]) => {
    mergeAndSetSf6ExcelCell(worksheet, range, value, {
      font: { size: 8 },
      alignment: {
        horizontal: "center",
        vertical: "bottom",
      },
      border: { bottom: SF6_MEDIUM_BORDER },
    });
  });

  const signatureLabels = [
    ["A19:G19", "(Signature of School Head/SCC Chair)"],
    ["L19:S19", "SCC-Vice Chair (Curriculum)"],
    ["V19:AA19", "SCC Member"],
    ["AC19:AH19", "SCC-Vice Chair (Generated thru LIS)"],
  ] as const;
  signatureLabels.forEach(([range, value]) => {
    mergeAndSetSf6ExcelCell(worksheet, range, value, {
      font: { size: 7 },
      alignment: {
        horizontal: "center",
        vertical: "top",
        wrapText: true,
      },
      border: {},
    });
  });

  mergeAndSetSf6ExcelCell(
    worksheet,
    "B20:I20",
    "(Additional slots may be added for SCC members.)",
    {
      font: { size: 7 },
      alignment: {
        horizontal: "left",
        vertical: "middle",
      },
      border: {},
    },
  );

  [
    50, 20, 4, 20, 25, 25, 25, 25, 25, 30, 30, 30, 30, 30, 30, 25, 20, 20, 20,
    20,
  ].forEach((height, index) => {
    worksheet.getRow(index + 1).height = height;
  });

  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.value !== null) {
        cell.font = {
          ...cell.font,
          name: "Arial",
        };
      }
    });
  });

  return workbook;
}

type Sf5ExcelLearnerRow = {
  lrn: string;
  name: string;
  generalAverage: number;
  action: string;
  didNotMeet: string;
};

type Sf5ExcelCount = {
  male: number;
  female: number;
  total: number;
};

type Sf5ExcelExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  curriculum: string;
  teacherName: string;
  schoolHead: string;
  learners: Sf5ExcelLearnerRow[];
  maleTotal: number;
  femaleTotal: number;
  promoted: Sf5ExcelCount;
  conditional: Sf5ExcelCount;
  retained: Sf5ExcelCount;
  gradeBands: Array<Sf5ExcelCount & { label: string }>;
};

const SF5_EXCEL_BLUE = "FFFFFFFF";
const SF5_EXCEL_WHITE = "FF000000";
const SF5_EXCEL_BLACK = "FF000000";
const SF5_EXCEL_LIGHT_GRAY = "FFFFFFFF";
const SF5_EXCEL_THIN_BORDER = {
  style: "thin" as const,
  color: { argb: SF5_EXCEL_BLACK },
};

function sf5ExcelBorder(): Partial<Borders> {
  return {
    top: SF5_EXCEL_THIN_BORDER,
    bottom: SF5_EXCEL_THIN_BORDER,
    left: SF5_EXCEL_THIN_BORDER,
    right: SF5_EXCEL_THIN_BORDER,
  };
}

function styleSf5ExcelCell(
  cell: Cell,
  options: {
    font?: Partial<Font>;
    alignment?: Partial<Alignment>;
    fillColor?: string;
    border?: Partial<Borders>;
    numberFormat?: string;
  } = {},
) {
  cell.font = {
    name: "Arial",
    size: 8,
    color: { argb: SF5_EXCEL_BLACK },
    ...options.font,
  };
  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
    ...options.alignment,
  };
  cell.border = options.border ?? sf5ExcelBorder();
  if (options.fillColor) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: options.fillColor },
    };
  }
  if (options.numberFormat) cell.numFmt = options.numberFormat;
}

function mergeAndSetSf5ExcelCell(
  worksheet: Worksheet,
  range: string,
  value: string | number | null,
  options: Parameters<typeof styleSf5ExcelCell>[1] = {},
) {
  const [firstCell, lastCell] = range.split(":");
  if (lastCell && firstCell !== lastCell) {
    worksheet.mergeCells(range);
  }

  const cell = worksheet.getCell(firstCell);
  cell.value = value === "" ? null : value;
  styleSf5ExcelCell(cell, options);
}

function setSf5ExcelCountRow(
  worksheet: Worksheet,
  row: number,
  label: string,
  counts: Sf5ExcelCount,
) {
  mergeAndSetSf5ExcelCell(worksheet, `H${row}:I${row}`, label, {
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
  });
  mergeAndSetSf5ExcelCell(worksheet, `J${row}`, counts.male, {
    numberFormat: "0",
  });
  mergeAndSetSf5ExcelCell(worksheet, `K${row}`, counts.female, {
    numberFormat: "0",
  });
  mergeAndSetSf5ExcelCell(worksheet, `L${row}`, counts.total, {
    font: { bold: true },
    numberFormat: "0",
  });
}

async function buildSf5ExcelWorkbook(options: Sf5ExcelExportOptions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.schoolName || "School Forms System";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject =
    "School Form 5 Report on Promotion and Learning Progress & Achievement";
  workbook.title = "School Form 5 (SF5)";

  const worksheet = workbook.addWorksheet("SF5", {
    views: [{ showGridLines: true, zoomScale: 90 }],
    pageSetup: {
      orientation: "landscape",
      paperSize: 5 as PaperSize,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: true,
      verticalCentered: false,
      margins: {
        left: 0.25,
        right: 0.25,
        top: 0.25,
        bottom: 0.25,
        header: 0.3,
        footer: 0.3,
      },
    },
  });

  worksheet.properties.defaultRowHeight = 18;
  [14, 16, 16, 16, 13, 14, 22, 18, 18, 9, 9, 9].forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  mergeAndSetSf5ExcelCell(
    worksheet,
    "A1:L1",
    "School Form 5 (SF 5) Report on Promotion and Learning Progress & Achievement",
    {
      font: { size: 14, bold: true },
      border: {},
    },
  );
  mergeAndSetSf5ExcelCell(
    worksheet,
    "A2:L2",
    "Revised to conform with the Instructions of DepEd Order 8, s. 2015",
    {
      font: { size: 8, italic: true },
      border: {},
    },
  );

  const infoRows = [
    ["A4:B4", "Region", options.region],
    ["C4:E4", "Division", options.division],
    ["F4:H4", "District", options.district],
    ["I4:J4", "School ID", options.schoolId],
    ["K4:L4", "School Year", options.schoolYear],
    ["A5:E5", "School Name", options.schoolName],
    ["F5:H5", "Grade Level", options.gradeLevel],
    ["I5:J5", "Section", options.section],
    ["K5:L5", "Curriculum", options.curriculum],
  ] as const;

  infoRows.forEach(([range, label, value]) => {
    mergeAndSetSf5ExcelCell(worksheet, range, `${label}: ${value || ""}`, {
      alignment: { horizontal: "left", vertical: "middle", wrapText: true },
      font: { size: 8 },
    });
  });

  const learnerHeaderRow = 7;
  mergeAndSetSf5ExcelCell(worksheet, `A${learnerHeaderRow}`, "LRN", {
    font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
    fillColor: SF5_EXCEL_BLUE,
  });
  mergeAndSetSf5ExcelCell(
    worksheet,
    `B${learnerHeaderRow}:D${learnerHeaderRow}`,
    "LEARNER'S NAME\n(Last Name, First Name, Middle Name)",
    {
      font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
      fillColor: SF5_EXCEL_BLUE,
    },
  );
  mergeAndSetSf5ExcelCell(
    worksheet,
    `E${learnerHeaderRow}`,
    "GENERAL AVERAGE\n(Whole numbers)",
    {
      font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
      fillColor: SF5_EXCEL_BLUE,
    },
  );
  mergeAndSetSf5ExcelCell(worksheet, `F${learnerHeaderRow}`, "ACTION TAKEN", {
    font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
    fillColor: SF5_EXCEL_BLUE,
  });
  mergeAndSetSf5ExcelCell(
    worksheet,
    `G${learnerHeaderRow}`,
    "Did Not Meet Expectations\nIncomplete / Failed Learning Areas",
    {
      font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
      fillColor: SF5_EXCEL_BLUE,
    },
  );

  mergeAndSetSf5ExcelCell(worksheet, "H7:I7", "SUMMARY TABLE", {
    font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
    fillColor: SF5_EXCEL_BLUE,
  });
  ["MALE", "FEMALE", "TOTAL"].forEach((label, index) => {
    mergeAndSetSf5ExcelCell(
      worksheet,
      `${String.fromCharCode(74 + index)}7`,
      label,
      {
        font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
        fillColor: SF5_EXCEL_BLUE,
      },
    );
  });

  const learnerRows = options.learners.slice(0, 25);
  const firstLearnerRow = 8;
  const minimumLearnerRows = 12;
  const totalLearnerRows = Math.max(minimumLearnerRows, learnerRows.length);
  for (let index = 0; index < totalLearnerRows; index += 1) {
    const rowNumber = firstLearnerRow + index;
    const learner = learnerRows[index];
    mergeAndSetSf5ExcelCell(worksheet, `A${rowNumber}`, learner?.lrn || "", {
      alignment: { horizontal: "left", vertical: "middle", wrapText: true },
    });
    mergeAndSetSf5ExcelCell(
      worksheet,
      `B${rowNumber}:D${rowNumber}`,
      learner?.name || "",
      {
        alignment: { horizontal: "left", vertical: "middle", wrapText: true },
      },
    );
    mergeAndSetSf5ExcelCell(
      worksheet,
      `E${rowNumber}`,
      learner?.generalAverage ? Math.round(learner.generalAverage) : "",
      { numberFormat: "0" },
    );
    mergeAndSetSf5ExcelCell(worksheet, `F${rowNumber}`, learner?.action || "", {
      font: { bold: Boolean(learner?.action) },
    });
    mergeAndSetSf5ExcelCell(
      worksheet,
      `G${rowNumber}`,
      learner?.didNotMeet || "",
      {
        alignment: { horizontal: "left", vertical: "middle", wrapText: true },
      },
    );
    worksheet.getRow(rowNumber).height = 18;
  }

  setSf5ExcelCountRow(worksheet, 8, "PROMOTED", options.promoted);
  setSf5ExcelCountRow(worksheet, 9, "*Conditional", options.conditional);
  setSf5ExcelCountRow(worksheet, 10, "RETAINED", options.retained);

  mergeAndSetSf5ExcelCell(
    worksheet,
    "H12:L12",
    "LEARNING PROGRESS AND ACHIEVEMENT\n(Based on Learners' General Average)",
    {
      font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
      fillColor: SF5_EXCEL_BLUE,
    },
  );
  mergeAndSetSf5ExcelCell(worksheet, "H13:I13", "Descriptors & Grading Scale", {
    font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
    fillColor: SF5_EXCEL_BLUE,
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
  });
  ["MALE", "FEMALE", "TOTAL"].forEach((label, index) => {
    mergeAndSetSf5ExcelCell(
      worksheet,
      `${String.fromCharCode(74 + index)}13`,
      label,
      {
        font: { bold: true, color: { argb: SF5_EXCEL_WHITE } },
        fillColor: SF5_EXCEL_BLUE,
      },
    );
  });
  options.gradeBands.forEach((band, index) => {
    setSf5ExcelCountRow(worksheet, 14 + index, band.label, band);
  });

  const totalRow = firstLearnerRow + totalLearnerRows;
  mergeAndSetSf5ExcelCell(
    worksheet,
    `A${totalRow}:G${totalRow}`,
    `TOTAL MALE: ${options.maleTotal}`,
    {
      font: { bold: true },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );
  mergeAndSetSf5ExcelCell(
    worksheet,
    `A${totalRow + 1}:G${totalRow + 1}`,
    `TOTAL FEMALE: ${options.femaleTotal}`,
    {
      font: { bold: true },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );
  mergeAndSetSf5ExcelCell(
    worksheet,
    `A${totalRow + 2}:G${totalRow + 2}`,
    `COMBINED TOTAL: ${options.learners.length}`,
    {
      font: { bold: true },
      alignment: { horizontal: "left", vertical: "middle" },
    },
  );

  const signatureStartRow = 21;
  const signatures = [
    [
      "H",
      "I",
      "PREPARED BY:",
      options.teacherName || "Class Adviser",
      "Class Adviser",
    ],
    [
      "J",
      "L",
      "CERTIFIED CORRECT & SUBMITTED:",
      options.schoolHead || "School Head",
      "School Head",
    ],
  ] as const;
  signatures.forEach(([startColumn, endColumn, label, name, role], index) => {
    const labelRow = signatureStartRow + index * 4;
    mergeAndSetSf5ExcelCell(
      worksheet,
      `${startColumn}${labelRow}:${endColumn}${labelRow}`,
      label,
      {
        font: { bold: true },
        alignment: { horizontal: "left", vertical: "middle" },
        border: {},
      },
    );
    mergeAndSetSf5ExcelCell(
      worksheet,
      `${startColumn}${labelRow + 2}:${endColumn}${labelRow + 2}`,
      name,
      {
        font: { bold: true },
        alignment: { horizontal: "center", vertical: "bottom", wrapText: true },
        border: { bottom: SF5_EXCEL_THIN_BORDER },
      },
    );
    mergeAndSetSf5ExcelCell(
      worksheet,
      `${startColumn}${labelRow + 3}:${endColumn}${labelRow + 3}`,
      role,
      {
        alignment: { horizontal: "center", vertical: "top", wrapText: true },
        border: {},
      },
    );
  });

  mergeAndSetSf5ExcelCell(worksheet, "H29:L29", "GUIDELINES:", {
    font: { bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: {},
  });
  [
    "1. Do not include Dropouts and Transferred Out.",
    "2. Prepare using the learner's Form 138.",
    "3. *Conditional means failed in not more than two learning areas.",
    "4. Did Not Meet Expectations means the learner failed the learning area.",
    "5. Validation and submission are under the SDO.",
  ].forEach((guide, index) => {
    mergeAndSetSf5ExcelCell(worksheet, `H${30 + index}:L${30 + index}`, guide, {
      alignment: { horizontal: "left", vertical: "middle", wrapText: true },
      border: {},
      font: { size: 7 },
    });
  });

  for (let row = 1; row <= 35; row += 1) {
    worksheet.getRow(row).height = row === 1 ? 24 : row === 2 ? 16 : 18;
  }

  worksheet.getRow(7).height = 32;
  worksheet.getRow(12).height = 32;
  worksheet.getRow(13).height = 28;
  worksheet.views = [{ showGridLines: true, zoomScale: 90 }];
  worksheet.pageSetup.printArea = `A1:L${Math.max(35, totalRow + 2)}`;

  [21, 22, 23, 24, 25, 26, 27, 28].forEach((rowNumber) => {
    ["H", "I", "J", "K", "L"].forEach((column) => {
      const cell = worksheet.getCell(`${column}${rowNumber}`);
      if (!cell.fill || (cell.fill as Fill).type !== "pattern") {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: SF5_EXCEL_LIGHT_GRAY },
        };
      }
    });
  });

  return workbook;
}

async function downloadSf5ExcelWorkbook(
  fileName: string,
  options: Sf5ExcelExportOptions,
) {
  const workbook = await buildSf5ExcelWorkbook(options);
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

async function downloadSf6ExcelWorkbook(
  fileName: string,
  options: Sf6ExportOptions,
) {
  const workbook = await buildSf6ExcelWorkbook(options);
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

type Sf9NewExcelExportOptions = {
  schoolName: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  municipality?: string;
  track?: string;
  learner: StudentRow | null;
  teacherName: string;
  schoolHead: string;
  subjects: Array<{
    subject: string;
    term1: number | null;
    term2: number | null;
    term3: number | null;
    final: number | null;
    remarks: string;
    units?: number | null;
    assignedTerm?: AdminSf9AssignedTerm | null;
    group?: "core" | "applied" | null;
  }>;
  generalAverage: number | null;
};

type Sf1ExcelExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  teacherName: string;
  schoolHead: string;
  male: Array<StudentRow & { age: string }>;
  female: Array<StudentRow & { age: string }>;
};

type Sf8ExcelExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  learners: StudentRow[];
};

type Sf10ExcelExportOptions = {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  learner: StudentRow | null;
  teacherName: string;
  schoolHead: string;
  subjects: Array<{
    subject: string;
    term1: number | null;
    term2: number | null;
    term3: number | null;
    term4: number | null;
    final: number | null;
    remarks: string;
  }>;
  generalAverage: number | null;
};

async function downloadSf10ExcelWorkbook(
  fileName: string,
  options: Sf10ExcelExportOptions,
) {
  const response = await fetch("/templates/SF10-admin.xlsx");
  if (!response.ok) {
    throw new Error(`Unable to load SF10 Excel template (${response.status}).`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const front = workbook.getWorksheet("Front") || workbook.worksheets[0];
  const back = workbook.getWorksheet("Back") || workbook.worksheets[2];
  if (!front || !back) {
    throw new Error(
      "The SF10 Excel template must contain Front and Back sheets.",
    );
  }

  const learner = options.learner;
  const rounded = (value: number | null) =>
    typeof value === "number" ? Math.round(value) : null;
  const set = (
    sheet: Worksheet,
    address: string,
    value: string | number | Date | null,
    numberFormat?: string,
  ) => {
    const cell = sheet.getCell(address);
    cell.value = value;
    if (numberFormat) cell.numFmt = numberFormat;
  };

  set(front, "G7", learner?.last_name || "");
  set(front, "W7", learner?.first_name || "");
  set(front, "AN7", "");
  set(front, "AX7", learner?.middle_name || "");
  set(front, "M8", learner?.lrn || "", "@");
  set(
    front,
    "AH8",
    learner?.birthdate ? new Date(`${learner.birthdate}T00:00:00`) : null,
    "mm/dd/yyyy",
  );
  set(front, "AV8", learner?.sex || "");

  const fillRecord = (
    sheet: Worksheet,
    headerRow: number,
    firstSubjectRow: number,
    averageRow: number,
  ) => {
    set(sheet, `E${headerRow}`, options.schoolName);
    set(sheet, `U${headerRow}`, options.schoolId, "@");
    set(sheet, `AD${headerRow}`, options.district);
    set(sheet, `AP${headerRow}`, options.division);
    set(sheet, `BB${headerRow}`, options.region);
    set(sheet, `I${headerRow + 1}`, options.gradeLevel);
    set(sheet, `N${headerRow + 1}`, options.section);
    set(sheet, `V${headerRow + 1}`, options.schoolYear);
    set(sheet, `AI${headerRow + 1}`, options.teacherName);

    for (let index = 0; index < 14; index += 1) {
      const row = firstSubjectRow + index;
      const subject = options.subjects[index];
      if (!subject) {
        ["B", "U", "Y", "AC", "AG", "AJ", "AP"].forEach((column) =>
          set(sheet, `${column}${row}`, null),
        );
        continue;
      }
      set(sheet, `B${row}`, subject.subject);
      set(sheet, `U${row}`, rounded(subject.term1));
      set(sheet, `Y${row}`, rounded(subject.term2));
      set(sheet, `AC${row}`, rounded(subject.term3));
      set(sheet, `AG${row}`, rounded(subject.term4));
      set(sheet, `AJ${row}`, rounded(subject.final));
      set(sheet, `AP${row}`, subject.remarks);
    }
    set(sheet, `AJ${averageRow}`, rounded(options.generalAverage));
    set(
      sheet,
      `AP${averageRow}`,
      typeof options.generalAverage === "number"
        ? options.generalAverage >= 75
          ? "PASSED"
          : "FAILED"
        : "",
    );
  };

  fillRecord(front, 21, 26, 40);
  fillRecord(front, 49, 54, 68);
  fillRecord(back, 3, 8, 22);
  fillRecord(back, 31, 36, 50);
  fillRecord(back, 59, 64, 78);

  set(front, "O79", learner ? studentName(learner) : "");
  set(front, "AF79", learner?.lrn || "", "@");
  set(front, "H80", options.schoolName);
  set(front, "AC80", options.schoolId, "@");
  set(front, "AT80", options.schoolYear);
  set(front, "S83", options.schoolHead);
  set(back, "O89", learner ? studentName(learner) : "");
  set(back, "AF89", learner?.lrn || "", "@");
  set(back, "H90", options.schoolName);
  set(back, "AC90", options.schoolId, "@");
  set(back, "AT90", options.schoolYear);
  set(back, "S93", options.schoolHead);

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

async function downloadSf8ExcelWorkbook(
  fileName: string,
  options: Sf8ExcelExportOptions,
) {
  const response = await fetch("/templates/SF8-admin.xlsx");
  if (!response.ok) {
    throw new Error(`Unable to load SF8 Excel template (${response.status}).`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const worksheet =
    workbook.getWorksheet("Nutritional Status") || workbook.worksheets[1];

  if (!worksheet) {
    throw new Error(
      "The SF8 Excel template does not contain the Nutritional Status sheet.",
    );
  }

  worksheet.getCell("F5").value = options.schoolName;
  worksheet.getCell("J5").value = options.district;
  worksheet.getCell("M5").value = options.division;
  worksheet.getCell("Q5").value = options.region;
  worksheet.getCell("C7").value = options.schoolId;
  worksheet.getCell("G7").value = options.gradeLevel;
  worksheet.getCell("I7").value = options.section;
  worksheet.getCell("Q7").value = options.schoolYear;
  worksheet.getCell("T7").value = new Date();
  worksheet.getCell("T7").numFmt = "m/d/yy";

  const male = options.learners.filter((learner) =>
    normalize(learner.sex).startsWith("m"),
  );
  const female = options.learners.filter((learner) =>
    normalize(learner.sex).startsWith("f"),
  );
  const unspecified = options.learners.filter((learner) => {
    const sex = normalize(learner.sex);
    return !sex.startsWith("m") && !sex.startsWith("f");
  });

  const clearRows = [
    ...Array.from({ length: 52 }, (_, index) => 12 + index),
    ...Array.from({ length: 44 }, (_, index) => 65 + index),
  ];
  clearRows.forEach((row) => {
    ["B", "C", "D", "H", "I", "J", "K", "L", "M", "N", "O", "P"].forEach(
      (column) => {
        worksheet.getCell(`${column}${row}`).value = null;
      },
    );
  });

  const writeLearner = (learner: StudentRow, row: number, number: number) => {
    worksheet.getCell(`B${row}`).value = number;
    worksheet.getCell(`C${row}`).value = learner.lrn || "";
    worksheet.getCell(`C${row}`).numFmt = "@";
    worksheet.getCell(`D${row}`).value = studentName(learner).toUpperCase();
    worksheet.getCell(`H${row}`).value = learner.birthdate
      ? new Date(`${learner.birthdate}T00:00:00`)
      : null;
    worksheet.getCell(`H${row}`).numFmt = "mm/dd/yyyy";
    worksheet.getCell(`I${row}`).value = calculateAge(learner.birthdate);
    worksheet.getCell(`P${row}`).value = learner.remarks || "";
  };

  male
    .slice(0, 52)
    .forEach((learner, index) => writeLearner(learner, 12 + index, index + 1));
  [...female, ...unspecified]
    .slice(0, 44)
    .forEach((learner, index) => writeLearner(learner, 65 + index, index + 1));

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

const SF1_ADMIN_LEGACY_TEMPLATE_URL = "/templates/SF1-class-adviser.xls";
const SF1_ADMIN_MALE_ROWS = 20;
const SF1_ADMIN_FEMALE_ROWS = 18;
const SF1_ADMIN_TEMPLATE_COLUMNS = [
  "A",
  "C",
  "G",
  "H",
  "J",
  "L",
  "N",
  "O",
  "P",
  "AB",
  "AF",
  "AK",
  "AP",
  "AR",
  "AS",
] as const;

type Sf1AdminLegacyCellUpdate = {
  row: number;
  column: number;
  value: string | number;
};

function sf1AdminExcelColumnIndex(column: string) {
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

function sf1AdminLegacyCellUpdate(
  address: string,
  value: string | number,
): Sf1AdminLegacyCellUpdate {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) {
    throw new Error(`Invalid SF1 template cell address: ${address}.`);
  }

  return {
    row: Number(match[2]) - 1,
    column: sf1AdminExcelColumnIndex(match[1]),
    value,
  };
}

function sf1AdminLearnerValues(
  learner: (StudentRow & { age: string }) | undefined,
  sex: "M" | "F",
): Array<string | number> {
  if (!learner) return SF1_ADMIN_TEMPLATE_COLUMNS.map(() => "");

  const fullName = [learner.last_name, learner.first_name, learner.middle_name]
    .filter(Boolean)
    .join(", ")
    .toUpperCase();

  return [
    learner.lrn || "",
    fullName,
    sex,
    learner.birthdate ? formatDate(learner.birthdate) : "",
    learner.age || "",
    learner.mother_tongue || "",
    learner.ip_ethnic_group || learner.ip_ethnic || "",
    learner.religion || "",
    learner.address || "",
    learner.father_name || "",
    learner.mother_name || "",
    learner.guardian || "",
    learner.contact_number || "",
    learner.learning_modality || "",
    learner.remarks || "",
  ];
}

function buildSf1AdminLegacyCellUpdates(options: Sf1ExcelExportOptions) {
  if (
    options.male.length > SF1_ADMIN_MALE_ROWS ||
    options.female.length > SF1_ADMIN_FEMALE_ROWS
  ) {
    throw new Error(
      `The original SF1 .xls template supports up to ${SF1_ADMIN_MALE_ROWS} male rows and ${SF1_ADMIN_FEMALE_ROWS} female rows.`,
    );
  }

  const updates: Sf1AdminLegacyCellUpdate[] = [
    sf1AdminLegacyCellUpdate("F3", options.schoolId),
    sf1AdminLegacyCellUpdate("K3", options.region),
    sf1AdminLegacyCellUpdate("T3", options.division),
    sf1AdminLegacyCellUpdate("F4", options.schoolName),
    sf1AdminLegacyCellUpdate("T4", options.schoolYear),
    sf1AdminLegacyCellUpdate("AE4", options.gradeLevel),
    sf1AdminLegacyCellUpdate("AM4", options.section),
  ];

  for (let index = 0; index < SF1_ADMIN_MALE_ROWS; index += 1) {
    const values = sf1AdminLearnerValues(options.male[index], "M");
    SF1_ADMIN_TEMPLATE_COLUMNS.forEach((column, columnIndex) => {
      updates.push(
        sf1AdminLegacyCellUpdate(
          `${column}${index + 7}`,
          values[columnIndex] ?? "",
        ),
      );
    });
  }

  for (let index = 0; index < SF1_ADMIN_FEMALE_ROWS; index += 1) {
    const values = sf1AdminLearnerValues(options.female[index], "F");
    SF1_ADMIN_TEMPLATE_COLUMNS.forEach((column, columnIndex) => {
      updates.push(
        sf1AdminLegacyCellUpdate(
          `${column}${index + 28}`,
          values[columnIndex] ?? "",
        ),
      );
    });
  }

  updates.push(
    sf1AdminLegacyCellUpdate("A27", options.male.length),
    sf1AdminLegacyCellUpdate("A46", options.female.length),
    sf1AdminLegacyCellUpdate(
      "A47",
      options.male.length + options.female.length,
    ),
    sf1AdminLegacyCellUpdate("X50", options.male.length),
    sf1AdminLegacyCellUpdate("X53", options.female.length),
    sf1AdminLegacyCellUpdate(
      "X55",
      options.male.length + options.female.length,
    ),
    sf1AdminLegacyCellUpdate("AE50", options.teacherName.toUpperCase()),
    sf1AdminLegacyCellUpdate("AN50", options.schoolHead.toUpperCase()),
    sf1AdminLegacyCellUpdate(
      "A59",
      `Generated on: ${new Intl.DateTimeFormat("en-PH", {
        dateStyle: "long",
      }).format(new Date())}`,
    ),
  );

  return updates;
}

function buildSf1AdminHiddenRows(options: Sf1ExcelExportOptions) {
  const hiddenRows = new Set<number>();

  // BIFF uses zero-based row indexes. Hide only unused learner rows so the
  // original SF1 .xls template remains intact.
  for (let row = 6 + options.male.length; row <= 25; row += 1) {
    hiddenRows.add(row);
  }
  for (let row = 27 + options.female.length; row <= 44; row += 1) {
    hiddenRows.add(row);
  }

  return hiddenRows;
}

function sf1ReadUInt16LE(bytes: Uint8Array, offset: number) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function sf1UInt16LE(value: number) {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function sf1ConcatBytes(chunks: Uint8Array[]) {
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

function sf1MakeBiffRecord(id: number, body: Uint8Array) {
  return sf1ConcatBytes([sf1UInt16LE(id), sf1UInt16LE(body.length), body]);
}

function sf1MakeBiffBlankRecord(row: number, column: number, xf: number) {
  return sf1MakeBiffRecord(
    0x0201,
    sf1ConcatBytes([sf1UInt16LE(row), sf1UInt16LE(column), sf1UInt16LE(xf)]),
  );
}

function sf1MakeBiffNumberRecord(
  row: number,
  column: number,
  xf: number,
  value: number,
) {
  const body = new Uint8Array(14);
  body.set(sf1UInt16LE(row), 0);
  body.set(sf1UInt16LE(column), 2);
  body.set(sf1UInt16LE(xf), 4);
  new DataView(body.buffer).setFloat64(6, value, true);
  return sf1MakeBiffRecord(0x0203, body);
}

function sf1MakeBiffLabelRecord(
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

  return sf1MakeBiffRecord(
    0x0204,
    sf1ConcatBytes([
      sf1UInt16LE(row),
      sf1UInt16LE(column),
      sf1UInt16LE(xf),
      sf1UInt16LE(text.length),
      Uint8Array.of(wide ? 1 : 0),
      characters,
    ]),
  );
}

function sf1MakeBiffValueRecord(
  row: number,
  column: number,
  xf: number,
  value: string | number,
) {
  if (value === "") return sf1MakeBiffBlankRecord(row, column, xf);

  if (typeof value === "number" && Number.isFinite(value)) {
    return sf1MakeBiffNumberRecord(row, column, xf, value);
  }

  return sf1MakeBiffLabelRecord(row, column, xf, String(value));
}

function sf1MakeBiffMulBlankRecord(
  row: number,
  firstColumn: number,
  xfs: number[],
) {
  if (xfs.length === 0) return undefined;

  return sf1MakeBiffRecord(
    0x00be,
    sf1ConcatBytes([
      sf1UInt16LE(row),
      sf1UInt16LE(firstColumn),
      ...xfs.map((xf) => sf1UInt16LE(xf)),
      sf1UInt16LE(firstColumn + xfs.length - 1),
    ]),
  );
}

function patchSf1AdminOriginalWorkbook(
  workbookBytes: Uint8Array,
  cellUpdates: Sf1AdminLegacyCellUpdate[],
  hiddenRows: ReadonlySet<number>,
) {
  const pending = new Map<string, Sf1AdminLegacyCellUpdate>(
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
    const id = sf1ReadUInt16LE(workbookBytes, offset);
    const length = sf1ReadUInt16LE(workbookBytes, offset + 2);
    const start = offset + 4;
    const end = start + length;

    if (end > workbookBytes.length) {
      throw new Error("The original SF1 .xls workbook stream is incomplete.");
    }

    if (id === 0x0809 && length >= 4) {
      const substreamType = sf1ReadUInt16LE(workbookBytes, start + 2);
      if (substreamType === 0x0010) {
        worksheetIndex += 1;
        inPrimaryWorksheet = worksheetIndex === 0;
      }
    }

    // INDEX and DBCELL contain byte offsets that become stale when records
    // are replaced. Excel rebuilds these optional performance records.
    if (id === 0x020b || id === 0x00d7) {
      offset = end;
      continue;
    }

    if (inPrimaryWorksheet && id === 0x0208 && length >= 16) {
      const row = sf1ReadUInt16LE(workbookBytes, start);

      if (hiddenRows.has(row)) {
        const hiddenRowRecord = workbookBytes.slice(offset, end);
        hiddenRowRecord[16] |= 0x20;
        chunks.push(hiddenRowRecord);
      } else {
        chunks.push(workbookBytes.slice(offset, end));
      }
    } else if (inPrimaryWorksheet && id === 0x00be && length >= 8) {
      const row = sf1ReadUInt16LE(workbookBytes, start);
      const firstColumn = sf1ReadUInt16LE(workbookBytes, start + 2);
      const lastColumn = sf1ReadUInt16LE(workbookBytes, end - 2);
      const xfs: number[] = [];

      for (let cursor = start + 4; cursor < end - 2; cursor += 2) {
        xfs.push(sf1ReadUInt16LE(workbookBytes, cursor));
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
          const before = sf1MakeBiffMulBlankRecord(
            row,
            cursorColumn,
            xfs.slice(cursorColumn - firstColumn, update.column - firstColumn),
          );
          if (before) chunks.push(before);

          chunks.push(
            sf1MakeBiffValueRecord(
              row,
              update.column,
              xfs[update.column - firstColumn],
              update.value,
            ),
          );
          pending.delete(`${update.row}:${update.column}`);
          cursorColumn = update.column + 1;
        });

        const after = sf1MakeBiffMulBlankRecord(
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
      const row = sf1ReadUInt16LE(workbookBytes, start);
      const column = sf1ReadUInt16LE(workbookBytes, start + 2);
      const key = `${row}:${column}`;
      const update = pending.get(key);

      if (update) {
        chunks.push(
          sf1MakeBiffValueRecord(
            row,
            column,
            sf1ReadUInt16LE(workbookBytes, start + 4),
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

    if (id === 0x000a && inPrimaryWorksheet) {
      inPrimaryWorksheet = false;
    }

    offset = end;
  }

  if (offset < workbookBytes.length) {
    chunks.push(workbookBytes.slice(offset));
  }

  if (pending.size > 0) {
    const first = pending.values().next().value as
      Sf1AdminLegacyCellUpdate | undefined;

    throw new Error(
      first
        ? `The original SF1 .xls template is missing cell row ${first.row + 1}, column ${first.column + 1}.`
        : "The original SF1 .xls template is missing an expected cell.",
    );
  }

  return sf1ConcatBytes(chunks);
}

async function downloadSf1ExcelWorkbook(
  fileName: string,
  options: Sf1ExcelExportOptions,
) {
  const response = await fetch(SF1_ADMIN_LEGACY_TEMPLATE_URL, {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(
      `Unable to load SF1 .xls template (${response.status}). Put SF1-class-adviser.xls inside public/templates/.`,
    );
  }

  const templateBytes = new Uint8Array(await response.arrayBuffer());
  const compoundFile = CFB.read(templateBytes, { type: "array" });
  const workbookEntry =
    CFB.find(compoundFile, "Workbook") || CFB.find(compoundFile, "Book");

  if (!workbookEntry?.content) {
    throw new Error("The SF1 .xls template has no BIFF workbook stream.");
  }

  const workbookBytes = patchSf1AdminOriginalWorkbook(
    new Uint8Array(workbookEntry.content),
    buildSf1AdminLegacyCellUpdates(options),
    buildSf1AdminHiddenRows(options),
  );

  workbookEntry.content = workbookBytes;
  workbookEntry.size = workbookBytes.length;

  const output = CFB.write(compoundFile, { type: "array" }) as Uint8Array;
  const outputBuffer = Uint8Array.from(output).buffer as ArrayBuffer;
  const blob = new Blob([outputBuffer], {
    type: "application/vnd.ms-excel",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${fileName}.xls`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const ADMIN_XLSX_MAIN_NS =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ADMIN_XLSX_REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ADMIN_XML_NS = "http://www.w3.org/XML/1998/namespace";

function getAdminXlsxCell(sheet: Document, reference: string) {
  return sheet.querySelector(`c[r="${reference}"]`);
}

function setAdminXlsxText(
  sheet: Document,
  reference: string,
  value: string | number | null | undefined,
) {
  const cell = getAdminXlsxCell(sheet, reference);
  if (!cell) throw new Error(`SF9 template is missing cell ${reference}.`);
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  const text = String(value ?? "");
  if (!text) {
    cell.removeAttribute("t");
    return;
  }
  cell.setAttribute("t", "inlineStr");
  const inline = sheet.createElementNS(ADMIN_XLSX_MAIN_NS, "is");
  const textNode = sheet.createElementNS(ADMIN_XLSX_MAIN_NS, "t");
  textNode.setAttributeNS(ADMIN_XML_NS, "xml:space", "preserve");
  textNode.textContent = text;
  inline.appendChild(textNode);
  cell.appendChild(inline);
}

function setAdminXlsxNumber(
  sheet: Document,
  reference: string,
  value: number | null | undefined,
) {
  const cell = getAdminXlsxCell(sheet, reference);
  if (!cell) throw new Error(`SF9 template is missing cell ${reference}.`);
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.removeAttribute("t");
  if (value == null || !Number.isFinite(Number(value))) return;
  const valueNode = sheet.createElementNS(ADMIN_XLSX_MAIN_NS, "v");
  valueNode.textContent = String(Number(value));
  cell.appendChild(valueNode);
}

function setAdminXlsxValue(
  sheet: Document,
  reference: string,
  value: string | number | null | undefined,
) {
  if (typeof value === "number" && Number.isFinite(value)) {
    setAdminXlsxNumber(sheet, reference, value);
  } else {
    setAdminXlsxText(sheet, reference, value);
  }
}

function setAdminXlsxRowHidden(sheet: Document, rowNumber: number, hidden: boolean) {
  const row = sheet.querySelector(`row[r="${rowNumber}"]`);
  if (!row) return;
  if (hidden) row.setAttribute("hidden", "1");
  else row.removeAttribute("hidden");
}

function applyAdminCenteredXlsxStyle(
  sheet: Document,
  styles: Document,
  reference: string,
) {
  const cell = getAdminXlsxCell(sheet, reference);
  if (!cell) return;
  const cellXfs = styles.getElementsByTagNameNS(ADMIN_XLSX_MAIN_NS, "cellXfs")[0];
  if (!cellXfs) return;
  const currentIndex = Number(cell.getAttribute("s") || "0");
  const current = cellXfs.children.item(currentIndex);
  if (!current) return;
  const cloned = current.cloneNode(true) as Element;
  cloned.setAttribute("numFmtId", "1");
  cloned.setAttribute("applyNumberFormat", "1");
  let alignment = cloned.getElementsByTagNameNS(ADMIN_XLSX_MAIN_NS, "alignment")[0];
  if (!alignment) {
    alignment = styles.createElementNS(ADMIN_XLSX_MAIN_NS, "alignment");
    cloned.appendChild(alignment);
  }
  alignment.setAttribute("horizontal", "center");
  alignment.setAttribute("vertical", "center");
  cloned.setAttribute("applyAlignment", "1");
  const newIndex = cellXfs.children.length;
  cellXfs.appendChild(cloned);
  cellXfs.setAttribute("count", String(cellXfs.children.length));
  cell.setAttribute("s", String(newIndex));
}

async function resolveAdminSf9Worksheet(zip: JSZip, preferredNames: string[]) {
  const workbookFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookFile || !relsFile) throw new Error("SF9 workbook metadata is missing.");
  const parser = new DOMParser();
  const workbook = parser.parseFromString(await workbookFile.async("text"), "application/xml");
  const rels = parser.parseFromString(await relsFile.async("text"), "application/xml");
  const sheets = Array.from(workbook.getElementsByTagNameNS("*", "sheet"));
  const wanted = preferredNames.map((name) => name.trim().toLowerCase());
  const selected = sheets.find((sheet) => wanted.includes(String(sheet.getAttribute("name") || "").trim().toLowerCase())) ?? sheets[0];
  if (!selected) throw new Error("SF9 worksheet was not found.");
  const relId = selected.getAttributeNS(ADMIN_XLSX_REL_NS, "id") || selected.getAttribute("r:id");
  const rel = Array.from(rels.getElementsByTagNameNS("*", "Relationship")).find((item) => item.getAttribute("Id") === relId);
  const target = rel?.getAttribute("Target");
  if (!target) throw new Error("SF9 worksheet relationship is missing.");
  const clean = target.replace(/^\/+/, "").replace(/^\.\//, "");
  const path = clean.startsWith("xl/") ? clean : `xl/${clean}`;
  const sheetFile = zip.file(path);
  if (!sheetFile) throw new Error("SF9 worksheet XML is missing.");
  const sheet = parser.parseFromString(await sheetFile.async("text"), "application/xml");
  return { path, sheet };
}

async function removeAdminSf9CalcChain(zip: JSZip) {
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  zip.remove("xl/calcChain.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (relsFile) {
    const rels = parser.parseFromString(await relsFile.async("text"), "application/xml");
    Array.from(rels.getElementsByTagNameNS("*", "Relationship")).forEach((rel) => {
      const type = rel.getAttribute("Type") || "";
      const target = rel.getAttribute("Target") || "";
      if (type.endsWith("/calcChain") || target.endsWith("calcChain.xml")) rel.parentNode?.removeChild(rel);
    });
    zip.file("xl/_rels/workbook.xml.rels", serializer.serializeToString(rels));
  }
  const contentTypesFile = zip.file("[Content_Types].xml");
  if (contentTypesFile) {
    const doc = parser.parseFromString(await contentTypesFile.async("text"), "application/xml");
    Array.from(doc.getElementsByTagNameNS("*", "Override")).forEach((item) => {
      if ((item.getAttribute("PartName") || "") === "/xl/calcChain.xml") item.parentNode?.removeChild(item);
    });
    zip.file("[Content_Types].xml", serializer.serializeToString(doc));
  }
  const workbookFile = zip.file("xl/workbook.xml");
  if (workbookFile) {
    const doc = parser.parseFromString(await workbookFile.async("text"), "application/xml");
    const root = doc.getElementsByTagNameNS(ADMIN_XLSX_MAIN_NS, "workbook")[0];
    if (root) {
      let calcPr = doc.getElementsByTagNameNS(ADMIN_XLSX_MAIN_NS, "calcPr")[0];
      if (!calcPr) {
        calcPr = doc.createElementNS(ADMIN_XLSX_MAIN_NS, "calcPr");
        root.appendChild(calcPr);
      }
      calcPr.setAttribute("calcMode", "auto");
      calcPr.setAttribute("fullCalcOnLoad", "1");
      calcPr.setAttribute("forceFullCalc", "1");
    }
    zip.file("xl/workbook.xml", serializer.serializeToString(doc));
  }
}

async function saveAdminPreservedXlsx(
  zip: JSZip,
  worksheetPath: string,
  sheet: Document,
  styles?: Document,
) {
  const serializer = new XMLSerializer();
  zip.file(worksheetPath, serializer.serializeToString(sheet));
  if (styles) zip.file("xl/styles.xml", serializer.serializeToString(styles));
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

function downloadAdminSf9Xlsx(buffer: ArrayBuffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${fileName.replace(/[<>:"/\\|?*]+/g, "_").replace(/\s+/g, "_")}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadSf9NewExcelWorkbook(
  fileName: string,
  options: Sf9NewExcelExportOptions,
) {
  const gradeKey = adminSf9GradeKey(options.gradeLevel);
  const isGrade11 = gradeKey === "grade 11" || gradeKey === "11";
  const isGrade12 = gradeKey === "grade 12" || gradeKey === "12";
  const rounded = (value: number | null | undefined) =>
    typeof value === "number" && Number.isFinite(value) ? Math.round(value) : null;

  if (isGrade11 || isGrade12) {
    const templateUrl = isGrade11
      ? ADMIN_SF9_GRADE11_TEMPLATE_URL
      : ADMIN_SF9_GRADE12_TEMPLATE_URL;
    const response = await fetch(templateUrl, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Unable to load ${templateUrl} (${response.status}).`);
    }

    const bytes = await response.arrayBuffer();
    const zip = await JSZip.loadAsync(bytes.slice(0));
    const { path: worksheetPath, sheet } = await resolveAdminSf9Worksheet(
      zip,
      isGrade11
        ? ["SF9 - GRADE 11 SHS", "SF9 - GRADE 12 ACADEMIC"]
        : ["SF9 - GRADE 11 ACADEMIC", "SF9 - GRADE 12 ACADEMIC", "SF9 - GRADE 12 SHS"],
    );
    const division = String(options.division || "").replace(/^SCHOOLS DIVISION OFFICE OF\s*/i, "");
    const municipality = options.municipality || "Prosperidad, Agusan del Sur";
    const learner = options.learner;

    // The supplied Grade 11 and Grade 12 workbooks use the same SHS header cells.
    setAdminXlsxText(sheet, "B7", options.region || "CARAGA Region");
    setAdminXlsxText(sheet, "B8", `SCHOOLS DIVISION OFFICE OF ${division}`.trim());
    setAdminXlsxText(sheet, "B9", options.district || "");
    setAdminXlsxText(sheet, "B10", municipality);
    setAdminXlsxText(sheet, "B12", options.schoolName);
    setAdminXlsxText(sheet, "B15", `School Year ${options.schoolYear}`.trim());
    setAdminXlsxText(sheet, "C17", learner ? studentName(learner).toUpperCase() : "");
    setAdminXlsxValue(sheet, "I17", learner ? calculateAge(learner.birthdate) : "");
    setAdminXlsxText(sheet, "K17", learner?.sex || "");
    setAdminXlsxText(sheet, "C18", learner?.lrn || "");
    setAdminXlsxText(sheet, "I18", options.gradeLevel);
    setAdminXlsxText(sheet, "K18", options.section);
    setAdminXlsxText(sheet, "E19", options.track || "");
    setAdminXlsxText(sheet, "B27", options.schoolHead);
    setAdminXlsxText(sheet, "H27", options.teacherName);

    if (isGrade11) {
      const stylesFile = zip.file("xl/styles.xml");
      const styles = stylesFile
        ? new DOMParser().parseFromString(await stylesFile.async("text"), "application/xml")
        : null;
      const elective = options.subjects.find(
        (row) => normalizeAdminSf9Subject(row.subject) === "elective subject",
      );
      const regular = options.subjects
        .filter((row) => normalizeAdminSf9Subject(row.subject) !== "elective subject")
        .slice(0, 7);

      for (let index = 0; index < 7; index += 1) {
        const rowNumber = 33 + index;
        const row = regular[index];
        setAdminXlsxRowHidden(sheet, rowNumber, !row);
        ["B", "F", "G", "H", "I", "J", "K"].forEach((column) =>
          setAdminXlsxText(sheet, `${column}${rowNumber}`, ""),
        );
        if (!row) continue;
        setAdminXlsxText(sheet, `B${rowNumber}`, row.subject);
        setAdminXlsxValue(sheet, `F${rowNumber}`, rounded(row.term1));
        setAdminXlsxValue(sheet, `G${rowNumber}`, rounded(row.term2));
        setAdminXlsxValue(sheet, `H${rowNumber}`, rounded(row.term3));
        setAdminXlsxValue(sheet, `I${rowNumber}`, row.units ?? null);
        setAdminXlsxValue(sheet, `J${rowNumber}`, rounded(row.final));
        setAdminXlsxText(sheet, `K${rowNumber}`, row.remarks ? row.remarks[0] + row.remarks.slice(1).toLowerCase() : "");
      }

      const electiveRows = [
        { row: 40, label: "Academic Elective 1", column: "F", grade: elective?.term1 ?? null },
        { row: 41, label: "Academic Elective 2", column: "G", grade: elective?.term2 ?? null },
        { row: 42, label: "Academic Elective 3", column: "H", grade: elective?.term3 ?? null },
      ] as const;
      electiveRows.forEach((item) => {
        setAdminXlsxRowHidden(sheet, item.row, false);
        setAdminXlsxText(sheet, `B${item.row}`, item.label);
        ["F", "G", "H", "I", "J", "K"].forEach((column) =>
          setAdminXlsxText(sheet, `${column}${item.row}`, ""),
        );
        setAdminXlsxValue(sheet, `${item.column}${item.row}`, rounded(item.grade));
        setAdminXlsxValue(sheet, `I${item.row}`, item.grade == null ? null : elective?.units ?? null);
        if (styles) {
          ["F", "G", "H", "I"].forEach((column) =>
            applyAdminCenteredXlsxStyle(sheet, styles, `${column}${item.row}`),
          );
        }
      });

      ["B", "F", "G", "H", "I", "J", "K"].forEach((column) =>
        setAdminXlsxText(sheet, `${column}43`, ""),
      );
      setAdminXlsxText(sheet, "B43", "General Average");
      setAdminXlsxValue(sheet, "J43", rounded(options.generalAverage));
      setAdminXlsxText(
        sheet,
        "K43",
        typeof options.generalAverage === "number"
          ? options.generalAverage >= 75
            ? "Promoted"
            : "Failed"
          : "",
      );
      if (styles) applyAdminCenteredXlsxStyle(sheet, styles, "J43");

      const buffer = await saveAdminPreservedXlsx(zip, worksheetPath, sheet, styles || undefined);
      downloadAdminSf9Xlsx(buffer, fileName);
      return;
    }

    // Grade 12: keep all original gray/white term cells and write only the
    // subject's assigned non-gray term, exactly like the Class Adviser SF9.
    ADMIN_SF9_GRADE12_SUBJECTS.forEach((spec) => {
      const row = options.subjects.find(
        (item) => normalizeAdminSf9Subject(item.subject) === normalizeAdminSf9Subject(spec.label),
      );
      setAdminXlsxText(sheet, `B${spec.excelRow}`, spec.label);
      ["F", "G", "H"].forEach((column) => setAdminXlsxText(sheet, `${column}${spec.excelRow}`, ""));
      const termColumn = spec.term === "1" ? "F" : spec.term === "2" ? "G" : "H";
      const termGrade =
        spec.term === "1" ? row?.term1 ?? null : spec.term === "2" ? row?.term2 ?? null : row?.term3 ?? null;
      setAdminXlsxValue(sheet, `${termColumn}${spec.excelRow}`, rounded(termGrade));
      setAdminXlsxValue(sheet, `I${spec.excelRow}`, row?.units ?? null);
      setAdminXlsxValue(sheet, `J${spec.excelRow}`, rounded(row?.final ?? termGrade));
      setAdminXlsxText(
        sheet,
        `K${spec.excelRow}`,
        typeof (row?.final ?? termGrade) === "number"
          ? Number(row?.final ?? termGrade) >= 75
            ? "Passed"
            : "Failed"
          : "",
      );
    });
    setAdminXlsxValue(sheet, "J53", rounded(options.generalAverage));
    setAdminXlsxText(
      sheet,
      "K53",
      typeof options.generalAverage === "number"
        ? options.generalAverage >= 75
          ? "Promoted"
          : "Failed"
        : "",
    );
    await removeAdminSf9CalcChain(zip);
    const buffer = await saveAdminPreservedXlsx(zip, worksheetPath, sheet);
    downloadAdminSf9Xlsx(buffer, fileName);
    return;
  }

  // Grades 7–10 keep the existing Admin SF9 export.
  const response = await fetch("/templates/sf9-3-terms-template.xlsx");
  if (!response.ok) {
    throw new Error(`Unable to load SF9 Excel template (${response.status}).`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const worksheet = workbook.getWorksheet("School Form 9") || workbook.worksheets[0];
  if (!worksheet) throw new Error("The SF9 Excel template does not contain a worksheet.");
  const learner = options.learner;
  const regionLabel = options.region
    ? /region$/i.test(options.region.trim())
      ? options.region.trim()
      : `${options.region.trim()} Region`
    : "";
  worksheet.getCell("B4").value = regionLabel;
  worksheet.getCell("B5").value = options.division;
  worksheet.getCell("B6").value = options.district;
  worksheet.getCell("B8").value = options.schoolName;
  worksheet.getCell("D11").value = options.schoolYear;
  worksheet.getCell("C13").value = learner ? studentName(learner) : "";
  worksheet.getCell("G13").value = learner ? calculateAge(learner.birthdate) : "";
  worksheet.getCell("C14").value = learner?.lrn || "";
  worksheet.getCell("G14").value = learner?.sex || "";
  worksheet.getCell("C15").value = "";
  worksheet.getCell("G15").value = options.gradeLevel;
  worksheet.getCell("G16").value = options.section;
  for (let index = 0; index < 11; index += 1) {
    const rowNumber = 22 + index;
    const subject = options.subjects[index];
    worksheet.getCell(`B${rowNumber}`).value = subject?.subject || "";
    worksheet.getCell(`C${rowNumber}`).value = rounded(subject?.term1 ?? null);
    worksheet.getCell(`D${rowNumber}`).value = rounded(subject?.term2 ?? null);
    worksheet.getCell(`E${rowNumber}`).value = rounded(subject?.term3 ?? null);
    worksheet.getCell(`F${rowNumber}`).value = rounded(subject?.final ?? null);
    worksheet.getCell(`G${rowNumber}`).value = subject?.remarks || "";
  }
  worksheet.getCell("F34").value = rounded(options.generalAverage);
  worksheet.getCell("G34").value =
    typeof options.generalAverage === "number"
      ? options.generalAverage >= 75
        ? "PROMOTED"
        : "FAILED"
      : "";
  worksheet.getCell("J34").value = options.schoolHead;
  worksheet.getCell("Q33").value = options.teacherName;
  worksheet.getCell("L39").value = options.schoolHead;
  const buffer = await workbook.xlsx.writeBuffer();
  downloadAdminSf9Xlsx(buffer as ArrayBuffer, fileName);
}

function formatDate(value: string | null | undefined) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return new Intl.DateTimeFormat("en-PH", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(date);
}

function formatDateTime(value: string | null | undefined) {
  if (!value) return { date: "-", time: "-" };

  const parsedDate = new Date(value);
  if (Number.isNaN(parsedDate.getTime())) return { date: "-", time: "-" };

  return {
    date: new Intl.DateTimeFormat("en-PH", {
      month: "short",
      day: "numeric",
      year: "numeric",
    }).format(parsedDate),
    time: new Intl.DateTimeFormat("en-PH", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    }).format(parsedDate),
  };
}

function formatDateTimeText(value: string | null | undefined) {
  const formatted = formatDateTime(value);
  return formatted.date === "-" ? "—" : `${formatted.date} ${formatted.time}`;
}

function studentName(student: StudentRow) {
  return (
    [student.last_name, student.first_name, student.middle_name]
      .filter(Boolean)
      .join(", ")
      .replace(", ,", ", ") || "No name"
  );
}

function studentTransferStatus(
  student: StudentRow,
): "in" | "out" | "none" {
  const status = normalize(student.enrollment_status)
    .replace(/[\s-]+/g, "_");

  if (status === "transferred_in" || status === "transfer_in") {
    return "in";
  }

  if (status === "transferred_out" || status === "transfer_out") {
    return "out";
  }

  return "none";
}

function StudentTransferStatusBadge({ student }: { student: StudentRow }) {
  const status = studentTransferStatus(student);

  if (status === "in") {
    return (
      <span
        className="inline-flex items-center rounded-full bg-emerald-100 px-2.5 py-1 text-[9px] font-bold text-emerald-700"
        title={
          student.previous_school
            ? `Transferred in from ${student.previous_school}`
            : "Transferred In"
        }
      >
        Transfer In
      </span>
    );
  }

  if (status === "out") {
    return (
      <span
        className="inline-flex items-center rounded-full bg-rose-100 px-2.5 py-1 text-[9px] font-bold text-rose-700"
        title={
          student.destination_school
            ? `Transferred out to ${student.destination_school}`
            : "Transferred Out"
        }
      >
        Transfer Out
      </span>
    );
  }

  return (
    <span className="text-[var(--admin-text-muted)]">
      —
    </span>
  );
}

function calculateAge(birthdate: string | null | undefined) {
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

type Sf1PreviewLearner = StudentRow & {
  age: string;
  sexGroup: string;
};

function Sf1SpreadsheetPreview({
  schoolName,
  schoolId,
  region,
  division,
  schoolYear,
  gradeLevel,
  section,
  teacherName,
  schoolHead,
  male,
  female,
}: {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  schoolYear: string;
  gradeLevel: string;
  section: string;
  teacherName: string;
  schoolHead: string;
  male: Sf1PreviewLearner[];
  female: Sf1PreviewLearner[];
}) {
  const learners = [
    ...male.map((learner) => ({ learner, sex: "M" as const })),
    ...female.map((learner) => ({ learner, sex: "F" as const })),
  ];

  const cellClass =
    "border border-black px-[2px] py-[2px] align-middle leading-[1.05]";
  const tinyCellClass =
    "border border-black px-[2px] py-[1px] align-middle leading-[1.05]";

  return (
    <div
      className="box-border h-full w-full bg-white px-4 pb-4 pt-3 text-[4.6px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="text-center">
        <div className="text-[15px] font-semibold leading-tight">
          School Form 1 (SF 1) School Register
        </div>
        <div className="mt-[2px] text-[5px] italic leading-tight">
          (This replaces Form 1, Master List &amp; STS Form 2-Family Background
          and Profile)
        </div>
      </div>

      <div className="mt-3 grid grid-cols-[1.08fr_.92fr] gap-x-12 gap-y-[3px] text-[5.6px]">
        <div className="grid grid-cols-[70px_1fr_1fr] items-center">
          <span className="pr-2 text-right">School ID</span>
          <span className="h-[17px] border border-black px-2 py-[2px] text-center">
            {schoolId}
          </span>
          <span className="h-[17px] border border-l-0 border-black px-2 py-[2px] text-center">
            {region}
          </span>
        </div>

        <div className="grid grid-cols-[55px_1fr] items-center">
          <span className="pr-2 text-right">Division</span>
          <span className="h-[17px] border border-black px-2 py-[2px] text-center">
            {division}
          </span>
        </div>

        <div className="grid grid-cols-[70px_1fr] items-center">
          <span className="pr-2 text-right">School Name</span>
          <span className="h-[17px] border border-black px-2 py-[2px] text-center">
            {schoolName}
          </span>
        </div>

        <div className="grid grid-cols-[58px_1fr_52px_1fr_36px_1fr] items-center">
          <span className="pr-2 text-right">School Year</span>
          <span className="h-[17px] border border-black px-1 py-[2px] text-center">
            {schoolYear}
          </span>
          <span className="pr-2 text-right">Grade Level</span>
          <span className="h-[17px] border border-l-0 border-black px-1 py-[2px] text-center">
            {gradeLevel}
          </span>
          <span className="pr-2 text-right">Section</span>
          <span className="h-[17px] border border-black px-1 py-[2px] text-center">
            {section}
          </span>
        </div>
      </div>

      <table className="mt-[3px] w-full table-fixed border-collapse text-[4.15px]">
        <colgroup>
          {[
            7, 12, 2.5, 5.5, 3, 4.5, 3.5, 4, 5, 5, 5, 5, 6, 6, 5.5, 4, 4.5, 4.5,
            4.5, 4.5,
          ].map((width, index) => (
            <col
              key={`sf1-class-adviser-col-${index}`}
              style={{ width: `${width}%` }}
            />
          ))}
        </colgroup>

        <thead>
          <tr className="h-[24px] text-center font-semibold">
            <th rowSpan={2} className={cellClass}>
              LRN
            </th>
            <th rowSpan={2} className={cellClass}>
              NAME
              <br />
              <span className="font-normal">
                (Last Name, First Name, Middle Name)
              </span>
            </th>
            <th rowSpan={2} className={cellClass}>
              Sex
              <br />
              (M/F)
            </th>
            <th rowSpan={2} className={cellClass}>
              BIRTH DATE
              <br />
              (mm/dd/yyyy)
            </th>
            <th rowSpan={2} className={cellClass}>
              AGE as of
              <br />
              1st Friday June
            </th>
            <th rowSpan={2} className={cellClass}>
              MOTHER TONGUE
              <br />
              (Grade 1 to 3 Only)
            </th>
            <th rowSpan={2} className={cellClass}>
              IP
              <br />
              (Ethnic Group)
            </th>
            <th rowSpan={2} className={cellClass}>
              RELIGION
            </th>
            <th colSpan={4} className={cellClass}>
              ADDRESS
            </th>
            <th colSpan={2} className={cellClass}>
              PARENTS
            </th>
            <th colSpan={2} className={cellClass}>
              GUARDIAN
              <br />
              (if Not Parent)
            </th>
            <th rowSpan={2} className={cellClass}>
              Contact Number of
              <br />
              Parent or Guardian
            </th>
            <th rowSpan={2} className={cellClass}>
              Learning Modality
            </th>
            <th colSpan={2} className={cellClass}>
              REMARKS
            </th>
          </tr>

          <tr className="h-[24px] text-center font-normal">
            <th className={cellClass}>House #/Street/Sitio/Purok</th>
            <th className={cellClass}>Barangay</th>
            <th className={cellClass}>Municipality/City</th>
            <th className={cellClass}>Province</th>
            <th className={cellClass}>Father&apos;s Name</th>
            <th className={cellClass}>Mother&apos;s Maiden Name</th>
            <th className={cellClass}>Name</th>
            <th className={cellClass}>Relationship</th>
            <th colSpan={2} className={cellClass}>
              (Please refer to the legend on last page)
            </th>
          </tr>
        </thead>

        <tbody>
          {learners.map(({ learner, sex }) => (
            <tr key={`sf1-class-adviser-${learner.id}`} className="h-[17px]">
              <td className={cellClass}>{learner.lrn || ""}</td>
              <td className={cellClass}>
                {[learner.last_name, learner.first_name, learner.middle_name]
                  .filter(Boolean)
                  .join(", ")
                  .toUpperCase()}
              </td>
              <td className={`${cellClass} text-center`}>{sex}</td>
              <td className={`${cellClass} text-center`}>
                {learner.birthdate ? formatDate(learner.birthdate) : ""}
              </td>
              <td className={`${cellClass} text-center`}>
                {learner.age || ""}
              </td>
              <td className={cellClass}>{learner.mother_tongue || ""}</td>
              <td className={cellClass}>
                {learner.ip_ethnic_group || learner.ip_ethnic || ""}
              </td>
              <td className={cellClass}>{learner.religion || ""}</td>
              <td className={cellClass}>{learner.address || ""}</td>
              <td className={cellClass} />
              <td className={cellClass} />
              <td className={cellClass} />
              <td className={cellClass}>{learner.father_name || ""}</td>
              <td className={cellClass}>{learner.mother_name || ""}</td>
              <td className={cellClass}>{learner.guardian || ""}</td>
              <td className={cellClass} />
              <td className={cellClass}>{learner.contact_number || ""}</td>
              <td className={cellClass}>{learner.learning_modality || ""}</td>
              <td colSpan={2} className={cellClass}>
                {learner.remarks || ""}
              </td>
            </tr>
          ))}

          <tr className="h-[14px] font-semibold">
            <td className={`${tinyCellClass} text-right`}>{male.length}</td>
            <td colSpan={19} className={tinyCellClass}>
              &lt;=== TOTAL MALE
            </td>
          </tr>
          <tr className="h-[14px] font-semibold">
            <td className={`${tinyCellClass} text-right`}>{female.length}</td>
            <td colSpan={19} className={tinyCellClass}>
              &lt;=== TOTAL FEMALE
            </td>
          </tr>
          <tr className="h-[14px] font-semibold">
            <td className={`${tinyCellClass} text-right`}>
              {male.length + female.length}
            </td>
            <td colSpan={19} className={tinyCellClass}>
              &lt;=== COMBINED
            </td>
          </tr>
        </tbody>
      </table>

      <div className="border-x border-b border-black py-[2px] text-center text-[4.6px]">
        List and Code of Indicators under REMARKS column
      </div>

      <div className="mt-[4px] grid grid-cols-[1.72fr_.48fr_1fr_1fr] gap-3 text-[4.2px]">
        <div>
          <table className="w-full border-collapse">
            <tbody>
              <tr>
                <td className={tinyCellClass}>Indicator</td>
                <td className={`${tinyCellClass} w-[38px] text-center`}>
                  Code
                </td>
                <td className={tinyCellClass}>Indicator</td>
                <td className={`${tinyCellClass} w-[38px] text-center`}>
                  Code
                </td>
              </tr>
              <tr>
                <td className={tinyCellClass}>Transferred Out</td>
                <td className={`${tinyCellClass} text-center`}>T/O</td>
                <td className={tinyCellClass}>CCT Recipient</td>
                <td className={`${tinyCellClass} text-center`}>CCT</td>
              </tr>
              <tr>
                <td className={tinyCellClass}>Transferred In</td>
                <td className={`${tinyCellClass} text-center`}>T/I</td>
                <td className={tinyCellClass}>Balik Aral</td>
                <td className={`${tinyCellClass} text-center`}>B/A</td>
              </tr>
              <tr>
                <td className={tinyCellClass}>Dropped</td>
                <td className={`${tinyCellClass} text-center`}>DRP</td>
                <td className={tinyCellClass}>Special Needs Education</td>
                <td className={`${tinyCellClass} text-center`}>SNED</td>
              </tr>
              <tr>
                <td className={tinyCellClass}>Late Enrollment</td>
                <td className={`${tinyCellClass} text-center`}>LE</td>
                <td className={tinyCellClass}>Accelerated</td>
                <td className={`${tinyCellClass} text-center`}>ACL</td>
              </tr>
            </tbody>
          </table>
        </div>

        <table className="h-fit w-full border-collapse text-center">
          <tbody>
            <tr>
              <td className={tinyCellClass}>REGISTERED</td>
              <td className={tinyCellClass}>BoSY</td>
              <td className={tinyCellClass}>EoSY</td>
            </tr>
            <tr>
              <td className={tinyCellClass}>MALE</td>
              <td className={tinyCellClass}>{male.length}</td>
              <td className={tinyCellClass} />
            </tr>
            <tr>
              <td className={tinyCellClass}>FEMALE</td>
              <td className={tinyCellClass}>{female.length}</td>
              <td className={tinyCellClass} />
            </tr>
            <tr>
              <td className={tinyCellClass}>TOTAL</td>
              <td className={tinyCellClass}>{male.length + female.length}</td>
              <td className={tinyCellClass} />
            </tr>
          </tbody>
        </table>

        <div className="pt-[2px] text-center">
          <div className="text-left">Prepared by:</div>
          <div className="mt-5 border-b border-black pb-[1px] font-semibold">
            {teacherName}
          </div>
          <div className="mt-[1px]">
            (Signature of Adviser over Printed Name)
          </div>

          <div className="mx-auto mt-4 grid max-w-[150px] grid-cols-2 gap-4">
            <div>
              <div className="border-b border-black">&nbsp;</div>
              <div className="mt-[1px]">BoSY Date</div>
            </div>
            <div>
              <div className="border-b border-black">&nbsp;</div>
              <div className="mt-[1px]">EoSY Date</div>
            </div>
          </div>
        </div>

        <div className="pt-[2px] text-center">
          <div className="text-left">Certified Correct:</div>
          <div className="mt-5 border-b border-black pb-[1px] font-semibold">
            {schoolHead}
          </div>
          <div className="mt-[1px]">
            (Signature of School Head over Printed Name)
          </div>

          <div className="mx-auto mt-4 grid max-w-[150px] grid-cols-2 gap-4">
            <div>
              <div className="border-b border-black">&nbsp;</div>
              <div className="mt-[1px]">BoSY Date</div>
            </div>
            <div>
              <div className="border-b border-black">&nbsp;</div>
              <div className="mt-[1px]">EoSY Date</div>
            </div>
          </div>

          <div className="mx-auto mt-4 max-w-[190px] border-b border-black pb-[1px]">
            Generated thru LIS
          </div>
        </div>
      </div>

      <div className="mt-5 text-[4px]">
        Generated on:{" "}
        {new Intl.DateTimeFormat("en-PH", {
          weekday: "long",
          month: "long",
          day: "numeric",
          year: "numeric",
        }).format(new Date())}
      </div>
    </div>
  );
}

function AdminDashboardSkeleton() {
  return (
    <div
      className="min-h-screen bg-[var(--admin-page-bg)] text-[var(--admin-color-2d211d)] lg:grid lg:grid-cols-[252px_minmax(0,1fr)]"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading admin dashboard...</span>

      <aside
        className="sticky top-0 hidden h-screen border-r border-[var(--admin-divider)] lg:block"
        aria-hidden="true"
      >
        <div className="flex h-full flex-col bg-[var(--admin-surface)]">
          <div className="flex items-center gap-3 border-b border-[var(--admin-divider)] px-5 py-5">
            <Skeleton className="size-12 shrink-0 rounded-full" />
            <div className="min-w-0 space-y-2">
              <Skeleton className="h-3 w-36" />
              <Skeleton className="h-3 w-28" />
            </div>
          </div>

          <nav className="flex-1 space-y-2 px-4 py-6">
            {Array.from({ length: 6 }).map((_, index) => (
              <div
                key={index}
                className={`flex h-10 items-center gap-3 rounded-lg px-3.5 ${
                  index === 0 ? "bg-[var(--admin-nav-hover)]" : ""
                }`}
              >
                <Skeleton className="size-4 shrink-0 rounded" />
                <Skeleton
                  className={`h-3 ${index % 2 === 0 ? "w-28" : "w-36"}`}
                />
              </div>
            ))}
          </nav>

          <div className="border-t border-[var(--admin-divider)] p-4">
            <Skeleton className="h-10 w-full rounded-lg" />
          </div>
        </div>
      </aside>

      <div className="min-w-0" aria-hidden="true">
        <header className="flex h-[72px] items-center justify-between border-b border-[var(--admin-divider)] bg-[var(--admin-surface-95)] px-4 sm:px-6 lg:px-8">
          <Skeleton className="size-10 rounded-lg lg:hidden" />
          <Skeleton className="hidden h-3 w-72 lg:block" />

          <div className="ml-auto flex items-center gap-3">
            <Skeleton className="size-10 rounded-xl" />
            <div className="flex items-center gap-2 rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-2.5 py-1.5">
              <Skeleton className="size-8 shrink-0 rounded-full" />
              <div className="hidden space-y-1.5 sm:block">
                <Skeleton className="h-3 w-28" />
                <Skeleton className="h-2.5 w-36" />
              </div>
              <Skeleton className="size-3.5 rounded" />
            </div>
          </div>
        </header>

        <main className="px-4 py-5 sm:px-6 lg:px-8 lg:py-7">
          <div className="space-y-5">
            <div className="space-y-2">
              <Skeleton className="h-7 w-64 max-w-full" />
              <Skeleton className="h-3.5 w-[30rem] max-w-full" />
            </div>

            <section className="rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-surface)] p-4 shadow-[var(--admin-shadow-panel)]">
              <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                <div className="flex items-center gap-3">
                  <Skeleton className="size-11 rounded-2xl" />
                  <div className="space-y-2">
                    <Skeleton className="h-4 w-36" />
                    <Skeleton className="h-3 w-72 max-w-full" />
                  </div>
                </div>
                <div className="grid gap-2 sm:grid-cols-[minmax(0,1fr)_150px] xl:w-[520px]">
                  <Skeleton className="h-11 w-full rounded-xl" />
                  <Skeleton className="h-11 w-full rounded-xl" />
                </div>
              </div>
              <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                {Array.from({ length: 3 }).map((_, index) => (
                  <Skeleton
                    key={`school-year-dashboard-skeleton-${index}`}
                    className="h-24 w-full rounded-2xl"
                  />
                ))}
              </div>
            </section>

            <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
              {Array.from({ length: 4 }).map((_, index) => (
                <div
                  key={index}
                  className="rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface)] p-4"
                >
                  <div className="flex items-start justify-between gap-3">
                    <div className="space-y-2">
                      <Skeleton className="h-3 w-28" />
                      <Skeleton className="h-7 w-12" />
                      <Skeleton className="h-2.5 w-10" />
                    </div>
                    <Skeleton className="size-8 shrink-0 rounded-lg" />
                  </div>
                </div>
              ))}
            </div>

            <section className="rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface)] p-4 shadow-[var(--admin-shadow-panel)]">
              <div className="mb-4 flex items-center justify-between gap-4">
                <Skeleton className="h-4 w-36" />
                <Skeleton className="h-3 w-12" />
              </div>

              <div className="divide-y divide-[var(--admin-color-f0e7df)]">
                {Array.from({ length: 5 }).map((_, index) => (
                  <div
                    key={index}
                    className="flex items-center gap-3 py-3 first:pt-0 last:pb-0"
                  >
                    <Skeleton className="size-8 shrink-0 rounded-lg" />
                    <div className="min-w-0 flex-1 space-y-1.5">
                      <Skeleton
                        className={`h-3 ${index % 2 === 0 ? "w-40" : "w-32"}`}
                      />
                      <Skeleton className="h-2.5 w-20" />
                    </div>
                    <Skeleton className="h-2.5 w-20 shrink-0" />
                  </div>
                ))}
              </div>
            </section>

            <section className="rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface)] p-4 shadow-[var(--admin-shadow-panel)]">
              <div className="mb-4 flex items-center justify-between gap-4">
                <Skeleton className="h-4 w-28" />
                <Skeleton className="h-3 w-12" />
              </div>

              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                {Array.from({ length: 4 }).map((_, index) => (
                  <div
                    key={index}
                    className="flex items-center gap-3 rounded-lg border border-[var(--admin-color-f0e6de)] bg-[var(--admin-surface-soft)] p-3"
                  >
                    <Skeleton className="size-8 shrink-0 rounded-lg" />
                    <div className="space-y-1.5">
                      <Skeleton className="h-2.5 w-20" />
                      <Skeleton className="h-5 w-10" />
                    </div>
                  </div>
                ))}
              </div>
            </section>
          </div>
        </main>
      </div>
    </div>
  );
}

function AdminDashboardPage() {
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const {
    onlineCount: onlineAdminCount,
    maxAdmins: maxOnlineAdmins,
    onlineAdminIds,
  } = useAdminOnlineStatus();
  const [activeSection, setActiveSection] = useState<AdminSection>("dashboard");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [editingProfile, setEditingProfile] = useState<ProfileRow | null>(null);
  const [createUserOpen, setCreateUserOpen] = useState(false);
  const [createStudentOpen, setCreateStudentOpen] = useState(false);
  const [editingStudent, setEditingStudent] = useState<StudentRow | null>(null);
  const [busyId, setBusyId] = useState<string | null>(null);

  const { data, isLoading, error } = useQuery<AdminData>({
    queryKey: ["admin-dashboard-data"],
    queryFn: async () => {
      const { data: userData, error: userError } =
        await supabase.auth.getUser();
      if (userError) throw userError;

      const currentUserId = userData.user?.id ?? null;
      const { data: isAdmin, error: adminError } =
        await supabase.rpc("is_admin");

      if (adminError || isAdmin !== true) {
        return {
          isAdmin: false,
          currentUserId,
          profiles: [],
          roles: [],
          classes: [],
          students: [],
          attendance: [],
          grades: [],
          credentialRequests: [],
        };
      }

      const [
        profilesRes,
        rolesRes,
        classesRes,
        studentsRes,
        gradesRes,
        credentialRequestsRes,
      ] = await Promise.all([
        supabase
          .from("profiles")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase.from("user_roles").select("user_id, role"),
        supabase
          .from("classes")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase
          .from("students")
          .select("*")
          .order("created_at", { ascending: false }),
        supabase.from("grades").select("*").limit(5000),
        (supabase as any)
          .from("student_credential_requests")
          .select("*")
          .order("request_date", { ascending: false }),
      ]);

      if (profilesRes.error) throw profilesRes.error;
      if (rolesRes.error) throw rolesRes.error;
      if (classesRes.error) throw classesRes.error;
      if (studentsRes.error) throw studentsRes.error;
      if (gradesRes.error) throw gradesRes.error;
      if (credentialRequestsRes.error) throw credentialRequestsRes.error;

      return {
        isAdmin: true,
        currentUserId,
        profiles: (profilesRes.data ?? []) as ProfileRow[],
        roles: (rolesRes.data ?? []) as RoleRow[],
        classes: (classesRes.data ?? []) as ClassRow[],
        students: (studentsRes.data ?? []) as StudentRow[],
        // attendance_records was removed from the database. Keep dependent
        // admin form views safe without querying the missing table.
        attendance: [],
        grades: (gradesRes.data ?? []) as GradeRow[],
        credentialRequests: (credentialRequestsRes.data ??
          []) as StudentCredentialRequestRow[],
      };
    },
  });

  const profiles = data?.profiles ?? EMPTY_PROFILES;
  const roles = data?.roles ?? EMPTY_ROLES;
  const classes = data?.classes ?? EMPTY_CLASSES;
  const students = data?.students ?? EMPTY_STUDENTS;
  const attendance = data?.attendance ?? EMPTY_ATTENDANCE;
  const grades = data?.grades ?? EMPTY_GRADES;
  const credentialRequests =
    data?.credentialRequests ?? EMPTY_CREDENTIAL_REQUESTS;

  const roleByUserId = useMemo(
    () => new Map(roles.map((item) => [item.user_id, item.role])),
    [roles],
  );
  const teacherProfiles = useMemo(
    () =>
      profiles.filter((profile) => roleByUserId.get(profile.id) !== "admin"),
    [profiles, roleByUserId],
  );
  const teacherById = useMemo(
    () => new Map(profiles.map((item) => [item.id, item])),
    [profiles],
  );
  const classById = useMemo(
    () => new Map(classes.map((item) => [item.id, item])),
    [classes],
  );
  const studentById = useMemo(
    () => new Map(students.map((item) => [item.id, item])),
    [students],
  );

  const currentProfile =
    profiles.find((profile) => profile.id === data?.currentUserId) ?? null;
  const currentAdminName = currentProfile?.full_name || "Admin";
  const currentAdminEmail = currentProfile?.email || "Super Administrator";

  const changeSection = (section: AdminSection) => {
    setActiveSection(section);
    setSearch("");
    setPage(1);
    setMobileMenuOpen(false);
  };

  const refreshData = async () => {
    await queryClient.invalidateQueries({ queryKey: ["admin-dashboard-data"] });
  };

  const signOut = async () => {
    await queryClient.cancelQueries();
    queryClient.clear();

    try {
      await releaseAdminAccess();
    } catch (releaseError) {
      console.warn("Unable to release the Admin online session:", releaseError);
    }

    await supabase.auth.signOut();
    navigate({ to: "/auth", replace: true });
  };

  const deleteProfile = async (profile: ProfileRow) => {
    if (profile.id === data?.currentUserId) {
      window.alert(
        "You cannot delete the administrator account that is currently signed in.",
      );
      return;
    }

    const confirmed = window.confirm(
      `Delete ${profile.full_name || profile.email || "this user"}? This also removes the user's classes and learner records.`,
    );
    if (!confirmed) return;

    setBusyId(profile.id);
    const { error: deleteError } = await supabase.rpc("admin_delete_user", {
      target_user_id: profile.id,
    });
    setBusyId(null);

    if (deleteError) {
      window.alert(deleteError.message);
      return;
    }

    await refreshData();
  };

  const deleteStudent = async (student: StudentRow) => {
    const confirmed = window.confirm(`Delete learner ${studentName(student)}?`);
    if (!confirmed) return;

    setBusyId(student.id);
    const { error: deleteError } = await supabase
      .from("students")
      .delete()
      .eq("id", student.id);
    setBusyId(null);

    if (deleteError) {
      window.alert(deleteError.message);
      return;
    }

    await refreshData();
  };

  const reviewCredentialRequest = async (
    request: StudentCredentialRequestRow,
    status: "approved" | "rejected",
    adminRemarks: string,
    resolvedStudent: StudentRow | null,
  ) => {
    setBusyId(request.id);
    const now = new Date().toISOString();
    const updatePayload: Record<string, unknown> = {
      status,
      admin_remarks: adminRemarks.trim() || null,
      reviewed_by: data?.currentUserId ?? null,
      reviewed_at: now,
    };

    if (resolvedStudent) {
      updatePayload.student_id = resolvedStudent.id;
      updatePayload.requesting_class_id = null;
      updatePayload.previous_class_id = resolvedStudent.class_id || null;
    }

    if (status === "approved") {
      updatePayload.processed_by = data?.currentUserId ?? null;
      updatePayload.processed_at = now;
      updatePayload.release_notes = adminRemarks.trim() || null;
    }

    const { error: updateError } = await (supabase as any)
      .from("student_credential_requests")
      .update(updatePayload)
      .eq("id", request.id);
    setBusyId(null);

    if (updateError) {
      window.alert(updateError.message);
      return;
    }

    await refreshData();
  };

  if (isLoading) {
    return <AdminDashboardSkeleton />;
  }

  if (error) {
    return (
      <div className="grid min-h-screen place-items-center bg-[var(--admin-page-bg)] p-6">
        <div className="max-w-lg rounded-2xl border border-[var(--admin-danger-border)] bg-[var(--admin-danger-bg)] p-6 text-sm text-[var(--admin-danger-700)]">
          {(error as Error).message}
        </div>
      </div>
    );
  }

  if (data?.isAdmin === false) {
    return (
      <div className="grid min-h-screen place-items-center bg-[var(--admin-page-bg)] p-6">
        <div className="max-w-lg rounded-2xl border border-[var(--admin-danger-border)] bg-[var(--admin-danger-bg)] p-6 text-sm text-[var(--admin-danger-700)]">
          You do not have admin access.
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[var(--admin-page-bg)] text-[var(--admin-color-2d211d)] lg:grid lg:grid-cols-[252px_minmax(0,1fr)]">
      <AdminSidebar
        activeSection={activeSection}
        onChangeSection={changeSection}
        onSignOut={signOut}
        mobileOpen={mobileMenuOpen}
        onCloseMobile={() => setMobileMenuOpen(false)}
      />

      <div className="min-w-0">
        <AdminTopbar
          key={activeSection}
          adminName={currentAdminName}
          adminSubtitle={currentAdminEmail}
          avatarUrl={currentProfile?.avatar_url || null}
          onOpenMenu={() => setMobileMenuOpen(true)}
          onOpenProfile={() => changeSection("profile")}
          onOpenNotifications={() => changeSection("notifications")}
        />

        <main className="px-4 py-5 sm:px-6 lg:px-8 lg:py-7">
          {activeSection === "dashboard" && (
            <DashboardSection
              profiles={profiles}
              teacherProfiles={teacherProfiles}
              classes={classes}
              students={students}
              grades={grades}
              roleByUserId={roleByUserId}
              onlineAdminCount={onlineAdminCount}
              maxOnlineAdmins={maxOnlineAdmins}
              onChangeSection={changeSection}
            />
          )}

          {activeSection === "users" && (
            <UsersSection
              profiles={profiles}
              roleByUserId={roleByUserId}
              search={search}
              onSearch={setSearch}
              page={page}
              onPage={setPage}
              currentUserId={data?.currentUserId ?? null}
              onlineAdminCount={onlineAdminCount}
              maxOnlineAdmins={maxOnlineAdmins}
              onlineAdminIds={onlineAdminIds}
              busyId={busyId}
              onAdd={() => setCreateUserOpen(true)}
              onEdit={setEditingProfile}
              onDelete={deleteProfile}
            />
          )}

          {activeSection === "students" && (
            <StudentsSection
              students={students}
              teacherById={teacherById}
              classById={classById}
              search={search}
              onSearch={setSearch}
              page={page}
              onPage={setPage}
              busyId={busyId}
              onAdd={() => setCreateStudentOpen(true)}
              onEdit={setEditingStudent}
              onDelete={deleteStudent}
            />
          )}

          {activeSection === "credential_requests" && (
            <CredentialRequestsAdminSection
              requests={credentialRequests}
              students={students}
              studentById={studentById}
              teacherById={teacherById}
              classById={classById}
              busyId={busyId}
              onReview={reviewCredentialRequest}
            />
          )}

          {activeSection === "forms" && (
            <SchoolFormsSection
              classes={classes}
              students={students}
              attendance={attendance}
              grades={grades}
              teacherById={teacherById}
              schoolProfile={currentProfile}
              search={search}
              onSearch={setSearch}
            />
          )}

          {activeSection === "analytics" && <AdminAnalyticsInsights />}

          {activeSection === "profile" && currentProfile && (
            <AdminProfileSection
              profile={currentProfile}
              onEdit={() => setEditingProfile(currentProfile)}
              onSaved={refreshData}
            />
          )}

          {activeSection === "notifications" && (
            <AdminNotificationsSection userId={data?.currentUserId ?? null} />
          )}
        </main>
      </div>

      <UserCreateDialog
        open={createUserOpen}
        onClose={() => setCreateUserOpen(false)}
        onSaved={refreshData}
      />

      <ProfileEditDialog
        profile={editingProfile}
        onClose={() => setEditingProfile(null)}
        onSaved={refreshData}
      />
      <StudentEditDialog
        student={editingStudent}
        classes={classes}
        onClose={() => setEditingStudent(null)}
        onSaved={refreshData}
      />
      <StudentCreateDialog
        open={createStudentOpen}
        onClose={() => setCreateStudentOpen(false)}
        onSaved={refreshData}
      />
    </div>
  );
}

type AdminNotificationRow = {
  id: string;
  title: string;
  message: string;
  category: string;
  related_path: string | null;
  read_at: string | null;
  created_at: string;
};

function AdminNotificationsSection({ userId }: { userId: string | null }) {
  const queryClient = useQueryClient();
  const [notificationSearch, setNotificationSearch] = useState("");
  const [notificationStatus, setNotificationStatus] = useState<
    "all" | "unread" | "read"
  >("all");

  const notificationQueryKey = ["admin-notifications", userId] as const;
  const { data: notifications = [], isLoading } = useQuery<
    AdminNotificationRow[]
  >({
    queryKey: notificationQueryKey,
    enabled: Boolean(userId),
    queryFn: async () => {
      if (!userId) return [];
      const { data: rows, error: notificationsError } = await (supabase as any)
        .from("notifications")
        .select(
          "id, title, message, category, related_path, read_at, created_at",
        )
        .eq("recipient_id", userId)
        .order("created_at", { ascending: false });

      if (notificationsError) throw notificationsError;
      return (rows ?? []) as AdminNotificationRow[];
    },
  });

  const visibleNotifications = useMemo(() => {
    const term = notificationSearch.trim().toLowerCase();
    return notifications.filter((notification) => {
      const matchesSearch =
        !term ||
        notification.title.toLowerCase().includes(term) ||
        notification.message.toLowerCase().includes(term) ||
        notification.category.toLowerCase().includes(term);
      const matchesStatus =
        notificationStatus === "all" ||
        (notificationStatus === "unread" && !notification.read_at) ||
        (notificationStatus === "read" && Boolean(notification.read_at));
      return matchesSearch && matchesStatus;
    });
  }, [notificationSearch, notificationStatus, notifications]);

  const unreadCount = notifications.filter(
    (notification) => !notification.read_at,
  ).length;

  const markOneRead = async (notification: AdminNotificationRow) => {
    if (notification.read_at) return;
    const { error: markError } = await (supabase as any).rpc(
      "mark_notification_read",
      {
        p_notification_id: notification.id,
      },
    );
    if (markError) {
      window.alert(markError.message);
      return;
    }
    await queryClient.invalidateQueries({ queryKey: notificationQueryKey });
    await queryClient.invalidateQueries({
      queryKey: ["notification-bell", userId],
    });
  };

  const markAllRead = async () => {
    const { error: markError } = await (supabase as any).rpc(
      "mark_all_notifications_read",
    );
    if (markError) {
      window.alert(markError.message);
      return;
    }
    await queryClient.invalidateQueries({ queryKey: notificationQueryKey });
    await queryClient.invalidateQueries({
      queryKey: ["notification-bell", userId],
    });
  };

  return (
    <section className="space-y-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-2xl font-bold text-[var(--admin-text-heading)]">
            Notifications
          </h1>
          <p className="mt-1 text-sm text-[var(--admin-text-muted)]">
            Admin alerts and activity history. {unreadCount} unread.
          </p>
        </div>
        <Button
          type="button"
          variant="outline"
          disabled={unreadCount === 0}
          onClick={markAllRead}
          className="gap-2"
        >
          <CircleCheck className="size-4" />
          Mark all as read
        </Button>
      </div>

      <div className="grid gap-3 rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-4 shadow-sm sm:grid-cols-[1fr_180px]">
        <div className="relative">
          <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          <Input
            value={notificationSearch}
            onChange={(event) => setNotificationSearch(event.target.value)}
            placeholder="Search notifications..."
            className="pl-9"
          />
        </div>
        <select
          value={notificationStatus}
          onChange={(event) =>
            setNotificationStatus(
              event.target.value as "all" | "unread" | "read",
            )
          }
          className="h-10 rounded-md border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-sm"
        >
          <option value="all">All notifications</option>
          <option value="unread">Unread only</option>
          <option value="read">Read only</option>
        </select>
      </div>

      <div className="overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-sm">
        {isLoading ? (
          <div className="flex items-center justify-center gap-2 p-12 text-sm text-[var(--admin-text-muted)]">
            <Loader2 className="size-5 animate-spin" /> Loading notifications...
          </div>
        ) : visibleNotifications.length === 0 ? (
          <div className="p-12 text-center text-sm text-[var(--admin-text-muted)]">
            No notifications found.
          </div>
        ) : (
          <div className="divide-y divide-[var(--admin-divider)]">
            {visibleNotifications.map((notification) => (
              <button
                key={notification.id}
                type="button"
                onClick={() => markOneRead(notification)}
                className={`block w-full p-5 text-left transition hover:bg-[var(--admin-nav-hover)] ${
                  notification.read_at ? "" : "bg-[var(--admin-color-fdf8f2)]"
                }`}
              >
                <div className="flex items-start gap-3">
                  <span
                    className={`mt-2 size-2 shrink-0 rounded-full ${
                      notification.read_at
                        ? "bg-transparent"
                        : "bg-[var(--admin-primary)]"
                    }`}
                  />
                  <div className="min-w-0 flex-1">
                    <div className="flex flex-col gap-1 sm:flex-row sm:items-start sm:justify-between">
                      <h2 className="font-semibold text-[var(--admin-text-heading)]">
                        {notification.title}
                      </h2>
                      <time className="shrink-0 text-xs text-[var(--admin-text-muted)]">
                        {formatDateTime(notification.created_at).date},{" "}
                        {formatDateTime(notification.created_at).time}
                      </time>
                    </div>
                    <p className="mt-1 text-sm leading-6 text-[var(--admin-color-735e55)]">
                      {notification.message}
                    </p>
                    <span className="mt-2 inline-block rounded-full bg-[var(--admin-nav-hover)] px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wide text-[var(--admin-primary)]">
                      {notification.category.split("_").join(" ")}
                    </span>
                  </div>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </section>
  );
}

function AdminSidebar({
  activeSection,
  onChangeSection,
  onSignOut,
  mobileOpen,
  onCloseMobile,
}: {
  activeSection: AdminSection;
  onChangeSection: (section: AdminSection) => void;
  onSignOut: () => void;
  mobileOpen: boolean;
  onCloseMobile: () => void;
}) {
  const content = (
    <div className="flex h-full flex-col bg-[var(--admin-surface)]">
      <div className="flex items-center gap-3 border-b border-[var(--admin-divider)] px-5 py-5">
        <img
          src={logo}
          alt="Agusan del Sur National Science High School"
          className="size-12 object-contain"
        />
        <div className="min-w-0 text-[11px] font-semibold leading-4 text-[var(--admin-text-value)]">
          Agusan del Sur National
          <br />
          Science High School
        </div>
      </div>

      <nav className="flex-1 space-y-2 px-4 py-6">
        {NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          const active = item.id === activeSection;
          return (
            <button
              key={item.id}
              type="button"
              onClick={() => onChangeSection(item.id)}
              className={`flex w-full items-center gap-3 rounded-lg px-3.5 py-3 text-left text-xs font-semibold transition ${
                active
                  ? "bg-[var(--admin-primary)] text-[var(--admin-white)] shadow-[var(--admin-shadow-active)]"
                  : "text-[var(--admin-color-4f403a)] hover:bg-[var(--admin-nav-hover)] hover:text-[var(--admin-primary)]"
              }`}
            >
              <Icon className="size-4" />
              <span>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div className="border-t border-[var(--admin-divider)] p-4">
        <button
          type="button"
          onClick={onSignOut}
          className="flex w-full items-center justify-center gap-2 rounded-lg border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 py-2.5 text-xs font-semibold text-[var(--admin-primary-strong)] hover:bg-[var(--admin-danger-bg)]"
        >
          <LogOut className="size-4" />
          Log out
        </button>
      </div>
    </div>
  );

  return (
    <>
      <aside className="sticky top-0 hidden h-screen border-r border-[var(--admin-divider)] lg:block">
        {content}
      </aside>

      {mobileOpen && (
        <div className="fixed inset-0 z-50 lg:hidden">
          <button
            type="button"
            aria-label="Close admin menu"
            className="absolute inset-0 bg-[var(--admin-overlay)]"
            onClick={onCloseMobile}
          />
          <aside className="relative h-full w-[280px] max-w-[86vw] border-r border-[var(--admin-divider)] shadow-2xl">
            <button
              type="button"
              aria-label="Close admin menu"
              onClick={onCloseMobile}
              className="absolute right-3 top-3 z-10 rounded-lg border border-[var(--admin-border)] bg-[var(--admin-white)] p-2 text-[var(--admin-text-icon)]"
            >
              <X className="size-4" />
            </button>
            {content}
          </aside>
        </div>
      )}
    </>
  );
}

function AdminTopbar({
  adminName,
  adminSubtitle,
  avatarUrl,
  onOpenMenu,
  onOpenProfile,
  onOpenNotifications,
}: {
  adminName: string;
  adminSubtitle: string;
  avatarUrl: string | null;
  onOpenMenu: () => void;
  onOpenProfile: () => void;
  onOpenNotifications: () => void;
}) {
  const [profileMenuOpen, setProfileMenuOpen] = useState(false);

  return (
    <header className="sticky top-0 z-30 flex h-[72px] items-center justify-between border-b border-[var(--admin-divider)] bg-[var(--admin-surface-95)] px-4 backdrop-blur sm:px-6 lg:px-8">
      <button
        type="button"
        onClick={onOpenMenu}
        className="rounded-lg border border-[var(--admin-border)] bg-[var(--admin-white)] p-2 text-[var(--admin-text-icon)] lg:hidden"
        aria-label="Open admin menu"
      >
        <Menu className="size-5" />
      </button>

      <div className="hidden text-xs text-[var(--admin-text-muted)] lg:block">
        School Information, Grading &amp; Learner's Assessment
      </div>

      <div className="ml-auto flex items-center gap-3">
        <div
          onClickCapture={(event) => {
            const target = event.target;
            if (!(target instanceof Element)) return;

            const action = target.closest("a, button");
            const isTeacherNotificationRoute =
              action?.getAttribute("href") === "/notifications";
            const isViewAllAction =
              action?.textContent?.trim().toLowerCase() ===
              "view all notifications";

            if (isTeacherNotificationRoute || isViewAllAction) {
              event.preventDefault();
              event.stopPropagation();
              onOpenNotifications();
            }
          }}
        >
          <NotificationBell className="border-0 bg-transparent text-[var(--admin-text-icon)] shadow-none hover:bg-[var(--admin-nav-hover)]" />
        </div>

        <div className="relative">
          <button
            type="button"
            onClick={() => setProfileMenuOpen((open) => !open)}
            className="flex items-center gap-2 rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-2.5 py-1.5 text-left shadow-sm transition hover:bg-[var(--admin-nav-hover)]"
            aria-label="Open admin profile"
            aria-expanded={profileMenuOpen}
          >
            {avatarUrl ? (
              <img
                src={avatarUrl}
                alt={adminName}
                className="size-8 shrink-0 rounded-full object-cover"
              />
            ) : (
              <div className="grid size-8 shrink-0 place-items-center rounded-full bg-[var(--admin-color-f5e7d7)] text-[var(--admin-primary)]">
                <UserRound className="size-4" />
              </div>
            )}
            <div className="hidden min-w-0 sm:block">
              <div className="max-w-[180px] truncate text-[11px] font-semibold text-[var(--admin-text-heading)]">
                {adminName}
              </div>
              <div className="max-w-[180px] truncate text-[9px] text-[var(--admin-text-muted)]">
                {adminSubtitle}
              </div>
            </div>
            <ChevronDown
              className={`size-3.5 text-[var(--admin-text-muted)] transition-transform ${
                profileMenuOpen ? "rotate-180" : ""
              }`}
            />
          </button>

          {profileMenuOpen && (
            <div className="absolute right-0 top-[calc(100%+8px)] z-50 w-56 overflow-hidden rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-1.5 shadow-lg">
              <div className="border-b border-[var(--admin-divider)] px-3 py-2 sm:hidden">
                <p className="truncate text-xs font-semibold text-[var(--admin-text-heading)]">
                  {adminName}
                </p>
                <p className="truncate text-[10px] text-[var(--admin-text-muted)]">
                  {adminSubtitle}
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setProfileMenuOpen(false);
                  onOpenProfile();
                }}
                className="flex w-full items-center gap-2 rounded-lg px-3 py-2.5 text-left text-xs font-semibold text-[var(--admin-color-4f403a)] hover:bg-[var(--admin-nav-hover)]"
              >
                <UserRound className="size-4 text-[var(--admin-primary)]" />
                My Profile
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}

function AdminProfileSection({
  profile,
  onEdit,
  onSaved,
}: {
  profile: ProfileRow;
  onEdit: () => void;
  onSaved: () => Promise<void>;
}) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [uploading, setUploading] = useState(false);
  const [deleting, setDeleting] = useState(false);

  const uploadProfileImage = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    if (!file.type.startsWith("image/")) {
      window.alert("Please select an image file.");
      return;
    }

    if (file.size > 5 * 1024 * 1024) {
      window.alert("The profile image must be 5 MB or smaller.");
      return;
    }

    setUploading(true);
    const extension = file.name.split(".").pop()?.toLowerCase() || "jpg";
    const filePath = `${profile.id}/admin-profile-${Date.now()}.${extension}`;
    const { error: uploadError } = await supabase.storage
      .from("avatars")
      .upload(filePath, file, { cacheControl: "3600", upsert: true });

    if (uploadError) {
      setUploading(false);
      window.alert(
        `${uploadError.message}\n\nMake sure the Supabase Storage bucket named "avatars" exists and allows the signed-in user to upload.`,
      );
      return;
    }

    const { data: publicUrlData } = supabase.storage
      .from("avatars")
      .getPublicUrl(filePath);
    const { error: updateError } = await (supabase as any)
      .from("profiles")
      .update({ avatar_url: publicUrlData.publicUrl })
      .eq("id", profile.id);

    setUploading(false);
    if (updateError) {
      await supabase.storage.from("avatars").remove([filePath]);
      window.alert(updateError.message);
      return;
    }

    await onSaved();
  };

  const deleteProfileImage = async () => {
    if (!profile.avatar_url) return;
    if (!window.confirm("Delete your admin profile image?")) return;

    setDeleting(true);
    const { error: updateError } = await (supabase as any)
      .from("profiles")
      .update({ avatar_url: null })
      .eq("id", profile.id);

    if (updateError) {
      setDeleting(false);
      window.alert(updateError.message);
      return;
    }

    const marker = "/storage/v1/object/public/avatars/";
    const storedPath = profile.avatar_url.includes(marker)
      ? decodeURIComponent(profile.avatar_url.split(marker)[1])
      : "";
    if (storedPath) {
      await supabase.storage.from("avatars").remove([storedPath]);
    }

    setDeleting(false);
    await onSaved();
  };

  const information = [
    {
      label: "Full Name",
      value: profile.full_name || "Not provided",
      icon: UserRound,
    },
    {
      label: "Email Address",
      value: profile.email || "Not provided",
      icon: IdCard,
    },
    { label: "User Type", value: "Administrator", icon: ShieldCheck },
    {
      label: "School Name",
      value:
        profile.school_name || "Agusan del Sur National Science High School",
      icon: GraduationCap,
    },
    {
      label: "School ID",
      value: profile.school_id || "Not provided",
      icon: IdCard,
    },
    {
      label: "Region",
      value: profile.region || "Not provided",
      icon: FolderOpen,
    },
    {
      label: "Division",
      value: profile.division || "Not provided",
      icon: FolderOpen,
    },
    {
      label: "District",
      value: profile.district || "Not provided",
      icon: FolderOpen,
    },
  ];

  return (
    <div className="mx-auto max-w-5xl space-y-5">
      <PageHeading
        title="My Profile"
        description="View and manage your administrator profile information."
      />

      <Panel className="overflow-hidden border-[#ead8b8] bg-[#fffdf9] p-0 shadow-[0_14px_35px_rgba(102,63,38,0.10)]">
        <div className="relative isolate overflow-hidden bg-gradient-to-r from-[#f5e5e2] via-[#fff8ef] to-[#fffdf8] p-6 sm:p-8">
          {/* Decorative background matching the Class Adviser / Subject Teacher profile */}
          <div className="pointer-events-none absolute -left-20 -bottom-28 size-56 rounded-full border border-[#e8bdb4]/70" />
          <div className="pointer-events-none absolute -left-10 -bottom-20 size-40 rounded-full bg-[#f5c8c2]/30" />

          <div className="pointer-events-none absolute -right-20 -top-24 size-60 rounded-full border border-[#e6c37f]/55" />
          <div className="pointer-events-none absolute -right-12 -top-16 size-44 rounded-full bg-[#f5c7b8]/35" />

          <div className="pointer-events-none absolute left-5 top-5 grid grid-cols-7 gap-2 opacity-35">
            {Array.from({ length: 35 }).map((_, index) => (
              <span
                key={`admin-profile-left-dot-${index}`}
                className="size-1.5 rounded-full bg-[#d98f79]"
              />
            ))}
          </div>

          <div className="pointer-events-none absolute right-8 bottom-5 grid grid-cols-6 gap-2 opacity-35">
            {Array.from({ length: 24 }).map((_, index) => (
              <span
                key={`admin-profile-right-dot-${index}`}
                className="size-1.5 rounded-full bg-[#d98f79]"
              />
            ))}
          </div>

          <div className="relative z-10 flex flex-col gap-6 sm:flex-row sm:items-center">
            <div className="size-28 shrink-0 overflow-hidden rounded-3xl bg-[#8f2928] text-white shadow-[0_12px_28px_rgba(143,41,40,0.24)]">
              {profile.avatar_url ? (
                <img
                  src={profile.avatar_url}
                  alt={profile.full_name || "Administrator"}
                  className="size-full object-cover"
                />
              ) : (
                <div className="grid size-full place-items-center">
                  <UserRound className="size-14" />
                </div>
              )}
            </div>

            <div className="min-w-0 flex-1">
              <h2 className="text-2xl font-bold text-[#2f1c15] sm:text-3xl">
                {profile.full_name || "Administrator"}
              </h2>

              <p className="mt-1 text-sm font-semibold text-[#8f2928]">
                Administrator
              </p>

              <p className="mt-1 text-sm text-[#7d6657]">
                {profile.school_name ||
                  "Agusan del Sur National Science High School"}
              </p>

              <div className="mt-5 flex flex-wrap gap-2">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/png,image/jpeg,image/webp"
                  className="hidden"
                  onChange={uploadProfileImage}
                />

                <Button
                  type="button"
                  variant="outline"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={uploading || deleting}
                  className="border-[#ead8b8] bg-[#fffaf0]/95 text-[#3c241b] shadow-sm hover:bg-white"
                >
                  {uploading ? (
                    <Loader2 className="animate-spin" />
                  ) : (
                    <Pencil />
                  )}
                  Edit Profile Image
                </Button>

                <Button
                  type="button"
                  variant="outline"
                  onClick={deleteProfileImage}
                  disabled={!profile.avatar_url || uploading || deleting}
                  className="border-[#f2cbc5] bg-[#fff8f6]/90 text-[#e25c59] shadow-sm hover:bg-[#fff1ee]"
                >
                  {deleting ? (
                    <Loader2 className="animate-spin" />
                  ) : (
                    <Trash2 />
                  )}
                  Delete Image
                </Button>
              </div>
            </div>
          </div>
        </div>
      </Panel>

      <Panel className="border-[#ead8b8] bg-[#fffdf9] shadow-[0_10px_28px_rgba(102,63,38,0.07)]">
        <PanelHeader
          title="Account Information"
          action={
            <Button
              type="button"
              variant="outline"
              onClick={onEdit}
              className="border-[#ead8b8] bg-[#fffaf0] text-[#8f2928] shadow-sm hover:bg-white"
            >
              <Pencil className="size-4" />
              Edit Information
            </Button>
          }
        />

        <div className="grid gap-3 sm:grid-cols-2">
          {information.map(({ label, value, icon: Icon }) => (
            <div
              key={label}
              className="flex items-start gap-3 rounded-2xl border border-[#ead8b8] bg-gradient-to-br from-[#fff9ee] to-[#fffdf8] p-4 shadow-[0_5px_14px_rgba(116,77,47,0.05)]"
            >
              <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-[#f7eadb] text-[#8f2928] shadow-sm">
                <Icon className="size-4" />
              </div>
              <div className="min-w-0">
                <p className="text-[9px] font-semibold uppercase tracking-wide text-[#7d6657]">
                  {label}
                </p>
                <p className="mt-1 break-words text-xs font-semibold text-[#2f1c15]">
                  {value}
                </p>
              </div>
            </div>
          ))}
        </div>
      </Panel>
    </div>
  );
}


function SchoolYearLibraryCard() {
  const queryClient = useQueryClient();
  const [newSchoolYear, setNewSchoolYear] = useState("");
  const [busyAction, setBusyAction] = useState<string | null>(null);
  const [schoolYearConfirm, setSchoolYearConfirm] = useState<{
    type: "lock" | "unlock" | "delete";
    row: AdminSchoolYearLibraryRow;
  } | null>(null);

  const {
    data: schoolYears = [],
    isLoading: schoolYearsLoading,
    error: schoolYearsError,
  } = useQuery<AdminSchoolYearLibraryRow[]>({
    queryKey: ["school-year-library"],
    queryFn: async () => {
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
    },
    refetchOnWindowFocus: true,
  });

  const {
    data: schoolYearSetting,
    isLoading: schoolYearSettingLoading,
    error: schoolYearSettingError,
  } = useQuery<AdminSchoolYearSetting>({
    queryKey: ["school-year-setting"],
    queryFn: async () => {
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
        updated_at: data?.updated_at ?? null,
      };
    },
    refetchOnWindowFocus: true,
  });

  const lockedSchoolYear =
    schoolYearSetting?.is_locked && schoolYearSetting.school_year
      ? schoolYearSetting.school_year
      : schoolYears.find((row) => row.is_locked)?.school_year ?? null;

  const refreshSchoolYears = async () => {
    await Promise.all([
      queryClient.invalidateQueries({ queryKey: ["school-year-library"] }),
      queryClient.invalidateQueries({ queryKey: ["school-year-setting"] }),
      queryClient.invalidateQueries({
        queryKey: ["admin-school-form-submissions"],
      }),
    ]);
  };

  const normalizedSchoolYear = (value: string) =>
    value.replace(/\s+/g, "").trim();

  const validateSchoolYear = (value: string) => {
    const normalized = normalizedSchoolYear(value);
    const match = /^(\d{4})-(\d{4})$/.exec(normalized);

    if (!match) {
      return {
        valid: false as const,
        value: normalized,
        message: "Use the School Year format YYYY-YYYY, for example 2027-2028.",
      };
    }

    const startYear = Number(match[1]);
    const endYear = Number(match[2]);

    if (endYear !== startYear + 1) {
      return {
        valid: false as const,
        value: normalized,
        message: "The second year must be exactly one year after the first year.",
      };
    }

    return {
      valid: true as const,
      value: `${startYear}-${endYear}`,
      message: "",
    };
  };

  const addSchoolYear = async () => {
    const validation = validateSchoolYear(newSchoolYear);

    if (!validation.valid) {
      window.alert(validation.message);
      return;
    }

    if (
      schoolYears.some(
        (row) => normalize(row.school_year) === normalize(validation.value),
      )
    ) {
      window.alert(`${validation.value} is already in the School Year Library.`);
      return;
    }

    setBusyAction("add");

    try {
      const { error } = await (supabase as any)
        .from("school_year_library")
        .insert({
          school_year: validation.value,
          is_locked: false,
        });

      if (error) throw error;

      setNewSchoolYear("");
      await refreshSchoolYears();
    } catch (error) {
      window.alert(
        error instanceof Error
          ? error.message
          : "Unable to add the School Year.",
      );
    } finally {
      setBusyAction(null);
    }
  };

  const lockSchoolYear = async (row: AdminSchoolYearLibraryRow) => {
    // The selected year is already the active locked School Year.
    if (
      lockedSchoolYear &&
      normalize(lockedSchoolYear) === normalize(row.school_year)
    ) {
      return;
    }

    // Do not switch automatically. The current locked School Year
    // must be unlocked before another one can be locked.
    if (lockedSchoolYear) {
      return;
    }

    setBusyAction(`lock-${row.id}`);

    try {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;

      // Safety rule: clear every existing lock before locking the selected row.
      const ids = schoolYears.map((item) => item.id);

      if (ids.length > 0) {
        const { error: clearLibraryError } = await (supabase as any)
          .from("school_year_library")
          .update({ is_locked: false })
          .in("id", ids);

        if (clearLibraryError) throw clearLibraryError;
      }

      const { error: lockLibraryError } = await (supabase as any)
        .from("school_year_library")
        .update({ is_locked: true })
        .eq("id", row.id);

      if (lockLibraryError) throw lockLibraryError;

      const { error: settingError } = await (supabase as any)
        .from("school_year_settings")
        .upsert(
          {
            id: 1,
            school_year: row.school_year,
            is_locked: true,
            updated_by: user?.id ?? null,
            updated_at: new Date().toISOString(),
          },
          { onConflict: "id" },
        );

      if (settingError) {
        // Remove the selected library lock if saving the active setting fails.
        await (supabase as any)
          .from("school_year_library")
          .update({ is_locked: false })
          .eq("id", row.id);

        throw settingError;
      }

      await refreshSchoolYears();
    } catch (error) {
      window.alert(
        error instanceof Error
          ? error.message
          : "Unable to lock the School Year.",
      );
    } finally {
      setBusyAction(null);
    }
  };

  const unlockSchoolYear = async (row: AdminSchoolYearLibraryRow) => {
    // Only the currently active locked School Year can be unlocked.
    if (
      !lockedSchoolYear ||
      normalize(lockedSchoolYear) !== normalize(row.school_year)
    ) {
      return;
    }

    setBusyAction(`unlock-${row.id}`);

    try {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;

      // Clear all lock flags, including any accidental old duplicate lock data.
      const ids = schoolYears.map((item) => item.id);

      if (ids.length > 0) {
        const { error: clearLibraryError } = await (supabase as any)
          .from("school_year_library")
          .update({ is_locked: false })
          .in("id", ids);

        if (clearLibraryError) throw clearLibraryError;
      }

      const { error: settingError } = await (supabase as any)
        .from("school_year_settings")
        .upsert(
          {
            id: 1,
            school_year: null,
            is_locked: false,
            updated_by: user?.id ?? null,
            updated_at: new Date().toISOString(),
          },
          { onConflict: "id" },
        );

      if (settingError) {
        // Restore the previous lock if updating the active setting fails.
        await (supabase as any)
          .from("school_year_library")
          .update({ is_locked: true })
          .eq("id", row.id);

        throw settingError;
      }

      await refreshSchoolYears();
    } catch (error) {
      window.alert(
        error instanceof Error
          ? error.message
          : "Unable to unlock the School Year.",
      );
    } finally {
      setBusyAction(null);
    }
  };

  const deleteSchoolYear = async (row: AdminSchoolYearLibraryRow) => {
    const isLocked =
      lockedSchoolYear &&
      normalize(lockedSchoolYear) === normalize(row.school_year);

    if (isLocked) {
      window.alert(
        `${row.school_year} is currently locked. Unlock it before deleting it.`,
      );
      return;
    }

    setBusyAction(`delete-${row.id}`);

    try {
      // Remove School Form submissions for this School Year first.
      // Historical classes, learners, and grades are intentionally preserved.
      const { error: submissionsDeleteError } = await (supabase as any)
        .from("school_form_submissions")
        .delete()
        .eq("school_year", row.school_year);

      if (submissionsDeleteError) throw submissionsDeleteError;

      const { error: libraryDeleteError } = await (supabase as any)
        .from("school_year_library")
        .delete()
        .eq("id", row.id);

      if (libraryDeleteError) throw libraryDeleteError;

      // If an old unlocked setting still points at the deleted year, clear it.
      if (
        schoolYearSetting?.school_year &&
        normalize(schoolYearSetting.school_year) === normalize(row.school_year)
      ) {
        const { error: settingCleanupError } = await (supabase as any)
          .from("school_year_settings")
          .update({
            school_year: null,
            is_locked: false,
            updated_at: new Date().toISOString(),
          })
          .eq("id", 1);

        if (settingCleanupError) {
          console.warn(
            "School Year was deleted, but the old School Year setting could not be cleared:",
            settingCleanupError,
          );
        }
      }

      await refreshSchoolYears();
    } catch (error) {
      window.alert(
        error instanceof Error
          ? error.message
          : "Unable to delete the School Year and its School Forms.",
      );
    } finally {
      setBusyAction(null);
    }
  };

  const loading = schoolYearsLoading || schoolYearSettingLoading;
  const loadError = schoolYearsError || schoolYearSettingError;

  return (
    <section className="overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-[var(--admin-shadow-panel)]">
      <div className="flex flex-col gap-4 border-b border-[var(--admin-divider)] px-5 py-4 xl:flex-row xl:items-center xl:justify-between">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-11 shrink-0 place-items-center rounded-2xl bg-amber-50 text-amber-700">
            <LockKeyhole className="size-5" />
          </div>

          <div className="min-w-0">
            <h2 className="text-base font-bold text-[var(--admin-text-heading)]">
              School Year Library
            </h2>
            <p className="mt-1 text-xs text-[var(--admin-text-muted)]">
              Add School Years to your library. Lock or unlock one School Year
              at a time. Unlock the current year before locking another.
            </p>
          </div>
        </div>

        <div className="grid w-full gap-2 sm:grid-cols-[minmax(0,1fr)_auto] xl:w-[580px]">
          <Input
            value={newSchoolYear}
            onChange={(event) =>
              setNewSchoolYear(
                event.target.value.replace(/[^\d-\s]/g, "").slice(0, 9),
              )
            }
            onKeyDown={(event) => {
              if (event.key === "Enter") void addSchoolYear();
            }}
            placeholder="Example: 2027-2028"
            inputMode="numeric"
            aria-label="New School Year"
            disabled={Boolean(busyAction)}
            className="h-11 rounded-xl"
          />

          <Button
            type="button"
            onClick={() => void addSchoolYear()}
            disabled={Boolean(busyAction) || !newSchoolYear.trim()}
            className="h-11 gap-2 rounded-xl bg-[var(--admin-primary)] px-5 hover:bg-[var(--admin-primary-hover)]"
          >
            {busyAction === "add" ? (
              <Loader2 className="size-4 animate-spin" />
            ) : (
              <Plus className="size-4" />
            )}
            Add School Year
          </Button>
        </div>
      </div>

      <div className="p-5">
        {loading ? (
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
            {Array.from({ length: 3 }).map((_, index) => (
              <div
                key={`school-year-loading-${index}`}
                className="rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-4"
              >
                <Skeleton className="h-5 w-28" />
                <Skeleton className="mt-3 h-6 w-20 rounded-full" />
              </div>
            ))}
          </div>
        ) : loadError ? (
          <div className="rounded-xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-700">
            {loadError instanceof Error
              ? loadError.message
              : "Unable to load the School Year Library."}
          </div>
        ) : schoolYears.length === 0 ? (
          <div className="rounded-xl border border-dashed border-[var(--admin-border)] bg-[var(--admin-surface-soft)] px-4 py-8 text-center text-sm text-[var(--admin-text-muted)]">
            No School Years yet. Add your first School Year above.
          </div>
        ) : (
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
            {schoolYears.map((row) => {
              const isLocked = lockedSchoolYear === row.school_year;
              const isBusy =
                busyAction === `lock-${row.id}` ||
                busyAction === `unlock-${row.id}` ||
                busyAction === `delete-${row.id}`;

              return (
                <div
                  key={row.id}
                  className={`flex items-center gap-4 rounded-2xl border p-4 transition ${
                    isLocked
                      ? "border-emerald-300 bg-emerald-50/40"
                      : "border-[var(--admin-border)] bg-[var(--admin-surface-soft)]"
                  }`}
                >
                  <div
                    className={`grid size-11 shrink-0 place-items-center rounded-2xl ${
                      isLocked
                        ? "bg-emerald-100 text-emerald-700"
                        : "bg-amber-50 text-amber-700"
                    }`}
                  >
                    {isLocked ? (
                      <LockKeyhole className="size-5" />
                    ) : (
                      <UnlockKeyhole className="size-5" />
                    )}
                  </div>

                  <div className="min-w-0 flex-1">
                    <div className="truncate text-base font-bold text-[var(--admin-text-heading)]">
                      {row.school_year}
                    </div>
                    <span
                      className={`mt-2 inline-flex rounded-full px-2.5 py-1 text-[9px] font-bold ${
                        isLocked
                          ? "bg-emerald-100 text-emerald-700"
                          : "bg-amber-100 text-amber-700"
                      }`}
                    >
                      {isLocked ? "Locked" : "Unlocked"}
                    </span>
                  </div>

                  <div className="flex shrink-0 items-center gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        setSchoolYearConfirm({
                          type: isLocked ? "unlock" : "lock",
                          row,
                        })
                      }
                      disabled={
                        Boolean(busyAction) ||
                        (!isLocked && Boolean(lockedSchoolYear))
                      }
                      title={
                        isLocked
                          ? `Unlock ${row.school_year}`
                          : lockedSchoolYear
                            ? `Unlock ${lockedSchoolYear} first before locking ${row.school_year}`
                            : `Lock ${row.school_year}`
                      }
                      aria-label={
                        isLocked
                          ? `Unlock School Year ${row.school_year}`
                          : lockedSchoolYear
                            ? `Cannot lock ${row.school_year}. Unlock ${lockedSchoolYear} first.`
                            : `Lock School Year ${row.school_year}`
                      }
                      className={`grid size-10 place-items-center rounded-xl border transition disabled:cursor-not-allowed disabled:opacity-50 ${
                        isLocked
                          ? "border-amber-300 bg-white text-amber-700 hover:bg-amber-50"
                          : "border-emerald-300 bg-white text-emerald-700 hover:bg-emerald-50"
                      }`}
                    >
                      {isBusy &&
                      (busyAction === `lock-${row.id}` ||
                        busyAction === `unlock-${row.id}`) ? (
                        <Loader2 className="size-4 animate-spin" />
                      ) : isLocked ? (
                        <UnlockKeyhole className="size-4" />
                      ) : (
                        <LockKeyhole className="size-4" />
                      )}
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setSchoolYearConfirm({
                          type: "delete",
                          row,
                        })
                      }
                      disabled={Boolean(busyAction) || isLocked}
                      title={
                        isLocked
                          ? "Unlock this School Year before deleting it."
                          : `Delete ${row.school_year}`
                      }
                      aria-label={`Delete School Year ${row.school_year}`}
                      className="grid size-10 place-items-center rounded-xl border border-rose-200 bg-white text-rose-600 transition hover:bg-rose-50 disabled:cursor-not-allowed disabled:opacity-40"
                    >
                      {busyAction === `delete-${row.id}` ? (
                        <Loader2 className="size-4 animate-spin" />
                      ) : (
                        <Trash2 className="size-4" />
                      )}
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div
        className={`border-t px-5 py-3 text-xs font-medium ${
          lockedSchoolYear
            ? "border-emerald-200 bg-emerald-50/70 text-emerald-800"
            : "border-amber-200 bg-amber-50/80 text-amber-800"
        }`}
      >
        {lockedSchoolYear ? (
          <span className="inline-flex items-center gap-2">
            <LockKeyhole className="size-4 shrink-0" />
            School Year {lockedSchoolYear} is locked. Unlock it first before
            locking a different School Year.
          </span>
        ) : (
          <span className="inline-flex items-center gap-2">
            <AlertTriangle className="size-4 shrink-0" />
            No School Year is locked. Teachers cannot create a new class until
            the Administrator locks one from the library.
          </span>
        )}
      </div>

      <Dialog
        open={Boolean(schoolYearConfirm)}
        onOpenChange={(open) => {
          if (!open && !busyAction) {
            setSchoolYearConfirm(null);
          }
        }}
      >
        <DialogContent className="overflow-hidden border-[#ead8b8] bg-[#fffdf9] p-0 sm:max-w-[500px]">
          {schoolYearConfirm ? (
            <>
              <div className="border-b border-[#efe1c8] px-6 py-5">
                <div
                  className={`mx-auto grid size-14 place-items-center rounded-full ${
                    schoolYearConfirm.type === "delete"
                      ? "bg-rose-100 text-rose-600"
                      : schoolYearConfirm.type === "unlock"
                        ? "bg-amber-100 text-amber-700"
                        : "bg-emerald-100 text-emerald-700"
                  }`}
                >
                  {schoolYearConfirm.type === "delete" ? (
                    <Trash2 className="size-7" />
                  ) : schoolYearConfirm.type === "unlock" ? (
                    <UnlockKeyhole className="size-7" />
                  ) : (
                    <LockKeyhole className="size-7" />
                  )}
                </div>

                <DialogHeader className="mt-4 space-y-2 text-center">
                  <DialogTitle className="text-xl font-bold text-[var(--admin-text-heading)]">
                    {schoolYearConfirm.type === "delete"
                      ? `Delete School Year ${schoolYearConfirm.row.school_year}?`
                      : schoolYearConfirm.type === "unlock"
                        ? `Unlock School Year ${schoolYearConfirm.row.school_year}?`
                        : `Lock School Year ${schoolYearConfirm.row.school_year}?`}
                  </DialogTitle>

                  <DialogDescription className="text-sm leading-6 text-[var(--admin-text-muted)]">
                    {schoolYearConfirm.type === "delete" ? (
                      <>
                        School Year{" "}
                        <span className="font-semibold text-[var(--admin-text-heading)]">
                          {schoolYearConfirm.row.school_year}
                        </span>{" "}
                        will be removed from the Admin School Year Library.
                        Its related School Form submissions will also be removed
                        from Available Forms and Submitted Forms. Historical
                        classes, learners, and grades will not be deleted.
                      </>
                    ) : schoolYearConfirm.type === "unlock" ? (
                      <>
                        School Year{" "}
                        <span className="font-semibold text-[var(--admin-text-heading)]">
                          {schoolYearConfirm.row.school_year}
                        </span>{" "}
                        will no longer be active. Teachers will not be able to
                        create a new class until another School Year is locked.
                      </>
                    ) : (
                      <>
                        School Year{" "}
                        <span className="font-semibold text-[var(--admin-text-heading)]">
                          {schoolYearConfirm.row.school_year}
                        </span>{" "}
                        will become the active School Year used by teachers when
                        creating new classes.
                      </>
                    )}
                  </DialogDescription>
                </DialogHeader>
              </div>

              <div className="px-6 py-4">
                <div
                  className={`rounded-xl border px-4 py-3 text-xs leading-5 ${
                    schoolYearConfirm.type === "delete"
                      ? "border-rose-200 bg-rose-50 text-rose-700"
                      : schoolYearConfirm.type === "unlock"
                        ? "border-amber-200 bg-amber-50 text-amber-800"
                        : "border-emerald-200 bg-emerald-50 text-emerald-800"
                  }`}
                >
                  <div className="flex items-start gap-2">
                    {schoolYearConfirm.type === "delete" ? (
                      <AlertTriangle className="mt-0.5 size-4 shrink-0" />
                    ) : schoolYearConfirm.type === "unlock" ? (
                      <AlertTriangle className="mt-0.5 size-4 shrink-0" />
                    ) : (
                      <LockKeyhole className="mt-0.5 size-4 shrink-0" />
                    )}

                    <span>
                      {schoolYearConfirm.type === "delete"
                        ? "This removes the School Year from the library and deletes its School Form submissions. The year must be unlocked first. Historical classes, learners, and grades are preserved."
                        : schoolYearConfirm.type === "unlock"
                          ? "No School Year will be active after this action."
                          : "Only one School Year can be locked at a time."}
                    </span>
                  </div>
                </div>
              </div>

              <DialogFooter className="gap-2 border-t border-[#efe1c8] bg-[#fffaf2] px-6 py-4 sm:justify-end">
                <Button
                  type="button"
                  variant="outline"
                  disabled={Boolean(busyAction)}
                  onClick={() => setSchoolYearConfirm(null)}
                  className="min-w-24 rounded-xl"
                >
                  Cancel
                </Button>

                <Button
                  type="button"
                  disabled={Boolean(busyAction)}
                  onClick={async () => {
                    const action = schoolYearConfirm;

                    if (action.type === "delete") {
                      await deleteSchoolYear(action.row);
                    } else if (action.type === "unlock") {
                      await unlockSchoolYear(action.row);
                    } else {
                      await lockSchoolYear(action.row);
                    }

                    setSchoolYearConfirm(null);
                  }}
                  className={`min-w-28 gap-2 rounded-xl text-white ${
                    schoolYearConfirm.type === "delete"
                      ? "bg-rose-600 hover:bg-rose-700"
                      : schoolYearConfirm.type === "unlock"
                        ? "bg-amber-600 hover:bg-amber-700"
                        : "bg-emerald-600 hover:bg-emerald-700"
                  }`}
                >
                  {busyAction ? (
                    <Loader2 className="size-4 animate-spin" />
                  ) : schoolYearConfirm.type === "delete" ? (
                    <Trash2 className="size-4" />
                  ) : schoolYearConfirm.type === "unlock" ? (
                    <UnlockKeyhole className="size-4" />
                  ) : (
                    <LockKeyhole className="size-4" />
                  )}

                  {schoolYearConfirm.type === "delete"
                    ? "Delete School Year"
                    : schoolYearConfirm.type === "unlock"
                      ? "Unlock"
                      : "Lock School Year"}
                </Button>
              </DialogFooter>
            </>
          ) : null}
        </DialogContent>
      </Dialog>
    </section>
  );
}

function DashboardSection({
  profiles,
  teacherProfiles,
  classes,
  students,
  grades,
  roleByUserId,
  onlineAdminCount,
  maxOnlineAdmins,
  onChangeSection,
}: {
  profiles: ProfileRow[];
  teacherProfiles: ProfileRow[];
  classes: ClassRow[];
  students: StudentRow[];
  grades: GradeRow[];
  roleByUserId: Map<string, string>;
  onlineAdminCount: number;
  maxOnlineAdmins: number;
  onChangeSection: (section: AdminSection) => void;
}) {
  const gradeLevels = new Set(
    classes.map((item) => item.grade_level).filter(Boolean),
  ).size;
  const sections = new Set(
    classes
      .map((item) => `${item.grade_level}-${item.section}`)
      .filter(Boolean),
  ).size;
  const recentProfiles = profiles.slice(0, 5);

  return (
    <div className="space-y-5">
      <PageHeading
        title="Welcome back, Admin!"
        description="Manage teachers, learners, classes, records, and school forms."
      />

      <SchoolYearLibraryCard />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
        <StatCard
          title="Teachers / Users"
          value={teacherProfiles.length}
          helper="Total"
          icon={Users}
          tone="orange"
        />
        <StatCard
          title="Classes"
          value={classes.length}
          helper="Total"
          icon={BookOpen}
          tone="green"
        />
        <StatCard
          title="Learners / Students"
          value={students.length}
          helper="Total"
          icon={GraduationCap}
          tone="blue"
        />
        <StatCard
          title="Grade Record Count"
          value={grades.length}
          helper="Total"
          icon={FileText}
          tone="violet"
        />
        <StatCard
          title="Admins Online"
          value={`${onlineAdminCount} / ${maxOnlineAdmins}`}
          helper="Currently active"
          icon={ShieldCheck}
          tone="green"
        />
      </div>

      <div>
        <Panel>
          <PanelHeader
            title="Recent Registrations"
            action={
              <button
                type="button"
                onClick={() => onChangeSection("users")}
                className="text-[10px] font-semibold text-[var(--admin-primary-strong)] hover:underline"
              >
                View all
              </button>
            }
          />
          <div className="divide-y divide-[var(--admin-color-f0e7df)]">
            {recentProfiles.map((profile, index) => (
              <div
                key={profile.id}
                className="flex items-center gap-3 py-3 first:pt-0 last:pb-0"
              >
                {profile.avatar_url ? (
                  <img
                    src={profile.avatar_url}
                    alt={profile.full_name || "User profile"}
                    className="size-9 shrink-0 rounded-full border border-[var(--admin-border)] object-cover shadow-sm"
                  />
                ) : (
                  <div
                    className={`grid size-9 shrink-0 place-items-center rounded-full ${registrationTone(index)}`}
                  >
                    <UserRound className="size-4" />
                  </div>
                )}

                <div className="min-w-0 flex-1">
                  <div className="truncate text-xs font-semibold text-[var(--admin-color-362722)]">
                    {profile.full_name || "No name"}
                  </div>
                  <div className="text-[10px] text-[var(--admin-color-8b776e)]">
                    {adminUserRoleLabel(profile, roleByUserId)}
                  </div>
                </div>

                <div className="shrink-0 text-[9px] text-[var(--admin-color-9b887f)]">
                  {formatDate(profile.created_at)}
                </div>
              </div>
            ))}
            {recentProfiles.length === 0 && (
              <EmptyState text="No user registrations found." />
            )}
          </div>
        </Panel>
      </div>

      <Panel>
        <PanelHeader
          title="Forms Overview"
          action={
            <button
              type="button"
              onClick={() => onChangeSection("forms")}
              className="text-[10px] font-semibold text-[var(--admin-primary-strong)] hover:underline"
            >
              View all
            </button>
          }
        />
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <MiniMetric
            icon={Files}
            label="Form Types"
            value={FORM_TYPES.length}
          />
          <MiniMetric
            icon={GraduationCap}
            label="Grade Levels"
            value={gradeLevels}
          />
          <MiniMetric icon={FolderOpen} label="Sections" value={sections} />
          <MiniMetric
            icon={FileStack}
            label="Class Records"
            value={classes.length}
          />
        </div>
      </Panel>
    </div>
  );
}

function UsersSection({
  profiles,
  roleByUserId,
  search,
  onSearch,
  page,
  onPage,
  currentUserId,
  onlineAdminCount,
  maxOnlineAdmins,
  onlineAdminIds,
  busyId,
  onAdd,
  onEdit,
  onDelete,
}: {
  profiles: ProfileRow[];
  roleByUserId: Map<string, string>;
  search: string;
  onSearch: (value: string) => void;
  page: number;
  onPage: (value: number) => void;
  currentUserId: string | null;
  onlineAdminCount: number;
  maxOnlineAdmins: number;
  onlineAdminIds: Set<string>;
  busyId: string | null;
  onAdd: () => void;
  onEdit: (profile: ProfileRow) => void;
  onDelete: (profile: ProfileRow) => void;
}) {
  const filtered = useMemo(() => {
    const q = normalize(search);
    if (!q) return profiles;
    return profiles.filter((profile) =>
      [
        profile.full_name,
        profile.email,
        profile.school_name,
        roleByUserId.get(profile.id),
        profile.teacher_type,
        adminUserRoleLabel(profile, roleByUserId),
      ].some((value) => normalize(value).includes(q)),
    );
  }, [profiles, roleByUserId, search]);

  const teacherCount = profiles.filter(
    (item) => roleByUserId.get(item.id) !== "admin",
  ).length;
  const adminCount = profiles.filter(
    (item) => roleByUserId.get(item.id) === "admin",
  ).length;
  const thirtyDaysAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
  const recentCount = profiles.filter(
    (item) => new Date(item.created_at).getTime() >= thirtyDaysAgo,
  ).length;
  const paged = paginate(filtered, page, PAGE_SIZE);

  return (
    <div className="space-y-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
        <PageHeading
          title="All Teachers / Users"
          description="Manage all teacher accounts and system users."
        />

        <button
          type="button"
          onClick={onAdd}
          className="inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[var(--admin-primary)] px-4 text-[11px] font-semibold text-[var(--admin-white)] shadow-[var(--admin-shadow-active)] hover:bg-[var(--admin-primary-hover)]"
        >
          <Plus className="size-4" />
          Add Teacher / User
        </button>
      </div>
      <SearchBar
        value={search}
        onChange={(value) => {
          onSearch(value);
          onPage(1);
        }}
        placeholder="Search teacher name, email, school, or role..."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
        <StatCard
          title="Total Users"
          value={profiles.length}
          helper="Registered"
          icon={UsersRound}
          tone="orange"
          compact
        />
        <StatCard
          title="Teachers"
          value={teacherCount}
          helper="Accounts"
          icon={UserCheck}
          tone="green"
          compact
        />
        <StatCard
          title="Administrators"
          value={adminCount}
          helper="Accounts"
          icon={ShieldCheck}
          tone="violet"
          compact
        />
        <StatCard
          title="New This Month"
          value={recentCount}
          helper="Registrations"
          icon={CalendarDays}
          tone="pink"
          compact
        />
        <StatCard
          title="Admins Online"
          value={`${onlineAdminCount} / ${maxOnlineAdmins}`}
          helper="Currently active"
          icon={ShieldCheck}
          tone="green"
          compact
        />
      </div>

      <Panel>
        <PanelHeader title="Teachers / Users List" />
        <DataTable minWidth="900px">
          <thead>
            <tr>
              <Th>#</Th>
              <Th>Name</Th>
              <Th>Email</Th>
              <Th>Role</Th>
              <Th>Registered</Th>
              <Th>Status</Th>
              <Th align="right">Action</Th>
            </tr>
          </thead>
          <tbody>
            {paged.items.map((profile, index) => {
              const role = adminUserRoleLabel(profile, roleByUserId);
              const isAdministrator =
                normalize(roleByUserId.get(profile.id)) === "admin";
              const isOnlineAdmin =
                isAdministrator && onlineAdminIds.has(profile.id);

              return (
                <tr
                  key={profile.id}
                  className="border-b border-[var(--admin-row-divider)] last:border-0 hover:bg-[var(--admin-surface-soft)]"
                >
                  <Td>{(paged.currentPage - 1) * PAGE_SIZE + index + 1}</Td>
                  <Td strong>{profile.full_name || "No name"}</Td>
                  <Td>{profile.email || "-"}</Td>
                  <Td>
                    <RoleBadge role={role} />
                  </Td>
                  <Td>{formatDate(profile.created_at)}</Td>
                  <Td>
                    {isAdministrator ? (
                      isOnlineAdmin ? (
                        <StatusBadge>Online</StatusBadge>
                      ) : (
                        <span className="inline-flex rounded-full bg-slate-100 px-2 py-1 text-[8px] font-semibold text-slate-600">
                          Offline
                        </span>
                      )
                    ) : (
                      <StatusBadge>Registered</StatusBadge>
                    )}
                  </Td>
                  <Td align="right">
                    <RowActions
                      busy={busyId === profile.id}
                      disableDelete={profile.id === currentUserId}
                      onEdit={() => onEdit(profile)}
                      onDelete={() => onDelete(profile)}
                    />
                  </Td>
                </tr>
              );
            })}
            {paged.items.length === 0 && (
              <TableEmpty
                colSpan={7}
                text="No matching teacher or user found."
              />
            )}
          </tbody>
        </DataTable>
        <PaginationFooter
          total={filtered.length}
          pageSize={PAGE_SIZE}
          page={paged.currentPage}
          totalPages={paged.totalPages}
          onPage={onPage}
        />
      </Panel>
    </div>
  );
}

function StudentsSection({
  students,
  teacherById,
  classById,
  search,
  onSearch,
  page,
  onPage,
  busyId,
  onAdd,
  onEdit,
  onDelete,
}: {
  students: StudentRow[];
  teacherById: Map<string, ProfileRow>;
  classById: Map<string, ClassRow>;
  search: string;
  onSearch: (value: string) => void;
  page: number;
  onPage: (value: number) => void;
  busyId: string | null;
  onAdd: () => void;
  onEdit: (student: StudentRow) => void;
  onDelete: (student: StudentRow) => void;
}) {
  const [selectedGradeLevel, setSelectedGradeLevel] = useState("all");

  const filtered = useMemo(() => {
    const q = normalize(search);

    return students
      .filter((student) => {
        const cls = classById.get(student.class_id);
        const gradeNumber = cls?.grade_level?.match(/\d+/)?.[0] ?? "";
        const matchesGrade =
          selectedGradeLevel === "all" || gradeNumber === selectedGradeLevel;

        if (!matchesGrade) return false;
        if (!q) return true;

        const teacher = teacherById.get(student.teacher_id);
        return [
          student.lrn,
          studentName(student),
          student.guardian,
          student.mother_name,
          student.father_name,
          student.contact_number,
          student.enrollment_status,
          student.previous_school,
          student.destination_school,
          student.transfer_reason,
          cls?.subject,
          cls?.grade_level,
          cls?.section,
          teacher?.full_name,
        ].some((value) => normalize(value).includes(q));
      })
      .sort((firstStudent, secondStudent) => {
        const firstClass = classById.get(firstStudent.class_id);
        const secondClass = classById.get(secondStudent.class_id);
        const firstGrade = Number(
          firstClass?.grade_level?.match(/\d+/)?.[0] ?? 999,
        );
        const secondGrade = Number(
          secondClass?.grade_level?.match(/\d+/)?.[0] ?? 999,
        );

        if (firstGrade !== secondGrade) return firstGrade - secondGrade;

        const sectionComparison = (firstClass?.section ?? "").localeCompare(
          secondClass?.section ?? "",
        );
        if (sectionComparison !== 0) return sectionComparison;

        return studentName(firstStudent).localeCompare(
          studentName(secondStudent),
        );
      });
  }, [students, search, selectedGradeLevel, classById, teacherById]);

  const male = students.filter(
    (student) => normalize(student.sex) === "male",
  ).length;
  const female = students.filter(
    (student) => normalize(student.sex) === "female",
  ).length;
  const withGuardian = students.filter(
    (student) => student.guardian || student.mother_name || student.father_name,
  ).length;
  const gradeEnrollment = useMemo(() => {
    const gradeLevels = [
      "Grade 7",
      "Grade 8",
      "Grade 9",
      "Grade 10",
      "Grade 11",
      "Grade 12",
    ];
    const counts = new Map(gradeLevels.map((grade) => [grade, 0]));

    students.forEach((student) => {
      const cls = classById.get(student.class_id);
      const gradeNumber = cls?.grade_level?.match(/\d+/)?.[0];

      if (!gradeNumber) return;

      const gradeLabel = `Grade ${gradeNumber}`;

      if (counts.has(gradeLabel)) {
        counts.set(gradeLabel, (counts.get(gradeLabel) ?? 0) + 1);
      }
    });

    return gradeLevels.map((grade) => ({
      grade,
      enrolled: counts.get(grade) ?? 0,
    }));
  }, [students, classById]);
  const paged = paginate(filtered, page, PAGE_SIZE);

  return (
    <div className="space-y-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
        <PageHeading
          title="All Learners / Students"
          description="Create and manage complete SF1 learner records."
        />
        <Button
          type="button"
          onClick={onAdd}
          className="gap-2 bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
        >
          <Plus className="size-4" />
          Add Student
        </Button>
      </div>
      <div className="grid gap-3 sm:grid-cols-[minmax(0,1fr)_220px]">
        <SearchBar
          value={search}
          onChange={(value) => {
            onSearch(value);
            onPage(1);
          }}
          placeholder="Search student name, LRN, class, teacher, or guardian..."
        />

        <div className="relative">
          <GraduationCap className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          <select
            value={selectedGradeLevel}
            onChange={(event) => {
              setSelectedGradeLevel(event.target.value);
              onPage(1);
            }}
            aria-label="Filter learners by grade level"
            className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-10 text-sm font-medium text-[var(--admin-color-2d211d)] outline-none transition focus:border-[var(--admin-primary)] focus:ring-2 focus:ring-[var(--admin-primary)]/15"
          >
            <option value="all">All Grade Levels</option>
            <option value="7">Grade 7</option>
            <option value="8">Grade 8</option>
            <option value="9">Grade 9</option>
            <option value="10">Grade 10</option>
            <option value="11">Grade 11</option>
            <option value="12">Grade 12</option>
          </select>
          <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
        </div>
      </div>

      <div className="grid gap-5 xl:grid-cols-[minmax(0,2fr)_minmax(340px,1fr)]">
        <div className="grid gap-3 sm:grid-cols-2">
          <StatCard
            title="Total Students"
            value={students.length}
            helper="Learners"
            icon={GraduationCap}
            tone="orange"
            compact
          />
          <StatCard
            title="Male"
            value={male}
            helper="Learners"
            icon={Mars}
            tone="blue"
            compact
          />
          <StatCard
            title="Female"
            value={female}
            helper="Learners"
            icon={Venus}
            tone="pink"
            compact
          />
          <StatCard
            title="With Guardians"
            value={withGuardian}
            helper="Records"
            icon={HeartHandshake}
            tone="violet"
            compact
          />
        </div>

        <GradeEnrollmentCard gradeEnrollment={gradeEnrollment} />
      </div>

      <Panel>
        <PanelHeader
          title={
            selectedGradeLevel === "all"
              ? "Students List"
              : `Grade ${selectedGradeLevel} Students`
          }
        />
        <DataTable minWidth="1240px">
          <thead>
            <tr>
              <Th>#</Th>
              <Th>LRN</Th>
              <Th>Student Name</Th>
              <Th>Class</Th>
              <Th>Transfer Status</Th>
              <Th>Teacher</Th>
              <Th>Guardian</Th>
              <Th>Contact</Th>
              <Th align="right">Action</Th>
            </tr>
          </thead>
          <tbody>
            {paged.items.map((student, index) => {
              const cls = classById.get(student.class_id);
              const teacher = teacherById.get(student.teacher_id);
              return (
                <tr
                  key={student.id}
                  className="border-b border-[var(--admin-row-divider)] last:border-0 hover:bg-[var(--admin-surface-soft)]"
                >
                  <Td>{(paged.currentPage - 1) * PAGE_SIZE + index + 1}</Td>
                  <Td>{student.lrn || "-"}</Td>
                  <Td strong>{studentName(student)}</Td>
                  <Td>
                    {cls
                      ? `${cls.grade_level || "-"} - ${cls.section || "-"}`
                      : "-"}
                  </Td>
                  <Td>
                    <StudentTransferStatusBadge student={student} />
                  </Td>
                  <Td>{teacher?.full_name || teacher?.email || "-"}</Td>
                  <Td>
                    {student.guardian ||
                      student.mother_name ||
                      student.father_name ||
                      "-"}
                  </Td>
                  <Td>{student.contact_number || "-"}</Td>
                  <Td align="right">
                    <RowActions
                      busy={busyId === student.id}
                      onEdit={() => onEdit(student)}
                      onDelete={() => onDelete(student)}
                    />
                  </Td>
                </tr>
              );
            })}
            {paged.items.length === 0 && (
              <TableEmpty colSpan={9} text="No matching learner found." />
            )}
          </tbody>
        </DataTable>
        <PaginationFooter
          total={filtered.length}
          pageSize={PAGE_SIZE}
          page={paged.currentPage}
          totalPages={paged.totalPages}
          onPage={onPage}
        />
      </Panel>
    </div>
  );
}

function GradeEnrollmentCard({
  gradeEnrollment,
}: {
  gradeEnrollment: Array<{ grade: string; enrolled: number }>;
}) {
  return (
    <div className="overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-sm">
      <div className="grid grid-cols-[1fr_auto] border-b border-[var(--admin-row-divider)] bg-[var(--admin-surface-soft)] px-4 py-3 text-[11px] font-bold uppercase tracking-wide text-[var(--admin-text-muted)]">
        <span>Grade Level</span>
        <span>Enrolled</span>
      </div>

      <div className="divide-y divide-[var(--admin-row-divider)]">
        {gradeEnrollment.map((item) => (
          <div
            key={item.grade}
            className="grid grid-cols-[1fr_auto] items-center px-4 py-3 text-sm"
          >
            <span className="font-semibold text-[var(--admin-text-heading)]">
              {item.grade}
            </span>
            <span className="font-bold text-[var(--admin-text-value)]">
              {item.enrolled}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

function CredentialRequestsAdminSection({
  requests,
  students,
  studentById,
  teacherById,
  classById,
  busyId,
  onReview,
}: {
  requests: StudentCredentialRequestRow[];
  students: StudentRow[];
  studentById: Map<string, StudentRow>;
  teacherById: Map<string, ProfileRow>;
  classById: Map<string, ClassRow>;
  busyId: string | null;
  onReview: (
    request: StudentCredentialRequestRow,
    status: "approved" | "rejected",
    adminRemarks: string,
    resolvedStudent: StudentRow | null,
  ) => Promise<void>;
}) {
  const [searchValue, setSearchValue] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [adviserFilter, setAdviserFilter] = useState("all");
  const [selectedRequest, setSelectedRequest] =
    useState<StudentCredentialRequestRow | null>(null);
  const [informationTab, setInformationTab] = useState<
    "personal" | "address" | "guardian" | "enrollment"
  >("personal");
  const [adminRemarks, setAdminRemarks] = useState("");
  const [adminStudentSearch, setAdminStudentSearch] = useState("");
  const [resolvedStudentId, setResolvedStudentId] = useState("");
  const [adminStudentListOpen, setAdminStudentListOpen] = useState(false);

  const normalizedStatus = (request: StudentCredentialRequestRow) => {
    const value = normalize(request.status);
    if (value === "pending") return "pending_review";
    if (value === "completed" || value === "released") return "approved";
    return value || "pending_review";
  };

  const requestCode = (request: StudentCredentialRequestRow) => {
    const year = new Date(request.request_date || request.created_at)
      .getFullYear()
      .toString();
    const suffix = request.id.replaceAll("-", "").slice(0, 6).toUpperCase();
    return `SF1-${year}-${suffix}`;
  };

  const statusLabel = (status: string) =>
    status === "pending_review"
      ? "Pending Review"
      : status === "approved"
        ? "Approved"
        : status === "rejected"
          ? "Rejected"
          : status.replaceAll("_", " ");

  const statusClass = (status: string) =>
    status === "approved"
      ? "bg-emerald-100 text-emerald-700"
      : status === "rejected"
        ? "bg-red-100 text-red-700"
        : "bg-amber-100 text-amber-700";

  const adviserOptions = useMemo(
    () =>
      Array.from(
        new Map(
          requests.map((request) => {
            const profile = teacherById.get(request.requester_id);
            return [
              request.requester_id,
              profile?.full_name || profile?.email || "Unknown adviser",
            ];
          }),
        ),
      ),
    [requests, teacherById],
  );

  const filteredRequests = useMemo(() => {
    const query = normalize(searchValue);
    return requests.filter((request) => {
      const student = request.student_id
        ? studentById.get(request.student_id)
        : null;
      const searchable = normalize(
        [
          requestCode(request),
          student ? studentName(student) : "",
          request.requested_student_name,
          student?.lrn,
        ].join(" "),
      );
      return (
        (!query || searchable.includes(query)) &&
        (statusFilter === "all" ||
          normalizedStatus(request) === statusFilter) &&
        (adviserFilter === "all" || request.requester_id === adviserFilter)
      );
    });
  }, [adviserFilter, requests, searchValue, statusFilter, studentById]);

  const pendingCount = requests.filter(
    (request) => normalizedStatus(request) === "pending_review",
  ).length;
  const approvedCount = requests.filter(
    (request) => normalizedStatus(request) === "approved",
  ).length;
  const rejectedCount = requests.filter(
    (request) => normalizedStatus(request) === "rejected",
  ).length;

  const openRequest = (request: StudentCredentialRequestRow) => {
    setSelectedRequest(request);
    setAdminRemarks(request.admin_remarks || request.release_notes || "");
    setResolvedStudentId(request.student_id || "");
    const linkedStudent = request.student_id
      ? studentById.get(request.student_id)
      : null;
    setAdminStudentSearch(
      linkedStudent
        ? studentName(linkedStudent)
        : request.requested_student_name || "",
    );
    setAdminStudentListOpen(false);
    setInformationTab("personal");
  };

  const closeRequest = () => {
    setSelectedRequest(null);
    setAdminRemarks("");
    setResolvedStudentId("");
    setAdminStudentSearch("");
    setAdminStudentListOpen(false);
  };

  const submitReview = async (status: "approved" | "rejected") => {
    if (!selectedRequest) return;
    if (status !== "rejected" && !selectedStudent) {
      window.alert("Search and select the correct student before approval.");
      return;
    }
    await onReview(selectedRequest, status, adminRemarks, selectedStudent);
    closeRequest();
  };

  const adminStudentMatches = useMemo(() => {
    const query = normalize(adminStudentSearch);
    if (resolvedStudentId) return [];
    if (!query) return students;

    return students.filter((student) => {
      const name = normalize(studentName(student));
      const firstNameFirst = normalize(
        [student.first_name, student.middle_name, student.last_name]
          .filter(Boolean)
          .join(" "),
      );
      return (
        name.includes(query) ||
        firstNameFirst.includes(query) ||
        normalize(student.lrn).includes(query)
      );
    });
  }, [adminStudentSearch, resolvedStudentId, students]);

  const selectedStudent = resolvedStudentId
    ? studentById.get(resolvedStudentId) || null
    : null;
  const selectedClass = selectedStudent
    ? classById.get(selectedStudent.class_id) || null
    : selectedRequest?.requesting_class_id
      ? classById.get(selectedRequest.requesting_class_id) || null
      : null;
  const selectedPreviousClass = selectedRequest?.previous_class_id
    ? classById.get(selectedRequest.previous_class_id)
    : null;
  const selectedRequester = selectedRequest
    ? teacherById.get(selectedRequest.requester_id)
    : null;
  const selectedStatus = selectedRequest
    ? normalizedStatus(selectedRequest)
    : "pending_review";

  const informationValue = (value: unknown) => {
    const text = String(value ?? "").trim();
    return text || "—";
  };

  const learnerAge = (() => {
    if (!selectedStudent?.birthdate) return "—";
    const birth = new Date(
      `${selectedStudent.birthdate.slice(0, 10)}T00:00:00Z`,
    );
    if (Number.isNaN(birth.getTime())) return "—";
    const yearMatch = selectedClass?.school_year?.match(/\d{4}/)?.[0];
    const schoolYear = yearMatch ? Number(yearMatch) : new Date().getFullYear();
    const reference = new Date(Date.UTC(schoolYear, 5, 1));
    reference.setUTCDate(
      reference.getUTCDate() + ((5 - reference.getUTCDay() + 7) % 7),
    );
    let age = reference.getUTCFullYear() - birth.getUTCFullYear();
    const birthdayPassed =
      reference.getUTCMonth() > birth.getUTCMonth() ||
      (reference.getUTCMonth() === birth.getUTCMonth() &&
        reference.getUTCDate() >= birth.getUTCDate());
    if (!birthdayPassed) age -= 1;
    return String(age);
  })();

  return (
    <div className="space-y-5">
      <PageHeading
        title="Student Credential Requests"
        description="Review and approve the School Form 1 information of the requested learner."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatCard
          title="Total Requests"
          value={requests.length}
          helper="All SF1 requests"
          icon={FileStack}
          tone="orange"
          compact
        />
        <StatCard
          title="Pending Review"
          value={pendingCount}
          helper="Needs admin review"
          icon={ClipboardCheck}
          tone="pink"
          compact
        />
        <StatCard
          title="Approved"
          value={approvedCount}
          helper="Ready for class assignment"
          icon={CircleCheck}
          tone="green"
          compact
        />
        <StatCard
          title="Rejected"
          value={rejectedCount}
          helper="Returned requests"
          icon={X}
          tone="pink"
          compact
        />
      </div>

      <Panel>
        <div className="grid gap-3 border-b border-[var(--admin-row-divider)] p-4 lg:grid-cols-[minmax(260px,1fr)_190px_220px]">
          <div className="relative">
            <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
            <input
              value={searchValue}
              onChange={(event) => setSearchValue(event.target.value)}
              placeholder="Search learner name, LRN, or request ID..."
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-white pl-9 pr-3 text-sm outline-none focus:border-[var(--admin-primary)]"
            />
          </div>
          <select
            value={statusFilter}
            onChange={(event) => setStatusFilter(event.target.value)}
            className="h-10 rounded-xl border border-[var(--admin-border)] bg-white px-3 text-sm outline-none focus:border-[var(--admin-primary)]"
          >
            <option value="all">All statuses</option>
            <option value="pending_review">Pending Review</option>
            <option value="approved">Approved</option>
            <option value="rejected">Rejected</option>
          </select>
          <select
            value={adviserFilter}
            onChange={(event) => setAdviserFilter(event.target.value)}
            className="h-10 rounded-xl border border-[var(--admin-border)] bg-white px-3 text-sm outline-none focus:border-[var(--admin-primary)]"
          >
            <option value="all">All requesting advisers</option>
            {adviserOptions.map(([id, label]) => (
              <option key={id} value={id}>
                {label}
              </option>
            ))}
          </select>
        </div>

        <DataTable minWidth="1100px">
          <thead>
            <tr>
              <Th>Request ID</Th>
              <Th>Learner</Th>
              <Th>LRN</Th>
              <Th>Requesting Adviser</Th>
              <Th>Current Class</Th>
              <Th>Request Date</Th>
              <Th>Status</Th>
              <Th align="right">Action</Th>
            </tr>
          </thead>
          <tbody>
            {filteredRequests.map((request) => {
              const student = request.student_id
                ? studentById.get(request.student_id)
                : null;
              const requester = teacherById.get(request.requester_id);
              const currentClass = request.requesting_class_id
                ? classById.get(request.requesting_class_id)
                : null;
              const status = normalizedStatus(request);

              return (
                <tr
                  key={request.id}
                  className="border-b border-[var(--admin-row-divider)] last:border-0 hover:bg-[var(--admin-surface-soft)]"
                >
                  <Td strong>{requestCode(request)}</Td>
                  <Td strong>
                    {student
                      ? studentName(student)
                      : request.requested_student_name || "Name not entered"}
                  </Td>
                  <Td>{student?.lrn || "—"}</Td>
                  <Td>{requester?.full_name || requester?.email || "-"}</Td>
                  <Td>
                    {currentClass
                      ? `${currentClass.grade_level || "-"} - ${currentClass.section || "-"}`
                      : "-"}
                  </Td>
                  <Td>
                    {formatDate(request.request_date || request.created_at)}
                  </Td>
                  <Td>
                    <span
                      className={`inline-flex rounded-full px-2.5 py-1 text-[9px] font-semibold ${statusClass(status)}`}
                    >
                      {statusLabel(status)}
                    </span>
                  </Td>
                  <Td align="right">
                    <button
                      type="button"
                      onClick={() => openRequest(request)}
                      className={`inline-flex items-center gap-1 rounded-lg px-3 py-2 text-[10px] font-semibold ${
                        status === "pending_review"
                          ? "bg-[var(--admin-primary)] text-white hover:bg-[var(--admin-primary-hover)]"
                          : "border border-[var(--admin-primary)] text-[var(--admin-primary)] hover:bg-[var(--admin-primary)]/5"
                      }`}
                    >
                      {status === "pending_review" ? "Review" : "View"}
                    </button>
                  </Td>
                </tr>
              );
            })}
            {filteredRequests.length === 0 && (
              <TableEmpty
                colSpan={8}
                text="No credential request matches the selected filters."
              />
            )}
          </tbody>
        </DataTable>
      </Panel>

      {selectedRequest && (
        <div className="fixed inset-0 z-[90] bg-black/25" role="presentation">
          <button
            type="button"
            className="absolute inset-0 cursor-default"
            onClick={closeRequest}
            aria-label="Close credential review"
          />
          <aside className="absolute inset-y-0 right-0 flex w-full max-w-[650px] flex-col overflow-hidden border-l border-[var(--admin-border)] bg-[var(--admin-page-bg)] shadow-2xl">
            <div className="flex items-start justify-between border-b border-[var(--admin-row-divider)] bg-white px-5 py-4">
              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-[var(--admin-primary)]">
                  Review Credential Request
                </p>
                <h2 className="mt-1 text-lg font-bold text-[var(--admin-text-heading)]">
                  {requestCode(selectedRequest)}
                </h2>
              </div>
              <button
                type="button"
                onClick={closeRequest}
                className="rounded-lg p-2 text-[var(--admin-text-muted)] hover:bg-[var(--admin-surface-soft)]"
                aria-label="Close review panel"
              >
                <X className="size-5" />
              </button>
            </div>

            <div className="flex-1 space-y-4 overflow-y-auto p-5">
              <section>
                <div className="mb-2 flex items-center gap-2 text-sm font-bold text-[var(--admin-text-heading)]">
                  <span className="grid size-5 place-items-center rounded-full border border-[var(--admin-primary)] text-[10px] text-[var(--admin-primary)]">
                    1
                  </span>
                  Learner Verification
                </div>
                <div className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
                  <p className="mb-2 text-xs text-[var(--admin-text-muted)]">
                    Requested name or LRN:{" "}
                    <span className="font-semibold text-[var(--admin-text-heading)]">
                      {selectedRequest.requested_student_name ||
                        (selectedStudent ? studentName(selectedStudent) : "—")}
                    </span>
                  </p>

                  <div className="relative mb-3">
                    <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
                    <input
                      value={adminStudentSearch}
                      onFocus={() => setAdminStudentListOpen(true)}
                      onChange={(event) => {
                        setAdminStudentSearch(event.target.value);
                        setResolvedStudentId("");
                        setAdminStudentListOpen(true);
                      }}
                      placeholder="Admin: search student name or LRN..."
                      disabled={
                        selectedStatus === "approved" ||
                        selectedStatus === "rejected"
                      }
                      className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-white pl-9 pr-3 text-sm outline-none focus:border-[var(--admin-primary)] disabled:bg-[var(--admin-surface-soft)]"
                    />

                    {adminStudentListOpen && !resolvedStudentId && (
                      <div className="absolute left-0 right-0 top-11 z-20 max-h-72 overflow-y-auto rounded-xl border border-[var(--admin-border)] bg-white shadow-xl">
                        {adminStudentMatches.length > 0 ? (
                          adminStudentMatches.map((student) => {
                            const studentClass = classById.get(
                              student.class_id,
                            );
                            return (
                              <button
                                key={student.id}
                                type="button"
                                onClick={() => {
                                  setResolvedStudentId(student.id);
                                  setAdminStudentSearch(studentName(student));
                                  setAdminStudentListOpen(false);
                                }}
                                className="flex w-full items-center justify-between gap-3 border-b px-3 py-2 text-left text-xs last:border-0 hover:bg-[var(--admin-surface-soft)]"
                              >
                                <span>
                                  <span className="block font-semibold">
                                    {studentName(student)}
                                  </span>
                                  <span className="text-[var(--admin-text-muted)]">
                                    LRN: {student.lrn || "Not entered"}
                                  </span>
                                </span>
                                <span className="shrink-0 text-[var(--admin-text-muted)]">
                                  {studentClass
                                    ? `${studentClass.grade_level || "—"} - ${studentClass.section || "—"}`
                                    : "—"}
                                </span>
                              </button>
                            );
                          })
                        ) : (
                          <p className="px-3 py-4 text-center text-xs text-[var(--admin-text-muted)]">
                            No student matched. Clear the search to view all
                            students.
                          </p>
                        )}
                      </div>
                    )}
                  </div>

                  {selectedStatus !== "approved" &&
                    selectedStatus !== "rejected" && (
                      <div className="mb-3 flex justify-end">
                        <button
                          type="button"
                          onClick={() => {
                            setAdminStudentSearch("");
                            setResolvedStudentId("");
                            setAdminStudentListOpen(true);
                          }}
                          className="text-xs font-semibold text-[var(--admin-primary)] hover:underline"
                        >
                          View all students
                        </button>
                      </div>
                    )}

                  <div className="flex items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="grid size-10 place-items-center rounded-full bg-[var(--admin-primary)]/10 text-[var(--admin-primary)]">
                        <UserRound className="size-5" />
                      </div>
                      <div>
                        <p className="font-bold text-[var(--admin-text-heading)]">
                          {selectedStudent
                            ? studentName(selectedStudent)
                            : "Select the correct learner"}
                        </p>
                        <p className="text-xs text-[var(--admin-text-muted)]">
                          LRN: {selectedStudent?.lrn || "Not entered"}
                        </p>
                      </div>
                    </div>
                    <span
                      className={`rounded-full px-2.5 py-1 text-xs font-semibold ${selectedStudent ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"}`}
                    >
                      {selectedStudent ? "Verified" : "Needs matching"}
                    </span>
                  </div>
                  <div className="mt-3 grid grid-cols-[90px_1fr] gap-2 border-t pt-3 text-xs">
                    <span className="text-[var(--admin-text-muted)]">
                      Class
                    </span>
                    <span className="font-semibold">
                      {selectedClass
                        ? `${selectedClass.grade_level || "—"} - ${selectedClass.section || "—"}`
                        : "—"}
                    </span>
                  </div>
                </div>
              </section>

              <section>
                <div className="mb-2 flex items-center gap-2 text-sm font-bold text-[var(--admin-text-heading)]">
                  <span className="grid size-5 place-items-center rounded-full border border-[var(--admin-primary)] text-[10px] text-[var(--admin-primary)]">
                    2
                  </span>
                  Complete Student Information
                </div>
                <div className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
                  <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                    {(
                      [
                        ["personal", "Personal Information"],
                        ["address", "Address"],
                        ["guardian", "Parent/Guardian"],
                        ["enrollment", "Enrollment Details"],
                      ] as const
                    ).map(([value, label]) => (
                      <button
                        key={value}
                        type="button"
                        onClick={() => setInformationTab(value)}
                        className={`rounded-lg border px-2 py-2 text-[10px] font-semibold ${
                          informationTab === value
                            ? "border-[var(--admin-primary)] bg-[var(--admin-primary)]/5 text-[var(--admin-primary)]"
                            : "border-[var(--admin-border)] text-[var(--admin-text-muted)]"
                        }`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>

                  <div className="mt-4 grid gap-x-4 gap-y-2 text-xs sm:grid-cols-2">
                    {informationTab === "personal" && (
                      <>
                        <InfoLine
                          label="Full Name"
                          value={
                            selectedStudent ? studentName(selectedStudent) : "—"
                          }
                        />
                        <InfoLine
                          label="Sex"
                          value={informationValue(selectedStudent?.sex)}
                        />
                        <InfoLine
                          label="Birthdate"
                          value={formatDate(selectedStudent?.birthdate)}
                        />
                        <InfoLine
                          label="Age as of first Friday of June"
                          value={learnerAge}
                        />
                        <InfoLine
                          label="Mother Tongue"
                          value={informationValue(
                            selectedStudent?.mother_tongue,
                          )}
                        />
                        <InfoLine
                          label="IP / Ethnic Group"
                          value={informationValue(
                            selectedStudent?.ip_ethnic_group ||
                              selectedStudent?.ip_ethnic,
                          )}
                        />
                        <InfoLine
                          label="Religion"
                          value={informationValue(selectedStudent?.religion)}
                        />
                      </>
                    )}
                    {informationTab === "address" && (
                      <>
                        <InfoLine
                          label="House # / Street / Sitio"
                          value={informationValue(
                            selectedStudent?.house_street,
                          )}
                        />
                        <InfoLine
                          label="Barangay"
                          value={informationValue(selectedStudent?.barangay)}
                        />
                        <InfoLine
                          label="Municipality / City"
                          value={informationValue(
                            selectedStudent?.municipality_city ||
                              selectedStudent?.municipality,
                          )}
                        />
                        <InfoLine
                          label="Province"
                          value={informationValue(selectedStudent?.province)}
                        />
                        <InfoLine
                          label="Complete Address"
                          value={informationValue(selectedStudent?.address)}
                        />
                      </>
                    )}
                    {informationTab === "guardian" && (
                      <>
                        <InfoLine
                          label="Father's Name"
                          value={informationValue(selectedStudent?.father_name)}
                        />
                        <InfoLine
                          label="Mother's Maiden Name"
                          value={informationValue(selectedStudent?.mother_name)}
                        />
                        <InfoLine
                          label="Parent / Guardian"
                          value={informationValue(selectedStudent?.guardian)}
                        />
                        <InfoLine
                          label="Relationship"
                          value={informationValue(
                            selectedStudent?.guardian_relationship,
                          )}
                        />
                        <InfoLine
                          label="Contact Number"
                          value={informationValue(
                            selectedStudent?.contact_number,
                          )}
                        />
                      </>
                    )}
                    {informationTab === "enrollment" && (
                      <>
                        <InfoLine
                          label="LRN"
                          value={informationValue(selectedStudent?.lrn)}
                        />
                        <InfoLine
                          label="Grade Level"
                          value={informationValue(selectedClass?.grade_level)}
                        />
                        <InfoLine
                          label="Section"
                          value={informationValue(selectedClass?.section)}
                        />
                        <InfoLine
                          label="School Year"
                          value={informationValue(selectedClass?.school_year)}
                        />
                        <InfoLine
                          label="Learning Modality"
                          value={informationValue(
                            selectedStudent?.learning_modality,
                          )}
                        />
                        <InfoLine
                          label="Remarks"
                          value={informationValue(selectedStudent?.remarks)}
                        />
                      </>
                    )}
                  </div>
                </div>
              </section>

              <section>
                <div className="mb-2 flex items-center gap-2 text-sm font-bold text-[var(--admin-text-heading)]">
                  <span className="grid size-5 place-items-center rounded-full border border-[var(--admin-primary)] text-[10px] text-[var(--admin-primary)]">
                    3
                  </span>
                  Request Details
                </div>
                <div className="grid gap-x-4 gap-y-2 rounded-2xl border border-[var(--admin-border)] bg-white p-4 text-xs sm:grid-cols-2">
                  <InfoLine
                    label="Credential"
                    value="School Form 1 (SF1) — School Register"
                  />
                  <InfoLine
                    label="Purpose"
                    value={informationValue(selectedRequest.reason)}
                  />
                  <InfoLine
                    label="Requesting Adviser"
                    value={
                      selectedRequester?.full_name ||
                      selectedRequester?.email ||
                      "—"
                    }
                  />
                  <InfoLine
                    label="Current Class"
                    value={
                      selectedClass
                        ? `${selectedClass.grade_level || "—"} - ${selectedClass.section || "—"}`
                        : "—"
                    }
                  />
                  <InfoLine
                    label="Previous Class"
                    value={
                      selectedPreviousClass
                        ? `${selectedPreviousClass.grade_level || "—"} - ${selectedPreviousClass.section || "—"}`
                        : "—"
                    }
                  />
                  <InfoLine
                    label="Request Date"
                    value={formatDateTimeText(
                      selectedRequest.request_date ||
                        selectedRequest.created_at,
                    )}
                  />
                </div>
              </section>

              <section>
                <div className="mb-2 flex items-center gap-2 text-sm font-bold text-[var(--admin-text-heading)]">
                  <span className="grid size-5 place-items-center rounded-full border border-[var(--admin-primary)] text-[10px] text-[var(--admin-primary)]">
                    4
                  </span>
                  Admin Remarks (Optional)
                </div>
                <div className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
                  <textarea
                    value={adminRemarks}
                    onChange={(event) =>
                      setAdminRemarks(event.target.value.slice(0, 250))
                    }
                    rows={4}
                    placeholder="Add remarks or instructions for this request..."
                    disabled={
                      selectedStatus === "approved" ||
                      selectedStatus === "rejected"
                    }
                    className="w-full resize-none rounded-xl border border-[var(--admin-border)] p-3 text-sm outline-none focus:border-[var(--admin-primary)] disabled:bg-[var(--admin-surface-soft)]"
                  />
                  <p className="mt-1 text-right text-[10px] text-[var(--admin-text-muted)]">
                    {adminRemarks.length} / 250
                  </p>
                </div>
              </section>

              <section>
                <div className="mb-2 flex items-center gap-2 text-sm font-bold text-[var(--admin-text-heading)]">
                  <span className="grid size-5 place-items-center rounded-full border border-[var(--admin-primary)] text-[10px] text-[var(--admin-primary)]">
                    5
                  </span>
                  Request Timeline
                </div>
                <div className="rounded-2xl border border-[var(--admin-border)] bg-white p-4 text-xs">
                  <div className="grid grid-cols-3 gap-3 text-center">
                    <TimelinePoint
                      active
                      label="Requested by Adviser"
                      detail={formatDateTimeText(
                        selectedRequest.request_date ||
                          selectedRequest.created_at,
                      )}
                    />
                    <TimelinePoint
                      active={selectedStatus !== "pending_review"}
                      label="Reviewed by Admin"
                      detail={formatDateTimeText(selectedRequest.reviewed_at)}
                    />
                    <TimelinePoint
                      active={selectedStatus === "approved"}
                      label="Approved"
                      detail={formatDateTimeText(
                        selectedRequest.processed_at ||
                          selectedRequest.reviewed_at,
                      )}
                    />
                  </div>
                  <div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 px-3 py-2 text-amber-800">
                    Approval allows the requesting adviser to choose one of
                    their classes and add this learner with the complete SF1
                    information.
                  </div>
                </div>
              </section>
            </div>

            <div className="flex flex-wrap justify-end gap-2 border-t border-[var(--admin-row-divider)] bg-white p-4">
              {selectedStatus === "pending_review" && (
                <>
                  <button
                    type="button"
                    onClick={() => submitReview("rejected")}
                    disabled={busyId === selectedRequest.id}
                    className="rounded-xl border border-red-500 px-4 py-2 text-sm font-semibold text-red-600 disabled:opacity-50"
                  >
                    Reject
                  </button>
                  <button
                    type="button"
                    onClick={() => submitReview("approved")}
                    disabled={busyId === selectedRequest.id || !selectedStudent}
                    className="rounded-xl bg-[var(--admin-primary)] px-4 py-2 text-sm font-semibold text-white disabled:opacity-50"
                  >
                    {busyId === selectedRequest.id && (
                      <Loader2 className="mr-2 inline size-4 animate-spin" />
                    )}
                    Approve Student
                  </button>
                </>
              )}
              {(selectedStatus === "approved" ||
                selectedStatus === "rejected") && (
                <button
                  type="button"
                  onClick={closeRequest}
                  className="rounded-xl border border-[var(--admin-border)] px-4 py-2 text-sm font-semibold"
                >
                  Close
                </button>
              )}
            </div>
          </aside>
        </div>
      )}
    </div>
  );
}

function InfoLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="grid grid-cols-[120px_1fr] gap-2">
      <span className="text-[var(--admin-text-muted)]">{label}</span>
      <span className="font-medium text-[var(--admin-text-heading)]">
        {value}
      </span>
    </div>
  );
}

function TimelinePoint({
  active,
  label,
  detail,
}: {
  active: boolean;
  label: string;
  detail: string;
}) {
  return (
    <div>
      <div
        className={`mx-auto mb-2 grid size-7 place-items-center rounded-full border-2 ${active ? "border-[var(--admin-primary)] bg-[var(--admin-primary)] text-white" : "border-gray-300 bg-gray-100 text-gray-400"}`}
      >
        <CircleCheck className="size-4" />
      </div>
      <p className="font-semibold">{label}</p>
      <p className="mt-1 text-[10px] text-[var(--admin-text-muted)]">
        {detail}
      </p>
    </div>
  );
}

function SubmittedFormsAdminView({
  onShowAvailable,
  onOpenFullPreview,
  classes,
  students,
  grades,
  schoolProfile,
}: {
  onShowAvailable: () => void;
  onOpenFullPreview: (submission: SchoolFormSubmissionRow) => void;
  classes: ClassRow[];
  students: StudentRow[];
  grades: GradeRow[];
  schoolProfile: ProfileRow | null;
}) {
  const queryClient = useQueryClient();
  const [selectedSchoolYear, setSelectedSchoolYear] = useState("all");
  const [selectedGrade, setSelectedGrade] = useState("all");
  const [selectedFormCode, setSelectedFormCode] = useState("all");
  const [selectedSubmission, setSelectedSubmission] =
    useState<SchoolFormSubmissionRow | null>(null);
  const [remarks, setRemarks] = useState("");
  const [reviewZoom, setReviewZoom] = useState(75);
  const [busyAction, setBusyAction] = useState<"approved" | "returned" | null>(
    null,
  );
  const [deletingSubmissionId, setDeletingSubmissionId] = useState<
    string | null
  >(null);
  const [deleteCandidate, setDeleteCandidate] =
    useState<SchoolFormSubmissionRow | null>(null);

  // Submitted Forms also follows the Admin School Year Library. Deleted years
  // must not remain in the dropdown or in the Submitted Forms UI.
  const { data: submittedSchoolYearLibrary = [] } = useQuery<
    AdminSchoolYearLibraryRow[]
  >({
    queryKey: ["school-year-library"],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("school_year_library")
        .select("id, school_year, is_locked")
        .order("school_year", { ascending: false });

      if (error) throw error;

      return (data ?? []).map((row: any) => ({
        id: Number(row.id),
        school_year: String(row.school_year ?? "").trim(),
        is_locked: Boolean(row.is_locked),
      }));
    },
    refetchOnWindowFocus: true,
  });

  const validSubmittedSchoolYears = useMemo(
    () =>
      new Set(
        submittedSchoolYearLibrary
          .map((row) => normalize(row.school_year))
          .filter(Boolean),
      ),
    [submittedSchoolYearLibrary],
  );

  useEffect(() => {
    if (
      selectedSchoolYear !== "all" &&
      !validSubmittedSchoolYears.has(normalize(selectedSchoolYear))
    ) {
      setSelectedSchoolYear("all");
      setSelectedSubmission(null);
      setRemarks("");
    }
  }, [selectedSchoolYear, validSubmittedSchoolYears]);

  const {
    data: submissions = [],
    isLoading,
    error,
  } = useQuery({
    queryKey: ["admin-school-form-submissions"],
    queryFn: async () => {
      const { data, error: submissionError } = await (supabase as any)
        .from("school_form_submissions")
        .select("*")
        .order("grade_level", { ascending: true })
        .order("adviser_name", { ascending: true })
        .order("submitted_at", { ascending: false });

      if (submissionError) throw submissionError;
      return (data ?? []) as SchoolFormSubmissionRow[];
    },
  });

  const normalizeFormCode = (code: string | null | undefined) => {
    const value = String(code ?? "").trim();

    if (value === "SOG" || value === "SOG Report") return "SOG Report";
    if (value === "SF9_NEW" || value === "SF9 (New)") return "SF9 (New)";
    if (value === "ANECDOTAL" || value === "Anecdotal") return "Anecdotal";

    return value;
  };

  const formLabel = (code: string | null | undefined) => {
    const normalizedCode = normalizeFormCode(code);

    return (
      [
        { code: "GSA", label: "GSA" },
        { code: "SF5", label: "SF5" },
        { code: "SOG Report", label: "SOG Report" },
        { code: "SF1", label: "SF1" },
        { code: "SF9 (New)", label: "SF9 (New)" },
        { code: "SF8", label: "SF8" },
        { code: "SF10", label: "SF10" },
        { code: "Anecdotal", label: "Anecdotal" },
      ].find((item) => item.code === normalizedCode)?.label ||
      normalizedCode ||
      "Unknown Form"
    );
  };

  const submittedSchoolYearOptions = useMemo(
    () =>
      submittedSchoolYearLibrary
        .map((row) => row.school_year)
        .filter(Boolean)
        .sort((firstYear, secondYear) => {
          const firstStartYear = Number(firstYear.match(/\d{4}/)?.[0] ?? 0);
          const secondStartYear = Number(secondYear.match(/\d{4}/)?.[0] ?? 0);
          return secondStartYear - firstStartYear;
        }),
    [submittedSchoolYearLibrary],
  );

  const librarySubmissions = useMemo(
    () =>
      submissions.filter((submission) =>
        validSubmittedSchoolYears.has(normalize(submission.school_year)),
      ),
    [submissions, validSubmittedSchoolYears],
  );

  const visibleSubmissions = useMemo(
    () =>
      librarySubmissions.filter((submission) => {
        const matchesSchoolYear =
          selectedSchoolYear === "all" ||
          normalize(submission.school_year) === normalize(selectedSchoolYear);

        const matchesGrade =
          selectedGrade === "all" ||
          submission.grade_level.match(/\d+/)?.[0] === selectedGrade;

        const matchesForm =
          selectedFormCode === "all" ||
          normalizeFormCode(submission.form_code) === selectedFormCode;

        return matchesSchoolYear && matchesGrade && matchesForm;
      }),
    [
      librarySubmissions,
      selectedFormCode,
      selectedGrade,
      selectedSchoolYear,
    ],
  );

  const groupedSubmissions = useMemo(() => {
    const groups = new Map<string, Map<string, SchoolFormSubmissionRow[]>>();

    visibleSubmissions.forEach((submission) => {
      const grade = submission.grade_level || "Unassigned Grade";
      const adviser = submission.adviser_name || "Unnamed Class Adviser";
      const adviserGroups = groups.get(grade) ?? new Map();

      adviserGroups.set(adviser, [
        ...(adviserGroups.get(adviser) ?? []),
        submission,
      ]);

      groups.set(grade, adviserGroups);
    });

    return [...groups.entries()].sort(([first], [second]) => {
      const firstNumber = Number(first.match(/\d+/)?.[0] ?? 999);
      const secondNumber = Number(second.match(/\d+/)?.[0] ?? 999);
      return firstNumber - secondNumber;
    });
  }, [visibleSubmissions]);

  const reviewSubmission = async (status: "approved" | "returned") => {
    if (!selectedSubmission) return;

    if (status === "returned" && !remarks.trim()) {
      window.alert(
        `Please enter the correction needed before returning ${formLabel(
          selectedSubmission.form_code,
        )}.`,
      );
      return;
    }

    // Approve immediately with no confirmation popup.
    // Keep confirmation only when returning a form for correction.
    if (status === "returned") {
      const confirmed = window.confirm(
        `Return ${formLabel(selectedSubmission.form_code)} to the Class Adviser for correction?`,
      );

      if (!confirmed) return;
    }

    setBusyAction(status);

    const { error: reviewError } = await supabase.rpc(
      "review_school_form_submission" as any,
      {
        p_submission_id: selectedSubmission.id,
        p_status: status,
        p_admin_remarks: remarks.trim() || null,
      } as any,
    );

    setBusyAction(null);

    if (reviewError) {
      window.alert(reviewError.message);
      return;
    }

    setSelectedSubmission(null);
    setRemarks("");

    await queryClient.invalidateQueries({
      queryKey: ["admin-school-form-submissions"],
    });
  };

  const deleteSubmission = async (submission: SchoolFormSubmissionRow) => {
    setDeletingSubmissionId(submission.id);

    const { error: deleteError } = await (supabase as any)
      .from("school_form_submissions")
      .delete()
      .eq("id", submission.id);

    setDeletingSubmissionId(null);

    if (deleteError) {
      window.alert(deleteError.message);
      return;
    }

    if (selectedSubmission?.id === submission.id) {
      setSelectedSubmission(null);
      setRemarks("");
      setReviewZoom(75);
    }

    setDeleteCandidate(null);

    await queryClient.invalidateQueries({
      queryKey: ["admin-school-form-submissions"],
    });
  };

  const statusCounts = {
    submitted: librarySubmissions.filter(
      (item) => item.status === "submitted" || item.status === "pending_review",
    ).length,
    approved: librarySubmissions.filter(
      (item) => item.status === "approved",
    ).length,
    returned: librarySubmissions.filter(
      (item) => item.status === "returned",
    ).length,
  };

  return (
    <div className="space-y-5">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
        <PageHeading
          title="School Forms"
          description="Review each submitted school form individually. Approval locks only the approved form for the Class Adviser."
        />

        <div className="grid w-full gap-2 sm:grid-cols-2 xl:w-auto xl:grid-cols-3">
          <div className="relative w-full sm:w-52">
            <CalendarRange className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-primary)]" />
            <select
              value={selectedSchoolYear}
              onChange={(event) => setSelectedSchoolYear(event.target.value)}
              aria-label="Filter submitted forms by school year"
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-9 text-sm font-medium outline-none focus:border-[var(--admin-primary)]"
            >
              <option value="all">All School Years</option>
              {submittedSchoolYearOptions.map((schoolYear) => (
                <option key={schoolYear} value={schoolYear}>
                  {schoolYear}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>

          <div className="relative w-full sm:w-52">
            <GraduationCap className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-primary)]" />
            <select
              value={selectedGrade}
              onChange={(event) => setSelectedGrade(event.target.value)}
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-9 text-sm font-medium outline-none focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Grade Levels</option>
              {[7, 8, 9, 10, 11, 12].map((grade) => (
                <option key={grade} value={grade}>
                  Grade {grade}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>

          <div className="relative w-full sm:w-52">
            <FileText className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-primary)]" />
            <select
              value={selectedFormCode}
              onChange={(event) => setSelectedFormCode(event.target.value)}
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-9 text-sm font-medium outline-none focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Forms</option>
              {[
                "GSA",
                "SF5",
                "SOG Report",
                "SF1",
                "SF9 (New)",
                "SF8",
                "SF10",
                "Anecdotal",
              ].map((code) => (
                <option key={code} value={code}>
                  {code}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>
        </div>
      </div>

      <div className="flex w-fit gap-1 rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-1">
        <button
          type="button"
          onClick={onShowAvailable}
          className="rounded-lg px-4 py-2 text-xs font-semibold text-[var(--admin-text-muted)] hover:bg-[var(--admin-surface-soft)]"
        >
          Available Forms
        </button>

        <button
          type="button"
          className="rounded-lg bg-[var(--admin-primary)] px-4 py-2 text-xs font-semibold text-white"
        >
          Submitted Forms
        </button>
      </div>

      <div className="grid gap-3 sm:grid-cols-3">
        {[
          [
            "Ready for Review",
            statusCounts.submitted,
            "text-amber-700 bg-amber-50",
          ],
          ["Approved", statusCounts.approved, "text-emerald-700 bg-emerald-50"],
          ["Returned", statusCounts.returned, "text-rose-700 bg-rose-50"],
        ].map(([label, count, tone]) => (
          <Panel key={String(label)}>
            <div className="flex items-center justify-between gap-3">
              <span className="text-xs font-semibold text-[var(--admin-text-muted)]">
                {label}
              </span>
              <span
                className={`rounded-full px-2.5 py-1 text-sm font-bold ${tone}`}
              >
                {count}
              </span>
            </div>
          </Panel>
        ))}
      </div>

      {isLoading && (
        <Panel>
          <div className="flex items-center gap-2 text-sm">
            <Loader2 className="size-4 animate-spin" />
            Loading submitted forms...
          </div>
        </Panel>
      )}

      {error && (
        <Panel>
          <div className="text-sm text-rose-700">
            {(error as Error).message}
          </div>
        </Panel>
      )}

      {!isLoading && !error && groupedSubmissions.length === 0 && (
        <Panel>
          <div className="py-10 text-center text-sm text-[var(--admin-text-muted)]">
            No submitted forms for the selected filters.
          </div>
        </Panel>
      )}

      {groupedSubmissions.map(([grade, adviserGroups]) => (
        <Panel key={grade}>
          <div className="mb-4 flex items-center gap-2 border-b border-[var(--admin-row-divider)] pb-3">
            <FolderOpen className="size-5 text-[var(--admin-primary)]" />
            <h2 className="text-base font-bold text-[var(--admin-text-heading)]">
              {grade}
            </h2>
            <span className="rounded-full bg-[var(--admin-color-f7e6e8)] px-2 py-1 text-[9px] font-bold text-[var(--admin-primary)]">
              {[...adviserGroups.values()].flat().length} form
              {[...adviserGroups.values()].flat().length === 1 ? "" : "s"}
            </span>
          </div>

          <div className="space-y-4">
            {[...adviserGroups.entries()].map(
              ([adviser, adviserSubmissions]) => (
                <div
                  key={adviser}
                  className="rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-3"
                >
                  <div className="mb-3 flex items-center gap-2">
                    <UserRound className="size-4 text-[var(--admin-primary)]" />
                    <div>
                      <div className="text-xs font-bold text-[var(--admin-text-heading)]">
                        {adviser}
                      </div>
                      <div className="text-[9px] text-[var(--admin-text-muted)]">
                        Class Adviser
                      </div>
                    </div>
                  </div>

                  <div className="grid gap-2 lg:grid-cols-2 xl:grid-cols-3">
                    {adviserSubmissions.map((submission) => {
                      const submitted = formatDateTime(submission.submitted_at);
                      const pending =
                        submission.status === "submitted" ||
                        submission.status === "pending_review";

                      return (
                        <div
                          key={submission.id}
                          className="rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-3"
                        >
                          <div className="flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="rounded-lg bg-[var(--admin-nav-hover)] px-2 py-1 text-[9px] font-bold text-[var(--admin-primary)]">
                                  {formLabel(submission.form_code)}
                                </span>
                                <span className="text-sm font-bold text-[var(--admin-text-heading)]">
                                  {submission.section || "No Section"}
                                </span>
                              </div>

                              <div className="mt-1 text-[9px] text-[var(--admin-text-muted)]">
                                {submission.school_year || "No school year"} ·
                                Submitted {submitted.date}, {submitted.time}
                              </div>
                            </div>

                            <span
                              className={`shrink-0 rounded-full px-2 py-1 text-[8px] font-bold capitalize ${
                                submission.status === "approved"
                                  ? "bg-emerald-50 text-emerald-700"
                                  : submission.status === "returned"
                                    ? "bg-rose-50 text-rose-700"
                                    : "bg-amber-50 text-amber-700"
                              }`}
                            >
                              {pending ? "Pending Review" : submission.status}
                            </span>
                          </div>

                          {submission.admin_remarks && (
                            <div className="mt-3 rounded-lg border border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-2 text-[9px] text-[var(--admin-text-muted)]">
                              <b>Admin remarks:</b> {submission.admin_remarks}
                            </div>
                          )}

                          <div className="mt-3 flex items-center justify-between gap-2">
                            <span className="text-[9px] font-semibold text-[var(--admin-text-muted)]">
                              Individual form submission
                            </span>

                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                onClick={() => setDeleteCandidate(submission)}
                                disabled={
                                  deletingSubmissionId === submission.id ||
                                  Boolean(busyAction)
                                }
                                aria-label={`Delete ${formLabel(
                                  submission.form_code,
                                )} submission`}
                                title="Delete submitted form"
                                className="grid size-7 place-items-center rounded-lg border border-rose-200 bg-white text-rose-600 transition hover:bg-rose-50 disabled:cursor-not-allowed disabled:opacity-50"
                              >
                                {deletingSubmissionId === submission.id ? (
                                  <Loader2 className="size-3.5 animate-spin" />
                                ) : (
                                  <Trash2 className="size-3.5" />
                                )}
                              </button>

                              <button
                                type="button"
                                onClick={() => {
                                  setSelectedSubmission(submission);
                                  setRemarks(submission.admin_remarks || "");
                                  setReviewZoom(75);
                                }}
                                disabled={
                                  deletingSubmissionId === submission.id
                                }
                                className="rounded-lg bg-[var(--admin-primary)] px-3 py-1.5 text-[9px] font-semibold text-white hover:bg-[var(--admin-primary-hover)] disabled:opacity-50"
                              >
                                {pending ? "Review Form" : "View Review"}
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ),
            )}
          </div>
        </Panel>
      ))}

      <Dialog
        open={Boolean(deleteCandidate)}
        onOpenChange={(open) => {
          if (!open && !deletingSubmissionId) {
            setDeleteCandidate(null);
          }
        }}
      >
        <DialogContent className="overflow-hidden p-0 sm:max-w-[500px]">
          <div className="p-6">
            <DialogHeader className="text-left">
              <div className="flex items-center gap-4">
                <div className="grid size-14 shrink-0 place-items-center rounded-full bg-rose-50 text-rose-600">
                  <Trash2 className="size-7" />
                </div>

                <DialogTitle className="text-2xl font-bold text-[var(--admin-text-heading)]">
                  Delete Submitted Form
                </DialogTitle>
              </div>
            </DialogHeader>

            <div className="mt-5 border-t border-[var(--admin-divider)] pt-5">
              <p className="text-sm font-medium text-[var(--admin-text-heading)]">
                Are you sure you want to delete this submitted form?
              </p>

              <div className="mt-4 rounded-xl bg-rose-50/70 p-4">
                <div className="grid grid-cols-[56px_1fr] gap-x-3 gap-y-2 text-sm">
                  <span className="font-semibold text-rose-700">Form:</span>
                  <span className="font-bold text-[var(--admin-text-heading)]">
                    {formLabel(deleteCandidate?.form_code)}
                  </span>

                  <span className="font-semibold text-rose-700">Class:</span>
                  <span className="font-bold text-[var(--admin-text-heading)]">
                    {deleteCandidate?.section || "No Section"}
                  </span>
                </div>
              </div>

              <p className="mt-4 text-sm leading-6 text-[var(--admin-text-muted)]">
                This will remove only this submitted-form record.
                <br />
                This action cannot be undone.
              </p>
            </div>

            <div className="mt-5 flex justify-end gap-3 border-t border-[var(--admin-divider)] pt-5">
              <Button
                type="button"
                variant="outline"
                onClick={() => setDeleteCandidate(null)}
                disabled={Boolean(deletingSubmissionId)}
                className="min-w-24 border-rose-300 text-rose-700 hover:bg-rose-50"
              >
                Cancel
              </Button>

              <Button
                type="button"
                onClick={() => {
                  if (deleteCandidate) {
                    void deleteSubmission(deleteCandidate);
                  }
                }}
                disabled={
                  !deleteCandidate ||
                  deletingSubmissionId === deleteCandidate?.id
                }
                className="min-w-32 bg-rose-700 text-white hover:bg-rose-800"
              >
                {deletingSubmissionId === deleteCandidate?.id ? (
                  <Loader2 className="mr-2 size-4 animate-spin" />
                ) : (
                  <Trash2 className="mr-2 size-4" />
                )}
                Delete Form
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(
          selectedSubmission &&
            formLabel(selectedSubmission.form_code) === "SF9 (New)",
        )}
        onOpenChange={(open) => {
          if (!open && !busyAction) {
            setSelectedSubmission(null);
            setRemarks("");
            setReviewZoom(75);
          }
        }}
      >
        {selectedSubmission &&
          formLabel(selectedSubmission.form_code) === "SF9 (New)" && (
            <Sf9ReviewWorkspace
              key={selectedSubmission.id}
              submission={selectedSubmission}
              classes={classes}
              students={students}
              grades={grades}
              schoolProfile={schoolProfile}
              remarks={remarks}
              setRemarks={setRemarks}
              reviewZoom={reviewZoom}
              setReviewZoom={setReviewZoom}
              busyAction={busyAction}
              onReview={reviewSubmission}
              onOpenFullPreview={() => onOpenFullPreview(selectedSubmission)}
              onClose={() => {
                setSelectedSubmission(null);
                setRemarks("");
                setReviewZoom(75);
              }}
            />
          )}
      </Dialog>

      <Dialog
        open={Boolean(
          selectedSubmission &&
            formLabel(selectedSubmission.form_code) !== "SF9 (New)",
        )}
        onOpenChange={(open) => {
          if (!open && !busyAction) {
            setSelectedSubmission(null);
            setRemarks("");
            setReviewZoom(75);
          }
        }}
      >
        <DialogContent className="max-h-[94vh] overflow-hidden p-0 sm:max-w-[1120px]">
          <div className="border-b border-[var(--admin-border)] px-5 py-4 pr-12">
            <DialogHeader className="space-y-1 text-left">
              <DialogTitle className="text-xl">
                Review {formLabel(selectedSubmission?.form_code)}
              </DialogTitle>

              <DialogDescription>
                {selectedSubmission?.grade_level} ·{" "}
                {selectedSubmission?.adviser_name || "Class Adviser"} ·{" "}
                {selectedSubmission?.section || "No Section"}
              </DialogDescription>
            </DialogHeader>
          </div>

          <div className="grid min-h-0 gap-4 overflow-y-auto p-5 lg:grid-cols-[300px_minmax(0,1fr)]">
            <div className="flex min-h-0 flex-col gap-4">
              <div className="rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-4">
                <div className="text-sm font-bold text-[var(--admin-text-heading)]">
                  Submitted form
                </div>
                <div className="mt-2 text-xs leading-5 text-[var(--admin-text-muted)]">
                  {formLabel(selectedSubmission?.form_code)} was submitted
                  individually. Approving it locks only this form for the Class
                  Adviser; the other school forms remain editable.
                </div>
              </div>

              <div className="flex min-h-[250px] flex-1 flex-col">
                <label
                  htmlFor="admin-form-remarks"
                  className="text-xs font-semibold text-[var(--admin-text-heading)]"
                >
                  Admin remarks
                </label>

                <textarea
                  id="admin-form-remarks"
                  value={remarks}
                  onChange={(event) => setRemarks(event.target.value)}
                  placeholder="Enter corrections or an optional approval note..."
                  disabled={
                    selectedSubmission?.status !== "submitted" &&
                    selectedSubmission?.status !== "pending_review"
                  }
                  className="mt-2 min-h-[220px] flex-1 resize-none rounded-2xl border border-[var(--admin-primary)]/60 bg-background px-3 py-3 text-sm outline-none focus:border-[var(--admin-primary)] disabled:opacity-60"
                />
              </div>

              <div className="grid gap-2">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => {
                    setSelectedSubmission(null);
                    setRemarks("");
                    setReviewZoom(75);
                  }}
                  disabled={Boolean(busyAction)}
                >
                  Close
                </Button>

                {(selectedSubmission?.status === "submitted" ||
                  selectedSubmission?.status === "pending_review") && (
                  <>
                    <Button
                      type="button"
                      variant="outline"
                      onClick={() => reviewSubmission("returned")}
                      disabled={Boolean(busyAction)}
                      className="border-rose-200 text-rose-700 hover:bg-rose-50"
                    >
                      {busyAction === "returned" && (
                        <Loader2 className="mr-2 size-4 animate-spin" />
                      )}
                      Return for Correction
                    </Button>

                    <Button
                      type="button"
                      onClick={() => reviewSubmission("approved")}
                      disabled={Boolean(busyAction)}
                      className="bg-[var(--admin-primary)] text-white hover:bg-[var(--admin-primary-hover)]"
                    >
                      {busyAction === "approved" && (
                        <Loader2 className="mr-2 size-4 animate-spin" />
                      )}
                      Approve Form
                    </Button>
                  </>
                )}
              </div>
            </div>

            <div className="min-w-0 rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-4">
              <div className="mb-3">
                <div className="text-sm font-bold text-[var(--admin-text-heading)]">
                  Bondpaper Preview
                </div>
                <div className="mt-1 text-xs text-[var(--admin-text-muted)]">
                  Preview of the submitted{" "}
                  {formLabel(selectedSubmission?.form_code)} as it will appear
                  on bondpaper.
                </div>
              </div>

              <div className="overflow-auto rounded-xl border border-[var(--admin-border)] bg-[#ece9e6] p-4">
                <div className="mx-auto flex min-h-[520px] min-w-[660px] items-start justify-center">
                  <div
                    className="origin-top bg-white shadow-xl transition-transform"
                    style={{
                      width: "760px",
                      minHeight: "500px",
                      transform: `scale(${reviewZoom / 100})`,
                      transformOrigin: "top center",
                    }}
                  >
                    {selectedSubmission && (
                      <SubmittedFormBondpaperMiniPreview
                        submission={selectedSubmission}
                        formLabel={formLabel(selectedSubmission.form_code)}
                      />
                    )}
                  </div>
                </div>
              </div>

              <div className="mt-3 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() =>
                      setReviewZoom((current) => Math.max(50, current - 10))
                    }
                    disabled={reviewZoom <= 50}
                    className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"
                    title="Zoom out"
                  >
                    <Minus className="size-4" />
                  </button>

                  <div className="min-w-16 rounded-lg border border-[var(--admin-border)] bg-white px-3 py-2 text-center text-xs font-semibold">
                    {reviewZoom}%
                  </div>

                  <button
                    type="button"
                    onClick={() =>
                      setReviewZoom((current) => Math.min(125, current + 10))
                    }
                    disabled={reviewZoom >= 125}
                    className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"
                    title="Zoom in"
                  >
                    <Plus className="size-4" />
                  </button>

                  <button
                    type="button"
                    onClick={() => setReviewZoom(75)}
                    className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white"
                    title="Reset preview zoom"
                  >
                    <Maximize2 className="size-4" />
                  </button>
                </div>

                <Button
                  type="button"
                  variant="outline"
                  onClick={() => {
                    if (selectedSubmission) {
                      onOpenFullPreview(selectedSubmission);
                    }
                  }}
                  disabled={!selectedSubmission || Boolean(busyAction)}
                  className="gap-2"
                >
                  Open Full Preview
                  <Maximize2 className="size-4" />
                </Button>
              </div>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
}


function Sf9ReviewWorkspace({
  submission,
  classes,
  students,
  grades,
  schoolProfile,
  remarks,
  setRemarks,
  reviewZoom,
  setReviewZoom,
  busyAction,
  onReview,
  onOpenFullPreview,
  onClose,
}: {
  submission: SchoolFormSubmissionRow;
  classes: ClassRow[];
  students: StudentRow[];
  grades: GradeRow[];
  schoolProfile: ProfileRow | null;
  remarks: string;
  setRemarks: (value: string) => void;
  reviewZoom: number;
  setReviewZoom: (value: number) => void;
  busyAction: "approved" | "returned" | null;
  onReview: (status: "approved" | "returned") => Promise<void>;
  onOpenFullPreview: () => void;
  onClose: () => void;
}) {
  const learners = useMemo(
    () =>
      students
        .filter((student) => student.class_id === submission.class_id)
        .sort((a, b) => studentName(a).localeCompare(studentName(b))),
    [students, submission.class_id],
  );

  const selectedClass =
    classes.find((item) => item.id === submission.class_id) ?? null;
  const [searchLearner, setSearchLearner] = useState("");
  const [page, setPage] = useState(1);
  const [selectedLearnerId, setSelectedLearnerId] = useState(
    learners[0]?.id ?? "",
  );
  const [checkedLearners, setCheckedLearners] = useState<Set<string>>(
    () => new Set(),
  );
  const [approveConfirmOpen, setApproveConfirmOpen] = useState(false);

  const isPending =
    submission.status === "submitted" ||
    submission.status === "pending_review";
  const isApproved = submission.status === "approved";
  const isReturned = submission.status === "returned";

  const filteredLearners = useMemo(() => {
    const q = normalize(searchLearner);
    if (!q) return learners;
    return learners.filter((learner) =>
      [studentName(learner), learner.lrn].some((value) =>
        normalize(value).includes(q),
      ),
    );
  }, [learners, searchLearner]);

  const pageSize = 10;
  const totalPages = Math.max(1, Math.ceil(filteredLearners.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const shownLearners = filteredLearners.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize,
  );
  const selectedLearner =
    learners.find((item) => item.id === selectedLearnerId) ?? learners[0] ?? null;
  const selectedIndex = selectedLearner
    ? learners.findIndex((item) => item.id === selectedLearner.id)
    : -1;
  const checkedCount = isApproved ? learners.length : checkedLearners.size;
  const progress = learners.length
    ? Math.round((checkedCount / learners.length) * 100)
    : 0;
  const canApprove =
    isPending && learners.length > 0 && checkedLearners.size === learners.length;

  const getStatus = (learner: StudentRow) => {
    if (isApproved) return ["Approved", "bg-emerald-50 text-emerald-700"] as const;
    if (isReturned) return ["Returned", "bg-rose-50 text-rose-700"] as const;
    if (checkedLearners.has(learner.id))
      return ["Checked", "bg-emerald-50 text-emerald-700"] as const;
    if (selectedLearner?.id === learner.id)
      return ["In Review", "bg-blue-50 text-blue-700"] as const;
    return ["Pending Review", "bg-amber-50 text-amber-700"] as const;
  };

  const markChecked = () => {
    if (!selectedLearner || !isPending) return;
    setCheckedLearners((current) => {
      const next = new Set(current);
      next.add(selectedLearner.id);
      return next;
    });

    const next = learners[selectedIndex + 1];
    if (next) setSelectedLearnerId(next.id);
  };

  const toggleAllChecked = () => {
    if (!isPending || learners.length === 0 || Boolean(busyAction)) return;

    const allChecked =
      checkedLearners.size === learners.length && learners.length > 0;

    if (allChecked) {
      // Cancel / unmark all learner checks.
      setCheckedLearners(new Set());
      return;
    }

    // Mark every learner as checked.
    setCheckedLearners(new Set(learners.map((learner) => learner.id)));
  };

  const openApproveConfirmation = () => {
    if (learners.length === 0 || Boolean(busyAction)) return;
    setApproveConfirmOpen(true);
  };

  const confirmApproveSf9 = async () => {
    if (!canApprove || Boolean(busyAction)) return;

    setApproveConfirmOpen(false);
    await onReview("approved");
  };

  const submitted = formatDateTime(submission.submitted_at);

  return (
    <DialogContent className="max-h-[96vh] overflow-y-auto border-[#ead8b8] bg-[#fffdf9] p-0 sm:max-w-[1500px]">
      <div className="border-b border-[var(--admin-divider)] px-5 py-4 pr-12 sm:px-7">
        <DialogHeader className="space-y-1 text-left">
          <div className="flex flex-wrap items-center gap-3">
            <DialogTitle className="text-2xl font-bold text-[var(--admin-text-heading)]">
              SF9 Review Forms
            </DialogTitle>
            <span className="rounded-full bg-rose-50 px-2.5 py-1 text-[9px] font-bold text-[var(--admin-primary)]">
              SF9 Only
            </span>
          </div>
          <DialogDescription>
            Review learner report cards submitted by the Class Adviser.
          </DialogDescription>
        </DialogHeader>
      </div>

      <div className="space-y-4 p-4 sm:p-5 lg:p-6">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="inline-flex items-center gap-2 rounded-xl border border-[var(--admin-border)] bg-white px-3 py-2 text-xs font-semibold">
            <Users className="size-4 text-[var(--admin-primary)]" />
            {submission.grade_level || selectedClass?.grade_level || "Grade"}
            <span className="text-[var(--admin-text-muted)]">/</span>
            {submission.section || selectedClass?.section || "Section"}
          </div>
          <div className="text-xs text-[var(--admin-text-muted)]">
            Class Adviser: <b className="text-[var(--admin-text-heading)]">{submission.adviser_name || "Class Adviser"}</b>
          </div>
        </div>

        <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
          <Sf9ReviewStat icon={FileText} label="Submitted" value={learners.length} helper="Learners in this class" tone="violet" />
          <Sf9ReviewStat icon={ClipboardCheck} label="In Review" value={isPending ? checkedLearners.size : 0} helper="Learners checked" tone="amber" />
          <Sf9ReviewStat icon={CircleCheck} label="Approved" value={isApproved ? learners.length : 0} helper="Ready for records" tone="green" />
          <Sf9ReviewStat icon={RefreshCcw} label="Returned" value={isReturned ? learners.length : 0} helper="For revision" tone="rose" />
        </div>

        <div className="grid min-h-[650px] gap-4 xl:grid-cols-[300px_minmax(0,1fr)_330px]">
          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-white">
            <div className="border-b border-[var(--admin-divider)] p-3">
              <div className="flex items-center justify-between gap-2">
                <h3 className="text-sm font-bold text-[var(--admin-text-heading)]">
                  Learner List ({learners.length})
                </h3>

                {isPending && (
                  <Button
                    type="button"
                    variant="outline"
                    onClick={toggleAllChecked}
                    disabled={learners.length === 0 || Boolean(busyAction)}
                    className={`h-8 shrink-0 px-2.5 text-[10px] font-semibold ${
                      checkedLearners.size === learners.length &&
                      learners.length > 0
                        ? "border-rose-200 text-rose-700 hover:bg-rose-50"
                        : "border-emerald-200 text-emerald-700 hover:bg-emerald-50"
                    }`}
                  >
                    {checkedLearners.size === learners.length &&
                    learners.length > 0 ? (
                      <>
                        <X className="mr-1.5 size-3.5" />
                        Unmark All
                      </>
                    ) : (
                      <>
                        <CircleCheck className="mr-1.5 size-3.5" />
                        Mark All
                      </>
                    )}
                  </Button>
                )}
              </div>

              <div className="relative mt-3">
                <Search className="pointer-events-none absolute left-3 top-1/2 size-3.5 -translate-y-1/2 text-[var(--admin-text-muted)]" />
                <Input
                  value={searchLearner}
                  onChange={(event) => {
                    setSearchLearner(event.target.value);
                    setPage(1);
                  }}
                  placeholder="Search learner..."
                  className="h-9 pl-9 text-xs"
                />
              </div>

              {isPending && learners.length > 0 && (
                <div className="mt-2 flex items-center justify-between text-[9px] text-[var(--admin-text-muted)]">
                  <span>
                    {checkedLearners.size} of {learners.length} checked
                  </span>

                  {checkedLearners.size === learners.length ? (
                    <span className="font-semibold text-emerald-700">
                      Ready for approval
                    </span>
                  ) : checkedLearners.size > 0 ? (
                    <span className="font-semibold text-amber-700">
                      Review in progress
                    </span>
                  ) : (
                    <span>Not yet checked</span>
                  )}
                </div>
              )}
            </div>

            <div className="max-h-[520px] flex-1 overflow-y-auto">
              {shownLearners.map((learner, index) => {
                const active = learner.id === selectedLearner?.id;
                const [statusLabel, statusClass] = getStatus(learner);
                const number = (currentPage - 1) * pageSize + index + 1;
                return (
                  <button
                    key={learner.id}
                    type="button"
                    onClick={() => setSelectedLearnerId(learner.id)}
                    className={`flex w-full items-center gap-2 border-b border-[var(--admin-divider)] px-3 py-3 text-left transition ${
                      active
                        ? "bg-[#fff5e8] shadow-[inset_3px_0_0_var(--admin-primary)]"
                        : "hover:bg-[var(--admin-surface-soft)]"
                    }`}
                  >
                    <span className="grid size-7 shrink-0 place-items-center rounded-full bg-[var(--admin-surface-soft)] text-[9px] font-bold text-[var(--admin-text-muted)]">{number}</span>
                    <span className="grid size-9 shrink-0 place-items-center rounded-full bg-[#fff0df] text-[10px] font-bold text-[#b76a25]">
                      {([learner.first_name?.[0], learner.last_name?.[0]].filter(Boolean).join("").toUpperCase() || "L")}
                    </span>
                    <span className="min-w-0 flex-1">
                      <span className="block truncate text-[11px] font-bold text-[var(--admin-text-heading)]">{studentName(learner)}</span>
                      <span className="mt-0.5 block truncate text-[9px] text-[var(--admin-text-muted)]">LRN: {learner.lrn || "Not provided"}</span>
                    </span>
                    <span className={`shrink-0 rounded-full px-2 py-1 text-[8px] font-bold ${statusClass}`}>{statusLabel}</span>
                  </button>
                );
              })}
              {shownLearners.length === 0 && (
                <div className="p-8 text-center text-xs text-[var(--admin-text-muted)]">No learners found.</div>
              )}
            </div>

            <div className="mt-auto flex items-center justify-between gap-2 border-t border-[var(--admin-divider)] p-3">
              <span className="text-[9px] text-[var(--admin-text-muted)]">
                {filteredLearners.length === 0
                  ? "0 learners"
                  : `${(currentPage - 1) * pageSize + 1}–${Math.min(currentPage * pageSize, filteredLearners.length)} of ${filteredLearners.length}`}
              </span>
              <div className="flex items-center gap-1">
                <button type="button" disabled={currentPage <= 1} onClick={() => setPage((value) => Math.max(1, value - 1))} className="grid size-8 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><ChevronLeft className="size-3.5" /></button>
                <span className="min-w-8 text-center text-[10px] font-semibold">{currentPage}/{totalPages}</span>
                <button type="button" disabled={currentPage >= totalPages} onClick={() => setPage((value) => Math.min(totalPages, value + 1))} className="grid size-8 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><ChevronRight className="size-3.5" /></button>
              </div>
            </div>
          </section>

          <section className="min-w-0 overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-white">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[var(--admin-divider)] px-4 py-3">
              <div>
                <h3 className="text-sm font-bold text-[var(--admin-text-heading)]">Bondpaper Preview</h3>
                <p className="mt-0.5 text-[9px] text-[var(--admin-text-muted)]">Preview of the selected learner&apos;s submitted SF9.</p>
              </div>
              <Button type="button" variant="outline" onClick={onOpenFullPreview} disabled={Boolean(busyAction)} className="h-9 gap-2 text-xs">
                Open Full Preview <Maximize2 className="size-3.5" />
              </Button>
            </div>

            <div className="overflow-auto bg-[#ece9e6] p-4">
              <div className="mx-auto flex min-h-[520px] min-w-[650px] items-start justify-center">
                <div className="origin-top bg-white shadow-xl transition-transform" style={{ width: "760px", minHeight: "500px", transform: `scale(${reviewZoom / 100})`, transformOrigin: "top center" }}>
                  {selectedLearner ? (
                    <Sf9LearnerBondpaperPreview learner={selectedLearner} submission={submission} selectedClass={selectedClass} classes={classes} grades={grades} schoolProfile={schoolProfile} />
                  ) : (
                    <div className="grid min-h-[500px] place-items-center p-8 text-center text-sm text-[var(--admin-text-muted)]">No learner is available for this class.</div>
                  )}
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-3 border-t border-[var(--admin-divider)] p-3">
              <div className="flex items-center gap-2">
                <button type="button" onClick={() => setReviewZoom(Math.max(50, reviewZoom - 10))} disabled={reviewZoom <= 50} className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><Minus className="size-4" /></button>
                <div className="min-w-16 rounded-lg border border-[var(--admin-border)] bg-white px-3 py-2 text-center text-xs font-semibold">{reviewZoom}%</div>
                <button type="button" onClick={() => setReviewZoom(Math.min(125, reviewZoom + 10))} disabled={reviewZoom >= 125} className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><Plus className="size-4" /></button>
                <button type="button" onClick={() => setReviewZoom(75)} className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white"><Maximize2 className="size-4" /></button>
              </div>
              <div className="flex items-center gap-2">
                <button type="button" disabled={selectedIndex <= 0} onClick={() => { const prev = learners[selectedIndex - 1]; if (prev) setSelectedLearnerId(prev.id); }} className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><ChevronLeft className="size-4" /></button>
                <button type="button" disabled={selectedIndex < 0 || selectedIndex >= learners.length - 1} onClick={() => { const next = learners[selectedIndex + 1]; if (next) setSelectedLearnerId(next.id); }} className="grid size-9 place-items-center rounded-lg border border-[var(--admin-border)] bg-white disabled:opacity-40"><ChevronRight className="size-4" /></button>
              </div>
            </div>
          </section>

          <aside className="flex min-h-0 flex-col gap-3">
            <section className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
              <h3 className="text-sm font-bold text-[var(--admin-text-heading)]">Review Details</h3>
              {selectedLearner && (
                <>
                  <div className="mt-3 flex items-center gap-3 rounded-xl border border-[var(--admin-divider)] bg-[var(--admin-surface-soft)] p-3">
                    <div className="grid size-11 shrink-0 place-items-center rounded-full bg-[#fff0df] text-sm font-bold text-[#b76a25]">{([selectedLearner.first_name?.[0], selectedLearner.last_name?.[0]].filter(Boolean).join("").toUpperCase() || "L")}</div>
                    <div className="min-w-0 flex-1">
                      <div className="truncate text-sm font-bold text-[var(--admin-text-heading)]">{studentName(selectedLearner)}</div>
                      <div className="mt-0.5 truncate text-[9px] text-[var(--admin-text-muted)]">LRN: {selectedLearner.lrn || "Not provided"}</div>
                    </div>
                  </div>
                  <div className="mt-3 space-y-2 text-[10px]">
                    <Sf9Detail label="Grade & Section" value={`${submission.grade_level || selectedClass?.grade_level || "—"} - ${submission.section || selectedClass?.section || "—"}`} />
                    <Sf9Detail label="Class Adviser" value={submission.adviser_name || "Class Adviser"} />
                    <Sf9Detail label="Submitted On" value={`${submitted.date}, ${submitted.time}`} />
                  </div>
                </>
              )}
            </section>

            <section className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
              <div className="flex items-center justify-between gap-3">
                <h3 className="text-sm font-bold text-[var(--admin-text-heading)]">Review Progress</h3>
                <span className="text-xs font-bold text-[var(--admin-primary)]">{progress}%</span>
              </div>
              <p className="mt-1 text-[9px] text-[var(--admin-text-muted)]">{checkedCount} of {learners.length} learners checked</p>
              <div className="mt-3 h-2 overflow-hidden rounded-full bg-[var(--admin-surface-soft)]"><div className="h-full rounded-full bg-[var(--admin-primary)] transition-[width]" style={{ width: `${progress}%` }} /></div>
            </section>

            <section className="flex min-h-0 flex-1 flex-col rounded-2xl border border-[var(--admin-border)] bg-white p-4">
              <label htmlFor="sf9-review-remarks" className="text-sm font-bold text-[var(--admin-text-heading)]">Remarks <span className="font-normal text-[var(--admin-text-muted)]">(Optional for approval)</span></label>
              <textarea id="sf9-review-remarks" value={remarks} onChange={(event) => setRemarks(event.target.value)} maxLength={500} disabled={!isPending} placeholder="Add remarks or notes about the submitted SF9..." className="mt-2 min-h-[120px] resize-none rounded-xl border border-[var(--admin-border)] bg-white px-3 py-3 text-xs outline-none focus:border-[var(--admin-primary)] disabled:bg-[var(--admin-surface-soft)] disabled:opacity-70" />
              <div className="mt-1 text-right text-[9px] text-[var(--admin-text-muted)]">{remarks.length}/500</div>

              <div className="mt-4 border-t border-[var(--admin-divider)] pt-4">
                <h4 className="text-xs font-bold text-[var(--admin-text-heading)]">Review Actions</h4>
                <div className="mt-3 grid gap-2">
                  {isPending ? (
                    <>
                      <Button type="button" variant="outline" onClick={markChecked} disabled={!selectedLearner || (selectedLearner ? checkedLearners.has(selectedLearner.id) : false) || Boolean(busyAction)} className="h-10 border-emerald-200 text-emerald-700 hover:bg-emerald-50"><CircleCheck className="mr-2 size-4" />{selectedLearner && checkedLearners.has(selectedLearner.id) ? "Learner Checked" : "Mark as Checked"}</Button>
                      <Button type="button" variant="outline" onClick={() => void onReview("returned")} disabled={Boolean(busyAction)} className="h-10 border-rose-200 text-rose-700 hover:bg-rose-50">{busyAction === "returned" ? <Loader2 className="mr-2 size-4 animate-spin" /> : <RefreshCcw className="mr-2 size-4" />}Return for Revision</Button>
                      <Button
                        type="button"
                        onClick={openApproveConfirmation}
                        disabled={Boolean(busyAction) || learners.length === 0}
                        className={`h-10 text-white ${
                          canApprove
                            ? "bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
                            : "bg-[var(--admin-primary)]/55 hover:bg-[var(--admin-primary)]/65"
                        }`}
                      >
                        {busyAction === "approved" ? (
                          <Loader2 className="mr-2 size-4 animate-spin" />
                        ) : (
                          <CircleCheck className="mr-2 size-4" />
                        )}
                        Approve SF9
                      </Button>
                      {!canApprove && learners.length > 0 && <p className="text-center text-[9px] leading-4 text-[var(--admin-text-muted)]">Check all learner SF9 reports before final approval.</p>}
                    </>
                  ) : (
                    <div className={`rounded-xl border p-3 text-center text-xs font-semibold ${isApproved ? "border-emerald-200 bg-emerald-50 text-emerald-700" : "border-rose-200 bg-rose-50 text-rose-700"}`}>{isApproved ? "This SF9 class submission is approved." : "This SF9 class submission was returned for revision."}</div>
                  )}
                  <Button type="button" variant="outline" onClick={onClose} disabled={Boolean(busyAction)} className="h-10">Close Review</Button>
                </div>
              </div>
            </section>
          </aside>
        </div>
      </div>

      <Dialog
        open={approveConfirmOpen}
        onOpenChange={(open) => {
          if (!busyAction) {
            setApproveConfirmOpen(open);
          }
        }}
      >
        <DialogContent className="overflow-hidden border-[#ead8b8] bg-[#fffdf9] p-0 sm:max-w-[520px]">
          <button
            type="button"
            aria-label="Close approval confirmation"
            onClick={() => setApproveConfirmOpen(false)}
            disabled={Boolean(busyAction)}
            className="absolute right-4 top-4 z-10 grid size-8 place-items-center rounded-lg text-[var(--admin-text-muted)] transition hover:bg-[var(--admin-surface-soft)] hover:text-[var(--admin-text-heading)] disabled:opacity-50"
          >
            <X className="size-4" />
          </button>

          <div className="px-6 pb-6 pt-7 sm:px-7">
            <DialogHeader className="items-center space-y-2 text-center">
              <div className="mb-2 grid size-16 place-items-center rounded-full border border-rose-200 bg-rose-50 text-[var(--admin-primary)]">
                <FileText className="size-7" />
              </div>

              <DialogTitle className="text-2xl font-bold text-[var(--admin-text-heading)]">
                Approve SF9 Reports?
              </DialogTitle>

              <DialogDescription className="max-w-sm text-center text-sm leading-6">
                Are you sure you want to approve the SF9 reports for this class?
              </DialogDescription>
            </DialogHeader>

            <div className="mt-6 overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-white">
              {[
                {
                  label: "Grade Level & Section",
                  value: `${submission.grade_level || selectedClass?.grade_level || "Grade"} - ${submission.section || selectedClass?.section || "Section"}`,
                  icon: Users,
                  className: "text-[#b76a25]",
                },
                {
                  label: "Total Learners",
                  value: learners.length,
                  icon: Users,
                  className: "text-[#b76a25]",
                },
                {
                  label: "Submitted",
                  value: learners.length,
                  icon: FileText,
                  className: "text-violet-600",
                },
                {
                  label: "In Review",
                  value: isPending ? checkedLearners.size : 0,
                  icon: ClipboardCheck,
                  className: "text-amber-600",
                },
                {
                  label: "Approved",
                  value: isApproved ? learners.length : 0,
                  icon: CircleCheck,
                  className: "text-emerald-600",
                },
                {
                  label: "Returned",
                  value: isReturned ? learners.length : 0,
                  icon: RefreshCcw,
                  className: "text-rose-600",
                },
              ].map((item, index, items) => {
                const Icon = item.icon;
                return (
                  <div
                    key={item.label}
                    className={`grid grid-cols-[1fr_auto] items-center gap-4 px-4 py-3 text-sm ${
                      index < items.length - 1
                        ? "border-b border-[var(--admin-divider)]"
                        : ""
                    }`}
                  >
                    <span className="flex items-center gap-2 text-[var(--admin-text-muted)]">
                      <Icon className={`size-4 ${item.className}`} />
                      {item.label}
                    </span>
                    <span className="font-bold text-[var(--admin-text-heading)]">
                      {item.value}
                    </span>
                  </div>
                );
              })}
            </div>

            <div
              className={`mt-4 rounded-2xl border px-4 py-3 text-xs leading-5 ${
                canApprove
                  ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                  : "border-rose-200 bg-rose-50 text-rose-700"
              }`}
            >
              <div className="flex items-start gap-3">
                {canApprove ? (
                  <CircleCheck className="mt-0.5 size-5 shrink-0" />
                ) : (
                  <RefreshCcw className="mt-0.5 size-5 shrink-0" />
                )}

                <p>
                  {canApprove
                    ? `All ${learners.length} learner SF9 reports have been checked. Once approved, the SF9 records will be finalized and ready for school records.`
                    : `Please check all ${learners.length} learner SF9 reports before approving the class submission. ${checkedLearners.size} of ${learners.length} reports are currently checked.`}
                </p>
              </div>
            </div>

            <DialogFooter className="mt-6 grid grid-cols-2 gap-3 sm:grid-cols-2">
              <Button
                type="button"
                variant="outline"
                onClick={() => setApproveConfirmOpen(false)}
                disabled={Boolean(busyAction)}
                className="h-11"
              >
                Cancel
              </Button>

              <Button
                type="button"
                onClick={() => void confirmApproveSf9()}
                disabled={!canApprove || Boolean(busyAction)}
                className="h-11 bg-[var(--admin-primary)] text-white hover:bg-[var(--admin-primary-hover)] disabled:bg-[var(--admin-primary)]/45"
              >
                {busyAction === "approved" ? (
                  <Loader2 className="mr-2 size-4 animate-spin" />
                ) : (
                  <CircleCheck className="mr-2 size-4" />
                )}
                Yes, Approve SF9
              </Button>
            </DialogFooter>

            {!canApprove && learners.length > 0 && (
              <p className="mt-3 text-center text-[10px] leading-4 text-[var(--admin-text-muted)]">
                Mark every learner as checked before the final approval button becomes available.
              </p>
            )}
          </div>
        </DialogContent>
      </Dialog>
    </DialogContent>
  );
}

function Sf9ReviewStat({
  icon: Icon,
  label,
  value,
  helper,
  tone,
}: {
  icon: ComponentType<{ className?: string }>;
  label: string;
  value: number;
  helper: string;
  tone: "violet" | "amber" | "green" | "rose";
}) {
  const toneClass = {
    violet: "border-violet-100 bg-violet-50 text-violet-700",
    amber: "border-amber-100 bg-amber-50 text-amber-700",
    green: "border-emerald-100 bg-emerald-50 text-emerald-700",
    rose: "border-rose-100 bg-rose-50 text-rose-700",
  }[tone];

  return (
    <div className="rounded-2xl border border-[var(--admin-border)] bg-white p-4 shadow-sm">
      <div className="flex items-center gap-3">
        <div className={`grid size-11 shrink-0 place-items-center rounded-full border ${toneClass}`}><Icon className="size-5" /></div>
        <div className="min-w-0">
          <div className="text-xs font-semibold text-[var(--admin-text-heading)]">{label}</div>
          <div className="text-xl font-bold text-[var(--admin-text-heading)]">{value}</div>
          <div className="text-[9px] text-[var(--admin-text-muted)]">{helper}</div>
        </div>
      </div>
    </div>
  );
}

function Sf9Detail({ label, value }: { label: string; value: string }) {
  return (
    <div className="grid grid-cols-[105px_minmax(0,1fr)] gap-2">
      <span className="text-[var(--admin-text-muted)]">{label}</span>
      <span className="font-semibold text-[var(--admin-text-heading)]">{value}</span>
    </div>
  );
}

function Sf9LearnerBondpaperPreview({
  learner,
  submission,
  selectedClass,
  classes,
  grades,
  schoolProfile,
}: {
  learner: StudentRow;
  submission: SchoolFormSubmissionRow;
  selectedClass: ClassRow | null;
  classes: ClassRow[];
  grades: GradeRow[];
  schoolProfile: ProfileRow | null;
}) {
  const schoolYear = submission.school_year || selectedClass?.school_year || "";
  const gradeLevel = submission.grade_level || selectedClass?.grade_level || "";
  const section = submission.section || selectedClass?.section || "";
  const schoolName = schoolProfile?.school_name || "Agusan del Sur National Science High School";

  const relatedClasses = classes.filter(
    (item) =>
      normalize(item.grade_level) === normalize(gradeLevel) &&
      normalize(item.section) === normalize(section) &&
      (!schoolYear || normalize(item.school_year) === normalize(schoolYear)),
  );
  const relatedClassIds = new Set(relatedClasses.map((item) => item.id));
  if (selectedClass?.id) relatedClassIds.add(selectedClass.id);

  const learnerGrades = grades.filter(
    (grade) =>
      grade.student_id === learner.id &&
      (relatedClassIds.size === 0 || relatedClassIds.has(grade.class_id)),
  );
  const subjectLabels = Array.from(
    new Map(
      [...relatedClasses.map((item) => item.subject), ...learnerGrades.map((item) => item.subject)]
        .filter((subject): subject is string => Boolean(subject?.trim()))
        .map((subject) => [normalize(subject), subject.trim()]),
    ).values(),
  ).slice(0, 11);

  const termNo = (value: string | null) => {
    const term = normalize(value).replaceAll(" ", "");
    if (["1", "1st", "first", "term1", "firstterm"].includes(term)) return 1;
    if (["2", "2nd", "second", "term2", "secondterm"].includes(term)) return 2;
    if (["3", "3rd", "third", "term3", "thirdterm"].includes(term)) return 3;
    return 0;
  };
  const rows = subjectLabels.map((subject) => {
    const matches = learnerGrades.filter((item) => normalize(item.subject) === normalize(subject));
    const averageTerm = (term: number) => {
      const values = matches
        .filter((item) => termNo(item.term) === term)
        .map((item) => item.score)
        .filter((value): value is number => typeof value === "number");
      return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
    };
    const term1 = averageTerm(1);
    const term2 = averageTerm(2);
    const term3 = averageTerm(3);
    const completed = [term1, term2, term3].filter((value): value is number => typeof value === "number");
    const final = completed.length ? completed.reduce((sum, value) => sum + value, 0) / completed.length : null;
    return { subject, term1, term2, term3, final };
  });
  const finals = rows.map((row) => row.final).filter((value): value is number => typeof value === "number");
  const generalAverage = finals.length ? finals.reduce((sum, value) => sum + value, 0) / finals.length : null;
  const showScore = (value: number | null) => (typeof value === "number" ? Math.round(value) : "");

  return (
    <div className="box-border min-h-[500px] bg-white px-5 py-4 text-black" style={{ fontFamily: "Arial, Helvetica, sans-serif" }}>
      <div className="grid grid-cols-[70px_minmax(0,1fr)_70px] items-start gap-3">
        <img src={deped} alt="DepEd" className="mx-auto size-11 object-contain" />
        <div className="text-center leading-tight">
          <div className="text-[5px]">Republic of the Philippines</div>
          <div className="text-[5px]">Department of Education</div>
          <div className="text-[5px]">{schoolProfile?.region || ""}</div>
          <div className="text-[5px]">{schoolProfile?.division || ""}</div>
          <div className="text-[6px] font-bold uppercase">{schoolName}</div>
          <div className="mt-1 text-[9px] font-bold">LEARNER&apos;S PERFORMANCE REPORT</div>
          <div className="text-[5px]">SY {schoolYear || "____-____"}</div>
        </div>
        <img src={logo} alt="School" className="mx-auto size-11 object-contain" />
      </div>

      <div className="mt-3 grid grid-cols-2 gap-4 text-[5px]">
        <div className="grid grid-cols-[55px_1fr] gap-y-1">
          <span>School Year</span><span className="border-b border-black">{schoolYear}</span>
          <span>Name</span><span className="border-b border-black font-semibold">{studentName(learner)}</span>
          <span>LRN</span><span className="border-b border-black">{learner.lrn || ""}</span>
          <span>Track/Strand</span><span className="border-b border-black">{selectedClass?.track_shs || ""}</span>
        </div>
        <div className="grid grid-cols-[32px_1fr] gap-y-1">
          <span>Age</span><span className="border-b border-black">{calculateAge(learner.birthdate)}</span>
          <span>Sex</span><span className="border-b border-black">{learner.sex || ""}</span>
          <span>Grade</span><span className="border-b border-black">{gradeLevel}</span>
          <span>Section</span><span className="border-b border-black">{section}</span>
        </div>
      </div>

      <div className="mt-3 grid grid-cols-[1.2fr_.8fr] gap-3">
        <div>
          <div className="border border-black py-1 text-center text-[5px] font-bold">LEARNING PROGRESS AND ACHIEVEMENTS</div>
          <table className="w-full table-fixed border-collapse text-[4.6px]">
            <thead><tr><th className="border border-black p-[2px] text-left">Learning Areas</th><th className="border border-black p-[2px]">1</th><th className="border border-black p-[2px]">2</th><th className="border border-black p-[2px]">3</th><th className="border border-black p-[2px]">Final</th><th className="border border-black p-[2px]">Remarks</th></tr></thead>
            <tbody>
              {Array.from({ length: 11 }, (_, index) => {
                const row = rows[index];
                return <tr key={index}><td className="h-[18px] border border-black px-[2px]">{row?.subject || ""}</td><td className="border border-black text-center">{showScore(row?.term1 ?? null)}</td><td className="border border-black text-center">{showScore(row?.term2 ?? null)}</td><td className="border border-black text-center">{showScore(row?.term3 ?? null)}</td><td className="border border-black text-center font-semibold">{showScore(row?.final ?? null)}</td><td className="border border-black text-center">{typeof row?.final === "number" ? (row.final >= 75 ? "Passed" : "Failed") : ""}</td></tr>;
              })}
              <tr><td colSpan={4} className="border border-black p-[2px] text-right font-bold">GENERAL AVERAGE</td><td className="border border-black text-center font-bold">{showScore(generalAverage)}</td><td className="border border-black text-center">{typeof generalAverage === "number" ? (generalAverage >= 75 ? "Promoted" : "Failed") : ""}</td></tr>
            </tbody>
          </table>

          <table className="mx-auto mt-2 w-[68%] border-collapse text-[4px]"><tbody>{[["Outstanding","90-100","Passed"],["Very Satisfactory","85-89","Passed"],["Satisfactory","80-84","Passed"],["Fairly Satisfactory","75-79","Passed"],["Did Not Meet Expectations","Below 75","Failed"]].map(([a,b,c]) => <tr key={a}><td className="border border-black px-1">{a}</td><td className="border border-black px-1 text-center">{b}</td><td className="border border-black px-1 text-center">{c}</td></tr>)}</tbody></table>
        </div>

        <div className="text-[4.3px]">
          <div className="border border-black py-1 text-center font-bold">ATTENDANCE RECORD</div>
          <table className="w-full border-collapse"><thead><tr><th className="border border-black p-[2px]">Month</th>{["Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan"].map((month) => <th key={month} className="border border-black p-[2px]">{month}</th>)}</tr></thead><tbody>{["School Days","Days Present","Days Absent"].map((label) => <tr key={label}><td className="border border-black p-[2px]">{label}</td>{Array.from({ length: 8 }, (_, index) => <td key={index} className="h-[15px] border border-black" />)}</tr>)}</tbody></table>
          <div className="mt-2 border border-black py-1 text-center font-bold">TEACHER&apos;S COMMENTS / REMARKS</div>
          {[1,2,3].map((term) => <div key={term} className="grid grid-cols-[34px_1fr] border-x border-b border-black"><div className="border-r border-black p-1">Term {term}</div><div className="h-[45px]" /></div>)}
          <div className="mt-3 text-center font-bold">PARENT&apos;S / GUARDIAN&apos;S SIGNATURE</div>
          {[1,2,3].map((term) => <div key={term} className="mt-3 flex items-end gap-2"><span>Term {term}</span><span className="flex-1 border-b border-black" /></div>)}
          <div className="mt-4 text-center font-bold">CERTIFICATE OF TRANSFER</div>
          <div className="mt-2 border-b border-black">&nbsp;</div>
          <div className="mt-2 grid grid-cols-2 gap-3"><div className="border-b border-black text-center">{schoolProfile?.principal || "School Head"}</div><div className="border-b border-black text-center">{submission.adviser_name || "Class Adviser"}</div></div>
          <div className="grid grid-cols-2 gap-3 text-center text-[3.7px]"><div>School Head</div><div>Class Adviser</div></div>
        </div>
      </div>
    </div>
  );
}

function SubmittedFormBondpaperMiniPreview({
  submission,
  formLabel,
}: {
  submission: SchoolFormSubmissionRow;
  formLabel: string;
}) {
  const snapshotClass =
    submission.snapshot &&
    typeof submission.snapshot === "object" &&
    "class" in submission.snapshot &&
    submission.snapshot.class &&
    typeof submission.snapshot.class === "object"
      ? (submission.snapshot.class as Record<string, unknown>)
      : null;

  const schoolYear =
    submission.school_year ||
    String(snapshotClass?.school_year ?? "") ||
    "____-____";
  const gradeLevel =
    submission.grade_level ||
    String(snapshotClass?.grade_level ?? "") ||
    "Grade";
  const section =
    submission.section || String(snapshotClass?.section ?? "") || "Section";

  const isSf5 = formLabel === "SF5";

  return (
    <div
      className="box-border min-h-[500px] bg-white p-5 text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="text-center">
        <div className="text-[13px] font-bold">
          {isSf5
            ? "School Form 5 (SF 5) Report on Promotion and Learning Progress & Achievement"
            : `${formLabel} School Form`}
        </div>
        <div className="mt-1 text-[7px] italic">
          Submitted form bondpaper preview
        </div>
      </div>

      <div className="mt-4 grid grid-cols-2 gap-x-5 gap-y-2 text-[7px]">
        <div className="flex items-center gap-2">
          <span className="w-16 text-right">School Year</span>
          <span className="flex-1 border border-black px-2 py-1 text-center">
            {schoolYear}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <span className="w-16 text-right">Grade Level</span>
          <span className="flex-1 border border-black px-2 py-1 text-center">
            {gradeLevel}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <span className="w-16 text-right">Adviser</span>
          <span className="flex-1 border border-black px-2 py-1 text-center">
            {submission.adviser_name || "Class Adviser"}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <span className="w-16 text-right">Section</span>
          <span className="flex-1 border border-black px-2 py-1 text-center">
            {section}
          </span>
        </div>
      </div>

      {isSf5 ? (
        <div className="mt-4 grid grid-cols-[2fr_1fr] gap-3">
          <table className="w-full border-collapse text-[6px]">
            <thead>
              <tr>
                {[
                  "LRN",
                  "LEARNER'S NAME",
                  "GENERAL AVERAGE",
                  "ACTION TAKEN",
                  "DID NOT MEET EXPECTATIONS",
                ].map((heading) => (
                  <th key={heading} className="border border-black p-1">
                    {heading}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: 8 }, (_, index) => (
                <tr key={index}>
                  <td className="h-6 border border-black" />
                  <td className="border border-black" />
                  <td className="border border-black" />
                  <td className="border border-black" />
                  <td className="border border-black" />
                </tr>
              ))}
              <tr>
                <td className="border border-black p-1 font-bold">0</td>
                <td className="border border-black p-1" colSpan={4}>
                  &lt;=== COMBINED
                </td>
              </tr>
            </tbody>
          </table>

          <div className="space-y-3">
            <table className="w-full border-collapse text-[6px]">
              <thead>
                <tr>
                  <th className="border border-black p-1">SUMMARY TABLE</th>
                  <th className="border border-black p-1">MALE</th>
                  <th className="border border-black p-1">FEMALE</th>
                  <th className="border border-black p-1">TOTAL</th>
                </tr>
              </thead>
              <tbody>
                {["PROMOTED", "PROMOTED WITH ACADEMIC AWARD", "RETAINED"].map(
                  (label) => (
                    <tr key={label}>
                      <td className="border border-black p-1">{label}</td>
                      <td className="border border-black text-center">0</td>
                      <td className="border border-black text-center">0</td>
                      <td className="border border-black text-center">0</td>
                    </tr>
                  ),
                )}
              </tbody>
            </table>

            <table className="w-full border-collapse text-[6px]">
              <thead>
                <tr>
                  <th className="border border-black p-1" colSpan={4}>
                    LEVEL OF PROFICIENCY AND ACHIEVEMENT
                  </th>
                </tr>
              </thead>
              <tbody>
                {[
                  "Did Not Meet Expectations",
                  "Fairly Satisfactory",
                  "Satisfactory",
                  "Very Satisfactory",
                  "Outstanding",
                ].map((label) => (
                  <tr key={label}>
                    <td className="border border-black p-1">{label}</td>
                    <td className="border border-black text-center">0</td>
                    <td className="border border-black text-center">0</td>
                    <td className="border border-black text-center">0</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="mt-5">
          <div className="border border-black p-2 text-center text-[8px] font-bold">
            {formLabel}
          </div>
          <table className="w-full border-collapse text-[6px]">
            <tbody>
              {Array.from({ length: 10 }, (_, index) => (
                <tr key={index}>
                  <td className="h-7 w-10 border border-black text-center">
                    {index + 1}
                  </td>
                  <td className="border border-black" />
                  <td className="w-24 border border-black" />
                  <td className="w-24 border border-black" />
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="mt-8 grid grid-cols-2 gap-16 text-[7px]">
        <div>
          <div>PREPARED BY:</div>
          <div className="mt-6 border-b border-black text-center font-semibold">
            {submission.adviser_name || "Class Adviser"}
          </div>
          <div className="text-center">Class Adviser</div>
        </div>
        <div>
          <div>CERTIFIED CORRECT &amp; SUBMITTED BY:</div>
          <div className="mt-6 border-b border-black" />
          <div className="text-center">School Head</div>
        </div>
      </div>
    </div>
  );
}

function SchoolFormsSection({
  classes,
  students,
  attendance,
  grades,
  teacherById,
  schoolProfile,
  search,
  onSearch,
}: {
  classes: ClassRow[];
  students: StudentRow[];
  attendance: AttendanceRow[];
  grades: GradeRow[];
  teacherById: Map<string, ProfileRow>;
  schoolProfile: ProfileRow | null;
  search: string;
  onSearch: (value: string) => void;
}) {
  const [formsView, setFormsView] = useState<"available" | "submitted">(
    "available",
  );
  const [selectedForm, setSelectedForm] = useState(FORM_TYPES[0]);
  const [selectedClassId, setSelectedClassId] = useState<string | null>(
    classes[0]?.id ?? null,
  );
  const [selectedSchoolYear, setSelectedSchoolYear] = useState("all");
  const [selectedGradeLevel, setSelectedGradeLevel] = useState("all");
  const [selectedDateRange, setSelectedDateRange] = useState("all");
  const [paperSize, setPaperSize] = useState<PaperSizeKey>("long");
  const [sogTerm, setSogTerm] = useState<SogTerm>("1");
  const [sf5ReportLength, setSf5ReportLength] =
    useState<Sf5ReportLength>("short");
  const [sf1ReportLength, setSf1ReportLength] =
    useState<Sf1ReportLength>("full");
  const [sf2Month, setSf2Month] = useState(
    () => new Date(new Date().getFullYear(), new Date().getMonth(), 1),
  );
  const [zoom, setZoom] = useState(90);
  const [selectedSf9LearnerId, setSelectedSf9LearnerId] = useState("");

  // The Admin School Forms School Year filter must follow the School Year
  // Library, not old class/submission records. Once a School Year is deleted
  // from the library, its Available Forms and Submitted Forms disappear from
  // this page while the historical class itself remains untouched.
  const { data: formSchoolYearLibrary = [] } = useQuery<
    AdminSchoolYearLibraryRow[]
  >({
    queryKey: ["school-year-library"],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("school_year_library")
        .select("id, school_year, is_locked")
        .order("school_year", { ascending: false });

      if (error) throw error;

      return (data ?? []).map((row: any) => ({
        id: Number(row.id),
        school_year: String(row.school_year ?? "").trim(),
        is_locked: Boolean(row.is_locked),
      }));
    },
    refetchOnWindowFocus: true,
  });

  const validFormSchoolYears = useMemo(
    () =>
      new Set(
        formSchoolYearLibrary
          .map((row) => normalize(row.school_year))
          .filter(Boolean),
      ),
    [formSchoolYearLibrary],
  );

  // If the currently selected School Year is deleted from the Library,
  // immediately return this filter to "All School Years".
  useEffect(() => {
    if (
      selectedSchoolYear !== "all" &&
      !validFormSchoolYears.has(normalize(selectedSchoolYear))
    ) {
      setSelectedSchoolYear("all");
      setSelectedClassId(null);
    }
  }, [selectedSchoolYear, validFormSchoolYears]);

  const visibleFormTypeDefinitions = useMemo(() => {
    return adminFormIdsForGradeLevel(selectedGradeLevel)
      .map((formId) =>
        FORM_TYPE_DEFINITIONS.find((form) => form.id === formId),
      )
      .filter((form): form is FormTypeDefinition => Boolean(form));
  }, [selectedGradeLevel]);

  // If the Admin changes from JHS to SHS (or vice versa) while a form that
  // does not belong to the new grade level is selected, automatically select
  // the first valid form for that grade level.
  useEffect(() => {
    if (visibleFormTypeDefinitions.length === 0) return;

    const selectedFormIsVisible = visibleFormTypeDefinitions.some(
      (form) => form.id === selectedForm,
    );

    if (!selectedFormIsVisible) {
      setSelectedForm(visibleFormTypeDefinitions[0].id);
      setSelectedClassId(null);
    }
  }, [selectedForm, visibleFormTypeDefinitions]);

  const selectedFormDefinition =
    visibleFormTypeDefinitions.find((form) => form.id === selectedForm) ??
    visibleFormTypeDefinitions[0] ??
    FORM_TYPE_DEFINITIONS[0];

  // Keep the existing Available Forms UI. We only read the individual
  // Class Adviser submissions so each existing file row can show its real
  // submission status.
  const { data: formSubmissions = [] } = useQuery<SchoolFormSubmissionRow[]>({
    queryKey: ["admin-school-form-submissions"],
    queryFn: async () => {
      const { data, error: submissionError } = await (supabase as any)
        .from("school_form_submissions")
        .select(
          "id,class_id,form_code,adviser_id,adviser_name,grade_level,section,school_year,status,admin_remarks,submitted_at,reviewed_at,reviewed_by,snapshot",
        )
        .order("submitted_at", { ascending: false });

      if (submissionError) throw submissionError;
      return (data ?? []) as SchoolFormSubmissionRow[];
    },
  });

  const selectedDatabaseFormCode =
    selectedForm === "SOG"
      ? "SOG Report"
      : selectedForm === "SF9_NEW"
        ? "SF9 (New)"
        : selectedForm === "ANECDOTAL"
          ? "Anecdotal"
          : selectedForm;

  const normalizeSubmissionFormCode = (formCode: string | null | undefined) => {
    const value = String(formCode ?? "").trim();

    if (value === "SOG") return "SOG Report";
    if (value === "SF9_NEW") return "SF9 (New)";
    if (value === "ANECDOTAL") return "Anecdotal";

    return value;
  };

  const latestSubmissionByClass = useMemo(() => {
    const submissionMap = new Map<string, SchoolFormSubmissionRow>();

    // The query is newest first, so the first matching row is the latest
    // submission for this exact class + exact form.
    formSubmissions.forEach((submission) => {
      if (
        normalizeSubmissionFormCode(submission.form_code) !==
        selectedDatabaseFormCode
      ) {
        return;
      }

      if (!submissionMap.has(submission.class_id)) {
        submissionMap.set(submission.class_id, submission);
      }
    });

    return submissionMap;
  }, [formSubmissions, selectedDatabaseFormCode]);

  const schoolYearOptions = useMemo(
    () =>
      formSchoolYearLibrary
        .map((row) => row.school_year)
        .filter(Boolean)
        .sort((firstYear, secondYear) => {
          const firstStartYear = Number(firstYear.match(/\d{4}/)?.[0] ?? 0);
          const secondStartYear = Number(secondYear.match(/\d{4}/)?.[0] ?? 0);
          return secondStartYear - firstStartYear;
        }),
    [formSchoolYearLibrary],
  );

  const filteredClasses = useMemo(() => {
    const q = normalize(search);
    return classes
      .filter((item) => {
        // A class can remain in the database for history, but it should no
        // longer generate an Admin Available Form after its School Year is
        // removed from the School Year Library.
        const classSchoolYear = normalize(item.school_year);
        const belongsToCurrentLibrary =
          Boolean(classSchoolYear) &&
          validFormSchoolYears.has(classSchoolYear);

        if (!belongsToCurrentLibrary) return false;

        const gradeNumber = item.grade_level?.match(/\d+/)?.[0] ?? "";
        const matchesGrade =
          selectedGradeLevel === "all" || gradeNumber === selectedGradeLevel;
        const matchesSchoolYear =
          selectedSchoolYear === "all" ||
          normalize(item.school_year) === normalize(selectedSchoolYear);

        if (!matchesGrade || !matchesSchoolYear) return false;

        if (selectedDateRange !== "all") {
          const createdAt = item.created_at
            ? new Date(item.created_at).getTime()
            : Number.NaN;
          if (Number.isNaN(createdAt)) return false;

          const now = Date.now();
          const rangeInMilliseconds: Record<string, number> = {
            now: 60 * 60 * 1000,
            day: 24 * 60 * 60 * 1000,
            week: 7 * 24 * 60 * 60 * 1000,
            month: 30 * 24 * 60 * 60 * 1000,
            year: 365 * 24 * 60 * 60 * 1000,
          };
          const selectedRange = rangeInMilliseconds[selectedDateRange];

          if (selectedRange && now - createdAt > selectedRange) return false;
        }

        if (!q) return true;

        const teacher = teacherById.get(item.teacher_id);
        return [
          item.subject,
          item.grade_level,
          item.section,
          item.school_year,
          teacher?.full_name,
        ].some((value) => normalize(value).includes(q));
      })
      .sort((firstClass, secondClass) => {
        const firstGrade = Number(
          firstClass.grade_level?.match(/\d+/)?.[0] ?? 999,
        );
        const secondGrade = Number(
          secondClass.grade_level?.match(/\d+/)?.[0] ?? 999,
        );

        if (firstGrade !== secondGrade) return firstGrade - secondGrade;

        return (firstClass.section ?? "").localeCompare(
          secondClass.section ?? "",
        );
      });
  }, [
    classes,
    search,
    selectedDateRange,
    selectedGradeLevel,
    selectedSchoolYear,
    teacherById,
    validFormSchoolYears,
  ]);

  // SF6 is an Admin-only summarized form. It does not depend on a
  // Class Adviser school_form_submissions record. Other forms keep the
  // existing submission-based selection flow unchanged.
  const selectedClass =
    selectedForm === "SF6"
      ? (filteredClasses.find((item) => item.id === selectedClassId) ??
        filteredClasses[0] ??
        null)
      : (filteredClasses.find(
          (item) =>
            item.id === selectedClassId && latestSubmissionByClass.has(item.id),
        ) ??
        filteredClasses.find((item) => latestSubmissionByClass.has(item.id)) ??
        null);

  const selectedClassSubmission = selectedClass
    ? (latestSubmissionByClass.get(selectedClass.id) ?? null)
    : null;

  const sf9ClassLearners = useMemo(() => {
    if (!selectedClass) return [];

    return students
      .filter((student) => student.class_id === selectedClass.id)
      .sort((first, second) =>
        studentName(first).localeCompare(studentName(second)),
      );
  }, [selectedClass, students]);

  const selectedSf9Learner =
    sf9ClassLearners.find(
      (learner) => learner.id === selectedSf9LearnerId,
    ) ??
    sf9ClassLearners[0] ??
    null;

  const isGsaPreview = selectedForm === "GSA";
  const isSogPreview = selectedForm === "SOG";
  const isSf5Preview = selectedForm === "SF5";
  const isSf1Preview = selectedForm === "SF1";
  const isSf9NewPreview = selectedForm === "SF9_NEW";
  const isGrade11Sf9NewPreview =
    isSf9NewPreview && adminSf9GradeKey(selectedClass?.grade_level) === "grade 11";
  const isGrade12Sf9NewPreview =
    isSf9NewPreview && adminSf9GradeKey(selectedClass?.grade_level) === "grade 12";
  const isShsSf9NewPreview = isGrade11Sf9NewPreview || isGrade12Sf9NewPreview;
  const isSf6Preview = selectedForm === "SF6";
  const isSf8Preview = selectedForm === "SF8";
  const isSf10Preview = selectedForm === "SF10";
  // Retained legacy rendering code is unreachable because these forms are no
  // longer present in FORM_TYPE_DEFINITIONS.
  const isMpsPreview = false;
  const isIlsPreview = false;
  const isGradeSlipPreview = false;
  const isSf2Preview = false;
  const isSf9Preview = false;
  const isLirPreview = false;
  const isIndexCardPreview = false;
  const isSchoolIdPreview = false;
  const isLandscapePreview =
    isGsaPreview ||
    isSogPreview ||
    isSf5Preview ||
    isSf1Preview ||
    isSf9NewPreview ||
    isSf6Preview ||
    isSf8Preview;

  const selectedPaperPreset = isSf5Preview
    ? PAPER_SIZE_PRESETS.long
    : PAPER_SIZE_PRESETS[paperSize];
  const selectedPaper = isLandscapePreview
    ? {
        ...selectedPaperPreset,
        width: selectedPaperPreset.height,
        height: selectedPaperPreset.width,
      }
    : selectedPaperPreset;
  const zoomScale = zoom / 100;
  const previewZoomScale =
    zoomScale * (isSf5Preview ? ADMIN_SF5_PREVIEW_SCALE : 1);
  const scaledPaperWidth = Math.ceil(selectedPaper.width * previewZoomScale);
  const scaledPaperHeight = Math.ceil(selectedPaper.height * previewZoomScale);

  const selectedTeacherName = selectedClass
    ? selectedClass.teacher_name ||
      teacherById.get(selectedClass.teacher_id)?.full_name ||
      teacherById.get(selectedClass.teacher_id)?.email ||
      "Teacher Name"
    : "Teacher Name";

  const sf2MonthValue = `${sf2Month.getFullYear()}-${String(
    sf2Month.getMonth() + 1,
  ).padStart(2, "0")}`;
  const sf2MonthLabel = new Intl.DateTimeFormat("en-PH", {
    month: "long",
    year: "numeric",
  }).format(sf2Month);
  const sf2MonthName = new Intl.DateTimeFormat("en-PH", {
    month: "long",
  }).format(sf2Month);

  const moveSf2Month = (amount: number) => {
    setSf2Month(
      (current) =>
        new Date(current.getFullYear(), current.getMonth() + amount, 1),
    );
  };

  const gsaMetrics = useMemo(() => {
    if (!selectedClass) return null;

    const classStudents = students.filter(
      (student) => student.class_id === selectedClass.id,
    );

    const selectedSubject = normalize(selectedClass.subject);
    const gradeByStudent = new Map<string, number>();

    grades.forEach((grade) => {
      const sameClass = grade.class_id === selectedClass.id;
      const sameTerm = normalize(grade.term) === "1";
      const sameSubject =
        !selectedSubject || normalize(grade.subject) === selectedSubject;

      if (
        sameClass &&
        sameTerm &&
        sameSubject &&
        typeof grade.score === "number"
      ) {
        gradeByStudent.set(grade.student_id, grade.score);
      }
    });

    const summarize = (sex: "male" | "female") => {
      const learners = classStudents.filter((student) =>
        normalize(student.sex).startsWith(sex[0]),
      );
      const scores = learners
        .map((student) => gradeByStudent.get(student.id))
        .filter((score): score is number => typeof score === "number");
      const totalGrades = scores.reduce((sum, score) => sum + score, 0);
      const passed = scores.filter((score) => score >= 75).length;
      const registered = learners.length;

      return {
        registered,
        passed,
        totalGrades,
        average: registered ? totalGrades / registered : 0,
        proficiency: registered ? Math.round((passed / registered) * 100) : 0,
      };
    };

    const male = summarize("male");
    const female = summarize("female");
    const totalRegistered = male.registered + female.registered;
    const totalPassed = male.passed + female.passed;
    const totalGrades = male.totalGrades + female.totalGrades;

    return {
      male,
      female,
      total: {
        registered: totalRegistered,
        passed: totalPassed,
        totalGrades,
        average: totalRegistered ? totalGrades / totalRegistered : 0,
        proficiency: totalRegistered
          ? Math.round((totalPassed / totalRegistered) * 100)
          : 0,
      },
    };
  }, [grades, selectedClass, students]);

  const mpsMetrics = useMemo(() => {
    if (!selectedClass) return null;

    const highestTestScore = 50;
    const passingScore = highestTestScore * 0.75;

    const classStudents = students.filter(
      (student) => student.class_id === selectedClass.id,
    );

    const selectedSubject = normalize(selectedClass.subject);
    const scoreByStudent = new Map<string, number>();

    grades.forEach((grade) => {
      const term = normalize(grade.term).replaceAll(" ", "");
      const sameClass = grade.class_id === selectedClass.id;
      const sameTerm = ["1", "1st", "first", "firstterm", "term1"].includes(
        term,
      );
      const sameSubject =
        !selectedSubject || normalize(grade.subject) === selectedSubject;

      if (
        sameClass &&
        sameTerm &&
        sameSubject &&
        typeof grade.score === "number"
      ) {
        const score =
          grade.score > highestTestScore
            ? (grade.score / 100) * highestTestScore
            : grade.score;

        scoreByStudent.set(grade.student_id, score);
      }
    });

    const summarize = (sex: "male" | "female") => {
      const learners = classStudents.filter((student) =>
        normalize(student.sex).startsWith(sex[0]),
      );

      const scores = learners
        .map((student) => scoreByStudent.get(student.id))
        .filter((score): score is number => typeof score === "number");

      const totalScore = scores.reduce((sum, score) => sum + score, 0);
      const registered = learners.length;
      const passed = scores.filter((score) => score >= passingScore).length;
      const meanScore = registered ? totalScore / registered : 0;
      const meanPercentage = highestTestScore
        ? (meanScore / highestTestScore) * 100
        : 0;

      return {
        registered,
        passed,
        totalScore,
        meanScore,
        meanPercentage,
      };
    };

    const male = summarize("male");
    const female = summarize("female");
    const totalRegistered = male.registered + female.registered;
    const totalPassed = male.passed + female.passed;
    const totalScore = male.totalScore + female.totalScore;
    const meanScore = totalRegistered ? totalScore / totalRegistered : 0;

    return {
      highestTestScore,
      male,
      female,
      total: {
        registered: totalRegistered,
        passed: totalPassed,
        totalScore,
        meanScore,
        meanPercentage: highestTestScore
          ? (meanScore / highestTestScore) * 100
          : 0,
      },
    };
  }, [grades, selectedClass, students]);

  const sogMetrics = useMemo(() => {
    if (!selectedClass) return null;

    const relatedClasses = classes.filter(
      (item) =>
        item.grade_level === selectedClass.grade_level &&
        item.section === selectedClass.section,
    );
    const relatedClassIds = new Set(relatedClasses.map((item) => item.id));
    relatedClassIds.add(selectedClass.id);

    const classStudents = students.filter((student) =>
      relatedClassIds.has(student.class_id),
    );

    const learnerGroups = new Map<
      string,
      { learner: StudentRow; studentIds: string[] }
    >();

    classStudents.forEach((student) => {
      const key = student.lrn?.trim()
        ? `lrn:${student.lrn.trim()}`
        : `name:${normalize(student.last_name)}|${normalize(student.first_name)}|${normalize(student.middle_name)}`;
      const existing = learnerGroups.get(key);

      if (existing) {
        existing.studentIds.push(student.id);
      } else {
        learnerGroups.set(key, {
          learner: student,
          studentIds: [student.id],
        });
      }
    });

    const relatedGrades = grades.filter(
      (grade) =>
        relatedClassIds.has(grade.class_id) && typeof grade.score === "number",
    );

    const termNumber = (term: string | null) => {
      const value = normalize(term).replaceAll(" ", "");
      if (["1", "1st", "first", "term1", "firstterm"].includes(value)) return 1;
      if (["2", "2nd", "second", "term2", "secondterm"].includes(value))
        return 2;
      if (["3", "3rd", "third", "term3", "thirdterm"].includes(value)) return 3;
      return 0;
    };

    const defaultSubjects = [
      "MAPEH",
      "Science",
      "Values Education",
      "Math",
      "Filipino",
      "English",
      "Mathematics",
      "Araling Panlipunan",
      "TLE",
    ];

    const subjectLabels = Array.from(
      new Map(
        [
          ...relatedClasses.map((item) => item.subject),
          ...relatedGrades.map((grade) => grade.subject),
          ...(relatedClasses.length || relatedGrades.length
            ? []
            : defaultSubjects),
        ]
          .filter((subject): subject is string => Boolean(subject?.trim()))
          .map((subject) => [normalize(subject), subject.trim()]),
      ).values(),
    ).slice(0, 11);

    const selectedTermNumber = Number(sogTerm);
    const descriptorFor = (average: number | null) => {
      if (typeof average !== "number") return "";
      if (average >= 90) return "Outstanding";
      if (average >= 85) return "Very Satisfactory";
      if (average >= 80) return "Satisfactory";
      if (average >= 75) return "Fairly Satisfactory";
      return "Did Not Meet Expectations";
    };

    const learners = Array.from(learnerGroups.values())
      .map(({ learner, studentIds }) => {
        const studentIdSet = new Set(studentIds);
        const subjectScores = subjectLabels.map((subject) => {
          const scores = relatedGrades
            .filter(
              (grade) =>
                studentIdSet.has(grade.student_id) &&
                termNumber(grade.term) === selectedTermNumber &&
                normalize(grade.subject) === normalize(subject),
            )
            .map((grade) => grade.score)
            .filter((score): score is number => typeof score === "number");

          const score = scores.length
            ? scores.reduce((sum, item) => sum + item, 0) / scores.length
            : null;

          return { subject, score };
        });

        const availableScores = subjectScores
          .map((item) => item.score)
          .filter((score): score is number => typeof score === "number");
        const termAverage = availableScores.length
          ? availableScores.reduce((sum, score) => sum + score, 0) /
            availableScores.length
          : null;
        const sexValue = normalize(learner.sex);
        const sexGroup = sexValue.startsWith("m")
          ? "male"
          : sexValue.startsWith("f")
            ? "female"
            : "unspecified";

        return {
          learner,
          name: studentName(learner),
          lrn: learner.lrn || "",
          sexGroup,
          subjectScores,
          termAverage,
          rank: null as number | null,
          descriptor: descriptorFor(termAverage),
        };
      })
      .sort((a, b) => a.name.localeCompare(b.name));

    const ranked = learners
      .filter(
        (
          learner,
        ): learner is (typeof learners)[number] & { termAverage: number } =>
          typeof learner.termAverage === "number",
      )
      .sort((a, b) => b.termAverage - a.termAverage);

    let previousAverage: number | null = null;
    let previousRank = 0;
    ranked.forEach((learner, index) => {
      const roundedAverage = Number(learner.termAverage.toFixed(4));
      const rank =
        previousAverage !== null && roundedAverage === previousAverage
          ? previousRank
          : index + 1;
      learner.rank = rank;
      previousAverage = roundedAverage;
      previousRank = rank;
    });

    return {
      term: sogTerm,
      termLabel: `TERM ${sogTerm}`,
      subjects: subjectLabels,
      learners,
    };
  }, [classes, grades, selectedClass, sogTerm, students]);

  const sf5Metrics = useMemo(() => {
    if (!selectedClass) return null;

    const classStudents = students.filter(
      (student) => student.class_id === selectedClass.id,
    );

    const fallbackSubject = selectedClass.subject?.trim();
    let summarySubjects = fallbackSubject ? [fallbackSubject] : [];

    try {
      const storedValue = window.localStorage.getItem(
        `sog-subjects-${selectedClass.id}`,
      );
      const parsedSubjects: unknown = storedValue
        ? JSON.parse(storedValue)
        : [];

      if (Array.isArray(parsedSubjects)) {
        const uniqueSubjects = Array.from(
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
        );

        if (uniqueSubjects.length > 0) summarySubjects = uniqueSubjects;
      }
    } catch {
      // Keep the selected class subject as the safe Summary-of-Grades fallback.
    }

    const learners = classStudents
      .map((student) => {
        const subjectAverages = summarySubjects.map((subject) => {
          const termScores = ["1", "2", "3"]
            .map(
              (term) =>
                grades.find(
                  (grade) =>
                    grade.class_id === selectedClass.id &&
                    grade.student_id === student.id &&
                    normalize(grade.subject) === normalize(subject) &&
                    grade.term === term,
                )?.score ?? null,
            )
            .filter((score): score is number => typeof score === "number");

          return termScores.length
            ? termScores.reduce((sum, score) => sum + score, 0) /
                termScores.length
            : null;
        });
        const availableSubjectAverages = subjectAverages.filter(
          (average): average is number => typeof average === "number",
        );
        const generalAverage = availableSubjectAverages.length
          ? availableSubjectAverages.reduce(
              (sum, average) => sum + average,
              0,
            ) / availableSubjectAverages.length
          : 0;
        const sexValue = normalize(student.sex);
        const sex = sexValue.startsWith("m")
          ? "male"
          : sexValue.startsWith("f")
            ? "female"
            : "unspecified";

        return {
          id: student.id,
          lrn: student.lrn || "-",
          name: studentName(student),
          sex,
          generalAverage,
          action: getSF5Action(
            availableSubjectAverages.length ? generalAverage : null,
          ),
          didNotMeet: subjectAverages
            .map((average, index) =>
              average != null && average < 75 ? summarySubjects[index] : null,
            )
            .filter((subject): subject is string => Boolean(subject))
            .join(", "),
        };
      })
      .sort((a, b) => a.name.localeCompare(b.name));

    const countBySex = (
      predicate: (learner: (typeof learners)[number]) => boolean,
    ) => {
      const matching = learners.filter(predicate);
      return {
        male: matching.filter((learner) => learner.sex === "male").length,
        female: matching.filter((learner) => learner.sex === "female").length,
        total: matching.length,
      };
    };

    const promoted = countBySex((learner) => learner.generalAverage >= 75);
    const conditional = { male: 0, female: 0, total: 0 };
    const retained = countBySex(
      (learner) => learner.generalAverage > 0 && learner.generalAverage < 75,
    );

    const gradeBands = [
      {
        label: "Outstanding (90-100)",
        ...countBySex(
          (learner) =>
            learner.generalAverage >= 90 && learner.generalAverage <= 100,
        ),
      },
      {
        label: "Very Satisfactory (85-89)",
        ...countBySex(
          (learner) =>
            learner.generalAverage >= 85 && learner.generalAverage < 90,
        ),
      },
      {
        label: "Satisfactory (80-84)",
        ...countBySex(
          (learner) =>
            learner.generalAverage >= 80 && learner.generalAverage < 85,
        ),
      },
      {
        label: "Fairly Satisfactory (75-79)",
        ...countBySex(
          (learner) =>
            learner.generalAverage >= 75 && learner.generalAverage < 80,
        ),
      },
      {
        label: "Did Not Meet Expectations (74 and below)",
        ...countBySex(
          (learner) =>
            learner.generalAverage > 0 && learner.generalAverage < 75,
        ),
      },
    ];

    return {
      learners,
      maleTotal: learners.filter((learner) => learner.sex === "male").length,
      femaleTotal: learners.filter((learner) => learner.sex === "female")
        .length,
      promoted,
      conditional,
      retained,
      gradeBands,
    };
  }, [grades, selectedClass, students]);

  const sf1Metrics = useMemo(() => {
    if (!selectedClass) return null;

    const learners = students
      .filter((student) => student.class_id === selectedClass.id)
      .map((student) => {
        const normalizedSex = normalize(student.sex);
        const sex = normalizedSex.startsWith("m")
          ? "male"
          : normalizedSex.startsWith("f")
            ? "female"
            : "unspecified";

        return {
          ...student,
          sexGroup: sex,
          age: calculateAge(student.birthdate),
        };
      })
      .sort((a, b) => {
        const lastNameComparison = normalize(a.last_name).localeCompare(
          normalize(b.last_name),
        );

        if (lastNameComparison !== 0) return lastNameComparison;

        return normalize(a.first_name).localeCompare(normalize(b.first_name));
      });

    const male = learners.filter((learner) => learner.sexGroup === "male");
    const female = learners.filter((learner) => learner.sexGroup === "female");
    const unspecified = learners.filter(
      (learner) => learner.sexGroup === "unspecified",
    );

    return {
      learners,
      male,
      female,
      unspecified,
      total: learners.length,
    };
  }, [selectedClass, students]);

  const sf2Metrics = useMemo(() => {
    if (!selectedClass) return null;

    const year = sf2Month.getFullYear();
    const month = sf2Month.getMonth();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const schoolDays = Array.from({ length: daysInMonth }, (_, index) => {
      const date = new Date(year, month, index + 1);
      return {
        date,
        day: index + 1,
        key: `${year}-${String(month + 1).padStart(2, "0")}-${String(index + 1).padStart(2, "0")}`,
        weekday: new Intl.DateTimeFormat("en-PH", {
          weekday: "narrow",
        }).format(date),
      };
    }).filter(({ date }) => date.getDay() !== 0 && date.getDay() !== 6);

    const classStudents = students
      .filter((student) => student.class_id === selectedClass.id)
      .sort((a, b) => studentName(a).localeCompare(studentName(b)));

    const attendanceByStudentAndDay = new Map<string, string>();

    attendance.forEach((record) => {
      if (record.class_id !== selectedClass.id || !record.date) return;

      const recordDate = new Date(`${record.date}T00:00:00`);
      if (
        Number.isNaN(recordDate.getTime()) ||
        recordDate.getFullYear() !== year ||
        recordDate.getMonth() !== month
      ) {
        return;
      }

      const key = `${record.student_id}-${record.date.slice(0, 10)}`;
      attendanceByStudentAndDay.set(key, normalize(record.status));
    });

    const markForStatus = (status: string) => {
      if (status === "absent") return "X";
      if (status === "late") return "L";
      if (status === "excused") return "E";
      return "";
    };

    const isPresentStatus = (status: string) =>
      status === "present" || status === "late" || status === "excused";

    const learners = classStudents.map((student) => {
      const dayStatuses = schoolDays.map((day) => {
        const status =
          attendanceByStudentAndDay.get(`${student.id}-${day.key}`) || "";
        return {
          ...day,
          status,
          mark: markForStatus(status),
        };
      });

      const present = dayStatuses.filter((day) =>
        isPresentStatus(day.status),
      ).length;
      const absent = dayStatuses.filter(
        (day) => day.status === "absent",
      ).length;

      return {
        ...student,
        name: studentName(student),
        sexGroup: normalize(student.sex).startsWith("m")
          ? "male"
          : normalize(student.sex).startsWith("f")
            ? "female"
            : "unspecified",
        dayStatuses,
        present,
        absent,
      };
    });

    const dailyTotals = schoolDays.map((day) => {
      const matchingStatuses = learners.map(
        (learner) =>
          learner.dayStatuses.find((status) => status.key === day.key)
            ?.status || "",
      );

      const countPresent = (sex?: "male" | "female") =>
        learners.filter((learner, learnerIndex) => {
          if (sex && learner.sexGroup !== sex) return false;
          return isPresentStatus(matchingStatuses[learnerIndex]);
        }).length;

      return {
        key: day.key,
        male: countPresent("male"),
        female: countPresent("female"),
        combined: countPresent(),
      };
    });

    const maleEnrollment = learners.filter(
      (learner) => learner.sexGroup === "male",
    ).length;
    const femaleEnrollment = learners.filter(
      (learner) => learner.sexGroup === "female",
    ).length;
    const totalEnrollment = learners.length;
    const totalPresent = learners.reduce(
      (sum, learner) => sum + learner.present,
      0,
    );
    const totalAbsent = learners.reduce(
      (sum, learner) => sum + learner.absent,
      0,
    );
    const averageDailyAttendance = schoolDays.length
      ? totalPresent / schoolDays.length
      : 0;
    const attendancePercentage =
      schoolDays.length && totalEnrollment
        ? (totalPresent / (schoolDays.length * totalEnrollment)) * 100
        : 0;

    return {
      schoolDays,
      learners,
      dailyTotals,
      maleEnrollment,
      femaleEnrollment,
      totalEnrollment,
      totalPresent,
      totalAbsent,
      averageDailyAttendance,
      attendancePercentage,
    };
  }, [attendance, selectedClass, sf2Month, students]);

  const ilsMetrics = useMemo<IlsMetricsData | null>(() => {
    if (!selectedClass) return null;

    const classStudents = students
      .filter((student) => student.class_id === selectedClass.id)
      .sort((a, b) => studentName(a).localeCompare(studentName(b)));
    const learner = classStudents[0] ?? null;
    const relatedClasses = classes.filter(
      (item) =>
        item.grade_level === selectedClass.grade_level &&
        item.section === selectedClass.section,
    );
    const relatedClassIds = new Set(relatedClasses.map((item) => item.id));
    relatedClassIds.add(selectedClass.id);

    const termNumber = (term: string | null) => {
      const value = normalize(term).replaceAll(" ", "");
      if (["1", "1st", "first", "term1", "firstterm"].includes(value)) return 1;
      if (["2", "2nd", "second", "term2", "secondterm"].includes(value))
        return 2;
      if (["3", "3rd", "third", "term3", "thirdterm"].includes(value)) return 3;
      return 0;
    };

    const learnerGrades = learner
      ? grades.filter(
          (grade) =>
            grade.student_id === learner.id &&
            relatedClassIds.has(grade.class_id) &&
            typeof grade.score === "number",
        )
      : [];

    const defaultSubjects = [
      "Science",
      "Values Education",
      "MAPEH",
      "Araling Panlipunan",
      "Mathematics",
      "English",
      "Filipino",
    ];

    const subjectLabels = Array.from(
      new Map(
        [
          ...relatedClasses.map((item) => item.subject),
          ...learnerGrades.map((grade) => grade.subject),
          ...defaultSubjects,
        ]
          .filter((subject): subject is string => Boolean(subject?.trim()))
          .map((subject) => [normalize(subject), subject.trim()]),
      ).values(),
    ).slice(0, 8);

    const scoreFor = (subject: string, term: number) => {
      const scores = learnerGrades
        .filter(
          (grade) =>
            normalize(grade.subject) === normalize(subject) &&
            termNumber(grade.term) === term,
        )
        .map((grade) => grade.score)
        .filter((score): score is number => typeof score === "number");

      return scores.length
        ? scores.reduce((sum, score) => sum + score, 0) / scores.length
        : null;
    };

    const subjects = subjectLabels.map((subject) => {
      const term1 = scoreFor(subject, 1);
      const term2 = scoreFor(subject, 2);
      const term3 = scoreFor(subject, 3);
      const completed = [term1, term2, term3].filter(
        (score): score is number => typeof score === "number",
      );
      const average = completed.length
        ? completed.reduce((sum, score) => sum + score, 0) / completed.length
        : null;
      const firstAvailable = completed[0] ?? null;
      const lastAvailable = completed[completed.length - 1] ?? null;
      const trend =
        firstAvailable === null || lastAvailable === null
          ? "→"
          : lastAvailable > firstAvailable
            ? "✓"
            : lastAvailable < firstAvailable
              ? "↓"
              : "→";

      return {
        subject,
        term1,
        term2,
        term3,
        average,
        trend,
        remarks:
          typeof average === "number"
            ? average >= 75
              ? "PASSED"
              : "FOR INTERVENTION"
            : "",
      };
    });

    const subjectAverages = subjects
      .map((subject) => subject.average)
      .filter((score): score is number => typeof score === "number");
    const generalAverage = subjectAverages.length
      ? subjectAverages.reduce((sum, score) => sum + score, 0) /
        subjectAverages.length
      : null;

    const averageForStudent = (studentId: string) => {
      const studentScores = grades
        .filter(
          (grade) =>
            grade.student_id === studentId &&
            relatedClassIds.has(grade.class_id) &&
            typeof grade.score === "number",
        )
        .map((grade) => grade.score as number);
      return studentScores.length
        ? studentScores.reduce((sum, score) => sum + score, 0) /
            studentScores.length
        : null;
    };

    const rankedAverages = classStudents
      .map((student) => ({
        id: student.id,
        average: averageForStudent(student.id),
      }))
      .filter(
        (item): item is { id: string; average: number } =>
          typeof item.average === "number",
      )
      .sort((a, b) => b.average - a.average);
    const learnerRankIndex = learner
      ? rankedAverages.findIndex((item) => item.id === learner.id)
      : -1;
    const classAverage = rankedAverages.length
      ? rankedAverages.reduce((sum, item) => sum + item.average, 0) /
        rankedAverages.length
      : null;

    const learnerAttendance = learner
      ? attendance.filter(
          (record) =>
            record.student_id === learner.id &&
            record.class_id === selectedClass.id,
        )
      : [];
    const absences = learnerAttendance.filter(
      (record) => normalize(record.status) === "absent",
    ).length;
    const lateHalfDays = learnerAttendance.filter((record) => {
      const status = normalize(record.status).replaceAll("_", " ");
      return (
        status === "late" ||
        status.includes("half day") ||
        status.includes("half-day")
      );
    }).length;
    const failingSubjectCount = subjects.filter(
      (subject) => typeof subject.average === "number" && subject.average < 75,
    ).length;
    const missingActivities = 0;
    const riskLevel: "LOW" | "MODERATE" | "HIGH" =
      failingSubjectCount >= 3 || absences >= 10
        ? "HIGH"
        : failingSubjectCount >= 1 || absences >= 3
          ? "MODERATE"
          : "LOW";

    const componentSubject =
      selectedClass.subject || subjectLabels[0] || "Learning Area";
    const componentRows = [1, 2].map((term) => {
      const score = scoreFor(componentSubject, term);
      const percentage =
        typeof score === "number" ? Math.max(0, Math.min(100, score)) : null;
      const display = (maximum: number) =>
        percentage === null
          ? "—"
          : `${Math.round((percentage / 100) * maximum)}/${maximum} (${Math.round(percentage)}%)`;

      return {
        termLabel: term === 1 ? "1st Term" : "2nd Term",
        writtenWorks: display(50),
        performanceTasks: display(75),
        summativeExam: display(130),
      };
    });

    const recommendations: string[] = [];
    if (failingSubjectCount > 0) {
      recommendations.push(
        `Provide focused remediation in ${failingSubjectCount} learning area${failingSubjectCount === 1 ? "" : "s"}.`,
      );
    }
    if (absences > 0 || lateHalfDays > 0) {
      recommendations.push(
        "Strengthen attendance monitoring and parent coordination.",
      );
    }
    if (recommendations.length === 0) {
      recommendations.push(
        "Continue to support the learner's academic progress.",
      );
    }

    return {
      learner,
      parentGuardian:
        learner?.guardian ||
        learner?.mother_name ||
        learner?.father_name ||
        "Parent/Guardian",
      subjects,
      generalAverage,
      classRank: learnerRankIndex >= 0 ? learnerRankIndex + 1 : null,
      classSize: classStudents.length,
      classAverage,
      failingSubjectCount,
      absences,
      lateHalfDays,
      missingActivities,
      riskLevel,
      componentSubject,
      componentRows,
      recommendations,
    };
  }, [attendance, classes, grades, selectedClass, students]);

  const sf9Metrics = useMemo(() => {
    if (!selectedClass) return null;

    const learner = selectedSf9Learner;
    const selectedGrade = adminSf9GradeKey(selectedClass.grade_level);
    const isGrade12 = selectedGrade === "grade 12" || selectedGrade === "12";

    const relatedClasses = classes.filter(
      (item) =>
        adminSf9GradeKey(item.grade_level) === selectedGrade &&
        normalize(item.section) === normalize(selectedClass.section) &&
        String(item.school_year || "") === String(selectedClass.school_year || ""),
    );
    const relatedClassIds = new Set(relatedClasses.map((item) => item.id));

    const learnerNameKey = (row: StudentRow) =>
      [row.last_name, row.first_name, row.middle_name]
        .map((value) => normalize(value))
        .join("|");
    const learnerIds = new Set<string>();
    if (learner) {
      learnerIds.add(learner.id);
      const lrn = String(learner.lrn || "").trim();
      const nameKey = learnerNameKey(learner);
      students.forEach((candidate) => {
        if (!relatedClassIds.has(candidate.class_id)) return;
        const candidateLrn = String(candidate.lrn || "").trim();
        const sameLearner = lrn && candidateLrn
          ? lrn === candidateLrn
          : learnerNameKey(candidate) === nameKey;
        if (sameLearner) learnerIds.add(candidate.id);
      });
    }

    const gradeRows = learner
      ? grades.filter((grade) => learnerIds.has(grade.student_id))
      : [];

    const defaultSubjects = [
      "MAPEH",
      "Science",
      "Values Education",
      "Mathematics",
      "Filipino",
      "English",
      "Araling Panlipunan",
      "TLE",
    ];

    const subjectLabels = isGrade12
      ? ADMIN_SF9_GRADE12_SUBJECTS.map((subject) => subject.label)
      : Array.from(
          new Map(
            [
              ...relatedClasses.map((item) => item.subject),
              ...gradeRows.map((grade) => grade.subject),
              ...(relatedClasses.length || gradeRows.length ? [] : defaultSubjects),
            ]
              .filter((subject): subject is string => Boolean(subject?.trim()))
              .map((subject) => [normalizeAdminSf9Subject(subject), subject.trim()]),
          ).values(),
        );

    const termNumber = (term: string | null) => {
      const value = normalize(term).replaceAll(" ", "");
      if (["1", "1st", "first", "term1", "firstterm"].includes(value)) return 1;
      if (["2", "2nd", "second", "term2", "secondterm"].includes(value)) return 2;
      if (["3", "3rd", "third", "term3", "thirdterm"].includes(value)) return 3;
      return 0;
    };

    const subjects = subjectLabels.map((subject) => {
      const matching = gradeRows.filter(
        (grade) => normalizeAdminSf9Subject(grade.subject) === normalizeAdminSf9Subject(subject),
      );
      const scoreForTerm = (term: number) => {
        const scores = matching
          .filter((grade) => termNumber(grade.term) === term)
          .map((grade) => grade.score)
          .filter((score): score is number => typeof score === "number");
        return scores.length
          ? scores.reduce((sum, score) => sum + score, 0) / scores.length
          : null;
      };
      const term1 = scoreForTerm(1);
      const term2 = scoreForTerm(2);
      const term3 = scoreForTerm(3);
      const grade12Spec = isGrade12
        ? ADMIN_SF9_GRADE12_SUBJECTS.find(
            (item) => normalizeAdminSf9Subject(item.label) === normalizeAdminSf9Subject(subject),
          )
        : undefined;
      const assignedTerm = grade12Spec?.term as AdminSf9AssignedTerm | undefined;
      const completedTerms = [term1, term2, term3].filter(
        (score): score is number => typeof score === "number",
      );
      const final = assignedTerm
        ? assignedTerm === "1"
          ? term1
          : assignedTerm === "2"
            ? term2
            : term3
        : completedTerms.length
          ? completedTerms.reduce((sum, score) => sum + score, 0) / completedTerms.length
          : null;
      const matchingClass = relatedClasses.find(
        (item) => normalizeAdminSf9Subject(item.subject) === normalizeAdminSf9Subject(subject),
      );
      const numericUnits = Number(matchingClass?.units);
      const units = matchingClass?.units != null && Number.isFinite(numericUnits)
        ? numericUnits
        : null;

      return {
        subject,
        term1,
        term2,
        term3,
        final,
        remarks: typeof final === "number" ? (final >= 75 ? "PASSED" : "FAILED") : "",
        units,
        assignedTerm: assignedTerm ?? null,
        group: grade12Spec?.group ?? null,
      };
    });

    const finalScores = subjects
      .map((subject) => subject.final)
      .filter((score): score is number => typeof score === "number");
    const generalAverage = finalScores.length
      ? finalScores.reduce((sum, score) => sum + score, 0) / finalScores.length
      : null;

    return { learner, subjects, generalAverage };
  }, [classes, grades, selectedClass, selectedSf9Learner, students]);

  const sf10Metrics = useMemo(() => {
    if (!selectedClass) return null;

    const learner =
      students
        .filter((student) => student.class_id === selectedClass.id)
        .sort((a, b) => studentName(a).localeCompare(studentName(b)))[0] ??
      null;

    const gradeRows = learner
      ? grades.filter((grade) => grade.student_id === learner.id)
      : [];

    const subjectLabels = classes
      .filter(
        (item) =>
          normalize(item.grade_level) ===
            normalize(selectedClass.grade_level) &&
          normalize(item.section) === normalize(selectedClass.section) &&
          normalize(item.school_year) === normalize(selectedClass.school_year),
      )
      .map((item) => item.subject?.trim())
      .filter((subject): subject is string => Boolean(subject));

    gradeRows.forEach((grade) => {
      const subject = grade.subject?.trim();
      if (
        subject &&
        !subjectLabels.some(
          (existingSubject) =>
            normalize(existingSubject) === normalize(subject),
        )
      ) {
        subjectLabels.push(subject);
      }
    });

    const termNumber = (term: string | null) => {
      const value = normalize(term).replaceAll(" ", "");
      if (["1", "1st", "first", "term1", "firstterm"].includes(value)) return 1;
      if (["2", "2nd", "second", "term2", "secondterm"].includes(value))
        return 2;
      if (["3", "3rd", "third", "term3", "thirdterm"].includes(value)) return 3;
      if (["4", "4th", "fourth", "term4", "fourthterm"].includes(value))
        return 4;
      return 0;
    };

    const subjects = subjectLabels.slice(0, 11).map((subject) => {
      const matching = gradeRows.filter(
        (grade) => normalize(grade.subject) === normalize(subject),
      );

      const scoreForTerm = (term: number) => {
        const scores = matching
          .filter((grade) => termNumber(grade.term) === term)
          .map((grade) => grade.score)
          .filter((score): score is number => typeof score === "number");

        return scores.length
          ? scores.reduce((sum, score) => sum + score, 0) / scores.length
          : null;
      };

      const term1 = scoreForTerm(1);
      const term2 = scoreForTerm(2);
      const term3 = scoreForTerm(3);
      const term4 = scoreForTerm(4);
      const completedTerms = [term1, term2, term3, term4].filter(
        (score): score is number => typeof score === "number",
      );
      const final = completedTerms.length
        ? completedTerms.reduce((sum, score) => sum + score, 0) /
          completedTerms.length
        : null;

      return {
        subject,
        term1,
        term2,
        term3,
        term4,
        final,
        remarks:
          typeof final === "number" ? (final >= 75 ? "PASSED" : "FAILED") : "",
      };
    });

    const finalScores = subjects
      .map((subject) => subject.final)
      .filter((score): score is number => typeof score === "number");
    const generalAverage = finalScores.length
      ? finalScores.reduce((sum, score) => sum + score, 0) / finalScores.length
      : null;

    return {
      learner,
      subjects,
      generalAverage,
    };
  }, [classes, grades, selectedClass, students]);

  const sf6Metrics = useMemo<Sf6MetricsData | null>(() => {
    if (!selectedClass) return null;

    const sexOf = (student: StudentRow) => {
      const value = normalize(student.sex);
      if (value.startsWith("m")) return "male" as const;
      if (value.startsWith("f")) return "female" as const;
      return "unspecified" as const;
    };

    const selectedSchoolYear = normalize(selectedClass.school_year);
    const schoolYearClasses = classes.filter(
      (item) =>
        !selectedSchoolYear ||
        normalize(item.school_year) === selectedSchoolYear,
    );
    const schoolYearClassIds = new Set(
      schoolYearClasses.map((item) => item.id),
    );
    const schoolYearGrades = grades.filter(
      (grade) =>
        schoolYearClassIds.has(grade.class_id) &&
        typeof grade.score === "number",
    );

    const countLearners = (
      learners: Array<{
        student: StudentRow;
        sex: "male" | "female" | "unspecified";
        generalAverage: number | null;
      }>,
      predicate: (learner: {
        student: StudentRow;
        sex: "male" | "female" | "unspecified";
        generalAverage: number | null;
      }) => boolean,
    ): Sf6Count => {
      const matching = learners.filter(predicate);
      const male = matching.filter((learner) => learner.sex === "male").length;
      const female = matching.filter(
        (learner) => learner.sex === "female",
      ).length;

      return {
        male,
        female,
        total: male + female,
      };
    };

    const gradeReports = SF6_GRADE_LEVELS.map((gradeLevel) => {
      const gradeClasses = schoolYearClasses.filter(
        (item) => sf6GradeNumber(item.grade_level) === gradeLevel.value,
      );
      const gradeClassIds = new Set(gradeClasses.map((item) => item.id));
      const gradeStudents = Array.from(
        new Map(
          students
            .filter((student) => gradeClassIds.has(student.class_id))
            .map((student) => [student.id, student]),
        ).values(),
      );

      const learnerSummaries = gradeStudents.map((student) => {
        const studentGrades = schoolYearGrades.filter(
          (grade) =>
            grade.student_id === student.id &&
            gradeClassIds.has(grade.class_id),
        );
        const gradesBySubject = new Map<string, GradeRow[]>();

        studentGrades.forEach((grade) => {
          const subjectKey =
            normalize(grade.subject) || `class:${grade.class_id}`;
          const current = gradesBySubject.get(subjectKey) ?? [];
          current.push(grade);
          gradesBySubject.set(subjectKey, current);
        });

        const subjectAverages = Array.from(gradesBySubject.values())
          .map((subjectGrades) => {
            const finalScores = subjectGrades
              .filter((grade) => normalize(grade.term).includes("final"))
              .map((grade) => grade.score)
              .filter((score): score is number => typeof score === "number");
            const scores = (
              finalScores.length
                ? finalScores
                : subjectGrades.map((grade) => grade.score)
            ).filter((score): score is number => typeof score === "number");

            return scores.length
              ? scores.reduce((sum, score) => sum + score, 0) / scores.length
              : null;
          })
          .filter((score): score is number => typeof score === "number");

        return {
          student,
          sex: sexOf(student),
          generalAverage: subjectAverages.length
            ? subjectAverages.reduce((sum, score) => sum + score, 0) /
              subjectAverages.length
            : null,
        };
      });

      return {
        grade: gradeLevel.value,
        label: gradeLevel.label,
        promoted: countLearners(
          learnerSummaries,
          (learner) =>
            typeof learner.generalAverage === "number" &&
            learner.generalAverage >= 75,
        ),
        conditional: { ...EMPTY_SF6_COUNT },
        retained: countLearners(
          learnerSummaries,
          (learner) =>
            typeof learner.generalAverage === "number" &&
            learner.generalAverage < 75,
        ),
        enrolled: countLearners(learnerSummaries, () => true),
        bands: SF6_PROGRESS_BANDS.map((band) =>
          countLearners(
            learnerSummaries,
            (learner) =>
              typeof learner.generalAverage === "number" &&
              band.matches(learner.generalAverage),
          ),
        ),
      };
    });

    const promotedByGrade = gradeReports.map((item) => item.promoted);
    const conditionalByGrade = gradeReports.map((item) => item.conditional);
    const retainedByGrade = gradeReports.map((item) => item.retained);
    const enrolledByGrade = gradeReports.map((item) => item.enrolled);
    const bandValues = SF6_PROGRESS_BANDS.map((_, bandIndex) =>
      gradeReports.map((item) => item.bands[bandIndex]),
    );

    return {
      gradeReports,
      promotedByGrade,
      conditionalByGrade,
      retainedByGrade,
      enrolledByGrade,
      bandValues,
    };
  }, [classes, grades, selectedClass, students]);

  const selectedTeacherProfile = selectedClass
    ? teacherById.get(selectedClass.teacher_id) || null
    : null;
  const sf6SchoolName =
    schoolProfile?.school_name ||
    selectedTeacherProfile?.school_name ||
    "Agusan del Sur National Science High School";
  const sf6SchoolId =
    schoolProfile?.school_id || selectedTeacherProfile?.school_id || "";
  const sf6Region =
    schoolProfile?.region || selectedTeacherProfile?.region || "";
  const sf6Division =
    schoolProfile?.division || selectedTeacherProfile?.division || "";
  const sf6District =
    schoolProfile?.district || selectedTeacherProfile?.district || "";
  const sf6SchoolHead =
    schoolProfile?.principal || selectedTeacherProfile?.principal || "";
  const sf6StatusRows = sf6Metrics
    ? [
        {
          label: "PROMOTED",
          values: sf6Metrics.promotedByGrade,
          total: addSf6Counts(sf6Metrics.promotedByGrade),
        },
        {
          label: "CONDITIONAL",
          values: sf6Metrics.conditionalByGrade,
          total: addSf6Counts(sf6Metrics.conditionalByGrade),
        },
        {
          label: "RETAINED",
          values: sf6Metrics.retainedByGrade,
          total: addSf6Counts(sf6Metrics.retainedByGrade),
        },
      ]
    : [];

  const sf8Learners = useMemo(
    () =>
      selectedClass
        ? students
            .filter((student) => student.class_id === selectedClass.id)
            .sort((first, second) =>
              studentName(first).localeCompare(studentName(second)),
            )
        : [],
    [selectedClass, students],
  );

  const exportSf8ToExcel = async () => {
    if (!selectedClass) {
      window.alert("Select an SF8 class record before exporting.");
      return;
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName =
        `SF8_${schoolYear || "School"}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}`
          .replace(/[<>:"/\\|?*]+/g, "_")
          .replace(/\s+/g, "_");

      await downloadSf8ExcelWorkbook(fileName, {
        schoolName:
          schoolProfile?.school_name ||
          selectedTeacherProfile?.school_name ||
          "Agusan del Sur National Science High School",
        schoolId:
          schoolProfile?.school_id || selectedTeacherProfile?.school_id || "",
        region: schoolProfile?.region || selectedTeacherProfile?.region || "",
        division:
          schoolProfile?.division || selectedTeacherProfile?.division || "",
        district:
          schoolProfile?.district || selectedTeacherProfile?.district || "",
        schoolYear,
        gradeLevel: selectedClass.grade_level || "",
        section: selectedClass.section || "",
        learners: sf8Learners,
      });
    } catch (error) {
      console.error("SF8 Excel export failed", error);
      window.alert(
        "Unable to export SF8 to Excel. Make sure SF8-admin.xlsx is inside public/templates.",
      );
    }
  };

  const exportSf10ToExcel = async () => {
    if (!selectedClass || !sf10Metrics) {
      window.alert("Select an SF10 class record before exporting.");
      return;
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName =
        `SF10_${schoolYear || "School"}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}_${
          sf10Metrics.learner ? studentName(sf10Metrics.learner) : "Learner"
        }`
          .replace(/[<>:"/\\|?*]+/g, "_")
          .replace(/\s+/g, "_");

      await downloadSf10ExcelWorkbook(fileName, {
        schoolName:
          schoolProfile?.school_name ||
          selectedTeacherProfile?.school_name ||
          "Agusan del Sur National Science High School",
        schoolId:
          schoolProfile?.school_id || selectedTeacherProfile?.school_id || "",
        region: schoolProfile?.region || selectedTeacherProfile?.region || "",
        division:
          schoolProfile?.division || selectedTeacherProfile?.division || "",
        district:
          schoolProfile?.district || selectedTeacherProfile?.district || "",
        schoolYear,
        gradeLevel: selectedClass.grade_level || "",
        section: selectedClass.section || "",
        learner: sf10Metrics.learner,
        teacherName: selectedTeacherName,
        schoolHead:
          schoolProfile?.principal ||
          selectedTeacherProfile?.principal ||
          "School Head",
        subjects: sf10Metrics.subjects,
        generalAverage: sf10Metrics.generalAverage,
      });
    } catch (error) {
      console.error("SF10 Excel export failed", error);
      window.alert(
        "Unable to export SF10 to Excel. Make sure SF10-admin.xlsx is inside public/templates.",
      );
    }
  };

  const exportSf1ToExcel = async () => {
    if (!selectedClass || !sf1Metrics) {
      window.alert("Select an SF1 class record before exporting.");
      return;
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName =
        `SF1_${schoolYear || "School"}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}`
          .replace(/[<>:"/\\|?*]+/g, "_")
          .replace(/\s+/g, "_");

      await downloadSf1ExcelWorkbook(fileName, {
        schoolName:
          schoolProfile?.school_name ||
          selectedTeacherProfile?.school_name ||
          "Agusan del Sur National Science High School",
        schoolId:
          schoolProfile?.school_id || selectedTeacherProfile?.school_id || "",
        region: schoolProfile?.region || selectedTeacherProfile?.region || "",
        division:
          schoolProfile?.division || selectedTeacherProfile?.division || "",
        schoolYear,
        gradeLevel: selectedClass.grade_level || "",
        section: selectedClass.section || "",
        teacherName: selectedTeacherName,
        schoolHead:
          schoolProfile?.principal ||
          selectedTeacherProfile?.principal ||
          "School Head",
        male: sf1Metrics.male,
        female: [...sf1Metrics.female, ...sf1Metrics.unspecified],
      });
    } catch (error) {
      console.error("SF1 Excel export failed", error);
      window.alert(
        "Unable to export SF1 to Excel. Make sure SF1-class-adviser.xls is inside public/templates.",
      );
    }
  };

  const exportSf6ToExcel = async () => {
    if (!selectedClass || !sf6Metrics) {
      window.alert("Select an SF6 class record before exporting.");
      return;
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName = `SF6_${schoolYear || "School"}_${sf6SchoolName}`
        .replace(/[<>:"/\\|?*]+/g, "_")
        .replace(/\s+/g, "_");

      await downloadSf6ExcelWorkbook(fileName, {
        schoolName: sf6SchoolName,
        schoolId: sf6SchoolId,
        region: sf6Region,
        division: sf6Division,
        district: sf6District,
        schoolYear,
        statusRows: sf6StatusRows,
        bandValues: sf6Metrics.bandValues,
        enrolledByGrade: sf6Metrics.enrolledByGrade,
        schoolHead: sf6SchoolHead,
      });
    } catch (error) {
      console.error("SF6 Excel export failed", error);
      window.alert("Unable to export SF6 to Excel.");
    }
  };

  const buildAdminSf5TemplateOptions = (): SF5TemplateOptions | null => {
    if (!selectedClass || !sf5Metrics) return null;

    const toTemplateRow = (
      learner: (typeof sf5Metrics.learners)[number],
    ): SF5TemplateRow => ({
      lrn: learner.lrn === "-" ? "" : learner.lrn,
      name: learner.name.toUpperCase(),
      average:
        learner.generalAverage > 0 ? Math.round(learner.generalAverage) : null,
      action: learner.action,
      failedAreas: learner.didNotMeet.toUpperCase(),
    });
    const maleRows = sf5Metrics.learners
      .filter((learner) => learner.sex === "male")
      .map(toTemplateRow);
    const femaleRows = sf5Metrics.learners
      .filter((learner) => learner.sex === "female")
      .map(toTemplateRow);
    const gradeLevel = selectedClass.grade_level?.trim() || "";

    return {
      region: schoolProfile?.region || selectedTeacherProfile?.region || "",
      division:
        schoolProfile?.division || selectedTeacherProfile?.division || "",
      schoolId:
        schoolProfile?.school_id || selectedTeacherProfile?.school_id || "",
      schoolYear: selectedClass.school_year || "",
      curriculum: "K to 12",
      schoolName:
        schoolProfile?.school_name ||
        selectedTeacherProfile?.school_name ||
        "Agusan del Sur National Science High School",
      gradeLevel:
        gradeLevel && !/^grade\b/i.test(gradeLevel)
          ? `Grade ${gradeLevel}`
          : gradeLevel,
      section: selectedClass.section || "",
      maleRows,
      femaleRows,
      summary: {
        promotedMale: sf5Metrics.promoted.male,
        promotedFemale: sf5Metrics.promoted.female,
        conditionalMale: 0,
        conditionalFemale: 0,
        retainedMale: sf5Metrics.retained.male,
        retainedFemale: sf5Metrics.retained.female,
      },
      progress: [...sf5Metrics.gradeBands].reverse().map((band) => ({
        label: band.label,
        male: band.male,
        female: band.female,
      })),
      adviser: selectedTeacherName,
      schoolHead:
        schoolProfile?.principal ||
        selectedTeacherProfile?.principal ||
        "School Head",
    };
  };

  const exportSf5ToExcel = async () => {
    const options = buildAdminSf5TemplateOptions();
    if (!selectedClass || !sf5Metrics || !options) {
      window.alert("Select an SF5 class record before exporting.");
      return;
    }

    try {
      await downloadSF5ExcelTemplate(
        `SF5_Promotion_${selectedClass.section || "class"}`,
        options,
      );
      return;
    } catch (templateError) {
      console.warn(
        "SF5 Excel template export failed; using the restored legacy exporter.",
        templateError,
      );
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName =
        `SF5_${schoolYear || "School"}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}`
          .replace(/[<>:"/\\|?*]+/g, "_")
          .replace(/\s+/g, "_");

      await downloadSf5ExcelWorkbook(fileName, {
        schoolName:
          schoolProfile?.school_name ||
          selectedTeacherProfile?.school_name ||
          "Agusan del Sur National Science High School",
        schoolId:
          schoolProfile?.school_id || selectedTeacherProfile?.school_id || "",
        region: schoolProfile?.region || selectedTeacherProfile?.region || "",
        division:
          schoolProfile?.division || selectedTeacherProfile?.division || "",
        district:
          schoolProfile?.district || selectedTeacherProfile?.district || "",
        schoolYear,
        gradeLevel: selectedClass.grade_level || "",
        section: selectedClass.section || "",
        curriculum: "K to 12",
        teacherName: selectedTeacherName,
        schoolHead:
          schoolProfile?.principal ||
          selectedTeacherProfile?.principal ||
          "School Head",
        learners: sf5Metrics.learners,
        maleTotal: sf5Metrics.maleTotal,
        femaleTotal: sf5Metrics.femaleTotal,
        promoted: sf5Metrics.promoted,
        conditional: sf5Metrics.conditional,
        retained: sf5Metrics.retained,
        gradeBands: sf5Metrics.gradeBands,
      });
    } catch (error) {
      console.error("SF5 Excel export failed", error);
      window.alert("Unable to export SF5 to Excel.");
    }
  };

  const exportSf9NewToExcel = async () => {
    if (!selectedClass || !sf9Metrics) {
      window.alert("Select an SF9 New class record before exporting.");
      return;
    }

    try {
      const schoolYear = selectedClass.school_year || "";
      const fileName =
        `SF9_New_${schoolYear || "School"}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}_${
          sf9Metrics.learner ? studentName(sf9Metrics.learner) : "Learner"
        }`
          .replace(/[<>:"/\\|?*]+/g, "_")
          .replace(/\s+/g, "_");

      await downloadSf9NewExcelWorkbook(fileName, {
        schoolName:
          schoolProfile?.school_name ||
          selectedTeacherProfile?.school_name ||
          "Agusan del Sur National Science High School",
        region: schoolProfile?.region || selectedTeacherProfile?.region || "",
        division:
          schoolProfile?.division || selectedTeacherProfile?.division || "",
        district:
          schoolProfile?.district || selectedTeacherProfile?.district || "",
        schoolYear,
        gradeLevel: selectedClass.grade_level || "",
        section: selectedClass.section || "",
        municipality:
          schoolProfile?.city_municipality_province ||
          schoolProfile?.municipality ||
          selectedTeacherProfile?.municipality ||
          "Prosperidad, Agusan del Sur",
        track: selectedClass.track_shs || "",
        learner: sf9Metrics.learner,
        teacherName: selectedTeacherName,
        schoolHead:
          schoolProfile?.principal ||
          selectedTeacherProfile?.principal ||
          "School Head",
        subjects: sf9Metrics.subjects,
        generalAverage: sf9Metrics.generalAverage,
      });
    } catch (error) {
      console.error("SF9 New Excel export failed", error);
      window.alert(
        "Unable to export SF9 New to Excel. Make sure the SF9 template is inside public/templates.",
      );
    }
  };

  const subjectCount = useMemo(() => {
    if (!selectedClass) return 0;

    const subjects = classes
      .filter(
        (item) =>
          item.grade_level === selectedClass.grade_level &&
          item.section === selectedClass.section,
      )
      .map((item) => normalize(item.subject))
      .filter(Boolean);

    return new Set(subjects).size;
  }, [classes, selectedClass]);

  const increaseZoom = () => setZoom((current) => Math.min(current + 10, 200));
  const decreaseZoom = () => setZoom((current) => Math.max(current - 10, 50));

  const resetPreview = () => {
    setPaperSize("long");
    setSogTerm("1");
    setSf5ReportLength("short");
    setSf1ReportLength("full");
    setSf2Month(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    setZoom(90);
  };

  const getPreviewClone = () => {
    const preview = document.getElementById("admin-school-form-paper");
    if (!preview) return null;

    const clone = preview.cloneNode(true) as HTMLElement;

    const applyComputedStyles = (source: Element, target: Element) => {
      const computed = window.getComputedStyle(source);
      const declarations: string[] = [];

      for (let index = 0; index < computed.length; index += 1) {
        const property = computed.item(index);
        declarations.push(
          `${property}: ${computed.getPropertyValue(property)};`,
        );
      }

      const existingStyle = target.getAttribute("style") || "";
      target.setAttribute("style", `${existingStyle};${declarations.join("")}`);

      const sourceChildren = Array.from(source.children);
      const targetChildren = Array.from(target.children);
      sourceChildren.forEach((child, index) => {
        const targetChild = targetChildren[index];
        if (targetChild) applyComputedStyles(child, targetChild);
      });
    };

    applyComputedStyles(preview, clone);
    clone.style.transform = "none";
    clone.style.boxShadow = "none";
    clone.style.margin = "0 auto";

    return clone;
  };

  const printSchoolForm = async () => {
    const previewClone = getPreviewClone();
    if (!previewClone || !selectedClass) return;

    if (!isSf5Preview) {
      const printWindow = window.open("", "_blank", "width=1100,height=900");
      if (!printWindow) {
        window.alert(
          "Please allow pop-ups to print or save this school form as PDF.",
        );
        return;
      }

      const legacyPageSize = isLandscapePreview
        ? paperSize === "long"
          ? "13in 8.5in"
          : paperSize === "letter"
            ? "11in 8.5in"
            : "11.69in 8.27in"
        : paperSize === "long"
          ? "8.5in 13in"
          : paperSize === "letter"
            ? "letter"
            : "A4";

      printWindow.document.write(`<!doctype html>
        <html>
          <head>
            <meta charset="utf-8" />
            <title>${selectedForm}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}</title>
            <style>
              @page { size: ${legacyPageSize}; margin: 0; }
              html, body { margin: 0; padding: 0; background: white; }
              body { display: flex; justify-content: center; align-items: flex-start; }
            </style>
          </head>
          <body>${previewClone.outerHTML}</body>
        </html>`);
      printWindow.document.close();
      printWindow.focus();
      window.setTimeout(() => printWindow.print(), 300);
      return;
    }

    const pageSize = "13in 8.5in";

    const printRootId = "admin-school-form-print-root";
    const printStyleId = "admin-school-form-print-style";
    document.getElementById(printRootId)?.remove();
    document.getElementById(printStyleId)?.remove();

    const printRoot = document.createElement("div");
    const printStyle = document.createElement("style");
    const originalTitle = document.title;
    printRoot.id = printRootId;
    previewClone.setAttribute("data-admin-print-target", "true");
    printRoot.appendChild(previewClone);

    printStyle.id = printStyleId;
    printStyle.textContent = `
      @page { size: ${pageSize}; margin: 0; }

      @media screen {
        #${printRootId} {
          position: fixed !important;
          left: -100000px !important;
          top: 0 !important;
          visibility: hidden !important;
          pointer-events: none !important;
        }
      }

      @media print {
        html, body {
          width: 100% !important;
          height: 100% !important;
          margin: 0 !important;
          padding: 0 !important;
          overflow: visible !important;
          background: #fff !important;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }

        body > *:not(#${printRootId}) { display: none !important; }

        #${printRootId},
        [data-admin-print-target="true"] {
          display: block !important;
          position: static !important;
          box-sizing: border-box !important;
          width: 100% !important;
          height: 100% !important;
          max-width: none !important;
          max-height: none !important;
          margin: 0 !important;
          padding: 0 !important;
          overflow: hidden !important;
          transform: none !important;
          box-shadow: none !important;
          break-inside: avoid-page !important;
          page-break-inside: avoid !important;
          background: #fff !important;
        }

        #${printRootId} table { border-collapse: collapse !important; }
      }
    `;

    const cleanup = () => {
      printRoot.remove();
      printStyle.remove();
      document.title = originalTitle;
    };

    document.head.appendChild(printStyle);
    document.body.appendChild(printRoot);
    document.title = `${selectedForm}_${selectedClass.grade_level || "Grade"}_${selectedClass.section || "Section"}`;

    try {
      const images = Array.from(previewClone.querySelectorAll("img"));
      await Promise.all([
        document.fonts?.ready ?? Promise.resolve(),
        ...images.map(
          (image) =>
            image.decode?.().catch(() => undefined) ?? Promise.resolve(),
        ),
      ]);
      await new Promise<void>((resolve) =>
        window.requestAnimationFrame(() =>
          window.requestAnimationFrame(() => resolve()),
        ),
      );
      window.print();
    } catch (error) {
      console.error("School form print failed", error);
      window.alert("Unable to open the print dialog.");
    } finally {
      cleanup();
    }
  };

  const copySchoolFormToWord = async () => {
    if (!selectedClass) return;

    try {
      const teacherName =
        selectedClass.teacher_name ||
        teacherById.get(selectedClass.teacher_id)?.full_name ||
        "-";

      const gradeLevel = selectedClass.grade_level || "Grade Level";
      const section = selectedClass.section || "-";
      const schoolYear = selectedClass.school_year || "-";
      const subject = selectedClass.subject || "-";
      const title = formTitle(selectedForm);

      const inchesToTwip = (inches: number) => Math.round(inches * 1440);

      const effectivePaperSize = selectedForm === "SF5" ? "long" : paperSize;

      const pageDimensions =
        effectivePaperSize === "long"
          ? {
              width: inchesToTwip(8.5),
              height: inchesToTwip(13),
            }
          : effectivePaperSize === "letter"
            ? {
                width: inchesToTwip(8.5),
                height: inchesToTwip(11),
              }
            : {
                width: inchesToTwip(8.27),
                height: inchesToTwip(11.69),
              };

      let logoBytes: Uint8Array | null = null;
      let logoType: "png" | "jpg" = "jpg";

      try {
        const logoResponse = await fetch(deped);

        if (logoResponse.ok) {
          const contentType = logoResponse.headers.get("content-type") || "";

          logoType = contentType.includes("png") ? "png" : "jpg";

          logoBytes = new Uint8Array(await logoResponse.arrayBuffer());
        }
      } catch (logoError) {
        console.warn("Unable to add logo to Word file.", logoError);
      }

      const noBorder = {
        style: BorderStyle.NONE,
        size: 0,
        color: "FFFFFF",
      };

      const thinBorder = {
        style: BorderStyle.SINGLE,
        size: 4,
        color: "DED5CD",
      };

      const headerCellBorders = {
        top: noBorder,
        bottom: noBorder,
        left: noBorder,
        right: noBorder,
      };

      // GSA uses its own Word layout so the downloaded document matches
      // the landscape GSA bond-paper preview instead of the generic form.
      if (selectedForm === "GSA" && gsaMetrics) {
        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "000000",
        };

        const gsaBorders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const makeGsaCell = ({
          text,
          width,
          bold = false,
          fontSize = 12,
          alignment = AlignmentType.CENTER,
          fill = "FFFFFF",
          rowSpan,
          columnSpan,
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          rowSpan?: number;
          columnSpan?: number;
        }) =>
          new TableCell({
            width: {
              size: width,
              type: WidthType.PERCENTAGE,
            },
            rowSpan,
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: {
              type: ShadingType.CLEAR,
              fill,
            },
            borders: gsaBorders,
            margins: {
              top: 45,
              bottom: 45,
              left: 35,
              right: 35,
            },
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize,
                    color: "000000",
                  }),
                ],
              }),
            ],
          });

        const makeValueCell = (
          value: string | number,
          width = 5.2,
          bold = false,
        ) =>
          makeGsaCell({
            text: String(value),
            width,
            bold,
            fontSize: 12,
          });

        const gsaHeaderTable = new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 100 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children:
                        logoBytes !== null
                          ? [
                              new ImageRun({
                                data: logoBytes,
                                type: logoType,
                                transformation: {
                                  width: 92,
                                  height: 54,
                                },
                              }),
                            ]
                          : [],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 56, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 0 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Republic of the Philippines",
                          font: "Arial",
                          size: 11,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "DEPARTMENT OF EDUCATION",
                          font: "Arial",
                          bold: true,
                          size: 16,
                          color: "0038A8",
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: schoolProfile?.region
                            ? `Region ${schoolProfile.region}`
                            : "Region XIII — CARAGA",
                          font: "Arial",
                          size: 11,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: (
                            schoolProfile?.school_name ||
                            "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL"
                          ).toUpperCase(),
                          font: "Arial",
                          bold: true,
                          size: 12,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
          ],
        });

        const gsaInfoTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 70, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Teacher:  ",
                          font: "Arial",
                          bold: true,
                          size: 12,
                        }),
                        new TextRun({
                          text: teacherName,
                          font: "Arial",
                          size: 12,
                          color: "0038A8",
                          underline: {},
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Term:  ",
                          font: "Arial",
                          bold: true,
                          size: 12,
                        }),
                        new TextRun({
                          text: "1st",
                          font: "Arial",
                          size: 12,
                          underline: {},
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        });

        const gsaTableRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeGsaCell({
                text: "Grade Level/\nSubjects",
                width: 10,
                bold: true,
                rowSpan: 2,
              }),
              makeGsaCell({
                text: "STRAND",
                width: 6,
                bold: true,
                rowSpan: 2,
              }),
              makeGsaCell({
                text: "SECTION",
                width: 6,
                bold: true,
                rowSpan: 2,
              }),
              makeGsaCell({
                text: "REGISTERED LEARNERS",
                width: 15.6,
                bold: true,
                columnSpan: 3,
              }),
              makeGsaCell({
                text: "LEARNERS WITH 75% AND ABOVE",
                width: 15.6,
                bold: true,
                columnSpan: 3,
              }),
              makeGsaCell({
                text: "TOTAL GRADES",
                width: 15.6,
                bold: true,
                columnSpan: 3,
              }),
              makeGsaCell({
                text: "AVERAGE GRADE",
                width: 15.6,
                bold: true,
                columnSpan: 3,
              }),
              makeGsaCell({
                text: "% OF PROFICIENCY (75%+)",
                width: 15.6,
                bold: true,
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: Array.from({ length: 5 }).flatMap(() =>
              ["MALE", "FEMALE", "TOTAL"].map((label) =>
                makeGsaCell({
                  text: label,
                  width: 5.2,
                  bold: true,
                  fontSize: 10,
                  fill: "F2F2F2",
                }),
              ),
            ),
          }),
          new TableRow({
            cantSplit: true,
            height: { value: 330, rule: HeightRule.ATLEAST },
            children: [
              makeGsaCell({
                text: selectedClass.subject || "Learning Area",
                width: 10,
                alignment: AlignmentType.LEFT,
                fontSize: 12,
              }),
              makeGsaCell({ text: "N/A", width: 6, fontSize: 12 }),
              makeGsaCell({
                text: selectedClass.section || "-",
                width: 6,
                fontSize: 12,
              }),
              makeValueCell(gsaMetrics.male.registered),
              makeValueCell(gsaMetrics.female.registered),
              makeValueCell(gsaMetrics.total.registered, 5.2, true),
              makeValueCell(gsaMetrics.male.passed),
              makeValueCell(gsaMetrics.female.passed),
              makeValueCell(gsaMetrics.total.passed, 5.2, true),
              makeValueCell(gsaMetrics.male.totalGrades.toFixed(0)),
              makeValueCell(gsaMetrics.female.totalGrades.toFixed(0)),
              makeValueCell(gsaMetrics.total.totalGrades.toFixed(0), 5.2, true),
              makeValueCell(
                gsaMetrics.male.average
                  ? gsaMetrics.male.average.toFixed(2)
                  : "0",
              ),
              makeValueCell(
                gsaMetrics.female.average
                  ? gsaMetrics.female.average.toFixed(2)
                  : "0",
              ),
              makeValueCell(
                gsaMetrics.total.average
                  ? gsaMetrics.total.average.toFixed(2)
                  : "0",
                5.2,
                true,
              ),
              makeValueCell(`${gsaMetrics.male.proficiency}%`),
              makeValueCell(`${gsaMetrics.female.proficiency}%`),
              makeValueCell(`${gsaMetrics.total.proficiency}%`, 5.2, true),
            ],
          }),
          ...Array.from({ length: 2 }).map(
            () =>
              new TableRow({
                cantSplit: true,
                height: { value: 430, rule: HeightRule.ATLEAST },
                children: [
                  makeGsaCell({ text: "", width: 10 }),
                  makeGsaCell({ text: "", width: 6 }),
                  makeGsaCell({ text: "", width: 6 }),
                  ...Array.from({ length: 15 }).map(() =>
                    makeGsaCell({ text: "", width: 5.2 }),
                  ),
                ],
              }),
          ),
        ];

        const gsaResultsTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: gsaTableRows,
        });

        const signatureItems = [
          ["Prepared by:", teacherName, "Teacher I"],
          ["Checked by:", "CHECKER NAME", "Master Teacher I"],
          ["Noted:", "NOTED BY NAME", "SHS Dept. Head"],
          ["Approved:", "PRINCIPAL NAME", "School Principal"],
        ];

        const gsaSignatureTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: signatureItems.map(
                ([label, name, position]) =>
                  new TableCell({
                    width: { size: 25, type: WidthType.PERCENTAGE },
                    borders: headerCellBorders,
                    margins: { top: 0, bottom: 0, left: 80, right: 80 },
                    children: [
                      new Paragraph({
                        spacing: { before: 0, after: 260 },
                        children: [
                          new TextRun({
                            text: label,
                            font: "Arial",
                            bold: true,
                            size: 11,
                          }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        border: {
                          top: {
                            style: BorderStyle.SINGLE,
                            size: 6,
                            color: "000000",
                          },
                        },
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({
                            text: name.toUpperCase(),
                            font: "Arial",
                            bold: true,
                            size: 12,
                          }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({
                            text: position,
                            font: "Arial",
                            size: 10,
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
          ],
        });

        const gsaDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: {
                  font: "Arial",
                  size: 12,
                  color: "000000",
                },
                paragraph: {
                  spacing: { before: 0, after: 0, line: 220 },
                },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    width: pageDimensions.height,
                    height: pageDimensions.width,
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: inchesToTwip(0.32),
                    bottom: inchesToTwip(0.32),
                    left: inchesToTwip(0.35),
                    right: inchesToTwip(0.35),
                  },
                },
              },
              children: [
                gsaHeaderTable,
                new Paragraph({
                  border: {
                    bottom: {
                      style: BorderStyle.SINGLE,
                      size: 8,
                      color: "000000",
                    },
                  },
                  spacing: { before: 20, after: 90 },
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({
                      text: "REPORT ON GENERAL SCHOLASTIC APTITUDE OF LEARNERS",
                      font: "Arial",
                      bold: true,
                      size: 18,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 10 },
                  children: [
                    new TextRun({
                      text: "PER TERM BY LEARNING AREA",
                      font: "Arial",
                      bold: true,
                      size: 18,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 100 },
                  children: [
                    new TextRun({
                      text: `School Year: ${schoolYear}`,
                      font: "Arial",
                      size: 11,
                    }),
                  ],
                }),
                gsaInfoTable,
                new Paragraph({ spacing: { before: 0, after: 35 } }),
                gsaResultsTable,
                new Paragraph({ spacing: { before: 0, after: 170 } }),
                gsaSignatureTable,
              ],
            },
          ],
        });

        const gsaWordBlob = await Packer.toBlob(gsaDocument);
        const gsaFileName =
          `GSA_${gradeLevel}_${section}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const gsaUrl = URL.createObjectURL(gsaWordBlob);
        const gsaAnchor = document.createElement("a");
        gsaAnchor.href = gsaUrl;
        gsaAnchor.download = gsaFileName;
        document.body.appendChild(gsaAnchor);
        gsaAnchor.click();
        gsaAnchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(gsaUrl), 1000);
        return;
      }

      // MPS uses its own Word layout so the downloaded document matches
      // the landscape MPS bond-paper preview instead of the generic form.
      if (selectedForm === "MPS" && mpsMetrics) {
        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "000000",
        };

        const mpsBorders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const makeMpsCell = ({
          text,
          width,
          bold = false,
          fontSize = 10,
          alignment = AlignmentType.CENTER,
          fill = "FFFFFF",
          rowSpan,
          columnSpan,
          height,
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          rowSpan?: number;
          columnSpan?: number;
          height?: number;
        }) =>
          new TableCell({
            width: {
              size: width,
              type: WidthType.PERCENTAGE,
            },
            rowSpan,
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: {
              type: ShadingType.CLEAR,
              fill,
            },
            borders: mpsBorders,
            margins: {
              top: height ? 55 : 40,
              bottom: height ? 55 : 40,
              left: 28,
              right: 28,
            },
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize,
                    color: "000000",
                  }),
                ],
              }),
            ],
          });

        const mpsHeaderTable = new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 100 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children:
                        logoBytes !== null
                          ? [
                              new ImageRun({
                                data: logoBytes,
                                type: logoType,
                                transformation: {
                                  width: 92,
                                  height: 54,
                                },
                              }),
                            ]
                          : [],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 56, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 0 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Republic of the Philippines",
                          font: "Arial",
                          size: 10,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "DEPARTMENT OF EDUCATION",
                          font: "Arial",
                          bold: true,
                          size: 15,
                          color: "0038A8",
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: schoolProfile?.region
                            ? `Region ${schoolProfile.region}`
                            : "Region XIII - CARAGA",
                          font: "Arial",
                          size: 10,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: (
                            schoolProfile?.school_name ||
                            "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL"
                          ).toUpperCase(),
                          font: "Arial",
                          bold: true,
                          size: 11,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
          ],
        });

        const departmentName =
          Number(
            String(selectedClass.grade_level || "").match(/\d+/)?.[0] || 0,
          ) >= 11
            ? "SENIOR HIGH SCHOOL"
            : "JUNIOR HIGH SCHOOL";

        const mpsInfoTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 33.33, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Grade:  ",
                          font: "Arial",
                          bold: true,
                          size: 11,
                        }),
                        new TextRun({
                          text: gradeLevel,
                          font: "Arial",
                          size: 11,
                          underline: {},
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 33.33, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Department:  ",
                          font: "Arial",
                          bold: true,
                          size: 11,
                        }),
                        new TextRun({
                          text: departmentName,
                          font: "Arial",
                          size: 11,
                          underline: {},
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 33.34, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Term:  ",
                          font: "Arial",
                          bold: true,
                          size: 11,
                        }),
                        new TextRun({
                          text: "FIRST",
                          font: "Arial",
                          size: 11,
                          underline: {},
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        });

        const smallWidth = 4.75;
        const mpsTableRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeMpsCell({
                text: "Grade Level/\nSubjects",
                width: 12,
                bold: true,
                columnSpan: 2,
              }),
              makeMpsCell({
                text: "SHS/ABM/\nSECTION",
                width: 12,
                bold: true,
                columnSpan: 2,
              }),
              makeMpsCell({
                text: "No. of Learners",
                width: smallWidth * 3,
                bold: true,
                columnSpan: 3,
              }),
              makeMpsCell({
                text: "Learners with scores 75% of HTS and above",
                width: smallWidth * 3,
                bold: true,
                columnSpan: 3,
              }),
              makeMpsCell({
                text: "MEAN score of Learners",
                width: smallWidth * 3,
                bold: true,
                columnSpan: 3,
              }),
              makeMpsCell({
                text: "Mean % of Learners Scores",
                width: smallWidth * 3,
                bold: true,
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeMpsCell({
                text: "Subjects Handled",
                width: 7,
                bold: true,
                fontSize: 9,
              }),
              makeMpsCell({ text: "HTS", width: 5, bold: true, fontSize: 9 }),
              makeMpsCell({
                text: "SHS/ABM",
                width: 6,
                bold: true,
                fontSize: 9,
              }),
              makeMpsCell({
                text: "SECTION",
                width: 6,
                bold: true,
                fontSize: 9,
              }),
              ...Array.from({ length: 4 }).flatMap(() =>
                ["MALE", "FEMALE", "TOTAL"].map((label) =>
                  makeMpsCell({
                    text: label,
                    width: smallWidth,
                    bold: true,
                    fontSize: 8,
                    fill: "F2F2F2",
                  }),
                ),
              ),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              makeMpsCell({ text: "", width: 61.5, columnSpan: 10 }),
              makeMpsCell({
                text: "Formula: Sum of Scores / No. of Learners",
                width: smallWidth * 3,
                columnSpan: 3,
                fontSize: 7,
              }),
              makeMpsCell({
                text: "Formula: (Mean Score / HTS) x 100",
                width: smallWidth * 3,
                columnSpan: 3,
                fontSize: 7,
              }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            height: { value: 360, rule: HeightRule.ATLEAST },
            children: [
              makeMpsCell({
                text: selectedClass.subject || "Learning Area",
                width: 7,
                alignment: AlignmentType.LEFT,
                fontSize: 10,
              }),
              makeMpsCell({
                text: String(mpsMetrics.highestTestScore),
                width: 5,
                fontSize: 10,
              }),
              makeMpsCell({ text: "N/A", width: 6, fontSize: 10 }),
              makeMpsCell({
                text: selectedClass.section || "-",
                width: 6,
                fontSize: 10,
              }),
              makeMpsCell({
                text: String(mpsMetrics.male.registered),
                width: smallWidth,
              }),
              makeMpsCell({
                text: String(mpsMetrics.female.registered),
                width: smallWidth,
              }),
              makeMpsCell({
                text: String(mpsMetrics.total.registered),
                width: smallWidth,
                bold: true,
              }),
              makeMpsCell({
                text: String(mpsMetrics.male.passed),
                width: smallWidth,
              }),
              makeMpsCell({
                text: String(mpsMetrics.female.passed),
                width: smallWidth,
              }),
              makeMpsCell({
                text: String(mpsMetrics.total.passed),
                width: smallWidth,
                bold: true,
              }),
              makeMpsCell({
                text: mpsMetrics.male.meanScore
                  ? mpsMetrics.male.meanScore.toFixed(2)
                  : "0",
                width: smallWidth,
              }),
              makeMpsCell({
                text: mpsMetrics.female.meanScore
                  ? mpsMetrics.female.meanScore.toFixed(2)
                  : "0",
                width: smallWidth,
              }),
              makeMpsCell({
                text: mpsMetrics.total.meanScore
                  ? mpsMetrics.total.meanScore.toFixed(2)
                  : "0",
                width: smallWidth,
                bold: true,
              }),
              makeMpsCell({
                text: mpsMetrics.male.meanPercentage.toFixed(2),
                width: smallWidth,
              }),
              makeMpsCell({
                text: mpsMetrics.female.meanPercentage.toFixed(2),
                width: smallWidth,
              }),
              makeMpsCell({
                text: mpsMetrics.total.meanPercentage.toFixed(2),
                width: smallWidth,
                bold: true,
              }),
            ],
          }),
          ...Array.from({ length: 4 }).map(
            () =>
              new TableRow({
                cantSplit: true,
                height: { value: 430, rule: HeightRule.ATLEAST },
                children: [
                  makeMpsCell({ text: "", width: 7 }),
                  makeMpsCell({ text: "", width: 5 }),
                  makeMpsCell({ text: "", width: 6 }),
                  makeMpsCell({ text: "", width: 6 }),
                  ...Array.from({ length: 12 }).map(() =>
                    makeMpsCell({ text: "", width: smallWidth }),
                  ),
                ],
              }),
          ),
        ];

        const mpsResultsTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: mpsTableRows,
        });

        const mpsSignatureItems = [
          ["Prepared by:", teacherName, "Teacher I"],
          ["Checked by:", "CHECKER NAME", "Master Teacher I"],
          ["Noted:", "NOTED BY NAME", "SHS Dept. Head"],
          ["Approved:", "PRINCIPAL NAME", "School Principal"],
        ];

        const mpsSignatureTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: {
            top: noBorder,
            bottom: noBorder,
            left: noBorder,
            right: noBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder,
          },
          rows: [
            new TableRow({
              cantSplit: true,
              children: mpsSignatureItems.map(
                ([label, name, position]) =>
                  new TableCell({
                    width: { size: 25, type: WidthType.PERCENTAGE },
                    borders: headerCellBorders,
                    margins: { top: 0, bottom: 0, left: 80, right: 80 },
                    children: [
                      new Paragraph({
                        spacing: { before: 0, after: 260 },
                        children: [
                          new TextRun({
                            text: label,
                            font: "Arial",
                            bold: true,
                            size: 10,
                          }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        border: {
                          top: {
                            style: BorderStyle.SINGLE,
                            size: 6,
                            color: "000000",
                          },
                        },
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({
                            text: name.toUpperCase(),
                            font: "Arial",
                            bold: true,
                            size: 11,
                          }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({
                            text: position,
                            font: "Arial",
                            size: 9,
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
          ],
        });

        const mpsDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: {
                  font: "Arial",
                  size: 11,
                  color: "000000",
                },
                paragraph: {
                  spacing: { before: 0, after: 0, line: 210 },
                },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    width: pageDimensions.height,
                    height: pageDimensions.width,
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: inchesToTwip(0.3),
                    bottom: inchesToTwip(0.3),
                    left: inchesToTwip(0.35),
                    right: inchesToTwip(0.35),
                  },
                },
              },
              children: [
                mpsHeaderTable,
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 90, after: 0 },
                  children: [
                    new TextRun({
                      text: "REPORT ON MEAN PERCENTAGE SCORE OF LEARNERS IN TERM EXAM",
                      font: "Arial",
                      bold: true,
                      size: 17,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 10 },
                  children: [
                    new TextRun({
                      text: "PER TERM BY LEARNING AREA",
                      font: "Arial",
                      bold: true,
                      size: 17,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 90 },
                  children: [
                    new TextRun({
                      text: `School Year: ${schoolYear}`,
                      font: "Arial",
                      size: 10,
                    }),
                  ],
                }),
                mpsInfoTable,
                new Paragraph({ spacing: { before: 0, after: 35 } }),
                mpsResultsTable,
                new Paragraph({ spacing: { before: 0, after: 170 } }),
                mpsSignatureTable,
              ],
            },
          ],
        });

        const mpsWordBlob = await Packer.toBlob(mpsDocument);
        const mpsFileName =
          `MPS_${gradeLevel}_${section}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const mpsUrl = URL.createObjectURL(mpsWordBlob);
        const mpsAnchor = document.createElement("a");
        mpsAnchor.href = mpsUrl;
        mpsAnchor.download = mpsFileName;
        document.body.appendChild(mpsAnchor);
        mpsAnchor.click();
        mpsAnchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(mpsUrl), 1000);
        return;
      }

      // SOG uses its own Word layout so the downloaded document matches
      // the landscape Summary of Grades bond-paper preview.
      if (selectedForm === "SOG" && sogMetrics) {
        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "000000",
        };

        const sogBorders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const borderlessTable = {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        };

        const sogBlue = "0B5A96";
        const sogGreen = "0A7A4B";
        const sogLightBlue = "EAF2F8";
        const sogLightGreen = "E8F7F0";
        const sogSubjects = sogMetrics.subjects.length
          ? sogMetrics.subjects
          : ["Learning Area"];
        const sogSubjectWidth = 57 / sogSubjects.length;
        const totalSogColumns = sogSubjects.length + 6;

        const makeSogCell = ({
          text,
          width,
          bold = false,
          fontSize = 8,
          alignment = AlignmentType.CENTER,
          fill = "FFFFFF",
          color = "000000",
          columnSpan,
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          color?: string;
          columnSpan?: number;
        }) =>
          new TableCell({
            width: {
              size: width,
              type: WidthType.PERCENTAGE,
            },
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: {
              type: ShadingType.CLEAR,
              fill,
            },
            borders: sogBorders,
            margins: {
              top: 40,
              bottom: 40,
              left: 35,
              right: 35,
            },
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize,
                    color,
                  }),
                ],
              }),
            ],
          });

        const sogRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeSogCell({
                text: "#",
                width: 3,
                bold: true,
                fontSize: 7,
                fill: sogBlue,
                color: "FFFFFF",
              }),
              makeSogCell({
                text: "LEARNER'S NAME",
                width: 13,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.LEFT,
                fill: sogBlue,
                color: "FFFFFF",
              }),
              makeSogCell({
                text: "LRN",
                width: 9,
                bold: true,
                fontSize: 7,
                fill: sogBlue,
                color: "FFFFFF",
              }),
              ...sogSubjects.map((subjectLabel) =>
                makeSogCell({
                  text: subjectLabel,
                  width: sogSubjectWidth,
                  bold: true,
                  fontSize: sogSubjects.length > 8 ? 5 : 6,
                  fill: sogBlue,
                  color: "FFFFFF",
                }),
              ),
              makeSogCell({
                text: "TERM AVG",
                width: 6,
                bold: true,
                fontSize: 6,
                fill: sogGreen,
                color: "FFFFFF",
              }),
              makeSogCell({
                text: "RANK",
                width: 4,
                bold: true,
                fontSize: 6,
                fill: sogBlue,
                color: "FFFFFF",
              }),
              makeSogCell({
                text: "DESCRIPTOR",
                width: 8,
                bold: true,
                fontSize: 6,
                fill: sogBlue,
                color: "FFFFFF",
              }),
            ],
          }),
        ];

        let sogRowNumber = 0;
        const sogGroups = [
          { key: "male", label: "MALE" },
          { key: "female", label: "FEMALE" },
          { key: "unspecified", label: "UNSPECIFIED" },
        ] as const;

        sogGroups.forEach((group) => {
          const groupLearners = sogMetrics.learners.filter(
            (learner) => learner.sexGroup === group.key,
          );

          if (groupLearners.length === 0) return;

          sogRows.push(
            new TableRow({
              cantSplit: true,
              children: [
                makeSogCell({
                  text: group.label,
                  width: 100,
                  columnSpan: totalSogColumns,
                  bold: true,
                  fontSize: 8,
                  alignment: AlignmentType.LEFT,
                  fill: sogLightBlue,
                }),
              ],
            }),
          );

          groupLearners.forEach((learner) => {
            sogRowNumber += 1;

            const scoreBySubject = new Map(
              learner.subjectScores.map((item) => [
                normalize(item.subject),
                item.score,
              ]),
            );

            sogRows.push(
              new TableRow({
                cantSplit: true,
                height: { value: 260, rule: HeightRule.ATLEAST },
                children: [
                  makeSogCell({
                    text: String(sogRowNumber),
                    width: 3,
                    fontSize: 7,
                  }),
                  makeSogCell({
                    text: learner.name.toLowerCase(),
                    width: 13,
                    bold: true,
                    fontSize: 7,
                    alignment: AlignmentType.LEFT,
                  }),
                  makeSogCell({
                    text: learner.lrn,
                    width: 9,
                    fontSize: 7,
                  }),
                  ...sogSubjects.map((subjectLabel) => {
                    const score = scoreBySubject.get(normalize(subjectLabel));
                    return makeSogCell({
                      text:
                        typeof score === "number"
                          ? String(Math.round(score))
                          : "",
                      width: sogSubjectWidth,
                      fontSize: 7,
                    });
                  }),
                  makeSogCell({
                    text:
                      typeof learner.termAverage === "number"
                        ? String(Math.round(learner.termAverage))
                        : "",
                    width: 6,
                    bold: true,
                    fontSize: 7,
                    fill: sogLightGreen,
                  }),
                  makeSogCell({
                    text: learner.rank ? String(learner.rank) : "",
                    width: 4,
                    bold: true,
                    fontSize: 7,
                  }),
                  makeSogCell({
                    text: learner.descriptor,
                    width: 8,
                    fontSize: 6,
                  }),
                ],
              }),
            );
          });
        });

        if (sogMetrics.learners.length === 0) {
          sogRows.push(
            new TableRow({
              cantSplit: true,
              height: { value: 520, rule: HeightRule.ATLEAST },
              children: [
                makeSogCell({
                  text: "No learner records found for this grade and section.",
                  width: 100,
                  columnSpan: totalSogColumns,
                  fontSize: 8,
                }),
              ],
            }),
          );
        }

        const sogResultsTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: sogRows,
        });

        const makeSogSignatureCell = (
          label: string,
          name: string,
          role: string,
        ) =>
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: headerCellBorders,
            verticalAlign: VerticalAlign.BOTTOM,
            margins: { top: 0, bottom: 0, left: 0, right: 180 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 150 },
                children: [
                  new TextRun({
                    text: label,
                    font: "Arial",
                    bold: true,
                    size: 8,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                border: {
                  bottom: {
                    style: BorderStyle.SINGLE,
                    size: 6,
                    color: "000000",
                  },
                },
                spacing: { before: 0, after: 18 },
                children: [
                  new TextRun({
                    text: name,
                    font: "Arial",
                    bold: Boolean(name.trim()),
                    size: 8,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: role,
                    font: "Arial",
                    size: 7,
                  }),
                ],
              }),
            ],
          });

        const sogSignatureTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeSogSignatureCell(
                  "Prepared by:",
                  teacherName,
                  "Class Adviser",
                ),
                makeSogSignatureCell("Noted by:", "", "School Head"),
              ],
            }),
          ],
        });

        const sogDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: {
                  font: "Arial",
                  size: 8,
                  color: "000000",
                },
                paragraph: {
                  spacing: { before: 0, after: 0, line: 180 },
                },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    width: pageDimensions.height,
                    height: pageDimensions.width,
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: inchesToTwip(0.2),
                    bottom: inchesToTwip(0.2),
                    left: inchesToTwip(0.2),
                    right: inchesToTwip(0.2),
                  },
                },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 90, after: 10 },
                  children: [
                    new TextRun({
                      text: `SUMMARY OF GRADES - ${sogMetrics.termLabel}`,
                      font: "Arial",
                      bold: true,
                      size: 16,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 110 },
                  children: [
                    new TextRun({
                      text: `${gradeLevel} - ${section}`,
                      font: "Arial",
                      size: 11,
                    }),
                  ],
                }),
                sogResultsTable,
                new Paragraph({ spacing: { before: 0, after: 120 } }),
                sogSignatureTable,
              ],
            },
          ],
        });

        const sogWordBlob = await Packer.toBlob(sogDocument);
        const sogFileName =
          `SOG_Report_${gradeLevel}_${section}_Term${sogTerm}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const sogUrl = URL.createObjectURL(sogWordBlob);
        const sogAnchor = document.createElement("a");
        sogAnchor.href = sogUrl;
        sogAnchor.download = sogFileName;
        document.body.appendChild(sogAnchor);
        sogAnchor.click();
        sogAnchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(sogUrl), 1000);
        return;
      }

      // SF5 uses its own Word layout so the downloaded document matches
      // the landscape SF5 bond-paper preview instead of the generic form.
      if (selectedForm === "SF5" && sf5Metrics) {
        const options = buildAdminSf5TemplateOptions();
        if (!options) {
          throw new Error("Select an SF5 class record before exporting.");
        }

        try {
          await downloadSF5WordTemplate(
            `SF5_Promotion_${selectedClass.section || "class"}`,
            options,
          );
          return;
        } catch (templateError) {
          console.warn(
            "SF5 Word template export failed; using the restored legacy exporter.",
            templateError,
          );
        }

        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "000000",
        };

        const sf5Borders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const borderlessTable = {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        };

        const sf5Blue = "0B5A96";

        const makeSf5Cell = ({
          text,
          width,
          bold = false,
          fontSize = 9,
          alignment = AlignmentType.LEFT,
          fill = "FFFFFF",
          color = "000000",
          columnSpan,
          height,
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          color?: string;
          columnSpan?: number;
          height?: number;
        }) =>
          new TableCell({
            width: {
              size: width,
              type: WidthType.PERCENTAGE,
            },
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: {
              type: ShadingType.CLEAR,
              fill,
            },
            borders: sf5Borders,
            margins: {
              top: 45,
              bottom: 45,
              left: 45,
              right: 45,
            },
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize,
                    color,
                  }),
                ],
              }),
            ],
          });

        const makeSf5HeaderCell = (text: string, width: number, fontSize = 8) =>
          makeSf5Cell({
            text,
            width,
            bold: true,
            fontSize,
            alignment: AlignmentType.CENTER,
            fill: sf5Blue,
            color: "FFFFFF",
          });

        const makeSf5InfoCell = (
          label: string,
          value: string,
          width = 33.33,
          columnSpan?: number,
        ) =>
          new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            borders: sf5Borders,
            margins: { top: 60, bottom: 60, left: 80, right: 80 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: `${label} `,
                    font: "Arial",
                    bold: true,
                    size: 9,
                  }),
                  new TextRun({
                    text: value,
                    font: "Arial",
                    size: 9,
                  }),
                ],
              }),
            ],
          });

        const sf5HeaderTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 18, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 90 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children:
                        logoBytes !== null
                          ? [
                              new ImageRun({
                                data: logoBytes,
                                type: logoType,
                                transformation: {
                                  width: 74,
                                  height: 43,
                                },
                              }),
                            ]
                          : [],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 64, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 8 },
                      children: [
                        new TextRun({
                          text: "School Form 5 (SF 5) Report on Promotion and Learning Progress & Achievement",
                          font: "Arial",
                          bold: true,
                          size: 16,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Revised to conform with the Instructions of DepEd Order 8, s. 2015",
                          font: "Arial",
                          italics: true,
                          size: 8,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 18, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
          ],
        });

        const sf5GradeSectionTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeSf5InfoCell("Grade Level", gradeLevel, 50),
                makeSf5InfoCell("Section", section, 50),
              ],
            }),
          ],
        });

        const sf5InfoTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeSf5InfoCell("Region", schoolProfile?.region || "", 33.33),
                makeSf5InfoCell(
                  "Division",
                  schoolProfile?.division || "",
                  33.33,
                ),
                makeSf5InfoCell(
                  "District",
                  schoolProfile?.district || "",
                  33.34,
                ),
              ],
            }),
            new TableRow({
              cantSplit: true,
              children: [
                makeSf5InfoCell(
                  "School ID",
                  schoolProfile?.school_id || "",
                  33.33,
                ),
                makeSf5InfoCell("School Year", schoolYear, 33.33),
                makeSf5InfoCell("Curriculum", "K to 12", 33.34),
              ],
            }),
            new TableRow({
              cantSplit: true,
              children: [
                makeSf5InfoCell(
                  "School Name",
                  schoolProfile?.school_name ||
                    "Agusan del Sur National Science High School",
                  66.66,
                  2,
                ),
                new TableCell({
                  width: { size: 33.34, type: WidthType.PERCENTAGE },
                  borders: sf5Borders,
                  margins: { top: 0, bottom: 0, left: 0, right: 0 },
                  children: [sf5GradeSectionTable],
                }),
              ],
            }),
          ],
        });

        const sf5VisibleLearners = sf5Metrics.learners.slice(
          0,
          sf5ReportLength === "full" ? 25 : 8,
        );
        const sf5TargetRows = sf5ReportLength === "full" ? 12 : 4;
        const sf5EmptyRows = Math.max(
          0,
          sf5TargetRows - sf5VisibleLearners.length,
        );

        const sf5LearnerRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeSf5HeaderCell("LRN", 13),
              makeSf5HeaderCell(
                "LEARNER'S NAME\n(Last Name, First Name, Middle Name)",
                35,
              ),
              makeSf5HeaderCell("GENERAL AVERAGE\n(Whole numbers)", 14),
              makeSf5HeaderCell("ACTION TAKEN", 14),
              makeSf5HeaderCell(
                "Did Not Meet Expectations\nIncomplete / Failed Learning Areas",
                24,
              ),
            ],
          }),
          ...sf5VisibleLearners.map(
            (learner) =>
              new TableRow({
                cantSplit: true,
                height: { value: 270, rule: HeightRule.ATLEAST },
                children: [
                  makeSf5Cell({ text: learner.lrn, width: 13, fontSize: 9 }),
                  makeSf5Cell({ text: learner.name, width: 35, fontSize: 9 }),
                  makeSf5Cell({
                    text: learner.generalAverage
                      ? String(Math.round(learner.generalAverage))
                      : "",
                    width: 14,
                    fontSize: 9,
                    alignment: AlignmentType.CENTER,
                  }),
                  makeSf5Cell({
                    text: learner.action,
                    width: 14,
                    bold: true,
                    fontSize: 9,
                    alignment: AlignmentType.CENTER,
                  }),
                  makeSf5Cell({
                    text: learner.didNotMeet,
                    width: 24,
                    fontSize: 8,
                  }),
                ],
              }),
          ),
          ...Array.from({ length: sf5EmptyRows }).map(
            () =>
              new TableRow({
                cantSplit: true,
                height: { value: 270, rule: HeightRule.ATLEAST },
                children: [
                  makeSf5Cell({ text: "", width: 13 }),
                  makeSf5Cell({ text: "", width: 35 }),
                  makeSf5Cell({ text: "", width: 14 }),
                  makeSf5Cell({ text: "", width: 14 }),
                  makeSf5Cell({ text: "", width: 24 }),
                ],
              }),
          ),
          new TableRow({
            cantSplit: true,
            children: [
              makeSf5Cell({
                text: `TOTAL MALE: ${sf5Metrics.maleTotal}`,
                width: 100,
                columnSpan: 5,
                bold: true,
                fontSize: 9,
              }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              makeSf5Cell({
                text: `TOTAL FEMALE: ${sf5Metrics.femaleTotal}`,
                width: 100,
                columnSpan: 5,
                bold: true,
                fontSize: 9,
              }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              makeSf5Cell({
                text: `COMBINED TOTAL: ${sf5Metrics.learners.length}`,
                width: 100,
                columnSpan: 5,
                bold: true,
                fontSize: 9,
              }),
            ],
          }),
        ];

        const sf5LearnersTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: sf5LearnerRows,
        });

        const summaryCounts = [
          ["PROMOTED", sf5Metrics.promoted],
          ["*Conditional", sf5Metrics.conditional],
          ["RETAINED", sf5Metrics.retained],
        ] as const;

        const sf5SummaryRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeSf5HeaderCell("SUMMARY TABLE", 52),
              makeSf5HeaderCell("MALE", 16, 7),
              makeSf5HeaderCell("FEMALE", 16, 7),
              makeSf5HeaderCell("TOTAL", 16, 7),
            ],
          }),
          ...summaryCounts.map(
            ([label, counts]) =>
              new TableRow({
                cantSplit: true,
                children: [
                  makeSf5Cell({ text: label, width: 52, fontSize: 8 }),
                  makeSf5Cell({
                    text: String(counts.male),
                    width: 16,
                    alignment: AlignmentType.CENTER,
                    fontSize: 8,
                  }),
                  makeSf5Cell({
                    text: String(counts.female),
                    width: 16,
                    alignment: AlignmentType.CENTER,
                    fontSize: 8,
                  }),
                  makeSf5Cell({
                    text: String(counts.total),
                    width: 16,
                    bold: true,
                    alignment: AlignmentType.CENTER,
                    fontSize: 8,
                  }),
                ],
              }),
          ),
          new TableRow({
            cantSplit: true,
            children: [
              makeSf5Cell({
                text: "LEARNING PROGRESS AND ACHIEVEMENT\n(Based on Learners' General Average)",
                width: 100,
                columnSpan: 4,
                bold: true,
                fontSize: 8,
                alignment: AlignmentType.CENTER,
                fill: sf5Blue,
                color: "FFFFFF",
              }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              makeSf5HeaderCell("Descriptors & Grading Scale", 52, 7),
              makeSf5HeaderCell("MALE", 16, 7),
              makeSf5HeaderCell("FEMALE", 16, 7),
              makeSf5HeaderCell("TOTAL", 16, 7),
            ],
          }),
          ...sf5Metrics.gradeBands.map(
            (band) =>
              new TableRow({
                cantSplit: true,
                children: [
                  makeSf5Cell({ text: band.label, width: 52, fontSize: 7 }),
                  makeSf5Cell({
                    text: String(band.male),
                    width: 16,
                    alignment: AlignmentType.CENTER,
                    fontSize: 7,
                  }),
                  makeSf5Cell({
                    text: String(band.female),
                    width: 16,
                    alignment: AlignmentType.CENTER,
                    fontSize: 7,
                  }),
                  makeSf5Cell({
                    text: String(band.total),
                    width: 16,
                    bold: true,
                    alignment: AlignmentType.CENTER,
                    fontSize: 7,
                  }),
                ],
              }),
          ),
        ];

        const sf5SummaryTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: sf5SummaryRows,
        });

        const makeSf5Signature = (
          label: string,
          name: string,
          role: string,
        ) => [
          new Paragraph({
            spacing: { before: 140, after: 220 },
            children: [
              new TextRun({
                text: label,
                font: "Arial",
                bold: true,
                size: 8,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: {
              top: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: "000000",
              },
            },
            spacing: { before: 0, after: 0 },
            children: [
              new TextRun({
                text: name,
                font: "Arial",
                bold: true,
                size: 8,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [
              new TextRun({
                text: role,
                font: "Arial",
                size: 7,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [
              new TextRun({
                text: "(Name and Signature)",
                font: "Arial",
                italics: true,
                size: 6,
              }),
            ],
          }),
        ];

        const sf5Guidelines = [
          "1. Do not include Dropouts and Transferred Out.",
          "2. Prepare using the learner's Form 138.",
          "3. *Conditional means failed in not more than two learning areas.",
          "4. Did Not Meet Expectations means the learner failed the learning area.",
          "5. Validation and submission are under the SDO.",
        ];

        const sf5RightChildren: Array<Paragraph | Table> = [
          sf5SummaryTable,
          ...makeSf5Signature("PREPARED BY:", teacherName, "Class Adviser"),
          ...makeSf5Signature(
            "CERTIFIED CORRECT & SUBMITTED:",
            "School Head",
            "",
          ),
          ...makeSf5Signature("REVIEWED BY:", "Division Representative", ""),
          new Paragraph({
            spacing: { before: 160, after: 35 },
            children: [
              new TextRun({
                text: "GUIDELINES:",
                font: "Arial",
                bold: true,
                size: 8,
              }),
            ],
          }),
          ...sf5Guidelines.map(
            (guideline) =>
              new Paragraph({
                spacing: { before: 0, after: 12, line: 180 },
                children: [
                  new TextRun({
                    text: guideline,
                    font: "Arial",
                    size: 6,
                  }),
                ],
              }),
          ),
        ];

        const sf5MainLayout = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 74, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  margins: { top: 0, bottom: 0, left: 0, right: 55 },
                  verticalAlign: VerticalAlign.TOP,
                  children: [
                    sf5LearnersTable,
                    ...(sf5ReportLength === "full"
                      ? [
                          new Paragraph({
                            spacing: { before: 120, after: 35 },
                            children: [
                              new TextRun({
                                text: "GUIDELINES:",
                                font: "Arial",
                                bold: true,
                                size: 8,
                              }),
                            ],
                          }),
                          ...sf5Guidelines.map(
                            (guideline) =>
                              new Paragraph({
                                spacing: { before: 0, after: 12, line: 180 },
                                children: [
                                  new TextRun({
                                    text: guideline,
                                    font: "Arial",
                                    size: 6,
                                  }),
                                ],
                              }),
                          ),
                        ]
                      : []),
                  ],
                }),
                new TableCell({
                  width: { size: 26, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  margins: { top: 0, bottom: 0, left: 55, right: 0 },
                  verticalAlign: VerticalAlign.TOP,
                  children: sf5RightChildren,
                }),
              ],
            }),
          ],
        });

        const sf5Document = new WordDocument({
          styles: {
            default: {
              document: {
                run: {
                  font: "Arial",
                  size: 9,
                  color: "000000",
                },
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
                    width: pageDimensions.height,
                    height: pageDimensions.width,
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: inchesToTwip(0.22),
                    bottom: inchesToTwip(0.22),
                    left: inchesToTwip(0.25),
                    right: inchesToTwip(0.25),
                  },
                },
              },
              children: [
                sf5HeaderTable,
                new Paragraph({ spacing: { before: 0, after: 45 } }),
                sf5InfoTable,
                new Paragraph({ spacing: { before: 0, after: 25 } }),
                sf5MainLayout,
              ],
            },
          ],
        });

        const sf5WordBlob = await Packer.toBlob(sf5Document);
        const sf5FileName =
          `SF5_${gradeLevel}_${section}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const sf5Url = URL.createObjectURL(sf5WordBlob);
        const sf5Anchor = document.createElement("a");
        sf5Anchor.href = sf5Url;
        sf5Anchor.download = sf5FileName;
        document.body.appendChild(sf5Anchor);
        sf5Anchor.click();
        sf5Anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(sf5Url), 1000);
        return;
      }

      // ILS uses its own Word layout so the downloaded document matches
      // the portrait Individual Learner Status bond-paper preview.
      if (selectedForm === "ILS" && ilsMetrics) {
        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "000000",
        };

        const ilsBorders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const borderlessTable = {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        };

        const learner = ilsMetrics.learner;
        const learnerDisplayName = learner
          ? (learner.last_name || studentName(learner)).toUpperCase()
          : "LEARNER NAME";
        const formattedToday = new Intl.DateTimeFormat("en-PH", {
          month: "long",
          day: "numeric",
          year: "numeric",
        }).format(new Date());

        const riskColors =
          ilsMetrics.riskLevel === "HIGH"
            ? {
                border: "DC2626",
                fill: "FEF2F2",
                text: "B91C1C",
              }
            : ilsMetrics.riskLevel === "MODERATE"
              ? {
                  border: "D97706",
                  fill: "FFFBEB",
                  text: "B45309",
                }
              : {
                  border: "16A34A",
                  fill: "F0FDF4",
                  text: "15803D",
                };

        const scoreText = (score: number | null) =>
          typeof score === "number" ? String(Math.round(score)) : "";

        const makeIlsCell = ({
          text,
          width,
          bold = false,
          fontSize = 8,
          alignment = AlignmentType.LEFT,
          fill = "FFFFFF",
          color = "000000",
          columnSpan,
          margins = { top: 58, bottom: 58, left: 55, right: 55 },
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          color?: string;
          columnSpan?: number;
          margins?: {
            top: number;
            bottom: number;
            left: number;
            right: number;
          };
        }) =>
          new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: { type: ShadingType.CLEAR, fill },
            borders: ilsBorders,
            margins,
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize * 2,
                    color,
                  }),
                ],
              }),
            ],
          });

        const ilsHeaderTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 0, bottom: 0, left: 0, right: 90 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children:
                        logoBytes !== null
                          ? [
                              new ImageRun({
                                data: logoBytes,
                                type: logoType,
                                transformation: { width: 72, height: 43 },
                              }),
                            ]
                          : [],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 56, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "REPUBLIC OF THE PHILIPPINES",
                          font: "Arial",
                          size: 18,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "DEPARTMENT OF EDUCATION",
                          font: "Arial",
                          bold: true,
                          size: 24,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: schoolProfile?.region
                            ? `Region ${schoolProfile.region}`
                            : "Region",
                          font: "Arial",
                          size: 16,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: schoolProfile?.division || "Division",
                          font: "Arial",
                          size: 16,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: (
                            schoolProfile?.school_name || "SCHOOL NAME"
                          ).toUpperCase(),
                          font: "Arial",
                          bold: true,
                          size: 18,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 22, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
          ],
        });

        const riskBorder = {
          style: BorderStyle.SINGLE,
          size: 7,
          color: riskColors.border,
        };

        const ilsRiskTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 48, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  shading: { type: ShadingType.CLEAR, fill: riskColors.fill },
                  borders: {
                    top: riskBorder,
                    bottom: riskBorder,
                    left: riskBorder,
                    right: noBorder,
                  },
                  margins: { top: 95, bottom: 95, left: 120, right: 50 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "●  ",
                          font: "Arial",
                          bold: true,
                          size: 24,
                          color: riskColors.border,
                        }),
                        new TextRun({
                          text: `RISK LEVEL: ${ilsMetrics.riskLevel}`,
                          font: "Arial",
                          bold: true,
                          size: 18,
                          color: riskColors.text,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 52, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  shading: { type: ShadingType.CLEAR, fill: riskColors.fill },
                  borders: {
                    top: riskBorder,
                    bottom: riskBorder,
                    left: noBorder,
                    right: riskBorder,
                  },
                  margins: { top: 95, bottom: 95, left: 50, right: 120 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: `${ilsMetrics.failingSubjectCount} failing subject(s) · ${ilsMetrics.absences} absence(s) · ${ilsMetrics.missingActivities} missing activities`,
                          font: "Arial",
                          size: 14,
                          color: "000000",
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        });

        const academicRows: TableRow[] = [
          new TableRow({
            cantSplit: true,
            children: [
              new TableCell({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnSpan: 7,
                shading: { type: ShadingType.CLEAR, fill: "F3F4F6" },
                borders: ilsBorders,
                margins: { top: 82, bottom: 82, left: 75, right: 75 },
                children: [
                  new Paragraph({
                    spacing: { before: 0, after: 0 },
                    children: [
                      new TextRun({
                        text: "1. ACADEMIC STATUS",
                        font: "Arial",
                        bold: true,
                        size: 18,
                      }),
                      new TextRun({
                        text: `                                      Class Rank: ${ilsMetrics.classRank || "—"} of ${ilsMetrics.classSize || 0} · Class Avg: ${typeof ilsMetrics.classAverage === "number" ? Math.round(ilsMetrics.classAverage) : "—"} · Student Avg: ${typeof ilsMetrics.generalAverage === "number" ? Math.round(ilsMetrics.generalAverage) : "—"}`,
                        font: "Arial",
                        size: 14,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeIlsCell({
                text: "Learning Area",
                width: 36,
                bold: true,
                fontSize: 7,
              }),
              makeIlsCell({
                text: "1st\nTerm",
                width: 8,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "2nd\nTerm",
                width: 8,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "3rd\nTerm",
                width: 8,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "Average",
                width: 10,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "Trend",
                width: 8,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "Remarks",
                width: 22,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
          ...ilsMetrics.subjects.map(
            (row) =>
              new TableRow({
                cantSplit: true,
                height: {
                  value: paperSize === "long" ? 300 : 255,
                  rule: HeightRule.ATLEAST,
                },
                children: [
                  makeIlsCell({
                    text: row.subject,
                    width: 36,
                    bold: true,
                    fontSize: 7,
                  }),
                  makeIlsCell({
                    text: scoreText(row.term1),
                    width: 8,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color:
                      typeof row.term1 === "number" && row.term1 < 75
                        ? "DC2626"
                        : "059669",
                  }),
                  makeIlsCell({
                    text: scoreText(row.term2),
                    width: 8,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color:
                      typeof row.term2 === "number" && row.term2 < 75
                        ? "DC2626"
                        : "059669",
                  }),
                  makeIlsCell({
                    text: scoreText(row.term3),
                    width: 8,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color:
                      typeof row.term3 === "number" && row.term3 < 75
                        ? "DC2626"
                        : "059669",
                  }),
                  makeIlsCell({
                    text: scoreText(row.average),
                    width: 10,
                    bold: true,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                  }),
                  makeIlsCell({
                    text: row.trend,
                    width: 8,
                    bold: true,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color:
                      row.trend === "↓"
                        ? "DC2626"
                        : row.trend === "✓"
                          ? "059669"
                          : "000000",
                  }),
                  makeIlsCell({
                    text: row.remarks,
                    width: 22,
                    bold: true,
                    fontSize: 6,
                    alignment: AlignmentType.CENTER,
                    color:
                      row.remarks === "FOR INTERVENTION" ? "DC2626" : "059669",
                  }),
                ],
              }),
          ),
          new TableRow({
            cantSplit: true,
            children: [
              makeIlsCell({
                text: "GENERAL AVERAGE",
                width: 60,
                columnSpan: 4,
                bold: true,
                fontSize: 7,
              }),
              makeIlsCell({
                text: scoreText(ilsMetrics.generalAverage),
                width: 10,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
                color: "059669",
              }),
              makeIlsCell({
                text: "—",
                width: 8,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text:
                  typeof ilsMetrics.generalAverage === "number"
                    ? ilsMetrics.generalAverage >= 75
                      ? "PASSED"
                      : "FOR INTERVENTION"
                    : "",
                width: 22,
                bold: true,
                fontSize: 6,
                alignment: AlignmentType.CENTER,
                color:
                  typeof ilsMetrics.generalAverage === "number" &&
                  ilsMetrics.generalAverage < 75
                    ? "DC2626"
                    : "059669",
              }),
            ],
          }),
        ];

        const academicTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: academicRows,
        });

        const componentRows: TableRow[] = [
          new TableRow({
            cantSplit: true,
            children: [
              makeIlsCell({
                text: `2. COMPONENT PERFORMANCE BREAKDOWN (${ilsMetrics.componentSubject})`,
                width: 100,
                columnSpan: 4,
                bold: true,
                fontSize: 8,
                fill: "F3F4F6",
              }),
            ],
          }),
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeIlsCell({
                text: "Term",
                width: 14,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
              }),
              makeIlsCell({
                text: "Written Works (20%)",
                width: 28.67,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
                fill: "EFF6FF",
              }),
              makeIlsCell({
                text: "Performance Tasks (50%)",
                width: 28.67,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
                fill: "F5F3FF",
              }),
              makeIlsCell({
                text: "Summative/Term Exams (30%)",
                width: 28.66,
                bold: true,
                fontSize: 7,
                alignment: AlignmentType.CENTER,
                fill: "FFF7ED",
              }),
            ],
          }),
          ...ilsMetrics.componentRows.map(
            (row) =>
              new TableRow({
                cantSplit: true,
                height: {
                  value: paperSize === "long" ? 310 : 265,
                  rule: HeightRule.ATLEAST,
                },
                children: [
                  makeIlsCell({
                    text: row.termLabel,
                    width: 14,
                    bold: true,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                  }),
                  makeIlsCell({
                    text: row.writtenWorks,
                    width: 28.67,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color: "059669",
                  }),
                  makeIlsCell({
                    text: row.performanceTasks,
                    width: 28.67,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color: "D97706",
                  }),
                  makeIlsCell({
                    text: row.summativeExam,
                    width: 28.66,
                    fontSize: 7,
                    alignment: AlignmentType.CENTER,
                    color: "D97706",
                  }),
                ],
              }),
          ),
        ];

        const componentTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: componentRows,
        });

        const attendanceTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeIlsCell({
                  text: "3. ATTENDANCE STATUS",
                  width: 100,
                  columnSpan: 2,
                  bold: true,
                  fontSize: 8,
                  fill: "F3F4F6",
                }),
              ],
            }),
            new TableRow({
              cantSplit: true,
              children: [
                makeIlsCell({
                  text: "Total Number of Days Absent:",
                  width: 70,
                  fontSize: 7,
                }),
                makeIlsCell({
                  text: String(ilsMetrics.absences),
                  width: 30,
                  bold: true,
                  fontSize: 7,
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
            new TableRow({
              cantSplit: true,
              children: [
                makeIlsCell({
                  text: "Total Number of Days Tardy / Half-Day:",
                  width: 70,
                  fontSize: 7,
                }),
                makeIlsCell({
                  text: String(ilsMetrics.lateHalfDays),
                  width: 30,
                  bold: true,
                  fontSize: 7,
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
          ],
        });

        const interventionTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeIlsCell({
                  text: "RECOMMENDED INTERVENTIONS",
                  width: 100,
                  bold: true,
                  fontSize: 8,
                  fill: "F3F4F6",
                }),
              ],
            }),
            ...ilsMetrics.recommendations.map(
              (recommendation) =>
                new TableRow({
                  cantSplit: true,
                  children: [
                    makeIlsCell({
                      text: `•  ${recommendation}`,
                      width: 100,
                      fontSize: 7,
                      margins: { top: 90, bottom: 90, left: 120, right: 70 },
                    }),
                  ],
                }),
            ),
          ],
        });

        const signatureTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  margins: { top: 0, bottom: 0, left: 100, right: 260 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      border: { top: blackBorder },
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: selectedTeacherName,
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Class Adviser",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  margins: { top: 0, bottom: 0, left: 260, right: 100 },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      border: { top: blackBorder },
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({ text: "", font: "Arial", size: 14 }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "School Head / Principal",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        });

        const acknowledgmentTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [
                    new Paragraph({
                      border: {
                        top: {
                          style: BorderStyle.DASHED,
                          size: 6,
                          color: "000000",
                        },
                      },
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 80, after: 100 },
                      children: [
                        new TextRun({
                          text: "Please sign and return this portion to the Class Adviser to confirm receipt.",
                          font: "Arial",
                          italics: true,
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 50 },
                      children: [
                        new TextRun({
                          text: `ACKNOWLEDGMENT SLIP — ${learner?.last_name || "Learner"}`.toUpperCase(),
                          font: "Arial",
                          bold: true,
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 20 },
                      children: [
                        new TextRun({
                          text: "☐ I acknowledge receipt of this Individual Learner Status & Intervention Report.",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 20 },
                      children: [
                        new TextRun({
                          text: "☐ I will attend the scheduled Parent-Teacher Conference on ____________________.",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 120 },
                      children: [
                        new TextRun({
                          text: "☐ I commit to monitoring my child's daily attendance and academic progress.",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 0 },
                      children: [
                        new TextRun({
                          text: "Parent/Guardian Signature: ______________________________     Contact No.: __________________     Date: __________________",
                          font: "Arial",
                          size: 14,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        });

        const ilsDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: { font: "Arial", size: 16, color: "000000" },
                paragraph: { spacing: { before: 0, after: 0, line: 205 } },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    width: pageDimensions.width,
                    height: pageDimensions.height,
                    orientation: PageOrientation.PORTRAIT,
                  },
                  margin: {
                    top: inchesToTwip(0.28),
                    bottom: inchesToTwip(0.28),
                    left: inchesToTwip(0.48),
                    right: inchesToTwip(0.48),
                  },
                },
              },
              children: [
                ilsHeaderTable,
                new Paragraph({
                  border: { bottom: blackBorder },
                  spacing: { before: 20, after: 95 },
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 20 },
                  children: [
                    new TextRun({
                      text: "INDIVIDUAL LEARNER STATUS & INTERVENTION REPORT",
                      font: "Arial",
                      bold: true,
                      underline: {},
                      size: 26,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 100 },
                  children: [
                    new TextRun({
                      text: `SY ${schoolYear}`,
                      font: "Arial",
                      size: 16,
                    }),
                  ],
                }),
                ilsRiskTable,
                new Paragraph({
                  spacing: { before: 110, after: 0 },
                  children: [
                    new TextRun({
                      text: "Date: ",
                      font: "Arial",
                      bold: true,
                      size: 16,
                    }),
                    new TextRun({
                      text: formattedToday,
                      font: "Arial",
                      size: 16,
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 30, after: 0 },
                  children: [
                    new TextRun({
                      text: "To the Parent/Guardian of:",
                      font: "Arial",
                      bold: true,
                      size: 16,
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 0, after: 80 },
                  children: [
                    new TextRun({
                      text: learnerDisplayName,
                      font: "Arial",
                      bold: true,
                      size: 20,
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 0, after: 80 },
                  children: [
                    new TextRun({
                      text: "Dear Parent/Guardian,",
                      font: "Arial",
                      size: 16,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.JUSTIFIED,
                  indent: { firstLine: 360 },
                  spacing: { before: 0, after: 95, line: 185 },
                  children: [
                    new TextRun({
                      text: "This formal communication serves to inform you regarding the current academic and attendance standing of your child. Based on our latest records, your child's academic standing is being monitored.",
                      font: "Arial",
                      size: 16,
                    }),
                  ],
                }),
                academicTable,
                new Paragraph({ spacing: { before: 0, after: 80 } }),
                componentTable,
                new Paragraph({ spacing: { before: 0, after: 80 } }),
                attendanceTable,
                new Paragraph({ spacing: { before: 0, after: 80 } }),
                interventionTable,
                new Paragraph({
                  alignment: AlignmentType.JUSTIFIED,
                  indent: { firstLine: 360 },
                  spacing: { before: 110, after: 50, line: 180 },
                  children: [
                    new TextRun({
                      text: "In light of this, we strongly request your presence for a ",
                      font: "Arial",
                      size: 16,
                    }),
                    new TextRun({
                      text: "Parent-Teacher Conference (PTC)",
                      font: "Arial",
                      bold: true,
                      size: 16,
                    }),
                    new TextRun({
                      text: " to discuss remediation strategies.",
                      font: "Arial",
                      size: 16,
                    }),
                  ],
                }),
                new Paragraph({
                  indent: { left: 500 },
                  spacing: { before: 0, after: 15 },
                  children: [
                    new TextRun({
                      text: "Date: ____________________",
                      font: "Arial",
                      size: 16,
                      color: "1859D6",
                    }),
                  ],
                }),
                new Paragraph({
                  indent: { left: 500 },
                  spacing: { before: 0, after: 15 },
                  children: [
                    new TextRun({
                      text: "Time: ____________________",
                      font: "Arial",
                      size: 16,
                      color: "DC2626",
                    }),
                  ],
                }),
                new Paragraph({
                  indent: { left: 500 },
                  spacing: {
                    before: 0,
                    after: paperSize === "long" ? 360 : 250,
                  },
                  children: [
                    new TextRun({
                      text: "Venue: ",
                      font: "Arial",
                      size: 16,
                      color: "DC2626",
                    }),
                    new TextRun({
                      text: "School Principal's Office",
                      font: "Arial",
                      size: 16,
                      bold: true,
                      underline: {},
                    }),
                  ],
                }),
                signatureTable,
                new Paragraph({
                  spacing: {
                    before: 0,
                    after: paperSize === "long" ? 260 : 180,
                  },
                }),
                acknowledgmentTable,
              ],
            },
          ],
        });

        const ilsWordBlob = await Packer.toBlob(ilsDocument);
        const ilsFileName =
          `Individual_Learner_Status_${gradeLevel}_${section}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const ilsUrl = URL.createObjectURL(ilsWordBlob);
        const ilsAnchor = document.createElement("a");
        ilsAnchor.href = ilsUrl;
        ilsAnchor.download = ilsFileName;
        document.body.appendChild(ilsAnchor);
        ilsAnchor.click();
        ilsAnchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(ilsUrl), 1000);
        return;
      }

      // Grade Slip uses its own Word layout so the downloaded document
      // matches the landscape Grade Slip bond-paper preview.
      if (selectedForm === "GRADE_SLIP" && sf9Metrics) {
        const blackBorder = {
          style: BorderStyle.SINGLE,
          size: 8,
          color: "000000",
        };

        const gradeSlipBorders = {
          top: blackBorder,
          bottom: blackBorder,
          left: blackBorder,
          right: blackBorder,
        };

        const borderlessTable = {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        };

        const gradeSlipBlue = "0B5A96";
        const gradeSlipGreen = "0A805B";
        const gradeSlipLightGreen = "D7F5E6";
        const gradeSlipGeneralFill = "F2F3F4";

        const descriptorForGradeSlip = (score: number | null) => {
          if (typeof score !== "number") return "";
          if (score >= 90) return "Outstanding";
          if (score >= 85) return "Very Satisfactory";
          if (score >= 80) return "Satisfactory";
          if (score >= 75) return "Fairly Satisfactory";
          return "Did Not Meet Expectations";
        };

        const scoreForGradeSlip = (score: number | null) =>
          typeof score === "number" ? String(Math.round(score)) : "";

        const gradeSlipLearner = sf9Metrics.learner;
        const gradeSlipLearnerName = gradeSlipLearner
          ? studentName(gradeSlipLearner).toLowerCase()
          : "";
        const gradeSlipVisibleSubjects = sf9Metrics.subjects.slice(0, 9);
        const gradeSlipEmptyRows = Math.max(
          0,
          9 - gradeSlipVisibleSubjects.length,
        );

        const makeGradeSlipCell = ({
          text,
          width,
          bold = false,
          fontSize = 16,
          alignment = AlignmentType.CENTER,
          fill = "FFFFFF",
          color = "000000",
          columnSpan,
          margins = { top: 55, bottom: 55, left: 45, right: 45 },
        }: {
          text: string;
          width: number;
          bold?: boolean;
          fontSize?: number;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
          fill?: string;
          color?: string;
          columnSpan?: number;
          margins?: {
            top: number;
            bottom: number;
            left: number;
            right: number;
          };
        }) =>
          new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            columnSpan,
            verticalAlign: VerticalAlign.CENTER,
            shading: { type: ShadingType.CLEAR, fill },
            borders: gradeSlipBorders,
            margins,
            children: [
              new Paragraph({
                alignment,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: "Arial",
                    bold,
                    size: fontSize,
                    color,
                  }),
                ],
              }),
            ],
          });

        // Start the framed page with a real paragraph instead of a nested
        // table. Microsoft Word otherwise inserts an empty editable paragraph
        // at the upper-left corner of the frame, which looks like a stray text
        // cursor when the downloaded document is opened.
        const gradeSlipHeaderParagraph = new Paragraph({
          alignment: AlignmentType.CENTER,
          keepNext: true,
          spacing: { before: 20, after: 90 },
          children: [
            ...(logoBytes !== null
              ? [
                  new ImageRun({
                    data: logoBytes,
                    type: logoType,
                    transformation: {
                      width: 70,
                      height: 42,
                    },
                  }),
                ]
              : []),
            new TextRun({
              text: "   GRADE SLIP",
              font: "Arial",
              bold: true,
              size: 24,
            }),
          ],
        });

        const makeGradeSlipFieldCell = (
          label: string,
          value: string,
          width: number,
        ) =>
          new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            borders: headerCellBorders,
            margins: { top: 0, bottom: 50, left: 0, right: 120 },
            children: [
              new Paragraph({
                border: { bottom: blackBorder },
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: `${label}:  `,
                    font: "Arial",
                    bold: true,
                    size: 16,
                  }),
                  new TextRun({
                    text: value,
                    font: "Arial",
                    bold: label === "Name",
                    size: 16,
                  }),
                ],
              }),
            ],
          });

        const gradeSlipDetailsTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeGradeSlipFieldCell("Name", gradeSlipLearnerName, 68),
                makeGradeSlipFieldCell("LRN", gradeSlipLearner?.lrn || "", 32),
              ],
            }),
            new TableRow({
              cantSplit: true,
              children: [
                makeGradeSlipFieldCell(
                  "Grade & Section",
                  `${gradeLevel} - ${section}`,
                  68,
                ),
                makeGradeSlipFieldCell("School Year", schoolYear, 32),
              ],
            }),
          ],
        });

        const gradeSlipRows: TableRow[] = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              makeGradeSlipCell({
                text: "Learning Area",
                width: 16,
                bold: true,
                fontSize: 16,
                alignment: AlignmentType.LEFT,
                fill: gradeSlipBlue,
                color: "FFFFFF",
              }),
              makeGradeSlipCell({
                text: "Term 1",
                width: 22,
                bold: true,
                fontSize: 16,
                fill: gradeSlipBlue,
                color: "FFFFFF",
              }),
              makeGradeSlipCell({
                text: "Term 2",
                width: 22,
                bold: true,
                fontSize: 16,
                fill: gradeSlipBlue,
                color: "FFFFFF",
              }),
              makeGradeSlipCell({
                text: "Term 3",
                width: 22,
                bold: true,
                fontSize: 16,
                fill: gradeSlipBlue,
                color: "FFFFFF",
              }),
              makeGradeSlipCell({
                text: "Final",
                width: 5,
                bold: true,
                fontSize: 15,
                fill: gradeSlipGreen,
                color: "FFFFFF",
              }),
              makeGradeSlipCell({
                text: "Descriptor",
                width: 13,
                bold: true,
                fontSize: 15,
                fill: gradeSlipBlue,
                color: "FFFFFF",
              }),
            ],
          }),
          ...gradeSlipVisibleSubjects.map(
            (row) =>
              new TableRow({
                cantSplit: true,
                height: { value: 300, rule: HeightRule.ATLEAST },
                children: [
                  makeGradeSlipCell({
                    text: row.subject,
                    width: 16,
                    bold: true,
                    fontSize: 16,
                    alignment: AlignmentType.LEFT,
                  }),
                  makeGradeSlipCell({
                    text: scoreForGradeSlip(row.term1),
                    width: 22,
                    fontSize: 16,
                  }),
                  makeGradeSlipCell({
                    text: scoreForGradeSlip(row.term2),
                    width: 22,
                    fontSize: 16,
                  }),
                  makeGradeSlipCell({
                    text: scoreForGradeSlip(row.term3),
                    width: 22,
                    fontSize: 16,
                  }),
                  makeGradeSlipCell({
                    text: scoreForGradeSlip(row.final),
                    width: 5,
                    bold: true,
                    fontSize: 16,
                    fill: gradeSlipLightGreen,
                  }),
                  makeGradeSlipCell({
                    text: descriptorForGradeSlip(row.final),
                    width: 13,
                    bold: true,
                    fontSize: 13,
                  }),
                ],
              }),
          ),
          ...Array.from({ length: gradeSlipEmptyRows }).map(
            () =>
              new TableRow({
                cantSplit: true,
                height: { value: 300, rule: HeightRule.ATLEAST },
                children: [
                  makeGradeSlipCell({ text: "", width: 16 }),
                  makeGradeSlipCell({ text: "", width: 22 }),
                  makeGradeSlipCell({ text: "", width: 22 }),
                  makeGradeSlipCell({ text: "", width: 22 }),
                  makeGradeSlipCell({
                    text: "",
                    width: 5,
                    fill: gradeSlipLightGreen,
                  }),
                  makeGradeSlipCell({ text: "", width: 13 }),
                ],
              }),
          ),
          new TableRow({
            cantSplit: true,
            height: { value: 310, rule: HeightRule.ATLEAST },
            children: [
              makeGradeSlipCell({
                text: "GENERAL AVERAGE",
                width: 82,
                columnSpan: 4,
                bold: true,
                fontSize: 17,
                fill: gradeSlipGeneralFill,
              }),
              makeGradeSlipCell({
                text: scoreForGradeSlip(sf9Metrics.generalAverage),
                width: 5,
                bold: true,
                fontSize: 16,
                fill: "BFEFD0",
              }),
              makeGradeSlipCell({
                text: descriptorForGradeSlip(sf9Metrics.generalAverage),
                width: 13,
                bold: true,
                fontSize: 13,
                fill: gradeSlipGeneralFill,
              }),
            ],
          }),
        ];

        const gradeSlipResultsTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: gradeSlipRows,
        });

        const makeGradeSlipSignatureCell = (
          name: string,
          role: string,
          width: number,
        ) =>
          new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            borders: headerCellBorders,
            verticalAlign: VerticalAlign.BOTTOM,
            margins: { top: 0, bottom: 0, left: 180, right: 180 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                border: { top: blackBorder },
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: name,
                    font: "Arial",
                    size: 15,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: role,
                    font: "Arial",
                    size: 14,
                  }),
                ],
              }),
            ],
          });

        const gradeSlipSignatureTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: borderlessTable,
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                makeGradeSlipSignatureCell(teacherName, "Class Adviser", 45),
                new TableCell({
                  width: { size: 10, type: WidthType.PERCENTAGE },
                  borders: headerCellBorders,
                  children: [new Paragraph({ text: "" })],
                }),
                makeGradeSlipSignatureCell("", "Parent/Guardian Signature", 45),
              ],
            }),
          ],
        });

        const gradeSlipFrameHeight = Math.max(
          1000,
          pageDimensions.width - inchesToTwip(0.48),
        );

        const gradeSlipFrameTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              cantSplit: true,
              height: {
                value: gradeSlipFrameHeight,
                rule: HeightRule.ATLEAST,
              },
              children: [
                new TableCell({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.TOP,
                  borders: gradeSlipBorders,
                  margins: {
                    top: 145,
                    bottom: 145,
                    left: 145,
                    right: 145,
                  },
                  children: [
                    gradeSlipHeaderParagraph,
                    gradeSlipDetailsTable,
                    new Paragraph({ spacing: { before: 0, after: 70 } }),
                    gradeSlipResultsTable,
                    new Paragraph({ spacing: { before: 0, after: 210 } }),
                    gradeSlipSignatureTable,
                  ],
                }),
              ],
            }),
          ],
        });

        const gradeSlipDocument = new WordDocument({
          styles: {
            default: {
              document: {
                run: {
                  font: "Arial",
                  size: 16,
                  color: "000000",
                },
                paragraph: {
                  spacing: { before: 0, after: 0, line: 180 },
                },
              },
            },
          },
          sections: [
            {
              properties: {
                page: {
                  size: {
                    width: pageDimensions.height,
                    height: pageDimensions.width,
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: inchesToTwip(0.15),
                    bottom: inchesToTwip(0.15),
                    left: inchesToTwip(0.15),
                    right: inchesToTwip(0.15),
                  },
                },
              },
              children: [gradeSlipFrameTable],
            },
          ],
        });

        const gradeSlipWordBlob = await Packer.toBlob(gradeSlipDocument);
        const gradeSlipFileName =
          `Grade_Slip_${gradeLevel}_${section}`
            .replaceAll(" ", "_")
            .replaceAll("/", "-") + ".docx";
        const gradeSlipUrl = URL.createObjectURL(gradeSlipWordBlob);
        const gradeSlipAnchor = document.createElement("a");
        gradeSlipAnchor.href = gradeSlipUrl;
        gradeSlipAnchor.download = gradeSlipFileName;
        document.body.appendChild(gradeSlipAnchor);
        gradeSlipAnchor.click();
        gradeSlipAnchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(gradeSlipUrl), 1000);
        return;
      }

      const createFieldCell = (label: string, value: string) =>
        new TableCell({
          width: {
            size: 50,
            type: WidthType.PERCENTAGE,
          },
          verticalAlign: VerticalAlign.CENTER,
          borders: {
            top: noBorder,
            left: noBorder,
            right: noBorder,
            bottom: {
              style: BorderStyle.SINGLE,
              size: 4,
              color: "D7CCC3",
            },
          },
          margins: {
            top: 60,
            bottom: 80,
            left: 20,
            right: 100,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: `${label}:`,
                  bold: true,
                  size: 15,
                  color: "55433C",
                }),
                new TextRun({
                  text: `     ${value}`,
                  size: 15,
                  color: "75635B",
                }),
              ],
            }),
          ],
        });

      const logoParagraph =
        logoBytes !== null
          ? new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new ImageRun({
                  data: logoBytes,
                  type: logoType,
                  transformation: {
                    width: 54,
                    height: 54,
                  },
                }),
              ],
            })
          : new Paragraph({
              text: "",
            });

      const headerTable = new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },

        layout: TableLayoutType.FIXED,

        borders: {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        },

        rows: [
          new TableRow({
            cantSplit: true,

            children: [
              // Left 20%: logo
              new TableCell({
                width: {
                  size: 25,
                  type: WidthType.PERCENTAGE,
                },

                borders: headerCellBorders,
                verticalAlign: VerticalAlign.CENTER,

                margins: {
                  top: 0,
                  bottom: 0,
                  left: 1000,
                  right: 0,
                },

                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: {
                      before: 0,
                      after: 0,
                    },

                    children:
                      logoBytes !== null
                        ? [
                            new ImageRun({
                              data: logoBytes,
                              type: logoType,

                              transformation: {
                                width: 48,
                                height: 48,
                              },
                            }),
                          ]
                        : [],
                  }),
                ],
              }),

              // Center 60%: heading text
              new TableCell({
                width: {
                  size: 50,
                  type: WidthType.PERCENTAGE,
                },

                borders: headerCellBorders,
                verticalAlign: VerticalAlign.CENTER,

                margins: {
                  top: 0,
                  bottom: 0,
                  left: 0,
                  right: 0,
                },

                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,

                    spacing: {
                      before: 0,
                      after: 24,
                    },

                    children: [
                      new TextRun({
                        text: "REPUBLIC OF THE PHILIPPINES",
                        font: "Arial",
                        size: 14,
                        color: "6F5C53",
                        characterSpacing: 30,
                      }),
                    ],
                  }),

                  new Paragraph({
                    alignment: AlignmentType.CENTER,

                    spacing: {
                      before: 0,
                      after: 24,
                    },

                    children: [
                      new TextRun({
                        text: "DEPARTMENT OF EDUCATION",
                        font: "Arial",
                        bold: true,
                        size: 18,
                      }),
                    ],
                  }),

                  new Paragraph({
                    alignment: AlignmentType.CENTER,

                    spacing: {
                      before: 0,
                      after: 0,
                    },

                    children: [
                      new TextRun({
                        text: "Agusan del Sur National Science High School",
                        font: "Arial",
                        bold: true,
                        size: 16,
                      }),
                    ],
                  }),
                ],
              }),

              // Right 20%: empty balancing column
              new TableCell({
                width: {
                  size: 25,
                  type: WidthType.PERCENTAGE,
                },

                borders: headerCellBorders,
                verticalAlign: VerticalAlign.CENTER,

                children: [
                  new Paragraph({
                    text: "",
                  }),
                ],
              }),
            ],
          }),
        ],
      });

      const detailsTable = new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        layout: TableLayoutType.FIXED,
        borders: {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        },
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              createFieldCell("Subject", subject),
              createFieldCell("Class Adviser", teacherName),
            ],
          }),

          new TableRow({
            cantSplit: true,
            children: [
              createFieldCell("Grade Level", gradeLevel),
              createFieldCell("Section", section),
            ],
          }),
        ],
      });

      const recordsHeader = new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          new TableCell({
            width: {
              size: 9,
              type: WidthType.PERCENTAGE,
            },
            shading: {
              type: ShadingType.CLEAR,
              fill: "F2ECE7",
            },
            borders: {
              top: thinBorder,
              bottom: thinBorder,
              left: thinBorder,
              right: thinBorder,
            },
            margins: {
              top: 100,
              bottom: 100,
              left: 100,
              right: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "NO.",
                    bold: true,
                    size: 14,
                    color: "5F4B43",
                  }),
                ],
              }),
            ],
          }),

          new TableCell({
            width: {
              size: 69,
              type: WidthType.PERCENTAGE,
            },
            shading: {
              type: ShadingType.CLEAR,
              fill: "F2ECE7",
            },
            borders: {
              top: thinBorder,
              bottom: thinBorder,
              left: thinBorder,
              right: thinBorder,
            },
            margins: {
              top: 100,
              bottom: 100,
              left: 100,
              right: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "LEARNER / RECORD",
                    bold: true,
                    size: 14,
                    color: "5F4B43",
                  }),
                ],
              }),
            ],
          }),

          new TableCell({
            width: {
              size: 22,
              type: WidthType.PERCENTAGE,
            },
            shading: {
              type: ShadingType.CLEAR,
              fill: "F2ECE7",
            },
            borders: {
              top: thinBorder,
              bottom: thinBorder,
              left: thinBorder,
              right: thinBorder,
            },
            margins: {
              top: 100,
              bottom: 100,
              left: 100,
              right: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "REMARKS",
                    bold: true,
                    size: 14,
                    color: "5F4B43",
                  }),
                ],
              }),
            ],
          }),
        ],
      });

      const recordRows = [1, 2, 3, 4, 5].map(
        (rowNumber) =>
          new TableRow({
            cantSplit: true,
            height: {
              value: 500,
              rule: HeightRule.ATLEAST,
            },
            children: [
              new TableCell({
                width: {
                  size: 9,
                  type: WidthType.PERCENTAGE,
                },
                borders: {
                  top: thinBorder,
                  bottom: thinBorder,
                  left: thinBorder,
                  right: thinBorder,
                },
                margins: {
                  top: 100,
                  bottom: 100,
                  left: 100,
                  right: 100,
                },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: String(rowNumber),
                        size: 14,
                        color: "8A7770",
                      }),
                    ],
                  }),
                ],
              }),

              new TableCell({
                width: {
                  size: 69,
                  type: WidthType.PERCENTAGE,
                },
                borders: {
                  top: thinBorder,
                  bottom: thinBorder,
                  left: thinBorder,
                  right: thinBorder,
                },
                children: [
                  new Paragraph({
                    text: "",
                  }),
                ],
              }),

              new TableCell({
                width: {
                  size: 22,
                  type: WidthType.PERCENTAGE,
                },
                borders: {
                  top: thinBorder,
                  bottom: thinBorder,
                  left: thinBorder,
                  right: thinBorder,
                },
                children: [
                  new Paragraph({
                    text: "",
                  }),
                ],
              }),
            ],
          }),
      );

      const recordsTable = new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        layout: TableLayoutType.FIXED,
        rows: [recordsHeader, ...recordRows],
      });

      const wordDocument = new WordDocument({
        styles: {
          default: {
            document: {
              run: {
                font: "Arial",
                size: 16,
                color: "3A2923",
              },
              paragraph: {
                spacing: {
                  after: 0,
                  line: 240,
                },
              },
            },
          },
        },
        sections: [
          {
            properties: {
              page: {
                size: {
                  width: pageDimensions.width,
                  height: pageDimensions.height,
                  orientation: PageOrientation.PORTRAIT,
                },
                margin: {
                  top: inchesToTwip(0.55),
                  bottom: inchesToTwip(0.55),
                  left: inchesToTwip(0.65),
                  right: inchesToTwip(0.65),
                },
              },
            },

            children: [
              headerTable,

              new Paragraph({
                border: {
                  bottom: {
                    style: BorderStyle.SINGLE,
                    size: 8,
                    color: "4A3A33",
                  },
                },
                spacing: {
                  after: 440,
                },
              }),

              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                  after: 80,
                },
                children: [
                  new TextRun({
                    text: title.toUpperCase(),
                    bold: true,
                    size: 24,
                  }),
                ],
              }),

              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                  after: 60,
                },
                children: [
                  new TextRun({
                    text: `${gradeLevel} · Section ${section}`,
                    bold: true,
                    size: 18,
                  }),
                ],
              }),

              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                  after: 400,
                },
                children: [
                  new TextRun({
                    text: `School Year ${schoolYear}`,
                    size: 15,
                    color: "75645D",
                  }),
                ],
              }),

              detailsTable,

              new Paragraph({
                spacing: {
                  after: 240,
                },
              }),

              recordsTable,

              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: {
                  before: 420,
                },
                children: [
                  new TextRun({
                    text: "Administrative preview · Open the teacher workspace to complete or print the official form.",
                    size: 13,
                    color: "99877E",
                  }),
                ],
              }),
            ],
          },
        ],
      });

      const wordBlob = await Packer.toBlob(wordDocument);

      const fileName =
        `${selectedFormDefinition.title}_${gradeLevel}_${section}`
          .replaceAll(" ", "_")
          .replaceAll("/", "-") + ".docx";

      const url = URL.createObjectURL(wordBlob);
      const anchor = document.createElement("a");

      anchor.href = url;
      anchor.download = fileName;

      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();

      window.setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 1000);
    } catch (wordError) {
      console.error(wordError);

      window.alert(
        "Unable to create the Word document. Check the browser console for details.",
      );
    }
  };

  const selectForm = (form: string) => {
    setSelectedForm(form);
    setSelectedClassId(null);
  };

  if (formsView === "submitted") {
    return (
      <SubmittedFormsAdminView
        classes={classes}
        students={students}
        grades={grades}
        schoolProfile={schoolProfile}
        onShowAvailable={() => setFormsView("available")}
        onOpenFullPreview={(submission) => {
          const normalizedCode = normalizeSubmissionFormCode(
            submission.form_code,
          );

          const formId =
            normalizedCode === "SOG Report"
              ? "SOG"
              : normalizedCode === "SF9 (New)"
                ? "SF9_NEW"
                : normalizedCode === "Anecdotal"
                  ? "ANECDOTAL"
                  : normalizedCode;

          if (FORM_TYPES.includes(formId)) {
            setSelectedForm(formId);
          }

          setSelectedClassId(submission.class_id);
          setFormsView("available");
        }}
      />
    );
  }

  return (
    <div className="space-y-5">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
        <PageHeading
          title="School Forms"
          description="Manage and organize school forms by form type, grade level, and section."
        />
        <div className="grid w-full gap-3 sm:grid-cols-2 xl:max-w-5xl xl:grid-cols-[minmax(0,1fr)_180px_190px_170px]">
          <SearchBar
            value={search}
            onChange={onSearch}
            placeholder="Search forms, grades, sections, or teachers..."
          />

          <div className="relative">
            <CalendarRange className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-primary)]" />
            <select
              value={selectedSchoolYear}
              onChange={(event) => {
                setSelectedSchoolYear(event.target.value);
                setSelectedClassId(null);
              }}
              aria-label="Filter school forms by school year"
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-10 text-sm font-medium text-[var(--admin-color-2d211d)] outline-none transition focus:border-[var(--admin-primary)] focus:ring-2 focus:ring-[var(--admin-primary)]/15"
            >
              <option value="all">All School Years</option>
              {schoolYearOptions.map((schoolYear) => (
                <option key={schoolYear} value={schoolYear}>
                  {schoolYear}
                </option>
              ))}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>

          <div className="relative">
            <GraduationCap className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
            <select
              value={selectedGradeLevel}
              onChange={(event) => {
                setSelectedGradeLevel(event.target.value);
                setSelectedClassId(null);
              }}
              aria-label="Filter school forms by grade level"
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-10 text-sm font-medium text-[var(--admin-color-2d211d)] outline-none transition focus:border-[var(--admin-primary)] focus:ring-2 focus:ring-[var(--admin-primary)]/15"
            >
              <option value="all">All Grade Levels</option>
              <option value="7">Grade 7</option>
              <option value="8">Grade 8</option>
              <option value="9">Grade 9</option>
              <option value="10">Grade 10</option>
              <option value="11">Grade 11</option>
              <option value="12">Grade 12</option>
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>

          <div className="relative">
            <CalendarDays className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-primary)]" />
            <select
              value={selectedDateRange}
              onChange={(event) => {
                setSelectedDateRange(event.target.value);
                setSelectedClassId(null);
              }}
              aria-label="Filter school forms by date"
              className="h-11 w-full appearance-none rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] pl-10 pr-10 text-sm font-medium text-[var(--admin-color-2d211d)] outline-none transition focus:border-[var(--admin-primary)] focus:ring-2 focus:ring-[var(--admin-primary)]/15"
            >
              <option value="all">All Dates</option>
              <option value="now">Now</option>
              <option value="day">Day</option>
              <option value="week">Week</option>
              <option value="month">Month</option>
              <option value="year">Year</option>
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
          </div>
        </div>
      </div>

      <div className="flex w-fit gap-1 rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-1">
        <button
          type="button"
          className="rounded-lg bg-[var(--admin-primary)] px-4 py-2 text-xs font-semibold text-white"
        >
          Available Forms
        </button>
        <button
          type="button"
          onClick={() => setFormsView("submitted")}
          className="rounded-lg px-4 py-2 text-xs font-semibold text-[var(--admin-text-muted)] hover:bg-[var(--admin-surface-soft)]"
        >
          Submitted Forms
        </button>
      </div>

      <div className="grid min-h-[680px] gap-4 xl:grid-cols-[280px_minmax(0,1fr)]">
        <Panel className="self-start xl:sticky xl:top-[92px]">
          <div className="mb-3 flex items-start justify-between gap-3">
            <div className="min-w-0">
              <h2 className="text-xs font-bold text-[var(--admin-text-title)]">
                Form Types
              </h2>
              <p className="mt-1 truncate text-[9px] font-medium text-[var(--admin-text-muted)]">
                {adminGradeLevelFormsLabel(selectedGradeLevel)}
              </p>
            </div>

            <span className="shrink-0 rounded-full bg-[var(--admin-color-f7e6e8)] px-2 py-1 text-[8px] font-bold text-[var(--admin-primary)]">
              {visibleFormTypeDefinitions.length}
            </span>
          </div>

          <div className="max-h-[700px] space-y-1 overflow-y-auto pr-1">
            {visibleFormTypeDefinitions.map((form) => {
              const Icon = form.icon;
              const active = selectedForm === form.id;

              return (
                <button
                  key={form.id}
                  type="button"
                  onClick={() => selectForm(form.id)}
                  title={form.description}
                  className={`flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left transition ${
                    active
                      ? "bg-[var(--admin-color-f9e8e9)] text-[var(--admin-primary)]"
                      : "text-[var(--admin-color-4e3c35)] hover:bg-[var(--admin-color-faf3ed)]"
                  }`}
                >
                  {active ? (
                    <ChevronDown className="size-3.5 shrink-0" />
                  ) : (
                    <ChevronRight className="size-3.5 shrink-0 text-[var(--admin-color-9a867d)]" />
                  )}

                  <span
                    className={`grid size-7 shrink-0 place-items-center rounded-md ${
                      active ? "bg-[var(--admin-white)]" : form.iconBackground
                    }`}
                  >
                    <Icon
                      className={`size-3.5 ${
                        active ? "text-[var(--admin-primary)]" : form.iconColor
                      }`}
                    />
                  </span>

                  <span className="min-w-0 flex-1 truncate text-xs font-semibold">
                    {form.title}
                  </span>

                  <span className="rounded-full bg-[var(--admin-white)] px-1.5 py-0.5 text-[8px] font-semibold text-[var(--admin-text-muted)] shadow-sm">
                    {filteredClasses.length}
                  </span>
                </button>
              );
            })}
          </div>
        </Panel>

        <div className="space-y-4">
          <Panel>
            <div className="flex flex-wrap items-center gap-2 text-[10px] text-[var(--admin-text-muted)]">
              <span>School Forms</span>
              <ChevronRight className="size-3" />
              <span className="font-semibold text-[var(--admin-primary)]">
                {selectedFormDefinition.title}
              </span>
              {selectedClass && (
                <>
                  <ChevronRight className="size-3" />
                  <span>{selectedClass.grade_level || "Unassigned"}</span>
                  <ChevronRight className="size-3" />
                  <span>{selectedClass.section || "No Section"}</span>
                </>
              )}
            </div>
            <div className="mt-3 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div>
                <h2 className="text-lg font-semibold text-[var(--admin-text-heading)]">
                  {selectedFormDefinition.title}
                </h2>
                <p className="mt-1 text-[10px] text-[var(--admin-text-muted)]">
                  {selectedFormDefinition.description}
                </p>
              </div>
              <StatusBadge>
                {filteredClasses.length} class record
                {filteredClasses.length === 1 ? "" : "s"}
              </StatusBadge>
            </div>
          </Panel>

          {selectedForm !== "SF6" && (
            <Panel>
              <PanelHeader title="Files" />
              <DataTable minWidth="900px">
                <thead>
                  <tr>
                    <Th>File Name</Th>
                    <Th>Teacher</Th>
                    <Th>Grade / Section</Th>
                    <Th>School Year</Th>
                    <Th>Date &amp; Time</Th>
                    <Th>Status</Th>
                    <Th align="right">Preview</Th>
                  </tr>
                </thead>
                <tbody>
                  {filteredClasses.map((item) => {
                    const teacher = teacherById.get(item.teacher_id);
                    const selected = selectedClass?.id === item.id;
                    const submission =
                      latestSubmissionByClass.get(item.id) ?? null;
                    const pendingReview =
                      submission?.status === "submitted" ||
                      submission?.status === "pending_review";
                    const approved = submission?.status === "approved";
                    const returned = submission?.status === "returned";
                    const hasSubmission = Boolean(submission);
                    const formDateTime = formatDateTime(
                      submission?.submitted_at ?? null,
                    );

                    return (
                      <tr
                        key={item.id}
                        className={`border-b border-[var(--admin-row-divider)] last:border-0 ${selected ? "bg-[var(--admin-color-fff8f4)]" : "hover:bg-[var(--admin-surface-soft)]"}`}
                      >
                        <Td strong>
                          <div className="flex items-center gap-2">
                            <div className="grid size-7 place-items-center rounded bg-[var(--admin-color-f7e6e8)] text-[var(--admin-primary-strong)]">
                              <FileText className="size-3.5" />
                            </div>
                            {selectedForm === "GSA"
                              ? `GSA_${item.section || "Section"}.pdf`
                              : selectedForm === "SOG"
                                ? `SOG_Report_${item.grade_level || "Grade"}_${item.section || "Section"}_Term${sogTerm}.pdf`
                                : selectedForm === "SF5"
                                  ? `SF5_Promotion_${item.section || "Section"}.pdf`
                                  : selectedForm === "SF1"
                                    ? `SF1_Register_${item.section || "Section"}.pdf`
                                    : selectedForm === "SF9_NEW"
                                      ? `SF9_New_Progress_Performance_${item.section || "Section"}.pdf`
                                      : selectedForm === "SF6"
                                        ? `SF6_Summarized_Promotion_${item.section || "Section"}.pdf`
                                        : selectedForm === "SF8"
                                          ? `SF8_Health_Nutrition_${item.section || "Section"}.pdf`
                                          : selectedForm === "SF10"
                                            ? `SF10_Permanent_Academic_Record_${item.section || "Section"}.pdf`
                                            : `${selectedFormDefinition.title.replaceAll(" ", "_")}_${item.grade_level || "Grade"}_${item.section || "Section"}.pdf`}
                          </div>
                        </Td>
                        <Td>
                          {item.teacher_name ||
                            teacher?.full_name ||
                            teacher?.email ||
                            "-"}
                        </Td>
                        <Td>
                          {item.grade_level || "-"} / {item.section || "-"}
                        </Td>
                        <Td>{item.school_year || "-"}</Td>
                        <Td>
                          <div className="flex items-center gap-2 whitespace-nowrap">
                            <CalendarDays className="size-3.5 shrink-0 text-[var(--admin-primary)]" />
                            <div>
                              <div className="text-[10px] font-semibold text-[var(--admin-text-heading)]">
                                {formDateTime.date}
                              </div>
                              <div className="text-[9px] text-[var(--admin-text-muted)]">
                                {formDateTime.time}
                              </div>
                            </div>
                          </div>
                        </Td>
                        <Td>
                          {pendingReview ? (
                            <span className="inline-flex items-center gap-1.5 whitespace-nowrap text-[9px] font-bold text-emerald-700">
                              <CircleCheck className="size-3.5" />
                              Submitted
                            </span>
                          ) : approved ? (
                            <span className="inline-flex items-center gap-1.5 whitespace-nowrap text-[9px] font-bold text-emerald-700">
                              <CircleCheck className="size-3.5" />
                              Approved
                            </span>
                          ) : returned ? (
                            <span className="inline-flex items-center gap-1.5 whitespace-nowrap text-[9px] font-bold text-amber-700">
                              <AlertTriangle className="size-3.5" />
                              Returned
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1.5 whitespace-nowrap text-[9px] font-bold text-rose-600">
                              <X className="size-3.5" />
                              Not Submitted
                            </span>
                          )}
                        </Td>
                        <Td align="right">
                          <button
                            type="button"
                            disabled={!hasSubmission}
                            onClick={() => {
                              if (!hasSubmission) return;
                              setSelectedClassId(item.id);
                            }}
                            className="rounded-md border border-[var(--admin-color-e3d7cc)] bg-[var(--admin-white)] px-2.5 py-1.5 text-[9px] font-semibold text-[var(--admin-primary)] hover:bg-[var(--admin-color-fff3f4)] disabled:cursor-not-allowed disabled:border-[var(--admin-row-divider)] disabled:bg-[var(--admin-surface-soft)] disabled:text-[var(--admin-text-muted)] disabled:opacity-60"
                          >
                            {hasSubmission ? "View" : "—"}
                          </button>
                        </Td>
                      </tr>
                    );
                  })}
                  {filteredClasses.length === 0 && (
                    <TableEmpty
                      colSpan={7}
                      text="No matching school form records found."
                    />
                  )}
                </tbody>
              </DataTable>
            </Panel>
          )}

          <Panel>
            <PanelHeader title="Preview" />
            {selectedClass && (isSf6Preview || selectedClassSubmission) ? (
              <div className="space-y-3">
                <div className="flex flex-col gap-3 rounded-xl border border-[var(--admin-color-ded2c7)] bg-[var(--admin-surface)] p-3 xl:flex-row xl:items-center xl:justify-between">
                  <div className="flex flex-wrap items-center gap-2">
                    {isSogPreview && (
                      <select
                        value={sogTerm}
                        onChange={(event) =>
                          setSogTerm(event.target.value as SogTerm)
                        }
                        className="h-9 min-w-[220px] rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-semibold text-[var(--admin-text-heading)] outline-none ring-0"
                        aria-label="SOG report term"
                      >
                        <option value="1">SUMMARY OF GRADES - TERM 1</option>
                        <option value="2">SUMMARY OF GRADES - TERM 2</option>
                        <option value="3">SUMMARY OF GRADES - TERM 3</option>
                      </select>
                    )}

                    {isSf5Preview && (
                      <select
                        value={sf5ReportLength}
                        onChange={(event) =>
                          setSf5ReportLength(
                            event.target.value as Sf5ReportLength,
                          )
                        }
                        className="h-9 rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-medium text-[var(--admin-text-heading)] outline-none ring-0"
                        aria-label="SF5 report length"
                      >
                        <option value="short">Short</option>
                        <option value="full">Full / Long</option>
                      </select>
                    )}

                    {isSf1Preview && (
                      <>
                        <span className="text-xs font-semibold text-[var(--admin-color-5f4b43)]">
                          Report length
                        </span>
                        <select
                          value={sf1ReportLength}
                          onChange={(event) =>
                            setSf1ReportLength(
                              event.target.value as Sf1ReportLength,
                            )
                          }
                          className="h-9 rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-medium text-[var(--admin-text-heading)] outline-none ring-0"
                          aria-label="SF1 report length"
                        >
                          <option value="full">Full / Long</option>
                          <option value="short">Short</option>
                        </select>
                      </>
                    )}

                    {isSf9NewPreview && selectedClassSubmission && (
                      <>
                        <span className="text-xs font-semibold text-[var(--admin-color-5f4b43)]">
                          Learner
                        </span>
                        <select
                          value={selectedSf9Learner?.id ?? ""}
                          onChange={(event) =>
                            setSelectedSf9LearnerId(event.target.value)
                          }
                          disabled={sf9ClassLearners.length === 0}
                          className="h-9 min-w-[250px] max-w-[360px] rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-semibold text-[var(--admin-text-heading)] outline-none ring-0 disabled:cursor-not-allowed disabled:opacity-60"
                          aria-label="Select learner for SF9 bondpaper preview"
                          title="Select which learner's submitted SF9 to preview"
                        >
                          {sf9ClassLearners.length === 0 ? (
                            <option value="">No learners available</option>
                          ) : (
                            sf9ClassLearners.map((learner) => (
                              <option key={learner.id} value={learner.id}>
                                {studentName(learner)}
                                {learner.lrn ? ` — ${learner.lrn}` : ""}
                              </option>
                            ))
                          )}
                        </select>
                      </>
                    )}

                    <span className="text-xs font-semibold text-[var(--admin-color-5f4b43)]">
                      Paper size
                    </span>
                    <select
                      value={isSf5Preview ? "long" : paperSize}
                      onChange={(event) =>
                        setPaperSize(event.target.value as PaperSizeKey)
                      }
                      disabled={isSf5Preview}
                      className="h-9 rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-medium text-[var(--admin-text-heading)] outline-none ring-0"
                      title={
                        isSf5Preview
                          ? "SF5 uses the original long-bond landscape template"
                          : "Select paper size"
                      }
                    >
                      {Object.entries(PAPER_SIZE_PRESETS).map(
                        ([key, preset]) => (
                          <option key={key} value={key}>
                            {preset.label}
                          </option>
                        ),
                      )}
                    </select>

                    {isSf2Preview && (
                      <div className="ml-1 flex items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => moveSf2Month(-1)}
                          className="grid size-9 place-items-center rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] text-[var(--admin-text-heading)] hover:bg-[var(--admin-color-faf3ed)]"
                          title="Previous month"
                          aria-label="Previous month"
                        >
                          <ChevronLeft className="size-4" />
                        </button>

                        <label className="relative flex h-9 min-w-[155px] items-center gap-2 rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs font-semibold uppercase text-[var(--admin-text-heading)]">
                          <CalendarDays className="size-4 text-[var(--admin-primary)]" />
                          <span className="pointer-events-none flex-1 text-center">
                            {sf2MonthLabel}
                          </span>
                          <input
                            type="month"
                            value={sf2MonthValue}
                            onChange={(event) => {
                              const [year, month] = event.target.value
                                .split("-")
                                .map(Number);
                              if (year && month) {
                                setSf2Month(new Date(year, month - 1, 1));
                              }
                            }}
                            className="absolute inset-0 cursor-pointer opacity-0"
                            aria-label="Select SF2 report month"
                          />
                        </label>

                        <button
                          type="button"
                          onClick={() => moveSf2Month(1)}
                          className="grid size-9 place-items-center rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] text-[var(--admin-text-heading)] hover:bg-[var(--admin-color-faf3ed)]"
                          title="Next month"
                          aria-label="Next month"
                        >
                          <ChevronRight className="size-4" />
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={resetPreview}
                      className="grid size-9 place-items-center rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] text-[#17345f] hover:bg-[var(--admin-color-faf3ed)]"
                      title="Reset preview"
                    >
                      <RefreshCcw className="size-4" />
                    </button>

                    <button
                      type="button"
                      onClick={printSchoolForm}
                      className="inline-flex h-9 items-center gap-2 rounded-md bg-[#17345f] px-3 text-[10px] font-semibold text-white shadow-sm hover:bg-[#102746]"
                    >
                      <Printer className="size-3.5" />
                      Print / Save as PDF
                    </button>

                    {isSf5Preview && (
                      <button
                        type="button"
                        onClick={exportSf5ToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    {isSf1Preview && (
                      <button
                        type="button"
                        onClick={exportSf1ToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    {isSf6Preview && (
                      <button
                        type="button"
                        onClick={exportSf6ToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    {isSf9NewPreview && (
                      <button
                        type="button"
                        onClick={exportSf9NewToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    {isSf8Preview && (
                      <button
                        type="button"
                        onClick={exportSf8ToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    {isSf10Preview && (
                      <button
                        type="button"
                        onClick={exportSf10ToExcel}
                        className="inline-flex h-9 items-center gap-2 rounded-md bg-[#15803d] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#166534]"
                      >
                        <FileSpreadsheet className="size-3.5" />
                        Export to Excel
                      </button>
                    )}

                    <button
                      type="button"
                      onClick={copySchoolFormToWord}
                      className="inline-flex h-9 items-center gap-2 rounded-md bg-[#1859d6] px-3 text-[10px] font-semibold uppercase text-white shadow-sm hover:bg-[#1249b5]"
                    >
                      <Copy className="size-3.5" />
                      Copy to Word
                    </button>

                    <span className="px-1 text-[10px] font-semibold text-[#17345f]">
                      {isSogPreview && sogMetrics
                        ? `${sogMetrics.learners.length} learner${sogMetrics.learners.length === 1 ? "" : "s"}`
                        : isSf9NewPreview && selectedClassSubmission
                          ? `${sf9ClassLearners.length} learner${sf9ClassLearners.length === 1 ? "" : "s"}`
                          : `${subjectCount} subject${subjectCount === 1 ? "" : "s"}`}
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={decreaseZoom}
                      disabled={zoom <= 50}
                      className="grid size-9 place-items-center rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] text-[var(--admin-text-heading)] hover:bg-[var(--admin-color-faf3ed)] disabled:cursor-not-allowed disabled:opacity-40"
                      title="Zoom out"
                    >
                      <Minus className="size-4" />
                    </button>
                    <div className="min-w-16 rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 py-2 text-center text-xs font-semibold text-[var(--admin-text-heading)]">
                      {zoom}%
                    </div>
                    <button
                      type="button"
                      onClick={increaseZoom}
                      disabled={zoom >= 200}
                      className="grid size-9 place-items-center rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] text-[var(--admin-text-heading)] hover:bg-[var(--admin-color-faf3ed)] disabled:cursor-not-allowed disabled:opacity-40"
                      title="Zoom in"
                    >
                      <Plus className="size-4" />
                    </button>
                  </div>
                </div>

                <div className="overflow-hidden rounded-lg border border-[var(--admin-color-ded2c7)] bg-[var(--admin-color-ece9e6)]">
                  <div
                    className="max-h-[calc(100vh-245px)] min-h-[520px] overflow-auto p-4 sm:p-6"
                    style={{
                      scrollbarGutter: "stable both-edges",
                      overscrollBehavior: "contain",
                    }}
                  >
                    <div className="flex min-w-full w-max justify-center">
                      <div
                        className="relative shrink-0"
                        style={{
                          width: `${scaledPaperWidth}px`,
                          height: `${scaledPaperHeight}px`,
                        }}
                      >
                        <div
                          id="admin-school-form-paper"
                          className="absolute left-0 top-0 origin-top-left bg-[var(--admin-white)] shadow-2xl"
                          style={{
                            width: `${selectedPaper.width}px`,
                            minHeight: `${selectedPaper.height}px`,
                            transform: `scale(${previewZoomScale})`,
                            transformOrigin: "top left",
                          }}
                        >
                          {isSogPreview && sogMetrics ? (
                            <SogBondPaper
                              selectedClass={selectedClass}
                              selectedTeacherName={selectedTeacherName}
                              metrics={sogMetrics}
                              paperHeight={selectedPaper.height}
                            />
                          ) : isGradeSlipPreview && sf9Metrics ? (
                            <GradeSlipBondPaper
                              selectedClass={selectedClass}
                              selectedTeacherName={selectedTeacherName}
                              metrics={sf9Metrics}
                              paperHeight={selectedPaper.height}
                            />
                          ) : isGsaPreview && gsaMetrics ? (
                            <div
                              className="box-border p-8 text-[10px] text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div className="grid grid-cols-[180px_minmax(0,1fr)_180px] items-center gap-4">
                                <div className="flex justify-end pr-4">
                                  <img
                                    src={deped}
                                    alt="Department of Education logo"
                                    className="h-16 w-auto object-contain"
                                  />
                                </div>

                                <div className="text-center leading-tight">
                                  <div className="text-[9px]">
                                    Republic of the Philippines
                                  </div>
                                  <div className="text-[11px] font-bold text-[#0038A8]">
                                    DEPARTMENT OF EDUCATION
                                  </div>
                                  <div className="text-[9px]">
                                    {schoolProfile?.region
                                      ? `Region ${schoolProfile.region}`
                                      : "Region XIII — CARAGA"}
                                  </div>
                                  <div className="text-[9px] font-semibold uppercase">
                                    {schoolProfile?.school_name ||
                                      "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL"}
                                  </div>
                                </div>

                                <div aria-hidden="true" />
                              </div>

                              <div className="mt-3 border-t border-black" />

                              <div className="mt-3 text-center text-[13px] font-bold leading-tight">
                                REPORT ON GENERAL SCHOLASTIC APTITUDE OF
                                LEARNERS
                                <div>PER TERM BY LEARNING AREA</div>
                              </div>

                              <div className="mt-1 text-center text-[9px]">
                                School Year:{" "}
                                {selectedClass.school_year || "____-____"}
                              </div>

                              <div className="mt-4 flex items-end justify-between text-[9px]">
                                <div>
                                  <b>Teacher:</b>{" "}
                                  <span className="ml-1 border-b border-black px-6 text-[#0038A8]">
                                    {selectedTeacherName}
                                  </span>
                                </div>
                                <div>
                                  <b>Term:</b> <u>1st</u>
                                </div>
                              </div>

                              <table className="mt-2 w-full table-fixed border-collapse border border-black text-[8px]">
                                <thead>
                                  <tr>
                                    <th
                                      rowSpan={2}
                                      className="w-[8.5%] border border-black px-1 py-1"
                                    >
                                      Grade Level/
                                      <br />
                                      Subjects
                                    </th>
                                    <th
                                      rowSpan={2}
                                      className="w-[6.5%] border border-black px-1 py-1"
                                    >
                                      STRAND
                                    </th>
                                    <th
                                      rowSpan={2}
                                      className="w-[6.5%] border border-black px-1 py-1"
                                    >
                                      SECTION
                                    </th>
                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-1"
                                    >
                                      REGISTERED LEARNERS
                                    </th>
                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-1"
                                    >
                                      LEARNERS WITH 75% AND ABOVE
                                    </th>
                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-1"
                                    >
                                      TOTAL GRADES
                                    </th>
                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-1"
                                    >
                                      AVERAGE GRADE
                                    </th>
                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-1"
                                    >
                                      % OF PROFICIENCY (75%+)
                                    </th>
                                  </tr>
                                  <tr>
                                    {Array.from({ length: 5 }).flatMap(
                                      (_, groupIndex) => [
                                        <th
                                          key={`${groupIndex}-male`}
                                          className="border border-black px-0.5 py-0.5"
                                        >
                                          MALE
                                        </th>,
                                        <th
                                          key={`${groupIndex}-female`}
                                          className="border border-black px-0.5 py-0.5"
                                        >
                                          FEMALE
                                        </th>,
                                        <th
                                          key={`${groupIndex}-total`}
                                          className="border border-black px-0.5 py-0.5"
                                        >
                                          TOTAL
                                        </th>,
                                      ],
                                    )}
                                  </tr>
                                </thead>

                                <tbody>
                                  <tr>
                                    <td className="border border-black px-1 py-1">
                                      {selectedClass.subject || "Learning Area"}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      N/A
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {selectedClass.section || "-"}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.male.registered}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.female.registered}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center font-bold">
                                      {gsaMetrics.total.registered}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.male.passed}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.female.passed}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center font-bold">
                                      {gsaMetrics.total.passed}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.male.totalGrades.toFixed(0)}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.female.totalGrades.toFixed(0)}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center font-bold">
                                      {gsaMetrics.total.totalGrades.toFixed(0)}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.male.average
                                        ? gsaMetrics.male.average.toFixed(2)
                                        : "0"}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.female.average
                                        ? gsaMetrics.female.average.toFixed(2)
                                        : "0"}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center font-bold">
                                      {gsaMetrics.total.average
                                        ? gsaMetrics.total.average.toFixed(2)
                                        : "0"}
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.male.proficiency}%
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center">
                                      {gsaMetrics.female.proficiency}%
                                    </td>
                                    <td className="border border-black px-1 py-1 text-center font-bold">
                                      {gsaMetrics.total.proficiency}%
                                    </td>
                                  </tr>

                                  {Array.from({ length: 2 }).map(
                                    (_, rowIndex) => (
                                      <tr key={`gsa-empty-${rowIndex}`}>
                                        {Array.from({ length: 18 }).map(
                                          (__, cellIndex) => (
                                            <td
                                              key={`gsa-empty-${rowIndex}-${cellIndex}`}
                                              className="border border-black px-1 py-3"
                                            >
                                              &nbsp;
                                            </td>
                                          ),
                                        )}
                                      </tr>
                                    ),
                                  )}
                                </tbody>
                              </table>

                              <div className="mt-10 grid grid-cols-4 gap-6 text-center text-[9px]">
                                {[
                                  [
                                    "Prepared by:",
                                    selectedTeacherName,
                                    "Teacher I",
                                  ],
                                  [
                                    "Checked by:",
                                    "CHECKER NAME",
                                    "Master Teacher I",
                                  ],
                                  ["Noted:", "NOTED BY NAME", "SHS Dept. Head"],
                                  [
                                    "Approved:",
                                    "PRINCIPAL NAME",
                                    "School Principal",
                                  ],
                                ].map(([label, name, title]) => (
                                  <div key={label}>
                                    <div className="text-left font-semibold">
                                      {label}
                                    </div>
                                    <div className="mt-6 border-t border-black pt-1 text-[10px] font-bold uppercase">
                                      {name}
                                    </div>
                                    <div className="text-[8px]">{title}</div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ) : isMpsPreview && mpsMetrics ? (
                            <div
                              className="box-border p-8 text-[8px] text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div className="grid grid-cols-[180px_minmax(0,1fr)_180px] items-center gap-4">
                                <div className="flex justify-end pr-4">
                                  <img
                                    src={deped}
                                    alt="Department of Education logo"
                                    className="h-16 w-auto object-contain"
                                  />
                                </div>

                                <div className="text-center leading-tight">
                                  <div className="text-[8px]">
                                    Republic of the Philippines
                                  </div>

                                  <div className="text-[11px] font-bold text-[#0038A8]">
                                    DEPARTMENT OF EDUCATION
                                  </div>

                                  <div className="text-[8px]">
                                    {schoolProfile?.region
                                      ? `Region ${schoolProfile.region}`
                                      : "Region XIII — CARAGA"}
                                  </div>

                                  <div className="text-[9px] font-semibold uppercase">
                                    {schoolProfile?.school_name ||
                                      "SCHOOL NAME"}
                                  </div>
                                </div>

                                <div aria-hidden="true" />
                              </div>

                              <div className="mt-3 text-center text-[13px] font-bold leading-tight">
                                REPORT ON MEAN PERCENTAGE SCORE OF LEARNERS IN
                                TERM EXAM
                                <div>PER TERM BY LEARNING AREA</div>
                              </div>

                              <div className="mt-1 text-center text-[9px]">
                                School Year:{" "}
                                {selectedClass.school_year || "____-____"}
                              </div>

                              <div className="mt-5 grid grid-cols-3 items-end text-[9px]">
                                <div>
                                  <b>Grade:</b>{" "}
                                  <span className="ml-1 border-b border-black px-4">
                                    {selectedClass.grade_level || "Grade Level"}
                                  </span>
                                </div>

                                <div className="text-center">
                                  <b>Department:</b>{" "}
                                  <span className="ml-1 border-b border-black px-4 uppercase">
                                    {Number(
                                      String(
                                        selectedClass.grade_level || "",
                                      ).match(/\d+/)?.[0] || 0,
                                    ) >= 11
                                      ? "Senior High School"
                                      : "Junior High School"}
                                  </span>
                                </div>

                                <div className="text-right">
                                  <b>Term:</b>{" "}
                                  <span className="ml-1 border-b border-black px-3 uppercase">
                                    First
                                  </span>
                                </div>
                              </div>

                              <table className="mt-3 w-full table-fixed border-collapse border border-black text-[7px]">
                                <thead>
                                  <tr>
                                    <th
                                      colSpan={2}
                                      className="border border-black px-1 py-2"
                                    >
                                      Grade Level/
                                      <br />
                                      Subjects
                                    </th>

                                    <th
                                      colSpan={2}
                                      className="border border-black px-1 py-2"
                                    >
                                      SHS/ABM/
                                      <br />
                                      SECTION
                                    </th>

                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-2"
                                    >
                                      No. of Learners
                                    </th>

                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-2"
                                    >
                                      Learners with scores 75% of HTS and above
                                    </th>

                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-2"
                                    >
                                      MEAN score of Learners
                                    </th>

                                    <th
                                      colSpan={3}
                                      className="border border-black px-1 py-2"
                                    >
                                      Mean % of Learners Scores
                                    </th>
                                  </tr>

                                  <tr>
                                    {[
                                      "Subjects Handled",
                                      "HTS",
                                      "SHS/ABM",
                                      "SECTION",
                                      "MALE",
                                      "FEMALE",
                                      "TOTAL",
                                      "MALE",
                                      "FEMALE",
                                      "TOTAL",
                                      "MALE",
                                      "FEMALE",
                                      "AVERAGE",
                                      "MALE",
                                      "FEMALE",
                                      "TOTAL",
                                    ].map((label, index) => (
                                      <th
                                        key={`mps-header-${index}`}
                                        className="border border-black px-0.5 py-1.5"
                                      >
                                        {label}
                                      </th>
                                    ))}
                                  </tr>
                                </thead>

                                <tbody>
                                  <tr>
                                    <td
                                      colSpan={10}
                                      className="border border-black px-1 py-1"
                                    >
                                      &nbsp;
                                    </td>

                                    <td
                                      colSpan={3}
                                      className="border border-black px-1 py-1 text-center text-[6px]"
                                    >
                                      Formula: Sum of Scores / No. of Learners
                                    </td>

                                    <td
                                      colSpan={3}
                                      className="border border-black px-1 py-1 text-center text-[6px]"
                                    >
                                      Formula: (Mean Score / HTS) × 100
                                    </td>
                                  </tr>

                                  <tr>
                                    <td className="border border-black px-1 py-2">
                                      {selectedClass.subject || "Learning Area"}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.highestTestScore}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      N/A
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {selectedClass.section || "-"}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.male.registered}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.female.registered}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center font-bold">
                                      {mpsMetrics.total.registered}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.male.passed}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.female.passed}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center font-bold">
                                      {mpsMetrics.total.passed}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.male.meanScore
                                        ? mpsMetrics.male.meanScore.toFixed(2)
                                        : "0"}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.female.meanScore
                                        ? mpsMetrics.female.meanScore.toFixed(2)
                                        : "0"}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center font-bold">
                                      {mpsMetrics.total.meanScore
                                        ? mpsMetrics.total.meanScore.toFixed(2)
                                        : "0"}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.male.meanPercentage.toFixed(
                                        2,
                                      )}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center">
                                      {mpsMetrics.female.meanPercentage.toFixed(
                                        2,
                                      )}
                                    </td>

                                    <td className="border border-black px-1 py-2 text-center font-bold">
                                      {mpsMetrics.total.meanPercentage.toFixed(
                                        2,
                                      )}
                                    </td>
                                  </tr>

                                  {Array.from({ length: 4 }).map(
                                    (_, rowIndex) => (
                                      <tr key={`mps-empty-${rowIndex}`}>
                                        {Array.from({ length: 16 }).map(
                                          (__, cellIndex) => (
                                            <td
                                              key={`mps-empty-${rowIndex}-${cellIndex}`}
                                              className="border border-black px-1 py-3"
                                            >
                                              &nbsp;
                                            </td>
                                          ),
                                        )}
                                      </tr>
                                    ),
                                  )}
                                </tbody>
                              </table>

                              <div className="mt-9 grid grid-cols-4 gap-7 text-center text-[8px]">
                                {[
                                  [
                                    "Prepared by:",
                                    selectedTeacherName,
                                    "Teacher I",
                                  ],
                                  [
                                    "Checked by:",
                                    "CHECKER NAME",
                                    "Master Teacher I",
                                  ],
                                  ["Noted:", "NOTED BY NAME", "SHS Dept. Head"],
                                  [
                                    "Approved:",
                                    "PRINCIPAL NAME",
                                    "School Principal",
                                  ],
                                ].map(([label, name, title]) => (
                                  <div key={label}>
                                    <div className="text-left font-semibold">
                                      {label}
                                    </div>

                                    <div className="mt-7 border-t border-black pt-1 text-[9px] font-bold uppercase">
                                      {name}
                                    </div>

                                    <div className="text-[7px]">{title}</div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ) : isSf5Preview && sf5Metrics ? (
                            <div
                              className="box-border h-full w-full overflow-hidden bg-white px-6 py-5 text-[8px] leading-tight text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div className="mx-auto w-[1040px]">
                                <div className="text-center text-[17px] font-bold leading-none">
                                  School Form 5 (SF 5) Report on Promotion and
                                  Level of Proficiency &amp; Achievement
                                </div>
                                <div className="mt-3 text-center text-[7px] italic">
                                  (This replaces Forms 18-E1, 18-E2, 18A and
                                  List of Graduates)
                                </div>

                                <div
                                  className="mt-2 grid items-stretch text-[8px]"
                                  style={{
                                    gridTemplateColumns:
                                      ADMIN_SF5_COLUMN_WIDTHS.map(
                                        (width) => `${width}fr`,
                                      ).join(" "),
                                    gridTemplateRows: "18px 18px 20px",
                                  }}
                                >
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{
                                      gridColumn: "1 / span 4",
                                      gridRow: 1,
                                    }}
                                  >
                                    Region
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{ gridColumn: "5", gridRow: 1 }}
                                  >
                                    {schoolProfile?.region ||
                                      selectedTeacherProfile?.region ||
                                      ""}
                                  </div>
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{ gridColumn: "6", gridRow: 1 }}
                                  >
                                    Division
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "7 / span 4",
                                      gridRow: 1,
                                    }}
                                  >
                                    {schoolProfile?.division ||
                                      selectedTeacherProfile?.division ||
                                      ""}
                                  </div>

                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{
                                      gridColumn: "1 / span 4",
                                      gridRow: 2,
                                    }}
                                  >
                                    School ID
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "5 / span 2",
                                      gridRow: 2,
                                    }}
                                  >
                                    {schoolProfile?.school_id ||
                                      selectedTeacherProfile?.school_id ||
                                      ""}
                                  </div>
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{
                                      gridColumn: "7 / span 2",
                                      gridRow: 2,
                                    }}
                                  >
                                    School Year
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "9 / span 2",
                                      gridRow: 2,
                                    }}
                                  >
                                    {selectedClass.school_year || ""}
                                  </div>
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{ gridColumn: "11", gridRow: 2 }}
                                  >
                                    Curriculum
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "12 / span 6",
                                      gridRow: 2,
                                    }}
                                  >
                                    K to 12
                                  </div>

                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{
                                      gridColumn: "1 / span 4",
                                      gridRow: 3,
                                    }}
                                  >
                                    School Name
                                  </div>
                                  <div
                                    className="flex items-center justify-center whitespace-nowrap border border-black px-1"
                                    style={{
                                      gridColumn: "5 / span 6",
                                      gridRow: 3,
                                    }}
                                  >
                                    {schoolProfile?.school_name ||
                                      selectedTeacherProfile?.school_name ||
                                      "Agusan del Sur National Science High School"}
                                  </div>
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{ gridColumn: "11", gridRow: 3 }}
                                  >
                                    Grade Level
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "12 / span 4",
                                      gridRow: 3,
                                    }}
                                  >
                                    {selectedClass.grade_level || ""}
                                  </div>
                                  <div
                                    className="flex items-center justify-end pr-1"
                                    style={{
                                      gridColumn: "16 / span 2",
                                      gridRow: 3,
                                    }}
                                  >
                                    Section
                                  </div>
                                  <div
                                    className="flex items-center justify-center border border-black px-1"
                                    style={{
                                      gridColumn: "19 / span 8",
                                      gridRow: 3,
                                    }}
                                  >
                                    {selectedClass.section || ""}
                                  </div>
                                </div>

                                <div className="mx-auto mt-[6px] grid w-[1038px] grid-cols-[720px_310px] items-start gap-2">
                                  <div>
                                    <table className="w-full table-fixed border-collapse bg-white text-[5px]">
                                      <thead>
                                        <tr>
                                          <th className="w-[15%] border border-black px-1 py-4 text-center font-bold">
                                            LRN
                                          </th>
                                          <th className="w-[37%] border border-black px-1 py-4 text-center font-bold leading-tight">
                                            LEARNER&apos;S NAME
                                            <div className="text-[4.4px] font-normal">
                                              (Last Name, First Name, Middle
                                              Name)
                                            </div>
                                          </th>
                                          <th className="w-[12%] border border-black px-1 py-4 text-center font-bold leading-tight">
                                            GENERAL
                                            <br />
                                            AVERAGE
                                          </th>
                                          <th className="w-[16%] border border-black px-1 py-4 text-center font-bold leading-tight">
                                            ACTION TAKEN:
                                            <br />
                                            PROMOTED,
                                            <br />
                                            CONDITIONAL
                                            <br />
                                            or RETAINED
                                          </th>
                                          <th className="w-[20%] border border-black px-1 py-4 text-center font-bold leading-tight">
                                            Did Not Meet Expectations
                                            <div className="text-[4.2px] font-normal">
                                              Incomplete / Failed Learning Areas
                                              as of end of Current School Year
                                            </div>
                                          </th>
                                        </tr>
                                      </thead>

                                      <tbody>
                                        {[
                                          {
                                            key: "male",
                                            label: "MALE",
                                            totalLabel: `===> TOTAL MALE: ${sf5Metrics.maleTotal}`,
                                            learners:
                                              sf5Metrics.learners.filter(
                                                (learner) =>
                                                  learner.sex === "male",
                                              ),
                                          },
                                          {
                                            key: "female",
                                            label: "FEMALE",
                                            totalLabel: `===> TOTAL FEMALE: ${sf5Metrics.femaleTotal}`,
                                            learners:
                                              sf5Metrics.learners.filter(
                                                (learner) =>
                                                  learner.sex === "female",
                                              ),
                                          },
                                        ].map((group) => {
                                          const visibleLearners =
                                            sf5ReportLength === "full"
                                              ? group.learners
                                              : group.learners.slice(0, 10);
                                          return (
                                            <Fragment key={group.key}>
                                              <tr>
                                                <td
                                                  colSpan={5}
                                                  className="border border-black px-1 py-1 font-bold"
                                                >
                                                  {group.label}
                                                </td>
                                              </tr>

                                              {visibleLearners.map(
                                                (learner) => (
                                                  <tr key={learner.id}>
                                                    <td className="border border-black px-1 py-1 leading-tight">
                                                      {learner.lrn}
                                                    </td>
                                                    <td className="border border-black px-1 py-1 uppercase leading-tight">
                                                      {learner.name}
                                                    </td>
                                                    <td className="border border-black px-1 py-1 text-center">
                                                      {learner.generalAverage
                                                        ? Math.round(
                                                            learner.generalAverage,
                                                          )
                                                        : ""}
                                                    </td>
                                                    <td className="border border-black px-1 py-1 text-center font-semibold uppercase leading-tight">
                                                      {learner.action}
                                                    </td>
                                                    <td className="border border-black px-1 py-1 leading-tight">
                                                      {learner.didNotMeet}
                                                    </td>
                                                  </tr>
                                                ),
                                              )}

                                              {visibleLearners.length === 0 && (
                                                <tr>
                                                  <td className="border border-black px-1 py-1">
                                                    &nbsp;
                                                  </td>
                                                  <td className="border border-black px-1 py-1">
                                                    &nbsp;
                                                  </td>
                                                  <td className="border border-black px-1 py-1">
                                                    &nbsp;
                                                  </td>
                                                  <td className="border border-black px-1 py-1">
                                                    &nbsp;
                                                  </td>
                                                  <td className="border border-black px-1 py-1">
                                                    &nbsp;
                                                  </td>
                                                </tr>
                                              )}

                                              <tr>
                                                <td className="border border-black px-1 py-1 text-right">
                                                  0
                                                </td>
                                                <td className="border border-black px-1 py-1 font-bold">
                                                  {group.totalLabel}
                                                </td>
                                                <td className="border border-black bg-[#d9d4ff]" />
                                                <td className="border border-black" />
                                                <td className="border border-black bg-[#d9d4ff]" />
                                              </tr>
                                            </Fragment>
                                          );
                                        })}

                                        <tr>
                                          <td className="border border-black px-1 py-1 text-right">
                                            0
                                          </td>
                                          <td className="border border-black px-1 py-1 font-bold">
                                            ===&gt; COMBINED
                                          </td>
                                          <td className="border border-black bg-[#d9d4ff]" />
                                          <td className="border border-black" />
                                          <td className="border border-black bg-[#d9d4ff]" />
                                        </tr>
                                      </tbody>
                                    </table>
                                  </div>

                                  <div>
                                    <table className="w-full table-fixed border-collapse bg-white text-[4.9px]">
                                      <thead>
                                        <tr>
                                          <th
                                            colSpan={4}
                                            className="border border-black px-1 py-1 text-center font-bold"
                                          >
                                            SUMMARY TABLE
                                          </th>
                                        </tr>
                                        <tr>
                                          <th className="w-[42%] border border-black px-1 py-1 text-center font-bold">
                                            STATUS
                                          </th>
                                          <th className="w-[19%] border border-black px-1 py-1 text-center font-bold">
                                            MALE
                                          </th>
                                          <th className="w-[20%] border border-black px-1 py-1 text-center font-bold">
                                            FEMALE
                                          </th>
                                          <th className="w-[19%] border border-black px-1 py-1 text-center font-bold">
                                            TOTAL
                                          </th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {[
                                          ["PROMOTED", sf5Metrics.promoted],
                                          [
                                            "CONDITIONALLY PROMOTED",
                                            sf5Metrics.conditional,
                                          ],
                                          ["RETAINED", sf5Metrics.retained],
                                        ].map(([label, counts]) => {
                                          const value = counts as {
                                            male: number;
                                            female: number;
                                            total: number;
                                          };
                                          return (
                                            <tr key={String(label)}>
                                              <td className="border border-black px-1 py-2 font-bold">
                                                {String(label)}
                                              </td>
                                              <td className="border border-black px-1 py-2 text-center">
                                                {value.male}
                                              </td>
                                              <td className="border border-black px-1 py-2 text-center">
                                                {value.female}
                                              </td>
                                              <td className="border border-black px-1 py-2 text-center font-bold">
                                                {value.total}
                                              </td>
                                            </tr>
                                          );
                                        })}

                                        <tr>
                                          <th
                                            colSpan={4}
                                            className="border border-black px-1 py-1 text-center font-bold"
                                          >
                                            LEVEL OF PROGRESS AND ACHIEVEMENT
                                          </th>
                                        </tr>
                                        <tr>
                                          <th className="border border-black px-1 py-1 text-left font-bold">
                                            Descriptor &amp;
                                            <br />
                                            Grading
                                          </th>
                                          <th className="border border-black px-1 py-1 text-center font-bold">
                                            MALE
                                          </th>
                                          <th className="border border-black px-1 py-1 text-center font-bold">
                                            FEMALE
                                          </th>
                                          <th className="border border-black px-1 py-1 text-center font-bold">
                                            TOTAL
                                          </th>
                                        </tr>
                                        {sf5Metrics.gradeBands.map((band) => (
                                          <tr key={band.label}>
                                            <td className="border border-black px-1 py-2 font-bold leading-tight">
                                              {band.label}
                                            </td>
                                            <td className="border border-black px-1 py-2 text-center">
                                              {band.male}
                                            </td>
                                            <td className="border border-black px-1 py-2 text-center">
                                              {band.female}
                                            </td>
                                            <td className="border border-black px-1 py-2 text-center font-bold">
                                              {band.total}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>

                                    <div className="mt-6 text-[4.8px] leading-tight">
                                      <div className="font-bold">
                                        Instructions:
                                      </div>
                                      <ol className="mt-2 list-decimal pl-5">
                                        <li>
                                          The SCC shall conduct checking in
                                          their own school, no swapping of SCC
                                          from one school to another is
                                          permitted.
                                        </li>
                                        <li>
                                          The name of SCC members shall be
                                          printed and put their signature on
                                          top.
                                        </li>
                                        <li>
                                          The school head is accountable and
                                          liable for any wrongful entry on the
                                          forms.
                                        </li>
                                        <li>
                                          Only LIS generated SF5 shall be
                                          recognized.
                                        </li>
                                        <li>
                                          This form shall be submitted to the
                                          DCC together with accomplished SFCR1.
                                        </li>
                                      </ol>

                                      <div className="mt-8 font-bold">
                                        PREPARED BY:
                                      </div>
                                      <div className="mt-5 border-t border-black pt-1 text-center">
                                        <b>{selectedTeacherName}</b>
                                        <div>Class Adviser</div>
                                        <div className="italic">
                                          (Name and Signature)
                                        </div>
                                      </div>

                                      <div className="mt-8 font-bold">
                                        CERTIFIED CORRECT &amp; SUBMITTED BY:
                                      </div>
                                      <div className="mt-5 border-t border-black pt-1 text-center">
                                        <b>
                                          {schoolProfile?.principal ||
                                            selectedTeacherProfile?.principal ||
                                            "School Head"}
                                        </b>
                                        <div>School Head / SCC Chair</div>
                                        <div className="italic">
                                          (Name and Signature)
                                        </div>
                                      </div>

                                      <div className="mt-8 font-bold">
                                        REVIEWED BY: SCC Members
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          ) : isSf1Preview && sf1Metrics ? (
                            <Sf1SpreadsheetPreview
                              schoolName={
                                schoolProfile?.school_name ||
                                selectedTeacherProfile?.school_name ||
                                "Agusan del Sur National Science High School"
                              }
                              schoolId={
                                schoolProfile?.school_id ||
                                selectedTeacherProfile?.school_id ||
                                ""
                              }
                              region={
                                schoolProfile?.region ||
                                selectedTeacherProfile?.region ||
                                ""
                              }
                              division={
                                schoolProfile?.division ||
                                selectedTeacherProfile?.division ||
                                ""
                              }
                              schoolYear={selectedClass.school_year || ""}
                              gradeLevel={selectedClass.grade_level || ""}
                              section={selectedClass.section || ""}
                              teacherName={selectedTeacherName}
                              schoolHead={
                                schoolProfile?.principal ||
                                selectedTeacherProfile?.principal ||
                                "School Head"
                              }
                              male={sf1Metrics.male}
                              female={[
                                ...sf1Metrics.female,
                                ...sf1Metrics.unspecified,
                              ]}
                            />
                          ) : isSf1Preview && sf1Metrics ? (
                            <div
                              className="box-border p-6 text-[7px] text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div className="grid grid-cols-[180px_minmax(0,1fr)_180px] items-center gap-3">
                                <div className="flex justify-end pr-4">
                                  <img
                                    src={deped}
                                    alt="Department of Education logo"
                                    className="h-16 w-auto object-contain"
                                  />
                                </div>

                                <div className="text-center leading-tight">
                                  <div className="text-[8px]">
                                    Republic of the Philippines
                                  </div>
                                  <div className="text-[11px] font-bold text-[#0038A8]">
                                    DEPARTMENT OF EDUCATION
                                  </div>
                                  <div className="mt-2 text-[11px] font-bold text-[#0038A8]">
                                    SCHOOL FORM 1 (SF1) SCHOOL REGISTER
                                  </div>
                                </div>

                                <div aria-hidden="true" />
                              </div>

                              <div className="mt-4 grid grid-cols-3 gap-x-8 gap-y-2 text-[8px]">
                                <div>
                                  <b>Region:</b> {schoolProfile?.region || ""}
                                </div>
                                <div>
                                  <b>Division:</b>{" "}
                                  {schoolProfile?.division || ""}
                                </div>
                                <div>
                                  <b>District:</b>{" "}
                                  {schoolProfile?.district || ""}
                                </div>
                                <div>
                                  <b>School Name:</b>{" "}
                                  {schoolProfile?.school_name ||
                                    "Agusan del Sur National Science High School"}
                                </div>
                                <div>
                                  <b>School ID:</b>{" "}
                                  {schoolProfile?.school_id || ""}
                                </div>
                                <div>
                                  <b>School Year:</b>{" "}
                                  <u>
                                    {selectedClass.school_year || "____-____"}
                                  </u>
                                </div>
                                <div className="col-span-3">
                                  <b>Grade &amp; Section:</b>{" "}
                                  <u>
                                    {selectedClass.grade_level || "Grade Level"}{" "}
                                    - {selectedClass.section || "Section"}
                                  </u>
                                </div>
                              </div>

                              <table className="mt-4 w-full table-fixed border-collapse border border-black text-[6px]">
                                <colgroup>
                                  {[
                                    3, 10, 7, 7, 7, 4, 6, 3, 7, 6, 6, 8, 6, 6,
                                    6, 6, 6,
                                  ].map((width, index) => (
                                    <col
                                      key={`sf1-col-${index}`}
                                      style={{ width: `${width}%` }}
                                    />
                                  ))}
                                </colgroup>
                                <thead>
                                  <tr className="bg-[#064AA6] text-white">
                                    {[
                                      "#",
                                      "LRN",
                                      "Last Name",
                                      "First Name",
                                      "Middle Name",
                                      "Sex",
                                      "Birthdate",
                                      "Age",
                                      "Mother Tongue",
                                      "IP/Ethnic",
                                      "Religion",
                                      "Address",
                                      "Father",
                                      "Mother",
                                      "Guardian",
                                      "Contact",
                                      "Remarks",
                                    ].map((label) => (
                                      <th
                                        key={label}
                                        className="border border-black px-0.5 py-2 text-center font-bold"
                                      >
                                        {label}
                                      </th>
                                    ))}
                                  </tr>
                                </thead>

                                <tbody>
                                  <tr className="bg-[#F8DCE8] text-[7px] font-bold text-[#A30F3C]">
                                    <td
                                      colSpan={17}
                                      className="border border-black px-1 py-1"
                                    >
                                      MALE
                                    </td>
                                  </tr>

                                  {sf1Metrics.male
                                    .slice(
                                      0,
                                      sf1ReportLength === "full" ? 30 : 10,
                                    )
                                    .map((learner, index) => (
                                      <tr key={`sf1-male-${learner.id}`}>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {index + 1}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.lrn || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 font-semibold lowercase">
                                          {learner.last_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.first_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.middle_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          M
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.birthdate
                                            ? formatDate(learner.birthdate)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.age}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.address || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.father_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.mother_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.guardian || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.contact_number || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                      </tr>
                                    ))}

                                  {Array.from({
                                    length: sf1ReportLength === "full" ? 6 : 2,
                                  }).map((_, rowIndex) => (
                                    <tr key={`sf1-male-empty-${rowIndex}`}>
                                      {Array.from({ length: 17 }).map(
                                        (__, cellIndex) => (
                                          <td
                                            key={`sf1-male-empty-${rowIndex}-${cellIndex}`}
                                            className="border border-black px-0.5 py-2"
                                          >
                                            &nbsp;
                                          </td>
                                        ),
                                      )}
                                    </tr>
                                  ))}

                                  <tr className="bg-[#F8DCE8] text-[7px] font-bold text-[#A30F3C]">
                                    <td
                                      colSpan={17}
                                      className="border border-black px-1 py-1"
                                    >
                                      FEMALE
                                    </td>
                                  </tr>

                                  {sf1Metrics.female
                                    .slice(
                                      0,
                                      sf1ReportLength === "full" ? 30 : 10,
                                    )
                                    .map((learner, index) => (
                                      <tr key={`sf1-female-${learner.id}`}>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {index + 1}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.lrn || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 font-semibold lowercase">
                                          {learner.last_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.first_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.middle_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          F
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.birthdate
                                            ? formatDate(learner.birthdate)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.age}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.address || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.father_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.mother_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.guardian || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.contact_number || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                      </tr>
                                    ))}

                                  {Array.from({
                                    length: sf1ReportLength === "full" ? 6 : 2,
                                  }).map((_, rowIndex) => (
                                    <tr key={`sf1-female-empty-${rowIndex}`}>
                                      {Array.from({ length: 17 }).map(
                                        (__, cellIndex) => (
                                          <td
                                            key={`sf1-female-empty-${rowIndex}-${cellIndex}`}
                                            className="border border-black px-0.5 py-2"
                                          >
                                            &nbsp;
                                          </td>
                                        ),
                                      )}
                                    </tr>
                                  ))}

                                  {sf1Metrics.unspecified.map(
                                    (learner, index) => (
                                      <tr key={`sf1-unspecified-${learner.id}`}>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {index + 1}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.lrn || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 font-semibold lowercase">
                                          {learner.last_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.first_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 lowercase">
                                          {learner.middle_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          -
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.birthdate
                                            ? formatDate(learner.birthdate)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1 text-center">
                                          {learner.age}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1" />
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.address || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.father_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.mother_name || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.guardian || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1">
                                          {learner.contact_number || ""}
                                        </td>
                                        <td className="border border-black px-0.5 py-1" />
                                      </tr>
                                    ),
                                  )}

                                  <tr className="font-bold">
                                    <td
                                      colSpan={17}
                                      className="border border-black px-1 py-1 text-center"
                                    >
                                      TOTAL: {sf1Metrics.total} (M:{" "}
                                      {sf1Metrics.male.length} / F:{" "}
                                      {sf1Metrics.female.length})
                                    </td>
                                  </tr>
                                </tbody>
                              </table>

                              <div className="mt-8 grid grid-cols-2 gap-8 text-center text-[8px]">
                                <div>
                                  <div className="border-t border-black pt-1 font-semibold">
                                    {selectedTeacherName}
                                  </div>
                                  <div>Class Adviser</div>
                                </div>
                                <div>
                                  <div className="border-t border-black pt-1 font-semibold">
                                    School Head
                                  </div>
                                </div>
                              </div>
                            </div>
                          ) : isSf2Preview && sf2Metrics ? (
                            <div
                              className="box-border p-8 text-[6px] text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div className="text-center">
                                <div className="text-[16px] font-bold leading-tight">
                                  School Form 2 (SF2) Daily Attendance Report of
                                  Learners
                                </div>
                                <div className="mt-1 text-[7px] italic">
                                  (This replaces Form 1, Form 2 and STS Form 4 -
                                  Absenteeism and Dropout Profile)
                                </div>
                              </div>

                              <div className="mt-5 grid grid-cols-[1fr_1fr_1.45fr] gap-x-5 gap-y-1 text-[7px]">
                                <div className="flex items-center gap-2">
                                  <b>School ID</b>
                                  <span className="h-5 flex-1 border border-black px-2 py-1 text-center">
                                    {schoolProfile?.school_id || ""}
                                  </span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <b>School Year</b>
                                  <span className="h-5 flex-1 border border-black px-2 py-1 text-center">
                                    {selectedClass.school_year || "____-____"}
                                  </span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <b>Report for the Month of</b>
                                  <span className="h-5 flex-1 border border-[#19B45B] px-2 py-1 text-center font-semibold uppercase">
                                    {sf2MonthName}
                                  </span>
                                </div>

                                <div className="col-span-2 flex items-center gap-2">
                                  <b>Name of School</b>
                                  <span className="h-5 flex-1 border border-black px-2 py-1 text-center">
                                    {schoolProfile?.school_name ||
                                      "Agusan del Sur National Science High School"}
                                  </span>
                                </div>
                                <div className="grid grid-cols-2 gap-3">
                                  <div className="flex items-center gap-2">
                                    <b>Grade Level</b>
                                    <span className="h-5 flex-1 border border-black px-2 py-1 text-center">
                                      {selectedClass.grade_level || "-"}
                                    </span>
                                  </div>
                                  <div className="flex items-center gap-2">
                                    <b>Section</b>
                                    <span className="h-5 flex-1 border border-black px-2 py-1 text-center">
                                      {selectedClass.section || "-"}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              <table className="mt-1 w-full table-fixed border-collapse border border-black text-[5px]">
                                <thead>
                                  <tr>
                                    <th
                                      rowSpan={3}
                                      className="w-[3%] border border-black px-0.5 py-1"
                                    >
                                      No.
                                    </th>
                                    <th
                                      rowSpan={3}
                                      className="w-[20%] border border-black px-1 py-1"
                                    >
                                      NAME
                                      <div className="font-normal">
                                        (Last Name, First Name, Middle Name)
                                      </div>
                                    </th>
                                    <th
                                      colSpan={sf2Metrics.schoolDays.length}
                                      className="border border-black px-1 py-1 font-normal"
                                    >
                                      (1st row for date)
                                    </th>
                                    <th
                                      colSpan={2}
                                      rowSpan={2}
                                      className="w-[8%] border border-black px-1 py-1"
                                    >
                                      Total for the Month
                                    </th>
                                    <th
                                      rowSpan={3}
                                      className="w-[16%] border border-black px-1 py-1"
                                    >
                                      REMARKS (if DROPPED OUT, state reason; if
                                      TRANSFERRED IN/OUT, write the school)
                                    </th>
                                  </tr>
                                  <tr>
                                    {sf2Metrics.schoolDays.map((day) => (
                                      <th
                                        key={`sf2-date-${day.key}`}
                                        className="border border-black px-0 py-0.5 text-center"
                                      >
                                        {day.day}
                                      </th>
                                    ))}
                                  </tr>
                                  <tr>
                                    {sf2Metrics.schoolDays.map((day) => (
                                      <th
                                        key={`sf2-weekday-${day.key}`}
                                        className="border border-black px-0 py-0.5 text-center"
                                      >
                                        {day.weekday}
                                      </th>
                                    ))}
                                    <th className="border border-black px-0.5 py-1">
                                      ABSENT
                                    </th>
                                    <th className="border border-black px-0.5 py-1">
                                      PRESENT
                                    </th>
                                  </tr>
                                </thead>

                                <tbody>
                                  {sf2Metrics.learners.map((learner, index) => (
                                    <tr key={`sf2-learner-${learner.id}`}>
                                      <td className="border border-black px-0.5 py-1 text-center">
                                        {index + 1}.
                                      </td>
                                      <td className="border border-black px-1 py-1 font-semibold uppercase">
                                        {learner.name}
                                      </td>
                                      {learner.dayStatuses.map((day) => (
                                        <td
                                          key={`sf2-mark-${learner.id}-${day.key}`}
                                          className="border border-black px-0 py-1 text-center"
                                        >
                                          {day.mark}
                                        </td>
                                      ))}
                                      <td className="border border-black px-0.5 py-1 text-center">
                                        {learner.absent || ""}
                                      </td>
                                      <td className="border border-black px-0.5 py-1 text-center">
                                        {learner.present || ""}
                                      </td>
                                      <td className="border border-black px-1 py-1" />
                                    </tr>
                                  ))}

                                  {Array.from({
                                    length: Math.max(
                                      0,
                                      8 - sf2Metrics.learners.length,
                                    ),
                                  }).map((_, rowIndex) => (
                                    <tr key={`sf2-empty-${rowIndex}`}>
                                      <td className="border border-black px-0.5 py-1 text-center">
                                        {sf2Metrics.learners.length +
                                          rowIndex +
                                          1}
                                        .
                                      </td>
                                      <td className="border border-black px-1 py-1">
                                        &nbsp;
                                      </td>
                                      {sf2Metrics.schoolDays.map((day) => (
                                        <td
                                          key={`sf2-empty-${rowIndex}-${day.key}`}
                                          className="border border-black px-0 py-1"
                                        />
                                      ))}
                                      <td className="border border-black px-0.5 py-1" />
                                      <td className="border border-black px-0.5 py-1" />
                                      <td className="border border-black px-1 py-1" />
                                    </tr>
                                  ))}

                                  {[
                                    {
                                      label: "MALE | TOTAL Per Day",
                                      values: sf2Metrics.dailyTotals.map(
                                        (item) => item.male,
                                      ),
                                    },
                                    {
                                      label: "FEMALE | TOTAL Per Day",
                                      values: sf2Metrics.dailyTotals.map(
                                        (item) => item.female,
                                      ),
                                    },
                                    {
                                      label: "Combined TOTAL Per Day",
                                      values: sf2Metrics.dailyTotals.map(
                                        (item) => item.combined,
                                      ),
                                    },
                                  ].map((row) => (
                                    <tr key={row.label} className="font-bold">
                                      <td className="border border-black px-0.5 py-1 text-center" />
                                      <td className="border border-black px-1 py-1">
                                        {row.label}
                                      </td>
                                      {row.values.map((value, dayIndex) => (
                                        <td
                                          key={`${row.label}-${dayIndex}`}
                                          className="border border-black px-0 py-1 text-center"
                                        >
                                          {value || ""}
                                        </td>
                                      ))}
                                      <td className="border border-black px-0.5 py-1 text-center" />
                                      <td className="border border-black px-0.5 py-1 text-center" />
                                      <td className="border border-black px-1 py-1" />
                                    </tr>
                                  ))}
                                </tbody>
                              </table>

                              <div className="grid grid-cols-[1.15fr_0.95fr_1.1fr] border-x border-b border-black text-[5px] leading-relaxed">
                                <div className="min-h-[245px] border-r border-black p-2">
                                  <div className="font-bold">GUIDELINES:</div>
                                  <ol className="mt-1 list-decimal space-y-1 pl-4">
                                    <li>
                                      The attendance shall be accomplished
                                      daily.
                                    </li>
                                    <li>
                                      Dates shall be written in the columns
                                      after Learner&apos;s Name.
                                    </li>
                                    <li>
                                      Compute percentage of enrolment, average
                                      daily attendance, and attendance
                                      percentage.
                                    </li>
                                    <li>
                                      At the end of the month, submit this
                                      report for recording.
                                    </li>
                                    <li>
                                      Record learners with prolonged absences
                                      for follow-up.
                                    </li>
                                    <li>
                                      Attendance performance shall be reflected
                                      in Forms 137 and 138.
                                    </li>
                                  </ol>
                                </div>

                                <div className="min-h-[245px] border-r border-black p-2">
                                  <div className="font-bold">
                                    1. CODES FOR CHECKING ATTENDANCE
                                  </div>
                                  <div>
                                    (blank) - Present; X - Absent; L - Late; E -
                                    Excused
                                  </div>
                                  <div className="mt-2 font-bold">
                                    2. REASONS/CAUSES FOR NON-ATTENDANCE
                                  </div>
                                  <div>a. Domestic-related factors</div>
                                  <div>b. Individual-related factors</div>
                                  <div>c. School-related factors</div>
                                  <div>d. Geographic/Environmental factors</div>
                                  <div>e. Financial-related factors</div>
                                </div>

                                <div className="min-h-[245px] p-2">
                                  <div className="grid grid-cols-[1fr_36px_36px_44px] border border-black text-center">
                                    <div className="border-r border-black px-1 py-1 text-left font-bold">
                                      Summary - {sf2MonthName}
                                    </div>
                                    <div className="border-r border-black px-1 py-1 font-bold">
                                      M
                                    </div>
                                    <div className="border-r border-black px-1 py-1 font-bold">
                                      F
                                    </div>
                                    <div className="px-1 py-1 font-bold">
                                      TOTAL
                                    </div>

                                    {[
                                      [
                                        "Registered Learners as of end of month",
                                        sf2Metrics.maleEnrollment,
                                        sf2Metrics.femaleEnrollment,
                                        sf2Metrics.totalEnrollment,
                                      ],
                                      [
                                        "Average Daily Attendance",
                                        "",
                                        "",
                                        sf2Metrics.averageDailyAttendance.toFixed(
                                          2,
                                        ),
                                      ],
                                      [
                                        "Percentage of Attendance for the month",
                                        "",
                                        "",
                                        `${sf2Metrics.attendancePercentage.toFixed(2)}%`,
                                      ],
                                      ["Dropped out", "", "", ""],
                                      ["Transferred out", "", "", ""],
                                      ["Transferred in", "", "", ""],
                                    ].flatMap((summaryRow, rowIndex) =>
                                      summaryRow.map((cell, cellIndex) => (
                                        <div
                                          key={`sf2-summary-${rowIndex}-${cellIndex}`}
                                          className={`${
                                            cellIndex < 3 ? "border-r" : ""
                                          } border-t border-black px-1 py-1 ${
                                            cellIndex === 0
                                              ? "text-left"
                                              : "text-center"
                                          }`}
                                        >
                                          {cell}
                                        </div>
                                      )),
                                    )}
                                  </div>

                                  <div className="mt-3 italic">
                                    I certify that this is a true and correct
                                    report.
                                  </div>
                                  <div className="mt-8 border-t border-black pt-1 text-center">
                                    {selectedTeacherName}
                                    <div>
                                      (Signature of Adviser over Printed Name)
                                    </div>
                                  </div>
                                  <div className="mt-4">Attested by:</div>
                                  <div className="mt-7 border-t border-black pt-1 text-center">
                                    School Head
                                    <div>
                                      (Signature of School Head over Printed
                                      Name)
                                    </div>
                                  </div>
                                </div>
                              </div>

                              <div className="mt-1 text-center text-[6px] font-bold">
                                Generated thru LIS
                              </div>
                            </div>
                          ) : isSf9Preview && sf9Metrics ? (
                            <div
                              className="box-border p-3 text-[8px] text-black"
                              style={{
                                fontFamily: "Arial, Helvetica, sans-serif",
                              }}
                            >
                              <div
                                className="border border-black p-3"
                                style={{
                                  minHeight: `${selectedPaper.height - 24}px`,
                                }}
                              >
                                <div className="flex items-center justify-center gap-4 text-center">
                                  <img
                                    src={deped}
                                    alt="Department of Education logo"
                                    className="h-9 w-auto object-contain"
                                  />
                                  <div className="text-[12px] font-bold">
                                    REPORT ON LEARNER&apos;S PROGRESS (SF9)
                                  </div>
                                </div>

                                <div className="mt-4 grid grid-cols-[1fr_0.42fr] gap-x-3 gap-y-2 text-[8px]">
                                  <div className="flex items-end gap-2">
                                    <b>Name:</b>
                                    <span className="min-h-5 flex-1 border-b border-black px-1 font-semibold lowercase">
                                      {sf9Metrics.learner
                                        ? studentName(sf9Metrics.learner)
                                        : ""}
                                    </span>
                                  </div>
                                  <div className="flex items-end gap-2">
                                    <b>LRN:</b>
                                    <span className="min-h-5 flex-1 border-b border-black px-1">
                                      {sf9Metrics.learner?.lrn || ""}
                                    </span>
                                  </div>
                                  <div className="flex items-end gap-2">
                                    <b>Grade &amp; Section:</b>
                                    <span className="min-h-5 flex-1 border-b border-black px-1">
                                      {selectedClass.grade_level ||
                                        "Grade Level"}{" "}
                                      - {selectedClass.section || "Section"}
                                    </span>
                                  </div>
                                  <div className="flex items-end gap-2">
                                    <b>School Year:</b>
                                    <span className="min-h-5 flex-1 border-b border-black px-1">
                                      {selectedClass.school_year || ""}
                                    </span>
                                  </div>
                                </div>

                                <table className="mt-4 w-full table-fixed border-collapse border border-black text-[8px]">
                                  <colgroup>
                                    <col style={{ width: "16%" }} />
                                    <col style={{ width: "23%" }} />
                                    <col style={{ width: "23%" }} />
                                    <col style={{ width: "23%" }} />
                                    <col style={{ width: "5%" }} />
                                    <col style={{ width: "10%" }} />
                                  </colgroup>
                                  <thead>
                                    <tr className="bg-[#0B5A96] text-white">
                                      <th className="border border-black px-1 py-1.5 text-left">
                                        Learning Area
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        Term 1
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        Term 2
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        Term 3
                                      </th>
                                      <th className="border border-black bg-[#0A805B] px-1 py-1.5">
                                        Final
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        Remarks
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {sf9Metrics.subjects.map((row) => (
                                      <tr key={row.subject}>
                                        <td className="border border-black px-1 py-1 font-semibold">
                                          {row.subject}
                                        </td>
                                        <td className="border border-black px-1 py-1 text-center">
                                          {typeof row.term1 === "number"
                                            ? Math.round(row.term1)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-1 py-1 text-center">
                                          {typeof row.term2 === "number"
                                            ? Math.round(row.term2)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-1 py-1 text-center">
                                          {typeof row.term3 === "number"
                                            ? Math.round(row.term3)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-1 py-1 text-center font-semibold">
                                          {typeof row.final === "number"
                                            ? Math.round(row.final)
                                            : ""}
                                        </td>
                                        <td className="border border-black px-1 py-1 text-center text-[7px] font-semibold">
                                          {row.remarks}
                                        </td>
                                      </tr>
                                    ))}
                                    {Array.from({
                                      length: Math.max(
                                        0,
                                        11 - sf9Metrics.subjects.length,
                                      ),
                                    }).map((_, rowIndex) => (
                                      <tr key={`sf9-empty-${rowIndex}`}>
                                        {Array.from({ length: 6 }).map(
                                          (__, cellIndex) => (
                                            <td
                                              key={`sf9-empty-${rowIndex}-${cellIndex}`}
                                              className="border border-black px-1 py-1.5"
                                            >
                                              &nbsp;
                                            </td>
                                          ),
                                        )}
                                      </tr>
                                    ))}
                                    <tr className="font-bold">
                                      <td
                                        colSpan={4}
                                        className="border border-black px-1 py-1 text-center"
                                      >
                                        GENERAL AVERAGE
                                      </td>
                                      <td className="border border-black bg-[#D7F5E6] px-1 py-1 text-center">
                                        {typeof sf9Metrics.generalAverage ===
                                        "number"
                                          ? Math.round(
                                              sf9Metrics.generalAverage,
                                            )
                                          : ""}
                                      </td>
                                      <td className="border border-black px-1 py-1 text-center">
                                        {typeof sf9Metrics.generalAverage ===
                                        "number"
                                          ? sf9Metrics.generalAverage >= 75 &&
                                            sf9Metrics.generalAverage <= 100
                                            ? "PROMOTED"
                                            : "FAILED"
                                          : ""}
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>

                                <div className="mt-2 text-[8px] font-semibold">
                                  REPORT ON LEARNER&apos;S OBSERVED VALUES
                                </div>

                                <table className="mt-1 w-full table-fixed border-collapse border border-black text-[7px]">
                                  <colgroup>
                                    <col style={{ width: "12%" }} />
                                    <col style={{ width: "64%" }} />
                                    <col style={{ width: "8%" }} />
                                    <col style={{ width: "8%" }} />
                                    <col style={{ width: "8%" }} />
                                  </colgroup>
                                  <thead>
                                    <tr className="bg-[#0B5A96] text-white">
                                      <th className="border border-black px-1 py-1.5">
                                        Core Value
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        Behavior Statement
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        T1
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        T2
                                      </th>
                                      <th className="border border-black px-1 py-1.5">
                                        T3
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {[
                                      [
                                        "Maka-Diyos",
                                        "Expresses one’s spiritual beliefs while respecting the spiritual beliefs of others",
                                      ],
                                      [
                                        "",
                                        "Shows adherence to ethical principles by upholding truth",
                                      ],
                                      [
                                        "Makatao",
                                        "Is sensitive to individual, social, and cultural differences",
                                      ],
                                      [
                                        "",
                                        "Demonstrates contributions toward solidarity",
                                      ],
                                      [
                                        "Maka-Kalikasan",
                                        "Cares for the environment and utilizes resources wisely",
                                      ],
                                      [
                                        "Makabansa",
                                        "Demonstrates pride in being a Filipino; exercises the rights and responsibilities of a Filipino citizen",
                                      ],
                                      [
                                        "",
                                        "Demonstrates appropriate behavior in carrying out activities in school, community, and country",
                                      ],
                                    ].map(([coreValue, behavior], rowIndex) => (
                                      <tr key={`sf9-value-${rowIndex}`}>
                                        <td className="border border-black px-1 py-1 font-semibold">
                                          {coreValue}
                                        </td>
                                        <td className="border border-black px-1 py-1">
                                          {behavior}
                                        </td>
                                        {[1, 2, 3].map((term) => (
                                          <td
                                            key={`sf9-value-${rowIndex}-${term}`}
                                            className="border border-black px-1 py-1"
                                          >
                                            <div className="flex items-center justify-center gap-2">
                                              <span className="text-[#64748B]">
                                                —
                                              </span>
                                              <ChevronDown className="size-3 text-[#17345F]" />
                                            </div>
                                          </td>
                                        ))}
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>

                                <div className="mt-5 grid grid-cols-2 gap-16 text-center text-[7px]">
                                  <div>
                                    <div className="border-t border-black pt-1">
                                      {selectedTeacherName}
                                    </div>
                                    <div>Class Adviser</div>
                                  </div>
                                  <div>
                                    <div className="border-t border-black pt-1">
                                      &nbsp;
                                    </div>
                                    <div>Parent/Guardian</div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          ) : isSf9NewPreview && sf9Metrics ? (
                            <div
                              className="mx-auto box-border bg-white text-[9.5px] text-black"
                              style={{
                                width: isShsSf9NewPreview ? "1124px" : "1224px",
                                minHeight: isShsSf9NewPreview ? "795px" : "800px",
                                height: isShsSf9NewPreview ? "795px" : "800px",
                                padding: "14px",
                                fontFamily:
                                  '"Bookman Old Style", "Times New Roman", serif',
                              }}
                            >
                              <div
                                className="h-full"
                                style={{ padding: "12px 14px" }}
                              >
                                <div
                                  className="grid h-full min-h-0 gap-3 overflow-hidden"
                                  style={{
                                    gridTemplateColumns: "48.5fr 51.5fr",
                                  }}
                                >
                                  <section className="flex min-w-0 flex-col">
                                    <div className="grid grid-cols-[82px_1fr_82px] items-start gap-2">
                                      <img
                                        src={deped}
                                        alt="DepEd logo"
                                        className="h-[76px] w-[76px] object-contain"
                                      />

                                      <div className="text-center text-[9.5px] leading-[1.18]">
                                        <div>Republic of the Philippines</div>
                                        <div>Department of Education</div>
                                        <div>
                                          {schoolProfile?.region ||
                                            "CARAGA Region"}
                                        </div>
                                        <div className="font-bold uppercase">
                                          SCHOOLS DIVISION OFFICE OF{" "}
                                          {String(
                                            schoolProfile?.division ||
                                              "AGUSAN DEL SUR",
                                          ).replace(
                                            /^SCHOOLS DIVISION OFFICE OF\s*/i,
                                            "",
                                          )}
                                        </div>
                                        <div>
                                          {schoolProfile?.district ||
                                            "Prosperidad District"}
                                        </div>
                                        <div>
                                          {schoolProfile?.city_municipality_province ||
                                            schoolProfile?.municipality ||
                                            "Prosperidad, Agusan del Sur"}
                                        </div>
                                        <div className="mt-1 font-semibold uppercase">
                                          {schoolProfile?.school_name ||
                                            "AGUSAN DEL SUR NATIONAL SCIENCE HIGH SCHOOL"}
                                        </div>
                                      </div>

                                      <img
                                        src={logo}
                                        alt="School logo"
                                        className="h-[80px] w-[80px] object-contain"
                                      />
                                    </div>

                                    <div className="mt-3 text-center text-[16px] font-bold">
                                      LEARNER&apos;S PERFORMANCE REPORT
                                    </div>

                                    <div className="mt-1 grid grid-cols-[96px_1fr] items-end gap-2 text-[10.5px]">
                                      <div className="font-bold">
                                        School Year
                                      </div>
                                      <div className="min-h-[18px] border-b border-black px-2 text-center">
                                        {selectedClass.school_year || ""}
                                      </div>
                                    </div>

                                    <div className="mt-3 grid grid-cols-[92px_1fr_50px_86px] items-end gap-x-2 gap-y-1 text-[10.5px]">
                                      <div className="font-bold">Name</div>
                                      <div className="min-h-[18px] border-b border-black px-1 font-semibold uppercase">
                                        {sf9Metrics.learner
                                          ? [
                                              sf9Metrics.learner.last_name,
                                              sf9Metrics.learner.first_name,
                                              sf9Metrics.learner.middle_name,
                                            ]
                                              .filter(Boolean)
                                              .join(", ")
                                          : ""}
                                      </div>
                                      <div className="font-bold">Age</div>
                                      <div className="min-h-[18px] border-b border-black px-1 text-center">
                                        {sf9Metrics.learner?.birthdate
                                          ? Math.max(
                                              0,
                                              new Date().getFullYear() -
                                                new Date(
                                                  sf9Metrics.learner.birthdate,
                                                ).getFullYear(),
                                            )
                                          : ""}
                                      </div>

                                      <div className="font-bold">LRN</div>
                                      <div className="min-h-[18px] border-b border-black px-1">
                                        {sf9Metrics.learner?.lrn || ""}
                                      </div>
                                      <div className="font-bold">Sex</div>
                                      <div className="min-h-[18px] border-b border-black px-1 text-center capitalize">
                                        {sf9Metrics.learner?.sex || ""}
                                      </div>

                                      <div className="font-bold">
                                        Track (SHS only)
                                      </div>
                                      <div className="min-h-[18px] border-b border-black px-1">
                                        {selectedClass.track_shs || ""}
                                      </div>
                                      <div className="font-bold">Grade</div>
                                      <div className="min-h-[18px] border-b border-black px-1 text-center">
                                        {selectedClass.grade_level || ""}
                                      </div>

                                      <div />
                                      <div />
                                      <div className="font-bold">Section</div>
                                      <div className="min-h-[18px] border-b border-black px-1 text-center">
                                        {selectedClass.section || ""}
                                      </div>
                                    </div>

                                    <div className="mt-3 text-[9.5px] italic leading-[1.18]">
                                      <div className="font-bold">
                                        Dear Parents,
                                      </div>
                                      <div className="indent-8 text-justify">
                                        This Performance Report shows the
                                        ability and progress your child has made
                                        in the different learning areas as well
                                        as his / her core values. The school
                                        welcomes you should you desire to know
                                        more about your child&apos;s progress.
                                      </div>
                                    </div>

                                    <div className="mt-3 text-center text-[11.5px] font-bold">
                                      LEARNING PROGRESS AND ACHIEVEMENT
                                    </div>
                                    <table className="mt-1 w-full table-fixed border-collapse border border-black text-[8.4px]">
                                      <colgroup>
                                        <col style={{ width: "36%" }} />
                                        <col style={{ width: "8%" }} />
                                        <col style={{ width: "8%" }} />
                                        <col style={{ width: "8%" }} />
                                        <col style={{ width: "9%" }} />
                                        <col style={{ width: "14%" }} />
                                        <col style={{ width: "17%" }} />
                                      </colgroup>
                                      <thead>
                                        <tr className="font-bold">
                                          <th rowSpan={2} className="border border-black px-1 py-1">Learning Areas</th>
                                          <th colSpan={3} className="border border-black px-1 py-0.5">TERM</th>
                                          <th rowSpan={2} className="border border-black px-1 py-1">Units</th>
                                          <th rowSpan={2} className="border border-black px-1 py-1">Final Grade</th>
                                          <th rowSpan={2} className="border border-black px-1 py-1">Remarks</th>
                                        </tr>
                                        <tr>
                                          {[1, 2, 3].map((term) => (
                                            <th key={term} className="border border-black py-0.5">{term}</th>
                                          ))}
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {isGrade12Sf9NewPreview ? (
                                          <>
                                            <tr>
                                              <td colSpan={7} className="border border-black bg-black/15 px-1 py-1 font-bold italic">Core Subjects</td>
                                            </tr>
                                            {sf9Metrics.subjects.filter((row) => row.group === "core").map((row) => (
                                              <tr key={`sf9-new-${row.subject}`} className="h-[18px]">
                                                <td className="border border-black px-1">{row.subject}</td>
                                                {[row.term1, row.term2, row.term3].map((score, index) => {
                                                  const term = String(index + 1);
                                                  const unavailable = Boolean(row.assignedTerm && row.assignedTerm !== term);
                                                  return (
                                                    <td key={`${row.subject}-${term}`} className={`border border-black text-center ${unavailable ? "bg-black/20" : ""}`}>
                                                      {!unavailable && typeof score === "number" ? Math.round(score) : ""}
                                                    </td>
                                                  );
                                                })}
                                                <td className="border border-black text-center">{row.units ?? ""}</td>
                                                <td className="border border-black text-center font-semibold">{typeof row.final === "number" ? Math.round(row.final) : ""}</td>
                                                <td className="border border-black text-center text-[6px]">{row.remarks}</td>
                                              </tr>
                                            ))}
                                            <tr>
                                              <td colSpan={7} className="border border-black bg-black/15 px-1 py-1 font-bold">Applied and Specialized Subjects</td>
                                            </tr>
                                            {sf9Metrics.subjects.filter((row) => row.group === "applied").map((row) => (
                                              <tr key={`sf9-new-${row.subject}`} className="h-[18px]">
                                                <td className="border border-black px-1">{row.subject}</td>
                                                {[row.term1, row.term2, row.term3].map((score, index) => {
                                                  const term = String(index + 1);
                                                  const unavailable = Boolean(row.assignedTerm && row.assignedTerm !== term);
                                                  return (
                                                    <td key={`${row.subject}-${term}`} className={`border border-black text-center ${unavailable ? "bg-black/20" : ""}`}>
                                                      {!unavailable && typeof score === "number" ? Math.round(score) : ""}
                                                    </td>
                                                  );
                                                })}
                                                <td className="border border-black text-center">{row.units ?? ""}</td>
                                                <td className="border border-black text-center font-semibold">{typeof row.final === "number" ? Math.round(row.final) : ""}</td>
                                                <td className="border border-black text-center text-[6px]">{row.remarks}</td>
                                              </tr>
                                            ))}
                                          </>
                                        ) : (
                                          sf9Metrics.subjects.slice(0, 10).map((row) => (
                                            <tr key={`sf9-new-${row.subject}`} className="h-[20px]">
                                              <td className="border border-black px-1">{row.subject}</td>
                                              {[row.term1, row.term2, row.term3].map((score, index) => (
                                                <td key={`${row.subject}-${index}`} className="border border-black text-center">
                                                  {typeof score === "number" ? Math.round(score) : ""}
                                                </td>
                                              ))}
                                              <td className="border border-black text-center">{row.units ?? ""}</td>
                                              <td className="border border-black text-center font-semibold">{typeof row.final === "number" ? Math.round(row.final) : ""}</td>
                                              <td className="border border-black text-center text-[6px]">{row.remarks}</td>
                                            </tr>
                                          ))
                                        )}
                                        <tr className="h-[22px] font-bold">
                                          <td colSpan={5} className="border border-black px-1 text-center italic">GENERAL AVERAGE</td>
                                          <td className="border border-black text-center">
                                            {typeof sf9Metrics.generalAverage === "number" ? Math.round(sf9Metrics.generalAverage) : ""}
                                          </td>
                                          <td className="border border-black text-center text-[6px]">
                                            {typeof sf9Metrics.generalAverage === "number"
                                              ? sf9Metrics.generalAverage >= 75 && sf9Metrics.generalAverage <= 100
                                                ? "PROMOTED"
                                                : "FAILED"
                                              : ""}
                                          </td>
                                        </tr>
                                      </tbody>
                                    </table>

                                    <div className="mt-3 grid grid-cols-[1fr_1.6fr] items-center gap-3">
                                      <div className="text-[9px]">
                                        Performance Descriptors
                                      </div>
                                      <table className="w-full border-collapse text-center text-[6px]">
                                        <thead>
                                          <tr>
                                            <th className="border border-black">
                                              Grading Scale
                                            </th>
                                            <th className="border border-black">
                                              Description
                                            </th>
                                            <th className="border border-black">
                                              Remarks
                                            </th>
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {[
                                            ["90 - 100", "Advancing", "Passed"],
                                            [
                                              "80 - 89",
                                              "Benchmarking",
                                              "Passed",
                                            ],
                                            ["75 - 79", "Connecting", "Passed"],
                                            ["65 - 74", "Developing", "Failed"],
                                            ["0 - 64", "Emerging", "Failed"],
                                          ].map((row) => (
                                            <tr key={row[0]}>
                                              {row.map((cell) => (
                                                <td
                                                  key={cell}
                                                  className="border border-black py-0.5"
                                                >
                                                  {cell}
                                                </td>
                                              ))}
                                            </tr>
                                          ))}
                                        </tbody>
                                      </table>
                                    </div>
                                  </section>

                                  <section className="flex min-w-0 max-w-full flex-col overflow-hidden px-0.5 [&>*]:max-w-full">
                                    <div className="text-center text-[13px] font-bold">
                                      ATTENDANCE RECORD
                                    </div>

                                    <table className="mt-1 w-full table-fixed border-collapse border border-black text-[6px]">
                                      <colgroup>
                                        <col style={{ width: "15%" }} />
                                        {[
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
                                        ].map((month) => (
                                          <col
                                            key={month}
                                            style={{ width: "6.35%" }}
                                          />
                                        ))}
                                        <col style={{ width: "5%" }} />
                                      </colgroup>
                                      <thead>
                                        <tr className="h-[58px]">
                                          <th className="border-t border-l border-b border-r border-black px-1 text-[9px] font-bold">
                                            Month
                                          </th>

                                          {[
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
                                          ].map((month) => (
                                            <th
                                              key={month}
                                              className="relative overflow-visible border-y border-black p-0"
                                            >
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

                                            {[
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
                                            ].map((month, index) => (
                                              <td
                                                key={month}
                                                className={
                                                  index === 0
                                                    ? "border-y border-r border-black"
                                                    : index === 10
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
                                        <div className="flex items-center">
                                          Term 1
                                        </div>
                                        <div className="flex items-center">
                                          Term 2
                                        </div>
                                        <div className="flex items-center">
                                          Term 3
                                        </div>
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
                                      {["TERM 1", "TERM 2", "TERM 3"].map(
                                        (term) => (
                                          <div
                                            key={term}
                                            className="grid min-w-0 grid-cols-[52px_minmax(0,1fr)] items-end gap-3"
                                          >
                                            <div>{term}</div>
                                            <div className="min-w-0 w-full border-b border-black">
                                              &nbsp;
                                            </div>
                                          </div>
                                        ),
                                      )}
                                    </div>

                                    <div className="mt-3 text-center text-[12px] font-normal">
                                      CERTIFICATE OF TRANSFER
                                    </div>
                                    <div className="mt-1 text-[9.5px] italic leading-snug">
                                      This is to certify that the above-named
                                      learner has satisfactorily completed the
                                      requirements for the grade level
                                      indicated.
                                    </div>

                                    <div className="mt-3 grid min-w-0 grid-cols-2 gap-5 overflow-hidden px-1 text-[10px] font-bold">
                                      <div className="grid min-w-0 grid-cols-[100px_minmax(0,1fr)] items-end gap-2">
                                        <div>Admitted to Grade</div>
                                        <div className="min-w-0 w-full border-b border-black">
                                          &nbsp;
                                        </div>
                                      </div>
                                      <div className="grid min-w-0 grid-cols-[155px_minmax(0,1fr)] items-end gap-2">
                                        <div>
                                          Eligible for Admission to Grade
                                        </div>
                                        <div className="min-w-0 w-full border-b border-black">
                                          &nbsp;
                                        </div>
                                      </div>
                                    </div>

                                    <div className="mt-3 grid min-w-0 grid-cols-[62px_minmax(0,1fr)] items-end gap-3 px-1 text-[10px] font-bold">
                                      <div>Approved:</div>
                                      <div className="min-w-0 w-full border-b border-black">
                                        &nbsp;
                                      </div>
                                    </div>

                                    <div className="mt-4 grid min-w-0 grid-cols-2 gap-10 px-1 text-center text-[9.5px]">
                                      <div>
                                        <div className="min-h-[18px] min-w-0 w-full border-b border-black px-1">
                                          {schoolProfile?.principal || ""}
                                        </div>
                                        <div>School Head</div>
                                      </div>
                                      <div>
                                        <div className="min-h-[18px] min-w-0 w-full border-b border-black px-1">
                                          {selectedTeacherName}
                                        </div>
                                        <div>Adviser</div>
                                      </div>
                                    </div>

                                    <div className="mt-3 text-center text-[11.5px] font-normal">
                                      CANCELLATION OF ELIGIBILITY TO TRANSFER
                                    </div>
                                    <div className="mt-2 grid min-w-0 grid-cols-[72px_minmax(0,1fr)_32px_minmax(0,1fr)] items-end gap-2 overflow-hidden px-1 text-[9.5px] font-bold">
                                      <div>Admitted in:</div>
                                      <div className="min-w-0 w-full border-b border-black">
                                        &nbsp;
                                      </div>
                                      <div>Date:</div>
                                      <div className="min-w-0 w-full border-b border-black">
                                        &nbsp;
                                      </div>
                                    </div>

                                    <div className="mx-auto mt-4 w-[58%] max-w-full text-center text-[9.5px]">
                                      <div className="min-h-[18px] border-b border-black">
                                        {schoolProfile?.principal || ""}
                                      </div>
                                      <div>School Head</div>
                                    </div>
                                  </section>
                                </div>
                              </div>
                            </div>
                          ) : isIlsPreview && ilsMetrics ? (
                            <IlsBondPaper
                              selectedClass={selectedClass}
                              schoolProfile={schoolProfile}
                              selectedTeacherName={selectedTeacherName}
                              metrics={ilsMetrics}
                              paperHeight={selectedPaper.height}
                            />
                          ) : isSf6Preview && sf6Metrics ? (
                            <Sf6BondPaper
                              schoolName={sf6SchoolName}
                              schoolId={sf6SchoolId}
                              region={sf6Region}
                              division={sf6Division}
                              district={sf6District}
                              schoolYear={selectedClass.school_year || ""}
                              gradeSection={`${selectedClass.grade_level || ""}${
                                selectedClass.section
                                  ? ` - ${selectedClass.section}`
                                  : ""
                              }`}
                              schoolHead={sf6SchoolHead}
                              statusRows={sf6StatusRows}
                              metrics={sf6Metrics}
                            />
                          ) : isSf8Preview ? (
                            <Sf8BondPaper
                              selectedClass={selectedClass}
                              schoolProfile={schoolProfile}
                              selectedTeacherProfile={selectedTeacherProfile}
                              learners={sf8Learners}
                              selectedTeacherName={selectedTeacherName}
                            />
                          ) : isSf10Preview && sf10Metrics ? (
                            <Sf10BondPaper
                              selectedClass={selectedClass}
                              selectedTeacherName={selectedTeacherName}
                              schoolProfile={schoolProfile}
                              selectedTeacherProfile={selectedTeacherProfile}
                              metrics={sf10Metrics}
                              paperHeight={selectedPaper.height}
                            />
                          ) : isLirPreview && sf10Metrics ? (
                            <LirBondPaper
                              selectedClass={selectedClass}
                              learner={sf10Metrics.learner}
                              paperHeight={selectedPaper.height}
                            />
                          ) : isIndexCardPreview && sf10Metrics ? (
                            <IndexCardBondPaper
                              selectedClass={selectedClass}
                              selectedTeacherName={selectedTeacherName}
                              metrics={sf10Metrics}
                            />
                          ) : isSchoolIdPreview && sf10Metrics ? (
                            <SchoolIdBondPaper
                              selectedClass={selectedClass}
                              learner={sf10Metrics.learner}
                            />
                          ) : (
                            <div className="px-8 py-9 sm:px-14">
                              <div className="grid grid-cols-[64px_minmax(0,1fr)_64px] items-center border-b-2 border-[var(--admin-color-4a3a33)] pb-5">
                                <div className="flex items-center justify-center translate-x-25 -translate-y-1">
                                  <img
                                    src={deped}
                                    alt="Department of Education logo"
                                    className="size-30 object-contain sm:size-14 lg:size-20"
                                  />
                                </div>

                                <div className="text-center">
                                  <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--admin-color-6f5c53)]">
                                    Republic of the Philippines
                                  </div>

                                  <div className="mt-1 text-xs font-semibold uppercase text-[var(--admin-text-heading)]">
                                    Department of Education
                                  </div>

                                  <div className="mt-1 text-[11px] font-semibold text-[var(--admin-text-heading)]">
                                    Agusan del Sur National Science High School
                                  </div>
                                </div>

                                <div aria-hidden="true" />
                              </div>

                              <div className="py-8 text-center">
                                <div className="text-sm font-bold uppercase tracking-wide">
                                  {formTitle(selectedForm)}
                                </div>
                                <div className="mt-1 text-xs font-semibold">
                                  {selectedClass.grade_level || "Grade Level"} ·
                                  Section {selectedClass.section || "-"}
                                </div>
                                <div className="mt-1 text-[10px] text-[var(--admin-color-75645d)]">
                                  School Year {selectedClass.school_year || "-"}
                                </div>
                              </div>

                              <div className="grid gap-3 text-[10px] sm:grid-cols-2">
                                <PreviewField
                                  label="Subject"
                                  value={selectedClass.subject || "-"}
                                />
                                <PreviewField
                                  label="Class Adviser"
                                  value={
                                    selectedClass.teacher_name ||
                                    teacherById.get(selectedClass.teacher_id)
                                      ?.full_name ||
                                    "-"
                                  }
                                />
                                <PreviewField
                                  label="Grade Level"
                                  value={selectedClass.grade_level || "-"}
                                />
                                <PreviewField
                                  label="Section"
                                  value={selectedClass.section || "-"}
                                />
                              </div>

                              <div className="mt-8 overflow-hidden rounded border border-[var(--admin-preview-border)]">
                                <div className="grid grid-cols-[56px_1fr_140px] bg-[var(--admin-color-f2ece7)] text-[9px] font-semibold uppercase text-[var(--admin-color-5f4b43)]">
                                  <div className="border-r border-[var(--admin-preview-border)] px-3 py-2">
                                    No.
                                  </div>
                                  <div className="border-r border-[var(--admin-preview-border)] px-3 py-2">
                                    Learner / Record
                                  </div>
                                  <div className="px-3 py-2">Remarks</div>
                                </div>
                                {[1, 2, 3, 4, 5].map((row) => (
                                  <div
                                    key={row}
                                    className="grid grid-cols-[56px_1fr_140px] border-t border-[var(--admin-preview-grid)] text-[9px] text-[var(--admin-color-8a7770)]"
                                  >
                                    <div className="border-r border-[var(--admin-preview-grid)] px-3 py-3">
                                      {row}
                                    </div>
                                    <div className="border-r border-[var(--admin-preview-grid)] px-3 py-3">
                                      &nbsp;
                                    </div>
                                    <div className="px-3 py-3">&nbsp;</div>
                                  </div>
                                ))}
                              </div>

                              <div className="mt-8 text-center text-[9px] text-[var(--admin-text-subtle)]">
                                Administrative preview · Open the teacher
                                workspace to complete or print the official
                                form.
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <EmptyState text="Select a class record to preview this form." />
            )}
          </Panel>
        </div>
      </div>
    </div>
  );
}

function SogBondPaper({
  selectedClass,
  selectedTeacherName,
  metrics,
  paperHeight,
}: {
  selectedClass: ClassRow;
  selectedTeacherName: string;
  metrics: {
    term: SogTerm;
    termLabel: string;
    subjects: string[];
    learners: Array<{
      learner: StudentRow;
      name: string;
      lrn: string;
      sexGroup: string;
      subjectScores: Array<{ subject: string; score: number | null }>;
      termAverage: number | null;
      rank: number | null;
      descriptor: string;
    }>;
  };
  paperHeight: number;
}) {
  const groups = [
    { key: "male", label: "MALE" },
    { key: "female", label: "FEMALE" },
    { key: "unspecified", label: "UNSPECIFIED" },
  ] as const;
  const subjectWidth = metrics.subjects.length
    ? `${57 / metrics.subjects.length}%`
    : "57%";
  let rowNumber = 0;

  return (
    <div
      className="box-border p-5 text-[7px] text-black"
      style={{
        fontFamily: "Arial, Helvetica, sans-serif",
        minHeight: `${paperHeight}px`,
      }}
    >
      <div className="pt-3 text-center leading-tight">
        <div className="text-[14px] font-bold uppercase">
          SUMMARY OF GRADES - {metrics.termLabel}
        </div>
        <div className="mt-1 text-[10px] font-medium">
          {selectedClass.grade_level || "Grade Level"} -{" "}
          {selectedClass.section || "Section"}
        </div>
      </div>

      <table className="mt-4 w-full table-fixed border-collapse border border-black text-[6px]">
        <colgroup>
          <col style={{ width: "3%" }} />
          <col style={{ width: "13%" }} />
          <col style={{ width: "9%" }} />
          {metrics.subjects.map((subject) => (
            <col key={`sog-col-${subject}`} style={{ width: subjectWidth }} />
          ))}
          <col style={{ width: "6%" }} />
          <col style={{ width: "4%" }} />
          <col style={{ width: "8%" }} />
        </colgroup>
        <thead>
          <tr className="bg-[#0B5A96] text-white">
            <th className="border border-black px-1 py-1.5">#</th>
            <th className="border border-black px-1 py-1.5 text-left">
              LEARNER&apos;S NAME
            </th>
            <th className="border border-black px-1 py-1.5">LRN</th>
            {metrics.subjects.map((subject) => (
              <th
                key={`sog-header-${subject}`}
                className="border border-black px-0.5 py-1.5 leading-tight"
              >
                {subject}
              </th>
            ))}
            <th className="border border-black bg-[#0A7A4B] px-0.5 py-1.5">
              TERM AVG
            </th>
            <th className="border border-black px-0.5 py-1.5">RANK</th>
            <th className="border border-black px-0.5 py-1.5">DESCRIPTOR</th>
          </tr>
        </thead>
        <tbody>
          {groups.map((group) => {
            const groupLearners = metrics.learners.filter(
              (learner) => learner.sexGroup === group.key,
            );
            if (groupLearners.length === 0) return null;

            return (
              <Fragment key={`sog-group-${group.key}`}>
                <tr className="bg-[#EAF2F8] text-[7px] font-bold">
                  <td
                    colSpan={metrics.subjects.length + 6}
                    className="border border-black px-1 py-1 text-left"
                  >
                    {group.label}
                  </td>
                </tr>
                {groupLearners.map((learner) => {
                  rowNumber += 1;
                  return (
                    <tr key={`sog-learner-${learner.learner.id}-${group.key}`}>
                      <td className="border border-black px-1 py-1 text-center">
                        {rowNumber}
                      </td>
                      <td className="border border-black px-1 py-1 font-semibold lowercase">
                        {learner.name}
                      </td>
                      <td className="border border-black px-1 py-1 text-center">
                        {learner.lrn}
                      </td>
                      {learner.subjectScores.map((subjectScore) => (
                        <td
                          key={`sog-score-${learner.learner.id}-${subjectScore.subject}`}
                          className="border border-black px-0.5 py-1 text-center"
                        >
                          {typeof subjectScore.score === "number"
                            ? Math.round(subjectScore.score)
                            : ""}
                        </td>
                      ))}
                      <td className="border border-black bg-[#E8F7F0] px-0.5 py-1 text-center font-bold">
                        {typeof learner.termAverage === "number"
                          ? Math.round(learner.termAverage)
                          : ""}
                      </td>
                      <td className="border border-black px-0.5 py-1 text-center font-bold">
                        {learner.rank || ""}
                      </td>
                      <td className="border border-black px-0.5 py-1 text-center text-[5px]">
                        {learner.descriptor}
                      </td>
                    </tr>
                  );
                })}
              </Fragment>
            );
          })}

          {metrics.learners.length === 0 && (
            <tr>
              <td
                colSpan={metrics.subjects.length + 6}
                className="border border-black px-2 py-5 text-center text-[7px]"
              >
                No learner records found for this grade and section.
              </td>
            </tr>
          )}
        </tbody>
      </table>

      <div className="mt-5 grid grid-cols-2 gap-20 text-[7px]">
        <div className="flex items-end gap-2">
          <b>Prepared by:</b>
          <div className="flex-1 text-center">
            <div className="border-b border-black px-2 pb-0.5 font-semibold">
              {selectedTeacherName}
            </div>
            <div className="mt-1 text-[6px]">Class Adviser</div>
          </div>
        </div>
        <div className="flex items-end gap-2">
          <b>Noted by:</b>
          <div className="flex-1 text-center">
            <div className="border-b border-black px-2 pb-0.5">&nbsp;</div>
            <div className="mt-1 text-[6px]">School Head</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function GradeSlipBondPaper({
  selectedClass,
  selectedTeacherName,
  metrics,
  paperHeight,
}: {
  selectedClass: ClassRow;
  selectedTeacherName: string;
  metrics: {
    learner: StudentRow | null;
    subjects: Array<{
      subject: string;
      term1: number | null;
      term2: number | null;
      term3: number | null;
      final: number | null;
      remarks: string;
    }>;
    generalAverage: number | null;
  };
  paperHeight: number;
}) {
  const descriptorFor = (score: number | null) => {
    if (typeof score !== "number") return "";
    if (score >= 90) return "Outstanding";
    if (score >= 85) return "Very Satisfactory";
    if (score >= 80) return "Satisfactory";
    if (score >= 75) return "Fairly Satisfactory";
    return "Did Not Meet Expectations";
  };

  const visibleSubjects = metrics.subjects.slice(0, 9);
  const emptyRows = Math.max(0, 9 - visibleSubjects.length);

  return (
    <div
      className="box-border p-3 text-[8px] text-black"
      style={{
        fontFamily: "Arial, Helvetica, sans-serif",
        minHeight: `${paperHeight}px`,
      }}
    >
      <div
        className="border-2 border-black p-3"
        style={{ minHeight: `${paperHeight - 24}px` }}
      >
        <div className="grid grid-cols-[1fr_auto_1fr] items-center">
          <div />
          <div className="flex items-center justify-center gap-5">
            <img
              src={deped}
              alt="Department of Education logo"
              className="h-9 w-auto object-contain"
            />
            <div className="text-[13px] font-bold uppercase">Grade Slip</div>
          </div>
          <div />
        </div>

        <div className="mt-3 grid grid-cols-[1fr_0.48fr] gap-x-3 gap-y-2 text-[9px]">
          <div className="flex items-end gap-2">
            <b>Name:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1 font-semibold lowercase">
              {metrics.learner ? studentName(metrics.learner) : ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>LRN:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {metrics.learner?.lrn || ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>Grade &amp; Section:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {selectedClass.grade_level || "Grade Level"} -{" "}
              {selectedClass.section || "Section"}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>School Year:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {selectedClass.school_year || ""}
            </span>
          </div>
        </div>

        <table className="mt-3 w-full table-fixed border-collapse border border-black text-[9px]">
          <colgroup>
            <col style={{ width: "16%" }} />
            <col style={{ width: "22%" }} />
            <col style={{ width: "22%" }} />
            <col style={{ width: "22%" }} />
            <col style={{ width: "5%" }} />
            <col style={{ width: "13%" }} />
          </colgroup>
          <thead>
            <tr className="bg-[#0B5A96] text-white">
              <th className="border border-black px-1 py-1.5 text-left">
                Learning Area
              </th>
              <th className="border border-black px-1 py-1.5">Term 1</th>
              <th className="border border-black px-1 py-1.5">Term 2</th>
              <th className="border border-black px-1 py-1.5">Term 3</th>
              <th className="border border-black bg-[#0A805B] px-1 py-1.5">
                Final
              </th>
              <th className="border border-black px-1 py-1.5">Descriptor</th>
            </tr>
          </thead>
          <tbody>
            {visibleSubjects.map((row) => (
              <tr key={`grade-slip-${row.subject}`}>
                <td className="border border-black px-1 py-1.5 font-semibold">
                  {row.subject}
                </td>
                <td className="border border-black px-1 py-1.5 text-center">
                  {typeof row.term1 === "number" ? Math.round(row.term1) : ""}
                </td>
                <td className="border border-black px-1 py-1.5 text-center">
                  {typeof row.term2 === "number" ? Math.round(row.term2) : ""}
                </td>
                <td className="border border-black px-1 py-1.5 text-center">
                  {typeof row.term3 === "number" ? Math.round(row.term3) : ""}
                </td>
                <td className="border border-black bg-[#D7F5E6] px-1 py-1.5 text-center font-bold">
                  {typeof row.final === "number" ? Math.round(row.final) : ""}
                </td>
                <td className="border border-black px-1 py-1.5 text-center text-[7px] font-semibold">
                  {descriptorFor(row.final)}
                </td>
              </tr>
            ))}

            {Array.from({ length: emptyRows }).map((_, rowIndex) => (
              <tr key={`grade-slip-empty-${rowIndex}`}>
                <td className="border border-black px-1 py-1.5">&nbsp;</td>
                <td className="border border-black px-1 py-1.5" />
                <td className="border border-black px-1 py-1.5" />
                <td className="border border-black px-1 py-1.5" />
                <td className="border border-black bg-[#D7F5E6] px-1 py-1.5" />
                <td className="border border-black px-1 py-1.5" />
              </tr>
            ))}

            <tr className="font-bold">
              <td
                colSpan={4}
                className="border border-black bg-[#F2F3F4] px-1 py-1.5 text-center"
              >
                GENERAL AVERAGE
              </td>
              <td className="border border-black bg-[#BFEFD0] px-1 py-1.5 text-center">
                {typeof metrics.generalAverage === "number"
                  ? Math.round(metrics.generalAverage)
                  : ""}
              </td>
              <td className="border border-black bg-[#F2F3F4] px-1 py-1.5 text-center text-[7px]">
                {descriptorFor(metrics.generalAverage)}
              </td>
            </tr>
          </tbody>
        </table>

        <div className="mt-5 grid grid-cols-2 gap-24 text-center text-[8px]">
          <div>
            <div className="border-t border-black pt-1">
              {selectedTeacherName}
            </div>
            <div>Class Adviser</div>
          </div>
          <div>
            <div className="border-t border-black pt-1">&nbsp;</div>
            <div>Parent/Guardian Signature</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function IlsBondPaper({
  selectedClass,
  schoolProfile,
  selectedTeacherName,
  metrics,
  paperHeight,
}: {
  selectedClass: ClassRow;
  schoolProfile: ProfileRow | null;
  selectedTeacherName: string;
  metrics: IlsMetricsData;
  paperHeight: number;
}) {
  const learnerName = metrics.learner
    ? (metrics.learner.last_name || studentName(metrics.learner)).toUpperCase()
    : "LEARNER NAME";
  const formattedToday = new Intl.DateTimeFormat("en-PH", {
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date());
  const riskTheme =
    metrics.riskLevel === "HIGH"
      ? { border: "#DC2626", background: "#FEF2F2", text: "#B91C1C" }
      : metrics.riskLevel === "MODERATE"
        ? { border: "#D97706", background: "#FFFBEB", text: "#B45309" }
        : { border: "#16A34A", background: "#F0FDF4", text: "#15803D" };
  const scoreCell = (score: number | null) =>
    typeof score === "number" ? Math.round(score) : "";

  return (
    <div
      className="box-border p-8 text-[7px] leading-[1.3] text-black"
      style={{
        fontFamily: "Arial, Helvetica, sans-serif",
        minHeight: `${paperHeight}px`,
      }}
    >
      <div className="grid grid-cols-[105px_minmax(0,1fr)_105px] items-center gap-3">
        <div className="flex justify-start">
          <img
            src={deped}
            alt="Department of Education logo"
            className="h-12 w-auto object-contain"
          />
        </div>
        <div className="text-center leading-tight">
          <div className="text-[7px] uppercase">
            Republic of the Philippines
          </div>
          <div className="text-[9px] font-bold uppercase">
            Department of Education
          </div>
          <div className="text-[7px]">
            {schoolProfile?.region
              ? `Region ${schoolProfile.region}`
              : "Region"}
          </div>
          <div className="text-[7px]">
            {schoolProfile?.division || "Division"}
          </div>
          <div className="text-[8px] font-bold uppercase">
            {schoolProfile?.school_name || "SCHOOL NAME"}
          </div>
        </div>
        <div aria-hidden="true" />
      </div>

      <div className="mt-2 border-t border-black" />
      <div className="mt-3 text-center text-[11px] font-bold uppercase underline">
        Individual Learner Status &amp; Intervention Report
      </div>
      <div className="mt-1 text-center text-[7px]">
        SY {selectedClass.school_year || "____-____"}
      </div>

      <div
        className="mt-4 flex items-center justify-between rounded border px-3 py-2 text-[7px]"
        style={{
          borderColor: riskTheme.border,
          backgroundColor: riskTheme.background,
          color: riskTheme.text,
        }}
      >
        <div className="flex items-center gap-2 font-bold">
          <span
            className="inline-block size-3 rounded-full"
            style={{ backgroundColor: riskTheme.border }}
          />
          RISK LEVEL: {metrics.riskLevel}
        </div>
        <div className="text-[6px] text-black">
          {metrics.failingSubjectCount} failing subject(s) · {metrics.absences}{" "}
          absence(s) · {metrics.missingActivities} missing activities
        </div>
      </div>

      <div className="mt-4 space-y-1">
        <div>
          <b>Date:</b> {formattedToday}
        </div>
        <div>
          <b>To the Parent/Guardian of:</b>
        </div>
        <div className="text-[9px] font-bold">{learnerName}</div>
      </div>

      <div className="mt-3">Dear Parent/Guardian,</div>
      <p className="mt-3 text-justify indent-8">
        This formal communication serves to inform you regarding the current
        academic and attendance standing of your child. Based on our latest
        records, your child&apos;s academic standing is being monitored.
      </p>

      <table className="mt-4 w-full table-fixed border-collapse border border-black text-[6px]">
        <colgroup>
          <col style={{ width: "36%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "10%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "22%" }} />
        </colgroup>
        <thead>
          <tr>
            <th
              colSpan={7}
              className="border border-black bg-[#F3F4F6] px-2 py-1.5 text-left text-[7px]"
            >
              <span className="font-bold">1. ACADEMIC STATUS</span>
              <span className="float-right font-normal">
                Class Rank: {metrics.classRank || "—"} of{" "}
                {metrics.classSize || 0} · Class Avg:{" "}
                {typeof metrics.classAverage === "number"
                  ? Math.round(metrics.classAverage)
                  : "—"}{" "}
                · Student Avg:{" "}
                {typeof metrics.generalAverage === "number"
                  ? Math.round(metrics.generalAverage)
                  : "—"}
              </span>
            </th>
          </tr>
          <tr>
            <th className="border border-black px-1 py-1 text-left">
              Learning Area
            </th>
            <th className="border border-black px-1 py-1">1st Term</th>
            <th className="border border-black px-1 py-1">2nd Term</th>
            <th className="border border-black px-1 py-1">3rd Term</th>
            <th className="border border-black px-1 py-1">Average</th>
            <th className="border border-black px-1 py-1">Trend</th>
            <th className="border border-black px-1 py-1">Remarks</th>
          </tr>
        </thead>
        <tbody>
          {metrics.subjects.map((row, index) => (
            <tr key={`ils-subject-${row.subject}-${index}`}>
              <td className="border border-black px-1 py-1 font-semibold">
                {row.subject}
              </td>
              <td
                className={`border border-black px-1 py-1 text-center ${typeof row.term1 === "number" && row.term1 < 75 ? "text-red-600" : "text-emerald-600"}`}
              >
                {scoreCell(row.term1)}
              </td>
              <td
                className={`border border-black px-1 py-1 text-center ${typeof row.term2 === "number" && row.term2 < 75 ? "text-red-600" : "text-emerald-600"}`}
              >
                {scoreCell(row.term2)}
              </td>
              <td
                className={`border border-black px-1 py-1 text-center ${typeof row.term3 === "number" && row.term3 < 75 ? "text-red-600" : "text-emerald-600"}`}
              >
                {scoreCell(row.term3)}
              </td>
              <td className="border border-black px-1 py-1 text-center font-bold">
                {scoreCell(row.average)}
              </td>
              <td
                className={`border border-black px-1 py-1 text-center font-bold ${row.trend === "↓" ? "text-red-600" : row.trend === "✓" ? "text-emerald-600" : ""}`}
              >
                {row.trend}
              </td>
              <td
                className={`border border-black px-1 py-1 text-center text-[5px] font-bold ${row.remarks === "FOR INTERVENTION" ? "text-red-600" : "text-emerald-600"}`}
              >
                {row.remarks}
              </td>
            </tr>
          ))}
          <tr className="font-bold">
            <td colSpan={4} className="border border-black px-1 py-1">
              GENERAL AVERAGE
            </td>
            <td className="border border-black px-1 py-1 text-center text-emerald-600">
              {scoreCell(metrics.generalAverage)}
            </td>
            <td className="border border-black px-1 py-1 text-center">—</td>
            <td className="border border-black px-1 py-1 text-center text-emerald-600">
              {typeof metrics.generalAverage === "number"
                ? metrics.generalAverage >= 75
                  ? "PASSED"
                  : "FOR INTERVENTION"
                : ""}
            </td>
          </tr>
        </tbody>
      </table>

      <table className="mt-3 w-full table-fixed border-collapse border border-black text-[6px]">
        <thead>
          <tr>
            <th
              colSpan={4}
              className="border border-black bg-[#F3F4F6] px-2 py-1.5 text-left text-[7px]"
            >
              2. COMPONENT PERFORMANCE BREAKDOWN ({metrics.componentSubject})
            </th>
          </tr>
          <tr>
            <th className="w-[14%] border border-black px-1 py-1">Term</th>
            <th className="border border-black bg-[#EFF6FF] px-1 py-1">
              Written Works (20%)
            </th>
            <th className="border border-black bg-[#F5F3FF] px-1 py-1">
              Performance Tasks (50%)
            </th>
            <th className="border border-black bg-[#FFF7ED] px-1 py-1">
              Summative/Term Exams (30%)
            </th>
          </tr>
        </thead>
        <tbody>
          {metrics.componentRows.map((row) => (
            <tr key={row.termLabel}>
              <td className="border border-black px-1 py-1 text-center font-bold">
                {row.termLabel}
              </td>
              <td className="border border-black px-1 py-1 text-center text-emerald-600">
                {row.writtenWorks}
              </td>
              <td className="border border-black px-1 py-1 text-center text-amber-600">
                {row.performanceTasks}
              </td>
              <td className="border border-black px-1 py-1 text-center text-amber-600">
                {row.summativeExam}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <table className="mt-3 w-full border-collapse border border-black text-[6px]">
        <thead>
          <tr>
            <th
              colSpan={2}
              className="border border-black bg-[#F3F4F6] px-2 py-1.5 text-left text-[7px]"
            >
              3. ATTENDANCE STATUS
            </th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td className="w-[70%] border border-black px-2 py-1">
              Total Number of Days Absent:
            </td>
            <td className="border border-black px-2 py-1 text-center font-bold">
              {metrics.absences}
            </td>
          </tr>
          <tr>
            <td className="border border-black px-2 py-1">
              Total Number of Days Tardy / Half-Day:
            </td>
            <td className="border border-black px-2 py-1 text-center font-bold">
              {metrics.lateHalfDays}
            </td>
          </tr>
        </tbody>
      </table>

      <div className="mt-3 border border-black text-[6px]">
        <div className="border-b border-black bg-[#F3F4F6] px-2 py-1.5 text-[7px] font-bold">
          RECOMMENDED INTERVENTIONS
        </div>
        <ul className="list-disc space-y-1 px-6 py-2">
          {metrics.recommendations.map((recommendation) => (
            <li key={recommendation}>{recommendation}</li>
          ))}
        </ul>
      </div>

      <p className="mt-4 text-justify indent-8">
        In light of this, we strongly request your presence for a{" "}
        <b>Parent-Teacher Conference (PTC)</b> to discuss remediation
        strategies.
      </p>

      <div className="mt-2 ml-10 space-y-0.5">
        <div>📅 Date: ____________________</div>
        <div>⏰ Time: ____________________</div>
        <div>
          📍 Venue: <u>School Principal&apos;s Office</u>
        </div>
      </div>

      <div className="mt-8 grid grid-cols-2 gap-20 text-center text-[6px]">
        <div>
          <div className="border-t border-black pt-1">
            {selectedTeacherName}
          </div>
          <div>Class Adviser</div>
        </div>
        <div>
          <div className="border-t border-black pt-1">&nbsp;</div>
          <div>School Head / Principal</div>
        </div>
      </div>

      <div className="mt-8 border-t border-dashed border-black pt-2 text-[6px]">
        <div className="text-center italic">
          Please sign and return this portion to the Class Adviser to confirm
          receipt.
        </div>
        <div className="mt-3 font-bold uppercase">
          Acknowledgment Slip — {metrics.learner?.last_name || "Learner"}
        </div>
        <div className="mt-2 space-y-1">
          <div>
            ☐ I acknowledge receipt of this Individual Learner Status &amp;
            Intervention Report.
          </div>
          <div>
            ☐ I will attend the scheduled Parent-Teacher Conference on
            ____________________.
          </div>
          <div>
            ☐ I commit to monitoring my child&apos;s daily attendance and
            academic progress.
          </div>
        </div>
        <div className="mt-5 grid grid-cols-[1fr_0.6fr_0.6fr] gap-5">
          <div>Parent/Guardian Signature: ______________________________</div>
          <div>Contact No.: __________________</div>
          <div>Date: __________________</div>
        </div>
      </div>
    </div>
  );
}

function SchoolIdBondPaper({
  selectedClass,
  learner,
}: {
  selectedClass: ClassRow;
  learner: StudentRow | null;
}) {
  const learnerDisplayName = learner
    ? learner.last_name || studentName(learner)
    : "";

  return (
    <div
      className="box-border p-4 text-[8px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="mx-auto w-[48%] rounded-[10px] border border-[#0B5A96] p-2.5">
        <div className="text-center text-[6px] leading-tight text-[#374151]">
          Republic of the Philippines
        </div>

        <div className="mt-7 border-t border-[#0B5A96]" />

        <div className="mt-2 grid grid-cols-[58px_minmax(0,1fr)] gap-2">
          <div className="grid h-[72px] place-items-center rounded border border-[#C9CED6] bg-[#F7F7F7] text-center text-[6px] text-[#8B8F97]">
            2x2 PHOTO
          </div>

          <div className="space-y-1 text-[7px] leading-tight">
            <div className="flex items-end gap-1">
              <b>Name:</b>
              <span className="min-h-4 flex-1 border-b border-[#C9CED6] px-1 lowercase">
                {learnerDisplayName}
              </span>
            </div>

            <div className="flex items-end gap-1">
              <b>LRN:</b>
              <span className="min-h-4 flex-1 border-b border-[#C9CED6] px-1">
                {learner?.lrn || ""}
              </span>
            </div>

            <div className="flex items-end gap-1">
              <b>Grade &amp; Section:</b>
              <span className="min-h-4 flex-1 border-b border-[#C9CED6] px-1">
                {selectedClass.grade_level || "Grade Level"} -{" "}
                {selectedClass.section || "Section"}
              </span>
            </div>

            <div className="flex items-end gap-1">
              <b>S.Y.:</b>
              <span className="min-h-4 w-24 border-b border-[#C9CED6] px-1">
                {selectedClass.school_year || ""}
              </span>
            </div>

            <div className="flex items-end gap-1">
              <b>Address:</b>
              <span className="min-h-4 flex-1 border-b border-[#C9CED6] px-1">
                {learner?.address || ""}
              </span>
            </div>

            <div className="flex items-end gap-1">
              <b>Contact:</b>
              <span className="min-h-4 flex-1 border-b border-[#C9CED6] px-1">
                {learner?.contact_number || ""}
              </span>
            </div>
          </div>
        </div>

        <div className="mx-auto mt-4 w-24 text-center text-[6px]">
          <div className="border-t border-black pt-1">&nbsp;</div>
          <div>School Head</div>
        </div>
      </div>
    </div>
  );
}

function IndexCardBondPaper({
  selectedClass,
  selectedTeacherName,
  metrics,
}: {
  selectedClass: ClassRow;
  selectedTeacherName: string;
  metrics: {
    learner: StudentRow | null;
    subjects: Array<{
      subject: string;
      term1: number | null;
      term2: number | null;
      term3: number | null;
      final: number | null;
      remarks: string;
    }>;
    generalAverage: number | null;
  };
}) {
  const learner = metrics.learner;
  const learnerDisplayName =
    learner?.last_name || (learner ? studentName(learner) : "");

  return (
    <div
      className="box-border p-4 text-[8px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="w-[47%] border border-black p-2">
        <div className="text-center text-[11px] font-bold uppercase leading-tight">
          Index Card
        </div>

        <div className="mt-2 grid grid-cols-[1fr_0.95fr] gap-x-3 gap-y-1 text-[7px]">
          <div className="flex items-end gap-1">
            <b>Name:</b>
            <span className="min-h-4 flex-1 border-b border-black px-1 lowercase">
              {learnerDisplayName}
            </span>
          </div>

          <div className="flex items-end gap-1">
            <b>LRN:</b>
            <span className="min-h-4 flex-1 border-b border-black px-1">
              {learner?.lrn || ""}
            </span>
          </div>

          <div className="col-span-2 flex items-end gap-1">
            <b>Grade &amp; Section:</b>
            <span className="min-h-4 border-b border-black px-1">
              {selectedClass.grade_level || "Grade Level"} -{" "}
              {selectedClass.section || "Section"}
            </span>
            <b className="ml-1">BY:</b>
            <span className="min-h-4 w-24 border-b border-black px-1" />
          </div>
        </div>

        <table className="mt-2 w-full table-fixed border-collapse border border-black text-[7px]">
          <colgroup>
            <col style={{ width: "50%" }} />
            <col style={{ width: "11%" }} />
            <col style={{ width: "11%" }} />
            <col style={{ width: "11%" }} />
            <col style={{ width: "17%" }} />
          </colgroup>

          <thead>
            <tr className="bg-[#0B5A96] text-white">
              <th className="border border-black px-1 py-1 text-left">
                Subject
              </th>
              <th className="border border-black px-1 py-1">T1</th>
              <th className="border border-black px-1 py-1">T2</th>
              <th className="border border-black px-1 py-1">T3</th>
              <th className="border border-black bg-[#08784F] px-1 py-1">
                Final
              </th>
            </tr>
          </thead>

          <tbody>
            {metrics.subjects.slice(0, 11).map((row, index) => (
              <tr key={`index-card-subject-${row.subject}-${index}`}>
                <td className="border border-black px-1 py-1 font-semibold">
                  {row.subject}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term1 === "number" ? Math.round(row.term1) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term2 === "number" ? Math.round(row.term2) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term3 === "number" ? Math.round(row.term3) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center font-semibold">
                  {typeof row.final === "number" ? Math.round(row.final) : ""}
                </td>
              </tr>
            ))}

            {Array.from({
              length: Math.max(0, 11 - metrics.subjects.length),
            }).map((_, rowIndex) => (
              <tr key={`index-card-empty-${rowIndex}`}>
                {Array.from({ length: 5 }).map((__, cellIndex) => (
                  <td
                    key={`index-card-empty-${rowIndex}-${cellIndex}`}
                    className="border border-black px-1 py-1"
                  >
                    &nbsp;
                  </td>
                ))}
              </tr>
            ))}

            <tr className="font-bold">
              <td
                colSpan={4}
                className="border border-black px-1 py-1 text-center"
              >
                GEN. AVE.
              </td>
              <td className="border border-black bg-[#D7F5E6] px-1 py-1 text-center">
                {typeof metrics.generalAverage === "number"
                  ? Math.round(metrics.generalAverage)
                  : ""}
              </td>
            </tr>
          </tbody>
        </table>

        <div className="mx-auto mt-5 w-28 text-center text-[6px]">
          <div className="border-t border-black pt-1">
            {selectedTeacherName}
          </div>
          <div>Class Adviser</div>
        </div>
      </div>
    </div>
  );
}

function LirBondPaper({
  selectedClass,
  learner,
  paperHeight,
}: {
  selectedClass: ClassRow;
  learner: StudentRow | null;
  paperHeight: number;
}) {
  const sexValue = normalize(learner?.sex);
  const sex = sexValue.startsWith("m")
    ? "M"
    : sexValue.startsWith("f")
      ? "F"
      : "";

  const rows: Array<[string, string]> = [
    ["LRN", learner?.lrn || ""],
    ["Last Name", learner?.last_name || ""],
    ["First Name", learner?.first_name || ""],
    ["Middle Name", learner?.middle_name || ""],
    ["Birthdate", learner?.birthdate ? formatDate(learner.birthdate) : ""],
    ["Address", learner?.address || ""],
    ["Father's Name", learner?.father_name || ""],
    ["Mother's Name", learner?.mother_name || ""],
    ["Parent/Guardian", learner?.guardian || ""],
    ["Contact #", learner?.contact_number || ""],
    ["Mother Tongue", learner?.mother_tongue || ""],
    ["IP/Ethnic Group", learner?.ip_ethnic_group || learner?.ip_ethnic || ""],
    ["Religion", learner?.religion || ""],
    ["Remarks", learner?.remarks || ""],
    [
      "Disability/Condition",
      learner?.disability_condition || learner?.disability || "",
    ],
    ["Learning Modality", learner?.learning_modality || ""],
    ["Age", calculateAge(learner?.birthdate)],
    ["Sex", sex],
    [
      "Grade & Section",
      `${selectedClass.grade_level || "Grade Level"} - ${selectedClass.section || "Section"}`,
    ],
    ["School Year", selectedClass.school_year || ""],
  ];

  return (
    <div
      className="box-border p-3 text-[9px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div
        className="border border-black p-2"
        style={{ minHeight: `${paperHeight - 24}px` }}
      >
        <div className="pb-2 text-center text-[11px] font-bold uppercase leading-tight">
          Learner Individual Record
        </div>

        <table className="w-full table-fixed border-collapse border border-black text-[9px]">
          <colgroup>
            <col style={{ width: "35%" }} />
            <col style={{ width: "65%" }} />
          </colgroup>
          <tbody>
            {rows.map(([label, value]) => (
              <tr key={label}>
                <th className="border border-black bg-[#F3F4F6] px-2 py-1.5 text-left font-bold">
                  {label}
                </th>
                <td className="border border-black px-2 py-1.5">
                  {value || "\u00A0"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Sf8BondPaper({
  selectedClass,
  schoolProfile,
  selectedTeacherProfile,
  learners,
  selectedTeacherName,
}: {
  selectedClass: ClassRow;
  schoolProfile: ProfileRow | null;
  selectedTeacherProfile: ProfileRow | null;
  learners: StudentRow[];
  selectedTeacherName: string;
}) {
  const profile = schoolProfile || selectedTeacherProfile;
  const male = learners.filter((learner) =>
    normalize(learner.sex).startsWith("m"),
  );
  const female = learners.filter((learner) =>
    normalize(learner.sex).startsWith("f"),
  );
  const unspecified = learners.filter((learner) => {
    const sex = normalize(learner.sex);
    return !sex.startsWith("m") && !sex.startsWith("f");
  });
  const visibleRows = (group: StudentRow[], minimum: number) => [
    ...group,
    ...Array.from({ length: Math.max(0, minimum - group.length) }, () => null),
  ];
  let learnerNumber = 0;

  const renderLearnerRows = (
    label: string,
    group: StudentRow[],
    minimum: number,
  ) => (
    <Fragment key={label}>
      <tr className="h-[17px] bg-[#a6a6a6] font-bold italic">
        <td colSpan={15} className="border border-black px-1 text-left">
          {label}
        </td>
      </tr>
      {visibleRows(group, minimum).map((learner, index) => {
        if (learner) learnerNumber += 1;
        return (
          <tr key={`${label}-${learner?.id || index}`} className="h-[17px]">
            <td className="border border-black px-0.5 text-center">
              {learner ? learnerNumber : ""}
            </td>
            <td className="border border-black px-0.5 text-center">
              {learner?.lrn || ""}
            </td>
            <td className="border border-black px-1 uppercase">
              {learner ? studentName(learner) : ""}
            </td>
            <td className="border border-black px-0.5 text-center">
              {learner?.birthdate ? formatDate(learner.birthdate) : ""}
            </td>
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center">
              {learner ? calculateAge(learner.birthdate) : ""}
            </td>
            <td className="border border-black px-0.5 text-center" />
            <td className="border border-black px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-0.5 text-center" />
            <td className="border border-black bg-[#d9d9d9] px-1">
              {learner?.remarks || ""}
            </td>
          </tr>
        );
      })}
    </Fragment>
  );

  return (
    <div
      className="box-border h-full w-full overflow-hidden bg-white px-5 py-4 text-[5.8px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="relative text-center leading-tight">
        <img
          src={deped}
          alt="Department of Education logo"
          className="absolute left-0 top-0 h-16 w-16 object-contain"
        />
        <div className="text-[7px]">SF 8</div>
        <div className="text-[11px] font-bold">Department of Education</div>
        <div className="text-[11px] font-bold">
          School Form 8 Learner&apos;s Basic Health and Nutrition Report (SF8)
        </div>
        <div className="text-[7px] italic">(For All Grade Levels)</div>
      </div>

      <div className="mt-3 grid grid-cols-[1.6fr_1fr_1fr_0.7fr] gap-x-5 gap-y-1 text-[6.5px]">
        <div className="flex items-end gap-1">
          <span>School Name</span>
          <b className="flex-1 border border-black px-2 py-1 text-center uppercase">
            {profile?.school_name ||
              "Agusan del Sur National Science High School"}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>District</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {profile?.district || ""}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>Division</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {profile?.division || ""}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>Region</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {profile?.region || ""}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>School ID</span>
          <b className="w-24 border border-black px-2 py-1 text-center">
            {profile?.school_id || ""}
          </b>
          <span className="ml-2">Grade</span>
          <b className="w-14 border border-black px-2 py-1 text-center">
            {selectedClass.grade_level || ""}
          </b>
          <span className="ml-2">Section</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {selectedClass.section || ""}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>Track/Strand (SHS)</span>
          <span className="flex-1 border border-black px-2 py-1">&nbsp;</span>
        </div>
        <div className="flex items-end gap-1">
          <span>School Year</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {selectedClass.school_year || ""}
          </b>
        </div>
        <div className="flex items-end gap-1">
          <span>Date of Measurement</span>
          <b className="flex-1 border border-black px-2 py-1 text-center">
            {new Intl.DateTimeFormat("en-PH").format(new Date())}
          </b>
        </div>
      </div>

      <table className="mt-2 w-full table-fixed border-collapse text-[5.5px]">
        <colgroup>
          <col style={{ width: "3%" }} />
          <col style={{ width: "9%" }} />
          <col style={{ width: "17%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "4.5%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "5.5%" }} />
          <col style={{ width: "5.5%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "9%" }} />
          <col style={{ width: "7%" }} />
          <col style={{ width: "7%" }} />
          <col style={{ width: "8.5%" }} />
        </colgroup>
        <thead>
          <tr className="h-[34px] text-center font-bold">
            <th rowSpan={2} className="border border-black">
              No.
            </th>
            <th rowSpan={2} className="border border-black">
              LRN
            </th>
            <th rowSpan={2} className="border border-black leading-tight">
              Learner&apos;s Name
              <div className="font-normal">
                (Last Name, First Name, Name Extension, Middle Name)
              </div>
            </th>
            <th rowSpan={2} className="border border-black">
              Birthdate
              <br />
              (MM/DD/YYYY)
            </th>
            <th rowSpan={2} className="border border-black">
              Age
            </th>
            <th rowSpan={2} className="border border-black">
              Weight
              <br />
              (kg)
            </th>
            <th rowSpan={2} className="border border-black">
              Height
              <br />
              (m)
            </th>
            <th rowSpan={2} className="border border-black">
              Height²
              <br />
              (m²)
            </th>
            <th colSpan={3} className="border border-black">
              Nutritional Status
            </th>
            <th rowSpan={2} className="border border-black">
              Height for Age (HFA)
            </th>
            <th rowSpan={2} className="border border-black">
              Deworming Status
            </th>
            <th rowSpan={2} className="border border-black">
              Health Condition
            </th>
            <th rowSpan={2} className="border border-black">
              Remarks
            </th>
          </tr>
          <tr className="h-[20px] text-center font-bold">
            <th className="border border-black">
              BMI
              <br />
              (kg/m²)
            </th>
            <th className="border border-black">BMI Category</th>
            <th className="border border-black">Assessment</th>
          </tr>
        </thead>
        <tbody>
          {renderLearnerRows("MALE", male, 8)}
          {renderLearnerRows("FEMALE", [...female, ...unspecified], 8)}
        </tbody>
      </table>

      <div className="mt-3 text-center text-[9px] font-bold">SUMMARY TABLE</div>
      <table className="mt-1 w-full table-fixed border-collapse text-center text-[5.5px]">
        <thead>
          <tr>
            <th rowSpan={2} className="border border-black">
              SEX
            </th>
            <th colSpan={6} className="border border-black">
              Nutritional Status Summary Table
            </th>
            <th colSpan={5} className="border border-black">
              Height for Age (HFA) Summary Table
            </th>
          </tr>
          <tr>
            {[
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
            ].map((label) => (
              <th key={label} className="border border-black px-1 py-1">
                {label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {[
            ["MALE", male.length],
            ["FEMALE", female.length + unspecified.length],
            ["TOTAL", learners.length],
          ].map(([label, total]) => (
            <tr key={String(label)}>
              <th className="border border-black text-left">{label}</th>
              {[0, 0, 0, 0, 0, total, 0, 0, 0, 0, total].map((value, index) => (
                <td key={index} className="border border-black">
                  {value}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      <div className="mt-6 grid grid-cols-4 gap-10 text-center text-[6px]">
        {[
          ["Date of Assessment", ""],
          ["Conducted/Assessed By", selectedTeacherName],
          ["Certified Correct By", profile?.principal || ""],
          ["Reviewed By", ""],
        ].map(([label, name]) => (
          <div key={label}>
            <div className="mb-5 text-left">{label}:</div>
            <div className="border-b border-black font-semibold uppercase">
              {name || "\u00A0"}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Sf10BondPaper({
  selectedClass,
  selectedTeacherName,
  schoolProfile,
  selectedTeacherProfile,
  metrics,
  paperHeight,
}: {
  selectedClass: ClassRow;
  selectedTeacherName: string;
  schoolProfile: ProfileRow | null;
  selectedTeacherProfile: ProfileRow | null;
  metrics: {
    learner: StudentRow | null;
    subjects: Array<{
      subject: string;
      term1: number | null;
      term2: number | null;
      term3: number | null;
      term4: number | null;
      final: number | null;
      remarks: string;
    }>;
    generalAverage: number | null;
  };
  paperHeight: number;
}) {
  const learner = metrics.learner;
  const sexValue = normalize(learner?.sex);
  const sex = sexValue.startsWith("m")
    ? "M"
    : sexValue.startsWith("f")
      ? "F"
      : "";
  const schoolName =
    schoolProfile?.school_name ||
    selectedTeacherProfile?.school_name ||
    "Agusan del Sur National Science High School";
  const schoolId =
    schoolProfile?.school_id || selectedTeacherProfile?.school_id || "";
  const region = schoolProfile?.region || selectedTeacherProfile?.region || "";
  const division =
    schoolProfile?.division || selectedTeacherProfile?.division || "";
  const district =
    schoolProfile?.district || selectedTeacherProfile?.district || "";
  const schoolHead =
    schoolProfile?.principal || selectedTeacherProfile?.principal || "";

  return (
    <div
      className="box-border p-3 text-[7px] text-black"
      style={{ fontFamily: "'Arial Narrow', Arial, sans-serif" }}
    >
      <div
        className="border border-black p-2"
        style={{ minHeight: `${paperHeight - 32}px` }}
      >
        <div className="text-center leading-tight">
          <div>Republic of the Philippines</div>
          <div>Department of Education</div>
          <div className="text-[11px] font-bold">
            Learner Permanent Academic Record for Junior High School (SF10-JHS)
          </div>
          <div className="italic">(Formerly Form 137)</div>
        </div>

        <div className="mt-1 border border-black bg-[#cbc7a5] py-0.5 text-center text-[9px] font-bold">
          LEARNER&apos;S INFORMATION
        </div>
        <div className="mt-1 grid grid-cols-4 gap-x-3 gap-y-1">
          <div className="flex items-end gap-2">
            <b>LAST NAME:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1 font-semibold lowercase">
              {learner?.last_name || ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>FIRST NAME:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {learner?.first_name || ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>MIDDLE NAME:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1 text-center">
              {learner?.middle_name || ""}
            </span>
          </div>

          <div className="flex items-end gap-2">
            <b>LRN:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {learner?.lrn || ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>Birthdate:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1 text-center">
              {learner?.birthdate ? formatDate(learner.birthdate) : ""}
            </span>
          </div>
          <div className="flex items-end gap-2">
            <b>Sex:</b>
            <span className="min-h-5 flex-1 border-b border-black px-1">
              {sex}
            </span>
          </div>
        </div>

        <div className="mt-2 border border-black bg-[#cbc7a5] py-0.5 text-center text-[9px] font-bold">
          ELIGIBILITY FOR JHS ENROLLMENT
        </div>
        <div className="grid grid-cols-3 border-x border-b border-black px-2 py-1">
          <div>Elementary School Completer</div>
          <div>General Average: __________</div>
          <div>Citation (if any): __________</div>
          <div className="col-span-2">
            Name of Elementary School: ____________________
          </div>
          <div>School ID: __________</div>
        </div>

        <div className="mt-2 border border-black bg-[#cbc7a5] py-0.5 text-center text-[9px] font-bold">
          SCHOLASTIC RECORD
        </div>
        <div className="grid grid-cols-5 gap-x-2 border-x border-black px-1 py-1">
          <div>
            School: <b>{schoolName}</b>
          </div>
          <div>
            School ID: <b>{schoolId}</b>
          </div>
          <div>
            District: <b>{district}</b>
          </div>
          <div>
            Division: <b>{division}</b>
          </div>
          <div>
            Region: <b>{region}</b>
          </div>
          <div>
            Grade: <b>{selectedClass.grade_level}</b>
          </div>
          <div>
            Section: <b>{selectedClass.section}</b>
          </div>
          <div>
            School Year: <b>{selectedClass.school_year}</b>
          </div>
          <div className="col-span-2">
            Adviser: <b>{selectedTeacherName}</b>
          </div>
        </div>
        <table className="w-full table-fixed border-collapse border border-black text-[7px]">
          <colgroup>
            <col style={{ width: "34%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "9%" }} />
            <col style={{ width: "12%" }} />
            <col style={{ width: "18%" }} />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-black px-1 py-1.5 text-left">
                Learning Area
              </th>
              <th className="border border-black px-1 py-1.5">1</th>
              <th className="border border-black px-1 py-1.5">2</th>
              <th className="border border-black px-1 py-1.5">3</th>
              <th className="border border-black px-1 py-1.5">4</th>
              <th className="border border-black px-1 py-1.5">Final Rating</th>
              <th className="border border-black px-1 py-1.5">Remarks</th>
            </tr>
          </thead>
          <tbody>
            {metrics.subjects.map((row, index) => (
              <tr key={`sf10-subject-${row.subject}-${index}`}>
                <td className="border border-black px-1 py-1 font-semibold">
                  {row.subject}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term1 === "number" ? Math.round(row.term1) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term2 === "number" ? Math.round(row.term2) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term3 === "number" ? Math.round(row.term3) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center">
                  {typeof row.term4 === "number" ? Math.round(row.term4) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center font-semibold">
                  {typeof row.final === "number" ? Math.round(row.final) : ""}
                </td>
                <td className="border border-black px-1 py-1 text-center text-[7px] font-semibold">
                  {row.remarks}
                </td>
              </tr>
            ))}

            {Array.from({
              length: Math.max(0, 11 - metrics.subjects.length),
            }).map((_, rowIndex) => (
              <tr key={`sf10-empty-${rowIndex}`}>
                {Array.from({ length: 7 }).map((__, cellIndex) => (
                  <td
                    key={`sf10-empty-${rowIndex}-${cellIndex}`}
                    className="border border-black px-1 py-1.5"
                  >
                    &nbsp;
                  </td>
                ))}
              </tr>
            ))}

            <tr className="font-bold">
              <td
                colSpan={5}
                className="border border-black px-1 py-1 text-center"
              >
                GENERAL AVERAGE
              </td>
              <td className="border border-black px-1 py-1 text-center">
                {typeof metrics.generalAverage === "number"
                  ? Math.round(metrics.generalAverage)
                  : ""}
              </td>
              <td className="border border-black px-1 py-1 text-center">
                {typeof metrics.generalAverage === "number"
                  ? metrics.generalAverage >= 75
                    ? "PASSED"
                    : "FAILED"
                  : ""}
              </td>
            </tr>
          </tbody>
        </table>

        <div className="mt-2 border-t-4 border-[#cbc7a5]">
          <div className="grid grid-cols-5 border border-black text-center font-bold">
            <div>Remedial Classes</div>
            <div>Final Rating</div>
            <div>Remedial Class Mark</div>
            <div>Recomputed Final Grade</div>
            <div>Remarks</div>
            {Array.from({ length: 10 }).map((_, index) => (
              <div key={index} className="h-4 border-t border-black">
                &nbsp;
              </div>
            ))}
          </div>
        </div>

        <div className="mt-2 border border-black">
          <div className="bg-[#cbc7a5] py-0.5 text-center text-[9px] font-bold">
            CERTIFICATION
          </div>
          <div className="px-2 py-2">
            I CERTIFY that this is a true record of{" "}
            <b>{learner ? studentName(learner) : "________________"}</b> with
            LRN <b>{learner?.lrn || "____________"}</b>.
          </div>
          <div className="grid grid-cols-3 gap-10 px-5 pb-2 pt-6 text-center">
            <div className="border-t border-black">Date</div>
            <div className="border-t border-black">
              {schoolHead || "Signature of Principal/School Head"}
            </div>
            <div className="border-t border-black">
              (Affix School Seal Here)
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function Sf6CountCells({
  value,
  strong = false,
}: {
  value: Sf6Count;
  strong?: boolean;
}) {
  return (
    <>
      <td
        className={`border border-black px-0.5 py-1.5 text-center tabular-nums ${
          strong ? "font-bold" : ""
        }`}
        style={{ borderLeftWidth: 2 }}
      >
        {value.male}
      </td>
      <td
        className={`border border-black px-0.5 py-1.5 text-center tabular-nums ${
          strong ? "font-bold" : ""
        }`}
      >
        {value.female}
      </td>
      <td
        className={`border border-black px-0.5 py-1.5 text-center tabular-nums ${
          strong ? "font-bold" : ""
        }`}
      >
        {value.total}
      </td>
    </>
  );
}

function Sf6BondPaperLogo() {
  return (
    <img
      src={SF6_DEPED_SEAL_URL}
      alt="Department of Education seal"
      className="absolute left-2 top-0 h-[78px] w-[78px] object-contain"
      onError={(event) => {
        event.currentTarget.onerror = null;
        event.currentTarget.src = logos;
      }}
    />
  );
}

function Sf6BondPaper({
  schoolName,
  schoolId,
  region,
  division,
  district,
  schoolYear,
  gradeSection,
  schoolHead,
  statusRows,
  metrics,
}: {
  schoolName: string;
  schoolId: string;
  region: string;
  division: string;
  district: string;
  schoolYear: string;
  gradeSection: string;
  schoolHead: string;
  statusRows: Array<{
    label: string;
    values: Sf6Count[];
    total: Sf6Count;
  }>;
  metrics: Sf6MetricsData;
}) {
  return (
    <div
      className="box-border h-full overflow-hidden p-[18px] text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="relative min-h-[112px]">
        <Sf6BondPaperLogo />

        <div className="px-[110px] pt-1 text-center">
          <div className="text-[18px] font-bold leading-tight">
            School Form 6 (SF6)
          </div>
          <div className="text-[16px] font-bold leading-tight">
            Summarized Report on Promotion and Learning Progress &amp;
            Achievement
          </div>
        </div>

        <div className="mx-auto mt-3 max-w-[920px] space-y-1.5 text-[8px]">
          <div className="grid grid-cols-[58px_130px_48px_130px_52px_130px_45px_130px] items-center gap-2">
            <span className="text-right">School ID</span>
            <span className="min-h-5 border border-black px-1.5 py-1">
              {schoolId}
            </span>

            <span className="text-right">Region</span>
            <span className="min-h-5 border border-black px-1.5 py-1">
              {region}
            </span>

            <span className="text-right">Division</span>
            <span className="min-h-5 border border-black px-1.5 py-1">
              {division}
            </span>

            <span className="text-right">District</span>
            <span className="min-h-5 border border-black px-1.5 py-1">
              {district}
            </span>
          </div>

          <div className="grid grid-cols-[70px_360px_64px_120px_82px_160px] items-center gap-2">
            <span className="text-right">School Name</span>
            <span className="min-h-5 border border-black px-1.5 py-1">
              {schoolName}
            </span>

            <span className="text-right">School Year</span>
            <span className="min-h-5 border border-black px-1.5 py-1 text-center">
              {schoolYear}
            </span>

            <span className="text-right">Grade &amp; Section</span>
            <span className="min-h-5 border border-black px-1.5 py-1 text-center">
              {gradeSection}
            </span>
          </div>
        </div>
      </div>

      <table className="mt-1 w-full table-fixed border-collapse border-2 border-black text-[7px] leading-tight">
        <colgroup>
          <col style={{ width: "12.5%" }} />
          {Array.from({ length: 21 }).map((_, index) => (
            <col key={`sf6-col-${index}`} />
          ))}
        </colgroup>
        <thead>
          <tr>
            <th
              rowSpan={2}
              className="border border-black px-1 py-2 text-left text-[9px] font-normal"
            >
              SUMMARY TABLE
            </th>
            {SF6_GRADE_LEVELS.map((grade) => (
              <th
                key={grade.value}
                colSpan={3}
                className="border border-black px-1 py-2 text-center text-[9px] font-normal"
                style={{ borderLeftWidth: 2 }}
              >
                {grade.label}
              </th>
            ))}
            <th
              colSpan={3}
              className="border border-black px-1 py-2 text-center text-[10px] font-normal"
              style={{ borderLeftWidth: 2 }}
            >
              TOTAL
            </th>
          </tr>

          <tr>
            {[...SF6_GRADE_LEVELS, { value: "total", label: "TOTAL" }].map(
              (grade) => (
                <Fragment key={grade.value}>
                  <th
                    className="border border-black px-0.5 py-1.5 text-center font-normal"
                    style={{ borderLeftWidth: 2 }}
                  >
                    Male
                  </th>
                  <th className="border border-black px-0.5 py-1.5 text-center font-normal">
                    Female
                  </th>
                  <th className="border border-black px-0.5 py-1.5 text-center font-normal">
                    Total
                  </th>
                </Fragment>
              ),
            )}
          </tr>
        </thead>

        <tbody>
          {statusRows.map((row) => (
            <tr key={row.label}>
              <td className="border border-black px-1 py-2 text-[8px]">
                {row.label}
              </td>
              {row.values.map((value, index) => (
                <Sf6CountCells
                  key={`${row.label}-${SF6_GRADE_LEVELS[index].value}`}
                  value={value}
                />
              ))}
              <Sf6CountCells value={row.total} strong />
            </tr>
          ))}

          <tr>
            <td className="border border-black px-1 py-1.5 text-[5.5px] font-bold uppercase">
              Learning Progress and
              <br />
              Achievement
            </td>
            {[...SF6_GRADE_LEVELS, { value: "total", label: "TOTAL" }].map(
              (grade) => (
                <Fragment key={`sf6-progress-${grade.value}`}>
                  <td
                    className="border border-black px-0.5 py-1.5 text-center"
                    style={{ borderLeftWidth: 2 }}
                  >
                    Male
                  </td>
                  <td className="border border-black px-0.5 py-1.5 text-center">
                    Female
                  </td>
                  <td className="border border-black px-0.5 py-1.5 text-center">
                    Total
                  </td>
                </Fragment>
              ),
            )}
          </tr>

          {SF6_PROGRESS_BANDS.map((band, bandIndex) => {
            const values = metrics.bandValues[bandIndex];

            return (
              <tr key={band.key}>
                <td className="border border-black px-1 py-1.5 text-[6px]">
                  <div>{band.label}</div>
                  <div>{band.detail}</div>
                </td>
                {values.map((value, index) => (
                  <Sf6CountCells
                    key={`${band.key}-${SF6_GRADE_LEVELS[index].value}`}
                    value={value}
                  />
                ))}
                <Sf6CountCells value={addSf6Counts(values)} strong />
              </tr>
            );
          })}

          <tr className="border-t-2 border-black">
            <td className="border border-black px-1 py-2 text-[9px] font-bold">
              TOTAL
            </td>
            {metrics.enrolledByGrade.map((value, index) => (
              <Sf6CountCells
                key={`sf6-enrolled-${SF6_GRADE_LEVELS[index].value}`}
                value={value}
                strong
              />
            ))}
            <Sf6CountCells
              value={addSf6Counts(metrics.enrolledByGrade)}
              strong
            />
          </tr>
        </tbody>
      </table>

      <div className="mt-1 text-[7px]">Prepared and Submitted by:</div>

      <div className="mt-6 grid grid-cols-4 gap-8 px-1 text-center text-[7px]">
        <div>
          <div className="min-h-[12px] text-[8px] uppercase">{schoolHead}</div>
          <div className="border-t-2 border-black pt-1">
            (Signature of School Head/SCC Chair)
          </div>
        </div>

        <div>
          <div className="min-h-[12px]" />
          <div className="border-t-2 border-black pt-1">
            SCC-Vice Chair (Curriculum)
          </div>
        </div>

        <div>
          <div className="min-h-[12px]" />
          <div className="border-t-2 border-black pt-1">SCC Member</div>
        </div>

        <div>
          <div className="min-h-[12px]" />
          <div className="border-t-2 border-black pt-1">
            SCC-Vice Chair (Generated thru LIS)
          </div>
        </div>
      </div>

      <div className="mt-4 text-[7px]">
        (Additional slots may be added for SCC members.)
      </div>
    </div>
  );
}

function UserCreateDialog({
  open,
  onClose,
  onSaved,
}: {
  open: boolean;
  onClose: () => void;
  onSaved: () => Promise<void>;
}) {
  const [saving, setSaving] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [successMessage, setSuccessMessage] = useState("");

  const save = async (formData: FormData) => {
    const fullName = String(formData.get("full_name") || "").trim();
    const email = String(formData.get("email") || "")
      .trim()
      .toLowerCase();
    const password = String(formData.get("password") || "");
    const teacherType = String(formData.get("teacher_type") || "class_adviser");

    // Keep the Supabase access role as "teacher".
    // teacher_type is used to separate the Class Adviser and Subject Teacher UI.
    const role = "teacher";

    setErrorMessage("");
    setSuccessMessage("");

    if (!fullName) {
      setErrorMessage("Enter the teacher or user name.");
      return;
    }

    if (!/^\S+@\S+\.\S+$/.test(email)) {
      setErrorMessage("Enter a valid Gmail or DepEd email address.");
      return;
    }

    if (password.length < 8) {
      setErrorMessage("Password must contain at least 8 characters.");
      return;
    }

    if (teacherType !== "class_adviser" && teacherType !== "subject_teacher") {
      setErrorMessage("Select Class Adviser or Subject Teacher.");
      return;
    }

    setSaving(true);

    const { data, error } = await supabase.functions.invoke(
      "admin-create-user",
      {
        body: {
          full_name: fullName,
          email,
          password,
          role,
          teacher_type: teacherType,
        },
      },
    );

    setSaving(false);

    if (error) {
      let message = error.message || "Unable to create the account.";

      try {
        const response = (error as { context?: Response }).context;
        if (response) {
          const body = (await response.clone().json()) as {
            error?: string;
            message?: string;
          };
          message = body.error || body.message || message;
        }
      } catch {
        // Keep the original Supabase function error message.
      }

      setErrorMessage(message);
      return;
    }

    if (!data?.user?.id) {
      setErrorMessage("The server did not return the newly created account.");
      return;
    }

    await onSaved();

    const teacherTypeLabel =
      teacherType === "class_adviser" ? "Class Adviser" : "Subject Teacher";

    setSuccessMessage(
      `${teacherTypeLabel} account created successfully. The user can now sign in with the email and password you entered.`,
    );
  };

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) => {
        if (!nextOpen && !saving) {
          setErrorMessage("");
          setSuccessMessage("");
          onClose();
        }
      }}
    >
      <DialogContent className="border-[var(--admin-dialog-border)] bg-[var(--admin-surface)] sm:max-w-lg">
        {successMessage ? (
          <>
            <DialogHeader>
              <DialogTitle className="text-center text-[var(--admin-text-heading)]">
                Account Created
              </DialogTitle>
            </DialogHeader>

            <div className="rounded-2xl border border-emerald-300 bg-emerald-50/80 px-6 py-7 text-center">
              <div className="mx-auto grid size-16 place-items-center rounded-full bg-emerald-600 text-white shadow-sm">
                <CircleCheck className="size-9" />
              </div>

              <p className="mt-4 text-base font-bold text-emerald-800">
                Success!
              </p>

              <p className="mx-auto mt-2 max-w-sm text-xs leading-5 text-emerald-950/75">
                {successMessage}
              </p>

              <Button
                type="button"
                onClick={() => {
                  setSuccessMessage("");
                  onClose();
                }}
                className="mt-5 min-w-28 bg-emerald-600 text-white hover:bg-emerald-700"
              >
                Done
              </Button>
            </div>
          </>
        ) : (
          <>
            <DialogHeader>
              <DialogTitle className="text-[var(--admin-text-heading)]">
                Add Teacher / User
              </DialogTitle>
            </DialogHeader>

            <form action={save} className="space-y-4">
              <div className="grid gap-4">
                <FormField label="Full Name" name="full_name" required />
                <FormField label="Email " name="email" required />

                <label className="space-y-1.5">
                  <span className="text-[10px] font-semibold text-[var(--admin-form-label)]">
                    Password
                  </span>
                  <Input
                    type="password"
                    name="password"
                    required
                    minLength={8}
                    autoComplete="new-password"
                    placeholder="At least 8 characters"
                    className="border-[var(--admin-input-border)] bg-[var(--admin-white)] text-xs focus-visible:ring-[var(--admin-primary-ring-25)]"
                  />
                  <span className="block text-[9px] text-[var(--admin-text-muted)]">
                    This is the user's sign-in password. It remains active until
                    the user changes or resets it.
                  </span>
                </label>

                <SelectField
                  label="Role"
                  name="teacher_type"
                  defaultValue="class_adviser"
                >
                  <option value="class_adviser">Class Adviser</option>
                  <option value="subject_teacher">Subject Teacher</option>
                </SelectField>
              </div>

              {errorMessage && (
                <div className="rounded-lg border border-[var(--admin-danger-border)] bg-[var(--admin-danger-bg)] px-3 py-2 text-[10px] text-[var(--admin-danger-700)]">
                  {errorMessage}
                </div>
              )}

              <DialogFooter>
                <Button
                  type="button"
                  variant="outline"
                  onClick={onClose}
                  disabled={saving}
                >
                  Cancel
                </Button>
                <Button
                  type="submit"
                  disabled={saving}
                  className="bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
                >
                  {saving ? <Loader2 className="animate-spin" /> : <Plus />}
                  Create Account
                </Button>
              </DialogFooter>
            </form>
          </>
        )}
      </DialogContent>
    </Dialog>
  );
}

function ProfileEditDialog({
  profile,
  onClose,
  onSaved,
}: {
  profile: ProfileRow | null;
  onClose: () => void;
  onSaved: () => Promise<void>;
}) {
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");

  useEffect(() => {
    setSaveError("");
  }, [profile?.id]);

  const save = async (formData: FormData) => {
    if (!profile || saving) return;

    const fullName = String(formData.get("full_name") || "").trim();
    const schoolName = String(formData.get("school_name") || "").trim();
    const schoolId = String(formData.get("school_id") || "").trim();
    const region = String(formData.get("region") || "").trim();
    const division = String(formData.get("division") || "").trim();
    const district = String(formData.get("district") || "").trim();

    if (!fullName) {
      setSaveError("Full Name is required.");
      return;
    }

    setSaving(true);
    setSaveError("");

    // Use a SECURITY DEFINER Admin RPC instead of a direct profiles UPDATE.
    // The profiles table has Admin SELECT RLS access, but a direct UPDATE of
    // another teacher can be filtered out by RLS and appear to save while
    // changing 0 rows.
    const { data: updatedProfile, error } = await (supabase as any).rpc(
      "admin_update_profile",
      {
        target_user_id: profile.id,
        new_full_name: fullName,
        new_school_name: schoolName || null,
        new_school_id: schoolId || null,
        new_region: region || null,
        new_division: division || null,
        new_district: district || null,
      },
    );

    if (error) {
      setSaving(false);
      setSaveError(
        error.message.includes("admin_update_profile")
          ? "The Admin profile update database function is not installed yet. Run the new SQL migration in Supabase, then try again."
          : error.message,
      );
      return;
    }

    const savedRow = Array.isArray(updatedProfile)
      ? updatedProfile[0]
      : updatedProfile;

    if (!savedRow?.id) {
      setSaving(false);
      setSaveError(
        "No profile row was updated. Please run the Admin profile update SQL migration and try again.",
      );
      return;
    }

    try {
      await onSaved();
      onClose();
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog open={Boolean(profile)} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-h-[88vh] overflow-y-auto border-[var(--admin-dialog-border)] bg-[var(--admin-surface)] sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle className="text-[var(--admin-text-heading)]">
            Edit Teacher / User
          </DialogTitle>
          <DialogDescription>
            Update the profile details. The sign-in email is managed in Supabase
            Authentication.
          </DialogDescription>
        </DialogHeader>
        {profile && (
          <form action={save} className="space-y-4">
            <div className="grid gap-4 sm:grid-cols-2">
              <FormField
                label="Full Name"
                name="full_name"
                defaultValue={profile.full_name || ""}
              />
              <FormField
                label="Sign-in Email"
                value={profile.email || ""}
                disabled
              />
              <FormField
                label="School Name"
                name="school_name"
                defaultValue={profile.school_name || ""}
              />
              <FormField
                label="School ID"
                name="school_id"
                defaultValue={profile.school_id || ""}
              />
              <FormField
                label="Region"
                name="region"
                defaultValue={profile.region || ""}
              />
              <FormField
                label="Division"
                name="division"
                defaultValue={profile.division || ""}
              />
              <FormField
                label="District"
                name="district"
                defaultValue={profile.district || ""}
              />
            </div>

            {saveError && (
              <div className="rounded-lg border border-[var(--admin-danger-border)] bg-[var(--admin-danger-bg)] px-3 py-2 text-[10px] leading-4 text-[var(--admin-danger-700)]">
                {saveError}
              </div>
            )}

            <DialogFooter>
              <Button
                type="button"
                variant="outline"
                onClick={onClose}
                disabled={saving}
              >
                Cancel
              </Button>
              <Button
                type="submit"
                disabled={saving}
                className="bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
              >
                {saving ? <Loader2 className="animate-spin" /> : <Save />}
                Save Changes
              </Button>
            </DialogFooter>
          </form>
        )}
      </DialogContent>
    </Dialog>
  );
}

function StudentCreateDialog({
  open,
  onClose,
  onSaved,
}: {
  open: boolean;
  onClose: () => void;
  onSaved: () => Promise<void>;
}) {
  const [saving, setSaving] = useState(false);

  const save = async (formData: FormData) => {
    const firstName = String(formData.get("first_name") || "").trim();
    const lastName = String(formData.get("last_name") || "").trim();
    const lrn = String(formData.get("lrn") || "").trim();

    if (!firstName || !lastName) {
      window.alert("Enter the student's first and last name.");
      return;
    }

    setSaving(true);

    if (lrn) {
      const { data: existingStudent, error: duplicateCheckError } = await (
        supabase as any
      )
        .from("students")
        .select("id")
        .eq("lrn", lrn)
        .maybeSingle();

      if (duplicateCheckError) {
        setSaving(false);
        window.alert(duplicateCheckError.message);
        return;
      }
      if (existingStudent) {
        setSaving(false);
        window.alert("A student with this LRN already exists.");
        return;
      }
    }

    const value = (name: string) =>
      String(formData.get(name) || "").trim() || null;
    const addressParts = [
      value("house_street"),
      value("barangay"),
      value("municipality_city"),
      value("province"),
    ].filter(Boolean);

    const payload = {
      class_id: null,
      teacher_id: null,
      lrn: lrn || null,
      first_name: firstName,
      middle_name: value("middle_name"),
      last_name: lastName,
      sex: value("sex"),
      birthdate: value("birthdate"),
      mother_tongue: value("mother_tongue"),
      ip_ethnic_group: value("ip_ethnic_group"),
      religion: value("religion"),
      house_street: value("house_street"),
      barangay: value("barangay"),
      municipality_city: value("municipality_city"),
      province: value("province"),
      address: addressParts.length > 0 ? addressParts.join(", ") : null,
      father_name: value("father_name"),
      mother_name: value("mother_name"),
      guardian: value("guardian"),
      guardian_relationship: value("guardian_relationship"),
      contact_number: value("contact_number"),
      learning_modality: value("learning_modality"),
      remarks: value("remarks"),
    };

    const { error } = await (supabase as any).from("students").insert(payload);
    setSaving(false);

    if (error) {
      window.alert(error.message);
      return;
    }

    await onSaved();
    onClose();
  };

  return (
    <Dialog open={open} onOpenChange={(nextOpen) => !nextOpen && onClose()}>
      <DialogContent className="max-h-[92vh] overflow-y-auto border-[var(--admin-dialog-border)] bg-[var(--admin-surface)] sm:max-w-5xl">
        <DialogHeader>
          <DialogTitle className="text-[var(--admin-text-heading)]">
            Add Student — SF1 Information
          </DialogTitle>
          <DialogDescription>
            Create the learner record used by School Form 1 and Credential
            Requests. The student becomes searchable by name or LRN after
            saving.
          </DialogDescription>
        </DialogHeader>

        <form action={save} className="space-y-5">
          <StudentFormSection title="Personal Information">
            <FormField
              label="LRN"
              name="lrn"
              placeholder="Enter learner reference number"
            />
            <FormField label="Last Name" name="last_name" required />
            <FormField label="First Name" name="first_name" required />
            <FormField label="Middle Name" name="middle_name" />
            <SelectField label="Sex" name="sex" defaultValue="">
              <option value="">Not specified</option>
              <option value="male">Male</option>
              <option value="female">Female</option>
            </SelectField>
            <FormField label="Birthdate" name="birthdate" type="date" />
            <FormField label="Mother Tongue" name="mother_tongue" />
            <FormField label="IP / Ethnic Group" name="ip_ethnic_group" />
            <FormField label="Religion" name="religion" />
          </StudentFormSection>

          <StudentFormSection title="Address">
            <FormField
              label="House No. / Street / Sitio / Purok"
              name="house_street"
            />
            <FormField label="Barangay" name="barangay" />
            <FormField label="Municipality / City" name="municipality_city" />
            <FormField label="Province" name="province" />
          </StudentFormSection>

          <StudentFormSection title="Parent / Guardian">
            <FormField label="Father's Name" name="father_name" />
            <FormField label="Mother's Maiden Name" name="mother_name" />
            <FormField label="Parent / Guardian" name="guardian" />
            <FormField label="Relationship" name="guardian_relationship" />
            <FormField label="Contact Number" name="contact_number" />
          </StudentFormSection>

          <DialogFooter>
            <Button type="button" variant="outline" onClick={onClose}>
              Cancel
            </Button>
            <Button
              type="submit"
              disabled={saving}
              className="gap-2 bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
            >
              {saving ? (
                <Loader2 className="size-4 animate-spin" />
              ) : (
                <Save className="size-4" />
              )}
              Save Student
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

function StudentFormSection({
  title,
  children,
}: {
  title: string;
  children: ReactNode;
}) {
  return (
    <section className="rounded-2xl border border-[var(--admin-border)] bg-white p-4">
      <h3 className="mb-3 text-sm font-bold text-[var(--admin-text-heading)]">
        {title}
      </h3>
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">{children}</div>
    </section>
  );
}

function StudentEditDialog({
  student,
  classes,
  onClose,
  onSaved,
}: {
  student: StudentRow | null;
  classes: ClassRow[];
  onClose: () => void;
  onSaved: () => Promise<void>;
}) {
  const [saving, setSaving] = useState(false);

  const save = async (formData: FormData) => {
    if (!student) return;
    const classId = String(formData.get("class_id") || student.class_id);
    const selectedClass = classes.find((item) => item.id === classId);
    setSaving(true);
    const patch = {
      class_id: classId,
      teacher_id: selectedClass?.teacher_id || student.teacher_id,
      lrn: String(formData.get("lrn") || "").trim() || null,
      first_name: String(formData.get("first_name") || "").trim(),
      middle_name: String(formData.get("middle_name") || "").trim() || null,
      last_name: String(formData.get("last_name") || "").trim(),
      sex: String(formData.get("sex") || "").trim() || null,
      guardian: String(formData.get("guardian") || "").trim() || null,
      contact_number:
        String(formData.get("contact_number") || "").trim() || null,
      address: String(formData.get("address") || "").trim() || null,
    };
    const { error } = await supabase
      .from("students")
      .update(patch)
      .eq("id", student.id);
    setSaving(false);
    if (error) {
      window.alert(error.message);
      return;
    }
    await onSaved();
    onClose();
  };

  return (
    <Dialog open={Boolean(student)} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-h-[88vh] overflow-y-auto border-[var(--admin-dialog-border)] bg-[var(--admin-surface)] sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle className="text-[var(--admin-text-heading)]">
            Edit Learner / Student
          </DialogTitle>
          <DialogDescription>
            Update the learner information and assigned class.
          </DialogDescription>
        </DialogHeader>
        {student && (
          <form action={save} className="space-y-4">
            <div className="grid gap-4 sm:grid-cols-2">
              <FormField
                label="LRN"
                name="lrn"
                defaultValue={student.lrn || ""}
              />
              <SelectField
                label="Class"
                name="class_id"
                defaultValue={student.class_id}
              >
                {classes.map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.grade_level} - {item.section} · {item.subject}
                  </option>
                ))}
              </SelectField>
              <FormField
                label="First Name"
                name="first_name"
                defaultValue={student.first_name || ""}
                required
              />
              <FormField
                label="Middle Name"
                name="middle_name"
                defaultValue={student.middle_name || ""}
              />
              <FormField
                label="Last Name"
                name="last_name"
                defaultValue={student.last_name || ""}
                required
              />
              <SelectField
                label="Sex"
                name="sex"
                defaultValue={student.sex || ""}
              >
                <option value="">Not specified</option>
                <option value="male">Male</option>
                <option value="female">Female</option>
              </SelectField>
              <FormField
                label="Guardian"
                name="guardian"
                defaultValue={student.guardian || ""}
              />
              <FormField
                label="Contact Number"
                name="contact_number"
                defaultValue={student.contact_number || ""}
              />
              <div className="sm:col-span-2">
                <FormField
                  label="Address"
                  name="address"
                  defaultValue={student.address || ""}
                />
              </div>
            </div>
            <DialogFooter>
              <Button type="button" variant="outline" onClick={onClose}>
                Cancel
              </Button>
              <Button
                type="submit"
                disabled={saving}
                className="bg-[var(--admin-primary)] hover:bg-[var(--admin-primary-hover)]"
              >
                {saving ? <Loader2 className="animate-spin" /> : <Save />}
                Save Changes
              </Button>
            </DialogFooter>
          </form>
        )}
      </DialogContent>
    </Dialog>
  );
}

function PageHeading({
  title,
  description,
}: {
  title: string;
  description: string;
}) {
  return (
    <div>
      <h1 className="text-xl font-bold tracking-[-0.025em] text-[var(--admin-text-title)] sm:text-2xl">
        {title}
      </h1>
      <p className="mt-1 text-[11px] text-[var(--admin-text-muted)]">
        {description}
      </p>
    </div>
  );
}

function Panel({
  children,
  className = "",
}: {
  children: ReactNode;
  className?: string;
}) {
  return (
    <section
      className={`rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface)] p-4 shadow-[var(--admin-shadow-panel)] ${className}`}
    >
      {children}
    </section>
  );
}

function PanelHeader({ title, action }: { title: string; action?: ReactNode }) {
  return (
    <div className="mb-4 flex items-center justify-between gap-4">
      <h2 className="text-xs font-bold text-[var(--admin-text-title)]">
        {title}
      </h2>
      {action}
    </div>
  );
}

type Tone = "orange" | "green" | "blue" | "pink" | "violet";

const toneClasses: Record<Tone, string> = {
  orange: "bg-[var(--admin-tone-orange-bg)] text-[var(--admin-color-da741d)]",
  green: "bg-[var(--admin-tone-green-bg)] text-[var(--admin-tone-green-text)]",
  blue: "bg-[var(--admin-tone-blue-bg)] text-[var(--admin-tone-blue-text)]",
  pink: "bg-[var(--admin-tone-pink-bg)] text-[var(--admin-tone-pink-text)]",
  violet:
    "bg-[var(--admin-tone-violet-bg)] text-[var(--admin-tone-violet-text)]",
};

function StatCard({
  title,
  value,
  helper,
  icon: Icon,
  tone,
  compact = false,
}: {
  title: string;
  value: number | string;
  helper: string;
  icon: ComponentType<{ className?: string }>;
  tone: Tone;
  compact?: boolean;
}) {
  return (
    <div
      className={`rounded-xl border border-[var(--admin-border)] bg-[var(--admin-surface)] ${compact ? "p-3.5" : "p-4"}`}
    >
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-[9px] font-medium text-[var(--admin-color-7e6a61)]">
            {title}
          </div>
          <div
            className={`${compact ? "mt-1 text-xl" : "mt-2 text-2xl"} font-bold leading-none text-[var(--admin-text-value)]`}
          >
            {value}
          </div>
          <div className="mt-1 text-[8px] text-[var(--admin-text-helper)]">
            {helper}
          </div>
        </div>
        <div
          className={`grid size-8 shrink-0 place-items-center rounded-lg ${toneClasses[tone]}`}
        >
          <Icon className="size-4" />
        </div>
      </div>
    </div>
  );
}

function SearchBar({
  value,
  onChange,
  placeholder,
}: {
  value: string;
  onChange: (value: string) => void;
  placeholder: string;
}) {
  return (
    <div className="relative">
      <Search className="pointer-events-none absolute left-3 top-1/2 size-3.5 -translate-y-1/2 text-[var(--admin-color-9b877e)]" />
      <Input
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        className="h-10 rounded-lg border-[var(--admin-input-border)] bg-[var(--admin-surface)] pl-9 text-[11px] placeholder:text-[var(--admin-color-ad9a91)] focus-visible:ring-[var(--admin-primary-ring-25)]"
      />
    </div>
  );
}

function MiniMetric({
  icon: Icon,
  label,
  value,
}: {
  icon: ComponentType<{ className?: string }>;
  label: string;
  value: number;
}) {
  return (
    <div className="flex items-center gap-3 rounded-lg border border-[var(--admin-color-f0e6de)] bg-[var(--admin-surface-soft)] p-3">
      <div className="grid size-8 place-items-center rounded-lg bg-[var(--admin-color-f6e9dd)] text-[var(--admin-primary)]">
        <Icon className="size-4" />
      </div>
      <div>
        <div className="text-[9px] text-[var(--admin-text-caption)]">
          {label}
        </div>
        <div className="text-lg font-bold text-[var(--admin-text-number)]">
          {value}
        </div>
      </div>
    </div>
  );
}

function DataTable({
  children,
  minWidth,
}: {
  children: ReactNode;
  minWidth: string;
}) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-left" style={{ minWidth }}>
        {children}
      </table>
    </div>
  );
}

function Th({
  children,
  align = "left",
}: {
  children: ReactNode;
  align?: "left" | "right";
}) {
  return (
    <th
      className={`border-b border-[var(--admin-color-e9ded5)] bg-[var(--admin-surface-soft)] px-3 py-2.5 text-[9px] font-bold text-[var(--admin-color-6d5a52)] ${align === "right" ? "text-right" : "text-left"}`}
    >
      {children}
    </th>
  );
}

function Td({
  children,
  align = "left",
  strong = false,
}: {
  children: ReactNode;
  align?: "left" | "right";
  strong?: boolean;
}) {
  return (
    <td
      className={`px-3 py-3 text-[10px] ${align === "right" ? "text-right" : "text-left"} ${strong ? "font-semibold text-[var(--admin-color-3b2b25)]" : "text-[var(--admin-color-68564e)]"}`}
    >
      {children}
    </td>
  );
}

function TableEmpty({ colSpan, text }: { colSpan: number; text: string }) {
  return (
    <tr>
      <td
        colSpan={colSpan}
        className="px-4 py-12 text-center text-[11px] text-[var(--admin-text-empty)]"
      >
        {text}
      </td>
    </tr>
  );
}

function EmptyState({ text }: { text: string }) {
  return (
    <div className="py-8 text-center text-[11px] text-[var(--admin-text-empty)]">
      {text}
    </div>
  );
}

function RoleBadge({ role }: { role: string }) {
  const normalizedRole = normalize(role);
  const admin = normalizedRole === "admin";
  const classAdviser = normalizedRole === "class adviser";

  return (
    <span
      className={`inline-flex rounded-full px-2 py-1 text-[8px] font-semibold ${
        admin
          ? "bg-[var(--admin-color-f0e9ff)] text-[var(--admin-color-7045b2)]"
          : classAdviser
            ? "bg-emerald-100 text-emerald-700"
            : "bg-[var(--admin-tone-orange-bg)] text-[var(--admin-color-ad651f)]"
      }`}
    >
      {role}
    </span>
  );
}

function StatusBadge({ children }: { children: ReactNode }) {
  return (
    <span className="inline-flex items-center gap-1 rounded-full bg-[var(--admin-color-e7f7eb)] px-2 py-1 text-[8px] font-semibold text-[var(--admin-color-27844a)]">
      <CircleCheck className="size-2.5" />
      {children}
    </span>
  );
}

function RowActions({
  busy,
  disableDelete = false,
  onEdit,
  onDelete,
}: {
  busy: boolean;
  disableDelete?: boolean;
  onEdit: () => void;
  onDelete: () => void;
}) {
  return (
    <div className="inline-flex items-center gap-1.5">
      <button
        type="button"
        onClick={onEdit}
        className="rounded-md border border-[var(--admin-color-ddd0c5)] bg-[var(--admin-white)] p-1.5 text-[var(--admin-text-action)] hover:border-[var(--admin-primary)] hover:text-[var(--admin-primary)]"
        title="Edit"
      >
        <Pencil className="size-3" />
      </button>
      <button
        type="button"
        onClick={onDelete}
        disabled={busy || disableDelete}
        className="rounded-md border border-[var(--admin-danger-border)] bg-[var(--admin-white)] p-1.5 text-[var(--admin-danger-500)] hover:bg-[var(--admin-danger-bg)] disabled:cursor-not-allowed disabled:opacity-40"
        title={
          disableDelete ? "You cannot delete your current account" : "Delete"
        }
      >
        {busy ? (
          <Loader2 className="size-3 animate-spin" />
        ) : (
          <Trash2 className="size-3" />
        )}
      </button>
    </div>
  );
}

function PaginationFooter({
  total,
  pageSize,
  page,
  totalPages,
  onPage,
}: {
  total: number;
  pageSize: number;
  page: number;
  totalPages: number;
  onPage: (page: number) => void;
}) {
  const start = total === 0 ? 0 : (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, total);
  const pages = paginationNumbers(page, totalPages);

  return (
    <div className="mt-4 flex flex-col gap-3 border-t border-[var(--admin-color-eee4dc)] pt-4 sm:flex-row sm:items-center sm:justify-between">
      <div className="text-[9px] text-[var(--admin-text-caption)]">
        Showing {start} to {end} of {total} entries
      </div>
      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={() => onPage(Math.max(1, page - 1))}
          disabled={page <= 1}
          className="grid size-7 place-items-center rounded-md border border-[var(--admin-pagination-border)] bg-[var(--admin-white)] text-[var(--admin-text-action)] disabled:opacity-35"
        >
          <ChevronLeft className="size-3" />
        </button>
        {pages.map((item, index) =>
          item === "..." ? (
            <span
              key={`dots-${index}`}
              className="grid size-7 place-items-center text-[9px] text-[var(--admin-text-subtle)]"
            >
              ...
            </span>
          ) : (
            <button
              key={item}
              type="button"
              onClick={() => onPage(item)}
              className={`grid size-7 place-items-center rounded-md border text-[9px] font-semibold ${item === page ? "border-[var(--admin-primary)] bg-[var(--admin-primary)] text-[var(--admin-white)]" : "border-[var(--admin-pagination-border)] bg-[var(--admin-white)] text-[var(--admin-text-action)]"}`}
            >
              {item}
            </button>
          ),
        )}
        <button
          type="button"
          onClick={() => onPage(Math.min(totalPages, page + 1))}
          disabled={page >= totalPages}
          className="grid size-7 place-items-center rounded-md border border-[var(--admin-pagination-border)] bg-[var(--admin-white)] text-[var(--admin-text-action)] disabled:opacity-35"
        >
          <ChevronRight className="size-3" />
        </button>
      </div>
    </div>
  );
}

function FormField({
  label,
  name,
  defaultValue,
  value,
  disabled,
  required,
  type = "text",
  placeholder,
}: {
  label: string;
  name?: string;
  defaultValue?: string;
  value?: string;
  disabled?: boolean;
  required?: boolean;
  type?: string;
  placeholder?: string;
}) {
  return (
    <label className="space-y-1.5">
      <span className="text-[10px] font-semibold text-[var(--admin-form-label)]">
        {label}
      </span>
      <Input
        type={type}
        name={name}
        defaultValue={defaultValue}
        value={value}
        disabled={disabled}
        required={required}
        placeholder={placeholder}
        readOnly={value !== undefined}
        className="border-[var(--admin-input-border)] bg-[var(--admin-white)] text-xs focus-visible:ring-[var(--admin-primary-ring-25)]"
      />
    </label>
  );
}

function SelectField({
  label,
  name,
  defaultValue,
  children,
}: {
  label: string;
  name: string;
  defaultValue: string;
  children: ReactNode;
}) {
  return (
    <label className="space-y-1.5">
      <span className="text-[10px] font-semibold text-[var(--admin-form-label)]">
        {label}
      </span>
      <select
        name={name}
        defaultValue={defaultValue}
        className="h-9 w-full rounded-md border border-[var(--admin-input-border)] bg-[var(--admin-white)] px-3 text-xs outline-none focus:ring-2 focus:ring-[var(--admin-primary-ring-20)]"
      >
        {children}
      </select>
    </label>
  );
}

function PreviewField({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center border-b border-[var(--admin-color-d7ccc3)] pb-1">
      <span className="w-24 font-semibold text-[var(--admin-color-55433c)]">
        {label}:
      </span>
      <span className="flex-1 text-[var(--admin-color-75635b)]">{value}</span>
    </div>
  );
}

function formTitle(form: string) {
  const titles: Record<string, string> = {
    GSA: "General School Agreement",
    SF5: "Report on Promotion and Level of Proficiency",
    SOG: "Summary of Grades per Term",
    SF1: "School Register",
    SF9_NEW: "Learner Progress and Performance Report",
    SF6: "Summarized Report on Promotion",
    SF8: "Learner Health and Nutrition Profile",
    SF10: "Learner's Permanent Academic Record",
    ANECDOTAL: "Anecdotal Record",
    ITEM_ANALYSIS: "Item Analysis Report",
  };

  return titles[form] || form;
}

function registrationTone(index: number) {
  const tones = [
    "bg-[var(--admin-tone-orange-bg)] text-[var(--admin-color-c76b1e)]",
    "bg-[var(--admin-tone-pink-bg)] text-[var(--admin-tone-pink-text)]",
    "bg-[var(--admin-tone-violet-bg)] text-[var(--admin-tone-violet-text)]",
    "bg-[var(--admin-tone-green-bg)] text-[var(--admin-tone-green-text)]",
    "bg-[var(--admin-tone-blue-bg)] text-[var(--admin-tone-blue-text)]",
  ];
  return tones[index % tones.length];
}

function paginate<T>(items: T[], requestedPage: number, pageSize: number) {
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  const currentPage = Math.min(Math.max(1, requestedPage), totalPages);
  const start = (currentPage - 1) * pageSize;
  return {
    items: items.slice(start, start + pageSize),
    currentPage,
    totalPages,
  };
}

function paginationNumbers(
  page: number,
  totalPages: number,
): Array<number | "..."> {
  if (totalPages <= 5)
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  if (page <= 3) return [1, 2, 3, "...", totalPages];
  if (page >= totalPages - 2)
    return [1, "...", totalPages - 2, totalPages - 1, totalPages];
  return [1, "...", page, "...", totalPages];
}
