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
import { Label } from "@/components/ui/label";
import { Skeleton } from "@/components/ui/skeleton";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from "@/components/ui/alert-dialog";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Atom,
  BookMarked,
  BookOpen,
  BriefcaseBusiness,
  Calculator,
  Check,
  Dumbbell,
  Eye,
  FileSearch,
  FlaskConical,
  Globe2,
  HeartHandshake,
  HeartPulse,
  Inbox,
  Landmark,
  Languages,
  Loader2,
  MoreHorizontal,
  Music2,
  Palette,
  Plus,
  Target,
  Users,
  X,
  type LucideIcon,
} from "lucide-react";
import { useState } from "react";
import { toast } from "sonner";
import { getUserId, type ClassRow } from "@/lib/data";

type SchoolYearSetting = {
  id: number;
  school_year: string | null;
  is_locked: boolean;
  updated_by: string | null;
  updated_at: string;
};

async function getSchoolYearSetting(): Promise<SchoolYearSetting> {
  // school_year_settings is created by the School Year migration.
  // `as any` is intentional because generated Supabase types may not
  // include this new table yet.
  const { data, error } = await (supabase as any)
    .from("school_year_settings")
    .select("id, school_year, is_locked, updated_by, updated_at")
    .eq("id", 1)
    .maybeSingle();

  if (error) throw error;

  return {
    id: 1,
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

async function resolveAllowedSchoolYear(
  selectedSchoolYear?: string | null,
): Promise<string> {
  const setting = await getSchoolYearSetting();

  if (setting.is_locked && setting.school_year) {
    return setting.school_year;
  }

  const requestedSchoolYear = String(selectedSchoolYear ?? "").trim();

  if (!requestedSchoolYear) {
    throw new Error(
      "Select a School Year from the Administrator's School Year Library.",
    );
  }

  const { data, error } = await (supabase as any)
    .from("school_year_library")
    .select("school_year")
    .eq("school_year", requestedSchoolYear)
    .maybeSingle();

  if (error) throw error;

  if (!data?.school_year) {
    throw new Error(
      "The selected School Year is not available in the Administrator's School Year Library.",
    );
  }

  return String(data.school_year);
}

export const Route = createFileRoute(
  "/_authenticated/subject-teacher-dashboard",
)({
  component: SubjectTeacherDashboard,
});

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

const UNITS_BY_GRADE: Record<string, readonly string[]> = {
  "Grade 11": ["2", "3", "6"],
  "Grade 12": ["3", "none"],
};

const emptyClassForm = {
  subject: "",
  grade_level: "",
  section: "",
  school_year: "",
  units: "",
};

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

function SubjectCardArtwork({ theme }: { theme: SubjectTheme }) {
  const commonClass = `absolute inset-y-0 right-0 w-[58%] overflow-hidden ${theme.artworkClass}`;

  if (theme.kind === "mathematics") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute -right-5 top-1 rotate-[-8deg] text-right font-serif text-lg font-bold leading-8">
          <div>∑ x² + y²</div>
          <div>f(x) = mx + b</div>
          <div>πr² · √144</div>
          <div>∫ 2x dx</div>
        </div>
        <Calculator className="absolute -bottom-7 right-24 size-28 rotate-12 opacity-40" />
      </div>
    );
  }

  if (theme.kind === "filipino") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute -right-8 -top-10 size-32 rounded-full border-[16px] border-amber-500/20" />
        <div className="absolute right-3 top-8 rotate-[-7deg] text-right font-serif font-bold tracking-widest">
          <div className="text-3xl">WIKA</div>
          <div className="mt-1 text-xs tracking-[0.28em]">
            PANITIKAN · KULTURA
          </div>
        </div>
        <Languages className="absolute -bottom-8 right-20 size-28 -rotate-6 opacity-45" />
      </div>
    );
  }

  if (theme.kind === "mapeh") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Music2 className="absolute right-8 top-4 size-14 -rotate-12" />
        <Palette className="absolute right-24 top-16 size-12 rotate-12" />
        <Dumbbell className="absolute bottom-3 right-5 size-14 rotate-6" />
        <HeartPulse className="absolute bottom-3 right-32 size-12 -rotate-6" />
      </div>
    );
  }

  if (theme.kind === "science") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Atom className="absolute -right-5 -top-4 size-36 rotate-12 opacity-65" />
        <FlaskConical className="absolute bottom-1 right-24 size-16 -rotate-6" />
        <div className="absolute bottom-3 right-2 font-mono text-sm font-bold">
          H₂O · F=ma · DNA
        </div>
      </div>
    );
  }

  if (theme.kind === "social-studies") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Globe2 className="absolute -right-5 -top-5 size-32 opacity-60" />
        <Landmark className="absolute bottom-0 right-24 size-20" />
        <div className="absolute right-2 top-24 text-[10px] font-bold tracking-[0.22em]">
          KASAYSAYAN · LIPUNAN
        </div>
      </div>
    );
  }

  if (theme.kind === "research") {
    return (
      <div className={commonClass} aria-hidden="true">
        <FileSearch className="absolute -right-4 top-0 size-32 rotate-6 opacity-60" />
        <div className="absolute bottom-3 right-5 font-mono text-xs font-bold tracking-widest">
          QUESTION · DATA · EVIDENCE
        </div>
      </div>
    );
  }

  if (theme.kind === "language") {
    return (
      <div className={commonClass} aria-hidden="true">
        <div className="absolute right-5 top-0 font-serif text-8xl font-bold leading-none opacity-60">
          “
        </div>
        <BookMarked className="absolute bottom-0 right-20 size-24 -rotate-6" />
        <div className="absolute bottom-4 right-2 text-[10px] font-bold tracking-[0.22em]">
          READ · WRITE · SPEAK
        </div>
      </div>
    );
  }

  if (theme.kind === "career") {
    return (
      <div className={commonClass} aria-hidden="true">
        <Target className="absolute -right-5 -top-5 size-32 opacity-60" />
        <BriefcaseBusiness className="absolute bottom-0 right-24 size-20 -rotate-6" />
        <div className="absolute bottom-4 right-2 text-[10px] font-bold tracking-[0.2em]">
          PLAN · GROW · LEAD
        </div>
      </div>
    );
  }

  if (theme.kind === "values") {
    return (
      <div className={commonClass} aria-hidden="true">
        <HeartHandshake className="absolute -right-2 -top-3 size-32 rotate-6 opacity-65" />
        <div className="absolute bottom-4 right-3 text-[10px] font-bold tracking-[0.2em]">
          DIGNITY · RESPECT · SERVICE
        </div>
      </div>
    );
  }

  return (
    <div className={commonClass} aria-hidden="true">
      <BookOpen className="absolute -bottom-8 -right-3 size-36 -rotate-6 opacity-60" />
    </div>
  );
}


type StudentAssignmentStatus = "pending" | "accepted" | "declined";

type StudentAssignment = {
  id: string;
  class_id: string;
  subject: string;
  status: StudentAssignmentStatus;
  sent_at: string;
  responded_at: string | null;
  decline_reason: string | null;
  grade_level: string;
  section: string;
  school_year: string;
  adviser_name: string;
  learner_count: number;
};

type AssignedLearner = {
  id: string;
  lrn: string | null;
  first_name: string;
  middle_name: string | null;
  last_name: string;
  suffix: string | null;
  sex: string | null;
};

function SubjectTeacherDashboardSkeleton() {
  return (
    <div
      className="space-y-6 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading classes and assignments...</span>

      <div className="space-y-2" aria-hidden="true">
        <Skeleton className="h-8 w-40" />
        <Skeleton className="h-4 w-[24rem] max-w-full" />
      </div>

      <div className="grid gap-3 sm:grid-cols-3" aria-hidden="true">
        {Array.from({ length: 3 }).map((_, index) => (
          <div
            key={index}
            className="min-h-[132px] rounded-2xl border border-amber-200/80 bg-gradient-to-r from-[#fffdf8] via-[#fffbf2] to-[#fff8e8] p-5 shadow-sm"
          >
            <div className="flex min-h-[92px] flex-col justify-between">
              <div className="flex items-start justify-between gap-3">
                <Skeleton className="h-4 w-32" />
                <Skeleton className="size-10 shrink-0 rounded-xl" />
              </div>
              <Skeleton className="h-8 w-12" />
            </div>
          </div>
        ))}
      </div>

      <section className="space-y-3" aria-hidden="true">
        <div className="space-y-2">
          <Skeleton className="h-6 w-40" />
          <Skeleton className="h-4 w-[28rem] max-w-full" />
        </div>

        <div className="grid gap-3 lg:grid-cols-2">
          {Array.from({ length: 2 }).map((_, index) => (
            <div key={index} className="rounded-2xl border bg-card p-4">
              <div className="flex items-start justify-between gap-3">
                <div className="space-y-2">
                  <div className="flex items-center gap-2">
                    <Skeleton className="h-5 w-28" />
                    <Skeleton className="h-6 w-32 rounded-full" />
                  </div>
                  <Skeleton className="h-4 w-52" />
                </div>
                <Skeleton className="size-10 shrink-0 rounded-lg" />
              </div>

              <div className="mt-4 grid grid-cols-2 gap-3 rounded-xl bg-muted/40 p-3">
                <div className="space-y-2">
                  <Skeleton className="h-3 w-24" />
                  <Skeleton className="h-4 w-32" />
                </div>
                <div className="space-y-2">
                  <Skeleton className="h-3 w-16" />
                  <Skeleton className="h-4 w-8" />
                </div>
              </div>

              <div className="mt-4 flex flex-wrap justify-end gap-2">
                <Skeleton className="h-9 w-32 rounded-md" />
                <Skeleton className="h-9 w-24 rounded-md" />
                <Skeleton className="h-9 w-32 rounded-md" />
              </div>
            </div>
          ))}
        </div>
      </section>

      <section className="space-y-3" aria-hidden="true">
        <div className="flex items-center justify-between gap-3">
          <Skeleton className="h-6 w-20" />
          <Skeleton className="h-9 w-28 rounded-md" />
        </div>

        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {Array.from({ length: 6 }).map((_, index) => (
            <div key={index} className="rounded-2xl border bg-card p-4">
              <div className="flex items-start justify-between gap-3">
                <div className="flex-1 space-y-3">
                  <Skeleton className="size-10 rounded-lg" />
                  <Skeleton
                    className={`h-4 ${index % 2 === 0 ? "w-32" : "w-24"}`}
                  />
                  <Skeleton className="h-3.5 w-48 max-w-full" />
                </div>
                <Skeleton className="size-8 rounded-md" />
              </div>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function AssignmentLearnersSkeleton() {
  return (
    <div
      className="divide-y"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading assigned students...</span>
      {Array.from({ length: 5 }).map((_, index) => (
        <div
          key={index}
          className="flex gap-3 px-4 py-3"
          aria-hidden="true"
        >
          <Skeleton className="h-4 w-6 shrink-0" />
          <div className="flex-1 space-y-2">
            <Skeleton
              className={`h-4 ${index % 2 === 0 ? "w-48" : "w-40"}`}
            />
            <Skeleton className="h-3 w-32" />
          </div>
        </div>
      ))}
    </div>
  );
}

function SubjectTeacherDashboard() {
  const qc = useQueryClient();
  const { data: classes = [], isLoading: classesLoading } = useQuery({
    queryKey: ["classes"],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("classes")
        .select("*")
        .order("created_at", { ascending: false });
      if (error) throw error;
      return data as ClassRow[];
    },
  });
  const { data: studentsCount = 0, isLoading: studentsCountLoading } = useQuery({
    queryKey: ["students-count"],
    queryFn: async () => {
      const { count } = await supabase
        .from("students")
        .select("*", { count: "exact", head: true });
      return count ?? 0;
    },
  });

  const { data: studentAssignments = [], isLoading: assignmentsLoading } =
    useQuery({
      queryKey: ["my-student-list-assignments"],
      queryFn: async () => {
        const { data, error } = await (supabase as any).rpc(
          "list_my_student_list_assignments",
        );
        if (error) throw error;
        return (data ?? []) as StudentAssignment[];
      },
    });

  const pendingAssignments = studentAssignments.filter(
    (assignment) => assignment.status === "pending",
  );
  const acceptedAssignments = studentAssignments.filter(
    (assignment) => assignment.status === "accepted",
  );

  const { data: schoolYearSetting, isLoading: schoolYearLoading } =
    useQuery<SchoolYearSetting>({
      queryKey: ["school-year-setting"],
      queryFn: getSchoolYearSetting,
      refetchInterval: 15000,
      refetchOnWindowFocus: true,
    });

  const activeSchoolYear =
    schoolYearSetting?.is_locked && schoolYearSetting.school_year
      ? schoolYearSetting.school_year
      : null;

  const { data: schoolYearLibrary = [], isLoading: schoolYearLibraryLoading } =
    useQuery<SchoolYearLibraryRow[]>({
      queryKey: ["school-year-library"],
      queryFn: getSchoolYearLibrary,
      refetchInterval: 15000,
      refetchOnWindowFocus: true,
    });

  const [open, setOpen] = useState(false);
  const [selectedAssignment, setSelectedAssignment] =
    useState<StudentAssignment | null>(null);
  const [declineReason, setDeclineReason] = useState("");
  const [form, setForm] = useState(emptyClassForm);

  const createClass = useMutation({
    mutationFn: async () => {
      const teacher_id = await getUserId();
      const allowedSchoolYear = await resolveAllowedSchoolYear(
        activeSchoolYear ?? form.school_year,
      );
      const units =
        form.units && form.units !== "none" ? Number(form.units) : null;
      const { error } = await (supabase as any).from("classes").insert({
        subject: form.subject,
        grade_level: form.grade_level,
        section: form.section,
        school_year: allowedSchoolYear,
        units,
        teacher_id,
      });
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success("Class created");
      setOpen(false);
      setForm(emptyClassForm);
      qc.invalidateQueries({ queryKey: ["classes"] });
    },
    onError: (e: Error) => toast.error(e.message),
  });

  const deleteClass = useMutation({
    mutationFn: async (klass: ClassRow) => {
      const assignmentId = (
        klass as ClassRow & {
          student_list_assignment_id?: string | null;
        }
      ).student_list_assignment_id;

      const { error } = assignmentId
        ? await (supabase as any).rpc("delete_student_list_assignment", {
            p_assignment_id: assignmentId,
          })
        : await supabase.from("classes").delete().eq("id", klass.id);
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success(
        "Class permanently deleted. The adviser can send it again.",
      );
      qc.invalidateQueries({ queryKey: ["classes"] });
      qc.invalidateQueries({ queryKey: ["students-count"] });
      qc.invalidateQueries({ queryKey: ["my-student-list-assignments"] });
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const respondToAssignment = useMutation({
    mutationFn: async ({
      assignmentId,
      response,
      reason,
    }: {
      assignmentId: string;
      response: "accepted" | "declined";
      reason?: string;
    }) => {
      const functionName =
        response === "accepted"
          ? "accept_student_list_assignment"
          : "decline_student_list_assignment";
      const params =
        response === "accepted"
          ? { p_assignment_id: assignmentId }
          : { p_assignment_id: assignmentId, p_reason: reason?.trim() || null };
      const { error } = await (supabase as any).rpc(functionName, params);
      if (error) throw error;
    },
    onSuccess: (_, variables) => {
      toast.success(
        variables.response === "accepted"
          ? "Students accepted. The assigned class is now available."
          : "Student assignment declined.",
      );
      setSelectedAssignment(null);
      setDeclineReason("");
      qc.invalidateQueries({ queryKey: ["my-student-list-assignments"] });
      qc.invalidateQueries({ queryKey: ["classes"] });
      qc.invalidateQueries({ queryKey: ["students-count"] });
    },
    onError: (error: Error) => toast.error(error.message),
  });

  if (
    classesLoading ||
    studentsCountLoading ||
    assignmentsLoading ||
    schoolYearLoading ||
    schoolYearLibraryLoading
  ) {
    return <SubjectTeacherDashboardSkeleton />;
  }

  return (
    <div className="space-y-6 pb-24 md:pb-6">
      <div>
        <h1 className="text-2xl font-semibold">My Classes</h1>
        <p className="text-sm text-muted-foreground">
          Everything you enter here flows into every DepEd form.
        </p>
      </div>

      <div className="grid gap-3 sm:grid-cols-3">
        <StatCard label="Classes" value={classes.length} icon={BookOpen} />
        <StatCard label="Students" value={studentsCount} icon={Users} />
        <StatCard
          label="Pending Assignments"
          value={pendingAssignments.length}
          icon={Inbox}
        />
      </div>

      <section className="space-y-3">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-lg font-semibold">Student Assigned</h2>
            <p className="text-sm text-muted-foreground">
              Accept the student list before managing your assigned subject.
            </p>
          </div>
        </div>

        {pendingAssignments.length === 0 ? (
          <div className="rounded-2xl border border-dashed p-5 text-sm text-muted-foreground">
            No pending student assigned.
          </div>
        ) : (
          <div className="grid gap-3 lg:grid-cols-2">
            {pendingAssignments.map((assignment) => (
              <div
                key={assignment.id}
                className="rounded-2xl border bg-card p-4"
              >
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <div className="flex flex-wrap items-center gap-2">
                      <h3 className="font-semibold">{assignment.subject}</h3>
                      <StatusBadge status={assignment.status} />
                    </div>
                    <p className="mt-1 text-sm text-muted-foreground">
                      {assignment.grade_level} · {assignment.section} ·{" "}
                      {assignment.school_year}
                    </p>
                  </div>
                  <div className="grid size-10 shrink-0 place-items-center rounded-lg bg-accent text-accent-foreground">
                    <Users className="size-5" />
                  </div>
                </div>

                <div className="mt-4 grid grid-cols-2 gap-3 rounded-xl bg-muted/40 p-3 text-sm">
                  <div>
                    <p className="text-xs text-muted-foreground">
                      Class Adviser
                    </p>
                    <p className="font-medium">
                      {assignment.adviser_name || "—"}
                    </p>
                  </div>
                  <div>
                    <p className="text-xs text-muted-foreground">Students</p>
                    <p className="font-medium">{assignment.learner_count}</p>
                  </div>
                </div>

                <div className="mt-4 flex flex-wrap justify-end gap-2">
                  <Button
                    variant="outline"
                    onClick={() => setSelectedAssignment(assignment)}
                  >
                    <Eye className="mr-1 size-4" /> View Students
                  </Button>
                  <Button
                    variant="outline"
                    className="text-destructive"
                    disabled={respondToAssignment.isPending}
                    onClick={() =>
                      respondToAssignment.mutate({
                        assignmentId: assignment.id,
                        response: "declined",
                      })
                    }
                  >
                    <X className="mr-1 size-4" /> Decline
                  </Button>
                  <Button
                    disabled={respondToAssignment.isPending}
                    onClick={() =>
                      respondToAssignment.mutate({
                        assignmentId: assignment.id,
                        response: "accepted",
                      })
                    }
                  >
                    <Check className="mr-1 size-4" /> Accept Students
                  </Button>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      <div className="flex items-center justify-between">
        <h2 className="text-lg font-semibold">Classes</h2>
        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button>
              <Plus className="size-4 mr-1" />
              Add class
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>New class</DialogTitle>
            </DialogHeader>
            <div className="space-y-3">
              <div>
                <Label>Grade level</Label>
                <Select
                  value={form.grade_level || undefined}
                  onValueChange={(v) =>
                    setForm((current) => ({
                      ...current,
                      grade_level: v,
                      subject: "",
                      units: "",
                    }))
                  }
                >
                  <SelectTrigger>
                    <SelectValue placeholder="-- Select Grade --" />
                  </SelectTrigger>
                  <SelectContent>
                    {[
                      "Grade 7",
                      "Grade 8",
                      "Grade 9",
                      "Grade 10",
                      "Grade 11",
                      "Grade 12",
                    ].map((g) => (
                      <SelectItem key={g} value={g}>
                        {g}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {UNITS_BY_GRADE[form.grade_level] && (
                <div>
                  <Label>Units</Label>
                  <Select
                    value={form.units || undefined}
                    onValueChange={(v) => setForm({ ...form, units: v })}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="-- Select Units --" />
                    </SelectTrigger>
                    <SelectContent>
                      {UNITS_BY_GRADE[form.grade_level].map((unit) => (
                        <SelectItem key={unit} value={unit}>
                          {unit === "none" ? "None" : `${unit} Units`}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Subject</Label>
                  <Select
                    value={form.subject || undefined}
                    onValueChange={(v) => setForm({ ...form, subject: v })}
                    disabled={!form.grade_level}
                  >
                    <SelectTrigger>
                      <SelectValue
                        placeholder={
                          form.grade_level
                            ? "-- Select Subject --"
                            : "-- Select Grade First --"
                        }
                      />
                    </SelectTrigger>
                    <SelectContent>
                      {(SUBJECTS_BY_GRADE[form.grade_level] ?? []).map(
                        (sub) => (
                          <SelectItem key={sub} value={sub}>
                            {sub}
                          </SelectItem>
                        ),
                      )}
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label>Section</Label>
                  <Input
                    value={form.section}
                    onChange={(e) =>
                      setForm({ ...form, section: e.target.value })
                    }
                    placeholder="Enter section"
                  />
                </div>
              </div>
              <div>
                <Label>School year</Label>

                {activeSchoolYear ? (
                  <Input
                    value={activeSchoolYear}
                    readOnly
                    aria-readonly="true"
                    className="cursor-not-allowed bg-emerald-50/50 text-foreground"
                  />
                ) : (
                  <Select
                    value={form.school_year || undefined}
                    onValueChange={(value) =>
                      setForm((current) => ({
                        ...current,
                        school_year: value,
                      }))
                    }
                    disabled={schoolYearLibrary.length === 0}
                  >
                    <SelectTrigger
                      className={
                        schoolYearLibrary.length === 0
                          ? "border-amber-300 bg-amber-50/60"
                          : ""
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
                  className={`mt-1 text-xs ${
                    activeSchoolYear
                      ? "text-emerald-700"
                      : schoolYearLibrary.length > 0
                        ? "text-amber-700"
                        : "text-destructive"
                  }`}
                >
                  {activeSchoolYear
                    ? `Locked by Admin: ${activeSchoolYear}. Only this School Year is allowed.`
                    : schoolYearLibrary.length > 0
                      ? "No School Year is locked. Select one from the Administrator's School Year Library."
                      : "No School Years are available. Ask the Administrator to add a School Year first."}
                </p>
              </div>
            </div>
            <DialogFooter>
              <Button variant="outline" onClick={() => setOpen(false)}>
                Cancel
              </Button>
              <Button
                onClick={() => createClass.mutate()}
                disabled={
                  !form.subject ||
                  !form.section ||
                  !(activeSchoolYear || form.school_year) ||
                  (Boolean(UNITS_BY_GRADE[form.grade_level]) && !form.units) ||
                  createClass.isPending
                }
                title={
                  !activeSchoolYear && !form.school_year
                    ? "Select a School Year from the Administrator's School Year Library."
                    : undefined
                }
              >
                {createClass.isPending ? "Creating..." : "Create"}
              </Button>
            </DialogFooter>
          </DialogContent>
        </Dialog>
      </div>

      {classes.length === 0 ? (
        <div className="rounded-2xl border border-dashed p-12 text-center">
          <div className="mx-auto grid size-12 place-items-center rounded-xl bg-accent text-accent-foreground">
            <BookOpen className="size-6" />
          </div>
          <p className="mt-3 font-medium">No classes yet</p>
          <p className="text-sm text-muted-foreground">
            Add your first class to get started.
          </p>
        </div>
      ) : (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {classes.map((c) => {
            const theme = getSubjectTheme(c.subject);
            const SubjectIcon = theme.Icon;

            return (
              <div
                key={c.id}
                className={`group relative isolate min-h-40 overflow-hidden rounded-2xl border p-4 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg ${theme.cardClass}`}
              >
                <SubjectCardArtwork theme={theme} />
                <div className="pointer-events-none absolute inset-0 bg-gradient-to-r from-white/90 via-white/70 to-white/5" />

                <div className="relative z-10 flex h-full items-start justify-between gap-3">
                  <Link
                    to="/classes/$classId"
                    params={{ classId: c.id }}
                    className="flex min-w-0 flex-1 flex-col self-stretch rounded-xl outline-none focus-visible:ring-2 focus-visible:ring-primary/60"
                  >
                    <div className="flex items-center gap-2">
                      <div
                        className={`grid size-11 shrink-0 place-items-center rounded-xl shadow-sm ring-1 backdrop-blur-sm ${theme.iconBackgroundClass}`}
                      >
                        <SubjectIcon className={`size-5 ${theme.iconClass}`} />
                      </div>

                      <span
                        className={`max-w-[150px] truncate rounded-full border px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wide backdrop-blur-sm ${theme.badgeClass}`}
                      >
                        {theme.category}
                      </span>
                    </div>

                    <div className="mt-auto pt-5">
                      <div className="truncate text-lg font-semibold text-foreground">
                        {c.subject}
                      </div>

                      <div className="mt-0.5 truncate text-sm text-foreground/65">
                        {c.grade_level} · {c.section} · {c.school_year}
                        {(c as ClassRow & { units?: number | null }).units != null
                          ? ` · ${(c as ClassRow & { units: number }).units} Units`
                          : ""}
                      </div>
                    </div>
                  </Link>

                  <DropdownMenu>
                    <DropdownMenuTrigger asChild>
                      <Button
                        variant="ghost"
                        size="icon"
                        aria-label={`Open ${c.subject} class actions`}
                        className="relative z-20 shrink-0 bg-white/45 backdrop-blur-sm hover:bg-white/80"
                      >
                        <MoreHorizontal className="size-4" />
                      </Button>
                    </DropdownMenuTrigger>

                    <DropdownMenuContent align="end">
                      <AlertDialog>
                        <AlertDialogTrigger asChild>
                          <DropdownMenuItem
                            onSelect={(e) => e.preventDefault()}
                            className="text-destructive"
                          >
                            Delete
                          </DropdownMenuItem>
                        </AlertDialogTrigger>

                        <AlertDialogContent>
                          <AlertDialogHeader>
                            <AlertDialogTitle>
                              Delete this class?
                            </AlertDialogTitle>

                            <AlertDialogDescription>
                              All students and grades in {c.subject} will be
                              permanently removed.
                            </AlertDialogDescription>
                          </AlertDialogHeader>

                          <AlertDialogFooter>
                            <AlertDialogCancel>Cancel</AlertDialogCancel>

                            <AlertDialogAction
                              onClick={() => deleteClass.mutate(c)}
                              disabled={deleteClass.isPending}
                              className="bg-destructive text-destructive-foreground"
                            >
                              Delete
                            </AlertDialogAction>
                          </AlertDialogFooter>
                        </AlertDialogContent>
                      </AlertDialog>
                    </DropdownMenuContent>
                  </DropdownMenu>
                </div>
              </div>
            );
          })}
        </div>
      )}

      <AssignmentStudentsDialog
        assignment={selectedAssignment}
        declineReason={declineReason}
        onDeclineReasonChange={setDeclineReason}
        onClose={() => {
          setSelectedAssignment(null);
          setDeclineReason("");
        }}
        onAccept={(id) =>
          respondToAssignment.mutate({ assignmentId: id, response: "accepted" })
        }
        onDecline={(id) =>
          respondToAssignment.mutate({
            assignmentId: id,
            response: "declined",
            reason: declineReason,
          })
        }
        isResponding={respondToAssignment.isPending}
      />
    </div>
  );
}

function AssignmentStudentsDialog({
  assignment,
  declineReason,
  onDeclineReasonChange,
  onClose,
  onAccept,
  onDecline,
  isResponding,
}: {
  assignment: StudentAssignment | null;
  declineReason: string;
  onDeclineReasonChange: (value: string) => void;
  onClose: () => void;
  onAccept: (id: string) => void;
  onDecline: (id: string) => void;
  isResponding: boolean;
}) {
  const { data: learners = [], isLoading } = useQuery({
    queryKey: ["student-list-assignment-learners", assignment?.id],
    enabled: Boolean(assignment?.id),
    queryFn: async () => {
      const { data, error } = await (supabase as any).rpc(
        "list_student_list_assignment_learners",
        { p_assignment_id: assignment!.id },
      );
      if (error) throw error;
      return (data ?? []) as AssignedLearner[];
    },
  });

  const maleLearners = learners.filter(
    (learner) => learner.sex?.trim().toLowerCase() === "male",
  );
  const femaleLearners = learners.filter(
    (learner) => learner.sex?.trim().toLowerCase() === "female",
  );
  const learnersWithoutSex = learners.filter((learner) => {
    const sex = learner.sex?.trim().toLowerCase();
    return sex !== "male" && sex !== "female";
  });

  const renderLearnerSection = (
    title: string,
    sectionLearners: AssignedLearner[],
    emptyMessage: string,
  ) => (
    <div>
      <div className="flex items-center justify-between border-b bg-muted/50 px-4 py-2">
        <p className="text-xs font-semibold tracking-wide">{title}</p>
        <span className="text-xs text-muted-foreground">
          {sectionLearners.length}{" "}
          {sectionLearners.length === 1 ? "learner" : "learners"}
        </span>
      </div>
      {sectionLearners.length === 0 ? (
        <p className="px-4 py-3 text-sm text-muted-foreground">
          {emptyMessage}
        </p>
      ) : (
        sectionLearners.map((learner, index) => (
          <div
            key={learner.id}
            className="flex gap-3 border-b px-4 py-3 text-sm last:border-b-0"
          >
            <span className="w-6 text-muted-foreground">{index + 1}.</span>
            <div className="flex-1">
              <p className="font-medium">
                {[
                  learner.last_name,
                  learner.first_name,
                  learner.middle_name,
                  learner.suffix,
                ]
                  .filter(Boolean)
                  .join(", ")}
              </p>
              <p className="text-xs text-muted-foreground">
                LRN: {learner.lrn || "—"}
              </p>
            </div>
          </div>
        ))
      )}
    </div>
  );

  return (
    <Dialog
      open={Boolean(assignment)}
      onOpenChange={(next) => !next && onClose()}
    >
      <DialogContent className="max-h-[90vh] max-w-2xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Received Student Assignment</DialogTitle>
        </DialogHeader>

        {assignment && (
          <div className="space-y-4">
            <div className="grid gap-3 rounded-xl border bg-muted/30 p-4 text-sm sm:grid-cols-2">
              <div>
                <span className="text-muted-foreground">Subject:</span>{" "}
                <strong>{assignment.subject}</strong>
              </div>
              <div>
                <span className="text-muted-foreground">Class:</span>{" "}
                <strong>
                  {assignment.grade_level} · {assignment.section}
                </strong>
              </div>
              <div>
                <span className="text-muted-foreground">School Year:</span>{" "}
                <strong>{assignment.school_year}</strong>
              </div>
              <div>
                <span className="text-muted-foreground">Class Adviser:</span>{" "}
                <strong>{assignment.adviser_name || "—"}</strong>
              </div>
            </div>

            <div>
              <div className="mb-2 flex items-center justify-between">
                <Label>Students ({assignment.learner_count})</Label>
                <span className="text-xs text-muted-foreground">
                  Read-only official list
                </span>
              </div>
              <div className="max-h-64 overflow-y-auto rounded-xl border">
                {isLoading ? (
                  <AssignmentLearnersSkeleton />
                ) : learners.length === 0 ? (
                  <p className="p-4 text-sm text-muted-foreground">
                    No students found.
                  </p>
                ) : (
                  <div>
                    {renderLearnerSection(
                      "MALE",
                      maleLearners,
                      "No male learners.",
                    )}
                    {renderLearnerSection(
                      "FEMALE",
                      femaleLearners,
                      "No female learners.",
                    )}
                    {learnersWithoutSex.length > 0 &&
                      renderLearnerSection(
                        "SEX NOT SET",
                        learnersWithoutSex,
                        "",
                      )}
                  </div>
                )}
              </div>
            </div>

            <div className="rounded-xl border border-primary/20 bg-primary/5 p-3 text-sm">
              After accepting, you can manage only{" "}
              <strong>{assignment.subject}</strong>. Student names, LRN, grade
              level, section, and other subjects remain read-only.
            </div>

            <div>
              <Label htmlFor="decline-reason">Decline reason (optional)</Label>
              <Input
                id="decline-reason"
                value={declineReason}
                onChange={(event) => onDeclineReasonChange(event.target.value)}
                placeholder="Tell the Class Adviser why you declined"
              />
            </div>
          </div>
        )}

        <DialogFooter>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="outline"
            className="text-destructive"
            disabled={!assignment || isResponding}
            onClick={() => assignment && onDecline(assignment.id)}
          >
            <X className="mr-1 size-4" /> Decline
          </Button>
          <Button
            disabled={!assignment || isResponding}
            onClick={() => assignment && onAccept(assignment.id)}
          >
            {isResponding ? (
              <Loader2 className="mr-1 size-4 animate-spin" />
            ) : (
              <Check className="mr-1 size-4" />
            )}
            Accept Students
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function StatusBadge({ status }: { status: StudentAssignmentStatus }) {
  const styles = {
    pending: "bg-amber-100 text-amber-800",
    accepted: "bg-emerald-100 text-emerald-800",
    declined: "bg-red-100 text-red-800",
  }[status];
  const label = {
    pending: "Pending Acceptance",
    accepted: "Accepted",
    declined: "Declined",
  }[status];
  return (
    <span className={`rounded-full px-2 py-1 text-xs font-medium ${styles}`}>
      {label}
    </span>
  );
}

function StatCard({
  label,
  value,
  icon: Icon,
}: {
  label: string;
  value: number | string;
  icon: React.ComponentType<{ className?: string }>;
}) {
  return (
    <div className="group min-h-[132px] overflow-hidden rounded-2xl border border-amber-200/80 bg-gradient-to-r from-[#fffdf8] via-[#fffbf2] to-[#fff8e8] shadow-sm transition-all duration-300 hover:-translate-y-0.5 hover:shadow-md">
      <div className="flex min-h-[132px] flex-col justify-between p-5">
        <div className="flex items-start justify-between gap-4">
          <span className="text-sm font-medium text-[#76533f]">{label}</span>

          <div className="grid size-10 shrink-0 place-items-center rounded-xl border border-amber-200 bg-[#fffaf0] text-[#8f2928] shadow-sm">
            <Icon className="size-5" />
          </div>
        </div>

        <div className="text-3xl font-semibold tracking-tight text-[#321c15]">
          {value}
        </div>
      </div>
    </div>
  );
}
