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
  Dumbbell,
  FileSearch,
  FlaskConical,
  Globe2,
  HeartHandshake,
  HeartPulse,
  Landmark,
  Languages,
  MoreHorizontal,
  Music2,
  Palette,
  Plus,
  Target,
  Users,
  type LucideIcon,
} from "lucide-react";
import { useMemo, useState } from "react";
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

export const Route = createFileRoute("/_authenticated/dashboard")({
  component: Dashboard,
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
    "Elective Subject",
  ],
  "Grade 12": [
    // Core Subjects — based on the Grade 12 SF9 template
    "Media and Information Literacy",
    "PE and Health 3",
    "Introduction to Human Philosophy",
    "Disaster Readiness and Risk Reduction",
    "Contemporary Philippine Arts",
    "PE and Health 4",

    // Applied and Specialized Subjects — based on the Grade 12 SF9 template
    "Filipino sa Piling Larang",
    "Practical Research 2",
    "General Biology 1",
    "General Physics 1",
    "English for Academic & Professional Purposes",
    "Entrepreneurship",
    "General Physics 2",
    "General Biology 2",
    "Inquiries, Investigation and Immersion",
    "Capstone Project",
  ],
};

const UNITS_BY_GRADE: Record<string, readonly string[]> = {
  // Grade 11 keeps the required Units selector.
  // Grade 12 intentionally has no Units selector. Its Subject is entered
  // manually and SF9 leaves the Units cells blank.
  "Grade 11": ["2", "3", "6"],
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

function DashboardSkeleton() {
  return (
    <div
      className="space-y-6 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading classes dashboard...</span>

      <div className="space-y-2">
        <Skeleton className="h-8 w-40" />
        <Skeleton className="h-4 w-full max-w-sm" />
      </div>

      <div className="grid gap-3 sm:grid-cols-2">
        {Array.from({ length: 2 }, (_, index) => (
          <div key={index} className="rounded-2xl border bg-card p-4">
            <div className="flex items-center justify-between">
              <Skeleton className="h-4 w-20" />
              <Skeleton className="size-8 rounded-lg" />
            </div>
            <Skeleton className="mt-3 h-8 w-16" />
          </div>
        ))}
      </div>

      <div className="flex items-center justify-between gap-3">
        <Skeleton className="h-6 w-24" />
        <Skeleton className="h-9 w-28 rounded-lg" />
      </div>

      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
        {Array.from({ length: 6 }, (_, index) => (
          <div key={index} className="rounded-2xl border bg-card p-4">
            <div className="flex items-start justify-between gap-3">
              <div className="flex-1">
                <Skeleton className="size-10 rounded-lg" />
                <Skeleton
                  className={`mt-3 h-5 ${index % 3 === 0 ? "w-40" : "w-28"}`}
                />
                <Skeleton className="mt-2 h-4 w-full max-w-56" />
              </div>
              <Skeleton className="size-9 rounded-lg" />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Dashboard() {
  const qc = useQueryClient();
  const { data: classes = [], isLoading: isClassesLoading } = useQuery({
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
  const { data: studentsCount = 0, isLoading: isStudentsCountLoading } =
    useQuery({
      queryKey: ["students-count"],
      queryFn: async () => {
        const { count } = await supabase
          .from("students")
          .select("*", { count: "exact", head: true });
        return count ?? 0;
      },
    });

  const { data: schoolYearSetting, isLoading: isSchoolYearLoading } =
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

  const { data: schoolYearLibrary = [], isLoading: isSchoolYearLibraryLoading } =
    useQuery<SchoolYearLibraryRow[]>({
      queryKey: ["school-year-library"],
      queryFn: getSchoolYearLibrary,
      refetchInterval: 15000,
      refetchOnWindowFocus: true,
    });

  const [open, setOpen] = useState(false);
  const [form, setForm] = useState(emptyClassForm);

  // Sort all class cards alphabetically by subject for both
  // Class Adviser and Subject Teacher dashboards.
  // If subjects are the same, sort by grade level, section,
  // and school year for a stable and predictable order.
  const sortedClasses = useMemo(() => {
    return [...classes].sort((a, b) => {
      const subjectCompare = String(a.subject ?? "").localeCompare(
        String(b.subject ?? ""),
        undefined,
        {
          sensitivity: "base",
          numeric: true,
        },
      );

      if (subjectCompare !== 0) return subjectCompare;

      const gradeCompare = String(a.grade_level ?? "").localeCompare(
        String(b.grade_level ?? ""),
        undefined,
        {
          sensitivity: "base",
          numeric: true,
        },
      );

      if (gradeCompare !== 0) return gradeCompare;

      const sectionCompare = String(a.section ?? "").localeCompare(
        String(b.section ?? ""),
        undefined,
        {
          sensitivity: "base",
          numeric: true,
        },
      );

      if (sectionCompare !== 0) return sectionCompare;

      return String(a.school_year ?? "").localeCompare(
        String(b.school_year ?? ""),
        undefined,
        {
          sensitivity: "base",
          numeric: true,
        },
      );
    });
  }, [classes]);

  const createClass = useMutation({
    mutationFn: async () => {
      const teacher_id = await getUserId();
      const allowedSchoolYear = await resolveAllowedSchoolYear(
        activeSchoolYear ?? form.school_year,
      );
      // Grade 12 classes do not use Units. Even if an old form value is
      // still present in state, always save Grade 12 units as null.
      const units =
        form.grade_level === "Grade 12"
          ? null
          : form.units && form.units !== "none"
            ? Number(form.units)
            : null;
      const { error } = await supabase.from("classes").insert({
        subject: form.subject.trim(),
        grade_level: form.grade_level,
        section: form.section.trim(),
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
    mutationFn: async (id: string) => {
      const { error } = await supabase.from("classes").delete().eq("id", id);
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success("Deleted");
      qc.invalidateQueries({ queryKey: ["classes"] });
    },
  });

  if (
    isClassesLoading ||
    isStudentsCountLoading ||
    isSchoolYearLoading ||
    isSchoolYearLibraryLoading
  ) {
    return <DashboardSkeleton />;
  }

  return (
    <div className="space-y-6 pb-24 md:pb-6">
      <div>
        <h1 className="text-2xl font-semibold">My Classes</h1>
        <p className="text-sm text-muted-foreground">
          Everything you enter here flows into every DepEd form.
        </p>
      </div>

      <div className="grid gap-3 sm:grid-cols-2">
        <StatCard
          label="Classes"
          value={classes.length}
          icon={BookOpen}
          category="classes"
        />
        <StatCard
          label="Students"
          value={studentsCount}
          icon={Users}
          category="students"
        />
      </div>

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
                          {unit === "none"
                            ? "None"
                            : `${unit} ${unit === "1" ? "Unit" : "Units"}`}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Subject</Label>

                  {form.grade_level === "Grade 12" ? (
                    <Input
                      value={form.subject}
                      onChange={(e) =>
                        setForm((current) => ({
                          ...current,
                          subject: e.target.value,
                        }))
                      }
                      placeholder="Enter subject"
                    />
                  ) : (
                    <Select
                      value={form.subject || undefined}
                      onValueChange={(v) =>
                        setForm((current) => ({
                          ...current,
                          subject: v,
                        }))
                      }
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
                      <SelectContent
                        position="popper"
                        side="bottom"
                        align="start"
                        sideOffset={4}
                        avoidCollisions={false}
                        className="max-h-72 overflow-y-auto"
                      >
                        {(SUBJECTS_BY_GRADE[form.grade_level] ?? []).map(
                          (sub) => (
                            <SelectItem key={sub} value={sub}>
                              {sub}
                            </SelectItem>
                          ),
                        )}
                      </SelectContent>
                    </Select>
                  )}
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
                  !form.subject.trim() ||
                  !form.section.trim() ||
                  !(activeSchoolYear || form.school_year) ||
                  (form.grade_level === "Grade 11" && !form.units) ||
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
          {sortedClasses.map((c) => {
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
                        {c.grade_level !== "Grade 12" && c.units != null
                          ? ` · ${c.units} Units`
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
                              onClick={() => deleteClass.mutate(c.id)}
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
    </div>
  );
}

function StatCard({
  label,
  value,
  icon: Icon,
  category,
}: {
  label: string;
  value: number | string;
  icon: React.ComponentType<{ className?: string }>;
  category: "classes" | "students";
}) {
  const theme =
    category === "classes"
      ? {
          border: "border-amber-200/80",
          surface: "bg-gradient-to-r from-[#fffdf8] via-[#fffbf2] to-[#fff8e8]",
          icon: "border-amber-200 bg-[#fffaf0] text-[#8f2928]",
          label: "text-[#76533f]",
          value: "text-[#321c15]",
        }
      : {
          border: "border-amber-200/80",
          surface: "bg-gradient-to-r from-[#fffdf8] via-[#fffbf2] to-[#fff8e8]",
          icon: "border-amber-200 bg-[#fffaf0] text-[#8f2928]",
          label: "text-[#76533f]",
          value: "text-[#321c15]",
        };

  return (
    <div
      className={`group min-h-[132px] overflow-hidden rounded-2xl border shadow-sm transition-all duration-300 hover:-translate-y-0.5 hover:shadow-md ${theme.border} ${theme.surface}`}
    >
      <div className="flex min-h-[132px] flex-col justify-between p-5">
        <div className="flex items-start justify-between gap-4">
          <span className={`text-sm font-medium ${theme.label}`}>{label}</span>
          <div
            className={`grid size-10 shrink-0 place-items-center rounded-xl border shadow-sm ${theme.icon}`}
          >
            <Icon className="size-5" />
          </div>
        </div>
        <div className={`text-3xl font-semibold tracking-tight ${theme.value}`}>
          {value}
        </div>
      </div>
    </div>
  );
}