import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useEffect, useMemo, useState } from "react";
import {
  ArrowLeft,
  ChevronRight,
  GraduationCap,
  FileText,
  BookOpen,
  HeartPulse,
  BadgeCheck,
  Users,
  School,
  StickyNote,
  Send,
  Clock3,
  RotateCcw,
  ShieldCheck,
  TriangleAlert,
  LockKeyhole,
  Pencil,
  Eye,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { Skeleton } from "@/components/ui/skeleton";
import type { ClassRow, StudentRow } from "@/lib/data";
import { toast } from "sonner";

export const Route = createFileRoute("/_authenticated/school-forms")({
  component: FormsPage,
});

type Tone =
  | "rose"
  | "violet"
  | "amber"
  | "sky"
  | "emerald"
  | "indigo"
  | "orange"
  | "teal"
  | "fuchsia"
  | "cyan"
  | "lime"
  | "pink"
  | "slate";

type SubmissionStatus =
  | "submitted"
  | "pending_review"
  | "approved"
  | "returned";

type SubmissionRow = {
  id: string;
  class_id: string;
  form_code: string;
  status: SubmissionStatus;
  admin_remarks: string | null;
  submitted_at: string;
  reviewed_at: string | null;
};

type FormReadinessRow = {
  form_code: string;
  is_ready: boolean;
};

type SchoolLevel = "jhs" | "shs" | "other";

const toneMap: Record<Tone, string> = {
  rose: "text-rose-600 bg-rose-50",
  violet: "text-violet-600 bg-violet-50",
  amber: "text-amber-600 bg-amber-50",
  sky: "text-sky-600 bg-sky-50",
  emerald: "text-emerald-600 bg-emerald-50",
  indigo: "text-indigo-600 bg-indigo-50",
  orange: "text-orange-600 bg-orange-50",
  teal: "text-teal-600 bg-teal-50",
  fuchsia: "text-fuchsia-600 bg-fuchsia-50",
  cyan: "text-cyan-600 bg-cyan-50",
  lime: "text-lime-600 bg-lime-50",
  pink: "text-pink-600 bg-pink-50",
  slate: "text-slate-600 bg-slate-50",
};

const forms: Array<{
  code: string;
  title: string;
  icon: LucideIcon;
  tone: Tone;
  to: string;
}> = [
  {
    code: "GSA",
    title: "General Scholastic Aptitude Report",
    icon: FileText,
    tone: "rose",
    to: "/gsa",
  },
  {
    code: "SF5",
    title: "Report on Promotion & Level of Proficiency",
    icon: GraduationCap,
    tone: "rose",
    to: "/sf5",
  },
  {
    code: "SOG Report",
    title: "Printable Summary of Grades Per Term",
    icon: FileText,
    tone: "violet",
    to: "/sog-report",
  },
  {
    code: "SF1",
    title: "School Register",
    icon: BookOpen,
    tone: "emerald",
    to: "/sf1",
  },
  {
    code: "SF9 (New)",
    title: "Learner's Progress/Performance Report",
    icon: FileText,
    tone: "teal",
    to: "/sf9-matatag",
  },
  {
    code: "SF8",
    title: "Learner Health & Nutrition Profile",
    icon: HeartPulse,
    tone: "cyan",
    to: "/sf8",
  },
  {
    code: "SF10",
    title: "Learner's Permanent Academic Record",
    icon: BookOpen,
    tone: "lime",
    to: "/sf10",
  },
  {
    code: "Anecdotal",
    title: "Anecdotal Record",
    icon: StickyNote,
    tone: "emerald",
    to: "/anecdotal-record",
  },
];

const seniorHighFormCodes = new Set(["SF5", "SF1", "SF9 (New)", "SF10"]);

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";
const SCHOOL_FORMS_NOTIFICATION_CLASS_EVENT =
  "school-forms-notification-class-change";

function readRememberedSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

function rememberSchoolFormsClassId(classId: string) {
  if (typeof window === "undefined" || !classId) return;
  window.sessionStorage.setItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY, classId);
}

function gradeNumberOf(classItem?: Pick<ClassRow, "grade_level"> | null) {
  return Number(String(classItem?.grade_level ?? "").match(/\d+/)?.[0] ?? 0);
}

function schoolLevelOf(classItem?: Pick<ClassRow, "grade_level"> | null): SchoolLevel {
  const gradeNumber = gradeNumberOf(classItem);

  if (gradeNumber >= 7 && gradeNumber <= 10) return "jhs";
  if (gradeNumber === 11 || gradeNumber === 12) return "shs";
  return "other";
}

function schoolLevelLabel(level: SchoolLevel) {
  if (level === "jhs") return "Junior High School";
  if (level === "shs") return "Senior High School";
  return "Other Classes";
}

function SchoolFormsSkeleton() {
  const formCards = Array.from({ length: 8 }, (_, index) => index);

  return (
    <div
      className="space-y-4 pb-24 md:pb-6"
      role="status"
      aria-busy="true"
      aria-label="Loading School Forms"
    >
      <span className="sr-only">Loading School Forms...</span>

      <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border bg-card px-4 py-3 shadow-sm">
        <div className="flex min-w-0 items-center gap-3">
          <Skeleton className="h-9 w-20 shrink-0 rounded-lg" />
          <Skeleton className="size-6 shrink-0 rounded-full" />
          <div className="space-y-2">
            <Skeleton className="h-6 w-36 rounded-md" />
            <Skeleton className="h-3 w-48 max-w-full rounded" />
          </div>
        </div>
        <Skeleton className="h-9 w-52 rounded-md" />
      </div>

      <div className="flex w-fit gap-1 rounded-xl border bg-card p-1">
        <Skeleton className="h-8 w-32 rounded-lg" />
        <Skeleton className="h-8 w-28 rounded-lg" />
      </div>

      <section className="rounded-2xl border bg-card p-4 shadow-sm">
        <div className="flex items-start gap-3">
          <Skeleton className="size-5 shrink-0 rounded-full" />
          <div className="flex-1 space-y-2">
            <Skeleton className="h-4 w-80 max-w-full rounded" />
            <Skeleton className="h-3 w-[32rem] max-w-full rounded" />
          </div>
        </div>
      </section>

      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-6">
        {formCards.map((card) => (
          <div
            key={card}
            className="rounded-2xl border bg-card p-4 shadow-sm"
          >
            <div className="flex items-start justify-between">
              <Skeleton className="size-10 rounded-xl" />
              <Skeleton className="size-4 rounded" />
            </div>

            <Skeleton
              className={`mt-3 h-4 rounded ${
                card % 3 === 0 ? "w-12" : card % 3 === 1 ? "w-16" : "w-20"
              }`}
            />
            <div className="mt-2 space-y-1.5">
              <Skeleton className="h-3 w-full rounded" />
              <Skeleton className="h-3 w-3/4 rounded" />
            </div>

            <div className="mt-4 space-y-2">
              <Skeleton className="h-8 w-full rounded-lg" />
              <Skeleton className="h-8 w-full rounded-lg" />
            </div>
          </div>
        ))}
      </div>

      <div className="flex items-start gap-2 rounded-xl border border-sky-100 bg-sky-50/60 p-3">
        <Skeleton className="size-4 shrink-0 rounded-full" />
        <div className="flex-1 space-y-1.5">
          <Skeleton className="h-3 w-full rounded" />
          <Skeleton className="h-3 w-3/5 rounded" />
        </div>
      </div>
    </div>
  );
}

function FormsPage() {
  const nav = useNavigate();
  const queryClient = useQueryClient();

  const [classId, setClassId] = useState(readRememberedSchoolFormsClassId);
  const [confirmFormCode, setConfirmFormCode] = useState<string | null>(null);

  const { data: session, isLoading: sessionLoading } = useQuery({
    queryKey: ["school-forms-user"],
    queryFn: async () => {
      const { data, error } = await supabase.auth.getUser();
      if (error) throw error;
      return data.user;
    },
  });

  const { data: teacherType, isLoading: teacherTypeLoading } = useQuery({
    enabled: Boolean(session?.id),
    queryKey: ["school-forms-teacher-type", session?.id],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("profiles")
        .select("teacher_type")
        .eq("id", session!.id)
        .maybeSingle();

      if (error) throw error;

      const profile = data as { teacher_type?: string | null } | null;

      return (
        profile?.teacher_type ??
        session?.user_metadata?.teacher_type ??
        "class_adviser"
      );
    },
  });

  const isClassAdviser = teacherType === "class_adviser";

  const { data: classes = [], isLoading: classesLoading } = useQuery({
    enabled: Boolean(session?.id),
    queryKey: ["school-forms-classes", session?.id],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("classes")
        .select("*")
        .eq("teacher_id", session!.id)
        .order("created_at", { ascending: false });

      if (error) throw error;
      return (data ?? []) as ClassRow[];
    },
  });

  const groupedClasses = useMemo(() => {
    const sorted = [...classes].sort((a, b) => {
      const gradeDifference = gradeNumberOf(a) - gradeNumberOf(b);
      if (gradeDifference !== 0) return gradeDifference;

      return String(a.section ?? "").localeCompare(String(b.section ?? ""), undefined, {
        sensitivity: "base",
      });
    });

    return {
      jhs: sorted.filter((item) => schoolLevelOf(item) === "jhs"),
      shs: sorted.filter((item) => schoolLevelOf(item) === "shs"),
      other: sorted.filter((item) => schoolLevelOf(item) === "other"),
    };
  }, [classes]);

  const active = classId || classes[0]?.id || "";
  const activeClass = classes.find((item) => item.id === active);

  const selectSchoolFormsClass = (nextClassId: string) => {
    setClassId(nextClassId);
    rememberSchoolFormsClassId(nextClassId);
  };

  useEffect(() => {
    if (typeof window === "undefined") return;

    const handleNotificationClassChange = (event: Event) => {
      const customEvent = event as CustomEvent<{ classId?: string }>;
      const nextClassId = customEvent.detail?.classId ?? "";

      if (!nextClassId) return;

      setClassId(nextClassId);
      rememberSchoolFormsClassId(nextClassId);
    };

    window.addEventListener(
      SCHOOL_FORMS_NOTIFICATION_CLASS_EVENT,
      handleNotificationClassChange,
    );

    return () => {
      window.removeEventListener(
        SCHOOL_FORMS_NOTIFICATION_CLASS_EVENT,
        handleNotificationClassChange,
      );
    };
  }, []);

  const activeSchoolLevel = schoolLevelOf(activeClass);
  const isSeniorHighActive = activeSchoolLevel === "shs";

  const visibleForms = useMemo(
    () =>
      isSeniorHighActive
        ? forms.filter((form) => seniorHighFormCodes.has(form.code))
        : forms,
    [isSeniorHighActive],
  );

  useEffect(() => {
    // Do not clear an incoming notification class while the teacher's classes
    // are still loading. Once classes are available, validate the selection.
    if (
      classId &&
      !classesLoading &&
      classes.length > 0 &&
      classes.every((item) => item.id !== classId)
    ) {
      setClassId("");
      if (typeof window !== "undefined") {
        window.sessionStorage.removeItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY);
      }
      return;
    }

    if (active) {
      rememberSchoolFormsClassId(active);
    }
  }, [active, classId, classes, classesLoading]);

  const { data: students = [], isLoading: studentsLoading } = useQuery({
    enabled: Boolean(active),
    queryKey: ["school-forms-students", session?.id, active],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("students")
        .select("*")
        .eq("class_id", active);

      if (error) throw error;
      return (data ?? []) as StudentRow[];
    },
  });

  const { data: readiness = [], isLoading: readinessLoading } =
    useQuery<FormReadinessRow[]>({
    enabled: Boolean(active) && isClassAdviser,
    queryKey: ["form-readiness", active],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("class_form_readiness")
        .select("form_code,is_ready")
        .eq("class_id", active);

      if (error) throw error;
      return data ?? [];
    },
  });

  const { data: submissions = [], isLoading: submissionsLoading } =
    useQuery<SubmissionRow[]>({
    enabled: Boolean(active) && isClassAdviser,
    queryKey: ["form-submissions", active],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("school_form_submissions")
        .select(
          "id,class_id,form_code,status,admin_remarks,submitted_at,reviewed_at",
        )
        .eq("class_id", active)
        .order("submitted_at", { ascending: false });

      if (error) throw error;
      return data ?? [];
    },
  });

  const readyCodes = useMemo(
    () =>
      new Set(
        readiness
          .filter((item) => item.is_ready)
          .map((item) => item.form_code),
      ),
    [readiness],
  );

  const latestSubmissionByForm = useMemo(() => {
    const map = new Map<string, SubmissionRow>();

    for (const item of submissions) {
      if (!map.has(item.form_code)) {
        map.set(item.form_code, item);
      }
    }

    return map;
  }, [submissions]);

  const selectedForm = confirmFormCode
    ? forms.find((item) => item.code === confirmFormCode) ?? null
    : null;

  const selectedSubmission = confirmFormCode
    ? latestSubmissionByForm.get(confirmFormCode) ?? null
    : null;

  const toggleReady = useMutation({
    mutationFn: async ({
      formCode,
      ready,
    }: {
      formCode: string;
      ready: boolean;
    }) => {
      const submission = latestSubmissionByForm.get(formCode);

      if (
        submission?.status === "approved" ||
        submission?.status === "submitted" ||
        submission?.status === "pending_review"
      ) {
        throw new Error(
          submission.status === "approved"
            ? "This form is already approved and locked."
            : "This form is already with the Admin for review.",
        );
      }

      const { error } = await (supabase as any).rpc(
        "set_class_form_readiness",
        {
          p_class_id: active,
          p_form_code: formCode,
          p_is_ready: ready,
        },
      );

      if (error) throw error;
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["form-readiness", active],
      });
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const submitSingleForm = useMutation({
    mutationFn: async (formCode: string) => {
      if (!readyCodes.has(formCode)) {
        throw new Error("Mark this form Ready before submitting.");
      }

      const { error } = await (supabase as any).rpc(
        "submit_school_form_to_admin",
        {
          p_class_id: active,
          p_form_code: formCode,
        },
      );

      if (error) throw error;
    },
    onSuccess: async (_, formCode) => {
      const previous = latestSubmissionByForm.get(formCode);

      setConfirmFormCode(null);

      await Promise.all([
        queryClient.invalidateQueries({
          queryKey: ["form-submissions", active],
        }),
        queryClient.invalidateQueries({
          queryKey: ["form-readiness", active],
        }),
      ]);

      toast.success(
        previous?.status === "returned"
          ? `${formCode} resubmitted to Admin.`
          : `${formCode} submitted to Admin.`,
      );
    },
    onError: (error: Error) => toast.error(error.message),
  });

  const schoolFormsLoading =
    sessionLoading ||
    (Boolean(session?.id) && (teacherTypeLoading || classesLoading)) ||
    (Boolean(active) &&
      (studentsLoading ||
        (isClassAdviser && (readinessLoading || submissionsLoading))));

  if (schoolFormsLoading) {
    return <SchoolFormsSkeleton />;
  }

  return (
    <div className="space-y-4 pb-24 md:pb-6">
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border bg-card px-4 py-3">
        <div className="flex items-center gap-3">
          <Link
            to="/dashboard"
            className="inline-flex items-center gap-1 rounded-lg border px-3 py-1.5 text-sm hover:bg-muted"
          >
            <ArrowLeft className="size-4" /> Back
          </Link>

          <div>
            <div className="flex items-center gap-2 text-lg font-semibold">
              <School className="size-5 text-primary" /> School Forms
            </div>
            <div className="text-xs text-muted-foreground">
              {activeClass
                ? `${schoolLevelLabel(activeSchoolLevel)} · ${activeClass.grade_level} · ${
                    activeClass.section || "No section"
                  }`
                : "Pick a class"}{" "}
              · {students.length} students
            </div>
          </div>
        </div>

        <Select value={active} onValueChange={selectSchoolFormsClass}>
          <SelectTrigger className="h-9 w-64">
            <SelectValue placeholder="Pick class" />
          </SelectTrigger>
          <SelectContent>
            {groupedClasses.jhs.length > 0 && (
              <>
                <div className="px-2 pb-1 pt-2 text-[10px] font-bold uppercase tracking-wider text-muted-foreground">
                  Junior High School · Grades 7–10
                </div>
                {groupedClasses.jhs.map((classItem) => (
                  <SelectItem key={classItem.id} value={classItem.id}>
                    {classItem.grade_level} · {classItem.section || "—"}
                  </SelectItem>
                ))}
              </>
            )}

            {groupedClasses.shs.length > 0 && (
              <>
                <div className="mt-1 border-t px-2 pb-1 pt-2 text-[10px] font-bold uppercase tracking-wider text-muted-foreground">
                  Senior High School · Grades 11–12
                </div>
                {groupedClasses.shs.map((classItem) => (
                  <SelectItem key={classItem.id} value={classItem.id}>
                    {classItem.grade_level} · {classItem.section || "—"}
                  </SelectItem>
                ))}
              </>
            )}

            {groupedClasses.other.length > 0 && (
              <>
                <div className="mt-1 border-t px-2 pb-1 pt-2 text-[10px] font-bold uppercase tracking-wider text-muted-foreground">
                  Other Classes
                </div>
                {groupedClasses.other.map((classItem) => (
                  <SelectItem key={classItem.id} value={classItem.id}>
                    {classItem.grade_level || "No grade"} · {classItem.section || "—"}
                  </SelectItem>
                ))}
              </>
            )}
          </SelectContent>
        </Select>
      </div>

      <div className="flex w-fit gap-1 rounded-xl border bg-card p-1">
        <button
          type="button"
          className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-1.5 text-sm font-medium text-primary-foreground"
        >
          <FileText className="size-4" /> School Forms
        </button>

        <button
          type="button"
          onClick={() => nav({ to: "/students" })}
          className="inline-flex items-center gap-2 rounded-lg px-4 py-1.5 text-sm font-medium hover:bg-muted"
        >
          <Users className="size-4" /> My Students
        </button>
      </div>

      {isClassAdviser && activeClass && (
        <section className="rounded-2xl border bg-card p-4 shadow-sm">
          <div className="flex items-start gap-3">
            <ShieldCheck className="mt-0.5 size-5 shrink-0 text-sky-600" />
            <div>
              <div className="font-semibold">
                Forms are reviewed individually by the Admin.
              </div>
              <p className="mt-1 text-sm text-muted-foreground">
                Complete each form, mark it Ready, then submit only that form
                to the Admin for review.
              </p>
            </div>
          </div>
        </section>
      )}

      {activeClass && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border bg-card px-4 py-2.5 text-sm">
          <div>
            <span className="font-semibold">{schoolLevelLabel(activeSchoolLevel)}</span>
            <span className="text-muted-foreground">
              {" "}
              · {activeClass.grade_level} · {activeClass.section || "No section"}
            </span>
          </div>
          <span className="rounded-full bg-muted px-2.5 py-1 text-xs font-semibold text-muted-foreground">
            {isSeniorHighActive ? "4 SHS forms" : `${visibleForms.length} JHS forms`}
          </span>
        </div>
      )}

      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-6">
        {visibleForms.map((form) => {
          const isReady = readyCodes.has(form.code);
          const submission = latestSubmissionByForm.get(form.code);

          const isPending =
            submission?.status === "submitted" ||
            submission?.status === "pending_review";

          const isApproved = submission?.status === "approved";
          const isReturned = submission?.status === "returned";

          const canEditForm = !isApproved && !isPending;
          const canToggleReady = canEditForm;
          const canSubmit =
            isClassAdviser &&
            activeClass &&
            isReady &&
            !isApproved &&
            !isPending;

          return (
            <div
              key={form.code}
              className="group relative rounded-2xl border bg-card p-4 transition hover:border-primary/50 hover:shadow-md"
            >
              {canEditForm ? (
                <Link
                  to={form.to}
                  onClick={() => rememberSchoolFormsClassId(active)}
                  className="block"
                >
                  <div className="flex items-start justify-between">
                    <div
                      className={`grid size-10 place-items-center rounded-xl ${
                        toneMap[form.tone]
                      }`}
                    >
                      <form.icon className="size-5" />
                    </div>

                    <ChevronRight
                      className={`size-4 opacity-0 transition group-hover:opacity-100 ${
                        toneMap[form.tone].split(" ")[0]
                      }`}
                    />
                  </div>

                  <div
                    className={`mt-3 text-sm font-semibold ${
                      toneMap[form.tone].split(" ")[0]
                    }`}
                  >
                    {form.code}
                  </div>

                  <div className="min-h-8 text-xs leading-snug text-muted-foreground">
                    {form.title}
                  </div>
                </Link>
              ) : (
                <div className="block">
                  <div className="flex items-start justify-between">
                    <div
                      className={`grid size-10 place-items-center rounded-xl ${
                        toneMap[form.tone]
                      }`}
                    >
                      <form.icon className="size-5" />
                    </div>

                    <LockKeyhole className="size-4 text-muted-foreground" />
                  </div>

                  <div
                    className={`mt-3 text-sm font-semibold ${
                      toneMap[form.tone].split(" ")[0]
                    }`}
                  >
                    {form.code}
                  </div>

                  <div className="min-h-8 text-xs leading-snug text-muted-foreground">
                    {form.title}
                  </div>
                </div>
              )}

              {isClassAdviser && activeClass && (
                <div className="mt-3 space-y-2">
                  {isApproved ? (
                    <>
                      <div className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2 py-1.5 text-xs font-semibold text-emerald-700">
                        <ShieldCheck className="size-3.5" />
                        Approved
                        <LockKeyhole className="size-3.5" />
                      </div>

                      <button
                        type="button"
                        disabled
                        className="inline-flex w-full cursor-not-allowed items-center justify-center gap-1.5 rounded-lg border px-2 py-1.5 text-xs font-medium opacity-45"
                      >
                        <Eye className="size-3.5" />
                        View (Read Only)
                      </button>
                    </>
                  ) : isPending ? (
                    <>
                      <div className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg border border-amber-200 bg-amber-50 px-2 py-1.5 text-xs font-semibold text-amber-800">
                        <Clock3 className="size-3.5" />
                        Pending Admin Review
                      </div>

                      <button
                        type="button"
                        disabled
                        className="inline-flex w-full cursor-not-allowed items-center justify-center gap-1.5 rounded-lg border px-2 py-1.5 text-xs font-medium opacity-45"
                      >
                        <Send className="size-3.5" />
                        Submitted
                      </button>
                    </>
                  ) : isReturned ? (
                    <>
                      <div className="rounded-lg border border-sky-200 bg-sky-50 p-2 text-xs text-sky-900">
                        <div className="flex items-center gap-1.5 font-semibold">
                          <RotateCcw className="size-3.5" />
                          Returned for Correction
                        </div>

                        {submission.admin_remarks && (
                          <p className="mt-1.5 leading-snug">
                            <b>Admin note:</b> {submission.admin_remarks}
                          </p>
                        )}
                      </div>

                      <Link
                        to={form.to}
                        onClick={() => rememberSchoolFormsClassId(active)}
                        className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg border px-2 py-1.5 text-xs font-medium hover:bg-muted"
                      >
                        <Pencil className="size-3.5" />
                        Edit Form
                      </Link>

                      <button
                        type="button"
                        disabled={toggleReady.isPending}
                        onClick={() =>
                          toggleReady.mutate({
                            formCode: form.code,
                            ready: !isReady,
                          })
                        }
                        className={`inline-flex w-full items-center justify-center gap-1.5 rounded-lg border px-2 py-1.5 text-xs font-medium disabled:opacity-50 ${
                          isReady
                            ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                            : "hover:bg-muted"
                        }`}
                      >
                        <BadgeCheck className="size-3.5" />
                        {isReady ? "Ready" : "Mark Ready"}
                      </button>

                      <button
                        type="button"
                        disabled={!canSubmit || submitSingleForm.isPending}
                        onClick={() => setConfirmFormCode(form.code)}
                        className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg bg-primary px-2 py-1.5 text-xs font-semibold text-primary-foreground disabled:cursor-not-allowed disabled:opacity-45"
                      >
                        <RotateCcw className="size-3.5" />
                        Resubmit to Admin
                      </button>
                    </>
                  ) : (
                    <>
                      <button
                        type="button"
                        disabled={!canToggleReady || toggleReady.isPending}
                        onClick={() =>
                          toggleReady.mutate({
                            formCode: form.code,
                            ready: !isReady,
                          })
                        }
                        className={`inline-flex w-full items-center justify-center gap-1.5 rounded-lg border px-2 py-1.5 text-xs font-medium disabled:opacity-50 ${
                          isReady
                            ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                            : "hover:bg-muted"
                        }`}
                      >
                        <BadgeCheck className="size-3.5" />
                        {isReady ? "Ready" : "Mark Ready"}
                      </button>

                      <button
                        type="button"
                        disabled={!canSubmit || submitSingleForm.isPending}
                        onClick={() => setConfirmFormCode(form.code)}
                        className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg bg-primary px-2 py-1.5 text-xs font-semibold text-primary-foreground disabled:cursor-not-allowed disabled:opacity-45"
                      >
                        <Send className="size-3.5" />
                        {isReady ? "Submit to Admin" : "Ready first"}
                      </button>
                    </>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {isClassAdviser && activeClass && (
        <div className="flex items-start gap-2 rounded-xl border border-sky-200 bg-sky-50 p-3 text-xs text-sky-900">
          <LockKeyhole className="mt-0.5 size-4 shrink-0" />
          <span>
            Approved forms are locked on this page. To make the individual
            SF1/SF5/SF8/SF9/SF10/GSA/Anecdotal form pages truly read-only too,
            those form routes must also check the approved status.
          </span>
        </div>
      )}

      <AlertDialog
        open={Boolean(confirmFormCode)}
        onOpenChange={(open) => {
          if (!open) setConfirmFormCode(null);
        }}
      >
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>
              {selectedSubmission?.status === "returned"
                ? `Resubmit ${selectedForm?.code ?? "form"}?`
                : `Submit ${selectedForm?.code ?? "form"} to Admin?`}
            </AlertDialogTitle>

            <AlertDialogDescription>
              Only this form will be sent for Admin review for{" "}
              {activeClass?.grade_level},{" "}
              {activeClass?.section || "No section"}, school year{" "}
              {activeClass?.school_year || "—"}. Other forms remain editable
              and can be submitted separately.
            </AlertDialogDescription>
          </AlertDialogHeader>

          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>

            <AlertDialogAction
              disabled={!confirmFormCode || submitSingleForm.isPending}
              onClick={(event) => {
                event.preventDefault();

                if (confirmFormCode) {
                  submitSingleForm.mutate(confirmFormCode);
                }
              }}
            >
              {submitSingleForm.isPending
                ? "Submitting…"
                : selectedSubmission?.status === "returned"
                  ? "Confirm Resubmission"
                  : "Confirm Submission"}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
