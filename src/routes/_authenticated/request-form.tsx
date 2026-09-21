import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Fragment, useMemo, useState } from "react";
import {
  ArrowLeft,
  BookOpen,
  CalendarDays,
  CheckCircle2,
  Clock3,
  Download,
  Eye,
  Flag,
  GraduationCap,
  Inbox,
  LoaderCircle,
  Mail,
  MessageSquareText,
  RefreshCw,
  Send,
  FileCheck2,
  Trash2,
  Trophy,
  UserRound,
  Users,
  XCircle,
} from "lucide-react";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
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
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

export const Route = createFileRoute("/_authenticated/request-form")({
  component: RequestFormPage,
});

type RequestStatus = "pending" | "in_progress" | "completed" | "rejected";
type Priority = "low" | "medium" | "high";
type DateFilterMode = "all" | "day" | "month" | "year";

type RequestRow = {
  id: string;
  requester_id: string;
  subject_teacher_id: string | null;
  subject: string;
  term: string;
  message: string;
  priority: Priority;
  status: RequestStatus;
  attachment_url: string | null;
  attachment_name: string | null;
  teacher_response: string | null;
  created_at: string;
  updated_at: string;
};

type ClassAdviserProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
};

type StudentSexRow = {
  first_name: string | null;
  last_name: string | null;
  sex: string | null;
};

type LearnerSex = "male" | "female" | "unknown";

type ParsedLearner = {
  number: number;
  name: string;
  sex: LearnerSex;
  grade: number | null;
  gradeText: string;
  rank: number | null;
  descriptor: string;
  award: {
    label: string;
    tone: string;
  } | null;
};

type ParsedRequestMessage = {
  heading: string;
  classDetails: string;
  learners: ParsedLearner[];
};

// Added Class Adviser -> Subject Teacher grade-request workflow types.
type GradingPeriod = "1" | "2" | "3" | "final";
type WorkflowRecipientRole = "class_adviser" | "subject_teacher";
type WorkflowOverallStatus = "pending" | "completed";
type WorkflowSubjectStatus = "pending" | "completed" | "rejected";
type WorkflowStatusFilter = "all" | WorkflowSubjectStatus;
type PreviousWorkflowStatusFilter =
  | "all"
  | "pending"
  | "completed"
  | "rejected"
  | "finalized";

type WorkflowContext = {
  userId: string;
  teacherType: string | null;
  fullName: string | null;
  email: string | null;
};

type WorkflowRecipientOption = {
  id: string;
  fullName: string;
  email: string;
  teacherType: "class_adviser" | "subject_teacher";
  assignedToClass: boolean;
  subjects: string[];
};

type AdvisoryClassRow = {
  id: string;
  teacher_id: string | null;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
};

type GradeWorkflowRow = {
  batch_id: string;
  adviser_id: string;
  adviser_name: string | null;
  adviser_email: string | null;
  advisory_class_id: string;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  grading_period: GradingPeriod;
  overall_status: WorkflowOverallStatus;
  is_finalized: boolean;
  requested_at: string;
  completed_at: string | null;
  finalized_at: string | null;
  subject_request_id: string;
  subject_class_id: string | null;
  subject_teacher_id: string;
  subject_teacher_name: string | null;
  subject_teacher_email: string | null;
  subject: string;
  subject_status: WorkflowSubjectStatus;
  subject_requested_at: string;
  submitted_at: string | null;
  teacher_note: string | null;
  submitted_grade_count: number;
  request_message: string | null;
  recipient_role?: WorkflowRecipientRole | null;
  recipient_id?: string | null;
};

type GradeWorkflowBatch = {
  batchId: string;
  adviserName: string;
  advisoryClassId: string;
  gradeLevel: string;
  section: string;
  schoolYear: string;
  gradingPeriod: GradingPeriod;
  overallStatus: WorkflowOverallStatus;
  isFinalized: boolean;
  requestedAt: string;
  completedAt: string | null;
  finalizedAt: string | null;
  subjects: GradeWorkflowRow[];
};

type SentWorkflowDeleteTarget = {
  batch: GradeWorkflowBatch;
  subject: GradeWorkflowRow;
};

type WorkflowReviewStudent = {
  id: string;
  first_name: string;
  middle_name: string | null;
  last_name: string;
  lrn: string | null;
  sex: LearnerSex;
};

type WorkflowReviewScore = {
  student_id: string;
  score: number;
};

const statusLabel: Record<RequestStatus, string> = {
  pending: "Pending",
  in_progress: "In Progress",
  completed: "Completed",
  rejected: "Rejected",
};

const statusStyle: Record<RequestStatus, string> = {
  pending: "border-amber-200 bg-amber-50 text-amber-700",
  in_progress: "border-blue-200 bg-blue-50 text-blue-700",
  completed: "border-emerald-200 bg-emerald-50 text-emerald-700",
  rejected: "border-red-200 bg-red-50 text-red-700",
};

const priorityStyle: Record<Priority, string> = {
  low: "border-emerald-200 bg-emerald-50 text-emerald-700",
  medium: "border-amber-200 bg-amber-50 text-amber-700",
  high: "border-red-200 bg-red-50 text-red-700",
};

const workflowOverallStyle: Record<WorkflowOverallStatus, string> = {
  pending: "border-amber-200 bg-amber-50 text-amber-700",
  completed: "border-emerald-200 bg-emerald-50 text-emerald-700",
};

const workflowSubjectStyle: Record<WorkflowSubjectStatus, string> = {
  pending: "border-amber-200 bg-amber-50 text-amber-700",
  completed: "border-emerald-200 bg-emerald-50 text-emerald-700",
  rejected: "border-red-200 bg-red-50 text-red-700",
};

function StatusIcon({ status }: { status: RequestStatus }) {
  if (status === "pending") {
    return <Clock3 className="size-3.5" />;
  }

  if (status === "in_progress") {
    return <LoaderCircle className="size-3.5 animate-spin" />;
  }

  if (status === "completed") {
    return <CheckCircle2 className="size-3.5" />;
  }

  return <XCircle className="size-3.5" />;
}

function RequestFormSkeleton({ showWorkflow }: { showWorkflow: boolean }) {
  const workflowRows = Array.from({ length: 3 }, (_, index) => index);
  const historyRows = Array.from({ length: 5 }, (_, index) => index);

  return (
    <div
      className="space-y-5 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading Request Form...</span>

      <section className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card px-5 py-4 shadow-sm">
        <div className="flex items-center gap-4">
          <Skeleton className="h-10 w-20 rounded-xl" />
          <div className="space-y-2">
            <div className="flex items-center gap-2">
              <Skeleton className="size-6 rounded-lg" />
              <Skeleton className="h-6 w-40" />
            </div>
            <Skeleton className="h-4 w-64 max-w-[60vw]" />
          </div>
        </div>
        <Skeleton className="h-10 w-28 rounded-xl" />
      </section>

      {showWorkflow && (
        <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
          <div className="flex flex-wrap items-start justify-between gap-3 border-b px-5 py-4">
            <div className="space-y-2">
              <div className="flex items-center gap-2">
                <Skeleton className="size-5 rounded-md" />
                <Skeleton className="h-5 w-72 max-w-[65vw]" />
              </div>
              <Skeleton className="h-4 w-[34rem] max-w-[75vw]" />
            </div>
            <Skeleton className="h-10 w-40 rounded-xl" />
          </div>

          <div className="border-b bg-muted/20 p-5">
            <div className="grid gap-3 lg:grid-cols-[minmax(220px,0.65fr)_minmax(320px,1fr)_minmax(180px,0.55fr)]">
              {Array.from({ length: 3 }, (_, index) => (
                <div key={index} className="space-y-2">
                  <Skeleton className="h-3 w-28" />
                  <Skeleton className="h-10 w-full rounded-md" />
                  {index < 2 && <Skeleton className="h-3 w-4/5" />}
                </div>
              ))}
            </div>

            <div className="mt-4 grid gap-3 lg:grid-cols-[minmax(300px,0.95fr)_minmax(0,1.4fr)_auto] lg:items-end">
              <div className="space-y-2">
                <Skeleton className="h-3 w-24" />
                <Skeleton className="h-10 w-full rounded-md" />
                <Skeleton className="h-3 w-3/4" />
              </div>
              <div className="space-y-2">
                <div className="flex justify-between gap-4">
                  <Skeleton className="h-3 w-28" />
                  <Skeleton className="h-3 w-14" />
                </div>
                <Skeleton className="h-20 w-full rounded-xl" />
              </div>
              <Skeleton className="h-10 w-48 rounded-md" />
            </div>
          </div>

          <div className="grid gap-3 border-b bg-background px-5 py-4 sm:grid-cols-2 lg:grid-cols-4">
            {Array.from({ length: 3 }, (_, index) => (
              <div key={index} className="rounded-xl border p-3">
                <Skeleton className="h-3 w-24" />
                <Skeleton className="mt-2 h-6 w-10" />
              </div>
            ))}
          </div>

          <div className="space-y-4 p-5">
            <div className="overflow-hidden rounded-2xl border bg-background">
              <div className="flex flex-wrap items-start justify-between gap-3 border-b px-4 py-4">
                <div className="space-y-2">
                  <div className="flex items-center gap-2">
                    <Skeleton className="h-5 w-36" />
                    <Skeleton className="h-7 w-16 rounded-lg" />
                    <Skeleton className="h-7 w-20 rounded-lg" />
                  </div>
                  <Skeleton className="h-4 w-56" />
                </div>
                <Skeleton className="h-9 w-32 rounded-lg" />
              </div>
              <div className="overflow-x-auto">
                <div className="min-w-[760px]">
                  <div className="grid grid-cols-[1fr_1.2fr_0.7fr_0.8fr_0.6fr] gap-4 bg-muted/40 px-4 py-3">
                    {Array.from({ length: 5 }, (_, index) => (
                      <Skeleton key={index} className="h-3 w-20" />
                    ))}
                  </div>
                  {workflowRows.map((row) => (
                    <div
                      key={row}
                      className="grid grid-cols-[1fr_1.2fr_0.7fr_0.8fr_0.6fr] items-center gap-4 border-t px-4 py-3"
                    >
                      <Skeleton className="h-4 w-28" />
                      <Skeleton className="h-4 w-40" />
                      <Skeleton className="h-4 w-14" />
                      <Skeleton className="h-7 w-20 rounded-lg" />
                      <Skeleton className="ml-auto size-8 rounded-lg" />
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
        {Array.from({ length: 5 }, (_, index) => (
          <div key={index} className="rounded-xl border bg-card p-4 shadow-sm">
            <Skeleton className="h-3 w-24" />
            <Skeleton className="mt-3 h-7 w-10" />
          </div>
        ))}
      </section>

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div className="space-y-2">
            <Skeleton className="h-5 w-40" />
            <Skeleton className="h-4 w-28" />
          </div>
          <div className="flex flex-wrap gap-2">
            <Skeleton className="h-10 w-44 rounded-md" />
            <Skeleton className="h-10 w-44 rounded-md" />
          </div>
        </div>

        <div className="overflow-x-auto">
          <div className="min-w-[1100px]">
            <div className="grid grid-cols-[1.1fr_0.7fr_0.5fr_0.6fr_1.4fr_0.7fr_0.8fr_0.5fr] gap-4 bg-muted/40 px-4 py-3">
              {Array.from({ length: 8 }, (_, index) => (
                <Skeleton key={index} className="h-3 w-20" />
              ))}
            </div>
            {historyRows.map((row) => (
              <div
                key={row}
                className="grid grid-cols-[1.1fr_0.7fr_0.5fr_0.6fr_1.4fr_0.7fr_0.8fr_0.5fr] items-center gap-4 border-t px-4 py-3"
              >
                <div className="flex items-center gap-2">
                  <Skeleton className="size-8 rounded-lg" />
                  <Skeleton className="h-4 w-28" />
                </div>
                <Skeleton className="h-4 w-20" />
                <Skeleton className="h-4 w-12" />
                <Skeleton className="h-7 w-16 rounded-lg" />
                <div className="space-y-1.5">
                  <Skeleton className="h-3 w-full" />
                  <Skeleton className="h-3 w-3/4" />
                </div>
                <Skeleton className="h-7 w-20 rounded-lg" />
                <Skeleton className="h-4 w-24" />
                <Skeleton className="ml-auto h-8 w-16 rounded-lg" />
              </div>
            ))}
          </div>
        </div>
      </section>
    </div>
  );
}

function RequestFormPage() {
  const queryClient = useQueryClient();
  const navigate = useNavigate();
  const today = new Date();

  const [statusFilter, setStatusFilter] = useState<"all" | RequestStatus>(
    "all",
  );
  const [dateFilterMode, setDateFilterMode] = useState<DateFilterMode>("all");
  const [selectedDay, setSelectedDay] = useState(() => toDateInputValue(today));
  const [selectedMonth, setSelectedMonth] = useState(() =>
    toMonthInputValue(today),
  );
  const [selectedYear, setSelectedYear] = useState(() =>
    String(today.getFullYear()),
  );
  const [selectedRequest, setSelectedRequest] = useState<RequestRow | null>(
    null,
  );
  const [deleteTarget, setDeleteTarget] = useState<RequestRow | null>(null);
  const [sentWorkflowDeleteTarget, setSentWorkflowDeleteTarget] =
    useState<SentWorkflowDeleteTarget | null>(null);
  const [sentWorkflowReviewTarget, setSentWorkflowReviewTarget] =
    useState<SentWorkflowDeleteTarget | null>(null);
  const [recipientMessageDialog, setRecipientMessageDialog] = useState<{
    recipientName: string;
    subject: string;
    classLabel: string;
    status: string;
    message: string;
  } | null>(null);

  // Added workflow controls. Existing request history state remains unchanged.
  const [workflowClassId, setWorkflowClassId] = useState("");
  const [workflowPeriod, setWorkflowPeriod] = useState<GradingPeriod>("1");
  const [workflowRecipientRole, setWorkflowRecipientRole] =
    useState<WorkflowRecipientRole>("subject_teacher");
  const [workflowDirectoryRecipientId, setWorkflowDirectoryRecipientId] =
    useState("");
  const [workflowMessage, setWorkflowMessage] = useState("");
  const [workflowStatusFilter, setWorkflowStatusFilter] =
    useState<WorkflowStatusFilter>("all");
  const [previousWorkflowStatusFilter, setPreviousWorkflowStatusFilter] =
    useState<PreviousWorkflowStatusFilter>("all");

  const {
    data: requestsData,
    isLoading,
    isPlaceholderData: requestsPlaceholder,
    error,
    refetch,
    isFetching,
  } = useQuery<RequestRow[]>({
    queryKey: ["my-grade-requests"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      const { data, error: requestError } = await (supabase as any)
        .from("grade_requests")
        .select("*")
        .eq("requester_id", user.id)
        .order("created_at", { ascending: false });

      if (requestError) throw requestError;

      return Array.isArray(data) ? (data as unknown as RequestRow[]) : [];
    },
    placeholderData: [],
  });

  const requests = Array.isArray(requestsData) ? requestsData : [];

  // ADDITIVE WORKFLOW: determine whether the current user is a Class Adviser.
  const { data: workflowContext, isLoading: workflowContextLoading } =
    useQuery<WorkflowContext>({
    queryKey: ["request-form-workflow-context"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      const { data, error: profileError } = await (supabase as any)
        .from("profiles")
        .select("teacher_type, full_name, email")
        .eq("id", user.id)
        .maybeSingle();

      if (profileError) throw profileError;

      return {
        userId: user.id,
        teacherType: data?.teacher_type ?? null,
        fullName: data?.full_name ?? null,
        email: data?.email ?? user.email ?? null,
      };
    },
  });

  const isClassAdviser = workflowContext?.teacherType === "class_adviser";

  // Keep the logged-in Class Adviser's own classes. These are used as the
  // authorized advisory-class records when a grade request is submitted.
  const {
    data: ownedAdvisoryClassesData,
    isLoading: ownedClassesLoading,
    isPlaceholderData: ownedClassesPlaceholder,
  } = useQuery<AdvisoryClassRow[]>({
    queryKey: ["request-form-owned-advisory-classes", workflowContext?.userId],
    enabled: Boolean(isClassAdviser && workflowContext?.userId),
    queryFn: async () => {
      const { data, error: classError } = await (supabase as any)
        .from("classes")
        .select("id, teacher_id, subject, grade_level, section, school_year")
        .eq("teacher_id", workflowContext!.userId)
        .order("created_at", { ascending: false });

      if (classError) throw classError;
      return Array.isArray(data) ? (data as unknown as AdvisoryClassRow[]) : [];
    },
    placeholderData: [],
  });

  const ownedAdvisoryClasses = useMemo(
    () => dedupeAdvisoryClasses(ownedAdvisoryClassesData ?? []),
    [ownedAdvisoryClassesData],
  );

  // The recipient directory RPC requires one class owned by the logged-in
  // Class Adviser. This class is only used to load the email directory.
  const directorySeedClassId = ownedAdvisoryClasses[0]?.id || "";

  const {
    data: workflowRecipientOptionsData,
    isLoading: workflowRecipientsLoading,
    isPlaceholderData: workflowRecipientsPlaceholder,
    error: workflowRecipientsError,
  } = useQuery<WorkflowRecipientOption[]>({
    queryKey: [
      "request-form-workflow-recipient-directory",
      directorySeedClassId,
      workflowContext?.userId,
      workflowContext?.email?.trim().toLowerCase(),
    ],
    enabled: Boolean(isClassAdviser && directorySeedClassId),
    queryFn: async () => {
      const { data, error: directoryError } = await (supabase as any).rpc(
        "list_grade_request_recipient_directory",
        { p_advisory_class_id: directorySeedClassId },
      );

      if (directoryError) throw directoryError;

      return (Array.isArray(data) ? data : [])
        .map<WorkflowRecipientOption>((row: any): WorkflowRecipientOption => ({
          id: String(row.teacher_id),
          fullName:
            row.full_name ||
            (row.teacher_type === "class_adviser"
              ? "Class Adviser"
              : "Subject Teacher"),
          email: row.email || "No email",
          teacherType:
            row.teacher_type === "class_adviser"
              ? "class_adviser"
              : "subject_teacher",
          assignedToClass: Boolean(row.assigned_to_class),
          subjects: Array.isArray(row.subjects)
            ? row.subjects.map((subject: unknown) => String(subject))
            : [],
        }))
        .filter(
          (recipient) =>
            recipient.id !== workflowContext?.userId &&
            recipient.email.trim().toLowerCase() !==
              workflowContext?.email?.trim().toLowerCase(),
        )
        .sort((first, second) => {
          if (first.teacherType !== second.teacherType) {
            return first.teacherType === "class_adviser" ? -1 : 1;
          }

          return first.email.localeCompare(second.email);
        });
    },
    placeholderData: [],
  });

  const workflowDirectoryOptions = Array.isArray(workflowRecipientOptionsData)
    ? workflowRecipientOptionsData
    : [];

  const workflowClassAdviserOptions = workflowDirectoryOptions.filter(
    (recipient) =>
      recipient.teacherType === "class_adviser" &&
      recipient.id !== workflowContext?.userId &&
      recipient.email.trim().toLowerCase() !==
        workflowContext?.email?.trim().toLowerCase(),
  );

  const workflowSubjectTeacherOptions = workflowDirectoryOptions.filter(
    (recipient) => recipient.teacherType === "subject_teacher",
  );

  const activeRecipientOptions =
    workflowRecipientRole === "class_adviser"
      ? workflowClassAdviserOptions
      : workflowSubjectTeacherOptions;

  const activeWorkflowDirectoryRecipientId = activeRecipientOptions.some(
    (recipient) => recipient.id === workflowDirectoryRecipientId,
  )
    ? workflowDirectoryRecipientId
    : "";

  const selectedWorkflowRecipient = workflowDirectoryOptions.find(
    (recipient) => recipient.id === activeWorkflowDirectoryRecipientId,
  );

  // Both teacher types follow the selected email. Direct table reads cannot
  // reliably see another teacher's classes through RLS, so the database
  // function returns every class owned by the selected account.
  const {
    data: selectedTeacherClassesData,
    isLoading: selectedTeacherClassesLoading,
    error: selectedTeacherClassesError,
  } = useQuery<AdvisoryClassRow[]>({
    queryKey: [
      "request-form-selected-teacher-classes",
      selectedWorkflowRecipient?.id,
    ],
    enabled: Boolean(selectedWorkflowRecipient?.id),
    queryFn: async () => {
      const { data, error: teacherClassError } = await (supabase as any).rpc(
        "list_selected_teacher_classes",
        { p_teacher_id: selectedWorkflowRecipient!.id },
      );

      if (teacherClassError) throw teacherClassError;
      return Array.isArray(data) ? (data as unknown as AdvisoryClassRow[]) : [];
    },
    placeholderData: [],
  });

  const selectedTeacherClasses = useMemo(
    () => dedupeAdvisoryClasses(selectedTeacherClassesData ?? []),
    [selectedTeacherClassesData],
  );

  // The class dropdown always shows every class owned by the teacher selected
  // in the email field, whether that account is a Class Adviser or Subject Teacher.
  const advisoryClasses = selectedTeacherClasses;

  const activeWorkflowClassId =
    workflowClassId &&
    advisoryClasses.some((classRow) => classRow.id === workflowClassId)
      ? workflowClassId
      : "";

  // Attach the request to the class shown in the dropdown. The SQL function
  // verifies that it belongs to the selected teacher account.
  const requestAdvisoryClassId = activeWorkflowClassId;

  const {
    data: workflowReviewStudentsData,
    isLoading: reviewStudentsLoading,
    error: workflowReviewError,
  } = useQuery<Array<WorkflowReviewStudent & { score: number }>>({
    queryKey: [
      "grade-request-review",
      sentWorkflowReviewTarget?.subject.subject_request_id,
    ],
    enabled: Boolean(sentWorkflowReviewTarget?.subject.subject_request_id),
    queryFn: async () => {
      const { data, error: reviewError } = await (supabase as any).rpc(
        "list_my_sent_grade_request_review_by_sex",
        {
          p_subject_request_id:
            sentWorkflowReviewTarget!.subject.subject_request_id,
        },
      );

      if (reviewError) throw reviewError;
      return Array.isArray(data)
        ? (data as unknown as Array<WorkflowReviewStudent & { score: number }>)
        : [];
    },
    placeholderData: [],
  });

  const workflowReviewStudents = Array.isArray(workflowReviewStudentsData)
    ? workflowReviewStudentsData
    : [];
  const workflowReviewStudentGroups = useMemo(
    () => [
      {
        label: "MALE",
        learners: workflowReviewStudents.filter(
          (student) => student.sex === "male",
        ),
      },
      {
        label: "FEMALE",
        learners: workflowReviewStudents.filter(
          (student) => student.sex === "female",
        ),
      },
      {
        label: "SEX NOT SET",
        learners: workflowReviewStudents.filter(
          (student) => student.sex !== "male" && student.sex !== "female",
        ),
      },
    ],
    [workflowReviewStudents],
  );
  const workflowReviewScoreByStudent = new Map(
    workflowReviewStudents.map((student) => [
      student.id,
      Number(student.score),
    ]),
  );
  const reviewScoresLoading = reviewStudentsLoading;

  const {
    data: workflowRowsData,
    isLoading: isWorkflowLoading,
    isPlaceholderData: workflowPlaceholder,
    error: workflowError,
    refetch: refetchWorkflow,
    isFetching: isWorkflowFetching,
  } = useQuery<GradeWorkflowRow[]>({
    queryKey: ["class-adviser-grade-workflow", workflowContext?.userId],
    enabled: Boolean(isClassAdviser && workflowContext?.userId),
    queryFn: async () => {
      const { data, error: workflowQueryError } = await (supabase as any).rpc(
        "list_my_sent_grade_request_workflow",
      );

      if (workflowQueryError) throw workflowQueryError;
      return Array.isArray(data) ? (data as unknown as GradeWorkflowRow[]) : [];
    },
    placeholderData: [],
  });

  const workflowRows = Array.isArray(workflowRowsData) ? workflowRowsData : [];

  const workflowBatches = useMemo<GradeWorkflowBatch[]>(() => {
    const batchMap = new Map<string, GradeWorkflowBatch>();

    workflowRows.forEach((row) => {
      const existing = batchMap.get(row.batch_id);

      if (existing) {
        existing.subjects.push(row);
        return;
      }

      batchMap.set(row.batch_id, {
        batchId: row.batch_id,
        adviserName: row.adviser_name || row.adviser_email || "Class Adviser",
        advisoryClassId: row.advisory_class_id,
        gradeLevel: row.grade_level || "—",
        section: row.section || "—",
        schoolYear: row.school_year || "—",
        gradingPeriod: row.grading_period,
        overallStatus: row.overall_status,
        isFinalized: Boolean(row.is_finalized),
        requestedAt: row.requested_at,
        completedAt: row.completed_at,
        finalizedAt: row.finalized_at,
        subjects: [row],
      });
    });

    return Array.from(batchMap.values())
      .map((batch) => ({
        ...batch,
        subjects: [...batch.subjects].sort((first, second) =>
          first.subject.localeCompare(second.subject),
        ),
      }))
      .sort(
        (first, second) =>
          new Date(second.requestedAt).getTime() -
          new Date(first.requestedAt).getTime(),
      );
  }, [workflowRows]);

  const workflowCounts = useMemo(
    () => ({
      all: workflowBatches.length,
      pending: workflowBatches.filter((batch) =>
        batch.subjects.some((subject) => subject.subject_status === "pending"),
      ).length,
      completed: workflowBatches.filter((batch) =>
        batch.subjects.every(
          (subject) => subject.subject_status === "completed",
        ),
      ).length,
      rejected: workflowBatches.filter((batch) =>
        batch.subjects.some((subject) => subject.subject_status === "rejected"),
      ).length,
    }),
    [workflowBatches],
  );

  const workflowCopy = {
    title: "Class Adviser Grade Request Workflow",
    description:
      "Send one grade request to one selected Subject Teacher or Class Adviser. The recipient enters grades, then completes or rejects the request.",
    button: "Send Grade Request",
    sending: "Sending Request...",
    all: "All Requests",
    pending: "Pending",
    completed: "Completed",
    rejected: "Rejected",
    emptyNoun: "grade requests",
  };

  // The upper workflow section is for active/current requests.
  // Rejected requests belong only in Previous Requests, so remove rejected
  // subject rows from the upper cards without deleting them from history.
  const currentWorkflowBatches = useMemo<GradeWorkflowBatch[]>(
    () =>
      workflowBatches
        .map((batch) => ({
          ...batch,
          subjects: batch.subjects.filter(
            (subject) => subject.subject_status !== "rejected",
          ),
        }))
        .filter((batch) => batch.subjects.length > 0),
    [workflowBatches],
  );

  const visibleWorkflowBatches = useMemo(() => {
    // Rejected requests are intentionally displayed only in Previous Requests.
    if (workflowStatusFilter === "rejected") return [];

    if (workflowStatusFilter === "all") return currentWorkflowBatches;

    if (workflowStatusFilter === "pending") {
      return currentWorkflowBatches.filter((batch) =>
        batch.subjects.some((subject) => subject.subject_status === "pending"),
      );
    }

    return currentWorkflowBatches.filter((batch) =>
      batch.subjects.every((subject) => subject.subject_status === "completed"),
    );
  }, [currentWorkflowBatches, workflowStatusFilter]);

  // Previous Requests for a Class Adviser now uses the SAME workflow data
  // created by create_selected_teacher_grade_request(). This makes every
  // newly sent Grade Request appear in the history automatically.
  const workflowHistoryRows = useMemo(
    () =>
      workflowBatches.flatMap((batch) =>
        batch.subjects.map((subject) => ({
          id: subject.subject_request_id,
          batch,
          subject,
          recipientName:
            subject.subject_teacher_name ||
            subject.subject_teacher_email ||
            (subject.recipient_role === "class_adviser"
              ? "Class Adviser"
              : "Subject Teacher"),
          requestedAt:
            subject.subject_requested_at || batch.requestedAt,
          submittedAt: subject.submitted_at,
          status: batch.isFinalized
            ? ("finalized" as const)
            : batch.overallStatus === "completed"
              ? ("completed" as const)
              : subject.subject_status,
        })),
      ),
    [workflowBatches],
  );

  const dateFilteredWorkflowHistory = useMemo(() => {
    if (dateFilterMode === "all") return workflowHistoryRows;

    return workflowHistoryRows.filter((row) => {
      const requestDate = new Date(row.requestedAt);

      if (Number.isNaN(requestDate.getTime())) return false;

      if (dateFilterMode === "day") {
        return toDateInputValue(requestDate) === selectedDay;
      }

      if (dateFilterMode === "month") {
        return toMonthInputValue(requestDate) === selectedMonth;
      }

      return String(requestDate.getFullYear()) === selectedYear;
    });
  }, [
    workflowHistoryRows,
    dateFilterMode,
    selectedDay,
    selectedMonth,
    selectedYear,
  ]);

  const filteredWorkflowHistory = useMemo(() => {
    if (previousWorkflowStatusFilter === "all") {
      return dateFilteredWorkflowHistory;
    }

    return dateFilteredWorkflowHistory.filter((row) => {
      if (previousWorkflowStatusFilter === "pending") {
        return row.subject.subject_status === "pending";
      }

      if (previousWorkflowStatusFilter === "completed") {
        return (
          row.subject.subject_status === "completed" &&
          !row.batch.isFinalized
        );
      }

      if (previousWorkflowStatusFilter === "rejected") {
        return (
          row.subject.subject_status === "rejected" &&
          !row.batch.isFinalized
        );
      }

      return row.batch.isFinalized;
    });
  }, [dateFilteredWorkflowHistory, previousWorkflowStatusFilter]);

  const workflowHistoryCounts = useMemo(
    () => ({
      all: dateFilteredWorkflowHistory.length,
      pending: dateFilteredWorkflowHistory.filter(
        (row) => row.subject.subject_status === "pending",
      ).length,
      completed: dateFilteredWorkflowHistory.filter(
        (row) =>
          row.subject.subject_status === "completed" &&
          !row.batch.isFinalized,
      ).length,
      rejected: dateFilteredWorkflowHistory.filter(
        (row) =>
          row.subject.subject_status === "rejected" &&
          !row.batch.isFinalized,
      ).length,
      finalized: dateFilteredWorkflowHistory.filter(
        (row) => row.batch.isFinalized,
      ).length,
    }),
    [dateFilteredWorkflowHistory],
  );

  const createWorkflowMutation = useMutation({
    mutationFn: async () => {
      if (!activeWorkflowClassId) {
        throw new Error("Select an advisory class first.");
      }

      if (!selectedWorkflowRecipient) {
        throw new Error(
          `Select one ${
            workflowRecipientRole === "class_adviser"
              ? "Class Adviser"
              : "Subject Teacher"
          } email first.`,
        );
      }

      const requestFunction = "create_selected_teacher_grade_request";
      const requestPayload = {
        p_advisory_class_id: requestAdvisoryClassId,
        p_grading_period: workflowPeriod,
        p_recipient_id: selectedWorkflowRecipient.id,
        p_recipient_role: selectedWorkflowRecipient.teacherType,
        p_message: workflowMessage.trim() || null,
      };

      const { data, error: createError } = await (supabase as any).rpc(
        requestFunction,
        requestPayload,
      );

      if (createError) throw createError;
      return String(data);
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["class-adviser-grade-workflow"],
      });
      toast.success("The Grade Request was sent to the selected teacher.");
      setWorkflowMessage("");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const finalizeWorkflowMutation = useMutation({
    mutationFn: async (batch: GradeWorkflowBatch) => {
      const finalizedSubject =
        batch.subjects.find((subject) => subject.submitted_grade_count > 0) ??
        batch.subjects[0];

      if (!finalizedSubject?.subject_request_id) {
        throw new Error(
          "This Grade Request has no submitted learner records to add as a subject.",
        );
      }

      const { error: finalizeError } = await (supabase as any).rpc(
        "finalize_grade_request_batch",
        {
          p_batch_id: batch.batchId,
        },
      );

      if (finalizeError) throw finalizeError;
      return {
        batch,
        subjectRequestId: finalizedSubject.subject_request_id,
      };
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["class-adviser-grade-workflow"],
      });
      await queryClient.invalidateQueries({
        queryKey: ["summary-finalized-grade-request-forms"],
      });

      toast.success(
        "Grades were finalized. You can add this finalized request later from Summary of Grades → Add Subject.",
      );
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const deleteSentWorkflowMutation = useMutation({
    mutationFn: async (subjectRequestId: string) => {
      const { error: deleteError } = await (supabase as any).rpc(
        "delete_sent_grade_request",
        {
          p_subject_request_id: subjectRequestId,
        },
      );

      if (deleteError) throw deleteError;
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["class-adviser-grade-workflow"],
      });
      setSentWorkflowDeleteTarget(null);
      toast.success("The sent Grade Request was permanently deleted.");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const classAdviserIds = useMemo(
    () => [
      ...new Set(
        requests
          .map((request) => request.subject_teacher_id)
          .filter((id): id is string => Boolean(id)),
      ),
    ],
    [requests],
  );

  const {
    data: classAdviserProfilesData,
    isLoading: classAdvisersLoading,
    isPlaceholderData: classAdvisersPlaceholder,
  } = useQuery<ClassAdviserProfile[]>({
    queryKey: ["request-class-advisers", classAdviserIds.join(",")],
    enabled: classAdviserIds.length > 0,
    queryFn: async () => {
      const { data, error: profileError } = await (supabase as any)
        .from("profiles")
        .select("id, full_name, email")
        .in("id", classAdviserIds);

      if (profileError) throw profileError;

      return Array.isArray(data)
        ? (data as unknown as ClassAdviserProfile[])
        : [];
    },
    placeholderData: [],
  });

  const classAdviserProfiles = Array.isArray(classAdviserProfilesData)
    ? classAdviserProfilesData
    : [];

  const classAdviserById = useMemo(
    () =>
      new Map(
        classAdviserProfiles.map((adviser) => [
          adviser.id,
          adviser.full_name || adviser.email || "Class Adviser",
        ]),
      ),
    [classAdviserProfiles],
  );

  const { data: studentSexRowsData } = useQuery<StudentSexRow[]>({
    queryKey: ["request-form-student-sex-directory"],
    queryFn: async () => {
      const { data, error: studentError } = await (supabase as any)
        .from("students")
        .select("first_name, last_name, sex");

      if (studentError) throw studentError;

      return Array.isArray(data) ? (data as unknown as StudentSexRow[]) : [];
    },
    placeholderData: [],
  });

  const studentSexRows = Array.isArray(studentSexRowsData)
    ? studentSexRowsData
    : [];

  const studentSexByName = useMemo(() => {
    const directory = new Map<string, LearnerSex>();

    studentSexRows.forEach((student) => {
      const fullName = `${student.last_name ?? ""}, ${student.first_name ?? ""}`;

      const normalizedName = normalizeLearnerName(fullName);
      const normalizedSex = String(student.sex ?? "")
        .trim()
        .toLowerCase();

      if (!normalizedName) return;

      directory.set(
        normalizedName,
        normalizedSex === "male"
          ? "male"
          : normalizedSex === "female"
            ? "female"
            : "unknown",
      );
    });

    return directory;
  }, [studentSexRows]);

  const availableYears = useMemo(() => {
    const years = requests
      .map((request) => new Date(request.created_at).getFullYear())
      .filter((year) => Number.isFinite(year));

    workflowHistoryRows.forEach((row) => {
      const year = new Date(row.requestedAt).getFullYear();
      if (Number.isFinite(year)) years.push(year);
    });

    years.push(today.getFullYear());

    return [...new Set(years)].sort((first, second) => second - first);
  }, [requests, workflowHistoryRows, today]);

  const dateFilteredRequests = useMemo(() => {
    if (dateFilterMode === "all") return requests;

    return requests.filter((request) => {
      const requestDate = new Date(request.created_at);

      if (Number.isNaN(requestDate.getTime())) return false;

      if (dateFilterMode === "day") {
        return toDateInputValue(requestDate) === selectedDay;
      }

      if (dateFilterMode === "month") {
        return toMonthInputValue(requestDate) === selectedMonth;
      }

      return String(requestDate.getFullYear()) === selectedYear;
    });
  }, [requests, dateFilterMode, selectedDay, selectedMonth, selectedYear]);

  const filteredRequests = useMemo(() => {
    if (statusFilter === "all") return dateFilteredRequests;

    return dateFilteredRequests.filter(
      (request) => request.status === statusFilter,
    );
  }, [dateFilteredRequests, statusFilter]);

  const counts = useMemo(
    () => ({
      all: dateFilteredRequests.length,
      pending: dateFilteredRequests.filter(
        (request) => request.status === "pending",
      ).length,
      in_progress: dateFilteredRequests.filter(
        (request) => request.status === "in_progress",
      ).length,
      completed: dateFilteredRequests.filter(
        (request) => request.status === "completed",
      ).length,
      rejected: dateFilteredRequests.filter(
        (request) => request.status === "rejected",
      ).length,
    }),
    [dateFilteredRequests],
  );

  const openAttachment = async (request: RequestRow) => {
    if (!request.attachment_url) return;

    const { data, error: signedUrlError } = await supabase.storage
      .from("grade-request-attachments")
      .createSignedUrl(request.attachment_url, 60);

    if (signedUrlError) {
      toast.error(signedUrlError.message);
      return;
    }

    window.open(data.signedUrl, "_blank", "noopener,noreferrer");
  };

  const deleteRequestMutation = useMutation({
    mutationFn: async (request: RequestRow) => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      const { error: deleteError } = await (supabase as any)
        .from("grade_requests")
        .delete()
        .eq("id", request.id)
        .eq("requester_id", user.id);

      if (deleteError) throw deleteError;

      let attachmentDeleteFailed = false;

      if (request.attachment_url) {
        const { error: attachmentError } = await supabase.storage
          .from("grade-request-attachments")
          .remove([request.attachment_url]);

        attachmentDeleteFailed = Boolean(attachmentError);
      }

      return { attachmentDeleteFailed };
    },
    onSuccess: async ({ attachmentDeleteFailed }) => {
      await queryClient.invalidateQueries({
        queryKey: ["my-grade-requests"],
      });

      setDeleteTarget(null);
      setSelectedRequest(null);

      if (attachmentDeleteFailed) {
        toast.warning(
          "The request was deleted, but its attachment could not be removed.",
        );
        return;
      }

      toast.success("Request permanently deleted.");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const requestFormLoading =
    isLoading ||
    requestsPlaceholder ||
    workflowContextLoading ||
    (classAdviserIds.length > 0 &&
      (classAdvisersLoading || classAdvisersPlaceholder)) ||
    (isClassAdviser &&
      (ownedClassesLoading ||
        ownedClassesPlaceholder ||
        isWorkflowLoading ||
        workflowPlaceholder ||
        (Boolean(directorySeedClassId) &&
          (workflowRecipientsLoading || workflowRecipientsPlaceholder))));

  if (requestFormLoading) {
    return (
      <RequestFormSkeleton
        showWorkflow={workflowContext?.teacherType !== "subject_teacher"}
      />
    );
  }

  return (
    <div className="space-y-5 pb-24 md:pb-6">
      <section className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card px-5 py-4 shadow-sm">
        <div className="flex items-center gap-4">
          <Link
            to="/subject-teacher-dashboard"
            className="inline-flex items-center gap-1.5 rounded-xl border px-3 py-2 text-sm font-medium hover:bg-muted"
          >
            <ArrowLeft className="size-4" />
            Back
          </Link>

          <div>
            <div className="flex items-center gap-2 text-xl font-semibold">
              <Inbox className="size-6 text-primary" />
              Request Form
            </div>

            <p className="mt-0.5 text-sm text-muted-foreground">
              Track Class Adviser and Subject Teacher requests in separate
              flows.
            </p>
          </div>
        </div>

        <Button
          variant="outline"
          onClick={() => refetch()}
          disabled={isFetching}
        >
          <RefreshCw
            className={`mr-2 size-4 ${isFetching ? "animate-spin" : ""}`}
          />
          Refresh
        </Button>
      </section>

      {/*
        Added only: Class Adviser grade-request workflow.
        The complete original Previous Requests UI below is preserved.
      */}
      {isClassAdviser && (
        <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
          <div className="flex flex-wrap items-start justify-between gap-3 border-b px-5 py-4">
            <div>
              <div className="flex items-center gap-2 text-lg font-semibold">
                <FileCheck2 className="size-5 text-primary" />
                {workflowCopy.title}
              </div>
              <p className="mt-1 max-w-3xl text-sm text-muted-foreground">
                {workflowCopy.description}
              </p>
            </div>

            <Button
              variant="outline"
              onClick={() => refetchWorkflow()}
              disabled={isWorkflowFetching}
            >
              <RefreshCw
                className={`mr-2 size-4 ${isWorkflowFetching ? "animate-spin" : ""}`}
              />
              Refresh Workflow
            </Button>
          </div>

          <div className="border-b bg-muted/20 p-5">
            <div className="grid gap-3 lg:grid-cols-[minmax(220px,0.65fr)_minmax(320px,1fr)_minmax(180px,0.55fr)] lg:items-start">
              <div>
                <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                  Teacher Type
                </label>
                <Select
                  value={workflowRecipientRole}
                  onValueChange={(value) => {
                    const role = value as WorkflowRecipientRole;
                    setWorkflowRecipientRole(role);
                    setWorkflowDirectoryRecipientId("");
                    setWorkflowClassId("");
                  }}
                >
                  <SelectTrigger className="bg-background">
                    <Users className="mr-2 size-4 text-primary" />
                    <SelectValue placeholder="Select teacher type" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="class_adviser">Class Adviser</SelectItem>
                    <SelectItem value="subject_teacher">
                      Subject Teacher
                    </SelectItem>
                  </SelectContent>
                </Select>
                <p className="mt-1 text-xs text-muted-foreground">
                  Select which teacher type will receive the Grade Request and
                  submit grades back to you.
                </p>
              </div>

              <div>
                <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                  {workflowRecipientRole === "class_adviser"
                    ? "Class Adviser Email"
                    : "Subject Teacher Email"}
                </label>
                <Select
                  value={activeWorkflowDirectoryRecipientId}
                  onValueChange={(recipientId) => {
                    setWorkflowDirectoryRecipientId(recipientId);
                    setWorkflowClassId("");
                  }}
                  disabled={workflowRecipientsLoading}
                >
                  <SelectTrigger className="bg-background">
                    <Mail className="mr-2 size-4 text-primary" />
                    <SelectValue
                      placeholder={
                        workflowRecipientRole === "class_adviser"
                          ? "Select a Class Adviser email"
                          : "Select a Subject Teacher email"
                      }
                    />
                  </SelectTrigger>
                  <SelectContent>
                    {activeRecipientOptions.map((recipient) => (
                      <SelectItem key={recipient.id} value={recipient.id}  className="justify-center pl-2 text-center [&>span:first-child]:hidden">
                        
                      {recipient.fullName}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                <p className="mt-1 text-xs text-muted-foreground">
                  Select one email to load all classes owned by that selected
                  teacher.
                </p>
              </div>

              <div>
                <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                  Grading Period
                </label>
                <Select
                  value={workflowPeriod}
                  onValueChange={(value) =>
                    setWorkflowPeriod(value as GradingPeriod)
                  }
                >
                  <SelectTrigger className="bg-background">
                    <SelectValue placeholder="Select grading period" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="1">Term 1</SelectItem>
                    <SelectItem value="2">Term 2</SelectItem>
                    <SelectItem value="3">Term 3</SelectItem>
                    <SelectItem value="final">Final Grade</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>

            <div className="mt-3 grid gap-3 lg:grid-cols-[minmax(300px,0.95fr)_minmax(0,1.4fr)_auto] lg:items-start">
              <div>
                <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                  Advisory Class
                </label>
                <Select
                  value={activeWorkflowClassId}
                  onValueChange={setWorkflowClassId}
                  disabled={
                    advisoryClasses.length === 0 ||
                    !selectedWorkflowRecipient ||
                    selectedTeacherClassesLoading
                  }
                >
                  <SelectTrigger className="bg-background">
                    <SelectValue
                      placeholder={
                        !selectedWorkflowRecipient
                          ? `Select ${
                              workflowRecipientRole === "class_adviser"
                                ? "Class Adviser"
                                : "Subject Teacher"
                            } email first`
                          : selectedTeacherClassesLoading
                            ? "Loading teacher classes..."
                            : "Select advisory class"
                      }
                    />
                  </SelectTrigger>
                  <SelectContent>
                    {advisoryClasses.map((classRow) => (
                      <SelectItem key={classRow.id} value={classRow.id}>
                        {formatAdvisoryClassLabel(classRow)}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                <p className="mt-1 text-xs text-muted-foreground">
                  Select a Grade Level and Section owned by the selected
                  teacher.
                </p>
              </div>

              <div>
                <div className="mb-1.5 flex items-center justify-between gap-3">
                  <label
                    htmlFor="workflow-request-message"
                    className="text-xs font-semibold uppercase tracking-wide text-muted-foreground"
                  >
                    Message (Optional)
                  </label>
                  <span className="text-xs text-muted-foreground">
                    {workflowMessage.length}/1000
                  </span>
                </div>
                <div className="relative">
                  <MessageSquareText className="pointer-events-none absolute left-3 top-3 size-4 text-muted-foreground" />
                  <textarea
                    id="workflow-request-message"
                    value={workflowMessage}
                    maxLength={1000}
                    onChange={(event) => setWorkflowMessage(event.target.value)}
                    placeholder="Example: Please submit Term 1 grades before Friday."
                    className="min-h-20 w-full resize-y rounded-xl border bg-background py-2.5 pl-10 pr-3 text-sm outline-none transition placeholder:text-muted-foreground focus:border-primary focus:ring-2 focus:ring-primary/20"
                  />
                </div>
              </div>

              <Button
                onClick={() => createWorkflowMutation.mutate()}
                disabled={
                  !activeWorkflowClassId ||
                  !requestAdvisoryClassId ||
                  !selectedWorkflowRecipient ||
                  workflowRecipientsLoading ||
                  selectedTeacherClassesLoading ||
                  createWorkflowMutation.isPending
                }
                className="min-w-48"
              >
                {createWorkflowMutation.isPending ? (
                  <LoaderCircle className="mr-2 size-4 animate-spin" />
                ) : (
                  <Send className="mr-2 size-4" />
                )}
                {createWorkflowMutation.isPending
                  ? workflowCopy.sending
                  : workflowCopy.button}
              </Button>
            </div>

            {workflowRecipientsError && (
              <p className="mt-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {(workflowRecipientsError as Error).message}
              </p>
            )}

            {selectedTeacherClassesError && (
              <p className="mt-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {(selectedTeacherClassesError as Error).message}. Run the
                selected teacher classes SQL migration, then refresh this page.
              </p>
            )}

            {selectedWorkflowRecipient &&
              !selectedTeacherClassesLoading &&
              !selectedTeacherClassesError &&
              selectedTeacherClasses.length === 0 && (
                <p className="mt-3 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
                  No classes are assigned to the selected teacher.
                </p>
              )}

            {workflowError && (
              <p className="mt-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {(workflowError as Error).message}. Apply the Grade Request
                Workflow SQL migration if the workflow tables have not been
                created yet.
              </p>
            )}
          </div>

          <div className="grid gap-3 border-b bg-background px-5 py-4 sm:grid-cols-2 lg:grid-cols-4">
            <CountCard
              label={workflowCopy.all}
              value={workflowCounts.all}
              active={workflowStatusFilter === "all"}
              onClick={() => setWorkflowStatusFilter("all")}
            />
            <CountCard
              label={workflowCopy.pending}
              value={workflowCounts.pending}
              active={workflowStatusFilter === "pending"}
              onClick={() => setWorkflowStatusFilter("pending")}
            />
            <CountCard
              label={workflowCopy.completed}
              value={workflowCounts.completed}
              active={workflowStatusFilter === "completed"}
              onClick={() => setWorkflowStatusFilter("completed")}
            />
            <CountCard
              label={workflowCopy.rejected}
              value={workflowCounts.rejected}
              active={workflowStatusFilter === "rejected"}
              onClick={() => setWorkflowStatusFilter("rejected")}
            />
          </div>

          <div className="space-y-4 p-5">
            {isWorkflowLoading && (
              <div className="py-12 text-center text-sm text-muted-foreground">
                <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                Loading grade-request workflow...
              </div>
            )}

            {!isWorkflowLoading &&
              visibleWorkflowBatches.map((batch) => {
                const submittedCount = batch.subjects.filter(
                  (subject) => subject.subject_status === "completed",
                ).length;
                const totalSubjects = batch.subjects.length;
                const progress = totalSubjects
                  ? Math.round((submittedCount / totalSubjects) * 100)
                  : 0;

                return (
                  <article
                    key={batch.batchId}
                    className="overflow-hidden rounded-2xl border bg-background"
                  >
                    <div className="flex flex-wrap items-start justify-between gap-3 border-b px-4 py-4">
                      <div>
                        <div className="flex flex-wrap items-center gap-2">
                          <h3 className="font-semibold">
                            Grade {batch.gradeLevel} · {batch.section}
                          </h3>
                          <span className="rounded-lg border bg-muted px-2.5 py-1 text-xs font-medium">
                            {formatTerm(batch.gradingPeriod)}
                          </span>
                          <span
                            className={`rounded-lg border px-2.5 py-1 text-xs font-medium ${
                              workflowOverallStyle[batch.overallStatus]
                            }`}
                          >
                            {batch.overallStatus === "completed"
                              ? "Completed"
                              : "Pending"}
                          </span>
                          {batch.isFinalized && (
                            <span className="rounded-lg border border-violet-200 bg-violet-50 px-2.5 py-1 text-xs font-medium text-violet-700">
                              Finalized
                            </span>
                          )}
                        </div>
                        <p className="mt-1 text-sm text-muted-foreground">
                          School Year {batch.schoolYear} · Requested{" "}
                          {formatDate(batch.requestedAt)}
                        </p>
                      </div>

                      <div className="flex flex-wrap gap-2">
                        {batch.overallStatus === "completed" &&
                          !batch.isFinalized && (
                            <Button
                              size="sm"
                              onClick={() =>
                                finalizeWorkflowMutation.mutate(batch)
                              }
                              disabled={finalizeWorkflowMutation.isPending}
                            >
                              {finalizeWorkflowMutation.isPending ? (
                                <LoaderCircle className="mr-2 size-4 animate-spin" />
                              ) : (
                                <CheckCircle2 className="mr-2 size-4" />
                              )}
                              Finalize Grades
                            </Button>
                          )}

                      </div>
                    </div>

                    <div className="px-4 py-3">
                      <div className="flex items-center justify-between gap-3 text-sm">
                        <span className="font-medium">
                          {submittedCount} of {totalSubjects} subjects completed
                        </span>
                        <span className="text-muted-foreground">
                          {progress}%
                        </span>
                      </div>
                      <div className="mt-2 h-2 overflow-hidden rounded-full bg-muted">
                        <div
                          className="h-full rounded-full bg-primary transition-all"
                          style={{ width: `${progress}%` }}
                        />
                      </div>
                    </div>

                    <div className="overflow-x-auto border-t">
                      <table className="w-full min-w-[850px] text-sm">
                        <thead className="bg-muted/30 text-xs uppercase text-muted-foreground">
                          <tr>
                            <th className="px-4 py-3 text-left">Recipient</th>
                            <th className="px-4 py-3 text-left">Subject</th>
                            <th className="px-4 py-3 text-left">Requested</th>
                            <th className="px-4 py-3 text-left">Completed</th>
                            <th className="px-4 py-3 text-center">Grades</th>
                            <th className="px-4 py-3 text-left">Message</th>
                            <th className="px-4 py-3 text-left">Status</th>
                            <th className="px-4 py-3 text-center">Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {batch.subjects.map((subject) => (
                            <tr
                              key={subject.subject_request_id}
                              className="border-t"
                            >
                              <td className="px-4 py-3">
                                <div className="font-medium">
                                  {subject.subject_teacher_name ||
                                    subject.subject_teacher_email ||
                                    (subject.recipient_role === "class_adviser"
                                      ? "Class Adviser"
                                      : "Subject Teacher")}
                                </div>
                                {subject.subject_teacher_name &&
                                  subject.subject_teacher_email && (
                                    <div className="text-xs text-muted-foreground">
                                      {subject.subject_teacher_email}
                                    </div>
                                  )}
                                <div className="mt-1 text-xs capitalize text-muted-foreground">
                                  {subject.recipient_role === "class_adviser"
                                    ? "Class Adviser"
                                    : "Subject Teacher"}
                                </div>
                              </td>
                              <td className="px-4 py-3 font-medium">
                                {subject.subject}
                              </td>
                              <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                                {formatDate(subject.subject_requested_at)}
                              </td>
                              <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                                {subject.submitted_at
                                  ? formatDate(subject.submitted_at)
                                  : "—"}
                              </td>
                              <td className="px-4 py-3 text-center">
                                {subject.submitted_grade_count || 0}
                              </td>
                              <td className="max-w-72 px-4 py-3 text-muted-foreground">
                                <p className="line-clamp-2">
                                  {subject.request_message || "—"}
                                </p>
                              </td>
                              <td className="px-4 py-3">
                                <span
                                  className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1 text-xs font-medium ${
                                    workflowSubjectStyle[subject.subject_status]
                                  }`}
                                >
                                  {subject.subject_status === "completed" ? (
                                    <CheckCircle2 className="size-3.5" />
                                  ) : subject.subject_status === "rejected" ? (
                                    <XCircle className="size-3.5" />
                                  ) : (
                                    <Clock3 className="size-3.5" />
                                  )}
                                  {subject.subject_status === "completed"
                                    ? "Completed"
                                    : subject.subject_status === "rejected"
                                      ? "Rejected"
                                      : "Pending"}
                                </span>
                              </td>
                              <td className="px-4 py-3 text-center">
                                <div className="flex items-center justify-center gap-1">
                                  {subject.subject_status === "completed" && (
                                    <Button
                                      type="button"
                                      size="icon"
                                      variant="ghost"
                                      className="text-primary hover:bg-primary/10"
                                      title="Review submitted grades"
                                      aria-label={`Review submitted grades for ${subject.subject}`}
                                      onClick={() =>
                                        setSentWorkflowReviewTarget({
                                          batch,
                                          subject,
                                        })
                                      }
                                    >
                                      <Eye className="size-4" />
                                    </Button>
                                  )}
                                  <Button
                                    type="button"
                                    size="icon"
                                    variant="ghost"
                                    className="text-destructive hover:bg-destructive/10 hover:text-destructive"
                                    title="Delete sent Grade Request"
                                    aria-label={`Delete sent Grade Request for ${subject.subject}`}
                                    onClick={() =>
                                      setSentWorkflowDeleteTarget({
                                        batch,
                                        subject,
                                      })
                                    }
                                  >
                                    <Trash2 className="size-4" />
                                  </Button>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </article>
                );
              })}

            {!isWorkflowLoading &&
              visibleWorkflowBatches.length === 0 &&
              !workflowError && (
                <div className="rounded-2xl border border-dashed px-5 py-12 text-center text-muted-foreground">
                  {workflowStatusFilter === "rejected" ? (
                    <>
                      Rejected requests are kept only in{" "}
                      <span className="font-medium text-foreground">
                        Previous Requests
                      </span>{" "}
                      below.
                    </>
                  ) : (
                    <>
                      No{" "}
                      {workflowStatusFilter === "all"
                        ? ""
                        : `${workflowStatusFilter} `}
                      {workflowCopy.emptyNoun} in this flow.
                    </>
                  )}
                </div>
              )}
          </div>
        </section>
      )}

      {isClassAdviser ? (
        <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
          <CountCard
            label="All Requests"
            value={workflowHistoryCounts.all}
            active={previousWorkflowStatusFilter === "all"}
            onClick={() => setPreviousWorkflowStatusFilter("all")}
          />

          <CountCard
            label="Pending"
            value={workflowHistoryCounts.pending}
            active={previousWorkflowStatusFilter === "pending"}
            onClick={() => setPreviousWorkflowStatusFilter("pending")}
          />

          <CountCard
            label="Completed"
            value={workflowHistoryCounts.completed}
            active={previousWorkflowStatusFilter === "completed"}
            onClick={() => setPreviousWorkflowStatusFilter("completed")}
          />

          <CountCard
            label="Rejected"
            value={workflowHistoryCounts.rejected}
            active={previousWorkflowStatusFilter === "rejected"}
            onClick={() => setPreviousWorkflowStatusFilter("rejected")}
          />

          <CountCard
            label="Finalized"
            value={workflowHistoryCounts.finalized}
            active={previousWorkflowStatusFilter === "finalized"}
            onClick={() => setPreviousWorkflowStatusFilter("finalized")}
          />
        </section>
      ) : (
        <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
          <CountCard
            label="All Requests"
            value={counts.all}
            active={statusFilter === "all"}
            onClick={() => setStatusFilter("all")}
          />

          <CountCard
            label="Pending"
            value={counts.pending}
            active={statusFilter === "pending"}
            onClick={() => setStatusFilter("pending")}
          />

          <CountCard
            label="In Progress"
            value={counts.in_progress}
            active={statusFilter === "in_progress"}
            onClick={() => setStatusFilter("in_progress")}
          />

          <CountCard
            label="Completed"
            value={counts.completed}
            active={statusFilter === "completed"}
            onClick={() => setStatusFilter("completed")}
          />

          <CountCard
            label="Rejected"
            value={counts.rejected}
            active={statusFilter === "rejected"}
            onClick={() => setStatusFilter("rejected")}
          />
        </section>
      )}

      {isClassAdviser ? (
        <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
            <div>
              <h2 className="text-lg font-semibold">Previous Requests</h2>
              <p className="text-sm text-muted-foreground">
                {filteredWorkflowHistory.length} request
                {filteredWorkflowHistory.length === 1 ? "" : "s"} shown
              </p>
            </div>

            <div className="flex flex-wrap items-center justify-end gap-2">
              <Select
                value={dateFilterMode}
                onValueChange={(value) =>
                  setDateFilterMode(value as DateFilterMode)
                }
              >
                <SelectTrigger className="w-44">
                  <CalendarDays className="mr-2 size-4 text-primary" />
                  <SelectValue placeholder="All Dates" />
                </SelectTrigger>

                <SelectContent>
                  <SelectItem value="all">All Dates</SelectItem>
                  <SelectItem value="day">Specific Day</SelectItem>
                  <SelectItem value="month">Specific Month</SelectItem>
                  <SelectItem value="year">Specific Year</SelectItem>
                </SelectContent>
              </Select>

              {dateFilterMode === "day" && (
                <input
                  type="date"
                  value={selectedDay}
                  onChange={(event) => setSelectedDay(event.target.value)}
                  className="h-10 rounded-md border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
                  aria-label="Filter workflow requests by day"
                />
              )}

              {dateFilterMode === "month" && (
                <input
                  type="month"
                  value={selectedMonth}
                  onChange={(event) => setSelectedMonth(event.target.value)}
                  className="h-10 rounded-md border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
                  aria-label="Filter workflow requests by month"
                />
              )}

              {dateFilterMode === "year" && (
                <Select value={selectedYear} onValueChange={setSelectedYear}>
                  <SelectTrigger className="w-32">
                    <SelectValue placeholder="Year" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableYears.map((year) => (
                      <SelectItem key={year} value={String(year)}>
                        {year}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              )}

              <Select
                value={previousWorkflowStatusFilter}
                onValueChange={(value) =>
                  setPreviousWorkflowStatusFilter(
                    value as PreviousWorkflowStatusFilter,
                  )
                }
              >
                <SelectTrigger className="w-44">
                  <SelectValue placeholder="All Status" />
                </SelectTrigger>

                <SelectContent>
                  <SelectItem value="all">All Status</SelectItem>
                  <SelectItem value="pending">Pending</SelectItem>
                  <SelectItem value="completed">Completed</SelectItem>
                  <SelectItem value="rejected">Rejected</SelectItem>
                  <SelectItem value="finalized">Finalized</SelectItem>
                </SelectContent>
              </Select>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full min-w-[1150px] text-sm">
              <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
                <tr>
                  <th className="px-4 py-3 text-left">Recipient</th>
                  <th className="px-4 py-3 text-left">Subject</th>
                  <th className="px-4 py-3 text-left">Class</th>
                  <th className="px-4 py-3 text-left">Term</th>
                  <th className="px-4 py-3 text-left">Message</th>
                  <th className="px-4 py-3 text-left">Status</th>
                  <th className="px-4 py-3 text-left">Requested</th>
                  <th className="px-4 py-3 text-left">Completed</th>
                  <th className="px-4 py-3 text-right">Action</th>
                </tr>
              </thead>

              <tbody>
                {isWorkflowLoading && (
                  <tr>
                    <td
                      colSpan={9}
                      className="px-4 py-16 text-center text-muted-foreground"
                    >
                      <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                      Loading previous requests...
                    </td>
                  </tr>
                )}

                {!isWorkflowLoading &&
                  filteredWorkflowHistory.map((row) => (
                    <tr key={row.id} className="border-t">
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-2">
                          <div className="grid size-8 place-items-center rounded-lg bg-primary/10 text-primary">
                            <UserRound className="size-4" />
                          </div>
                          <div>
                            <div className="font-medium">{row.recipientName}</div>
                            <div className="text-xs capitalize text-muted-foreground">
                              {row.subject.recipient_role === "class_adviser"
                                ? "Class Adviser"
                                : "Subject Teacher"}
                            </div>
                          </div>
                        </div>
                      </td>

                      <td className="px-4 py-3 font-medium">
                        {row.subject.subject}
                      </td>

                      <td className="px-4 py-3">
                        {row.batch.gradeLevel} · {row.batch.section}
                      </td>

                      <td className="px-4 py-3">
                        {formatTerm(row.batch.gradingPeriod)}
                      </td>

                      <td className="w-[260px] max-w-[260px] px-4 py-3">
                        {(() => {
                          const message =
                            row.status === "rejected" ||
                            row.status === "completed" ||
                            row.status === "finalized"
                              ? row.subject.teacher_note?.trim() ||
                                row.subject.request_message?.trim() ||
                                ""
                              : row.subject.request_message?.trim() || "";

                          if (!message) {
                            return (
                              <span className="text-muted-foreground">—</span>
                            );
                          }

                          const isLong = message.length > 42;

                          return (
                            <div className="flex min-w-0 items-center gap-2">
                              <span
                                className="min-w-0 flex-1 truncate text-muted-foreground"
                                title={message}
                              >
                                {message}
                              </span>

                              {isLong && (
                                <Button
                                  type="button"
                                  variant="outline"
                                  size="sm"
                                  className="h-7 shrink-0 px-2 text-xs"
                                  onClick={() =>
                                    setRecipientMessageDialog({
                                      recipientName: row.recipientName,
                                      subject: row.subject.subject,
                                      classLabel: `${row.batch.gradeLevel} · ${row.batch.section}`,
                                      status:
                                        row.status === "finalized"
                                          ? "Finalized"
                                          : row.status === "completed"
                                            ? "Completed"
                                            : row.status === "rejected"
                                              ? "Rejected"
                                              : "Pending",
                                      message,
                                    })
                                  }
                                >
                                  <MessageSquareText className="mr-1 size-3.5" />
                                  View
                                </Button>
                              )}
                            </div>
                          );
                        })()}
                      </td>

                      <td className="px-4 py-3">
                        {row.batch.isFinalized ? (
                          <span className="inline-flex items-center gap-1.5 rounded-lg border border-violet-200 bg-violet-50 px-2.5 py-1 text-xs font-medium text-violet-700">
                            <CheckCircle2 className="size-3.5" />
                            Finalized
                          </span>
                        ) : row.batch.overallStatus === "completed" ? (
                          <span className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-700">
                            <CheckCircle2 className="size-3.5" />
                            Completed
                          </span>
                        ) : (
                          <span
                            className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1 text-xs font-medium ${
                              workflowSubjectStyle[row.subject.subject_status]
                            }`}
                          >
                            {row.subject.subject_status === "completed" ? (
                              <CheckCircle2 className="size-3.5" />
                            ) : row.subject.subject_status === "rejected" ? (
                              <XCircle className="size-3.5" />
                            ) : (
                              <Clock3 className="size-3.5" />
                            )}
                            {row.subject.subject_status === "completed"
                              ? "Completed"
                              : row.subject.subject_status === "rejected"
                                ? "Rejected"
                                : "Pending"}
                          </span>
                        )}
                      </td>

                      <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                        <div>{formatDate(row.requestedAt)}</div>
                        <div className="mt-0.5 text-xs">
                          {formatRequestAge(row.requestedAt)}
                        </div>
                      </td>

                      <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                        {row.submittedAt ? formatDate(row.submittedAt) : "—"}
                      </td>

                      <td className="px-4 py-3 text-right">
                        <div className="flex items-center justify-end gap-2">
                          {row.subject.subject_status === "completed" && (
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() =>
                                setSentWorkflowReviewTarget({
                                  batch: row.batch,
                                  subject: row.subject,
                                })
                              }
                              className="text-primary"
                            >
                              <Eye className="mr-1 size-4" />
                              View
                            </Button>
                          )}

                          <Button
                            variant="outline"
                            size="icon"
                            onClick={() =>
                              setSentWorkflowDeleteTarget({
                                batch: row.batch,
                                subject: row.subject,
                              })
                            }
                            className="border-destructive/40 text-destructive hover:bg-destructive/10"
                            title="Permanently delete sent request"
                            aria-label={`Permanently delete ${row.subject.subject} request`}
                          >
                            <Trash2 className="size-4" />
                          </Button>
                        </div>
                      </td>
                    </tr>
                  ))}

                {!isWorkflowLoading &&
                  filteredWorkflowHistory.length === 0 &&
                  !workflowError && (
                    <tr>
                      <td
                        colSpan={9}
                        className="px-4 py-20 text-center text-muted-foreground"
                      >
                        No previous requests found for this status.
                      </td>
                    </tr>
                  )}

                {workflowError && (
                  <tr>
                    <td
                      colSpan={9}
                      className="px-4 py-16 text-center text-destructive"
                    >
                      {(workflowError as Error).message}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      ) : (
      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">Previous Requests</h2>

            <p className="text-sm text-muted-foreground">
              {filteredRequests.length} request
              {filteredRequests.length === 1 ? "" : "s"} shown
            </p>
          </div>

          <div className="flex flex-wrap items-center justify-end gap-2">
            <Select
              value={dateFilterMode}
              onValueChange={(value) =>
                setDateFilterMode(value as DateFilterMode)
              }
            >
              <SelectTrigger className="w-44">
                <CalendarDays className="mr-2 size-4 text-primary" />
                <SelectValue placeholder="All Dates" />
              </SelectTrigger>

              <SelectContent>
                <SelectItem value="all">All Dates</SelectItem>
                <SelectItem value="day">Specific Day</SelectItem>
                <SelectItem value="month">Specific Month</SelectItem>
                <SelectItem value="year">Specific Year</SelectItem>
              </SelectContent>
            </Select>

            {dateFilterMode === "day" && (
              <input
                type="date"
                value={selectedDay}
                onChange={(event) => setSelectedDay(event.target.value)}
                className="h-10 rounded-md border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
                aria-label="Filter requests by day"
              />
            )}

            {dateFilterMode === "month" && (
              <input
                type="month"
                value={selectedMonth}
                onChange={(event) => setSelectedMonth(event.target.value)}
                className="h-10 rounded-md border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
                aria-label="Filter requests by month"
              />
            )}

            {dateFilterMode === "year" && (
              <Select value={selectedYear} onValueChange={setSelectedYear}>
                <SelectTrigger className="w-32">
                  <SelectValue placeholder="Year" />
                </SelectTrigger>

                <SelectContent>
                  {availableYears.map((year) => (
                    <SelectItem key={year} value={String(year)}>
                      {year}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}

            <Select
              value={statusFilter}
              onValueChange={(value) =>
                setStatusFilter(value as "all" | RequestStatus)
              }
            >
              <SelectTrigger className="w-44">
                <SelectValue placeholder="All Status" />
              </SelectTrigger>

              <SelectContent>
                <SelectItem value="all">All Status</SelectItem>
                <SelectItem value="pending">Pending</SelectItem>
                <SelectItem value="in_progress">In Progress</SelectItem>
                <SelectItem value="completed">Completed</SelectItem>
                <SelectItem value="rejected">Rejected</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full min-w-[1100px] text-sm">
            <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
              <tr>
                <th className="px-4 py-3 text-left">Class Adviser</th>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Term</th>
                <th className="px-4 py-3 text-left">Priority</th>
                <th className="px-4 py-3 text-left">Message</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Date</th>
                <th className="px-4 py-3 text-right">Action</th>
              </tr>
            </thead>

            <tbody>
              {isLoading && (
                <tr>
                  <td
                    colSpan={8}
                    className="px-4 py-16 text-center text-muted-foreground"
                  >
                    <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                    Loading previous requests...
                  </td>
                </tr>
              )}

              {!isLoading &&
                filteredRequests.map((request) => (
                  <tr key={request.id} className="border-t">
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2">
                        <div className="grid size-8 place-items-center rounded-lg bg-primary/10 text-primary">
                          <UserRound className="size-4" />
                        </div>

                        <div className="font-medium">
                          {request.subject_teacher_id
                            ? classAdviserById.get(
                                request.subject_teacher_id,
                              ) || "Class Adviser"
                            : "Not assigned"}
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3 font-medium">{request.subject}</td>

                    <td className="px-4 py-3">{formatTerm(request.term)}</td>

                    <td className="px-4 py-3">
                      <span
                        className={`inline-flex rounded-lg border px-2.5 py-1 text-xs font-medium capitalize ${
                          priorityStyle[request.priority]
                        }`}
                      >
                        {request.priority}
                      </span>
                    </td>

                    <td className="max-w-72 px-4 py-3">
                      <p className="line-clamp-2">{request.message}</p>
                    </td>

                    <td className="px-4 py-3">
                      <span
                        className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1 text-xs font-medium ${
                          statusStyle[request.status]
                        }`}
                      >
                        <StatusIcon status={request.status} />
                        {statusLabel[request.status]}
                      </span>
                    </td>

                    <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                      <div className="flex items-start gap-2">
                        <CalendarDays className="mt-0.5 size-4 shrink-0 text-primary" />

                        <div>
                          <div>{formatDate(request.created_at)}</div>
                          <div className="mt-0.5 text-xs text-muted-foreground">
                            {formatRequestAge(request.created_at)}
                          </div>
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => setSelectedRequest(request)}
                          className="text-primary"
                        >
                          <Eye className="mr-1 size-4" />
                          View
                        </Button>

                        <Button
                          variant="outline"
                          size="icon"
                          onClick={() => setDeleteTarget(request)}
                          className="border-destructive/40 text-destructive hover:bg-destructive/10"
                          title="Permanently delete request"
                          aria-label="Permanently delete request"
                        >
                          <Trash2 className="size-4" />
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))}

              {!isLoading && filteredRequests.length === 0 && !error && (
                <tr>
                  <td
                    colSpan={8}
                    className="px-4 py-20 text-center text-muted-foreground"
                  >
                    No previous requests found for this status.
                  </td>
                </tr>
              )}

              {error && (
                <tr>
                  <td
                    colSpan={8}
                    className="px-4 py-16 text-center text-destructive"
                  >
                    {(error as Error).message}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      )}

      <Dialog
        open={Boolean(recipientMessageDialog)}
        onOpenChange={(open) => {
          if (!open) setRecipientMessageDialog(null);
        }}
      >
        <DialogContent className="w-[calc(100vw-2rem)] max-w-xl overflow-hidden p-0 sm:max-w-xl">
          <div className="min-w-0 px-6 pt-6">
            <DialogHeader className="min-w-0 text-left">
              <DialogTitle className="flex min-w-0 items-center gap-2">
                <MessageSquareText className="size-5 shrink-0 text-primary" />
                <span className="truncate">Recipient Message</span>
              </DialogTitle>
              <DialogDescription className="pr-6">
                Full message sent by this request recipient.
              </DialogDescription>
            </DialogHeader>
          </div>

          {recipientMessageDialog && (
            <div className="min-w-0 space-y-5 px-6 pb-6">
              <div className="grid min-w-0 grid-cols-1 gap-x-6 gap-y-4 rounded-xl border bg-muted/20 p-4 text-sm sm:grid-cols-2">
                <div className="min-w-0">
                  <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    Recipient
                  </div>
                  <div className="mt-1 break-words font-medium leading-5">
                    {recipientMessageDialog.recipientName}
                  </div>
                </div>

                <div className="min-w-0">
                  <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    Status
                  </div>
                  <div className="mt-1 break-words font-medium leading-5">
                    {recipientMessageDialog.status}
                  </div>
                </div>

                <div className="min-w-0">
                  <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    Subject
                  </div>
                  <div className="mt-1 break-words font-medium leading-5">
                    {recipientMessageDialog.subject}
                  </div>
                </div>

                <div className="min-w-0">
                  <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    Class
                  </div>
                  <div className="mt-1 break-words font-medium leading-5">
                    {recipientMessageDialog.classLabel}
                  </div>
                </div>
              </div>

              <div className="min-w-0">
                <div className="mb-2 text-sm font-semibold">Message</div>
                <div className="max-h-64 w-full min-w-0 max-w-full overflow-y-auto overflow-x-hidden whitespace-pre-wrap break-all rounded-xl border bg-background p-4 text-sm leading-6 [overflow-wrap:anywhere]">
                  {recipientMessageDialog.message}
                </div>
              </div>
            </div>
          )}

          <DialogFooter className="m-0 border-t bg-muted/10 px-6 py-4 sm:justify-end">
            <Button
              type="button"
              className="min-w-20"
              onClick={() => setRecipientMessageDialog(null)}
            >
              Close
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(selectedRequest)}
        onOpenChange={(open) => {
          if (!open) setSelectedRequest(null);
        }}
      >
        <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-5xl">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-xl">
              <Inbox className="size-6 text-primary" />
              Previous Request Details
            </DialogTitle>
          </DialogHeader>

          {selectedRequest && (
            <PreviousRequestDetails
              request={selectedRequest}
              classAdviserName={
                selectedRequest.subject_teacher_id
                  ? classAdviserById.get(selectedRequest.subject_teacher_id) ||
                    "Class Adviser"
                  : "Not assigned"
              }
              studentSexByName={studentSexByName}
              onOpenAttachment={() => openAttachment(selectedRequest)}
            />
          )}
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(deleteTarget)}
        onOpenChange={(open) => {
          if (!open && !deleteRequestMutation.isPending) {
            setDeleteTarget(null);
          }
        }}
      >
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-destructive">
              <Trash2 className="size-5" />
              Permanently Delete Request?
            </DialogTitle>

            <DialogDescription>
              This permanently removes the request and its attachment. This
              action cannot be undone.
            </DialogDescription>
          </DialogHeader>

          {deleteTarget && (
            <div className="rounded-xl border bg-muted/30 p-4 text-sm">
              <div className="font-medium">
                {deleteTarget.subject} · {formatTerm(deleteTarget.term)}
              </div>

              <div className="mt-1 text-muted-foreground">
                Sent {formatRequestAge(deleteTarget.created_at).toLowerCase()}
              </div>
            </div>
          )}

          <DialogFooter>
            <Button
              variant="outline"
              onClick={() => setDeleteTarget(null)}
              disabled={deleteRequestMutation.isPending}
            >
              Cancel
            </Button>

            <Button
              variant="destructive"
              onClick={() => {
                if (deleteTarget) {
                  deleteRequestMutation.mutate(deleteTarget);
                }
              }}
              disabled={!deleteTarget || deleteRequestMutation.isPending}
            >
              {deleteRequestMutation.isPending ? (
                <LoaderCircle className="mr-2 size-4 animate-spin" />
              ) : (
                <Trash2 className="mr-2 size-4" />
              )}

              {deleteRequestMutation.isPending
                ? "Deleting..."
                : "Delete Permanently"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(sentWorkflowReviewTarget)}
        onOpenChange={(open) => {
          if (!open) setSentWorkflowReviewTarget(null);
        }}
      >
        <DialogContent className="max-h-[94vh] overflow-y-auto sm:max-w-5xl">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <BookOpen className="size-5 text-primary" />
              View {sentWorkflowReviewTarget?.subject.subject ||
                "Submitted"}{" "}
              Grades
            </DialogTitle>
            <DialogDescription>
              View the learner grades submitted in response to your Grade
              Request.
            </DialogDescription>
          </DialogHeader>

          {sentWorkflowReviewTarget && (
            <div className="space-y-4">
              <div className="grid gap-x-6 gap-y-5 rounded-2xl border bg-muted/20 p-6 sm:grid-cols-2 lg:grid-cols-4">
                <DetailItem
                  label="Class Adviser"
                  value={
                    sentWorkflowReviewTarget.subject.subject_teacher_name ||
                    sentWorkflowReviewTarget.subject.subject_teacher_email ||
                    "Teacher"
                  }
                />
                <DetailItem
                  label="Class"
                  value={`${sentWorkflowReviewTarget.batch.gradeLevel} · ${sentWorkflowReviewTarget.batch.section}`}
                />
                <DetailItem
                  label="Subject"
                  value={sentWorkflowReviewTarget.subject.subject}
                />
                <DetailItem
                  label="Period"
                  value={formatTerm(
                    sentWorkflowReviewTarget.batch.gradingPeriod,
                  )}
                />
                <DetailItem label="Request Status" value="Submitted" />
                <DetailItem
                  label="Requested"
                  value={formatDate(
                    sentWorkflowReviewTarget.subject.subject_requested_at,
                  )}
                />
                <DetailItem
                  label="Submitted"
                  value={
                    sentWorkflowReviewTarget.subject.submitted_at
                      ? formatDate(
                          sentWorkflowReviewTarget.subject.submitted_at,
                        )
                      : "—"
                  }
                />
                <DetailItem
                  label="Overall Status"
                  value={
                    sentWorkflowReviewTarget.batch.overallStatus === "completed"
                      ? "Completed"
                      : "Pending"
                  }
                />
              </div>

              <div className="rounded-xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900">
                <div className="font-semibold">Submission Note</div>
                <p className="mt-1 whitespace-pre-wrap">
                  {sentWorkflowReviewTarget.subject.teacher_note?.trim() ||
                    "No submission note."}
                </p>
              </div>

              <div className="overflow-hidden rounded-2xl border">
                <div className="border-b px-4 py-4">
                  <h3 className="font-semibold">Learner Grades</h3>
                  <p className="mt-1 text-sm text-muted-foreground">
                    These are the grades submitted for{" "}
                    {sentWorkflowReviewTarget.subject.subject}.
                  </p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full min-w-[620px] text-sm">
                    <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
                      <tr>
                        <th className="px-4 py-3 text-left">#</th>
                        <th className="px-4 py-3 text-left">Learner</th>
                        <th className="px-4 py-3 text-left">LRN</th>
                        <th className="px-4 py-3 text-right">Grade</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(reviewStudentsLoading || reviewScoresLoading) && (
                        <tr>
                          <td
                            colSpan={4}
                            className="px-4 py-12 text-center text-muted-foreground"
                          >
                            <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                            Loading submitted grades...
                          </td>
                        </tr>
                      )}

                      {!reviewStudentsLoading &&
                        !reviewScoresLoading &&
                        workflowReviewError && (
                          <tr>
                            <td
                              colSpan={4}
                              className="px-4 py-12 text-center text-destructive"
                            >
                              {(workflowReviewError as Error).message}
                            </td>
                          </tr>
                        )}

                      {!reviewStudentsLoading &&
                        !reviewScoresLoading &&
                        workflowReviewStudentGroups.map((group) =>
                          group.learners.length > 0 ? (
                            <Fragment key={group.label}>
                              <tr className="border-t bg-primary/10">
                                <td
                                  colSpan={4}
                                  className="px-4 py-2 font-bold text-primary"
                                >
                                  {group.label} ({group.learners.length})
                                </td>
                              </tr>
                              {group.learners.map((student, index) => (
                                <tr key={student.id} className="border-t">
                                  <td className="px-4 py-3 text-muted-foreground">
                                    {index + 1}
                                  </td>
                                  <td className="px-4 py-3 font-medium">
                                    {student.last_name}, {student.first_name}{" "}
                                    {student.middle_name || ""}
                                  </td>
                                  <td className="px-4 py-3 text-muted-foreground">
                                    {student.lrn || "—"}
                                  </td>
                                  <td className="px-4 py-2 text-right">
                                    <div className="ml-auto w-36 rounded-xl border bg-background px-3 py-2 text-left font-medium shadow-sm">
                                      {workflowReviewScoreByStudent.has(
                                        student.id,
                                      )
                                        ? workflowReviewScoreByStudent.get(
                                            student.id,
                                          )
                                        : "—"}
                                    </div>
                                  </td>
                                </tr>
                              ))}
                            </Fragment>
                          ) : null,
                        )}

                      {!reviewStudentsLoading &&
                        !reviewScoresLoading &&
                        !workflowReviewError &&
                        workflowReviewStudents.length === 0 && (
                          <tr>
                            <td
                              colSpan={4}
                              className="px-4 py-12 text-center text-muted-foreground"
                            >
                              No submitted learner grades were found for this
                              request.
                            </td>
                          </tr>
                        )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          <DialogFooter>
            <Button
              type="button"
              onClick={() => setSentWorkflowReviewTarget(null)}
            >
              Close
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(sentWorkflowDeleteTarget)}
        onOpenChange={(open) => {
          if (!open && !deleteSentWorkflowMutation.isPending) {
            setSentWorkflowDeleteTarget(null);
          }
        }}
      >
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-destructive">
              <Trash2 className="size-5" />
              Delete Sent Grade Request?
            </DialogTitle>
            <DialogDescription>
              Delete only this selected sent Grade Request. If it was already
              submitted, its saved grades will also be deleted. This action
              cannot be undone.
            </DialogDescription>
          </DialogHeader>

          {sentWorkflowDeleteTarget && (
            <div className="rounded-xl border bg-muted/30 p-4 text-sm">
              <div className="font-medium">
                {sentWorkflowDeleteTarget.subject.subject} ·{" "}
                {formatTerm(sentWorkflowDeleteTarget.batch.gradingPeriod)}
              </div>
              <div className="mt-1 text-muted-foreground">
                Grade {sentWorkflowDeleteTarget.batch.gradeLevel} ·{" "}
                {sentWorkflowDeleteTarget.batch.section}
              </div>
              <div className="mt-1 text-muted-foreground">
                Sent to{" "}
                {sentWorkflowDeleteTarget.subject.subject_teacher_name ||
                  sentWorkflowDeleteTarget.subject.subject_teacher_email ||
                  (sentWorkflowDeleteTarget.subject.recipient_role ===
                  "class_adviser"
                    ? "Class Adviser"
                    : "Subject Teacher")}
              </div>
            </div>
          )}

          <DialogFooter>
            <Button
              variant="outline"
              onClick={() => setSentWorkflowDeleteTarget(null)}
              disabled={deleteSentWorkflowMutation.isPending}
            >
              Cancel
            </Button>
            <Button
              variant="destructive"
              onClick={() => {
                if (sentWorkflowDeleteTarget) {
                  deleteSentWorkflowMutation.mutate(
                    sentWorkflowDeleteTarget.subject.subject_request_id,
                  );
                }
              }}
              disabled={
                !sentWorkflowDeleteTarget ||
                deleteSentWorkflowMutation.isPending
              }
            >
              {deleteSentWorkflowMutation.isPending ? (
                <LoaderCircle className="mr-2 size-4 animate-spin" />
              ) : (
                <Trash2 className="mr-2 size-4" />
              )}
              {deleteSentWorkflowMutation.isPending
                ? "Deleting..."
                : "Delete Request"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function PreviousRequestDetails({
  request,
  classAdviserName,
  studentSexByName,
  onOpenAttachment,
}: {
  request: RequestRow;
  classAdviserName: string;
  studentSexByName: Map<string, LearnerSex>;
  onOpenAttachment: () => void;
}) {
  const parsed = parseGradeRequestMessage(request.message, studentSexByName);

  const maleLearners = parsed.learners.filter(
    (learner) => learner.sex === "male",
  );

  const femaleLearners = parsed.learners.filter(
    (learner) => learner.sex === "female",
  );

  const unidentifiedLearners = parsed.learners.filter(
    (learner) => learner.sex === "unknown",
  );

  return (
    <div className="space-y-5">
      <section className="grid gap-3 rounded-2xl border bg-muted/20 p-4 sm:grid-cols-2">
        <MetadataCard
          icon={UserRound}
          label="Class Adviser"
          value={classAdviserName}
          helper="Request recipient"
          iconTone="bg-rose-50 text-rose-700"
        />

        <MetadataCard
          icon={CalendarDays}
          label="Date Requested"
          value={formatDate(request.created_at)}
          iconTone="bg-red-50 text-red-600"
        />

        <MetadataCard
          icon={BookOpen}
          label="Subject"
          value={request.subject}
          iconTone="bg-blue-50 text-blue-700"
        />

        <MetadataCard
          icon={GraduationCap}
          label="Term"
          value={formatTerm(request.term)}
          iconTone="bg-sky-50 text-sky-700"
        />

        <MetadataCard
          icon={Flag}
          label="Priority"
          value={capitalize(request.priority)}
          badgeClass={priorityStyle[request.priority]}
          iconTone="bg-amber-50 text-amber-600"
        />

        <MetadataCard
          icon={Clock3}
          label="Current Status"
          value={statusLabel[request.status]}
          badgeClass={statusStyle[request.status]}
          iconTone="bg-orange-50 text-orange-600"
        />
      </section>

      <section className="overflow-hidden rounded-2xl border bg-background">
        <div className="border-b px-4 py-3">
          <div className="flex items-center gap-2 text-sm font-semibold uppercase tracking-wide text-primary">
            <Inbox className="size-4" />
            Request Message
          </div>

          <p className="mt-3 font-medium">{parsed.heading}</p>

          {parsed.classDetails && (
            <p className="mt-1 text-sm text-muted-foreground">
              {parsed.classDetails}
            </p>
          )}

          <div className="mt-3 inline-flex items-center gap-2 rounded-lg bg-muted px-3 py-1.5 text-sm font-medium">
            <Users className="size-4 text-primary" />
            Learners ({parsed.learners.length})
          </div>
        </div>

        {parsed.learners.length > 0 ? (
          <div className="space-y-4 p-4">
            <div className="grid gap-4 xl:grid-cols-2">
              <LearnerGroupTable
                title="MALE"
                learners={maleLearners}
                headerClassName="border-blue-200 bg-blue-50 text-blue-700"
                bodyClassName="border-blue-100 bg-blue-50/20"
              />

              <LearnerGroupTable
                title="FEMALE"
                learners={femaleLearners}
                headerClassName="border-pink-200 bg-pink-50 text-pink-700"
                bodyClassName="border-pink-100 bg-pink-50/20"
              />
            </div>

            {unidentifiedLearners.length > 0 && (
              <LearnerGroupTable
                title="UNIDENTIFIED"
                learners={unidentifiedLearners}
                headerClassName="border-slate-200 bg-slate-50 text-slate-700"
                bodyClassName="border-slate-100 bg-slate-50/20"
              />
            )}
          </div>
        ) : (
          <div className="whitespace-pre-wrap px-4 py-4 text-sm leading-relaxed">
            {request.message}
          </div>
        )}

        {parsed.learners.length > 0 && (
          <div className="flex flex-wrap gap-4 border-t bg-muted/20 px-4 py-3 text-xs">
            <span className="inline-flex items-center gap-1 font-medium text-emerald-600">
              <Trophy className="size-3.5" />
              Promoted
            </span>

            <span className="inline-flex items-center gap-1 font-medium text-amber-600">
              <Trophy className="size-3.5" />
              Promoted with Honors
            </span>
          </div>
        )}
      </section>

      {request.teacher_response && (
        <section>
          <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            Class Adviser Response
          </div>

          <div className="mt-1.5 rounded-xl border border-emerald-200 bg-emerald-50 p-4 leading-relaxed text-emerald-900">
            {request.teacher_response}
          </div>
        </section>
      )}

      {request.attachment_url && (
        <Button variant="outline" onClick={onOpenAttachment}>
          <Download className="mr-2 size-4" />
          Open {request.attachment_name || "Attachment"}
        </Button>
      )}
    </div>
  );
}

function LearnerGroupTable({
  title,
  learners,
  headerClassName,
  bodyClassName,
}: {
  title: string;
  learners: ParsedLearner[];
  headerClassName: string;
  bodyClassName: string;
}) {
  return (
    <section className={`overflow-hidden rounded-2xl border ${bodyClassName}`}>
      <div
        className={`flex items-center gap-2 border-b px-4 py-3 ${headerClassName}`}
      >
        <UserRound className="size-4" />
        <span className="text-sm font-semibold">
          {title} ({learners.length})
        </span>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[700px] text-xs">
          <thead className="bg-background/80 uppercase text-muted-foreground">
            <tr>
              <th className="px-2 py-2 text-left">#</th>
              <th className="px-2 py-2 text-left">Learner's Name</th>
              <th className="px-2 py-2 text-center">Grade</th>
              <th className="px-2 py-2 text-center">Avg</th>
              <th className="px-2 py-2 text-center">Rank</th>
              <th className="px-2 py-2 text-center">Descriptor</th>
              <th className="px-2 py-2 text-center">Award</th>
            </tr>
          </thead>

          <tbody className="bg-background">
            {learners.map((learner, index) => (
              <tr
                key={`${title}-${learner.number}-${learner.name}`}
                className="border-t"
              >
                <td className="px-2 py-2 text-muted-foreground">{index + 1}</td>

                <td className="whitespace-nowrap px-2 py-2 font-medium">
                  {learner.name}
                </td>

                <td className="px-2 py-2 text-center font-semibold">
                  {learner.gradeText}
                </td>

                <td className="px-2 py-2 text-center font-semibold">
                  {learner.grade ?? "—"}
                </td>

                <td className="px-2 py-2 text-center">{learner.rank ?? "—"}</td>

                <td className="px-2 py-2 text-center">{learner.descriptor}</td>

                <td className="px-2 py-2 text-center">
                  {learner.award ? (
                    <span
                      className={`inline-flex items-center gap-1 font-medium ${learner.award.tone}`}
                    >
                      <Trophy className="size-3.5" />
                      {learner.award.label}
                    </span>
                  ) : (
                    "—"
                  )}
                </td>
              </tr>
            ))}

            {learners.length === 0 && (
              <tr>
                <td
                  colSpan={7}
                  className="px-3 py-8 text-center text-muted-foreground"
                >
                  No {title.toLowerCase()} learners found.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function MetadataCard({
  icon: Icon,
  label,
  value,
  helper,
  iconTone,
  badgeClass,
}: {
  icon: typeof UserRound;
  label: string;
  value: string;
  helper?: string;
  iconTone: string;
  badgeClass?: string;
}) {
  return (
    <div className="flex items-start gap-3 rounded-xl bg-background/70 p-3">
      <div
        className={`grid size-10 shrink-0 place-items-center rounded-full ${iconTone}`}
      >
        <Icon className="size-5" />
      </div>

      <div className="min-w-0">
        <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
          {label}
        </div>

        {badgeClass ? (
          <span
            className={`mt-1 inline-flex rounded-lg border px-2.5 py-1 text-sm font-medium ${badgeClass}`}
          >
            {value}
          </span>
        ) : (
          <div className="mt-1 font-medium">{value}</div>
        )}

        {helper && (
          <div className="mt-0.5 text-xs text-muted-foreground">{helper}</div>
        )}
      </div>
    </div>
  );
}

function parseGradeRequestMessage(
  message: string,
  studentSexByName: Map<string, LearnerSex>,
): ParsedRequestMessage {
  const lines = message
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const heading = lines[0] || "Grade request";

  const studentHeaderIndex = lines.findIndex((line) =>
    /^students\s*\(\d+\)\s*:?\s*$/i.test(line),
  );

  const classDetails =
    studentHeaderIndex > 1
      ? lines.slice(1, studentHeaderIndex).join(" · ")
      : lines[1] || "";

  const learnerLines =
    studentHeaderIndex >= 0 ? lines.slice(studentHeaderIndex + 1) : [];

  const baseLearners = learnerLines
    .map((line) => {
      const match = line.match(/^(\d+)\.\s*(.+?)\s+[—-]\s+[^:]+:\s*(.+)$/u);

      if (!match) return null;

      const gradeText = match[3].trim();
      const numericGrade = Number(gradeText);
      const grade = Number.isFinite(numericGrade) ? numericGrade : null;

      return {
        number: Number(match[1]),
        name: match[2].trim(),
        grade,
        gradeText: grade ?? "Not yet submitted",
      };
    })
    .filter(
      (
        learner,
      ): learner is {
        number: number;
        name: string;
        grade: number | null;
        gradeText: string | number;
      } => learner !== null,
    );

  const rankedGrades = [
    ...new Set(
      baseLearners
        .map((learner) => learner.grade)
        .filter((grade): grade is number => grade !== null),
    ),
  ].sort((a, b) => b - a);

  const learners: ParsedLearner[] = baseLearners.map((learner) => ({
    number: learner.number,
    name: learner.name,
    sex: studentSexByName.get(normalizeLearnerName(learner.name)) ?? "unknown",
    grade: learner.grade,
    gradeText: String(learner.gradeText),
    rank:
      learner.grade === null ? null : rankedGrades.indexOf(learner.grade) + 1,
    descriptor: gradeDescriptor(learner.grade),
    award: gradeAward(learner.grade),
  }));

  return {
    heading,
    classDetails,
    learners,
  };
}

function normalizeLearnerName(value: string) {
  return value
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function gradeDescriptor(grade: number | null) {
  if (grade === null) return "—";
  if (grade >= 90) return "ADVANCING";
  if (grade >= 85) return "PROFICIENT";
  if (grade >= 80) return "APPROACHING";
  if (grade >= 75) return "DEVELOPING";
  return "BEGINNING";
}

function gradeAward(grade: number | null) {
  if (grade === null) return null;

  if (grade >= 98) {
    return {
      label: "PROMOTED WITH HIGH HONORS",
      tone: "text-amber-600",
    };
  }

  if (grade >= 90) {
    return {
      label: "PROMOTED WITH HONORS",
      tone: "text-amber-600",
    };
  }

  if (grade >= 75) {
    return {
      label: "PROMOTED",
      tone: "text-emerald-600",
    };
  }

  return null;
}

function CountCard({
  label,
  value,
  active,
  onClick,
}: {
  label: string;
  value: number;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`rounded-2xl border p-4 text-left transition ${
        active
          ? "border-primary bg-primary text-primary-foreground shadow-md ring-2 ring-primary/20"
          : "bg-card hover:border-primary/40 hover:shadow-sm"
      }`}
    >
      <div
        className={`text-sm ${active ? "text-primary-foreground/80" : "text-muted-foreground"}`}
      >
        {label}
      </div>

      <div className="mt-1 text-2xl font-semibold">{value}</div>
    </button>
  );
}

function DetailItem({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
        {label}
      </div>

      <div className="mt-1 font-medium">{value}</div>
    </div>
  );
}

function advisoryClassKey(classRow: AdvisoryClassRow) {
  return [
    classRow.grade_level ?? "",
    classRow.subject ?? "",
    classRow.section ?? "",
    classRow.school_year ?? "",
  ]
    .map((value) => value.trim().toLowerCase())
    .join("|");
}

function dedupeAdvisoryClasses(classRows: AdvisoryClassRow[]) {
  const uniqueClasses = new Map<string, AdvisoryClassRow>();

  classRows.forEach((classRow) => {
    const key = advisoryClassKey(classRow);
    if (!uniqueClasses.has(key)) uniqueClasses.set(key, classRow);
  });

  return Array.from(uniqueClasses.values());
}

function formatGradeLevel(value: string | null) {
  const normalized = String(value ?? "").trim();

  if (!normalized) return "Grade —";
  if (/^grade\s+/i.test(normalized)) return normalized;

  return `Grade ${normalized}`;
}

function formatAdvisoryClassLabel(classRow: AdvisoryClassRow) {
  return [
    formatGradeLevel(classRow.grade_level),
    classRow.subject?.trim() || "No subject",
    classRow.section?.trim() || "No section",
    classRow.school_year?.trim() || "No school year",
  ].join(" · ");
}

function capitalize(value: string) {
  if (!value) return "—";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function formatTerm(value: string) {
  const normalized = value.trim().toLowerCase();

  if (normalized === "final") {
    return "Final Grade";
  }

  if (normalized.startsWith("term")) {
    return value;
  }

  return `Term ${value}`;
}

function toDateInputValue(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function toMonthInputValue(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}`;
}

function formatRequestAge(value: string) {
  const requestDate = new Date(value);

  if (Number.isNaN(requestDate.getTime())) return "Unknown date";

  const now = new Date();
  const difference = Math.max(0, now.getTime() - requestDate.getTime());
  const days = Math.floor(difference / 86_400_000);

  if (days === 0) return "Today";
  if (days === 1) return "1 day ago";
  if (days < 31) return `${days} days ago`;

  const months = Math.floor(days / 30);

  if (months === 1) return "1 month ago";
  if (months < 12) return `${months} months ago`;

  const years = Math.floor(months / 12);

  return years === 1 ? "1 year ago" : `${years} years ago`;
}

function formatDate(value: string) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) return "—";

  return date.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}