import { createFileRoute, Link } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Fragment, useEffect, useMemo, useState } from "react";
import {
  ArrowLeft,
  BookOpen,
  Calendar,
  CalendarDays,
  CheckCircle2,
  Clock3,
  Download,
  Eye,
  Flag,
  GraduationCap,
  History,
  Inbox,
  LoaderCircle,
  MessageSquareText,
  RefreshCw,
  Save,
  Send,
  Trash2,
  Trophy,
  UserRound,
  Users,
  XCircle,
} from "lucide-react";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
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
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

export const Route = createFileRoute("/_authenticated/received-requests")({
  component: ReceivedRequestsPage,
});

type TeacherType = "class_adviser" | "subject_teacher";
type RequestStatus = "pending" | "in_progress" | "completed" | "rejected";
type Priority = "low" | "medium" | "high";
type DateFilter = "all" | "today" | "last_7_days" | "this_month" | "this_year";

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

type RequesterProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type?: TeacherType | null;
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

type WorkflowFilterStatus = "all" | "pending" | "completed" | "rejected";
type WorkflowSubjectStatus = "pending" | "completed" | "rejected";
type WorkflowDecision = "complete" | "reject";
type GradingPeriod = "1" | "2" | "3" | "final";

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
  overall_status: "pending" | "completed";
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
  recipient_role: TeacherType;
  recipient_id: string | null;
};

type WorkflowStudentRow = {
  id: string;
  class_id: string;
  first_name: string;
  middle_name: string | null;
  last_name: string;
  lrn: string | null;
  sex: LearnerSex;
};

type WorkflowGradeRow = {
  student_id: string;
  class_id: string;
  subject: string;
  term: string;
  score: number | null;
};

type WorkflowGradeSourceClassRow = {
  id: string;
  teacher_id: string;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
};

type WorkflowScoreRow = {
  student_id: string;
  score: number;
};

type WorkflowHistoryRow = {
  id: string;
  batch_id: string;
  subject_request_id: string | null;
  actor_id: string | null;
  event_type: string;
  old_status: string | null;
  new_status: string | null;
  details: Record<string, unknown>;
  created_at: string;
};

const statusLabel: Record<RequestStatus, string> = {
  pending: "Pending Review",
  in_progress: "Reviewing",
  completed: "Accepted",
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

function SkeletonPageHeader({ showBack = true }: { showBack?: boolean }) {
  return (
    <section className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card px-5 py-4 shadow-sm">
      <div className="flex items-center gap-4">
        {showBack && <Skeleton className="h-10 w-20 rounded-xl" />}
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <Skeleton className="size-6 rounded-lg" />
            <Skeleton className="h-6 w-56 max-w-[58vw]" />
          </div>
          <Skeleton className="h-4 w-[32rem] max-w-[68vw]" />
        </div>
      </div>
      <Skeleton className="h-10 w-28 rounded-xl" />
    </section>
  );
}

function SkeletonCountCards({ count }: { count: number }) {
  return (
    <>
      {Array.from({ length: count }, (_, index) => (
        <div key={index} className="rounded-xl border bg-card p-4 shadow-sm">
          <Skeleton className="h-3 w-24" />
          <Skeleton className="mt-3 h-7 w-10" />
        </div>
      ))}
    </>
  );
}

function WorkflowInboxSkeleton({ embedded = false }: { embedded?: boolean }) {
  const rows = Array.from({ length: 4 }, (_, index) => index);

  return (
    <div
      className={embedded ? "space-y-5" : "space-y-5 pb-24 md:pb-6"}
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading received grade requests...</span>
      <SkeletonPageHeader showBack={!embedded} />

      <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
        <SkeletonCountCards count={3} />
      </section>

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div className="space-y-2">
            <Skeleton className="h-5 w-52" />
            <Skeleton className="h-4 w-28" />
          </div>
          <div className="flex flex-wrap gap-2">
            <Skeleton className="h-10 w-44 rounded-md" />
            <Skeleton className="h-10 w-44 rounded-md" />
          </div>
        </div>

        <div className="overflow-x-auto">
          <div className="min-w-[1120px]">
            <div className="grid grid-cols-[1.1fr_1fr_0.7fr_0.6fr_0.7fr_0.5fr_0.8fr_0.8fr_0.7fr] gap-4 bg-muted/40 px-4 py-3">
              {Array.from({ length: 9 }, (_, index) => (
                <Skeleton key={index} className="h-3 w-20" />
              ))}
            </div>

            {rows.map((row) => (
              <div
                key={row}
                className="grid grid-cols-[1.1fr_1fr_0.7fr_0.6fr_0.7fr_0.5fr_0.8fr_0.8fr_0.7fr] items-center gap-4 border-t px-4 py-3"
              >
                <div className="flex items-center gap-2">
                  <Skeleton className="size-8 rounded-lg" />
                  <div className="space-y-1.5">
                    <Skeleton className="h-4 w-24" />
                    <Skeleton className="h-3 w-28" />
                  </div>
                </div>
                <div className="space-y-1.5">
                  <Skeleton className="h-4 w-20" />
                  <Skeleton className="h-3 w-28" />
                </div>
                <Skeleton className="h-4 w-20" />
                <Skeleton className="h-4 w-14" />
                <Skeleton className="h-7 w-20 rounded-lg" />
                <Skeleton className="h-4 w-8" />
                <Skeleton className="h-4 w-24" />
                <Skeleton className="h-4 w-24" />
                <Skeleton className="ml-auto h-9 w-28 rounded-lg" />
              </div>
            ))}
          </div>
        </div>
      </section>
    </div>
  );
}

function ReceivedRequestsSkeleton() {
  const rows = Array.from({ length: 5 }, (_, index) => index);

  return (
    <div
      className="space-y-5 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading Receive Form...</span>
      <SkeletonPageHeader />
      <WorkflowInboxSkeleton embedded />

      <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
        <SkeletonCountCards count={5} />
      </section>

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div className="space-y-2">
            <Skeleton className="h-5 w-64" />
            <Skeleton className="h-4 w-28" />
          </div>
          <div className="flex flex-wrap gap-2">
            <Skeleton className="h-10 w-44 rounded-md" />
            <Skeleton className="h-10 w-44 rounded-md" />
          </div>
        </div>

        <div className="overflow-x-auto">
          <div className="min-w-[1000px]">
            <div className="grid grid-cols-[1.1fr_0.7fr_0.5fr_0.6fr_1.4fr_0.7fr_0.8fr_0.8fr] gap-4 bg-muted/40 px-4 py-3">
              {Array.from({ length: 8 }, (_, index) => (
                <Skeleton key={index} className="h-3 w-20" />
              ))}
            </div>

            {rows.map((row) => (
              <div
                key={row}
                className="grid grid-cols-[1.1fr_0.7fr_0.5fr_0.6fr_1.4fr_0.7fr_0.8fr_0.8fr] items-center gap-4 border-t px-4 py-3"
              >
                <div className="flex items-center gap-2">
                  <Skeleton className="size-8 rounded-lg" />
                  <div className="space-y-1.5">
                    <Skeleton className="h-4 w-24" />
                    <Skeleton className="h-3 w-16" />
                  </div>
                </div>
                <Skeleton className="h-4 w-20" />
                <Skeleton className="h-4 w-12" />
                <Skeleton className="h-7 w-16 rounded-lg" />
                <div className="space-y-1.5">
                  <Skeleton className="h-3 w-full" />
                  <Skeleton className="h-3 w-3/4" />
                </div>
                <Skeleton className="h-7 w-20 rounded-lg" />
                <div className="space-y-1.5">
                  <Skeleton className="h-4 w-24" />
                  <Skeleton className="h-3 w-16" />
                </div>
                <div className="flex justify-end gap-2">
                  <Skeleton className="h-9 w-28 rounded-lg" />
                  <Skeleton className="size-9 rounded-lg" />
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>
    </div>
  );
}

function ReceivedRequestsPage() {
  const queryClient = useQueryClient();

  const [statusFilter, setStatusFilter] = useState<"all" | RequestStatus>("all");
  const [dateFilter, setDateFilter] = useState<DateFilter>("all");
  const [selectedRequest, setSelectedRequest] = useState<RequestRow | null>(null);
  const [requestToDelete, setRequestToDelete] = useState<RequestRow | null>(null);
  const [responseStatus, setResponseStatus] = useState<RequestStatus>("pending");
  const [teacherResponse, setTeacherResponse] = useState("");

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

        if (profileTeacherType === "class_adviser" || profileTeacherType === "subject_teacher") {
          return profileTeacherType;
        }

        if (metadataTeacherType === "class_adviser" || metadataTeacherType === "subject_teacher") {
          return metadataTeacherType;
        }

        return "class_adviser";
      },
      staleTime: 5 * 60 * 1000,
    });

  const dashboardPath = "/dashboard";

  const {
    data: requestsData,
    isLoading,
    isPlaceholderData: requestsPlaceholder,
    error,
    refetch,
    isFetching,
  } = useQuery<RequestRow[]>({
    enabled: !teacherTypeLoading && teacherType === "class_adviser",
    queryKey: ["received-grade-requests"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      // grade_requests may not yet exist in generated Supabase TypeScript types.
      const { data, error: requestError } = await (supabase as any)
        .from("grade_requests")
        .select("*")
        .eq("subject_teacher_id", user.id)
        .order("created_at", { ascending: false });

      if (requestError) throw requestError;

      return Array.isArray(data) ? (data as unknown as RequestRow[]) : [];
    },
    placeholderData: [],
  });

  const requests = Array.isArray(requestsData) ? requestsData : [];

  const requesterIds = useMemo(
    () => [...new Set(requests.map((request) => request.requester_id))],
    [requests],
  );

  const {
    data: requesterProfilesData,
    isLoading: requesterProfilesLoading,
    isPlaceholderData: requesterProfilesPlaceholder,
  } = useQuery<RequesterProfile[]>({
    queryKey: ["grade-request-requesters", requesterIds.join(",")],
    enabled: requesterIds.length > 0,
    queryFn: async () => {
      const { data, error: profileError } = await (supabase as any)
        .from("profiles")
        .select("id, full_name, email, teacher_type")
        .in("id", requesterIds);

      if (profileError) throw profileError;

      return Array.isArray(data) ? (data as unknown as RequesterProfile[]) : [];
    },
    placeholderData: [],
  });

  const requesterProfiles = Array.isArray(requesterProfilesData) ? requesterProfilesData : [];

  const requesterById = useMemo(
    () =>
      new Map(
        requesterProfiles.map((profile) => [
          profile.id,
          profile.full_name || profile.email || "Teacher",
        ]),
      ),
    [requesterProfiles],
  );

  const requesterProfileById = useMemo(
    () => new Map(requesterProfiles.map((profile) => [profile.id, profile])),
    [requesterProfiles],
  );

  const { data: studentSexRowsData } = useQuery<StudentSexRow[]>({
    queryKey: ["received-request-student-sex-directory"],
    queryFn: async () => {
      const { data, error: studentError } = await (supabase as any)
        .from("students")
        .select("first_name, last_name, sex");

      if (studentError) throw studentError;

      return Array.isArray(data) ? (data as unknown as StudentSexRow[]) : [];
    },
    placeholderData: [],
  });

  const studentSexRows = Array.isArray(studentSexRowsData) ? studentSexRowsData : [];

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
        normalizedSex === "male" ? "male" : normalizedSex === "female" ? "female" : "unknown",
      );
    });

    return directory;
  }, [studentSexRows]);

  const filteredRequests = useMemo(() => {
    return requests.filter((request) => {
      const matchesStatus = statusFilter === "all" || request.status === statusFilter;

      const matchesDate = matchesDateFilter(request.created_at, dateFilter);

      return matchesStatus && matchesDate;
    });
  }, [requests, statusFilter, dateFilter]);

  const counts = useMemo(
    () => ({
      all: requests.length,
      pending: requests.filter((request) => request.status === "pending").length,
      in_progress: requests.filter((request) => request.status === "in_progress").length,
      completed: requests.filter((request) => request.status === "completed").length,
      rejected: requests.filter((request) => request.status === "rejected").length,
    }),
    [requests],
  );

  const openRequest = (request: RequestRow) => {
    setSelectedRequest(request);
    setResponseStatus(request.status);
    setTeacherResponse(request.teacher_response ?? "");
  };

  const updateRequestMutation = useMutation({
    mutationFn: async () => {
      if (!selectedRequest) {
        throw new Error("No request selected.");
      }

      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      const { error: updateError } = await (supabase as any)
        .from("grade_requests")
        .update({
          status: responseStatus,
          teacher_response: teacherResponse.trim() || null,
        })
        .eq("id", selectedRequest.id)
        .eq("subject_teacher_id", user.id);

      if (updateError) throw updateError;
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["received-grade-requests"],
      });

      toast.success("Request updated successfully.");
      setSelectedRequest(null);
      setTeacherResponse("");
      setResponseStatus("pending");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const deleteRequestMutation = useMutation({
    mutationFn: async (request: RequestRow) => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      if (request.attachment_url) {
        const { error: storageError } = await supabase.storage
          .from("grade-request-attachments")
          .remove([request.attachment_url]);

        if (storageError) {
          throw storageError;
        }
      }

      const { error: deleteError } = await (supabase as any)
        .from("grade_requests")
        .delete()
        .eq("id", request.id)
        .eq("subject_teacher_id", user.id);

      if (deleteError) throw deleteError;
    },
    onSuccess: async () => {
      await queryClient.invalidateQueries({
        queryKey: ["received-grade-requests"],
      });

      toast.success("Request permanently deleted.");
      setRequestToDelete(null);
      setSelectedRequest(null);
    },
    onError: (deleteError: Error) => {
      toast.error(deleteError.message);
    },
  });

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

  const receivedRequestsLoading =
    isLoading ||
    requestsPlaceholder ||
    (requesterIds.length > 0 &&
      (requesterProfilesLoading || requesterProfilesPlaceholder));

  if (teacherTypeLoading) return <ReceivedRequestsSkeleton />;

  if (teacherType === "subject_teacher") {
    return <SubjectTeacherGradeRequestInbox recipientRole="subject_teacher" />;
  }

  if (receivedRequestsLoading) return <ReceivedRequestsSkeleton />;

  return (
    <div className="space-y-5 pb-24 md:pb-6">
      <section className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card px-5 py-4 shadow-sm">
        <div className="flex items-center gap-4">
          <Link
            to={dashboardPath}
            className="inline-flex items-center gap-1.5 rounded-xl border px-3 py-2 text-sm font-medium hover:bg-muted"
          >
            <ArrowLeft className="size-4" />
            Back
          </Link>

          <div>
            <div className="flex items-center gap-2 text-xl font-semibold">
              <Inbox className="size-6 text-primary" />
              Grade Requests Received
            </div>

            <p className="mt-0.5 text-sm text-muted-foreground">
              Enter and submit grades for requests sent directly to your Class Adviser account.
            </p>
          </div>
        </div>

        <Button variant="outline" onClick={() => refetch()} disabled={isFetching}>
          <RefreshCw className={`mr-2 size-4 ${isFetching ? "animate-spin" : ""}`} />
          Refresh
        </Button>
      </section>

      <SubjectTeacherGradeRequestInbox recipientRole="class_adviser" embedded />

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

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">Submissions from Subject Teachers</h2>
            <p className="text-sm text-muted-foreground">
              {filteredRequests.length} request
              {filteredRequests.length === 1 ? "" : "s"} shown
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <Select
              value={dateFilter}
              onValueChange={(value) => setDateFilter(value as DateFilter)}
            >
              <SelectTrigger className="w-44">
                <Calendar className="mr-2 size-4 text-primary" />
                <SelectValue placeholder="All Dates" />
              </SelectTrigger>

              <SelectContent>
                <SelectItem value="all">All Dates</SelectItem>
                <SelectItem value="today">Today</SelectItem>
                <SelectItem value="last_7_days">Last 7 Days</SelectItem>
                <SelectItem value="this_month">This Month</SelectItem>
                <SelectItem value="this_year">This Year</SelectItem>
              </SelectContent>
            </Select>

            <Select
              value={statusFilter}
              onValueChange={(value) => setStatusFilter(value as "all" | RequestStatus)}
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
          <table className="w-full min-w-[1000px] text-sm">
            <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
              <tr>
                <th className="px-4 py-3 text-left">Requester</th>
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
                  <td colSpan={8} className="px-4 py-16 text-center text-muted-foreground">
                    <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                    Loading received requests...
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

                        <div>
                          <div className="font-medium">
                            {requesterById.get(request.requester_id) || "Teacher"}
                          </div>
                          <div className="text-xs text-muted-foreground">
                            {request.requester_id.slice(0, 8)}...
                          </div>
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3 font-medium">{request.subject}</td>

                    <td className="px-4 py-3">{request.term}</td>

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
                          <div className="text-xs">{formatRelativeDate(request.created_at)}</div>
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => openRequest(request)}
                          className="text-primary"
                        >
                          <Eye className="mr-1 size-4" />
                          View / Respond
                        </Button>

                        <Button
                          variant="outline"
                          size="icon"
                          onClick={() => setRequestToDelete(request)}
                          className="border-destructive/40 text-destructive hover:bg-destructive/10"
                          aria-label="Delete request permanently"
                          title="Delete request permanently"
                        >
                          <Trash2 className="size-4" />
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))}

              {!isLoading && filteredRequests.length === 0 && !error && (
                <tr>
                  <td colSpan={8} className="px-4 py-16 text-center text-muted-foreground">
                    No grade submissions received for this status.
                  </td>
                </tr>
              )}

              {error && (
                <tr>
                  <td colSpan={8} className="px-4 py-16 text-center text-destructive">
                    {(error as Error).message}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      <Dialog
        open={Boolean(selectedRequest)}
        onOpenChange={(open) => {
          if (!open) {
            setSelectedRequest(null);
            setTeacherResponse("");
            setResponseStatus("pending");
          }
        }}
      >
        <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-5xl">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-xl">
              <MessageSquareText className="size-6 text-primary" />
              Review Request
            </DialogTitle>
          </DialogHeader>

          {selectedRequest && (
            <ReviewRequestContent
              request={selectedRequest}
              requester={requesterProfileById.get(selectedRequest.requester_id)}
              studentSexByName={studentSexByName}
              responseStatus={responseStatus}
              teacherResponse={teacherResponse}
              onStatusChange={setResponseStatus}
              onResponseChange={setTeacherResponse}
              onOpenAttachment={() => openAttachment(selectedRequest)}
            />
          )}

          <DialogFooter className="border-t pt-4">
            <Button variant="outline" onClick={() => setSelectedRequest(null)}>
              Cancel
            </Button>

            <Button
              onClick={() => updateRequestMutation.mutate()}
              disabled={updateRequestMutation.isPending}
            >
              {updateRequestMutation.isPending ? (
                <LoaderCircle className="mr-2 size-4 animate-spin" />
              ) : (
                <Save className="mr-2 size-4" />
              )}

              {updateRequestMutation.isPending ? "Saving..." : "Save Response"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(requestToDelete)}
        onOpenChange={(open) => {
          if (!open && !deleteRequestMutation.isPending) {
            setRequestToDelete(null);
          }
        }}
      >
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-destructive">
              <Trash2 className="size-5" />
              Permanently Delete Request?
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-3 text-sm">
            <p>
              This request will be permanently removed from both the Class Adviser Receive Form and
              the Subject Teacher Request Form.
            </p>

            {requestToDelete && (
              <div className="rounded-xl border bg-muted/30 p-3">
                <div className="font-medium">
                  {requestToDelete.subject} · {formatTerm(requestToDelete.term)}
                </div>

                <div className="mt-1 text-muted-foreground">
                  Sent by {requesterById.get(requestToDelete.requester_id) || "Teacher"}
                </div>
              </div>
            )}

            <p className="font-medium text-destructive">This action cannot be undone.</p>
          </div>

          <DialogFooter>
            <Button
              variant="outline"
              onClick={() => setRequestToDelete(null)}
              disabled={deleteRequestMutation.isPending}
            >
              Cancel
            </Button>

            <Button
              variant="destructive"
              onClick={() => {
                if (requestToDelete) {
                  deleteRequestMutation.mutate(requestToDelete);
                }
              }}
              disabled={deleteRequestMutation.isPending}
            >
              {deleteRequestMutation.isPending ? (
                <LoaderCircle className="mr-2 size-4 animate-spin" />
              ) : (
                <Trash2 className="mr-2 size-4" />
              )}

              {deleteRequestMutation.isPending ? "Deleting..." : "Delete Permanently"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function SubjectTeacherGradeRequestInbox({
  recipientRole,
  embedded = false,
}: {
  recipientRole: TeacherType;
  embedded?: boolean;
}) {
  const queryClient = useQueryClient();
  const [workflowStatusFilter, setWorkflowStatusFilter] = useState<WorkflowFilterStatus>("all");
  const [workflowDateFilter, setWorkflowDateFilter] = useState<DateFilter>("all");
  const [selectedWorkflowRequest, setSelectedWorkflowRequest] = useState<GradeWorkflowRow | null>(
    null,
  );
  const [workflowRequestToDelete, setWorkflowRequestToDelete] = useState<GradeWorkflowRow | null>(
    null,
  );
  const [workflowScores, setWorkflowScores] = useState<Record<string, string>>({});
  const [workflowTeacherNote, setWorkflowTeacherNote] = useState("");
  const [workflowDecision, setWorkflowDecision] = useState<WorkflowDecision>("complete");

  const {
    data: subjectTeacherId = "",
    isLoading: subjectTeacherLoading,
    error: subjectTeacherError,
  } = useQuery<string>({
    queryKey: ["grade-request-current-recipient-id", recipientRole],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      return user.id;
    },
    staleTime: 5 * 60 * 1000,
  });

  const {
    data: workflowRequestsData,
    isLoading: workflowLoading,
    isPlaceholderData: workflowPlaceholder,
    error: workflowError,
    refetch: refetchWorkflow,
    isFetching: workflowFetching,
  } = useQuery<GradeWorkflowRow[]>({
    enabled: Boolean(subjectTeacherId),
    queryKey: ["grade-request-subject-inbox", subjectTeacherId, recipientRole],
    queryFn: async () => {
      let requestQuery = (supabase as any)
        .from("grade_request_workflow_view")
        .select("*")
        .order("subject_requested_at", { ascending: false });

      requestQuery = requestQuery
        .eq("recipient_role", recipientRole)
        .eq("subject_teacher_id", subjectTeacherId);

      const { data, error } = await requestQuery;

      if (error) throw error;

      return Array.isArray(data) ? (data as unknown as GradeWorkflowRow[]) : [];
    },
    placeholderData: [],
  });

  const workflowRequests = Array.isArray(workflowRequestsData) ? workflowRequestsData : [];

  const filteredWorkflowRequests = useMemo(() => {
    return workflowRequests.filter((request) => {
      const matchesStatus =
        workflowStatusFilter === "all" || request.subject_status === workflowStatusFilter;
      const matchesDate = matchesDateFilter(request.subject_requested_at, workflowDateFilter);

      return matchesStatus && matchesDate;
    });
  }, [workflowRequests, workflowStatusFilter, workflowDateFilter]);

  const workflowCounts = useMemo(
    () => ({
      all: workflowRequests.length,
      pending: workflowRequests.filter((request) => request.subject_status === "pending").length,
      completed: workflowRequests.filter((request) => request.subject_status === "completed").length,
      rejected: workflowRequests.filter((request) => request.subject_status === "rejected").length,
    }),
    [workflowRequests],
  );

  const { data: workflowStudentsData, isLoading: workflowStudentsLoading } = useQuery<
    WorkflowStudentRow[]
  >({
    enabled: Boolean(selectedWorkflowRequest?.advisory_class_id),
    queryKey: ["grade-request-submission-students", selectedWorkflowRequest?.advisory_class_id],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("students")
        .select("id, class_id, first_name, middle_name, last_name, lrn, sex")
        .eq("class_id", selectedWorkflowRequest!.advisory_class_id)
        .order("last_name", { ascending: true })
        .order("first_name", { ascending: true });

      if (error) throw error;

      return Array.isArray(data) ? (data as unknown as WorkflowStudentRow[]) : [];
    },
    placeholderData: [],
  });

  const workflowStudents = Array.isArray(workflowStudentsData) ? workflowStudentsData : [];
  const workflowStudentGroups = useMemo(
    () => [
      {
        label: "MALE",
        learners: workflowStudents.filter((student) => student.sex === "male"),
      },
      {
        label: "FEMALE",
        learners: workflowStudents.filter((student) => student.sex === "female"),
      },
      {
        label: "SEX NOT SET",
        learners: workflowStudents.filter(
          (student) => student.sex !== "male" && student.sex !== "female",
        ),
      },
    ],
    [workflowStudents],
  );

  const { data: submittedWorkflowScoresData } = useQuery<WorkflowScoreRow[]>({
    enabled: Boolean(selectedWorkflowRequest?.subject_request_id),
    queryKey: ["grade-request-submitted-scores", selectedWorkflowRequest?.subject_request_id],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grade_request_scores")
        .select("student_id, score")
        .eq("subject_request_id", selectedWorkflowRequest!.subject_request_id);

      if (error) throw error;

      return Array.isArray(data) ? (data as unknown as WorkflowScoreRow[]) : [];
    },
    placeholderData: [],
  });

  const submittedWorkflowScores = Array.isArray(submittedWorkflowScoresData)
    ? submittedWorkflowScoresData
    : [];

  const { data: workflowGradeSourceClassesData } = useQuery<WorkflowGradeSourceClassRow[]>({
    enabled: Boolean(selectedWorkflowRequest && subjectTeacherId),
    queryKey: [
      "grade-request-grade-source-classes",
      subjectTeacherId,
      selectedWorkflowRequest?.grade_level,
      selectedWorkflowRequest?.section,
      selectedWorkflowRequest?.school_year,
      selectedWorkflowRequest?.subject,
    ],
    queryFn: async () => {
      let query = (supabase as any)
        .from("classes")
        .select("id, teacher_id, subject, grade_level, section, school_year")
        .eq("teacher_id", subjectTeacherId);

      if (selectedWorkflowRequest?.grade_level) {
        query = query.eq("grade_level", selectedWorkflowRequest.grade_level);
      }
      if (selectedWorkflowRequest?.section) {
        query = query.eq("section", selectedWorkflowRequest.section);
      }
      if (selectedWorkflowRequest?.school_year) {
        query = query.eq("school_year", selectedWorkflowRequest.school_year);
      }

      const { data, error } = await query;
      if (error) throw error;

      const rows = Array.isArray(data)
        ? (data as unknown as WorkflowGradeSourceClassRow[])
        : [];
      const requestedSubject = normalizeWorkflowValue(selectedWorkflowRequest?.subject);
      const matchingSubjectRows = rows.filter(
        (row) => normalizeWorkflowValue(row.subject) === requestedSubject,
      );

      // Prefer classes for the exact requested subject, but keep the same
      // grade/section/year classes as a fallback for older duplicated class data.
      return matchingSubjectRows.length > 0 ? matchingSubjectRows : rows;
    },
    placeholderData: [],
  });

  const workflowGradeSourceClasses = Array.isArray(workflowGradeSourceClassesData)
    ? workflowGradeSourceClassesData
    : [];

  const workflowGradeSourceClassIds = useMemo(() => {
    const ids = workflowGradeSourceClasses.map((row) => row.id).filter(Boolean);

    // Older requests may only contain subject_class_id/advisory_class_id.
    // Use them only as a fallback when no matching teacher class is visible.
    if (ids.length === 0) {
      const fallbackIds = [
        selectedWorkflowRequest?.subject_class_id,
        selectedWorkflowRequest?.advisory_class_id,
      ].filter((value): value is string => Boolean(value));
      return [...new Set(fallbackIds)];
    }

    return [...new Set(ids)];
  }, [workflowGradeSourceClasses, selectedWorkflowRequest]);

  const { data: workflowGradeSourceStudentsData } = useQuery<WorkflowStudentRow[]>({
    enabled: workflowGradeSourceClassIds.length > 0,
    queryKey: [
      "grade-request-grade-source-students",
      workflowGradeSourceClassIds.join("|"),
    ],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("students")
        .select("id, class_id, first_name, middle_name, last_name, lrn, sex")
        .in("class_id", workflowGradeSourceClassIds);

      if (error) throw error;
      return Array.isArray(data) ? (data as unknown as WorkflowStudentRow[]) : [];
    },
    placeholderData: [],
  });

  const workflowGradeSourceStudents = Array.isArray(workflowGradeSourceStudentsData)
    ? workflowGradeSourceStudentsData
    : [];

  const { data: existingWorkflowGradesData } = useQuery<WorkflowGradeRow[]>({
    enabled: Boolean(
      selectedWorkflowRequest?.grading_period && workflowGradeSourceClassIds.length > 0,
    ),
    queryKey: [
      "grade-request-existing-grades",
      workflowGradeSourceClassIds.join("|"),
      selectedWorkflowRequest?.subject,
      selectedWorkflowRequest?.grading_period,
    ],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, class_id, subject, term, score")
        .in("class_id", workflowGradeSourceClassIds)
        .eq("term", selectedWorkflowRequest!.grading_period);

      if (error) throw error;

      const rows = Array.isArray(data) ? (data as unknown as WorkflowGradeRow[]) : [];
      const requestedSubject = normalizeWorkflowValue(selectedWorkflowRequest?.subject);
      const matchingRows = rows.filter(
        (grade) => normalizeWorkflowValue(grade.subject) === requestedSubject,
      );

      // Exact subject matches are preferred. If an older class only has one
      // subject's grades under a slightly different label, use the term rows.
      return matchingRows.length > 0 ? matchingRows : rows;
    },
    placeholderData: [],
  });

  const existingWorkflowGrades = Array.isArray(existingWorkflowGradesData)
    ? existingWorkflowGradesData
    : [];

  useEffect(() => {
    if (!selectedWorkflowRequest) {
      setWorkflowScores({});
      setWorkflowTeacherNote("");
      return;
    }

    const submittedByStudent = new Map(
      submittedWorkflowScores.map((item) => [item.student_id, Number(item.score)]),
    );
    const existingByStudent = new Map(
      existingWorkflowGrades.map((grade) => [
        grade.student_id,
        grade.score == null ? null : Number(grade.score),
      ]),
    );

    const sourceStudentById = new Map(
      workflowGradeSourceStudents.map((student) => [student.id, student]),
    );
    const gradeByLrn = new Map<string, number>();
    const gradeByName = new Map<string, number>();

    existingWorkflowGrades.forEach((grade) => {
      const score = grade.score == null ? null : Number(grade.score);
      if (score == null || !Number.isFinite(score)) return;

      const sourceStudent = sourceStudentById.get(grade.student_id);
      if (!sourceStudent) return;

      const lrnKey = normalizeWorkflowLrn(sourceStudent.lrn);
      if (lrnKey && !gradeByLrn.has(lrnKey)) {
        gradeByLrn.set(lrnKey, score);
      }

      const nameKey = workflowStudentNameKey(sourceStudent);
      if (nameKey && !gradeByName.has(nameKey)) {
        gradeByName.set(nameKey, score);
      }
    });

    const nextScores: Record<string, string> = {};

    workflowStudents.forEach((student) => {
      const submittedScore = submittedByStudent.get(student.id);
      const directExistingScore = existingByStudent.get(student.id);
      const lrnExistingScore = gradeByLrn.get(normalizeWorkflowLrn(student.lrn));
      const nameExistingScore = gradeByName.get(workflowStudentNameKey(student));

      const score =
        typeof submittedScore === "number" && Number.isFinite(submittedScore)
          ? submittedScore
          : typeof directExistingScore === "number" && Number.isFinite(directExistingScore)
            ? directExistingScore
            : typeof lrnExistingScore === "number" && Number.isFinite(lrnExistingScore)
              ? lrnExistingScore
              : typeof nameExistingScore === "number" && Number.isFinite(nameExistingScore)
                ? nameExistingScore
                : null;

      nextScores[student.id] = score == null ? "" : String(score);
    });

    setWorkflowScores(nextScores);
    setWorkflowTeacherNote(selectedWorkflowRequest.teacher_note || "");
    setWorkflowDecision(selectedWorkflowRequest.subject_status === "rejected" ? "reject" : "complete");
  }, [
    selectedWorkflowRequest,
    submittedWorkflowScores,
    existingWorkflowGrades,
    workflowStudents,
    workflowGradeSourceStudents,
  ]);

  const selectedWorkflowBatchId = selectedWorkflowRequest?.batch_id || "";

  const { data: workflowHistoryData } = useQuery<WorkflowHistoryRow[]>({
    enabled: Boolean(selectedWorkflowBatchId),
    queryKey: ["grade-request-subject-history", selectedWorkflowBatchId],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grade_request_history")
        .select("*")
        .eq("batch_id", selectedWorkflowBatchId)
        .order("created_at", { ascending: false });

      if (error) throw error;

      return Array.isArray(data) ? (data as unknown as WorkflowHistoryRow[]) : [];
    },
    placeholderData: [],
  });

  const workflowHistory = Array.isArray(workflowHistoryData) ? workflowHistoryData : [];

  const submitWorkflowGradesMutation = useMutation({
    mutationFn: async () => {
      if (!selectedWorkflowRequest) {
        throw new Error("No grade request selected.");
      }

      if (selectedWorkflowRequest.is_finalized) {
        throw new Error("This grading period has already been finalized by the Class Adviser.");
      }

      if (workflowStudents.length === 0) {
        throw new Error("No learners were found in the advisory class.");
      }

      const scoreRows = workflowStudents.map((student) => {
        const rawValue = workflowScores[student.id]?.trim() || "";
        const score = Number(rawValue);

        if (!rawValue || !Number.isFinite(score) || score < 0 || score > 100) {
          throw new Error(
            `Enter a valid grade from 0 to 100 for ${student.last_name}, ${student.first_name}.`,
          );
        }

        return {
          student_id: student.id,
          score,
        };
      });

      const { error } = await (supabase as any).rpc("complete_grade_request_subject", {
        p_subject_request_id: selectedWorkflowRequest.subject_request_id,
        p_scores: scoreRows,
        p_teacher_note: workflowTeacherNote.trim() || null,
      });

      if (error) throw error;
    },
    onSuccess: async () => {
      await Promise.all([
        queryClient.invalidateQueries({
          queryKey: ["grade-request-subject-inbox"],
        }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-submitted-scores"],
        }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-subject-history"],
        }),
        queryClient.invalidateQueries({ queryKey: ["grades-all"] }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-adviser-workflow"],
        }),
      ]);

      toast.success("Grade request completed successfully.");
      setSelectedWorkflowRequest(null);
      setWorkflowScores({});
      setWorkflowTeacherNote("");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const rejectWorkflowRequestMutation = useMutation({
    mutationFn: async () => {
      if (!selectedWorkflowRequest) {
        throw new Error("No grade request selected.");
      }

      if (selectedWorkflowRequest.is_finalized) {
        throw new Error("This grading period has already been finalized by the Class Adviser.");
      }

      const { error } = await (supabase as any).rpc("reject_grade_request_subject", {
        p_subject_request_id: selectedWorkflowRequest.subject_request_id,
        p_teacher_note: workflowTeacherNote.trim() || null,
      });

      if (error) throw error;
    },
    onSuccess: async () => {
      await Promise.all([
        queryClient.invalidateQueries({ queryKey: ["grade-request-subject-inbox"] }),
        queryClient.invalidateQueries({ queryKey: ["grade-request-subject-history"] }),
        queryClient.invalidateQueries({ queryKey: ["grade-request-adviser-workflow"] }),
        queryClient.invalidateQueries({ queryKey: ["class-adviser-grade-workflow"] }),
      ]);

      toast.success("Grade request rejected. Wait for the Class Adviser to send another request.");
      setSelectedWorkflowRequest(null);
      setWorkflowScores({});
      setWorkflowTeacherNote("");
      setWorkflowDecision("complete");
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const deleteAssignedWorkflowRequestMutation = useMutation({
    mutationFn: async (request: GradeWorkflowRow) => {
      const deleteFunction =
        recipientRole === "class_adviser"
          ? "delete_received_class_adviser_submission"
          : "delete_assigned_grade_request";

      const { data, error } = await (supabase as any).rpc(deleteFunction, {
        p_subject_request_id: request.subject_request_id,
      });

      if (error) throw error;
      if (data !== true) {
        throw new Error(
          recipientRole === "class_adviser"
            ? "The Class Adviser submission could not be deleted."
            : "The assigned grade request could not be deleted.",
        );
      }
    },
    onSuccess: async () => {
      await Promise.all([
        queryClient.invalidateQueries({
          queryKey: ["grade-request-subject-inbox"],
        }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-adviser-workflow"],
        }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-submitted-scores"],
        }),
        queryClient.invalidateQueries({
          queryKey: ["grade-request-subject-history"],
        }),
      ]);

      toast.success("Assigned grade request permanently deleted.");
      setWorkflowRequestToDelete(null);
      setSelectedWorkflowRequest(null);
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  if (subjectTeacherLoading || workflowLoading || workflowPlaceholder) {
    return <WorkflowInboxSkeleton embedded={embedded} />;
  }

  return (
    <div className={embedded ? "space-y-5" : "space-y-5 pb-24 md:pb-6"}>
      <section className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-card px-5 py-4 shadow-sm">
        <div className="flex items-center gap-4">
          {!embedded && (
            <Link
              to={recipientRole === "class_adviser" ? "/dashboard" : "/subject-teacher-dashboard"}
              className="inline-flex items-center gap-1.5 rounded-xl border px-3 py-2 text-sm font-medium hover:bg-muted"
            >
              <ArrowLeft className="size-4" />
              Back
            </Link>
          )}

          <div>
            <div className="flex items-center gap-2 text-xl font-semibold">
              <Inbox className="size-6 text-primary" />
              {recipientRole === "class_adviser" ? "Receive Form" : "Receive Form"}
            </div>

            <p className="mt-0.5 text-sm text-muted-foreground">
              {recipientRole === "class_adviser"
                ? "Receive Grade Requests sent to your Class Adviser account, enter the learner grades, then complete or reject the request."
                : "Receive requests sent directly to your Subject Teacher account by the Admin or Class Adviser, then complete or reject the request for the selected class and subject."}
            </p>
          </div>
        </div>

        <Button variant="outline" onClick={() => refetchWorkflow()} disabled={workflowFetching}>
          <RefreshCw className={`mr-2 size-4 ${workflowFetching ? "animate-spin" : ""}`} />
          Refresh
        </Button>
      </section>

      <section className="grid gap-3 sm:grid-cols-3">
        <CountCard
          label="All Requests"
          value={workflowCounts.all}
          active={workflowStatusFilter === "all"}
          onClick={() => setWorkflowStatusFilter("all")}
        />

        <CountCard
          label="Pending"
          value={workflowCounts.pending}
          active={workflowStatusFilter === "pending"}
          onClick={() => setWorkflowStatusFilter("pending")}
        />

        <CountCard
          label="Completed"
          value={workflowCounts.completed}
          active={workflowStatusFilter === "completed"}
          onClick={() => setWorkflowStatusFilter("completed")}
        />

        <CountCard
          label="Rejected"
          value={workflowCounts.rejected}
          active={workflowStatusFilter === "rejected"}
          onClick={() => setWorkflowStatusFilter("rejected")}
        />
      </section>

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">
              {recipientRole === "class_adviser"
                ? "Assigned Grade Requests"
                : "Assigned Subject Requests"}
            </h2>

            <p className="text-sm text-muted-foreground">
              {filteredWorkflowRequests.length} request
              {filteredWorkflowRequests.length === 1 ? "" : "s"} shown
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <Select
              value={workflowDateFilter}
              onValueChange={(value) => setWorkflowDateFilter(value as DateFilter)}
            >
              <SelectTrigger className="w-44">
                <Calendar className="mr-2 size-4 text-primary" />
                <SelectValue placeholder="All Dates" />
              </SelectTrigger>

              <SelectContent>
                <SelectItem value="all">All Dates</SelectItem>
                <SelectItem value="today">Today</SelectItem>
                <SelectItem value="last_7_days">Last 7 Days</SelectItem>
                <SelectItem value="this_month">This Month</SelectItem>
                <SelectItem value="this_year">This Year</SelectItem>
              </SelectContent>
            </Select>

            <Select
              value={workflowStatusFilter}
              onValueChange={(value) => setWorkflowStatusFilter(value as WorkflowFilterStatus)}
            >
              <SelectTrigger className="w-44">
                <SelectValue placeholder="All Status" />
              </SelectTrigger>

              <SelectContent>
                <SelectItem value="all">All Status</SelectItem>
                <SelectItem value="pending">Pending</SelectItem>
                <SelectItem value="completed">Completed</SelectItem>
                <SelectItem value="rejected">Rejected</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full min-w-[1120px] text-sm">
            <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
              <tr>
                <th className="px-4 py-3 text-left">Class Adviser</th>
                <th className="px-4 py-3 text-left">Class</th>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Period</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Grades</th>
                <th className="px-4 py-3 text-left">Requested</th>
                <th className="px-4 py-3 text-left">Completed</th>
                <th className="px-4 py-3 text-right">Action</th>
              </tr>
            </thead>

            <tbody>
              {workflowLoading && (
                <tr>
                  <td colSpan={9} className="px-4 py-16 text-center text-muted-foreground">
                    <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                    Loading grade requests...
                  </td>
                </tr>
              )}

              {!workflowLoading &&
                filteredWorkflowRequests.map((request) => (
                  <tr key={request.subject_request_id} className="border-t">
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2">
                        <div className="grid size-8 place-items-center rounded-lg bg-primary/10 text-primary">
                          <UserRound className="size-4" />
                        </div>

                        <div>
                          <div className="font-medium">
                            {request.adviser_name || "Class Adviser"}
                          </div>
                          <div className="text-xs text-muted-foreground">
                            {request.adviser_email || "—"}
                          </div>
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3">
                      <div className="font-medium">{request.grade_level || "Grade —"}</div>
                      <div className="text-xs text-muted-foreground">
                        {request.section || "No section"} ·{" "}
                        {request.school_year || "No school year"}
                      </div>
                    </td>

                    <td className="px-4 py-3 font-medium">{request.subject}</td>

                    <td className="px-4 py-3">{formatWorkflowPeriod(request.grading_period)}</td>

                    <td className="px-4 py-3">
                      <WorkflowStatusBadge
                        status={request.subject_status}
                        finalized={request.is_finalized}
                      />
                    </td>

                    <td className="px-4 py-3">{request.submitted_grade_count}</td>

                    <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                      {formatWorkflowDate(request.subject_requested_at)}
                    </td>

                    <td className="whitespace-nowrap px-4 py-3 text-muted-foreground">
                      {formatWorkflowDate(request.submitted_at)}
                    </td>

                    <td className="px-4 py-3 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <Button
                          variant={request.subject_status === "pending" ? "default" : "outline"}
                          size="sm"
                          onClick={() => setSelectedWorkflowRequest(request)}
                        >
                          {request.subject_status === "pending" ? (
                            <Send className="mr-1 size-4" />
                          ) : (
                            <Eye className="mr-1 size-4" />
                          )}

                          {request.is_finalized
                            ? "View Grades"
                            : recipientRole === "class_adviser"
                              ? request.subject_status === "completed"
                                ? "Review / Edit"
                                : "Enter Grades"
                              : request.subject_status === "pending"
                                ? "Enter Grades"
                                : request.subject_status === "rejected"
                                  ? "View Rejected"
                                  : "Review / Edit"}
                        </Button>

                        <Button
                          type="button"
                          variant="destructive"
                          size="icon"
                          className="size-9"
                          aria-label={`Delete ${request.subject} request`}
                          title={
                            recipientRole === "class_adviser"
                              ? "Delete Class Adviser submission"
                              : "Delete assigned request"
                          }
                          onClick={() => setWorkflowRequestToDelete(request)}
                        >
                          <Trash2 className="size-4" />
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))}

              {!workflowLoading &&
                filteredWorkflowRequests.length === 0 &&
                !workflowError &&
                !subjectTeacherError && (
                  <tr>
                    <td colSpan={9} className="px-4 py-16 text-center text-muted-foreground">
                      No grade requests found for this status.
                    </td>
                  </tr>
                )}

              {(workflowError || subjectTeacherError) && (
                <tr>
                  <td colSpan={9} className="px-4 py-16 text-center text-destructive">
                    {((workflowError || subjectTeacherError) as Error).message}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      <Dialog
        open={Boolean(selectedWorkflowRequest)}
        onOpenChange={(open) => {
          if (!open && !submitWorkflowGradesMutation.isPending) {
            setSelectedWorkflowRequest(null);
            setWorkflowScores({});
            setWorkflowTeacherNote("");
          }
        }}
      >
        <DialogContent className="max-h-[94vh] overflow-y-auto sm:max-w-5xl">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-xl">
              <BookOpen className="size-6 text-primary" />
              {selectedWorkflowRequest?.is_finalized ? "Review" : "Manage"}{" "}
              {selectedWorkflowRequest?.subject} Grades
            </DialogTitle>
          </DialogHeader>

          {selectedWorkflowRequest && (
            <div className="space-y-5">
              <section className="grid gap-3 rounded-2xl border bg-muted/20 p-4 sm:grid-cols-2 lg:grid-cols-4">
                <WorkflowMetadata
                  label="Class Adviser"
                  value={selectedWorkflowRequest.adviser_name || "—"}
                />
                <WorkflowMetadata
                  label="Class"
                  value={`${
                    selectedWorkflowRequest.grade_level || "Grade —"
                  } · ${selectedWorkflowRequest.section || "No section"}`}
                />
                <WorkflowMetadata label="Subject" value={selectedWorkflowRequest.subject} />
                <WorkflowMetadata
                  label="Period"
                  value={formatWorkflowPeriod(selectedWorkflowRequest.grading_period)}
                />
                <WorkflowMetadata
                  label="Request Status"
                  value={
                    selectedWorkflowRequest.subject_status === "completed"
                      ? "Completed"
                      : selectedWorkflowRequest.subject_status === "rejected"
                        ? "Rejected"
                        : "Pending"
                  }
                />
                <WorkflowMetadata
                  label="Requested"
                  value={formatWorkflowDate(selectedWorkflowRequest.subject_requested_at)}
                />
                <WorkflowMetadata
                  label="Completed"
                  value={formatWorkflowDate(selectedWorkflowRequest.submitted_at)}
                />
                <WorkflowMetadata
                  label="Overall Status"
                  value={
                    selectedWorkflowRequest.is_finalized
                      ? "Finalized"
                      : selectedWorkflowRequest.overall_status === "completed"
                        ? "Completed"
                        : "Pending"
                  }
                />
              </section>

              {selectedWorkflowRequest.request_message && (
                <section className="rounded-2xl border border-blue-200 bg-blue-50 p-4">
                  <div className="flex items-center gap-2 text-sm font-semibold text-blue-800">
                    <MessageSquareText className="size-4" />
                    Message from Requester
                  </div>
                  <p className="mt-2 whitespace-pre-wrap text-sm leading-relaxed text-blue-900">
                    {selectedWorkflowRequest.request_message}
                  </p>
                </section>
              )}

              {selectedWorkflowRequest.is_finalized && (
                <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-800">
                  The requester has finalized this grading period. The completed grades are now
                  read-only.
                </div>
              )}

              <section className="overflow-hidden rounded-2xl border">
                <div className="border-b px-4 py-3">
                  <h3 className="font-semibold">Learner Grades</h3>
                  <p className="text-sm text-muted-foreground">
                    Enter one grade from 0 to 100 for every learner. These grades are saved only for{" "}
                    {selectedWorkflowRequest.subject}.
                  </p>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full min-w-[700px] text-sm">
                    <thead className="bg-muted/40 text-xs uppercase text-muted-foreground">
                      <tr>
                        <th className="px-4 py-3 text-left">#</th>
                        <th className="px-4 py-3 text-left">Learner</th>
                        <th className="px-4 py-3 text-left">LRN</th>
                        <th className="px-4 py-3 text-left">Grade</th>
                      </tr>
                    </thead>

                    <tbody>
                      {workflowStudentsLoading && (
                        <tr>
                          <td colSpan={4} className="px-4 py-12 text-center text-muted-foreground">
                            <LoaderCircle className="mx-auto mb-2 size-5 animate-spin" />
                            Loading learners...
                          </td>
                        </tr>
                      )}

                      {!workflowStudentsLoading &&
                        workflowStudentGroups.map((group) =>
                          group.learners.length > 0 ? (
                            <Fragment key={group.label}>
                              <tr className="border-t bg-primary/10">
                                <td colSpan={4} className="px-4 py-2 font-bold text-primary">
                                  {group.label} ({group.learners.length})
                                </td>
                              </tr>
                              {group.learners.map((student, index) => (
                                <tr key={student.id} className="border-t">
                                  <td className="px-4 py-2 text-muted-foreground">{index + 1}</td>

                                  <td className="px-4 py-2 font-medium">
                                    {student.last_name}, {student.first_name}{" "}
                                    {student.middle_name || ""}
                                  </td>

                                  <td className="px-4 py-2 text-muted-foreground">
                                    {student.lrn || "—"}
                                  </td>

                                  <td className="w-48 px-4 py-2">
                                    <Input
                                      type="number"
                                      min={0}
                                      max={100}
                                      step="0.01"
                                      value={workflowScores[student.id] ?? ""}
                                      disabled={selectedWorkflowRequest.is_finalized || selectedWorkflowRequest.subject_status === "rejected"}
                                      onChange={(event) => {
                                        const value = event.target.value;

                                        if (value === "" || /^\d{0,3}(?:\.\d{0,2})?$/.test(value)) {
                                          setWorkflowScores((current) => ({
                                            ...current,
                                            [student.id]: value,
                                          }));
                                        }
                                      }}
                                      placeholder="0–100"
                                    />
                                  </td>
                                </tr>
                              ))}
                            </Fragment>
                          ) : null,
                        )}

                      {!workflowStudentsLoading && workflowStudents.length === 0 && (
                        <tr>
                          <td colSpan={4} className="px-4 py-12 text-center text-muted-foreground">
                            No learners found in this advisory class.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section>
                <div className="mb-1.5 flex items-center justify-between">
                  <label htmlFor="workflow-teacher-note" className="text-sm font-medium">
                    Submission Note{" "}
                    <span className="font-normal text-muted-foreground">(optional)</span>
                  </label>

                  <span className="text-xs text-muted-foreground">
                    {workflowTeacherNote.length}/500
                  </span>
                </div>

                <textarea
                  id="workflow-teacher-note"
                  value={workflowTeacherNote}
                  disabled={selectedWorkflowRequest.is_finalized || selectedWorkflowRequest.subject_status === "rejected"}
                  maxLength={500}
                  onChange={(event) => setWorkflowTeacherNote(event.target.value)}
                  placeholder="Add a note for the Class Adviser..."
                  className="min-h-24 w-full resize-none rounded-xl border bg-background px-3 py-2 text-sm outline-none transition placeholder:text-muted-foreground focus:border-primary focus:ring-2 focus:ring-primary/20 disabled:cursor-not-allowed disabled:opacity-70"
                />
              </section>

              <section className="rounded-2xl border bg-muted/20 p-4">
                <div className="flex items-center gap-2 font-semibold">
                  <History className="size-5 text-primary" />
                  Request History
                </div>

                <div className="mt-3 space-y-2">
                  {workflowHistory.slice(0, 12).map((history) => (
                    <div
                      key={history.id}
                      className="flex flex-wrap justify-between gap-2 rounded-lg border bg-background px-3 py-2 text-sm"
                    >
                      <div>
                        <span className="font-medium capitalize">
                          {history.event_type.replaceAll("_", " ")}
                        </span>

                        {(history.old_status || history.new_status) && (
                          <span className="ml-2 text-xs text-muted-foreground">
                            {history.old_status || "—"} → {history.new_status || "—"}
                          </span>
                        )}
                      </div>

                      <span className="text-muted-foreground">
                        {formatWorkflowDate(history.created_at)}
                      </span>
                    </div>
                  ))}

                  {workflowHistory.length === 0 && (
                    <p className="text-sm text-muted-foreground">No history entries yet.</p>
                  )}
                </div>
              </section>
            </div>
          )}

          <DialogFooter className="border-t pt-4">
            <Button
              variant="outline"
              onClick={() => setSelectedWorkflowRequest(null)}
              disabled={submitWorkflowGradesMutation.isPending}
            >
              Close
            </Button>

            {!selectedWorkflowRequest?.is_finalized &&
              selectedWorkflowRequest?.subject_status !== "rejected" && (
                <>
                  <Select
                    value={workflowDecision}
                    onValueChange={(value) => setWorkflowDecision(value as WorkflowDecision)}
                    disabled={
                      submitWorkflowGradesMutation.isPending ||
                      rejectWorkflowRequestMutation.isPending
                    }
                  >
                    <SelectTrigger className="w-48">
                      <SelectValue placeholder="Select action" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="complete">Complete Request</SelectItem>
                      <SelectItem value="reject">Reject Request</SelectItem>
                    </SelectContent>
                  </Select>

                  <Button
                    variant={workflowDecision === "reject" ? "destructive" : "default"}
                    onClick={() =>
                      workflowDecision === "reject"
                        ? rejectWorkflowRequestMutation.mutate()
                        : submitWorkflowGradesMutation.mutate()
                    }
                    disabled={
                      submitWorkflowGradesMutation.isPending ||
                      rejectWorkflowRequestMutation.isPending ||
                      (workflowDecision === "complete" &&
                        (workflowStudentsLoading || workflowStudents.length === 0))
                    }
                  >
                    {submitWorkflowGradesMutation.isPending ||
                    rejectWorkflowRequestMutation.isPending ? (
                      <LoaderCircle className="mr-2 size-4 animate-spin" />
                    ) : workflowDecision === "reject" ? (
                      <XCircle className="mr-2 size-4" />
                    ) : (
                      <CheckCircle2 className="mr-2 size-4" />
                    )}

                    {rejectWorkflowRequestMutation.isPending
                      ? "Rejecting..."
                      : submitWorkflowGradesMutation.isPending
                        ? "Completing..."
                        : workflowDecision === "reject"
                          ? "Reject Request"
                          : selectedWorkflowRequest?.subject_status === "completed"
                            ? "Update Completed Grades"
                            : "Complete Request"}
                  </Button>
                </>
              )}
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog
        open={Boolean(workflowRequestToDelete)}
        onOpenChange={(open) => {
          if (!open && !deleteAssignedWorkflowRequestMutation.isPending) {
            setWorkflowRequestToDelete(null);
          }
        }}
      >
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-destructive">
              <Trash2 className="size-5" />
              {recipientRole === "class_adviser"
                ? "Delete Assigned Grade Request?"
                : "Delete Assigned Subject Request?"}
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-3 text-sm">
            <p>
              This will permanently delete only this selected{" "}
              {recipientRole === "class_adviser" ? "assigned Grade Request" : "assigned request"}{" "}
              and its submitted grades and history. Other requests will not be changed. This action
              cannot be undone.
            </p>

            {workflowRequestToDelete && (
              <div className="rounded-xl border bg-muted/30 p-3">
                <div className="font-semibold">{workflowRequestToDelete.subject}</div>
                <div className="mt-1 text-muted-foreground">
                  {workflowRequestToDelete.grade_level || "Grade —"} ·{" "}
                  {workflowRequestToDelete.section || "No section"} ·{" "}
                  {formatWorkflowPeriod(workflowRequestToDelete.grading_period)}
                </div>
              </div>
            )}
          </div>

          <DialogFooter>
            <Button
              type="button"
              variant="outline"
              onClick={() => setWorkflowRequestToDelete(null)}
              disabled={deleteAssignedWorkflowRequestMutation.isPending}
            >
              Cancel
            </Button>

            <Button
              type="button"
              variant="destructive"
              onClick={() => {
                if (workflowRequestToDelete) {
                  deleteAssignedWorkflowRequestMutation.mutate(workflowRequestToDelete);
                }
              }}
              disabled={deleteAssignedWorkflowRequestMutation.isPending}
            >
              {deleteAssignedWorkflowRequestMutation.isPending ? (
                <LoaderCircle className="mr-2 size-4 animate-spin" />
              ) : (
                <Trash2 className="mr-2 size-4" />
              )}
              {deleteAssignedWorkflowRequestMutation.isPending
                ? "Deleting..."
                : "Delete Permanently"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function WorkflowStatusBadge({
  status,
  finalized,
}: {
  status: WorkflowSubjectStatus;
  finalized: boolean;
}) {
  if (finalized) {
    return (
      <span className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-300 bg-emerald-100 px-2.5 py-1 text-xs font-medium text-emerald-800">
        <CheckCircle2 className="size-3.5" />
        Finalized
      </span>
    );
  }

  if (status === "completed") {
    return (
      <span className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-700">
        <CheckCircle2 className="size-3.5" />
        Completed
      </span>
    );
  }

  if (status === "rejected") {
    return (
      <span className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-red-50 px-2.5 py-1 text-xs font-medium text-red-700">
        <XCircle className="size-3.5" />
        Rejected
      </span>
    );
  }

  return (
    <span className="inline-flex items-center gap-1.5 rounded-lg border border-amber-200 bg-amber-50 px-2.5 py-1 text-xs font-medium text-amber-700">
      <Clock3 className="size-3.5" />
      Pending
    </span>
  );
}

function WorkflowMetadata({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl bg-background/70 p-3">
      <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
        {label}
      </div>
      <div className="mt-1 font-medium">{value}</div>
    </div>
  );
}

function formatWorkflowPeriod(value: GradingPeriod | string) {
  if (value === "final") return "Final Grade";

  const normalized = String(value).trim().toLowerCase();
  return normalized.startsWith("term") ? String(value) : `Term ${value}`;
}

function formatWorkflowDate(value: string | null | undefined) {
  if (!value) return "—";
  return formatDate(value);
}

function ReviewRequestContent({
  request,
  requester,
  studentSexByName,
  responseStatus,
  teacherResponse,
  onStatusChange,
  onResponseChange,
  onOpenAttachment,
}: {
  request: RequestRow;
  requester?: RequesterProfile;
  studentSexByName: Map<string, LearnerSex>;
  responseStatus: RequestStatus;
  teacherResponse: string;
  onStatusChange: (status: RequestStatus) => void;
  onResponseChange: (value: string) => void;
  onOpenAttachment: () => void;
}) {
  const parsed = parseGradeRequestMessage(request.message, studentSexByName);

  const maleLearners = parsed.learners.filter((learner) => learner.sex === "male");

  const femaleLearners = parsed.learners.filter((learner) => learner.sex === "female");

  const unidentifiedLearners = parsed.learners.filter((learner) => learner.sex === "unknown");

  const requesterRole =
    requester?.teacher_type === "subject_teacher"
      ? "Subject Teacher"
      : requester?.teacher_type === "class_adviser"
        ? "Class Adviser"
        : "Teacher";

  return (
    <div className="space-y-5">
      <section className="grid gap-3 rounded-2xl border bg-muted/20 p-4 sm:grid-cols-2">
        <MetadataCard
          icon={UserRound}
          label="Requester"
          value={requester?.full_name || requester?.email || "Teacher"}
          helper={`${requesterRole} · ${request.subject}`}
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
            <p className="mt-1 text-sm text-muted-foreground">{parsed.classDetails}</p>
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

      {request.attachment_url && (
        <Button variant="outline" onClick={onOpenAttachment}>
          <Download className="mr-2 size-4" />
          Open {request.attachment_name || "Attachment"}
        </Button>
      )}

      <div>
        <div className="mb-1.5 text-sm font-medium">Update Status</div>

        <Select
          value={responseStatus}
          onValueChange={(value) => onStatusChange(value as RequestStatus)}
        >
          <SelectTrigger>
            <SelectValue placeholder="Select status" />
          </SelectTrigger>

          <SelectContent>
            <SelectItem value="pending">Pending</SelectItem>
            <SelectItem value="in_progress">In Progress</SelectItem>
            <SelectItem value="completed">Completed</SelectItem>
            <SelectItem value="rejected">Rejected</SelectItem>
          </SelectContent>
        </Select>
      </div>

      <div>
        <div className="mb-1.5 flex items-center justify-between">
          <label htmlFor="teacher-response" className="text-sm font-medium">
            Teacher Response <span className="text-muted-foreground">(optional)</span>
          </label>

          <span className="text-xs text-muted-foreground">{teacherResponse.length}/1000</span>
        </div>

        <textarea
          id="teacher-response"
          value={teacherResponse}
          maxLength={1000}
          onChange={(event) => onResponseChange(event.target.value)}
          placeholder="Enter your response or remarks..."
          className="min-h-28 w-full resize-none rounded-xl border bg-background px-3 py-2 text-sm outline-none transition placeholder:text-muted-foreground focus:border-primary focus:ring-2 focus:ring-primary/20"
        />
      </div>
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
      <div className={`flex items-center gap-2 border-b px-4 py-3 ${headerClassName}`}>
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
              <tr key={`${title}-${learner.number}-${learner.name}`} className="border-t">
                <td className="px-2 py-2 text-muted-foreground">{index + 1}</td>

                <td className="whitespace-nowrap px-2 py-2 font-medium">{learner.name}</td>

                <td className="px-2 py-2 text-center font-semibold">{learner.gradeText}</td>

                <td className="px-2 py-2 text-center font-semibold">{learner.grade ?? "—"}</td>

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
                <td colSpan={7} className="px-3 py-8 text-center text-muted-foreground">
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
      <div className={`grid size-10 shrink-0 place-items-center rounded-full ${iconTone}`}>
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

        {helper && <div className="mt-0.5 text-xs text-muted-foreground">{helper}</div>}
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
  const studentHeaderIndex = lines.findIndex((line) => /^students\s*\(\d+\)\s*:?\s*$/i.test(line));

  const classDetails =
    studentHeaderIndex > 1 ? lines.slice(1, studentHeaderIndex).join(" · ") : lines[1] || "";

  const learnerLines = studentHeaderIndex >= 0 ? lines.slice(studentHeaderIndex + 1) : [];

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
    rank: learner.grade === null ? null : rankedGrades.indexOf(learner.grade) + 1,
    descriptor: gradeDescriptor(learner.grade),
    award: gradeAward(learner.grade),
  }));

  return {
    heading,
    classDetails,
    learners,
  };
}

function normalizeWorkflowValue(value: string | null | undefined) {
  return String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeWorkflowLrn(value: string | null | undefined) {
  return String(value ?? "").replace(/[^0-9a-zA-Z]/g, "").toLowerCase();
}

function workflowStudentNameKey(student: {
  first_name: string | null;
  last_name: string | null;
}) {
  // First + last is intentionally used as the fallback because older
  // duplicated rosters do not always store middle names consistently.
  return normalizeWorkflowValue(`${student.last_name ?? ""}|${student.first_name ?? ""}`);
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

function formatTerm(value: string) {
  return value === "final" ? "Final Grade" : `Term ${value}`;
}

function capitalize(value: string) {
  if (!value) return "—";
  return value.charAt(0).toUpperCase() + value.slice(1);
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
      <div className={`text-sm ${active ? "text-primary-foreground/80" : "text-muted-foreground"}`}>
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

function matchesDateFilter(value: string, filter: DateFilter) {
  if (filter === "all") return true;

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) return false;

  const now = new Date();

  if (filter === "today") {
    return (
      date.getFullYear() === now.getFullYear() &&
      date.getMonth() === now.getMonth() &&
      date.getDate() === now.getDate()
    );
  }

  if (filter === "last_7_days") {
    const sevenDaysAgo = new Date(now);
    sevenDaysAgo.setDate(now.getDate() - 7);
    return date >= sevenDaysAgo && date <= now;
  }

  if (filter === "this_month") {
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  }

  return date.getFullYear() === now.getFullYear();
}

function formatRelativeDate(value: string) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) return "";

  const now = new Date();
  const differenceMs = now.getTime() - date.getTime();
  const dayMs = 24 * 60 * 60 * 1000;
  const days = Math.max(0, Math.floor(differenceMs / dayMs));

  if (days === 0) return "Today";
  if (days === 1) return "1 day ago";
  if (days < 30) return `${days} days ago`;

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