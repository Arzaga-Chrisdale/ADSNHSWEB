import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useMemo, useState } from "react";
import {
  Archive,
  BookOpen,
  CheckCircle2,
  Clock3,
  History,
  Loader2,
  Search,
  Send,
  UserRound,
  X,
} from "lucide-react";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Skeleton } from "@/components/ui/skeleton";
import { Textarea } from "@/components/ui/textarea";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import type { ClassRow, StudentRow } from "@/lib/data";

export const Route = createFileRoute("/_authenticated/credential-requests")({
  component: CredentialRequestsPage,
});

type CredentialRequestStatus =
  | "pending"
  | "pending_review"
  | "approved"
  | "released"
  | "rejected"
  | "completed"
  | string;

type SF1StudentRow = StudentRow & {
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
  learning_modality?: string | null;
  remarks?: string | null;
};

type CredentialRequestRow = {
  id: string;
  student_id: string | null;
  requested_student_name: string | null;
  requester_id: string;
  requesting_class_id: string | null;
  previous_class_id: string | null;
  reason: string | null;
  status: CredentialRequestStatus;
  release_notes: string | null;
  admin_remarks: string | null;
  reviewed_at: string | null;
  released_at: string | null;
  processed_by: string | null;
  processed_at: string | null;
  request_date: string;
  created_at: string;
};

type CredentialSubmissionRow = {
  id: string;
  student_id: string;
  submitted_by: string;
  school_year: string;
  final_term: string | null;
  document_notes: string | null;
  status: string;
  created_at: string;
};

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

function studentName(student: StudentRow | null | undefined) {
  if (!student) return "Unknown learner";

  return (
    [student.last_name, student.first_name, student.middle_name]
      .filter(Boolean)
      .join(", ")
      .replace(", ,", ", ") || "No name"
  );
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

function CredentialRequestsSkeleton() {
  return (
    <div
      className="space-y-6 pb-24 md:pb-6"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading credential requests...</span>

      <div className="space-y-2" aria-hidden="true">
        <Skeleton className="h-8 w-72 max-w-full" />
        <Skeleton className="h-4 w-[28rem] max-w-full" />
      </div>

      <div className="grid gap-3 sm:grid-cols-3" aria-hidden="true">
        {Array.from({ length: 3 }).map((_, index) => (
          <div
            key={index}
            className="rounded-3xl border bg-card p-5 shadow-sm"
          >
            <div className="flex items-center justify-between gap-4">
              <div className="space-y-3">
                <Skeleton className="h-4 w-36" />
                <Skeleton className="h-9 w-12" />
              </div>
              <Skeleton className="size-11 shrink-0 rounded-full" />
            </div>
          </div>
        ))}
      </div>

      <section
        className="rounded-3xl border bg-card p-5 shadow-sm"
        aria-hidden="true"
      >
        <div className="mb-5 flex items-start justify-between gap-3">
          <div className="flex min-w-0 items-start gap-2">
            <Skeleton className="mt-0.5 size-5 shrink-0 rounded-md" />
            <div className="space-y-2">
              <Skeleton className="h-5 w-52 max-w-full" />
              <Skeleton className="h-3.5 w-[30rem] max-w-full" />
            </div>
          </div>
          <Skeleton className="h-6 w-28 shrink-0 rounded-full" />
        </div>

        <div className="mb-5 grid gap-2 md:grid-cols-3">
          {Array.from({ length: 3 }).map((_, index) => (
            <div key={index} className="space-y-2 rounded-2xl border p-3">
              <Skeleton className="h-3 w-12" />
              <Skeleton className="h-4 w-36 max-w-full" />
            </div>
          ))}
        </div>

        <div className="space-y-4">
          <div className="space-y-2">
            <Skeleton className="h-4 w-36" />
            <Skeleton className="h-10 w-full rounded-md" />
            <Skeleton className="h-3 w-[32rem] max-w-full" />
          </div>

          <div className="space-y-2">
            <Skeleton className="h-4 w-16" />
            <Skeleton className="h-24 w-full rounded-md" />
          </div>

          <Skeleton className="h-10 w-full rounded-md" />
        </div>
      </section>

      <section
        className="rounded-3xl border bg-card p-5 shadow-sm"
        aria-hidden="true"
      >
        <div className="mb-5 flex items-center gap-2">
          <Skeleton className="size-5 rounded-md" />
          <Skeleton className="h-5 w-32" />
        </div>

        <div className="overflow-hidden rounded-xl border">
          <div className="grid min-w-[850px] grid-cols-[1.35fr_1.2fr_0.8fr_0.8fr_0.9fr_1.35fr_0.8fr] gap-4 border-b bg-muted/30 px-3 py-3">
            {Array.from({ length: 7 }).map((_, index) => (
              <Skeleton key={index} className="h-3 w-20" />
            ))}
          </div>

          {Array.from({ length: 4 }).map((_, rowIndex) => (
            <div
              key={rowIndex}
              className="grid min-w-[850px] grid-cols-[1.35fr_1.2fr_0.8fr_0.8fr_0.9fr_1.35fr_0.8fr] items-center gap-4 border-b px-3 py-4 last:border-0"
            >
              <Skeleton className="h-4 w-32" />
              <Skeleton className="h-4 w-36" />
              <Skeleton className="h-4 w-24" />
              <Skeleton className="h-6 w-24 rounded-full" />
              <Skeleton className="h-4 w-24" />
              <Skeleton className="h-4 w-40" />
              <Skeleton className="ml-auto h-8 w-24 rounded-md" />
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function CredentialRequestsPage() {
  const queryClient = useQueryClient();
  const navigate = useNavigate();

  const [studentSearch, setStudentSearch] = useState("");

  const [requestReason, setRequestReason] = useState(
    "Enrollment credential request",
  );
  const [assigningRequest, setAssigningRequest] =
    useState<CredentialRequestRow | null>(null);
  const [selectedClassId, setSelectedClassId] = useState("");

  const { data, isLoading, error } = useQuery({
    queryKey: ["student-credential-workspace"],

    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;

      if (!user) {
        throw new Error("You must be signed in.");
      }

      const [profileRes, classesRes, requestsRes, submissionsRes] =
        await Promise.all([
          supabase
            .from("profiles")
            .select("teacher_type")
            .eq("id", user.id)
            .maybeSingle(),

          supabase.from("classes").select("*").eq("teacher_id", user.id),

          (supabase as any)
            .from("student_credential_requests")
            .select("*")
            .eq("requester_id", user.id)
            .order("request_date", { ascending: false }),

          (supabase as any)
            .from("student_credentials")
            .select("*")
            .eq("submitted_by", user.id)
            .order("created_at", { ascending: false }),
        ]);

      if (profileRes.error) throw profileRes.error;
      if (classesRes.error) throw classesRes.error;
      if (requestsRes.error) throw requestsRes.error;
      if (submissionsRes.error) throw submissionsRes.error;

      const classes = (classesRes.data ?? []) as ClassRow[];
      const requests = (requestsRes.data ?? []) as CredentialRequestRow[];

      const classIds = classes.map((item) => item.id);
      const requestStudentIds = Array.from(
        new Set(
          requests
            .map((request) => request.student_id)
            .filter((studentId): studentId is string => Boolean(studentId)),
        ),
      );

      let students: SF1StudentRow[] = [];

      if (classIds.length) {
        const { data: studentsData, error: studentsError } = await supabase
          .from("students")
          .select("*")
          .in("class_id", classIds)
          .order("last_name", { ascending: true });

        if (studentsError) throw studentsError;

        students = (studentsData ?? []) as SF1StudentRow[];
      }

      if (requestStudentIds.length) {
        const { data: requestedStudentsData, error: requestedStudentsError } =
          await supabase
            .from("students")
            .select("*")
            .in("id", requestStudentIds)
            .order("last_name", { ascending: true });

        if (requestedStudentsError) throw requestedStudentsError;

        students = Array.from(
          new Map(
            [
              ...students,
              ...((requestedStudentsData ?? []) as SF1StudentRow[]),
            ].map((student) => [student.id, student]),
          ).values(),
        );
      }

      return {
        currentUserId: user.id,
        teacherType: profileRes.data?.teacher_type ?? null,
        classes,
        students,

        requests,

        submissions: (submissionsRes.data ?? []) as CredentialSubmissionRow[],
      };
    },
  });

  const classes = data?.classes ?? [];
  const students = data?.students ?? [];
  const requests = data?.requests ?? [];
  const submissions = data?.submissions ?? [];
  const isClassAdviser = data?.teacherType === "class_adviser";

  const classById = useMemo(
    () => new Map(classes.map((item) => [item.id, item])),
    [classes],
  );

  const studentById = useMemo(
    () => new Map(students.map((item) => [item.id, item])),
    [students],
  );

  const requestCredentials = useMutation({
    mutationFn: async () => {
      if (!isClassAdviser) {
        throw new Error("Only a Class Adviser can request an SF1 credential.");
      }

      const requestedStudentName = studentSearch.trim();

      if (requestedStudentName.length < 2) {
        throw new Error(
          "Enter the student's name or LRN before sending the request.",
        );
      }

      const hasOpenRequest = requests.some(
        (request) =>
          normalize(
            request.requested_student_name ||
              studentName(
                request.student_id ? studentById.get(request.student_id) : null,
              ),
          ) === normalize(requestedStudentName) &&
          ["pending", "pending_review"].includes(normalize(request.status)),
      );

      if (hasOpenRequest) {
        throw new Error("This learner already has an open credential request.");
      }

      const { error: insertError } = await (supabase as any)
        .from("student_credential_requests")
        .insert({
          student_id: null,
          requested_student_name: requestedStudentName,
          requester_id: data?.currentUserId,
          requesting_class_id: null,
          previous_class_id: null,
          reason: requestReason,
          status: "pending_review",
        });

      if (insertError) throw insertError;
    },

    onSuccess: () => {
      toast.success("Credential request sent to Admin.");

      setStudentSearch("");

      setRequestReason("Enrollment credential request");

      queryClient.invalidateQueries({
        queryKey: ["student-credential-workspace"],
      });
    },

    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  const assignApprovedStudent = useMutation({
    mutationFn: async () => {
      if (!assigningRequest || !selectedClassId) {
        throw new Error("Select the class you want to open.");
      }

      const { error: assignmentError } = await (supabase as any).rpc(
        "assign_approved_credential_student",
        {
          p_request_id: assigningRequest.id,
          p_class_id: selectedClassId,
        },
      );

      if (assignmentError) throw assignmentError;
    },
    onSuccess: async () => {
      const classId = selectedClassId;
      toast.success("Opening the selected class.");
      setAssigningRequest(null);
      setSelectedClassId("");
      await queryClient.invalidateQueries({
        queryKey: ["student-credential-workspace"],
      });
      navigate({ to: "/classes/$classId", params: { classId } });
    },
    onError: (mutationError: Error) => {
      toast.error(mutationError.message);
    },
  });

  if (isLoading) {
    return <CredentialRequestsSkeleton />;
  }

  if (error) {
    return (
      <div className="rounded-2xl border border-red-200 bg-red-50 p-5 text-sm text-red-700">
        {(error as Error).message}
      </div>
    );
  }

  const pendingRequests = requests.filter((request) =>
    ["pending", "pending_review"].includes(normalize(request.status)),
  ).length;

  const approvedRequests = requests.filter((request) =>
    ["approved", "completed", "released"].includes(normalize(request.status)),
  ).length;

  return (
    <div className="space-y-6 pb-24 md:pb-6">
      {/* =========================================================
          PAGE HEADER
      ========================================================== */}

      <div>
        <h1 className="text-2xl font-semibold">
          Student Credential Request System
        </h1>

        <p className="text-sm text-muted-foreground">
          Request individual learner credentials from Admin for enrollment.
        </p>
      </div>

      {/* =========================================================
          SUMMARY CARDS
      ========================================================== */}

      <div className="grid gap-3 sm:grid-cols-3">
        <SummaryCard
          icon={Archive}
          label="Submitted Credentials"
          value={submissions.length}
        />

        <SummaryCard
          icon={Clock3}
          label="Pending Requests"
          value={pendingRequests}
        />

        <SummaryCard
          icon={CheckCircle2}
          label="Approved Requests"
          value={approvedRequests}
        />
      </div>

      {/* =========================================================
          REQUEST STUDENT CREDENTIALS
          This is now the ONLY credential action UI for adviser.
      ========================================================== */}

      <section className="rounded-3xl border bg-card p-5 shadow-sm">
        <div className="mb-5 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <UserRound className="size-5 text-primary" />

            <div>
              <h2 className="font-semibold">Request Student Credentials</h2>

              <p className="text-xs text-muted-foreground">
                Admin approves the learner's SF1 record, then you select the
                class where the student will be added.
              </p>
            </div>
          </div>

          <div className="rounded-full bg-primary/10 px-3 py-1 text-xs text-primary">
            Credential Request
          </div>
        </div>

        <div className="mb-5 grid gap-2 md:grid-cols-3">
          <div className="rounded-2xl border p-3">
            <p className="text-xs text-muted-foreground">Step 1</p>
            <p className="font-medium">Enter name or LRN</p>
          </div>
          <div className="rounded-2xl border p-3">
            <p className="text-xs text-muted-foreground">Step 2</p>
            <p className="font-medium">Admin verifies SF1</p>
          </div>
          <div className="rounded-2xl border p-3">
            <p className="text-xs text-muted-foreground">Step 3</p>
            <p className="font-medium">Submit request</p>
          </div>
        </div>

        {!isClassAdviser ? (
          <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
            Only a Class Adviser account can search for a learner and submit an
            SF1 credential request.
          </div>
        ) : (
          <div className="space-y-4">
            <div>
              <Label htmlFor="credential-student-search">
                Student name or LRN
              </Label>

              <div className="relative mt-1.5">
                <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  id="credential-student-search"
                  value={studentSearch}
                  onChange={(event) => setStudentSearch(event.target.value)}
                  placeholder="Enter the student's name or LRN..."
                  className="pl-9 pr-9"
                />
                {studentSearch && (
                  <button
                    type="button"
                    onClick={() => setStudentSearch("")}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
                    aria-label="Clear learner search"
                  >
                    <X className="size-4" />
                  </button>
                )}
              </div>

              <p className="mt-2 text-xs text-muted-foreground">
                Admin will search this name or LRN and connect the request to
                the correct SF1 learner record.
              </p>
            </div>

            {/* Reason */}

            <div>
              <Label>Reason</Label>

              <Textarea
                value={requestReason}
                onChange={(event) => setRequestReason(event.target.value)}
                rows={4}
                placeholder="Enter reason for requesting the student's credentials..."
              />
            </div>

            {/* Send Request */}

            <Button
              type="button"
              onClick={() => requestCredentials.mutate()}
              disabled={
                requestCredentials.isPending || studentSearch.trim().length < 2
              }
              className="w-full"
            >
              {requestCredentials.isPending ? (
                <Loader2 className="mr-2 size-4 animate-spin" />
              ) : (
                <Send className="mr-2 size-4" />
              )}
              Send Request to Admin
            </Button>
          </div>
        )}
      </section>

      {/* =========================================================
          REQUEST HISTORY
      ========================================================== */}

      <section className="rounded-3xl border bg-card p-5 shadow-sm">
        <div className="mb-5 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <History className="size-5 text-primary" />

            <h2 className="font-semibold">Request History</h2>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full min-w-[850px] text-sm">
            <thead>
              <tr className="border-b text-left text-xs uppercase text-muted-foreground">
                <th className="px-3 py-2">Student</th>

                <th className="px-3 py-2">Class</th>

                <th className="px-3 py-2">Request Date</th>

                <th className="px-3 py-2">Status</th>

                <th className="px-3 py-2">Approved Date</th>

                <th className="px-3 py-2">Notes</th>

                <th className="px-3 py-2 text-right">Action</th>
              </tr>
            </thead>

            <tbody>
              {requests.map((request) => {
                const student = request.student_id
                  ? studentById.get(request.student_id)
                  : null;

                const cls = request.requesting_class_id
                  ? classById.get(request.requesting_class_id)
                  : null;
                const approved = ["approved", "completed", "released"].includes(
                  normalize(request.status),
                );

                return (
                  <tr key={request.id} className="border-b last:border-0">
                    <td className="px-3 py-3 font-medium">
                      {student
                        ? studentName(student)
                        : request.requested_student_name ||
                          "Student name not entered"}
                    </td>

                    <td className="px-3 py-3">
                      {cls
                        ? `${cls.subject || "Class"} · ${cls.grade_level || "-"} - ${
                            cls.section || "-"
                          }`
                        : "-"}
                    </td>

                    <td className="px-3 py-3">
                      {formatDate(request.request_date || request.created_at)}
                    </td>

                    <td className="px-3 py-3">
                      <StatusPill status={request.status} />
                    </td>

                    <td className="px-3 py-3">
                      {formatDate(
                        request.released_at ||
                          request.processed_at ||
                          request.reviewed_at,
                      )}
                    </td>

                    <td className="px-3 py-3">
                      {request.admin_remarks ||
                        request.release_notes ||
                        request.reason ||
                        "-"}
                    </td>

                    <td className="px-3 py-3 text-right">
                      {approved && request.student_id ? (
                        <Button
                          type="button"
                          size="sm"
                          variant="outline"
                          onClick={() => {
                            setAssigningRequest(request);
                            setSelectedClassId("");
                          }}
                        >
                          <BookOpen className="mr-2 size-4" />
                          Open Class
                        </Button>
                      ) : (
                        <span className="text-xs text-muted-foreground">—</span>
                      )}
                    </td>
                  </tr>
                );
              })}

              {requests.length === 0 && (
                <tr>
                  <td
                    colSpan={7}
                    className="px-3 py-10 text-center text-muted-foreground"
                  >
                    No credential requests yet.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      <Dialog
        open={Boolean(assigningRequest)}
        onOpenChange={(open) => {
          if (!open && !assignApprovedStudent.isPending) {
            setAssigningRequest(null);
            setSelectedClassId("");
          }
        }}
      >
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Select a class to open</DialogTitle>
            <DialogDescription>
              Choose MAPEH or another class. The approved learner and complete
              SF1 information will be connected to that class before it opens.
            </DialogDescription>
          </DialogHeader>

          <div className="space-y-3 py-2">
            <div className="rounded-2xl border bg-muted/30 p-3 text-sm">
              <p className="text-xs text-muted-foreground">Approved learner</p>
              <p className="font-semibold">
                {assigningRequest?.student_id &&
                studentById.get(assigningRequest.student_id)
                  ? studentName(studentById.get(assigningRequest.student_id))
                  : assigningRequest?.requested_student_name ||
                    "Selected student"}
              </p>
            </div>

            <div>
              <Label>Class</Label>
              <Select
                value={selectedClassId}
                onValueChange={setSelectedClassId}
              >
                <SelectTrigger className="mt-1.5">
                  <SelectValue placeholder="Select subject, grade, and section" />
                </SelectTrigger>
                <SelectContent>
                  {classes.map((cls) => (
                    <SelectItem key={cls.id} value={cls.id}>
                      {cls.subject || "Class"} · {cls.grade_level || "—"} -{" "}
                      {cls.section || "—"} · {cls.school_year || "—"}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              {classes.length === 0 && (
                <p className="mt-2 text-xs text-amber-700">
                  Create a class first before opening this approved learner.
                </p>
              )}
            </div>
          </div>

          <DialogFooter>
            <Button
              type="button"
              variant="outline"
              onClick={() => setAssigningRequest(null)}
              disabled={assignApprovedStudent.isPending}
            >
              Cancel
            </Button>
            <Button
              type="button"
              onClick={() => assignApprovedStudent.mutate()}
              disabled={!selectedClassId || assignApprovedStudent.isPending}
            >
              {assignApprovedStudent.isPending ? (
                <Loader2 className="mr-2 size-4 animate-spin" />
              ) : (
                <BookOpen className="mr-2 size-4" />
              )}
              Open Selected Class
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function SummaryCard({
  icon: Icon,
  label,
  value,
}: {
  icon: typeof Archive;
  label: string;
  value: number;
}) {
  return (
    <div className="rounded-3xl border bg-card p-5 shadow-sm">
      <div className="flex items-center justify-between">
        <div>
          <p className="text-sm text-muted-foreground">{label}</p>

          <p className="mt-2 text-3xl font-bold">{value}</p>
        </div>

        <div className="grid size-11 place-items-center rounded-full bg-primary/10 text-primary">
          <Icon className="size-5" />
        </div>
      </div>
    </div>
  );
}

function StatusPill({ status }: { status: CredentialRequestStatus }) {
  const normalizedStatus = normalize(status);
  const approved = ["approved", "completed", "released"].includes(
    normalizedStatus,
  );
  const rejected = normalizedStatus === "rejected";
  const label =
    normalizedStatus === "pending" || normalizedStatus === "pending_review"
      ? "Pending Review"
      : approved
        ? "Approved"
        : status || "Pending Review";

  return (
    <span
      className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-semibold ${
        approved
          ? "bg-emerald-50 text-emerald-700"
          : rejected
            ? "bg-red-50 text-red-700"
            : "bg-amber-50 text-amber-700"
      }`}
    >
      {approved ? (
        <CheckCircle2 className="size-3.5" />
      ) : (
        <Clock3 className="size-3.5" />
      )}

      <span className="capitalize">{String(label).replaceAll("_", " ")}</span>
    </span>
  );
}