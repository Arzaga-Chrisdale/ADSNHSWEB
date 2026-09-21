import { supabase } from "@/integrations/supabase/client";

export type TeacherType = "class_adviser" | "subject_teacher";
export type GradingPeriod = "1" | "2" | "3" | "final";
export type OverallRequestStatus = "pending" | "completed";
export type SubjectRequestStatus = "pending" | "completed" | "rejected";

export type GradeRequestWorkflowRow = {
  batch_id: string;
  adviser_id: string;
  adviser_name: string | null;
  adviser_email: string | null;
  advisory_class_id: string;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  grading_period: GradingPeriod;
  overall_status: OverallRequestStatus;
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
  subject_status: SubjectRequestStatus;
  subject_requested_at: string;
  submitted_at: string | null;
  teacher_note: string | null;
  submitted_grade_count: number;
};

export type GradeRequestHistoryRow = {
  id: string;
  batch_id: string;
  subject_request_id: string | null;
  actor_id: string | null;
  event_type: string;
  old_status: string | null;
  new_status: string | null;
  details: Record<string, unknown> | null;
  created_at: string;
};

export type GradeRequestScoreRow = {
  student_id: string;
  score: number;
};

export function gradingPeriodLabel(period: GradingPeriod | string) {
  if (period === "final") return "Final Grade";
  return `Term ${period}`;
}

export function requestStatusLabel(status: string) {
  if (status === "completed") return "Completed";
  if (status === "rejected") return "Rejected";
  if (status === "completed") return "Completed";
  return "Pending";
}

export function formatRequestDate(value: string | null | undefined) {
  if (!value) return "—";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";

  return new Intl.DateTimeFormat("en-PH", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

export async function getCurrentUserId() {
  const {
    data: { user },
    error,
  } = await supabase.auth.getUser();

  if (error) throw error;
  if (!user) throw new Error("You must be signed in.");

  return user.id;
}

export async function getCurrentTeacherType(): Promise<TeacherType> {
  const {
    data: { user },
    error: userError,
  } = await supabase.auth.getUser();

  if (userError) throw userError;
  if (!user) throw new Error("You must be signed in.");

  const metadataTeacherType = user.user_metadata?.teacher_type;

  const { data, error } = await (supabase as any)
    .from("profiles")
    .select("teacher_type")
    .eq("id", user.id)
    .maybeSingle();

  if (error) throw error;

  const profileTeacherType = data?.teacher_type;

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
}

export async function createGradeRequestBatch(
  advisoryClassId: string,
  gradingPeriod: GradingPeriod,
) {
  if (!advisoryClassId) {
    throw new Error("Select an advisory class.");
  }

  const { data, error } = await (supabase as any).rpc(
    "create_grade_request_batch",
    {
      p_advisory_class_id: advisoryClassId,
      p_grading_period: gradingPeriod,
    },
  );

  if (error) throw error;
  if (!data) throw new Error("The grade request could not be created.");

  return String(data);
}

export async function submitSubjectGrades(args: {
  subjectRequestId: string;
  scores: GradeRequestScoreRow[];
  teacherNote?: string;
}) {
  if (!args.subjectRequestId) {
    throw new Error("No subject request was selected.");
  }

  if (!Array.isArray(args.scores) || args.scores.length === 0) {
    throw new Error("Enter at least one learner grade before submitting.");
  }

  const invalidScore = args.scores.find(
    (item) =>
      !item.student_id ||
      !Number.isFinite(item.score) ||
      item.score < 0 ||
      item.score > 100,
  );

  if (invalidScore) {
    throw new Error("Every learner grade must be between 0 and 100.");
  }

  const { error } = await (supabase as any).rpc(
    "complete_grade_request_subject",
    {
      p_subject_request_id: args.subjectRequestId,
      p_scores: args.scores,
      p_teacher_note: args.teacherNote?.trim() || null,
    },
  );

  if (error) throw error;
}

export async function finalizeGradeRequestBatch(batchId: string) {
  if (!batchId) {
    throw new Error("No grade request batch was selected.");
  }

  const { error } = await (supabase as any).rpc(
    "finalize_grade_request_batch",
    {
      p_batch_id: batchId,
    },
  );

  if (error) throw error;
}
