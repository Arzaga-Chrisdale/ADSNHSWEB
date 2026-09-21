import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import type { GradingPeriod } from "@/lib/grade-request-workflow";

export function useGradeRequestFinalization(
  advisoryClassId: string,
  gradingPeriod: GradingPeriod,
) {
  return useQuery({
    enabled: Boolean(advisoryClassId && gradingPeriod),
    queryKey: [
      "grade-request-finalization-gate",
      advisoryClassId,
      gradingPeriod,
    ],
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grade_request_batches")
        .select(
          "id, status, is_finalized, requested_at, completed_at, finalized_at",
        )
        .eq("advisory_class_id", advisoryClassId)
        .eq("grading_period", gradingPeriod)
        .order("requested_at", { ascending: false })
        .limit(1)
        .maybeSingle();

      if (error) throw error;

      return {
        batchId: data?.id ? String(data.id) : null,
        status: data?.status === "completed" ? "completed" : "pending",
        isFinalized: Boolean(data?.is_finalized),
        requestedAt: data?.requested_at ?? null,
        completedAt: data?.completed_at ?? null,
        finalizedAt: data?.finalized_at ?? null,
        canGenerateReportCards:
          data?.status === "completed" && Boolean(data?.is_finalized),
      };
    },
  });
}
