import { useEffect, useMemo, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { BarChart3, BookOpen } from "lucide-react";
import { Skeleton } from "@/components/ui/skeleton";
import { supabase } from "@/integrations/supabase/client";
import {
  type ActivityScore,
  type GradeActivity,
  type GradeComponent,
  type StudentRow,
} from "@/lib/data";
import type {
  AnalyticsFeatureVector,
  NaiveBayesForecastUiData,
  NaiveBayesPredictionInput,
  NaiveBayesTrainingRow,
} from "@/lib/analytics-types";
import {
  buildNaiveBayesForecastUiData,
  MIN_NAIVE_BAYES_TRAINING_ROWS,
} from "@/lib/forecast-utils";
import {
  PROFICIENCY_DOT_COLORS,
  proficiencyRangeLabel,
} from "@/lib/proficiency";

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

function roundInitialGrade(initialGrade: number | null | undefined): number | null {
  if (typeof initialGrade !== "number" || Number.isNaN(initialGrade)) {
    return null;
  }

  const grade = Math.max(0, Math.min(100, initialGrade));
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

/* ---------------- Analytics ---------------- */
type AnalyticsTerm = "1" | "2" | "3" | "final";

type ProficiencyLevel =
  | "Advancing"
  | "Benchmarking"
  | "Connecting"
  | "Developing"
  | "Emerging";

type ProficiencyDistribution = Record<ProficiencyLevel, number>;

type ProficiencySeries = {
  term: AnalyticsTerm;
  label: string;
  distribution: ProficiencyDistribution;
  className: string;
};


const PROFICIENCY_LEVELS: Array<{
  key: ProficiencyLevel;
  short: string;
  cardClass: string;
  valueClass: string;
}> = [
  {
    key: "Advancing",
    short: "90–100",
    cardClass: "border-emerald-200 bg-emerald-50/55",
    valueClass: "text-emerald-600",
  },
  {
    key: "Benchmarking",
    short: "80–89",
    cardClass: "border-blue-200 bg-blue-50/55",
    valueClass: "text-blue-600",
  },
  {
    key: "Connecting",
    short: "75–79",
    cardClass: "border-amber-200 bg-amber-50/55",
    valueClass: "text-amber-600",
  },
  {
    key: "Developing",
    short: "65–74",
    cardClass: "border-orange-200 bg-orange-50/55",
    valueClass: "text-orange-600",
  },
  {
    key: "Emerging",
    short: "0–64",
    cardClass: "border-red-200 bg-red-50/55",
    valueClass: "text-red-600",
  },
];

function emptyProficiencyDistribution(): ProficiencyDistribution {
  return {
    Advancing: 0,
    Benchmarking: 0,
    Connecting: 0,
    Developing: 0,
    Emerging: 0,
  };
}

function proficiencyForGrade(grade: number): ProficiencyLevel {
  if (grade >= 90) return "Advancing";
  if (grade >= 80) return "Benchmarking";
  if (grade >= 75) return "Connecting";
  if (grade >= 65) return "Developing";
  return "Emerging";
}

function analyticsTermLabel(term: AnalyticsTerm) {
  if (term === "1") return "1st Term";
  if (term === "2") return "2nd Term";
  if (term === "3") return "3rd Term";
  return "Final";
}

export function AnalyticsInsightsPanel({
  classId,
  students,
  subject,
  subjectOptions = [],
  onSubjectChange,
  summarySubjectGrades,
  allowComponentForecast = true,
  initialTerm = "1",
  adminView = false,
  classLabel,
  teacherLabel,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  /** Teacher-only subject selection; Admin keeps its existing class directory UI. */
  subjectOptions?: readonly string[];
  onSubjectChange?: (subject: string) => void;
  /**
   * Actual saved grades for a different subject in Summary of Grades.
   * Never substitute the selected class's WW/PT/STE scores for another subject.
   */
  summarySubjectGrades?: Record<
    string,
    Partial<Record<AnalyticsTerm, number | null>>
  >;
  /** False if this subject does not have its own matching component records. */
  allowComponentForecast?: boolean;
  initialTerm?: AnalyticsTerm;
  /**
   * Deprecated compatibility prop from the previous Random Forest version.
   * It is intentionally ignored by the Naive Bayes flow so older callers can
   * still compile while they are being migrated.
   */
  forecastData?: unknown;
  /**
   * Admin mode reuses the exact same grade/proficiency calculations and the
   * same Naive Bayes Term 3 forecast while keeping the view read-only.
   */
  adminView?: boolean;
  classLabel?: string;
  teacherLabel?: string;
}) {
  const [selectedTerm, setSelectedTerm] = useState<AnalyticsTerm>(initialTerm);

  useEffect(() => {
    setSelectedTerm(initialTerm);
  }, [initialTerm]);
  const isMapeh = normalize(subject) === "mapeh";

  const storageTerms = useMemo(
    () =>
      isMapeh
        ? ["1", "PEH_T1", "MA_T2", "PEH_T2", "MA_T3", "PEH_T3"]
        : ["1", "2", "3"],
    [isMapeh],
  );

  const { data: activities = [], isLoading: activitiesLoading } = useQuery({
    queryKey: ["analytics-activities", classId, storageTerms.join("|")],
    enabled: Boolean(classId && allowComponentForecast),
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

  const { data: components = [], isLoading: componentsLoading } = useQuery({
    queryKey: ["analytics-components", classId, storageTerms.join("|")],
    enabled: Boolean(classId && allowComponentForecast),
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

  const activityIds = useMemo(() => activities.map((activity) => activity.id), [activities]);

  const { data: scores = [], isLoading: scoresLoading } = useQuery({
    queryKey: ["analytics-scores", classId, activityIds.join("|")],
    enabled: Boolean(classId && allowComponentForecast && activityIds.length),
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

  const { data: savedTermGradeBases = [], isLoading: gradesLoading } = useQuery({
    queryKey: ["analytics-term-grade-bases", classId, subject],
    enabled: Boolean(classId && allowComponentForecast),
    queryFn: async () => {
      const { data, error } = await (supabase as any)
        .from("grades")
        .select("student_id, subject, term, term_grade_base")
        .eq("class_id", classId);
      if (error) throw error;
      return (data ?? []) as Array<{
        student_id: string;
        subject: string | null;
        term: string;
        term_grade_base: number | null;
      }>;
    },
  });

  const isLoading = activitiesLoading || componentsLoading || scoresLoading || gradesLoading;

  const scoreMap = useMemo(() => {
    const map = new Map<string, number | null>();
    scores.forEach((score) => map.set(`${score.activity_id}|${score.student_id}`, score.score));
    return map;
  }, [scores]);

  const savedBaseMap = useMemo(() => {
    const map = new Map<string, number>();
    savedTermGradeBases.forEach((row) => {
      if (row.term_grade_base == null) return;
      const value = Number(row.term_grade_base);
      if (Number.isFinite(value)) {
        map.set(`${row.student_id}|${row.term}|${normalize(row.subject)}`, value);
      }
    });
    return map;
  }, [savedTermGradeBases]);

  const termScopes = useMemo(() => {
    if (!isMapeh) {
      return {
        "1": [{ storageTerm: "1", scopeLabel: "" }],
        "2": [{ storageTerm: "2", scopeLabel: "" }],
        "3": [{ storageTerm: "3", scopeLabel: "" }],
        final: [],
      } as Record<AnalyticsTerm, Array<{ storageTerm: string; scopeLabel: string }>>;
    }

    return {
      "1": [
        { storageTerm: "1", scopeLabel: "MA" },
        { storageTerm: "PEH_T1", scopeLabel: "PEH" },
      ],
      "2": [
        { storageTerm: "MA_T2", scopeLabel: "MA" },
        { storageTerm: "PEH_T2", scopeLabel: "PEH" },
      ],
      "3": [
        { storageTerm: "MA_T3", scopeLabel: "MA" },
        { storageTerm: "PEH_T3", scopeLabel: "PEH" },
      ],
      final: [],
    } as Record<AnalyticsTerm, Array<{ storageTerm: string; scopeLabel: string }>>;
  }, [isMapeh]);

  const currentScopes = termScopes[selectedTerm];

  const activitiesForStorageTerm = (
    storageTerm: string,
    component: "WW" | "PT" | "QA",
  ) =>
    activities
      .filter((activity) => activity.term === storageTerm && activity.component === component)
      .sort((a, b) => a.position - b.position)
      .slice(0, component === "QA" ? 3 : undefined);

  const componentWeight = (storageTerm: string, component: "WW" | "PT" | "QA") =>
    components.find((row) => row.term === storageTerm && row.component === component)?.weight ??
    (component === "WW" ? 20 : component === "PT" ? 50 : 30);

  const componentPercentageForStorageTerm = (
    studentId: string,
    storageTerm: string,
    component: "WW" | "PT" | "QA",
  ): number | null => {
    const componentActivities = activitiesForStorageTerm(
      storageTerm,
      component,
    );

    if (componentActivities.length === 0) return null;

    let raw = 0;
    let hps = 0;
    let hasScore = false;

    if (component === "QA") {
      const qaPartWeights = [30, 30, 40] as const;
      let weightedPercentage = 0;
      let usedWeight = 0;

      componentActivities.forEach((activity, index) => {
        const score = scoreMap.get(`${activity.id}|${studentId}`);
        const activityHps = Number(activity.hps || 0);

        if (typeof score !== "number" || activityHps <= 0) return;

        const partWeight = qaPartWeights[index] ?? 0;
        weightedPercentage +=
          (score / activityHps) * 100 * (partWeight / 100);
        usedWeight += partWeight;
        hasScore = true;
      });

      if (!hasScore || usedWeight <= 0) return null;

      // The current E-Class Record stores the summative test/exam component
      // under QA. Analytics exposes it as STE without changing database keys.
      return Math.max(
        0,
        Math.min(100, weightedPercentage * (100 / usedWeight)),
      );
    }

    componentActivities.forEach((activity) => {
      const score = scoreMap.get(`${activity.id}|${studentId}`);
      const activityHps = Number(activity.hps || 0);

      if (typeof score !== "number" || activityHps <= 0) return;

      raw += score;
      hps += activityHps;
      hasScore = true;
    });

    if (!hasScore || hps <= 0) return null;

    return Math.max(0, Math.min(100, (raw / hps) * 100));
  };

  const storageTermGrade = (
    studentId: string,
    storageTerm: string,
    scopeLabel: string,
    analyticsTerm: "1" | "2" | "3",
  ): number | null => {
    const weightedParts = (["WW", "PT", "QA"] as const).map((component) => {
      const componentActivities = activitiesForStorageTerm(storageTerm, component);
      let raw = 0;
      let hps = 0;
      let hasScore = false;
      let qaPercentageScore = 0;
      const qaPartWeights = [30, 30, 40] as const;

      componentActivities.forEach((activity, index) => {
        const score = scoreMap.get(`${activity.id}|${studentId}`);
        if (typeof score !== "number") return;
        raw += score;
        hps += Number(activity.hps || 0);
        hasScore = true;
        if (component === "QA" && Number(activity.hps) > 0) {
          qaPercentageScore +=
            (score / Number(activity.hps)) * 100 * ((qaPartWeights[index] ?? 0) / 100);
        }
      });

      if (!hasScore || hps <= 0) return null;
      const percentageScore = component === "QA" ? qaPercentageScore : (raw / hps) * 100;
      return percentageScore * (componentWeight(storageTerm, component) / 100);
    });

    if (weightedParts.some((part) => part == null)) return null;

    const initialGrade =
      Math.round((weightedParts as number[]).reduce((sum, value) => sum + value, 0) * 100) / 100;
    const computedBase = roundInitialGrade(initialGrade);
    if (computedBase == null) return null;

    if (!isMapeh) {
      const saved = savedBaseMap.get(`${studentId}|${storageTerm}|${normalize(subject)}`);
      return resolveTermGradeBase(computedBase, saved);
    }

    const expectedSubject = normalize(`${subject}_${scopeLabel}_T${analyticsTerm}_ALL`);
    const saved = savedTermGradeBases.find(
      (row) =>
        row.student_id === studentId &&
        normalize(row.subject) === expectedSubject &&
        row.term_grade_base != null,
    );

    return resolveTermGradeBase(
      computedBase,
      saved?.term_grade_base == null ? undefined : Number(saved.term_grade_base),
    );
  };

  const gradeForAnalyticsTerm = (
    studentId: string,
    analyticsTerm: "1" | "2" | "3",
  ): number | null => {
    const scopes = termScopes[analyticsTerm];
    const grades = scopes.map((scope) =>
      storageTermGrade(
        studentId,
        scope.storageTerm,
        scope.scopeLabel,
        analyticsTerm,
      ),
    );

    if (!isMapeh) return grades[0] ?? null;
    if (grades.some((grade) => grade == null)) return null;

    return Math.round(
      (grades as number[]).reduce((sum, grade) => sum + grade, 0) /
        grades.length,
    );
  };

  const featuresForAnalyticsTerm = (
    studentId: string,
    analyticsTerm: "1" | "2",
  ): AnalyticsFeatureVector | null => {
    const scopes = termScopes[analyticsTerm];

    const componentAverage = (
      component: "WW" | "PT" | "QA",
    ): number | null => {
      const values = scopes.map((scope) =>
        componentPercentageForStorageTerm(
          studentId,
          scope.storageTerm,
          component,
        ),
      );

      if (
        values.length === 0 ||
        values.some((value) => typeof value !== "number")
      ) {
        return null;
      }

      return (
        (values as number[]).reduce((sum, value) => sum + value, 0) /
        values.length
      );
    };

    const ww = componentAverage("WW");
    const pt = componentAverage("PT");
    const ste = componentAverage("QA");
    const finalGrade = gradeForAnalyticsTerm(studentId, analyticsTerm);

    if (
      ww == null ||
      pt == null ||
      ste == null ||
      finalGrade == null
    ) {
      return null;
    }

    return {
      ww: Math.round(ww * 100) / 100,
      pt: Math.round(pt * 100) / 100,
      ste: Math.round(ste * 100) / 100,
      finalGrade,
    };
  };

  const studentAnalyticsName = (student: StudentRow) => {
    const middle = String(student.middle_name ?? "").trim();
    return `${student.last_name}, ${student.first_name}${
      middle ? ` ${middle}` : ""
    }`.trim();
  };

  const selectedTermGrade = (studentId: string): number | null => {
    if (selectedTerm !== "final") {
      return gradeForAnalyticsTerm(studentId, selectedTerm);
    }

    const termGrades = (["1", "2", "3"] as const).map((term) =>
      gradeForAnalyticsTerm(studentId, term),
    );

    if (termGrades.some((grade) => grade == null)) return null;

    return Math.round(
      (termGrades as number[]).reduce((sum, grade) => sum + grade, 0) /
        termGrades.length,
    );
  };

  const gradeForSelectedAnalyticsTerm = (
    studentId: string,
    analyticsTerm: AnalyticsTerm,
  ): number | null => {
    if (summarySubjectGrades) {
      const savedGrade = summarySubjectGrades[studentId]?.[analyticsTerm];
      return typeof savedGrade === "number" && Number.isFinite(savedGrade)
        ? savedGrade
        : null;
    }

    if (analyticsTerm !== "final") {
      return gradeForAnalyticsTerm(studentId, analyticsTerm);
    }

    const termGrades = (["1", "2", "3"] as const).map((term) =>
      gradeForAnalyticsTerm(studentId, term),
    );

    if (termGrades.some((grade) => grade == null)) return null;

    return Math.round(
      (termGrades as number[]).reduce((sum, grade) => sum + grade, 0) /
        termGrades.length,
    );
  };

  const selectedTermGrades = students
    .map((student) => ({
      student,
      grade: gradeForSelectedAnalyticsTerm(student.id, selectedTerm),
    }))
    .filter(
      (row): row is { student: StudentRow; grade: number } =>
        typeof row.grade === "number",
    );

  const classAverage = selectedTermGrades.length
    ? Math.round(
        (selectedTermGrades.reduce((sum, row) => sum + row.grade, 0) /
          selectedTermGrades.length) *
          100,
      ) / 100
    : null;

  const highestGrade = selectedTermGrades.length
    ? Math.max(...selectedTermGrades.map((row) => row.grade))
    : null;

  const lowestGrade = selectedTermGrades.length
    ? Math.min(...selectedTermGrades.map((row) => row.grade))
    : null;

  const passCount = selectedTermGrades.filter((row) => row.grade >= 75).length;
  const failCount = selectedTermGrades.filter((row) => row.grade < 75).length;

  const distributionForTerm = (analyticsTerm: AnalyticsTerm) => {
    const distribution = emptyProficiencyDistribution();
    let graded = 0;

    students.forEach((student) => {
      const grade = gradeForSelectedAnalyticsTerm(student.id, analyticsTerm);
      if (typeof grade !== "number") return;

      distribution[proficiencyForGrade(grade)] += 1;
      graded += 1;
    });

    return { distribution, graded };
  };

  const selectedDistribution = distributionForTerm(selectedTerm);

  const term2Distribution = distributionForTerm("2");

  const naiveBayesTrainingRows: NaiveBayesTrainingRow[] = allowComponentForecast
    ? students.flatMap((student) => {
      const features = featuresForAnalyticsTerm(student.id, "1");
      const term2Grade = gradeForAnalyticsTerm(student.id, "2");

      if (!features || term2Grade == null) return [];

      return [
        {
          studentId: student.id,
          studentName: studentAnalyticsName(student),
          features,
          label: proficiencyForGrade(term2Grade),
        },
      ];
    })
    : [];

  const naiveBayesPredictionInputs: NaiveBayesPredictionInput[] =
    allowComponentForecast ? students.flatMap((student) => {
      const features = featuresForAnalyticsTerm(student.id, "2");
      const term2Grade = gradeForAnalyticsTerm(student.id, "2");

      if (!features || term2Grade == null) return [];

      return [
        {
          studentId: student.id,
          studentName: studentAnalyticsName(student),
          features,
          actualTerm2Category: proficiencyForGrade(term2Grade),
        },
      ];
    }) : [];

  const naiveBayesForecast = buildNaiveBayesForecastUiData({
    trainingRows: naiveBayesTrainingRows,
    predictionInputs: naiveBayesPredictionInputs,
    actualTerm2Distribution: term2Distribution.distribution,
  });

  // The chart is cumulative by the selected tab:
  // 1st Term -> Term 1 only
  // 2nd Term -> Term 1 + Term 2
  // 3rd Term -> Term 1 + Term 2 + Term 3
  // Final -> Term 1 + Term 2 + Term 3 + Final
  const visibleTerms: AnalyticsTerm[] =
    selectedTerm === "1"
      ? ["1"]
      : selectedTerm === "2"
        ? ["1", "2"]
        : selectedTerm === "3"
          ? ["1", "2", "3"]
          : ["1", "2", "3", "final"];

  const seriesClass: Record<AnalyticsTerm, string> = {
    "1": "bg-emerald-600",
    "2": "bg-sky-700",
    "3": "bg-amber-500",
    final: "bg-primary",
  };

  const comparisonSeries: ProficiencySeries[] = visibleTerms.map((term) => ({
    term,
    label: `${analyticsTermLabel(term)} Actual`,
    distribution: distributionForTerm(term).distribution,
    className: seriesClass[term],
  }));

  const expectedScores =
    selectedTerm === "final"
      ? students.length * 3
      : currentScopes.reduce((total, scope) => {
          const scopeActivities = (["WW", "PT", "QA"] as const).flatMap(
            (component) =>
              activitiesForStorageTerm(scope.storageTerm, component),
          );
          return total + scopeActivities.length * students.length;
        }, 0);

  const enteredScores =
    selectedTerm === "final"
      ? students.reduce((total, student) => {
          const completed = (["1", "2", "3"] as const).filter(
            (term) => gradeForAnalyticsTerm(student.id, term) != null,
          ).length;
          return total + completed;
        }, 0)
      : currentScopes.reduce((total, scope) => {
          const scopeActivities = (["WW", "PT", "QA"] as const).flatMap(
            (component) =>
              activitiesForStorageTerm(scope.storageTerm, component),
          );

          return (
            total +
            scopeActivities.reduce(
              (scoreTotal, activity) =>
                scoreTotal +
                students.filter(
                  (student) =>
                    typeof scoreMap.get(`${activity.id}|${student.id}`) ===
                    "number",
                ).length,
              0,
            )
          );
        }, 0);

  const missingScores = summarySubjectGrades
    ? Math.max(0, students.length - selectedTermGrades.length)
    : Math.max(0, expectedScores - enteredScores);
  const readiness = summarySubjectGrades
    ? (students.length > 0
        ? Math.round((selectedTermGrades.length / students.length) * 100)
        : 0)
    : expectedScores > 0
      ? Math.round((enteredScores / expectedScores) * 100)
      : selectedTermGrades.length > 0
        ? Math.round((selectedTermGrades.length / Math.max(students.length, 1)) * 100)
        : 0;

  const termLabel = analyticsTermLabel(selectedTerm);

  const comparisonTitle =
    selectedTerm === "1"
      ? "1st Term Proficiency Distribution"
      : selectedTerm === "2"
        ? "1st Term vs 2nd Term"
        : selectedTerm === "3"
          ? "1st Term vs 2nd Term vs 3rd Term"
          : "Term 1–3 vs Final";

  const comparisonDescription =
    selectedTerm === "1"
      ? "Actual proficiency distribution for the selected subject in 1st Term."
      : `Cumulative comparison up to ${termLabel} for the selected subject.`;

  if (isLoading) return <AnalyticsHubSkeleton />;

  return (
    <div className="space-y-4">
      {adminView && (
        <div className="flex flex-col gap-2 rounded-2xl border border-emerald-200 bg-emerald-50/55 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="min-w-0">
            <div className="text-[10px] font-bold uppercase tracking-wide text-emerald-700">
              Admin read-only analytics
            </div>
            <div className="mt-0.5 truncate text-sm font-bold text-emerald-950">
              {classLabel || "Selected Class"}
              {subject ? ` · ${subject}` : ""}
            </div>
            {teacherLabel ? (
              <div className="mt-0.5 truncate text-xs text-emerald-800/80">
                Teacher: {teacherLabel}
              </div>
            ) : null}
          </div>
          <span className="inline-flex w-fit rounded-full border border-emerald-200 bg-white px-3 py-1 text-[10px] font-semibold text-emerald-700">
            All-class access
          </span>
        </div>
      )}

      <div className="rounded-2xl border bg-card p-3 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 flex-1 flex-wrap items-center gap-3">
            <div className="inline-flex items-center gap-2 rounded-xl border border-emerald-600 bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm">
              <BookOpen className="size-4" />
              {adminView ? "Admin Analytics" : "Grades Hub"}
            </div>

            {!adminView && subjectOptions.length > 1 && onSubjectChange && (
              <label className="flex min-w-[220px] flex-1 flex-wrap items-center gap-2 text-xs font-semibold text-[#633d33] sm:flex-initial">
                <BookOpen className="size-4 text-[#8f3b34]" aria-hidden="true" />
                Subject
                <select
                  aria-label="Select analytics subject"
                  value={subject}
                  onChange={(event) => onSubjectChange(event.target.value)}
                  className="h-10 min-w-0 flex-1 rounded-xl border border-[#c88b84] bg-white px-3 text-xs font-semibold text-[#3f2c28] shadow-sm outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-200 sm:w-[210px] sm:flex-none"
                >
                  {subjectOptions.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </label>
            )}

            {!adminView && subject && (
              <div className="min-w-0 rounded-xl border border-emerald-100 bg-emerald-50 px-3 py-2 text-xs text-emerald-900">
                <span className="font-medium text-emerald-700">Selected Subject: </span>
                <span className="font-bold">{subject}</span>
                <div className="mt-0.5 text-[10px] text-emerald-700">
                  Analytics below are for this subject only.
                </div>
              </div>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {(["1", "2", "3", "final"] as AnalyticsTerm[]).map((term) => (
              <button
                key={term}
                type="button"
                onClick={() => setSelectedTerm(term)}
                className={`rounded-lg border px-3 py-1.5 text-xs font-semibold transition ${
                  selectedTerm === term
                    ? "border-primary bg-primary text-primary-foreground"
                    : "bg-background text-muted-foreground hover:bg-muted"
                }`}
              >
                {term === "1"
                  ? "1st Term"
                  : term === "2"
                    ? "2nd Term"
                    : term === "3"
                      ? "3rd Term"
                      : "Final"}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="overflow-hidden rounded-2xl border border-emerald-200 bg-card shadow-sm">
        <div className="border-t-4 border-emerald-600 px-4 pb-5 pt-4 sm:px-5">
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-6">
            <AnalyticsMetric
              label="Class Average"
              value={formatAnalyticsNumber(classAverage)}
            />
            <AnalyticsMetric
              label="Highest Grade"
              value={formatAnalyticsNumber(highestGrade)}
              tone="success"
            />
            <AnalyticsMetric
              label="Lowest Grade"
              value={formatAnalyticsNumber(lowestGrade)}
              tone="danger"
            />
            <AnalyticsMetric
              label="Pass / Fail"
              value={`${passCount} / ${failCount}`}
            />
            <AnalyticsMetric
              label="Missing Scores"
              value={String(missingScores)}
              tone={missingScores ? "danger" : "success"}
            />
            <AnalyticsMetric
              label="Readiness"
              value={`${readiness}%`}
              tone={readiness === 100 ? "success" : "info"}
            />
          </div>

          <div className="mt-5 rounded-2xl border bg-background/70 p-3 sm:p-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <div>
                <div className="text-[11px] font-bold uppercase tracking-wide text-muted-foreground">
                  Proficiency Distribution ({termLabel})
                </div>
                <div className="mt-1 text-xs text-muted-foreground">
                  Based on the selected subject's actual term grades.
                </div>
              </div>

              <span className="rounded-full border bg-card px-3 py-1 text-[10px] font-semibold text-muted-foreground">
                {selectedDistribution.graded}/{students.length} graded
              </span>
            </div>

            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
              {PROFICIENCY_LEVELS.map((level) => (
                <div
                  key={level.key}
                  className={`rounded-xl border p-3 ${level.cardClass}`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className={`text-sm font-semibold ${level.valueClass}`}>
                      {level.key}
                    </span>
                    <span className="text-[10px] font-medium text-muted-foreground">
                      {level.short}
                    </span>
                  </div>
                  <div className={`mt-2 text-3xl font-extrabold ${level.valueClass}`}>
                    {selectedDistribution.distribution[level.key]}
                  </div>
                  <div className="mt-1 text-[10px] text-muted-foreground">
                    student
                    {selectedDistribution.distribution[level.key] === 1 ? "" : "s"}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Full-width analytics panels.
              Term analytics stay cumulative by the selected tab. The forecast
              flow itself is fixed: T1 features -> T2 proficiency label ->
              Gaussian Naive Bayes -> T2 features -> predicted T3 category. */}
          <div className="mt-5 space-y-4">
            <div className="w-full">
              <ProficiencyComparisonChart
                animationKey={`proficiency-${selectedTerm}`}
                title={comparisonTitle}
                description={comparisonDescription}
                series={comparisonSeries}
              />
            </div>

            <div className="w-full">
              <NaiveBayesForecastPanel
                forecastData={naiveBayesForecast}
                actualTerm2Distribution={term2Distribution.distribution}
                trainingRowCount={naiveBayesTrainingRows.length}
                predictionInputCount={naiveBayesPredictionInputs.length}
                totalStudents={students.length}
                unavailableReason={
                  allowComponentForecast
                    ? undefined
                    : "Only this subject's actual Summary grades are displayed. Its own comparable Term 1 and Term 2 WW, PT and STE inputs are not available for forecasting in this view. Another subject's scores are never used."
                }
              />
            </div>
          </div>

          <div className="mt-4 grid gap-3 xl:grid-cols-2">
            <div className="rounded-xl border border-sky-200 bg-sky-50/70 px-4 py-3 text-xs text-sky-950">
              <div className="font-bold text-sky-900">How to Read This</div>
              <div className="mt-1 leading-relaxed">
                The model learns from each learner&apos;s Term 1 WW, PT, STE,
                and Final Grade together with the learner&apos;s actual Term 2
                proficiency category. It then applies that trained relationship
                to Term 2 features to forecast the Term 3 category.
              </div>
            </div>

            <div className="rounded-xl border border-amber-200 bg-amber-50/70 px-4 py-3 text-xs text-amber-950">
              <div className="font-bold text-amber-900">Model Note</div>
              <div className="mt-1 leading-relaxed">
                {!allowComponentForecast ? (
                  <>This view displays this subject&apos;s actual Summary grades only.
                  Term 3 prediction requires comparable WW, PT, and STE data
                  for the same subject across Term 1 and Term 2.</>
                ) : (
                  <>A forecast is shown only when at least{" "}
                {MIN_NAIVE_BAYES_TRAINING_ROWS} learners have complete Term 1
                features and Term 2 labels. Full validation uses 5-fold
                cross-validation. Missing E-Class Record inputs are never
                replaced with invented values.</>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ProficiencyComparisonChart({
  animationKey,
  title,
  description,
  series,
}: {
  animationKey: string;
  title: string;
  description: string;
  series: ProficiencySeries[];
}) {
  const [animateBars, setAnimateBars] = useState(false);

  const signature = series
    .map((item) =>
      [
        item.label,
        ...PROFICIENCY_LEVELS.map(
          (level) => item.distribution[level.key],
        ),
      ].join(":"),
    )
    .join("|");

  useEffect(() => {
    setAnimateBars(false);
    const timer = window.setTimeout(() => setAnimateBars(true), 60);
    return () => window.clearTimeout(timer);
  }, [animationKey, signature]);

  const maxValue = Math.max(
    1,
    ...series.flatMap((item) =>
      PROFICIENCY_LEVELS.map((level) => item.distribution[level.key]),
    ),
  );

  const chartMax = Math.max(5, Math.ceil(maxValue / 5) * 5);
  const tickValues = Array.from({ length: 6 }, (_, index) =>
    Math.round((chartMax / 5) * index),
  );

  return (
    <div className="h-full overflow-hidden rounded-2xl border border-emerald-200 bg-background/70 shadow-sm">
      <div className="flex flex-wrap items-start justify-between gap-3 border-b bg-muted/10 px-4 py-3">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-emerald-100 text-emerald-700">
            <BarChart3 className="size-4" />
          </div>
          <div className="min-w-0">
            <h4 className="font-bold text-emerald-700">{title}</h4>
            <p className="mt-0.5 text-xs text-muted-foreground">
              {description}
            </p>
          </div>
        </div>

        <span className="rounded-full border bg-card px-3 py-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
          Animated
        </span>
      </div>

      <div className="overflow-x-auto bg-gradient-to-b from-sky-50 via-white to-sky-50/70 px-4 pb-4 pt-5">
        <div className="relative h-[330px] min-w-[560px]">
          <div className="absolute bottom-[72px] left-0 top-2 w-10 text-[10px] text-muted-foreground">
            {[...tickValues].reverse().map((tick, index) => (
              <span
                key={`${tick}-${index}`}
                className="absolute right-1 -translate-y-1/2 tabular-nums"
                style={{ top: `${index * 20}%` }}
              >
                {tick}
              </span>
            ))}
          </div>

          <div className="absolute bottom-[72px] left-11 right-3 top-2">
            {tickValues.map((tick) => (
              <div
                key={tick}
                className="absolute left-0 right-0 border-t border-dashed border-sky-300/70"
                style={{ bottom: `${(tick / chartMax) * 100}%` }}
              />
            ))}

            <div className="absolute bottom-0 left-0 right-0 border-t-2 border-sky-500/60" />

            <div className="absolute inset-0 flex items-end justify-around gap-5 px-6">
              {PROFICIENCY_LEVELS.map((level) => (
                <div
                  key={level.key}
                  className="relative flex h-full min-w-[92px] flex-1 items-end justify-center gap-1.5"
                >
                  {series.map((item, seriesIndex) => {
                    const value = item.distribution[level.key];
                    const height = animateBars
                      ? Math.max(0, Math.min(100, (value / chartMax) * 100))
                      : 0;

                    return (
                      <div
                        key={`${item.term}-${level.key}`}
                        className="relative flex h-full w-9 items-end justify-center"
                        title={`${item.label} · ${level.key}: ${value} student${
                          value === 1 ? "" : "s"
                        }`}
                      >
                        <span className="absolute bottom-full mb-1 text-[11px] font-extrabold tabular-nums text-foreground">
                          {value}
                        </span>
                        <div
                          className={`w-full rounded-t-md shadow-sm ${item.className}`}
                          style={{
                            height: `${height}%`,
                            minHeight: value > 0 && animateBars ? "4px" : "0px",
                            transition:
                              "height 750ms cubic-bezier(0.22, 1, 0.36, 1)",
                            transitionDelay: `${seriesIndex * 70}ms`,
                          }}
                        />
                      </div>
                    );
                  })}

                  <div className="absolute -bottom-[52px] left-1/2 w-[96px] -translate-x-1/2 text-center text-[10px] font-semibold leading-3 text-foreground">
                    {level.key}
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="absolute bottom-0 left-11 right-3 flex flex-wrap items-center justify-center gap-x-5 gap-y-1 text-[10px] font-semibold text-muted-foreground">
            {series.map((item) => (
              <span key={item.term} className="inline-flex items-center gap-1.5">
                <span className={`size-2.5 rounded-sm ${item.className}`} />
                {item.label}
              </span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}


function formatPercent(value: number | null) {
  if (value == null || !Number.isFinite(value)) return "—";
  return `${Math.round(value * 100)}%`;
}

function probabilitySummary(
  probabilities: NaiveBayesForecastUiData["predictions"][number]["probabilities"],
) {
  return PROFICIENCY_LEVELS.map((level) => {
    const percent = Math.round((probabilities[level.key] ?? 0) * 100);
    return `${level.key}: ${percent}%`;
  }).join(" · ");
}

function NaiveBayesForecastPanel({
  forecastData,
  actualTerm2Distribution,
  trainingRowCount,
  predictionInputCount,
  totalStudents,
  unavailableReason,
}: {
  forecastData: NaiveBayesForecastUiData | null;
  actualTerm2Distribution: ProficiencyDistribution;
  trainingRowCount: number;
  predictionInputCount: number;
  totalStudents: number;
  unavailableReason?: string;
}) {
  const validationAccuracy =
    forecastData?.validation.accuracy ?? null;

  return (
    <div className="overflow-hidden rounded-2xl border border-violet-200 bg-gradient-to-b from-violet-50/70 via-background to-background shadow-sm">
      <div className="flex flex-wrap items-start justify-between gap-3 border-b border-violet-100 bg-violet-50/45 px-4 py-3">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-violet-100 text-violet-700">
            <BarChart3 className="size-4" />
          </div>
          <div className="min-w-0">
            <h4 className="font-bold text-violet-950">
              Gaussian Naive Bayes · Term 3 Proficiency Forecast
            </h4>
            <p className="mt-0.5 text-xs text-muted-foreground">
              Trains on Term 1 WW, PT, STE and Final Grade → actual Term 2
              proficiency, then predicts each learner&apos;s Term 3 category
              from Term 2 features.
            </p>
          </div>
        </div>

        <span className="rounded-full border border-violet-200 bg-white/80 px-3 py-1 text-[10px] font-semibold uppercase tracking-wide text-violet-700">
          Term 3 Forecast
        </span>
      </div>

      <div className="border-b border-violet-100 bg-white/60 px-4 py-3">
        <div className="grid gap-2 text-center sm:grid-cols-5">
          {[
            ["1", "Term 1 Features", "WW · PT · STE · Final"],
            ["2", "Term 2 Label", "Rule-based proficiency"],
            ["3", "Train Model", "Gaussian Naive Bayes"],
            ["4", "Validate", "5-fold cross-validation"],
            ["5", "Predict Term 3", "Term 2 features → category"],
          ].map(([step, title, detail]) => (
            <div
              key={step}
              className="rounded-xl border border-violet-100 bg-white px-3 py-2"
            >
              <div className="text-[9px] font-extrabold uppercase tracking-wide text-violet-500">
                Step {step}
              </div>
              <div className="mt-0.5 text-[11px] font-bold text-violet-950">
                {title}
              </div>
              <div className="mt-0.5 text-[9px] leading-4 text-muted-foreground">
                {detail}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="grid gap-3 p-4 sm:grid-cols-2 xl:grid-cols-4">
        <ModelMetric
          label="Training Rows"
          value={`${trainingRowCount}/${totalStudents}`}
          detail="Complete T1 → T2 learner rows"
        />
        <ModelMetric
          label="Prediction Inputs"
          value={`${predictionInputCount}/${totalStudents}`}
          detail="Complete Term 2 feature rows"
        />
        <ModelMetric
          label="5-Fold Accuracy"
          value={formatPercent(validationAccuracy)}
          detail={
            forecastData?.validation.foldsUsed === 5
              ? `${forecastData.validation.correct}/${forecastData.validation.sampleCount} held-out predictions correct`
              : `Needs at least ${MIN_NAIVE_BAYES_TRAINING_ROWS} complete training rows`
          }
        />
        <ModelMetric
          label="Learned Categories"
          value={String(forecastData?.learnedCategories.length ?? 0)}
          detail={
            forecastData?.learnedCategories.join(", ") ||
            "Waiting for complete training data"
          }
        />
      </div>

      {!forecastData ? (
        <div className="mx-4 mb-4 rounded-xl border border-dashed border-violet-200 bg-white/80 px-5 py-8 text-center">
          <div className="text-sm font-bold text-violet-950">
            Term 3 forecast is not ready yet
          </div>
          <div className="mx-auto mt-2 max-w-2xl text-xs leading-5 text-muted-foreground">
            {unavailableReason || (
              <>
                Complete both Term 1 and Term 2 E-Class Record inputs for at least{" "}
                {MIN_NAIVE_BAYES_TRAINING_ROWS} learners. Each usable learner needs
                WW, PT, STE and a final term grade. The model will then train,
                run 5-fold cross-validation and generate learner-level Term 3
                probabilities automatically.
              </>
            )}
          </div>
        </div>
      ) : (
        <>
          <div className="grid gap-4 px-4 pb-4 2xl:grid-cols-[minmax(0,1.55fr)_minmax(330px,0.75fr)]">
            <div className="min-w-0 overflow-hidden rounded-xl border bg-white/80">
              <div className="border-b px-4 py-3">
                <div className="text-sm font-bold text-foreground">
                  Student-Level Output
                </div>
                <div className="mt-0.5 text-[10px] text-muted-foreground">
                  Actual Term 2 category, predicted Term 3 category, confidence
                  and probability per category.
                </div>
              </div>

              <div className="max-h-[430px] overflow-auto">
                <table className="w-full min-w-[820px] border-collapse text-left">
                  <thead className="sticky top-0 z-10 bg-muted/95 text-[9px] font-bold uppercase tracking-wide text-muted-foreground">
                    <tr>
                      <th className="px-3 py-2.5">Student</th>
                      <th className="px-3 py-2.5">T2 Actual</th>
                      <th className="px-3 py-2.5">T3 Predicted</th>
                      <th className="px-3 py-2.5">Confidence</th>
                      <th className="px-3 py-2.5">Probability per Category</th>
                    </tr>
                  </thead>
                  <tbody>
                    {forecastData.predictions.map((prediction) => (
                      <tr
                        key={prediction.studentId}
                        className="border-t text-[11px]"
                      >
                        <td className="px-3 py-2.5 font-semibold text-foreground">
                          {prediction.studentName}
                        </td>
                        <td className="px-3 py-2.5">
                          <ProficiencyBadge
                            level={prediction.actualTerm2Category}
                          />
                        </td>
                        <td className="px-3 py-2.5">
                          <ProficiencyBadge
                            level={prediction.predictedTerm3Category}
                          />
                        </td>
                        <td className="px-3 py-2.5 font-bold tabular-nums">
                          {formatPercent(prediction.confidence)}
                        </td>
                        <td
                          className="max-w-[430px] px-3 py-2.5 text-[10px] leading-4 text-muted-foreground"
                          title={probabilitySummary(prediction.probabilities)}
                        >
                          {probabilitySummary(prediction.probabilities)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="space-y-4">
              <ForecastDistributionCard
                actualTerm2={actualTerm2Distribution}
                forecastTerm3={forecastData.distribution}
              />

              <div className="rounded-xl border bg-white/80 p-4">
                <div className="text-sm font-bold text-foreground">
                  Actionable Insights
                </div>
                <div className="mt-3 space-y-2">
                  {forecastData.insights.map((insight) => (
                    <div
                      key={insight.id}
                      className="rounded-lg border bg-muted/20 px-3 py-2 text-[11px] leading-5 text-foreground"
                    >
                      {insight.message}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </>
      )}

      <div className="flex flex-wrap items-center justify-center gap-x-4 gap-y-1 border-t border-violet-100 px-3 py-2 text-[9px] font-medium text-muted-foreground">
        {PROFICIENCY_LEVELS.map((level) => (
          <span key={level.key} className="inline-flex items-center gap-1.5">
            <span
              className="size-2.5 rounded-full"
              style={{ backgroundColor: PROFICIENCY_DOT_COLORS[level.key] }}
            />
            {level.key} ({proficiencyRangeLabel(level.key)})
          </span>
        ))}
      </div>
    </div>
  );
}

function ModelMetric({
  label,
  value,
  detail,
}: {
  label: string;
  value: string;
  detail: string;
}) {
  return (
    <div className="rounded-xl border bg-white/80 px-3 py-3">
      <div className="text-[9px] font-bold uppercase tracking-wide text-muted-foreground">
        {label}
      </div>
      <div className="mt-1 text-xl font-extrabold text-foreground">{value}</div>
      <div className="mt-1 line-clamp-2 text-[9px] leading-4 text-muted-foreground">
        {detail}
      </div>
    </div>
  );
}

function ProficiencyBadge({ level }: { level: ProficiencyLevel }) {
  const item =
    PROFICIENCY_LEVELS.find((entry) => entry.key === level) ??
    PROFICIENCY_LEVELS[PROFICIENCY_LEVELS.length - 1];

  return (
    <span
      className={`inline-flex rounded-full border px-2 py-1 text-[9px] font-bold ${item.cardClass} ${item.valueClass}`}
    >
      {level}
    </span>
  );
}

function ForecastDistributionCard({
  actualTerm2,
  forecastTerm3,
}: {
  actualTerm2: ProficiencyDistribution;
  forecastTerm3: ProficiencyDistribution;
}) {
  const maxValue = Math.max(
    1,
    ...PROFICIENCY_LEVELS.flatMap((level) => [
      actualTerm2[level.key],
      forecastTerm3[level.key],
    ]),
  );

  return (
    <div className="rounded-xl border bg-white/80 p-4">
      <div className="text-sm font-bold text-foreground">
        Section-Level Output
      </div>
      <div className="mt-0.5 text-[10px] text-muted-foreground">
        Actual Term 2 versus forecasted Term 3 proficiency distribution.
      </div>

      <div className="mt-4 space-y-3">
        {PROFICIENCY_LEVELS.map((level) => {
          const actual = actualTerm2[level.key];
          const forecast = forecastTerm3[level.key];

          return (
            <div key={level.key}>
              <div className="mb-1 flex items-center justify-between gap-3 text-[10px]">
                <span className={`font-bold ${level.valueClass}`}>
                  {level.key}
                </span>
                <span className="font-semibold tabular-nums text-muted-foreground">
                  T2 {actual} · T3 {forecast}
                </span>
              </div>

              <div className="grid gap-1.5">
                <div className="h-2.5 overflow-hidden rounded-full bg-slate-100">
                  <div
                    className="h-full rounded-full bg-slate-400 transition-all"
                    style={{ width: `${(actual / maxValue) * 100}%` }}
                  />
                </div>
                <div className="h-3 overflow-hidden rounded-full bg-violet-50">
                  <div
                    className="h-full rounded-full transition-all"
                    style={{
                      width: `${(forecast / maxValue) * 100}%`,
                      backgroundColor: PROFICIENCY_DOT_COLORS[level.key],
                    }}
                  />
                </div>
              </div>
            </div>
          );
        })}
      </div>

      <div className="mt-3 flex flex-wrap gap-3 border-t pt-3 text-[9px] text-muted-foreground">
        <span className="inline-flex items-center gap-1.5">
          <span className="h-2 w-4 rounded bg-slate-400" />
          Term 2 Actual
        </span>
        <span className="inline-flex items-center gap-1.5">
          <span className="h-2 w-4 rounded bg-violet-500" />
          Term 3 Forecast
        </span>
      </div>
    </div>
  );
}

function AnalyticsMetric({
  label,
  value,
  tone = "default",
}: {
  label: string;
  value: string;
  tone?: "default" | "success" | "danger" | "info";
}) {
  const valueClass =
    tone === "success"
      ? "text-emerald-600"
      : tone === "danger"
        ? "text-red-500"
        : tone === "info"
          ? "text-sky-500"
          : "text-foreground";
  return (
    <div className="min-w-0">
      <div className="text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">{label}</div>
      <div className={`mt-1 text-2xl font-extrabold ${valueClass}`}>{value}</div>
    </div>
  );
}

function formatAnalyticsNumber(value: number | null) {
  if (value == null || !Number.isFinite(value)) return "—";
  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}


function AnalyticsHubSkeleton() {
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-2 rounded-2xl border bg-card p-3">
        <Skeleton className="h-9 w-32 rounded-xl" />
        <div className="ml-auto flex gap-2">
          <Skeleton className="h-8 w-20 rounded-lg" />
          <Skeleton className="h-8 w-20 rounded-lg" />
          <Skeleton className="h-8 w-20 rounded-lg" />
          <Skeleton className="h-8 w-20 rounded-lg" />
        </div>
      </div>

      <div className="rounded-2xl border bg-card p-5">
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-6">
          {Array.from({ length: 6 }).map((_, index) => (
            <div key={index}>
              <Skeleton className="mb-2 h-3 w-20" />
              <Skeleton className="h-7 w-16" />
            </div>
          ))}
        </div>

        <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
          {Array.from({ length: 5 }).map((_, index) => (
            <Skeleton key={index} className="h-28 rounded-xl" />
          ))}
        </div>

        <Skeleton className="mt-5 h-80 w-full rounded-xl" />
      </div>
    </div>
  );
}
