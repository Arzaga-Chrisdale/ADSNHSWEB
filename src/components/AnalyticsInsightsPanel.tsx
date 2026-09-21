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
import type { MultiOutputForecastUiData } from "@/lib/analytics-types";
import {
  PROFICIENCY_AXIS_ORDER,
  PROFICIENCY_DOT_COLORS,
  proficiencyBandPosition,
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
  initialTerm = "1",
  forecastData = null,
  adminView = false,
  classLabel,
  teacherLabel,
}: {
  classId: string;
  students: StudentRow[];
  subject: string;
  initialTerm?: AnalyticsTerm;
  /**
   * Optional UI-ready forecast returned by the trained model/API.
   *
   * Existing callers do not need to pass this prop, so this update remains
   * backward compatible. When it is null, the Random Forest card still
   * renders but clearly shows that model output has not been connected yet.
   */
  forecastData?: MultiOutputForecastUiData | null;
  /**
   * Admin mode reuses the exact same grade/proficiency calculations while
   * showing a small read-only context banner for the class selected by Admin.
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

  const forecastSourceTerm =
    selectedTerm === "final" ? null : selectedTerm;

  const forecastTargetTerm =
    selectedTerm === "1"
      ? "2"
      : selectedTerm === "2"
        ? "3"
        : selectedTerm === "3"
          ? "final"
          : null;

  const activeForecast =
    forecastData &&
    forecastSourceTerm &&
    forecastTargetTerm &&
    forecastData.sourceTerm === forecastSourceTerm &&
    forecastData.targetTerm === forecastTargetTerm
      ? forecastData
      : null;

  const actualAtLeast88 = selectedTermGrades.filter(
    (row) => row.grade >= 88,
  ).length;

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

  const missingScores = Math.max(0, expectedScores - enteredScores);
  const readiness =
    expectedScores > 0
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
          <div className="flex flex-wrap gap-2">
            <div className="inline-flex items-center gap-2 rounded-xl border border-emerald-600 bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-sm">
              <BookOpen className="size-4" />
              {adminView ? "Admin Analytics" : "Grades Hub"}
            </div>
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
              Do not split the proficiency chart and Random Forest UI into
              two narrow columns. Each panel gets the full modal width. */}
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
              <RandomForestForecastPanel
                sourceTerm={forecastSourceTerm}
                targetTerm={forecastTargetTerm}
                actualDistribution={selectedDistribution.distribution}
                actualAtLeast88={actualAtLeast88}
                forecastData={activeForecast}
              />
            </div>
          </div>

          <div className="mt-4 grid gap-3 xl:grid-cols-2">
            <div className="rounded-xl border border-sky-200 bg-sky-50/70 px-4 py-3 text-xs text-sky-950">
              <div className="font-bold text-sky-900">How to Read This</div>
              <div className="mt-1 leading-relaxed">
                Each dot represents a student. The horizontal and vertical axes
                show the average grade range. Dots near the diagonal line
                indicate closer actual-versus-predicted grades.
              </div>
            </div>

            <div className="rounded-xl border border-amber-200 bg-amber-50/70 px-4 py-3 text-xs text-amber-950">
              <div className="font-bold text-amber-900">Note</div>
              <div className="mt-1 leading-relaxed">
                Forecast values must come from the trained Multi-Output Random
                Forest Regression model/API. This panel never invents forecast
                values when model output is unavailable.
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


type ForecastSourceTerm = "1" | "2" | "3";
type ForecastTargetTerm = "2" | "3" | "final";

function forecastTermLabel(term: ForecastTargetTerm | null) {
  if (term === "2") return "Term 2";
  if (term === "3") return "Term 3";
  if (term === "final") return "Final";
  return "Next Term";
}

function signedCount(value: number) {
  if (value > 0) return `+${value}`;
  return String(value);
}

function RandomForestForecastPanel({
  sourceTerm,
  targetTerm,
  actualDistribution,
  actualAtLeast88,
  forecastData,
}: {
  sourceTerm: ForecastSourceTerm | null;
  targetTerm: ForecastTargetTerm | null;
  actualDistribution: ProficiencyDistribution;
  actualAtLeast88: number;
  forecastData: MultiOutputForecastUiData | null;
}) {
  const targetLabel = forecastTermLabel(targetTerm);
  const hasForecast = Boolean(forecastData && targetTerm && sourceTerm);
  const modelLabel =
    forecastData?.modelLabel ?? "Multi-Output Random Forest Regression";

  return (
    <div className="h-full overflow-hidden rounded-2xl border border-violet-200 bg-gradient-to-b from-violet-50/70 via-background to-background shadow-sm">
      <div className="flex flex-wrap items-start justify-between gap-3 border-b border-violet-100 bg-violet-50/45 px-4 py-3">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-violet-100 text-violet-700">
            <BarChart3 className="size-4" />
          </div>
          <div className="min-w-0">
            <h4 className="font-bold text-violet-950">
              Multi-Output Random Forest Regression
            </h4>
            <p className="mt-0.5 text-xs text-muted-foreground">
              Predicts the number of students per proficiency level and 88+
              using the selected source term.
            </p>
          </div>
        </div>

        <span className="rounded-full border border-violet-200 bg-white/80 px-3 py-1 text-[10px] font-semibold uppercase tracking-wide text-violet-700">
          {targetTerm ? `${targetLabel} Forecast` : "Final Actual"}
        </span>
      </div>

      <div className="grid gap-4 p-4 xl:grid-cols-[minmax(0,2.7fr)_minmax(300px,0.85fr)]">
        <div className="min-w-0 rounded-xl border bg-white/75 p-5">
          <div className="text-base font-bold text-foreground">
            Actual vs Predicted {targetTerm ? `(${targetLabel})` : ""}
          </div>
          <div className="mt-2">
            <ForecastScatterPlot
              points={forecastData?.scatterPoints ?? []}
              hasForecast={hasForecast}
            />
          </div>
        </div>

        <div className="rounded-xl border bg-white/75 p-4">
          <div className="text-sm font-bold text-foreground">
            Predicted Distribution {targetTerm ? `(${targetLabel})` : ""}
          </div>

          <div className="mt-2 divide-y">
            {PROFICIENCY_LEVELS.map((level) => {
              const predicted = forecastData?.counts[level.key] ?? null;
              const difference =
                predicted == null
                  ? null
                  : predicted - actualDistribution[level.key];

              return (
                <div
                  key={level.key}
                  className="grid grid-cols-[1fr_auto_auto] items-center gap-2 py-2 text-[11px]"
                >
                  <div className="flex min-w-0 items-center gap-2">
                    <span
                      className="size-3 shrink-0 rounded-full"
                      style={{
                        backgroundColor:
                          PROFICIENCY_DOT_COLORS[level.key],
                      }}
                    />
                    <span className="truncate">
                      {level.key} ({level.short})
                    </span>
                  </div>

                  <span className="min-w-5 text-right text-sm font-extrabold tabular-nums">
                    {predicted ?? "—"}
                  </span>

                  <span
                    className={`min-w-8 text-right font-semibold tabular-nums ${
                      difference == null
                        ? "text-muted-foreground"
                        : difference > 0
                          ? "text-emerald-600"
                          : difference < 0
                            ? "text-red-500"
                            : "text-muted-foreground"
                    }`}
                  >
                    {difference == null ? "—" : `(${signedCount(difference)})`}
                  </span>
                </div>
              );
            })}
          </div>

          <div className="mt-2 grid grid-cols-[1fr_auto_auto] items-center gap-2 border-t pt-3 text-[11px]">
            <div className="font-bold text-foreground">★ ≥ 88 (Predicted)</div>
            <span className="min-w-5 text-right text-sm font-extrabold tabular-nums">
              {forecastData?.predictedAtLeast88 ?? "—"}
            </span>
            <span
              className={`min-w-8 text-right font-semibold tabular-nums ${
                forecastData?.predictedAtLeast88 == null
                  ? "text-muted-foreground"
                  : forecastData.predictedAtLeast88 - actualAtLeast88 > 0
                    ? "text-emerald-600"
                    : forecastData.predictedAtLeast88 - actualAtLeast88 < 0
                      ? "text-red-500"
                      : "text-muted-foreground"
              }`}
            >
              {forecastData?.predictedAtLeast88 == null
                ? "—"
                : `(${signedCount(
                    forecastData.predictedAtLeast88 - actualAtLeast88,
                  )})`}
            </span>
          </div>

          <div className="mt-3 text-[10px] leading-relaxed text-muted-foreground">
            {hasForecast
              ? `Values are forecasts based on ${analyticsTermLabel(
                  sourceTerm as AnalyticsTerm,
                )} data.`
              : targetTerm
                ? `Waiting for ${modelLabel} output for ${targetLabel}.`
                : "Final is an actual distribution; no next-term forecast is shown."}
          </div>
        </div>
      </div>

      <div className="flex flex-wrap items-center justify-center gap-x-4 gap-y-1 border-t border-violet-100 px-3 py-2 text-[9px] font-medium text-muted-foreground">
        {PROFICIENCY_AXIS_ORDER.map((level) => (
          <span key={level} className="inline-flex items-center gap-1.5">
            <span
              className="size-2.5 rounded-full"
              style={{ backgroundColor: PROFICIENCY_DOT_COLORS[level] }}
            />
            {level} ({proficiencyRangeLabel(level)})
          </span>
        ))}
      </div>
    </div>
  );
}

function ForecastScatterPlot({
  points,
  hasForecast,
}: {
  points: MultiOutputForecastUiData["scatterPoints"];
  hasForecast: boolean;
}) {
  const width = 620;
  const height = 390;
  const left = 78;
  const right = 24;
  const top = 24;
  const bottom = 64;
  const plotWidth = width - left - right;
  const plotHeight = height - top - bottom;

  const axisBands = PROFICIENCY_AXIS_ORDER.map((level, index) => ({
    level,
    label: proficiencyRangeLabel(level),
    position: (index + 0.5) / PROFICIENCY_AXIS_ORDER.length,
  }));

  const toX = (grade: number) => {
    const position = proficiencyBandPosition(grade) ?? 0;
    return left + position * plotWidth;
  };

  const toY = (grade: number) => {
    const position = proficiencyBandPosition(grade) ?? 0;
    return top + (1 - position) * plotHeight;
  };

  return (
    <div className="relative overflow-hidden rounded-lg bg-white">
      <svg
        viewBox={`0 0 ${width} ${height}`}
        className="h-[390px] w-full min-w-[520px] lg:h-[420px]"
        role="img"
        aria-label="Actual versus predicted average grade scatter plot"
      >
        {axisBands.map((band) => {
          const x = left + band.position * plotWidth;
          const y = top + (1 - band.position) * plotHeight;

          return (
            <g key={band.level}>
              <line
                x1={x}
                x2={x}
                y1={top}
                y2={top + plotHeight}
                stroke="#bae6fd"
                strokeDasharray="3 3"
              />
              <line
                x1={left}
                x2={left + plotWidth}
                y1={y}
                y2={y}
                stroke="#bae6fd"
                strokeDasharray="3 3"
              />
              <text
                x={x}
                y={height - 26}
                textAnchor="middle"
                fontSize="13"
                fill="#475569"
              >
                {band.label}
              </text>
              <text
                x={left - 7}
                y={y + 3}
                textAnchor="end"
                fontSize="13"
                fill="#475569"
              >
                {band.label}
              </text>
            </g>
          );
        })}

        <line
          x1={left}
          y1={top + plotHeight}
          x2={left + plotWidth}
          y2={top}
          stroke="#94a3b8"
          strokeDasharray="5 4"
        />

        <line
          x1={left}
          y1={top}
          x2={left}
          y2={top + plotHeight}
          stroke="#64748b"
        />
        <line
          x1={left}
          y1={top + plotHeight}
          x2={left + plotWidth}
          y2={top + plotHeight}
          stroke="#64748b"
        />

        {points.map((point, index) => (
          <circle
            key={`${point.studentId ?? "student"}-${index}`}
            cx={toX(point.actualGrade)}
            cy={toY(point.predictedGrade)}
            r="7"
            fill={PROFICIENCY_DOT_COLORS[point.level]}
            stroke="#ffffff"
            strokeWidth="1.5"
          >
            <title>
              {`Actual ${point.actualGrade} · Predicted ${point.predictedGrade} · ${point.level}`}
            </title>
          </circle>
        ))}

        <text
          x={left + plotWidth / 2}
          y={height - 5}
          textAnchor="middle"
          fontSize="13"
          fontWeight="600"
          fill="#334155"
        >
          Actual Average Grade
        </text>

        <text
          x="20"
          y={top + plotHeight / 2}
          textAnchor="middle"
          fontSize="13"
          fontWeight="600"
          fill="#334155"
          transform={`rotate(-90 20 ${top + plotHeight / 2})`}
        >
          Predicted Average Grade
        </text>
      </svg>

      {!hasForecast && (
        <div className="pointer-events-none absolute left-1/2 top-1/2 w-[62%] -translate-x-1/2 -translate-y-1/2 rounded-xl border border-dashed bg-white/95 px-5 py-4 text-center text-sm font-medium leading-relaxed text-muted-foreground shadow-sm">
          Connect the trained model/API to display learner-level predicted
          points here.
        </div>
      )}
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
