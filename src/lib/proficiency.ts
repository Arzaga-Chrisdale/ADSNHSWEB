/**
 * SIGLA proficiency helpers.
 *
 * Proficiency bands:
 * - Advancing:    90–100
 * - Benchmarking: 80–89
 * - Connecting:   75–79
 * - Developing:   65–74
 * - Emerging:      0–64
 *
 * Rule-based suggested actions use the learner's ACTUAL current
 * proficiency classification. They are not ML-generated forecasts.
 */

export type ProficiencyLevel =
  | "Advancing"
  | "Benchmarking"
  | "Connecting"
  | "Developing"
  | "Emerging";

export type ProficiencyDistribution = Record<ProficiencyLevel, number>;

export type ProficiencyLevelDefinition = {
  key: ProficiencyLevel;
  min: number;
  max: number;
  short: string;
  cardClass: string;
  valueClass: string;
};

export type SuggestedAction = {
  level: ProficiencyLevel;
  title: string;
  actions: string[];
};

export const PROFICIENCY_LEVEL_ORDER: ProficiencyLevel[] = [
  "Advancing",
  "Benchmarking",
  "Connecting",
  "Developing",
  "Emerging",
];

/**
 * Axis order used by the Random Forest actual-vs-predicted scatter plot.
 * The plot reads from the lowest proficiency band to the highest.
 */
export const PROFICIENCY_AXIS_ORDER: ProficiencyLevel[] = [
  "Emerging",
  "Developing",
  "Connecting",
  "Benchmarking",
  "Advancing",
];

/**
 * Stable colors shared by the Analytics & Insights forecast UI.
 */
export const PROFICIENCY_DOT_COLORS: Record<ProficiencyLevel, string> = {
  Advancing: "#16a34a",
  Benchmarking: "#0d6efd",
  Connecting: "#f5a700",
  Developing: "#f97316",
  Emerging: "#ef172c",
};

export const PROFICIENCY_LEVELS: ProficiencyLevelDefinition[] = [
  {
    key: "Advancing",
    min: 90,
    max: 100,
    short: "90–100",
    cardClass: "border-emerald-200 bg-emerald-50/55",
    valueClass: "text-emerald-600",
  },
  {
    key: "Benchmarking",
    min: 80,
    max: 89,
    short: "80–89",
    cardClass: "border-blue-200 bg-blue-50/55",
    valueClass: "text-blue-600",
  },
  {
    key: "Connecting",
    min: 75,
    max: 79,
    short: "75–79",
    cardClass: "border-amber-200 bg-amber-50/55",
    valueClass: "text-amber-600",
  },
  {
    key: "Developing",
    min: 65,
    max: 74,
    short: "65–74",
    cardClass: "border-orange-200 bg-orange-50/55",
    valueClass: "text-orange-600",
  },
  {
    key: "Emerging",
    min: 0,
    max: 64,
    short: "0–64",
    cardClass: "border-red-200 bg-red-50/55",
    valueClass: "text-red-600",
  },
];

/**
 * Returns the short proficiency-range label for a level.
 */
export function proficiencyRangeLabel(level: ProficiencyLevel): string {
  return PROFICIENCY_LEVELS.find((item) => item.key === level)?.short ?? "";
}

/**
 * Returns the definition for a proficiency level.
 */
export function proficiencyDefinition(
  level: ProficiencyLevel,
): ProficiencyLevelDefinition {
  return (
    PROFICIENCY_LEVELS.find((item) => item.key === level) ??
    PROFICIENCY_LEVELS[PROFICIENCY_LEVELS.length - 1]
  );
}

/**
 * Converts a grade to its display-band position for the categorical
 * proficiency scatter plot.
 *
 * Each proficiency band occupies the same visual width/height even though
 * the numerical grade ranges have different sizes.
 */
export function proficiencyBandPosition(
  grade: number | string | null | undefined,
): number | null {
  const normalized = normalizeTermGrade(grade);
  if (normalized == null) return null;

  const level = proficiencyForGrade(normalized);
  if (level == null) return null;

  const bandIndex = PROFICIENCY_AXIS_ORDER.indexOf(level);
  const definition = proficiencyDefinition(level);
  const span = Math.max(1, definition.max - definition.min);
  const withinBand = Math.max(
    0,
    Math.min(1, (normalized - definition.min) / span),
  );

  return (bandIndex + withinBand) / PROFICIENCY_AXIS_ORDER.length;
}

/**
 * Converts a stored/calculated grade into a safe whole-number term grade.
 *
 * - null / undefined / blank / invalid values -> null
 * - valid numeric values are rounded
 * - values are clamped to 0–100
 */
export function normalizeTermGrade(
  value: number | string | null | undefined,
): number | null {
  if (value == null) return null;

  if (typeof value === "string" && value.trim() === "") {
    return null;
  }

  const numericValue =
    typeof value === "number" ? value : Number(value.trim());

  if (!Number.isFinite(numericValue)) {
    return null;
  }

  return Math.max(0, Math.min(100, Math.round(numericValue)));
}

/**
 * Returns the proficiency level for a valid numeric grade.
 */
export function proficiencyForGrade(grade: number): ProficiencyLevel;
export function proficiencyForGrade(
  grade: number | string | null | undefined,
): ProficiencyLevel | null;
export function proficiencyForGrade(
  grade: number | string | null | undefined,
): ProficiencyLevel | null {
  const normalized = normalizeTermGrade(grade);

  if (normalized == null) return null;

  if (normalized >= 90) return "Advancing";
  if (normalized >= 80) return "Benchmarking";
  if (normalized >= 75) return "Connecting";
  if (normalized >= 65) return "Developing";

  return "Emerging";
}

/**
 * Alias kept for components/utilities that use "classify" wording.
 */
export const classifyProficiency = proficiencyForGrade;

/**
 * Creates a fresh zero-filled distribution.
 */
export function emptyProficiencyDistribution(): ProficiencyDistribution {
  return {
    Advancing: 0,
    Benchmarking: 0,
    Connecting: 0,
    Developing: 0,
    Emerging: 0,
  };
}

/**
 * Counts valid grades only.
 */
export function countValidGrades(
  grades: Array<number | null | undefined>,
): number {
  return grades.reduce<number>((count, value) => {
    return normalizeTermGrade(value) == null ? count : count + 1;
  }, 0);
}

/**
 * Builds the class proficiency distribution from actual grades.
 * Missing/invalid grades are ignored.
 */
export function proficiencyDistributionForGrades(
  grades: Array<number | null | undefined>,
): ProficiencyDistribution {
  return grades.reduce<ProficiencyDistribution>((distribution, value) => {
    const level = proficiencyForGrade(value);

    if (level == null) {
      return distribution;
    }

    distribution[level] += 1;
    return distribution;
  }, emptyProficiencyDistribution());
}

/**
 * Alias for places that use "build" wording.
 */
export const buildProficiencyDistribution =
  proficiencyDistributionForGrades;

/**
 * Converts a distribution to the fixed model-vector order:
 * [Advancing, Benchmarking, Connecting, Developing, Emerging]
 */
export function proficiencyVectorForDistribution(
  distribution: ProficiencyDistribution,
): [number, number, number, number, number] {
  return [
    distribution.Advancing,
    distribution.Benchmarking,
    distribution.Connecting,
    distribution.Developing,
    distribution.Emerging,
  ];
}

/**
 * Builds the fixed model-vector directly from actual grades.
 */
export function proficiencyVectorForGrades(
  grades: Array<number | null | undefined>,
): [number, number, number, number, number] {
  return proficiencyVectorForDistribution(
    proficiencyDistributionForGrades(grades),
  );
}

/**
 * Counts how many learners have exactly a target whole-number grade.
 *
 * Example:
 * countSpecificGrade([88, 87.6, 90, null], 88) === 2
 *
 * IMPORTANT:
 * reduce<number> explicitly makes the accumulator a number.
 * This fixes the TypeScript error shown in VS Code when the source
 * array contains number | null | undefined.
 */
export function countSpecificGrade(
  grades: Array<number | null | undefined>,
  targetGrade: number,
): number {
  const target = normalizeTermGrade(targetGrade);

  if (target == null) return 0;

  return grades.reduce<number>((count, value) => {
    const grade = normalizeTermGrade(value);

    return count + (grade === target ? 1 : 0);
  }, 0);
}

/**
 * Counts grades inside an inclusive range.
 *
 * Example:
 * countGradeRange([80, 85, 89, 90, null], 80, 89) === 3
 *
 * minGrade/maxGrade can be supplied in either order.
 */
export function countGradeRange(
  grades: Array<number | null | undefined>,
  minGrade: number,
  maxGrade: number,
): number {
  const min = Math.max(
    0,
    Math.min(100, Math.round(Math.min(minGrade, maxGrade))),
  );

  const max = Math.max(
    0,
    Math.min(100, Math.round(Math.max(minGrade, maxGrade))),
  );

  return grades.reduce<number>((count, value) => {
    const grade = normalizeTermGrade(value);

    if (grade == null) return count;

    return count + (grade >= min && grade <= max ? 1 : 0);
  }, 0);
}

/**
 * Counts grades close to a target grade.
 *
 * Example:
 * target=88, tolerance=1 counts 87, 88, and 89.
 */
export function countGradesNearTarget(
  grades: Array<number | null | undefined>,
  targetGrade: number,
  tolerance = 0,
): number {
  const target = normalizeTermGrade(targetGrade);

  if (target == null) return 0;

  const safeTolerance = Math.max(0, Math.round(tolerance));

  return countGradeRange(
    grades,
    target - safeTolerance,
    target + safeTolerance,
  );
}

/**
 * Returns the percentage of graded learners in each proficiency band.
 * The percentages use only valid grades as the denominator.
 */
export function proficiencyPercentagesForGrades(
  grades: Array<number | null | undefined>,
): Record<ProficiencyLevel, number> {
  const distribution = proficiencyDistributionForGrades(grades);
  const total = PROFICIENCY_LEVEL_ORDER.reduce<number>(
    (sum, level) => sum + distribution[level],
    0,
  );

  if (total === 0) {
    return emptyProficiencyDistribution();
  }

  return PROFICIENCY_LEVEL_ORDER.reduce<Record<ProficiencyLevel, number>>(
    (result, level) => {
      result[level] =
        Math.round((distribution[level] / total) * 10000) / 100;
      return result;
    },
    emptyProficiencyDistribution(),
  );
}

/**
 * Rule-based recommendations.
 *
 * These recommendations depend on the learner's ACTUAL current
 * proficiency classification. They are not generated by the forecast model.
 */
export const PROFICIENCY_SUGGESTED_ACTIONS: Record<
  ProficiencyLevel,
  SuggestedAction
> = {
  Advancing: {
    level: "Advancing",
    title: "Enrichment",
    actions: [
      "Provide enrichment and advanced activities.",
      "Assign higher-order or independent tasks.",
      "Offer extension work that deepens the learner's mastery.",
    ],
  },

  Benchmarking: {
    level: "Benchmarking",
    title: "Deeper Application",
    actions: [
      "Provide deeper application activities.",
      "Strengthen remaining competencies.",
      "Use challenging guided practice to move toward Advancing.",
    ],
  },

  Connecting: {
    level: "Connecting",
    title: "Guided Practice",
    actions: [
      "Provide guided practice on weak competencies.",
      "Give additional exercises with feedback.",
      "Review errors before moving to more difficult tasks.",
    ],
  },

  Developing: {
    level: "Developing",
    title: "Targeted Remediation",
    actions: [
      "Provide targeted remediation.",
      "Use scaffolded activities and smaller learning steps.",
      "Focus practice on the learner's weakest competencies.",
    ],
  },

  Emerging: {
    level: "Emerging",
    title: "Intensive Support",
    actions: [
      "Provide intensive and structured support.",
      "Re-teach prerequisite or foundational competencies.",
      "Use frequent guided practice and immediate feedback.",
    ],
  },
};

/**
 * Returns the rule-based suggested action for a proficiency level.
 */
export function suggestedActionForLevel(
  level: ProficiencyLevel | null | undefined,
): SuggestedAction | null {
  if (level == null) return null;

  return PROFICIENCY_SUGGESTED_ACTIONS[level] ?? null;
}

/**
 * Returns the rule-based suggested action from an actual grade.
 */
export function suggestedActionForGrade(
  grade: number | null | undefined,
): SuggestedAction | null {
  const level = proficiencyForGrade(grade);

  return suggestedActionForLevel(level);
}

/**
 * Helpful aliases for components that use "get..." naming.
 */
export const getProficiencyLevel = proficiencyForGrade;
export const getSuggestedActionForGrade = suggestedActionForGrade;
export const getSuggestedActionForLevel = suggestedActionForLevel;
