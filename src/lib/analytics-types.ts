// src/lib/analytics-types.ts
// Shared types for SIGLA Analytics & Insights.
//
// This file contains types only. It does not perform database access,
// Machine Learning training, or UI rendering.

export type AnalyticsTerm = "1" | "2" | "3" | "final";

export type HistoricalTerm = "1" | "2" | "3";

export type ForecastTargetTerm = "2" | "3" | "final";

export type ProficiencyLevel =
  | "Advancing"
  | "Benchmarking"
  | "Connecting"
  | "Developing"
  | "Emerging";

export type ProficiencyDistribution = Record<ProficiencyLevel, number>;

export type ProficiencyDefinition = {
  level: ProficiencyLevel;
  min: number;
  max: number;
  label: string;
};

export type StudentTermGrade = {
  studentId: string;
  grade: number;
};

export type ProficiencyCountRow = {
  level: ProficiencyLevel;
  count: number;
  percentage: number;
};

export type ProficiencySeries = {
  term: AnalyticsTerm;
  label: string;
  distribution: ProficiencyDistribution;
};

export type GradeFrequencyRow = {
  label: string;
  min: number;
  max: number;
  count: number;
  percentage: number;
};

export type ExactGradeFrequencyRow = {
  grade: number;
  count: number;
  percentage: number;
};

export type ForecastContext = {
  gradeLevel: string;
  section: string;
  subject: string;
  sourceTerm: HistoricalTerm;
  targetTerm: ForecastTargetTerm;
  totalStudents: number;
};

export type ForecastModelInput = {
  context: ForecastContext;
  inputDistribution: ProficiencyDistribution;
};

/**
 * Numerical outputs returned by the trained Multi-Output Random Forest.
 *
 * The five proficiency outputs are required.
 * AtLeast88 is optional so the current five-output API remains compatible.
 * If the trained model is updated to predict >=88 as a sixth target,
 * return it here.
 */
export type RawForecastOutput = {
  Advancing: number;
  Benchmarking: number;
  Connecting: number;
  Developing: number;
  Emerging: number;
  AtLeast88?: number;
};

export type ForecastScatterPoint = {
  studentId?: string;
  actualGrade: number;
  predictedGrade: number;
  level: ProficiencyLevel;
};

export type ForecastOutput = {
  context: ForecastContext;
  raw: RawForecastOutput;
  counts: ProficiencyDistribution;
  predictedAtLeast88?: number;
  scatterPoints?: ForecastScatterPoint[];
};

/**
 * UI-ready data for the Random Forest card.
 *
 * This is intentionally separate from RawForecastOutput because the scatter
 * plot requires learner-level actual/predicted grade pairs. Those pairs must
 * come from the trained model/API; they cannot be reconstructed from the five
 * class-count outputs alone.
 */
export type MultiOutputForecastUiData = {
  sourceTerm: HistoricalTerm;
  targetTerm: ForecastTargetTerm;
  counts: ProficiencyDistribution;
  predictedAtLeast88: number | null;
  scatterPoints: ForecastScatterPoint[];
  modelLabel?: string;
};

export type ForecastComparisonRow = {
  level: ProficiencyLevel;
  actual: number;
  forecast: number;
  signedError: number;
  absoluteError: number;
};

export type RegressionMetrics = {
  mae: number | null;
  rmse: number | null;
  r2: number | null;
};

export type PerCategoryRegressionMetrics = Record<
  ProficiencyLevel,
  RegressionMetrics
>;

export type ForecastEvaluation = {
  overall: RegressionMetrics;
  perCategory: PerCategoryRegressionMetrics;
  countComparison?: ForecastComparisonRow[];
};

export type HistoricalForecastDataPoint = {
  id?: string;
  gradeLevel: string;
  section: string;
  subject: string;

  term1: ProficiencyDistribution;
  term2: ProficiencyDistribution;

  totalStudents: number;
};

export type NumericInsight = {
  id: string;
  message: string;
  level?: ProficiencyLevel;
  direction?: "increase" | "decrease" | "same" | "info";
};

export type SuggestedAction = {
  level: ProficiencyLevel;
  title: string;
  actions: string[];
};
