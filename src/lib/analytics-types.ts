// src/lib/analytics-types.ts
// Shared types for SIGLA Analytics & Insights.
//
// The forecasting flow in this version is learner-level Gaussian Naive Bayes:
// Term 1 features -> Term 2 proficiency label, then Term 2 features ->
// predicted Term 3 proficiency category.

export type AnalyticsTerm = "1" | "2" | "3" | "final";

export type HistoricalTerm = "1" | "2" | "3";

export type ForecastTargetTerm = "3";

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

/**
 * Four learner-level variables used by the Naive Bayes model.
 *
 * `ste` is the Summative Test / Examination input. The existing E-Class
 * Record stores that component under the internal `QA` key, so the Analytics
 * panel maps QA -> STE when it builds this feature object.
 */
export type AnalyticsFeatureVector = {
  ww: number;
  pt: number;
  ste: number;
  finalGrade: number;
};

export type AnalyticsFeatureKey = keyof AnalyticsFeatureVector;

/**
 * One supervised training row:
 * X = Term 1 WW, PT, STE, Final Grade
 * y = Term 2 proficiency category
 */
export type NaiveBayesTrainingRow = {
  studentId: string;
  studentName: string;
  features: AnalyticsFeatureVector;
  label: ProficiencyLevel;
};

/**
 * One learner whose Term 2 features are ready for a Term 3 prediction.
 */
export type NaiveBayesPredictionInput = {
  studentId: string;
  studentName: string;
  features: AnalyticsFeatureVector;
  actualTerm2Category: ProficiencyLevel;
};

export type ProficiencyProbability = Record<ProficiencyLevel, number>;

export type NaiveBayesPrediction = {
  studentId: string;
  studentName: string;
  actualTerm2Category: ProficiencyLevel;
  predictedTerm3Category: ProficiencyLevel;
  confidence: number;
  probabilities: ProficiencyProbability;
  term2Features: AnalyticsFeatureVector;
};

export type NaiveBayesFeatureStats = {
  mean: number;
  variance: number;
};

export type NaiveBayesClassStats = {
  label: ProficiencyLevel;
  count: number;
  prior: number;
  features: Record<AnalyticsFeatureKey, NaiveBayesFeatureStats>;
};

export type NaiveBayesModel = {
  modelLabel: "Gaussian Naive Bayes";
  trainingSize: number;
  labels: ProficiencyLevel[];
  classStats: Partial<Record<ProficiencyLevel, NaiveBayesClassStats>>;
};

export type NaiveBayesCrossValidation = {
  foldsRequested: 5;
  foldsUsed: number;
  sampleCount: number;
  correct: number;
  accuracy: number | null;
};

export type NaiveBayesForecastUiData = {
  modelLabel: "Gaussian Naive Bayes";
  sourceTerm: "2";
  targetTerm: "3";
  trainingSize: number;
  predictionInputCount: number;
  learnedCategories: ProficiencyLevel[];
  validation: NaiveBayesCrossValidation;
  predictions: NaiveBayesPrediction[];
  distribution: ProficiencyDistribution;
  insights: NumericInsight[];
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
