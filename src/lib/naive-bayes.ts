// src/lib/naive-bayes.ts
// Learner-level Gaussian Naive Bayes implementation for SIGLA Analytics.
//
// Training flow:
//   X = Term 1 WW, PT, STE, Final Grade
//   y = Term 2 proficiency category
//
// Forecast flow:
//   Apply the trained model to Term 2 WW, PT, STE, Final Grade
//   -> predict each learner's Term 3 proficiency category.

import type {
  AnalyticsFeatureKey,
  AnalyticsFeatureVector,
  NaiveBayesClassStats,
  NaiveBayesCrossValidation,
  NaiveBayesModel,
  NaiveBayesTrainingRow,
  ProficiencyLevel,
  ProficiencyProbability,
} from "@/lib/analytics-types";

const FEATURE_KEYS: AnalyticsFeatureKey[] = [
  "ww",
  "pt",
  "ste",
  "finalGrade",
];

const PROFICIENCY_ORDER: ProficiencyLevel[] = [
  "Advancing",
  "Benchmarking",
  "Connecting",
  "Developing",
  "Emerging",
];

const MIN_VARIANCE = 1;

function emptyProbabilities(): ProficiencyProbability {
  return {
    Advancing: 0,
    Benchmarking: 0,
    Connecting: 0,
    Developing: 0,
    Emerging: 0,
  };
}

function mean(values: number[]) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function sampleVariance(values: number[], fallback = MIN_VARIANCE) {
  if (values.length <= 1) return Math.max(MIN_VARIANCE, fallback);

  const average = mean(values);
  const variance =
    values.reduce((sum, value) => {
      const delta = value - average;
      return sum + delta * delta;
    }, 0) /
    (values.length - 1);

  return Number.isFinite(variance)
    ? Math.max(MIN_VARIANCE, variance)
    : Math.max(MIN_VARIANCE, fallback);
}

function globalVariance(
  rows: NaiveBayesTrainingRow[],
  feature: AnalyticsFeatureKey,
) {
  return sampleVariance(
    rows.map((row) => row.features[feature]),
    MIN_VARIANCE,
  );
}

export function trainGaussianNaiveBayes(
  rows: NaiveBayesTrainingRow[],
): NaiveBayesModel | null {
  if (rows.length === 0) return null;

  const observedLabels = PROFICIENCY_ORDER.filter((label) =>
    rows.some((row) => row.label === label),
  );

  if (observedLabels.length === 0) return null;

  const globalVariances = Object.fromEntries(
    FEATURE_KEYS.map((feature) => [feature, globalVariance(rows, feature)]),
  ) as Record<AnalyticsFeatureKey, number>;

  const classStats: Partial<
    Record<ProficiencyLevel, NaiveBayesClassStats>
  > = {};

  observedLabels.forEach((label) => {
    const classRows = rows.filter((row) => row.label === label);

    const features = Object.fromEntries(
      FEATURE_KEYS.map((feature) => {
        const values = classRows.map((row) => row.features[feature]);
        return [
          feature,
          {
            mean: mean(values),
            variance: sampleVariance(values, globalVariances[feature]),
          },
        ];
      }),
    ) as NaiveBayesClassStats["features"];

    classStats[label] = {
      label,
      count: classRows.length,
      prior: classRows.length / rows.length,
      features,
    };
  });

  return {
    modelLabel: "Gaussian Naive Bayes",
    trainingSize: rows.length,
    labels: observedLabels,
    classStats,
  };
}

function gaussianLogDensity(value: number, meanValue: number, variance: number) {
  const safeVariance = Math.max(MIN_VARIANCE, variance);
  const delta = value - meanValue;

  return (
    -0.5 * Math.log(2 * Math.PI * safeVariance) -
    (delta * delta) / (2 * safeVariance)
  );
}

export function predictGaussianNaiveBayes(
  model: NaiveBayesModel,
  features: AnalyticsFeatureVector,
): {
  label: ProficiencyLevel;
  confidence: number;
  probabilities: ProficiencyProbability;
} {
  const logScores = model.labels.map((label) => {
    const stats = model.classStats[label];

    if (!stats) {
      return { label, score: Number.NEGATIVE_INFINITY };
    }

    let score = Math.log(Math.max(stats.prior, Number.EPSILON));

    FEATURE_KEYS.forEach((feature) => {
      const featureStats = stats.features[feature];
      score += gaussianLogDensity(
        features[feature],
        featureStats.mean,
        featureStats.variance,
      );
    });

    return { label, score };
  });

  const maxLogScore = Math.max(...logScores.map((item) => item.score));
  const exponentials = logScores.map((item) => ({
    label: item.label,
    value: Number.isFinite(item.score)
      ? Math.exp(item.score - maxLogScore)
      : 0,
  }));
  const denominator = exponentials.reduce(
    (sum, item) => sum + item.value,
    0,
  );

  const probabilities = emptyProbabilities();

  exponentials.forEach((item) => {
    probabilities[item.label] =
      denominator > 0 ? item.value / denominator : 0;
  });

  const best = model.labels
    .map((label) => ({ label, probability: probabilities[label] }))
    .sort((left, right) => right.probability - left.probability)[0];

  return {
    label: best?.label ?? model.labels[0] ?? "Emerging",
    confidence: best?.probability ?? 0,
    probabilities,
  };
}

/**
 * Deterministic 5-fold cross-validation.
 *
 * A full validation score is returned only when at least five complete
 * training rows are available. This prevents a misleading "5-fold" score
 * from being shown for a dataset that cannot actually be split five ways.
 */
export function crossValidateGaussianNaiveBayes(
  rows: NaiveBayesTrainingRow[],
): NaiveBayesCrossValidation {
  const foldsRequested = 5 as const;

  if (rows.length < foldsRequested) {
    return {
      foldsRequested,
      foldsUsed: 0,
      sampleCount: rows.length,
      correct: 0,
      accuracy: null,
    };
  }

  const orderedRows = [...rows].sort((left, right) =>
    `${left.label}|${left.studentId}`.localeCompare(
      `${right.label}|${right.studentId}`,
    ),
  );

  const folds = Array.from(
    { length: foldsRequested },
    () => [] as NaiveBayesTrainingRow[],
  );

  orderedRows.forEach((row, index) => {
    folds[index % foldsRequested].push(row);
  });

  let correct = 0;
  let tested = 0;
  let foldsUsed = 0;

  folds.forEach((testRows, foldIndex) => {
    if (testRows.length === 0) return;

    const trainingRows = folds.flatMap((fold, index) =>
      index === foldIndex ? [] : fold,
    );

    const model = trainGaussianNaiveBayes(trainingRows);
    if (!model) return;

    foldsUsed += 1;

    testRows.forEach((row) => {
      const prediction = predictGaussianNaiveBayes(model, row.features);
      tested += 1;
      if (prediction.label === row.label) correct += 1;
    });
  });

  return {
    foldsRequested,
    foldsUsed,
    sampleCount: tested,
    correct,
    accuracy: tested > 0 ? correct / tested : null,
  };
}
