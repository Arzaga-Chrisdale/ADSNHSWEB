// src/lib/forecast-utils.ts
// Utilities for SIGLA's learner-level Naive Bayes Term 3 forecasting.

import type {
  NaiveBayesForecastUiData,
  NaiveBayesPrediction,
  NaiveBayesPredictionInput,
  NaiveBayesTrainingRow,
  NumericInsight,
  ProficiencyDistribution,
  ProficiencyLevel,
} from "./analytics-types";

import {
  crossValidateGaussianNaiveBayes,
  predictGaussianNaiveBayes,
  trainGaussianNaiveBayes,
} from "./naive-bayes";

import {
  PROFICIENCY_LEVEL_ORDER,
  emptyProficiencyDistribution,
} from "./proficiency";

export const MIN_NAIVE_BAYES_TRAINING_ROWS = 5;

export function distributionTotal(
  distribution: ProficiencyDistribution,
): number {
  return PROFICIENCY_LEVEL_ORDER.reduce<number>((total, level) => {
    const value = Number(distribution[level]);

    return total + (Number.isFinite(value) ? Math.max(0, value) : 0);
  }, 0);
}

export function aggregateNaiveBayesPredictions(
  predictions: NaiveBayesPrediction[],
): ProficiencyDistribution {
  const distribution = emptyProficiencyDistribution();

  predictions.forEach((prediction) => {
    distribution[prediction.predictedTerm3Category] += 1;
  });

  return distribution;
}

export function generateNaiveBayesInsights(
  actualTerm2: ProficiencyDistribution,
  forecastTerm3: ProficiencyDistribution,
): NumericInsight[] {
  const insights: NumericInsight[] = [];

  const largestForecast = PROFICIENCY_LEVEL_ORDER
    .map((level) => ({
      level,
      count: forecastTerm3[level],
    }))
    .sort((left, right) => right.count - left.count)[0];

  if (largestForecast) {
    insights.push({
      id: "largest-term3-group",
      level: largestForecast.level,
      direction: "info",
      message: `${largestForecast.level} is the largest predicted Term 3 proficiency group with ${largestForecast.count} learner${
        largestForecast.count === 1 ? "" : "s"
      }.`,
    });
  }

  PROFICIENCY_LEVEL_ORDER.forEach((level) => {
    const current = Math.max(0, Number(actualTerm2[level]) || 0);
    const forecast = Math.max(0, Number(forecastTerm3[level]) || 0);
    const difference = forecast - current;

    if (difference === 0) return;

    insights.push({
      id: `${level}-${difference > 0 ? "increase" : "decrease"}`,
      level,
      direction: difference > 0 ? "increase" : "decrease",
      message: `${level} is predicted to ${
        difference > 0 ? "increase" : "decrease"
      } from ${current} learner${
        current === 1 ? "" : "s"
      } in Term 2 to ${forecast} in Term 3.`,
    });
  });

  const supportCount =
    forecastTerm3.Developing + forecastTerm3.Emerging;

  if (supportCount > 0) {
    insights.push({
      id: "support-priority",
      direction: "info",
      message: `${supportCount} learner${
        supportCount === 1 ? " is" : "s are"
      } predicted in Developing or Emerging for Term 3. Prioritize targeted remediation and close monitoring for these learners.`,
    });
  }

  return insights.slice(0, 6);
}

/**
 * Builds the complete UI-ready Term 3 forecast.
 *
 * Returns null until at least five complete T1 -> T2 training rows and at
 * least one complete Term 2 prediction input are available. This keeps the
 * UI from inventing model output when the E-Class Record is incomplete.
 */
export function buildNaiveBayesForecastUiData({
  trainingRows,
  predictionInputs,
  actualTerm2Distribution,
}: {
  trainingRows: NaiveBayesTrainingRow[];
  predictionInputs: NaiveBayesPredictionInput[];
  actualTerm2Distribution: ProficiencyDistribution;
}): NaiveBayesForecastUiData | null {
  if (
    trainingRows.length < MIN_NAIVE_BAYES_TRAINING_ROWS ||
    predictionInputs.length === 0
  ) {
    return null;
  }

  const model = trainGaussianNaiveBayes(trainingRows);

  if (!model) {
    return null;
  }

  const validation =
    crossValidateGaussianNaiveBayes(trainingRows);

  const predictions = predictionInputs.map<NaiveBayesPrediction>(
    (input) => {
      const prediction = predictGaussianNaiveBayes(
        model,
        input.features,
      );

      return {
        studentId: input.studentId,
        studentName: input.studentName,
        actualTerm2Category: input.actualTerm2Category,
        predictedTerm3Category: prediction.label,
        confidence: prediction.confidence,
        probabilities: prediction.probabilities,
        term2Features: input.features,
      };
    },
  );

  const distribution =
    aggregateNaiveBayesPredictions(predictions);

  return {
    modelLabel: "Gaussian Naive Bayes",
    sourceTerm: "2",
    targetTerm: "3",
    trainingSize: trainingRows.length,
    predictionInputCount: predictionInputs.length,
    learnedCategories: model.labels,
    validation,
    predictions,
    distribution,
    insights: generateNaiveBayesInsights(
      actualTerm2Distribution,
      distribution,
    ),
  };
}

export function highestProbabilityLabel(
  probabilities: Record<ProficiencyLevel, number>,
): ProficiencyLevel {
  return PROFICIENCY_LEVEL_ORDER.reduce(
    (best, level) =>
      probabilities[level] > probabilities[best]
        ? level
        : best,
  );
}
