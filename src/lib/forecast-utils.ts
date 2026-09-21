// src/lib/forecast-utils.ts
// Utilities for SIGLA's class-level proficiency forecasting.
//
// This file does NOT train a Random Forest in the browser.
// It accepts numerical output produced by an already-trained model/API and
// safely post-processes that output for the Analytics & Insights UI.

import type {
  ForecastComparisonRow,
  ForecastScatterPoint,
  ForecastTargetTerm,
  HistoricalTerm,
  MultiOutputForecastUiData,
  NumericInsight,
  ProficiencyDistribution,
  ProficiencyLevel,
  RawForecastOutput,
  RegressionMetrics,
} from "@/lib/analytics-types";

import {
  PROFICIENCY_LEVEL_ORDER as PROFICIENCY_LEVELS,
  emptyProficiencyDistribution,
  normalizeTermGrade,
  proficiencyForGrade,
} from "@/lib/proficiency";

/**
 * Returns the total number of students represented by
 * a proficiency distribution.
 *
 * Kept locally so forecast-utils.ts does not depend on
 * distributionTotal being exported from proficiency.ts.
 */
function distributionTotal(
  distribution: ProficiencyDistribution,
): number {
  return PROFICIENCY_LEVELS.reduce<number>((total, level) => {
    const value = Number(distribution[level]);

    return total + (Number.isFinite(value) ? Math.max(0, value) : 0);
  }, 0);
}

/**
 * Converts a raw Multi-Output regression prediction to valid student counts.
 *
 * Requirements implemented:
 * - negative / invalid predictions become 0
 * - output values are integers
 * - final output always sums to totalStudents
 *
 * Method:
 * 1. Clamp each prediction to a non-negative number.
 * 2. Scale all predictions so their sum equals totalStudents.
 * 3. Take floors.
 * 4. Allocate remaining students using the largest-remainder method.
 *
 * This is safer than independently rounding all five outputs because
 * independent rounding may produce a total different from the section size.
 */
export function toTotalPreservingIntegerCounts(
  raw: RawForecastOutput,
  totalStudents: number,
): ProficiencyDistribution {
  const safeTotal = Math.max(0, Math.round(totalStudents));

  if (safeTotal === 0) {
    return emptyProficiencyDistribution();
  }

  const values = PROFICIENCY_LEVELS.map((level) => {
    const numeric = Number(raw[level]);

    return {
      level,
      raw: Number.isFinite(numeric) ? Math.max(0, numeric) : 0,
    };
  });

  const rawSum = values.reduce<number>(
    (sum, item) => sum + item.raw,
    0,
  );

  if (rawSum <= 0) {
    throw new Error(
      "The forecasting model returned no positive counts, so a valid class distribution cannot be created.",
    );
  }

  const scaled = values.map((item) => {
    const exact = (item.raw / rawSum) * safeTotal;
    const floor = Math.floor(exact);

    return {
      level: item.level,
      exact,
      floor,
      fraction: exact - floor,
    };
  });

  let remaining =
    safeTotal -
    scaled.reduce<number>(
      (sum, item) => sum + item.floor,
      0,
    );

  const allocationOrder = [...scaled].sort((first, second) => {
    if (second.fraction !== first.fraction) {
      return second.fraction - first.fraction;
    }

    return (
      PROFICIENCY_LEVELS.indexOf(first.level) -
      PROFICIENCY_LEVELS.indexOf(second.level)
    );
  });

  const result = emptyProficiencyDistribution();

  scaled.forEach((item) => {
    result[item.level] = item.floor;
  });

  let index = 0;

  while (remaining > 0 && allocationOrder.length > 0) {
    const target =
      allocationOrder[index % allocationOrder.length];

    result[target.level] += 1;

    remaining -= 1;
    index += 1;
  }

  return result;
}

/**
 * Checks whether a forecast distribution:
 * - contains only valid non-negative integer counts
 * - sums exactly to the total number of students
 */
export function isValidForecastDistribution(
  distribution: ProficiencyDistribution,
  totalStudents: number,
): boolean {
  const safeTotal = Math.max(0, Math.round(totalStudents));

  return (
    PROFICIENCY_LEVELS.every((level) => {
      const value = distribution[level];

      return (
        Number.isInteger(value) &&
        Number.isFinite(value) &&
        value >= 0
      );
    }) && distributionTotal(distribution) === safeTotal
  );
}

/**
 * Compares the actual proficiency distribution with the forecast.
 */
export function compareDistributions(
  actual: ProficiencyDistribution,
  forecast: ProficiencyDistribution,
): ForecastComparisonRow[] {
  return PROFICIENCY_LEVELS.map((level) => {
    const actualCount = Math.max(
      0,
      Number(actual[level]) || 0,
    );

    const forecastCount = Math.max(
      0,
      Number(forecast[level]) || 0,
    );

    const signedError = forecastCount - actualCount;

    return {
      level,
      actual: actualCount,
      forecast: forecastCount,
      signedError,
      absoluteError: Math.abs(signedError),
    };
  });
}

/**
 * Generates short insights from NUMERICAL RESULTS ONLY.
 */
export function generateForecastInsights(
  actualTerm: ProficiencyDistribution,
  forecastTerm: ProficiencyDistribution,
  targetTermLabel = "Term 2",
): NumericInsight[] {
  const insights: NumericInsight[] = [];

  const largestForecast = PROFICIENCY_LEVELS
    .map((level) => ({
      level,
      count: Math.max(
        0,
        Number(forecastTerm[level]) || 0,
      ),
    }))
    .sort(
      (first, second) =>
        second.count - first.count,
    )[0];

  if (largestForecast) {
    insights.push({
      id: "largest-forecast-group",
      level: largestForecast.level,
      direction: "info",
      message: `${largestForecast.level} is expected to be the largest proficiency group in ${targetTermLabel} with approximately ${largestForecast.count} student${
        largestForecast.count === 1 ? "" : "s"
      }.`,
    });
  }

  PROFICIENCY_LEVELS.forEach((level) => {
    const current = Math.max(
      0,
      Number(actualTerm[level]) || 0,
    );

    const forecast = Math.max(
      0,
      Number(forecastTerm[level]) || 0,
    );

    const difference = forecast - current;

    if (difference > 0) {
      insights.push({
        id: `${level}-increase`,
        level,
        direction: "increase",
        message: `${level} is forecasted to increase from ${current} in the current term to ${forecast} in ${targetTermLabel}.`,
      });

      return;
    }

    if (difference < 0) {
      insights.push({
        id: `${level}-decrease`,
        level,
        direction: "decrease",
        message: `${level} is forecasted to decrease from ${current} in the current term to ${forecast} in ${targetTermLabel}.`,
      });

      return;
    }

    insights.push({
      id: `${level}-same`,
      level,
      direction: "same",
      message: `${level} is forecasted to remain at ${forecast} student${
        forecast === 1 ? "" : "s"
      } in ${targetTermLabel}.`,
    });
  });

  return insights;
}

/**
 * Mean Absolute Error.
 */
export function meanAbsoluteError(
  actual: number[],
  predicted: number[],
): number | null {
  if (
    actual.length === 0 ||
    actual.length !== predicted.length
  ) {
    return null;
  }

  const errors = actual.map((value, index) =>
    Math.abs(value - predicted[index]),
  );

  return (
    errors.reduce<number>(
      (sum, value) => sum + value,
      0,
    ) / errors.length
  );
}

/**
 * Root Mean Squared Error.
 */
export function rootMeanSquaredError(
  actual: number[],
  predicted: number[],
): number | null {
  if (
    actual.length === 0 ||
    actual.length !== predicted.length
  ) {
    return null;
  }

  const mse =
    actual.reduce<number>((sum, value, index) => {
      const error = value - predicted[index];

      return sum + error * error;
    }, 0) / actual.length;

  return Math.sqrt(mse);
}

/**
 * R² coefficient of determination.
 */
export function rSquared(
  actual: number[],
  predicted: number[],
): number | null {
  if (
    actual.length === 0 ||
    actual.length !== predicted.length
  ) {
    return null;
  }

  const mean =
    actual.reduce<number>(
      (sum, value) => sum + value,
      0,
    ) / actual.length;

  const totalSumOfSquares =
    actual.reduce<number>((sum, value) => {
      const delta = value - mean;

      return sum + delta * delta;
    }, 0);

  if (totalSumOfSquares === 0) {
    return null;
  }

  const residualSumOfSquares =
    actual.reduce<number>(
      (sum, value, index) => {
        const error = value - predicted[index];

        return sum + error * error;
      },
      0,
    );

  return (
    1 -
    residualSumOfSquares / totalSumOfSquares
  );
}

/**
 * Calculates regression evaluation metrics.
 */
export function regressionMetrics(
  actual: number[],
  predicted: number[],
): RegressionMetrics {
  return {
    mae: meanAbsoluteError(actual, predicted),
    rmse: rootMeanSquaredError(actual, predicted),
    r2: rSquared(actual, predicted),
  };
}

/**
 * Evaluates multiple historical multi-output test samples.
 */
export function evaluateMultiOutputPredictions(
  actualRows: number[][],
  predictedRows: number[][],
): {
  overall: RegressionMetrics;
  perCategory: Record<
    ProficiencyLevel,
    RegressionMetrics
  >;
} {
  const emptyMetrics: RegressionMetrics = {
    mae: null,
    rmse: null,
    r2: null,
  };

  const emptyPerCategory = (): Record<
    ProficiencyLevel,
    RegressionMetrics
  > =>
    Object.fromEntries(
      PROFICIENCY_LEVELS.map((level) => [
        level,
        {
          mae: null,
          rmse: null,
          r2: null,
        },
      ]),
    ) as Record<
      ProficiencyLevel,
      RegressionMetrics
    >;

  if (
    actualRows.length === 0 ||
    actualRows.length !== predictedRows.length
  ) {
    return {
      overall: emptyMetrics,
      perCategory: emptyPerCategory(),
    };
  }

  const validPairs = actualRows
    .map((actual, index) => ({
      actual,
      predicted: predictedRows[index],
    }))
    .filter(
      ({ actual, predicted }) =>
        Array.isArray(actual) &&
        Array.isArray(predicted) &&
        actual.length === PROFICIENCY_LEVELS.length &&
        predicted.length ===
          PROFICIENCY_LEVELS.length &&
        actual.every(Number.isFinite) &&
        predicted.every(Number.isFinite),
    );

  if (validPairs.length === 0) {
    return {
      overall: emptyMetrics,
      perCategory: emptyPerCategory(),
    };
  }

  const flattenedActual =
    validPairs.flatMap((row) => row.actual);

  const flattenedPredicted =
    validPairs.flatMap((row) => row.predicted);

  const perCategory = Object.fromEntries(
    PROFICIENCY_LEVELS.map(
      (level, categoryIndex) => {
        const actual = validPairs.map(
          (row) => row.actual[categoryIndex],
        );

        const predicted = validPairs.map(
          (row) => row.predicted[categoryIndex],
        );

        return [
          level,
          regressionMetrics(
            actual,
            predicted,
          ),
        ];
      },
    ),
  ) as Record<
    ProficiencyLevel,
    RegressionMetrics
  >;

  return {
    overall: regressionMetrics(
      flattenedActual,
      flattenedPredicted,
    ),
    perCategory,
  };
}

/**
 * Converts a proficiency distribution to the
 * fixed five-value vector expected by the
 * Multi-Output model.
 */
export function distributionToVector(
  distribution: ProficiencyDistribution,
): [
  number,
  number,
  number,
  number,
  number,
] {
  return [
    distribution.Advancing,
    distribution.Benchmarking,
    distribution.Connecting,
    distribution.Developing,
    distribution.Emerging,
  ];
}

/**
 * Converts a five-value vector back
 * to a named proficiency distribution.
 */
export function vectorToDistribution(
  vector: readonly number[],
): ProficiencyDistribution {
  if (vector.length !== 5) {
    throw new Error(
      "A proficiency forecast vector must contain exactly five values.",
    );
  }

  return {
    Advancing: Number.isFinite(Number(vector[0]))
      ? Math.max(0, Number(vector[0]))
      : 0,

    Benchmarking: Number.isFinite(Number(vector[1]))
      ? Math.max(0, Number(vector[1]))
      : 0,

    Connecting: Number.isFinite(Number(vector[2]))
      ? Math.max(0, Number(vector[2]))
      : 0,

    Developing: Number.isFinite(Number(vector[3]))
      ? Math.max(0, Number(vector[3]))
      : 0,

    Emerging: Number.isFinite(Number(vector[4]))
      ? Math.max(0, Number(vector[4]))
      : 0,
  };
}

/**
 * Returns the forecast target that follows a completed historical term.
 */
export function forecastTargetForSourceTerm(
  sourceTerm: HistoricalTerm,
): ForecastTargetTerm {
  if (sourceTerm === "1") return "2";
  if (sourceTerm === "2") return "3";
  return "final";
}

/**
 * Counts learners at or above a grade threshold.
 *
 * This is used by the forecast UI to compare the model's optional >=88
 * output with the source term's actual >=88 count.
 */
export function countGradesAtOrAbove(
  grades: Array<number | null | undefined>,
  threshold = 88,
): number {
  const safeThreshold = Math.max(0, Math.min(100, Math.round(threshold)));

  return grades.reduce<number>((count, value) => {
    const grade = normalizeTermGrade(value);
    if (grade == null) return count;
    return count + (grade >= safeThreshold ? 1 : 0);
  }, 0);
}

/**
 * Normalizes the model's optional >=88 output to a valid whole-student count.
 */
export function normalizeForecastCount(
  value: number | null | undefined,
  totalStudents: number,
): number | null {
  if (value == null) return null;

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return null;

  const safeTotal = Math.max(0, Math.round(totalStudents));
  return Math.max(0, Math.min(safeTotal, Math.round(numeric)));
}

/**
 * Builds learner-level actual-vs-predicted points for the scatter plot.
 *
 * IMPORTANT:
 * These predicted grades must come from the trained model/API.
 * The class-level five-output distribution alone is not enough to recreate
 * learner-level predicted grades.
 */
export function buildForecastScatterPoints(
  actualGrades: Array<number | null | undefined>,
  predictedGrades: Array<number | null | undefined>,
  studentIds: string[] = [],
): ForecastScatterPoint[] {
  const length = Math.min(actualGrades.length, predictedGrades.length);
  const points: ForecastScatterPoint[] = [];

  for (let index = 0; index < length; index += 1) {
    const actualGrade = normalizeTermGrade(actualGrades[index]);
    const predictedGrade = normalizeTermGrade(predictedGrades[index]);

    if (actualGrade == null || predictedGrade == null) {
      continue;
    }

    const level = proficiencyForGrade(predictedGrade);
    if (level == null) continue;

    points.push({
      studentId: studentIds[index],
      actualGrade,
      predictedGrade,
      level,
    });
  }

  return points;
}

/**
 * Creates the complete UI payload used by the Multi-Output Random Forest
 * card in Analytics & Insights.
 *
 * The trained model/API supplies:
 * - raw class-count predictions for the five proficiency outputs
 * - optionally a sixth >=88 count
 * - optionally learner-level predicted grades for the scatter plot
 */
export function buildMultiOutputForecastUiData({
  sourceTerm,
  raw,
  totalStudents,
  actualGrades = [],
  predictedGrades = [],
  studentIds = [],
}: {
  sourceTerm: HistoricalTerm;
  raw: RawForecastOutput;
  totalStudents: number;
  actualGrades?: Array<number | null | undefined>;
  predictedGrades?: Array<number | null | undefined>;
  studentIds?: string[];
}): MultiOutputForecastUiData {
  return {
    sourceTerm,
    targetTerm: forecastTargetForSourceTerm(sourceTerm),
    counts: toTotalPreservingIntegerCounts(raw, totalStudents),
    predictedAtLeast88: normalizeForecastCount(
      raw.AtLeast88,
      totalStudents,
    ),
    scatterPoints: buildForecastScatterPoints(
      actualGrades,
      predictedGrades,
      studentIds,
    ),
    modelLabel: "Multi-Output Random Forest Regression",
  };
}

