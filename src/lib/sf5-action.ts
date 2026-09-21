export const SF5_ACADEMIC_EXCELLENCE_ACTION = "PROMOTED WITH ACADEMIC EXCELLENCE AWARD";

export function getSF5Action(generalAverage: number | null, actionBelowPassing = "RETAINED") {
  if (generalAverage == null) return "";
  const roundedAverage = Math.round(generalAverage);
  if (roundedAverage >= 90) return SF5_ACADEMIC_EXCELLENCE_ACTION;
  if (roundedAverage >= 75) return "PROMOTED";
  return actionBelowPassing;
}