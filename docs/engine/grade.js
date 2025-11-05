import { STRINGS } from "./messages.js";
import { runMissionChecks } from "./rules/mission.js";
import { runAeroChecks } from "./rules/aero.js";
import { runThrustAndTakeoff } from "./rules/thrust.js";
import { runConstraintChecks } from "./rules/constraints.js";
import { runAttachmentChecks } from "./rules/attachments.js";
import { runStabilityChecks } from "./rules/stability.js";
import { runFuelVolumeChecks } from "./rules/fuel.js";
import { runRecurringCostChecks } from "./rules/cost.js";
import { runLandingGearChecks } from "./rules/landingGear.js";
import { getCell, asNumber } from "./parseUtils.js";
import { format } from "./format.js";

const TOLERANCES = {
  tol: 1e-3,
  mach: 1e-2,
  dist: 1e-3,
};

function computeBonuses(workbook) {
  const messages = [];
  let points = 0;

  const main = workbook.sheets.main;

  const radius = asNumber(getCell(main, "Y37"));
  if (Number.isFinite(radius) && radius >= 410 - TOLERANCES.dist) {
    points += 1;
    messages.push(format(STRINGS.bonus.radius, radius));
  }

  const aim120 = asNumber(getCell(main, "AB3"));
  const aim9 = asNumber(getCell(main, "AB4"));
  if (
    Number.isFinite(aim120) &&
    Number.isFinite(aim9) &&
    aim120 >= 8 - TOLERANCES.tol &&
    aim9 >= 2 - TOLERANCES.tol
  ) {
    points += 1;
    messages.push(format(STRINGS.bonus.payload, aim120, aim9));
  }

  const takeoffDist = asNumber(getCell(main, "X12"));
  if (Number.isFinite(takeoffDist) && takeoffDist <= 2500 + TOLERANCES.dist) {
    points += 1;
    messages.push(format(STRINGS.bonus.takeoff, takeoffDist));
  }

  const landingDist = asNumber(getCell(main, "X13"));
  if (Number.isFinite(landingDist) && landingDist <= 3500 + TOLERANCES.dist) {
    points += 1;
    messages.push(format(STRINGS.bonus.landing, landingDist));
  }

  const maxMach = asNumber(getCell(main, "U3"));
  if (Number.isFinite(maxMach) && maxMach >= 2.2 - TOLERANCES.mach) {
    points += 1;
    messages.push(format(STRINGS.bonus.maxMach, maxMach));
  }

  const supercruiseMach = asNumber(getCell(main, "U4"));
  if (Number.isFinite(supercruiseMach) && supercruiseMach >= 1.8 - TOLERANCES.mach) {
    points += 1;
    messages.push(format(STRINGS.bonus.supercruise, supercruiseMach));
  }

  const psHigh = asNumber(getCell(main, "X8"));
  if (Number.isFinite(psHigh) && psHigh >= 500 - TOLERANCES.dist) {
    points += 1;
    messages.push(format(STRINGS.bonus.psHigh, psHigh));
  }

  const psLow = asNumber(getCell(main, "X9"));
  if (Number.isFinite(psLow) && psLow >= 500 - TOLERANCES.dist) {
    points += 1;
    messages.push(format(STRINGS.bonus.psLow, psLow));
  }

  const gHigh = asNumber(getCell(main, "V6"));
  if (Number.isFinite(gHigh) && gHigh >= 4 - TOLERANCES.tol) {
    points += 1;
    messages.push(format(STRINGS.bonus.gHigh, gHigh));
  }

  const gLow = asNumber(getCell(main, "V7"));
  if (Number.isFinite(gLow) && gLow >= 4.5 - TOLERANCES.tol) {
    points += 1;
    messages.push(format(STRINGS.bonus.gLow, gLow));
  }

  const cost = asNumber(getCell(main, "Q31"));
  const numAircraft = asNumber(getCell(main, "N31"));
  if (Number.isFinite(cost) && Number.isFinite(numAircraft)) {
    if (numAircraft === 187 && cost <= 80 + TOLERANCES.tol) {
      points += 1;
      messages.push(format(STRINGS.bonus.cost, cost));
    } else if (numAircraft === 800 && cost <= 50 + TOLERANCES.tol) {
      points += 1;
      messages.push(format(STRINGS.bonus.cost, cost));
    }
  }

  return { points, messages };
}

export function gradeWorkbook(workbook, rules) {
  const feedback = [];
  let baseScore = 40;

  if (workbook.fileName) {
    feedback.push(workbook.fileName);
  }

  const aeroResult = runAeroChecks(workbook, rules);
  baseScore += aeroResult.delta;
  feedback.push(...aeroResult.feedback);

  const missionResult = runMissionChecks(workbook, rules);
  baseScore += missionResult.delta;
  feedback.push(...missionResult.feedback);

  const thrustResult = runThrustAndTakeoff(workbook, rules);
  baseScore += thrustResult.delta;
  feedback.push(...thrustResult.feedback);

  const constraintResult = runConstraintChecks(workbook, rules);
  baseScore += constraintResult.delta;
  baseScore += constraintResult.payloadDelta;
  baseScore += constraintResult.curveDelta;
  feedback.push(...constraintResult.feedback);

  const attachmentResult = runAttachmentChecks(workbook, rules);
  baseScore += attachmentResult.delta;
  feedback.push(...attachmentResult.feedback);

  const stabilityResult = runStabilityChecks(workbook, rules);
  baseScore += stabilityResult.delta;
  feedback.push(...stabilityResult.feedback);

  const fuelResult = runFuelVolumeChecks(workbook, rules);
  baseScore += fuelResult.delta;
  feedback.push(...fuelResult.feedback);

  const costResult = runRecurringCostChecks(workbook, rules);
  baseScore += costResult.delta;
  feedback.push(...costResult.feedback);

  const gearResult = runLandingGearChecks(workbook, rules);
  baseScore += gearResult.delta;
  feedback.push(...gearResult.feedback);

  const clampedBase = Math.max(0, baseScore);

  const bonusResult = computeBonuses(workbook);
  feedback.push(...bonusResult.messages);

  const finalScore = clampedBase + bonusResult.points;

  const scoreLine = format(STRINGS.summary.base, clampedBase);
  const bonusLine =
    bonusResult.points > 0
      ? format(STRINGS.summary.bonusEarned, bonusResult.points, finalScore)
      : format(STRINGS.summary.bonusNone, finalScore);

  feedback.push(scoreLine);
  feedback.push(bonusLine);

  return {
    score: finalScore,
    maxScore: 40,
    scoreLine,
    bonusLine,
    feedbackLog: feedback.join("\n"),
  };
}
