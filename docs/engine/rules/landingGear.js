import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const PERCENT_TOL = 1e-3;
const ANGLE_TOL = 1e-2;
const SPEED_TOL = 1e-3;

export function runLandingGearChecks(workbook) {
  const feedback = [];
  let failures = 0;

  const gear = workbook.sheets.gear;

  const noseRule = asNumber(getCell(gear, "J20"));
  if (!Number.isFinite(noseRule) || noseRule < 80 - PERCENT_TOL || noseRule > 95 + PERCENT_TOL) {
    feedback.push(format(STRINGS.gear.nose, noseRule));
    failures += 1;
  }

  const tipbackUpper = asNumber(getCell(gear, "L20"));
  const tipbackLower = asNumber(getCell(gear, "L21"));
  if (
    !Number.isFinite(tipbackUpper) ||
    !Number.isFinite(tipbackLower) ||
    tipbackUpper >= tipbackLower - ANGLE_TOL
  ) {
    feedback.push(format(STRINGS.gear.tipback, tipbackUpper, tipbackLower));
    failures += 1;
  }

  const rolloverUpper = asNumber(getCell(gear, "M20"));
  const rolloverLower = asNumber(getCell(gear, "M21"));
  if (
    !Number.isFinite(rolloverUpper) ||
    !Number.isFinite(rolloverLower) ||
    rolloverUpper >= rolloverLower - ANGLE_TOL
  ) {
    feedback.push(format(STRINGS.gear.rollover, rolloverUpper, rolloverLower));
    failures += 1;
  }

  const rotationSpeed = asNumber(getCell(gear, "N20"));
  if (!Number.isFinite(rotationSpeed) || rotationSpeed >= 200 - SPEED_TOL) {
    feedback.push(format(STRINGS.gear.rotation, rotationSpeed));
    failures += 1;
  }

  if (failures > 0) {
    const deduction = Math.min(4, failures);
    feedback.push(format(STRINGS.gear.deduction, deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
