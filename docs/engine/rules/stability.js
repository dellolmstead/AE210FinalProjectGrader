import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

export function runStabilityChecks(workbook) {
  const feedback = [];
  let failures = 0;

  const main = workbook.sheets.main;

  const sm = asNumber(getCell(main, "M10"));
  const clb = asNumber(getCell(main, "O10"));
  const cnb = asNumber(getCell(main, "P10"));
  const ratio = asNumber(getCell(main, "Q10"));

  if (!(sm >= -0.1 && sm <= 0.11)) {
    feedback.push(format(STRINGS.stability.sm, sm ?? NaN));
    failures += 1;
    if (Number.isFinite(sm) && sm < 0) {
      feedback.push(STRINGS.stability.smWarn);
    }
  } else if (Number.isFinite(sm) && sm < 0) {
    feedback.push(STRINGS.stability.smWarn);
  }

  if (!(clb < -0.001)) {
    feedback.push(format(STRINGS.stability.clb, clb ?? NaN));
    failures += 1;
  }

  if (!(cnb > 0.002)) {
    feedback.push(format(STRINGS.stability.cnb, cnb ?? NaN));
    failures += 1;
  }

  if (!(ratio >= -1 && ratio <= -0.3)) {
    feedback.push(format(STRINGS.stability.ratio, ratio ?? NaN));
    failures += 1;
  }

  if (failures > 0) {
    const deduction = Math.min(3, failures);
    feedback.push(format(STRINGS.stability.deduction, deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
