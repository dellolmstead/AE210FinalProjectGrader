import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

export function runRecurringCostChecks(workbook) {
  const feedback = [];
  let delta = 0;

  const main = workbook.sheets.main;
  const cost = asNumber(getCell(main, "Q31"));
  const numAircraft = asNumber(getCell(main, "N31"));

  if (numAircraft === 187) {
    if (!Number.isFinite(cost) || cost > 115) {
      feedback.push(format(STRINGS.cost.over187, cost ?? 0));
      delta -= 5;
    } else if (cost <= 100) {
      feedback.push(format(STRINGS.cost.obj187, cost));
    }
  } else if (numAircraft === 800) {
    if (!Number.isFinite(cost) || cost > 75) {
      feedback.push(format(STRINGS.cost.over800, cost ?? 0));
      delta -= 5;
    } else if (cost <= 63) {
      feedback.push(format(STRINGS.cost.obj800, cost));
    }
  } else {
    feedback.push(format(STRINGS.cost.invalid, numAircraft ?? NaN));
    delta -= 5;
  }

  return { delta, feedback };
}
