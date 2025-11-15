import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const THRUST_TOL = 1e-3;

const THRUST_CELLS = [
  ["C48", "C49"],
  ["D48", "D49"],
  ["E48", "E49"],
  ["F48", "F49"],
  ["G48", "G49"],
  ["H48", "H49"],
  ["I48", "I49"],
  ["J48", "J49"],
  ["K48", "K49"],
  ["L48", "L49"],
  ["M48", "M49"],
  ["N48", "N49"],
];

export function runThrustAndTakeoff(workbook) {
  const feedback = [];
  let failures = 0;

  const miss = workbook.sheets.miss;

  THRUST_CELLS.forEach(([dragRef, thrustRef]) => {
    const drag = asNumber(getCell(miss, dragRef));
    const thrust = asNumber(getCell(miss, thrustRef));
    if (Number.isFinite(drag) && Number.isFinite(thrust) && thrust <= drag + THRUST_TOL) {
      failures += 1;
    }
  });

  if (failures > 0) {
    const deduction = Math.min(3, failures);
    feedback.push(format(STRINGS.thrustShortfall, deduction, failures));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
