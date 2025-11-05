import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const LEG_COLUMNS = [11, 12, 13, 14, 16, 18, 19, 22, 23];
const ROWS = {
  altitude: 33,
  mach: 35,
  afterburner: 36,
  distance: 38,
  time: 39,
};

export function runMissionChecks(workbook) {
  const feedback = [];
  let errors = 0;

  const main = workbook.sheets.main;
  const constraintsMach = asNumber(getCell(main, "U4"));

  const readRowValues = (rowIndex) =>
    LEG_COLUMNS.map((col) => asNumber(getCellByIndex(main, rowIndex, col)));

  const altitude = readRowValues(ROWS.altitude);
  const mach = readRowValues(ROWS.mach);
  const afterburner = readRowValues(ROWS.afterburner);
  const distance = readRowValues(ROWS.distance);
  const time = readRowValues(ROWS.time);

  const pushIf = (condition, message) => {
    if (condition) {
      feedback.push(message);
      errors += 1;
    }
  };

  pushIf(altitude[0] !== 0 || afterburner[0] !== 100, STRINGS.missionLegs[0]);
  pushIf(!(altitude[1] >= altitude[0] && altitude[1] <= altitude[2]), STRINGS.missionLegs[1]);
  pushIf(!(mach[1] >= mach[0] && mach[1] <= mach[2]), STRINGS.missionLegs[2]);
  pushIf(afterburner[1] !== 0, STRINGS.missionLegs[3]);

  pushIf(altitude[2] < 35000 || mach[2] !== 0.9 || afterburner[2] !== 0, STRINGS.missionLegs[4]);
  pushIf(altitude[3] < 35000 || mach[3] !== 0.9 || afterburner[3] !== 0, STRINGS.missionLegs[5]);

  pushIf(
    altitude[4] < 35000 ||
      !Number.isFinite(constraintsMach) ||
      Math.abs(mach[4] - constraintsMach) > 0.01 ||
      afterburner[4] !== 0 ||
      distance[4] < 150,
    STRINGS.missionLegs[6]
  );

  pushIf(altitude[5] < 30000 || mach[5] < 1.2 || afterburner[5] !== 100 || time[5] < 2, STRINGS.missionLegs[7]);

  pushIf(
    altitude[6] < 35000 ||
      !Number.isFinite(constraintsMach) ||
      Math.abs(mach[6] - constraintsMach) > 0.01 ||
      afterburner[6] !== 0 ||
      distance[6] < 150,
    STRINGS.missionLegs[8]
  );

  pushIf(altitude[7] < 35000 || mach[7] !== 0.9 || afterburner[7] !== 0, STRINGS.missionLegs[9]);
  pushIf(altitude[8] !== 10000 || mach[8] !== 0.4 || afterburner[8] !== 0 || time[8] !== 20, STRINGS.missionLegs[10]);

  const radius = asNumber(getCell(main, "Y37"));
  if (Number.isFinite(radius) && radius < 375) {
    feedback.push(format(STRINGS.constraint.radiusLow, radius));
    errors += 1;
  }

  if (errors > 0) {
    const deduction = Math.min(2, errors);
    feedback.push(STRINGS.missionSummary.replace("%d", deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
