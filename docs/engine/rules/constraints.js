import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const MACH_TOL = 1e-2;
const VALUE_TOL = 1e-3;
const PS_TOL = 1;
const CDX_TOL = 1e-3;
const BETA_DEFAULT = 0.87620980519917;

const CONSTRAINTS = [
  {
    label: "MaxMach",
    row: 3,
    altMin: 35000,
    machMin: 2.0,
    machObj: 2.2,
    nEq: 1,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "CruiseMach",
    row: 4,
    altMin: 35000,
    machMin: 1.5,
    machObj: 1.8,
    nEq: 1,
    abEq: 0,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Cmbt Turn1",
    row: 6,
    machEq: 1.2,
    altEq: 30000,
    nMin: 3.0,
    nObj: 4.0,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Cmbt Turn2",
    row: 7,
    machEq: 0.9,
    altEq: 10000,
    nMin: 4.0,
    nObj: 4.5,
    abEq: 100,
    psEq: 0,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Ps1",
    row: 8,
    machEq: 1.15,
    altEq: 30000,
    nEq: 1,
    abEq: 100,
    psMin: 400,
    psObj: 500,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Ps2",
    row: 9,
    machEq: 0.9,
    altEq: 10000,
    nEq: 1,
    abEq: 0,
    psMin: 400,
    psObj: 500,
    cdxEq: 0,
    betaEq: BETA_DEFAULT,
  },
  {
    label: "Takeoff",
    row: 12,
    altEq: 0,
    machEq: 1.2,
    nEq: 0.03,
    abEq: 100,
    psEq: 0,
    cdxAllowed: [0, 0.035],
    betaEq: 1,
  },
  {
    label: "Landing",
    row: 13,
    altEq: 0,
    machEq: 1.3,
    nEq: 0.5,
    abEq: 0,
    psEq: 0,
    cdxAllowed: [0, 0.045],
    betaEq: 1,
  },
];

const CURVE_ROWS = [
  { row: 23, label: "MaxMach" },
  { row: 24, label: "Supercruise" },
  { row: 26, label: "CombatTurn1" },
  { row: 27, label: "CombatTurn2" },
  { row: 28, label: "Ps1" },
  { row: 29, label: "Ps2" },
  { row: 32, label: "Takeoff" },
];

function interpolate(xList, yList, x) {
  const pairs = xList
    .map((value, idx) => ({ x: asNumber(value), y: asNumber(yList[idx]) }))
    .filter((pair) => Number.isFinite(pair.x) && Number.isFinite(pair.y))
    .sort((a, b) => a.x - b.x);

  if (pairs.length === 0) {
    return null;
  }
  if (x <= pairs[0].x) {
    if (pairs.length === 1) {
      return pairs[0].y;
    }
    const [p0, p1] = pairs;
    const slope = (p1.y - p0.y) / (p1.x - p0.x);
    return p0.y + slope * (x - p0.x);
  }
  if (x >= pairs[pairs.length - 1].x) {
    const [p0, p1] = pairs.slice(-2);
    const slope = (p1.y - p0.y) / (p1.x - p0.x);
    return p1.y + slope * (x - p1.x);
  }
  for (let i = 0; i < pairs.length - 1; i += 1) {
    const p0 = pairs[i];
    const p1 = pairs[i + 1];
    if (x >= p0.x && x <= p1.x) {
      const slope = (p1.y - p0.y) / (p1.x - p0.x);
      return p0.y + slope * (x - p0.x);
    }
  }
  return null;
}

export function runConstraintChecks(workbook) {
  const feedback = [];
  let tableErrors = 0;
  let payloadPenalty = 0;
  let curvePenalty = 0;

  const main = workbook.sheets.main;
  const consts = workbook.sheets.consts;

  const radius = asNumber(getCell(main, "Y37"));
  if (Number.isFinite(radius)) {
    if (radius < 375) {
      feedback.push(format(STRINGS.constraint.radiusLow, radius));
      tableErrors += 1;
    } else if (radius >= 410) {
      feedback.push(format(STRINGS.constraint.radiusObj, radius));
    }
  }

  const aim120 = asNumber(getCell(main, "AB3"));
  const aim9 = asNumber(getCell(main, "AB4"));
  if (!Number.isFinite(aim120) || aim120 < 8) {
    const value = Number.isFinite(aim120) ? aim120 : 0;
    payloadPenalty -= 4;
    feedback.push(format(STRINGS.constraint.payloadPenalty, value));
  } else if (Number.isFinite(aim9) && aim9 >= 2) {
    feedback.push(format(STRINGS.constraint.payloadObj, aim120, aim9));
  }

  const takeoffDist = asNumber(getCell(main, "X12"));
  if (Number.isFinite(takeoffDist)) {
    if (takeoffDist > 3000) {
      feedback.push(format(STRINGS.constraint.takeoffHigh, takeoffDist));
      tableErrors += 1;
    } else if (takeoffDist <= 2500) {
      feedback.push(format(STRINGS.constraint.takeoffObj, takeoffDist));
    }
  }

  const landingDist = asNumber(getCell(main, "X13"));
  if (Number.isFinite(landingDist)) {
    if (landingDist > 5000) {
      feedback.push(format(STRINGS.constraint.landingHigh, landingDist));
      tableErrors += 1;
    } else if (landingDist <= 3500) {
      feedback.push(format(STRINGS.constraint.landingObj, landingDist));
    }
  }

  CONSTRAINTS.forEach((constraint) => {
    const row = constraint.row;
    const beta = asNumber(getCellByIndex(main, row, 19));
    const altitude = asNumber(getCellByIndex(main, row, 20));
    const mach = asNumber(getCellByIndex(main, row, 21));
    const n = asNumber(getCellByIndex(main, row, 22));
    const ab = asNumber(getCellByIndex(main, row, 23));
    const ps = asNumber(getCellByIndex(main, row, 24));
    const cdx = asNumber(getCellByIndex(main, row, 25));

    const exceedsTol = (value, expected, tol) =>
      !Number.isFinite(value) || Math.abs(value - expected) > tol;

    if (constraint.machEq != null) {
      if (exceedsTol(mach, constraint.machEq, MACH_TOL)) {
        feedback.push(format(STRINGS.constraint.machEq, constraint.label, mach ?? NaN, constraint.machEq));
        tableErrors += 1;
      }
    } else if (constraint.machMin != null) {
      if (!Number.isFinite(mach) || mach < constraint.machMin - MACH_TOL) {
        feedback.push(format(STRINGS.constraint.machMin, constraint.label, mach ?? NaN, constraint.machMin));
        tableErrors += 1;
      } else if (constraint.machObj != null && Number.isFinite(mach) && mach >= constraint.machObj - MACH_TOL) {
        feedback.push(format(STRINGS.constraint.machObj, constraint.label, constraint.machObj, mach));
      }
    }

    if (constraint.altEq != null) {
      if (exceedsTol(altitude, constraint.altEq, 1)) {
        feedback.push(format(STRINGS.constraint.altEq, constraint.label, altitude ?? NaN, constraint.altEq));
        tableErrors += 1;
      }
    } else if (constraint.altMin != null) {
      if (!Number.isFinite(altitude) || altitude < constraint.altMin - 1) {
        feedback.push(format(STRINGS.constraint.altMin, constraint.label, altitude ?? NaN, constraint.altMin));
        tableErrors += 1;
      }
    }

    if (constraint.nEq != null) {
      if (exceedsTol(n, constraint.nEq, VALUE_TOL)) {
        feedback.push(format(STRINGS.constraint.nEq, constraint.label, n ?? NaN, constraint.nEq));
        tableErrors += 1;
      }
    } else if (constraint.nMin != null) {
      if (!Number.isFinite(n) || n < constraint.nMin - VALUE_TOL) {
        feedback.push(format(STRINGS.constraint.nMin, constraint.label, n ?? NaN, constraint.nMin));
        tableErrors += 1;
      } else if (constraint.nObj != null && Number.isFinite(n) && n >= constraint.nObj - VALUE_TOL) {
        feedback.push(format(STRINGS.constraint.nObj, constraint.label, constraint.nObj, n));
      }
    }

    if (constraint.abEq != null) {
      if (exceedsTol(ab, constraint.abEq, VALUE_TOL)) {
        feedback.push(format(STRINGS.constraint.abEq, constraint.label, ab ?? NaN, constraint.abEq));
        tableErrors += 1;
      }
    }

    if (constraint.psEq != null) {
      if (exceedsTol(ps, constraint.psEq, PS_TOL)) {
        feedback.push(format(STRINGS.constraint.psEq, constraint.label, ps ?? NaN, constraint.psEq));
        tableErrors += 1;
      }
    } else if (constraint.psMin != null) {
      if (!Number.isFinite(ps) || ps < constraint.psMin - PS_TOL) {
        feedback.push(format(STRINGS.constraint.psMin, constraint.label, ps ?? NaN, constraint.psMin));
        tableErrors += 1;
      } else if (constraint.psObj != null && Number.isFinite(ps) && ps >= constraint.psObj - PS_TOL) {
        feedback.push(format(STRINGS.constraint.psObj, constraint.label, constraint.psObj, ps));
      }
    }

    if (constraint.betaEq != null) {
      if (exceedsTol(beta, constraint.betaEq, VALUE_TOL)) {
        feedback.push(format(STRINGS.constraint.betaEq, constraint.label, beta ?? NaN));
        tableErrors += 1;
      }
    }

    if (constraint.cdxAllowed) {
      const allowedMatch = constraint.cdxAllowed.some(
        (allowed) => Number.isFinite(cdx) && Math.abs(cdx - allowed) <= CDX_TOL
      );
      if (!allowedMatch) {
        const allowedText = constraint.cdxAllowed.map((value) => value.toFixed(3)).join(", ");
        feedback.push(format(STRINGS.constraint.cdxAllowed, constraint.label, cdx ?? NaN, allowedText));
        tableErrors += 1;
      }
    } else if (constraint.cdxEq != null) {
      if (exceedsTol(cdx, constraint.cdxEq, CDX_TOL)) {
        feedback.push(format(STRINGS.constraint.cdxEq, constraint.label, cdx ?? NaN, constraint.cdxEq));
        tableErrors += 1;
      }
    }
  });

  const deduction = Math.min(2, tableErrors);
  let delta = 0;
  if (deduction > 0) {
    feedback.push(format(STRINGS.constraintSummary, deduction));
    delta -= deduction;
  }

  try {
    const wsAxis = [];
    for (let col = 11; col <= 31; col += 1) {
      wsAxis.push(getCellByIndex(consts, 22, col));
    }
    const wsDesign = asNumber(getCell(main, "P13"));
    const twDesign = asNumber(getCell(main, "Q13"));

    const failures = [];
    if (Number.isFinite(wsDesign) && Number.isFinite(twDesign)) {
      CURVE_ROWS.forEach(({ row, label }) => {
        const twCurve = [];
        for (let col = 11; col <= 31; col += 1) {
          twCurve.push(getCellByIndex(consts, row, col));
        }
        const requiredTW = interpolate(wsAxis, twCurve, wsDesign);
        if (requiredTW != null && twDesign < requiredTW) {
          failures.push(label);
        }
      });

      const wsLimitLanding = asNumber(getCell(consts, "L33"));
      if (Number.isFinite(wsLimitLanding) && wsDesign > wsLimitLanding) {
        failures.push("Landing");
        feedback.push(format(STRINGS.constraint.landingCurve, wsDesign, wsLimitLanding));
      }

      if (failures.length === 1) {
        const deduction = 4;
        curvePenalty -= deduction;
        feedback.push(
          format(STRINGS.constraint.curveFailure, deduction, "", failures[0]) + STRINGS.constraint.curveSuffixFew
        );
      } else if (failures.length >= 2) {
        const deduction = 8;
        curvePenalty -= deduction;
        const joined = failures.join(", ");
        const suffix =
          failures.length > 6 ? STRINGS.constraint.curveSuffixMany : STRINGS.constraint.curveSuffixFew;
        feedback.push(format(STRINGS.constraint.curveFailure, deduction, "s", joined) + suffix);
      }
    }
  } catch (err) {
    feedback.push(format(STRINGS.constraint.curveError, err.message));
  }

  return {
    delta,
    payloadDelta: payloadPenalty,
    curveDelta: curvePenalty,
    feedback,
  };
}
