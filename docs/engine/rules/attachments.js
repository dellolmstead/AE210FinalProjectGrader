import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const DEG_TO_RAD = Math.PI / 180;

export function runAttachmentChecks(workbook) {
  const feedback = [];
  let controlFailures = 0;

  const main = workbook.sheets.main;
  const geom = workbook.sheets.geom;

  const fuselageLength = asNumber(getCell(main, "B32"));
  const pcsX = asNumber(getCell(main, "C23"));
  const pcsRootChord = asNumber(getCell(geom, "C8"));
  if (
    !Number.isFinite(fuselageLength) ||
    !Number.isFinite(pcsX) ||
    !Number.isFinite(pcsRootChord)
  ) {
    feedback.push("Unable to verify PCS placement due to missing geometry data");
    controlFailures += 1;
  } else if (pcsX > fuselageLength - 0.25 * pcsRootChord) {
    feedback.push(STRINGS.attachment.pcsX);
    controlFailures += 1;
  }

  const vtX = asNumber(getCell(main, "H23"));
  const vtRootChord = asNumber(getCell(geom, "C10"));
  if (
    !Number.isFinite(fuselageLength) ||
    !Number.isFinite(vtX) ||
    !Number.isFinite(vtRootChord)
  ) {
    feedback.push("Unable to verify vertical tail placement due to missing geometry data");
    controlFailures += 1;
  } else if (vtX > fuselageLength - 0.25 * vtRootChord) {
    feedback.push(STRINGS.attachment.vtX);
    controlFailures += 1;
  }

  const pcsZ = asNumber(getCell(main, "C25"));
  const fuseZCenter = asNumber(getCell(main, "D52"));
  const fuseZHeight = asNumber(getCell(main, "F52"));
  if (
    Number.isFinite(pcsZ) &&
    Number.isFinite(fuseZCenter) &&
    Number.isFinite(fuseZHeight)
  ) {
    if (pcsZ < fuseZCenter - fuseZHeight / 2 || pcsZ > fuseZCenter + fuseZHeight / 2) {
      feedback.push(STRINGS.attachment.pcsZ);
      controlFailures += 1;
    }
  } else {
    feedback.push("Unable to verify PCS vertical placement due to missing geometry data");
    controlFailures += 1;
  }

  const vtY = asNumber(getCell(main, "H24"));
  const fuseWidth = asNumber(getCell(main, "E52"));
  if (Number.isFinite(vtY) && Number.isFinite(fuseWidth)) {
    if (vtY > fuseWidth / 2) {
      feedback.push(STRINGS.attachment.vtY);
      controlFailures += 1;
    }
  } else {
    feedback.push("Unable to verify vertical tail lateral placement due to missing geometry data");
    controlFailures += 1;
  }

  const strakeArea = asNumber(getCell(main, "D18"));
  if (Number.isFinite(strakeArea) && strakeArea > 1) {
    const sweep = asNumber(getCell(geom, "K15"));
    const y = asNumber(getCell(geom, "M152"));
    const strake = asNumber(getCell(geom, "L155"));
    const apex = asNumber(getCell(geom, "L38"));
    if (
      Number.isFinite(sweep) &&
      Number.isFinite(y) &&
      Number.isFinite(strake) &&
      Number.isFinite(apex)
    ) {
      const wing = y / Math.tan((90 - sweep) * DEG_TO_RAD) + apex;
      if (!(wing < strake + 0.5)) {
        feedback.push(STRINGS.attachment.strake);
        controlFailures += 1;
      }
    } else {
      feedback.push("Unable to verify strake attachment due to missing geometry data");
      controlFailures += 1;
    }
  }

  if (Number.isFinite(fuselageLength)) {
    const componentPositions = [];
    for (let col = 2; col <= 8; col += 1) {
      const val = asNumber(getCellByIndex(main, 23, col));
      if (Number.isFinite(val)) {
        componentPositions.push(val);
      }
    }
    if (componentPositions.some((value) => value >= fuselageLength)) {
      feedback.push(format(STRINGS.attachment.fuselage, fuselageLength));
      controlFailures += 1;
    }
  }

  if (controlFailures > 0) {
    const deduction = Math.min(2, controlFailures);
    feedback.push(format(STRINGS.attachment.deduction, deduction));
    return { delta: -deduction, feedback };
  }

  return { delta: 0, feedback };
}
