import { loadWorkbook } from "./engine/loadWorkbook.js";
import { gradeWorkbook } from "./engine/grade.js";
import { RULES } from "./engine/rules.js";

const fileInput = document.getElementById("file-input");
const dropZone = document.getElementById("drop-zone");
const resultsSection = document.getElementById("results");
const statusMessage = document.getElementById("status-message");
const comparisonOutcome = document.getElementById("comparison-outcome");
const diffOutput = document.getElementById("diff-output");

const WEIRD_GE = String.fromCharCode(0xe2) + String.fromCharCode(0x2030) + String.fromCharCode(0xa5);
const WEIRD_LE = String.fromCharCode(0xe2) + String.fromCharCode(0x2030) + String.fromCharCode(0xa4);
const CHECK_MARK = "\u2705";
const CROSS_MARK = "\u274C";

let baselines = new Map();

async function loadBaselines() {
  const response = await fetch("./testdata/matlab_expected.json");
  if (!response.ok) {
    throw new Error(`Unable to load baseline data (${response.status})`);
  }
  const data = await response.json();
  baselines = new Map(data.map((item) => [item.file, item]));
}

function resetUI() {
  resultsSection.classList.add("hidden");
  statusMessage.classList.add("hidden");
  statusMessage.textContent = "";
  comparisonOutcome.textContent = "";
  diffOutput.innerHTML = "";
}

function showStatus(text, level = "info") {
  statusMessage.textContent = text;
  statusMessage.classList.remove("hidden", "info", "warning", "error");
  statusMessage.classList.add(level);
}

function normalizeLine(line) {
  if (!line && line !== "") {
    return "";
  }
  return line
    .replaceAll(WEIRD_GE, "≥")
    .replaceAll(WEIRD_LE, "≤")
    .replace(/\r/g, "")
    .replace(/\s+$/g, "");
}

function normalizeLog(log) {
  const lines = Array.isArray(log) ? log : log.split(/\r?\n/);
  const cleaned = lines.map((line) => normalizeLine(line));
  while (cleaned.length > 0 && cleaned[cleaned.length - 1] === "") {
    cleaned.pop();
  }
  return cleaned;
}

function compareLogs(expected, actual) {
  const exp = normalizeLog(expected);
  const act = normalizeLog(actual);
  const max = Math.max(exp.length, act.length);
  const rows = [];
  let mismatches = 0;

  for (let i = 0; i < max; i += 1) {
    const expectedLine = exp[i] ?? "";
    const actualLine = act[i] ?? "";
    const match = expectedLine === actualLine;
    if (!match) {
      mismatches += 1;
    }
    rows.push({
      index: i + 1,
      expected: expectedLine,
      actual: actualLine,
      match,
    });
  }

  return { rows, mismatches };
}

function renderDiff(rows) {
  if (!rows.length) {
    diffOutput.textContent = "No lines to compare.";
    return;
  }

  const table = document.createElement("table");
  table.className = "diff-table";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["Line", "Expected", "Actual"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (!row.match) {
      tr.classList.add("diff-row-mismatch");
    }
    const lineCell = document.createElement("td");
    lineCell.textContent = row.index;
    const expectedCell = document.createElement("td");
    expectedCell.textContent = row.expected;
    const actualCell = document.createElement("td");
    actualCell.textContent = row.actual;
    tr.append(lineCell, expectedCell, actualCell);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  diffOutput.innerHTML = "";
  diffOutput.appendChild(table);
}

async function runComparison(file) {
  resetUI();

  if (!file) {
    return;
  }

  const lowerName = file.name.toLowerCase();
  if (!lowerName.endsWith(".xlsm") && !lowerName.endsWith(".xlsx")) {
    showStatus("Only .xlsm or .xlsx files are supported.", "error");
    return;
  }

  const baseline = baselines.get(file.name);
  if (!baseline) {
    showStatus(`No MATLAB baseline found for "${file.name}".`, "warning");
    return;
  }

  showStatus("Running grader…", "info");

  try {
    const workbook = await loadWorkbook(file, RULES);
    const result = gradeWorkbook(workbook, RULES);
    const actualLog = result.feedbackLog.split(/\r?\n/);

    const { rows, mismatches } = compareLogs(baseline.logLines, actualLog);
    const outcomeText =
      mismatches === 0
        ? `${CHECK_MARK} Match: ${file.name}`
        : `${CROSS_MARK} ${mismatches} mismatched line${mismatches === 1 ? "" : "s"} for ${file.name}`;

    comparisonOutcome.textContent = outcomeText;
    renderDiff(rows);
    resultsSection.classList.remove("hidden");
    showStatus("Comparison complete.", "info");
  } catch (error) {
    console.error(error);
    showStatus(`Error: ${error.message || "Unable to grade this file."}`, "error");
  }
}

function handleFiles(files) {
  if (!files || files.length === 0) {
    return;
  }
  const [file] = files;
  runComparison(file);
}

fileInput.addEventListener("change", (event) => {
  handleFiles(event.target.files);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  handleFiles(event.dataTransfer.files);
});

loadBaselines()
  .then(() => {
    showStatus("Baseline data loaded. Upload a file to compare.", "info");
  })
  .catch((error) => {
    console.error(error);
    showStatus(`Failed to load baselines: ${error.message}`, "error");
  });
