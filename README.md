# AE210 Final Project Web Grader

Browser-hosted Jet11 grader for the AE210 Final Project. Students drag-and-drop their Jet11 `.xlsm` file; the app executes the same checks as the MATLAB autograder (`GE5_autograde_Olmstead_Fall_2025_v01`), reports the 40-point base score, and flags bonus objectives.

## What’s inside

- `docs/index.html`, `styles.css`, `app.js` – Front-end shell.
- `docs/engine/*` – Workbook loading, rule engine, and scoring logic.
- `docs/test_runner.html` – Parity harness to compare browser output against MATLAB logs.
- `docs/testdata/matlab_expected.json` – Reference logs for regression tests.

## Usage

1. Open `docs/index.html` (or publish the `docs/` folder with GitHub Pages).
2. Drop a Jet11 `.xlsm` file. Everything runs client-side; no files are uploaded.
3. Review the score summary, deductions, and bonus lines.

## Parity testing

1. Open `docs/test_runner.html`.
2. Upload a workbook whose name matches an entry in `docs/testdata/matlab_expected.json`.
3. The runner highlights any line differences between the browser log and MATLAB baseline.
*** End Patch*** }    
