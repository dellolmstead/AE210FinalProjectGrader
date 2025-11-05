
% Final Project Autograder: Grades AE210 Jet11 Excel submissions, logs feedback, and optionally exports Blackboard-compatible scores.


%--------------------------------------------------------------------------
% AE210 Final Project Autograder Script â€“ Fall 2025
%
% Description:
% This script automates grading for the AE210 Final Project by processing Jet11 Excel files (*.xlsm). It evaluates 
% multiple design criteria, generates detailed feedback, and outputs both a
% summary log and an optional Blackboard-compatible grade import file.
%
% Key Features:
% - Supports both single-file and batch-folder grading via GUI
% - Parallel-safe execution using MATLAB's parpool
% - Robust Excel reading with fallback for missing data
% - Detailed feedback log per cadet with scoring breakdown
% - Optional export to Blackboard offline grade format (SMART_TEXT)
% - Histogram visualization of score distribution
%
% Inputs:
% - User-selected Excel file or folder of files
%
% Outputs:
% - Text log file: textout_<timestamp>.txt
% - Histogram of scores
% - Optional Blackboard CSV: FinalProject_Blackboard_Offline_<timestamp>.csv
%
% Embedded Functions:
% - gradeCadet: Grades a single cadet's file and returns score and feedback
% - loadAllJet11Sheets: Loads all required sheets from a Jet11 Excel file
% - safeReadMatrix: Robustly reads numeric data from Excel, with fallback to readcell
% - cell2sub: Converts Excel cell references (e.g., 'G4') to row/col indices
% - sub2excel: Converts row/col indices back to Excel cell references
% - logf: Appends formatted text to a log string
% - selectRunMode: GUI for selecting single file or folder mode
% - promptAndGenerateBlackboardCSV: Dialog + export to Blackboard SMART_TEXT format
%
% Author: Lt Col Dell Olmstead, based on work by Capt Carol Bryant and Capt Anna Mason
% Heavy ChatGPT CODEX help in Nov 2025
% Last Updated: 4 Nov 2025
%--------------------------------------------------------------------------
clear; close all; clc;


%% Choose directory and get Excel files
% fprintf('Executing %s\n',mfilename);
% I recommend updating the below line to point to your Final Project files. It works
% as is, but will default to the right place if this is updated.

% folderAnalyzed = uigetdir('C:\Users\dell.olmstead\OneDrive - afacademy.af.edu\Documents 1\01 Classes\AE210 FA24\Design Project\Final Project files');
% fprintf('%s\n\n', folderAnalyzed);
% files = dir(fullfile(folderAnalyzed, '*.xlsm'));



%% Select run mode: single file or folder, start parallel pool if folder
[mode, selectedPath] = selectRunMode();
tic
if strcmp(mode, 'cancelled')
    disp('Operation cancelled by user.');
    return;
elseif strcmp(mode, 'single')
    folderAnalyzed = fileparts(selectedPath);
    files = dir(selectedPath);  % single file
elseif strcmp(mode, 'folder')
    % Ensure a process-based parallel pool is active
    poolobj = gcp('nocreate'); % Get the current pool, if any
    if isempty(poolobj)
        % Create a new local pool, ensuring process-based if possible
        try
            p = parpool('local'); % Try the simplest form first
        catch ME
            if contains(ME.message, 'ExecutionMode') % Check for specific error message
                p = parpool('local', 'ExecutionMode', 'Processes'); % Use ExecutionMode if supported
            else
                rethrow(ME); % If it's a different error, re-throw it
            end
        end

        if ~isempty(p)
            if isa(p, 'parallel.ThreadPool')
                warning('Created a thread-based pool despite requesting "local". Attempting to delete and recreate as process-based.');
                delete(p);
                parpool('local', 'ExecutionMode', 'Processes'); % Explicitly use ExecutionMode
            elseif isa(p, 'parallel.Pool')
                fprintf('Successfully created a process-based local parallel pool.\n');
            end
        end
    elseif isa(poolobj, 'parallel.ThreadPool')
        % If an existing pool is thread-based, delete it and create a process-based one
        warning('Existing parallel pool is thread-based. Deleting and creating a process-based local pool.');
        delete(poolobj);
        parpool('local', 'ExecutionMode', 'Processes'); % Explicitly use ExecutionMode
    elseif isa(poolobj, 'parallel.Pool')
        fprintf('A process-based local parallel pool is already running.\n');
    end
    folderAnalyzed = selectedPath;
    files = [dir(fullfile(folderAnalyzed, '*.xlsm')); dir(fullfile(folderAnalyzed, '*.xlsx')); dir(fullfile(folderAnalyzed, '*.xls'))];
else
    error('Unknown mode selected.');
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%% Iterate through cadets %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
textout = strings(numel(files), 1);
points = 10*ones(numel(files),1);  % Initialize points for each file
feedback = cell(1,numel(files));

fprintf('Reading %d files\n', numel(files));

if strcmp(mode, 'folder')        
    % Combined parallel read + grade
    
parfor cadetIdx = 1:numel(files)
    filename = fullfile(folderAnalyzed, files(cadetIdx).name);
    try
        [pt, fb] = gradeCadet(filename);
        points(cadetIdx) = pt;
        feedback{cadetIdx} = fb;
    catch
        points(cadetIdx) = NaN;
        feedback{cadetIdx} = sprintf('Error reading or grading file: %s', files(cadetIdx).name);
    end
end

else %       %%% Use the below code to run a single cadet

    filename = fullfile(folderAnalyzed, files(1).name);
    [points, feedback{1}] = gradeCadet(filename);

end


%% Set up log file
timestamp = char(datetime('now', 'Format', 'yyyy-MM-dd_HH-mm-ss'));
logFilePath = fullfile(folderAnalyzed, ['textout_', timestamp, '.txt']);
finalout = fopen(logFilePath,'w');

% Log file header
fprintf(finalout, 'Final Project Autograder Log\n');
fprintf(finalout, 'Script Name: %s.m\n', mfilename);
fprintf(finalout, 'Run Date: %s\n', string(datetime('now', 'Format', 'yyyy-MM-dd HH:mm:ss')));
fprintf(finalout, 'Analyzed Folder: %s\n', folderAnalyzed);
fprintf(finalout, 'Files to Analyze (%d):\n', numel(files));
for i = 1:numel(files)
    fprintf(finalout, '  - %s\n', files(i).name);
end
fprintf(finalout, '\n');

%% Concatenate all outputs into one text file and write it.
allLogText = strjoin(string(feedback(:)), '\n\n');
fprintf(finalout, '%s', allLogText); % Write accumulated log text
fclose(finalout);


%% Prompt user to export Blackboard CSV

promptAndGenerateBlackboardCSV(folderAnalyzed, files, points, feedback, timestamp);



%%  Create a histogram with 10 bins
figure;  % Open a new figure window
histogram(points, 10);
% Add labels and title
xlabel('Scores');
ylabel('Count');
title('Distribution of Scores');

duration=toc;
fprintf('Average time was %0.1f seconds per cadet\n',duration/numel(files))
%% Give link to the log file
fprintf('Open the output file here:\n <a href="matlab:system(''notepad %s'')">%s</a>\n', logFilePath, logFilePath);




%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%Embedded functions%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% This is the main code that does all the evaluations. It is here so
% it can be called using a for loop for one file, and a parfor loop for
% many files.
function [pt, fb] = gradeCadet(filename) % Read the sheet

sheets = loadAllJet11Sheets(filename);

Aero = sheets.Aero;
Miss = sheets.Miss;
Main = sheets.Main;
Consts = sheets.Consts;
Gear = sheets.Gear;
Geom = sheets.Geom;

pt = 40;
logText = "";
bonusPoints = 0;

[~, name, ext] = fileparts(filename);
logText = logf(logText, '%s\n', [name, ext]);

tol = 1e-3;
altTol = 1;
machTol = 1e-2;
timeTol = 1e-2;
distTol = 1e-3;
betaDefault = 0.87620980519917;

ConstraintsMach = Main(4, 21);
radius = Main(37, 25);
aim120 = Main(3, 28);
aim9 = Main(4, 28);
takeoff_dist = Main(12, 24);
landing_dist = Main(13, 24);
fuel_available = Main(18, 15);
fuel_required = Main(40, 24);
volume_remaining = Main(23, 17);
cost = Main(31, 17);
numaircraft = Main(31, 14);

% Aero tab programming (3 pts)
aeroIssues = 0;
if isequal(Aero(3,7), Aero(4,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(10,7), Aero(11,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(15,1), Aero(16,1)), aeroIssues = aeroIssues + 1; end

aeroDeduction = min(3, aeroIssues);
if aeroDeduction > 0
    pt = pt - aeroDeduction;
    logText = logf(logText, '-%d pts Aero tab formulas not active in cells A15, G3, or G10\n', aeroDeduction);
end
% Mission table (2 pts max deduction)
MissionArray = Main(33:44, 11:25);
colIdx = [1, 2, 3, 4, 6, 8, 9, 12, 13];
alt = MissionArray(1, colIdx);
mach = MissionArray(3, colIdx);
ab = MissionArray(4, colIdx);
dist = MissionArray(6, colIdx);
timeLeg = MissionArray(7, colIdx);

missionErrors = 0;

if abs(alt(1)) > altTol || abs(ab(1) - 100) > tol
    logText = logf(logText, 'Mission leg 1 must use altitude 0 ft and AB = 100%%\n');
    missionErrors = missionErrors + 1;
end

if alt(2) < min(alt(1), alt(3)) - altTol || alt(2) > max(alt(1), alt(3)) + altTol
    logText = logf(logText, 'Mission leg 2 altitude must remain between legs 1 and 3\n');
    missionErrors = missionErrors + 1;
end
if mach(2) < min(mach(1), mach(3)) - machTol || mach(2) > max(mach(1), mach(3)) + machTol
    logText = logf(logText, 'Mission leg 2 Mach must remain between legs 1 and 3\n');
    missionErrors = missionErrors + 1;
end
if abs(ab(2)) > tol
    logText = logf(logText, 'Mission leg 2 must use AB = 0%%\n');
    missionErrors = missionErrors + 1;
end

if alt(3) < 35000 - altTol || abs(mach(3) - 0.9) > machTol || abs(ab(3)) > tol
    logText = logf(logText, 'Mission leg 3 must be >= 35,000 ft, Mach 0.9, AB = 0%%\n');
    missionErrors = missionErrors + 1;
end

if alt(4) < 35000 - altTol || abs(mach(4) - 0.9) > machTol || abs(ab(4)) > tol
    logText = logf(logText, 'Mission leg 4 must be >= 35,000 ft, Mach 0.9, AB = 0%%\n');
    missionErrors = missionErrors + 1;
end

if alt(5) < 35000 - altTol || abs(mach(5) - ConstraintsMach) > machTol || abs(ab(5)) > tol || dist(5) < 150 - distTol
    logText = logf(logText, 'Mission leg 5 must be >= 35,000 ft, match the constraint supercruise Mach, AB = 0%%, distance >= 150 nm\n');
    missionErrors = missionErrors + 1;
end

if alt(6) < 30000 - altTol || mach(6) < 1.2 - machTol || abs(ab(6) - 100) > tol || timeLeg(6) < 2 - timeTol
    logText = logf(logText, 'Mission leg 6 must be >= 30,000 ft, Mach >= 1.2, AB = 100%%, time >= 2 min\n');
    missionErrors = missionErrors + 1;
end

if alt(7) < 35000 - altTol || abs(mach(7) - ConstraintsMach) > machTol || abs(ab(7)) > tol || dist(7) < 150 - distTol
    logText = logf(logText, 'Mission leg 7 must be >= 35,000 ft, match the constraint supercruise Mach, AB = 0%%, distance >= 150 nm\n');
    missionErrors = missionErrors + 1;
end

if alt(8) < 35000 - altTol || abs(mach(8) - 0.9) > machTol || abs(ab(8)) > tol
    logText = logf(logText, 'Mission leg 8 must be >= 35,000 ft, Mach 0.9, AB = 0%%\n');
    missionErrors = missionErrors + 1;
end

if abs(alt(9) - 10000) > altTol || abs(mach(9) - 0.4) > machTol || abs(ab(9)) > tol || abs(timeLeg(9) - 20) > timeTol
    logText = logf(logText, 'Mission leg 9 must be 10,000 ft, Mach 0.4, AB = 0%%, time = 20 min\n');
    missionErrors = missionErrors + 1;
end

if radius < 375 - distTol
    logText = logf(logText, 'Mission radius must be at least 375 nm (found %.1f)\n', radius);
    missionErrors = missionErrors + 1;
end

missionDeduction = min(2, missionErrors);
if missionDeduction > 0
    pt = pt - missionDeduction;
    logText = logf(logText, '-%d pts Mission profile inputs incorrect (max 2)\n', missionDeduction);
end

% Tavailable > D check (3 pts)
thrust_drag = Miss(48:49, 3:14);
thrustShort = thrust_drag(2, :) <= thrust_drag(1, :);
thrustFailures = sum(thrustShort);
if thrustFailures > 0
    deduction = min(3, thrustFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Not enough thrust for %d mission segment(s) (Tavailable <= D)\n', deduction, thrustFailures);
end
% Constraint table values (2 pts max deduction)
constraintErrors = 0;

if Main(3,20) < 35000 - altTol
    logText = logf(logText, 'Constraint Max Mach altitude must be >= 35,000 ft (found %.0f)\n', Main(3,20));
    constraintErrors = constraintErrors + 1;
end
if Main(3,21) < 2.0 - machTol
    logText = logf(logText, 'Constraint Max Mach requires Mach >= 2.0 (found %.2f)\n', Main(3,21));
    constraintErrors = constraintErrors + 1;
end
if Main(3,22) < 1 - tol
    logText = logf(logText, 'Constraint Max Mach requires n >= 1 (found %.2f)\n', Main(3,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(3,23) - 100) > tol
    logText = logf(logText, 'Constraint Max Mach requires AB = 100%% (found %.0f%%)\n', Main(3,23));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(3,24)) > tol
    logText = logf(logText, 'Constraint Max Mach requires Ps = 0 (found %.0f)\n', Main(3,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(3,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Max Mach W/WTO must remain at default (found %.3f)\n', Main(3,19));
    constraintErrors = constraintErrors + 1;
end

if Main(4,20) < 35000 - altTol
    logText = logf(logText, 'Constraint Supercruise altitude must be >= 35,000 ft (found %.0f)\n', Main(4,20));
    constraintErrors = constraintErrors + 1;
end
if Main(4,21) < 1.5 - machTol
    logText = logf(logText, 'Constraint Supercruise requires Mach >= 1.5 (found %.2f)\n', Main(4,21));
    constraintErrors = constraintErrors + 1;
end
if Main(4,22) < 1 - tol
    logText = logf(logText, 'Constraint Supercruise requires n >= 1 (found %.2f)\n', Main(4,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(4,23)) > tol
    logText = logf(logText, 'Constraint Supercruise requires AB = 0%% (found %.0f%%)\n', Main(4,23));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(4,24)) > tol
    logText = logf(logText, 'Constraint Supercruise requires Ps = 0 (found %.0f)\n', Main(4,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(4,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Supercruise W/WTO must remain at default (found %.3f)\n', Main(4,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(6,20) - 30000) > altTol
    logText = logf(logText, 'Constraint Combat Turn 1 altitude must equal 30,000 ft (found %.0f)\n', Main(6,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(6,21) - 1.2) > machTol
    logText = logf(logText, 'Constraint Combat Turn 1 requires Mach = 1.2 (found %.2f)\n', Main(6,21));
    constraintErrors = constraintErrors + 1;
end
if Main(6,22) < 3 - tol
    logText = logf(logText, 'Constraint Combat Turn 1 requires n >= 3 (found %.2f)\n', Main(6,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(6,23) - 100) > tol
    logText = logf(logText, 'Constraint Combat Turn 1 requires AB = 100%% (found %.0f%%)\n', Main(6,23));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(6,24)) > tol
    logText = logf(logText, 'Constraint Combat Turn 1 requires Ps = 0 (found %.0f)\n', Main(6,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(6,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Combat Turn 1 W/WTO must remain at default (found %.3f)\n', Main(6,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(7,20) - 10000) > altTol
    logText = logf(logText, 'Constraint Combat Turn 2 altitude must equal 10,000 ft (found %.0f)\n', Main(7,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(7,21) - 0.9) > machTol
    logText = logf(logText, 'Constraint Combat Turn 2 requires Mach = 0.9 (found %.2f)\n', Main(7,21));
    constraintErrors = constraintErrors + 1;
end
if Main(7,22) < 4 - tol
    logText = logf(logText, 'Constraint Combat Turn 2 requires n >= 4 (found %.2f)\n', Main(7,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(7,23) - 100) > tol
    logText = logf(logText, 'Constraint Combat Turn 2 requires AB = 100%% (found %.0f%%)\n', Main(7,23));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(7,24)) > tol
    logText = logf(logText, 'Constraint Combat Turn 2 requires Ps = 0 (found %.0f)\n', Main(7,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(7,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Combat Turn 2 W/WTO must remain at default (found %.3f)\n', Main(7,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(8,20) - 30000) > altTol
    logText = logf(logText, 'Constraint Ps1 altitude must equal 30,000 ft (found %.0f)\n', Main(8,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(8,21) - 1.15) > machTol
    logText = logf(logText, 'Constraint Ps1 requires Mach = 1.15 (found %.2f)\n', Main(8,21));
    constraintErrors = constraintErrors + 1;
end
if Main(8,22) < 1 - tol
    logText = logf(logText, 'Constraint Ps1 requires n >= 1 (found %.2f)\n', Main(8,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(8,23) - 100) > tol
    logText = logf(logText, 'Constraint Ps1 requires AB = 100%% (found %.0f%%)\n', Main(8,23));
    constraintErrors = constraintErrors + 1;
end
if Main(8,24) < 400 - distTol
    logText = logf(logText, 'Constraint Ps1 requires Ps >= 400 ft/s (found %.0f)\n', Main(8,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(8,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Ps1 W/WTO must remain at default (found %.3f)\n', Main(8,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(9,20) - 10000) > altTol
    logText = logf(logText, 'Constraint Ps2 altitude must equal 10,000 ft (found %.0f)\n', Main(9,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(9,21) - 0.9) > machTol
    logText = logf(logText, 'Constraint Ps2 requires Mach = 0.9 (found %.2f)\n', Main(9,21));
    constraintErrors = constraintErrors + 1;
end
if Main(9,22) < 1 - tol
    logText = logf(logText, 'Constraint Ps2 requires n >= 1 (found %.2f)\n', Main(9,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(9,23)) > tol
    logText = logf(logText, 'Constraint Ps2 requires AB = 0%% (found %.0f%%)\n', Main(9,23));
    constraintErrors = constraintErrors + 1;
end
if Main(9,24) < 400 - distTol
    logText = logf(logText, 'Constraint Ps2 requires Ps >= 400 ft/s (found %.0f)\n', Main(9,24));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(9,19) - betaDefault) > 1e-3
    logText = logf(logText, 'Constraint Ps2 W/WTO must remain at default (found %.3f)\n', Main(9,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(12,20)) > altTol
    logText = logf(logText, 'Constraint Takeoff altitude must equal 0 ft (found %.0f)\n', Main(12,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,21) - 1.2) > machTol
    logText = logf(logText, 'Constraint Takeoff requires V/Vstall = 1.2 (found %.2f)\n', Main(12,21));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,22) - 0.03) > 5e-4
    logText = logf(logText, 'Constraint Takeoff requires mu = 0.03 (found %.3f)\n', Main(12,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,23) - 100) > tol
    logText = logf(logText, 'Constraint Takeoff requires AB = 100%% (found %.0f%%)\n', Main(12,23));
    constraintErrors = constraintErrors + 1;
end
if takeoff_dist > 3000 + distTol
    logText = logf(logText, 'Constraint Takeoff distance must be <= 3000 ft (found %.0f)\n', takeoff_dist);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,19) - 1) > tol
    logText = logf(logText, 'Constraint Takeoff W/WTO must remain at default (found %.3f)\n', Main(12,19));
    constraintErrors = constraintErrors + 1;
end

cdxTakeoff = Main(12,25);
if ~(abs(cdxTakeoff) <= tol || abs(cdxTakeoff - 0.035) <= tol)
    logText = logf(logText, 'Constraint Takeoff CDx must be 0 or 0.035 (found %.3f)\n', cdxTakeoff);
    constraintErrors = constraintErrors + 1;
end

if abs(Main(13,20)) > altTol
    logText = logf(logText, 'Constraint Landing altitude must equal 0 ft (found %.0f)\n', Main(13,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,21) - 1.3) > machTol
    logText = logf(logText, 'Constraint Landing requires V/Vstall = 1.3 (found %.2f)\n', Main(13,21));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,22) - 0.5) > tol
    logText = logf(logText, 'Constraint Landing requires mu = 0.5 (found %.3f)\n', Main(13,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,23)) > tol
    logText = logf(logText, 'Constraint Landing requires AB = 0%% (found %.0f%%)\n', Main(13,23));
    constraintErrors = constraintErrors + 1;
end
if landing_dist > 5000 + distTol
    logText = logf(logText, 'Constraint Landing distance must be <= 5000 ft (found %.0f)\n', landing_dist);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,19) - 1) > tol
    logText = logf(logText, 'Constraint Landing W/WTO must remain at default (found %.3f)\n', Main(13,19));
    constraintErrors = constraintErrors + 1;
end

cdxLanding = Main(13,25);
if ~(abs(cdxLanding) <= tol || abs(cdxLanding - 0.045) <= tol)
    logText = logf(logText, 'Constraint Landing CDx must be 0 or 0.045 (found %.3f)\n', cdxLanding);
    constraintErrors = constraintErrors + 1;
end

constraintDeduction = min(2, constraintErrors);
if constraintDeduction > 0
    pt = pt - constraintDeduction;
    logText = logf(logText, '-%d pts Constraint table entries do not match the RFP (max 2)\n', constraintDeduction);
end
% Constraint curve compliance (8 pts)
constraintCurveFailures = 0;
failedCurves = {};

try
    WS_axis = Consts(22, 11:31);
    WS_axis = double(WS_axis);

    constraintRows = [23, 24, 26, 27, 28, 29, 32];
    columnLabels = {"MaxMach", "Supercruise", "CombatTurn1", "CombatTurn2", "Ps1", "Ps2", "Takeoff"};

    WS_design = Main(13, 16);
    TW_design = Main(13, 17);

    for idx = 1:numel(constraintRows)
        row = constraintRows(idx);
        TW_curve = Consts(row, 11:31);
        TW_curve = double(TW_curve);
        estimatedTWvalue = interp1(WS_axis, TW_curve, WS_design, 'pchip', 'extrap');
        if TW_design < estimatedTWvalue
            constraintCurveFailures = constraintCurveFailures + 1;
            failedCurves{end+1} = columnLabels{idx}; %#ok<AGROW>
        end
    end

    WS_limit_landing = Consts(33, 12);
    if WS_design > WS_limit_landing
        constraintCurveFailures = constraintCurveFailures + 1;
        failedCurves{end+1} = 'Landing'; %#ok<AGROW>
        logText = logf(logText, 'Landing constraint violated: W/S = %.2f exceeds limit of %.2f\n', WS_design, WS_limit_landing);
    end
catch ME
    logText = logf(logText, 'Constraint curve check skipped (error: %s)\n', ME.message);
    constraintCurveFailures = 0;
    failedCurves = {};
end

if constraintCurveFailures == 1
    pt = pt - 4;
    logText = logf(logText, '-4 pts Design falls below constraint %s\n', char(failedCurves{1}));
elseif constraintCurveFailures >= 2
    pt = pt - 8;
    summary = strjoin(failedCurves, ', ');
    logText = logf(logText, '-8 pts Design falls below multiple constraints: %s\n', char(summary));
end

% Payload (4 pts)
if isnan(aim120) || aim120 < 8 - tol
    count = aim120;
    if isnan(count), count = 0; end
    pt = pt - 4;
    logText = logf(logText, '-4 pts Payload must include at least 8 AIM-120Ds (found %.0f)\n', count);
end
% Control surface attachment (2 pts)
controlFailures = 0;

fuselage_length = Main(32, 2);
fuselage_end = fuselage_length;
PCS_x = Main(23, 3);
PCS_root_chord = Geom(8, 3);
if any(isnan([fuselage_end, PCS_x, PCS_root_chord]))
    logText = logf(logText, 'Unable to verify PCS placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_x > (fuselage_end - 0.25 * PCS_root_chord)
    logText = logf(logText, 'PCS X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end

VT_x = Main(23, 8);
VT_root_chord = Geom(10, 3);
if any(isnan([fuselage_end, VT_x, VT_root_chord]))
    logText = logf(logText, 'Unable to verify vertical tail placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif VT_x > (fuselage_end - 0.25 * VT_root_chord)
    logText = logf(logText, 'VT X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end

PCS_z = Main(25, 3);
fuse_z_center = Main(52, 4);
fuse_z_height = Main(52, 6);
if any(isnan([PCS_z, fuse_z_center, fuse_z_height]))
    logText = logf(logText, 'Unable to verify PCS vertical placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_z < (fuse_z_center - fuse_z_height/2) || PCS_z > (fuse_z_center + fuse_z_height/2)
    logText = logf(logText, 'PCS Z-location outside fuselage vertical bounds.\n');
    controlFailures = controlFailures + 1;
end

VT_y = Main(24, 8);
fuse_width = Main(52, 5);
if any(isnan([VT_y, fuse_width]))
    logText = logf(logText, 'Unable to verify vertical tail lateral placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif VT_y > fuse_width/2
    logText = logf(logText, 'VT Y-location outside fuselage width.\n');
    controlFailures = controlFailures + 1;
end

if Main(18, 4) > 1
    sweep = Geom(15, 11);
    y = Geom(152, 13);
    strake = Geom(155, 12);
    apex = Geom(38, 12);
    if any(isnan([sweep, y, strake, apex]))
        logText = logf(logText, 'Unable to verify strake attachment due to missing geometry data\n');
        controlFailures = controlFailures + 1;
    else
        wing = (y / tand(90 - sweep) + apex);
        if wing >= (strake + 0.5)
            logText = logf(logText, 'Strake disconnected\n');
            controlFailures = controlFailures + 1;
        end
    end
end

component_positions = Main(23, 2:8);
if any(component_positions >= fuselage_end)
    logText = logf(logText, 'One or more component X-locations extend beyond the fuselage end (B32 = %.2f)\n', fuselage_end);
    controlFailures = controlFailures + 1;
end

if controlFailures > 0
    deduction = min(2, controlFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Control surface placement issues (max 2)\n', deduction);
end

% Stability (3 pts)
SM = Main(10, 13);
clb = Main(10, 15);
cnb = Main(10, 16);
rat = Main(10, 17);

stabilityErrors = 0;
if ~(SM >= -0.1 && SM <= 0.11)
    logText = logf(logText, 'Static margin out of bounds (M10 = %.3f)\n', SM);
    stabilityErrors = stabilityErrors + 1;
    if SM < 0
        logText = logf(logText, 'Warning: aircraft is statically unstable (SM < 0)\n');
    end
end
if clb >= -0.001
    logText = logf(logText, 'Clb must be < -0.001 (O10 = %.6f)\n', clb);
    stabilityErrors = stabilityErrors + 1;
end
if cnb <= 0.002
    logText = logf(logText, 'Cnb must be > 0.002 (P10 = %.6f)\n', cnb);
    stabilityErrors = stabilityErrors + 1;
end
if ~(rat >= -1 && rat <= -0.3)
    logText = logf(logText, 'Cnb/Clb ratio must be between -1 and -0.3 (Q10 = %.3f)\n', rat);
    stabilityErrors = stabilityErrors + 1;
end

stabilityDeduction = min(3, stabilityErrors);
if stabilityDeduction > 0
    pt = pt - stabilityDeduction;
    logText = logf(logText, '-%d pts Stability parameters outside limits (max 3)\n', stabilityDeduction);
end

% Fuel (2 pts) and volume (2 pts)
if isnan(fuel_available) || isnan(fuel_required) || fuel_available + tol < fuel_required
    pt = pt - 2;
    if isnan(fuel_available) || isnan(fuel_required)
        logText = logf(logText, '-2 pts Unable to verify fuel available versus required\n');
    else
        logText = logf(logText, '-2 pts Fuel available (%.1f) is less than required (%.1f)\n', fuel_available, fuel_required);
    end
end

if isnan(volume_remaining) || volume_remaining <= 0
    pt = pt - 2;
    logText = logf(logText, '-2 pts Volume remaining must be positive (Q23 = %.2f)\n', volume_remaining);
end

% Recurring cost (5 pts)
costDeduction = 0;
if abs(numaircraft - 187) < 1e-3
    if isnan(cost)
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost missing for 187-aircraft estimate\n');
    elseif cost > 115 + tol
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost exceeds $115M (found $%.1fM)\n', cost);
    elseif cost <= 100 + tol
        logText = logf(logText, 'Recurring cost meets objective (<=$100M): $%.1fM\n', cost);
    end
elseif abs(numaircraft - 800) < 1e-3
    if isnan(cost)
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost missing for 800-aircraft estimate\n');
    elseif cost > 75 + tol
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost exceeds $75M (found $%.1fM)\n', cost);
    elseif cost <= 63 + tol
        logText = logf(logText, 'Recurring cost meets objective (<=$63M): $%.1fM\n', cost);
    end
else
    costDeduction = 5;
    logText = logf(logText, '-5 pts Number of aircraft (N31) must be 187 or 800 (found %.0f)\n', numaircraft);
end

if costDeduction > 0
    pt = pt - costDeduction;
end

% Landing gear geometry (4 pts)
gearFailures = 0;

g90 = Gear(20, 10);
if isnan(g90) || g90 < 80 - tol || g90 > 95 + tol
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Nose gear 90/10 rule must be between 80%% and 95%% (found %.1f%%)\n', g90);
end

tipbackActual = Gear(20, 12);
tipbackLimit = Gear(21, 12);
if isnan(tipbackActual) || isnan(tipbackLimit) || tipbackActual >= tipbackLimit - 1e-2
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Tipback angle must remain below limit (%.2f >= %.2f)\n', tipbackActual, tipbackLimit);
end

rolloverActual = Gear(20, 13);
rolloverLimit = Gear(21, 13);
if isnan(rolloverActual) || isnan(rolloverLimit) || rolloverActual >= rolloverLimit - 1e-2
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Rollover angle must remain below limit (%.2f >= %.2f)\n', rolloverActual, rolloverLimit);
end

rotationSpeed = Gear(20, 14);
if isnan(rotationSpeed) || rotationSpeed >= 200 - tol
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Takeoff rotation speed must be less than 200 knots (found %.1f)\n', rotationSpeed);
end

if gearFailures > 0
    deduction = min(4, gearFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Landing gear geometry outside limits (max 4)\n', deduction);
end
% Final score and bonuses
baseScore = max(0, pt);
pt = baseScore;
bonusPoints = 0;

if radius >= 410 - distTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Mission radius objective met (%.1f nm)\n', radius);
end
if ~isnan(aim120) && ~isnan(aim9) && aim120 >= 8 - tol && aim9 >= 2 - tol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Payload objective met (%.0f AIM-120, %.0f AIM-9)\n', aim120, aim9);
end
if takeoff_dist <= 2500 + distTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Takeoff distance objective met (%.0f ft)\n', takeoff_dist);
end
if landing_dist <= 3500 + distTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Landing distance objective met (%.0f ft)\n', landing_dist);
end
if Main(3,21) >= 2.2 - machTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Max Mach objective met (%.2f)\n', Main(3,21));
end
if Main(4,21) >= 1.8 - machTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Supercruise Mach objective met (%.2f)\n', Main(4,21));
end
if Main(8,24) >= 500 - distTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Ps objective met at 30k ft (%.0f ft/s)\n', Main(8,24));
end
if Main(9,24) >= 500 - distTol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Ps objective met at 10k ft (%.0f ft/s)\n', Main(9,24));
end
if Main(6,22) >= 4 - tol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Combat turn objective met at 30k ft (%.2f g)\n', Main(6,22));
end
if Main(7,22) >= 4.5 - tol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Combat turn objective met at 10k ft (%.2f g)\n', Main(7,22));
end
if ~isnan(cost) && cost <= 100 + tol
    bonusPoints = bonusPoints + 1;
    logText = logf(logText, '+1 bonus Cost objective met (%.1f M$)\n', cost);
end

pt = baseScore + bonusPoints;

logText = logf(logText, 'Jet11 base score: %d out of 40\n', baseScore);
if bonusPoints > 0
    logText = logf(logText, 'Bonus objectives earned: +%d (final score %d)\n', bonusPoints, pt);
else
    logText = logf(logText, 'No bonus objectives earned; final score %d\n', pt);
end

fprintf('%s completed\n', [name, ext]);
fprintf('Jet11 base score: %d / 40\n', baseScore);
fprintf('Jet11 final score (bonus applied): %d\n\n', pt);

fb = char(logText);
end
%% Function to read all useful sheets, and verify numbers returned for used cells
function sheets = loadAllJet11Sheets(filename)
%LOADALLJET11SHEETS Load all required Jet11 sheets using safeReadMatrix
%   sheets = loadAllJet11Sheets(filename) returns a struct with fields:
%   Aero, Miss, Main, Consts, Gear, Geom

sheets.Aero   = safeReadMatrix(filename, 'Aero',   {'G3','G4','G10','G11','A15','A16'});
sheets.Miss   = safeReadMatrix(filename, 'Miss',   {'C48','C49'});
sheets.Main   = safeReadMatrix(filename, 'Main',   {'S3','T3','U3','V3','W3','X3','Y3','S4','T4','U4','V4','W4','X4','Y4',...
    'S5','S6','S7','S8','S9','T6','U6','V6','W6','X6','Y6','T7','U7','V7',...
    'W7','X7','Y7','T8','U8','V8','W8','X8','Y8','T9','U9','V9','W9','X9',...
    'Y9','S12','S13','AB3','AB4','X12','X13','Y37','M10','O10','P10','Q10',...
    'O18','X40','Q23','Q31','N31','B32','C23','H23','D18','D23','D52','F52',...
    'H24','E52'});
sheets.Consts = safeReadMatrix(filename, 'Consts', {'K22','K23','K24','K26','K27','K28','K29','K32','AO42','AQ41','K33'});
sheets.Gear   = safeReadMatrix(filename, 'Gear',   {'J20','L20','L21','M20','M21','N20'});
sheets.Geom   = safeReadMatrix(filename, 'Geom',   {'C8','C10','M152','K15','L155','L38'});

% Constants is off by three rows. Row 22 of the Consts tab comes in as
% row 19 in matlab Consts variable. Adding three rows of NaN to the top
% so it can be addressed accurately.

% sheets.Consts = [NaN(3, size(sheets.Consts, 2)); sheets.Consts];

end

%% Function to read the data from the excel sheets as quickly and accurately as possible
function data = safeReadMatrix(filename, sheetname, fallbackCells)
% safeReadMatrix - Efficiently reads numeric data from an Excel sheet.
%   Attempts fast readmatrix first. If key cells are NaN, falls back to readcell.
%
% Inputs:
%   filename      - Excel file path
%   sheetname     - Sheet name to read
%   fallbackCells - Cell array of cell references to verify (e.g., {'G4', 'G10'})
%
% Output:
%   data - Numeric matrix with fallback values patched in if needed

% Try fast read
% if strcmp(sheetname,'Gear')
%     data = readmatrix(filename, 'Sheet', sheetname,'DataRange','A1:M155');
% else
%     data = readmatrix(filename, 'Sheet', sheetname,'DataRange','A1:AQ52');
% end
data = readmatrix(filename, 'Sheet', sheetname,'DataRange','A1:AQ155');


% Convert cell references to row/col indices
fallbackIndices = cellfun(@(c) cell2sub(c), fallbackCells, 'UniformOutput', false);

% Check for NaNs in fallback cells
needsPatch = false;
for i = 1:numel(fallbackIndices)
    idx = fallbackIndices{i};
    if idx(1) > size(data,1) || idx(2) > size(data,2) || isnan(data(idx(1), idx(2)))
        needsPatch = true;
        %                 fprintf('Patching %s %d\n',sheetname, idx)
        fprintf('Patched %s cell %s with value %.4f\n', sheetname, sub2excel(idx(1), idx(2)), data(idx(1), idx(2)));
        break;
    end
end

% If needed, patch from readcell
if needsPatch
    raw = readcell(filename, 'Sheet', sheetname);
    for i = 1:numel(fallbackIndices)
        idx = fallbackIndices{i};
        if idx(1) <= size(raw,1) && idx(2) <= size(raw,2)
            val = raw{idx(1), idx(2)};
            if isnumeric(val)
                data(idx(1), idx(2)) = val;
            elseif ischar(val) || isstring(val)
                data(idx(1), idx(2)) = str2double(val);
            end
        end
    end
end
end

function idx = cell2sub(cellref)
% Converts Excel cell reference (e.g., 'G4') to row/col indices
col = regexp(cellref, '[A-Z]+', 'match', 'once');
row = str2double(regexp(cellref, '\d+', 'match', 'once'));
colNum = 0;
for i = 1:length(col)
    colNum = colNum * 26 + (double(col(i)) - double('A') + 1);
end
idx = [row, colNum];
end

function ref = sub2excel(row, col)
letters = '';
while col > 0
    rem = mod(col - 1, 26);
    letters = [char(65 + rem), letters]; %#ok<AGROW>
    col = floor((col - 1) / 26);
end
ref = sprintf('%s%d', letters, row);
end

%% Function to do an fprintf like function to a local variable for future use
function logText = logf(logText, varargin)
logEntry = sprintf(varargin{:});  % Format input like fprintf
logText = [logText, logEntry];      % Append to string
end


function [mode, selectedPath] = selectRunMode()
% SELECTRUNMODE - Launches a GUI to choose between single file or folder mode



cursorPos = get(0, 'PointerLocation');
dialogWidth = 300;
dialogHeight = 150;

% Position just below the cursor
dialogLeft = cursorPos(1) - dialogWidth / 2;
dialogBottom = cursorPos(2) - dialogHeight - 20;  % 20 pixels below the cursor

d = dialog('Position', [dialogLeft, dialogBottom, dialogWidth, dialogHeight], ...
    'Name', 'Select Run Mode');


txt = uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 90 260 40],...
    'String','Choose how you want to run the autograder:',...
    'FontSize',10); %#ok<NASGU>

btn1 = uicontrol('Parent',d,...
    'Position',[30 40 100 30],...
    'String','Single File',...
    'Callback',@singleFile); %#ok<NASGU>

btn2 = uicontrol('Parent',d,...
    'Position',[170 40 100 30],...
    'String','Folder of Files',...
    'Callback',@folderRun); %#ok<NASGU>

mode = '';
selectedPath = '';

uiwait(d);  % Wait for user to close dialog

    function singleFile(~,~)
        [file, path] = uigetfile('*.xls*','Select a Jet11 Excel file');
        if isequal(file,0)
            mode = 'cancelled';
        else
            mode = 'single';
            selectedPath = fullfile(path, file);
        end
        delete(d);
    end

    function folderRun(~,~)
        path = uigetdir(pwd, 'Select folder containing Jet11 files');
        if isequal(path,0)
            mode = 'cancelled';
        else
            mode = 'folder';
            selectedPath = path;
        end
        delete(d);
    end
end

%% Prompt user and generate Blackboard CSV (combined function)
function promptAndGenerateBlackboardCSV(folderAnalyzed, files, points, feedback, timestamp)
% Position dialog below cursor
cursorPos = get(0, 'PointerLocation');
dialogWidth = 300;
dialogHeight = 150;
dialogLeft = cursorPos(1) - dialogWidth / 2;
dialogBottom = cursorPos(2) - dialogHeight - 20;

% Create dialog
d = dialog('Position', [dialogLeft, dialogBottom, dialogWidth, dialogHeight], ...
    'Name', 'Blackboard Export');

uicontrol('Parent', d, ...
    'Style', 'text', ...
    'Position', [20 90 260 40], ...
    'String', 'Generate Blackboard CSV for grade import?', ...
    'FontSize', 10);

uicontrol('Parent', d, ...
    'Position', [30 40 100 30], ...
    'String', 'Yes', ...
    'Callback', @(~,~) doExport(true, d));

uicontrol('Parent', d, ...
    'Position', [170 40 100 30], ...
    'String', 'No', ...
    'Callback', @(~,~) doExport(false, d));

    function doExport(shouldExport, dialogHandle)
        delete(dialogHandle);
        if shouldExport
            %% Create Blackboard Offline Grade CSV (SMART_TEXT format)
            csvFilename = fullfile(folderAnalyzed, ['FinalProject_Blackboard_Offline_', timestamp, '.csv']);
            fid = fopen(csvFilename, 'w');

            % Assignment title column (update if needed)
            assignmentTitle = 'Final Project: AATF Design Iteration 1 & Cutout [Total Pts: 15 Score]';

            % Write header
            fprintf(fid, '"Last Name","First Name","Username","%s","Grading Notes","Notes Format","Feedback to Learner","Feedback Format"\n', assignmentTitle);

            for i = 1:numel(files)
                fname = files(i).name;

                % Extract username from filename
                tokens = regexp(fname, '_([a-z0-9\\.]+)_attempt_.*?_(.*?)_', 'tokens');
                if ~isempty(tokens)
                    username = tokens{1}{1};
                    fullName = strsplit(tokens{1}{2});
                    if numel(fullName) >= 2
                        firstName = strjoin(fullName(1:end-1), ' ');
                        lastName = fullName{end};
                    else
                        firstName = '';
                        lastName = fullName{1};
                    end
                else
                    username = 'UNKNOWN';
                    firstName = '';
                    lastName = '';
                end

                % Get score and feedback
                score = points(i) + 5;
                fbText = feedback{i};

                % Sanitize feedback for SMART_TEXT (HTML-safe but readable)
                fbText = strrep(fbText, 'â‰¥', '&ge;');
                fbText = strrep(fbText, 'â‰¤', '&le;');
                fbText = strrep(fbText, 'â‰ ', '&ne;');
                fbText = strrep(fbText, 'âœ”', '&#10004;');
                fbText = strrep(fbText, 'âœ˜', '&#10008;');
                fbText = strrep(fbText, 'âœ…', '&#9989;');
                fbText = strrep(fbText, 'âŒ', '&#10060;');
                fbText = strrep(fbText, '<', '&lt;');
                fbText = strrep(fbText, '>', '&gt;');
                fbText = strrep(fbText, '"', '&quot;');
                fbText = strrep(fbText, newline, '<br>');

                % Write row
                fprintf(fid, '"%s","%s","%s","%.2f","","","%s","SMART_TEXT"\n', ...
                    lastName, firstName, username, score, fbText);
            end

            fclose(fid);
            fprintf('Blackboard offline grade CSV created: %s\n', csvFilename);

        end
    end
end



