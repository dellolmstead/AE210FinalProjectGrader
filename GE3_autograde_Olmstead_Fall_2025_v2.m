
% GE3 Autograder: Grades AE210 Jet10 Excel submissions, logs feedback, and optionally exports Blackboard-compatible scores.


%--------------------------------------------------------------------------
% AE210 GE3 Autograder Script – Fall 2025
%
% Description:
% This script automates grading for AE210 preliminary Design Project 
% submissions (GE3) by processing Jet10 Excel files (*.xlsm). It evaluates 
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
% - Optional Blackboard CSV: GE3_Blackboard_Offline_<timestamp>.csv
%
% Embedded Functions:
% - gradeCadet: Grades a single cadet's file and returns score and feedback
% - loadAllJet10Sheets: Loads all required sheets from a Jet10 Excel file
% - safeReadMatrix: Robustly reads numeric data from Excel, with fallback to readcell
% - cell2sub: Converts Excel cell references (e.g., 'G4') to row/col indices
% - sub2excel: Converts row/col indices back to Excel cell references
% - logf: Appends formatted text to a log string
% - selectRunMode: GUI for selecting single file or folder mode
% - promptAndGenerateBlackboardCSV: Dialog + export to Blackboard SMART_TEXT format
%
% Author: Lt Col Dell Olmstead, based on work by Capt Carol Bryant and Capt Anna Mason
% Last Updated: 22 Jul 2025
%--------------------------------------------------------------------------
clear; close all; clc;


%% Choose directory and get Excel files
% fprintf('Executing %s\n',mfilename);
% I recommend updating the below line to point to your GE3 files. It works
% as is, but will default to the right place if this is updated.

% folderAnalyzed = uigetdir('C:\Users\dell.olmstead\OneDrive - afacademy.af.edu\Documents 1\01 Classes\AE210 FA24\Design Project\GE3 files');
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
fprintf(finalout, 'GE3 Autograder Log\n');
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

sheets = loadAllJet10Sheets(filename);

Aero = sheets.Aero;
Miss = sheets.Miss;
Main = sheets.Main;
Consts = sheets.Consts;
Gear = sheets.Gear;
Geom = sheets.Geom;


% Initialize local variables
pt = 10;  % Start with full score
logText = ""; % create the blank logtext for this entry analysis

% filename = fullfile(folderAnalyzed, files(cadetnum).name);
% fprintf('%s started\n', files(cadetnum).name);
[~, name, ext] = fileparts(filename);
logText = logf(logText, '%s\n', [name, ext]);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% --- Aero Tab Check (2 points) --- %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Logic: The aero check cells for programming are a rigid number so any
% changes to the aircraft will create a mismatch on those initial check
% numbers indicating sufficient success. If less than all are incorrect
% (likely didn't get coded) then deduct 1 point per error, maximum of 2 points.

flag = 0;
if isequal(Aero(3,7), Aero(4,7)), flag = flag + 1; end % Check cells G3 and G4 to make sure they no longer match indicating a live cell G3
if isequal(Aero(10,7), Aero(11,7)), flag = flag + 1; end
if isequal(Aero(15,1), Aero(16,1)), flag = flag + 1; end

pointdeduction = min(2, flag);  % Deduct 1 point per error, max 2
if pointdeduction > 0
    pt = pt - pointdeduction;
    logText = logf(logText, '-%d point Mismatch in Aero A15, G3, and G10  \n', pointdeduction);
end
%     fprintf('Aero Check Complete\n')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% --- Mission Table Checks From Attachment 1 (0 points)%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Logic: compare to the RFP Attachment 1 and ensure all specified
% parameters are met. Most legs have inequalities, so each leg is an
% independent logical comparison. The legs are not all of the legs in
% JET10, so it must be parsed to compare the correct legs.

MissionInputFailed = 0;
MissionArray = Main(33:44, 11:25); % J32:Y44
ConstraintsMach = Main(4, 21); % Constraint supercruise Mach cell

% Column mapping for legs 1–9: K, L, M, N, P, R, S, V, W
%K=Takeoff	Accel	Climb	Cruise	Patrol	Supercruise	Patrol	Combat	Supercruise	Patrol	Climb	Cruise 	W=Loiter

colIdx = [1, 2, 3, 4, 6, 8, 9, 12, 13];

% Extract mission data
alt = MissionArray(1, colIdx);
mach = MissionArray(3, colIdx);
ab = MissionArray(4, colIdx);
dist = MissionArray(6, colIdx);
time = MissionArray(7, colIdx);

%%%%%%%%%%%%%%% Leg 1: Preflight & Takeoff
if alt(1) ~= 0 || ab(1) ~= 100
    logText = logf(logText, 'Leg 1: Altitude must be 0 and AB = 100\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 2: Acceleration to climb
if ~(alt(2) >= alt(1) && alt(2) <= alt(3))
    logText = logf(logText, 'Leg 2: Altitude must be between Leg 1 and Leg 3\n');
    MissionInputFailed = MissionInputFailed + 1;
end
if ~(mach(2) >= mach(1) && mach(2) <= mach(3))
    logText = logf(logText, 'Leg 2: Mach must be between Leg 1 and Leg 3\n');
    MissionInputFailed = MissionInputFailed + 1;
end
if ab(2) ~= 0
    logText = logf(logText, 'Leg 2: AB must be 0\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 3: Climb to cruise
if alt(3) < 35000 || mach(3) ~= 0.9 || ab(3) ~= 0
    logText = logf(logText, 'Leg 3: Must be ≥35,000 ft, Mach = 0.9, AB = 0\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 4: Subsonic cruise
if alt(4) < 35000 || mach(4) ~= 0.9 || ab(4) ~= 0
    logText = logf(logText, 'Leg 4: Must be ≥35,000 ft, Mach = 0.9, AB = 0\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 5: Supercruise to target
if alt(5) < 35000 || abs(mach(5) - ConstraintsMach) > 0.01 || ab(5) ~= 0 || dist(5) < 150
    logText = logf(logText, 'Leg 5: Must be ≥35,000 ft, Mach = Contraints block Supercruise Mach (cell U4), AB = 0, Distance ≥ 150 nm\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 6: Combat
if alt(6) < 30000 || mach(6) < 1.2 || ab(6) ~= 100 || time(6) < 2
    logText = logf(logText, 'Leg 6: Must be ≥30,000 ft, Mach ≥ 1.2, AB = 100, Time ≥ 2 min\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 7: Supercruise egress
if alt(7) < 35000 || abs(mach(7) - ConstraintsMach) > 0.01 || ab(7) ~= 0 || dist(7) < 150
    logText = logf(logText, 'Leg 7: Must be ≥35,000 ft, Mach = Contraints block Supercruise Mach (cell U4), AB = 0, Distance ≥ 150 nm\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 8: Subsonic return cruise
if alt(8) < 35000 || mach(8) ~= 0.9 || ab(8) ~= 0
    logText = logf(logText, 'Leg 8: Must be ≥35,000 ft, Mach = 0.9, AB = 0\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Leg 9: Descent
if alt(9) ~= 10000 || mach(9) ~= 0.4 || ab(9) ~= 0 || time(9) ~= 20
    logText = logf(logText, 'Leg 9: Must be 10,000 ft, Mach = 0.4, AB = 0, Time = 20 min\n');
    MissionInputFailed = MissionInputFailed + 1;
end

%%%%%%%%%%%%%%% Final deduction
if MissionInputFailed > 0
    %         pt = pt - 1;
    logText = logf(logText, 'There is an error with your inputs to the OCA Mission that must be corrected. \n');
end
%%%%%%%%%

%     fprintf('Mission Analysis Check is Complete\n')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Mission Analysis (T > D), takeoff roll (1 point max)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

thrust_drag = Miss(48:49, 3:14); % C:N
if any(thrust_drag(1,:) > thrust_drag(2,:))
    pt = pt -1;
    logText = logf(logText,'-1 Point Not enough thrust for at least one leg!\n');
else
    takeoff_d = Main(38, 11); % K38
    takeoff_rq = Main(12, 24); % X12
    if takeoff_d > takeoff_rq
        logText = logf(logText,'-1 Point Too long for takeoff roll\n');
        pt = pt - 1;
    end
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Constraint Compliance Checks from RFP Attachment 2 (1 point)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% This checks radius, payload, TO distance, Landing distance
% independently, and then assesses all the criteria in the constraint
% table to ensure all meet threshold, and give feedback if failing or
% meeting objective.
constraintInputsFailed = 0;

% Mission Radius (Y37)
radius = Main(37, 25);
if radius < 375
    logText = logf(logText, 'Mission radius below threshold (375 nm): %.1f\n', radius);
    constraintInputsFailed=constraintInputsFailed+1;
elseif radius >= 410
    logText = logf(logText, 'Mission radius meets objective (410 nm): %.1f\n', radius);
else
    %         logText = logf(logText, 'Mission radius meets threshold: %.1f\n', radius);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Weapons Payload (AB3 = AIM-120s, AB4 = AIM-9s)
aim120 = Main(3, 28); % AB3
aim9 = Main(4, 28);   % AB4
if aim120 < 8
    logText = logf(logText, 'Fewer than 8 AIM-120s: %d\n', aim120);
    constraintInputsFailed=constraintInputsFailed+1;
elseif aim9 >= 2
    logText = logf(logText, 'Payload meets objective: %d AIM-120s + %d AIM-9s\n', aim120, aim9);
else
    %         logText = logf(logText, 'Payload meets threshold: %d AIM-120s\n', aim120);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Takeoff Distance (X13)
takeoff_dist = Main(12, 24);
if takeoff_dist > 3000
    logText = logf(logText, 'Takeoff distance exceeds threshold (3000 ft): %.0f\n', takeoff_dist);
    constraintInputsFailed=constraintInputsFailed+1;
elseif takeoff_dist <= 2500
    logText = logf(logText, 'Takeoff distance meets objective (≤2500 ft): %.0f\n', takeoff_dist);
else
    %         logText = logf(logText, 'Takeoff distance meets threshold: %.0f\n', takeoff_dist);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Landing Distance (X14)
landing_dist = Main(13, 24);
if landing_dist > 5000
    logText = logf(logText, 'Landing distance exceeds threshold (5000 ft): %.0f\n', landing_dist);
    constraintInputsFailed=constraintInputsFailed+1;
elseif landing_dist <= 3500
    logText = logf(logText, 'Landing distance meets objective (≤3500 ft): %.0f\n', landing_dist);
else
    %         logText = logf(logText, 'Landing distance meets threshold: %.0f\n', landing_dist);
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Integrated Constraint Table Validation ---
% Logic: Check all constraints in the Jet constraint table against
% thresholds, objectives, or exact matches as specified in the RFP
% Attachment 2. For objective/threshold values, the expected exact is a
% NaN that triggers it to be ignored in the exact comparison. For exact
% parameters, it has NaN in the threshold/objective columns so only the
% exact value will be compared. This will then loop through all 7
% specified constraints. 1 point is deducted if there are any errors.

% Define expected values for rows 1–7.
% ** Change this block to change the mission***************************
% Format: [row, Mach_min, Mach_obj, Mach_eq, Alt_eq, n_eq, n_min, n_obj, AB_eq, Ps_eq, Ps_min, Ps_obj, CDx_eq]
expected_constraints = [
    1, 2.0, 2.2, NaN, NaN, 1, NaN, NaN, 100, 0, NaN, NaN, 0;       % Max Mach
    2, 1.5, 1.8, NaN, NaN, 1, NaN, NaN,   0, 0, NaN, NaN, 0;       % Supercruise
    4, NaN, NaN, 1.20, 30000, NaN, 3.0, 4.0, 100, 0, NaN, NaN, 0;  % Combat Turn 1
    5, NaN, NaN, 0.90, 10000, NaN, 4.0, 4.5, 100, 0, NaN, NaN, 0;  % Combat Turn 2
    6, NaN, NaN, 1.15, 30000, 1, NaN, NaN, 100, NaN, 400, 500, 0;  % Ps1
    7, NaN, NaN, 0.90, 10000, 1, NaN, NaN,   0, NaN, 400, 500, 0   % Ps2
    ];

% Define constraint labels for rows 1–10
constraintLabels = {'MaxMach', 'CruiseMach','Cmbt Turn1', 'Cmbt Turn2', 'Ps1', 'Ps2'};

for constraintnum = 1:size(expected_constraints, 1)
    % Collect each requirement from the RFP into a unique variable
    row = expected_constraints(constraintnum, 1) + 2; % Main(3:10,...)
    mach_min = expected_constraints(constraintnum, 2);
    mach_obj = expected_constraints(constraintnum, 3);
    mach_eq  = expected_constraints(constraintnum, 4);
    alt_eq   = expected_constraints(constraintnum, 5);
    n_eq     = expected_constraints(constraintnum, 6);
    n_min    = expected_constraints(constraintnum, 7);
    n_obj    = expected_constraints(constraintnum, 8);
    ab_eq    = expected_constraints(constraintnum, 9);
    ps_eq    = expected_constraints(constraintnum,10);
    ps_min   = expected_constraints(constraintnum,11);
    ps_obj   = expected_constraints(constraintnum,12);
    cdx_eq   = expected_constraints(constraintnum,13);

    % Actual values from the JET sheet
    mach = Main(row, 21);   % T
    alt  = Main(row, 20);   % U
    n    = Main(row, 22);   % V
    ab   = Main(row, 23);   % W
    ps   = Main(row, 24);   % X
    cdx  = Main(row, 25);   % Y

    label = constraintLabels{constraintnum};  % i is the loop index for the constraint

    % Mach equality or threshold
    if ~isnan(mach_eq)
        if abs(mach - mach_eq) > 0.01
            logText = logf(logText, '%s: Mach = %.2f, expected %.2f\n', label, mach, mach_eq);
            constraintInputsFailed = constraintInputsFailed + 1;
        end
    elseif ~isnan(mach_min)
        if mach < mach_min
            logText = logf(logText, '%s: Mach = %.2f, must be ≥ %.2f\n', label, mach, mach_min);
            constraintInputsFailed = constraintInputsFailed + 1;
        elseif ~isnan(mach_obj) && mach >= mach_obj
            logText = logf(logText, '%s: Mach meets objective (≥ %.2f): %.2f\n', label, mach_obj, mach);
        end
    end

    % Altitude equality
    if ~isnan(alt_eq) && alt ~= alt_eq
        logText = logf(logText, '%s: Altitude = %.0f, expected %.0f\n', label, alt, alt_eq);
        constraintInputsFailed = constraintInputsFailed + 1;
    end

    % Load factor
    if ~isnan(n_eq)
        if n ~= n_eq
            logText = logf(logText, '%s: n = %.1f, expected %.1f\n', label, n, n_eq);
            constraintInputsFailed = constraintInputsFailed + 1;
        end
    elseif ~isnan(n_min)
        if n < n_min
            logText = logf(logText, '%s: g-load = %.1f, must be ≥ %.1f\n', label, n, n_min);
            constraintInputsFailed = constraintInputsFailed + 1;
        elseif ~isnan(n_obj) && n >= n_obj
            logText = logf(logText, '%s: g-load meets objective (≥ %.1f): %.1f\n', label, n_obj, n);
        end
    end

    % Afterburner equality
    if ~isnan(ab_eq) && ab ~= ab_eq
        logText = logf(logText, '%s: AB = %.0f%%, expected %.0f%%\n', label, ab, ab_eq);
        constraintInputsFailed = constraintInputsFailed + 1;
    end

    % Ps equality or threshold
    if ~isnan(ps_eq)
        if ps ~= ps_eq
            logText = logf(logText, '%s: Ps = %.0f, expected %.0f\n', label, ps, ps_eq);
            constraintInputsFailed = constraintInputsFailed + 1;
        end
    elseif ~isnan(ps_min)
        if ps < ps_min
            logText = logf(logText, '%s: Ps = %.0f, must be ≥ %.0f\n', label, ps, ps_min);
            constraintInputsFailed = constraintInputsFailed + 1;
        elseif ~isnan(ps_obj) && ps >= ps_obj
            logText = logf(logText, '%s: Ps meets objective (≥ %.0f): %.0f\n', label, ps_obj, ps);
        end
    end

    % CDx equality
    if ~isnan(cdx_eq) && abs(cdx - cdx_eq) > 0.001
        logText = logf(logText, '%s: CDx = %.3f, expected %.3f\n', label, cdx, cdx_eq);
        constraintInputsFailed = constraintInputsFailed + 1;
    end

end

if constraintInputsFailed > 0
    logText = logf(logText, '-1 Point One or more constraints mentioned above are incorrect\n');
    pt=pt-1;
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Constraint above curve check (0 points), if applicable
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
try
    % Extract W/S axis (x-values) from row 22, columns K–AE (11:31)
    WS_axis = Consts(22, 11:31);  % L22:AE22 (columns 11-31)
    WS_axis = double(WS_axis);

    % Define constraint rows - exclude 25 (MaxAlt), 30 (Ps3), and 31 (blank)
    constraintRows = [23, 24, 26, 27, 28, 29, 32];  % Skip rows 25, 30, 31
    columnLabels = ["MaxMach", "Supercruise", "CombatTurn1", ...
        "CombatTurn2", "Ps1", "Ps2", "Takeoff"];

    % Read design point from Main sheet
    WS_design = Main(13, 16);     % P13 - Design W/S
    TW_design = Main(13, 17);     % Q13 - Design T/W
    

    ConstraintsFailed = 0;
    WhichFailed = strings(1, length(constraintRows) + 1);  % +1 for landing
    numFailed = 0;

    % Loop through each constraint row (curves)
    for idx = 1:length(constraintRows)
        row = constraintRows(idx);
        TW_curve = Consts(row, 11:31);  % L:AE for this constraint (columns 11-31)
        TW_curve = double(TW_curve);

        % Interpolate required T/W at design W/S (no sorting - preserve order for plotting)
        estimatedTWvalue = interp1(WS_axis, TW_curve, WS_design, 'pchip', 'extrap');

        if TW_design < estimatedTWvalue
            % Design point is below the curve — fail
            ConstraintsFailed = ConstraintsFailed + 1;
            numFailed = numFailed + 1;
            WhichFailed(numFailed) = columnLabels(idx);
        end
    end

    % Landing constraint: Special case - vertical line at W/S limit
    % Design must be to the LEFT of (less than) this vertical line
    WS_limit_landing = Consts(33, 12);  % L33 - Landing W/S limit

    if WS_design > WS_limit_landing
        ConstraintsFailed = ConstraintsFailed + 1;
        numFailed = numFailed + 1;
        WhichFailed(numFailed) = "Landing";
        logText = logf(logText, 'Landing constraint violated: W/S = %.2f exceeds limit of %.2f\n', ...
            WS_design, WS_limit_landing);
    end

    % Trim WhichFailed array to actual size
    WhichFailed = WhichFailed(1:numFailed);

    % --- Generate Error Message if Constraints Were Violated ---
    if ConstraintsFailed > 0
        msg = "Design did not meet the following constraint" + ...
            (ConstraintsFailed > 1)*"s" + ": " + strjoin(WhichFailed, ', ');
        if ConstraintsFailed > 6
            msg = msg + ", among other issues. Consider seeking EI.";
        else
            msg = msg + ". Consider lowering your standards if above threshold.";
        end
        logText = logf(logText, '%s\n', msg);
        % pt = pt - 1;  % Uncomment if you want to deduct points
    end

catch ME
    logText = logf(logText, 'Could not perform constraint curve check due to error: %s\n', ME.message);
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% ---  Control Surface Attachment Check (1 point) %%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disconnected = 0;

% Check x-direction
fuselage_length = Main(32, 2); % B32
PCS_x = Main(23, 3);           % C23
PCS_root_chord = Geom(8, 3);   % C8

if PCS_x > (fuselage_length - 0.25 * PCS_root_chord)
    logText = logf(logText, 'PCS X-location too far aft. Must overlap at least 25%% of root chord.\n');
    disconnected = disconnected + 1;
end

VT_x = Main(23, 8);           % H23
VT_root_chord = Geom(10, 3); % C10

if VT_x > (fuselage_length - 0.25 * VT_root_chord)
    logText = logf(logText, 'VT X-location too far aft. Must overlap at least 25%% of root chord.\n');
    disconnected = disconnected + 1;
end

% Check PCS height to ensure inside fuselage
PCS_z = Main(25, 3);          % C25
fuse_z_center = Main(52, 4);  % D52
fuse_z_height = Main(52, 6);  % F52

if PCS_z < (fuse_z_center - fuse_z_height/2) || PCS_z > (fuse_z_center + fuse_z_height/2)
    logText = logf(logText, 'PCS Z-location outside fuselage vertical bounds.\n');
    disconnected = disconnected + 1;
end

% Check VT spacing to ensure inside fuselage
VT_y = Main(24, 8);           % H24
fuse_width = Main(52, 5);     % E52

if VT_y > fuse_width/2
    logText = logf(logText, 'VT Y-location outside fuselage width.\n');
    disconnected = disconnected + 1;
end

% Strakes
if Main(18, 4) > 1  % D18 > 1 indicating area in the strakes
    sweep = Geom(15, 11);  % K15
    y = Geom(152, 13);     % M152
    strake = Geom(155, 12); % L155
    apex = Geom(38, 12);    % L38
    wing = (y / tand(90 - sweep) + apex);
    if wing < (strake + 0.5)
        % Connected
    else
        logText = logf(logText, 'Strake disconnected\n');
        disconnected = disconnected + 1;
    end
end

component_positions = Main(23, 2:8);  % B23:H23

% Check if any component is behind or at the end of the fuselage
if any(component_positions >= fuselage_length)
    logText = logf(logText, 'One or more components X Location are not ahead of the fuselage end (B32 = %.2f)\n', fuselage_end);
    pt = pt - 1;
else
    %     logText = logf(logText, 'All components are correctly positioned ahead of the fuselage end (B32 = %.2f)\n', fuselage_end);
end


if disconnected > 0
    logText = logf(logText, '-1 Point Some control surface is disconnected! Move them closer to the fuselage so they can be attached and not parts existing in a nearby location \n');
    pt=pt-1;
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% --- Stability Checks (1 point) %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Ensure adequate static margin and acceptable lat/dir stability
% derivatives

SM = Main(10, 13);   % M10
clb = Main(10, 15);  % O10
cnb = Main(10, 16);  % P10
rat = Main(10, 17);  % Q10

StableAircraft = 1;
if ~(SM >= -0.1 && SM <= 0.11)
    logText = logf(logText, 'Static Margin out of bounds\n');
    StableAircraft = 0;
elseif SM < 0
    logText = logf(logText, 'Warning: Unstable aircraft - Recommend to increase SM above 0 or your glider will not fly \n');
end
if clb >= -0.001
    logText = logf(logText, 'Clb out of bounds\n');
    StableAircraft = 0;
end
if cnb <= 0.002
    logText = logf(logText, 'Cnb out of bounds\n');
    StableAircraft = 0;
end
if ~(rat >= -1 && rat <= -0.3)
    logText = logf(logText, 'Cnb/Clb ratio out of bounds\n');
    StableAircraft = 0;
end

if StableAircraft == 0
    logText = logf(logText, '-1 Point Unstable! Adjust your aircraft to achieve flyable stability parameters (M10-Q10)\n');
    pt=pt-1;
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Fuel and Volume Check (2 points)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Ensure more fuel available than required

fuel_available = Main(18, 15);  % O18
fuel_required = Main(40, 24);   % X40

if fuel_available < fuel_required
    logText = logf(logText, '-1 Point Insufficient fuel: Available = %.1f, Required = %.1f\n', fuel_available, fuel_required);
    pt = pt - 1;
end

volume_remaining = Main(23, 17); % Q23
if volume_remaining > 0
    %     logText = logf(logText, 'Volume check passed: %.2f remaining\n', volume_remaining);
else
    logText = logf(logText, '-1 Point Insufficient volume remaining: %.2f additional required\n', volume_remaining);
    pt = pt - 1;
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Recurring Cost (Q31) (1 pt)%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

cost = Main(31, 17); % Q31
numaircraft = Main(31,14); % N31
costFailed = 0;

if numaircraft == 187
    if cost > 90
        logText = logf(logText, '-1 Point Recurring cost exceeds threshold ($90M): $%.1fM\n', cost);
        costFailed=costFailed+1;
    elseif cost <= 80
        logText = logf(logText, 'Recurring cost meets objective (≤$80M): $%.1fM\n', cost);
    else
        %             logText = logf(logText, 'Recurring cost meets threshold: $%.1fM\n', cost);
    end
elseif numaircraft == 800
    if cost > 60
        logText = logf(logText, '-1 Point Recurring cost exceeds threshold ($60M): $%.1fM\n', cost);
        costFailed=costFailed+1;
    elseif cost <= 50
        logText = logf(logText, 'Recurring cost meets objective (≤$50M): $%.1fM\n', cost);
    else
        %             logText = logf(logText, 'Recurring cost meets threshold: $%.1fM\n', cost);
    end
else
    logText = logf(logText, '-1 Point $d is not a valid number of aircraft for cost estimation\n', numaircraft);
    costFailed=costFailed+1;
end

if costFailed>0
    pt = pt - 1;
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Landing Gear Checks (1 pt)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

LandingGearGood = 1;

if Gear(19, 10) < 10 || Gear(19, 10) > 20 % J19 between 10 and 20%
    logText = logf(logText, 'Violates Nose gear 90/10 rule \n');
    LandingGearGood = 0;
end

if Gear(19, 12) < Gear(20, 12) % L19 is less than L20 so it won't tail sit
    % Tipback okay
else
    logText = logf(logText, 'Violates Tipback angle requirement \n');
    LandingGearGood = 0;
end

if Gear(19, 13) < Gear(20, 13) %M19 is less than M20 so it won't roll over
    % Rollover okay
else
    logText = logf(logText, 'Violates Rollover angle requirement \n');
    LandingGearGood = 0;
end

if Gear(19, 14) < 200 %M19 is less than 200 kts TO speed
    % Rotation speed okay
else
    logText = logf(logText, 'Violates Takeoff Rotation Speed \n');
    LandingGearGood = 0;
end

if LandingGearGood ~= 1
    pt = pt - 1;
    logText = logf(logText, '-1 point Something is wrong with the landing gear/takeoff speed, see the hints in the "Gear" tab! \n');
end
%     fprintf('Geometry Check Complete\n')

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Output total points(cadetnum) and store log text
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
logText = logf(logText,'Jet10 Score is %d out of 10\n',pt);
logText = logf(logText,'Cutout is 5 out of 5\n\n');
[~, name, ext] = fileparts(filename);
fprintf('%s completed\n', [name, ext]);
fprintf('Jet10 Score is %d out of 10\n',pt)
fprintf('Cutout is 5 out of 5\n\n')
logText = strjoin(logText, ''); % Join logText into a single string, using newline as delimiter.
fb = char(logText); % Accumulate log text
end

%% Function to read all useful sheets, and verify numbers returned for used cells
function sheets = loadAllJet10Sheets(filename)
%LOADALLJET10SHEETS Load all required Jet10 sheets using safeReadMatrix
%   sheets = loadAllJet10Sheets(filename) returns a struct with fields:
%   Aero, Miss, Main, Consts, Gear, Geom

sheets.Aero   = safeReadMatrix(filename, 'Aero',   {'G3','G4','G10','G11','A15','A16'});
sheets.Miss   = safeReadMatrix(filename, 'Miss',   {'C48','C49'});
sheets.Main   = safeReadMatrix(filename, 'Main',   {'T3','U3','V3','W3','X3','Y3','T4','U4','V4','W4','X4','Y4','T6','U6',...
    'V6','W6','X6','Y6','T7','U7','V7','W7','X7','Y7','T8','U8','V8','W8',...
    'X8','Y8','T9','U9','V9','W9','X9','Y9','AB3','AB4','X12','X13','M10',...
    'O10','P10','Q10','O18','X40','Q23','Q31','N31','B32','C23','H23',...
    'D18','D23','D52','F52','H24','E52'});
sheets.Consts = safeReadMatrix(filename, 'Consts', {'K22','K23','K24','K26','K27','K28','K29','K32','AO42','AQ41','K33'});
sheets.Gear   = safeReadMatrix(filename, 'Gear',   {'J19','L19','L20','M19','M20','N19','N20'});
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
        [file, path] = uigetfile('*.xls*','Select a Jet10 Excel file');
        if isequal(file,0)
            mode = 'cancelled';
        else
            mode = 'single';
            selectedPath = fullfile(path, file);
        end
        delete(d);
    end

    function folderRun(~,~)
        path = uigetdir(pwd, 'Select folder containing Jet10 files');
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
            csvFilename = fullfile(folderAnalyzed, ['GE3_Blackboard_Offline_', timestamp, '.csv']);
            fid = fopen(csvFilename, 'w');

            % Assignment title column (update if needed)
            assignmentTitle = 'GE 3: AATF Design Iteration 1 & Cutout [Total Pts: 15 Score]';

            % Write header
            fprintf(fid, '"Last Name","First Name","Username","%s","Grading Notes","Notes Format","Feedback to Learner","Feedback Format"\n', assignmentTitle);

            for i = 1:numel(files)
                fname = files(i).name;

                % Extract username from filename (e.g., c27charlesh.allen)
                tokens = regexp(fname, '_(c\d{2}[a-z]+[a-z]\.[a-z]+)_attempt', 'tokens');
                if ~isempty(tokens)
                    username = tokens{1}{1};

                    % Extract last and first name from username
                    nameParts = regexp(username, 'c\d{2}([a-z]+)[a-z]\.([a-z]+)', 'tokens');
                    if ~isempty(nameParts)
                        rawLast = nameParts{1}{1};
                        rawFirst = nameParts{1}{2};

                        % Capitalize first letter manually
                        lastName = [upper(rawLast(1)), lower(rawLast(2:end))];
                        firstName = [upper(rawFirst(1)), lower(rawFirst(2:end))];
                    else
                        firstName = '';
                        lastName = '';
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
                fbText = strrep(fbText, '≥', '&ge;');
                fbText = strrep(fbText, '≤', '&le;');
                fbText = strrep(fbText, '≠', '&ne;');
                fbText = strrep(fbText, '✔', '&#10004;');
                fbText = strrep(fbText, '✘', '&#10008;');
                fbText = strrep(fbText, '✅', '&#9989;');
                fbText = strrep(fbText, '❌', '&#10060;');
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
