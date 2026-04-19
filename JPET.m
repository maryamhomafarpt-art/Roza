%% =========================================================================
%  CERVICAL JPE — BATCH ANALYSIS: ALL SUBJECTS, TRIALS 2–4
%
%  For each subject (S01 to S25):
%    - Loads trials 2, 3, 4  (trial 1 skipped = familiarisation)
%    - Computes JPE 3D, Axial, FlexExt, LatBend for events 2–5 vs event 1
%    - Appends results to one master Excel file
%
%  Excel output format (one row per subject × event):
%    Subject | Event      | Trial_2 | Trial_3 | Trial_4
%    S01     | Evt1→Evt2  |  2.10   |  1.95   |  2.30
%    S01     | Evt1→Evt3  |  6.32   |  5.89   |  6.78
%    S01     | Evt1→Evt4  |  3.41   |  3.25   |  3.60
%    S01     | Evt1→Evt5  |  3.53   |  3.40   |  3.70
%    S02     | Evt1→Evt2  |  2.05   |  ...    |  ...
%    ...
%
%  Author: Generated for multi-subject JPE batch analysis
%  Date:   2026
% =========================================================================

clc; clear; close all;

%% ─── 1. USER SETTINGS  (only edit this section) ──────────────────────────

n_subjects     = 25;                 % total number of subjects
subject_prefix = 'S';                % prefix: 'S' → S01, S02, ... S25
trial_start    = 2;                  % first trial to analyse (skip trial 1)
trial_end      = 4;                  % last trial to analyse
file_ext       = 'tsv';              % 'tsv' or 'xlsx'
half_win       = 5;                  % ±frames averaged per event
THRESHOLD      = 4.5;                % clinical impairment threshold (°)

% Root folder that contains one subfolder per subject (S01, S02, ...)
% If all files are in ONE flat folder, set use_subfolders = false
data_folder    = '/Users/rozahomafar/Desktop/data roza spring/all';
use_subfolders = false;              % true  → data_folder/S01/S01-JP0002.tsv
                                     % false → data_folder/S01-JP0002.tsv

% Output Excel file (created fresh each run — rename to keep old versions)
out_file       = 'JPE_AllSubjects_Results.xlsx';

%% ─── 2. DERIVED SETTINGS  (do not edit) ──────────────────────────────────

trial_list   = trial_start : trial_end;   % [2 3 4]
n_trials     = length(trial_list);         % 3
event_labels = {'Evt1->Evt2','Evt1->Evt3','Evt1->Evt4','Evt1->Evt5'};
n_events     = 4;

% Delete old output file so we start fresh
if exist(out_file, 'file'), delete(out_file); end

%% ─── 3. MASTER RESULTS TABLE  (pre-allocate) ─────────────────────────────
%  Each row = one subject × one event
%  Columns  = Subject, Event, T2_JPE3D, T3_JPE3D, T4_JPE3D,
%              T2_Axial, T3_Axial, T4_Axial,
%              T2_FlexExt, T3_FlexExt, T4_FlexExt,
%              T2_LatBend, T3_LatBend, T4_LatBend

all_rows = {};   % will grow as subjects are processed

%% ─── 4. MAIN SUBJECT LOOP ────────────────────────────────────────────────

fprintf('\n%s\n', repmat('=',1,72));
fprintf('  BATCH JPE ANALYSIS — %d SUBJECTS  (trials %d to %d)\n', ...
    n_subjects, trial_start, trial_end);
fprintf('%s\n\n', repmat('=',1,72));

for si = 1:n_subjects

    subject_id = sprintf('%s%02d', subject_prefix, si);   % e.g. 'S01'
    fprintf('Processing %s ...\n', subject_id);

    % ── 4a. Locate trial files for this subject ──────────────────────────
    trial_files = cell(n_trials, 1);
    found       = false(n_trials, 1);

    for ki = 1:n_trials
        k = trial_list(ki);   % actual trial number (2, 3, 4)

        if use_subfolders
            fname = fullfile(data_folder, subject_id, ...
                sprintf('%s_JP%04d.%s', subject_id, k, file_ext));
        else
            fname = fullfile(data_folder, ...
                sprintf('%s_JP%04d.%s', subject_id, k, file_ext));
        end

        if exist(fname, 'file')
            trial_files{ki} = fname;
            found(ki)       = true;
            fprintf('  Trial %d: found\n', k);
        else
            fprintf('  Trial %d: NOT FOUND (%s)\n', k, fname);
        end
    end

    n_found = sum(found);
    if n_found == 0
        fprintf('  !! Skipping %s — no trial files found.\n\n', subject_id);
        continue;
    end

    % ── 4b. Compute JPE for each trial ───────────────────────────────────
    % Storage: n_trials rows × 4 event-comparison cols
    JPE_3D  = nan(n_trials, n_events);
    Axial   = nan(n_trials, n_events);
    FlexExt = nan(n_trials, n_events);
    LatBend = nan(n_trials, n_events);

    for ki = 1:n_trials
        if ~found(ki), continue; end

        k = trial_list(ki);

        % Load file
        [event_frames, frames, FH, TH, LH] = ...
            load_jpe_file(trial_files{ki}, file_ext);

        % Build neutral reference frame from Event 1
        FH1 = mean_markers(FH, event_frames(1), frames, half_win);
        TH1 = mean_markers(TH, event_frames(1), frames, half_win);
        LH1 = mean_markers(LH, event_frames(1), frames, half_win);
        R_neutral = build_head_frame(TH1, FH1, LH1);

        % Compute JPE for events 2 → 5 vs event 1
        for ev = 2:5
            col = ev - 1;
            FH_t = mean_markers(FH, event_frames(ev), frames, half_win);
            TH_t = mean_markers(TH, event_frames(ev), frames, half_win);
            LH_t = mean_markers(LH, event_frames(ev), frames, half_win);
            R_tgt = build_head_frame(TH_t, FH_t, LH_t);

            [j3d, ax, fe, lb] = compute_jpe(R_neutral, R_tgt);

            JPE_3D(ki, col)  = j3d;
            Axial(ki, col)   = ax;
            FlexExt(ki, col) = fe;
            LatBend(ki, col) = lb;
        end

        fprintf('  Trial %d computed OK\n', k);
    end

    % ── 4c. Add this subject to master table ─────────────────────────────
    % One row per event comparison (4 rows per subject)
    for ev_col = 1:n_events
        % Build one row: Subject | Event | T2 | T3 | T4 (for each metric)
        row = {subject_id, event_labels{ev_col}};

        % JPE 3D values for each trial
        for ki = 1:n_trials
            row{end+1} = JPE_3D(ki, ev_col);   %#ok
        end
        % Axial values
        for ki = 1:n_trials
            row{end+1} = Axial(ki, ev_col);     %#ok
        end
        % FlexExt values
        for ki = 1:n_trials
            row{end+1} = FlexExt(ki, ev_col);   %#ok
        end
        % LatBend values
        for ki = 1:n_trials
            row{end+1} = LatBend(ki, ev_col);   %#ok
        end

        all_rows(end+1, :) = row;               %#ok
    end

    fprintf('  %s done (%d/%d trials)\n\n', subject_id, n_found, n_trials);
end

%% ─── 5. BUILD COLUMN NAMES ───────────────────────────────────────────────

% Trial labels  e.g. 'Trial_2', 'Trial_3', 'Trial_4'
trial_col_names = arrayfun(@(k) sprintf('Trial_%d', k), trial_list, ...
    'UniformOutput', false);

% Build full variable name list
var_names = [{'Subject', 'Event'}, ...
    strcat('JPE3D_',    trial_col_names), ...
    strcat('Axial_',    trial_col_names), ...
    strcat('FlexExt_',  trial_col_names), ...
    strcat('LatBend_',  trial_col_names)];

%% ─── 6. WRITE EXCEL FILE ─────────────────────────────────────────────────

if isempty(all_rows)
    error('No data collected — check data_folder and file naming.');
end

T_all = cell2table(all_rows, 'VariableNames', var_names);
writetable(T_all, out_file, 'Sheet', 'JPE_Results');

fprintf('%s\n', repmat('=',1,72));
fprintf('  DONE\n');
fprintf('  %d subjects processed\n', size(all_rows,1) / n_events);
fprintf('  Results saved to: %s\n', out_file);
fprintf('%s\n', repmat('=',1,72));

%% =========================================================================
%  LOCAL FUNCTIONS
% =========================================================================

function [event_frames, frames, FH, TH, LH] = load_jpe_file(filepath, ext)
%LOAD_JPE_FILE  Load a Qualisys JPE export (tsv or xlsx).
%   Returns event_frames (1×5), frames (N×1), FH/TH/LH (N×3, mm).

    if strcmp(ext, 'tsv')
        raw = readcell(filepath, 'FileType','text', ...
                       'Delimiter','\t', 'NumHeaderLines', 0);
    else
        try
            [~, fname, ~] = fileparts(filepath);
            raw = readcell(filepath, 'Sheet', fname, 'NumHeaderLines', 0);
        catch
            raw = readcell(filepath, 'Sheet', 1, 'NumHeaderLines', 0);
        end
    end

    % Find the 5 EVENT rows
    event_frames = zeros(1, 5);
    event_count  = 0;
    for r = 1:min(20, size(raw,1))
        val = raw{r,1};
        if ischar(val) && strcmpi(strtrim(val), 'EVENT')
            event_count = event_count + 1;
            col3 = raw{r,3};
            if isnumeric(col3)
                event_frames(event_count) = round(col3);
            else
                event_frames(event_count) = round(str2double(string(col3)));
            end
            if event_count == 5, break; end
        end
    end

    % Find the header row (contains 'Frame')
    header_row = 0;
    for r = 1:25
        if ischar(raw{r,1}) && strcmpi(strtrim(raw{r,1}), 'Frame')
            header_row = r;
            break;
        end
    end
    if header_row == 0
        error('Cannot find Frame header in %s', filepath);
    end

    % Map column names
    headers = raw(header_row, :);
    get_col = @(name) find(cellfun(@(h) ischar(h) && ...
        strcmpi(strtrim(h), name), headers), 1);

    c_frame = get_col('Frame');
    c_FH_X  = get_col('FH X');  c_FH_Y = get_col('FH Y');  c_FH_Z = get_col('FH Z');
    c_TH_X  = get_col('TH X');  c_TH_Y = get_col('TH Y');  c_TH_Z = get_col('TH Z');
    c_LH_X  = get_col('LH X');  c_LH_Y = get_col('LH Y');  c_LH_Z = get_col('LH Z');

    % Read numeric data rows
    data_rows = raw(header_row+1:end, :);
    to_num = @(col) cellfun(@(x) ...
        double(x)  * double(isnumeric(x)) + ...
        str2double(string(x)) * double(~isnumeric(x)), ...
        data_rows(:, col));

    frames   = to_num(c_frame);
    valid    = ~isnan(frames);
    frames   = frames(valid);

    mk = @(cx,cy,cz) [to_num(cx), to_num(cy), to_num(cz)];
    FH_all = mk(c_FH_X, c_FH_Y, c_FH_Z);  FH = FH_all(valid,:);
    TH_all = mk(c_TH_X, c_TH_Y, c_TH_Z);  TH = TH_all(valid,:);
    LH_all = mk(c_LH_X, c_LH_Y, c_LH_Z);  LH = LH_all(valid,:);
end

% ─────────────────────────────────────────────────────────────────────────

function R = build_head_frame(TH_pos, FH_pos, LH_pos)
%BUILD_HEAD_FRAME  Orthonormal head frame from TH, FH, LH markers.
%   z = superior (centroid→TH), y = anterior (LH→FH), x = right (cross)
    centroid = (FH_pos + LH_pos) / 2.0;
    z = TH_pos - centroid;   z = z / norm(z);
    y = FH_pos - LH_pos;     y = y / norm(y);
    x = cross(y, z);          x = x / norm(x);
    y = cross(z, x);          y = y / norm(y);
    R = [x(:), y(:), z(:)];
end

% ─────────────────────────────────────────────────────────────────────────

function [jpe3d, axial, flex_ext, lat_bend] = compute_jpe(R_ref, R_tgt)
%COMPUTE_JPE  3D JPE and ZXY Euler components (degrees).
%   axial    = Yaw  (Z) = axial rotation  left/right
%   flex_ext = Pitch(Y) = flexion/extension
%   lat_bend = Roll (X) = lateral bending
    R_rel    = R_ref' * R_tgt;
    trace_v  = max(-1, min(1, (trace(R_rel)-1)/2));
    jpe3d    = acosd(trace_v);
    flex_ext = asind(max(-1, min(1, R_rel(3,2))));
    if abs(cosd(flex_ext)) > 1e-6
        axial    = atan2d(-R_rel(1,2), R_rel(2,2));
        lat_bend = atan2d(-R_rel(3,1), R_rel(3,3));
    else
        axial    = atan2d( R_rel(1,3), R_rel(1,1));
        lat_bend = 0;
    end
end

% ─────────────────────────────────────────────────────────────────────────

function pos = mean_markers(arr, frame_num, frames, half_win)
%MEAN_MARKERS  Average marker positions over ±half_win frames.
    idx = find(frames >= frame_num - half_win & ...
               frames <= frame_num + half_win);
    if isempty(idx)
        [~, nearest] = min(abs(frames - frame_num));
        idx = nearest;
    end
    pos = mean(arr(idx,:), 1, 'omitnan');
end