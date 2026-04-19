%% =========================================================================
%  STEPPING-IN-PLACE (STIP) ANALYSIS
%  Based on: Malmström et al. (2017), Exp Brain Res 235:2755–2766
%
%  Computes:
%    1. Moved distance   — total XY displacement from start (Pythagorean)
%    2. Body orientation — yaw rotation from shoulder vector (Rsh→Lsh)
%    3. Gait steps       — step time, step length, cadence per step
%
%  Markers (columns in data table):
%    Rank X/Y/Z  — Right ankle
%    Lank X/Y/Z  — Left ankle
%    Sac  X/Y/Z  — Sacrum
%    Rsh  X/Y/Z  — Right shoulder
%    Lsh  X/Y/Z  — Left shoulder
%    C7   X/Y/Z  — Cervical 7
%    FH   X/Y/Z  — Front head
%    TH   X/Y/Z  — Top head
%    LH   X/Y/Z  — Left head
%
%  File format: Qualisys .xlsx export (header rows 1–10, data from row 11)
%  Column order: Frame | Time | Rank(X,Y,Z) | Lank(X,Y,Z) | Sac(X,Y,Z) |
%                Rsh(X,Y,Z) | Lsh(X,Y,Z) | C7(X,Y,Z) | FH(X,Y,Z) |
%                TH(X,Y,Z) | LH(X,Y,Z)
%
%  Author: Generated for S04-ST0001
%  Date:   2026
% =========================================================================

clc; clear; close all;

%% ─── 1. USER SETTINGS ────────────────────────────────────────────────────

filename     = 'S04-ST0001.xlsx';   % <-- update path if needed
sheet_name   = 'S04-ST0001';
data_row_start = 11;                % Row where column headers are
freq         = 100;                 % Sampling frequency (Hz)
ground_pct   = 5;                   % Percentile for ground-level estimation
ground_thresh_mm = 40;              % mm above ground = still "on ground"
min_ds_frames    = 2;               % Minimum frames for a double-support event
outlier_k        = 5;               % MAD multiplier for outlier step detection

%% ─── 2. LOAD DATA ────────────────────────────────────────────────────────

fprintf('Loading data from: %s ...\n', filename);

% Read header row to get column names
raw_headers = readcell(filename, 'Sheet', sheet_name, ...
    'Range', sprintf('A%d:AD%d', data_row_start, data_row_start));

% Read numeric data (rows after header)
opts = detectImportOptions(filename, 'Sheet', sheet_name);
opts.DataRange = sprintf('A%d', data_row_start + 1);
opts.VariableNamesRange = sprintf('A%d', data_row_start);
T = readtable(filename, opts);

col = T.Properties.VariableNames;

% Helper to find column index by name fragment
find_col = @(name) find(contains(col, name, 'IgnoreCase', true), 1);

% Extract arrays — fallback to positional indices if names not found
try
    frames = T{:, find_col('Frame')};
    t      = T{:, find_col('Time')};
    Rank   = T{:, [find_col('RankX'), find_col('RankY'), find_col('RankZ')]};
    Lank   = T{:, [find_col('LankX'), find_col('LankY'), find_col('LankZ')]};
    Sac    = T{:, [find_col('SacX'),  find_col('SacY'),  find_col('SacZ')]};
    Rsh    = T{:, [find_col('RshX'),  find_col('RshY'),  find_col('RshZ')]};
    Lsh    = T{:, [find_col('LshX'),  find_col('LshY'),  find_col('LshZ')]};
    C7     = T{:, [find_col('C7X'),   find_col('C7Y'),   find_col('C7Z')]};
    FH     = T{:, [find_col('FHX'),   find_col('FHY'),   find_col('FHZ')]};
    TH     = T{:, [find_col('THX'),   find_col('THY'),   find_col('THZ')]};
    LH     = T{:, [find_col('LHX'),   find_col('LHY'),   find_col('LHZ')]};
catch
    warning('Column name matching failed — using fixed positional indices.');
    frames = T{:,1};
    t      = T{:,2};
    Rank   = T{:,3:5};
    Lank   = T{:,6:8};
    Sac    = T{:,9:11};
    Rsh    = T{:,12:14};
    Lsh    = T{:,15:17};
    C7     = T{:,18:20};
    FH     = T{:,22:24};
    TH     = T{:,25:27};
    LH     = T{:,28:30};
end

N = length(t);
fprintf('  Loaded %d frames, %.2f – %.2f s\n', N, t(1), t(end));

%% ─── 3. DOUBLE-SUPPORT DETECTION ─────────────────────────────────────────
% A double-support (DS) event = both feet simultaneously on the ground.
% "On ground" = ankle Z < (5th-percentile Z) + ground_thresh_mm

rank_z = Rank(:,3);
lank_z = Lank(:,3);

rank_ground = prctile(rank_z, ground_pct);
lank_ground = prctile(lank_z, ground_pct);

rank_on = rank_z < (rank_ground + ground_thresh_mm);
lank_on = lank_z < (lank_ground + ground_thresh_mm);
both_on = rank_on & lank_on;

% Find contiguous segments
transitions = diff([0; double(both_on); 0]);
ds_starts = find(transitions ==  1);   % rising edge
ds_ends   = find(transitions == -1) - 1; % falling edge

% Keep only events >= min_ds_frames long
valid_ds = (ds_ends - ds_starts + 1) >= min_ds_frames;
ds_starts = ds_starts(valid_ds);
ds_ends   = ds_ends(valid_ds);
n_ds = length(ds_starts);

fprintf('  Detected %d double-support events\n', n_ds);

%% ─── 4. HELPER FUNCTIONS ─────────────────────────────────────────────────

% Mean position over a frame range (ignores NaN)
mean_pos = @(arr, s, e) mean(arr(s:e, :), 1, 'omitnan');

% Body yaw angle (degrees) from shoulder vector Rsh→Lsh in XY plane
body_yaw_deg = @(rsh, lsh) atan2d(lsh(2)-rsh(2), lsh(1)-rsh(1));

%% ─── 5. REFERENCE POSITION (first DS event) ──────────────────────────────

ref_s = ds_starts(1);
ref_e = ds_ends(1);

ref_Rank = mean_pos(Rank, ref_s, ref_e);
ref_Lank = mean_pos(Lank, ref_s, ref_e);
ref_ankle_mid = (ref_Rank + ref_Lank) / 2.0;   % XY reference point

ref_Rsh = mean_pos(Rsh, ref_s, ref_e);
ref_Lsh = mean_pos(Lsh, ref_s, ref_e);
ref_yaw = body_yaw_deg(ref_Rsh, ref_Lsh);

fprintf('\nReference position (first DS event):\n');
fprintf('  Time:        %.3f – %.3f s\n', t(ref_s), t(ref_e));
fprintf('  Ankle mid:   [%.1f, %.1f, %.1f] mm\n', ref_ankle_mid);
fprintf('  Ref yaw:     %.2f deg\n', ref_yaw);

%% ─── 6. MOVED DISTANCE AND BODY ORIENTATION PER DS EVENT ─────────────────

ds_event    = zeros(n_ds, 1);
ds_t_start  = zeros(n_ds, 1);
ds_t_end    = zeros(n_ds, 1);
ds_t_mid    = zeros(n_ds, 1);
moved_dist  = zeros(n_ds, 1);
sagittal_mm = zeros(n_ds, 1);
lateral_mm  = zeros(n_ds, 1);
rotation_deg= zeros(n_ds, 1);

for i = 1:n_ds
    s = ds_starts(i);
    e = ds_ends(i);

    ankle_mid = (mean_pos(Rank, s, e) + mean_pos(Lank, s, e)) / 2.0;
    rsh_pos   = mean_pos(Rsh, s, e);
    lsh_pos   = mean_pos(Lsh, s, e);

    % XY displacement from reference
    disp_xy = ankle_mid(1:2) - ref_ankle_mid(1:2);

    ds_event(i)     = i;
    ds_t_start(i)   = t(s);
    ds_t_end(i)     = t(e);
    ds_t_mid(i)     = (t(s) + t(e)) / 2.0;
    moved_dist(i)   = norm(disp_xy);                         % mm
    sagittal_mm(i)  = ankle_mid(2) - ref_ankle_mid(2);       % Y
    lateral_mm(i)   = ankle_mid(1) - ref_ankle_mid(1);       % X
    rotation_deg(i) = body_yaw_deg(rsh_pos, lsh_pos) - ref_yaw;  % deg
end

fprintf('\n=== MOVED DISTANCE AND BODY ORIENTATION ===\n');
fprintf('%5s %9s %10s %13s %13s %12s\n', ...
    'Event', 't_mid(s)', 'MovedDist', 'Sagittal(mm)', 'Lateral(mm)', 'Rotation(°)');
fprintf('%s\n', repmat('-',1,65));
for i = 1:n_ds
    fprintf('%5d %9.2f %10.1f %13.1f %13.1f %12.2f\n', ...
        ds_event(i), ds_t_mid(i), moved_dist(i), ...
        sagittal_mm(i), lateral_mm(i), rotation_deg(i));
end

%% ─── 7. GAIT STEP ANALYSIS ────────────────────────────────────────────────
% Each step = interval between two consecutive DS events.
% The stepping foot = the one with the higher peak Z between events.
% Step length = XY displacement of that ankle between the two DS events.

n_steps = n_ds - 1;

step_num      = zeros(n_steps, 1);
step_t_start  = zeros(n_steps, 1);
step_t_end    = zeros(n_steps, 1);
step_time_s   = zeros(n_steps, 1);
step_foot     = cell(n_steps, 1);
step_length   = zeros(n_steps, 1);
step_cadence  = zeros(n_steps, 1);

for i = 1:n_steps
    s1 = ds_starts(i);   e1 = ds_ends(i);
    s2 = ds_starts(i+1); e2 = ds_ends(i+1);

    t1 = ds_t_mid(i);
    t2 = ds_t_mid(i+1);
    dt = t2 - t1;

    % Determine which foot stepped
    seg_rank_z = rank_z(e1:s2);
    seg_lank_z = lank_z(e1:s2);
    max_rank = max(seg_rank_z, [], 'omitnan');
    max_lank = max(seg_lank_z, [], 'omitnan');

    rank_at_1 = mean_pos(Rank, s1, e1);
    lank_at_1 = mean_pos(Lank, s1, e1);
    rank_at_2 = mean_pos(Rank, s2, e2);
    lank_at_2 = mean_pos(Lank, s2, e2);

    if max_rank >= max_lank
        foot = 'Right';
        step_len = norm(rank_at_2(1:2) - rank_at_1(1:2));
    else
        foot = 'Left';
        step_len = norm(lank_at_2(1:2) - lank_at_1(1:2));
    end

    step_num(i)     = i;
    step_t_start(i) = t1;
    step_t_end(i)   = t2;
    step_time_s(i)  = dt;
    step_foot{i}    = foot;
    step_length(i)  = step_len;
    step_cadence(i) = 60.0 / dt;
end

% Outlier removal (step length > median + outlier_k × MAD)
sl_med = median(step_length);
sl_mad = median(abs(step_length - sl_med));
outlier_threshold = sl_med + outlier_k * sl_mad;
is_outlier = step_length > outlier_threshold;

valid_idx   = ~is_outlier;
sl_valid    = step_length(valid_idx);
st_valid    = step_time_s(valid_idx);
cad_valid   = step_cadence(valid_idx);
right_idx   = valid_idx & strcmp(step_foot, 'Right');
left_idx    = valid_idx & strcmp(step_foot, 'Left');

fprintf('\n=== GAIT STEPS ===\n');
fprintf('%5s %8s %8s %8s %6s %12s %13s %8s\n', ...
    'Step','t_start','t_end','Time(s)','Foot','Length(mm)','Cadence(spm)','Flag');
fprintf('%s\n', repmat('-',1,72));
for i = 1:n_steps
    flag = '';
    if is_outlier(i), flag = '<OUTLIER>'; end
    fprintf('%5d %8.2f %8.2f %8.3f %6s %12.1f %13.1f %8s\n', ...
        step_num(i), step_t_start(i), step_t_end(i), step_time_s(i), ...
        step_foot{i}, step_length(i), step_cadence(i), flag);
end

%% ─── 8. SUMMARY STATISTICS ────────────────────────────────────────────────

fprintf('\n%s\n', repmat('=',1,60));
fprintf('  SUMMARY\n');
fprintf('%s\n', repmat('=',1,60));
fprintf('  Trial duration:          %.2f s\n', t(end));
fprintf('  Double-support events:   %d\n', n_ds);
fprintf('\n  --- FINAL DISPLACEMENT (start → last DS event) ---\n');
fprintf('  Total moved distance:    %.1f mm\n', moved_dist(end));
fprintf('  Sagittal displacement:   %.1f mm\n', sagittal_mm(end));
fprintf('  Lateral displacement:    %.1f mm\n', lateral_mm(end));
fprintf('  Body rotation:           %.2f deg\n', rotation_deg(end));
fprintf('\n  --- GAIT (valid steps: %d / %d) ---\n', sum(valid_idx), n_steps);
fprintf('  Step time:    mean=%.3fs, SD=%.3fs\n', mean(st_valid), std(st_valid));
fprintf('  Step length:  mean=%.1fmm, SD=%.1fmm\n', mean(sl_valid), std(sl_valid));
fprintf('  Cadence:      mean=%.1f spm, SD=%.1f spm\n', mean(cad_valid), std(cad_valid));
fprintf('  Right step:   mean=%.1fmm\n', mean(step_length(right_idx)));
fprintf('  Left step:    mean=%.1fmm\n', mean(step_length(left_idx)));
fprintf('%s\n', repmat('=',1,60));

%% ─── 9. PLOTS ─────────────────────────────────────────────────────────────

figure('Name','STIP Analysis — S04-ST0001','NumberTitle','off', ...
    'Color','w','Position',[80 80 1200 800]);

% ── (a) Ankle Z trajectories with DS events ──
subplot(3,2,1);
plot(t, rank_z, 'Color',[0.18 0.46 0.71], 'LineWidth',0.8); hold on;
plot(t, lank_z, 'Color',[0.85 0.33 0.10], 'LineWidth',0.8);
for i = 1:n_ds
    x_fill = [t(ds_starts(i)) t(ds_ends(i)) t(ds_ends(i)) t(ds_starts(i))];
    y_fill = [0 0 500 500];
    fill(x_fill, y_fill, [0.6 0.9 0.6], 'FaceAlpha',0.25, 'EdgeColor','none');
end
xlabel('Time (s)'); ylabel('Z (mm)');
title('Ankle height — double-support (green shading)');
legend('Right ankle','Left ankle','Location','NorthEast','FontSize',8);
grid on; box off;

% ── (b) Moved distance over time ──
subplot(3,2,2);
plot(ds_t_mid, moved_dist, '-o', 'Color',[0.18 0.46 0.71], ...
    'LineWidth',1.5, 'MarkerSize',4, 'MarkerFaceColor',[0.18 0.46 0.71]);
xlabel('Time (s)'); ylabel('Moved distance (mm)');
title('Moved distance from start');
grid on; box off;

% ── (c) Sagittal vs lateral trajectory (top view) ──
subplot(3,2,3);
scatter(lateral_mm, sagittal_mm, 30, ds_t_mid, 'filled');
hold on;
plot(lateral_mm, sagittal_mm, '-', 'Color',[0.6 0.6 0.6], 'LineWidth',0.5);
plot(0, 0, 'k+', 'MarkerSize',12, 'LineWidth',2);
colorbar; xlabel('Lateral (mm)'); ylabel('Sagittal (mm)');
title('Body trajectory (top view)  — colour = time');
axis equal; grid on; box off;

% ── (d) Body rotation over time ──
subplot(3,2,4);
bar(ds_t_mid, rotation_deg, 'FaceColor',[0.49 0.18 0.56], 'EdgeColor','none');
yline(0,'k--','LineWidth',1);
xlabel('Time (s)'); ylabel('Rotation (°)');
title('Body rotation (+ right, − left)');
grid on; box off;

% ── (e) Step length over steps ──
subplot(3,2,5);
bar_colors = zeros(n_steps, 3);
for i = 1:n_steps
    if is_outlier(i)
        bar_colors(i,:) = [0.9 0.2 0.2];  % red = outlier
    elseif strcmp(step_foot{i},'Right')
        bar_colors(i,:) = [0.18 0.46 0.71];
    else
        bar_colors(i,:) = [0.85 0.33 0.10];
    end
end
for i = 1:n_steps
    bar(i, step_length(i), 'FaceColor', bar_colors(i,:), 'EdgeColor','none'); hold on;
end
yline(outlier_threshold,'r--','LineWidth',1.2,'DisplayName','Outlier threshold');
xlabel('Step number'); ylabel('Step length (mm)');
title('Step length (blue=Right, orange=Left, red=Outlier)');
grid on; box off;

% ── (f) Cadence over time ──
subplot(3,2,6);
plot(step_t_start(valid_idx), cad_valid, '-s', ...
    'Color',[0.2 0.6 0.2], 'LineWidth',1.5, ...
    'MarkerSize',4, 'MarkerFaceColor',[0.2 0.6 0.2]);
yline(mean(cad_valid),'--','Color',[0.2 0.6 0.2],'LineWidth',1);
xlabel('Time (s)'); ylabel('Cadence (steps/min)');
title('Cadence over time');
grid on; box off;

sgtitle('STIP Analysis — S04-ST0001', 'FontSize',13, 'FontWeight','bold');

%% ─── 10. EXPORT RESULTS TO EXCEL ──────────────────────────────────────────

out_file = 'STIP_Results_MATLAB.xlsx';

% Sheet 1: displacement per DS event
T_disp = table(ds_event, ds_t_start, ds_t_end, ds_t_mid, ...
    moved_dist, sagittal_mm, lateral_mm, rotation_deg, ...
    'VariableNames', {'Event','t_start_s','t_end_s','t_mid_s', ...
    'MovedDist_mm','Sagittal_mm','Lateral_mm','Rotation_deg'});
writetable(T_disp, out_file, 'Sheet', 'Displacement_per_DS_Event');

% Sheet 2: gait steps
outlier_flag = repmat({''},n_steps,1);
outlier_flag(is_outlier) = {'OUTLIER'};
T_gait = table(step_num, step_t_start, step_t_end, step_time_s, ...
    step_foot, step_length, step_cadence, outlier_flag, ...
    'VariableNames', {'Step','t_start_s','t_end_s','StepTime_s', ...
    'Foot','StepLength_mm','Cadence_spm','Flag'});
writetable(T_gait, out_file, 'Sheet', 'Gait_Steps');

% Sheet 3: summary statistics
param_names = {'Total moved distance (mm)';'Sagittal displacement (mm)';
    'Lateral displacement (mm)';'Body rotation (deg)';'Trial duration (s)';
    'Double-support events (n)';'Valid steps (n)';
    'Step time mean (s)';'Step time SD (s)';
    'Step length mean (mm)';'Step length SD (mm)';
    'Cadence mean (spm)';'Cadence SD (spm)';
    'Right step length mean (mm)';'Left step length mean (mm)'};
values = [moved_dist(end); sagittal_mm(end); lateral_mm(end);
    rotation_deg(end); t(end); n_ds; sum(valid_idx);
    mean(st_valid); std(st_valid);
    mean(sl_valid); std(sl_valid);
    mean(cad_valid); std(cad_valid);
    mean(step_length(right_idx)); mean(step_length(left_idx))];
T_sum = table(param_names, round(values,3), 'VariableNames',{'Parameter','Value'});
writetable(T_sum, out_file, 'Sheet', 'Summary');

fprintf('\nResults saved to: %s\n', out_file);
saveas(gcf, 'STIP_Analysis_Plot.png');
fprintf('Figure saved to: STIP_Analysis_Plot.png\n');

%% =========================================================================
%  LOCAL FUNCTIONS
% =========================================================================

function pos = get_mean_pos(arr, idx_start, idx_end)
% Average rows idx_start:idx_end of arr, ignoring NaN.
    pos = mean(arr(idx_start:idx_end, :), 1, 'omitnan');
end

function yaw = compute_yaw_deg(rsh, lsh)
% Yaw angle of body from Rsh->Lsh shoulder vector in XY plane (degrees).
    yaw = atan2d(lsh(2)-rsh(2), lsh(1)-rsh(1));
end
