%% =========================================================================
%  PAIRED T-TEST ANALYSIS — Pre vs Post (STIP / Balance / JPE)
%
%  What this script does:
%    - Loads your Excel summary file
%    - Runs paired t-test AND Wilcoxon signed-rank for each variable
%    - Checks normality with Shapiro-Wilk
%    - Computes Cohen's d effect size
%    - Prints a clean results table to the Command Window
%    - Saves results to Excel
%
%  HOW TO CHANGE VARIABLES — read the USER SETTINGS section below
%
%  Author: STIP Batch Analysis
%  Date:   2026
% =========================================================================

clc; clear; close all;

%% =========================================================================
%  >>>  USER SETTINGS — ONLY EDIT THIS SECTION  <<<
% =========================================================================

% ── 1. INPUT FILE ─────────────────────────────────────────────────────────
%  Path to your Excel summary file (the one with Pre and Post columns)
input_file = 'JPE_FJPE_Organized18.xlsx';

%  Sheet name inside the Excel file
sheet_name = 'JPE_FJPE_Organized';

% ── 2. SUBJECT FILTER ─────────────────────────────────────────────────────
%  Leave empty {} to include ALL subjects
%  Or list specific subjects to include, e.g. {'S01','S06','S18'}
include_subjects = {};     % ALL subjects
% include_subjects = {'S01','S06','S18','S19','S22','S24','S25'};   % N=7
%include_subjects = {'S01','S03','S04','S06','S07','S10','S13', ...
                    %'S14','S16','S17','S18','S19','S20','S21', ...
                   %'S22','S23','S24','S25'};                      % N=18

% ── 3. COLUMN NAMES ───────────────────────────────────────────────────────
%  Column in your Excel that contains subject IDs
subject_col = 'Subject';

%  Column name for PRE values
pre_col  = 'JPE';

%  Column name for POST values
post_col = 'FJPE';

%  !! IF YOUR EXCEL HAS DIFFERENT COLUMN NAMES change pre_col and post_col
%  For example for the JPE vs FJPE file:
%    pre_col  = 'JPE';
%    post_col = 'FJPE';

% ── 4. VARIABLE COLUMN NAMES TO TEST ──────────────────────────────────────
%  This is where you choose WHICH variables to run statistics on.
%  List the Variable names exactly as they appear in your Excel.
%
%  FOR STIP DATA — use these:
%variables = { ...
    'n_DS_events', ...
    'n_steps_total', ...
    'n_steps_valid', ...
    'Duration_s', ...
    'Final_MovedDist_mm', ...
    'Final_Sagittal_mm', ...
    'Final_Lateral_mm', ...
    'Final_Rotation_deg', ...
    'StepTime_mean_s', ...
    'StepTime_SD_s', ...
    'StepLen_mean_mm', ...
    'StepLen_SD_mm', ...
    'Cadence_mean_spm', ...
    'Cadence_SD_spm' ...};

%  FOR BALANCE DATA — comment out the block above and use:
 %variables = {'X SD', 'X D Avg', 'Y SD', 'Y D Avg', 'V Avg','Area95'};
%  Note: Area is handled separately because each subject has a unique column

%  FOR JPE vs FJPE DATA — comment out the block above and use:
variables = {'JPE_3D_E', 'JPE_3D_F', 'JPE_3D_L','JPE_3D_R','LatBend_E','LatBend_F','LatBend_L','LatBend_R','LatBend_R','Axial_E','Axial_F','Axial_L','Axial_R',...
    'FlexExt_E','FlexExt_F','FlexExt_L','FlexExt_R'};

% ── 5. SIGNIFICANCE LEVEL ─────────────────────────────────────────────────
alpha = 0.05;              % uncorrected significance threshold
bonferroni = true;         % apply Bonferroni correction? true / false

% ── 6. OUTPUT FILE ────────────────────────────────────────────────────────
output_file = 'PairedTTest_ResultsJPE.xlsx';

% =========================================================================
%  END OF USER SETTINGS — do not edit below this line
% =========================================================================

%% ── LOAD DATA ─────────────────────────────────────────────────────────────
fprintf('\nLoading: %s (sheet: %s)\n', input_file, sheet_name);
T = readtable(input_file, 'Sheet', sheet_name, ...
    'VariableNamingRule', 'preserve');
fprintf('  Loaded %d rows, %d columns\n', height(T), width(T));

%% ── FILTER SUBJECTS ───────────────────────────────────────────────────────
if ~isempty(include_subjects)
    keep = ismember(T.(subject_col), include_subjects);
    T    = T(keep, :);
    fprintf('  Filtered to %d subjects\n', height(T));
else
    fprintf('  Using all %d subjects\n', height(T));
end

n_subjects = height(T);
n_vars     = length(variables);
bonf_alpha = alpha / n_vars;   % Bonferroni corrected threshold

%% ── RUN STATISTICS ────────────────────────────────────────────────────────
fprintf('\n%s\n', repmat('=', 1, 80));
fprintf('  PAIRED T-TEST RESULTS  (N=%d)\n', n_subjects);
if bonferroni
    fprintf('  Bonferroni correction: p < %.4f (alpha/%.0f)\n', bonf_alpha, n_vars);
end
fprintf('%s\n\n', repmat('=', 1, 80));
fprintf('  %-28s %7s %7s %7s %7s %7s %7s %8s %6s\n', ...
    'Variable', 'PreMean', 'PostMean', 'Diff', 'SW_p', 't_p', 'W_p', 'Cohens_d', 'Sig');
fprintf('  %s\n', repmat('-', 1, 76));

% Pre-allocate results storage
results = table();
results.Variable   = cell(n_vars, 1);
results.N          = zeros(n_vars, 1);
results.Pre_Mean   = zeros(n_vars, 1);
results.Pre_SD     = zeros(n_vars, 1);
results.Post_Mean  = zeros(n_vars, 1);
results.Post_SD    = zeros(n_vars, 1);
results.Diff_Mean  = zeros(n_vars, 1);
results.Diff_SD    = zeros(n_vars, 1);
results.SW_stat    = zeros(n_vars, 1);
results.SW_p       = zeros(n_vars, 1);
results.Normal     = cell(n_vars, 1);
results.t_stat     = zeros(n_vars, 1);
results.t_p        = zeros(n_vars, 1);
results.W_stat     = zeros(n_vars, 1);
results.W_p        = zeros(n_vars, 1);
results.Cohens_d   = zeros(n_vars, 1);
results.Effect     = cell(n_vars, 1);
results.Sig_uncorr = cell(n_vars, 1);
results.Sig_Bonf   = cell(n_vars, 1);

for vi = 1:n_vars
    var_name = variables{vi};

    % ── Get Pre and Post values for this variable ─────────────────────────
    %  Strategy: find rows where Variable column matches, then read Pre/Post
    %  Works for long-format data (one row per variable per subject)
    %  AND wide-format data (one column per variable per subject)

    if ismember('Variable', T.Properties.VariableNames)
        % LONG FORMAT: rows filtered by Variable column
        rows = strcmp(T.Variable, var_name);
        pre_vals  = T.(pre_col)(rows);
        post_vals = T.(post_col)(rows);
    else
        % WIDE FORMAT: columns named by variable
        pre_col_name  = [var_name '_Pre'];
        post_col_name = [var_name '_Post'];
        if ismember(pre_col_name, T.Properties.VariableNames)
            pre_vals  = T.(pre_col_name);
            post_vals = T.(post_col_name);
        elseif ismember(var_name, T.Properties.VariableNames)
            % fallback: variable is a single column (shouldn't happen for paired)
            fprintf('  %-28s  !! Cannot find paired Pre/Post columns\n', var_name);
            continue;
        else
            fprintf('  %-28s  !! Variable not found in table\n', var_name);
            continue;
        end
    end

    pre_vals  = double(pre_vals);
    post_vals = double(post_vals);
    diff_vals = post_vals - pre_vals;

    % Remove NaN pairs
    valid = ~isnan(pre_vals) & ~isnan(post_vals);
    pre_v  = pre_vals(valid);
    post_v = post_vals(valid);
    diff_v = diff_vals(valid);
    n      = length(diff_v);

    if n < 3
        fprintf('  %-28s  !! Too few valid pairs (N=%d)\n', var_name, n);
        continue;
    end

    % ── Descriptives ──────────────────────────────────────────────────────
    pre_mean  = mean(pre_v);    pre_sd  = std(pre_v);
    post_mean = mean(post_v);   post_sd = std(post_v);
    diff_mean = mean(diff_v);   diff_sd = std(diff_v);

    % ── Shapiro-Wilk normality test on difference scores ──────────────────
    [~, sw_p] = lillietest(diff_v);  sw_stat = NaN;  sw_h = sw_p < 0.05;   % requires swtest.m
    % FALLBACK if swtest not installed:
    % [~, sw_p] = lillietest(diff_v);  sw_stat = NaN;  sw_h = sw_p < 0.05;
    is_normal = ~sw_h;

    % ── Paired t-test ─────────────────────────────────────────────────────
    [~, t_p, ~, t_stats] = ttest(pre_v, post_v);
    t_stat_val = t_stats.tstat;

    % ── Wilcoxon signed-rank test ─────────────────────────────────────────
    try
        [w_p, ~, w_stats] = signrank(pre_v, post_v);
        w_stat_val = w_stats.signedrank;
    catch
        w_p        = NaN;
        w_stat_val = NaN;
    end

    % ── Cohen's d (on difference scores) ─────────────────────────────────
    if diff_sd > 0
        cohens_d = diff_mean / diff_sd;
    else
        cohens_d = NaN;
    end
    if      abs(cohens_d) >= 0.8,  effect = 'Large';
    elseif  abs(cohens_d) >= 0.5,  effect = 'Medium';
    else,                           effect = 'Small';
    end

    % ── Significance flags ─────────────────────────────────────────────────
    sig_uncorr = (t_p < alpha)      | (w_p < alpha);
    sig_bonf   = (t_p < bonf_alpha) | (w_p < bonf_alpha);
    sig_str    = '';
    if sig_uncorr, sig_str = '*';  end
    if sig_bonf,   sig_str = '**'; end

    % ── Print to Command Window ────────────────────────────────────────────
    fprintf('  %-28s %7.3f %7.3f %+7.3f %7.4f %7.4f %7.4f %8.3f %6s\n', ...
        var_name, pre_mean, post_mean, diff_mean, sw_p, t_p, w_p, cohens_d, sig_str);

    % ── Store results ──────────────────────────────────────────────────────
    results.Variable{vi}   = var_name;
    results.N(vi)          = n;
    results.Pre_Mean(vi)   = pre_mean;
    results.Pre_SD(vi)     = pre_sd;
    results.Post_Mean(vi)  = post_mean;
    results.Post_SD(vi)    = post_sd;
    results.Diff_Mean(vi)  = diff_mean;
    results.Diff_SD(vi)    = diff_sd;
    results.SW_stat(vi)    = sw_stat;
    results.SW_p(vi)       = sw_p;
    results.Normal{vi}     = matlab.lang.makeValidName(string(is_normal));
    results.t_stat(vi)     = t_stat_val;
    results.t_p(vi)        = t_p;
    results.W_stat(vi)     = w_stat_val;
    results.W_p(vi)        = w_p;
    results.Cohens_d(vi)   = cohens_d;
    results.Effect{vi}     = effect;
    results.Sig_uncorr{vi} = sig_str;
    results.Sig_Bonf{vi}   = ternary(sig_bonf, '**', '');

end   % end variable loop

%% ── PRINT LEGEND ──────────────────────────────────────────────────────────
fprintf('\n  %s\n', repmat('-', 1, 76));
fprintf('  *  = p < %.2f (uncorrected)\n', alpha);
fprintf('  ** = p < %.4f (Bonferroni corrected for %d tests)\n', bonf_alpha, n_vars);
fprintf('  SW_p = Shapiro-Wilk on difference scores (> %.2f = normal)\n', alpha);
fprintf('  Diff = Post - Pre  (positive = increase after fatigue)\n');
fprintf('%s\n\n', repmat('=', 1, 80));

%% ── SAVE TO EXCEL ─────────────────────────────────────────────────────────
% Remove empty rows (variables not found)
keep_rows = ~cellfun(@isempty, results.Variable);
results   = results(keep_rows, :);

writetable(results, output_file, 'Sheet', 'Paired_TTest');
fprintf('  Results saved to: %s\n\n', output_file);

%% =========================================================================
%  LOCAL HELPER
% =========================================================================
function out = ternary(cond, a, b)
    if cond, out = a; else, out = b; end
end