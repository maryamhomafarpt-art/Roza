# Roza
MATLAB analysis of cervical proprioception, gait, and balance using motion capture data (JPE and STIP).
# Cervical Proprioception & STIP Analysis (MATLAB)

## Overview
This project analyzes cervical proprioception, gait, and balance using motion capture data.

It includes:
- STIP (Stepping in Place) analysis
- Cervical Joint Position Error (JPE) analysis
- Statistical analysis (paired t-test and non-parametric tests)

## Files Description

### 1. STIP Analysis
- File: STIP_Analysis_withplot.m
- Computes:
  - Moved distance
  - Body rotation
  - Step length and cadence
- Generates plots and exports results to Excel

### 2. JPE test Analysis
- File: alljp.m
- Processes multiple subjects and trials
- Computes:
  - JPE 3D
  - Axial rotation
  - Flexion/Extension
  - Lateral bending
- Outputs results to Excel

### 3. Statistical Analysis
- File: stat.m
- Performs:
  - Paired t-test
  - Wilcoxon signed-rank test
  - Effect size (Cohen’s d)
- Saves results to Excel

## Data Requirements
- Motion capture data (Qualisys export)
- File formats:
  - .xlsx (STIP)
  - .tsv or .xlsx (JPE)

## How to Run

1. Open MATLAB
2. Set working directory to project folder
3. Run scripts:

