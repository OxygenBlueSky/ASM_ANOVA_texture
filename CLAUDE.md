# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an R-based statistical analysis project for ASM (Arnica Stability Method) experiments. The project analyzes texture and fractal data from homeopathic crystallization experiments, comparing treatment solutions (verum) against water controls (SNC - Systematic Negative Control).

**Experimental Design:**
- 10 experiments total (5 verum, 5 SNC)
- Each experiment tests 6 coded solutions (A-F): 3 different solutions × 2 pairs each
- Each solution has 7 replicates (crystallization plates)
- SNC experiments: 3 water preparations (water_1, water_2, water_3)
- Verum experiments: 3 homeopathic solutions (Lactose, Stannum, Silicea)

## Core Workflow

The analysis follows a two-script workflow:

### 1. Data Decoding (`260114_ASM_TA_FA_data_decoding.r`)

**Purpose:** Decode raw texture and fractal data files by mapping coded solution letters (A-F) to actual solutions and pair numbers.

**Run this first** before any statistical analysis.

**Input files:**
- `20250910_ASM_1_10_texture_data.csv` - Raw texture analysis data from ACIA
- `20251218_ASM_1_10_fractal_data.csv` - Raw fractal analysis data from FracLac

**Output files:**
- `20250910_ASM_1_10_texture_data_DECODED.csv`
- `20251218_ASM_1_10_fractal_data_DECODED.csv`

**What it does:**
- Parses experiment names and potency codes from raw data
- Applies decoding tables to map coded letters (A-F) to solutions (water_1/2/3 or Lactose/Stannum/Silicea) and pairs (1/2)
- Adds metadata columns: `experiment_name`, `experiment_number`, `verum`, `SNC`, `sub_exp_number`, `decoded_solution`, `pair`
- For fractal data: parses filenames and joins with texture decoding info via `series_name` and `nr_in_chamber`

### 2. Unified ANOVA Analysis (`260114_ANOVA_ASM_unified.r`)

**Purpose:** Run complete statistical analysis (ANOVA, post-hoc tests, Cohen's d, plots) on either texture or fractal data.

**Key Configuration Toggles (lines 24-33):**
```r
analysis_type <- "texture"          # "texture" or "fractal"
remove_outliers <- FALSE            # TRUE or FALSE
outliers_file_name <- "20260115_ASM_outliers.csv"
```

**Output files** (created in timestamped folders like `20260114_ASM_TA_stats/`):
- `*_ASM_TA_ANOVA.xlsx` - Type III ANOVA p-values for all parameters across 6 scenarios
- `*_ASM_TA_posthoc_cohensd.xlsx` - Post-hoc pairwise comparisons and effect sizes
- `*_ASM_TA_parameters_grid.png` - Line plots showing parameter trends across sub-experiments

**Analysis scenarios:**
- ALL DATA - SNC / Verum
- JZ ONLY - SNC / Verum (Experimenter JZ: experiments 3-10)
- AS ONLY - SNC / Verum (Experimenter AS: experiments 1-2)

**Statistical approach:**
- Type III ANOVA with factors: `decoded_solution * experiment_number`
- Post-hoc pairwise comparisons using `emmeans` (no multiple comparison adjustment)
- Cohen's d effect sizes calculated from residual SD
- Color-coded Excel output: red (p<0.01), orange (p<0.05), lilac (p<0.10)

## Key Data Structures

**Experiment numbering:**
- Experiments 1-10 alternate between SNC and verum
- SNC: 1, 3, 6, 8, 9
- Verum: 2, 4, 5, 7, 10
- `sub_exp_number`: renumbered 1-5 within each condition

**Solution decoding:**
- Coded letters (A-F) map to different solutions across experiments (counterbalanced design)
- Each solution produced in duplicate pairs
- Decoding tables are hardcoded in `260114_ASM_TA_FA_data_decoding.r` (lines 166-266)

**Texture parameters** (15 total):
- GLCM features: cluster_prominence, cluster_shade, correlation, diagonal_moment, difference_energy, difference_entropy, energy, entropy, inertia, inverse_different_moment, kappa, maximum_probability, sum_energy, sum_entropy, sum_variance
- Plus: evaporation_duration
- Some parameters scaled by 1e6 for numerical stability

**Fractal parameters** (3 total):
- `db_mean`, `dm_mean`, `dx_mean` (different fractal dimension measures)

## Running Analyses

**Standard texture analysis:**
```r
# 1. Decode data (only needed once or when input data changes)
source("260114_ASM_TA_FA_data_decoding.r")

# 2. Run analysis
source("260114_ANOVA_ASM_unified.r")
```

**Fractal analysis:**
```r
# Change toggle in 260114_ANOVA_ASM_unified.r:
analysis_type <- "fractal"
# Then source the file
```

**With outlier removal:**
```r
# Ensure outlier CSV exists, then:
remove_outliers <- TRUE
outliers_file_name <- "20260115_ASM_outliers.csv"
```

**Outlier file format:**
- CSV with columns: `series_name`, `nr_in_chamber`
- Delimiter: semicolon (`;`)

## Required R Packages

```r
library(openxlsx)   # Excel file export
library(readxl)     # Excel file import
library(car)        # Type III ANOVA (Anova function)
library(here)       # File path management
library(dplyr)      # Data manipulation
library(emmeans)    # Post-hoc comparisons
library(ggplot2)    # Plotting
library(plotrix)    # std.error function
library(gridExtra)  # grid.arrange for multi-panel plots
library(colorspace) # darken() for color manipulation
library(RColorBrewer) # Color palettes
```

## File Organization

**Active analysis files:**
- `260114_ASM_TA_FA_data_decoding.r` - Data decoding script
- `260114_ANOVA_ASM_unified.r` - Unified ANOVA analysis

**Data files:**
- Raw data: `20250910_ASM_1_10_texture_data.csv`, `20251218_ASM_1_10_fractal_data.csv`
- Decoded data: `*_DECODED.csv` (excluded from git via .gitignore)
- Outlier lists: `*_outliers.csv`

**Output folders:**
- Pattern: `{date}_ASM_{TA/FA}_stats[_excl_{outliers}]/`
- Excluded from git via .gitignore

**OLD directory:**
- Contains archived versions and previous analysis iterations
- Excluded from git via .gitignore

## Important Notes

- The decoding tables in `260114_ASM_TA_FA_data_decoding.r` are **experiment-specific** and should not be modified unless the experimental design changes
- Always run the decoding script before analysis if input data has changed
- The unified ANOVA script uses configuration toggles at the top - modify these rather than changing code throughout
- Output files include data source information in headers for traceability
- Excel files use color coding for statistical significance - legend included on each sheet
- Fractal data must be joined with texture decoding via `series_name` and `nr_in_chamber` matching
