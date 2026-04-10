# ASM Unified ANOVA and Post-Hoc Analysis Script
# Created: 2026-01-14
# Purpose: Apply ANOVA and Cohen's d analysis on either texture or fractal data
# Toggle between analysis types using the configuration section below
# Run 260114_ASM_TA_FA_data_decoding.r first to have decoded files

#===== Load required packages =================================================

library(openxlsx)  # For Excel file handling
library(readxl)    # For reading Excel files
library(car)       # For Type III ANOVA (Anova function)
library(here)      # For file path management
library(dplyr)     # For data manipulation
library(emmeans)   # For post-hoc comparisons
library(ggplot2)   # For plotting
library(plotrix)   # For std.error
library(gridExtra) # For grid.arrange
library(colorspace) # For darken() function
library(GGally)    # For correlation matrix plots (ggpairs)


#===== CONFIGURATION - USER TOGGLES ==========================================

# USER TOGGLE 1: Analysis type - "texture" or "fractal"
analysis_type <- "texture"

# USER TOGGLE 2: Input files (pre-decoded from 260114_ASM_texture_data_decoding.r)
input_file_texture <- "20250910_ASM_1_10_texture_data_DECODED.csv"
input_file_fractal <- "20251218_ASM_1_10_fractal_data_DECODED.csv"

# USER TOGGLE 3: Outlier removal - TRUE or FALSE
remove_outliers <- TRUE
outliers_file_name <- "20260115_ASM_exp3.csv"

# USER TOGGLE 4: Create correlation matrix plots - TRUE or FALSE
create_correlation_plots <- FALSE

# USER TOGGLE 5: Pair-pooling diagnostic - TRUE or FALSE
# Runs a nested ANOVA (decoded_solution / pair) to test whether the two
# independently prepared pairs within each solution differ systematically.
# If non-significant, pooling pairs in the main ANOVA is justified.
run_pair_diagnostic <- TRUE

#===== DERIVED CONFIGURATION (based on toggles above) ========================

# Select input file based on analysis type
if (analysis_type == "texture") {
  input_file <- input_file_texture
} else {
  input_file <- input_file_fractal
}

# Use today's date as prefix for all output files and folders
date <- Sys.Date()
date2 <- gsub("-| |UTC", "", date)
input_prefix <- date2

# Create outlier suffix (only used if remove_outliers is TRUE)
# e.g., "20260108_ASM_outliers.csv" -> "20260108_ASM_outliers"
if (remove_outliers) {
  outliers_suffix <- tools::file_path_sans_ext(basename(outliers_file_name))
} else {
  outliers_suffix <- NULL
}

# Data-specific parameter configuration
if (analysis_type == "texture") {

  # Output file prefix
  output_prefix <- "_ASM_TA_"

  # Four focal parameters for detailed post-hoc and plotting

  analysis_vars <- c("cluster_shade", "entropy", "maximum_probability", "diagonal_moment", "evaporation_duration")

  # All 15 texture parameters for comprehensive Excel ANOVA summary
  all_params <- c(
    "cluster_prominence", "cluster_shade", "correlation",
    "diagonal_moment", "difference_energy", "difference_entropy",
    "energy", "entropy", "inertia", "inverse_different_moment",
    "kappa", "maximum_probability", "sum_energy",
    "sum_entropy", "sum_variance", "evaporation_duration"
  )

  # Texture values are very small, scale for numerical stability
  scale_factor <- 1e6
  params_to_scale <- c("maximum_probability", "energy", "sum_energy")

  # Plot grid dimensions
  plot_nrow <- 4
  plot_ncol <- 2
  plot_height <- 35  # cm

} else if (analysis_type == "fractal") {

  # Input file (pre-decoded from 260114_ASM_texture_data_decoding.r)
  input_file <- "20251218_ASM_1_10_fractal_data_DECODED.csv"

  # Output file prefix
  output_prefix <- "_ASM_FA_"

  # Fractal parameters for post-hoc and plotting
  analysis_vars <- c("db_mean", "dm_mean", "dx_mean")

  # All fractal parameters for ANOVA summary
  all_params <- c("db_mean", "dm_mean", "dx_mean")

  # Fractal values don't need scaling
  scale_factor <- 1

  params_to_scale <- c()

  # Plot grid dimensions (3 params instead of 4)
  plot_nrow <- 3
  plot_ncol <- 2
  plot_height <- 28  # cm

} else {
  stop("Invalid analysis_type. Must be 'texture' or 'fractal'")
}

cat("\n")
cat("################################################################################\n")
cat(sprintf("ANALYSIS TYPE: %s\n", toupper(analysis_type)))
cat("################################################################################\n\n")
cat(sprintf("Input file: %s\n", input_file))
cat(sprintf("Output prefix: %s\n", output_prefix))
cat(sprintf("Parameters to analyze: %s\n", paste(analysis_vars, collapse = ", ")))
cat(sprintf("Total parameters for ANOVA: %d\n\n", length(all_params)))


# Create output folder based on analysis type and outlier settings
# Format: {input_prefix}_ASM_{TA/FA}_stats[_excl_{outliers_suffix}]
if (analysis_type == "texture") {
  base_folder <- paste0(input_prefix, "_ASM_TA_stats")
} else {
  base_folder <- paste0(input_prefix, "_ASM_FA_stats")
}

if (!is.null(outliers_suffix)) {
  output_folder <- paste0(base_folder, "_excl_", outliers_suffix)
  outlier_file_suffix <- paste0("_excl_", outliers_suffix)
} else {
  output_folder <- base_folder
  outlier_file_suffix <- ""
}

# Create folder if it doesn't exist
if (!dir.exists(output_folder)) {
  dir.create(output_folder)
  cat(sprintf("Created output folder: %s\n", output_folder))
} else {
  cat(sprintf("Output folder exists: %s\n", output_folder))
}


#===== DATA LOADING ===========================================================

cat("========================================================================\n")
cat("LOADING PRE-DECODED DATA\n")
cat("========================================================================\n\n")

# Load pre-decoded data
df2 <- read.delim(here(input_file))

cat(sprintf("Data loaded from: %s\n", input_file))
cat(sprintf("Total rows: %d\n", nrow(df2)))
cat(sprintf("Total columns: %d\n\n", ncol(df2)))

# Verify required columns exist
required_cols <- c("experiment_number", "experiment_name", "potency",
                   "verum", "SNC", "sub_exp_number", "decoded_solution", "pair")

missing_cols <- setdiff(required_cols, colnames(df2))
if (length(missing_cols) > 0) {
  stop(sprintf("Missing required columns: %s", paste(missing_cols, collapse = ", ")))
}

cat("All required decoding columns present.\n\n")

# Rename cryptic fractal column names

if (analysis_type == "fractal") {
  colnames(df2)[6] <- "db_mean"   # Column 6: D̅ for Dʙ
  colnames(df2)[17] <- "dm_mean"  # Column 17: D̅ for Dʍ  
  colnames(df2)[26] <- "dx_mean"  # Column 26: D̅ for D͞ᵪ (best r²)
}

#===== OUTLIER REMOVAL ========================================================

cat("\n")
cat("========================================================================\n")
cat("OUTLIER REMOVAL\n")
cat("========================================================================\n")

if (!remove_outliers) {
  cat("Outlier removal: OFF\n")
  cat("No outliers removed\n")
  outlier_note <- "Outliers: not removed"

} else {
  # Read outlier file (contains series_name and nr_in_chamber columns)
  outliers_file <- here(outliers_file_name)

  if (!file.exists(outliers_file)) {
    cat(sprintf("Outlier file not found: %s\n", outliers_file))
    cat("No outliers removed\n")
    outlier_note <- "Outliers: file not found"

  } else {
    outliers <- read.csv(outliers_file, sep = ";")

    if (nrow(outliers) == 0) {
      cat("No outliers removed - outlier file is empty\n")
      outlier_note <- "Outliers: file empty"

    } else {
      cat(sprintf("Outliers file: %s\n", basename(outliers_file)))
      cat(sprintf("Outliers in file: %d\n\n", nrow(outliers)))

      n_before <- nrow(df2)

      # Match on series_name and nr_in_chamber (works for both TA and FA)
      df2 <- df2 %>%
        anti_join(outliers, by = c("series_name", "nr_in_chamber"))

      n_after <- nrow(df2)
      n_removed <- n_before - n_after

      cat(sprintf("Rows before removal: %d\n", n_before))
      cat(sprintf("Rows removed: %d\n", n_removed))
      cat(sprintf("Rows after removal: %d\n", n_after))

      if (n_removed != nrow(outliers)) {
        cat(sprintf("\nNote: %d outliers in file, %d removed (some may not exist in this dataset)\n",
                    nrow(outliers), n_removed))
      }

      outlier_note <- sprintf("Outliers removed: %s (%d removed)", basename(outliers_file), n_removed)
    }
  }
}

cat("========================================================================\n\n")

# Create data source note for use in output files
data_source_note <- sprintf("Input: %s | %s", basename(input_file), outlier_note)


#===== PREP FOR ANOVA =========================================================

cat("========================================================================\n")
cat("PREPARING DATA FOR ANOVA\n")
cat("========================================================================\n\n")

# Convert categorical variables to factors
df2$decoded_solution <- as.factor(df2$decoded_solution)
df2$experiment_number <- as.factor(df2$experiment_number)
df2$sub_exp_number <- as.factor(df2$sub_exp_number)
df2$verum <- as.factor(df2$verum)
df2$pair <- as.factor(df2$pair)

cat("Categorical variables converted to factors.\n")

# Set contrast options for Type III ANOVA
options(contrasts = c("contr.sum", "contr.poly"))
cat("Contrasts set for Type III ANOVA.\n")

# Scale small-value parameters if needed
if (length(params_to_scale) > 0) {
  cat(sprintf("\nScaling parameters by %.0e: %s\n",
              scale_factor, paste(params_to_scale, collapse = ", ")))

  for (param in params_to_scale) {
    if (param %in% colnames(df2)) {
      df2[[param]] <- df2[[param]] * scale_factor
    }
  }
}

cat("\nData preparation complete.\n\n")


# Data integrity check

cat("========================================================================\n")
cat("DATA INTEGRITY CHECK\n")
cat("========================================================================\n")
cat("Total observations in df2:", nrow(df2), "\n")
cat("Expected: 10 experiments x 6 potencies x 7 replicates = 420 observations\n\n")

# Check counts by experiment_number
cat("--- Observations per experiment_number ---\n")
exp_counts <- table(df2$experiment_number)
print(exp_counts)
cat("\n")

# Check counts by experiment_number and decoded_solution
cat("--- Observations per experiment_number x decoded_solution ---\n")
exp_sol_counts <- table(df2$experiment_number, df2$decoded_solution)
print(exp_sol_counts)
cat("\n")

cat("========================================================================\n\n")


#===== ANOVA HELPER FUNCTION ==================================================

# Function to run ANOVA and extract p-values
run_anova_summary <- function(data_subset, param_name) {

  # Check if parameter exists in dataset
  if (!(param_name %in% colnames(data_subset))) {
    return(data.frame(
      Experiment = "Error: param not found",
      Solution = "Error: param not found",
      Interaction = "Error: param not found",
      stringsAsFactors = FALSE
    ))
  }

  # Check number of unique experiments in this subset
  n_experiments <- length(unique(data_subset$experiment_number))

  # Initialize results data frame
  result <- data.frame(
    Experiment = NA,
    Solution = NA,
    Interaction = NA,
    stringsAsFactors = FALSE
  )

  # Use tryCatch to handle any ANOVA failures
  tryCatch({

    if (n_experiments > 1) {
      # Two-way ANOVA with interaction
      formula_str <- paste(param_name, "~ decoded_solution * experiment_number")
      model <- aov(as.formula(formula_str), data = data_subset)
      anova_results <- Anova(model, type = "III")

      # Extract p-values from ANOVA table
      result$Experiment <- anova_results["experiment_number", "Pr(>F)"]
      result$Solution <- anova_results["decoded_solution", "Pr(>F)"]
      result$Interaction <- anova_results["decoded_solution:experiment_number", "Pr(>F)"]

    } else {
      # Only one experiment - run one-way ANOVA for solution effect only
      formula_str <- paste(param_name, "~ decoded_solution")
      model <- aov(as.formula(formula_str), data = data_subset)
      anova_results <- Anova(model, type = "III")

      result$Experiment <- "NA (single exp)"
      result$Solution <- anova_results["decoded_solution", "Pr(>F)"]
      result$Interaction <- "NA (single exp)"
    }

  }, error = function(e) {
    cat(sprintf("  Warning: ANOVA failed for %s - %s\n", param_name, e$message))
    result$Experiment <<- "Error"
    result$Solution <<- "Error"
    result$Interaction <<- "Error"
  })

  return(result)
}


#===== ANOVA SUMMARY FOR ALL PARAMETERS =======================================

cat("========================================================================\n")
cat("CREATING ANOVA SUMMARY FOR ALL PARAMETERS\n")
cat("========================================================================\n\n")

# Create a new workbook
wb <- createWorkbook()

# Define color styles for significance levels
style_red <- createStyle(fgFill = "#FFB3BA")      # Light red for p < 0.01
style_orange <- createStyle(fgFill = "#FFDFBA")   # Light orange for p < 0.05
style_lilac <- createStyle(fgFill = "#E0BBE4")    # Light lilac for p < 0.10
style_header <- createStyle(textDecoration = "bold", fgFill = "#D3D3D3")

# Define scenarios
scenarios_excel <- list(
  list(name = "ALL DATA - SNC", short_name = "ALL_SNC",
       data = df2[df2$verum == 0, ]),
  list(name = "ALL DATA - Verum", short_name = "ALL_Verum",
       data = df2[df2$verum == 1, ]),

  list(name = "JZ ONLY - SNC", short_name = "JZ_SNC",
       data = df2[df2$experiment_name == "ASM_JZ" & df2$verum == 0, ]),
  list(name = "JZ ONLY - Verum", short_name = "JZ_Verum",
       data = df2[df2$experiment_name == "ASM_JZ" & df2$verum == 1, ]),

  list(name = "AS ONLY - SNC", short_name = "AS_SNC",
       data = df2[df2$experiment_name == "ASM_AS" & df2$verum == 0, ]),
  list(name = "AS ONLY - Verum", short_name = "AS_Verum",
       data = df2[df2$experiment_name == "ASM_AS" & df2$verum == 1, ])
)

# Loop through each scenario and create a sheet
for (scenario in scenarios_excel) {

  cat(sprintf("Processing %s for Excel export...\n", scenario$name))

  # Add a worksheet with the short name
  addWorksheet(wb, scenario$short_name)

  # Initialize results list
  results_list <- list()

  # Run ANOVA for all parameters
  for (param in all_params) {
    results_list[[param]] <- run_anova_summary(scenario$data, param)
  }

  # Combine all results into one data frame
  results_df <- do.call(rbind, results_list)

  # Add parameter names as first column
  results_df <- data.frame(
    Parameter = rownames(results_df),
    results_df,
    stringsAsFactors = FALSE
  )

  # Format numeric columns to 6 decimals
  for (col in c("Experiment", "Solution", "Interaction")) {
    for (row in 1:nrow(results_df)) {
      cell_value <- results_df[row, col]
      if (!is.na(suppressWarnings(as.numeric(cell_value)))) {
        results_df[row, col] <- sprintf("%.6f", as.numeric(cell_value))
      }
    }
  }

  # Write title and data source note to worksheet
  writeData(wb, scenario$short_name,
            paste0("ANOVA Summary: ", scenario$name, " (", toupper(analysis_type), " analysis)"),
            startRow = 1, startCol = 1)
  writeData(wb, scenario$short_name, data_source_note,
            startRow = 2, startCol = 1)

  # Write data starting at row 4
  writeData(wb, scenario$short_name, results_df, startRow = 4, rowNames = FALSE)

  # Apply header style
  addStyle(wb, scenario$short_name, style_header,
           rows = 4, cols = 1:4, gridExpand = TRUE)

  # Apply color coding based on p-values
  for (row_idx in 1:nrow(results_df)) {
    for (col_idx in 2:4) {

      col_name <- colnames(results_df)[col_idx]
      cell_value <- results_df[row_idx, col_name]

      if (grepl("Error|NA", cell_value)) {
        next
      }

      p_val <- suppressWarnings(as.numeric(cell_value))

      if (!is.na(p_val)) {
        excel_row <- row_idx + 4
        excel_col <- col_idx

        if (p_val < 0.01) {
          addStyle(wb, scenario$short_name, style_red,
                   rows = excel_row, cols = excel_col)
        } else if (p_val < 0.05) {
          addStyle(wb, scenario$short_name, style_orange,
                   rows = excel_row, cols = excel_col)
        } else if (p_val < 0.10) {
          addStyle(wb, scenario$short_name, style_lilac,
                   rows = excel_row, cols = excel_col)
        }
      }
    }
  }

  # Set column widths
  setColWidths(wb, scenario$short_name, cols = 1:4, widths = c(25, 15, 15, 15))

  # Add legend at the bottom
  legend_row <- nrow(results_df) + 6
  writeData(wb, scenario$short_name, "Color Legend:",
            startRow = legend_row, startCol = 1)
  writeData(wb, scenario$short_name, "Light lilac = p < 0.10",
            startRow = legend_row + 1, startCol = 1)
  writeData(wb, scenario$short_name, "Light orange = p < 0.05",
            startRow = legend_row + 2, startCol = 1)
  writeData(wb, scenario$short_name, "Light red = p < 0.01",
            startRow = legend_row + 3, startCol = 1)

  addStyle(wb, scenario$short_name, style_lilac,
           rows = legend_row + 1, cols = 1)
  addStyle(wb, scenario$short_name, style_orange,
           rows = legend_row + 2, cols = 1)
  addStyle(wb, scenario$short_name, style_red,
           rows = legend_row + 3, cols = 1)

  cat(sprintf("  %s complete - %d parameters analyzed\n", scenario$short_name, nrow(results_df)))
}

# Save workbook
output_filename_anova <- file.path(output_folder, paste0(input_prefix, output_prefix, "ANOVA", outlier_file_suffix, ".xlsx"))
saveWorkbook(wb, output_filename_anova, overwrite = TRUE)

cat("\n")
cat("========================================================================\n")
cat(sprintf("ANOVA summary exported to: %s\n", output_filename_anova))
cat("========================================================================\n\n")


#===== POST-HOC TESTS WITH COHEN'S D ==========================================

cat("========================================================================\n")
cat("POST-HOC TESTS WITH COHEN'S D\n")
cat("========================================================================\n\n")

# Computation function for emmeans contrasts
compute_emmeans_contrasts <- function(data_subset, variable) {

  n_experiments <- length(unique(data_subset$experiment_number))
  solutions <- levels(data_subset$decoded_solution)
  n_solutions <- length(solutions)

  cat(sprintf("  Solutions in this analysis: %s\n", paste(solutions, collapse = ", ")))
  cat(sprintf("  Number of experiments: %d\n", n_experiments))

  # Initialize n×n pairwise comparison matrices (only lower triangle will be filled;
  # upper triangle stays NA to avoid redundancy, since A-vs-B = B-vs-A).
  #   potency_matrix:     unadjusted p-values for main-effect pairwise contrasts
  #   interaction_matrix:  p-values testing whether each pairwise difference is
  #                        consistent across experiments (per-pair interaction)
  #   cohens_d_matrix:     standardized effect sizes for each pairwise comparison
  potency_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
                           dimnames = list(solutions, solutions))
  interaction_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
                               dimnames = list(solutions, solutions))
  cohens_d_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
                            dimnames = list(solutions, solutions))

  # Determine model based on number of experiments
  if (n_experiments > 1) {
    formula_str <- paste(variable, "~ decoded_solution * experiment_number")
  } else {
    formula_str <- paste(variable, "~ decoded_solution")
  }

  # Fit ANOVA model
  model <- aov(as.formula(formula_str), data = data_subset)

  # --- Main-effect pairwise comparisons ---
  # Statistical goal: estimate the mean difference between every pair of
  # decoded_solution levels, averaging over experiments. P-values are
  # unadjusted (no Bonferroni/Tukey) because we report them alongside
  # Cohen's d — the effect size is more informative than corrected p-values
  # when the number of comparisons is modest.
  #
  # emmeans averages over the experiment factor, giving one estimated
  # marginal mean per solution level; pairs() computes all pairwise
  # differences and their t-tests.
  emm_solution <- emmeans(model, ~ decoded_solution)
  pairs_solution <- pairs(emm_solution, adjust = "none")
  pairs_summary <- summary(pairs_solution)

  # Parse each contrast name (e.g., "water_1 - water_2") into its two solution
  # names, then place the p-value into the lower triangle of potency_matrix
  # at [row, col] where row > col (ensures consistent lower-triangle layout).
  for (i in 1:nrow(pairs_summary)) {
    contrast_name <- as.character(pairs_summary$contrast[i])
    parts <- strsplit(contrast_name, " - ")[[1]]
    solution1 <- trimws(parts[1])
    solution2 <- trimws(parts[2])

    p_val <- pairs_summary$p.value[i]

    idx1 <- which(solutions == solution1)
    idx2 <- which(solutions == solution2)

    if (idx1 > idx2) {
      potency_matrix[idx1, idx2] <- p_val
    } else {
      potency_matrix[idx2, idx1] <- p_val
    }
  }

  # --- Cohen's d = mean difference / residual SD ---
  
  # Statistical goal: quantify pairwise effect sizes on a standardized scale
  # so they are comparable across parameters with different units/ranges.
  # We use the ANOVA model's residual SD (pooled within-group SD) as the
  # denominator — this is the standard choice for balanced/near-balanced
  # designs and gives the same denominator for all pairs within a model.
  
  residual_sd <- sigma(model)

  for (i in 1:nrow(pairs_summary)) {
    contrast_name <- as.character(pairs_summary$contrast[i])
    parts <- strsplit(contrast_name, " - ")[[1]]
    solution1 <- trimws(parts[1])
    solution2 <- trimws(parts[2])

    mean_diff <- pairs_summary$estimate[i]
    cohens_d <- mean_diff / residual_sd

    idx1 <- which(solutions == solution1)
    idx2 <- which(solutions == solution2)

    if (idx1 > idx2) {
      cohens_d_matrix[idx1, idx2] <- cohens_d
    } else {
      cohens_d_matrix[idx2, idx1] <- cohens_d
    }
  }

  # --- Interaction contrasts (only if multiple experiments) ---
  
  # Statistical goal: test whether each pairwise solution difference is
  # consistent across experiments, or whether the treatment effect depends
  # on which experiment we look at (a per-pair solution x experiment
  # interaction). A significant result means the difference between those
  # two solutions is not stable across experiments — i.e., the effect
  # does not replicate.
  #
  # Method: For each pair of solutions (e.g., water_1 vs water_2):
  #   1. Compute per-experiment contrasts via emmeans (emm_by_exp gives
  #      one estimated mean per solution per experiment; pairs() gives
  #      the difference within each experiment separately).
  #   2. Find the rows in pairs_by_exp matching this specific contrast
  #      name — one row per experiment.
  #   3. Apply pairs() again on those per-experiment contrasts to get
  #      differences-of-differences (how much the contrast changes
  #      between experiments).
  #   4. Run a joint F-test via test(joint = TRUE) to get a single
  #      omnibus p-value for whether those differences-of-differences
  #      are collectively zero.
  # The p-values fill the lower triangle of interaction_matrix
  # (same indexing convention as potency_matrix).
  
  if (n_experiments > 1) {
    tryCatch({
      emm_by_exp <- emmeans(model, ~ decoded_solution | experiment_number)
      pairs_by_exp <- pairs(emm_by_exp)
      pairs_by_exp_summary <- summary(pairs_by_exp)

      for (i in 1:(n_solutions-1)) {
        for (j in (i+1):n_solutions) {
          solution1 <- solutions[i]
          solution2 <- solutions[j]

          comparison_name <- paste(solution1, "-", solution2)
          matching_rows <- grep(paste0("^", comparison_name, "$"),
                                pairs_by_exp_summary$contrast,
                                fixed = FALSE)

          if (length(matching_rows) >= 2) {
            tryCatch({
              specific_contrasts <- pairs_by_exp[matching_rows]
              pairs_of_contrasts <- pairs(specific_contrasts)
              joint_test <- test(pairs_of_contrasts, joint = TRUE)
              p_val_int <- joint_test$p.value
              interaction_matrix[j, i] <- p_val_int
            }, error = function(e) {
              interaction_matrix[j, i] <<- "Error//pair"
            })
          } else {
            interaction_matrix[j, i] <- "Error/no of exp"
          }
        }
      }
    }, error = function(e) {
      interaction_matrix[lower.tri(interaction_matrix)] <- "Error//function"
    })
  } else {
    interaction_matrix[lower.tri(interaction_matrix)] <- "NA (single exp)"
  }

  return(list(
    potency = potency_matrix,
    interaction = interaction_matrix,
    cohens_d = cohens_d_matrix
  ))
}


# Create Excel workbook for post-hoc results
wb_posthoc <- createWorkbook()

style_red <- createStyle(fgFill = "#FFB3BA")
style_orange <- createStyle(fgFill = "#FFDFBA")
style_lilac <- createStyle(fgFill = "#E0BBE4")
style_header <- createStyle(textDecoration = "bold",
                            fgFill = "#D3D3D3",
                            halign = "center",
                            valign = "center")
style_bold <- createStyle(textDecoration = "bold")

# Define scenarios for post-hoc
scenarios_posthoc <- list(
  list(name = "ALL DATA - SNC", short_name = "ALL_SNC",
       data = droplevels(df2[df2$verum == 0, ])),
  list(name = "ALL DATA - Verum", short_name = "ALL_Verum",
       data = droplevels(df2[df2$verum == 1, ])),
  list(name = "JZ ONLY - SNC", short_name = "JZ_SNC",
       data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 0, ])),
  list(name = "JZ ONLY - Verum", short_name = "JZ_Verum",
       data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 1, ])),
  list(name = "AS ONLY - SNC", short_name = "AS_SNC",
       data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 0, ])),
  list(name = "AS ONLY - Verum", short_name = "AS_Verum",
       data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 1, ]))
)


# P-VALUE WORKSHEETS

cat("Creating p-value worksheets...\n\n")

for (scenario in scenarios_posthoc) {

  cat(sprintf("Processing: %s\n", scenario$name))

  addWorksheet(wb_posthoc, scenario$short_name)
  current_row <- 1

  writeData(wb_posthoc, scenario$short_name,
            paste0("Post Hoc Tests: ", scenario$name, " (", toupper(analysis_type), " analysis)"),
            startRow = current_row, startCol = 1)
  addStyle(wb_posthoc, scenario$short_name, style_bold,
           rows = current_row, cols = 1)

  current_row <- current_row + 1
  writeData(wb_posthoc, scenario$short_name, data_source_note,
            startRow = current_row, startCol = 1)

  current_row <- current_row + 2

  for (var in analysis_vars) {

    cat(sprintf("  Parameter: %s\n", var))

    # Check if variable exists
    if (!(var %in% colnames(scenario$data))) {
      cat(sprintf("    WARNING: Variable %s not found in data, skipping\n", var))
      next
    }

    results <- compute_emmeans_contrasts(scenario$data, var)

    potency_matrix <- results$potency
    interaction_matrix <- results$interaction

    solutions <- rownames(potency_matrix)
    n_solutions <- length(solutions)

    # Table header
    writeData(wb_posthoc, scenario$short_name, var,
              startRow = current_row, startCol = 1)

    for (j in 1:n_solutions) {
      writeData(wb_posthoc, scenario$short_name, solutions[j],
                startRow = current_row, startCol = 1 + (j-1)*2 + 1)
      mergeCells(wb_posthoc, scenario$short_name,
                 rows = current_row,
                 cols = (1 + (j-1)*2 + 1):(1 + (j-1)*2 + 2))
    }

    addStyle(wb_posthoc, scenario$short_name, style_header,
             rows = current_row, cols = 1:(2*n_solutions + 1), gridExpand = TRUE)

    current_row <- current_row + 1

    writeData(wb_posthoc, scenario$short_name, "",
              startRow = current_row, startCol = 1)

    for (j in 1:n_solutions) {
      writeData(wb_posthoc, scenario$short_name, "solution",
                startRow = current_row, startCol = 1 + (j-1)*2 + 1)
      writeData(wb_posthoc, scenario$short_name, "interaction",
                startRow = current_row, startCol = 1 + (j-1)*2 + 2)
    }

    addStyle(wb_posthoc, scenario$short_name, style_header,
             rows = current_row, cols = 1:(2*n_solutions + 1), gridExpand = TRUE)

    current_row <- current_row + 1

    # Write p-values into the lower triangle of the n×n solution comparison
    # matrix. Each cell holds one pairwise contrast; upper triangle is left
    # empty to avoid redundancy. Each solution gets two columns: "solution"
    # (main-effect p) and "interaction" (per-pair experiment interaction p).
    
    for (i in 1:n_solutions) {
      writeData(wb_posthoc, scenario$short_name, solutions[i],
                startRow = current_row, startCol = 1)

      for (j in 1:n_solutions) {
        col_sol <- 1 + (j-1)*2 + 1
        col_int <- col_sol + 1

        if (j < i) {
          p_sol <- potency_matrix[i, j]
          p_int <- interaction_matrix[i, j]

          if (is.numeric(p_sol) && !is.na(p_sol)) {
            writeData(wb_posthoc, scenario$short_name, sprintf("%.4f", p_sol),
                      startRow = current_row, startCol = col_sol)

            if (p_sol < 0.01) {
              addStyle(wb_posthoc, scenario$short_name, style_red,
                       rows = current_row, cols = col_sol, stack = TRUE)
            } else if (p_sol < 0.05) {
              addStyle(wb_posthoc, scenario$short_name, style_orange,
                       rows = current_row, cols = col_sol, stack = TRUE)
            } else if (p_sol < 0.10) {
              addStyle(wb_posthoc, scenario$short_name, style_lilac,
                       rows = current_row, cols = col_sol, stack = TRUE)
            }
          } else if (!is.na(p_sol)) {
            writeData(wb_posthoc, scenario$short_name, as.character(p_sol),
                      startRow = current_row, startCol = col_sol)
          }

          if (is.numeric(p_int) && !is.na(p_int)) {
            writeData(wb_posthoc, scenario$short_name, sprintf("%.4f", p_int),
                      startRow = current_row, startCol = col_int)

            if (p_int < 0.01) {
              addStyle(wb_posthoc, scenario$short_name, style_red,
                       rows = current_row, cols = col_int, stack = TRUE)
            } else if (p_int < 0.05) {
              addStyle(wb_posthoc, scenario$short_name, style_orange,
                       rows = current_row, cols = col_int, stack = TRUE)
            } else if (p_int < 0.10) {
              addStyle(wb_posthoc, scenario$short_name, style_lilac,
                       rows = current_row, cols = col_int, stack = TRUE)
            }
          } else if (!is.na(p_int)) {
            writeData(wb_posthoc, scenario$short_name, as.character(p_int),
                      startRow = current_row, startCol = col_int)
          }
        }
      }

      current_row <- current_row + 1
    }

    current_row <- current_row + 2
  }

  # Set column widths
  setColWidths(wb_posthoc, scenario$short_name,
               cols = 1:(2*n_solutions + 1),
               widths = c(15, rep(10, 2*n_solutions)))

  # Add legend
  legend_row <- current_row + 1
  writeData(wb_posthoc, scenario$short_name, "Color Legend:",
            startRow = legend_row, startCol = 1)
  writeData(wb_posthoc, scenario$short_name, "Light lilac = p < 0.10",
            startRow = legend_row + 1, startCol = 1)
  addStyle(wb_posthoc, scenario$short_name, style_lilac,
           rows = legend_row + 1, cols = 1)
  writeData(wb_posthoc, scenario$short_name, "Light orange = p < 0.05",
            startRow = legend_row + 2, startCol = 1)
  addStyle(wb_posthoc, scenario$short_name, style_orange,
           rows = legend_row + 2, cols = 1)
  writeData(wb_posthoc, scenario$short_name, "Light red = p < 0.01",
            startRow = legend_row + 3, startCol = 1)
  addStyle(wb_posthoc, scenario$short_name, style_red,
           rows = legend_row + 3, cols = 1)
}


# COHEN'S D WORKSHEETS

cat("\nCreating Cohen's d worksheets...\n\n")

for (scenario in scenarios_posthoc) {

  cat(sprintf("Processing Cohen's d: %s\n", scenario$name))

  sheet_name_d <- paste0(scenario$short_name, "_d")
  addWorksheet(wb_posthoc, sheet_name_d)

  current_row <- 1

  writeData(wb_posthoc, sheet_name_d,
            paste0("Cohen's d Effect Sizes: ", scenario$name, " (", toupper(analysis_type), " analysis)"),
            startRow = current_row, startCol = 1)
  addStyle(wb_posthoc, sheet_name_d, style_bold,
           rows = current_row, cols = 1)

  current_row <- current_row + 1
  writeData(wb_posthoc, sheet_name_d, data_source_note,
            startRow = current_row, startCol = 1)

  current_row <- current_row + 2

  for (var in analysis_vars) {

    cat(sprintf("  Parameter: %s\n", var))

    if (!(var %in% colnames(scenario$data))) {
      cat(sprintf("    WARNING: Variable %s not found in data, skipping\n", var))
      next
    }

    results <- compute_emmeans_contrasts(scenario$data, var)
    cohens_d_matrix <- results$cohens_d

    solutions <- rownames(cohens_d_matrix)
    n_solutions <- length(solutions)

    # Table header
    writeData(wb_posthoc, sheet_name_d, var,
              startRow = current_row, startCol = 1)

    for (j in 1:n_solutions) {
      writeData(wb_posthoc, sheet_name_d, solutions[j],
                startRow = current_row, startCol = 1 + j)
    }

    addStyle(wb_posthoc, sheet_name_d, style_header,
             rows = current_row, cols = 1:(n_solutions + 1), gridExpand = TRUE)

    current_row <- current_row + 1

    # Write Cohen's d values into lower triangle of the comparison matrix
    # (same layout as the p-value matrix above, one column per solution).
    for (i in 1:n_solutions) {
      writeData(wb_posthoc, sheet_name_d, solutions[i],
                startRow = current_row, startCol = 1)

      for (j in 1:n_solutions) {
        col_d <- 1 + j

        if (j < i) {
          d_val <- cohens_d_matrix[i, j]

          if (is.numeric(d_val) && !is.na(d_val)) {
            writeData(wb_posthoc, sheet_name_d, sprintf("%.3f", d_val),
                      startRow = current_row, startCol = col_d)
          } else if (!is.na(d_val)) {
            writeData(wb_posthoc, sheet_name_d, as.character(d_val),
                      startRow = current_row, startCol = col_d)
          }
        }
      }

      current_row <- current_row + 1
    }

    current_row <- current_row + 2
  }

  # Set column widths
  setColWidths(wb_posthoc, sheet_name_d,
               cols = 1:(n_solutions + 1),
               widths = c(15, rep(10, n_solutions)))

  # Add interpretation note
  legend_row <- current_row + 1
  writeData(wb_posthoc, sheet_name_d, "Cohen's d interpretation:",
            startRow = legend_row, startCol = 1)
  writeData(wb_posthoc, sheet_name_d, "Small effect: |d| = 0.2",
            startRow = legend_row + 1, startCol = 1)
  writeData(wb_posthoc, sheet_name_d, "Medium effect: |d| = 0.5",
            startRow = legend_row + 2, startCol = 1)
  writeData(wb_posthoc, sheet_name_d, "Large effect: |d| = 0.8",
            startRow = legend_row + 3, startCol = 1)
}

# Save workbook
output_filename_posthoc <- file.path(output_folder, paste0(input_prefix, output_prefix, "posthoc_cohensd", outlier_file_suffix, ".xlsx"))
saveWorkbook(wb_posthoc, output_filename_posthoc, overwrite = TRUE)

cat("\n")
cat("========================================================================\n")
cat(sprintf("Post hoc tests exported to: %s\n", output_filename_posthoc))
cat("========================================================================\n\n")


#===== PLOTTING ===============================================================

cat("========================================================================\n")
cat("CREATING PLOTS\n")
cat("========================================================================\n\n")

# Define conditions
conditions <- c(0, 1)
condition_labels <- c("SNC (Water Controls)", "Verum (Treatment)")

# Create solution-pair combination variable
df2 <- df2 %>%
  mutate(
    # Self-describing label of the form "<solution>-<pair>" so that future
    # solution renames flow through automatically (e.g. water_1-1, Lactose-2)
    solution_pair = paste0(decoded_solution, "-", pair),
    base_solution = case_when(
      decoded_solution == "water_1" ~ "sol1",
      decoded_solution == "water_2" ~ "sol2",
      decoded_solution == "water_3" ~ "sol3",
      decoded_solution == "Lactose" ~ "sol1",
      decoded_solution == "Stannum" ~ "sol2",
      decoded_solution == "Silicea" ~ "sol3",
      TRUE ~ NA_character_
    )
  )

cat("Solution-pair combinations created.\n")

# Filter to only include analysis_vars that exist in the data
existing_vars <- analysis_vars[analysis_vars %in% colnames(df2)]

if (length(existing_vars) == 0) {
  cat("WARNING: None of the analysis_vars found in data. Skipping plots.\n")
} else {

  cat(sprintf("Plotting parameters: %s\n", paste(existing_vars, collapse = ", ")))

  # Calculate summary statistics
  summary_data <- df2 %>%
    group_by(verum, sub_exp_number, decoded_solution, pair, solution_pair, base_solution) %>%
    summarise(
      across(all_of(existing_vars),
             list(mean = ~mean(.x, na.rm = TRUE),
                  se = ~std.error(.x, na.rm = TRUE)),
             .names = "{.col}_{.fn}"),
      n = n(),
      .groups = "drop"
    )

  cat("Summary statistics calculated.\n")

  # Determine y-axis limits
  y_limits <- list()
  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")

    y_min <- min(summary_data[[mean_col]] - summary_data[[se_col]], na.rm = TRUE)
    y_max <- max(summary_data[[mean_col]] + summary_data[[se_col]], na.rm = TRUE)

    y_range <- y_max - y_min
    y_limits[[var]] <- c(y_min - 0.05 * y_range, y_max + 0.05 * y_range)
  }

  # Color palette: 3 base colors (one per solution), darkened variant for pair 2,
  # so replicates of the same solution are visually linked but distinguishable.
  dark2_colors <- RColorBrewer::brewer.pal(8, "Dark2")
  base_colors <- dark2_colors[4:6]
  darkened_colors <- darken(base_colors, amount = 0.3)

  # Keys must match the solution_pair strings built via paste0(decoded_solution, "-", pair)
  # Pair 1 gets the base color, pair 2 gets the darkened variant of the same hue,
  # so the two replicates of each solution stay visually linked but distinguishable.
  color_mapping_snc <- c(
    "water_1-1" = base_colors[1], "water_1-2" = darkened_colors[1],
    "water_2-1" = base_colors[2], "water_2-2" = darkened_colors[2],
    "water_3-1" = base_colors[3], "water_3-2" = darkened_colors[3]
  )

  # Slot order must match color_mapping_verum_pooled (Lactose, Stannum, Silicea)
  # so that each substance keeps the same hue across the pair-level, pooled,
  # normalized and scatter plots. ggplot will still sort the legend entries
  # alphabetically (Lactose, Silicea, Stannum) — that's a display-order issue,
  # not a color-assignment issue; the hue stays glued to the substance.
  color_mapping_verum <- c(
    "Lactose-1" = base_colors[1], "Lactose-2" = darkened_colors[1],
    "Stannum-1" = base_colors[2], "Stannum-2" = darkened_colors[2],
    "Silicea-1" = base_colors[3], "Silicea-2" = darkened_colors[3]
  )

  # Create plots
  plot_list <- list()
  plot_counter <- 1

  for (var in existing_vars) {

    cat(sprintf("  Creating plots for: %s\n", var))

    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")

    for (condition_idx in seq_along(conditions)) {

      verum_value <- conditions[condition_idx]
      condition_label <- condition_labels[condition_idx]

      plot_data <- summary_data %>%
        filter(verum == verum_value)

      if (verum_value == 0) {
        color_map <- color_mapping_snc
      } else {
        color_map <- color_mapping_verum
      }

      dodge_width <- 0.3

      p <- ggplot(plot_data, aes(x = as.numeric(as.character(sub_exp_number)),
                                 y = .data[[mean_col]],
                                 color = solution_pair,
                                 group = solution_pair)) +
        geom_line(linewidth = 0.7, position = position_dodge(width = dodge_width)) +
        geom_errorbar(aes(ymin = .data[[mean_col]] - .data[[se_col]],
                          ymax = .data[[mean_col]] + .data[[se_col]]),
                      width = 0.2, linewidth = 0.5,
                      position = position_dodge(width = dodge_width)) +
        geom_point(size = 2.5, position = position_dodge(width = dodge_width)) +
        scale_x_continuous(breaks = 1:5, labels = 1:5) +
        scale_y_continuous(limits = y_limits[[var]]) +
        scale_color_manual(values = color_map, name = "Solution") +
        labs(
          title = condition_label,
          x = "Sub-experiment Number",
          y = var,
          caption = "Note: Sub-exp 1 = AS data; Sub-exp 2-5 = JZ data"
        ) +
        theme_bw() +
        theme(
          plot.title = element_text(size = 11, face = "bold"),
          axis.title = element_text(size = 10),
          axis.text = element_text(size = 9),
          legend.title = element_text(size = 9),
          legend.text = element_text(size = 8),
          plot.caption = element_text(size = 8, hjust = 0),
          # Lock panel shape (height/width) so widening the figure to fit the
          # 12-entry legend doesn't also stretch the data area. 0.58 ≈ a ~20%
          # horizontal squeeze relative to the unlocked 19x8.75 cm grid cell.
          aspect.ratio = 0.58
        )

      plot_list[[plot_counter]] <- p
      plot_counter <- plot_counter + 1
    }
  }

  cat("All plots created.\n")

  # Arrange plots in grid
  # Adjust grid dimensions based on number of parameters
  actual_nrow <- length(existing_vars)
  actual_height <- plot_height * (actual_nrow / plot_nrow)

  grid_plot <- grid.arrange(
    grobs = plot_list,
    nrow = actual_nrow,
    ncol = 2,
    top = paste0("ASM ", toupper(analysis_type), " Parameters\n", data_source_note)
  )

  # Export plot
  output_filename_plot <- file.path(output_folder, paste0(input_prefix, output_prefix, "parameters_grid", outlier_file_suffix, ".png"))

  ggsave(
    filename = output_filename_plot,
    plot = grid_plot,
    width = 30,  # aspect.ratio in the theme locks panel shape, so the legend
                 # fits inside the leftover cell width without needing a wider figure
    height = actual_height,
    dpi = 300,
    units = "cm",
    limitsize = FALSE
  )

  cat(sprintf("Plot saved as: %s\n", output_filename_plot))


  #----- PLOT 2: Pairs-pooled (3 lines per panel) ------------------------------

  cat("\n--- Creating pairs-pooled plot (3 solutions, pairs averaged) ---\n")

  # Pool pairs: mean and SE across all ~14 replicates per solution per sub-experiment
  summary_pooled <- df2 %>%
    group_by(verum, sub_exp_number, decoded_solution, base_solution) %>%
    summarise(
      across(all_of(existing_vars),
             list(mean = ~mean(.x, na.rm = TRUE),
                  se = ~std.error(.x, na.rm = TRUE)),
             .names = "{.col}_{.fn}"),
      n = n(),
      .groups = "drop"
    )

  # Color mapping for 3-line plots: one color per decoded solution
  color_mapping_snc_pooled <- c(
    "water_1" = base_colors[1],
    "water_2" = base_colors[2],
    "water_3" = base_colors[3]
  )
  color_mapping_verum_pooled <- c(
    "Lactose" = base_colors[1],
    "Stannum" = base_colors[2],
    "Silicea" = base_colors[3]
  )

  # Y-axis limits for pooled plot (same approach as original)
  y_limits_pooled <- list()
  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")
    y_min <- min(summary_pooled[[mean_col]] - summary_pooled[[se_col]], na.rm = TRUE)
    y_max <- max(summary_pooled[[mean_col]] + summary_pooled[[se_col]], na.rm = TRUE)
    y_range <- y_max - y_min
    y_limits_pooled[[var]] <- c(y_min - 0.05 * y_range, y_max + 0.05 * y_range)
  }

  # Build pooled plot panels
  plot_list_pooled <- list()
  plot_counter_pooled <- 1

  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")

    for (condition_idx in seq_along(conditions)) {
      verum_value <- conditions[condition_idx]
      condition_label <- condition_labels[condition_idx]

      plot_data <- summary_pooled %>% filter(verum == verum_value)
      color_map <- if (verum_value == 0) color_mapping_snc_pooled else color_mapping_verum_pooled
      dodge_width <- 0.3

      p <- ggplot(plot_data, aes(x = as.numeric(as.character(sub_exp_number)),
                                 y = .data[[mean_col]],
                                 color = decoded_solution,
                                 group = decoded_solution)) +
        geom_line(linewidth = 0.7, position = position_dodge(width = dodge_width)) +
        geom_errorbar(aes(ymin = .data[[mean_col]] - .data[[se_col]],
                          ymax = .data[[mean_col]] + .data[[se_col]]),
                      width = 0.2, linewidth = 0.5,
                      position = position_dodge(width = dodge_width)) +
        geom_point(size = 2.5, position = position_dodge(width = dodge_width)) +
        scale_x_continuous(breaks = 1:5, labels = 1:5) +
        scale_y_continuous(limits = y_limits_pooled[[var]]) +
        scale_color_manual(values = color_map, name = "Solution") +
        labs(
          title = condition_label,
          x = "Sub-experiment Number",
          y = var,
          caption = "Note: Sub-exp 1 = AS data; Sub-exp 2-5 = JZ data"
        ) +
        theme_bw() +
        theme(
          plot.title = element_text(size = 11, face = "bold"),
          axis.title = element_text(size = 10),
          axis.text = element_text(size = 9),
          legend.title = element_text(size = 9),
          legend.text = element_text(size = 8),
          plot.caption = element_text(size = 8, hjust = 0)
        )

      plot_list_pooled[[plot_counter_pooled]] <- p
      plot_counter_pooled <- plot_counter_pooled + 1
    }
  }

  # Arrange and save pooled plot
  grid_plot_pooled <- grid.arrange(
    grobs = plot_list_pooled,
    nrow = actual_nrow,
    ncol = 2,
    top = paste0("ASM ", toupper(analysis_type), " Parameters (pairs pooled)\n", data_source_note)
  )

  output_filename_pooled <- file.path(output_folder, paste0(input_prefix, output_prefix, "parameters_grid_pooled", outlier_file_suffix, ".png"))
  ggsave(
    filename = output_filename_pooled,
    plot = grid_plot_pooled,
    width = 30,
    height = actual_height,
    dpi = 300,
    units = "cm",
    limitsize = FALSE
  )
  cat(sprintf("Pooled plot saved as: %s\n", output_filename_pooled))


  #----- PLOT 3: Normalized to reference solution (% of reference mean) --------

  cat("\n--- Creating normalized plot (% relative to reference) ---\n")
  cat("  Reference solutions: Lactose (verum), water_1 (SNC)\n")

  # Normalize each parameter to the reference solution's mean within each
  # (verum, sub_exp_number) combination. SE is scaled by the same factor so
  # error bars remain proportional.
  summary_normalized <- summary_pooled

  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")

    # Extract reference means: Lactose for verum, water_1 for SNC
    ref_means <- summary_pooled %>%
      filter((verum == 1 & decoded_solution == "Lactose") |
             (verum == 0 & decoded_solution == "water_1")) %>%
      select(verum, sub_exp_number, ref_mean = !!sym(mean_col))

    summary_normalized <- summary_normalized %>%
      left_join(ref_means, by = c("verum", "sub_exp_number")) %>%
      mutate(
        !!se_col   := (.data[[se_col]]   / abs(ref_mean)) * 100,
        !!mean_col := (.data[[mean_col]] / ref_mean) * 100
      ) %>%
      select(-ref_mean)
  }

  # Y-axis limits for normalized plot
  y_limits_norm <- list()
  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")
    y_min <- min(summary_normalized[[mean_col]] - summary_normalized[[se_col]], na.rm = TRUE)
    y_max <- max(summary_normalized[[mean_col]] + summary_normalized[[se_col]], na.rm = TRUE)
    y_range <- y_max - y_min
    y_limits_norm[[var]] <- c(y_min - 0.05 * y_range, y_max + 0.05 * y_range)
  }

  # Build normalized plot panels
  plot_list_norm <- list()
  plot_counter_norm <- 1

  for (var in existing_vars) {
    mean_col <- paste0(var, "_mean")
    se_col <- paste0(var, "_se")

    for (condition_idx in seq_along(conditions)) {
      verum_value <- conditions[condition_idx]
      condition_label <- condition_labels[condition_idx]

      # Label for y-axis: which solution is the 100% reference
      ref_label <- if (verum_value == 0) "water_1" else "Lactose"

      plot_data <- summary_normalized %>% filter(verum == verum_value)
      color_map <- if (verum_value == 0) color_mapping_snc_pooled else color_mapping_verum_pooled
      dodge_width <- 0.3

      p <- ggplot(plot_data, aes(x = as.numeric(as.character(sub_exp_number)),
                                 y = .data[[mean_col]],
                                 color = decoded_solution,
                                 group = decoded_solution)) +
        geom_hline(yintercept = 100, linetype = "dashed", color = "grey50", linewidth = 0.4) +
        geom_line(linewidth = 0.7, position = position_dodge(width = dodge_width)) +
        geom_errorbar(aes(ymin = .data[[mean_col]] - .data[[se_col]],
                          ymax = .data[[mean_col]] + .data[[se_col]]),
                      width = 0.2, linewidth = 0.5,
                      position = position_dodge(width = dodge_width)) +
        geom_point(size = 2.5, position = position_dodge(width = dodge_width)) +
        scale_x_continuous(breaks = 1:5, labels = 1:5) +
        scale_y_continuous(limits = y_limits_norm[[var]]) +
        scale_color_manual(values = color_map, name = "Solution") +
        labs(
          title = condition_label,
          x = "Sub-experiment Number",
          y = paste0(var, "\n(% of ", ref_label, ")"),
          caption = "Note: Sub-exp 1 = AS data; Sub-exp 2-5 = JZ data"
        ) +
        theme_bw() +
        theme(
          plot.title = element_text(size = 11, face = "bold"),
          axis.title = element_text(size = 10),
          axis.text = element_text(size = 9),
          legend.title = element_text(size = 9),
          legend.text = element_text(size = 8),
          plot.caption = element_text(size = 8, hjust = 0)
        )

      plot_list_norm[[plot_counter_norm]] <- p
      plot_counter_norm <- plot_counter_norm + 1
    }
  }

  # Arrange and save normalized plot
  grid_plot_norm <- grid.arrange(
    grobs = plot_list_norm,
    nrow = actual_nrow,
    ncol = 2,
    top = paste0("ASM ", toupper(analysis_type), " Parameters (normalized to reference)\n", data_source_note)
  )

  output_filename_norm <- file.path(output_folder, paste0(input_prefix, output_prefix, "parameters_grid_normalized", outlier_file_suffix, ".png"))
  ggsave(
    filename = output_filename_norm,
    plot = grid_plot_norm,
    width = 30,
    height = actual_height,
    dpi = 300,
    units = "cm",
    limitsize = FALSE
  )
  cat(sprintf("Normalized plot saved as: %s\n", output_filename_norm))

}


#===== CORRELATION MATRIX PLOTS ================================================

# Create five ggpairs correlation matrices (ALL / AS / JZ / SNC / Verum)
# showing pairwise relationships between all parameters + evaporation_duration.
#
# Plot structure:
#   - Upper triangle: Pearson correlation coefficients with green color gradient
#   - Lower triangle: Scatterplots of individual observations
#   - Diagonal: Parameter names
#
# Adapted from ASPS analysis script (git_ANOVA_ASPS_fractal_and_texture.r)

if (create_correlation_plots) {

  cat("\n")
  cat("################################################################################\n")
  cat("CREATING CORRELATION MATRIX PLOTS\n")
  cat("################################################################################\n\n")

  # --- Parameter order for correlation matrix display ---
  # Custom order groups conceptually related parameters together
  if (analysis_type == "texture") {
    corr_param_order <- c(
      "cluster_shade",             # Cluster-based features
      "diagonal_moment",
      "maximum_probability",       # Probability-based features
      "sum_energy",
      "cluster_prominence",        # Additional cluster feature
      "entropy",                   # Entropy features
      "kappa",
      "energy",                    # Energy features
      "correlation",
      "difference_energy",
      "difference_entropy",        # Difference features
      "inertia",                   # Moment-based features
      "inverse_different_moment",
      "sum_entropy",               # Sum features
      "sum_variance",
      "evaporation_duration"       # Evaporation duration
    )
    corr_plot_size <- 40  # cm — large for 16x16 grid
  } else {
    corr_param_order <- c(
      "db_mean", "dm_mean", "dx_mean",
      "evaporation_duration"
    )
    corr_plot_size <- 20  # cm — smaller for 4x4 grid
  }

  cat(sprintf("Parameters for correlation matrix (%d total):\n", length(corr_param_order)))
  cat(paste(corr_param_order, collapse = ", "), "\n\n")


  # --- Helper functions for ggpairs panels ---

  # upper_cor: Pearson r with green color gradient (light = weak, dark = strong)
  upper_cor <- function(data, mapping, ...) {
    x <- eval_data_col(data, mapping$x)
    y <- eval_data_col(data, mapping$y)
    cor_val <- cor(x, y, use = "complete.obs")

    # Green gradient: light green (#90EE90) to dark green (#006400)
    abs_cor <- abs(cor_val)
    red_component   <- round(144 * (1 - abs_cor))
    green_component <- round(238 - 138 * abs_cor)
    blue_component  <- round(144 * (1 - abs_cor))
    color_hex <- sprintf("#%02X%02X%02X", red_component, green_component, blue_component)

    ggplot(data, mapping) +
      annotate("text", x = 0.5, y = 0.5,
               label = sprintf("%.2f", cor_val),
               color = color_hex, size = 6) +
      xlim(0, 1) + ylim(0, 1) +
      theme_void()
  }

  # lower_scatter: Scatterplots colored by experimenter (AS = blue, JZ = red)
  lower_scatter <- function(data, mapping, ...) {
    ggplot(data, mapping) +
      geom_point(aes(color = experiment_name), alpha = 0.2, size = 0.5) +
      scale_color_manual(values = c("ASM_AS" = "blue", "ASM_JZ" = "red"),
                         guide = "none") +
      theme_minimal() +
      theme(
        panel.grid = element_blank(),
        axis.text = element_blank(),
        axis.ticks = element_blank()
      )
  }

  # diag_label: Parameter name on diagonal (underscores replaced with newlines)
  diag_label <- function(data, mapping, ...) {
    var_name <- as.character(mapping$x)[2]
    display_name <- gsub("_", "\n", var_name)

    ggplot(data, mapping) +
      annotate("text", x = 0.5, y = 0.5,
               label = display_name,
               size = 2.5, fontface = "bold", lineheight = 0.8) +
      xlim(0, 1) + ylim(0, 1) +
      theme_void()
  }


  # --- Create correlation output subfolder ---
  corr_folder <- file.path(output_folder, "correlation_matrices")
  if (!dir.exists(corr_folder)) {
    dir.create(corr_folder, recursive = TRUE)
    cat(sprintf("Created subfolder: %s\n", corr_folder))
  }


  # --- Define data subsets for the 5 correlation plots ---
  corr_subsets <- list(
    list(name = "ALL",   label = "All Data",    filter_expr = rep(TRUE, nrow(df2))),
    list(name = "AS",    label = "AS Only",      filter_expr = df2$experiment_name == "ASM_AS"),
    list(name = "JZ",    label = "JZ Only",      filter_expr = df2$experiment_name == "ASM_JZ"),
    list(name = "SNC",   label = "SNC Only (JZ)", filter_expr = df2$verum == 0 & df2$experiment_name == "ASM_JZ"),
    list(name = "Verum", label = "Verum Only (JZ)", filter_expr = df2$verum == 1 & df2$experiment_name == "ASM_JZ")
  )


  # --- Generate and save each correlation plot ---
  for (subset_info in corr_subsets) {

    subset_data <- df2[subset_info$filter_expr, ]

    # Print experiment numbers for verification
    exp_nums <- sort(unique(as.numeric(as.character(subset_data$experiment_number))))
    cat(sprintf("\nCorrelation plot %s: experiments %s (n = %d observations)\n",
                subset_info$name,
                paste(exp_nums, collapse = ", "),
                nrow(subset_data)))

    # Select correlation parameters + experiment_name (for scatter coloring)
    df_ordered <- subset_data[, c(corr_param_order, "experiment_name")]

    # Build the ggpairs correlation matrix (columns arg excludes experiment_name)
    corr_plot <- ggpairs(
      df_ordered,
      columns = seq_along(corr_param_order),
      upper = list(continuous = upper_cor),
      lower = list(continuous = lower_scatter),
      diag  = list(continuous = diag_label)
    ) +
      theme(
        plot.title = element_text(size = 14, face = "bold", hjust = 0.5),
        strip.text = element_blank(),
        panel.spacing = unit(0.1, "lines")
      ) +
      labs(title = sprintf("Correlation Matrix: %s (%s Parameters)",
                           subset_info$label, tools::toTitleCase(analysis_type)))

    # Save to subfolder
    output_file_corr <- file.path(
      corr_folder,
      paste0(date2, "_", analysis_type, "_correlation_", subset_info$name,
             outlier_file_suffix, ".png")
    )

    ggsave(filename = output_file_corr,
           plot = corr_plot,
           width = corr_plot_size, height = corr_plot_size,
           dpi = 300, units = "cm")

    cat(sprintf("  Saved: %s\n", basename(output_file_corr)))
  }

  cat("\nCorrelation matrix plots completed successfully\n")
  cat("================================================================================\n\n")

} else {
  cat("\nCorrelation matrix plots: SKIPPED (create_correlation_plots = FALSE)\n\n")
}


#===== PAIR-POOLING DIAGNOSTIC ================================================

# Tests whether the two independently prepared pairs within each solution
# differ systematically. Uses a nested fixed-effect ANOVA:
#   param ~ decoded_solution / pair + experiment_number + decoded_solution:experiment_number
# The nested term (decoded_solution:pair) captures pair-level shifts within
# each solution. Per-solution contrasts (pair 1 vs pair 2 within each solution)
# are obtained via emmeans.
#
# Non-significant results justify pooling pairs in the main ANOVA.
# Significant results suggest escalating to a mixed model (lmer with random
# pair intercept) for formal reporting.
# Added 2026-04-09.

if (run_pair_diagnostic) {

  cat("\n")
  cat("========================================================================\n")
  cat("PAIR-POOLING DIAGNOSTIC\n")
  cat("========================================================================\n\n")

  # --- Helper function: nested ANOVA with per-solution pair contrasts ---
  # Returns a named list with omnibus p-value for the nested pair term,
  # plus individual p-values for pair 1 vs pair 2 within each solution.
  run_pair_diagnostic_anova <- function(data_subset, param_name) {

    if (!(param_name %in% colnames(data_subset))) {
      return(list(omnibus_p = NA, per_solution = setNames(rep(NA, 3), rep("NA", 3))))
    }

    n_experiments <- length(unique(data_subset$experiment_number))
    solutions <- levels(droplevels(data_subset$decoded_solution))

    # Initialize per-solution results
    per_solution <- setNames(rep(NA_real_, length(solutions)), solutions)

    tryCatch({
      # Build nested model: decoded_solution / pair expands to
      # decoded_solution + decoded_solution:pair
      if (n_experiments > 1) {
        formula_str <- paste(param_name,
          "~ decoded_solution / pair + experiment_number + decoded_solution:experiment_number")
      } else {
        formula_str <- paste(param_name, "~ decoded_solution / pair")
      }

      model <- aov(as.formula(formula_str), data = data_subset)
      anova_results <- Anova(model, type = "III")

      # Omnibus test for the nested pair term
      omnibus_p <- anova_results["decoded_solution:pair", "Pr(>F)"]

      # Per-solution pair contrasts via emmeans:
      # Compare pair 1 vs pair 2 separately within each decoded_solution level
      emm_pairs <- emmeans(model, pairwise ~ pair | decoded_solution)
      contrasts_summary <- summary(emm_pairs$contrasts)

      for (sol in solutions) {
        row <- contrasts_summary[contrasts_summary$decoded_solution == sol, ]
        if (nrow(row) == 1) {
          per_solution[sol] <- row$p.value
        }
      }

      return(list(omnibus_p = omnibus_p, per_solution = per_solution))

    }, error = function(e) {
      cat(sprintf("    WARNING: Model failed for %s — %s\n", param_name, e$message))
      return(list(omnibus_p = NA, per_solution = per_solution))
    })
  }

  # --- Scenario definitions (same 6 as main ANOVA, with droplevels) ---
  scenarios_pair <- list(
    list(name = "ALL DATA - SNC", short_name = "ALL_SNC",
         data = droplevels(df2[df2$verum == 0, ])),
    list(name = "ALL DATA - Verum", short_name = "ALL_Verum",
         data = droplevels(df2[df2$verum == 1, ])),
    list(name = "JZ ONLY - SNC", short_name = "JZ_SNC",
         data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 0, ])),
    list(name = "JZ ONLY - Verum", short_name = "JZ_Verum",
         data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 1, ])),
    list(name = "AS ONLY - SNC", short_name = "AS_SNC",
         data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 0, ])),
    list(name = "AS ONLY - Verum", short_name = "AS_Verum",
         data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 1, ]))
  )

  # --- Workbook setup (reuse existing color styles) ---
  wb_pair <- createWorkbook()

  # Collect per-scenario verdicts for the final console summary
  pair_verdicts <- list()

  for (scenario in scenarios_pair) {

    cat(sprintf("Processing pair diagnostic: %s\n", scenario$name))
    addWorksheet(wb_pair, scenario$short_name)

    # Identify solution names in this scenario's data
    solutions_in_data <- levels(scenario$data$decoded_solution)

    # Column names: Parameter + omnibus + one column per solution
    pair_col_names <- c("Parameter", "pair(omnibus)",
                        paste0("pair(", solutions_in_data, ")"))

    results_list <- list()

    for (param in all_params) {
      if (!(param %in% colnames(scenario$data))) next

      cat(sprintf("  Parameter: %s\n", param))
      result <- run_pair_diagnostic_anova(scenario$data, param)

      # Build one-row data.frame for this parameter
      row_vals <- c(result$omnibus_p)
      for (sol in solutions_in_data) {
        row_vals <- c(row_vals, result$per_solution[sol])
      }
      row_df <- data.frame(
        Parameter = param,
        t(row_vals),
        stringsAsFactors = FALSE
      )
      colnames(row_df) <- pair_col_names
      results_list[[param]] <- row_df
    }

    results_df <- do.call(rbind, results_list)
    rownames(results_df) <- NULL

    # Format numeric p-values as strings for Excel display
    for (col in pair_col_names[-1]) {
      for (row in 1:nrow(results_df)) {
        cell_value <- results_df[row, col]
        if (!is.na(suppressWarnings(as.numeric(cell_value)))) {
          results_df[row, col] <- sprintf("%.6f", as.numeric(cell_value))
        }
      }
    }

    # --- Write to Excel sheet ---
    writeData(wb_pair, scenario$short_name,
              paste0("Pair-Pooling Diagnostic: ", scenario$name,
                     " (", toupper(analysis_type), " analysis)"),
              startRow = 1, startCol = 1)
    writeData(wb_pair, scenario$short_name, data_source_note,
              startRow = 2, startCol = 1)

    writeData(wb_pair, scenario$short_name, results_df,
              startRow = 4, rowNames = FALSE)

    # Header style
    addStyle(wb_pair, scenario$short_name, style_header,
             rows = 4, cols = 1:ncol(results_df), gridExpand = TRUE)

    # Color-code p-values (columns 2 onwards)
    for (row_idx in 1:nrow(results_df)) {
      for (col_idx in 2:ncol(results_df)) {
        cell_value <- results_df[row_idx, colnames(results_df)[col_idx]]

        if (is.na(cell_value) || grepl("Error|NA", cell_value)) next

        p_val <- suppressWarnings(as.numeric(cell_value))
        if (!is.na(p_val)) {
          excel_row <- row_idx + 4  # +4 because data starts at row 5
          if (p_val < 0.01) {
            addStyle(wb_pair, scenario$short_name, style_red,
                     rows = excel_row, cols = col_idx)
          } else if (p_val < 0.05) {
            addStyle(wb_pair, scenario$short_name, style_orange,
                     rows = excel_row, cols = col_idx)
          } else if (p_val < 0.10) {
            addStyle(wb_pair, scenario$short_name, style_lilac,
                     rows = excel_row, cols = col_idx)
          }
        }
      }
    }

    # Column widths
    col_widths <- c(25, rep(18, ncol(results_df) - 1))
    setColWidths(wb_pair, scenario$short_name,
                 cols = 1:ncol(results_df), widths = col_widths)

    # Legend
    legend_row <- nrow(results_df) + 6
    writeData(wb_pair, scenario$short_name, "Color Legend:",
              startRow = legend_row, startCol = 1)
    writeData(wb_pair, scenario$short_name, "Light lilac = p < 0.10",
              startRow = legend_row + 1, startCol = 1)
    writeData(wb_pair, scenario$short_name, "Light orange = p < 0.05",
              startRow = legend_row + 2, startCol = 1)
    writeData(wb_pair, scenario$short_name, "Light red = p < 0.01",
              startRow = legend_row + 3, startCol = 1)
    writeData(wb_pair, scenario$short_name,
              "Omnibus = nested pair term from Type III ANOVA; per-solution = emmeans pair contrasts",
              startRow = legend_row + 4, startCol = 1)

    # --- Console summary for this scenario ---
    # Count parameters with omnibus pair effect at p < 0.05
    omnibus_col <- "pair(omnibus)"
    omnibus_pvals <- suppressWarnings(as.numeric(results_df[[omnibus_col]]))
    sig_params <- results_df$Parameter[!is.na(omnibus_pvals) & omnibus_pvals < 0.05]
    n_sig <- length(sig_params)
    n_total <- sum(!is.na(omnibus_pvals))

    if (n_sig == 0) {
      verdict <- "Pooling appears justified (no pair effects at p < 0.05)"
    } else {
      verdict <- sprintf("Pair effects detected at p < 0.05 in %d/%d parameters — consider mixed model",
                         n_sig, n_total)
    }

    cat(sprintf("  >> %s: %s\n", scenario$short_name, verdict))
    if (n_sig > 0) {
      cat(sprintf("     Affected: %s\n", paste(sig_params, collapse = ", ")))
    }
    pair_verdicts[[scenario$short_name]] <- verdict
  }

  # --- Save workbook ---
  output_filename_pair <- file.path(output_folder,
    paste0(input_prefix, output_prefix, "pair_diagnostic", outlier_file_suffix, ".xlsx"))
  saveWorkbook(wb_pair, output_filename_pair, overwrite = TRUE)

  cat("\n")
  cat("========================================================================\n")
  cat(sprintf("Pair diagnostic exported to: %s\n", output_filename_pair))
  cat("========================================================================\n\n")

  # --- Pair agreement scatter plots ---
  # For each parameter, plot mean(pair 1) vs mean(pair 2) per solution × experiment.
  # Points near the identity line indicate that the two independently prepared
  # pairs agree well, visually supporting the pooling assumption.

  cat("--- Creating pair agreement scatter plots ---\n")

  # Colors: Dark2 positions 4-6 (pink, green, yellow) — one per solution
  dark2_pal <- RColorBrewer::brewer.pal(8, "Dark2")
  pair_scatter_colors <- dark2_pal[4:6]

  # Filter to parameters that exist in the data
  scatter_vars <- analysis_vars[analysis_vars %in% colnames(df2)]

  if (length(scatter_vars) > 0) {

    # Compute pair-level means: one value per (verum, experiment_number,
    # decoded_solution, pair) — averaging over the ~7 replicates within each pair
    pair_means <- df2 %>%
      group_by(verum, experiment_number, decoded_solution, pair) %>%
      summarise(
        across(all_of(scatter_vars), ~mean(.x, na.rm = TRUE)),
        .groups = "drop"
      )

    # Pivot to wide: separate columns for pair 1 and pair 2 means
    pair1_data <- pair_means %>%
      filter(pair == 1) %>%
      rename_with(~paste0(.x, "_p1"), all_of(scatter_vars)) %>%
      select(-pair)

    pair2_data <- pair_means %>%
      filter(pair == 2) %>%
      rename_with(~paste0(.x, "_p2"), all_of(scatter_vars)) %>%
      select(-pair)

    scatter_data <- inner_join(pair1_data, pair2_data,
                               by = c("verum", "experiment_number", "decoded_solution"))

    # Color mappings for SNC and Verum (same order as base_solution sol1/sol2/sol3)
    snc_solutions <- c("water_1", "water_2", "water_3")
    verum_solutions <- c("Lactose", "Stannum", "Silicea")
    color_map_snc_scatter <- setNames(pair_scatter_colors, snc_solutions)
    color_map_verum_scatter <- setNames(pair_scatter_colors, verum_solutions)

    # Experiment shapes: use distinct shapes for up to 5 sub-experiments
    all_exp_numbers <- sort(unique(as.numeric(as.character(scatter_data$experiment_number))))
    exp_shapes <- c(16, 17, 15, 18, 8, 3, 4, 7, 9, 10)  # circle, triangle, square, diamond, ...
    shape_map <- setNames(exp_shapes[seq_along(all_exp_numbers)],
                          as.character(all_exp_numbers))

    # Conditions: SNC (verum == 0) and Verum (verum == 1)
    scatter_conditions <- c(0, 1)
    scatter_condition_labels <- c("SNC (Water Controls)", "Verum (Treatment)")

    scatter_plot_list <- list()
    scatter_counter <- 1

    for (var in scatter_vars) {
      p1_col <- paste0(var, "_p1")
      p2_col <- paste0(var, "_p2")

      for (cond_idx in seq_along(scatter_conditions)) {
        verum_val <- scatter_conditions[cond_idx]
        cond_label <- scatter_condition_labels[cond_idx]

        plot_df <- scatter_data %>% filter(verum == verum_val)
        color_map <- if (verum_val == 0) color_map_snc_scatter else color_map_verum_scatter

        # Shared axis range so the identity line is a true diagonal
        all_vals <- c(plot_df[[p1_col]], plot_df[[p2_col]])
        axis_range <- range(all_vals, na.rm = TRUE)
        axis_pad <- diff(axis_range) * 0.05
        axis_lims <- c(axis_range[1] - axis_pad, axis_range[2] + axis_pad)

        p <- ggplot(plot_df, aes(x = .data[[p1_col]],
                                 y = .data[[p2_col]],
                                 color = decoded_solution,
                                 shape = experiment_number)) +
          # Identity line: points on this line mean pair 1 == pair 2
          geom_abline(intercept = 0, slope = 1,
                      linetype = "dashed", color = "grey50", linewidth = 0.4) +
          geom_point(size = 3, alpha = 0.85) +
          scale_color_manual(values = color_map, name = "Solution") +
          scale_shape_manual(values = shape_map, name = "Experiment") +
          coord_fixed(xlim = axis_lims, ylim = axis_lims) +
          labs(
            title = cond_label,
            x = paste0(var, " (pair 1 mean)"),
            y = paste0(var, " (pair 2 mean)")
          ) +
          theme_bw() +
          theme(
            plot.title = element_text(size = 11, face = "bold"),
            axis.title = element_text(size = 9),
            axis.text = element_text(size = 8),
            legend.title = element_text(size = 9),
            legend.text = element_text(size = 8)
          )

        scatter_plot_list[[scatter_counter]] <- p
        scatter_counter <- scatter_counter + 1
      }
    }

    # Arrange in grid: one row per parameter, two columns (SNC | Verum)
    scatter_nrow <- length(scatter_vars)
    scatter_height <- 8 * scatter_nrow  # cm per row
    if (scatter_height < 20) scatter_height <- 20

    grid_scatter <- grid.arrange(
      grobs = scatter_plot_list,
      nrow = scatter_nrow,
      ncol = 2,
      top = paste0("ASM ", toupper(analysis_type),
                    " — Pair Agreement (pair 1 mean vs pair 2 mean)\n",
                    data_source_note)
    )

    output_filename_pair_scatter <- file.path(output_folder,
      paste0(input_prefix, output_prefix, "pair_scatter", outlier_file_suffix, ".png"))

    ggsave(
      filename = output_filename_pair_scatter,
      plot = grid_scatter,
      width = 30,
      height = scatter_height,
      dpi = 300,
      units = "cm",
      limitsize = FALSE
    )

    cat(sprintf("Pair scatter plot saved as: %s\n", output_filename_pair_scatter))

  } else {
    cat("WARNING: No analysis_vars found in data. Skipping pair scatter plots.\n")
  }

} else {
  cat("\nPair-pooling diagnostic: SKIPPED (run_pair_diagnostic = FALSE)\n\n")
}


#===== FINAL SUMMARY ==========================================================

cat("\n")
cat("################################################################################\n")
cat("ANALYSIS COMPLETE\n")
cat("################################################################################\n\n")
cat(sprintf("Analysis type: %s\n", toupper(analysis_type)))
cat(sprintf("Input file: %s\n", input_file))
cat("\nOutput files:\n")
cat(sprintf("  1. ANOVA summary: %s\n", output_filename_anova))
cat(sprintf("  2. Post-hoc & Cohen's d: %s\n", output_filename_posthoc))
if (exists("output_filename_plot")) {
  cat(sprintf("  3. Parameter plots: %s\n", output_filename_plot))
}
if (create_correlation_plots) {
  cat(sprintf("  4. Correlation matrices: %s/\n", corr_folder))
}
if (run_pair_diagnostic) {
  cat(sprintf("  5. Pair-pooling diagnostic: %s\n", output_filename_pair))
  if (exists("output_filename_pair_scatter")) {
    cat(sprintf("  6. Pair scatter plots: %s\n", output_filename_pair_scatter))
  }
}
cat("\nDone!\n")
