# ASM Unified ANOVA and Post-Hoc Analysis Script
# Created: 2026-01-14
# Purpose: Apply ANOVA and Cohen's d analysis on either texture or fractal data
# Toggle between analysis types using the configuration section below
# Run 260114_ASM_TA_FA_data_decoding.r first to have decoded files

# Load required packages -------
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


# ============================================================================
# CONFIGURATION - USER TOGGLES
# ============================================================================

# USER TOGGLE 1: Analysis type - "texture" or "fractal"
analysis_type <- "fractal"

# USER TOGGLE 2: Input files (pre-decoded from 260114_ASM_texture_data_decoding.r)
input_file_texture <- "20250910_ASM_1_10_texture_data_DECODED.csv"
input_file_fractal <- "20251218_ASM_1_10_fractal_data_DECODED.csv"

# USER TOGGLE 3: Outlier removal - TRUE or FALSE
remove_outliers <- TRUE
outliers_file_name <- "20260108_ASM_outliers.csv"

# ============================================================================
# DERIVED CONFIGURATION (based on toggles above)
# ============================================================================

# Select input file based on analysis type
if (analysis_type == "texture") {
  input_file <- input_file_texture
} else {
  input_file <- input_file_fractal
}

# Create file prefix from input filename (extract date portion)
# e.g., "20250910_ASM_1_10_texture_data_DECODED.csv" -> "20250910"
input_prefix <- sub("_.*", "", basename(input_file))

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

  analysis_vars <- c("cluster_shade", "entropy", "maximum_probability", "kappa")

  # All 15 texture parameters for comprehensive Excel ANOVA summary
  all_params <- c(
    "cluster_prominence", "cluster_shade", "correlation",
    "diagonal_moment", "difference_energy", "difference_entropy",
    "energy", "entropy", "inertia", "inverse_different_moment",
    "kappa", "maximum_probability", "sum_energy",
    "sum_entropy", "sum_variance"
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


# Determine export date and create output folder
date <- Sys.Date()
date2 <- gsub("-| |UTC", "", date)

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


# ============================================================================
# DATA LOADING
# ============================================================================

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

# ============================================================================
# OUTLIER REMOVAL
# ============================================================================

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


# ============================================================================
# PREP FOR ANOVA
# ============================================================================

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


# ============================================================================
# ANOVA HELPER FUNCTION
# ============================================================================

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


# ============================================================================
# ANOVA SUMMARY FOR ALL PARAMETERS
# ============================================================================

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


# ============================================================================
# POST-HOC TESTS WITH COHEN'S D
# ============================================================================

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

  # Initialize result matrices
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

  # Compute main effect pairwise comparisons
  emm_solution <- emmeans(model, ~ decoded_solution)
  pairs_solution <- pairs(emm_solution, adjust = "none")
  pairs_summary <- summary(pairs_solution)

  # Fill potency matrix with p-values
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

  # Cohen's d calculation
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

  # Interaction contrasts (only if multiple experiments)
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

    # Data rows
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

    # Data rows
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


# ============================================================================
# PLOTTING
# ============================================================================

cat("========================================================================\n")
cat("CREATING PLOTS\n")
cat("========================================================================\n\n")

# Define conditions
conditions <- c(0, 1)
condition_labels <- c("SNC (Water Controls)", "Verum (Treatment)")

# Create solution-pair combination variable
df2 <- df2 %>%
  mutate(
    solution_pair = case_when(
      decoded_solution == "water_1" & pair == 1 ~ "W1-1",
      decoded_solution == "water_1" & pair == 2 ~ "W1-2",
      decoded_solution == "water_2" & pair == 1 ~ "W2-1",
      decoded_solution == "water_2" & pair == 2 ~ "W2-2",
      decoded_solution == "water_3" & pair == 1 ~ "W3-1",
      decoded_solution == "water_3" & pair == 2 ~ "W3-2",
      decoded_solution == "potency_x" & pair == 1 ~ "X1",
      decoded_solution == "potency_x" & pair == 2 ~ "X2",
      decoded_solution == "potency_y" & pair == 1 ~ "Y1",
      decoded_solution == "potency_y" & pair == 2 ~ "Y2",
      decoded_solution == "potency_z" & pair == 1 ~ "Z1",
      decoded_solution == "potency_z" & pair == 2 ~ "Z2",
      TRUE ~ NA_character_
    ),
    base_solution = case_when(
      decoded_solution == "water_1" ~ "sol1",
      decoded_solution == "water_2" ~ "sol2",
      decoded_solution == "water_3" ~ "sol3",
      decoded_solution == "potency_x" ~ "sol1",
      decoded_solution == "potency_y" ~ "sol2",
      decoded_solution == "potency_z" ~ "sol3",
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

  # Define color palette
  dark2_colors <- RColorBrewer::brewer.pal(8, "Dark2")
  base_colors <- dark2_colors[4:6]
  darkened_colors <- darken(base_colors, amount = 0.3)

  color_mapping_snc <- c(
    "W1-1" = base_colors[1], "W1-2" = darkened_colors[1],
    "W2-1" = base_colors[2], "W2-2" = darkened_colors[2],
    "W3-1" = base_colors[3], "W3-2" = darkened_colors[3]
  )

  color_mapping_verum <- c(
    "X1" = base_colors[1], "X2" = darkened_colors[1],
    "Y1" = base_colors[2], "Y2" = darkened_colors[2],
    "Z1" = base_colors[3], "Z2" = darkened_colors[3]
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
          plot.caption = element_text(size = 8, hjust = 0)
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
    width = 30,
    height = actual_height,
    dpi = 300,
    units = "cm"
  )

  cat(sprintf("Plot saved as: %s\n", output_filename_plot))
}


# ============================================================================
# FINAL SUMMARY
# ============================================================================

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
cat("\nDone!\n")
