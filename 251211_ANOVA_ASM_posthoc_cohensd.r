# 2025.10.27 - Created sensitivity analysis script with solution decoding
# Based on: 251002_ANOVA_2_variables_ASM_filtering_and_sensitivity.r
# Original: 2023.06.06
# Remade by: Paul 2025.09.30
# Updated by: Anezka for ASM data
# Updated 2025.11.03 - Added emmeans post-hoc analysis with Cohen's d effect sizes
# Updated 25.12.11 with graphs of all TA variables

# Load required packages -------
library(openxlsx)  # For Excel file handling
library(readxl)    # For reading Excel files
library(car)       # For Type III ANOVA (Anova function)
library(here)      # For file path management
library(dplyr)     # For data manipulation


# Determine export date

date <- Sys.Date() 
date2 <- gsub("-| |UTC", "", date)


# Data files with Texture data from ACIA + DECODING --------

df <- read.delim(here("ASM1-10-reduced-data copy.csv"))


# Hard coding on ASM data
# Filter df for scale and separate out experiment number and potency codes
# AS and JZ named the series differently
# Using column "substance_group" for name extraction and "series_name" for experiment number

# Filter to keep only scale = 1 data
# The square brackets [rows, columns] allow us to subset the dataframe
# We're selecting rows where scale equals 1, and keeping all columns
df_scale1 <- df[df$scale == 1, ]

# Define the mapping for AS series
# The name "EI" maps to value 2, "HOM3_ASM1" maps to 1
# Corrected from Exp 1+2 changed up, 2025-11-10
series_to_number <- c(
  "EI" = 2,
  "HOM3_ASM1" = 1
)

# Create a function to parse substance_group
# This function extracts experiment number and potency letter from the substance_group text
# It handles two different naming patterns (AS and JZ)
parse_substance_group <- function(substance_group, series_name) {
  
  # Check if it's the JZ pattern (starts with "HOM3 MEX")
  # grepl() returns TRUE if the pattern is found
  if (grepl("^HOM3 MEX", substance_group)) {
    # JZ pattern: "HOM3 MEX3-A"
    # Extract number after MEX and letter after dash
    # regexec finds the pattern and returns the positions
    # regmatches extracts the actual text at those positions
    matches <- regmatches(substance_group, 
                          regexec("HOM3 MEX([0-9]+)-([A-F])", substance_group))[[1]]
    
    # Check if we found both the number and the letter
    # matches[1] is the full match, matches[2] is the number, matches[3] is the letter
    if (length(matches) == 3) {
      return(list(
        experiment_name = "ASM_JZ",
        experiment_number = as.integer(matches[2]),  # Convert text to integer
        potency = matches[3]                         # The letter (A-F)
      ))
    }
  } else {
    # AS pattern: "Cress A4 potA"
    # Extract everything after "pot"
    # regexpr finds the position of the pattern
    potency_match <- regmatches(substance_group, 
                                regexpr("pot[A-Z]+$", substance_group))
    
    if (length(potency_match) > 0) {
      # Remove "pot" prefix to get just the letter(s)
      # sub() replaces the first occurrence of a pattern
      potency <- sub("^pot", "", potency_match)
    } else {
      # Fallback if nothing is extracted
      potency <- "XX"
    }
    
    return(list(
      experiment_name = "ASM_AS",
      experiment_number = series_to_number[series_name],  # Look up experiment number using series_name
      potency = potency
    ))
  }
}

# Apply new columns to dataframe
# The %>% operator (pipe) passes the result from one line to the next
# This makes code more readable by avoiding nested functions
df2 <- df_scale1 %>%
  rowwise() %>%  # Apply the function to each row individually
  mutate(
    # Call our parsing function and store the result as a list
    parsed = list(parse_substance_group(substance_group, series_name)),
    # Extract individual elements from the parsed list into separate columns
    experiment_name = parsed$experiment_name,
    experiment_number = parsed$experiment_number,
    potency = parsed$potency
  ) %>%
  select(-parsed) %>%  # Remove the temporary parsed column
  ungroup()            # Remove rowwise grouping (important for performance)

# Add verum and SNC columns
# ifelse() is a vectorized if-else: ifelse(condition, value_if_true, value_if_false)
# %in% checks if a value is in a vector
df2 <- df2 %>%
  mutate(
    # verum = 1 for treatment experiments (2, 4, 5, 7, 10)
    verum = ifelse(experiment_number %in% c(2, 4, 5, 7, 10), 1, 0),
    # SNC = 1 for systematic negative control experiments (1, 3, 6, 8, 9)
    SNC = ifelse(experiment_number %in% c(1, 3, 6, 8, 9), 1, 0)
  )

# Create sub_exp_number (1-5 within each verum/SNC group)
# This renumbers experiments within each condition to 1-5
# Useful when you want to treat experiments as levels within a condition
# case_when() is like a multi-way if-else statement
# Verum experiments: 2, 4, 5, 7, 10 become 1, 2, 3, 4, 5
# SNC experiments: 1, 3, 6, 8, 9 become 1, 2, 3, 4, 5
df2 <- df2 %>%
  mutate(
    sub_exp_number = case_when(
      experiment_number == 2 ~ 1,   # Verum 1
      experiment_number == 4 ~ 2,   # Verum 2
      experiment_number == 5 ~ 3,   # Verum 3
      experiment_number == 7 ~ 4,   # Verum 4
      experiment_number == 10 ~ 5,  # Verum 5
      experiment_number == 1 ~ 1,   # SNC 1
      experiment_number == 3 ~ 2,   # SNC 2
      experiment_number == 6 ~ 3,   # SNC 3
      experiment_number == 8 ~ 4,   # SNC 4
      experiment_number == 9 ~ 5,   # SNC 5
      TRUE ~ NA_real_              # If none match, return NA
    )
  )


# DECODE PAIRED SOLUTIONS
# 
# Background:
# - Letters A-F represent groups to be compared (coded solutions)
# - Each experiment has 6 groups with 7 crystallization plates
# - 3 different solutions (3 potencies or 3 SNCs) are tested, each produced in duplicate (pairs) 
#
# SNC experiments: 3 water preparations (water_1, water_2, water_3) each produced twice
# Verum experiments: 3 homeopathic solutions (potency_x, potency_y, potency_z) each produced twice
# 
# The decoding tables below map: experiment_number + coded solution (A-F) → solution + pair (1 or 2)

cat("\n")
cat("========================================================================\n")
cat("SOLUTION DECODING TABLES\n")
cat("========================================================================\n\n")

cat("--- SNC Mapping (Water Controls) ---\n")
cat("Experiment | Potency Letters | Decoded Solution | Pair\n")
cat("-----------|-----------------|------------------|-----\n")

# Create SNC mapping table
# This data.frame explicitly maps each combination of experiment + coded solution letter
# to its corresponding solution and pair number
snc_mapping <- data.frame(
  # List all experiment numbers (each appears 6 times, once for each solution A-F)
  experiment_number = c(
    # Day 1: A+D (water_1), B+E (water_2), C+F (water_3)
    1, 1, 1, 1, 1, 1,
    # Day 3: F+A (water_1), D+B (water_2), E+C (water_3)
    3, 3, 3, 3, 3, 3,
    # Day 6: C+F (water_1), A+E (water_2), B+D (water_3)
    6, 6, 6, 6, 6, 6,
    # Day 8: E+C (water_1), F+A (water_2), D+B (water_3)
    8, 8, 8, 8, 8, 8,
    # Day 9: A+E (water_1), C+F (water_2), B+D (water_3)
    9, 9, 9, 9, 9, 9
  ),
  # List the coded solution letters in the order they correspond to solutions
  # Order matches the comments above (first two letters = water_1, next two = water_2, last two = water_3)
  potency = c(
    "A", "D", "B", "E", "C", "F",  # Day 1
    "F", "A", "D", "B", "E", "C",  # Day 3
    "C", "F", "A", "E", "B", "D",  # Day 6
    "E", "C", "F", "A", "D", "B",  # Day 8
    "A", "E", "C", "F", "B", "D"   # Day 9
  ),
  # Assign the decoded solution name
  # Pattern repeats: water_1, water_1, water_2, water_2, water_3, water_3
  decoded_solution = c(
    "water_1", "water_1", "water_2", "water_2", "water_3", "water_3",
    "water_1", "water_1", "water_2", "water_2", "water_3", "water_3",
    "water_1", "water_1", "water_2", "water_2", "water_3", "water_3",
    "water_1", "water_1", "water_2", "water_2", "water_3", "water_3",
    "water_1", "water_1", "water_2", "water_2", "water_3", "water_3"
  ),
  # Assign pair numbers (1 for first solution, 2 for second solution of each pair. The pairs are separately produced from 29X)
  # Pattern repeats: 1, 2, 1, 2, 1, 2
  pair = c(
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2
  )
)

# Print SNC mapping for verification
for (exp in unique(snc_mapping$experiment_number)) {
  exp_data <- snc_mapping[snc_mapping$experiment_number == exp, ]
  for (sol in c("water_1", "water_2", "water_3")) {
    sol_data <- exp_data[exp_data$decoded_solution == sol, ]
    letters <- paste(sol_data$potency, collapse = "+")
    pairs <- paste(sol_data$pair, collapse = ",")
    cat(sprintf("Day %-2d     | %-15s | %-16s | %s\n", 
                exp, letters, sol, pairs))
  }
}

cat("\n--- Verum Mapping (Treatment Solutions) ---\n")
cat("Experiment | Potency Letters | Decoded Solution | Pair\n")
cat("-----------|-----------------|------------------|-----\n")

# Create Verum mapping table
# Same structure as SNC mapping, but for treatment experiments
# potency_x, potency_y, potency_z will later be decoded to silicea, lactose, stannum
verum_mapping <- data.frame(
  # List all experiment numbers (each appears 6 times, once for each coded solution A-F)
  experiment_number = c(
    # Day 2: A+D (potency_x), B+E (potency_y), C+F (potency_z)
    2, 2, 2, 2, 2, 2,
    # Day 4: F+A (potency_x), D+B (potency_y), E+C (potency_z)
    4, 4, 4, 4, 4, 4,
    # Day 5: C+F (potency_x), A+E (potency_y), B+D (potency_z)
    5, 5, 5, 5, 5, 5,
    # Day 7: E+C (potency_x), F+A (potency_y), D+B (potency_z)
    7, 7, 7, 7, 7, 7,
    # Day 10: A+E (potency_x), C+F (potency_y), B+D (potency_z)
    10, 10, 10, 10, 10, 10
  ),
  # List the coded solution letters
  potency = c(
    "A", "D", "B", "E", "C", "F",  # Day 2
    "F", "A", "D", "B", "E", "C",  # Day 4
    "C", "F", "A", "E", "B", "D",  # Day 5
    "E", "C", "F", "A", "D", "B",  # Day 7
    "A", "E", "C", "F", "B", "D"   # Day 10
  ),
  # Assign decoded solution names
  decoded_solution = c(
    "potency_x", "potency_x", "potency_y", "potency_y", "potency_z", "potency_z",
    "potency_x", "potency_x", "potency_y", "potency_y", "potency_z", "potency_z",
    "potency_x", "potency_x", "potency_y", "potency_y", "potency_z", "potency_z",
    "potency_x", "potency_x", "potency_y", "potency_y", "potency_z", "potency_z",
    "potency_x", "potency_x", "potency_y", "potency_y", "potency_z", "potency_z"
  ),
  # Assign pair numbers
  pair = c(
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2,
    1, 2, 1, 2, 1, 2
  )
)

# Print Verum mapping for verification
# Loop through each unique experiment number in the mapping
for (exp in unique(verum_mapping$experiment_number)) {
  # Filter to get rows for this experiment only
  exp_data <- verum_mapping[verum_mapping$experiment_number == exp, ]
  # Loop through each solution type
  for (sol in c("potency_x", "potency_y", "potency_z")) {
    # Filter to get rows for this solution only
    sol_data <- exp_data[exp_data$decoded_solution == sol, ]
    # Combine the coded solution letters with a + sign
    letters <- paste(sol_data$potency, collapse = "+")
    # Combine the pair numbers with a comma
    pairs <- paste(sol_data$pair, collapse = ",")
    # Print formatted output
    cat(sprintf("Day %-2d     | %-15s | %-16s | %s\n", 
                exp, letters, sol, pairs))
  }
}

cat("\n========================================================================\n\n")


# Apply decoding to data

# Combine mapping tables
# bind_rows() stacks data frames vertically (combines rows from both tables)
# Now we have one complete mapping for all experiments
full_mapping <- bind_rows(snc_mapping, verum_mapping)

# Join with df2 to add decoded columns
# left_join() matches rows based on specified columns (experiment_number and potency)
# It adds the decoded_solution and pair columns from full_mapping to df2
# "left" join means: keep all rows from df2, add matching info from full_mapping
df2 <- df2 %>%
  left_join(full_mapping, by = c("experiment_number", "potency"))

# Verify decoding was successful
cat("Decoding verification:\n")
cat("Total rows in df2:", nrow(df2), "\n")
cat("Rows with decoded_solution:", sum(!is.na(df2$decoded_solution)), "\n")
cat("Rows with pair assignment:", sum(!is.na(df2$pair)), "\n\n")

# Check if all rows were decoded
# is.na() checks for missing values (NA = Not Available)
# any() returns TRUE if at least one value is TRUE
if (any(is.na(df2$decoded_solution))) {
  cat("WARNING: Some rows were not decoded!\n")
  cat("Rows without decoding:\n")
  # Filter to show only rows that failed to decode
  # select() chooses which columns to display
  # distinct() removes duplicate rows
  print(df2 %>% 
          filter(is.na(decoded_solution)) %>% 
          select(experiment_number, potency, substance_group) %>%
          distinct())
  cat("\n")
}

# OUTLIER REMOVAL -------

# 2025.10.28 - Remove visually identified outliers

# # Read outlier file
# outliers_file <- here("ASM1_10_reduced_data_excl_W1-1_only.xlsx")
# outliers <- read_excel(outliers_file)

# MANUAL: No removal
outliers <- df[0, ]

# Check if outlier file has any entries
if (nrow(outliers) == 0) {
  cat("\n")
  cat("========================================================================\n")
  cat("OUTLIER REMOVAL\n")
  cat("========================================================================\n")
  cat("No outliers removed - outlier file is empty\n")
  cat("========================================================================\n\n")
  
} else {
  
  # Create summary table of outliers BEFORE removal (with decoded information)
  outliers_decoded <- df2 %>%
    filter(name %in% outliers$name) %>%
    select(name, experiment_number, decoded_solution, pair) %>%
    arrange(experiment_number, decoded_solution, pair)
  
  # Count outliers before removal
  n_outliers <- nrow(outliers_decoded)
  n_before <- nrow(df2)
  
  # Remove outliers from df2
  df2 <- df2 %>%
    filter(!(name %in% outliers$name))
  
  # Count after removal
  n_after <- nrow(df2)
  n_removed <- n_before - n_after
  
  # Print results to console
  cat("\n")
  cat("========================================================================\n")
  cat("OUTLIER REMOVAL\n")
  cat("========================================================================\n")
  cat(sprintf("Outliers identified in file: %d\n", nrow(outliers)))
  cat(sprintf("Outliers found and removed: %d\n", n_removed))
  cat(sprintf("Rows before removal: %d\n", n_before))
  cat(sprintf("Rows after removal: %d\n\n", n_after))
  
  cat("Removed outliers (decoded):\n")
  cat("--------------------------------------------------------------------------------\n")
  cat(sprintf("%-50s %5s %-20s %5s\n", "Name", "Day", "Solution", "Pair"))
  cat("--------------------------------------------------------------------------------\n")
  
  for (i in 1:nrow(outliers_decoded)) {
    cat(sprintf("%-50s %5d %-20s %5d\n",
                outliers_decoded$name[i],
                outliers_decoded$experiment_number[i],
                outliers_decoded$decoded_solution[i],
                outliers_decoded$pair[i]))
  }
  
  cat("========================================================================\n\n")
  
  # Warning if outliers in file were not found in data
  if (n_removed < nrow(outliers)) {
    cat("WARNING: Some outliers from file were not found in df2\n")
    cat(sprintf("  Expected to remove: %d\n", nrow(outliers)))
    cat(sprintf("  Actually removed: %d\n\n", n_removed))
  }
}



# PREP for ANOVA -------

# Convert categorical variables to factors
# 
# In R, factors are special data types for categorical variables
# ANOVA and other statistical tests require categorical predictors to be factors
# as.factor() converts a variable to factor type
df2$decoded_solution <- as.factor(df2$decoded_solution)
df2$experiment_number <- as.factor(df2$experiment_number)
df2$sub_exp_number <- as.factor(df2$sub_exp_number)
df2$verum <- as.factor(df2$verum)
df2$pair <- as.factor(df2$pair)

# Set contrast options for Type III ANOVA
# Contrasts determine how R handles categorical variables in linear models
# "contr.sum" = sum-to-zero contrasts (required for Type III SS)
# "contr.poly" = polynomial contrasts (not used in Type III, but specified for completeness)
options(contrasts = c("contr.sum", "contr.poly"))


# Define 4 key texture parameters for post-hoc analysis

analysis_vars <- c("cluster_shade", "entropy", "maximum_probability", "kappa")


# Data integrity check

cat("\n")
cat("========================================================================\n")
cat("DATA INTEGRITY CHECK\n")
cat("========================================================================\n")
cat("Total observations in df2:", nrow(df2), "\n")
cat("Expected: 10 experiments × 6 potencies × 7 replicates = 420 observations\n\n")

# Check counts by experiment_number
cat("--- Observations per experiment_number ---\n")
exp_counts <- table(df2$experiment_number)
print(exp_counts)
cat("\n")

# Check counts by experiment_number and decoded_solution
# This table shows which solutions were tested in each experiment
# SNC experiments (1,3,6,8,9) should only have water_1, water_2, water_3
# Verum experiments (2,4,5,7,10) should only have potency_x, potency_y, potency_z
cat("--- Observations per experiment_number × decoded_solution ---\n")
exp_sol_counts <- table(df2$experiment_number, df2$decoded_solution)
print(exp_sol_counts)
cat("\n")

cat("========================================================================\n\n")


# SENSITIVITY ANALYSIS ---------------------

# Test robustness across different data subsets
# This checks whether conclusions depend on including AS vs JZ data
#
# Purpose: Determine if results are stable regardless of which data subset we use
# If p-values are similar across scenarios, results are robust
# If p-values differ substantially, experimenter effects may be present

cat("\n")
cat("################################################################################\n")
cat("SENSITIVITY ANALYSIS - Using decoded solutions (pairs pooled)\n")
cat("################################################################################\n\n")

# Define three analysis scenarios
# Each scenario is a list with: name, filter_condition (function), and description
scenarios <- list(
  list(name = "ALL DATA (Experiments 1-10, both AS and JZ)",
       filter_condition = function(df) df,  # No filtering - return all data
       description = "Full dataset including both experimenters"),

  list(name = "JZ ONLY (Experiments 3-10)",
       filter_condition = function(df) df[df$experiment_name == "ASM_JZ", ],  # Keep only JZ rows
       description = "Only JZ data: 4 SNC + 4 Verum experiments"),

  list(name = "AS ONLY (Experiments 1-2)",
       filter_condition = function(df) df[df$experiment_name == "ASM_AS", ],  # Keep only AS rows
       description = "Only AS data: 1 SNC + 1 Verum experiment (limited power)")
)

# Storage for results comparison
# This list will hold p-values for each scenario
sensitivity_results <- list()

# Loop through each scenario
# seq_along() generates a sequence 1, 2, 3 for the three scenarios
for (scenario_idx in seq_along(scenarios)) {

  # Get the current scenario from the list
  scenario <- scenarios[[scenario_idx]]

  cat("================================================================================\n")
  cat("SCENARIO", scenario_idx, ":", scenario$name, "\n")
  cat(scenario$description, "\n")
  cat("================================================================================\n\n")

  # Apply filter to get subset
  # Call the filter_condition function defined in the scenario
  df_scenario <- scenario$filter_condition(df2)

  cat("Number of observations:", nrow(df_scenario), "\n")
  cat("Experiments included:", paste(sort(unique(df_scenario$experiment_number)), collapse = ", "), "\n\n")

  # Store scenario info
  # Create a nested list structure to hold results for this scenario
  sensitivity_results[[scenario_idx]] <- list(
    scenario_name = scenario$name,
    n_obs = nrow(df_scenario),
    cluster_shade = list(),  # Will hold p-values for cluster_shade
    entropy = list()         # Will hold p-values for entropy
  )

  # Run analysis separately for SNC and Verum within this scenario
  # Loop through verum_value: 0 = SNC, 1 = Verum
  for (verum_value in c(0, 1)) {

    # Create readable label for output
    condition_label <- ifelse(verum_value == 0, "SNC", "Verum")

    # Filter for current condition
    # Keep only rows where verum matches the current value
    df_subset <- df_scenario[df_scenario$verum == verum_value, ]

    cat("--------------------------------------------------------------------------------\n")
    cat(condition_label, "condition\n")
    cat("--------------------------------------------------------------------------------\n")
    cat("Observations:", nrow(df_subset), "\n")

    # Check if we have enough data to run analysis
    if (nrow(df_subset) < 10) {
      cat("⚠️  WARNING: Very small sample size - results may be unreliable\n\n")
    }

    # Check how many experiments we have in this subset
    # length(unique()) counts the number of distinct experiment numbers
    n_experiments <- length(unique(df_subset$experiment_number))

    if (n_experiments > 1) {
      # Run two-way ANOVA if we have multiple experiments
      # Formula: outcome ~ predictor1 * predictor2
      # The * includes main effects AND interaction
      # This tests: main effect of decoded_solution, main effect of experiment_number, and their interaction
      two_way_clu_sha <- aov(cluster_shade ~ decoded_solution * experiment_number, data = df_subset)
      two_way_entropy <- aov(entropy ~ decoded_solution * experiment_number, data = df_subset)

      # Generate ANOVA tables with Type III Sum of Squares
      # Type III tests each effect after controlling for all others
      # This is appropriate for unbalanced designs
      results_clu_sha <- Anova(two_way_clu_sha, type = "III")
      results_entropy <- Anova(two_way_entropy, type = "III")

      cat("\n--- Cluster Shade ---\n")
      print(results_clu_sha, digits = 6)

      # Store key p-value for decoded_solution main effect
      # Extract p-value from the "decoded_solution" row, "Pr(>F)" column
      # This p-value tells us if solutions differ significantly
      potency_p_clu_sha <- results_clu_sha["decoded_solution", "Pr(>F)"]
      sensitivity_results[[scenario_idx]]$cluster_shade[[condition_label]] <- potency_p_clu_sha

      cat("\n--- Entropy ---\n")
      print(results_entropy, digits = 6)

      potency_p_entropy <- results_entropy["decoded_solution", "Pr(>F)"]
      sensitivity_results[[scenario_idx]]$entropy[[condition_label]] <- potency_p_entropy

    } else {
      # Only one experiment - can only test decoded_solution main effect
      # Cannot test experiment effects or interactions with only one experiment
      cat("\nOnly 1 experiment in this subset - running one-way ANOVA (decoded_solution only)\n")

      # One-way ANOVA: outcome ~ single predictor
      # This tests whether the three solutions differ from each other
      one_way_clu_sha <- aov(cluster_shade ~ decoded_solution, data = df_subset)
      one_way_entropy <- aov(entropy ~ decoded_solution, data = df_subset)

      results_clu_sha <- Anova(one_way_clu_sha, type = "III")
      results_entropy <- Anova(one_way_entropy, type = "III")

      cat("\n--- Cluster Shade ---\n")
      print(results_clu_sha, digits = 6)

      potency_p_clu_sha <- results_clu_sha["decoded_solution", "Pr(>F)"]
      sensitivity_results[[scenario_idx]]$cluster_shade[[condition_label]] <- potency_p_clu_sha

      cat("\n--- Entropy ---\n")
      print(results_entropy, digits = 6)

      potency_p_entropy <- results_entropy["decoded_solution", "Pr(>F)"]
      sensitivity_results[[scenario_idx]]$entropy[[condition_label]] <- potency_p_entropy
    }

    cat("\n")
  }

  cat("\n\n")
}


# Summary comparison of sensitivity analysis
#
# This section displays all p-values in a compact table format
# Allows easy comparison across scenarios to assess robustness

cat("################################################################################\n")
cat("SENSITIVITY ANALYSIS SUMMARY: Solution effect p-values across scenarios\n")
cat("################################################################################\n\n")

cat("--- Cluster Shade ---\n")
# sprintf() formats strings with specified width and alignment
# %-45s = left-aligned string with 45 character width
# %15s = right-aligned string with 15 character width
cat(sprintf("%-45s %15s %15s\n", "Scenario", "SNC p-value", "Verum p-value"))
cat(strrep("-", 80), "\n")  # strrep() repeats a string n times
# Loop through results and print each scenario's p-values
for (i in seq_along(sensitivity_results)) {
  result <- sensitivity_results[[i]]
  # ifelse() handles cases where p-value might be missing (NULL)
  snc_p <- ifelse(is.null(result$cluster_shade$SNC), "N/A",
                  sprintf("%.6f", result$cluster_shade$SNC))  # Format to 6 decimal places
  verum_p <- ifelse(is.null(result$cluster_shade$Verum), "N/A",
                    sprintf("%.6f", result$cluster_shade$Verum))
  # substr() truncates long scenario names to 45 characters
  cat(sprintf("%-45s %15s %15s\n",
              substr(result$scenario_name, 1, 45), snc_p, verum_p))
}

cat("\n--- Entropy ---\n")
cat(sprintf("%-45s %15s %15s\n", "Scenario", "SNC p-value", "Verum p-value"))
cat(strrep("-", 80), "\n")
for (i in seq_along(sensitivity_results)) {
  result <- sensitivity_results[[i]]
  snc_p <- ifelse(is.null(result$entropy$SNC), "N/A",
                  sprintf("%.6f", result$entropy$SNC))
  verum_p <- ifelse(is.null(result$entropy$Verum), "N/A",
                    sprintf("%.6f", result$entropy$Verum))
  cat(sprintf("%-45s %15s %15s\n",
              substr(result$scenario_name, 1, 45), snc_p, verum_p))
}

cat("\n")

# ANOVAs for all Texture Analysis variables ----------------------

# 2025.10.27 14:30 - ANOVA summary table for all texture parameters
# Exports p-values for all texture variables across sensitivity scenarios
# Color-coded Excel output: red (p<0.01), orange (p<0.05), lilac (p<0.10)

library(openxlsx)
library(car)
library(dplyr)


# Function to run ANOVA and extract p-values
# This function runs either two-way or one-way ANOVA depending on number of experiments
# Returns a data frame with p-values for Day (experiment), Solution, and their Interaction

run_anova_summary <- function(data_subset, param_name) {

  # Check if parameter exists in dataset
  if (!(param_name %in% colnames(data_subset))) {
    return(data.frame(
      Experiment = "Error",
      Solution = "Error",
      Interaction = "Error",
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

  # Use tryCatch to handle any ANOVA failures (e.g., singular design matrix)
  tryCatch({

    if (n_experiments > 1) {
      # Two-way ANOVA with interaction
      # Tests: main effect of experiment_number, main effect of decoded_solution, and their interaction
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
    # If ANOVA fails, return error message
    cat(sprintf("  Warning: ANOVA failed for %s - %s\n", param_name, e$message))
    result$Experiment <- "Error"
    result$Solution <- "Error"
    result$Interaction <- "Error"
  })

  return(result)
}


# Scale small-value variables to avoid numerical issues in ANOVA
# sum_energy and maximum_probability have very small values that can cause problems
# Multiply by 1e6 to bring them to a more manageable scale for the ANOVA package

cat("\n")
cat("========================================================================\n")
cat("SCALING SMALL-VALUE VARIABLES\n")
cat("========================================================================\n")
cat("Multiplying energy and maximum_probability by 1e6 to avoid numerical issues\n")
cat("This scaling does not affect p-values, only helps with numerical stability\n\n")

df2$energy <- df2$energy * 1e6
df2$maximum_probability <- df2$maximum_probability * 1e6


# Create Excel workbook with color-coded p-values

cat("========================================================================\n")
cat("CREATING ANOVA SUMMARY FOR ALL TEXTURE PARAMETERS\n")
cat("========================================================================\n\n")

# Create a new workbook
wb <- createWorkbook()

# Define color styles for significance levels
style_red <- createStyle(fgFill = "#FFB3BA")      # Light red for p < 0.01
style_orange <- createStyle(fgFill = "#FFDFBA")   # Light orange for p < 0.05
style_lilac <- createStyle(fgFill = "#E0BBE4")    # Light lilac for p < 0.10
style_header <- createStyle(textDecoration = "bold", fgFill = "#D3D3D3")

# Define scenarios matching the sensitivity analysis structure
# Each scenario has: name, short_name (for sheet name), and filtered data
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

# All texture parameters available in the dataset
all_texture_params <- c(
  "cluster_prominence",
  "cluster_shade",
  "correlation",
  "diagonal_moment",
  "difference_energy",
  "difference_entropy",
  "energy",
  "entropy",
  "inertia",
  "inverse_different_moment",
  "kappa",
  "maximum_probability",
  "sum_energy",
  "sum_entropy",
  "sum_variance"
)

# Loop through each scenario and create a sheet
for (scenario in scenarios_excel) {

  cat(sprintf("Processing %s for Excel export...\n", scenario$name))

  # Add a worksheet with the short name
  addWorksheet(wb, scenario$short_name)

  # Initialize results list to store ANOVA results for each parameter
  results_list <- list()

  # Run ANOVA for all texture parameters
  for (param in all_texture_params) {
    results_list[[param]] <- run_anova_summary(scenario$data, param)
  }

  # Combine all results into one data frame
  # do.call(rbind, list) stacks all data frames vertically
  results_df <- do.call(rbind, results_list)

  # Add parameter names as first column
  # rownames(results_df) contains the parameter names from the list
  results_df <- data.frame(
    Parameter = rownames(results_df),
    results_df,
    stringsAsFactors = FALSE
  )

  # Format numeric columns to 6 decimals where applicable
  # Loop through the p-value columns
  for (col in c("Experiment", "Solution", "Interaction")) {
    for (row in 1:nrow(results_df)) {
      cell_value <- results_df[row, col]

      # Only format if it's a numeric value (not "Error" or "NA (single exp)")
      # suppressWarnings() prevents warning messages from as.numeric conversion attempts
      if (!is.na(suppressWarnings(as.numeric(cell_value)))) {
        results_df[row, col] <- sprintf("%.6f", as.numeric(cell_value))
      }
    }
  }

  # Write title to worksheet (row 1)
  writeData(wb, scenario$short_name,
            paste0("ANOVA Summary: ", scenario$name, " (Outliers removed per ", basename(outliers_file), ")"),
            startRow = 1, startCol = 1)

  # Write data starting at row 3 (leaving row 2 blank)
  writeData(wb, scenario$short_name, results_df, startRow = 3, rowNames = FALSE)

  # Apply header style to column names
  addStyle(wb, scenario$short_name, style_header,
           rows = 3, cols = 1:4, gridExpand = TRUE)

  # Apply color coding based on p-values
  # Loop through each data row and each p-value column
  for (row_idx in 1:nrow(results_df)) {
    for (col_idx in 2:4) {  # Columns: Experiment, Solution, Interaction

      col_name <- colnames(results_df)[col_idx]
      cell_value <- results_df[row_idx, col_name]

      # Skip if it's an error or NA message
      if (grepl("Error|NA", cell_value)) {
        next
      }

      # Convert to numeric for comparison
      p_val <- suppressWarnings(as.numeric(cell_value))

      if (!is.na(p_val)) {
        excel_row <- row_idx + 3  # +3 because of title and header rows
        excel_col <- col_idx

        # Apply color based on p-value thresholds
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

  # Set column widths for better readability
  setColWidths(wb, scenario$short_name, cols = 1:4, widths = c(25, 15, 15, 15))

  # Add legend at the bottom
  legend_row <- nrow(results_df) + 5
  writeData(wb, scenario$short_name, "Color Legend:",
            startRow = legend_row, startCol = 1)
  writeData(wb, scenario$short_name, "Light lilac = p < 0.10",
            startRow = legend_row + 1, startCol = 1)
  writeData(wb, scenario$short_name, "Light orange = p < 0.05",
            startRow = legend_row + 2, startCol = 1)
  writeData(wb, scenario$short_name, "Light red = p < 0.01",
            startRow = legend_row + 3, startCol = 1)

  # Apply colors to legend text for visual reference
  addStyle(wb, scenario$short_name, style_lilac,
           rows = legend_row + 1, cols = 1)
  addStyle(wb, scenario$short_name, style_orange,
           rows = legend_row + 2, cols = 1)
  addStyle(wb, scenario$short_name, style_red,
           rows = legend_row + 3, cols = 1)

  # Report any errors found during ANOVA execution
  error_cells <- sum(sapply(results_df[, c("Experiment", "Solution", "Interaction")],
                            function(x) sum(x == "Error", na.rm = TRUE)))
  if (error_cells > 0) {
    cat(sprintf("  Note: %d ANOVA analyses failed due to data issues\n", error_cells/3))
  }

  cat(sprintf("  %s complete - %d parameters analyzed\n", scenario$short_name, nrow(results_df)))
}

# Save workbook with timestamp in filename
output_filename <- paste0(date2, "_ASM_TA_ANOVA.xlsx")
saveWorkbook(wb, output_filename, overwrite = TRUE)

cat("\n")
cat("========================================================================\n")
cat(sprintf("ANOVA summary exported to: %s\n", output_filename))
cat("File contains 6 sheets:\n")
cat("  - ALL_SNC, ALL_Verum (all experiments)\n")
cat("  - JZ_SNC, JZ_Verum (JZ data only)\n")
cat("  - AS_SNC, AS_Verum (AS data only)\n")
cat("P-values are color-coded:\n")
cat("  - Light lilac = p < 0.10 (trend)\n")
cat("  - Light orange = p < 0.05 (significant)\n")
cat("  - Light red = p < 0.01 (highly significant)\n")
cat("========================================================================\n\n")




# # PAIR EFFECT ANALYSIS ----------------
# 
# # 2025.10.27 20:00 - Pair effect analysis
# # Tests whether pair 1 and pair 2 (duplicate productions) systematically differ
# # Controls for solution and experiment effects
# 
# library(car)
# library(dplyr)
# 
# 
# # PAIR EFFECT ANALYSIS
# # Test if pair 1 and pair 2 differ systematically
# # This would indicate production batch effects
# #
# # Expected result: No significant pair effect (pairs should be equivalent)
# # If pair is significant: suggests systematic differences between duplicate productions
# 
# cat("\n")
# cat("################################################################################\n")
# cat("PAIR EFFECT ANALYSIS - Testing for systematic differences between pairs\n")
# cat("################################################################################\n\n")
# 
# # Define three analysis scenarios
# scenarios <- list(
#   list(name = "ALL DATA (Experiments 1-10, both AS and JZ)",
#        filter_condition = function(df) df,
#        description = "Full dataset including both experimenters"),
# 
#   list(name = "JZ ONLY (Experiments 3-10)",
#        filter_condition = function(df) df[df$experiment_name == "ASM_JZ", ],
#        description = "Only JZ data: 4 SNC + 4 Verum experiments"),
# 
#   list(name = "AS ONLY (Experiments 1-2)",
#        filter_condition = function(df) df[df$experiment_name == "ASM_AS", ],
#        description = "Only AS data: 1 SNC + 1 Verum experiment")
# )
# 
# # Storage for results comparison
# pair_results <- list()
# 
# # Loop through each scenario
# for (scenario_idx in seq_along(scenarios)) {
# 
#   scenario <- scenarios[[scenario_idx]]
# 
#   cat("================================================================================\n")
#   cat("SCENARIO", scenario_idx, ":", scenario$name, "\n")
#   cat(scenario$description, "\n")
#   cat("================================================================================\n\n")
# 
#   # Apply filter to get subset
#   df_scenario <- scenario$filter_condition(df2)
# 
#   cat("Number of observations:", nrow(df_scenario), "\n")
#   cat("Experiments included:", paste(sort(unique(df_scenario$experiment_number)), collapse = ", "), "\n\n")
# 
#   # Store scenario info
#   pair_results[[scenario_idx]] <- list(
#     scenario_name = scenario$name,
#     n_obs = nrow(df_scenario),
#     cluster_shade = list(),
#     entropy = list()
#   )
# 
#   # Run analysis separately for SNC and Verum within this scenario
#   for (verum_value in c(0, 1)) {
# 
#     condition_label <- ifelse(verum_value == 0, "SNC", "Verum")
# 
#     # Filter for current condition
#     df_subset <- df_scenario[df_scenario$verum == verum_value, ]
# 
#     cat("--------------------------------------------------------------------------------\n")
#     cat(condition_label, "condition\n")
#     cat("--------------------------------------------------------------------------------\n")
#     cat("Observations:", nrow(df_subset), "\n")
# 
#     # Check if we have enough data to run analysis
#     if (nrow(df_subset) < 10) {
#       cat("⚠️  WARNING: Very small sample size - results may be unreliable\n\n")
#     }
# 
#     # Check how many experiments we have in this subset
#     n_experiments <- length(unique(df_subset$experiment_number))
# 
#     if (n_experiments > 1) {
#       # Multiple experiments: control for both solution and experiment effects
#       # Model: outcome ~ pair + decoded_solution + experiment_number
#       # This asks: "Is there a pair effect after controlling for solution and experiment differences?"
#       pair_clu_sha <- aov(cluster_shade ~ pair + decoded_solution + experiment_number, data = df_subset)
#       pair_entropy <- aov(entropy ~ pair + decoded_solution + experiment_number, data = df_subset)
# 
#       results_clu_sha <- Anova(pair_clu_sha, type = "III")
#       results_entropy <- Anova(pair_entropy, type = "III")
# 
#       cat("\n--- Cluster Shade ---\n")
#       print(results_clu_sha, digits = 6)
# 
#       # Extract p-value for pair effect
#       pair_p_clu_sha <- results_clu_sha["pair", "Pr(>F)"]
#       pair_results[[scenario_idx]]$cluster_shade[[condition_label]] <- pair_p_clu_sha
# 
#       cat("\n--- Entropy ---\n")
#       print(results_entropy, digits = 6)
# 
#       pair_p_entropy <- results_entropy["pair", "Pr(>F)"]
#       pair_results[[scenario_idx]]$entropy[[condition_label]] <- pair_p_entropy
# 
#     } else {
#       # Only one experiment: control for solution effect only
#       # Model: outcome ~ pair + decoded_solution
#       cat("\nOnly 1 experiment in this subset - controlling for solution effect only\n")
# 
#       pair_clu_sha <- aov(cluster_shade ~ pair + decoded_solution, data = df_subset)
#       pair_entropy <- aov(entropy ~ pair + decoded_solution, data = df_subset)
# 
#       results_clu_sha <- Anova(pair_clu_sha, type = "III")
#       results_entropy <- Anova(pair_entropy, type = "III")
# 
#       cat("\n--- Cluster Shade ---\n")
#       print(results_clu_sha, digits = 6)
# 
#       pair_p_clu_sha <- results_clu_sha["pair", "Pr(>F)"]
#       pair_results[[scenario_idx]]$cluster_shade[[condition_label]] <- pair_p_clu_sha
# 
#       cat("\n--- Entropy ---\n")
#       print(results_entropy, digits = 6)
# 
#       pair_p_entropy <- results_entropy["pair", "Pr(>F)"]
#       pair_results[[scenario_idx]]$entropy[[condition_label]] <- pair_p_entropy
#     }
# 
#     cat("\n")
#   }
# 
#   cat("\n\n")
# }
# 
# 
# # Summary comparison of pair effect analysis
# 
# cat("################################################################################\n")
# cat("PAIR EFFECT ANALYSIS SUMMARY: Pair p-values across scenarios\n")
# cat("################################################################################\n\n")
# 
# cat("--- Cluster Shade ---\n")
# cat(sprintf("%-45s %15s %15s\n", "Scenario", "SNC p-value", "Verum p-value"))
# cat(strrep("-", 80), "\n")
# for (i in seq_along(pair_results)) {
#   result <- pair_results[[i]]
#   snc_p <- ifelse(is.null(result$cluster_shade$SNC), "N/A",
#                   sprintf("%.6f", result$cluster_shade$SNC))
#   verum_p <- ifelse(is.null(result$cluster_shade$Verum), "N/A",
#                     sprintf("%.6f", result$cluster_shade$Verum))
#   cat(sprintf("%-45s %15s %15s\n",
#               substr(result$scenario_name, 1, 45), snc_p, verum_p))
# }
# 
# cat("\n--- Entropy ---\n")
# cat(sprintf("%-45s %15s %15s\n", "Scenario", "SNC p-value", "Verum p-value"))
# cat(strrep("-", 80), "\n")
# for (i in seq_along(pair_results)) {
#   result <- pair_results[[i]]
#   snc_p <- ifelse(is.null(result$entropy$SNC), "N/A",
#                   sprintf("%.6f", result$entropy$SNC))
#   verum_p <- ifelse(is.null(result$entropy$Verum), "N/A",
#                     sprintf("%.6f", result$entropy$Verum))
#   cat(sprintf("%-45s %15s %15s\n",
#               substr(result$scenario_name, 1, 45), snc_p, verum_p))
# }
# 
# cat("\n")
# cat("INTERPRETATION:\n")
# cat("- Non-significant p-values (p > 0.05): Pairs are equivalent (good!)\n")
# cat("- Significant p-values (p < 0.05): Systematic pair differences exist\n")
# cat("  (suggests production batch effects)\n")
# cat("\n")
# 
# 
# # POST-HOC TESTS WITH COHEN'S D: emmeans pairwise comparisons -----
# 
# # Updated 2025.11.03 16:00 CET: Added emmeans post-hoc analysis with Cohen's d
# 
# library(emmeans)
# library(openxlsx)
# 
# 
# # COMPUTATION FUNCTION
# 
# compute_emmeans_contrasts <- function(data_subset, variable) {
#   # Computes pairwise comparisons, interaction contrasts, and Cohen's d using emmeans
#   #
#   # Args:
#   #   data_subset: dataframe containing the data for analysis
#   #   variable: character string naming the dependent variable
#   #
#   # Returns:
#   #   List with three matrices:
#   #     $potency: main effect p-values for pairwise comparisons
#   #     $interaction: interaction p-values testing if differences vary by experiment
#   #     $cohens_d: standardized effect sizes using model residual SD
# 
#   # Check number of experiments in this dataset
#   n_experiments <- length(unique(data_subset$experiment_number))
# 
#   # Get solution names from factor levels
#   solutions <- levels(data_subset$decoded_solution)
#   n_solutions <- length(solutions)
# 
#   # Debug output: show which solutions are being analyzed
#   cat(sprintf("  Solutions in this analysis: %s\n", paste(solutions, collapse = ", ")))
#   cat(sprintf("  Number of experiments: %d\n", n_experiments))
# 
#   # Initialize result matrices (will fill lower triangle only)
#   potency_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
#                            dimnames = list(solutions, solutions))
#   interaction_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
#                                dimnames = list(solutions, solutions))
#   cohens_d_matrix <- matrix(NA, nrow = n_solutions, ncol = n_solutions,
#                             dimnames = list(solutions, solutions))
# 
# 
#   # Determine model based on number of experiments
#   if (n_experiments > 1) {
#     # Two-way ANOVA with interaction
#     formula_str <- paste(variable, "~ decoded_solution * experiment_number")
#   } else {
#     # One-way ANOVA (only one experiment)
#     formula_str <- paste(variable, "~ decoded_solution")
#   }
# 
#   # Fit ANOVA model
#   model <- aov(as.formula(formula_str), data = data_subset)
# 
# 
#   # STEP 1: Compute main effect pairwise comparisons
# 
#   emm_solution <- emmeans(model, ~ decoded_solution)
# 
#   # Compute all pairwise comparisons without adjustment
#   pairs_solution <- pairs(emm_solution, adjust = "none")
#   pairs_summary <- summary(pairs_solution)
# 
#   # Fill potency matrix with p-values
#   for (i in 1:nrow(pairs_summary)) {
#     # Parse contrast name (e.g., "water_1 - water_2")
#     contrast_name <- as.character(pairs_summary$contrast[i])
#     parts <- strsplit(contrast_name, " - ")[[1]]
#     solution1 <- trimws(parts[1])
#     solution2 <- trimws(parts[2])
# 
#     # Extract p-value
#     p_val <- pairs_summary$p.value[i]
# 
#     # Find indices in solution list
#     idx1 <- which(solutions == solution1)
#     idx2 <- which(solutions == solution2)
# 
#     # Fill lower triangle only (where row index > column index)
#     if (idx1 > idx2) {
#       potency_matrix[idx1, idx2] <- p_val
#     } else {
#       potency_matrix[idx2, idx1] <- p_val
#     }
#   }
# 
# 
#   # COHEN'S D CALCULATION
#   # Calculate standardized effect sizes using model residual SD
# 
#   # Get model residual SD (sigma)
#   residual_sd <- sigma(model)
# 
#   # Fill Cohen's d matrix with effect sizes
#   for (i in 1:nrow(pairs_summary)) {
#     # Parse contrast name
#     contrast_name <- as.character(pairs_summary$contrast[i])
#     parts <- strsplit(contrast_name, " - ")[[1]]
#     solution1 <- trimws(parts[1])
#     solution2 <- trimws(parts[2])
# 
#     # Extract mean difference (estimate)
#     mean_diff <- pairs_summary$estimate[i]
# 
#     # Calculate Cohen's d = mean difference / pooled SD
#     cohens_d <- mean_diff / residual_sd
# 
#     # Find indices in solution list
#     idx1 <- which(solutions == solution1)
#     idx2 <- which(solutions == solution2)
# 
#     # Fill lower triangle only (where row index > column index)
#     if (idx1 > idx2) {
#       cohens_d_matrix[idx1, idx2] <- cohens_d
#     } else {
#       cohens_d_matrix[idx2, idx1] <- cohens_d
#     }
#   }
# 
# 
#   # STEP 2: Compute interaction contrasts (only if multiple experiments)
# 
#   if (n_experiments > 1) {
# 
#     tryCatch({
#       # Get emmeans for solution within each experiment separately
#       emm_by_exp <- emmeans(model, ~ decoded_solution | experiment_number)
# 
#       # Compute pairwise comparisons within each experiment
#       pairs_by_exp <- pairs(emm_by_exp)
#       pairs_by_exp_summary <- summary(pairs_by_exp)
# 
#       # For each pair of solutions, test if their difference varies by experiment
#       for (i in 1:(n_solutions-1)) {
#         for (j in (i+1):n_solutions) {
#           solution1 <- solutions[i]
#           solution2 <- solutions[j]
# 
#           # Find all instances of this comparison across experiments
#           comparison_name <- paste(solution1, "-", solution2)
#           matching_rows <- grep(paste0("^", comparison_name, "$"),
#                                 pairs_by_exp_summary$contrast,
#                                 fixed = FALSE)
# 
#           # Need at least 2 experiments to test for variation
#           if (length(matching_rows) >= 2) {
#             tryCatch({
#               # Get the subset of contrasts for this solution pair
#               specific_contrasts <- pairs_by_exp[matching_rows]
# 
#               # Test if these contrasts differ across experiments
#               pairs_of_contrasts <- pairs(specific_contrasts)
# 
#               # Joint F-test: Are all these experiment differences zero?
#               joint_test <- test(pairs_of_contrasts, joint = TRUE)
#               p_val_int <- joint_test$p.value
# 
#               # Fill matrix (lower triangle: row > col)
#               interaction_matrix[j, i] <- p_val_int
# 
#             }, error = function(e) {
#               # If computation fails for this pair, mark as "Error"
#               interaction_matrix[j, i] <- "Error//pair"
#             })
#           } else {
#             # Not enough experiments for comparison
#             interaction_matrix[j, i] <- "Error/no of exp"
#           }
#         }
#       }
# 
#     }, error = function(e) {
#       # If overall interaction computation fails, mark all as "Error"
#       interaction_matrix[lower.tri(interaction_matrix)] <- "Error//function"
#     })
# 
#   } else {
#     # Only one experiment - cannot test for variation across experiments
#     interaction_matrix[lower.tri(interaction_matrix)] <- "NA (single exp)"
#   }
# 
#   # Return all three matrices
#   return(list(
#     potency = potency_matrix,
#     interaction = interaction_matrix,
#     cohens_d = cohens_d_matrix
#   ))
# }
# 
# 
# # EXCEL WORKBOOK SETUP
# 
# # Create new Excel workbook
# wb_posthoc <- createWorkbook()
# 
# # Define cell styles for color coding significance levels
# style_red <- createStyle(fgFill = "#FFB3BA")      # p < 0.01
# style_orange <- createStyle(fgFill = "#FFDFBA")   # p < 0.05
# style_lilac <- createStyle(fgFill = "#E0BBE4")    # p < 0.10
# style_header <- createStyle(textDecoration = "bold",
#                             fgFill = "#D3D3D3",
#                             halign = "center",
#                             valign = "center")
# style_bold <- createStyle(textDecoration = "bold")
# 
# # Define number format styles
# style_4dec <- createStyle(numFmt = "0.0000")
# style_3dec <- createStyle(numFmt = "0.000")
# 
# # Define six analysis scenarios (same as ANOVA summary)
# scenarios_posthoc <- list(
#   list(name = "ALL DATA - SNC", short_name = "ALL_SNC",
#        data = droplevels(df2[df2$verum == 0, ])),
#   list(name = "ALL DATA - Verum", short_name = "ALL_Verum",
#        data = droplevels(df2[df2$verum == 1, ])),
#   list(name = "JZ ONLY - SNC", short_name = "JZ_SNC",
#        data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 0, ])),
#   list(name = "JZ ONLY - Verum", short_name = "JZ_Verum",
#        data = droplevels(df2[df2$experiment_name == "ASM_JZ" & df2$verum == 1, ])),
#   list(name = "AS ONLY - SNC", short_name = "AS_SNC",
#        data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 0, ])),
#   list(name = "AS ONLY - Verum", short_name = "AS_Verum",
#        data = droplevels(df2[df2$experiment_name == "ASM_AS" & df2$verum == 1, ]))
# )
# 
# 
# # MAIN LOOP: CREATE P-VALUE WORKSHEETS
# 
# cat(sprintf("\n*******************************************************\n"))
# cat(sprintf("Library emmeans procedure - P-VALUES\n"))
# cat(sprintf("\n*******************************************************\n"))
# 
# for (scenario in scenarios_posthoc) {
# 
#   cat(sprintf("\n========================================\n"))
#   cat(sprintf("Processing: %s\n", scenario$name))
#   cat(sprintf("========================================\n"))
# 
#   # Add new worksheet to workbook
#   addWorksheet(wb_posthoc, scenario$short_name)
# 
#   # Initialize row counter for writing data
#   current_row <- 1
# 
#   # Write worksheet title
#   writeData(wb_posthoc, scenario$short_name,
#             paste0("Post Hoc Tests: ", scenario$name, " (Outliers removed per ", basename(outliers_file), ")"),
#             startRow = current_row, startCol = 1)
#   addStyle(wb_posthoc, scenario$short_name, style_bold,
#            rows = current_row, cols = 1)
# 
#   current_row <- current_row + 2  # Skip a row
# 
# 
#   # Loop through the four texture parameters
#   for (var in analysis_vars) {
# 
#     cat(sprintf("\n--- Parameter: %s ---\n", var))
# 
#     # Compute both main effect and interaction contrasts
#     results <- compute_emmeans_contrasts(scenario$data, var)
# 
#     potency_matrix <- results$potency
#     interaction_matrix <- results$interaction
# 
#     # Get solution names for this analysis
#     solutions <- rownames(potency_matrix)
#     n_solutions <- length(solutions)
# 
# 
#     # CREATE TABLE HEADER
# 
#     # Row 1: Parameter name + solution names (spanning 2 columns each)
#     writeData(wb_posthoc, scenario$short_name, var,
#               startRow = current_row, startCol = 1)
# 
#     for (j in 1:n_solutions) {
#       # Write solution name
#       writeData(wb_posthoc, scenario$short_name, solutions[j],
#                 startRow = current_row, startCol = 1 + (j-1)*2 + 1)
# 
#       # Merge the two columns (potency + interaction) under this solution
#       mergeCells(wb_posthoc, scenario$short_name,
#                  rows = current_row,
#                  cols = (1 + (j-1)*2 + 1):(1 + (j-1)*2 + 2))
#     }
# 
#     # Apply header style to entire first row
#     addStyle(wb_posthoc, scenario$short_name, style_header,
#              rows = current_row, cols = 1:(2*n_solutions + 1), gridExpand = TRUE)
# 
#     current_row <- current_row + 1
# 
#     # Row 2: Empty cell + alternating "potency" and "interaction" labels
#     writeData(wb_posthoc, scenario$short_name, "",
#               startRow = current_row, startCol = 1)
# 
#     for (j in 1:n_solutions) {
#       writeData(wb_posthoc, scenario$short_name, "solution",
#                 startRow = current_row, startCol = 1 + (j-1)*2 + 1)
#       writeData(wb_posthoc, scenario$short_name, "interaction",
#                 startRow = current_row, startCol = 1 + (j-1)*2 + 2)
#     }
# 
#     # Apply header style to second row
#     addStyle(wb_posthoc, scenario$short_name, style_header,
#              rows = current_row, cols = 1:(2*n_solutions + 1), gridExpand = TRUE)
# 
#     current_row <- current_row + 1
# 
# 
#     # FILL DATA ROWS
#     # Create triangular table with solution names as row labels
#     # Fill only lower triangle with p-values
# 
#     for (i in 1:n_solutions) {
# 
#       # Write row label (solution name)
#       writeData(wb_posthoc, scenario$short_name, solutions[i],
#                 startRow = current_row, startCol = 1)
# 
#       # Write p-values for this row
#       for (j in 1:n_solutions) {
# 
#         # Calculate column positions for this solution
#         col_sol <- 1 + (j-1)*2 + 1    # "solution" column
#         col_int <- col_sol + 1          # "interaction" column
# 
#         # Only fill lower triangle (where row > column)
#         if (j < i) {
# 
#           # Get p-values from matrices
#           p_sol <- potency_matrix[i, j]
#           p_int <- interaction_matrix[i, j]
# 
#           # WRITE SOLUTION P-VALUE (main effect)
#           if (is.numeric(p_sol) && !is.na(p_sol)) {
#             # Format and write numeric value with explicit decimal format
#             writeData(wb_posthoc, scenario$short_name, sprintf("%.4f", p_sol),
#                       startRow = current_row, startCol = col_sol)
# 
#             # Apply color coding based on significance
#             if (p_sol < 0.01) {
#               addStyle(wb_posthoc, scenario$short_name, style_red,
#                        rows = current_row, cols = col_sol, stack = TRUE)
#             } else if (p_sol < 0.05) {
#               addStyle(wb_posthoc, scenario$short_name, style_orange,
#                        rows = current_row, cols = col_sol, stack = TRUE)
#             } else if (p_sol < 0.10) {
#               addStyle(wb_posthoc, scenario$short_name, style_lilac,
#                        rows = current_row, cols = col_sol, stack = TRUE)
#             }
#           } else if (!is.na(p_sol)) {
#             # Write non-numeric value (e.g., "Error")
#             writeData(wb_posthoc, scenario$short_name, as.character(p_sol),
#                       startRow = current_row, startCol = col_sol)
#           }
# 
#           # WRITE INTERACTION P-VALUE
#           if (is.numeric(p_int) && !is.na(p_int)) {
#             # Format and write numeric value with explicit decimal format
#             writeData(wb_posthoc, scenario$short_name, sprintf("%.4f", p_int),
#                       startRow = current_row, startCol = col_int)
# 
#             # Apply color coding based on significance
#             if (p_int < 0.01) {
#               addStyle(wb_posthoc, scenario$short_name, style_red,
#                        rows = current_row, cols = col_int, stack = TRUE)
#             } else if (p_int < 0.05) {
#               addStyle(wb_posthoc, scenario$short_name, style_orange,
#                        rows = current_row, cols = col_int, stack = TRUE)
#             } else if (p_int < 0.10) {
#               addStyle(wb_posthoc, scenario$short_name, style_lilac,
#                        rows = current_row, cols = col_int, stack = TRUE)
#             }
#           } else if (!is.na(p_int)) {
#             # Write non-numeric value (e.g., "Error", "NA (single exp)")
#             writeData(wb_posthoc, scenario$short_name, as.character(p_int),
#                       startRow = current_row, startCol = col_int)
#           }
#         }
#         # Upper triangle and diagonal remain empty
#       }
# 
#       current_row <- current_row + 1
#     }
# 
#     # Add spacing between parameters
#     current_row <- current_row + 2
#   }
# 
# 
#   # WORKSHEET FOOTER
# 
#   # Set column widths for readability
#   setColWidths(wb_posthoc, scenario$short_name,
#                cols = 1:(2*n_solutions + 1),
#                widths = c(15, rep(10, 2*n_solutions)))
# 
#   # Add color legend
#   legend_row <- current_row + 1
# 
#   writeData(wb_posthoc, scenario$short_name, "Color Legend:",
#             startRow = legend_row, startCol = 1)
# 
#   writeData(wb_posthoc, scenario$short_name, "Light lilac = p < 0.10",
#             startRow = legend_row + 1, startCol = 1)
#   addStyle(wb_posthoc, scenario$short_name, style_lilac,
#            rows = legend_row + 1, cols = 1)
# 
#   writeData(wb_posthoc, scenario$short_name, "Light orange = p < 0.05",
#             startRow = legend_row + 2, startCol = 1)
#   addStyle(wb_posthoc, scenario$short_name, style_orange,
#            rows = legend_row + 2, cols = 1)
# 
#   writeData(wb_posthoc, scenario$short_name, "Light red = p < 0.01",
#             startRow = legend_row + 3, startCol = 1)
#   addStyle(wb_posthoc, scenario$short_name, style_red,
#            rows = legend_row + 3, cols = 1)
# 
#   # Add interpretation note
#   writeData(wb_posthoc, scenario$short_name,
#             "Note: 'solution' = Main effect p-value (marginal comparison averaged over experiments)",
#             startRow = legend_row + 5, startCol = 1)
# 
#   writeData(wb_posthoc, scenario$short_name,
#             "      'interaction' = Does this comparison vary across experiments? (interaction test)",
#             startRow = legend_row + 6, startCol = 1)
# 
#   writeData(wb_posthoc, scenario$short_name,
#             "      If interaction p-value is significant, the main effect may be misleading.",
#             startRow = legend_row + 7, startCol = 1)
# }
# 
# 
# # COHEN'S D WORKSHEETS
# 
# cat(sprintf("\n*******************************************************\n"))
# cat(sprintf("Creating Cohen's d effect size worksheets\n"))
# cat(sprintf("\n*******************************************************\n"))
# 
# for (scenario in scenarios_posthoc) {
# 
#   cat(sprintf("\n========================================\n"))
#   cat(sprintf("Processing Cohen's d: %s\n", scenario$name))
#   cat(sprintf("========================================\n"))
# 
#   # Add new worksheet with "_d" suffix
#   sheet_name_d <- paste0(scenario$short_name, "_d")
#   addWorksheet(wb_posthoc, sheet_name_d)
# 
#   # Initialize row counter
#   current_row <- 1
# 
#   # Write worksheet title
#   writeData(wb_posthoc, sheet_name_d,
#             paste0("Cohen's d Effect Sizes: ", scenario$name, " (Outliers removed per ", basename(outliers_file), ")"),
#             startRow = current_row, startCol = 1)
#   addStyle(wb_posthoc, sheet_name_d, style_bold,
#            rows = current_row, cols = 1)
# 
#   current_row <- current_row + 2  # Skip a row
# 
# 
#   # Loop through the four texture parameters
#   for (var in analysis_vars) {
# 
#     cat(sprintf("\n--- Parameter: %s ---\n", var))
# 
#     # Compute contrasts (reuse same function)
#     results <- compute_emmeans_contrasts(scenario$data, var)
# 
#     cohens_d_matrix <- results$cohens_d
# 
#     # Get solution names for this analysis
#     solutions <- rownames(cohens_d_matrix)
#     n_solutions <- length(solutions)
# 
# 
#     # CREATE TABLE HEADER
# 
#     # Row 1: Parameter name + solution names (one column each for Cohen's d)
#     writeData(wb_posthoc, sheet_name_d, var,
#               startRow = current_row, startCol = 1)
# 
#     for (j in 1:n_solutions) {
#       # Write solution name
#       writeData(wb_posthoc, sheet_name_d, solutions[j],
#                 startRow = current_row, startCol = 1 + j)
#     }
# 
#     # Apply header style to entire first row
#     addStyle(wb_posthoc, sheet_name_d, style_header,
#              rows = current_row, cols = 1:(n_solutions + 1), gridExpand = TRUE)
# 
#     current_row <- current_row + 1
# 
# 
#     # FILL DATA ROWS
#     # Create triangular table with solution names as row labels
#     # Fill only lower triangle with Cohen's d values
# 
#     for (i in 1:n_solutions) {
# 
#       # Write row label (solution name)
#       writeData(wb_posthoc, sheet_name_d, solutions[i],
#                 startRow = current_row, startCol = 1)
# 
#       # Write Cohen's d values for this row
#       for (j in 1:n_solutions) {
# 
#         # Calculate column position
#         col_d <- 1 + j
# 
#         # Only fill lower triangle (where row > column)
#         if (j < i) {
# 
#           # Get Cohen's d value from matrix
#           d_val <- cohens_d_matrix[i, j]
# 
#           # Write Cohen's d value
#           if (is.numeric(d_val) && !is.na(d_val)) {
#             # Format and write numeric value with explicit decimal format
#             writeData(wb_posthoc, sheet_name_d, sprintf("%.3f", d_val),
#                       startRow = current_row, startCol = col_d)
#           } else if (!is.na(d_val)) {
#             # Write non-numeric value (e.g., "Error")
#             writeData(wb_posthoc, sheet_name_d, as.character(d_val),
#                       startRow = current_row, startCol = col_d)
#           }
#         }
#         # Upper triangle and diagonal remain empty
#       }
# 
#       current_row <- current_row + 1
#     }
# 
#     # Add spacing between parameters
#     current_row <- current_row + 2
#   }
# 
# 
#   # WORKSHEET FOOTER
# 
#   # Set column widths for readability
#   setColWidths(wb_posthoc, sheet_name_d,
#                cols = 1:(n_solutions + 1),
#                widths = c(15, rep(10, n_solutions)))
# 
#   # Add interpretation note
#   legend_row <- current_row + 1
# 
#   writeData(wb_posthoc, sheet_name_d, "Cohen's d interpretation:",
#             startRow = legend_row, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d, "Small effect: |d| ≈ 0.2",
#             startRow = legend_row + 1, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d, "Medium effect: |d| ≈ 0.5",
#             startRow = legend_row + 2, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d, "Large effect: |d| ≈ 0.8",
#             startRow = legend_row + 3, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d,
#             "Note: Cohen's d calculated using model residual SD (pooled across all groups and experiments)",
#             startRow = legend_row + 5, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d,
#             "      Positive d = row solution has higher value than column solution",
#             startRow = legend_row + 6, startCol = 1)
# 
#   writeData(wb_posthoc, sheet_name_d,
#             "      Negative d = row solution has lower value than column solution",
#             startRow = legend_row + 7, startCol = 1)
# }
# 
# 
# # SAVE WORKBOOK
# 
# output_filename_posthoc <- paste0(date2, "_ASM_TA_posthoc_cohensd.xlsx")
# saveWorkbook(wb_posthoc, output_filename_posthoc, overwrite = TRUE)
# 
# cat(sprintf("\n\n================================================================================\n"))
# cat(sprintf("Post hoc tests exported to: %s\n", output_filename_posthoc))
# cat("================================================================================\n")
# cat("File contains 12 sheets:\n")
# cat("  P-VALUES: ALL_SNC, ALL_Verum, JZ_SNC, JZ_Verum, AS_SNC, AS_Verum\n")
# cat("  COHEN'S D: ALL_SNC_d, ALL_Verum_d, JZ_SNC_d, JZ_Verum_d, AS_SNC_d, AS_Verum_d\n")
# cat("\nEach sheet has 4 parameters (cluster_shade, entropy, maximum_probability, kappa)\n")
# cat("Each parameter shows triangular table with solution comparisons\n")
# cat("\nP-value sheets:\n")
# cat("  - Two values per comparison: 'solution' (main effect) and 'interaction'\n")
# cat("  - Color coding: lilac (p<0.10), orange (p<0.05), red (p<0.01)\n")
# cat("\nCohen's d sheets:\n")
# cat("  - Standardized effect sizes using model residual SD\n")
# cat("  - Interpretation: small (|d|≈0.2), medium (|d|≈0.5), large (|d|≈0.8)\n")
# cat(sprintf("Outliers removed per %s\n", basename(outliers_file)))
# cat("================================================================================\n\n")
# 

# # Calculate average diagonal_moment for each solution ------
# 
# library(dplyr)
# 
# diagonal_moment_summary <- df2 %>%
#   group_by(decoded_solution) %>%
#   summarise(
#     n = n(),
#     mean = mean(diagonal_moment, na.rm = TRUE),
#     sd = sd(diagonal_moment, na.rm = TRUE)
#   ) %>%
#   arrange(decoded_solution)
# 
# cat("\n")
# cat("========================================================================\n")
# cat("DIAGONAL MOMENT - Average by Solution (all pairs combined)\n")
# cat("========================================================================\n")
# 
# # Set option to show 6 decimals
# options(pillar.sigfig = 7)
# print(diagonal_moment_summary, n = Inf)
# 
# cat("========================================================================\n\n")


# # ASM Texture Analysis Plotting Script ------
# # Creates 4x2 grid showing parameter trends across experiments for SNC and Verum
# # Adapted from ASPS plotting approach
# # Created: 2025-11-11
# 
# # This script assumes you've already run the ASM ANOVA script which created:
# # - df2: the main dataframe with decoded solutions and pairs
# 
# library(ggplot2)
# library(plotrix)     # for std.error
# library(gridExtra)   # for grid.arrange
# library(dplyr)
# library(colorspace)  # for darken() function
# 
# 
# # Export date for filename
# 
# date <- Sys.Date()
# date2 <- gsub("-| |UTC", "", date)
# 
# 
# # Define the four parameters to plot
# 
# analysis_vars <- c("cluster_shade", "entropy", "maximum_probability", "kappa")
# 
# 
# # Define conditions (SNC vs Verum)
# 
# conditions <- c(0, 1)  # 0 = SNC, 1 = Verum
# condition_labels <- c("SNC (Water Controls)", "Verum (Treatment)")
# 
# 
# # Create solution-pair combination variable
# 
# cat("Creating solution-pair combination variable...\n")
# 
# df2 <- df2 %>%
#   mutate(
#     # Create shortened labels for legend
#     solution_pair = case_when(
#       decoded_solution == "water_1" & pair == 1 ~ "W1-1",
#       decoded_solution == "water_1" & pair == 2 ~ "W1-2",
#       decoded_solution == "water_2" & pair == 1 ~ "W2-1",
#       decoded_solution == "water_2" & pair == 2 ~ "W2-2",
#       decoded_solution == "water_3" & pair == 1 ~ "W3-1",
#       decoded_solution == "water_3" & pair == 2 ~ "W3-2",
#       decoded_solution == "potency_x" & pair == 1 ~ "X1",
#       decoded_solution == "potency_x" & pair == 2 ~ "X2",
#       decoded_solution == "potency_y" & pair == 1 ~ "Y1",
#       decoded_solution == "potency_y" & pair == 2 ~ "Y2",
#       decoded_solution == "potency_z" & pair == 1 ~ "Z1",
#       decoded_solution == "potency_z" & pair == 2 ~ "Z2",
#       TRUE ~ NA_character_
#     ),
#     # Also create a base solution identifier for color mapping
#     base_solution = case_when(
#       decoded_solution == "water_1" ~ "sol1",
#       decoded_solution == "water_2" ~ "sol2",
#       decoded_solution == "water_3" ~ "sol3",
#       decoded_solution == "potency_x" ~ "sol1",
#       decoded_solution == "potency_y" ~ "sol2",
#       decoded_solution == "potency_z" ~ "sol3",
#       TRUE ~ NA_character_
#     )
#   )
# 
# cat("Solution-pair combinations created.\n\n")
# 
# 
# # Calculate summary statistics
# 
# cat("Calculating summary statistics for plotting...\n")
# 
# summary_data <- df2 %>%
#   group_by(verum, sub_exp_number, decoded_solution, pair, solution_pair, base_solution) %>%
#   summarise(
#     across(all_of(analysis_vars),
#            list(mean = ~mean(.x, na.rm = TRUE),
#                 se = ~std.error(.x, na.rm = TRUE)),
#            .names = "{.col}_{.fn}"),
#     n = n(),
#     .groups = "drop"
#   )
# 
# cat("Summary statistics calculated.\n")
# cat("Unique combinations:", nrow(summary_data), "\n\n")
# 
# 
# # Determine y-axis limits for each parameter
# 
# cat("Determining y-axis limits for each parameter...\n")
# 
# y_limits <- list()
# 
# for (var in analysis_vars) {
#   # Get column names for mean and SE
#   mean_col <- paste0(var, "_mean")
#   se_col <- paste0(var, "_se")
#   
#   # Calculate range including error bars across BOTH SNC and Verum
#   y_min <- min(summary_data[[mean_col]] - summary_data[[se_col]], na.rm = TRUE)
#   y_max <- max(summary_data[[mean_col]] + summary_data[[se_col]], na.rm = TRUE)
#   
#   # Add 5% padding for visual clarity
#   y_range <- y_max - y_min
#   y_limits[[var]] <- c(y_min - 0.05 * y_range, y_max + 0.05 * y_range)
#   
#   cat(sprintf("  %s: [%.6f, %.6f]\n", var, y_limits[[var]][1], y_limits[[var]][2]))
# }
# 
# cat("\n")
# 
# 
# # Define color palette
# 
# cat("Setting up color palette...\n")
# 
# # Extract Dark2 palette colors at positions 4, 5, 6
# dark2_colors <- RColorBrewer::brewer.pal(8, "Dark2")
# base_colors <- dark2_colors[4:6]  # Positions 4, 5, 6
# 
# cat("Base colors (Dark2 positions 4-6):\n")
# cat(sprintf("  Position 4 (sol1): %s\n", base_colors[1]))
# cat(sprintf("  Position 5 (sol2): %s\n", base_colors[2]))
# cat(sprintf("  Position 6 (sol3): %s\n", base_colors[3]))
# 
# # Create darkened versions for pair 2
# darkened_colors <- darken(base_colors, amount = 0.3)
# 
# cat("\nDarkened colors (for pair 2):\n")
# cat(sprintf("  Position 4 dark: %s\n", darkened_colors[1]))
# cat(sprintf("  Position 5 dark: %s\n", darkened_colors[2]))
# cat(sprintf("  Position 6 dark: %s\n", darkened_colors[3]))
# 
# # Build complete color mapping
# # This maps solution-pair combinations to specific colors
# color_mapping_snc <- c(
#   "W1-1" = base_colors[1],      # water_1 pair 1
#   "W1-2" = darkened_colors[1],  # water_1 pair 2
#   "W2-1" = base_colors[2],      # water_2 pair 1
#   "W2-2" = darkened_colors[2],  # water_2 pair 2
#   "W3-1" = base_colors[3],      # water_3 pair 1
#   "W3-2" = darkened_colors[3]   # water_3 pair 2
# )
# 
# color_mapping_verum <- c(
#   "X1" = base_colors[1],      # potency_x pair 1
#   "X2" = darkened_colors[1],  # potency_x pair 2
#   "Y1" = base_colors[2],      # potency_y pair 1
#   "Y2" = darkened_colors[2],  # potency_y pair 2
#   "Z1" = base_colors[3],      # potency_z pair 1
#   "Z2" = darkened_colors[3]   # potency_z pair 2
# )
# 
# cat("\nColor mappings created.\n\n")
# 
# 
# # Create plots
# 
# cat("Creating plots...\n")
# 
# plot_list <- list()
# plot_counter <- 1
# 
# # Loop through parameters (rows in final grid)
# for (var in analysis_vars) {
#   
#   cat(sprintf("  Processing parameter: %s\n", var))
#   
#   # Get column names for this parameter
#   mean_col <- paste0(var, "_mean")
#   se_col <- paste0(var, "_se")
#   
#   # Loop through conditions (columns in final grid)
#   for (condition_idx in seq_along(conditions)) {
#     
#     verum_value <- conditions[condition_idx]
#     condition_label <- condition_labels[condition_idx]
#     
#     # Filter data for this condition
#     plot_data <- summary_data %>%
#       filter(verum == verum_value)
#     
#     # Select appropriate color mapping
#     if (verum_value == 0) {
#       color_map <- color_mapping_snc
#     } else {
#       color_map <- color_mapping_verum
#     }
#     
#     # Define position dodge for clearer visualization
#     dodge_width <- 0.3
#     
#     # Create the plot
#     p <- ggplot(plot_data, aes(x = as.numeric(as.character(sub_exp_number)),
#                                y = .data[[mean_col]],
#                                color = solution_pair,
#                                group = solution_pair)) +
#       # Lines connecting points
#       geom_line(linewidth = 0.7, position = position_dodge(width = dodge_width)) +
#       # Error bars
#       geom_errorbar(aes(ymin = .data[[mean_col]] - .data[[se_col]],
#                         ymax = .data[[mean_col]] + .data[[se_col]]),
#                     width = 0.2, linewidth = 0.5,
#                     position = position_dodge(width = dodge_width)) +
#       # Points
#       geom_point(size = 2.5, position = position_dodge(width = dodge_width)) +
#       # Scales and labels
#       scale_x_continuous(breaks = 1:5, 
#                          labels = 1:5) +
#       scale_y_continuous(limits = y_limits[[var]]) +
#       scale_color_manual(values = color_map, 
#                          name = "Solution") +
#       labs(
#         title = condition_label,
#         x = "Sub-experiment Number",
#         y = var,
#         caption = "Note: Sub-exp 1 = AS data; Sub-exp 2-5 = JZ data"
#       ) +
#       theme_bw() +
#       theme(
#         plot.title = element_text(size = 11, face = "bold"),
#         axis.title = element_text(size = 10),
#         axis.text = element_text(size = 9),
#         legend.title = element_text(size = 9),
#         legend.text = element_text(size = 8),
#         plot.caption = element_text(size = 8, hjust = 0)
#       )
#     
#     # Store plot in list
#     plot_list[[plot_counter]] <- p
#     plot_counter <- plot_counter + 1
#   }
# }
# 
# cat("All plots created.\n\n")
# 
# 
# # Arrange plots in grid
# 
# cat("Arranging plots in grid...\n")
# 
# # Create the grid arrangement
# # 4 rows (parameters) × 2 columns (SNC vs Verum)
# grid_plot <- grid.arrange(
#   grobs = plot_list,
#   nrow = 4,
#   ncol = 2,
#   top = paste0("ASM_texture_parameters", outliers_file)
#   #top = paste0("ASM_texture_parameters")                 #Alternative when outliers are fully included
# )
# 
# 
# # Export plot
# 
# output_filename <- paste0(date2, "_ASM_texture_parameters_grid.png")
# 
# ggsave(
#   filename = output_filename,
#   plot = grid_plot,
#   width = 30,      # Wide enough for 2 columns with legends
#   height = 35,     # Tall enough for 4 rows
#   dpi = 300,
#   units = "cm"
# )
# 
# cat(sprintf("Plot saved as: %s\n", output_filename))
# cat("Done!\n")
# # ASM Texture Analysis - Full Parameter Plotting Script ================
# # Extends the 4-parameter plot to all 15 texture parameters
# # Run after main analysis script (df2 and all variables must exist)
# # 2025-12-11
# 
# # Setup
# 
# # Expand analysis_vars to include all texture parameters
# 
# all_texture_params <- c(
#   "cluster_prominence",
#   "cluster_shade",
#   "correlation",
#   "diagonal_moment",
#   "difference_energy",
#   "difference_entropy",
#   "energy",
#   "entropy",
#   "inertia",
#   "inverse_different_moment",
#   "kappa",
#   "maximum_probability",
#   "sum_energy",
#   "sum_entropy",
#   "sum_variance"
# )
# 
# # Parameters requiring scaling (very small values)
# # Scale by 1e6 for readable axis labels
# 
# params_to_scale <- c("energy", "maximum_probability")
# 
# cat("\nScaling parameters by 1e6:", paste(params_to_scale, collapse = ", "), "\n")
# 
# for (param in params_to_scale) {
#   df2[[param]] <- df2[[param]] * 1e6
# }
# 
# # Create axis labels with scaling notation where applicable
# 
# axis_labels <- setNames(all_texture_params, all_texture_params)
# axis_labels["energy"] <- "energy (×10⁶)"
# axis_labels["maximum_probability"] <- "maximum_probability (×10⁶)"
# 
# 
# # Recalculate summary statistics for all parameters
# 
# cat("\nCalculating summary statistics for all 15 parameters...\n")
# 
# summary_data_full <- df2 %>%
#   group_by(verum, sub_exp_number, decoded_solution, pair, solution_pair, base_solution) %>%
#   summarise(
#     across(all_of(all_texture_params),
#            list(mean = ~mean(.x, na.rm = TRUE),
#                 se = ~std.error(.x, na.rm = TRUE)),
#            .names = "{.col}_{.fn}"),
#     n = n(),
#     .groups = "drop"
#   )
# 
# cat("Summary statistics calculated:", nrow(summary_data_full), "unique combinations\n")
# 
# 
# # Determine y-axis limits for all parameters 
# 
# # Shared y-limits across SNC and Verum for visual comparability
# 
# cat("\nDetermining y-axis limits...\n")
# 
# y_limits_full <- list()
# 
# for (var in all_texture_params) {
#   mean_col <- paste0(var, "_mean")
#   se_col <- paste0(var, "_se")
#   
#   y_min <- min(summary_data_full[[mean_col]] - summary_data_full[[se_col]], na.rm = TRUE)
#   y_max <- max(summary_data_full[[mean_col]] + summary_data_full[[se_col]], na.rm = TRUE)
#   
#   # Add 5% padding
#   y_range <- y_max - y_min
#   y_limits_full[[var]] <- c(y_min - 0.05 * y_range, y_max + 0.05 * y_range)
# }
# 
# 
# #Create plots for all parameters 
# 
# cat("\nCreating plots for all 15 parameters...\n")
# 
# plot_list_full <- list()
# plot_counter <- 1
# 
# for (var in all_texture_params) {
#   
#   cat(sprintf("  %s\n", var))
#   
#   mean_col <- paste0(var, "_mean")
#   se_col <- paste0(var, "_se")
#   
#   for (condition_idx in seq_along(conditions)) {
#     
#     verum_value <- conditions[condition_idx]
#     condition_label <- condition_labels[condition_idx]
#     
#     plot_data <- summary_data_full %>% filter(verum == verum_value)
#     
#     color_map <- if (verum_value == 0) color_mapping_snc else color_mapping_verum
#     
#     dodge_width <- 0.3
#     
#     p <- ggplot(plot_data, aes(x = as.numeric(as.character(sub_exp_number)),
#                                y = .data[[mean_col]],
#                                color = solution_pair,
#                                group = solution_pair)) +
#       geom_line(linewidth = 0.7, position = position_dodge(width = dodge_width)) +
#       geom_errorbar(aes(ymin = .data[[mean_col]] - .data[[se_col]],
#                         ymax = .data[[mean_col]] + .data[[se_col]]),
#                     width = 0.2, linewidth = 0.5,
#                     position = position_dodge(width = dodge_width)) +
#       geom_point(size = 2.5, position = position_dodge(width = dodge_width)) +
#       scale_x_continuous(breaks = 1:5, labels = 1:5) +
#       scale_y_continuous(limits = y_limits_full[[var]]) +
#       scale_color_manual(values = color_map, name = "Solution") +
#       labs(
#         title = condition_label,
#         x = "Sub-experiment Number",
#         y = axis_labels[[var]],
#         caption = "Note: Sub-exp 1 = AS data; Sub-exp 2-5 = JZ data"
#       ) +
#       theme_bw() +
#       theme(
#         plot.title = element_text(size = 11, face = "bold"),
#         axis.title = element_text(size = 10),
#         axis.text = element_text(size = 9),
#         legend.title = element_text(size = 9),
#         legend.text = element_text(size = 8),
#         plot.caption = element_text(size = 8, hjust = 0)
#       )
#     
#     plot_list_full[[plot_counter]] <- p
#     plot_counter <- plot_counter + 1
#   }
# }
# 
# cat("All", length(plot_list_full), "plots created\n")
# 
# 
# # Arrange and export 
# 
# cat("\nArranging plots in grid...\n")
# 
# # Figure height: ~8.75 cm per row (same ratio as original 35cm/4rows)
# 
# n_rows <- length(all_texture_params)
# fig_height <- n_rows * 8.75
# 
# grid_plot_full <- grid.arrange(
#   grobs = plot_list_full,
#   nrow = n_rows,
#   ncol = 2,
#   top = paste0("ASM Texture Analysis - All Parameters (Outliers: ", basename(outliers_file), ")")
# )
# 
# output_filename_full <- paste0(date2, "_ASM_texture_parameters_ALL_grid.png")
# 
# ggsave(
#   filename = output_filename_full,
#   plot = grid_plot_full,
#   width = 30,
#   height = fig_height,
#   dpi = 300,
#   units = "cm",
#   limitsize = FALSE
# )
# 
# cat(sprintf("\nPlot saved as: %s\n", output_filename_full))
# cat(sprintf("Dimensions: 30 x %.1f cm\n", fig_height))
# cat("Done!\n")

# ASM Texture Analysis - Correlation Matrix Plot ---------------
# Creates 15x15 scatterplot matrix showing pairwise correlations between texture parameters
# Run after main analysis script (df2 must exist with outliers removed)
# 2025-12-11

#
# Setup
#

library(GGally)
library(ggplot2)
library(dplyr)
library(RColorBrewer)

# Export date for filename

date <- Sys.Date()
date2 <- gsub("-| |UTC", "", date)


#
# Define parameters and colors
#

all_texture_params <- c(
  "cluster_prominence",
  "cluster_shade",
  "correlation",
  "diagonal_moment",
  "difference_energy",
  "difference_entropy",
  "energy",
  "entropy",
  "inertia",
  "inverse_different_moment",
  "kappa",
  "maximum_probability",
  "sum_energy",
  "sum_entropy",
  "sum_variance"
)

# Create experiment color mapping
# SNC experiments (1, 3, 6, 8, 9): blue shades
# Verum experiments (2, 4, 5, 7, 10): red shades

blue_shades <- colorRampPalette(c("#08519c", "#6baed6"))(5)  # Dark to light blue
red_shades <- colorRampPalette(c("#a50f15", "#fc9272"))(5)   # Dark to light red

# Map experiment numbers to colors
# SNC: 1->blue1, 3->blue2, 6->blue3, 8->blue4, 9->blue5
# Verum: 2->red1, 4->red2, 5->red3, 7->red4, 10->red5

experiment_colors <- c(
  "1" = blue_shades[1],
  "3" = blue_shades[2],
  "6" = blue_shades[3],
  "8" = blue_shades[4],
  "9" = blue_shades[5],
  "2" = red_shades[1],
  "4" = red_shades[2],
  "5" = red_shades[3],
  "7" = red_shades[4],
  "10" = red_shades[5]
)

# Create legend labels for clarity

experiment_labels <- c(
  "1" = "SNC Exp 1",
  "3" = "SNC Exp 3",
  "6" = "SNC Exp 6",
  "8" = "SNC Exp 8",
  "9" = "SNC Exp 9",
  "2" = "Verum Exp 2",
  "4" = "Verum Exp 4",
  "5" = "Verum Exp 5",
  "7" = "Verum Exp 7",
  "10" = "Verum Exp 10"
)


#
# Prepare data
#

cat("\nPreparing data for correlation matrix...\n")

# Select only the texture parameters and experiment number
# Convert experiment_number to character for color mapping

plot_data <- df2 %>%
  select(all_of(all_texture_params), experiment_number) %>%
  mutate(experiment_number = as.character(experiment_number))

cat("Data prepared:", nrow(plot_data), "observations\n")
cat("Parameters:", length(all_texture_params), "\n")


#
# Custom panel functions
#

# Lower triangle: scatterplot with regression line and R² with significance stars

lower_scatter <- function(data, mapping, ...) {
  
  # Extract variable names from mapping
  x_var <- as.character(mapping$x)[2]
  y_var <- as.character(mapping$y)[2]
  
  # Calculate R² and p-value for significance stars
  x_vals <- data[[x_var]]
  y_vals <- data[[y_var]]
  
  # Remove NAs for correlation calculation
  complete_cases <- complete.cases(x_vals, y_vals)
  x_clean <- x_vals[complete_cases]
  y_clean <- y_vals[complete_cases]
  
  # Fit linear model and extract R² and p-value
  if (length(x_clean) > 2) {
    fit <- lm(y_clean ~ x_clean)
    r_squared <- summary(fit)$r.squared
    p_value <- summary(fit)$coefficients[2, 4]
    
    # Determine significance stars
    stars <- ""
    if (p_value < 0.001) stars <- "***"
    else if (p_value < 0.01) stars <- "**"
    else if (p_value < 0.05) stars <- "*"
    
    r2_label <- sprintf("R²=%.2f%s", r_squared, stars)
  } else {
    r2_label <- "R²=NA"
  }
  
  # Create the plot
  p <- ggplot(data, mapping) +
    geom_point(aes(color = experiment_number), alpha = 0.5, size = 0.5) +
    geom_smooth(method = "lm", se = FALSE, color = "black", linewidth = 0.5) +
    scale_color_manual(values = experiment_colors) +
    annotate("text", x = -Inf, y = Inf, label = r2_label,
             hjust = -0.1, vjust = 1.5, size = 2) +
    theme_minimal() +
    theme(
      panel.grid = element_blank(),
      axis.text = element_blank(),
      axis.ticks = element_blank()
    )
  
  return(p)
}


# Upper triangle: correlation coefficient as text

upper_cor <- function(data, mapping, ...) {
  
  # Extract variable names from mapping
  x_var <- as.character(mapping$x)[2]
  y_var <- as.character(mapping$y)[2]
  
  # Calculate correlation
  x_vals <- data[[x_var]]
  y_vals <- data[[y_var]]
  
  cor_val <- cor(x_vals, y_vals, use = "complete.obs")
  
  # Color based on correlation strength
  # Blue for negative, red for positive
  cor_color <- colorRampPalette(c("#001A00", "#00FF00"))(100)[max(1, round(abs(cor_val) * 99) + 1)]
  
  # Create plot with correlation text
  p <- ggplot(data, mapping) +
    annotate("text", x = 0.5, y = 0.5, 
             label = sprintf("%.2f", cor_val),
             size = 3, color = cor_color, fontface = "bold") +
    xlim(0, 1) + ylim(0, 1) +
    theme_void()
  
  return(p)
}


# Diagonal: variable names

diag_label <- function(data, mapping, ...) {
  
  # Extract variable name from mapping
  var_name <- as.character(mapping$x)[2]
  
  # Shorten long variable names for display
  display_name <- gsub("_", "\n", var_name)
  
  p <- ggplot(data, mapping) +
    annotate("text", x = 0.5, y = 0.5, label = display_name,
             size = 2.5, fontface = "bold", lineheight = 0.8) +
    xlim(0, 1) + ylim(0, 1) +
    theme_void()
  
  return(p)
}


#
# Create correlation matrix plot
#

cat("\nCreating correlation matrix plot...\n")
cat("This may take a few minutes with 15 variables and", nrow(plot_data), "observations...\n")

# Create the ggpairs plot

cor_matrix <- ggpairs(
  plot_data,
  columns = 1:15,
  upper = list(continuous = upper_cor),
  lower = list(continuous = lower_scatter),
  diag = list(continuous = diag_label),
  legend = c(2, 1),
  progress = TRUE
) +
  theme(
    strip.text = element_blank(),
    panel.spacing = unit(0.1, "lines")
  )

cat("Correlation matrix created.\n")


#
# Create separate legend
#

# GGally doesn't handle legends well, so create a separate one

legend_data <- data.frame(
  experiment = factor(names(experiment_labels), levels = names(experiment_labels)),
  label = experiment_labels,
  x = 1:10,
  y = 1:10
)

legend_plot <- ggplot(legend_data, aes(x = x, y = y, color = experiment)) +
  geom_point(size = 3) +
  scale_color_manual(
    values = experiment_colors,
    labels = experiment_labels,
    name = "Experiment"
  ) +
  guides(color = guide_legend(ncol = 2)) +
  theme_void() +
  theme(
    legend.position = "bottom",
    legend.title = element_text(face = "bold"),
    legend.text = element_text(size = 8)
  )

# Extract just the legend
legend_grob <- cowplot::get_legend(legend_plot)


#
# Export plot
#

cat("\nExporting plot...\n")

# Figure dimensions: 15 panels × 2 cm + margin for labels ≈ 35 cm

output_filename_cor <- paste0(date2, "_ASM_TA_correlation_matrix.png")

# Save the main matrix

ggsave(
  filename = output_filename_cor,
  plot = cor_matrix,
  width = 35,
  height = 35,
  dpi = 300,
  units = "cm"
)

cat(sprintf("Correlation matrix saved as: %s\n", output_filename_cor))


# Save legend separately (can be combined in presentation software)

output_filename_legend <- paste0(date2, "_ASM_TA_correlation_matrix_legend.png")

ggsave(
  filename = output_filename_legend,
  plot = legend_grob,
  width = 10,
  height = 5,
  dpi = 300,
  units = "cm"
)

cat(sprintf("Legend saved as: %s\n", output_filename_legend))

cat("\nSignificance levels: * p<0.05, ** p<0.01, *** p<0.001\n")
cat("Done!\n")