# ASM Texture and Fractal Data Decoding Script
# Created: 2026-01-14
# Purpose: Decode ASM texture and fractal data and export with decoding information
# Based on: 260109_ANOVA_ASM_posthoc_cohensd.r

# Load required packages -------
library(readxl)    # For reading Excel files (if needed)
library(here)      # For file path management
library(dplyr)     # For data manipulation


# ============================================================================
# PART 1: TEXTURE ANALYSIS DATA
# ============================================================================

cat("\n")
cat("################################################################################\n")
cat("TEXTURE ANALYSIS DATA PROCESSING\n")
cat("################################################################################\n\n")

# Data files with Texture data from ACIA --------

df <- read.delim(here("20250910_ASM_1_10_texture_data.csv"))


# Hard coding on ASM data
# Filter df for scale and separate out experiment number and potency codes
# AS and JZ named the series differently
# Using column "substance_group" for name extraction and "series_name" for experiment number

# Filter to keep only scale = 1 data
# The square brackets [rows, columns] allow us to subset the dataframe
# We're selecting rows where scale equals 1, and keeping all columns
df_scale1 <- df[df$scale == 1, ]

# Define the mapping for AS series
# This is a named vector: the name "EI" maps to value 2, "HOM3_ASM1" maps to 1
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
df_TA <- df_scale1 %>%
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
df_TA <- df_TA %>%
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
df_TA <- df_TA %>%
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

# Join with df_TA to add decoded columns
# left_join() matches rows based on specified columns (experiment_number and potency)
# It adds the decoded_solution and pair columns from full_mapping to df_TA
# "left" join means: keep all rows from df_TA, add matching info from full_mapping
df_TA <- df_TA %>%
  left_join(full_mapping, by = c("experiment_number", "potency"))

# Verify decoding was successful
cat("Decoding verification:\n")
cat("Total rows in df_TA:", nrow(df_TA), "\n")
cat("Rows with decoded_solution:", sum(!is.na(df_TA$decoded_solution)), "\n")
cat("Rows with pair assignment:", sum(!is.na(df_TA$pair)), "\n\n")

# Check if all rows were decoded
# is.na() checks for missing values (NA = Not Available)
# any() returns TRUE if at least one value is TRUE
if (any(is.na(df_TA$decoded_solution))) {
  cat("WARNING: Some rows were not decoded!\n")
  cat("Rows without decoding:\n")
  # Filter to show only rows that failed to decode
  # select() chooses which columns to display
  # distinct() removes duplicate rows
  print(df_TA %>%
          filter(is.na(decoded_solution)) %>%
          select(experiment_number, potency, substance_group) %>%
          distinct())
  cat("\n")
}


# Data integrity check

cat("\n")
cat("========================================================================\n")
cat("DATA INTEGRITY CHECK - TEXTURE ANALYSIS\n")
cat("========================================================================\n")
cat("Total observations in df_TA:", nrow(df_TA), "\n")
cat("Expected: 10 experiments × 6 potencies × 7 replicates = 420 observations\n\n")

# Check counts by experiment_number
cat("--- Observations per experiment_number ---\n")
exp_counts <- table(df_TA$experiment_number)
print(exp_counts)
cat("\n")

# Check counts by experiment_number and decoded_solution
# This table shows which solutions were tested in each experiment
# SNC experiments (1,3,6,8,9) should only have water_1, water_2, water_3
# Verum experiments (2,4,5,7,10) should only have potency_x, potency_y, potency_z
cat("--- Observations per experiment_number × decoded_solution ---\n")
exp_sol_counts <- table(df_TA$experiment_number, df_TA$decoded_solution)
print(exp_sol_counts)
cat("\n")

cat("========================================================================\n\n")


# Create final columns for export - Texture Analysis

df_TA_export <- df_TA 

# Export Texture Analysis to CSV

# Create output filename by adding _DECODED to the input filename
input_filename_TA <- "20250910_ASM_1_10_texture_data.csv"
output_filename_TA <- sub("\\.csv$", "_DECODED.csv", input_filename_TA)

# Write the decoded dataframe to tab-delimited CSV
write.table(df_TA_export, output_filename_TA, row.names = FALSE, sep = "\t", quote = FALSE)

cat("\n")
cat("========================================================================\n")
cat("TEXTURE ANALYSIS EXPORT COMPLETE\n")
cat("========================================================================\n")
cat(sprintf("File exported: %s\n", output_filename_TA))
cat(sprintf("Total rows: %d\n", nrow(df_TA_export)))
cat(sprintf("Total columns: %d\n", ncol(df_TA_export)))
cat("\nDecoded columns in Texture Analysis data:\n")
cat("  - experiment_name (ASM_AS or ASM_JZ)\n")
cat("  - experiment_number (1-10)\n")
cat("  - potency (A-F coded letters)\n")
cat("  - verum (0 or 1)\n")
cat("  - SNC (0 or 1)\n")
cat("  - sub_exp_number (1-5)\n")
cat("  - decoded_solution (water_1/2/3 or potency_x/y/z)\n")
cat("  - pair (1 or 2)\n")
cat("========================================================================\n\n")


# ============================================================================
# PART 2: FRACTAL ANALYSIS DATA
# ============================================================================

cat("\n")
cat("################################################################################\n")
cat("FRACTAL ANALYSIS DATA PROCESSING\n")
cat("################################################################################\n\n")

# Load Fractal data from FracLac analysis --------

df_FA <- read.delim(here("20251218_ASM_1_10_fractal_data.csv"))

cat("Fractal data loaded.\n")
cat("Total rows:", nrow(df_FA), "\n")
cat("Total columns:", ncol(df_FA), "\n\n")


# Parse filename column to extract series_name and nr_in_chamber
#
# The first column contains filenames with this pattern:
# [H or P][8 or 9-digit date][series_name]-[nr_in_chamber]-[rest of filename]
#
# Note: Some series have 8-digit dates (EI, HOM3_ASM1) while others have 9-digit dates (0GER, 0GEW, etc.)
#
# Examples:
# - H120231220GER-01-20231222_10_42cuttifŞ1_ (0, 0_2112x2112) -6,-6---6,-6  (9-digit date)
# - P20170123EI-16-20170125_11_45cuttifŞ1_ (0, 0_2112x2112) -6,-6---6,-6   (8-digit date)
# - P20170116HOM3_ASM1-43-20170118_13_50cuttifŞ1_ (0, 0_2112x2112) -6,-6---6,-6  (8-digit date)
#

cat("Parsing filename to extract series_name and nr_in_chamber...\n\n")

# Get the name of the first column
first_col_name <- colnames(df_FA)[1]

# Create parsing function for fractal filenames
parse_fractal_filename <- function(filename) {
  # Pattern: [H or P][8 or 9 digits][series_name]-[number]-[rest]
  # The 8 or 9 comes from the chamber name being "P" and "H1"
  # Some series have 8-digit dates (EI, HOM3_ASM1) and some have "1" + 8-digit dates (GER series)
  # Regex: ^[HP]\d{8,9}([A-Z0-9_]+)-(\d+)-

  pattern <- "^[HP]\\d{8,9}([A-Z0-9_]+)-(\\d+)-"
  matches <- regmatches(filename, regexec(pattern, filename))[[1]]

  if (length(matches) == 3) {
    return(list(
      series_name = matches[2],           # The series name (any alphanumeric combination)
      nr_in_chamber = as.integer(matches[3])  # The number (converted to integer)
    ))
  } else {
    # If pattern doesn't match, return NA
    return(list(
      series_name = NA_character_,
      nr_in_chamber = NA_integer_
    ))
  }
}

# Apply parsing to fractal dataframe
df_FA <- df_FA %>%
  rowwise() %>%
  mutate(
    # Parse the first column (filename)
    parsed = list(parse_fractal_filename(.data[[first_col_name]])),
    # Extract series_name and nr_in_chamber
    series_name = parsed$series_name,
    nr_in_chamber = parsed$nr_in_chamber
  ) %>%
  select(-parsed) %>%  # Remove temporary parsed column
  ungroup()

# Verify parsing was successful
cat("Parsing verification:\n")
cat("Total rows in df_FA:", nrow(df_FA), "\n")
cat("Rows with series_name:", sum(!is.na(df_FA$series_name)), "\n")
cat("Rows with nr_in_chamber:", sum(!is.na(df_FA$nr_in_chamber)), "\n\n")

# Check if all rows were parsed
if (any(is.na(df_FA$series_name))) {
  cat("WARNING: Some rows were not parsed!\n")
  cat("Rows without parsing:\n")
  print(df_FA %>%
          filter(is.na(series_name)) %>%
          select(1) %>%  # Show first column (filename)
          distinct())
  cat("\n")
}


# Data integrity check - Fractal Analysis

cat("\n")
cat("========================================================================\n")
cat("DATA INTEGRITY CHECK - FRACTAL ANALYSIS\n")
cat("========================================================================\n")

# Check counts by series_name
cat("--- Observations per series_name ---\n")
series_counts <- table(df_FA$series_name)
print(series_counts)
cat("\n")

# Check nr_in_chamber range for each series
cat("--- Nr_in_chamber range per series_name ---\n")
for (series in unique(df_FA$series_name)) {
  # Skip NA series
  if (is.na(series)) {
    next
  }
  series_data <- df_FA %>% filter(series_name == series)
  nr_range <- range(series_data$nr_in_chamber, na.rm = TRUE)
  cat(sprintf("%-12s: %2d to %2d (n = %d)\n",
              series, nr_range[1], nr_range[2], nrow(series_data)))
}
cat("\n")

cat("========================================================================\n\n")


# Join Fractal data with Texture decoding information

cat("\n")
cat("========================================================================\n")
cat("JOINING FRACTAL DATA WITH TEXTURE DECODING\n")
cat("========================================================================\n\n")

cat("Joining fractal data with texture decoding using series_name and nr_in_chamber...\n\n")

# Select only the decoding columns from texture data
texture_decoding <- df_TA %>%
  select(series_name, nr_in_chamber,
         experiment_name, experiment_number, potency,
         verum, SNC, sub_exp_number,
         decoded_solution, pair) %>%
  distinct()  # Remove duplicates if any

# Left join: keep all fractal rows, add decoding where available
df_FA_decoded <- df_FA %>%
  left_join(texture_decoding, by = c("series_name", "nr_in_chamber"))

# Check how many rows matched
n_matched <- sum(!is.na(df_FA_decoded$experiment_number))
n_unmatched <- sum(is.na(df_FA_decoded$experiment_number))

cat(sprintf("Fractal rows matched with texture decoding: %d\n", n_matched))
cat(sprintf("Fractal rows WITHOUT texture decoding: %d\n\n", n_unmatched))

# Report unmatched rows and delete them
if (n_unmatched > 0) {
  cat("========================================================================\n")
  cat("FRACTAL ROWS WITHOUT TEXTURE DECODING\n")
  cat("========================================================================\n\n")

  unmatched_rows <- df_FA_decoded %>%
    filter(is.na(experiment_number)) %>%
    select(series_name, nr_in_chamber) %>%
    distinct() %>%
    arrange(series_name, nr_in_chamber)

  cat("Series_name | Nr_in_chamber\n")
  cat("------------|---------------\n")
  for (i in 1:nrow(unmatched_rows)) {
    cat(sprintf("%-11s | %d\n",
                unmatched_rows$series_name[i],
                unmatched_rows$nr_in_chamber[i]))
  }
  cat("\n")
  cat(sprintf("Total unique unmatched combinations: %d\n", nrow(unmatched_rows)))
  cat("\nNon-matching rows are presumed to be the reference plates and deleted.\n")
  cat("========================================================================\n\n")

  # Delete unmatched rows from df_FA_decoded
  df_FA_decoded <- df_FA_decoded %>%
    filter(!is.na(experiment_number))

  cat(sprintf("Rows remaining after deletion: %d\n\n", nrow(df_FA_decoded)))
}

cat("Decoding columns added to fractal data:\n")
cat("  - experiment_name\n")
cat("  - experiment_number\n")
cat("  - potency\n")
cat("  - verum\n")
cat("  - SNC\n")
cat("  - sub_exp_number\n")
cat("  - decoded_solution\n")
cat("  - pair\n\n")


# Export Fractal Analysis to CSV

# Create output filename by adding _DECODED to the input filename
input_filename_FA <- "20251218_ASM_1_10_fractal_data.csv"
output_filename_FA <- sub("\\.csv$", "_DECODED.csv", input_filename_FA)

# Write the decoded dataframe to tab-delimited CSV
write.table(df_FA_decoded, output_filename_FA, row.names = FALSE, sep = "\t", quote = FALSE)

cat("\n")
cat("========================================================================\n")
cat("FRACTAL ANALYSIS EXPORT COMPLETE\n")
cat("========================================================================\n")
cat(sprintf("File exported: %s\n", output_filename_FA))
cat(sprintf("Total rows: %d\n", nrow(df_FA_decoded)))
cat(sprintf("Total columns: %d\n", ncol(df_FA_decoded)))
cat("\nNew parsed columns in Fractal Analysis data:\n")
cat("  - series_name (e.g., GER, EI, HOM3_ASM1)\n")
cat("  - nr_in_chamber (integer, range varies by series)\n")
cat("\nDecoding columns joined from Texture Analysis:\n")
cat("  - experiment_name, experiment_number, potency\n")
cat("  - verum, SNC, sub_exp_number\n")
cat("  - decoded_solution, pair\n")
cat("========================================================================\n\n")


# ============================================================================
# FINAL SUMMARY
# ============================================================================

cat("\n")
cat("################################################################################\n")
cat("ALL PROCESSING COMPLETE\n")
cat("################################################################################\n\n")
cat("Files exported:\n")
cat(sprintf("  1. %s (%d rows, %d columns)\n",
            output_filename_TA, nrow(df_TA_export), ncol(df_TA_export)))
cat(sprintf("  2. %s (%d rows, %d columns)\n",
            output_filename_FA, nrow(df_FA_decoded), ncol(df_FA_decoded)))
cat("\n")
cat("Done!\n")
