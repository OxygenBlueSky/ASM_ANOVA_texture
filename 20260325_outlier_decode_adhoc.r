# Read outliers
outliers <- read.csv2("20260108_ASM_outliers.csv")  # semicolon-separated

# Read decoded data
df <- read.delim("20250910_ASM_1_10_texture_data_DECODED.csv")

# Merge to get decoded info for each outlier plate
outlier_info <- merge(outliers, df, by = c("series_name", "nr_in_chamber"))

# Keep just the identifying columns (one row per plate)
outlier_summary <- unique(outlier_info[, c("series_name", "nr_in_chamber", 
                                           "experiment_number", "experiment_name",
                                           "potency", "decoded_solution", "pair",
                                           "verum", "SNC")])

outlier_summary[order(outlier_summary$experiment_number, 
                      outlier_summary$nr_in_chamber), ]