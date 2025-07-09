import os
import pandas as pd
from utils import processing, cache
from model.sheet_config import SheetConfig

# === CONFIGURATION ===

OUTPUT_FILE = "combined_output.csv"
OUTPUT_DIR = "output"
SUBJECT_OUTPUT_FILE = "subject_names.csv"

exclude_keywords = [
    "Awards in the optional Writing Element for Gaelic (Learners) are made at grades 1 to 4 only.",
    "Awards in the optional Writing elements for Modern Languages and Gaelic (Learners) are made at grades 1 to 4 only.",
    "as percentages",
    "FEMALE CANDIDATES",
    "FEMALE LEARNERS",
    "MALE CANDIDATES",
    "MALE LEARNERS",
    "SUBJECT",
    "Subtotal",
    "Subtotals",
    "TITLE",
    "Total",
    "Totals"
]

sheet_configs = [
    
    SheetConfig("Standard_Grade", r"data\Standard Grade\2000-2002", ["SG4a", "SG4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Standard_Grade", r"data\Standard Grade\2003-2009", ["SG5a", "SG5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Standard_Grade", r"data\Standard Grade\2010-2010", ["SG5a", "SG5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Standard_Grade", r"data\Standard Grade\2011-2013", ["SG5a", "SG5b"], gender_cell="A6", year_cell="A3"),
    
    SheetConfig("Intermediate_2", r"data\Intermediate 2\2000-2002", ["IB4a", "IB4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Intermediate_2", r"data\Intermediate 2\2003-2009", ["IB5a","IB5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Intermediate_2", r"data\Intermediate 2\2010-2010", ["IB5a","IB5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Intermediate_2", r"data\Intermediate 2\2011-2013", ["IB5a","IB5b"], gender_cell="A6", year_cell="A3"),
    
    SheetConfig("Intermediate_1", r"data\Intermediate 1\2000-2002", ["IA4a", "IA4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Intermediate_1", r"data\Intermediate 1\2003-2009", ["IA5a", "IA5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Intermediate_1", r"data\Intermediate 1\2010-2010", ["IA5a", "IA5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Intermediate_1", r"data\Intermediate 1\2011-2013", ["IA5a", "IA5b"], gender_cell="A6", year_cell="A3"),
    
    SheetConfig("Higher", r"data\Higher\2000-2002", ["NH4a", "NH4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2003-2009", ["NH5a", "NH5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2010-2010", ["NH5a", "NH5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2011-2013", ["NH5a", "NH5b"], gender_cell="A6", year_cell="A3"),
    
    SheetConfig("CSYS", r"data\Advanced Higher\2000-2000", ["CS3a", "CS3b"], gender_cell="A4", year_cell="A1"),
    
    SheetConfig("Advanced_Higher", r"data\Advanced Higher\2001-2002", ["AH4a", "AH4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Advanced_Higher", r"data\Advanced Higher\2003-2009", ["AH5a", "AH5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Advanced_Higher", r"data\Advanced Higher\2010-2010", ["AH5a", "AH5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Advanced_Higher", r"data\Advanced Higher\2011-2013", ["AH5a", "AH5b"], gender_cell="A6", year_cell="A3"),
]

# === MAIN EXECUTION ===

processed_cache = cache.load_cache()
combined_df = pd.DataFrame()

for config in sheet_configs:
    df = processing.process_file(
        config, 
        processed_cache, 
        OUTPUT_DIR,
        exclude_keywords
    )
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# Append to existing file if needed
if os.path.exists(OUTPUT_FILE):
    try:
        existing_df = pd.read_csv(OUTPUT_FILE)
        combined_df = pd.concat([existing_df, combined_df], ignore_index=True)
    except pd.errors.EmptyDataError:
        print(f"⚠️  Output file '{OUTPUT_FILE}' exists but is empty.")

# De-duplicate and save
# Clean up for deduplication
combined_df["Subject"] = combined_df["Subject"].astype(str).str.strip()
combined_df["Level"] = combined_df["Level"].astype(str).str.strip()
combined_df["Gender"] = combined_df["Gender"].astype(str).str.strip()
combined_df["Entries"] = pd.to_numeric(combined_df["Entries"], errors='coerce')

# Drop duplicates across the key columns
combined_df.drop_duplicates(
    subset=["Year", "Gender", "Subject", "Level", "Entries"],
    inplace=True
)

combined_df.to_csv(OUTPUT_FILE, index=False)
print("✅ Final combined data saved to:", OUTPUT_FILE)

# Save cache and subjects
cache.save_cache(processed_cache)
processing.save_subject_names(SUBJECT_OUTPUT_FILE)
print("📘 Subject names saved to:", SUBJECT_OUTPUT_FILE)
