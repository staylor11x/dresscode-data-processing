import os
import pandas as pd
from utils import processing, cache
from sheet_config import SheetConfig

# === OVERALL CONFIG ===
subjects_to_include = [
    "Biology", "Chemistry", "Computing",
    "English and Communication", "French",
    "Mathematics", "Physics","English"
]

OUTPUT_FILE = "combined_output.csv"
processed_cache = cache.load_cache()
combined_df = pd.DataFrame()

# === SHEET CONFIG ===
# sheet_configs = [
#     {"level": "Standard Grade", "folder": r"data\Standard Grade\SG", "sheets": ["SG5a", "SG5b"] },
#     {"level": "Standard Grade", "folder": r"data\Standard Grade\IM2", "sheets": ["IB4a", "IB4b"]},
#     {"level": "Higher", "folder": r"data\Higher\2000-2002", "sheets": ["NH4a", "NH4b"]},
#     {"level": "Higher", "folder": r"data\Higher\2003-2010", "sheets": ["NH5a", "NH5b"]},
#     {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2000-2000", "sheets": ["CS3a", "CS3b"]},
#     {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2001-2002", "sheets": ["AH4a", "AH4b"]},
#     {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2003-2010", "sheets": ["AH5a", "AH5b"]},
# ]

sheet_configs = [
    SheetConfig("Standard Grade", r"data\Standard Grade\SG", ["SG5a", "SG5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Standard Grade", r"data\Standard Grade\IM2", ["IB4a", "IB4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Standard Grade", r"data\Standard Grade\2010-2010", ["SG5a", "SG5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2000-2002", ["NH4a", "NH4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2003-2009", ["NH5a", "NH5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Higher", r"data\Higher\2010-2010", ["NH5a", "NH5b"], gender_cell="A4", year_cell="A1"),
    SheetConfig("Advanced Higher", r"data\Advanced Higher\2000-2000", ["CS3a", "CS3b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Advanced Higher", r"data\Advanced Higher\2001-2002", ["AH4a", "AH4b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Advanced Higher", r"data\Advanced Higher\2003-2009", ["AH5a", "AH5b"], gender_cell="A3", year_cell="A1"),
    SheetConfig("Advanced Higher", r"data\Advanced Higher\2010-2010", ["AH5a", "AH5b"], gender_cell="A4", year_cell="A1"),
]


# === PROCESSING ===
for config in sheet_configs:
    df = processing.process_sheet_with_params(
        config,
        subjects_to_include,
        processed_cache
    )
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# === APPEND TO EXISTING OUTPUT (IF ANY) ===
if os.path.exists(OUTPUT_FILE):
    try:
        existing_df = pd.read_csv(OUTPUT_FILE)
        combined_df = pd.concat([existing_df, combined_df], ignore_index=True)
    except pd.errors.EmptyDataError:
        print(f"⚠️  Output file '{OUTPUT_FILE}' exists but is empty. Proceeding with new data only.")


# === OPTIONAL: DE-DUPLICATE ===
combined_df.drop_duplicates(
    subset=["Year", "Gender", "Subject", "Level", "Entries"],
    inplace=True
)

# === SAVE OUTPUT ===
combined_df.to_csv(OUTPUT_FILE, index=False)
print("\n✅ Final combined data saved to:", OUTPUT_FILE)

# === SAVE CACHE ===
cache.save_cache(processed_cache)
