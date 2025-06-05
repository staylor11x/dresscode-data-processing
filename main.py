import pandas as pd
from utils import processing

# === CONFIGURATION ===
subjects_to_include = [
    "Biology", "Chemistry", "Computing",
    "English and Communication", "French",
    "Mathematics", "Physics"
]

sheet_configs = [
    {"level": "Standard Grade", "folder": r"data\Standard Grade\SG", "sheets": ["SG5a", "SG5b"]},
    {"level": "Standard Grade", "folder": r"data\Standard Grade\IM2", "sheets": ["IB4a", "IB4b"]},
    {"level": "Higher", "folder": r"data\Higher\2000-2002", "sheets": ["NH4a", "NH4b"]},
    {"level": "Higher", "folder": r"data\Higher\2003-2005", "sheets": ["NH5a", "NH5b"]},
    {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2000-2000", "sheets": ["CS3a", "CS3b"]},
    {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2001-2002", "sheets": ["AH4a", "AH4b"]},
    {"level": "Advanced Higher", "folder": r"data\Advanced Higher\2003-2005", "sheets": ["AH5a", "AH5b"]},
]

# === PROCESSING ===
combined_df = pd.DataFrame()

for config in sheet_configs:
    df = processing.process_sheet_with_params(
        config["folder"],
        config["sheets"],
        subjects_to_include,
        config["level"]
    )
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# === OUTPUT ===
print("\n✅ Final Combined Data:")
print(combined_df)
combined_df.to_csv("combined_output.csv", index=False)
