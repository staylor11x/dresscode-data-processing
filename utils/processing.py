import os
import pandas as pd
from utils import readexcel, cache

ALL_SUBJECT_NAMES = set()

def extract_metadata(df, year_coords, gender_coords):
    import re
    year_cell = str(df.iat[year_coords[0], year_coords[1]])
    match = re.search(r"\b(19|20)\d{2}\b", year_cell)
    year = int(match.group(0)) if match else None

    gender_cell = str(df.iat[gender_coords[0], gender_coords[1]])
    gender = gender_cell.split()[0].capitalize()
    
    return year, gender

def is_valid_subject(subject, exclude_keywords):
    if not isinstance(subject, str):
        return False
    subject = subject.strip().lower()
    return not any(keyword.lower() in subject for keyword in exclude_keywords)


def save_subject_names(filepath="subject_names.csv"):
    df = pd.DataFrame(sorted(ALL_SUBJECT_NAMES), columns=["Subject"])
    df.to_csv(filepath, index=False)

def process_file(config, processed_cache, output_dir, exclude_keywords):
    overall_output_df = pd.DataFrame()
    
    for filename in os.listdir(config.folder):
        if not filename.endswith((".xls", ".xlsx")):
            continue

        file_path = os.path.join(config.folder, filename)
        if cache.was_processed(processed_cache, filename):
            print(f"🔄 Skipping already processed file: {filename}")
            continue

        print(f"\n📄 Processing file: {filename}")
        processed_df, year = process_sheet(config, file_path, filename, exclude_keywords)
        overall_output_df = pd.concat([overall_output_df, processed_df], ignore_index=True)
        
        if not processed_df.empty:
            save_to_file(processed_df, config, output_dir, year)

        cache.mark_processed(processed_cache, filename)

    return overall_output_df

def process_sheet(config, file_path, filename, exclude_keywords):
    combined_df = pd.DataFrame()
    
    for sheet_name in config.sheets:
        print(f"📄 Reading sheet: {sheet_name}")
        df = readexcel.read_sheet_with_xlwings(file_path, sheet_name)
        if df.empty:
            print(f"⚠️  Skipping {sheet_name} — empty or failed to load.")
            continue

        try:
            year, gender = extract_metadata(df, config.year_coords, config.gender_coords)
            
            df_filtered = (
                df.iloc[3:, [0, 1]]
                .dropna(how="all")
                .rename(columns={0: "Subject", 1: "Entries"})
            )

            subjects = df_filtered["Subject"].astype(str).str.strip().tolist()
            ALL_SUBJECT_NAMES.update(subjects)

            df_filtered = df_filtered[df_filtered["Subject"].apply(lambda x: is_valid_subject(x, exclude_keywords))]

            df_filtered["Gender"] = gender
            df_filtered["Year"] = year
            df_filtered["Qualification"] = config.qualification

            combined_df = pd.concat([combined_df, df_filtered], ignore_index=True)

        except Exception as e:
            print(f"❌ Error in {sheet_name} ({filename}): {e}")

    return combined_df, year


def save_to_file(df, c, output_dir, year):
   
    # Check output file exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Construct output file name
    safe_qualification = c.qualification.replace(" ", "_")
    output_filename = f"{safe_qualification}_{year}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    # Always overwrite the file
    df.to_excel(output_path, index=False)
    print(f"✅ Saved to {output_filename}")

