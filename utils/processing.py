import os
import pandas as pd
from utils import readexcel
from utils import cache

SUBJECT_NAMES = []

def save_subject_names(s, filepath="subject_names.csv"):
    # De-duplicate the list
    unique_subjects = sorted(set(s))  # sorted is optional, just for neatness

    # Save to CSV using pandas
    df = pd.DataFrame(unique_subjects, columns=["Subject"])
    df.to_csv(filepath, index=False)


# Takes in one excel file and returns the processed output of all the included sheets
def process_file(c, processed_cache, output_dir, rows_to_exclude):

    overall_output_df = pd.DataFrame()
    
    for filename in os.listdir(c.folder):
        combined_df = pd.DataFrame()
        if filename.endswith((".xls", "xlsx")):
            file_path = os.path.join(c.folder, filename)

            if cache.was_processed(processed_cache, file_path):
                print(f"🔄 Skipping already processed file: {filename}")
                continue
            
            print(f"\nProcessing file: {filename}")
            processed_sheet, year = process_sheet(c, file_path, filename, rows_to_exclude)
            
            combined_df = pd.concat([combined_df, processed_sheet], ignore_index=True)
            overall_output_df = pd.concat([combined_df, processed_sheet], ignore_index=True)
        # save_to_file(combined_df, c, output_dir, year)
        cache.mark_processed(processed_cache, file_path)

    return overall_output_df

# Takes in one sheet and returns the processed output
def process_sheet(c, file_path, filename, rows_to_exclude):
    
    combined_df = pd.DataFrame()    # male and females
    
    for sheet_name in c.sheets:
        print(f"Reading sheet: {sheet_name}")
        df = readexcel.read_sheet_with_xlwings(file_path, sheet_name)
        if df.empty:
            print(f"Skipping {sheet_name} — sheet is empty or failed to load.")
            continue

        try: 
            year_cell = str(df.iat[c.year_coords[0], c.year_coords[1]])
            year = year_cell.split(",")[-1].strip()
            year_cell = str(df.iat[c.year_coords[0], c.year_coords[1]])
            year_text = year_cell.split(",")[-1].strip()
            
            # Ensure the year is a 4-digit number
            import re
            match = re.search(r"\b(19|20)\d{2}\b", year_text)
            if match:
                year = int(match.group(0))
            else:
                raise ValueError(f"Could not extract valid year from cell: '{year_cell}'")
            
            gender_cell = str(df.iat[c.gender_coords[0], c.gender_coords[1]])
            gender = gender_cell.split()[0].capitalize()
            
            df_filtered = (
                df.iloc[3:, [0, 1]]
                .dropna(how="all")
                .rename(columns={0: "Subject", 1: "Entries"})
            )
            s = df_filtered["Subject"].tolist()
            save_subject_names(s)

            # drop rows that are not subjects
            df_filtered = df_filtered[~df_filtered["Subject"].isin(rows_to_exclude)]
            
            
            df_filtered["Gender"] = gender
            df_filtered["Year"] = year
            df_filtered["Level"] = c.level
            
            combined_df = pd.concat([combined_df, df_filtered], ignore_index=True)
        
        except Exception as e:
            print(f"Error parsing {sheet_name} in {filename}: {e}")
    
    return combined_df, year

# Saves the output to a file in the format `LEVEL_YEAR_xlsx` i.e `Intermediate_2_2000.xlsx`
def save_to_file(df, c, output_dir, year):
    
    # Construct output file name
    safe_level = c.level.replace(" ", "_")
    output_filename = f"{safe_level}_{year}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    # Save or append to file
    if os.path.exists(output_path):
        existing_df = pd.read_excel(output_path)
        combined = pd.concat([existing_df, df], ignore_index=True)
    else:
        combined = df
        combined.to_excel(output_path, index=False)
        print(f"✅ Saved to {output_filename}")
    