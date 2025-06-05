import os
import pandas as pd
import xlwings as xw
from utils import readexcel

def process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level):

    combined_df = pd.DataFrame()

    for filename in os.listdir(folder_path):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            print(f"\nProcessing file: {filename}")

            for sheet_name in sheet_names:
                print(f"  -> Reading sheet: {sheet_name}")

                df = readexcel.read_sheet_with_xlwings(file_path, sheet_name)

                if df.empty:
                    print(f"  !! Skipping {sheet_name} — sheet is empty or failed to load.")
                    continue

                try:
                    year_cell = str(df.iat[0, 0])
                    year = year_cell.split(",")[-1].strip()

                    gender_cell = str(df.iat[2, 0])
                    gender = gender_cell.split()[0].capitalize()

                    df_clean = df.iloc[3:, [0, 1]].copy()
                    df_clean.columns = ['Subject', 'Entries']
                    df_clean.dropna(how='all', inplace=True)

                    df_filtered = df_clean[df_clean['Subject'].isin(subjects_to_include)]

                    df_filtered['Gender'] = gender
                    df_filtered['Year'] = year
                    df_filtered['Level'] = level

                    combined_df = pd.concat([combined_df, df_filtered], ignore_index=True)

                except Exception as e:
                    print(f"  !! Error parsing data from {sheet_name} in {filename}: {e}")

    return combined_df

