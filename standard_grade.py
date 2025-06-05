import os
import pandas as pd
import xlwings as xw

# === CONFIGURATION ===
folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Standard Grade\SG"
sheet_names = ["SG5a", "SG5b"]  # Process both sheets
subjects_to_include = [
    "Biology", "Chemistry", "Computing",
    "English and Communication", "French",
    "Mathematics", "Physics"
]

def read_sheet_with_xlwings(file_path, sheet_name):
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        sht = wb.sheets[sheet_name]

        data = sht.used_range.value
        wb.close()
        app.quit()

        return pd.DataFrame(data)
    except Exception as e:
        print(f"Failed to read {sheet_name} in {file_path} with xlwings: {e}")
        return pd.DataFrame()

# === MAIN PROCESSING ===
combined_df = pd.DataFrame()

for filename in os.listdir(folder_path):
    if filename.endswith(".xls") or filename.endswith(".xlsx"):
        file_path = os.path.join(folder_path, filename)
        print(f"\nProcessing file: {filename}")

        for sheet_name in sheet_names:
            print(f"  -> Reading sheet: {sheet_name}")

            df = read_sheet_with_xlwings(file_path, sheet_name)

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

                combined_df = pd.concat([combined_df, df_filtered], ignore_index=True)

            except Exception as e:
                print(f"  !! Error parsing data from {sheet_name} in {filename}: {e}")

# === OUTPUT ===
print("\n✅ Final Combined Data:")
print(combined_df)
# combined_df.to_csv("combined_output.csv", index=False)
