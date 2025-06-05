from utils import processing
import pandas as pd

# === OVERALL CONFIG === 
subjects_to_include = [
    "Biology", "Chemistry", "Computing",
    "English and Communication", "French",
    "Mathematics", "Physics"
]

combined_df = pd.DataFrame()

# === SHEET CONFIG ===

# = STD GRADE =
level = "Standard Grade"

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Standard Grade\SG"
sheet_names = ["SG5a", "SG5b"]  # Process both sheets
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Standard Grade\IM2"
sheet_names = ["IB4a", "IB4b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

# = HIGHER =
level = "Higher"

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Higher\2000-2002"
sheet_names =["NH4a", "NH4b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Higher\2003-2005"
sheet_names = ["NH5a", "NH5b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

# = ADV HIGHER =
level = "Advanced Higher"

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Advanced Higher\2000-2000"
sheet_names = ["CS3a", "CS3b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Advanced Higher\2001-2002"
sheet_names = ["AH4a", "AH4b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

folder_path = r"C:\Users\scott\OneDrive\Employment\JP Morgan\dressCode\Raw Data 2000 to 2005\Advanced Higher\2003-2005"
sheet_names = ["AH5a", "AH5b"]
df = processing.process_sheet_with_params(folder_path, sheet_names, subjects_to_include, level)
combined_df = pd.concat([combined_df, df], ignore_index=True)

# === OUTPUT ===
print("\n✅ Final Combined Data:")
print(combined_df)
combined_df.to_csv("combined_output.csv", index=False)

