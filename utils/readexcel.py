import pandas as pd
import xlwings as xw

def read_sheet_with_xlwings(file_path, sheet_name):
    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            sht = wb.sheets[sheet_name]
            data = sht.used_range.value
            wb.close()

        return pd.DataFrame(data)

    except Exception as e:
        print(f"Failed to read {sheet_name} in {file_path} with xlwings: {e}")
        return pd.DataFrame()
