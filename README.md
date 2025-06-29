# 📊 Scottish Qualification Data Processor

This project extracts, cleans, and aggregates subject-level entry data from a variety of Excel sources representing different Scottish qualification levels (e.g. Intermediate 2, Higher, Advanced Higher). The output is saved both as a combined CSV and as individual Excel files per level and year.

## 🧰 Features

- ✅ Reads and processes multiple Excel sheets using `xlwings`
- ✅ Extracts metadata like **year** and **gender** from configurable cells
- ✅ Filters out non-subject rows using keyword exclusion
- ✅ Saves:
  - A combined `.csv` file of all data (`combined_output.csv`)
  - Per-level Excel files like `Intermediate_2_2000.xlsx`
- ✅ Tracks processed files with a local cache to avoid duplication
- ✅ Exports a list of unique subject names (`subject_names.csv`)
- ✅ Robust deduplication logic for output


## ⚙️ Requirements

- Python 3.9+
- Excel (required by `xlwings`)
- Pip packages:

```bash
pip install pandas xlwings openpyxl
```

Or
```bash
pip install requirements.txt
```

## 🚀 How to Run

1. Place Excel files into the appropriate subfolders under ``data/``.

2. Configure the sheets and cell locations in `main.py` via SheetConfig.

3. Run the script:

```bash
python main.py
```

4. Check ``output/`` and root directory for the result files.

## 📝 Configuration: SheetConfig

Each Excel source is defined with:

```python
SheetConfig(
    level="Intermediate_2",
    folder="data/Standard Grade/IM2",
    sheets=["IB4a", "IB4b"],
    gender_cell="A3",
    year_cell="A1"
)
```

You can easily extend this by uncommenting or adding more configs in the `sheet_configs` list in `main.py`.

## 🔍 Exclusion Logic

```python
exclude_keywords = [
    "FEMALE CANDIDATES", "TOTAL", "SUBJECT", "Subtotal", ...
]
```

These are checked case-insensitively against the "Subject" column.

## 🧼 Deduplication

The output CSV (combined_output.csv) is de-duplicated on:

- Year
- Gender
- Subject
- Level
- Entries

This ensures consistent aggregation across re-runs or expanded source data.

## 🧠 Notes

- Make sure Excel is installed and available on your machine (required by xlwings)
- Output files are replaced, not appended — no duplicates will accumulate
- Processed files are tracked using a simple JSON cache (``processed_log.json``)

## 👤 Author
Scott Taylor — ChatGPT & automation enthusiast