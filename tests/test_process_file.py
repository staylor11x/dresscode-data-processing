# === tests/test_process_file.py ===
import os
from model.sheet_config import SheetConfig
from utils import processing, cache


def test_process_file_integration(tmp_path):
    # Arrange
    test_data_dir = os.path.join(os.path.dirname(__file__), "test_data")
    config = SheetConfig(
        qualification="Test_Level",
        folder=test_data_dir,
        sheets=["Sheet1"],
        gender_cell="A3",
        year_cell="A1"
    )
    test_cache = {}
    exclude_keywords = ["total"]

    # Act
    df = processing.process_file(config, test_cache, tmp_path, exclude_keywords)

    # Assert
    assert not df.empty
    assert "Subject" in df.columns
    output_files = list(tmp_path.glob("*.xlsx"))
    assert len(output_files) == 1
    assert output_files[0].name.startswith("Test_Level_")