import pandas as pd
from utils import processing

def test_extract_metadata():
    df = pd.DataFrame([
        ["Year: 2001"],
        [None],
        ["Male candidates"]
    ])
    year, gender = processing.extract_metadata(df, (0, 0), (2, 0))
    assert year == 2001
    assert gender == "Male"