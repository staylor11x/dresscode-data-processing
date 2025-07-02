from utils import processing
import pandas as pd

def test_is_valid_subject():
    assert processing.is_valid_subject("English", ["total", "subtotal"])
    assert not processing.is_valid_subject("Total", ["total", "subtotal"])
    assert not processing.is_valid_subject(None, ["total"])
