import pytest
import os
from src.excel_reader import read_excel, _clean, _parse_date
import pandas as pd
import time

def test_read_excel_missing_file():
    # Test that the app returns an error string when the file doesn't exist
    employees, leaves, err = read_excel("non_existent_file.xlsx")
    assert err is not None
    assert "File not found" in err

def test_data_types():
    # Verify that the dataclasses are correctly populated
    # (Assuming you have a test template in your folder for this to run)
    path = os.path.join("src", "Leave_Tracker_Template.xlsx")
    if os.path.exists(path):
        employees, leaves, err = read_excel(path)
        assert err is None
        assert isinstance(employees, list)
        assert isinstance(leaves, list)

def test_read_empty_file():
    import pandas as pd
    empty_path = "tests/empty.xlsx"
    pd.DataFrame().to_excel(empty_path)
    
    try:
        employees, leaves, err = read_excel(empty_path)
        assert err is not None
        assert "missing" in err.lower()
    finally:
        # Give the system a split second to release the file handle
        time.sleep(0.5) 
        if os.path.exists(empty_path):
            try:
                os.remove(empty_path)
            except PermissionError:
                # If it's still locked, just log a warning instead of failing the test
                print(f"\nWarning: Could not delete {empty_path} due to file lock.")

def test_invalid_date_format():
    # Test that invalid dates return None, which your reader handles
    assert _parse_date("not-a-date") is None
    assert _parse_date("2026/13/45") is None

def test_data_cleaning():
    assert _clean("  Dev Team  ") == "Dev Team"
    assert _clean("none") == ""
    assert _clean(None) == ""