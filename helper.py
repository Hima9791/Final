import pandas as pd
import requests
from difflib import SequenceMatcher
import io
import os

def load_file(file_path):
    """
    Load CSV or Excel file from local path or GitHub raw URL.
    """
    if file_path.startswith("http://") or file_path.startswith("https://"):
        response = requests.get(file_path)
        response.raise_for_status()
        file_bytes = io.BytesIO(response.content)
        if file_path.endswith(".csv"):
            return pd.read_csv(file_bytes)
        return pd.read_excel(file_bytes)
    else:
        if file_path.endswith(".csv"):
            return pd.read_csv(file_path)
        return pd.read_excel(file_path)

def similarity_ratio(a, b):
    """Return similarity ratio as percentage with 2 decimal places."""
    return round(SequenceMatcher(None, a, b).ratio() * 100, 2)

def normalize_series(series):
    """Normalize series name for comparison."""
    if pd.isna(series):
        return ""
    return str(series).strip().upper()
