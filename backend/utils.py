import pandas as pd

def extract_organismo(file_path):
    try:
        if "/" in file_path:
            organismo = file_path.split("/")[0].strip()
            return organismo if organismo else None
        
        return None
    except Exception as e:
        return None
    
def safe_text(row_dict, key, default=None):
    value = row_dict.get(key, default)
    if value is None or pd.isna(value):
        return default

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return default

    return text

def safe_datetime(row_dict, key):
    value = row_dict.get(key)
    if value is None or pd.isna(value):
        return None

    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed

def safe_int(row_dict, key):
    value = row_dict.get(key)
    if value is None or pd.isna(value):
        return None

    if isinstance(value, str):
        value = value.strip()
        if not value or value.lower() == "nan":
            return None

    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None