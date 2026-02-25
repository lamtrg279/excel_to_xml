
import pandas as pd

REQUIRED_COLUMNS = ["shipmentType", "job", "shipVia"]

# Helper functions
def safe_text(value):
    return "" if pd.isna(value) else str(value)

# Validate required columns and check for empty values
def validate_columns_rows(df, required_columns):
    # Check that required columns are present
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
    # Check that required columns are not empty
    empty_columns = [col for col in required_columns if df[col].dropna().empty]
    if empty_columns:
        raise ValueError(f"Data is missing in these columns: {', '.join(empty_columns)}")
