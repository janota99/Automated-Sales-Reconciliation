import pandas as pd

def stage_excel_file(file_path, sheet_name=0):
    """
    Reads an Excel file, cleans basic formatting issues, 
    and prepares it for the reconciliation staging window.
    """
    try:
        # Load the file into a Pandas DataFrame using openpyxl
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Clean up column names (strip whitespace and convert to uppercase for consistency)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        return df
    except Exception as e:
        print(f"Error loading file at {file_path}: {e}")
        return None

def generate_staging_summary(df, amount_column):
    """
    Calculates baseline metrics for a dataset before matching begins.
    This mimics your visual 'reconciliation window' summary.
    """
    # Ensure our target column is uppercase to match our cleaning step
    amt_col_clean = amount_column.strip().upper()
    
    if amt_col_clean not in df.columns:
        return {
            "Total Rows": len(df),
            "Total Absolute Volume": 0,
            "Net Balance": 0,
            "Error": f"Column '{amount_column}' not found."
        }
    
    # Force the amount column to be numeric, turning errors (like text) into NaN
    numeric_amounts = pd.to_numeric(df[amt_col_clean], errors='coerce').fillna(0)
    
    summary = {
        "Total Rows": len(df),
        "Total Absolute Volume": float(numeric_amounts.abs().sum()),
        "Net Balance": float(numeric_amounts.sum()),
        "Error": None
    }
    
    return summary