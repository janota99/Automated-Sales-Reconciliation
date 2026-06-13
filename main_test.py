import os
import pandas as pd
from engine.ingestion import stage_excel_file
from engine.matching import execute_sales_reconciliation

# 1. Setup Your Raw Material File Targets (Flipped to correct alignment!)
DATA_DIR = "data"
FILE_AS400 = "Infinium Data NEW.xlsx"  # This is your AS400/Infinium data
FILE_QB = "QBO Data NEW.xlsx"          # This is your Quickbooks Online export

# 2. Build Your Dynamic System File Paths
path_qb = os.path.join(DATA_DIR, FILE_QB)
path_as400 = os.path.join(DATA_DIR, FILE_AS400)

print("⚡ Initializing Failsafe Multi-Pass Matching Execution...")

# 3. Verify Both Staging Targets Exist in the Data Window
if not os.path.exists(path_qb) or not os.path.exists(path_as400):
    print("❌ Error: Missing file targets in your 'data/' folder.")
else:
    # 4. Load Data Frames Natively Into Local Memory
    # QBO Data has a 4-row header banner, so we tell openpyxl to skip to Row 5 (index 4)
    df_qb = pd.read_excel(path_qb, header=4)
    df_qb.columns = df_qb.columns.astype(str).str.strip().str.upper()

    # Infinium/AS400 Data starts immediately on Row 1, so it uses default header processing
    df_as400 = pd.read_excel(path_as400)
    df_as400.columns = df_as400.columns.astype(str).str.strip().str.upper()
    # 🔍 5. Diagnostic Header Inspection Print Window
    print("\n📋 ACTUAL QB COLUMNS IN MEMORY:")
    print(list(df_qb.columns))
    print("\n📋 ACTUAL AS400 COLUMNS IN MEMORY:")
    print(list(df_as400.columns))
    print("-" * 60)

    # 6. Map the Standard Column Header Selectors (Validated Uppercase Strings)
    QB_COLUMNS = {
        'po': 'P.O. NUMBER',  
        'invoice': 'NUM',
        'amount': 'AMOUNT'     
    }

    AS400_COLUMNS = {
        'po': 'OHDESC', 
        'invoice': 'OHOBNO',
        'amount': 'OHTOTA'} 
    # 7. Fire the Vectorized Sales Reconciliation Module (Fixed continuous line syntax)
    reconciled_qb, reconciled_as400 = execute_sales_reconciliation(df_qb, df_as400, QB_COLUMNS, AS400_COLUMNS)
    # 8. Output Summary Calculations Natively to Console
    cleared = len(reconciled_qb[reconciled_qb['IS_MATCHED'] == True])
    print("\n🏁 RECONCILIATION PROCESSING COMPLETE:")
    print(f" -> Total Base Records Processed: {len(reconciled_qb)}")
    print(f" -> ✅ Cleared Autonomously:       {cleared}")
    print(f" -> ❌ Remaining Exceptions:       {len(reconciled_qb) - cleared}")