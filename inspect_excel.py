import pandas as pd
import os

file_path = r'c:\Python-Projetos\BETMGM - Uplifts\Plano de Midia_Fevereiro a Outubro 2025_V6.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    print(f"File found: {file_path}")
    print(f"Sheet names: {xl.sheet_names}")
    
    print(f"Sheet names: {xl.sheet_names}")
    
    target_sheet = 'Fev_25'
    if target_sheet in xl.sheet_names:
        print(f"\n--- Searching for header in Sheet: {target_sheet} ---")
        df = xl.parse(target_sheet, header=None)
        
        # Find the first row with at least 3 non-null values
        header_row_idx = -1
        for i, row in df.iterrows():
            if row.count() > 3:
                header_row_idx = i
                break
        
        if header_row_idx != -1:
            print(f"Found potential header at row index: {header_row_idx}")
            # Reload with correct header
            df = xl.parse(target_sheet, header=header_row_idx, nrows=10)
            print("Columns:", list(df.columns))
            print("First 5 rows of data:")
            print(df.head().to_string())
        else:
            print("Could not find a clear header row.")

    if 'Resumo' in xl.sheet_names:
         print(f"\n--- Inspection of Sheet: Resumo ---")
         df_resumo = xl.parse('Resumo', header=None)
         # Find header for Resumo too
         header_row_idx = -1
         for i, row in df_resumo.iterrows():
             if row.count() > 3:
                 header_row_idx = i
                 break
         
         if header_row_idx != -1:
             print(f"Found potential header for Resumo at row index: {header_row_idx}")
             df_resumo = xl.parse('Resumo', header=header_row_idx, nrows=10)
             print("Columns:", list(df_resumo.columns))
             print(df_resumo.head().to_string())

except Exception as e:
    print(f"Error reading Excel file: {e}")
