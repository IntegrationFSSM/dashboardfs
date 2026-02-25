import pandas as pd
import os

path = r'c:\Users\yassi\OneDrive\Bureau\bilan'
files = [f for f in os.listdir(path) if f.endswith(('.xlsx', '.ods'))]

with open(r'c:\Users\yassi\OneDrive\Bureau\bilan\data_summary.txt', 'w', encoding='utf-8') as out:
    for f in files:
        full_path = os.path.join(path, f)
        out.write(f"\n{'='*50}\nFILE: {f}\n{'='*50}\n")
        try:
            engine = "odf" if f.endswith(".ods") else "openpyxl"
            xls = pd.ExcelFile(full_path, engine=engine)
            out.write(f"Sheets: {xls.sheet_names}\n")
            for sheet in xls.sheet_names:
                out.write(f"\n--- Sheet: {sheet} ---\n")
                df = pd.read_excel(full_path, sheet_name=sheet, engine=engine, nrows=10)
                out.write(df.to_string() + "\n")
        except Exception as e:
            out.write(f"Error reading file: {e}\n")
