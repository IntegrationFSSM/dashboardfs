import pandas as pd
import json

filename = r"c:\Users\yassi\OneDrive\Bureau\bilan\Capacité d'accueil_FSSM_ 2025-2026 - Locaux d'enseignement.xlsx"
out_filename = r"c:\Users\yassi\OneDrive\Bureau\bilan\full_dump.txt"

xls = pd.ExcelFile(filename)
with open(out_filename, "w", encoding="utf-8") as f:
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet).fillna("")
        f.write(f"\n{'='*50}\nSheet: {sheet}\n{'='*50}\n")
        f.write(df.to_string())
        f.write("\n")
