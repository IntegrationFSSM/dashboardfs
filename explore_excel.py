import pandas as pd
import json

file_path = r'c:\Users\yassi\OneDrive\Bureau\bilan\Stat- filière digitale.xlsx'
xls = pd.ExcelFile(file_path)

out = {}
for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet)
    # Convert to orient='records' directly
    out[sheet] = df.head(50).to_dict(orient='records')

with open(r'c:\Users\yassi\OneDrive\Bureau\bilan\temp_excel_out.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("Data exported to temp_excel_out.json")
