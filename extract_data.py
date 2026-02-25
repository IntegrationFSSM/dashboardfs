import pandas as pd
import json

def clean_sum(series):
    return pd.to_numeric(series, errors='coerce').sum()

data = {
    'personnel': {
        '2024-2025': {'Admin': 93, 'Enseignants': 293},
        '2025-2026': {'Admin': 92, 'Enseignants': 261}
    },
    'gender_personnel': {
        'Admin': {'F': 49, 'H': 44},
        'Enseignants': {'F': 85, 'H': 207}
    },
    'students_2024': {
        'Licences': 10187,
        'Masters': 706,
        'Doctorat': 515
    },
    'students_2025': {
        'Licences': 11915,
        'Masters': 381,
        'Doctorat': 0
    },
    'capacity_2024': {},
    'capacity_2025': {}
}

import os
path = r"c:\Users\yassi\OneDrive\Bureau\bilan"

file_24_25 = os.path.join(path, "Capacité d'accueil_FSSM_ 2024-2025.xlsx")
file_25_26 = os.path.join(path, "Capacité d'accueil_FSSM_ 2025-2026 - Locaux d'enseignement.xlsx")

def extract_capacity(file_path):
    res = {}
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=9)
        # Assuming last column is total capacity, or first numeric column > 10
        total_cap = 0
        names_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        # just sum all numeric columns and pick the max assuming the largest sum is the total capacity across all rooms
        numeric_sums = [clean_sum(df[c]) for c in df.columns]
        if numeric_sums:
             total_cap = max(numeric_sums) 
             res[sheet] = total_cap
    return res

if os.path.exists(file_24_25):
    data['capacity_2024'] = extract_capacity(file_24_25)
if os.path.exists(file_25_26):
    data['capacity_2025'] = extract_capacity(file_25_26)

with open(os.path.join(path, 'extracted_clean_data.json'), 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=4)
print("Data extracted successfully")
