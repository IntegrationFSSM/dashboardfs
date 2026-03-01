import pandas as pd
import json

file_path = r'c:\Users\yassi\OneDrive\Bureau\bilan\Stat-langues et power skills.xlsx'
df = pd.read_excel(file_path, sheet_name=0, header=6)

# The columns starting at index 6 look like:
# 0: Université
# 1: Etablissement
# 2: Diplôme
# 3: Filière
# 4: Semestre
# 5: Effectif global (Langues)
# 6: Garçons (Langues)
# 7: Filles (Langues)
# 8: Effectif global (Power Skills)
# 9: Garçons (Power Skills)
# 10: Filles (Power Skills)

# Drop rows where 'Filière' is NaN to filter out headers/footers if needed
# Wait, Filière is column 3 (index 3). Sometimes it's NaN if it's the same Filière but different Semestre.
# We will forward fill the 'Filière' column.
df.iloc[:, 3] = df.iloc[:, 3].ffill()

# Drop rows where both 'Effectif global (Langues)' and 'Effectif global (Power Skills)' are NaN
df = df.dropna(subset=[df.columns[5], df.columns[8]], how='all')

# Clean data (convert to numeric)
for col in range(5, 11):
    df.iloc[:, col] = pd.to_numeric(df.iloc[:, col], errors='coerce').fillna(0)

# Aggregates
total_langues = int(df.iloc[:, 5].sum())
total_langues_f = int(df.iloc[:, 7].sum())
total_langues_m = int(df.iloc[:, 6].sum())

total_ps = int(df.iloc[:, 8].sum())
total_ps_f = int(df.iloc[:, 10].sum())
total_ps_m = int(df.iloc[:, 9].sum())

# Per Filière Aggregates
filiere_stats = []
for filiere in df.iloc[:, 3].unique():
    if pd.isna(filiere) or str(filiere).strip() == '': continue
    sub = df[df.iloc[:, 3] == filiere]
    filiere_stats.append({
        "filiere": str(filiere).strip(),
        "langues_total": int(sub.iloc[:, 5].sum()),
        "langues_f": int(sub.iloc[:, 7].sum()),
        "ps_total": int(sub.iloc[:, 8].sum()),
        "ps_f": int(sub.iloc[:, 10].sum())
    })

out = {
    "quick_stats": {
        "total_langues": total_langues,
        "total_langues_f": total_langues_f,
        "total_langues_m": total_langues_m,
        "total_ps": total_ps,
        "total_ps_f": total_ps_f,
        "total_ps_m": total_ps_m
    },
    "filiere_stats": filiere_stats
}

with open(r'c:\Users\yassi\OneDrive\Bureau\bilan\langues_ps_summary.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("Aggregates extracted to langues_ps_summary.json")
