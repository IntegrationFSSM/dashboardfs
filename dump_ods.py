import pandas as pd
pd.set_option('display.max_rows', None, 'display.max_columns', None, 'display.width', 1000)
df = pd.read_excel('Donnée Stat -Synthèse.ods', sheet_name='Sheet1')
with open('output_ods_utf8.txt', 'w', encoding='utf-8') as f:
    f.write(df.to_string())
