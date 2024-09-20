import numpy as np
import pandas as pd
from datetime import datetime

file_name = "Data Stock 20 Sep 2024.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)

df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')
df = df[["AREA", "KOTA", "TSH", "SITE CODE", "PT", "STORE NAME", "Article code no color", "Stock","Sales 30 days", "DOS 30 days"]]
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days', 'DOS 30 days']].map(lambda x: np.nan if x < 0 else x)

# era
pt = 'EAR'
list_kota = sorted(list(set(df["KOTA"])))
kota = list_kota[0] # KAB. BANDUNG

min_dos_df = df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] <= 15)]
min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])


min_dos_code_unique = min_dos_df['SITE CODE'].drop_duplicates().tolist()

code = min_dos_code_unique[0]
min_dos_df = min_dos_df[(min_dos_df['SITE CODE'] == code)]

min_dos_article = list(set(min_dos_df['Article code no color']))

rotation_df = df[
    (df['KOTA'] == kota) &
    (df['DOS 30 days'] >= 45) 
    # (df['Article code no color'].isin(min_dos_article))
]

# print(len(min_dos_article))
print(rotation_df)