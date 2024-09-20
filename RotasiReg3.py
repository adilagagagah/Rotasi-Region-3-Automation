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
min_dos_df = min_dos_df.sort_values(by='DOS 30 days', ascending=True)

min_dos_code = list(min_dos_df['SITE CODE'])
# min_dos_store = list(min_dos_df['STORE NAME'])[0] + f" ({code})"
min_dos_product = list(min_dos_df['Article code no color'])

print(min_dos_df)