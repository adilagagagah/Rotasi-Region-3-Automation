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

# era
list_kota = list(set(df["KOTA"]))
# print(len(kota))


#print(df)