import numpy as np
import pandas as pd

file_name = "Data Stock 20 Sep 2024.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)

df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')
df = df[["KOTA", "SITE CODE", "PT", "STORE NAME","Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].map(lambda x: np.nan if x < 0 else x)

result_rows = []
pt = 'EAR'
list_kota = sorted(df["KOTA"].unique())
kota = list_kota[0]  # KAB. BANDUNG

min_dos_df = df[
    (df['PT'] == pt) & 
    (df['KOTA'] == kota) & 
    (df['DOS 30 days'] <= 15)
]
min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
min_dos_code_unique = min_dos_df['SITE CODE'].unique()

# Loop melalui SITE CODE unik
for code in min_dos_code_unique:
    min_dos_code_df = min_dos_df[min_dos_df['SITE CODE'] == code]
    store_name_code = f"{min_dos_code_df['STORE NAME'].iloc[0]} ({code})"
    
    # Loop melalui artikel unik di toko tersebut
    for article in min_dos_code_df['Article code no color'].unique():
        article_data = min_dos_code_df[min_dos_code_df['Article code no color'] == article]
        stock_asal, sales_asal, dos_asal = article_data[['Stock', 'Sales 30 days', 'DOS 30 days']].values[0]

        # Cari toko untuk rotasi (DOS >= 45)
        rotation_df = df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] >= 45) & (df['Article code no color'] == article)]
        if not rotation_df.empty:
            best_rotation = rotation_df.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False]).iloc[0]
            result_rows.append([
                store_name_code, article, stock_asal, sales_asal, dos_asal,
                f"{best_rotation['STORE NAME']} ({best_rotation['SITE CODE']})", best_rotation['Stock'], 
                best_rotation['Sales 30 days'], best_rotation['DOS 30 days']
            ])

# Buat DataFrame dari hasil dan simpan ke file
result_df = pd.DataFrame(result_rows, columns=[
    'STORE ASAL', 'ARTICLE', 'STOCK ASAL', 'SALES ASAL', 'DOS ASAL',
    'STORE ROTASI', 'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
])

# Sort hasil akhir berdasarkan DOS ASAL dan SALES ASAL
result_df = result_df.sort_values(by=['DOS ASAL', 'SALES ASAL'], ascending=[True, False]).drop_duplicates(subset=['ARTICLE'], keep='first')
# result_df.to_excel("result_optimized.xlsx", index=False)

print(result_df)
