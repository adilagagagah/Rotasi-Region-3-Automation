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
list_kota = sorted(list(set(df["KOTA"])))
kota = list_kota[0]  # KAB. BANDUNG

min_dos_df = df[
    (df['PT'] == pt) & 
    (df['KOTA'] == kota) & 
    (df['DOS 30 days'] <= 15)
]

min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
min_dos_code_unique = min_dos_df['SITE CODE'].drop_duplicates().tolist()

code = min_dos_code_unique[0] #E851
min_dos_code_df = min_dos_df[(min_dos_df['SITE CODE'] == code)]

min_dos_code_store = list(min_dos_code_df['STORE NAME'])[0] + f" ({code})"
min_dos_code_article = list(min_dos_code_df['Article code no color'])

rotation_df = df[
    (df['KOTA'] == kota) &
    (df['DOS 30 days'] >= 45) &
    (df['Article code no color'].isin(min_dos_code_article))
]

for article in min_dos_code_article:
    min_dos_article_data = min_dos_code_df[min_dos_code_df['Article code no color'] == article]

    min_stock = min_dos_article_data['Stock'].values[0]
    min_sales = min_dos_article_data['Sales 30 days'].values[0]
    min_dos = min_dos_article_data['DOS 30 days'].values[0]

    # Data dari rotation_df untuk produk yang cocok
    rotation_article_data = rotation_df[rotation_df['Article code no color'] == article]
    rotation_article_data = rotation_article_data.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

    if not rotation_article_data.empty:
        first_row = rotation_article_data.iloc[0]

        rotation_store = first_row['STORE NAME'] + f" ({first_row['SITE CODE']})"
        rotation_stock = first_row['Stock']
        rotation_sales = first_row['Sales 30 days']
        rotation_dos = first_row['DOS 30 days']

        # Menambahkan ke list hasil
        result_rows.append([
            min_dos_code_store, article, min_stock, min_sales, min_dos,
            rotation_store, rotation_stock, rotation_sales, rotation_dos
        ])

result_df = pd.DataFrame(result_rows, columns=[
    'STORE ASAL', 'ARTICLE', 'STOCK ASAL', 'SALES ASAL', 'DOS ASAL',
    'STORE ROTASI', 'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
])

output_file = "result.xlsx"
result_df.to_excel(output_file, index=False)

print(result_df)