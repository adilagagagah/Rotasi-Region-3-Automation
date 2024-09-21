import numpy as np
import pandas as pd
import re

# Load data
df_origin = pd.read_excel("../ROTASI REGION 3/Data Stock 20 Sep 2024.xlsx", sheet_name="dos by store-brand type & area")
df_master = pd.read_excel("../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx", sheet_name="REGION 3", header=1)

# Merge dengan kolom SITE CODE dan PT
df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')

# Filter kolom relevan dan ganti nilai negatif dengan NaN
df = df[["AREA", "KOTA", "TSH", "SITE CODE", "PT", "STORE NAME", "Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].applymap(lambda x: np.nan if x < 0 else x)

# Inisialisasi variabel
result_rows = []
pt, kota = 'EAR', sorted(df["KOTA"].unique())[0]  # KAB. BANDUNG contoh

# Toko dengan DOS <= 15
min_dos_df = df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] <= 15)].sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
min_dos_code_unique = min_dos_df['SITE CODE'].drop_duplicates().tolist()

# Iterasi toko dengan DOS rendah
for code in min_dos_code_unique:
    min_dos_code_df = min_dos_df[min_dos_df['SITE CODE'] == code]
    min_dos_store = f"{min_dos_code_df['STORE NAME'].iloc[0]} ({code})"
    min_dos_article = min_dos_code_df['Article code no color'].drop_duplicates().tolist()

    # Kandidat rotasi (toko dengan DOS >= 45)
    rotation_df = df[(df['KOTA'] == kota) & (df['DOS 30 days'] >= 45) & (df['Article code no color'].isin(min_dos_article))]

    for article in min_dos_article:
        min_dos_article_data = min_dos_df[min_dos_df['Article code no color'] == article].iloc[0]
        min_stock, min_sales, min_dos = min_dos_article_data[['Stock', 'Sales 30 days', 'DOS 30 days']]

        rotation_article_data = rotation_df[rotation_df['Article code no color'] == article].sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

        # Lakukan rotasi
        for _, row in rotation_article_data.iterrows():
            stock_sales_diff = row['Stock'] - row['Sales 30 days']

            if stock_sales_diff > 0:
                result_rows.append([min_dos_store, f"{row['STORE NAME']} ({row['SITE CODE']})", article, min_stock, min_sales, min_dos,
                                    row['Stock'], row['Sales 30 days'], row['DOS 30 days']])
                stock_sales_diff -= 1

            if stock_sales_diff <= 0:
                break

# Hasil DataFrame
result_df = pd.DataFrame(result_rows, columns=[
    'STORE ASAL', 'ROTASI DARI', 'ARTICLE', 'STOCK ASAL', 'SALES ASAL', 'DOS ASAL',
    'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI']).drop_duplicates(subset=['ARTICLE'])

# Ekstrak kode dari nama toko
def extract_code(store_name):
    return re.search(r'\((\w+)\)', store_name).group(1) if re.search(r'\((\w+)\)', store_name) else ''

result_df['CODE STORE ASAL'] = result_df['STORE ASAL'].apply(extract_code)
result_df['CODE ROTASI DARI'] = result_df['ROTASI DARI'].apply(extract_code)

# Sortir dan tambah baris kosong
result_df_sorted = result_df.sort_values(by=['CODE STORE ASAL', 'CODE ROTASI DARI', 'ARTICLE']).drop(columns=['CODE STORE ASAL', 'CODE ROTASI DARI'])

combined_rows, previous_store_asal, previous_rotasi_dari = [], None, None

for _, row in result_df_sorted.iterrows():
    current_store_asal, current_rotasi_dari = row['STORE ASAL'], row['ROTASI DARI']
    if current_store_asal != previous_store_asal or current_rotasi_dari != previous_rotasi_dari:
        if previous_store_asal is not None:
            combined_rows.append([''] * len(row))
    combined_rows.append(row.tolist())
    previous_store_asal, previous_rotasi_dari = current_store_asal, current_rotasi_dari

# DataFrame akhir
final_df = pd.DataFrame(combined_rows, columns=result_df_sorted.columns)

# Simpan ke Excel
with pd.ExcelWriter("Result_Stock_Rotation_sorted_by_code_no_duplicates.xlsx") as writer:
    result_df.to_excel(writer, sheet_name="Original Rotasi", index=False)
    final_df.to_excel(writer, sheet_name="Sorted Rotasi by Code", index=False)

print("DataFrame telah disimpan")
