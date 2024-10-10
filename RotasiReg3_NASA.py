import numpy as np
import pandas as pd

file_name = "Data Stock 9 Oct 2024.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)

df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')
df = df[["KOTA", "TSH", "SITE CODE", "PT", "STORE NAME","Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].map(lambda x: np.nan if x < 0 else x)
nasa_df = []
nasa_df_2 = []

pt = 'NASA'
df = df[df['PT'] == pt]
list_tsh = sorted(df["TSH"].unique())

jumlah_baris_kosong = df['Article code no color'].isnull().sum()
print(f"NASA : Jumlah baris yang kolom 'Article'-nya NaN: {jumlah_baris_kosong}")
df = df.dropna(subset=['Article code no color'])

toko_per_kota = df.groupby('KOTA')['STORE NAME'].nunique()
toko_per_kota_filtered = toko_per_kota[toko_per_kota >= 5].index
list_kota = sorted(toko_per_kota_filtered)

for tsh in list_tsh:
    result_rows = []

    min_dos_df = df[
        (df['TSH'] == tsh) & 
        (df['DOS 30 days'] <= 23)
    ]

    min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
    min_dos_code_unique = min_dos_df['SITE CODE'].unique()

    for code in min_dos_code_unique:
        min_dos_code_df = min_dos_df[(min_dos_df['SITE CODE'] == code)]

        min_dos_code_store = list(min_dos_code_df['STORE NAME'])[0] + f" ({code})"
        min_dos_code_article = list(min_dos_code_df['Article code no color'])

        rotation_df = df[
            (df['TSH'] == tsh) & 
            (df['DOS 30 days'] >= 30) &
            (df['Article code no color'].isin(min_dos_code_article))
        ]

        for article in min_dos_code_article:
            min_dos_article_data = min_dos_code_df[min_dos_code_df['Article code no color'] == article]

            min_stock, min_sales, min_dos = min_dos_article_data[['Stock', 'Sales 30 days', 'DOS 30 days']].values[0]
            
            rotation_df = rotation_df.copy()
            rotation_df.loc[:, 'Cleaned Article'] = rotation_df['Article code no color'].str.replace('[ /]', '', regex=True)
            cleaned_article = article.replace(' ', '').replace('/', '')

            # Melakukan pengecekan dengan kolom baru yang 'bersih'
            rotation_article_data = rotation_df[rotation_df['Cleaned Article'] == cleaned_article]
            rotation_article_data = rotation_article_data.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

            for _, first_row in rotation_article_data.iterrows():

                rotation_store = first_row['STORE NAME'] + f" ({first_row['SITE CODE']})"
                rotation_stock = first_row['Stock']
                rotation_sales = first_row['Sales 30 days']
                rotation_dos = first_row['DOS 30 days']

                result_rows.append([
                    tsh, min_dos_code_store, article, min_stock, min_sales, min_dos,
                    rotation_store, rotation_stock, rotation_sales, rotation_dos
                ])

    result_df = pd.DataFrame(result_rows, columns=[
            'TSH', 'STORE TUJUAN', 'ARTICLE', 'STOCK TUJUAN', 'SALES TUJUAN', 'DOS TUJUAN',
            'STORE ROTASI', 'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
        ])

    # Fungsi untuk proses sorting dan pairing untuk setiap artikel
    def process_article(dataframe, article):
        df_article = dataframe[dataframe['ARTICLE'] == article]
        sorted_asal = df_article.sort_values(by=['DOS TUJUAN', 'SALES TUJUAN'], ascending=[True, False]).drop_duplicates('STORE TUJUAN')
        sorted_rotasi = df_article.sort_values(by=['DOS ROTASI', 'STOCK ROTASI'], ascending=[False, False]).drop_duplicates('STORE ROTASI')
        sorted_asal.reset_index(drop=True, inplace=True)
        sorted_rotasi.reset_index(drop=True, inplace=True)
        result = pd.DataFrame({
            "TSH": sorted_asal['TSH'],
            "STORE TUJUAN": sorted_asal['STORE TUJUAN'],
            "ARTICLE": sorted_asal['ARTICLE'],
            "STOCK TUJUAN": sorted_asal['STOCK TUJUAN'],
            "SALES TUJUAN": sorted_asal['SALES TUJUAN'],
            "DOS TUJUAN": sorted_asal['DOS TUJUAN'],
            "STORE ROTASI": sorted_rotasi['STORE ROTASI'],
            "STOCK ROTASI": sorted_rotasi['STOCK ROTASI'],
            "SALES ROTASI": sorted_rotasi['SALES ROTASI'],
            "DOS ROTASI": sorted_rotasi['DOS ROTASI']
        })
        return result

    articles = result_df['ARTICLE'].unique()
    if not result_df.empty:
        result_df = pd.concat([process_article(result_df, article) for article in articles])

    # Menghapus baris dengan nilai NaN (jika ada)
    result_df.dropna(inplace=True)
    result_df.reset_index(drop=True, inplace=True)

    nasa_df.append(result_df)

filtered_nasa_df = [df.dropna(how='all', axis=1) for df in nasa_df] 
result_df = pd.concat(filtered_nasa_df, ignore_index=True)
result_df = result_df.sort_values(by=['TSH', 'STORE TUJUAN', 'ARTICLE'], ascending=[True, True, True])

##################
for kota in list_kota:
    result_rows = []

    min_dos_df = df[
        (df['KOTA'] == kota) & 
        (df['DOS 30 days'] <= 23)
    ]

    min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
    min_dos_code_unique = min_dos_df['SITE CODE'].unique()

    for code in min_dos_code_unique:
        min_dos_code_df = min_dos_df[(min_dos_df['SITE CODE'] == code)]

        min_dos_code_store = list(min_dos_code_df['STORE NAME'])[0] + f" ({code})"
        min_dos_code_article = list(min_dos_code_df['Article code no color'])

        rotation_df = df[
            (df['KOTA'] == kota) & 
            (df['DOS 30 days'] > 30) &
            (df['Article code no color'].isin(min_dos_code_article))
        ]

        for article in min_dos_code_article:
            min_dos_article_data = min_dos_code_df[min_dos_code_df['Article code no color'] == article]

            min_stock, min_sales, min_dos = min_dos_article_data[['Stock', 'Sales 30 days', 'DOS 30 days']].values[0]

            rotation_df = rotation_df.copy()
            rotation_df.loc[:, 'Cleaned Article'] = rotation_df['Article code no color'].str.replace('[ /]', '', regex=True)
            cleaned_article = article.replace(' ', '').replace('/', '')

            # Melakukan pengecekan dengan kolom baru yang 'bersih'
            rotation_article_data = rotation_df[rotation_df['Cleaned Article'] == cleaned_article]
            rotation_article_data = rotation_article_data.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

            for _, first_row in rotation_article_data.iterrows():

                rotation_store = first_row['STORE NAME'] + f" ({first_row['SITE CODE']})"
                rotation_stock = first_row['Stock']
                rotation_sales = first_row['Sales 30 days']
                rotation_dos = first_row['DOS 30 days']

                result_rows.append([
                    kota, min_dos_code_store, article, min_stock, min_sales, min_dos,
                    rotation_store, rotation_stock, rotation_sales, rotation_dos
                ])

    result_df_2 = pd.DataFrame(result_rows, columns=[
            'KOTA', 'STORE TUJUAN', 'ARTICLE', 'STOCK TUJUAN', 'SALES TUJUAN', 'DOS TUJUAN',
            'STORE ROTASI', 'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
        ])

    # Fungsi untuk proses sorting dan pairing untuk setiap artikel
    def process_article(dataframe, article):
        df_article = dataframe[dataframe['ARTICLE'] == article]
        sorted_asal = df_article.sort_values(by=['DOS TUJUAN', 'SALES TUJUAN'], ascending=[True, False]).drop_duplicates('STORE TUJUAN')
        sorted_rotasi = df_article.sort_values(by=['DOS ROTASI', 'STOCK ROTASI'], ascending=[False, False]).drop_duplicates('STORE ROTASI')
        sorted_asal.reset_index(drop=True, inplace=True)
        sorted_rotasi.reset_index(drop=True, inplace=True)
        result = pd.DataFrame({
            "KOTA": sorted_asal['KOTA'],
            "STORE TUJUAN": sorted_asal['STORE TUJUAN'],
            "ARTICLE": sorted_asal['ARTICLE'],
            "STOCK TUJUAN": sorted_asal['STOCK TUJUAN'],
            "SALES TUJUAN": sorted_asal['SALES TUJUAN'],
            "DOS TUJUAN": sorted_asal['DOS TUJUAN'],
            "STORE ROTASI": sorted_rotasi['STORE ROTASI'],
            "STOCK ROTASI": sorted_rotasi['STOCK ROTASI'],
            "SALES ROTASI": sorted_rotasi['SALES ROTASI'],
            "DOS ROTASI": sorted_rotasi['DOS ROTASI']
        })
        return result

    articles = result_df_2['ARTICLE'].unique()
    if not result_df_2.empty:
        result_df_2 = pd.concat([process_article(result_df_2, article) for article in articles])

    # Menghapus baris dengan nilai NaN (jika ada)
    result_df_2.dropna(inplace=True)
    result_df_2.reset_index(drop=True, inplace=True)

    nasa_df_2.append(result_df_2)

filtered_nasa_df = [df.dropna(how='all', axis=1) for df in nasa_df_2] 
result_df_2 = pd.concat(filtered_nasa_df, ignore_index=True)
result_df_2 = result_df_2.sort_values(by=['KOTA', 'STORE TUJUAN', 'ARTICLE'], ascending=[True, True, True])

output_file = "rotasi NASA.xlsx"
with pd.ExcelWriter(output_file) as writer:
    result_df.to_excel(writer, sheet_name='ByTSH', index=False)
    result_df_2.to_excel(writer, sheet_name='ByKOTA', index=False)

# print(result_df_2)
