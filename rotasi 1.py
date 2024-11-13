import numpy as np
import pandas as pd
import math

file_name = "Data Stock 9 Oct 2024.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)

df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')
df = df[["KOTA", "TSH", "SITE CODE", "PT", "STORE NAME","Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].clip(lower=0).fillna(0)
# ear_df_tsh = []
ear_df_city = []

# pt = 'EAR'
# df = df[df['PT'] == pt]

for pt in ['EAR', 'DCM', 'NASA']:
    df_pt = df[df['PT'] == pt]
    print(df_pt.shape)



jumlah_baris_kosong = df['Article code no color'].isnull().sum()
print(f"EAR : Jumlah baris yang kolom 'Article'-nya NaN: {jumlah_baris_kosong}")
df = df.dropna(subset=['Article code no color'])

def custom_round(n):
    # Mengecek bagian desimal dari n
    if n - math.floor(n) >= 0.5:
        return math.ceil(n)
    else:
        return math.floor(n)

def calculate_dos(stock, sales):
    return custom_round((stock / sales * 30)) if sales != 0 else 0

def rotate_stock(df_asal, df_tujuan, log_rotasi):
    i = 0
    while True:
        dos_df_asal = list(df_asal['DOS 30 days'])
        dos_df_tujuan = list(df_tujuan['DOS 30 days'])
        
        # if all(x != 0 for x in dos_df_tujuan):
        #     print('hentikan karena store tujuan tidak ada 0')
        #     break

        if all(x >= 23 for x in dos_df_tujuan):
            print('hentikan karena store tujuan sudah lebih dari 23 semua')
            break

        if df_asal.empty or df_tujuan.empty:
            print('hentikan karena data kosong')
            break
        
        # Pilih store asal dengan melihat dos simulasi terbesar
        df_asal = df_asal.sort_values(by=['DOS 30 days','Stock'], ascending=[False, False]).reset_index(drop=True)

        df_asal['Simulated Sales 30 days'] = df_asal['Sales 30 days'].where(df_asal['Sales 30 days'] != 0, 1)
        simulated_doses = []
        for index, row in df_asal.iterrows():
            simulated_stock = row['Stock'] - 1
            simulated_dos = calculate_dos(simulated_stock, row['Simulated Sales 30 days'])
            simulated_doses.append(simulated_dos)
        df_asal.drop('Simulated Sales 30 days', axis=1, inplace=True)
      
        selected_index = simulated_doses.index(max(simulated_doses))
        asal = df_asal.loc[selected_index]

        if (simulated_doses[selected_index] < 23):
            print('hentikan karena store terakhir yang dapat dirotasikan sudah kurang dari 23')
            break

        # Pilih store tujuan dengan DOS terendah yang memiliki sales tertinggi
        df_tujuan = df_tujuan.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False]).reset_index(drop=True)
        tujuan = df_tujuan.loc[0]

        # Update stok dan hitung ulang DOS untuk store asal dan tujuan
        df_asal.at[selected_index, 'Stock'] -= 1
        df_tujuan.at[0, 'Stock'] += 1
        df_asal.at[selected_index, 'DOS 30 days'] = calculate_dos(df_asal.at[selected_index, 'Stock'], df_asal.at[selected_index, 'Sales 30 days'])
        df_tujuan.at[0, 'DOS 30 days'] = calculate_dos(df_tujuan.at[0, 'Stock'], df_tujuan.at[0, 'Sales 30 days'])

        new_log_entry = pd.DataFrame({
            "KOTA": [asal['KOTA']],
            "ARTICLE": [asal['Article code no color']],
            "STORE ASAL": [f"{asal['STORE NAME']}"],
            "BARANG TEROTASI": [1],
            "STORE TUJUAN": [f"{tujuan['STORE NAME']}"]
        })
        
        log_rotasi = pd.concat([log_rotasi, new_log_entry], ignore_index=True)
        
        i += 1
        print(f'{i}.',asal['STORE NAME'], 'ke' ,tujuan['STORE NAME'])
    
    return df_asal, df_tujuan, log_rotasi

def split_store(df):
    df_asal = df[
        ((df['DOS 30 days'] >= 30) & (df['Stock'] >= 2)) |
        ((df['Sales 30 days'] >= 0) & (df['Stock'] >= 2))
    ]
    df_asal = df_asal.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

    df_tujuan = df[
        (df['DOS 30 days'] <= 23)&
        (df['Sales 30 days'] != 0)
    ]
    df_tujuan = df_tujuan.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
    
    return df_asal, df_tujuan

def rotate_and_merge_stock(df_asal, df_asal_updated, df_tujuan, df_tujuan_updated, log_rotasi):
    # Summing up rotated goods in log_rotasi
    log_rotasi = log_rotasi.groupby(["KOTA", "ARTICLE", "STORE ASAL", "STORE TUJUAN"], as_index=False).agg({
        "BARANG TEROTASI": "sum"
    })
    log_rotasi["TOTAL"] = log_rotasi.groupby(["KOTA", "STORE ASAL", "ARTICLE"])["BARANG TEROTASI"].transform("sum")
    log_rotasi = log_rotasi[["KOTA", "STORE ASAL", "ARTICLE", "TOTAL", "BARANG TEROTASI", "STORE TUJUAN"]]

    merged_data = pd.merge(
        log_rotasi, df_asal[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']],left_on='STORE ASAL', right_on='STORE NAME', 
        how='left', suffixes=('', '_asal')).drop(columns=['STORE NAME'])  

    merged_data = pd.merge(
        merged_data, df_asal_updated[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], left_on='STORE ASAL', right_on='STORE NAME', 
        how='left',suffixes=('_asal_sebelum', '_asal_setelah')).drop(columns=['STORE NAME'])

    merged_data = pd.merge(
        merged_data, df_tujuan[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], left_on='STORE TUJUAN', right_on='STORE NAME', 
        how='left',suffixes=('', '_tujuan_sebelum')).drop(columns=['STORE NAME'])

    # Merge keempat, menambahkan kolom tujuan setelah rotasi
    merged_data = pd.merge(
        merged_data, 
        df_tujuan_updated[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], 
        left_on='STORE TUJUAN', right_on='STORE NAME', 
        how='left',
        suffixes=('_tujuan_sebelum', '_tujuan_setelah')
    ).drop(columns=['STORE NAME'])

    # Renaming columns for clarity
    merged_data.rename(columns={
        "Stock_asal_sebelum": "Stock", "Sales 30 days_asal_sebelum": "Sales", "DOS 30 days_asal_sebelum": "Dos",
        "Stock_asal_setelah": "Stock Akhir", "Sales 30 days_asal_setelah": "Sales", "DOS 30 days_asal_setelah": "Dos Akhir",
        "Stock_tujuan_sebelum": "Stock Tujuan", "Sales 30 days_tujuan_sebelum": "Sales Tujuan", "DOS 30 days_tujuan_sebelum": "Dos Tujuan",
        "Stock_tujuan_setelah": "Stock Akhir Tujuan", "Sales 30 days_tujuan_setelah": "Sales Akhir Tujuan", "DOS 30 days_tujuan_setelah": "Dos Akhir Tujuan"
    }, inplace=True)

    # Selecting relevant columns
    merged_data = merged_data[[
        "KOTA", "ARTICLE", "STORE ASAL", "Stock", "Sales", "Dos", "Stock Akhir", "Sales", "Dos Akhir", 
        "TOTAL", "BARANG TEROTASI", "STORE TUJUAN", "Stock Tujuan", "Sales Tujuan", "Dos Tujuan", 
        "Stock Akhir Tujuan", "Sales Akhir Tujuan", "Dos Akhir Tujuan"
    ]]

    return merged_data

#------------------------------------------------------------------------------------------

#------------------------------------------------------------------------------------------
cities = df["KOTA"].unique()
log_rotasi = pd.DataFrame(columns=["KOTA", "ARTICLE", "STORE ASAL", "BARANG TEROTASI","STORE TUJUAN"])

# for city in cities:

city = 'KAB. KARAWANG'
df_percity = df[df['KOTA'] == city]
articles = df_percity["Article code no color"].unique().tolist()

# for article in articles:
article = 'OPPO A3X 4/64GB'
df_percity_perarticle = df_percity[df_percity['Article code no color'] == article]
df_percity_perarticle = df_percity_perarticle[["KOTA", "Article code no color", "STORE NAME", "Stock", "Sales 30 days", "DOS 30 days"]]

df_asal, df_tujuan = split_store(df_percity_perarticle)
df_asal_updated, df_tujuan_updated, log_rotasi  = rotate_stock(df_asal, df_tujuan, log_rotasi)

# log_rotasi = log_rotasi.groupby(["KOTA", "ARTICLE", "STORE ASAL", "STORE TUJUAN"], as_index=False).agg({
#         "BARANG TEROTASI": "sum"
# })
# log_rotasi["TOTAL"] = log_rotasi.groupby(["KOTA", "STORE ASAL", "ARTICLE"])["BARANG TEROTASI"].transform("sum")
# log_rotasi = log_rotasi[["KOTA", "STORE ASAL", "ARTICLE", "TOTAL", "BARANG TEROTASI", "STORE TUJUAN"]]

# if df_asal_updated.empty or df_tujuan_updated.empty:
#     continue
print("\n\n\nStore Asal sebelum Rotasi:")
print(df_asal)
print("\nStore Tujuan sebelum Rotasi:")
print(df_tujuan)
print("\n\nStore Asal setelah Rotasi:")
print(df_asal_updated)
print("\nStore Tujuan setelah Rotasi:")
print(df_tujuan_updated)
print("\n\nlog Rotasi:")
print(log_rotasi)

# merged_data = rotate_and_merge_stock(df_asal, df_asal_updated, df_tujuan, df_tujuan_updated, log_rotasi)

# with pd.ExcelWriter('TES3.xlsx') as writer:
#     merged_data.to_excel(writer, sheet_name='ByKota', index=False)