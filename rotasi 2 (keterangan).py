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
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].clip(lower=0).fillna(0)
# ear_df_tsh = []
ear_df_city = []

pt = 'EAR'
df = df[df['PT'] == pt]

jumlah_baris_kosong = df['Article code no color'].isnull().sum()
print(f"EAR : Jumlah baris yang kolom 'Article'-nya NaN: {jumlah_baris_kosong}")
df = df.dropna(subset=['Article code no color'])




city = 'KOTA SEMARANG'
df = df[df['KOTA'] == city]

def calculate_dos(stock, sales):
    return round((stock / sales * 30),0) if sales != 0 else 0

df_asal = df[
    (df['DOS 30 days'] >= 30) &
    (df['Stock'] >= 3)
]

df_asal = df_asal[df_asal['Article code no color'] == 'Redmi Note 13Pro5G 12/512']
df_asal = df_asal[["KOTA","Article code no color", "STORE NAME", "SITE CODE", "Stock", "Sales 30 days", "DOS 30 days"]]

# print(df_asal.sort_values(by=['DOS 30 days'], ascending=[False]))


print('--------------------------------------------\n')

df_tujuan = df[
    (df['DOS 30 days'] <= 23)&
    (df['Sales 30 days'] != 0)
]

df_tujuan = df_tujuan[df_tujuan['Article code no color'] == 'Redmi Note 13Pro5G 12/512']
df_tujuan = df_tujuan[["KOTA","Article code no color", "STORE NAME", "SITE CODE", "Stock", "Sales 30 days", "DOS 30 days"]]

# print(df_tujuan.sort_values(by=['DOS 30 days'], ascending=[True]))


print('--------------------------------------------\n')

# dos_df_tujuan = list(df_tujuan['DOS 30 days'])
# print(any(x < 23 for x in dos_df_tujuan))

# data_asal = {
#     'STORE NAME': [
#         'ERAFONE PARAGON CITY MALL'
#     ],
#     'Stock': [10],
#     'Sales 30 days': [3.0],
#     'DOS 30 days': [100.0]
# }

# data_tujuan = {
#     'STORE NAME': [
#         'ERAFONE SOEHAT SEMARANG', 'ERAFONE PURIANJASMORO SEMARANG', 'ERAFONE AND MORE NGALIYAN SEMARANG',
#         'ERAFONE MIJEN SEMARANG', 'ERAFONE WOLTER MONGINSIDI SEMARANG', 'ERAFONE RUKO KLIPANG SEMARANG',
#         'MEGASTORE RUKO TELOGOSARI SEMARANG'
#     ],
#     'Stock': [0, 0, 0, 2, 1, 1, 4],
#     'Sales 30 days': [1.0, 2.0, 2.0, 7.0, 3.0, 2.0, 7.0],
#     'DOS 30 days': [0.0, 0.0, 0.0, 9.0, 10.0, 15.0, 17.0]
# }

# df_asal = pd.DataFrame(data_asal).sort_values(by=['DOS 30 days','Stock'], ascending=[False, False]).reset_index(drop=True)
# df_tujuan = pd.DataFrame(data_tujuan).sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False]).reset_index(drop=True)

# print(df_asal)
# print('--------------------------------------------\n')
# print(df_tujuan,'\n')

log_rotasi = pd.DataFrame(columns=["KOTA", "ARTICLE", "STORE ASAL", "BARANG TEROTASI","STORE TUJUAN"])

def rotate_stock(df_asal, df_tujuan, log_rotasi):
    i = 0
    while True:
        dos_df_asal = list(df_asal['DOS 30 days'])
        dos_df_tujuan = list(df_tujuan['DOS 30 days'])
        
        if all(x != 0 for x in dos_df_tujuan):
            if any(x < 30 for x in dos_df_asal):
                break

        if all(x >= 23 for x in dos_df_tujuan):
            break
        
        # Pilih store asal dengan melihat dos simulasi terbesar
        df_asal = df_asal.sort_values(by=['DOS 30 days','Stock'], ascending=[False, False]).reset_index(drop=True)
        simulated_doses = []
        for index, row in df_asal.iterrows():
            simulated_stock = row['Stock'] - 1
            simulated_dos = calculate_dos(simulated_stock, row['Sales 30 days'])
            simulated_doses.append(simulated_dos)
        selected_index = simulated_doses.index(max(simulated_doses))
        asal = df_asal.loc[selected_index]

        if simulated_doses[selected_index] < 20:
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
        # print(f'{i}.',asal['STORE NAME'], 'mengirim 1 barang ke' ,tujuan['STORE NAME'])
    
    return df_asal, df_tujuan, log_rotasi

# Jalankan rotasi stok
df_asal_updated, df_tujuan_updated, log_rotasi  = rotate_stock(df_asal, df_tujuan, log_rotasi)
log_rotasi = log_rotasi.groupby(["KOTA", "ARTICLE", "STORE ASAL", "STORE TUJUAN"], as_index=False).agg({
    "BARANG TEROTASI": "sum"
})
log_rotasi["TOTAL"] = log_rotasi.groupby(["KOTA", "STORE ASAL", "ARTICLE"])["BARANG TEROTASI"].transform("sum")
log_rotasi = log_rotasi[["KOTA", "STORE ASAL", "ARTICLE", "TOTAL", "BARANG TEROTASI", "STORE TUJUAN"]]



# Merge pertama, menambahkan kolom asal sebelum rotasi
merged_data = pd.merge(
    log_rotasi, 
    df_asal[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], 
    left_on='STORE ASAL', right_on='STORE NAME', 
    how='left',
    suffixes=('', '_asal')
).drop(columns=['STORE NAME'])  # Menghapus duplikat STORE NAME

# Merge kedua, menambahkan kolom asal setelah rotasi
merged_data = pd.merge(
    merged_data, 
    df_asal_updated[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], 
    left_on='STORE ASAL', right_on='STORE NAME', 
    how='left',
    suffixes=('_asal_sebelum', '_asal_setelah')
).drop(columns=['STORE NAME'])

# Merge ketiga, menambahkan kolom tujuan sebelum rotasi
merged_data = pd.merge(
    merged_data, 
    df_tujuan[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], 
    left_on='STORE TUJUAN', right_on='STORE NAME', 
    how='left',
    suffixes=('', '_tujuan_sebelum')
).drop(columns=['STORE NAME'])

# Merge keempat, menambahkan kolom tujuan setelah rotasi
merged_data = pd.merge(
    merged_data, 
    df_tujuan_updated[['STORE NAME', 'Stock', 'Sales 30 days', 'DOS 30 days']], 
    left_on='STORE TUJUAN', right_on='STORE NAME', 
    how='left',
    suffixes=('_tujuan_sebelum', '_tujuan_setelah')
).drop(columns=['STORE NAME'])

merged_data.rename(columns={
    "Stock_asal_sebelum": "Stock",
    "Sales 30 days_asal_sebelum": "Sales",
    "DOS 30 days_asal_sebelum": "Dos",
    
    "Stock_asal_setelah": "Stock Akhir",
    "Sales 30 days_asal_setelah": "Sales",
    "DOS 30 days_asal_setelah": "Dos Akhir",
    
    "Stock_tujuan_sebelum": "Stock ",
    "Sales 30 days_tujuan_sebelum": "Sales ",
    "DOS 30 days_tujuan_sebelum": "Dos ",
    
    "Stock_tujuan_setelah": "Stock Akhir ",
    "Sales 30 days_tujuan_setelah": "Sales ",
    "DOS 30 days_tujuan_setelah": "Dos Akhir ",
    }, inplace=True)

merged_data = merged_data[["KOTA", "ARTICLE", "STORE ASAL", "Stock", "Sales", "Dos", "Stock Akhir", "Sales", "Dos Akhir", 
                           "TOTAL", "BARANG TEROTASI", "STORE TUJUAN", "Stock ", "Sales ", "Dos ", "Stock Akhir ", "Sales ", "Dos Akhir "]]


with pd.ExcelWriter('TES3.xlsx') as writer:
    merged_data.to_excel(writer, sheet_name='ByKota', index=False)