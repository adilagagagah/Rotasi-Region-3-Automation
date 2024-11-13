import numpy as np
import pandas as pd
import math
from tqdm import tqdm

def custom_round(n):
    return math.ceil(n) if n % 1 >= 0.5 else math.floor(n)

def calculate_dos(stock, sales):
    return custom_round((stock / sales * 30)) if sales != 0 else 0

def split_store(df, category_col_name):
    df_asal = df[
        ((df['DOS 30 days'] >= 30) & (df['Stock'] >= 2)) |
        ((df['Sales 30 days'] >= 0) & (df['Stock'] >= 2))
    ]
    df_asal = df_asal[[f"{category_col_name}", "STORE", "Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
    
    df_tujuan = df[
        (df['DOS 30 days'] <= 23) &
        (df['Sales 30 days'] != 0)
    ]
    df_tujuan = df_tujuan[[f"{category_col_name}", "STORE", "Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
    
    return df_asal, df_tujuan

def rotate_stock(df_asal, df_tujuan, category_col_name):
    log_rotasi = pd.DataFrame(columns=[f"{category_col_name}", "STORE ASAL", "ARTICLE", "BARANG TEROTASI", "STORE TUJUAN"])
    i = 0
    # store asal dan tujuan tidak boleh kosong
    while True:
        dos_df_asal = list(df_asal['DOS 30 days'])
        dos_df_tujuan = list(df_tujuan['DOS 30 days'])
        
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

        # buat kondisi dimana rotasi stok di hentikan
        if (simulated_doses[selected_index] < 23) or all(x >= 23 for x in dos_df_tujuan):
            # dos asal yang disimulasikan akan di rotasi kurang dari 23, stop
            # semua store tujuan sudah memiliki dos lebih dari 23, stop
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
            f"{category_col_name}": [asal[f'{category_col_name}']],
            "STORE ASAL": [f"{asal['STORE']}"],
            "ARTICLE": [asal['Article code no color']],
            "BARANG TEROTASI": [1],
            "STORE TUJUAN": [f"{tujuan['STORE']}"]
        })
        
        log_rotasi = pd.concat([log_rotasi, new_log_entry], ignore_index=True)
        
        i += 1
        # print(f'{i}.',asal['STORE'], 'ke' ,tujuan['STORE'])
    
    return df_asal, df_tujuan, log_rotasi

############################
def split_log_rotasi(log_rotasi, category_col_name):
    # Membuat tabel store asal beserta stok yang bisa dirotasikan
    tabel_asal = log_rotasi.groupby([f"{category_col_name}", "ARTICLE", "STORE ASAL"])['BARANG TEROTASI'].sum().reset_index()
    tabel_asal.rename(columns={'BARANG TEROTASI': 'TOTAL TEROTASI'}, inplace=True)

    # Membuat tabel store tujuan beserta stok yang dibutuhkan
    tabel_tujuan = log_rotasi.groupby([f"{category_col_name}", "ARTICLE", "STORE TUJUAN"])['BARANG TEROTASI'].sum().reset_index()
    tabel_tujuan.rename(columns={'BARANG TEROTASI': 'TOTAL BUTUH'}, inplace=True)

    return tabel_asal, tabel_tujuan

def reorder(log_rotasi):
    df_sorted = log_rotasi.sort_values(by=['TOTAL ASAL'], ascending=[False]).reset_index(drop=True)

    df_temp = pd.DataFrame(df_sorted.iloc[0]).transpose()
    df_sorted.drop(0, inplace=True)

    threshold = 0
    while not df_sorted.empty:
        df_sorted.reset_index(drop=True, inplace=True)
        last_name = df_temp['STORE TUJUAN'].iloc[-1] if not df_temp.empty else None

        for index_sorted, row_sorted in df_sorted.iterrows():
            if last_name == row_sorted['STORE TUJUAN']:
                df_temp = pd.concat([df_temp, pd.DataFrame([row_sorted])], ignore_index=True)
                df_sorted.drop(index_sorted, inplace=True)
                break

        new_threshold = len(df_temp)
        if threshold == new_threshold:
            new_row = df_sorted.iloc[[0]] 
            df_temp = pd.concat([df_temp, new_row], ignore_index=True)
            df_sorted.drop(0, inplace=True)
            df_sorted.reset_index(drop=True, inplace=True)
        threshold = new_threshold

    return df_temp

def process_pairing_1(tabel_asal, tabel_tujuan, paired, category_col_name):
    for index_asal, row_asal in tabel_asal.iterrows():
        for index_tujuan, row_tujuan in tabel_tujuan.iterrows():
            if row_asal['TOTAL TEROTASI'] == row_tujuan['TOTAL BUTUH']:
                new_entry = pd.DataFrame({
                    f"{category_col_name}": [row_asal[f"{category_col_name}"]],
                    "STORE ASAL": [row_asal['STORE ASAL']],
                    "ARTICLE": [row_asal['ARTICLE']],
                    "TOTAL ASAL": [row_asal['TOTAL TEROTASI']],
                    "BARANG TEROTASI": [min(row_asal['TOTAL TEROTASI'], row_tujuan['TOTAL BUTUH'])],
                    "TOTAL TUJUAN": [row_tujuan['TOTAL BUTUH']],
                    "STORE TUJUAN": [row_tujuan['STORE TUJUAN']]
                })
                paired = pd.concat([paired, new_entry], ignore_index=True)
                tabel_asal.drop(index_asal, inplace=True)
                tabel_tujuan.drop(index_tujuan, inplace=True)
                
                return tabel_asal, tabel_tujuan, paired 
    return tabel_asal, tabel_tujuan, paired

def process_pairing_2(tabel_asal, tabel_tujuan, paired, category_col_name):
    tabel_asal['Simulated TOTAL TEROTASI'] = tabel_asal['TOTAL TEROTASI']
    tabel_tujuan['Simulated TOTAL BUTUH'] = tabel_tujuan['TOTAL BUTUH']

    paired_temp = pd.DataFrame(columns=[f"{category_col_name}", 'STORE ASAL', 'ARTICLE', 'TOTAL ASAL', 'BARANG TEROTASI', 'TOTAL TUJUAN', 'STORE TUJUAN'])

    while True:
        tabel_asal = tabel_asal.sort_values(by=['Simulated TOTAL TEROTASI'], ascending=[False]).reset_index(drop=True)
        tabel_tujuan = tabel_tujuan.sort_values(by=['Simulated TOTAL BUTUH'], ascending=[False]).reset_index(drop=True)

        new_entry = pd.DataFrame({
            f"{category_col_name}" : [tabel_asal.loc[0, f"{category_col_name}"]],
            'STORE ASAL' : [tabel_asal.loc[0, 'STORE ASAL']],
            'ARTICLE' : [tabel_asal.loc[0, 'ARTICLE']],
            'TOTAL ASAL' : [tabel_asal.loc[0, 'TOTAL TEROTASI']],
            'BARANG TEROTASI' : min(tabel_asal.loc[0, 'Simulated TOTAL TEROTASI'], tabel_tujuan.loc[0, 'Simulated TOTAL BUTUH']),
            'TOTAL TUJUAN' : [tabel_tujuan.loc[0, 'TOTAL BUTUH']],
            'STORE TUJUAN' : [tabel_tujuan.loc[0, 'STORE TUJUAN']]
        })
        
        paired_temp = pd.concat([paired_temp, new_entry], ignore_index=True)

        tabel_asal.loc[0, 'Simulated TOTAL TEROTASI'] -= new_entry['BARANG TEROTASI'].item()
        tabel_tujuan.loc[0, 'Simulated TOTAL BUTUH'] -= new_entry['BARANG TEROTASI'].item()
            
        if all(x == 0 for x in tabel_tujuan['Simulated TOTAL BUTUH']) : 
            tabel_asal = pd.DataFrame(columns=tabel_asal.columns)
            tabel_tujuan = pd.DataFrame(columns=tabel_tujuan.columns)
            break        

    paired_temp = reorder(paired_temp)
    paired = pd.concat([paired, paired_temp], ignore_index=True)
    
    return tabel_asal, tabel_tujuan, paired

def efficiency(log_rotasi, category_col_name):
    paired = pd.DataFrame(columns=[f"{category_col_name}", 'STORE ASAL', 'ARTICLE', 'TOTAL ASAL', 'BARANG TEROTASI', 'TOTAL TUJUAN', 'STORE TUJUAN'])
    tabel_asal, tabel_tujuan = split_log_rotasi(log_rotasi, category_col_name)

    # Loop hingga salah satu atau kedua tabel kosong
    threshold = 0
    while not tabel_asal.empty and not tabel_tujuan.empty:
        tabel_asal, tabel_tujuan, paired = process_pairing_1(tabel_asal, tabel_tujuan, paired, category_col_name)
        
        new_threshold = len(tabel_asal) * len(tabel_tujuan)
        if threshold == new_threshold and new_threshold != 0 :
            tabel_asal, tabel_tujuan, paired = process_pairing_2(tabel_asal, tabel_tujuan, paired, category_col_name)
        threshold = new_threshold

    return paired

############################
def add_dos_information(df_asal, df_asal_updated, df_tujuan, df_tujuan_updated, log_rotasi, category_col_name):
    merged_data = pd.merge(
        log_rotasi, df_asal[['STORE', 'Stock', 'Sales 30 days', 'DOS 30 days']],left_on='STORE ASAL', right_on='STORE', 
        how='left', suffixes=('', '_asal')).drop(columns=['STORE'])  

    merged_data = pd.merge(
        merged_data, df_asal_updated[['STORE', 'Stock', 'Sales 30 days', 'DOS 30 days']], left_on='STORE ASAL', right_on='STORE', 
        how='left',suffixes=('_asal_sebelum', '_asal_setelah')).drop(columns=['STORE'])

    merged_data = pd.merge(
        merged_data, df_tujuan[['STORE', 'Stock', 'Sales 30 days', 'DOS 30 days']], left_on='STORE TUJUAN', right_on='STORE', 
        how='left',suffixes=('', '_tujuan_sebelum')).drop(columns=['STORE'])

    merged_data = pd.merge(
        merged_data, df_tujuan_updated[['STORE', 'Stock', 'Sales 30 days', 'DOS 30 days']], left_on='STORE TUJUAN', right_on='STORE', 
        how='left',suffixes=('_tujuan_sebelum', '_tujuan_setelah')).drop(columns=['STORE'])

    merged_data.rename(columns={
        "Stock_asal_sebelum": "Stock", "Sales 30 days_asal_sebelum": "Sales", "DOS 30 days_asal_sebelum": "Dos",
        "Stock_asal_setelah": "Stock Akhir", "Sales 30 days_asal_setelah": "Sales Akhir", "DOS 30 days_asal_setelah": "Dos Akhir",
        "Stock_tujuan_sebelum": "Stock Tujuan", "Sales 30 days_tujuan_sebelum": "Sales Tujuan", "DOS 30 days_tujuan_sebelum": "Dos Tujuan",
        "Stock_tujuan_setelah": "Stock Akhir Tujuan", "Sales 30 days_tujuan_setelah": "Sales Akhir Tujuan", "DOS 30 days_tujuan_setelah": "Dos Akhir Tujuan"
    }, inplace=True)

    merged_data = merged_data[[
        f"{category_col_name}", "STORE ASAL", "ARTICLE", "Stock", "Sales", "Dos", "Stock Akhir", "Sales Akhir", "Dos Akhir", 
        "TOTAL ASAL", "BARANG TEROTASI", "TOTAL TUJUAN", "STORE TUJUAN", "Stock Tujuan", "Sales Tujuan", "Dos Tujuan", 
        "Stock Akhir Tujuan", "Sales Akhir Tujuan", "Dos Akhir Tujuan"
    ]]

    return merged_data

def rotate_by(df_pt, category_list, pt, category_col_name):
    # category bisa by kota ataupun by tsh
    rotasi = pd.DataFrame(columns=[
        f"{category_col_name}", "STORE ASAL", "ARTICLE", "Stock", "Sales", "Dos", "Stock Akhir", "Sales Akhir", "Dos Akhir", 
        "TOTAL ASAL", "BARANG TEROTASI", "TOTAL TUJUAN", "STORE TUJUAN", "Stock Tujuan", "Sales Tujuan", "Dos Tujuan", 
        "Stock Akhir Tujuan", "Sales Akhir Tujuan", "Dos Akhir Tujuan"
    ])

    for category in tqdm(category_list, desc=f'Rotating {pt} by {category_col_name}'):
        df_category = df_pt[df_pt[f'{category_col_name}'] == category]
        articles = df_category["Article code no color"].unique()

        for article in articles:        
            df_article = df_category[df_category['Article code no color'] == article]
            df_asal, df_tujuan = split_store(df_article, category_col_name)
            if not df_asal.empty and not df_tujuan.empty:
                df_asal_updated, df_tujuan_updated, log_rotasi = rotate_stock(df_asal, df_tujuan, category_col_name)
                log_rotasi = efficiency(log_rotasi, category_col_name)
                log_rotasi = add_dos_information(df_asal, df_asal_updated, df_tujuan, df_tujuan_updated, log_rotasi, category_col_name)

                rotasi = pd.concat([rotasi, log_rotasi], ignore_index=True)

    # matikan jika memerlukan data total asal dan total tujuan
    rotasi = rotasi.drop(columns=['TOTAL ASAL', 'TOTAL TUJUAN'])
    return rotasi

def merge_cells(worksheet, result_df, start_row, merge_format):
    result_df = result_df.reset_index(drop=True)
    result_df['dummy'] = result_df['STORE ASAL'] + "_" + result_df['ARTICLE']
    same_store_article = [(i,dummy) for i, dummy in enumerate(result_df['dummy'], start=start_row)]
    indeks_dummy_list = []
    dummy_list = []
    
    for row in same_store_article:
        indeks_dummy_list.append(row)
        dummy_list.append(row[1])
            
        if len(set(dummy_list)) > 1:
            dummy_list_eksekusi = indeks_dummy_list[:-1]
            n_dummy_list_eksekusi = len(dummy_list_eksekusi)
            i = dummy_list_eksekusi[0][0]
            base_index = i - start_row
            worksheet.merge_range(i, 1, i + n_dummy_list_eksekusi - 1, 1, result_df['STORE ASAL'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 2, i + n_dummy_list_eksekusi - 1, 2, result_df['ARTICLE'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 3, i + n_dummy_list_eksekusi - 1, 3, result_df['Stock'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 4, i + n_dummy_list_eksekusi - 1, 4, result_df['Sales'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 5, i + n_dummy_list_eksekusi - 1, 5, result_df['Dos'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 6, i + n_dummy_list_eksekusi - 1, 6, result_df['Stock Akhir'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 7, i + n_dummy_list_eksekusi - 1, 7, result_df['Sales Akhir'].iloc[base_index], merge_format)
            worksheet.merge_range(i, 8, i + n_dummy_list_eksekusi - 1, 8, result_df['Dos Akhir'].iloc[base_index], merge_format)

            indeks_dummy_list = [indeks_dummy_list[-1]]
            dummy_list = [dummy_list[-1]]


#############################################
# LOAD DATA
file_name = "Data Stock 9 Oct 2024.xlsx"
output_file = file_name[11:-5]
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)
df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')
df['STORE'] = df['STORE NAME'] + " (" + df['SITE CODE'].astype(str) + ")"
df = df.drop(columns=['STORE NAME', 'SITE CODE'])
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].clip(lower=0).fillna(0)

#############################################
# KODE UTAMA
with pd.ExcelWriter(f'RotasiR3 {output_file}.xlsx', engine='xlsxwriter') as writer:
    for pt in tqdm(['EAR', 'DCM', 'NASA'], desc='Processing PT'):
        for category_col_name in ['KOTA', 'TSH']:
            df_pt = df[df['PT'] == pt]
            
            jumlah_baris_kosong = df_pt['Article code no color'].isnull().sum()
            print(f"\n{pt} : Jumlah baris yang kolom 'Article'-nya NaN: {jumlah_baris_kosong}")
            df_pt = df_pt.dropna(subset=['Article code no color'])
            print(df_pt.shape)

            category_list = df_pt[f"{category_col_name}"].unique()
            rotasi = rotate_by(df_pt, category_list, pt, category_col_name)

            # Menyimpan hasil ke sheet yang berbeda dalam satu file Excel
            rotasi.to_excel(writer, sheet_name=f'{pt} By {category_col_name}', index=False)

            # merge cell
            workbook = writer.book
            worksheet = writer.sheets[f'{pt} By {category_col_name}']
            
            merge_format = workbook.add_format({'valign': 'vcenter',})
            merge_cells(worksheet, rotasi, 1, merge_format)

#############################################
# PROBLEM : TERDAPAT ARTICLE CODE NO COLOR YANG DUPLIKAT UNTUK SATU TOKO DI SATU KOTA
# CONTOH : KOTA SEMARANG, M095, Redmi Note 13Pro5G 12/512, MEMILIKI STOK SALES DOS YANG BERBEDA

#############################################