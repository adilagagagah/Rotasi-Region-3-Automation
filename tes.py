import numpy as np
import pandas as pd

def load_data(file_path, master_path):
    df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
    df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)
    df_master_selected = df_master[['SITE CODE', 'PT']]
    return pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')

def preprocess_data(df):
    df = df[["KOTA", "SITE CODE", "PT", "STORE NAME", "Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]
    df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days', 'DOS 30 days']].applymap(lambda x: np.nan if x < 0 else x)
    return df

def filter_data(df, pt, kota, dos_threshold):
    return df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] <= dos_threshold)]

def find_rotation_candidates(df, pt, kota, min_dos_code_article):
    return df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] >= 45) & (df['Article code no color'].isin(min_dos_code_article))]

def process_rotation(min_dos_df, rotation_df):
    result_rows = []
    for code in min_dos_df['SITE CODE'].unique():
        code_df = min_dos_df[min_dos_df['SITE CODE'] == code]
        store_info = code_df.iloc[0]['STORE NAME'] + f" ({code})"
        for _, row in code_df.iterrows():
            article = row['Article code no color']
            rotation_articles = rotation_df[rotation_df['Article code no color'] == article]
            for _, rot_row in rotation_articles.iterrows():
                result_rows.append([
                    store_info, article, row['Stock'], row['Sales 30 days'], row['DOS 30 days'],
                    rot_row['STORE NAME'] + f" ({rot_row['SITE CODE']})", rot_row['Stock'], rot_row['Sales 30 days'], rot_row['DOS 30 days']
                ])
    return pd.DataFrame(result_rows, columns=[
        'STORE ASAL', 'ARTICLE', 'STOCK ASAL', 'SALES ASAL', 'DOS ASAL',
        'STORE ROTASI', 'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
    ])

# Load and preprocess data
file_path = "../ROTASI REGION 3/Data Stock 20 Sep 2024.xlsx"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"
df = load_data(file_path, master_path)
df = preprocess_data(df)

# Filter data for specific PT and city
pt = 'EAR'
kota = sorted(df["KOTA"].unique())[0]  # KAB. BANDUNG
min_dos_df = filter_data(df, pt, kota, 15)
rotation_df = find_rotation_candidates(df, pt, kota, min_dos_df['Article code no color'].unique())

# Process rotation for min DOS and rotation candidates
result_df = process_rotation(min_dos_df, rotation_df)
result_df.dropna(inplace=True)
result_df.sort_values(by=['ARTICLE'], inplace=True)
result_df.reset_index(drop=True, inplace=True)

# Save results
output_file = "result.xlsx"
result_df.to_excel(output_file, index=False)

print(result_df)
