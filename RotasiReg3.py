import numpy as np
import pandas as pd

file_name = "Data Stock 20 Sep 2024.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
master_path = "../ROTASI REGION 3/MASTER REGION STO_NEXT VERSION (62).xlsx"

# Load data
df_origin = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
df_master = pd.read_excel(master_path, sheet_name="REGION 3", header=1)

# Merge data with SITE CODE and PT information
df_master_selected = df_master[['SITE CODE', 'PT']]
df = pd.merge(df_origin, df_master_selected, on='SITE CODE', how='left')

# Filter relevant columns
df = df[["AREA", "KOTA", "TSH", "SITE CODE", "PT", "STORE NAME", "Article code no color", "Stock", "Sales 30 days", "DOS 30 days"]]

# Replace negative values in 'Sales' and 'DOS' columns with NaN
df[['Sales 30 days', 'DOS 30 days']] = df[['Sales 30 days','DOS 30 days']].map(lambda x: np.nan if x < 0 else x)

result_rows = []

# Define parameters for filtering
pt = 'EAR'
list_kota = sorted(list(set(df["KOTA"])))
kota = list_kota[0]  # Example: KAB. BANDUNG

# Find stores with low DOS (<= 15)
min_dos_df = df[(df['PT'] == pt) & (df['KOTA'] == kota) & (df['DOS 30 days'] <= 15)]
min_dos_df = min_dos_df.sort_values(by=['DOS 30 days', 'Sales 30 days'], ascending=[True, False])
min_dos_code_unique = min_dos_df['SITE CODE'].drop_duplicates().tolist()

# Iterate over stores with low DOS
for code in min_dos_code_unique:
    min_dos_code_df = min_dos_df[min_dos_df['SITE CODE'] == code]
    min_dos_store = list(min_dos_code_df['STORE NAME'])[0] + f" ({code})"
    min_dos_article = min_dos_code_df['Article code no color'].drop_duplicates().tolist()

    # Find rotation candidates (stores with DOS >= 45) for matching articles
    rotation_df = df[(df['KOTA'] == kota) & (df['DOS 30 days'] >= 45) & (df['Article code no color'].isin(min_dos_article))]

    for article in min_dos_article:
        min_dos_article_data = min_dos_df[min_dos_df['Article code no color'] == article]
        min_stock = min_dos_article_data['Stock'].values[0]
        min_sales = min_dos_article_data['Sales 30 days'].values[0]
        min_dos = min_dos_article_data['DOS 30 days'].values[0]

        # Get corresponding rotation data
        rotation_article_data = rotation_df[rotation_df['Article code no color'] == article]
        rotation_article_data = rotation_article_data.sort_values(by=['DOS 30 days', 'Stock'], ascending=[False, False])

        # Perform rotation based on stock-sales difference from rotation store
        for _, row in rotation_article_data.iterrows():
            rotation_store = row['STORE NAME'] + f" ({row['SITE CODE']})"
            rotation_stock = row['Stock']
            rotation_sales = row['Sales 30 days']
            rotation_dos = row['DOS 30 days']

            # Calculate stock-sales difference
            stock_sales_diff = rotation_stock - rotation_sales

            if stock_sales_diff > 0:
                # Only allow rotation based on stock-sales difference
                result_rows.append([min_dos_store, rotation_store, article, min_stock, min_sales, min_dos,
                                    rotation_stock, rotation_sales, rotation_dos])

                # Decrease the stock-sales difference for subsequent rotations
                stock_sales_diff -= 1

            # Stop if no more rotations are allowed
            if stock_sales_diff <= 0:
                break

# Save results to DataFrame with new column order
result_df = pd.DataFrame(result_rows, columns=[
    'STORE ASAL', 'ROTASI DARI', 'ARTICLE', 'STOCK ASAL', 'SALES ASAL', 'DOS ASAL',
    'STOCK ROTASI', 'SALES ROTASI', 'DOS ROTASI'
])

result_df = result_df.drop_duplicates(subset=['ARTICLE'], keep='first')

# Sort based on STORE ASAL, ROTASI DARI, and ARTICLE
result_df_sorted = result_df.sort_values(by=['STORE ASAL', 'ROTASI DARI', 'ARTICLE'])

# Insert blank rows between different STORE ASAL and ROTASI DARI combinations
combined_rows = []
previous_store_asal = None
previous_rotasi_dari = None

for idx, row in result_df_sorted.iterrows():
    current_store_asal = row['STORE ASAL']
    current_rotasi_dari = row['ROTASI DARI']
    
    # Check if the STORE ASAL or ROTASI DARI has changed
    if current_store_asal != previous_store_asal or current_rotasi_dari != previous_rotasi_dari:
        # Append a blank row for separation
        if previous_store_asal is not None:
            combined_rows.append([''] * len(row))
    
    # Append the current row
    combined_rows.append(row.tolist())
    
    # Update previous values
    previous_store_asal = current_store_asal
    previous_rotasi_dari = current_rotasi_dari

# Convert combined_rows to DataFrame
final_df = pd.DataFrame(combined_rows, columns=result_df_sorted.columns)

# Save both sheets to an Excel file
with pd.ExcelWriter("Result_Stock_Rotation_sorted_with_blanks.xlsx") as writer:
    # Original data
    result_df.to_excel(writer, sheet_name="Original Rotasi", index=False)
    
    # Sorted data by store asal and rotasi dari with blank rows
    final_df.to_excel(writer, sheet_name="Sorted Rotasi", index=False)

print("DataFrame telah disimpan ke file dengan sheet baru berisi data yang diurutkan dan jeda baris kosong antar rotasi.")
