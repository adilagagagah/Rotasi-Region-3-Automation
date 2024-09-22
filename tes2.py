import pandas as pd

# Data yang diberikan
data = {
    "STORE ASAL": ["ERAFONE DAYEUHKOLOT BANDUNG (E870)", "ERAFONE DAYEUHKOLOT BANDUNG (E870)",
                   "MEGASTORE RUKO BOJONGSOANG (M137)", "MEGASTORE RUKO BOJONGSOANG (M137)",
                   "ERAFONE RUKO CICALENGKA (E817)", "ERAFONE RUKO CICALENGKA (E817)",
                   "ERAFONE RUKO CIWIDEY BANDUNG (E916)", "ERAFONE RUKO CIWIDEY BANDUNG (E916)",
                   "ERAFONE TAMAN KOPO INDAH 2 (E851)", "ERAFONE TAMAN KOPO INDAH 2 (E851)"],
    "ARTICLE": ["Samsung Galaxy Tab A9 4/64"] * 10,
    "STOCK ASAL": [0, 0, 0, 0, 0, 0, 1, 1, 1, 1],
    "SALES ASAL": [2, 2, 2, 2, 1, 1, 3, 3, 2, 2],
    "DOS ASAL": [0, 0, 0, 0, 0, 0, 10, 10, 15, 15],
    "STORE ROTASI": ["ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE AND MORE SOREANG BANDUNG (M210)",
                     "ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE AND MORE SOREANG BANDUNG (M210)",
                     "ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE AND MORE SOREANG BANDUNG (M210)",
                     "ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE AND MORE SOREANG BANDUNG (M210)",
                     "ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE AND MORE SOREANG BANDUNG (M210)"],
    "STOCK ROTASI": [4, 2, 4, 2, 4, 2, 4, 2, 4, 2],
    "SALES ROTASI": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    "DOS ROTASI": [120, 60, 120, 60, 120, 60, 120, 60, 120, 60]
}

data_new = {
    "STORE ASAL": ["ERAFONE AND MORE SOREANG BANDUNG (M210)"] * 6 + ["ERAFONE DANGDEUR RANCAEKEK BANDUNG (E918)"] * 6,
    "ARTICLE": ["VIVO V40 12/256GB"] * 12,
    "STOCK ASAL": [0] * 12,
    "SALES ASAL": [1] * 12,
    "DOS ASAL": [0] * 12,
    "STORE ROTASI": ["MEGASTORE RUKO BOJONGSOANG (M137)", "CARREFOUR TRANSMART BUAH BATU SQUARE (G270)",
                     "ERAFONE RUKO KOPO BANDUNG (E260)", "ERAFONE DAYEUHKOLOT BANDUNG (E870)",
                     "ERAFONE RUKO BANJARAN (E571)", "ERAFONE RUKO LASWI CIPARAY (E633)"] * 2,
    "STOCK ROTASI": [11, 4, 6, 2, 5, 3] * 2,
    "SALES ROTASI": [2, 1, 3, 1, 3, 2] * 2,
    "DOS ROTASI": [165, 120, 60, 60, 50, 45] * 2
}

data_2 = {
    "STORE ASAL": [
        "ERAFONE DAYEUHKOLOT BANDUNG (E870)", "ERAFONE DAYEUHKOLOT BANDUNG (E870)",
        "ERAFONE DAYEUHKOLOT BANDUNG (E870)", "ERAFONE DAYEUHKOLOT BANDUNG (E870)",
        "ERAFONE TAMAN KOPO INDAH 2 (E851)", "ERAFONE TAMAN KOPO INDAH 2 (E851)",
        "ERAFONE TAMAN KOPO INDAH 2 (E851)", "ERAFONE TAMAN KOPO INDAH 2 (E851)",
        "ERAFONE DANGDEUR RANCAEKEK BANDUNG (E918)", "ERAFONE DANGDEUR RANCAEKEK BANDUNG (E918)",
        "ERAFONE DANGDEUR RANCAEKEK BANDUNG (E918)", "ERAFONE DANGDEUR RANCAEKEK BANDUNG (E918)",
        "ERAFONE RUKO CICALENGKA (E817)", "ERAFONE RUKO CICALENGKA (E817)",
        "ERAFONE RUKO CICALENGKA (E817)", "ERAFONE RUKO CICALENGKA (E817)",
        "ERAFONE RUKO LASWI CIPARAY (E633)", "ERAFONE RUKO LASWI CIPARAY (E633)",
        "ERAFONE RUKO LASWI CIPARAY (E633)", "ERAFONE RUKO LASWI CIPARAY (E633)",
        "ERAFONE RUKO BALEENDAH BANDUNG (E733)", "ERAFONE RUKO BALEENDAH BANDUNG (E733)",
        "ERAFONE RUKO BALEENDAH BANDUNG (E733)", "ERAFONE RUKO BALEENDAH BANDUNG (E733)",
        "ERAFONE RUKO CIWIDEY BANDUNG (E916)", "ERAFONE RUKO CIWIDEY BANDUNG (E916)",
        "ERAFONE RUKO CIWIDEY BANDUNG (E916)", "ERAFONE RUKO CIWIDEY BANDUNG (E916)",
        "MEGASTORE RUKO BOJONGSOANG (M137)", "MEGASTORE RUKO BOJONGSOANG (M137)",
        "MEGASTORE RUKO BOJONGSOANG (M137)", "MEGASTORE RUKO BOJONGSOANG (M137)"
    ],
    "ARTICLE": ["VIVO Y03 4/64GB"] * 32,
    "STOCK ASAL": [0] * 24 + [1] * 4 + [3] * 4,
    "SALES ASAL": [6] * 4 + [5] * 4 + [3] * 12 + [4] * 4 + [7] * 4 + [6] * 4,
    "DOS ASAL": [0] * 20 + [8] * 4 + [13] * 4 + [15] * 4,
    "STORE ROTASI": [
        "ERAFONE AND MORE SOREANG BANDUNG (M210)", "ERAFONE CILAMPENI KETAPANG BANDUNG (E979)",
        "ERAFONE RUKO MAJALAYA BANDUNG (F072)", "ERAFONE RUKO BANJARAN (E571)"
    ] * 8,  # Replikasi ini harus sesuai dengan jumlah total baris di kolom lain
    "STOCK ROTASI": [8, 4, 2, 3] * 8,
    "SALES ROTASI": [2, 1, 1, 2] * 8,
    "DOS ROTASI": [120, 120, 60, 45] * 8
}

# Membuat DataFrame dari data baru
df1 = pd.DataFrame(data)
df2 = pd.DataFrame(data_new)
df3 = pd.DataFrame(data_2)

# Menggabungkan DataFrame
df = pd.concat([df1, df2, df3], ignore_index=True)
# print(df)

# Fungsi untuk proses sorting dan pairing untuk setiap artikel
def process_article(dataframe, article):
    df_article = dataframe[dataframe['ARTICLE'] == article]
    sorted_asal = df_article.sort_values(by=['DOS ASAL', 'SALES ASAL'], ascending=[True, False]).drop_duplicates('STORE ASAL')
    sorted_rotasi = df_article.sort_values(by=['DOS ROTASI', 'STOCK ROTASI'], ascending=[False, True]).drop_duplicates('STORE ROTASI')
    sorted_asal.reset_index(drop=True, inplace=True)
    sorted_rotasi.reset_index(drop=True, inplace=True)
    result = pd.DataFrame({
        "STORE ASAL": sorted_asal['STORE ASAL'],
        "ARTICLE": sorted_asal['ARTICLE'],
        "STOCK ASAL": sorted_asal['STOCK ASAL'],
        "SALES ASAL": sorted_asal['SALES ASAL'],
        "DOS ASAL": sorted_asal['DOS ASAL'],
        "STORE ROTASI": sorted_rotasi['STORE ROTASI'],
        "STOCK ROTASI": sorted_rotasi['STOCK ROTASI'],
        "SALES ROTASI": sorted_rotasi['SALES ROTASI'],
        "DOS ROTASI": sorted_rotasi['DOS ROTASI']
    })
    return result

# List of articles to process
articles = df['ARTICLE'].unique()

# Applying the function to each article
results = pd.concat([process_article(df, article) for article in articles])

# Menghapus baris dengan nilai NaN (jika ada)
results.dropna(inplace=True)
results.sort_values(by=['ARTICLE'], ascending=[True])
results.reset_index(drop=True, inplace=True)

# Output the results
print(results)