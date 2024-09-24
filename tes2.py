import numpy as np
import pandas as pd

file_name = "EAR_result.xlsx"
file_path = f"../ROTASI REGION 3/{file_name}"
df = pd.read_excel(file_path)

cetak = df["KOTA"].unique()

print(len(cetak))
