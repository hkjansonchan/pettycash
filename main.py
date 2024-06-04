import pandas as pd

file = r"C:\Users\hkjan\Downloads\IvyWalletReport-20240604-2215.csv"
df = pd.read_csv(file)
print(df.to_string())