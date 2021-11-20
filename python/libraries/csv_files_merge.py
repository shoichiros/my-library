import pandas as pd
import glob


folder_path = r"C:\Users\falcon\Desktop\test"
files = glob.glob(folder_path + "\*.csv")
lists = []

for file in files:
    df = pd.read_csv(file, index_col=None, header=0)
    lists.append(df)

df = pd.concat(lists, axis=0, ignore_index=True)
df.to_csv("output.csv", index=False, encoding='utf-8')
