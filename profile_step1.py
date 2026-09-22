import pandas as pd

# Read the .xlsb file - Sheet1
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print(f"Dtypes:\n{df.dtypes}")
print(f"\nFirst 5 rows:")
print(df.head())
print(f"\nNull counts:")
print(df.isnull().sum())