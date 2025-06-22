import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
import joblib

excel_file = './SupplyChainEmissionFactorsforUSIndustriesCommodities.xlsx'
years = list(range(2010, 2017))
print(f"Selected year: {years[2]}")

def load_sheet(year, detail_type):
    sheet_name = f"{year}_Detail_{detail_type}"
    return pd.read_excel(excel_file, sheet_name=sheet_name)

try:
    df_1 = load_sheet(years[0], "Commodity")
    print("First Commodity Sheet:")
    print(df_1.head())
    df_2 = load_sheet(years[0], "Industry")
    print("First Industry Sheet:")
    print(df_2.head())
except Exception as e:
    print(f"Error reading initial sheets: {e}")

all_data = []

for year in years:
    for detail_type, code_col, name_col, source in [
        ("Commodity", "Commodity Code", "Commodity Name", "Commodity"),
        ("Industry", "Industry Code", "Industry Name", "Industry")
    ]:
        try:
            df = load_sheet(year, detail_type)
            df['Source'] = source
            df['Year'] = year
            df.columns = df.columns.str.strip()
            df.rename(columns={
                code_col: 'Code',
                name_col: 'Name'
            }, inplace=True)
            all_data.append(df)
        except Exception as e:
            print(f"Error processing {detail_type} for year {year}: {e}")

if all_data:
    df = pd.concat(all_data, ignore_index=True)
    print("\nCombined Data Sample:")
    print(df.head(10))
    print(f"Total rows: {len(df)}")
    print("Columns:", df.columns.tolist())
    print("Missing values per column:\n", df.isnull().sum())
    # Example visualization
    if 'Year' in df.columns and 'Source' in df.columns:
        plt.figure(figsize=(8,4))
        sns.countplot(data=df, x='Year', hue='Source')
        plt.title('Record Count by Year and Source')
        plt.show()
else:
    print("No data loaded from Excel.")
