import pandas as pd
from sklearn.preprocessing import MinMaxScaler

# Function to clean and convert columns to numerical types if possible
def convert_to_numeric(df):
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

excel_path = 'Sector-Data.xlsx'  # Change this to the path where your Excel file is located
sheet_names = pd.ExcelFile(excel_path).sheet_names

sector_data = {}

# Loop through each sheet to extract relevant metrics
for sheet in sheet_names:
    if sheet == 'Definitions':
        continue
    df = pd.read_excel(excel_path, sheet_name=sheet)
    industry_col = None
    for col in df.columns:
        if isinstance(col, str) and ('Industry' in col or 'Sector' in col):
            industry_col = col
            break
    if industry_col:
        df.rename(columns={industry_col: 'Industry'}, inplace=True)
        for _, row in df.iterrows():
            industry = row['Industry']
            if pd.isna(industry):
                continue
            if industry not in sector_data:
                sector_data[industry] = {}
            for col in df.columns:
                if col != 'Industry':
                    value = row[col]
                    if pd.notna(value):
                        sector_data[industry][f"{sheet}_{col}"] = value

# Convert the nested dictionary to a DataFrame for easier manipulation
sector_df = pd.DataFrame.from_dict(sector_data, orient='index')

# Clean and convert the data to numeric types
cleaned_sector_df = convert_to_numeric(sector_df.copy())

# Initialize MinMaxScaler and normalize the data
scaler = MinMaxScaler()
scaled_data = scaler.fit_transform(cleaned_sector_df.fillna(0))
scaled_df = pd.DataFrame(scaled_data, columns=cleaned_sector_df.columns, index=cleaned_sector_df.index)

weights_df_complete = pd.read_csv('Venture-Capital-Weights.csv')

weights_dict = weights_df_complete.set_index('Metric')['Weight'].to_dict()

scaled_df = scaled_df.drop(['Total Market', 'Total Market (without financials)'], errors='ignore')

# Initialize the composite score as zeros
composite_score_updated = pd.Series([0]*scaled_df.shape[0], index=scaled_df.index)

# Update the composite score calculation using individual weights
for metric in scaled_df.columns:
    if metric in weights_dict:
        composite_score_updated += scaled_df[metric] * weights_dict[metric]

# Sort the sectors based on the composite score
sorted_sectors = composite_score_updated.sort_values(ascending=False)

print("Top 10 Sectors for VC Investments based on Composite Score:")
print(sorted_sectors.head(10))