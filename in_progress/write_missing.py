import os
import pandas as pd
from datetime import datetime

def select_and_write_missing(df, column, output_folder, output_filename):
    missing_df = df[df[column].astype(int) == 0][['set', 'name', 'number']]
    output_path = os.path.join(output_folder, output_filename)
    missing_df.to_excel(output_path, index=False)

# Load CSV data into a DataFrame and select only the first 6 columns
selected_columns = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']
df = pd.read_csv('tcghub_collection.csv', usecols=selected_columns)

# Create a folder with the current date (YYYY-MM-DD format)
current_date = datetime.now().strftime('%Y-%m-%d')
output_folder = os.path.join('missing', current_date)

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Select and write missing cards for 'promo' where the 'promo' column is 0
select_and_write_missing(df[df['set'].str.contains('promo', case=False)], 'promo', output_folder, 'promo_missing.xlsx')

# Remove rows with 'promo' in the 'set' column
df = df[~df['set'].str.contains('promo', case=False)]

# Select and write missing cards for 'unlimited'
select_and_write_missing(df, 'unlimited', output_folder, 'unlimited_missing.xlsx')

# Select and write missing cards for 'reverse'
select_and_write_missing(df, 'reverse', output_folder, 'reverse_missing.xlsx')