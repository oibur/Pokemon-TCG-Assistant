import os
import pandas as pd
from datetime import datetime

def filter_and_write(df, column, writer, sheet_name, columns):
    # Filter and write the specified columns to the Excel sheet
    df[df[column].astype(int) > 1][columns].to_excel(writer, sheet_name=sheet_name, index=False)

current_date = datetime.now().strftime('%Y-%m-%d')
base_folder = 'History'
output_file_path = os.path.join(base_folder, current_date, 'doubles.xlsx')
input_file_path = os.path.join(base_folder, current_date, 'tcghub_collection.csv')

columns_to_use = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']

# Load CSV data into a DataFrame and open the Excel writer
df = pd.read_csv(input_file_path, usecols=columns_to_use)
with pd.ExcelWriter(output_file_path) as writer:
    # Filter and write doubles for each column
    filter_and_write(df, 'unlimited', writer, 'unlimited_doubles', columns_to_use[:4])
    filter_and_write(df, 'reverse', writer, 'reverse_doubles', columns_to_use[:3] + [columns_to_use[4]])
    filter_and_write(df, 'promo', writer, 'promo_doubles', columns_to_use[:3] + [columns_to_use[5]])
