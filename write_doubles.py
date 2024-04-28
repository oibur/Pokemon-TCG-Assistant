import os
import pandas as pd
from datetime import datetime

def select_and_write_doubles(df, column, excel_writer, sheet_name, columns_to_write):
    # Filter the DataFrame to find rows where the specified column value is greater than 1
    doubles_df = df[df[column].astype(int) > 1]
    # Select only the specified columns to write to the Excel sheet
    doubles_df = doubles_df[columns_to_write]
    # Write the filtered DataFrame with the specified columns to the specified Excel sheet
    doubles_df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Load CSV data into a DataFrame and select only the first 6 columns
selected_columns = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']
df = pd.read_csv('tcghub_collection.csv', usecols=selected_columns)

# Get the current date in YYYY-MM-DD format
current_date = datetime.now().strftime('%Y-%m-%d')

# Define the base folder as 'History'
base_folder = 'History'

# Define the output folder as a subfolder of the base folder for the current date
output_folder = os.path.join(base_folder, current_date)

# Define the output file path to save the workbook
output_file_path = os.path.join(output_folder, 'doubles.xlsx')

# Create an Excel writer object to write data to the same workbook
with pd.ExcelWriter(output_file_path) as excel_writer:
    # Select and write doubles for 'unlimited' cards with specified columns
    columns_unlimited = ['set', 'name', 'number', 'unlimited']
    select_and_write_doubles(df, 'unlimited', excel_writer, 'unlimited_doubles', columns_unlimited)
    
    # Select and write doubles for 'reverse' cards with specified columns
    columns_reverse = ['set', 'name', 'number', 'reverse']
    select_and_write_doubles(df, 'reverse', excel_writer, 'reverse_doubles', columns_reverse)
    
    # Select and write doubles for 'promo' cards with specified columns
    columns_promo = ['set', 'name', 'number', 'promo']
    select_and_write_doubles(df, 'promo', excel_writer, 'promo_doubles', columns_promo)