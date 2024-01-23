import os
import pandas as pd
from datetime import datetime

def select_and_write_doubles(df, column, output_folder, output_filename):
    doubles_df = df[df[column].astype(int) > 1][['set', 'name', 'number']]
    output_path = os.path.join(output_folder, output_filename)
    doubles_df.to_excel(output_path, index=False)

# Load CSV data into a DataFrame and select only the first 6 columns
selected_columns = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']
df = pd.read_csv('my_cards.csv', usecols=selected_columns)

# Create a folder with the current date (YYYY-MM-DD format)
current_date = datetime.now().strftime('%Y-%m-%d')
output_folder = os.path.join('doubles', current_date)

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Select and write doubles for 'unlimited' cards
select_and_write_doubles(df, 'unlimited', output_folder, 'unlimited_doubles.xlsx')

# Select and write doubles for 'reverse' cards
select_and_write_doubles(df, 'reverse', output_folder, 'reverse_doubles.xlsx')

# Select and write doubles for 'promo' cards
select_and_write_doubles(df, 'promo', output_folder, 'promo_doubles.xlsx')
