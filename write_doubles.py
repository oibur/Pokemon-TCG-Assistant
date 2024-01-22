import pandas as pd

def select_and_write_doubles(df, column, output_filename):
    doubles_df = df[df[column].astype(int) > 1][['set', 'name', 'number']]
    doubles_df.to_excel(output_filename, index=False)

# Load CSV data into a DataFrame and select only the first 6 columns
selected_columns = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']
df = pd.read_csv('my_cards.csv', usecols=selected_columns)

# Select and write doubles for 'unlimited' cards
select_and_write_doubles(df, 'unlimited', 'unlimited_doubles.xlsx')

# Select and write doubles for 'reverse' cards
select_and_write_doubles(df, 'reverse', 'reverse_doubles.xlsx')

# Select and write doubles for 'promo' cards
select_and_write_doubles(df, 'promo', 'promo_doubles.xlsx')
