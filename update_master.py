import os
import pandas as pd

def update_prices(master_path, new_path, output_path):
    # Read the existing master prices Excel file
    master_df = pd.read_excel(master_path)

    # Read the new CSV with the latest cards
    new_df = pd.read_csv(new_path)

    # Merge the DataFrames based on common keys ('set' and 'number')
    merged_df = pd.merge(master_df, new_df, on=['set', 'number'], how='outer', suffixes=('', '_new'))

    # Columns to update
    update_columns = ['unlimited', 'reverse', 'promo', '1st']

    # Update existing rows where the values are different
    for col in update_columns:
        col_master = col
        col_new = col + '_new'
        master_df[col_master] = merged_df.apply(lambda row: row[col_new] if pd.notnull(row[col_new]) else row[col_master], axis=1)

    # Sort the final DataFrame by 'set' and then 'number'
    master_df = master_df.sort_values(by=['set', 'number'])

    # Save the updated master list to a new Excel file
    master_df.to_excel(output_path, index=False)

if __name__ == "__main__":
    master_prices_path = 'prices/prices.xlsx'
    new_cards_path = 'my_cards.csv'
    prices_output_path = 'prices/prices.xlsx'

    update_prices(master_prices_path, new_cards_path, prices_output_path)
