import os
import pandas as pd

def update_prices(master_path, new_path, output_path):
    # Read the existing master prices Excel file
    master_df = pd.read_excel(master_path)

    # Read the new CSV with the latest cards
    new_df = pd.read_csv(new_path)

    # Identify new cards by comparing against the master list using the first 7 columns
    columns_to_compare = ['set', 'name', 'number', 'unlimited', 'reverse', 'promo']
    new_cards_df = new_df[~new_df[columns_to_compare].isin(master_df[columns_to_compare]).all(1)]

    # Only select the first 7 columns for the new cards
    new_cards_df = new_cards_df[columns_to_compare]

    # Append the new cards to the master list
    updated_master_df = pd.concat([master_df, new_cards_df], ignore_index=True)

    # Save the updated master list to a new Excel file
    updated_master_df.to_excel(output_path, index=False)

if __name__ == "__main__":
    master_prices_path = 'prices/prices.xlsx'
    new_cards_path = 'my_cards.csv'
    output_path = 'prices/prices.xlsx'

    update_prices(master_prices_path, new_cards_path, output_path)
