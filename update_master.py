import os
import pandas as pd

def update_prices(master_path, new_path, output_path):
    # Read the existing master prices CSV
    master_df = pd.read_csv(master_path)

    # Read the new CSV with the latest cards
    new_df = pd.read_csv(new_path)

    # Identify new cards by comparing against the master list
    new_cards_df = new_df[~new_df.isin(master_df.to_dict('list')).all(1)]

    # Append the new cards to the master list
    updated_master_df = pd.concat([master_df, new_cards_df], ignore_index=True)

    # Save the updated master list to a new CSV file
    updated_master_df.to_csv(output_path, index=False)

if __name__ == "__main__":
    master_prices_path = 'prices/prices.csv'
    new_cards_path = 'my_cards.csv'
    output_path = 'prices/prices.csv'

    update_prices(master_prices_path, new_cards_path, output_path)
