import pandas as pd
import requests

def fetch_card_data(api_key, set_id):
    url = f'https://api.pokemontcg.io/v2/cards?q=set.id:{set_id}'
    headers = {'X-Api-Key': api_key}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()["data"]
    else:
        print(f"Error: Unable to fetch data for set {set_id}. Status Code: {response.status_code}")
        return []

def extract_data(card_data, set_name):
    extracted_data = []

    for card in card_data:
        card_id, card_name = card["id"], card["name"]
        rarity = card.get("rarity", None)

        prices_data = {f"{format_name}_market_price": format_data.get("market", None)
                       for format_name, format_data in card.get("tcgplayer", {}).get("prices", {}).items()}

        # Add the 'reverseHolofoil_market_price' column
        prices_data["reverseHolofoil_market_price"] = prices_data.get("reverseHolofoil_market_price", None)

        # Extract numeric part of the card number and convert to integer
        card_number = int(''.join(filter(str.isdigit, card_id)))

        extracted_data.append({"set": set_name, "number": card_number, "name": card_name, "rarity": rarity, **prices_data})

    return extracted_data

def main():
    excel_file_path, set_sheet_name, cards_sheet_name = 'CARDS.xlsx', 'sets', 'cards'

    try:
        existing_df_sets = pd.read_excel(excel_file_path, sheet_name=set_sheet_name, engine='openpyxl')
    except FileNotFoundError:
        existing_df_sets = pd.DataFrame()

    # Check if 'id' column exists in existing_df_sets
    if 'id' not in existing_df_sets.columns:
        print("Error: 'id' column not found in existing_df_sets.")
        return

    # Fetch unique set IDs
    set_ids = existing_df_sets['id'].tolist()

    existing_df_cards = pd.DataFrame(columns=["set", "number", "name", "rarity", "holofoil_market_price"])

    for set_id in set_ids:
        try:
            card_data = fetch_card_data(api_key='c2eaa76b-c34c-4d3a-8f33-da95a230d9ea', set_id=set_id)
        except Exception as e:
            print(f"Error fetching data for set {set_id}: {e}")
            continue

        if card_data:
            extracted_data = extract_data(card_data, set_name=set_id)
            df = pd.DataFrame(extracted_data)

            existing_df_cards = pd.concat([existing_df_cards, df], ignore_index=True)

            print(existing_df_cards)

    existing_df_cards.sort_values(by=['set', 'number'], inplace=True)

    with pd.ExcelWriter(excel_file_path, engine='openpyxl', date_format='m/d/yyyy') as writer:
        existing_df_cards.to_excel(writer, sheet_name=cards_sheet_name, index=False)
        existing_df_sets.to_excel(writer, sheet_name=set_sheet_name, index=True)

    print("Data written to Excel.")

if __name__ == "__main__":
    main()