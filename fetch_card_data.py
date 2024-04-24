import pandas as pd
import requests
from openpyxl import Workbook
from typing import List, Dict, Any

# Constants
API_KEY = 'c2eaa76b-c34c-4d3a-8f33-da95a230d9ea'
API_URL = 'https://api.pokemontcg.io/v2'
EXCEL_FILE_PATH = 'CARDS.xlsx'
SET_SHEET_NAME = 'sets'
CARDS_SHEET_NAME = 'cards'
COLUMN_NAMES = ["id", "series", "name", "total", "releaseDate"]

def fetch_data(url: str, headers: Dict[str, str]) -> Dict[str, Any]:
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()["data"]

def fetch_set_data() -> List[Dict[str, Any]]:
    url = f'{API_URL}/sets'
    headers = {'X-Api-Key': API_KEY}
    return fetch_data(url, headers)

def fetch_card_data(set_id: str) -> List[Dict[str, Any]]:
    url = f'{API_URL}/cards?q=set.id:{set_id}'
    headers = {'X-Api-Key': API_KEY}
    return fetch_data(url, headers)

def write_set_data_to_excel(data: List[Dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SET_SHEET_NAME
    ws.append(COLUMN_NAMES)
    sorted_data = sorted(data, key=lambda x: x["releaseDate"])
    for entry in sorted_data:
        row = [entry[col] for col in COLUMN_NAMES]
        ws.append(row)
    wb.save(filename=EXCEL_FILE_PATH)
    print(f"Excel file '{EXCEL_FILE_PATH}' has been created successfully.")

def extract_card_data(card_data: List[Dict[str, Any]], set_name: str) -> List[Dict[str, Any]]:
    extracted_data = []
    for card in card_data:
        card_id, card_name = card["id"], card["name"]
        rarity = card.get("rarity", None)
        prices_data = {f"{format_name}_market_price": format_data.get("market", None)
                       for format_name, format_data in card.get("tcgplayer", {}).get("prices", {}).items()}
        prices_data["reverseHolofoil_market_price"] = prices_data.get("reverseHolofoil_market_price", None)
        card_number = int(''.join(filter(str.isdigit, card_id)))
        extracted_data.append({"set": set_name, "number": card_number, "name": card_name, "rarity": rarity, **prices_data})
    return extracted_data

def main() -> None:
    try:
        set_data = fetch_set_data()
        write_set_data_to_excel(set_data)
    except requests.exceptions.RequestException as e:
        print(f"Error: Unable to fetch data from the API. {e}")
        return

    existing_df_sets = pd.DataFrame()
    try:
        existing_df_sets = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SET_SHEET_NAME, engine='openpyxl')
        if 'id' not in existing_df_sets.columns:
            raise ValueError("'id' column not found in existing_df_sets.")
    except FileNotFoundError:
        pass

    existing_df_cards = pd.DataFrame(columns=["set", "number", "name", "rarity", "holofoil_market_price"])
    for set_id in existing_df_sets['id'].tolist():
        try:
            card_data = fetch_card_data(set_id=set_id)
            extracted_data = extract_card_data(card_data, set_name=set_id)
            df = pd.DataFrame(extracted_data)
            existing_df_cards = pd.concat([existing_df_cards, df], ignore_index=True)
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data for set {set_id}: {e}")
            continue

    existing_df_cards.sort_values(by=['set', 'number'], inplace=True)

    with pd.ExcelWriter(EXCEL_FILE_PATH, engine='openpyxl', date_format='m/d/yyyy') as writer:
        existing_df_cards.to_excel(writer, sheet_name=CARDS_SHEET_NAME, index=False)
        existing_df_sets.to_excel(writer, sheet_name=SET_SHEET_NAME, index=False)

    print("Data written to Excel.")

if __name__ == "__main__":
    main()
