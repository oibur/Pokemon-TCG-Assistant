import pandas as pd
import requests
from openpyxl import load_workbook

# Load Excel and find sheet
excel_file_path = 'CARDS.xlsx'
sheet_name = 'sets'
existing_df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Fetch set data from API
api_key = 'c2eaa76b-c34c-4d3a-8f33-da95a230d9ea'
url = 'https://api.pokemontcg.io/v2/sets'
headers = {'X-Api-Key': api_key}
response = requests.get(url, headers=headers)

# Compare API data to Excel data and update differences
if response.status_code == 200:
    api_data = response.json()["data"]
    api_ids = set(entry["id"] for entry in api_data)
    existing_ids = set(existing_df["id"])
    new_ids = api_ids - existing_ids

    for entry in api_data:
        if entry["id"] in new_ids:
            new_row = {
                "id": entry["id"],
                "series": entry["series"],
                "name": entry["name"],
                "total": entry["total"],
                "releaseDate": entry["releaseDate"]
            }
            existing_df = existing_df.append(new_row, ignore_index=True)

    # Sort by "releaseDate"
    existing_df.sort_values(by='releaseDate', inplace=True)

    # Save the Pandas Excel writer to the disk
    book = load_workbook(excel_file_path)
    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')
    existing_df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.book = book
    writer.save()

else:
    print(f"Error: Unable to fetch data from the API. Status Code: {response.status_code}")