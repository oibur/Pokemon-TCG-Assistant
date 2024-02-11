import pandas as pd
import requests
from openpyxl import Workbook

# Fetch set data from API
api_key = 'c2eaa76b-c34c-4d3a-8f33-da95a230d9ea'
url = 'https://api.pokemontcg.io/v2/sets'
headers = {'X-Api-Key': api_key}
response = requests.get(url, headers=headers)

# Create a new Excel workbook
excel_file_path = 'CARDS.xlsx'
wb = Workbook()

# Create a new sheet named "sets"
ws = wb.active
ws.title = 'sets'

# Write set data to the sheet with column names
if response.status_code == 200:
    api_data = response.json()["data"]
    
    # Write column names
    column_names = ["id", "series", "name", "total", "releaseDate"]
    ws.append(column_names)
    
    # Sort the data by "releaseDate"
    api_data.sort(key=lambda x: x["releaseDate"])
    
    # Write set data
    for entry in api_data:
        row = [entry["id"], entry["series"], entry["name"], entry["total"], entry["releaseDate"]]
        ws.append(row)

    # Save the workbook
    wb.save(filename=excel_file_path)
    print(f"Excel file '{excel_file_path}' has been created successfully.")
else:
    print(f"Error: Unable to fetch data from the API. Status Code: {response.status_code}")
