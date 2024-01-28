import pandas as pd
import requests
from datetime import datetime

# Step 1: Read the existing Excel sheet
df = pd.read_csv('test.csv')

# Step 2: Eliminate the last 3 columns and rows where all columns are 0
df = df.iloc[:, :-3]  # Exclude the last 3 columns
df = df[(df.iloc[:, 3:] != 0).any(axis=1)]  # Exclude rows where all columns are 0

# Step 3: Fetch prices using the TCGplayer API
def get_card_prices(card_name):
    # Replace 'your_api_key' with your actual TCGplayer API key
    api_key = 'c2eaa76b-c34c-4d3a-8f33-da95a230d9ea'
    url = f'https://api.pokemontcg.io/v2/cards?q=name:{card_name}'
    headers = {'X-Api-Key': api_key}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        if data['data']:
            prices = data['data'][0]['tcgplayer']['prices']
            return prices.get('normal', {}).get('market'), prices.get('reverseHolofoil', {}).get('market')
    return None, None

# Apply the function to each row in the DataFrame to fetch prices
df['unlimited_price'], df['reverse_price'] = zip(*df['name'].apply(get_card_prices))

# Step 4: Calculate the total price for cards with multiple formats
df['total_price'] = df['unlimited'] * df['unlimited_price'] + df['reverse'] * df['reverse_price']

# Step 5: Create a new Excel sheet
output_folder = 'prices'
current_date = datetime.now().strftime('%Y-%m-%d')
output_path = f'{output_folder}/master_{current_date}.xlsx'

df.to_excel(output_path, index=False)