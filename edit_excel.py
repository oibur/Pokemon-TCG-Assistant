import pandas as pd

# Load the workbook
file_path = 'CARDS.xlsx'  # replace with your actual file path
new_file_path = 'EDIT.xlsx'  # new file path to save the results

# Load the relevant sheets
workbook = pd.ExcelFile(file_path)
cards_df = workbook.parse('Cards')
set_sheets = workbook.sheet_names[2:]  # Skip the first two sheets

# Ensure 'Set / #' in 'Cards' is formatted as a string
cards_df['Set / #'] = cards_df['Set / #'].astype(str)

# Identify duplicates
duplicate_rows = cards_df[cards_df.duplicated(['Set / #'], keep=False)]
if not duplicate_rows.empty:
    print("Duplicate entries found in 'Set / #':")
    print(duplicate_rows)

    # Remove duplicates by keeping the first occurrence
    cards_df = cards_df.drop_duplicates(['Set / #'], keep='first')

# Ensure 'Set / #' is unique in cards_df
if not cards_df['Set / #'].is_unique:
    raise ValueError("The 'Set / #' column in 'Cards' is not unique after handling duplicates.")

# Process each set sheet
modified_dfs = {}

for sheet_name in set_sheets:
    set_df = workbook.parse(sheet_name)
    
    # Create the new 'Set / #' column by concatenating 'set-name' and 'number'
    set_df['set-name'] = set_df['set-name'].astype(str)
    set_df['number'] = set_df['number'].astype(str)
    set_df.insert(0, 'Set / #', set_df['set-name'] + ' ' + set_df['number'])
    
    # Loop through the columns F and onwards in the set sheet
    for col in set_df.columns[5:]:
        if col in cards_df.columns:
            # Create the quantity column
            qnty_col = f"{col}_Qty"
            set_df[qnty_col] = set_df['Set / #'].map(cards_df.set_index('Set / #')[col]).fillna(0).astype(int)
            
            # Create the total value column
            total_col = f"{col}_total_$"
            set_df[total_col] = set_df[col] * set_df[qnty_col]
    
    # Store the modified dataframe in the dictionary
    modified_dfs[sheet_name] = set_df

# Save the modified dataframes back to a new Excel file
with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
    # Write the original sheets
    for sheet_name in workbook.sheet_names[:2]:
        workbook.parse(sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Write the modified set sheets
    for sheet_name, df in modified_dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Processed data has been saved to {new_file_path}")
