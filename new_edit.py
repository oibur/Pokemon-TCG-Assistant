import pandas as pd
import xlsxwriter

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

# Function to convert column index to Excel column letters
def col_num_to_col_letters(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

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
    
    # Add a new column that sums all columns ending in '_total_$'
    total_value_columns = [col for col in set_df.columns if col.endswith('_total_$')]
    set_df['Total_Value_$'] = set_df[total_value_columns].sum(axis=1)
    
    # Move the 'Total_Value_$' column to the beginning
    cols = ['Total_Value_$'] + [col for col in set_df.columns if col != 'Total_Value_$']
    set_df = set_df[cols]
    
    # Store the modified dataframe in the dictionary
    modified_dfs[sheet_name] = set_df

# Save the modified dataframes back to a new Excel file
with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
    # Write the original sheets
    for sheet_name in workbook.sheet_names[:2]:
        workbook.parse(sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Write the modified set sheets
    for sheet_name, df in modified_dfs.items():
        # Sort by column C (set-releaseDate) and then by column E (number)
        df = df.sort_values(by=['set-releaseDate', 'number'])
        
        # Write to Excel
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Get the workbook and worksheet objects
        workbook_obj = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Get the header format
        header_format = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
        
        # Get the green format for total value columns
        green_format = workbook_obj.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
        
        # Write the header with the specified format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
            # Highlight columns ending in $ green and 'Total_Value_$'
            if value.endswith('_total_$') or value == 'Total_Value_$':
                worksheet.set_column(col_num, col_num, None, green_format)

print(f"Processed data has been saved to {new_file_path}")
