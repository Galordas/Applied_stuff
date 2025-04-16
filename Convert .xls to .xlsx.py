import pandas as pd
from pathlib import Path

# Define source and target directories
source_dir = Path(r'')
target_dir = Path(r'')

# Ensure the target directory exists
target_dir.mkdir(parents=True, exist_ok=True)

# Iterate over all .xls files in the source directory
for file_path in source_dir.glob('*.xls'):
    # Read all sheets from the Excel file
    sheets_dict = pd.read_excel(file_path, sheet_name=None)

    # Create an empty list to store DataFrames
    dfs = []

    # Iterate over each sheet
    for sheet_name, df in sheets_dict.items():
        # Drop the first 3 rows
        df = df.drop(range(3))
        # Append the modified DataFrame to the list
        dfs.append(df)

    # Concatenate all DataFrames into a single DataFrame
    merged_df = pd.concat(dfs, ignore_index=True)

    # Define the new file path with .xlsx extension
    new_file_path = target_dir / file_path.name.replace('.xls', '.xlsx')

    # Save the merged DataFrame to a new Excel file
    merged_df.to_excel(new_file_path, index=False)

print("Processing complete.")