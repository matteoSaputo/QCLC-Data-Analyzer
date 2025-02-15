import pandas as pd
import glob

# Path to the folder containing the roster files
folder_path = 'FA24 Grade Rosters/' # Adjust to relevant semester
file_pattern = folder_path + '*.xlsx'  # Adjust extension if CSV or other format

# Get a list of all roster files in the folder
roster_files = glob.glob(file_pattern)

all_rosters = []
for file in roster_files:
    # Read each roster file (header=1 to skip the first row if needed)
    df = pd.read_excel(file, header=1)
    all_rosters.append(df)

# Concatenate all DataFrames
combined_roster = pd.concat(all_rosters, ignore_index=True)

# Standardize column names
combined_roster.columns = combined_roster.columns.str.lower().str.strip()

# Save the combined DataFrame to a new file
combined_roster.to_excel('combined_Roster.xlsx', index=False)  # Save as Excel

print("All rosters have been combined and saved!")
