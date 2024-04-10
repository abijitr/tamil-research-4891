import pandas as pd

# Load the Excel file
excel_file = "modified_combined_data.xlsx"
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Initialize an Excel writer object
with pd.ExcelWriter("consolidated_data.xlsx") as writer:
    # Initialize empty dataframes for odd and even sheets
    odd_combined = pd.DataFrame()
    even_combined = pd.DataFrame()

    # Iterate over all sheets
    for i, (sheet_name, df) in enumerate(all_sheets.items()):
        # Check if the sheet is odd or even
        if i % 2 == 0:
            # Even sheet
            even_combined = pd.concat([even_combined, df], ignore_index=True)
        else:
            # Odd sheet
            odd_combined = pd.concat([odd_combined, df], ignore_index=True)

    # Write the combined dataframes to the same Excel file
    odd_combined.to_excel(writer, sheet_name="Before_Priming", index=False)
    even_combined.to_excel(writer, sheet_name="After_Priming", index=False)

import subprocess

# Specify the path to the Python file you want to run
python_file_path = 'perform-anova.py'

# Run the Python file
print('running: ', python_file_path)
subprocess.run(['python', python_file_path])