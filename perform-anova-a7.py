import pandas as pd
from scipy.stats import f_oneway
import openpyxl
import re

# Define the ANOVA function
def perform_anova(data1, data2, column_name, output_file):
    # Convert column data to numeric format
    data1_column = pd.to_numeric(data1[column_name], errors='coerce').dropna()
    data2_column = pd.to_numeric(data2[column_name], errors='coerce').dropna()
    
    with open(output_file, 'a') as f:
        f.write(f"Column: {column_name}\n")
        f_statistic, p_value = f_oneway(data1_column, data2_column)
        f_statistic_str = str(f_statistic)
        f.write(f"F-statistic: {f_statistic_str}\n")
        f.write(f"P-value: {p_value}\n")
        f.write("Useful result?: ")
        if p_value < 0.05:
            f.write("Yes\n")
            f.write("Reject null hypothesis: There is a significant difference between groups.\n")
        else:
            f.write("No\n")
            f.write("Fail to reject null hypothesis: There is no significant difference between groups.\n")
        f.write("\n")

# Load the Excel workbook
workbook = openpyxl.load_workbook('modified_combined_data.xlsx')

# Read the Excel file and identify pairs of sheets
sheet_names = workbook.sheetnames
pairs = []
pattern = re.compile(r'Data - (.+)_1')

for sheet_name in sheet_names:
    match = pattern.match(sheet_name)
    if match:
        participant_name = match.group(1)
        sheet1 = pd.read_excel('consolidated_data.xlsx', 0)
        sheet2 = pd.read_excel('consolidated_data.xlsx', 1)
        pairs.append((participant_name, sheet1, sheet2))

# Loop through each pair of sheets
for name, sheet1, sheet2 in pairs:
    # Combine data from before and after priming into a single DataFrame
    combined_data = pd.concat([sheet1, sheet2], ignore_index=True)
    
    # Output file for ANOVA
    anova_output_file = f'anova-output/anova_results_conditions_{name}.txt'
    open(anova_output_file, 'w').close()  # Clear contents of the output file

    # Specify columns for ANOVA
    columns_for_anova = ['F3_Hz_Before', 'F3_Hz_After', 'F4_Hz_Before', 'F4_Hz_After']

    # Perform ANOVA for each specified column
    for column_name in columns_for_anova:
        perform_anova(sheet1, sheet2, column_name, anova_output_file)

    print(f"ANOVA results for pair '{name}' have been saved to '{anova_output_file}'.")

import subprocess

# Specify the path to the Python file you want to run
python_file_path = 'perform-stats.py'

# Run the Python file
print('running: ', python_file_path)
subprocess.run(['python', python_file_path])