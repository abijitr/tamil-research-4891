import pandas as pd
from scipy.stats import f_oneway
import openpyxl

def perform_anova(data, worksheet_name, output_file):
    before_columns = ['F3_Hz_Before', 'F4_Hz_Before']
    after_columns = ['F3_Hz_After', 'F4_Hz_After']
    
    with open(output_file, 'a') as f:
        for i, (before_col, after_col) in enumerate(zip(before_columns, after_columns), 3):
            f.write(f"Worksheet: {worksheet_name}\n")
            f.write(f"ANOVA for Formant F{i}\n")
            before_data = pd.to_numeric(data[before_col], errors='coerce').dropna()
            after_data = pd.to_numeric(data[after_col], errors='coerce').dropna()
            f_statistic, p_value = f_oneway(before_data, after_data)
            f_statistic_str = str(f_statistic)
            f.write(f"F-statistic: {f_statistic_str}\n")
            f.write(f"P-value: {p_value}\n")
            if p_value < 0.05:
                f.write("Reject null hypothesis: There is a significant difference between groups.\n")
            else:
                f.write("Fail to reject null hypothesis: There is no significant difference between groups.\n")
            f.write("\n")

# Load the Excel workbook
workbook = openpyxl.load_workbook('summarized-formant-data.xlsx')

# Output file
output_file = 'anova_results.txt'

# Clear contents of the output file
open(output_file, 'w').close()

# Iterate over each sheet in the workbook
for sheet_name in workbook.sheetnames:
    print(f"Processing sheet: {sheet_name}")
    sheet = workbook[sheet_name]
    
    # Convert sheet data to DataFrame
    data = pd.DataFrame(sheet.values, columns=[cell.value for cell in sheet[1]])

    # Perform ANOVA and write results to output file
    perform_anova(data, sheet_name, output_file)

print("ANOVA results have been saved to 'anova_results.txt'.")
