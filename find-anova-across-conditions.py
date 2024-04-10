import pandas as pd
from scipy.stats import f_oneway
import openpyxl


def perform_anova(data1, data2, column_name, output_file):
    ready_arr = [4, 5, 17, 18]
    ready_dict = {4 : "F3_Before", 5: "F4_Before", 17: "F3_After", 18: "F4_After"}
    if column_name in ready_arr:
        # Convert column data to numeric format
        data1_column = pd.to_numeric(data1[column_name], errors='coerce').dropna()
        data2_column = pd.to_numeric(data2[column_name], errors='coerce').dropna()
        
        with open(output_file, 'a') as f:
            f.write(f"Column: {ready_dict[column_name]}\n")
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
workbook = openpyxl.load_workbook('summarized-formant-data.xlsx')

participant_names = ['Aa', 'H', 'Kesh', 'Krish', 'Krith', 'Sh']
sheet_name1_arr = []
sheet_name2_arr = []
for name in participant_names:
    sheet_name1_arr.append(name + '_English 1')
    sheet_name1_arr.append(name + '_Tamil 1 60bpm')
    sheet_name1_arr.append(name + '_Tamil 1 102bpm')
    sheet_name2_arr.append(name + '_English 2')
    sheet_name2_arr.append(name + '_Tamil 2 60bpm')
    sheet_name2_arr.append(name + '_Tamil 2 102bpm')
sheet_name1_arr = ['Data - ' + item for item in sheet_name1_arr]
sheet_name2_arr = ['Data - ' + item for item in sheet_name2_arr]

for sheet1, sheet2 in zip(sheet_name1_arr, sheet_name2_arr):
    # Output file
    output_file = f'anova-output/anova_results_conditions_{sheet1}_{sheet2}.txt'

    # Clear contents of the output file
    open(output_file, 'w').close()

    # Provide the names of the two sheets
    sheet_name1 = sheet1
    sheet_name2 = sheet2

    # Get the corresponding sheets
    sheet1 = workbook[sheet_name1]
    sheet2 = workbook[sheet_name2]

    # Convert sheet data to DataFrames
    data1 = pd.DataFrame(sheet1.values)
    data2 = pd.DataFrame(sheet2.values)

    # Print debug information
    print("Data1:")
    print(data1.head())
    print("Data2:")
    print(data2.head())

    # Get common columns
    common_columns = set(data1.columns).intersection(data2.columns)

    # Perform ANOVA for each common column
    for column_name in common_columns:
        perform_anova(data1, data2, column_name, output_file)

    print("ANOVA results have been saved to 'anova_results_conditions.txt'.")


