import pandas as pd

# Load the Excel file
excel_file = "summarized-formant-data.xlsx"
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Initialize an empty dictionary to store combined data
combined_data = {}

# Iterate through all sheets
for sheet_name, sheet_data in all_sheets.items():
    try:
        # Extract participant name and number from the sheet name
        participant_name, type_of_experiment = sheet_name.split("_", 1)
        number = type_of_experiment.split()[1]

        # Remove optional bpm information
        number = number.split("[")[0]

        # Create a key to identify the participant and number combination
        key = (participant_name, number)

        # If key exists in combined_data, append the current sheet data to it
        if key in combined_data:
            combined_data[key] = pd.concat([combined_data[key], sheet_data], ignore_index=True)
        # Otherwise, initialize the combined_data with the current sheet data
        else:
            combined_data[key] = sheet_data
    except ValueError:
        # Ignore sheets that don't have an underscore in their names
        pass

# Save the combined data to a new Excel file
with pd.ExcelWriter("combined_data.xlsx") as writer:
    for key, value in combined_data.items():
        participant_name, number = key
        sheet_name = f"{participant_name}_{number}"
        value.to_excel(writer, sheet_name=sheet_name, index=False)

import subprocess

# Specify the path to the Python file you want to run
python_file_path = 'label-sounds.py'

# Run the Python file
print('running: ', python_file_path)
subprocess.run(['python', python_file_path])