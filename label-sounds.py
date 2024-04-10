import pandas as pd

# Define the mapping of terms to sounds
term_to_sound = {
    'Butter': 'Non retroflex',
    'Valet': 'Non retroflex',
    'Puny': 'Non retroflex',
    'Tiny': 'Non retroflex',
    'Gala': 'Non retroflex',
    'Kannaa': 'n',
    'Pattam : Kite': 't',
    'Vellai': 'l',
    'Kullam : pond': 'l',
    'Padattum': 't',
    'Pallam : pit': 'l',
    'Pillai : Son': 'l',
    'Kannaadi : mirror': 'n',
    'Ketta : Bad': 't',
    'Vettu : cut / chop': 't',
    'Kollai': 'l',
    'Aattam : Dance': 't',
    'Vattum : circle': 't',
    'Mattum': 't',
    'Sattam': 't',
    'Vellam : Flood': 'l',
    'Vannam : color': 'n',
    'Kulaai : Pipe': 'l',
    'Ullai : inside': 'l',
    'Kattam : Crossword': 't',
    'Ennai : oil': 'n',
    'Tunnivu : courage': 'n',
    'Anna : Elder Brother': 'n',
    'Kannam : cheek': 'n',
    'Thanni : water': 'n',
    'Katti': 't'
}

def determine_sound(term):
    # Strip leading and trailing whitespace from the term
    term = term.strip()
    # Check if the term exists in the mapping
    if term in term_to_sound:
        return term_to_sound[term]
    else:
        raise ValueError(f"Term '{term}' not found in the mapping dictionary.")


# Load the Excel file
excel_file = "combined_data.xlsx"
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Iterate through all sheets
for sheet_name, sheet_data in all_sheets.items():
    # Add a new column "Sound" based on the "Term" column using determine_sound function
    sheet_data["Sound"] = sheet_data["Term"].apply(determine_sound)
    
# Write the modified DataFrames to the new Excel file
output_excel_file = "modified_combined_data.xlsx"
with pd.ExcelWriter(output_excel_file) as writer:
    for sheet_name, sheet_data in all_sheets.items():
        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)