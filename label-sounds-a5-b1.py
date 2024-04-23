import pandas as pd

# Define the mapping of terms to sounds
term_to_sound = {
    'Butter': 't',
    'Valet': 'l',
    'Puny': 'n',
    'Tiny': 'n',
    'Gala': 't',
    'Kannaa': 'ɳ',
    'Pattam : Kite': 'ʈ',
    'Vellai': 'ɭ',
    'Kullam : pond': 'ɭ',
    'Padattum': 'ʈ',
    'Pallam : pit': 'ɭ',
    'Pillai : Son': 'ɭ',
    'Kannaadi : mirror': 'ɳ',
    'Ketta : Bad': 'ʈ',
    'Vettu : cut / chop': 'ʈ',
    'Kollai': 'ɭ',
    'Aattam : Dance': 'ʈ',
    'Vattum : circle': 'ʈ',
    'Mattum': 'ʈ',
    'Sattam': 'ʈ',
    'Vellam : Flood': 'ɭ',
    'Vannam : color': 'ɳ',
    'Kulaai : Pipe': 'ɭ',
    'Ullai : inside': 'ɭ',
    'Kattam : Crossword': 'ʈ',
    'Ennai : oil': 'ɳ',
    'Tunnivu : courage': 'ɳ',
    'Anna : Elder Brother': 'ɳ',
    'Kannam : cheek': 'ɳ',
    'Thanni : water': 'ɳ',
    'Katti': 'ʈ'
}

term_to_preconsonantal_vowel = {
    'Butter': 'ə',               
    'Valet': 'æ',
    'Puny': 'u',
    'Tiny': 'a',
    'Gala': 'æ',
    'Kannaa': 'ə',               
    'Pattam : Kite': 'ə',      
    'Vellai': 'ɛ',
    'Kullam : pond': 'u',
    'Padattum': 'ə',            
    'Pallam : pit': 'ə',       
    'Pillai : Son': 'ɪ',
    'Kannaadi : mirror': 'ə', 
    'Ketta : Bad': 'ɛ',
    'Vettu : cut / chop': 'ɛ',
    'Kollai': 'o',
    'Aattam : Dance': 'a',
    'Vattum : circle': 'ə',     
    'Mattum': 'ə',               
    'Sattam': 'ə',               
    'Vellam : Flood': 'ɛ',
    'Vannam : color': 'ə',     
    'Kulaai : Pipe': 'u',
    'Ullai : inside': 'u',
    'Kattam : Crossword': 'ə',   
    'Ennai : oil': 'ɛ',
    'Tunnivu : courage': 'u',
    'Anna : Elder Brother': 'ə', 
    'Kannam : cheek': 'ə',       
    'Thanni : water': 'ə',       
    'Katti': 'ə'                
}


# Define the mapping of terms to sounds
term_to_postconsonantal_vowel = {
    'Butter': 'ə',
    'Valet': 'e',
    'Puny': 'i',
    'Tiny': 'i',
    'Gala': 'ə',
    'Kannaa': 'a',
    'Pattam : Kite': 'ə',
    'Vellai': 'ɛ',
    'Kullam : pond': 'ə',
    'Padattum': 'ə',
    'Pallam : pit': 'ə',
    'Pillai : Son': 'ɛ',
    'Kannaadi : mirror': 'a',
    'Ketta : Bad': 'ɛ',
    'Vettu : cut / chop': 'u',
    'Kollai': 'ɛ',
    'Aattam : Dance': 'ə',
    'Vattum : circle': 'ə',
    'Mattum': 'ə',
    'Sattam': 'ə',
    'Vellam : Flood': 'ə',
    'Vannam : color': 'ə',
    'Kulaai : Pipe': 'ɛ',
    'Ullai : inside': 'ɛ',
    'Kattam : Crossword': 'ə',
    'Ennai : oil': 'ɛ',
    'Tunnivu : courage': 'i',
    'Anna : Elder Brother': 'a',
    'Kannam : cheek': 'ə',
    'Thanni : water': 'i',
    'Katti': 'i'
}

def determine_sound(term):
    # Strip leading and trailing whitespace from the term
    term = term.strip()
    # Check if the term exists in the mapping
    if term in term_to_sound:
        return term_to_sound[term]
    else:
        raise ValueError(f"Intervocalic Consonant '{term}' not found in the mapping dictionary.")
    
def determine_pre(term):
    # Strip leading and trailing whitespace from the term
    term = term.strip()
    # Check if the term exists in the mapping
    if term in term_to_preconsonantal_vowel:
        return term_to_preconsonantal_vowel[term]
    else:
        raise ValueError(f"Intervocalic Consonant '{term}' not found in the mapping dictionary.")
    
def determine_post(term):
    # Strip leading and trailing whitespace from the term
    term = term.strip()
    # Check if the term exists in the mapping
    if term in term_to_postconsonantal_vowel:
        return term_to_postconsonantal_vowel[term]
    else:
        raise ValueError(f"Intervocalic Consonant '{term}' not found in the mapping dictionary.")

# Load the Excel file
excel_file = "summarized-formant-data.xlsx"
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Iterate through all sheets
for sheet_name, sheet_data in all_sheets.items():
    try:
        # Add a new column "Sound" based on the "Term" column using determine_sound function
        sheet_data["Intervocalic Consonant"] = sheet_data["Term"].apply(determine_sound)
        sheet_data["Vowel Before Consonant"] = sheet_data["Term"].apply(determine_pre)
        sheet_data["Vowel After Consonant"] = sheet_data["Term"].apply(determine_post)
    except KeyError:
        # If the "Term" column does not exist in the current sheet, handle the KeyError here
        # You can choose to skip this sheet or perform any other desired action
        print(f"KeyError: 'Term' column not found in sheet '{sheet_name}'. Skipping...")
        continue
    
    # Remove columns that start with "Unnamed"
    sheet_data = sheet_data.loc[:, ~sheet_data.columns.str.startswith('Unnamed')]
    all_sheets[sheet_name] = sheet_data

# Remove the first two sheets
sheets_to_delete = list(all_sheets.keys())[:2]
for sheet_name in sheets_to_delete:
    all_sheets.pop(sheet_name)

# Write the modified DataFrames to the new Excel file
output_excel_file = "modified_combined_data.xlsx"
with pd.ExcelWriter(output_excel_file) as writer:
    for sheet_name, sheet_data in all_sheets.items():
        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
