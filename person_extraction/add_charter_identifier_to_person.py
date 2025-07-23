import pandas as pd
import numpy as np

def normalize_string(s):
    """Normalize string for comparison"""
    if pd.isna(s) or s == '':
        return ''
    return str(s).strip().lower()

def main():
    # Read the Excel files
    print("Reading Excel files...")
    urkunden_df = pd.read_excel("../datasheets/Urkunden_Repertorium_v4.xlsx")
    personen_df = pd.read_excel("../datasheets/personen/Personen_Empfaenger_Aussteller_final.xlsx")
    
    print(f"Charter document has {len(urkunden_df)} rows")
    print(f"Person list has {len(personen_df)} rows")
    
    # Add Nr. Rep column at the beginning of person dataframe
    personen_df.insert(0, 'Nr. Rep', '')
    
    # Get all name columns (Name 1, Name 2, Name 3, etc.)
    name_columns = [col for col in personen_df.columns if col.startswith('Name ')]
    # Also add the "Name" column if it exists
    if 'Name' in personen_df.columns:
        name_columns.append('Name')
    
    # Get all function and role columns
    funktion_columns = [col for col in personen_df.columns if col.startswith('Funktion')]
    rolle_columns = [col for col in personen_df.columns if col.startswith('Rolle in Urkunde')]
    
    print(f"Found name columns: {name_columns}")
    print(f"Found function columns: {funktion_columns}")
    print(f"Found role columns: {rolle_columns}")

    # Process each row in the charter document
    for idx, row in urkunden_df.iterrows():
        nr_rep = row.get('Nr. Rep', '')
        if pd.isna(nr_rep):
            continue

        for i in range(1, 99):  # Assuming max 99 persons per charter
            person_col = f'Lateinische Schreibweise genannte Person {i}'
            funktion_col = f'Funktion {i}'
            rolle_col = f'Rolle in Urkunde {i}'

            if person_col not in urkunden_df.columns:
                break

            person_name = row.get(person_col, '')
            funktion = row.get(funktion_col, '')
            rolle = row.get(rolle_col, '')
            urkunde_datum = row.get('Datum', '')
            urkunde_edition_angaben = row.get('Angaben zu Drucken/Editionen', '')

            # Skip if person name is empty
            if pd.isna(person_name) or person_name == '':
                continue

            print(f"Searching for: '{person_name}', Funktion: '{funktion}', Rolle: '{rolle}'")

            # 1. Nur nach Namen suchen
            name_matches = []
            for person_idx, person_row in personen_df.iterrows():
                for name_col in name_columns:
                    if name_col in personen_df.columns:
                        name_value = person_row.get(name_col, '')
                        if normalize_string(name_value) == normalize_string(person_name) and normalize_string(name_value) != '':
                            name_matches.append(person_idx)
                            break

            matches = []

            # 2. Falls mehrere Matches, versuche über Datum oder Edition zuzuordnen
            if len(name_matches) == 1:
                matches = name_matches
            elif len(name_matches) > 1:
                # Versuche über "Datum Funktion i"
                date_refined_matches = []
                if not pd.isna(urkunde_datum):
                    for person_idx in name_matches:
                        person_row = personen_df.loc[person_idx]
                        for k in range(1, 28):
                            datum_f_col = f'Datum Funktion {k}'
                            if datum_f_col in person_row:
                                person_datum_f = person_row.get(datum_f_col, '')
                                if not pd.isna(person_datum_f) and str(urkunde_datum) == str(person_datum_f):
                                    date_refined_matches.append(person_idx)
                                    break  # Nächster person_idx

                if len(date_refined_matches) == 1:
                    matches = date_refined_matches
                    print(f"DEBUG: Disambiguated '{person_name}' for Nr. Rep {nr_rep} by Datum Funktion: {urkunde_datum}")
                else:
                    # Wenn Datum nicht eindeutig war, versuche Edition als Substring
                    candidates = date_refined_matches if len(date_refined_matches) > 0 else name_matches
                    edition_refined_matches = []
                    if not pd.isna(urkunde_edition_angaben):
                        for person_idx in candidates:
                            person_row = personen_df.loc[person_idx]
                            for k in range(1, 28):
                                edition_col = f'Edition {k}'
                                if edition_col in person_row:
                                    person_edition = person_row.get(edition_col, '')
                                    if not pd.isna(person_edition) and str(person_edition).strip() and normalize_string(str(person_edition)) in normalize_string(str(urkunde_edition_angaben)):
                                        edition_refined_matches.append(person_idx)
                                        break  # Nächster person_idx
                    
                    if len(edition_refined_matches) == 1:
                        matches = edition_refined_matches
                        print(f"DEBUG: Disambiguated '{person_name}' for Nr. Rep {nr_rep} by Edition substring.")
                    elif len(edition_refined_matches) > 1:
                        print(f"WARNING: Match für '{person_name}' (Nr. Rep {nr_rep}) gefunden, aber Zuordnung über Edition nicht eindeutig: {edition_refined_matches}")
                        matches = [] # Nicht zuordnen
                    else: # 0 edition matches
                        if len(date_refined_matches) > 1:
                             print(f"WARNING: Match für '{person_name}' (Nr. Rep {nr_rep}) gefunden, aber Zuordnung über Datum nicht eindeutig und Edition hat nicht geholfen. Kandidaten: {date_refined_matches}")
                        else:
                             print(f"WARNING: Match für '{person_name}' (Nr. Rep {nr_rep}) gefunden, aber Zuordnung nicht möglich. Kandidaten: {name_matches}")
                        matches = [] # Nicht zuordnen

            # Handle matches
            if len(matches) == 1:
                # Exactly one match - add Nr. Rep
                current_nr_rep = personen_df.loc[matches[0], 'Nr. Rep']
                if pd.isna(current_nr_rep) or current_nr_rep == '':
                    personen_df.loc[matches[0], 'Nr. Rep'] = str(nr_rep)
                else:
                    # Check if nr_rep is already in the list
                    existing_values = str(current_nr_rep).split(';')
                    if str(nr_rep) not in existing_values:
                        personen_df.loc[matches[0], 'Nr. Rep'] = str(current_nr_rep) + ';' + str(nr_rep)
                print(f"Added Nr. Rep {nr_rep} to person: {person_name} at index {matches[0]}")
            elif len(name_matches) > 1 and len(matches) != 1:
                # Multiple matches that could not be resolved to one.
                # The warning has already been printed inside the disambiguation logic.
                pass
            elif len(name_matches) == 0:
                print(f"No match found for: '{person_name}', Funktion: '{funktion}', Rolle: '{rolle}'")
    
    # Save the updated person list
    output_file = "../datasheets/personen/Personen_Empfaenger_Aussteller_final_with_nr_rep.xlsx"
    personen_df.to_excel(output_file, index=False)
    print(f"Updated person list saved to: {output_file}")

if __name__ == "__main__":
    main()
