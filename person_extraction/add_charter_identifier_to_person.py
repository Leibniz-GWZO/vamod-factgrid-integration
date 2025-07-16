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
            
        # Extract person information from charter
        for i in range(1, 99):  # Assuming max 99 persons per charter
            person_col = f'Lateinische Schreibweise genannte Person {i}'
            funktion_col = f'Funktion {i}'
            rolle_col = f'Rolle in Urkunde {i}'
            

            if person_col not in urkunden_df.columns:
                break
                
            person_name = row.get(person_col, '')
            funktion = row.get(funktion_col, '')
            rolle = row.get(rolle_col, '')

            print(person_name, funktion, rolle)
            
            # Skip if person name is empty
            if pd.isna(person_name) or person_name == '':
                continue
                
            print(f"Searching for: '{person_name}', Funktion: '{funktion}', Rolle: '{rolle}'")
            
            # Search for matches in person list
            matches = []
            
            
            for person_idx, person_row in personen_df.iterrows():

                # Check if any name variant matches
                name_match = False
                for name_col in name_columns:
                    if name_col in personen_df.columns:
                        name_value = person_row.get(name_col, '')
                        if normalize_string(name_value) == normalize_string(person_name) and normalize_string(name_value) != '':
                            name_match = True
                            break
                
                if not name_match:
                    continue
                    
                # Check if function and role also match (must be from same index)
                funktion_rolle_match = False
                
                # Check all function/role combinations with same index
                for funktion_col in funktion_columns:
                    person_funktion = person_row.get(funktion_col, '')
                    
                    # Extract index from function column name
                    if ' ' in funktion_col:
                        try:
                            func_index = funktion_col.split(' ')[1]
                            corresponding_rolle_col = f'Rolle in Urkunde {func_index}'
                        except:
                            corresponding_rolle_col = 'Rolle in Urkunde'
                    else:
                        corresponding_rolle_col = 'Rolle in Urkunde'
                    
                    # Check if corresponding role column exists
                    if corresponding_rolle_col in personen_df.columns:
                        person_rolle = person_row.get(corresponding_rolle_col, '')
                        
                        funktion_normalized = normalize_string(funktion)
                        rolle_normalized = normalize_string(rolle)
                        person_funktion_normalized = normalize_string(person_funktion)
                        person_rolle_normalized = normalize_string(person_rolle)
                        
                        funktion_match = funktion_normalized == person_funktion_normalized
                        rolle_match = rolle_normalized == person_rolle_normalized
                        
                        # Both function and role must match (empty values can match empty values)
                        if funktion_match and rolle_match:
                            funktion_rolle_match = True
                            break
                
                if funktion_rolle_match:
                    matches.append(person_idx)

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
                print(f"Added Nr. Rep {nr_rep} to person: {person_name}")
            elif len(matches) > 1:
                # Multiple matches - print warning
                print(f"WARNING: Multiple matches found for person '{person_name}' with function '{funktion}' and role '{rolle}' from Nr. Rep {nr_rep}")
                print(f"  Matched rows: {matches}")
            else:
                print(f"No match found for: '{person_name}', Funktion: '{funktion}', Rolle: '{rolle}'")
    
    # Save the updated person list
    output_file = "../datasheets/personen/Personen_Empfaenger_Aussteller_final_with_nr_rep.xlsx"
    personen_df.to_excel(output_file, index=False)
    print(f"Updated person list saved to: {output_file}")

if __name__ == "__main__":
    main()
