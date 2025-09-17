import pandas as pd
import os
from typing import List, Optional

def extract_ade_entries(row: pd.Series) -> List[str]:
    """Extract all entries from columns starting with 'Ade'"""
    ade_entries = []
    for col in row.index:
        if col.startswith('Ade') and pd.notna(row[col]) and str(row[col]).strip():
            entry = str(row[col]).strip()
            ade_entries.append(entry)
    
    # If there's only one entry total, allow single-part names
    if len(ade_entries) <= 1:
        return ade_entries
    
    # If there are multiple entries, only keep names with at least two parts
    filtered_entries = []
    for entry in ade_entries:
        if len(entry.split(" ")) >= 2:
            filtered_entries.append(entry)
    
    print(filtered_entries)
    
    return filtered_entries

def find_matching_person(ade_names: List[str], repertorium_row: pd.Series) -> Optional[str]:
    """Find matching person name in Repertorium row and return standardized version"""
    matched_names = []
    
    # Check columns "Lateinische Schreibweise genannte Person 1" to "Lateinische Schreibweise genannte Person 6"
    for i in range(1, 7):
        lat_col = f"Lateinische Schreibweise genannte Person {i}"
        person_col = f"Genannte Person {i}"
        
        if lat_col in repertorium_row.index and person_col in repertorium_row.index:
            lat_name = repertorium_row[lat_col]
            person_name = repertorium_row[person_col]
            
            if pd.notna(lat_name) and pd.notna(person_name):
                lat_name_str = str(lat_name).strip()
                person_name_str = str(person_name).strip()
                
                # Check if any of the Ade names exactly match the Latin name
                for ade_name in ade_names:
                    if ade_name == lat_name_str:
                        if person_name_str not in matched_names:
                            matched_names.append(person_name_str)
    
    # Return all matches joined with semicolon, or None if no matches
    return "; ".join(matched_names) if matched_names else None

def process_sheet(df: pd.DataFrame, sheet_name: str, repertorium_dict: dict) -> int:
    """Process a single sheet and return number of matches found"""

    
    matches_found = 0
    for idx in range(3, len(df)):  # Starting from row 4 (index 3)
        row = df.iloc[idx]
        
        # Extract Ade entries
        ade_names = extract_ade_entries(row)
        if not ade_names:
            continue
            
        # Get qal90 value
        qal90_value = row.get("qal90 (Nr. im Rep.)")
        if pd.isna(qal90_value):
            continue
            
        # Split qal90 by semicolon to handle multiple entries
        qal90_entries = [entry.strip() for entry in str(qal90_value).split(';') if entry.strip()]
        
        all_matched_names = []
        
        # Check each qal90 entry
        for qal90_str in qal90_entries:
            # Find matching repertorium entry
            if qal90_str in repertorium_dict:
                repertorium_row = repertorium_dict[qal90_str]
                
                # Find matching person name
                standardized_name = find_matching_person(ade_names, repertorium_row)
                
                if standardized_name:
                    # Split multiple names and add to collection
                    names = [name.strip() for name in standardized_name.split(';')]
                    for name in names:
                        if name not in all_matched_names:
                            all_matched_names.append(name)
        
        if all_matched_names:
            # Update the Lde column with all found matches
            final_standardized_names = "; ".join(all_matched_names)
            df.at[idx, "Lde"] = final_standardized_names
            matches_found += 1
            #print(f"{sheet_name} Row {idx + 1}: Found match for {ade_names} -> {final_standardized_names}")
        
    
    return matches_found

def main():
    # File paths
    personen_file = "../datasheets/Imports/Personen.xlsx"
    repertorium_file = "../datasheets/Urkunden_Repertorium_v4.xlsx"
    
    # Read the repertorium file
    print("Reading Urkunden_Repertorium_v4.xlsx...")
    repertorium_df = pd.read_excel(repertorium_file)
    
    # Create a dictionary for faster lookup of repertorium entries
    repertorium_dict = {}
    for idx, row in repertorium_df.iterrows():
        nr_rep = row.get("Nr. Rep")
        if pd.notna(nr_rep):
            repertorium_dict[str(nr_rep).strip()] = row
    
    # Read all sheets from the original file
    print("Reading all sheets from Personen.xlsx...")
    all_sheets = pd.read_excel(personen_file, sheet_name=None)
    
    # Process both target sheets
    sheets_to_process = ["Ausgangstab_ohne_rot_markiert", "Gruppen_Städte_Bistüm_Konvent", "Personen"]
    total_matches = 0
    
    for sheet_name in sheets_to_process:
        if sheet_name in all_sheets:
            #print(f"\nProcessing sheet: {sheet_name}")
            matches = process_sheet(all_sheets[sheet_name], sheet_name, repertorium_dict)
            total_matches += matches
            #print(f"Found {matches} matches in {sheet_name}")
    
    #print(f"\nTotal matches found: {total_matches}")
    
    # Save the updated file with all original sheets preserved
    output_file = "../datasheets/Imports/Personen_v2.xlsx"
    
    # Save all sheets to the new file
    print("Saving all sheets to new file...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Updated file saved as: {output_file}")
    print(f"Preserved {len(all_sheets)} sheets from original file")

if __name__ == "__main__":
    main()
