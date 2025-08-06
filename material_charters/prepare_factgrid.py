#!/usr/bin/env python3
"""
Script to prepare Excel data for Factgrid import by converting date formats.
Only converts dates already in YYYY-MM-DD format to Factgrid format (+YYYY-MM-DDTHH:MM:SSZ/precision).
Other questionable date formats are left unchanged.
"""

import pandas as pd
import re
from pathlib import Path


def convert_date_to_factgrid_format(date_str):
    """
    Convert date string to Factgrid format if it's already in YYYY-MM-DD format.
    
    Args:
        date_str: Date string to convert
        
    Returns:
        Converted date string in Factgrid format or original string if not YYYY-MM-DD
    """
    if pd.isna(date_str) or not isinstance(date_str, str):
        return date_str
    
    # Check if date is in YYYY-MM-DD format
    date_pattern = r'^\d{4}-\d{2}-\d{2}$'
    
    if re.match(date_pattern, date_str.strip()):
        # Convert to Factgrid format: +YYYY-MM-DDTHH:MM:SSZ/precision
        # Using precision 11 for day-level precision
        factgrid_date = f"+{date_str.strip()}T00:00:00Z/11"
        return factgrid_date
    
    # Return unchanged if not in YYYY-MM-DD format
    return date_str


def convert_originality_to_factgrid(value):
    """
    Convert originality values to Factgrid Q-IDs.
    
    Args:
        value: Original value from the Originalität column
        
    Returns:
        Corresponding Q-ID or original value if no match
    """
    if pd.isna(value):
        return value
    
    value_str = str(value).strip()
    
    if value_str == "Original":
        return "Q11177"
    elif value_str == "Kopial":
        return "Q1207589"
    else:
        return value


def prepare_factgrid_data():
    """
    Process the Excel file and convert date formats for Factgrid import.
    """
    input_file = Path("../datasheets/Imports/Urkunden_Metadatenliste_v5.xlsx")
    output_file = Path("../datasheets/Imports/Urkunden_Metadatenliste_v6.xlsx")
    
    try:
        # Read the Excel file
        print(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file)
        
        # Display column names to verify the date column
        print(f"\nColumns in the file:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        
        # Find the date column
        date_column = "Datum (P106) (P43 Datum vor, P41 Datum nach)"
        
        if date_column not in df.columns:
            print(f"\nWarning: Column '{date_column}' not found!")
            print("Available columns:")
            for col in df.columns:
                if 'datum' in col.lower() or 'date' in col.lower():
                    print(f"  - {col}")
            return
        
        print(f"\nProcessing date column: {date_column}")
        
        # Show some example values before conversion
        print(f"\nSample values before conversion:")
        sample_values = df[date_column].dropna().head(10)
        for i, value in enumerate(sample_values):
            print(f"  {i+1}: {value}")
        
        # Count dates in YYYY-MM-DD format
        yyyy_mm_dd_pattern = r'^\d{4}-\d{2}-\d{2}$'
        convertible_dates = df[date_column].astype(str).str.match(yyyy_mm_dd_pattern, na=False)
        convertible_count = convertible_dates.sum()
        
        print(f"\nFound {convertible_count} dates in YYYY-MM-DD format that will be converted")
        print(f"Total non-null dates: {df[date_column].notna().sum()}")
        
        # Apply the date conversion
        df[date_column] = df[date_column].apply(convert_date_to_factgrid_format)
        
        # Show some example values after date conversion
        print(f"\nSample values after date conversion:")
        sample_values_after = df[date_column].dropna().head(10)
        for i, value in enumerate(sample_values_after):
            print(f"  {i+1}: {value}")
        
        # Process other columns with specific transformations
        print(f"\nProcessing other columns...")
        
        # Transform Originalität column
        originality_column = "Originalität (P115)"
        if originality_column in df.columns:
            print(f"Converting {originality_column} values...")
            before_originality = df[originality_column].value_counts()
            print(f"Before conversion: {dict(before_originality)}")
            df[originality_column] = df[originality_column].apply(convert_originality_to_factgrid)
            after_originality = df[originality_column].value_counts()
            print(f"After conversion: {dict(after_originality)}")
        else:
            print(f"Warning: Column '{originality_column}' not found!")
        
        # Replace all values in specific columns with fixed Q-IDs
        column_replacements = {
            "P2": "Q517290",
            "Werktyp (P121)": "Q517290", 
            "Gelistet in (P124)": "Q1213987",
            "Forschungsprojekte, die zu diesem Datensatz beitrugen (P ?)": "Q1206913"
        }
        
        for column_name, replacement_value in column_replacements.items():
            if column_name in df.columns:
                non_null_count = df[column_name].notna().sum()
                print(f"Replacing all {non_null_count} non-null values in '{column_name}' with '{replacement_value}'")
                df.loc[df[column_name].notna(), column_name] = replacement_value
            else:
                print(f"Warning: Column '{column_name}' not found!")
        
        # Count total transformations
        converted_dates = df[date_column].astype(str).str.contains(r'^\+\d{4}-\d{2}-\d{2}T00:00:00Z/11$', na=False)
        total_date_conversions = converted_dates.sum()
        
        # Save the modified file
        print(f"\nSaving converted data to: {output_file}")
        df.to_excel(output_file, index=False)
        
        print(f"\nConversion completed successfully!")
        print(f"- Input file: {input_file}")
        print(f"- Output file: {output_file}")
        print(f"- Converted {convertible_count} dates to Factgrid format")
        print(f"- Total dates now in Factgrid format: {total_date_conversions}")
        print(f"- Processed additional columns with Q-ID replacements")
        
    except FileNotFoundError:
        print(f"Error: File {input_file} not found!")
        print("Please make sure the file exists in the correct location.")
    except Exception as e:
        print(f"Error processing file: {str(e)}")


if __name__ == "__main__":
    prepare_factgrid_data()