import pandas as pd
import os
import time
import requests
from tqdm import tqdm

def find_recipient(row):
    """Find recipient from 'Genannte Person n' columns where 'Rolle in Urkunde n' is 'Empfänger'"""
    # Get all Genannte Person columns and extract their numbers
    person_cols = [col for col in row.index if col.startswith('Genannte Person ')]
    
    for person_col in person_cols:
        # Extract the number from the column name (e.g., "Genannte Person 1" -> "1")
        try:
            person_num = person_col.replace('Genannte Person ', '').strip()
            rolle_col = f'Rolle in Urkunde {person_num}'
            
            # Check if corresponding rolle column exists
            if rolle_col in row.index:
                rolle_value = row[rolle_col]
                if pd.notna(rolle_value) and str(rolle_value).strip().lower() == 'empfänger':
                    person_value = row[person_col]
                    if pd.notna(person_value):
                        return str(person_value).strip()
        except:
            continue
    
    return ""

def translate_with_deepl(text, target_lang, max_retries=3):
    """
    Translate text using DeepL API with rate limiting and retry logic
    """
    if not text or text.strip() == "":
        return text
    
    import requests
    import time
    
    api_key = "b695037b-e738-4ce8-a99d-aab8f3aa9529:fx"
    url = "https://api-free.deepl.com/v2/translate"
    
    for attempt in range(max_retries):
        try:
            data = {
                'auth_key': api_key,
                'text': text,
                'source_lang': 'DE',
                'target_lang': target_lang
            }
            
            response = requests.post(url, data=data)
            
            if response.status_code == 429:  # Too Many Requests
                wait_time = (attempt + 1) * 2  # Exponential backoff: 2, 4, 6 seconds
                print(f"Rate limit hit, waiting {wait_time} seconds before retry {attempt + 1}/{max_retries}")
                time.sleep(wait_time)
                continue
            
            response.raise_for_status()
            
            result = response.json()
            translated_text = result['translations'][0]['text']
            
            # Add small delay between successful requests to avoid rate limiting
            time.sleep(0.2)
            
            return translated_text
            
        except requests.exceptions.RequestException as e:
            if attempt == max_retries - 1:  # Last attempt
                print(f"DeepL translation failed for '{text}' to {target_lang} after {max_retries} attempts: {e}")
                return text
            else:
                wait_time = (attempt + 1) * 2
                print(f"Request failed, retrying in {wait_time} seconds... (attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
        except Exception as e:
            print(f"DeepL translation failed for '{text}' to {target_lang}: {e}")
            return text
    
    return text

def translate_lde_components(aussteller, recipient, datum):
    """
    Translate aussteller and recipient separately using DeepL API
    """
    try:
        import time
        
        # Translate aussteller
        aussteller_en = translate_with_deepl(aussteller, 'EN') if aussteller else ""
        time.sleep(0.1)  # Small delay between requests
        aussteller_pl = translate_with_deepl(aussteller, 'PL') if aussteller else ""
        time.sleep(0.1)
        aussteller_uk = translate_with_deepl(aussteller, 'UK') if aussteller else ""
        time.sleep(0.1)
        
        # Translate recipient  
        recipient_en = translate_with_deepl(recipient, 'EN') if recipient else ""
        time.sleep(0.1)
        recipient_pl = translate_with_deepl(recipient, 'PL') if recipient else ""
        time.sleep(0.1)
        recipient_uk = translate_with_deepl(recipient, 'UK') if recipient else ""
        
        # Construct translated LDE strings
        en_lde = f"{aussteller_en} - {recipient_en} - {datum}" if aussteller_en and recipient_en and datum else ""
        pl_lde = f"{aussteller_pl} - {recipient_pl} - {datum}" if aussteller_pl and recipient_pl and datum else ""
        uk_lde = f"{aussteller_uk} - {recipient_uk} - {datum}" if aussteller_uk and recipient_uk and datum else ""
        
        return en_lde, pl_lde, uk_lde
    except Exception as e:
        print(f"Translation failed: {e}")
        return "", "", ""

def create_enriched_metadata():
    # Read Excel files
    metadata_file = "../datasheets/Imports/Urkunden_Metadaten.xlsx"
    repertorium_file = "../datasheets/Urkunden_Repertorium_v4.xlsx"
    
    try:
        df_metadata = pd.read_excel(metadata_file)
        df_repertorium = pd.read_excel(repertorium_file)
    except FileNotFoundError as e:
        print(f"Error reading files: {e}")
        return
    
    # Create new DataFrame with metadata structure (keep headers, remove data rows)
    df_new = df_metadata.head(0).copy()  # Keep only headers
    
    # Iterate through repertorium data with progress bar
    for _, rep_row in tqdm(df_repertorium.iterrows(), total=len(df_repertorium), desc="Processing documents"):
        new_row = {}
        
        # Copy all column headers from metadata
        for col in df_metadata.columns:
            new_row[col] = ""
        
        # Add Nr. Rep mapping (exact copy)
        if 'Nr. Rep' in df_new.columns:
            new_row['Nr. Rep'] = rep_row.get('Nr. Rep', "")
        
        if 'Datum (P106) (P43 Datum vor, P41 Datum nach)' in df_new.columns:
            new_row['Datum (P106) (P43 Datum vor, P41 Datum nach)'] = rep_row.get('Datum', "")
        
        if 'Ausstellungsort (P926)' in df_new.columns:
            new_row['Ausstellungsort (P926)'] = rep_row.get('Ausstellungsort', "")
        
        if 'Zusammenfassung (P724)' in df_new.columns:
            new_row['Zusammenfassung (P724)'] = rep_row.get('Regest', "")
        
        # Add Dde mapping with Regest content
        if 'Dde (wird mit Regest gefüllt)' in df_new.columns:
            new_row['Dde (wird mit Regest gefüllt)'] = rep_row.get('Regest', "")
        
        # Originalität: only text until first comma
        if 'Originalität (P115)' in df_new.columns:
            originality = str(rep_row.get('Angaben zur Überlieferung', ""))
            if originality and originality != "nan":
                new_row['Originalität (P115)'] = originality.split(',')[0].strip()
        
        # Fixed values
        if 'Gelistet in (P124)' in df_new.columns:
            new_row['Gelistet in (P124)'] = "Repertorium diplomatum terrae Russie Regni Poloniae (Q1213987)"
        
        if 'Notiz (P73), alles was im Rep. als Kommentar aufgeführt ist' in df_new.columns:
            new_row['Notiz (P73), alles was im Rep. als Kommentar aufgeführt ist'] = rep_row.get('Kommentar', "")
        
        if 'Forschungsprojekte, die zu diesem Datensatz beitrugen (P ?)' in df_new.columns:
            new_row['Forschungsprojekte, die zu diesem Datensatz beitrugen (P ?)'] = "VAMOD | Vormoderne Ambiguitäten modellieren (Q1206913)"
        
        if 'P2' in df_new.columns:
            new_row['P2'] = "Dokument (Q10671)"
        
        if 'Werktyp (P121)' in df_new.columns:
            new_row['Werktyp (P121)'] = "Urkunde (Q517290)"
        
        # Create Lde field: "Aussteller - Empfänger - Datum"
        if 'Lde' in df_new.columns:
            aussteller = str(rep_row.get('Aussteller', "")).strip()
            recipient = find_recipient(rep_row)
            datum = str(rep_row.get('Datum', "")).strip()
            
            if aussteller != "nan" and recipient and datum != "nan":
                lde_content = f"{aussteller} - {recipient} - {datum}"
                # Fix spacing issues in original content - preserve date formats
                import re
                # Find date patterns (YYYY-MM-DD or similar)
                date_pattern = r'\b\d{4}-\d{1,2}-\d{1,2}\b'
                dates = re.findall(date_pattern, lde_content)
                # Replace dates with placeholders
                temp_content = lde_content
                for i, date in enumerate(dates):
                    temp_content = temp_content.replace(date, f"__DATE_{i}__")
                # Fix spacing around remaining dashes
                temp_content = temp_content.replace('.-', '. -').replace('-', ' - ').replace('  -  ', ' - ').replace('  - ', ' - ').replace(' -  ', ' - ')
                # Restore dates
                for i, date in enumerate(dates):
                    temp_content = temp_content.replace(f"__DATE_{i}__", date)
                lde_content = temp_content
                new_row['Lde'] = lde_content
                
                # Add translations with progress indication
                tqdm.write(f"Translating: {aussteller[:30]}... - {recipient[:30]}...")
                en_translation, pl_translation, ukr_translation = translate_lde_components(aussteller, recipient, datum)
                
                if 'Len' in df_new.columns:
                    new_row['Len'] = en_translation
                
                if 'Lpl' in df_new.columns:
                    new_row['Lpl'] = pl_translation
                
                if 'Lukr' in df_new.columns:
                    new_row['Lukr'] = ukr_translation
        
        # Add row to new DataFrame
        df_new = pd.concat([df_new, pd.DataFrame([new_row])], ignore_index=True)
    
    # Save to new Excel file
    output_file = "../datasheets/Imports/Urkunden_Metadatenliste_v2.xlsx"
    df_new.to_excel(output_file, index=False)
    print(f"Created enriched metadata file: {output_file}")
    print(f"Total rows processed: {len(df_new)}")

if __name__ == "__main__":
    create_enriched_metadata()
