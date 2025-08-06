import pandas as pd
import os
import time
import requests
from tqdm import tqdm

def translate_text(content, target_lang, max_retries=3):
    """
    Translate using DeepL API
    """
    if not content or content.strip() == "":
        return ""
    
    api_key = "b695037b-e738-4ce8-a99d-aab8f3aa9529:fx"
    url = "https://api-free.deepl.com/v2/translate"
    
    target_lang_map = {
        'EN-US': 'EN',
        'PL': 'PL', 
        'UK': 'UK'
    }
    
    deepl_target_lang = target_lang_map.get(target_lang, target_lang)
    
    for attempt in range(max_retries):
        try:
            data = {
                'auth_key': api_key,
                'text': content,
                'source_lang': 'DE',
                'target_lang': deepl_target_lang
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
            
            # Fix spacing issues: ensure space after periods before "-"
            translated_text = translated_text.replace('.-', '. -')
            # Also fix general spacing around dashes
            translated_text = translated_text.replace('-', ' - ').replace('  -  ', ' - ').replace('  - ', ' - ').replace(' -  ', ' - ')
            
            # Add small delay between successful requests to avoid rate limiting
            time.sleep(0.2)
            
            return translated_text
            
        except requests.exceptions.RequestException as e:
            if attempt == max_retries - 1:  # Last attempt
                print(f"DeepL translation failed for '{content[:50]}...' to {target_lang} after {max_retries} attempts: {e}")
                return content
            else:
                wait_time = (attempt + 1) * 2
                print(f"Request failed, retrying in {wait_time} seconds... (attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
        except Exception as e:
            print(f"DeepL translation failed for '{content[:50]}...' to {target_lang}: {e}")
            return content
    
    return content

def translate_lde_content(lde_content):
    """
    Translate Lde content to multiple languages using DeepL API
    """
    if not lde_content or lde_content.strip() == "":
        return "", "", ""
    
    en_translation = translate_text(lde_content, 'EN-US')
    pl_translation = translate_text(lde_content, 'PL')
    uk_translation = translate_text(lde_content, 'UK')
    
    return en_translation, pl_translation, uk_translation

def fix_lde_spacing_and_add_dde_translations():
    # Read the existing metadata file
    metadata_file = "../datasheets/Imports/Urkunden_Metadatenliste_v2.xlsx"
    
    try:
        df = pd.read_excel(metadata_file)
        print(f"Loaded {len(df)} rows from {metadata_file}")
    except FileNotFoundError as e:
        print(f"Error reading file: {e}")
        return
    
    # Fix Lde spacing issue and translate if needed
    if 'Lde' in df.columns:
        print("Fixing Lde spacing issues...")
        for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing Lde spacing"):
            lde_content = str(row['Lde']).strip()
            if lde_content and lde_content != "nan":
                # Fix spacing: ensure there's a space before and after each "-" and after periods
                fixed_lde = lde_content.replace('.-', '. -').replace('-', ' - ').replace('  -  ', ' - ').replace('  - ', ' - ').replace(' -  ', ' - ')
                df.at[idx, 'Lde'] = fixed_lde
                
                # Only translate if translations are missing
                need_translation = False
                if 'Len' in df.columns and (pd.isna(row['Len']) or str(row['Len']).strip() == ""):
                    need_translation = True
                if 'Lpl' in df.columns and (pd.isna(row['Lpl']) or str(row['Lpl']).strip() == ""):
                    need_translation = True
                if 'Lukr' in df.columns and (pd.isna(row['Lukr']) or str(row['Lukr']).strip() == ""):
                    need_translation = True
                
                if need_translation:
                    tqdm.write(f"Translating Lde: {fixed_lde[:50]}...")
                    en_translation, pl_translation, uk_translation = translate_lde_content(fixed_lde)
                    
                    if 'Len' in df.columns:
                        df.at[idx, 'Len'] = en_translation
                    if 'Lpl' in df.columns:
                        df.at[idx, 'Lpl'] = pl_translation
                    if 'Lukr' in df.columns:
                        df.at[idx, 'Lukr'] = uk_translation
    
    # Translate Dde content to Den, Dpl, Dukr
    if 'Dde (wird mit Regest gefüllt)' in df.columns:
        print("Translating Dde content...")
        
        # Ensure target columns exist
        if 'Den' not in df.columns:
            df['Den'] = ""
        if 'Dpl' not in df.columns:
            df['Dpl'] = ""
        if 'Dukr' not in df.columns:
            df['Dukr'] = ""
        
        for idx, row in tqdm(df.iterrows(), total=len(df), desc="Translating Dde content"):
            dde_content = str(row['Dde (wird mit Regest gefüllt)']).strip()
            if dde_content and dde_content != "nan":
                # Only translate if translations are missing
                need_translation = False
                if pd.isna(row['Den']) or str(row['Den']).strip() == "":
                    need_translation = True
                if pd.isna(row['Dpl']) or str(row['Dpl']).strip() == "":
                    need_translation = True
                if pd.isna(row['Dukr']) or str(row['Dukr']).strip() == "":
                    need_translation = True
                
                if need_translation:
                    tqdm.write(f"Translating Dde: {dde_content[:50]}...")
                    
                    # Translate to English (Den)
                    if pd.isna(row['Den']) or str(row['Den']).strip() == "":
                        en_translation = translate_text(dde_content, 'EN-US')
                        df.at[idx, 'Den'] = en_translation
                    
                    # Translate to Polish (Dpl)
                    if pd.isna(row['Dpl']) or str(row['Dpl']).strip() == "":
                        pl_translation = translate_text(dde_content, 'PL')
                        df.at[idx, 'Dpl'] = pl_translation
                    
                    # Translate to Ukrainian (Dukr)
                    if pd.isna(row['Dukr']) or str(row['Dukr']).strip() == "":
                        uk_translation = translate_text(dde_content, 'UK')
                        df.at[idx, 'Dukr'] = uk_translation
                    
                    # Add small delay to avoid rate limiting (already handled in translate_text function)
                    pass
    
    # Save the updated file
    output_file = "../datasheets/Imports/Urkunden_Metadatenliste_v3.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Updated metadata file saved as: {output_file}")
    print(f"Total rows processed: {len(df)}")

if __name__ == "__main__":
    fix_lde_spacing_and_add_dde_translations()
