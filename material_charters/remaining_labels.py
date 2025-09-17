import os
import pandas as pd
import re
import asyncio
from openai import AsyncOpenAI

# -----------------------------------------------------------------------------
# API-Konfiguration
# -----------------------------------------------------------------------------
# (a) API-Key direkt gesetzt
api_key = "API KEY HIER EINFÜGEN"
model = "gpt-4o"

# (b) OpenAI-Client initialisieren
client = AsyncOpenAI(
    api_key=api_key
)

# Lade die Excel-Dateien
metadata_df = pd.read_excel('../datasheets/Imports/Urkunden_Metadatenliste_v9.xlsx')
repertorium_df = pd.read_excel('../datasheets/Urkunden_Repertorium_v4.xlsx')

# Funktion zum Extrahieren und Formatieren des Datums
def extract_and_format_date(date_str):
    if pd.isna(date_str):
        return ''
    # Entferne unnötige Teile wie (P106), (P43), etc.
    date_str = re.sub(r'\s*\([^)]*\)\s*', '', str(date_str)).strip()
    # Suche nach Datumsformaten wie YYYY-MM-DD, YYYY-MM, YYYY
    match = re.search(r'(\d{4})(?:-(\d{2})(?:-(\d{2}))?)?', date_str)
    if match:
        year = match.group(1)
        month = match.group(2) if match.group(2) else '00'
        day = match.group(3) if match.group(3) else '00'
        if day != '00':
            return f"{year}-{month}-{day}"
        elif month != '00':
            return f"{year}-{month}"
        else:
            return year
    # Spezielle Behandlung für /8 (z.B. 1410-00-00T00 -> 1410-1420)
    if '/8' in date_str:
        match = re.search(r'(\d{4})', date_str)
        if match:
            start_year = int(match.group(1))
            return f"{start_year}-{start_year + 9}"
    return date_str  # Fallback

# Iteriere durch die Metadatenliste
for index, row in metadata_df.iterrows():
    if pd.isna(row.get('Lde', '')) or row['Lde'] == '':
        nr_rep = row.get('Nr. Rep')
        if pd.notna(nr_rep):
            # Finde entsprechende Zeile in Repertorium
            rep_row = repertorium_df[repertorium_df['Nr. Rep'] == nr_rep]
            if not rep_row.empty:
                rep_row = rep_row.iloc[0]
                aussteller = rep_row.get('Aussteller', '')
                
                # Finde Empfänger
                empfänger = ''
                # Zuerst prüfe Genannte Person 1, wenn vorhanden
                person1 = str(rep_row.get('Genannte Person 1', '')).strip()
                if person1:
                    empfänger = person1
                else:
                    for i in range(1, 4):
                        rolle_col = f'Rolle in Urkunde {i}'
                        person_col = f'Genannte Person {i}'
                        rolle = str(rep_row.get(rolle_col, '')).strip()
                        person = str(rep_row.get(person_col, '')).strip()
                        if rolle and ('Empfänger' in rolle or 'Käufer' in rolle):
                            empfänger = person
                            break
                        elif not rolle and person:
                            empfänger = person
                            break
                # Wenn 'nan', setze zu leer
                if empfänger.lower() == 'nan':
                    empfänger = ''
                
                # Extrahiere und formatiere Datum
                datum_col = 'Datum (P106) (P43 Datum vor, P41 Datum nach)'
                datum = extract_and_format_date(row.get(datum_col, ''))
                
                # Kombiniere in Lde
                lde_value = f"{aussteller} - {empfänger} - {datum}".strip(' - ')
                metadata_df.at[index, 'Lde'] = lde_value

# Speichere die aktualisierte Metadatenliste
metadata_df.to_excel('../datasheets/Imports/Urkunden_Metadatenliste_v10.xlsx', index=False)
print("Aktualisierung abgeschlossen. Datei gespeichert als Urkunden_Metadatenliste_v10.xlsx")

# -----------------------------------------------------------------------------
# Übersetzungen für Len, Lpl, Lukr
# -----------------------------------------------------------------------------
async def translate_row(index, lde, semaphore):
    async with semaphore:
        print(f"Übersetze für Zeile {index}: {lde}")
        
        len_text = ''
        lpl_text = ''
        lukr_text = ''
        
        # Übersetze ins Englische
        try:
            response = await client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": f"Übersetze den folgenden Text ins Englische, behalte die Struktur mit Bindestrichen und Leerzeichen bei, übersetze nur die Namen, lasse das Datum unverändert: {lde}"}],
                temperature=0.0,
                max_tokens=256
            )
            len_text = response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Fehler bei Englisch-Übersetzung für Zeile {index}: {e}")
        
        # Übersetze ins Polnische
        try:
            response = await client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": f"Übersetze den folgenden Text ins Polnische, behalte die Struktur mit Bindestrichen und Leerzeichen bei, übersetze nur die Namen, lasse das Datum unverändert: {lde}"}],
                temperature=0.0,
                max_tokens=256
            )
            lpl_text = response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Fehler bei Polnisch-Übersetzung für Zeile {index}: {e}")
        
        # Übersetze ins Ukrainische
        try:
            response = await client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": f"Übersetze den folgenden Text ins Ukrainische, behalte die Struktur mit Bindestrichen und Leerzeichen bei, übersetze nur die Namen, lasse das Datum unverändert: {lde}"}],
                temperature=0.0,
                max_tokens=256
            )
            lukr_text = response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Fehler bei Ukrainisch-Übersetzung für Zeile {index}: {e}")
        
        return index, len_text, lpl_text, lukr_text

async def main():
    print("Starte Übersetzungen...")
    
    semaphore = asyncio.Semaphore(3)  # Limitiere auf 3 gleichzeitige Anfragen
    tasks = []
    
    for index, row in metadata_df.iterrows():
        lde = row.get('Lde', '')
        if lde and (pd.isna(row.get('Len', '')) or row['Len'] == '') and (pd.isna(row.get('Lpl', '')) or row['Lpl'] == '') and (pd.isna(row.get('Lukr', '')) or row['Lukr'] == ''):
            tasks.append(translate_row(index, lde, semaphore))
    
    results = await asyncio.gather(*tasks)
    
    for index, len_text, lpl_text, lukr_text in results:
        metadata_df.at[index, 'Len'] = len_text
        metadata_df.at[index, 'Lpl'] = lpl_text
        metadata_df.at[index, 'Lukr'] = lukr_text
    
    # Speichere die finale aktualisierte Metadatenliste
    metadata_df.to_excel('../datasheets/Imports/Urkunden_Metadatenliste_v10.xlsx', index=False)
    print("Übersetzungen abgeschlossen. Finale Datei gespeichert als Urkunden_Metadatenliste_v10.xlsx")

if __name__ == "__main__":
    asyncio.run(main())