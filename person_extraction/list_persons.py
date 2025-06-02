import pandas as pd
import re
from datetime import datetime

# Pfad zur Excel-Datei
excel_file = "datasheets/180809_Urkunden_Personen_korrekt_Sortiert2.xlsx"

# Mapping: Blattname -> Header-Zeile (0-basiert)
header_mapping = {
    "EL": 2,
    "WO": 2,
    "LA": 2,
    "MU": 2,
    "CR": 2,
    "WII": 2,
    "Erzb.": 2,
    "JA": 1,
    "Gf.Litauen": 0,           
    # "alt WII 10-1399 - 08-1409": 0,  # Diese Mappe wird ignoriert
    "KIII": 0
}

def parse_date(row, sheet):
    """
    Erzeugt ein Datum aus den Spalten 'Jahr' und 'Datum'.
    - Bei Blatt 'KIII' wird ausschließlich die Spalte 'Jahr' genutzt.
      Falls das Datum das Muster "XX" enthält (z. B. "1361-XX-XX ?"), wird nur das Jahr verwendet.
    - Bei allen anderen Blättern: 
        * Es werden die ersten 4 Zeichen aus 'Jahr' als Jahr entnommen.
        * Es wird versucht, aus 'Datum' Tag und Monat zu extrahieren. Scheitert dies,
          wird als Fallback nur das Jahr genutzt.
    """
    if sheet == "KIII":
        year_str = str(row["Jahr"]).strip()
        # Entferne ein eventuell vorhandenes Fragezeichen am Ende
        year_str = re.sub(r'\?$', '', year_str)
        # Falls der String ungültige Teile enthält, wie "XX", nutze nur die ersten 4 Zeichen
        if "XX" in year_str:
            year_extracted = year_str[:4]
            try:
                dt = datetime.strptime(year_extracted, "%Y")
                return dt
            except Exception as e:
                return pd.NaT
        else:
            try:
                if len(year_str) > 4:
                    # Erwartetes Format: YYYY-MM-DD
                    dt = datetime.strptime(year_str, "%Y-%m-%d")
                else:
                    dt = datetime.strptime(year_str, "%Y")
                return dt
            except Exception as e:
                return pd.NaT
    else:
        # Für andere Blätter: Verwende Jahr aus 'Jahr' als Basis
        year_val = str(row["Jahr"]).strip()
        year_extracted = year_val[:4]
        datum_val = str(row["Datum"]).strip() if "Datum" in row and pd.notnull(row["Datum"]) else ""
        datum_val = re.sub(r'\?$', '', datum_val).rstrip('.')  # Entferne Fragezeichen und überflüssige Punkte
        
        # Falls kein Datum angegeben ist, verwende nur das Jahr
        if datum_val == "":
            try:
                dt = datetime.strptime(year_extracted, "%Y")
                return dt
            except Exception as e:
                return pd.NaT
        
        # Versuch, Tag und Monat aus Datum zu extrahieren
        parts = datum_val.split('.')
        if len(parts) >= 2:
            try:
                day = int(parts[0])
                month = int(parts[1])
                dt = datetime(year=int(year_extracted), month=month, day=day)
                return dt
            except Exception as e:
                # Fallback: Wenn Tag/Monat nicht geparst werden können, verwende nur das Jahr
                try:
                    dt = datetime.strptime(year_extracted, "%Y")
                    return dt
                except Exception as e:
                    return pd.NaT
        # Falls das Datum-Format unerwartet ist, verwende als Fallback nur das Jahr
        try:
            dt = datetime.strptime(year_extracted, "%Y")
            return dt
        except Exception as e:
            return pd.NaT

# Alle Arbeitsblätter einlesen
xls = pd.ExcelFile(excel_file)
sheet_names = xls.sheet_names

# Überspringe die Mappe "alt WII 10-1399 - 08-1409"
sheet_names = [sheet for sheet in sheet_names if sheet != "alt WII 10-1399 - 08-1409"]

all_dfs = []

for sheet in sheet_names:
    header_row = header_mapping.get(sheet, 0)
    df = pd.read_excel(excel_file, sheet_name=sheet, header=header_row)
    
    df = df.dropna(subset=["Genannte Personen", "Funktion"], how='all')
    
    # Sicherstellen, dass die erforderlichen Spalten vorhanden sind
    for col in ["Genannte Personen", "Funktion", "Rolle in Urkunde"]:
        if col not in df.columns:
            raise ValueError(f"Spalte '{col}' nicht gefunden in Arbeitsblatt '{sheet}'.")
    
    if sheet == "KIII":
        if "Geschlecht" not in df.columns:
            raise ValueError("Spalte 'Geschlecht' nicht gefunden in Arbeitsblatt 'KIII'.")
    elif sheet == "CR":
        if "Familie" not in df.columns:
            raise ValueError("Spalte 'Familie' nicht gefunden in Arbeitsblatt 'CR'.")
        df = df.rename(columns={"Familie": "Geschlecht"})
    else:
        df["Geschlecht"] = ""
    
    # Erzeuge die Spalte 'ausfuehrungsdatum' mithilfe der Funktion parse_date
    df["ausfuehrungsdatum"] = df.apply(lambda row: parse_date(row, sheet), axis=1)
    
    all_dfs.append(df)

# Alle DataFrames zusammenführen
combined_df = pd.concat(all_dfs, ignore_index=True)

## DEBUGGING: Ausgabe aller Zeilen, in denen 'ausfuehrungsdatum' NaT ist
for idx, wert in combined_df["ausfuehrungsdatum"].items():
    if pd.isna(wert):
         print(f"Zeile {idx}:")
         print(combined_df.loc[idx])
         print("-" * 50)

# Konvertiere weitere relevante Spalten in Kleinbuchstaben
columns_to_lower = ["Genannte Personen", "Funktion", "Geschlecht", "Rolle in Urkunde"]
for col in columns_to_lower:
    combined_df[col] = combined_df[col].fillna('').astype(str).str.lower()

combined_df.columns = combined_df.columns.str.lower()

# Optional: Zeilen ohne Angabe in "genannte personen" herausfiltern
combined_df = combined_df[combined_df["genannte personen"].str.strip() != '']

# Gruppierung: Aggregation nach Person, Funktion, Geschlecht und Rolle in Urkunde
aggregated = combined_df.groupby(
    ["genannte personen", "funktion", "geschlecht", "rolle in urkunde"],
    dropna=False
).agg(
    genannte_häufigkeit=('ausfuehrungsdatum', 'size'),
    früheste_nennung=('ausfuehrungsdatum', lambda x: x[x.notna()].min() if x[x.notna()].size > 0 else pd.NaT),
    späteste_nennung=('ausfuehrungsdatum', lambda x: x[x.notna()].max() if x[x.notna()].size > 0 else pd.NaT)
).reset_index()

# Sortierung: Zuerst alphabetisch nach "genannte personen", dann absteigend nach Häufigkeit
aggregated = aggregated.sort_values(by=["genannte personen", "genannte_häufigkeit"], ascending=[True, False])

# Formatierung der Datumsangaben im gewünschten Format YYYY-MM-DD
aggregated["früheste_nennung"] = aggregated["früheste_nennung"].apply(lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "")
aggregated["späteste_nennung"] = aggregated["späteste_nennung"].apply(lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "")

# Umbenennen der Spalte "genannte personen" in "person"
aggregated = aggregated.rename(columns={"genannte personen": "person"})

# Hier erfolgt die Erweiterung der Funktion, wenn in der Rolle "vorgänger" vorkommt
mask = aggregated["rolle in urkunde"].str.contains("vorgänger", case=False, na=False)
aggregated.loc[mask, "funktion"] = aggregated.loc[mask, "funktion"].apply(
    lambda x: f"{x},vorgängerprivileg" if x.strip() != "" else "vorgängerprivileg"
)

# Speichern als Excel-Datei
excel_output = "datasheets/Personenliste.xlsx"
aggregated.to_excel(excel_output, index=False)

print(f"Excel-Datei wurde erfolgreich als '{excel_output}' gespeichert.")