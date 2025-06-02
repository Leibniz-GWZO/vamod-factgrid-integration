import pandas as pd
import re

# Ergebnisliste, in der alle Zeilen als Dictionaries gesammelt werden
result = []

# ---------------------------
# 1. Tabelle: datasheets/180809_Urkunden_Personen_korrekt_Sortiert2.xlsx
# ---------------------------
file1 = "datasheets/180809_Urkunden_Personen_korrekt_Sortiert2.xlsx"

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
    "KIII": 0
}

print("Starte Verarbeitung von Datei 1...")

for sheet, header in header_mapping.items():
    try:
        df = pd.read_excel(file1, sheet_name=sheet, header=header)
        print(f"Sheet '{sheet}' erfolgreich eingelesen mit {df.shape[0]} Zeilen und {df.shape[1]} Spalten.")
    except Exception as e:
        print(f"Fehler beim Einlesen von {sheet} in {file1}: {e}")
        continue

    # Für jede Zeile die Spalten "Ausstellungsort" und "Zielort" auswerten
    for col in ["Ausstellungsort", "Zielort"]:
        if col in df.columns:
            for index, row in df.iterrows():
                cell_value = row[col]
                if pd.isna(cell_value):
                    continue
                # Entferne "u. a." aus dem Zelleninhalt, damit es nicht als Trennzeichen missverstanden wird
                cell_value_clean = str(cell_value).replace("u. a.", "")
                cell_value_clean = str(cell_value).replace("u.a.", "")
                # Debug: Zeige den bereinigten Originalinhalt der Zelle
                print(f"Sheet '{sheet}', Zeile {index}, Spalte '{col}' (bereinigt): {cell_value_clean}")
                # Splitte den bereinigten Zelleninhalt bei "/", "," oder ";" oder bei " u. " (als abgekürztes 'und')
                delimiter_pattern = r'[/,;]|(?:\s+u\.\s+)'
                orte = [x.strip() for x in re.split(delimiter_pattern, cell_value_clean) if x.strip()]
                for ort in orte:
                    result.append({
                        "Ort": ort,
                        "Edition": row.get("Edition", None),
                        "Land Rotreußen?": None,
                        "Lage": None,
                        "Region": None,
                        "Geodaten": None,
                        "alternative/historische Schreibweisen": None
                    })
    print(f"Nach Sheet '{sheet}' ist die Anzahl der Einträge: {len(result)}")

# ---------------------------
# 2. Tabelle: datasheets/180322_Identifikation_Objektorte3.xlsx
# ---------------------------
file2 = "datasheets/180322_Identifikation_Objektorte3.xlsx"
sheets2 = ["KIII", "WO", "WII"]

print("Starte Verarbeitung von Datei 2...")

for sheet in sheets2:
    try:
        df = pd.read_excel(file2, sheet_name=sheet, header=2)
        print(f"Sheet '{sheet}' in Datei 2 eingelesen. Gefundene Spalten: {df.columns.tolist()}")
    except Exception as e:
        print(f"Fehler beim Einlesen von {sheet} in {file2}: {e}")
        continue

    # Debug: Ausgabe der ersten paar Werte der Spalte "Objektort"
    print(f"Erste Werte in 'Objektort' in Sheet '{sheet}':")
    print(df["Objektort"].head())

    for index, row in df.iterrows():
        ort = row.get("Objektort", None)
        if pd.isna(ort):
            continue
        result.append({
            "Ort": ort,
            "Heutiger Name": row.get("Heutiger Name", None),
            "Edition": row.get("Edition", None),
            "Land Rotreußen?": row.get("Land Rotreußen?", None),
            "Lage": row.get("Lage", None),
            "Region": None,
            "Geodaten": row.get("Geodaten", None),
            "alternative/historische Schreibweisen": None
        })
    print(f"Nach Sheet '{sheet}' von Datei 2 ist die Anzahl der Einträge: {len(result)}")

# ---------------------------
# 3. Tabelle: datasheets/190817_Ortstrias-neu-GESAMT3.xlsx
# ---------------------------
file3 = "datasheets/190817_Ortstrias-neu-GESAMT3.xlsx"
sheets3 = ["KIII", "WO neu", "WII"]

print("Starte Verarbeitung von Datei 3...")

# (a) Erste Einlesung: Edition, Ausstellungsort (Heutiger Name) und Land Rotreußen?
for sheet in sheets3:
    try:
        df = pd.read_excel(file3, sheet_name=sheet, header=0)  # Daten beginnen in Zeile 2
        print(f"Sheet '{sheet}' (a) in Datei 3 eingelesen mit {df.shape[0]} Zeilen.")
    except Exception as e:
        print(f"Fehler beim Einlesen von {sheet} (a) in {file3}: {e}")
        continue

    # Je nach Mappe können die Spaltennamen variieren:
    if sheet in ["WO neu", "WII"]:
        edition_col = "Edition (grün, wenn im Text eingebaut)"
        ort_col = "Ausstellungsort (Heutiger Name)"
    else:
        edition_col = "Edition"
        ort_col = "Ausstellungsort"

    for index, row in df.iterrows():
        ort = row.get(ort_col, None)
        if pd.isna(ort):
            continue
        result.append({
            "Ort": ort,
            "Edition": row.get(edition_col, None),
            "Land Rotreußen?": row.get("Land Rotreußen?", None),
            "Lage": None,
            "Region": None,
            "Geodaten": None,
            "alternative/historische Schreibweisen": None
        })
    print(f"Nach Sheet '{sheet}' (a) von Datei 3 ist die Anzahl der Einträge: {len(result)}")

# (b) Zweite Einlesung: Objektort und Region (wobei "Objektort" wieder den Ort bezeichnet)
for sheet in sheets3:
    try:
        df = pd.read_excel(file3, sheet_name=sheet, header=0)
        print(f"Sheet '{sheet}' (b) in Datei 3 eingelesen mit {df.shape[0]} Zeilen.")
    except Exception as e:
        print(f"Fehler beim (zweiten) Einlesen von {sheet} in {file3}: {e}")
        continue

    for index, row in df.iterrows():
        ort = row.get("Objektort", None)
        if pd.isna(ort):
            continue
        result.append({
            "Ort": ort,
            "Edition": None,
            "Land Rotreußen?": None,
            "Lage": None,
            "Region": row.get("Region", None),
            "Geodaten": None,
            "alternative/historische Schreibweisen": None
        })
    print(f"Nach Sheet '{sheet}' (b) von Datei 3 ist die Anzahl der Einträge: {len(result)}")

# ---------------------------
# 4. Tabelle: Ortsdaten_Repertorium-aufbereitet.xlsx
# ---------------------------
file4 = "datasheets/Ortsdaten_Repertorium-aufbereitet.xlsx"
print("Starte Verarbeitung von Datei 4...")
try:
    # Lese nur die ersten 2344 Zeilen ein
    df = pd.read_excel(file4, nrows=2344)
    print(f"Datei 4 eingelesen mit {df.shape[0]} Zeilen und {df.shape[1]} Spalten (nur bis Zeile 2344).")
except Exception as e:
    print(f"Fehler beim Einlesen von {file4}: {e}")
else:
    # Es werden hier die Spalten "Ort", "Edition", "Region" und "alternative/historische Schreibweisen" vorausgesetzt
    df = df[["Ort", "Edition", "Region", "alternative/historische Schreibweisen"]]
    for index, row in df.iterrows():
        ort = row.get("Ort", None)
        if pd.isna(ort):
            continue
        result.append({
            "Ort": ort,
            "Edition": row.get("Edition", None),
            "Land Rotreußen?": None,
            "Lage": None,
            "Region": row.get("Region", None),
            "Geodaten": None,
            "alternative/historische Schreibweisen": row.get("alternative/historische Schreibweisen", None)
        })
    print(f"Nach Datei 4 ist die Gesamtanzahl der Einträge: {len(result)}")

# ---------------------------
# Alle Ergebnisse zu einem DataFrame zusammenführen
# ---------------------------
df_result = pd.DataFrame(result)
print("Erstellter DataFrame mit Spalten:", df_result.columns.tolist())
print("Erster Blick auf die Daten:")
print(df_result.head())

# Optional: Zeilen entfernen, bei denen "Ort" leer ist
df_result = df_result[df_result["Ort"].notna()]

# Ziel-Datei abspeichern
output_file = "datasheets/orte/Ortsliste.xlsx"
df_result.to_excel(output_file, index=False)
print(f"Die Zieltabelle wurde erfolgreich erstellt und unter {output_file} gespeichert.")
