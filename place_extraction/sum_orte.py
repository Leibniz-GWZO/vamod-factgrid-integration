import pandas as pd
import unicodedata

def normalize_ort(ort):
    """
    Normalisiert den Ortsnamen:
    - Entfernt führende und abschließende Leerzeichen.
    - Wandelt in Kleinbuchstaben um.
    - Entfernt diakritische Zeichen.
    """
    if pd.isna(ort):
        return None
    ort_str = str(ort).strip().lower()
    ort_norm = unicodedata.normalize('NFKD', ort_str)
    ort_norm = "".join([c for c in ort_norm if not unicodedata.combining(c)])
    return ort_norm

# Lese die erstellte Zieltabelle ein
input_file = "datasheets/orte/Ortsliste.xlsx"
df = pd.read_excel(input_file)

# Entferne Zeilen ohne Ortsnamen
df = df[df["Ort"].notna()]

# Füge eine Normalisierungsspalte hinzu, die führende/trailing Whitespaces entfernt,
# diakritische Zeichen ignoriert und in Kleinbuchstaben umwandelt.
df["Ort_norm"] = df["Ort"].apply(normalize_ort)

# Aggregation: Gruppiere nach der normalisierten Version der Ortsnamen
consolidated_rows = []
grouped = df.groupby("Ort_norm")

for norm, group in grouped:
    group = group.copy()
    # Zähle die nicht-leeren Felder (außer "Ort" und "Ort_norm")
    group['non_na_count'] = group.drop(columns=["Ort", "Ort_norm"]).notna().sum(axis=1)
    # Wähle die Zeile mit den meisten Informationen als Basis
    base_row = group.loc[group['non_na_count'].idxmax()]
    
    # Speichere den Original-Ortsnamen aus der Basiszeile in der finalen Tabelle
    consolidated = {"Ort": base_row["Ort"]}
    
    # Für alle anderen Spalten (außer "Ort" und "Ort_norm")
    for col in df.columns:
        if col in ["Ort", "Ort_norm"]:
            continue
        
        values = group[col].dropna().unique().tolist()
        
        # Falls der Basiswert vorhanden ist, diesen an den Anfang setzen
        base_val = base_row[col]
        if pd.notna(base_val) and base_val in values:
            values.remove(base_val)
            values.insert(0, base_val)
        
        # Erzeuge immer nummerierte Spalten – auch wenn nur ein Wert vorliegt
        if len(values) == 0:
            consolidated[f"{col} 1"] = None
        else:
            for i, val in enumerate(values, start=1):
                consolidated[f"{col} {i}"] = val

    consolidated_rows.append(consolidated)

df_consolidated = pd.DataFrame(consolidated_rows)

# Umordnung der Spalten:
# - "Ort" steht immer an erster Stelle.
# - Alle Spalten, deren Name mit "Edition" beginnt, sollen am Ende stehen.
# - Die übrigen nummerierten Spalten kommen davor.
other_cols = [col for col in df_consolidated.columns if col != "Ort" and not col.startswith("Edition")]
edition_cols = [col for col in df_consolidated.columns if col.startswith("Edition")]

def edition_sort_key(col):
    parts = col.split(" ")
    try:
        return int(parts[1])
    except (IndexError, ValueError):
        return 0

edition_cols = sorted(edition_cols, key=edition_sort_key)

new_order = ["Ort"] + other_cols + edition_cols
df_consolidated = df_consolidated[new_order]

# Speichere den konsolidierten DataFrame in eine neue Excel-Datei
output_file = "datasheets/orte/Ortsliste_aggregated_sorted.xlsx"
df_consolidated.to_excel(output_file, index=False)
print(f"Die konsolidierte und sortierte Zieltabelle wurde erfolgreich erstellt und unter {output_file} gespeichert.")
