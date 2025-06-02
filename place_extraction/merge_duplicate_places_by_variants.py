import pandas as pd
import numpy as np

# ----- 1. Datei einlesen und vorbereiten -----
file_path = "datasheets/orte/Orte_Identifikation.xlsx"
df = pd.read_excel(file_path)

# NaN-Werte durch leere Strings ersetzen und alle Werte als Strings behandeln
df.fillna("", inplace=True)
df = df.astype(str)

n = len(df)  # Gesamtanzahl Zeilen

# ----- 2. Union-Find-Struktur zur Gruppierung -----
# Jede Zeile beginnt in einer eigenen Gruppe.
parent = list(range(n))

def find(i):
    if parent[i] != i:
        parent[i] = find(parent[i])
    return parent[i]

def union(i, j):
    ri, rj = find(i), find(j)
    if ri != rj:
        parent[rj] = ri

# Hilfsfunktion: Zerteilt einen Zelleninhalt an Semikolon und entfernt Leerzeichen.
def split_alt(value):
    return [x.strip() for x in value.split(";") if x.strip()]

# Vergleiche alle Zeilenpaare. Falls:
#   - Der Wert in "Schreibweise Ortsregister" einer Zeile
#     in der (gesplitteten) Zelle "alternative/historische Schreibweisen" der anderen Zeile vorkommt,
#   - oder umgekehrt,
# vereinige die Zeilen.
for i in range(n):
    ort_i = df.iloc[i]["Schreibweise Ortsregister"].strip()
    alt_i = split_alt(df.iloc[i]["alternative/historische Schreibweisen"])
    for j in range(i + 1, n):
        ort_j = df.iloc[j]["Schreibweise Ortsregister"].strip()
        alt_j = split_alt(df.iloc[j]["alternative/historische Schreibweisen"])
        if (ort_i and ort_i in alt_j) or (ort_j and ort_j in alt_i):
            union(i, j)

# Gruppiere Zeilen anhand ihres repräsentativen Eltern-Knotens
groups = {}
for i in range(n):
    rep = find(i)
    groups.setdefault(rep, []).append(i)

# ----- 3. Funktion zum Mergen von Zelleninhalten für Spalten ohne Mehrfachwerte -----
def merge_values(values):
    """
    Nimmt eine Liste von Strings (z. B. alle Werte einer Spalte aus den zu mergeenden Zeilen),
    entfernt leere Strings und Duplikate und führt sie per "; " zusammen.
    """
    s = set()
    for v in values:
        v_clean = v.strip()
        if v_clean and v_clean.lower() != "nan":
            s.add(v_clean)
    if not s:
        return ""
    else:
        return "; ".join(sorted(s))

# ----- 4. Merge der Gruppen über alle Spalten -----
merged_rows = []
print("Folgende Merge-Gruppen wurden gefunden:")

for rep, indices in groups.items():
    # Falls nur eine Zeile in der Gruppe ist, übernehmen wir sie direkt.
    if len(indices) == 1:
        merged_rows.append(df.iloc[indices[0]].to_dict())
    else:
        print("----- Gruppe (Zeilen: {}) -----".format(indices))
        for i in indices:
            print(f"Zeile {i}: {df.iloc[i].to_dict()}")
        
        # Wähle als Basis die Zeile mit den meisten nicht-leeren Feldern.
        best_idx = max(indices, key=lambda i: sum(1 for cell in df.loc[i] if str(cell).strip() != ""))
        base_row = df.loc[best_idx].copy()

        merged_row = {}
        for col in df.columns:
            if col == "Schreibweise Ortsregister":
                # Hier nehmen wir ausschließlich den Basiswert (ohne Merge)
                merged_row[col] = base_row[col].strip()
            elif col == "alternative/historische Schreibweisen":
                # Für die alternative Schreibweise: 
                #   - Zerteile alle Zellen (in dieser Spalte) der Gruppe an "; " 
                #   - Füge zusätzlich alle Werte der Spalte "Schreibweise Ortsregister" hinzu.
                s = set()
                for i in indices:
                    # Aus der Spalte "alternative/historische Schreibweisen":
                    for part in split_alt(df.iloc[i][col]):
                        s.add(part)
                    # Aus der Spalte "Schreibweise Ortsregister":
                    ort_val = df.iloc[i]["Schreibweise Ortsregister"].strip()
                    if ort_val:
                        s.add(ort_val)
                # Entferne den Basiswert aus "Schreibweise Ortsregister", falls er dabei ist
                base_val = base_row["Schreibweise Ortsregister"].strip()
                if base_val in s:
                    s.remove(base_val)
                merged_alt = "; ".join(sorted(s))
                merged_row[col] = merged_alt
            else:
                # Für alle anderen Spalten: Wir vereinigen die vorhandenen Werte
                merged_row[col] = merge_values([df.iloc[i][col] for i in indices])
        
        print("→ Gemergte Zeile:", merged_row, "\n")
        merged_rows.append(merged_row)

# Neuer DataFrame mit den gemergten Zeilen
merged_df = pd.DataFrame(merged_rows)

# ----- 5. Ergebnis speichern -----
output_path = "datasheets/orte/Orte_Identifikation_merged.xlsx"
merged_df.to_excel(output_path, index=False)

print(f"\nMerge abgeschlossen. Die gemergte Datei wurde gespeichert: {output_path}")
