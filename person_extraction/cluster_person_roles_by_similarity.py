from tqdm import tqdm
import pandas as pd
import difflib

# === Thresholds (einfach hier ändern) ===
NAME_THRESHOLD = 0.86
FUNC_THRESHOLD_MULTIWORD = 0.9   # wenn Funktion aus >1 Wort besteht
FUNC_THRESHOLD_SINGLEWORD = 0.85  # wenn Funktion aus 1 Wort besteht

# === Fuzzy-Matching-Funktion ===
try:
    from rapidfuzz import fuzz
    def compute_ratio(a: str, b: str) -> float:
        return fuzz.ratio(a, b) / 100
except ImportError:
    def compute_ratio(a: str, b: str) -> float:
        return difflib.SequenceMatcher(None, a, b).ratio()

# 1. Excel einlesen (Aussteller ergänzt)
df = pd.read_excel(
    '../datasheets/personen/Personen_Empfänger_Austeller.xlsx',
    usecols=["Name", "Funktion", "Rolle in Urkunde", "Geschlecht", "Edition", "Datum Funktion", "Aussteller"]
)

# 2. Fehlende Werte und lowercase (für Aussteller analog)
for col in ["Name", "Funktion", "Aussteller"]:
    df[col] = df[col].fillna("").astype(str)
df["name_lc"]       = df["Name"].str.lower().str.strip()
df["funktion_lc"]   = df["Funktion"].str.lower().str.strip()
df["aussteller_lc"] = df["Aussteller"].str.lower().str.strip()

# 3. Clustering
clusters = []
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Clustering Namen"):
    name         = row["name_lc"]
    funktion     = row["funktion_lc"]
    aussteller   = row["aussteller_lc"]
    name_words   = name.split()
    assigned     = False

    for cluster in clusters:
        rep_name       = cluster["rep_name"]
        rep_funktion   = cluster["rep_funktion"]
        # rep_aussteller ist hier nicht nötig fürs Matching
        name_sim       = compute_ratio(name, rep_name)

        if len(name_words) > 1:
            # Mehr-Wort-Name → nur Name matchen
            if name_sim >= NAME_THRESHOLD:
                cluster["members"].append(idx)
                assigned = True
                break
        else:
            # Ein-Wort-Name → zusätzlich Funktion
            if not funktion or not rep_funktion:
                continue
            funk_words  = funktion.split()
            func_thresh = (FUNC_THRESHOLD_MULTIWORD
                           if len(funk_words) > 1
                           else FUNC_THRESHOLD_SINGLEWORD)
            func_sim    = compute_ratio(funktion, rep_funktion)
            if name_sim >= NAME_THRESHOLD and func_sim >= func_thresh:
                cluster["members"].append(idx)
                assigned = True
                break

    if not assigned:
        clusters.append({
            "rep_name":       name,
            "rep_funktion":   funktion,
            "rep_index":      idx,
            # rep_aussteller optional, aber wir speichern's der Vollständigkeit halber
            "rep_aussteller": aussteller,
            "members":        [idx]
        })

# 4. Aggregierte DataFrame aufbauen
rows = []
max_cluster_size    = max(len(c["members"]) for c in clusters)
max_name_variants   = 0

# Max. Anzahl Name-Varianten ermitteln
for cluster in clusters:
    variants = {df.at[m, "Name"] for m in cluster["members"]}
    max_name_variants = max(max_name_variants, len(variants))

for cluster in tqdm(clusters, desc="Aufbauen aggregierte Zeilen"):
    members  = cluster["members"]
    variants = list(dict.fromkeys(df.at[m, "Name"] for m in members))
    
    row = {"Häufigkeit": len(members)}
    # Name-Varianten-Spalten
    for i, variant in enumerate(variants, start=1):
        row[f"Name {i}"] = variant
    for i in range(len(variants)+1, max_name_variants+1):
        row[f"Name {i}"] = ""
    
    # Funktions-, Aussteller- und Metadaten-Spalten
    for i, m in enumerate(members, start=1):
        row[f"Funktion {i}"]         = df.at[m, "Funktion"]
        row[f"Aussteller {i}"]       = df.at[m, "Aussteller"]
        row[f"Datum Funktion {i}"]   = df.at[m, "Datum Funktion"]
        row[f"Edition {i}"]          = df.at[m, "Edition"]
        row[f"Rolle in Urkunde {i}"] = df.at[m, "Rolle in Urkunde"]
        row[f"Geschlecht {i}"]       = df.at[m, "Geschlecht"]
    rows.append(row)

# Spalten-Reihenfolge zusammenstellen (Aussteller-Spalten ergänzen)
columns = ["Häufigkeit"]
columns += [f"Name {i}" for i in range(1, max_name_variants+1)]
for i in range(1, max_cluster_size+1):
    columns += [
        f"Funktion {i}",
        f"Aussteller {i}",
        f"Datum Funktion {i}",
        f"Edition {i}",
        f"Rolle in Urkunde {i}",
        f"Geschlecht {i}"
    ]

agg_df = pd.DataFrame(rows, columns=columns)

# 5. Ergebnis speichern
agg_df.to_excel(
    "../datasheets/personen/Personen_Empfaenger_Aussteller_aggregiert.xlsx",
    index=False
)
print("Fertig! Aggregierte Tabelle inkl. Aussteller in 'Personen_Empfaenger_Aussteller_aggregiert.xlsx'.")
