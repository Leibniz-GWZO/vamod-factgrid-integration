import pandas as pd
import difflib
from tqdm import tqdm
from functools import lru_cache
import re

# --- Zentrale Thresholds ---
NAME_THRESHOLD = 0.88
FUNCTION_THRESHOLD = 0.85
STRONG_FUNCTION_THRESHOLD = 0.96

# ------------------------------------------------------
# 1) HILFSFUNKTIONEN MIT CACHING
# ------------------------------------------------------
@lru_cache(maxsize=None)
def remove_diacritics(s: str) -> str:
    s = str(s).lower().strip()
    replacements = {
        'ó': 'o', 'ö': 'o', 'ò': 'o', 'ő': 'o',
        'ä': 'a', 'á': 'a', 'à': 'a', 'ă': 'a', 'ą': 'a',
        'é': 'e', 'ë': 'e', 'ě': 'e', 'è': 'e',
        'ü': 'u', 'ű': 'u', 'ú': 'u', 'ù': 'u',
        'ć': 'c', 'č': 'c', 'ç': 'c',
        'ł': 'l', 'ń': 'n', 'ñ': 'n', 'ß': 'ss',
        'ź': 'z', 'ż': 'z', 'ž': 'z'
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    return s

@lru_cache(maxsize=None)
def fuzzy_name_similarity(name_a: str, name_b: str) -> float:
    a = remove_diacritics(name_a.strip())
    b = remove_diacritics(name_b.strip())
    return difflib.SequenceMatcher(None, a, b).ratio()

def names_are_similar(name_a: str, name_b: str, threshold: float = NAME_THRESHOLD) -> bool:
    a = remove_diacritics(name_a.strip())
    b = remove_diacritics(name_b.strip())
    if {a, b} == {"iohannes", "johannes"}:
        return True
    return fuzzy_name_similarity(name_a, name_b) >= threshold

def split_functions(func_str: str) -> list:
    pattern = r"\s*(?:,|;|\+|/)\s*"
    functions = re.split(pattern, func_str.strip())
    return [f.strip() for f in functions if f.strip()]

def normalize_function_for_matching(f: str) -> str:
    norm = f.strip()
    if norm.endswith("ensis"):
        norm = norm[:-len("ensis")].strip()
    return norm

# ------------------------------------------------------
# 2) DATEN EINLESEN UND SPALTEN BEREINIGEN
# ------------------------------------------------------
try:
    # Excel-Datei einlesen ohne dtype=str und mit parse_dates
    df = pd.read_excel('datasheets/Personenliste.xlsx', parse_dates=['früheste_nennung', 'späteste_nennung'])
    # Fülle nur die Spalten, die keine Datumsspalten sind, mit leeren Strings
    non_date_columns = [col for col in df.columns if col not in ['früheste_nennung', 'späteste_nennung']]
    df[non_date_columns] = df[non_date_columns].fillna("")
    df.columns = df.columns.str.replace(" ", "_").str.lower()
except Exception as e:
    print("Fehler beim Einlesen der Datei:", e)
    exit()

# ------------------------------------------------------
# 3) CLUSTERING / ZUSAMMENFÜHREN
# ------------------------------------------------------
clusters = []
cluster_id = 0

for row in tqdm(df.itertuples(index=False), total=len(df)):
    person_name = row.person.strip()
    raw_funktion = row.funktion.strip()
    row_functions = split_functions(raw_funktion) if raw_funktion else []
    
    geschlecht = row.geschlecht.strip() if row.geschlecht else ""
    rolle = row.rolle_in_urkunde.strip() if row.rolle_in_urkunde else ""
    try:
        genannt = int(row.genannte_häufigkeit)
    except:
        genannt = 1  # fallback

    # Extrahiere die Datumswerte
    row_earliest = row.früheste_nennung if pd.notnull(row.früheste_nennung) else None
    row_latest = row.späteste_nennung if pd.notnull(row.späteste_nennung) else None

    found_cluster = None

    # Suche in den existierenden Clustern
    for cluster in clusters:
        if cluster["geschlecht"] and geschlecht and cluster["geschlecht"] != geschlecht:
            continue

        name_match = any(names_are_similar(person_name, existing_name) for existing_name in cluster["alle_namen"])
        
        compatibility_scores = []
        strong_match_found = False
        for row_func in row_functions:
            for func_entry in cluster["funktionen"]:
                existing_func = func_entry["funktion"]
                norm_row_func = normalize_function_for_matching(row_func)
                norm_existing_func = normalize_function_for_matching(existing_func)
                score = difflib.SequenceMatcher(
                    None,
                    remove_diacritics(norm_row_func),
                    remove_diacritics(norm_existing_func)
                ).ratio()
                compatibility_scores.append(score)
                # Neue Bedingung: Wenn beide Funktionen mit "capitaneus" beginnen, als starker Match werten
                # Bestehende Bedingung für "capitaneus" angepasst:
                if (row_func.lower().startswith("capitaneus") and 
                    existing_func.lower().startswith("capitaneus") and 
                    score >= STRONG_FUNCTION_THRESHOLD):
                    strong_match_found = True

                # Bestehende Bedingung für "ensis"
                if (len(row_func.split()) >= 2 and row_func.split()[-1].endswith("ensis") and
                    len(existing_func.split()) >= 2 and existing_func.split()[-1].endswith("ensis") and
                    score >= STRONG_FUNCTION_THRESHOLD):
                    strong_match_found = True

        if compatibility_scores:
            if strong_match_found:
                func_match = True
            else:
                compatible_count = sum(1 for score in compatibility_scores if score >= FUNCTION_THRESHOLD)
                func_match = (compatible_count / len(compatibility_scores)) >= 0.5
        else:
            func_match = False

        if len(person_name.split()) > 1:
            match_condition = name_match
        else:
            match_condition = name_match and func_match

        if match_condition:
            found_cluster = cluster
            break

    if found_cluster:
        if person_name not in found_cluster["alle_namen"]:
            found_cluster["alle_namen"].append(person_name)
        for row_func in row_functions:
            found_func_entry = None
            for func_entry in found_cluster["funktionen"]:
                if normalize_function_for_matching(row_func) == normalize_function_for_matching(func_entry["funktion"]):
                    found_func_entry = func_entry
                    break
            if found_func_entry is not None:
                if row_earliest is not None:
                    if (found_func_entry["früheste_nennung"] is None) or (row_earliest <= found_func_entry["früheste_nennung"]):
                        found_func_entry["früheste_nennung"] = row_earliest
                if row_latest is not None:
                    if (found_func_entry["späteste_nennung"] is None) or (row_latest > found_func_entry["späteste_nennung"]):
                        found_func_entry["späteste_nennung"] = row_latest
            else:
                found_cluster["funktionen"].append({
                    "funktion": row_func,
                    "früheste_nennung": row_earliest,
                    "späteste_nennung": row_latest
                })
        if rolle and rolle not in found_cluster["rollen_in_urkunden"]:
            found_cluster["rollen_in_urkunden"].append(rolle)
        found_cluster["summe_genannt"] += genannt
    else:
        cluster_id += 1
        new_cluster = {
            "cluster_id": cluster_id,
            "hauptname": person_name,
            "alle_namen": [person_name],
            "funktionen": [],
            "geschlecht": geschlecht,
            "rollen_in_urkunden": [rolle] if rolle else [],
            "summe_genannt": genannt
        }
        for row_func in row_functions:
            new_cluster["funktionen"].append({
                "funktion": row_func,
                "früheste_nennung": row_earliest,
                "späteste_nennung": row_latest
            })
        clusters.append(new_cluster)

# ------------------------------------------------------
# 4) ERGEBNISSE ALS DICTIONARY
# ------------------------------------------------------
cluster_dict = {}
for c in clusters:
    cid = c["cluster_id"]
    cluster_dict[cid] = {
        "summe_genannt": c["summe_genannt"],
        "hauptname": c["hauptname"],
        "alle_namen": c["alle_namen"],
        "geschlecht": c["geschlecht"],
        "funktionen": c["funktionen"],
        "rollen_in_urkunden": c["rollen_in_urkunden"]
    }

# ------------------------------------------------------
# 5) AUSGABE ALS CSV - ERWEITERTE FORMATIERUNG
# ------------------------------------------------------
max_names = max(len(v['alle_namen']) for v in cluster_dict.values())
max_funcs = max(len(v['funktionen']) for v in cluster_dict.values())
max_roles = max(len(v['rollen_in_urkunden']) for v in cluster_dict.values())

rows = []
for cid, data in cluster_dict.items():
    # "Häufigkeit" als erste Spalte und "hauptname" separat
    row = {
        "Häufigkeit": data["summe_genannt"],
        "hauptname": data["hauptname"],
    }
    # Zusätzliche Namen (überspringe den Hauptnamen, der bereits separat ausgegeben wird)
    for i in range(1, len(data["alle_namen"])):
        row[f"Name {i+1}"] = data["alle_namen"][i]
    # Falls es in anderen Clustern mehr Namen gibt, auffüllen
    for i in range(len(data["alle_namen"]), max_names):
        row[f"Name {i+1}"] = ""
    # Funktionen-Spalten
    for i in range(max_funcs):
        if i < len(data["funktionen"]):
            func_entry = data["funktionen"][i]
            row[f"Funktion {i+1}"] = func_entry["funktion"]
            row[f"früheste_nennung Funktion {i+1}"] = func_entry["früheste_nennung"]
            row[f"späteste_nennung Funktion {i+1}"] = func_entry["späteste_nennung"]
        else:
            row[f"Funktion {i+1}"] = ""
            row[f"früheste_nennung Funktion {i+1}"] = ""
            row[f"späteste_nennung Funktion {i+1}"] = ""
    # Rollen-Spalten
    for i in range(max_roles):
        row[f"Rolle {i+1}"] = data["rollen_in_urkunden"][i] if i < len(data["rollen_in_urkunden"]) else ""
    # "geschlecht" als letzte Spalte hinzufügen
    row["geschlecht"] = data["geschlecht"]
    rows.append(row)

output_df = pd.DataFrame(rows)

# ------------------------------------------------------
# 6) SPEICHERN ALS EXCEL-DATEI (.xlsx)
# ------------------------------------------------------
excel_datei = 'datasheets/Personen_Matching.xlsx'
output_df.to_excel(excel_datei, index=False)
print(f"Die aggregierten Ergebnisse wurden als Excel-Datei gespeichert: {excel_datei}")
