import pandas as pd
import re

# Datei-Pfade
df_path = '../datasheets/Urkunden_Repertorium.xlsx'
registry_path = '../datasheets/orts_namen_register.xlsx'
output_path = '../datasheets/Urkunden_Repertorium_v1.xlsx'

# Registry mit Orten und Personen einlesen
reg_df = pd.read_excel(registry_path)
places = reg_df['Ort'].dropna().astype(str).unique().tolist()
persons = reg_df['Person'].dropna().astype(str).unique().tolist()

# Lowercase-Maps für Substring-Matches
place_map = {p.strip().lower(): p for p in places}
person_map = {p.strip().lower(): p for p in persons}

# Urkunden-Repertorium einlesen
urkunden_df = pd.read_excel(df_path)

# Spalten "Begünstigte" und "Objektorte" (wenn vorhanden) entfernen
urkunden_df = urkunden_df.drop(columns=['Begünstigte', 'Objektorte'], errors='ignore')

def find_non_overlapping(text, keys_map):
    """
    Sucht alle nicht-überlappenden Vorkommen aus keys_map in text und gibt
    die Original-Namen (Werte aus keys_map) in einer Liste zurück.
    """
    lower = text.lower()
    spans = []
    results = []
    # Sortiere nach Länge absteigend, damit längere Treffer Vorrang haben.
    for key, original in sorted(keys_map.items(), key=lambda x: len(x[0]), reverse=True):
        start = lower.find(key)
        if start != -1:
            end = start + len(key)
            # nur hinzufügen, wenn sich dieser Span mit keinem vorhandenen überschneidet
            if all(end <= s or start >= e for s, e in spans):
                results.append(original)
                spans.append((start, end))
    return results

def strip_issuer(text, issuer):
    """
    Entfernt (falls vorhanden) den Aussteller (issuer) aus dem Text,
    damit er später nicht versehentlich als "genannte Person" erkannt wird.
    """
    if not issuer or pd.isna(issuer):
        return text
    # \b stellt sicher, dass nur ganze Wörter abgeglichen werden
    pattern = rf'\b{re.escape(str(issuer))}\b'
    return re.sub(pattern, '', str(text), flags=re.IGNORECASE)

def match_beguenstigte(regest, issuer=None):
    """
    Findet alle in der Registry erfassten Personen im Regest-Text, 
    entfernt dabei zuerst den Aussteller (issuer), damit dieser nicht mitgezählt wird.
    """
    txt = strip_issuer(regest, issuer)
    # Kopie der Registry-Keys (lowercase) erstellen und Aussteller darin entfernen
    keys = person_map.copy()
    if issuer:
        # issuer in lowercase vergleichen mit den Werten aus person_map (ebenfalls original in person_map.values())
        issuer_lc = str(issuer).strip().lower()
        keys = {
            k: v 
            for k, v in keys.items() 
            if v.strip().lower() != issuer_lc
        }
    return find_non_overlapping(txt, keys)

# Die Spalte "Aussteller" existiert bereits in urkunden_df, daher passen wir direkt an:
# --------------------------------------------------------------------------------------
# 1. Wir prüfen, ob in urkunden_df bereits eine Spalte "Aussteller" vorhanden ist.
#    Wenn nicht, kann man sie hier manuell ausfüllen oder vorher hinzufügen.
# 2. Wir übernehmen einfach die vorhandenen Aussteller-Werte aus der Excel-Datei.
# --------------------------------------------------------------------------------------

# Prüfen, ob die Spalte "Aussteller" vorhanden ist:
if 'Aussteller' not in urkunden_df.columns:
    raise KeyError("Die Spalte 'Aussteller' fehlt in der Excel-Datei. Bitte vorher hinzufügen.")

# Jetzt extrahieren wir für jedes Regest die "genannten Personen" und "Orte",
# wobei der Aussteller (aus Spalte 'Aussteller') nicht in 'genannte Person' auftaucht.

# Temporäre Listen für rohe Ergebnisse
raw_places = []
raw_persons = []

for idx, row in urkunden_df.iterrows():
    regest_text = row['Regest']
    issuer = row['Aussteller']  # statt aus fixer Liste holen wir ihn hier direkt aus der Spalte

    # 1. Alle Personen (aus Registry) finden, ohne den Aussteller selbst
    persons_found = match_beguenstigte(regest_text, issuer)
    raw_persons.append(persons_found)

    # 2. Alle Orte (aus Registry) finden, ebenfalls ohne Aussteller
    txt_ohne_issuer = strip_issuer(regest_text, issuer)
    places_found = find_non_overlapping(txt_ohne_issuer, place_map)
    raw_places.append(places_found)

# In Series umwandeln, damit wir später filtern können
all_persons = pd.Series(raw_persons)
all_places  = pd.Series(raw_places)

# --------------------------------------------------------------------------------------
# Anpassung: Wir verzichten auf den Ausschluss von Orten, die in mehrteiligen Personennamen
# vorkommen. Somit sollen alle gefundenen Orte unverändert in 'Genannte Orte' übernommen werden.
# --------------------------------------------------------------------------------------
# Statt:
# filtered_places = []
# for places_list, persons_list in zip(all_places, all_persons):
#     exclude = []
#     for person in persons_list:
#         if len(person.split()) > 2:
#             # finde alle Orte, die im Personennamen stecken
#             exclude += find_non_overlapping(person, place_map)
#     exclude = set(exclude)
#     filtered = [p for p in places_list if p not in exclude]
#     filtered_places.append(filtered)

# Nun direkt übernehmen – jede Liste aus all_places bleibt unverändert:
filtered_places = all_places.tolist()

# Dynamisch Spalten für alle Treffer anlegen
max_places  = max((len(lst) for lst in filtered_places), default=0)
max_persons = max((len(lst) for lst in all_persons), default=0)

for i in range(1, max_places + 1):
    urkunden_df[f'Genannter Ort {i}'] = None
for i in range(1, max_persons + 1):
    urkunden_df[f'Genannte Person {i}'] = None

# Ergebnisse in die DataFrame-Spalten schreiben
for idx, _ in urkunden_df.iterrows():
    # Orte eintragen
    for i, ort in enumerate(filtered_places[idx], start=1):
        urkunden_df.at[idx, f'Genannter Ort {i}'] = ort
    # Personen eintragen
    for i, person in enumerate(all_persons.iloc[idx], start=1):
        urkunden_df.at[idx, f'Genannte Person {i}'] = person

# Speichern
urkunden_df.to_excel(output_path, index=False)
print(f"Fertig: Ausgabe in {output_path}")
