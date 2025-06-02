import pandas as pd
import re

# Liste der Aussteller
aussteller_list = [
    "Kazimierz III.",
    "Ełżbieta Łokietkówna",
    "Władysław Opolczyk",
    "Ludwig von Anjou",
    "Maria von Ungarn",
    "Elisabeth von Bosnien",
    "Władysław II. Jagiełło",
    "Jadwiga von Anjou",
    "Dmytro Det’ko",
    "Otto von Pilica",
    "Johannes Kmita",
    "Andreas von Barabás",
    "Johannes von Sprowa",
    "Emerik Bebek",
    "Johannes von Tarnów",
    "Gnewosius von Dalewice",
    "Florian von Korytnica",
    "Ivan von Obiechów",
    "Spytek von Tarnów",
    "Petrus von Charbinowice",
    "Johannes Mężyk von Dąbrowa",
    "Vincentius von Szamotuły"
]
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

def find_non_overlapping(text, keys_map):
    lower = text.lower()
    spans = []
    results = []
    for key, original in sorted(keys_map.items(), key=lambda x: len(x[0]), reverse=True):
        start = lower.find(key)
        if start != -1:
            end = start + len(key)
            if all(end <= s or start >= e for s, e in spans):
                results.append(original)
                spans.append((start, end))
    return results

def strip_issuer(text, issuer):
    """Entfernt den Aussteller aus dem Text, falls vorhanden."""
    if not issuer:
        return text
    pattern = rf'\b{re.escape(issuer)}\b'
    return re.sub(pattern, '', text, flags=re.IGNORECASE)

def match_aussteller(regest):
    txt = regest.lower()
    for name in aussteller_list:
        if name.lower() in txt:
            return name
    return None

def match_beguenstigte(regest, issuer=None):
    txt = strip_issuer(str(regest), issuer)
    # aus person_map kopieren und Aussteller herausfiltern
    keys = person_map.copy()
    if issuer:
        keys = {k: v for k, v in keys.items() if v.lower() != issuer.lower()}
    return find_non_overlapping(txt, keys)

# Extraktion Aussteller
urkunden_df['Aussteller'] = urkunden_df['Regest'].astype(str).apply(match_aussteller)

# Zunächst alle Roh-Listen ermitteln
raw_places = []
raw_persons = []
for idx, row in urkunden_df.iterrows():
    text = row['Regest']
    issuer = row['Aussteller']
    # erst Aussteller raus, dann Personen
    persons = match_beguenstigte(text, issuer)
    raw_persons.append(persons)
    # dann Orte (ohne Aussteller)
    txt_ohne_issuer = strip_issuer(text, issuer)
    places = find_non_overlapping(txt_ohne_issuer, place_map)
    raw_places.append(places)

all_persons = pd.Series(raw_persons)
all_places  = pd.Series(raw_places)

# Filter: bei mehrteiligen Personennamen Orte des Namens entfernen
filtered_places = []
for places, persons in zip(all_places, all_persons):
    exclude = []
    for person in persons:
        if len(person.split()) > 2:
            # finde alle Orte, die im Personennamen stecken
            exclude += find_non_overlapping(person, place_map)
    # entferne Duplikate und filtere
    exclude = set(exclude)
    filtered = [p for p in places if p not in exclude]
    filtered_places.append(filtered)

# Dynamisch Spalten für alle Treffer anlegen
max_places  = max(len(lst) for lst in filtered_places) if filtered_places else 0
max_persons = max(len(lst) for lst in all_persons)    if not all_persons.empty else 0

for i in range(1, max_places + 1):
    urkunden_df[f'Objektort {i}'] = None
for i in range(1, max_persons + 1):
    urkunden_df[f'Begünstigte Person {i}'] = None

# Ergebnisse eintragen
for idx, row in urkunden_df.iterrows():
    for i, ort in enumerate(filtered_places[idx], start=1):
        urkunden_df.at[idx, f'Objektort {i}'] = ort
    for i, person in enumerate(all_persons.iloc[idx], start=1):
        urkunden_df.at[idx, f'Begünstigte Person {i}'] = person

# Speichern
urkunden_df.to_excel(output_path, index=False)
print(f"Fertig: Ausgabe in {output_path}")