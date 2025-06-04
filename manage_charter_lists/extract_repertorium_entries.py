import re
import pandas as pd
import pdfplumber
from datetime import datetime

# Pfade zu PDF und Excel
PDF_PATH = "../datasheets/Repertorium.pdf"
OUTPUT_XLSX = "../datasheets/Urkunden_Repertorium.xlsx"

# Mapping deutscher Monatsabkürzungen auf Zahlen
MONTH_MAP = {
    'Jan.': 1, 'Feb.': 2, 'März': 3, 'Apr.': 4, 'Mai': 5,
    'Juni': 6, 'Juli': 7, 'Jul.': 7, 'Aug.': 8, 'Sep.': 9,
    'Okt.': 10, 'Nov.': 11, 'Dez.': 12
}

# Liste möglicher Aussteller (in der Reihenfolge, wie geprüft werden soll)
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

# Spalten der Ausgabetabelle
COLUMNS = [
    'Nr. Rep', 'Datum', 'Aussteller', 'Ausstellungsort',
    'Datumssicherheit', 'Sicherheit Status Urkunde', 'Begünstigte',
    'Objektorte', 'Regest', 'Angaben zur Überlieferung',
    'Angaben zu Drucken/Editionen', 'Kommentar'
]

records = []

# Regulärer Ausdruck für den Einstieg jeder Urkunde (inklusive "Nr. Nr.")
entry_pattern = re.compile(
    r'^(Nr\.\s+(?:Nr\.\s+)?[AB]\d+)'                  # Nr. Rep oder Nr. Nr. Rep
    r'(?:\s+(?:Fälschung|Verfälscht)(?:\s*(?:\(\?\)|\?))?)?'  # optional Kennzeichnung
    r'\s+([^,]+),',                                          # Ausstellungsort bis Komma
    re.IGNORECASE
)

# Muster zum Überspringen unerwünschter Kopfzeilen und reiner Zahlen
skip_pattern1 = re.compile(r'^A – Die Herrscherinnen und Herrscher \| \d+$')
skip_pattern2 = re.compile(r'^\d+ \| Repertorium diplomatum terrae Russie Regni Poloniae$')
skip_pattern3 = re.compile(r'^B – Urkunden des Capitaneus Russie \| \d+$')

# PDF öffnen und Textzeilen sammeln
with pdfplumber.open(PDF_PATH) as pdf:
    # Seite 63 gefiltert ausgeben
    if len(pdf.pages) >= 63:
        text63 = pdf.pages[62].extract_text() or ''
        lines63 = [ln.strip() for ln in text63.splitlines()]
        filtered63 = [ln for ln in lines63 if not (
            skip_pattern1.match(ln)
            or skip_pattern2.match(ln)
            or skip_pattern3.match(ln)
            or re.fullmatch(r'\d{4}', ln)
        )]
        print("--- Seite 63 nach Bereinigung ---")
        for ln in filtered63:
            print(ln)
    # Alle Seiten sammeln
    lines = []
    for page in pdf.pages:
        text = page.extract_text() or ''
        lines.extend([ln.strip() for ln in text.splitlines()])

# Unerwünschte Linien entfernen (Kopfzeilen und reine vierstellige Zahlen)
lines = [ln for ln in lines if not (
    skip_pattern1.match(ln)
    or skip_pattern2.match(ln)
    or skip_pattern3.match(ln)
    or re.fullmatch(r'\d{4}', ln)
)]

# Durchlauf durch Zeilen
i = 0
while i < len(lines):
    line = lines[i]
    m = entry_pattern.match(line)
    if not m:
        i += 1
        continue

    # Nummer und Ort
    nr_rep = m.group(1)
    if nr_rep.lower().startswith('dep.'):
        i += 1
        continue
    ausstellungsort = m.group(2).strip()

    # Datumsteil nach Komma
    parts = line.split(',', 1)
    datum_raw = parts[1].strip() if len(parts) > 1 else ''
    status_urkunde = 'unsicher' if re.search(r'fälsch', line, re.IGNORECASE) else 'sicher'

    # Datum parsen
    date_str = ''
    date_sicher = 'unsicher'
    try:
        dparts = datum_raw.split()
        if len(dparts) == 3:
            y, mon, d = dparts
            mon_val = MONTH_MAP.get(mon)
            if mon_val:
                date_str = datetime(int(y), mon_val, int(d)).strftime('%Y-%m-%d')
                date_sicher = 'sicher'
            else:
                date_str = datum_raw
        else:
            date_str = datum_raw
    except:
        date_str = datum_raw

    # Regest sammeln (mit Hyphenations-Korrektur)
    regest_lines = []
    j = i + 1
    while j < len(lines):
        tl = lines[j]
        # Abbruch, wenn nächster Eintrag oder Überlieferungs­start erreicht
        if tl.startswith(('Kopial', 'Original', 'Transsumpt', 'Verlorenes Original')):
            break
        if re.search(r'\(T.{0,10}\)\.$', tl):
            # Letzte Regest­zeile, anhängen und beenden
            if regest_lines and regest_lines[-1].endswith('-'):
                # vorhergehende Zeile ohne '-' + komplette Zeile
                regest_lines[-1] = regest_lines[-1][:-1] + tl.lstrip()
            else:
                regest_lines.append(tl)
            j += 1
            break

        if tl:
            if regest_lines and regest_lines[-1].endswith('-'):
                # letzte Zeile endet auf Bindestrich → Wortbruch aufheben
                # Bindestrich entfernen und direkt mit dem Beginn von tl verbinden
                # tl.lstrip() entfernt führende Leerzeichen
                regest_lines[-1] = regest_lines[-1][:-1] + tl.lstrip()
            else:
                regest_lines.append(tl)
        j += 1

    regest_text = ' '.join(regest_lines)

    # Aussteller aus dem Regest extrahieren (der zuerst genannte in der Text‐Reihenfolge)
    aussteller = ''
    min_pos = len(regest_text) + 1
    for name in aussteller_list:
        pos = regest_text.find(name)
        if pos != -1 and pos < min_pos:
            min_pos = pos
            aussteller = name

    # Überlieferung sammeln
    ueber_lines = []
    if j < len(lines) and lines[j].startswith(('Kopial', 'Original', 'Transsumpt', 'Verlorenes Original')):
        k = j
        while k < len(lines):
            curr = lines[k]
            if (
                entry_pattern.match(curr)
                or re.match(r'^[A-ZÄÖÜ]{2}', curr)
                or curr.startswith((
                    'Ungedruckt','ungedruckt','Dokumenty','Dodatek',
                    'Bei der Revision','Oleś','Dodatki','||','ArchSang',
                    'Aneks.','Codex.','Szyszka','Prochaska','Tęgowski',
                    'MatHist','Fragmenty'
                ))
            ):
                break
            ueber_lines.append(curr)
            k += 1
        j = k
    ueber_text = ' '.join(ueber_lines)

    # Angaben zu Drucken/Editionen sammeln
    edition_markers = [
        'Die','Vgl.','Vgl','Dieser','Höchstwarscheinlich','Höchstwahrscheinlich','Das','Gąsiorowski,',
        'vgl.','Laut','Szyszka','Kurtyka','Der','Zeitgleich','Kapral’','Tylus','Sperka',
        'Ein Stanislaus von','Datum','Als Fälschung','Zur Datierung','Hruševs',
        'In','Zum','Aufgrund','Sowohl','Diese','Bereits','Siehe','Deperdita','Klebowicz','Červone',
        'Nr. Nr.','Abraham','Beide','Starunja','Rutkowska','Bei','Jarosław','Wilamowski',
        'Im','Auf','Prochaska','Ebenfalls'
    ]
    edition_lines = []
    if j < len(lines):
        k = j
        while k < len(lines):
            curr = lines[k]
            if entry_pattern.match(curr) or any(curr.startswith(m) for m in edition_markers) or curr.startswith('||'):
                break
            edition_lines.append(curr)
            k += 1
        j = k
    edition_text = ' '.join(edition_lines)

    # Kommentar sammeln
    comment_lines = []
    if j < len(lines):
        k = j
        while k < len(lines):
            curr = lines[k]
            if entry_pattern.match(curr) \
               or curr.lower().startswith('dep.') \
               or curr.startswith('Deperdita') \
               or curr.startswith('||') \
               or re.fullmatch(r'\d{4}', curr) \
               or curr.startswith('Nr. Nr.'):
                break
            comment_lines.append(curr)
            k += 1
        j = k
    comment_text = ' '.join(comment_lines)

    # Datensatz speichern
    records.append({
        'Nr. Rep': nr_rep,
        'Datum': date_str,
        'Aussteller': aussteller,
        'Ausstellungsort': ausstellungsort,
        'Datumssicherheit': date_sicher,
        'Sicherheit Status Urkunde': status_urkunde,
        'Begünstigte': '',
        'Objektorte': '',
        'Regest': regest_text,
        'Angaben zur Überlieferung': ueber_text,
        'Angaben zu Drucken/Editionen': edition_text,
        'Kommentar': comment_text
    })
    i = j

# Export nach Excel
if records:
    df = pd.DataFrame(records, columns=COLUMNS)
    with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Repertorium')
    print(f"Erzeugt: {OUTPUT_XLSX} mit {len(df)} Einträgen.")
else:
    print("Keine Datensätze gefunden.")
