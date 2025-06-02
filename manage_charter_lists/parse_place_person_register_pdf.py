import pdfplumber
import re
import pandas as pd


def read_pdf_columns(pdf_path):
    """
    Liest eine zweispaltige PDF ein und gibt eine Liste aller Zeilen (erst links, dann rechts) zurück.
    """
    extracted_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            width, height = page.width, page.height
            mid_x = width / 2

            # linke Spalte
            left = page.crop((0, 0, mid_x, height)).extract_text() or ""
            # rechte Spalte
            right = page.crop((mid_x, 0, width, height)).extract_text() or ""

            # Zeilen splitten und sammeln
            left_lines = left.split("\n")
            right_lines = right.split("\n")
            extracted_lines.extend(left_lines + right_lines)

    return extracted_lines


def main():
    pdf_path = "../datasheets/orts_namen_register.pdf"
    ident_excel = "../datasheets/orte/Orte_Identifikation_factgrid.xlsx"
    output_excel = "../datasheets/orts_namen_register.xlsx"

    # PDF lesen
    lines = read_pdf_columns(pdf_path)

    # Marker-Zeile finden
    marker = "Orts- und Personennamen zu entnehmen."
    start_idx = None
    for idx, ln in enumerate(lines):
        if marker in ln:
            start_idx = idx + 1
            break
    if start_idx is None:
        print(f"Marker-Zeile '{marker}' nicht gefunden.")
        return

    # Identifikationstabelle einlesen
    df_ident = pd.read_excel(ident_excel)
    required_col = 'Schreibweise Ortsregister'
    if required_col not in df_ident.columns:
        raise KeyError(f"Spalte '{required_col}' nicht gefunden in Identifikationstabelle")

    # Orte aus mehreren Spalten sammeln
    place_names = set()
    # Hauptschreibweise
    for val in df_ident[required_col].dropna().astype(str):
        name = val.strip()
        if name:
            place_names.add(name)
    # alternative/historische Schreibweisen
    hist_col = 'alternative/historische Schreibweisen'
    if hist_col in df_ident.columns:
        for cell in df_ident[hist_col].dropna().astype(str):
            for part in cell.split(';'):
                part = part.strip()
                if part:
                    place_names.add(part)
    # Heutiger Name (kann mehrere durch ';' getrennte Einträge enthalten)
    today_col = 'Heutiger Name'
    if today_col in df_ident.columns:
        for cell in df_ident[today_col].dropna().astype(str):
            for part in cell.split(';'):
                part = part.strip()
                if part:
                    place_names.add(part)

    place_list = list(place_names)
    # Für case-insensitive Abgleich
    place_list_lower = [p.lower() for p in place_list]

    # Header-Marker zum Überspringen
    header_markers = {"| Orts- und Personenregister", "Orts- und Personenregister |"}

    # Ab Marker: Zeilen bereinigen (Ziffern entfernen, Header skippen)
    cleaned_lines = []
    for ln in lines[start_idx:]:
        cleaned = re.sub(r"\d+", "", ln).strip()
        if not cleaned or cleaned in header_markers:
            continue
        cleaned_lines.append(cleaned)

    # Enthaltene Orte und Personen extrahieren
    place_entries = []
    person_entries = []
    last_base_entry = None
    en_dash = '–'

    for ln in cleaned_lines:
        entry = ln.split(',')[0].strip()
        if not entry:
            continue

        if ln.startswith(en_dash):
            # Entferne Gedankenstrich und reduziere Leerraum
            person_name = re.sub(r'\s+', ' ', entry.lstrip(en_dash)).strip()
            if last_base_entry:
                full_name = f"{person_name} {last_base_entry}"
            else:
                full_name = person_name
            full_name = re.sub(r'\s+', ' ', full_name).strip()
            person_entries.append(full_name)
        else:
            last_base_entry = entry
            entry_lower = entry.lower()
            if entry_lower in place_list_lower:
                matched_place = place_list[place_list_lower.index(entry_lower)]
                place_entries.append(matched_place)
            else:
                person_entries.append(entry)

    # Debug-Ausgaben
    print("Extracted places:", place_entries)
    print("Extracted persons:", person_entries)

    # Excel-Ausgabe vorbereiten
    df_places = pd.DataFrame({
        'Ort': place_entries,
        'Person': [''] * len(place_entries)
    })
    df_persons = pd.DataFrame({
        'Ort': [''] * len(person_entries),
        'Person': person_entries
    })
    df_output = pd.concat([df_places, df_persons], ignore_index=True)

    # Entferne Einträge, die nur 'f.' enthalten
    df_output = df_output[~((df_output['Ort'].str.strip() == 'f.') | (df_output['Person'].str.strip() == 'f.'))]

    # Strip whitespace und reduziere mehrfach-Leerzeichen
    df_output['Ort'] = df_output['Ort'] \
        .astype(str) \
        .str.strip() \
        .str.replace(r'\s+', ' ', regex=True)
    df_output['Person'] = df_output['Person'] \
        .astype(str) \
        .str.strip() \
        .str.replace(r'\s+', ' ', regex=True)

    # Entferne Personen, die wirklich Orte sind
    place_set = {o for o in df_output['Ort'] if o}
    df_output = df_output[~df_output['Person'].isin(place_set)]

    # Speichern
    df_output.to_excel(output_excel, index=False)
    print(f"Excel gespeichert: {output_excel}")


if __name__ == "__main__":
    main()
