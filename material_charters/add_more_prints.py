# -*- coding: utf-8 -*-
"""
Skript zur Verarbeitung der Editionen (Editionen auf FactGrid)
Sucht in der Repertoriumstabelle in der Spalte 'Angaben zu Drucken/Editionen' nach Einträgen
und trägt den Wert aus Spalte C der Editionen-Tabelle in 'Publiziert in (P64)' der Metadatenliste v8 ein.
Bei mehreren Einträgen, getrennt durch '=', werden die Teile in die Spalten
'Nr. (qal90)', 'Nr. (qal90).1' und 'Nr. (qal90).2' geschrieben.
"""

import pandas as pd
from pathlib import Path

def process_edition_assignments():
    """Verknüpft Editionen mit Urkunden-Metadaten basierend auf Repertorium-Matches"""

    print("Lade Excel-Dateien für Editionen...")

    # Editionen auf FactGrid lesen (Sheet-Name: 'Editionen auf FactGrid')
    editionen_file = Path(__file__).parent.parent / "datasheets" / "Imports" / "Archive_Editionen_FactGrid.xlsx"
    repertorium_file = Path(__file__).parent.parent / "datasheets" / "Urkunden_Repertorium_v4.xlsx"

    # Input Metadaten: v7 (als Basis) -> Output: v8
    metadaten_input = Path(__file__).parent.parent / "datasheets" / "Imports" / "Urkunden_Metadatenliste_v7.xlsx"
    metadaten_output = Path(__file__).parent.parent / "datasheets" / "Imports" / "Urkunden_Metadatenliste_v8.xlsx"

    # Lese Editionen-Sheet
    try:
        editionen = pd.read_excel(editionen_file, sheet_name="Editionen auf FactGrid", usecols=[0,1,2,3])
        editionen.columns = ['Spalte_A', 'Spalte_B', 'Spalte_C', 'Spalte_D']
    except Exception as e:
        print(f"Fehler beim Lesen der Editionen-Datei: {e}")
        raise

    # Lese Repertorium und Metadaten
    repertorium = pd.read_excel(repertorium_file)
    metadaten = pd.read_excel(metadaten_input)

    print(f"Editionen geladen: {len(editionen)} Zeilen")
    print(f"Repertorium geladen: {len(repertorium)} Zeilen")
    print(f"Metadaten geladen: {len(metadaten)} Zeilen (v7)")

    # Spaltennamen, die wir befüllen wollen
    publiziert_spalte = "Publiziert in (P64)"
    nr_qal90_spalte = "Nr. (qal90)"
    nr_qal90_1_spalte = "Nr. (qal90).1"
    nr_qal90_2_spalte = "Nr. (qal90).2"

    if publiziert_spalte not in metadaten.columns:
        print(f"Warnung: Spalte '{publiziert_spalte}' nicht gefunden in Metadaten. Sie wird angelegt.")
        metadaten[publiziert_spalte] = pd.NA

    # Bereite Suchstrings aus Editionen vor (Spalte B bevorzugt, sonst A)
    editione_suchstrings = []
    for idx, row in editionen.iterrows():
        spalte_b = str(row['Spalte_B']).strip() if pd.notna(row['Spalte_B']) else ""
        spalte_a = str(row['Spalte_A']).strip() if pd.notna(row['Spalte_A']) else ""
        spalte_c = str(row['Spalte_C']).strip() if pd.notna(row['Spalte_C']) else ""

        suchstring = spalte_b if spalte_b and spalte_b.lower() != 'nan' else spalte_a
        if suchstring and suchstring.lower() != 'nan':
            editione_suchstrings.append({
                'suchstring': suchstring,
                'publiziert_wert': spalte_c,
                'zeile': idx + 1
            })

    print(f"Gefundene Editionen-Suchstrings: {len(editione_suchstrings)}")

    matches_gefunden = 0
    aktualisierte_zeilen = 0
    total_repertorium_processed = 0

    # Durchsuche Repertorium nach den Editionen-Suchstrings in 'Angaben zu Drucken/Editionen'
    for entry in editione_suchstrings:
        such = entry['suchstring']
        publiziert_wert = entry['publiziert_wert']
        print(f"\nSuche nach: '{such}' (Publiziert-Wert: {publiziert_wert})")

        repertorium_matches = []
        for idx, rep_row in repertorium.iterrows():
            angaben_drucke = rep_row.get('Angaben zu Drucken/Editionen', '')
            angaben_str = str(angaben_drucke).strip() if pd.notna(angaben_drucke) else ''
            nr_rep = str(rep_row.get('Nr. Rep', '')).strip()

            if angaben_str and such.lower() in angaben_str.lower():
                repertorium_matches.append({
                    'nr_rep': nr_rep,
                    'angaben': angaben_str,
                    'zeile': idx + 1
                })
                print(f"  Match in Repertorium Zeile {idx+1}, Nr. Rep: {nr_rep}")

        # Für jeden Match die entsprechenden Zeilen in Metadaten aktualisieren
        for match in repertorium_matches:
            nr_rep = match['nr_rep']
            angaben = match['angaben']

            total_repertorium_processed += 1
            metadaten_mask = metadaten['Nr. Rep'].astype(str).str.strip() == nr_rep
            matching_rows = metadaten[metadaten_mask]

            if len(matching_rows) > 0:
                print(f"    Aktualisiere {len(matching_rows)} Zeilen in Metadaten für Nr. Rep {nr_rep}")

                # Setze Publiziert in (P64)
                metadaten.loc[metadaten_mask, publiziert_spalte] = publiziert_wert

                # Teile die Angaben zu Drucken/Editionen an '=' und schreibe in die Nr.-Spalten
                teile = [t.strip() for t in angaben.split('=')]
                if len(teile) >= 1 and nr_qal90_spalte in metadaten.columns:
                    metadaten.loc[metadaten_mask, nr_qal90_spalte] = teile[0]
                if len(teile) >= 2 and nr_qal90_1_spalte in metadaten.columns:
                    metadaten.loc[metadaten_mask, nr_qal90_1_spalte] = teile[1]
                if len(teile) >= 3 and nr_qal90_2_spalte in metadaten.columns:
                    metadaten.loc[metadaten_mask, nr_qal90_2_spalte] = teile[2]

                aktualisierte_zeilen += len(matching_rows)
            else:
                print(f"    Keine entsprechenden Zeilen in Metadaten für Nr. Rep {nr_rep}")

        if repertorium_matches:
            matches_gefunden += 1

    # Zusätzliche Zuordnung: Fülle 'Publiziert in (P64).1' und '.2' basierend auf 'Nr. (qal90).1'/'Nr. (qal90).2'
    pub1_col = publiziert_spalte + ".1"
    pub2_col = publiziert_spalte + ".2"

    if pub1_col not in metadaten.columns:
        metadaten[pub1_col] = pd.NA
    if pub2_col not in metadaten.columns:
        metadaten[pub2_col] = pd.NA

    # Baue Suchliste aus Editionen (Spalte B bevorzugt, sonst A)
    search_entries = []
    for idx, row in editionen.iterrows():
        a = str(row['Spalte_A']).strip() if pd.notna(row['Spalte_A']) else ''
        b = str(row['Spalte_B']).strip() if pd.notna(row['Spalte_B']) else ''
        c = str(row['Spalte_C']).strip() if pd.notna(row['Spalte_C']) else ''
        if b and b.lower() != 'nan':
            search_entries.append({'such': b, 'publ': c})
        if a and a.lower() != 'nan':
            search_entries.append({'such': a, 'publ': c})

    def find_publ_for_nr(nr_value):
        if pd.isna(nr_value):
            return None
        nrstr = str(nr_value).strip()
        if not nrstr:
            return None
        for entry in search_entries:
            try:
                if entry['such'].lower() in nrstr.lower():
                    return entry['publ']
            except Exception:
                continue
        return None

    # Durchlaufe Metadaten und setze Publiziert in (P64).1/.2 wenn passende Nr gefunden
    gesetzt_pub1 = 0
    gesetzt_pub2 = 0
    for idx, mrow in metadaten.iterrows():
        nr1 = mrow.get(nr_qal90_1_spalte, '') if nr_qal90_1_spalte in metadaten.columns else ''
        nr2 = mrow.get(nr_qal90_2_spalte, '') if nr_qal90_2_spalte in metadaten.columns else ''

        publ1 = find_publ_for_nr(nr1)
        if publ1:
            metadaten.at[idx, pub1_col] = publ1
            gesetzt_pub1 += 1

        publ2 = find_publ_for_nr(nr2)
        if publ2:
            metadaten.at[idx, pub2_col] = publ2
            gesetzt_pub2 += 1

    print(f"\nPubliziert in (P64).1 gesetzt: {gesetzt_pub1}")
    print(f"Publiziert in (P64).2 gesetzt: {gesetzt_pub2}")

    # Speichere die aktualisierten Metadaten als v8
    metadaten.to_excel(metadaten_output, index=False)

    print("\n=== ZUSAMMENFASSUNG ===")
    print(f"Editionen-Suchstrings verarbeitet: {len(editione_suchstrings)}")
    print(f"Matches in Repertorium gefunden: {matches_gefunden}")
    print(f"Repertorium-Einträge verarbeitet: {total_repertorium_processed}")
    print(f"Aktualisierte Zeilen in Metadaten: {aktualisierte_zeilen}")
    print(f"Aktualisierte Datei gespeichert: {metadaten_output}")

    return metadaten


if __name__ == '__main__':
    try:
        result = process_edition_assignments()
        print("\nSkript erfolgreich ausgeführt!")
    except Exception as e:
        print(f"Fehler beim Ausführen des Skripts: {e}")
        import traceback
        traceback.print_exc()