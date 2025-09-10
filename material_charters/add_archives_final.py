import pandas as pd
import numpy as np

def process_archive_assignments():
    """
    Verknüpft Archive-Editionen mit Urkunden-Metadaten basierend auf Repertorium-Matches
    """

    # 1. Lade die Excel-Dateien
    print("Lade Excel-Dateien...")

    # Archive Editionen (Spalten A, B, C)
    archive_editionen = pd.read_excel("../datasheets/Imports/Archive_Editionen_FactGrid_SJ.xlsx",
                                     usecols=[0, 1, 2],  # Spalten A, B, C
                                     sheet_name="Archive auf FactGrid")
    archive_editionen.columns = ['Spalte_A', 'Spalte_B', 'Spalte_C']

    # Repertorium (Nr. Rep und Angaben zu Drucken/Editionen)
    repertorium = pd.read_excel("../datasheets/Urkunden_Repertorium_v4.xlsx")

    # Metadatenliste
    metadaten = pd.read_excel("../datasheets/Imports/Urkunden_Metadatenliste_v10.xlsx")

    print(f"Archive Editionen geladen: {len(archive_editionen)} Zeilen")
    print(f"Repertorium geladen: {len(repertorium)} Zeilen")
    print(f"Metadaten geladen: {len(metadaten)} Zeilen")

    # 2. Suche die richtige Spalte für "Aktueller Standort"
    standort_spalte = None
    for col in metadaten.columns:
        if "Aktueller Standort" in str(col) and "P329" in str(col):
            standort_spalte = col
            break

    if standort_spalte is None:
        print("Warnung: Spalte 'Aktueller Standort (P329)' nicht gefunden in Metadaten")
        # Fallback: erste Spalte die "Standort" enthält
        for col in metadaten.columns:
            if "Standort" in str(col):
                standort_spalte = col
                break

    print(f"Verwende Standort-Spalte: {standort_spalte}")

    # Suche die Spalte für "Inventarposition"
    inventar_spalte = None
    for col in metadaten.columns:
        if "Inventarposition" in str(col) and "qal10" in str(col):
            inventar_spalte = col
            break

    if inventar_spalte is None:
        print("Warnung: Spalte 'Inventarposition (qal10)' nicht gefunden in Metadaten")
        # Fallback: erste Spalte die "Inventar" enthält
        for col in metadaten.columns:
            if "Inventar" in str(col):
                inventar_spalte = col
                break

    print(f"Verwende Inventar-Spalte: {inventar_spalte}")

    # Suche die Spalte für "Notiz, bei unbekannter Position"
    notiz_spalte = None
    for col in metadaten.columns:
        if "Notiz" in str(col) and "qal73" in str(col):
            notiz_spalte = col
            break

    if notiz_spalte is None:
        print("Warnung: Spalte 'Notiz, bei unbekannter Position (qal73)' nicht gefunden in Metadaten")
        # Fallback: erste Spalte die "Notiz" enthält
        for col in metadaten.columns:
            if "Notiz" in str(col):
                notiz_spalte = col
                break

    print(f"Verwende Notiz-Spalte: {notiz_spalte}")

    # 3. Bereite Archive-Editionen vor (Spalte B oder A als Suchstring)
    archive_suchstrings = []
    for idx, row in archive_editionen.iterrows():
        spalte_b = str(row['Spalte_B']).strip() if pd.notna(row['Spalte_B']) else ""
        spalte_a = str(row['Spalte_A']).strip() if pd.notna(row['Spalte_A']) else ""
        spalte_c = str(row['Spalte_C']).strip() if pd.notna(row['Spalte_C']) else ""

        # Wenn Spalte B leer ist, nimm Spalte A
        suchstring = spalte_b if spalte_b and spalte_b != 'nan' else spalte_a

        if suchstring and suchstring != 'nan':
            archive_suchstrings.append({
                'suchstring': suchstring,
                'archiv_wert': spalte_c,  # FactGrid-Item aus Spalte C
                'zeile': idx + 1
            })

    print(f"Gefundene Archive-Suchstrings: {len(archive_suchstrings)}")

    # 4. Suche Matches zwischen Archive-Editionen und Repertorium
    matches_gefunden = 0
    aktualisierte_zeilen = 0

    for archive_entry in archive_suchstrings:
        suchstring = archive_entry['suchstring']
        archiv_wert = archive_entry['archiv_wert']

        print(f"\nSuche nach: '{suchstring}'")

        # Suche in Repertorium Spalte "Angaben zur Überlieferung"
        repertorium_matches = []

        for idx, row in repertorium.iterrows():
            angaben_editionen = str(row.get('Angaben zur Überlieferung', '')).strip()
            nr_rep = str(row.get('Nr. Rep', '')).strip()

            if pd.notna(angaben_editionen) and suchstring.lower() in angaben_editionen.lower():
                # Extrahiere Inventarposition falls Abkürzung mit Komma gefunden wird
                inventar_position = None
                suchstring_lower = suchstring.lower()

                # Debug-Ausgabe
                print(f"  Prüfe Text: '{angaben_editionen}'")
                print(f"  Suche nach: '{suchstring_lower}'")

                # Suche nach der Abkürzung gefolgt von einem Komma (mit optionalen Whitespaces)
                import re
                # Vereinfachtes Pattern ohne word boundary für bessere Treffer
                pattern = re.escape(suchstring_lower) + r'\s*,\s*(.+)$'
                match = re.search(pattern, angaben_editionen.lower())  # Suche in lowercase

                if match:
                    # Extrahiere aus dem Original-String
                    original_match = re.search(re.escape(suchstring) + r'\s*,\s*(.+)$', angaben_editionen)
                    if original_match:
                        inventar_position = original_match.group(1).strip()
                        print(f"  Inventarposition gefunden: '{inventar_position}'")
                    else:
                        print("  Fehler: Konnte Inventarposition nicht aus Original-String extrahieren")
                else:
                    print(f"  Kein Komma nach '{suchstring}' gefunden")

                repertorium_matches.append({
                    'nr_rep': nr_rep,
                    'angaben': angaben_editionen,
                    'zeile': idx + 1,
                    'inventar_position': inventar_position
                })
                print(f"  Match gefunden in Repertorium Zeile {idx + 1}, Nr. Rep: {nr_rep}")

        # 5. Für jeden Match, finde entsprechende Zeilen in Metadaten und aktualisiere
        for match in repertorium_matches:
            nr_rep = match['nr_rep']
            inventar_position = match.get('inventar_position')

            # Finde alle Zeilen in Metadaten mit derselben Nr. Rep
            metadaten_mask = metadaten['Nr. Rep'].astype(str).str.strip() == nr_rep
            matching_rows = metadaten[metadaten_mask]

            if len(matching_rows) > 0:
                print(f"    Aktualisiere {len(matching_rows)} Zeilen in Metadaten für Nr. Rep {nr_rep}")
                print(f"    Setze Archiv-Wert: {archiv_wert}")

                # Aktualisiere die Standort-Spalte
                metadaten.loc[metadaten_mask, standort_spalte] = archiv_wert
                aktualisierte_zeilen += len(matching_rows)

                # Aktualisiere die Inventarposition-Spalte falls gefunden
                if inventar_position and inventar_spalte:
                    metadaten.loc[metadaten_mask, inventar_spalte] = inventar_position
                    print(f"    Setze Inventarposition: {inventar_position}")
            else:
                print(f"    Keine entsprechenden Zeilen in Metadaten für Nr. Rep {nr_rep}")

        if repertorium_matches:
            matches_gefunden += 1

    # 5.5 Zusätzliche Verarbeitung: Fülle Notiz-Spalte für leere Standort-Einträge
    print("\n=== VERARBEITUNG LEERER STANDORT-EINTRÄGE ===")
    notiz_updates = 0
    
    for idx, row in metadaten.iterrows():
        nr_rep = str(row.get('Nr. Rep', '')).strip()
        standort_wert = str(row.get(standort_spalte, '')).strip() if standort_spalte else ""
        
        # Prüfe ob Standort-Spalte leer ist
        if not standort_wert or standort_wert == 'nan' or standort_wert == '':
            # Finde entsprechenden Eintrag im Repertorium
            repertorium_row = repertorium[repertorium['Nr. Rep'].astype(str).str.strip() == nr_rep]
            
            if not repertorium_row.empty:
                angaben_ueberlieferung = str(repertorium_row.iloc[0].get('Angaben zur Überlieferung', '')).strip()
                
                if angaben_ueberlieferung and angaben_ueberlieferung != 'nan':
                    # Trage den Wert in die Notiz-Spalte ein
                    if notiz_spalte:
                        metadaten.at[idx, notiz_spalte] = angaben_ueberlieferung
                        notiz_updates += 1
                        print(f"  Notiz für Nr. Rep {nr_rep}: '{angaben_ueberlieferung}'")

    print(f"Notiz-Einträge für leere Standorte: {notiz_updates}")

    # 5.6 Lösche die Spalte "Datumssicherheit (qal155)"
    datumssicherheit_spalte = None
    for col in metadaten.columns:
        if "Datumssicherheit" in str(col) and "qal155" in str(col):
            datumssicherheit_spalte = col
            break
    
    if datumssicherheit_spalte:
        print(f"\nLösche Spalte: {datumssicherheit_spalte}")
        metadaten = metadaten.drop(columns=[datumssicherheit_spalte])
    else:
        print("\nSpalte 'Datumssicherheit (qal155)' nicht gefunden - nichts zu löschen")

    # 6. Speichere aktualisierte Metadaten
    output_file = "../datasheets/Imports/Urkunden_Metadatenliste_v11.xlsx"
    metadaten.to_excel(output_file, index=False)

    print(f"\n=== ZUSAMMENFASSUNG ===")
    print(f"Archive-Suchstrings verarbeitet: {len(archive_suchstrings)}")
    print(f"Matches in Repertorium gefunden: {matches_gefunden}")
    print(f"Aktualisierte Zeilen in Metadaten: {aktualisierte_zeilen}")
    print(f"Aktualisierte Datei gespeichert: {output_file}")

    return metadaten

if __name__ == "__main__":
    try:
        result = process_archive_assignments()
        print("\nSkript erfolgreich ausgeführt!")
    except Exception as e:
        print(f"Fehler beim Ausführen des Skripts: {e}")
        import traceback
        traceback.print_exc()