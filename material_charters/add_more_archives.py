import pandas as pd
import numpy as np

def process_archive_assignments():
    """
    Verknüpft Archive-Editionen mit Urkunden-Metadaten basierend auf Repertorium-Matches
    """
    
    # 1. Lade die Excel-Dateien
    print("Lade Excel-Dateien...")
    
    # Archive Editionen (nur Spalten A, B, C (FactGrid-Item) und D und nur bis Zeile 17)
    archive_editionen = pd.read_excel("../datasheets/Imports/Archive_Editionen_FactGrid.xlsx", 
                                     usecols=[0, 1, 2, 3],  # Spalten A, B, C, D
                                     nrows=17,
                                     sheet_name="Archive auf FactGrid")
    archive_editionen.columns = ['Spalte_A', 'Spalte_B', 'Spalte_C', 'Spalte_D']
    
    # Repertorium (Nr. Rep und Angaben zu Drucken/Editionen)
    repertorium = pd.read_excel("../datasheets/Urkunden_Repertorium_v4.xlsx")
    
    # Metadatenliste
    metadaten = pd.read_excel("../datasheets/Imports/Urkunden_Metadatenliste_v6.xlsx")
    
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
    
    # 3. Bereite Archive-Editionen vor (Spalte B oder A als Suchstring)
    archive_suchstrings = []
    for idx, row in archive_editionen.iterrows():
        spalte_b = str(row['Spalte_B']).strip() if pd.notna(row['Spalte_B']) else ""
        spalte_a = str(row['Spalte_A']).strip() if pd.notna(row['Spalte_A']) else ""
        spalte_c = str(row['Spalte_C']).strip() if pd.notna(row['Spalte_C']) else ""
        spalte_d = str(row['Spalte_D']).strip() if pd.notna(row['Spalte_D']) else ""
        
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
            angaben_ueberlieferung = row.get('Angaben zur Überlieferung', '')
            angaben_ueberlieferung_str = str(angaben_ueberlieferung).strip() if pd.notna(angaben_ueberlieferung) else ''
            nr_rep = str(row.get('Nr. Rep', '')).strip()
            
            if angaben_ueberlieferung_str and suchstring.lower() in angaben_ueberlieferung_str.lower():
                repertorium_matches.append({
                    'nr_rep': nr_rep,
                    'angaben': angaben_ueberlieferung_str,
                    'zeile': idx + 1
                })
                print(f"  Match gefunden in Repertorium Zeile {idx + 1}, Nr. Rep: {nr_rep}")
        
        # 5. Für jeden Match, finde entsprechende Zeilen in Metadaten und aktualisiere
        for match in repertorium_matches:
            nr_rep = match['nr_rep']
            
            # Finde alle Zeilen in Metadaten mit derselben Nr. Rep
            metadaten_mask = metadaten['Nr. Rep'].astype(str).str.strip() == nr_rep
            matching_rows = metadaten[metadaten_mask]
            
            if len(matching_rows) > 0:
                print(f"    Aktualisiere {len(matching_rows)} Zeilen in Metadaten für Nr. Rep {nr_rep}")
                print(f"    Setze Archiv-Wert: {archiv_wert}")
                
                # Aktualisiere die Standort-Spalte
                metadaten.loc[metadaten_mask, standort_spalte] = archiv_wert
                aktualisierte_zeilen += len(matching_rows)
            else:
                print(f"    Keine entsprechenden Zeilen in Metadaten für Nr. Rep {nr_rep}")
        
        if repertorium_matches:
            matches_gefunden += 1
    
    # 6. Speichere aktualisierte Metadaten
    output_file = "../datasheets/Imports/Urkunden_Metadatenliste_v7.xlsx"
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