#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skript zur Verarbeitung der Ortsdaten in den Urkunden-Metadaten
Liest Ortsdaten und aktualisiert die Ausstellungsorte mit FactGrid QIDs
"""

import pandas as pd
import re
import os
from pathlib import Path

def main():
    """
    Hauptfunktion des Skripts
    """
    # Dateipfade definieren
    base_path = Path(__file__).parent.parent / "datasheets"
    
    orte_file = base_path / "orte" / "Orte_Identifikation_factgrid.xlsx"
    metadaten_v4_file = base_path / "Imports" / "Urkunden_Metadatenliste_v4.xlsx"
    metadaten_v5_file = base_path / "Imports" / "Urkunden_Metadatenliste_v5.xlsx"
    
    try:
        # Excel-Dateien einlesen
        print("Lade Orte_Identifikation_factgrid.xlsx...")
        orte_df = pd.read_excel(orte_file)
        
        print("Lade Urkunden_Metadatenliste_v4.xlsx...")
        metadaten_v4_df = pd.read_excel(metadaten_v4_file)
        
        # Spalten überprüfen und anzeigen
        print("\n=== Verfügbare Spalten ===")
        print("Ortsdaten:", list(orte_df.columns))
        print("Metadaten v4:", list(metadaten_v4_df.columns))
        
        # Kopie der v4 Metadaten für v5 erstellen
        metadaten_v5_df = metadaten_v4_df.copy()
        
        # Dictionary für Ortsname -> FactGrid QID mapping erstellen
        orte_mapping = {}
        for _, row in orte_df.iterrows():
            schreibweise = str(row['Schreibweise Ortsregister']).strip() if pd.notna(row['Schreibweise Ortsregister']) else ''
            factgrid_qid = str(row['Factgrid QID']).strip() if pd.notna(row['Factgrid QID']) else ''
            if schreibweise and factgrid_qid:
                orte_mapping[schreibweise] = factgrid_qid
        
        print(f"\n{len(orte_mapping)} Ortsnamen geladen:")
        for ort, qid in list(orte_mapping.items())[:5]:  # Erste 5 anzeigen
            print(f"  '{ort}' -> {qid}")
        if len(orte_mapping) > 5:
            print("  ...")
        
        # Spaltenname für Ausstellungsort
        ausstellungsort_spalte = "Ausstellungsort (P926)"
        
        # Prüfen ob die Spalte existiert
        if ausstellungsort_spalte not in metadaten_v5_df.columns:
            print(f"\nFehler: Spalte '{ausstellungsort_spalte}' nicht gefunden!")
            print("Verfügbare Spalten:")
            for col in metadaten_v5_df.columns:
                if 'ort' in col.lower() or 'ausstellung' in col.lower():
                    print(f"  - {col}")
            return
        
        # Verarbeitung der Ausstellungsorte
        print(f"\n=== Verarbeitung der Spalte '{ausstellungsort_spalte}' ===")
        matches_found = 0
        total_processed = 0
        
        for index, row in metadaten_v5_df.iterrows():
            ausstellungsort = str(row[ausstellungsort_spalte]) if pd.notna(row[ausstellungsort_spalte]) else ''
            
            if not ausstellungsort or ausstellungsort == 'nan':
                continue
                
            total_processed += 1
            
            # Exakte Übereinstimmung prüfen
            if ausstellungsort in orte_mapping:
                factgrid_qid = orte_mapping[ausstellungsort]
                metadaten_v5_df.at[index, ausstellungsort_spalte] = factgrid_qid
                matches_found += 1
                print(f"Match gefunden: '{ausstellungsort}' -> {factgrid_qid}")
            else:
                # Fuzzy matching - prüfe ob der Ausstellungsort in einem der Ortsnamen enthalten ist
                found_match = False
                for ortsname, qid in orte_mapping.items():
                    # Prüfe sowohl Ausstellungsort in Ortsname als auch Ortsname in Ausstellungsort
                    if (ausstellungsort.lower() in ortsname.lower() or 
                        ortsname.lower() in ausstellungsort.lower()):
                        metadaten_v5_df.at[index, ausstellungsort_spalte] = qid
                        matches_found += 1
                        print(f"Fuzzy Match gefunden: '{ausstellungsort}' ~ '{ortsname}' -> {qid}")
                        found_match = True
                        break
                
                if not found_match:
                    print(f"Kein Match für: '{ausstellungsort}'")
        
        # Ergebnisse speichern
        print(f"\n=== Ergebnisse ===")
        print(f"Verarbeitete Ausstellungsorte: {total_processed}")
        print(f"Gefundene Matches: {matches_found}")
        print(f"Match-Rate: {matches_found/total_processed*100:.1f}%" if total_processed > 0 else "Match-Rate: 0%")
        
        metadaten_v5_df.to_excel(metadaten_v5_file, index=False)
        print(f"\nNeue Datei gespeichert: {metadaten_v5_file}")
        
        # Zusammenfassung der Änderungen
        if matches_found > 0:
            print(f"\n=== Änderungen in v5 ===")
            # Zähle wie viele Einträge jetzt FactGrid QIDs haben (beginnen mit Q)
            qid_entries = metadaten_v5_df[ausstellungsort_spalte].astype(str).str.startswith('Q').sum()
            print(f"Einträge mit FactGrid QIDs (beginnen mit 'Q'): {qid_entries}")
        
    except FileNotFoundError as e:
        print(f"Fehler: Datei nicht gefunden - {e}")
        print(f"Aktueller Pfad: {Path.cwd()}")
        print(f"Gesuchte Datei: {e.filename if hasattr(e, 'filename') else 'unbekannt'}")
    except Exception as e:
        print(f"Fehler bei der Verarbeitung: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()