#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skript zur Verarbeitung der Urkunden-Metadaten
Liest verschiedene Excel-Dateien und erstellt eine neue Version v4 der Metadatenliste
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
    
    repertorium_file = base_path / "Urkunden_Repertorium_v4.xlsx"
    archive_editionen_file = base_path / "Imports" / "Archive_Editionen_FactGrid.xlsx"
    metadaten_v3_file = base_path / "Imports" / "Urkunden_Metadatenliste_v3.xlsx"
    metadaten_v4_file = base_path / "Imports" / "Urkunden_Metadatenliste_v4.xlsx"
    
    try:
        # Excel-Dateien einlesen
        print("Lade Urkunden_Repertorium_v4.xlsx...")
        repertorium_df = pd.read_excel(repertorium_file)
        
        print("Lade Archive_Editionen_FactGrid.xlsx...")
        archive_editionen_df = pd.read_excel(archive_editionen_file)
        
        print("Lade Urkunden_Metadatenliste_v3.xlsx...")
        metadaten_v3_df = pd.read_excel(metadaten_v3_file)
        
        # Spalten überprüfen und anzeigen
        print("\n=== Verfügbare Spalten ===")
        print("Repertorium:", list(repertorium_df.columns))
        print("Archive/Editionen:", list(archive_editionen_df.columns))
        print("Metadaten v3:", list(metadaten_v3_df.columns))
        
        # Kopie der v3 Metadaten für v4 erstellen
        metadaten_v4_df = metadaten_v3_df.copy()
        
        # Dictionary für Abkürzung -> FactGrid-Item mapping erstellen
        abkuerzung_mapping = {}
        for _, row in archive_editionen_df.iterrows():
            abkuerzung = str(row['Abkürzung']).strip() if pd.notna(row['Abkürzung']) else ''
            factgrid_item = str(row['FactGrid-Item']).strip() if pd.notna(row['FactGrid-Item']) else ''
            if abkuerzung and factgrid_item:
                abkuerzung_mapping[abkuerzung] = factgrid_item
        
        print(f"\n{len(abkuerzung_mapping)} Abkürzungen geladen:")
        for abk, fg in list(abkuerzung_mapping.items())[:5]:  # Erste 5 anzeigen
            print(f"  {abk} -> {fg}")
        if len(abkuerzung_mapping) > 5:
            print("  ...")
        
        # Spalten-Mapping für bessere Lesbarkeit
        publiziert_spalte = "Publiziert in (P64)"
        nr_qal90_spalte = "Nr. (qal90)"
        nr_qal90_1_spalte = "Nr. (qal90).1"
        nr_qal90_2_spalte = "Nr. (qal90).2"
        originalitaet_spalte = "Originalität (P115)"
        
        # Verarbeitung der Repertorium-Einträge
        print("\n=== Verarbeitung beginnt ===")
        matches_found = 0
        total_processed = 0
        
        for _, rep_row in repertorium_df.iterrows():
            nr_rep = rep_row['Nr. Rep'] if pd.notna(rep_row['Nr. Rep']) else None
            angaben_drucke = str(rep_row['Angaben zu Drucken/Editionen']) if pd.notna(rep_row['Angaben zu Drucken/Editionen']) else ''
            
            if nr_rep is None or not angaben_drucke:
                continue
                
            total_processed += 1
            
            # Suche nach Abkürzungen in den Angaben zu Drucken/Editionen
            for abkuerzung, factgrid_item in abkuerzung_mapping.items():
                if abkuerzung in angaben_drucke:
                    # Finde die entsprechende Zeile in der Metadatenliste
                    mask = metadaten_v4_df['Nr. Rep'] == nr_rep
                    matching_rows = metadaten_v4_df[mask]
                    
                    if not matching_rows.empty:
                        row_index = matching_rows.index[0]
                        
                        # FactGrid-Item in Publiziert-Spalte eintragen
                        if publiziert_spalte in metadaten_v4_df.columns:
                            metadaten_v4_df.at[row_index, publiziert_spalte] = factgrid_item
                        
                        # Gesamte Zeile nach "=" aufteilen und in entsprechende Nr. (qal90) Spalten eintragen
                        teile = angaben_drucke.split("=")
                        if len(teile) >= 1 and nr_qal90_spalte in metadaten_v4_df.columns:
                            metadaten_v4_df.at[row_index, nr_qal90_spalte] = teile[0].strip()
                        if len(teile) >= 2 and nr_qal90_1_spalte in metadaten_v4_df.columns:
                            metadaten_v4_df.at[row_index, nr_qal90_1_spalte] = teile[1].strip()
                        if len(teile) >= 3 and nr_qal90_2_spalte in metadaten_v4_df.columns:
                            metadaten_v4_df.at[row_index, nr_qal90_2_spalte] = teile[2].strip()
                        
                        matches_found += 1
                        print(f"Match gefunden: Nr. Rep {nr_rep}, Abkürzung '{abkuerzung}' -> {factgrid_item}")
                        break  # Nur erste gefundene Abkürzung verwenden
        
        # Originalität-Spalte bearbeiten: nur erstes Wort behalten (außer bei "verlorenes original")
        if originalitaet_spalte in metadaten_v4_df.columns:
            def extract_first_word(text):
                if pd.isna(text):
                    return text
                text_str = str(text).strip()
                if not text_str:
                    return text_str
                
                words = text_str.split()
                if not words:
                    return text_str
                
                # Prüfe ob die ersten zwei Wörter "verlorenes original" sind (case-insensitive)
                if len(words) >= 2 and words[0].lower() == "verlorenes" and words[1].lower() == "original":
                    return " ".join(words[:2])
                else:
                    return words[0]
            
            metadaten_v4_df[originalitaet_spalte] = metadaten_v4_df[originalitaet_spalte].apply(extract_first_word)
            print(f"\nOriginalität-Spalte bearbeitet: erstes Wort behalten (außer 'verlorenes original')")
        
        # Ergebnisse speichern
        print(f"\n=== Ergebnisse ===")
        print(f"Verarbeitete Repertorium-Einträge: {total_processed}")
        print(f"Gefundene Matches: {matches_found}")
        
        metadaten_v4_df.to_excel(metadaten_v4_file, index=False)
        print(f"\nNeue Datei gespeichert: {metadaten_v4_file}")
        
        # Zusammenfassung der Änderungen
        if matches_found > 0:
            print(f"\n=== Änderungen in v4 ===")
            if publiziert_spalte in metadaten_v4_df.columns:
                nicht_leer_publiziert = metadaten_v4_df[publiziert_spalte].notna().sum()
                print(f"Einträge mit Publiziert-Information: {nicht_leer_publiziert}")
            
            if nr_qal90_spalte in metadaten_v4_df.columns:
                nicht_leer_nr_qal90 = metadaten_v4_df[nr_qal90_spalte].notna().sum()
                print(f"Einträge mit Nr. (qal90): {nicht_leer_nr_qal90}")
            
            if nr_qal90_1_spalte in metadaten_v4_df.columns:
                nicht_leer_nr_qal90_1 = metadaten_v4_df[nr_qal90_1_spalte].notna().sum()
                print(f"Einträge mit Nr. (qal90).1: {nicht_leer_nr_qal90_1}")
            
            if nr_qal90_2_spalte in metadaten_v4_df.columns:
                nicht_leer_nr_qal90_2 = metadaten_v4_df[nr_qal90_2_spalte].notna().sum()
                print(f"Einträge mit Nr. (qal90).2: {nicht_leer_nr_qal90_2}")
        
    except FileNotFoundError as e:
        print(f"Fehler: Datei nicht gefunden - {e}")
    except Exception as e:
        print(f"Fehler bei der Verarbeitung: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()