import pandas as pd
import numpy as np

def match_archive_editions():
    """
    Lädt drei Excel-Tabellen und führt Zuordnung von Archiv-Editionen durch.
    """
    
    # Excel-Dateien laden
    print("Lade Excel-Dateien...")
    
    try:
        # Archive Editionen FactGrid laden
        archive_df = pd.read_excel("../datasheets/Imports/Archive_Editionen_FactGrid.xlsx")
        print(f"Archive Editionen geladen: {len(archive_df)} Zeilen")
        
        # Urkunden Metadatenliste laden
        metadata_df = pd.read_excel("../datasheets/Imports/Urkunden_Metadatenliste_v6.xlsx")
        print(f"Metadatenliste geladen: {len(metadata_df)} Zeilen")
        
        # Urkunden Repertorium laden
        repertorium_df = pd.read_excel("../datasheets/Urkunden_Repertorium_v4.xlsx")
        print(f"Repertorium geladen: {len(repertorium_df)} Zeilen")
        
    except FileNotFoundError as e:
        print(f"Fehler beim Laden der Dateien: {e}")
        return
    
    # Archive Editionen vorbereiten (nur bis Zeile 17)
    archive_subset = archive_df.head(17).copy()
    
    # Suchstrings aus Spalte B erstellen, bei leeren Zellen Spalte A verwenden
    search_strings = []
    for idx, row in archive_subset.iterrows():
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip():  # Spalte B (Index 1)
            search_strings.append((idx, str(row.iloc[1]).strip(), row.iloc[3]))  # Index, String, Spalte D
        elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():  # Spalte A (Index 0)
            search_strings.append((idx, str(row.iloc[0]).strip(), row.iloc[3]))  # Index, String, Spalte D
    
    print(f"Gefundene Suchstrings: {len(search_strings)}")
    
    # Spalten identifizieren
    target_column = "Aktueller Standort (P329), wenn nicht bekannt: Q400468 (vielleicht besser: Q893727)"
    
    # Prüfen ob die Zielspalte existiert
    if target_column not in metadata_df.columns:
        print(f"Warnung: Spalte '{target_column}' nicht in Metadatenliste gefunden.")
        print("Verfügbare Spalten:")
        for col in metadata_df.columns:
            print(f"  - {col}")
        return
    
    matches_found = 0
    updates_made = 0
    
    # Für jeden Suchstring
    for idx, search_string, column_d_value in search_strings:
        print(f"\nSuche nach: '{search_string}'")
        
        # In Repertorium nach Matches in "Angaben zu Drucken/Editionen" suchen
        if "Angaben zu Drucken/Editionen" in repertorium_df.columns:
            mask = repertorium_df["Angaben zu Drucken/Editionen"].astype(str).str.contains(
                search_string, case=False, na=False, regex=False
            )
            matching_rows = repertorium_df[mask]
            
            if len(matching_rows) > 0:
                matches_found += 1
                print(f"  Gefunden in {len(matching_rows)} Repertorium-Zeilen")
                
                # Für jede gefundene Nr Rep
                for _, rep_row in matching_rows.iterrows():
                    nr_rep = rep_row["Nr. Rep"]
                    print(f"    Nr. Rep: {nr_rep}")
                    
                    # In Metadatenliste nach gleicher Nr Rep suchen
                    if "Nr Rep" in metadata_df.columns:
                        metadata_mask = metadata_df["Nr Rep"] == nr_rep
                        matching_metadata_rows = metadata_df[metadata_mask]
                        
                        if len(matching_metadata_rows) > 0:
                            print(f"      Aktualisiere {len(matching_metadata_rows)} Metadaten-Zeilen")
                            
                            # Spalte D Wert in Zielspalte eintragen
                            metadata_df.loc[metadata_mask, target_column] = column_d_value
                            updates_made += len(matching_metadata_rows)
                        else:
                            print(f"      Keine passende Nr Rep '{nr_rep}' in Metadatenliste gefunden")
                    else:
                        print("      Spalte 'Nr Rep' nicht in Metadatenliste gefunden")
            else:
                print(f"  Keine Matches gefunden")
        else:
            print("  Spalte 'Angaben zu Drucken/Editionen' nicht im Repertorium gefunden")
    
    print(f"\n=== ZUSAMMENFASSUNG ===")
    print(f"Suchstrings verarbeitet: {len(search_strings)}")
    print(f"Matches im Repertorium: {matches_found}")
    print(f"Aktualisierungen in Metadatenliste: {updates_made}")
    
    # Aktualisierte Metadatenliste speichern
    if updates_made > 0:
        output_file = "../datasheets/Imports/Urkunden_Metadatenliste_v7.xlsx"
        metadata_df.to_excel(output_file, index=False)
        print(f"\nAktualisierte Metadatenliste gespeichert als: {output_file}")
    else:
        print("\nKeine Änderungen vorgenommen, keine neue Datei erstellt.")

if __name__ == "__main__":
    match_archive_editions()