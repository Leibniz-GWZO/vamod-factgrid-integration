import pandas as pd
import openpyxl
from openpyxl import load_workbook
import re

def add_regex_sheet():
    """
    Fügt eine neue Mappe 'Regex' zur LLM_Evaluierung.xlsx hinzu und kopiert
    spezifische Spalten aus der Urkunden_Repertorium_v2.xlsx basierend auf
    übereinstimmenden 'Nr. Rep' Werten.
    """
    
    # Dateipfade
    llm_file = "LLM_Evaluierung.xlsx"
    source_file = "../datasheets/Urkunden_Repertorium_v2.xlsx"
    
    try:
        # 1. Ground_Truth Mappe aus LLM_Evaluierung.xlsx laden
        print("Lade Ground_Truth Daten...")
        ground_truth_df = pd.read_excel(llm_file, sheet_name="Ground_Truth")
        
        # Nr. Rep Werte aus Ground_Truth extrahieren
        if "Nr. Rep" not in ground_truth_df.columns:
            raise ValueError("Spalte 'Nr. Rep' nicht in Ground_Truth gefunden!")
        
        nr_rep_values = ground_truth_df["Nr. Rep"].dropna().unique()
        print(f"Gefundene Nr. Rep Werte: {len(nr_rep_values)}")
        
        # 2. Quell-Datei laden
        print("Lade Urkunden_Repertorium_v2.xlsx...")
        source_df = pd.read_excel(source_file)
        
        # 3. Relevante Spalten identifizieren
        columns_to_copy = []
        
        # Alle Spalten durchgehen und nach den gewünschten Mustern suchen
        for col in source_df.columns:
            # Genannte Person mit Index (Genannte Person 1, Genannte Person 2, etc.)
            if re.match(r"Genannte Person \d+", str(col)):
                columns_to_copy.append(col)
            # Empfängersitz mit Index
            elif re.match(r"Empfängersitz \d+", str(col)):
                columns_to_copy.append(col)
            # Objektort mit Index
            elif re.match(r"Objektort \d+", str(col)):
                columns_to_copy.append(col)
            # Nr. Rep wird auch benötigt für die Filterung
            elif col == "Nr. Rep":
                columns_to_copy.append(col)
        
        print(f"Gefundene relevante Spalten: {columns_to_copy}")
        
        # 4. Nur die Spalten auswählen, die existieren
        available_columns = [col for col in columns_to_copy if col in source_df.columns]
        
        if not available_columns:
            raise ValueError("Keine der gewünschten Spalten in der Quelldatei gefunden!")
        
        # 5. Daten basierend auf Nr. Rep filtern
        print("Filtere Daten basierend auf Nr. Rep Werten...")
        filtered_df = source_df[source_df["Nr. Rep"].isin(nr_rep_values)][available_columns]
        
        print(f"Gefilterte Zeilen: {len(filtered_df)}")
        
        # 6. Neue Mappe zur LLM_Evaluierung.xlsx hinzufügen
        print("Füge neue 'Regex' Mappe hinzu...")
        
        # Workbook laden
        with pd.ExcelWriter(llm_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            # Wenn 'Regex' Sheet bereits existiert, wird es ersetzt
            filtered_df.to_excel(writer, sheet_name='Regex', index=False)
        
        print("✅ Erfolgreich abgeschlossen!")
        print(f"Neue Mappe 'Regex' mit {len(filtered_df)} Zeilen und {len(available_columns)} Spalten erstellt.")
        print(f"Kopierte Spalten: {available_columns}")
        
        # Optional: Übersicht der kopierten Daten anzeigen
        if not filtered_df.empty:
            print("\nÜbersicht der ersten 5 Zeilen:")
            print(filtered_df.head())
        
    except FileNotFoundError as e:
        print(f"❌ Datei nicht gefunden: {e}")
        print("Bitte überprüfen Sie die Dateipfade:")
        print(f"- {llm_file}")
        print(f"- {source_file}")
        
    except ValueError as e:
        print(f"❌ Wertfehler: {e}")
        
    except Exception as e:
        print(f"❌ Unerwarteter Fehler: {e}")

def show_available_columns():
    """
    Hilfsfunktion um verfügbare Spalten in beiden Dateien anzuzeigen
    """
    try:
        print("=== Verfügbare Spalten in LLM_Evaluierung.xlsx (Ground_Truth) ===")
        df1 = pd.read_excel("LLM_Evaluierung.xlsx", sheet_name="Ground_Truth")
        for i, col in enumerate(df1.columns, 1):
            print(f"{i}: {col}")
        
        print("\n=== Verfügbare Spalten in Urkunden_Repertorium_v2.xlsx ===")
        df2 = pd.read_excel("../datasheets/Urkunden_Repertorium_v2.xlsx")
        for i, col in enumerate(df2.columns, 1):
            print(f"{i}: {col}")
            
    except Exception as e:
        print(f"Fehler beim Anzeigen der Spalten: {e}")

if __name__ == "__main__":
    # Hauptfunktion ausführen
    add_regex_sheet()
    
    # Uncomment um verfügbare Spalten anzuzeigen:
    # show_available_columns()