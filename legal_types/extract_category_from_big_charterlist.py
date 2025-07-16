import pandas as pd
import os

try:
    # Prüfen ob die Excel-Datei existiert
    excel_path = "../datasheets/Urkunden_gesamt_neu.xlsx"
    if not os.path.exists(excel_path):
        print(f"FEHLER: Datei '{excel_path}' nicht gefunden!")
        print(f"Aktuelles Verzeichnis: {os.getcwd()}")
        print("Verfügbare Dateien im aktuellen Verzeichnis:")
        for file in os.listdir("."):
            print(f"  - {file}")
        exit()
    
    print(f"Excel-Datei gefunden: {excel_path}")
    
    # Excel-Datei laden
    df = pd.read_excel(excel_path)
    print(f"Excel-Datei geladen. Anzahl Zeilen: {len(df)}")
    print(f"Verfügbare Spalten: {list(df.columns)}")
    
    # Prüfen ob "Betreff" Spalte existiert
    if "Betreff" not in df.columns:
        print("FEHLER: Spalte 'Betreff' nicht gefunden!")
        print("Verfügbare Spalten:", list(df.columns))
        exit()
    
    # Unique Werte aus der Spalte "Betreff" extrahieren (ohne leere Werte)
    unique_betreff = df["Betreff"].dropna().unique()
    print(f"Gefundene unique Werte: {len(unique_betreff)}")
    
    # Werte sortieren
    unique_sorted = sorted(unique_betreff)
    
    # In Textdatei speichern
    output_file = "Betreffe_große_Urkundenliste.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        for wert in unique_sorted:
            f.write(str(wert) + "\n")
    
    # Prüfen ob Datei erstellt wurde
    if os.path.exists(output_file):
        file_size = os.path.getsize(output_file)
        print(f"✓ Erfolgreich {len(unique_sorted)} unique Werte in '{output_file}' gespeichert.")
        print(f"✓ Dateigröße: {file_size} Bytes")
        print(f"✓ Datei gespeichert in: {os.path.abspath(output_file)}")
    else:
        print("FEHLER: Textdatei wurde nicht erstellt!")

except FileNotFoundError as e:
    print(f"FEHLER: Datei nicht gefunden - {e}")
except PermissionError as e:
    print(f"FEHLER: Keine Berechtigung - {e}")
except Exception as e:
    print(f"FEHLER: {e}")
    import traceback
    traceback.print_exc()