import pandas as pd
import random
import numpy as np
from pathlib import Path

def extract_ground_truth_samples(input_file, output_file, n_samples=50, max_row=553):
    """
    Extrahiert zufällige Samples aus der Urkunden-Tabelle für Ground Truth Erstellung.
    
    Args:
        input_file (str): Pfad zur Eingabedatei
        output_file (str): Pfad zur Ausgabedatei
        n_samples (int): Anzahl der zu extrahierenden Samples
        max_row (int): Maximale Zeilennummer (1-basiert)
    """
    
    try:
        # Excel-Datei einlesen
        print(f"Lade Datei: {input_file}")
        df = pd.read_excel(input_file)
        
        # Nur die ersten max_row-1 Zeilen berücksichtigen (0-basiert)
        df_filtered = df.iloc[:max_row-1].copy()
        
        print(f"Verfügbare Zeilen nach Filterung: {len(df_filtered)}")
        print(f"Spalten in der Datei: {list(df_filtered.columns)}")
        
        # Überprüfen ob genügend Zeilen vorhanden sind
        if len(df_filtered) < n_samples:
            print(f"Warnung: Nur {len(df_filtered)} Zeilen verfügbar, aber {n_samples} angefordert.")
            n_samples = len(df_filtered)
        
        # Zufällige Samples auswählen
        random.seed(42)  # Für Reproduzierbarkeit
        sample_indices = random.sample(range(len(df_filtered)), n_samples)
        df_samples = df_filtered.iloc[sample_indices].copy()
        
        # Relevante Spalten identifizieren
        required_columns = ['Nr. Rep', 'Datum', 'Regest', 'Angaben zur Überlieferung']
        
        # Dynamisch alle "Genannte Person" Spalten finden
        genannte_person_cols = [col for col in df_samples.columns if 'Genannte Person' in str(col)]
        
        # Dynamisch alle "Empfängersitz" Spalten finden
        empfaengersitz_cols = [col for col in df_samples.columns if 'Empfängersitz' in str(col)]
        
        # Dynamisch alle "Objektort" Spalten finden
        objektort_cols = [col for col in df_samples.columns if 'Objektort' in str(col)]
        
        # Alle gewünschten Spalten zusammenfassen
        desired_columns = required_columns + genannte_person_cols + empfaengersitz_cols + objektort_cols
        
        # Prüfen welche Spalten tatsächlich existieren
        existing_columns = [col for col in desired_columns if col in df_samples.columns]
        missing_columns = [col for col in desired_columns if col not in df_samples.columns]
        
        if missing_columns:
            print(f"Warnung: Folgende Spalten wurden nicht gefunden: {missing_columns}")
        
        print(f"Extrahierte Spalten: {existing_columns}")
        
        # Ground Truth Datensatz erstellen
        ground_truth = df_samples[existing_columns].copy()
        
        # Index zurücksetzen
        ground_truth.reset_index(drop=True, inplace=True)
        
        # Ausgabedatei speichern
        print(f"Speichere Ground Truth mit {len(ground_truth)} Samples nach: {output_file}")
        ground_truth.to_excel(output_file, index=False)
        
        # Statistiken ausgeben
        print("\n=== STATISTIKEN ===")
        print(f"Anzahl Samples: {len(ground_truth)}")
        print(f"Anzahl Spalten: {len(ground_truth.columns)}")
        
        # Analyse der Entitätsspalten
        print(f"\nGefundene 'Genannte Person' Spalten: {len(genannte_person_cols)}")
        for col in genannte_person_cols[:5]:  # Zeige erste 5
            non_empty = ground_truth[col].notna().sum()
            print(f"  - {col}: {non_empty}/{len(ground_truth)} ausgefüllt")
        
        print(f"\nGefundene 'Empfängersitz' Spalten: {len(empfaengersitz_cols)}")
        for col in empfaengersitz_cols[:5]:  # Zeige erste 5
            non_empty = ground_truth[col].notna().sum()
            print(f"  - {col}: {non_empty}/{len(ground_truth)} ausgefüllt")
            
        print(f"\nGefundene 'Objektort' Spalten: {len(objektort_cols)}")
        for col in objektort_cols[:5]:  # Zeige erste 5
            non_empty = ground_truth[col].notna().sum()
            print(f"  - {col}: {non_empty}/{len(ground_truth)} ausgefüllt")
        
        # Beispiel-Einträge anzeigen
        print("\n=== BEISPIEL-EINTRÄGE ===")
        for i in range(min(3, len(ground_truth))):
            print(f"\nSample {i+1}:")
            print(f"  Nr. Rep: {ground_truth.iloc[i].get('Nr. Rep', 'N/A')}")
            print(f"  Datum: {ground_truth.iloc[i].get('Datum', 'N/A')}")
            print(f"  Regest: {str(ground_truth.iloc[i].get('Regest', 'N/A'))[:100]}...")
            print(f"  Überlieferung: {str(ground_truth.iloc[i].get('Angaben zur Überlieferung', 'N/A'))[:50]}...")
        
        return ground_truth
        
    except FileNotFoundError:
        print(f"Fehler: Datei '{input_file}' nicht gefunden.")
        return None
    except Exception as e:
        print(f"Fehler beim Verarbeiten der Datei: {str(e)}")
        return None

# Hauptfunktion
if __name__ == "__main__":
    # Pfade definieren
    input_file = "../datasheets/Urkunden_Repertorium_v2.xlsx"
    output_file = "Urkunden_Ground_Truth.xlsx"
    
    # Ground Truth Samples extrahieren
    result = extract_ground_truth_samples(
        input_file=input_file,
        output_file=output_file,
        n_samples=50,
        max_row=553
    )
    
    if result is not None:
        print(f"\n✅ Ground Truth erfolgreich erstellt: {output_file}")
    else:
        print("\n❌ Fehler beim Erstellen der Ground Truth")