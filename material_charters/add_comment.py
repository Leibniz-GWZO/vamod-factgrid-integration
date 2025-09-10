import pandas as pd

# Pfade zu den Dateien
metadata_path = "../datasheets/Imports/Urkunden_Metadatenliste_v10.xlsx"
repertorium_path = "../datasheets/Urkunden_Repertorium_v4.xlsx"

# Lade die Excel-Dateien
metadata_df = pd.read_excel(metadata_path)
metadata_df.drop(columns=['Datum_v9'], inplace=True)
repertorium_df = pd.read_excel(repertorium_path)

# Merge der DataFrames über "Nr. Rep"
merged_df = pd.merge(metadata_df, repertorium_df[['Nr. Rep', 'Kommentar']], on='Nr. Rep', how='left')

# Übertrage Kommentare in "Notiz (P73), alles was im Rep. als Kommentar aufgeführt ist", falls vorhanden
merged_df['Notiz (P73), alles was im Rep. als Kommentar aufgeführt ist'] = merged_df['Notiz (P73), alles was im Rep. als Kommentar aufgeführt ist'].fillna('') + merged_df['Kommentar'].fillna('')

# Entferne die temporäre Kommentar-Spalte
merged_df.drop(columns=['Kommentar'], inplace=True)

# Speichere die aktualisierte Metadatenliste
merged_df.to_excel(metadata_path, index=False)

print("Kommentare erfolgreich hinzugefügt und Datei gespeichert.")