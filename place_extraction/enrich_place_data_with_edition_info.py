import pandas as pd

# Einlesen der Excel-Dateien
# Stelle sicher, dass sich beide Dateien im gleichen Verzeichnis wie das Skript befinden oder passe die Pfade entsprechend an.
df_ident = pd.read_excel("datasheets/orte/Orte_Identifikation_merged.xlsx")
df_orts = pd.read_excel("datasheets/Ortsdaten_Repertorium-aufbereitet.xlsx")

def lookup_info(row):
    """
    Für eine gegebene Zeile aus df_ident:
    - Extrahiere die Werte aus 'Schreibweise Ortsregister' und 'alternative/historische Schreibweise'.
    - Bei 'alternative/historische Schreibweise' werden durch Semikolon getrennte Werte aufgeteilt.
    - Suche in df_orts nach einem passenden Eintrag in der Spalte 'Ort' und
      übernehme die Spalten 'Edition' und 'Nr. Rep. Druck', wenn ein Treffer gefunden wird.
    - Wird kein Treffer erzielt, werden die neuen Spalten mit None belegt.
    """
    # Liste der möglichen Namen zusammenstellen
    candidates = []
    if pd.notnull(row.get("Schreibweise Ortsregister")):
        candidates.append(row["Schreibweise Ortsregister"].strip())
    if pd.notnull(row.get("alternative/historische Schreibweise")):
        # Aufteilen, falls mehrere Werte vorhanden sind (durch Semikolon getrennt)
        parts = row["alternative/historische Schreibweise"].split(";")
        candidates.extend([part.strip() for part in parts if part.strip()])
    
    # Suche in df_orts nach einem passenden Eintrag
    for candidate in candidates:
        # Hier wird ein exakter Vergleich durchgeführt. Sollte ein unsicherer Abgleich benötigt werden,
        # kann man die Strings auch z.B. in Kleinbuchstaben umwandeln.
        match = df_orts[df_orts["Ort"] == candidate]
        if not match.empty:
            # Es wird der erste gefundene Treffer übernommen.
            edition = match.iloc[0]["Edition"] if "Edition" in match.columns else None
            nr_rep = match.iloc[0]["Nr. Rep. Druck"] if "Nr. Rep. Druck" in match.columns else None
            return pd.Series({"Edition": edition, "Nr. Rep. Druck": nr_rep})
    
    # Falls kein Treffer gefunden wird, gebe None zurück.
    return pd.Series({"Edition": None, "Nr. Rep. Druck": None})

# Wende die lookup-Funktion zeilenweise auf df_ident an.
lookup_results = df_ident.apply(lookup_info, axis=1)

# Füge die neuen Spalten 'Edition' und 'Nr. Rep. Druck' an den vorhandenen DataFrame an.
df_ident = pd.concat([df_ident, lookup_results], axis=1)

# Speichere den erweiterten DataFrame als neue Excel-Datei ab.
df_ident.to_excel("datasheets/orte/Orte_Identifikation_merged.xlsx", index=False)

print("Die neue Excel-Datei 'Orte_Identifikation_merged.xlsx' wurde erfolgreich erstellt.")
