import pandas as pd
import re
from tqdm import tqdm

def main():
    # Pfade zu den Excel-Dateien
    path_v1 = "../datasheets/Urkunden_Repertorium_v1.xlsx"
    path_aufbereitet = "../datasheets/Ortsdaten_Repertorium-aufbereitet.xlsx"
    output_path = "../datasheets/Urkunden_Repertorium_v2.xlsx"
    
    # Einlesen der Excel-Dateien; dtype=str stellt sicher, dass alle Zellen als Strings gelesen werden
    df1 = pd.read_excel(path_v1, dtype=str)
    df2 = pd.read_excel(path_aufbereitet, dtype=str)

    # 1. In df1 alle Spalten identifizieren, die mit "Genannte Orte" beginnen
    genannte_cols = [col for col in df1.columns if col.startswith("Genannter Ort")]

    # 2. Für jede Zeile in df1 extrahieren wir:
    #    - den Repertorium-Schlüssel (z.B. "A123" oder "B045") aus der Spalte "Nr. Rep"
    #    - alle Ortsnamen aus den Spalten "Genannte Orte i"
    #    - wir ordnen jeden Ort entweder Empfängersitz oder Objektort zu, basierend auf df2
    
    all_recipients = []  # Liste von Listen mit Empfängersitzen pro Zeile
    all_objects = []     # Liste von Listen mit Objektorten pro Zeile

    for idx, row in tqdm(df1.iterrows()):
        # 2a. "Nr. Rep" auslesen und AXXX/BXXX extrahieren
        rep_cell = row.get("Nr. Rep", "")
        match = re.search(r"([AB]\d+)", str(rep_cell))
        if match:
            rep_key = match.group(1)
        else:
            rep_key = None

        # 2b. Alle Einträge der Spalten "Genannte Orte i" sammeln (ohne leere Zellen)
        gen_orte = []
        for col in genannte_cols:
            val = row.get(col, "")
            if pd.notnull(val) and str(val).strip() != "":
                gen_orte.append(str(val).strip())

        recipients = []
        objects = []

        # 2c. Wenn rep_key existiert, suchen wir alle Zeilen in df2 mit "Nr. Rep. Druck" == rep_key
        if rep_key is not None:
            df2_rep = df2[df2["Nr. Rep. Druck"] == rep_key]
            # Für jeden Ort im v1-Datensatz prüfen, ob er in df2_rep vorkommt
            for ort in gen_orte:
                # exakter Abgleich auf die Spalte "Ort" in df2_rep
                df2_match = df2_rep[df2_rep["Ort"] == ort]
                if not df2_match.empty:
                    # Wenn mindestens eine Zeile gefunden wurde, nehmen wir den ersten Eintrag
                    funktion = df2_match.iloc[0].get("Funktion in Urkunde?", "")
                    # Je nach Funktion entscheiden, ob Empfängersitz oder Objektort
                    if str(funktion).strip().lower() == "empfängersitz":
                        recipients.append(ort)
                    else:
                        objects.append(ort)
                else:
                    # Kein Match in df2: Ort wird als Objektort klassifiziert
                    objects.append(ort)
        else:
            # Falls kein gültiger rep_key gefunden wurde: alle Orte zu Objektorten
            objects = gen_orte.copy()

        all_recipients.append(recipients)
        all_objects.append(objects)

    # 3. Zwischenergebnisse in df1 einfügen (als temporäre Spalten)
    df1["_temp_recipients"] = all_recipients
    df1["_temp_objects"] = all_objects

    # 4. Die originalen "Genannte Orte i"-Spalten entfernen
    df1 = df1.drop(columns=genannte_cols)

    # 5. Ermitteln, wie viele Empfängersitze bzw. Objektorte die längsten Listen haben
    max_rec = max(len(lst) for lst in all_recipients)
    max_obj = max(len(lst) for lst in all_objects)

    # 6. Neue Spalten "Empfängersitz 1", "Empfängersitz 2", ..., "Objektort 1", ... erstellen
    for i in range(max_rec):
        colname = f"Empfängersitz {i+1}"
        df1[colname] = [
            lst[i] if len(lst) > i else ""  for lst in all_recipients
        ]

    for i in range(max_obj):
        colname = f"Objektort {i+1}"
        df1[colname] = [
            lst[i] if len(lst) > i else ""  for lst in all_objects
        ]

    # 7. Temporäre Spalten entfernen
    df1 = df1.drop(columns=["_temp_recipients", "_temp_objects"])

    # 8. Ergebnis als neue Excel-Datei speichern
    df1.to_excel(output_path, index=False)
    print(f"Die Datei wurde gespeichert unter: {output_path}")

if __name__ == "__main__":
    main()
