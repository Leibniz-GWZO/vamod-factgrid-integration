import pandas as pd
import sys
import datetime


def main():
    input_path = "../datasheets/Urkunden_gesamt_neu.xlsx"
    output_path = "../datasheets/personen/Personen_Empfänger_Austeller.xlsx"

    try:
        df = pd.read_excel(input_path, engine="openpyxl")
    except FileNotFoundError:
        print(f"Fehler: Datei nicht gefunden: {input_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Fehler beim Einlesen der Datei: {e}")
        sys.exit(1)

    # Sicherstellen, dass nötige Spalten existieren
    required_cols = [
        "Genannte Personen",
        "Rolle in Urkunde",
        "Funktion",
        "Geschlecht",
        "Edition",
        "Datum",
        "Aussteller"
    ]
    for col in required_cols:
        if col not in df.columns:
            print(f"Fehler: Spalte '{col}' nicht gefunden in der Datei.")
            sys.exit(1)

    # Filter nach relevanten Rollen anhand regulärem Ausdruck
    # Non-capturing group, um UserWarning zu vermeiden
    patterns = ["aussteller", "begünst", "empfänger", "käufer", "vorgänger", "mitaussteller"]
    role_series = df["Rolle in Urkunde"].astype(str).str.lower()
    regex = r"\b(?:" + "|".join(patterns) + r")"
    mask = role_series.str.contains(regex, na=False)

    df_filtered = df.loc[mask, required_cols].copy()

    # Personen extrahieren und aufsplitten
    df_filtered["Name"] = df_filtered["Genannte Personen"].astype(str)
    df_filtered["Name"] = df_filtered["Name"].str.split(r"[;]")

    df_exploded = df_filtered.explode("Name")
    df_exploded["Name"] = df_exploded["Name"].str.strip()

    # Neue Spalte 'Datum Funktion'
    df_exploded["Datum Funktion"] = df_exploded["Datum"]

    # Endgültige Spaltenauswahl inkl. Aussteller
    result = df_exploded[
        ["Name", "Funktion", "Rolle in Urkunde", "Geschlecht", "Edition", "Aussteller", "Datum Funktion"]
    ]

    # Speichern der neuen Excel-Datei
    result.to_excel(output_path, index=False)
    print(f"Neue Datei gespeichert: {output_path}")

    # Ungültige Datumseinträge ermitteln (nur Format)
    orig_dates = df["Datum"]
    normalized_dates = orig_dates.apply(
        lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else str(x)
    )
    invalid_dates = normalized_dates[~normalized_dates.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)].unique()

    if len(invalid_dates) > 0:
        print("\nDatumseinträge, die nicht dem Format YYYY-MM-DD entsprechen:")
        for d in invalid_dates:
            print(d)
    else:
        print("\nAlle Datumseinträge sind im Format YYYY-MM-DD.")


if __name__ == "__main__":
    main()
