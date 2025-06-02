import pandas as pd

# Pfade zu den Dateien
aemter_path   = "../datasheets/aemter/Aemter_Matching_SJ_LP.xlsx"
personen_path = "../datasheets/personen/Personen_Empfaenger_Aussteller_aggregiert.xlsx"
output_path   = "../datasheets/aemter/Aemter_Receiver_Issuer.xlsx"

# 1. Aemter-Matching-Liste einlesen
df_aemter = pd.read_excel(
    aemter_path,
    sheet_name="Aufb. Ämter ohne Durchgestr.",
    usecols="C:K",
    header=0
)

# 2. Personen-Empfaenger-Liste einlesen
funktion_cols = [f"Funktion {i}" for i in range(1, 16)]
df_personen = pd.read_excel(
    personen_path,
    usecols=funktion_cols
)

# 3. Alle Funktionsbezeichnungen in eine Menge sammeln (lower & strip)
function_set = {
    func.strip().lower()
    for col in funktion_cols
    for cell in df_personen[col].dropna().astype(str)
    for func in cell.split(";")
}

# 4. Filterfunktion: 
#    – normalized row_parts: lower & strip
#    – match, wenn row_part == func oder row_part in func (Substring-Vergleich)
def row_matches_functions(row):
    # Alle Zellen der Zeile aufsplitten, strip & lower anwenden
    row_parts = [
        part.strip().lower()
        for cell in row.dropna().astype(str)
        for part in cell.split(";")
    ]
    for part in row_parts:
        if not part:
            continue
        # exakter Treffer in Personen-Funktionen
        if part in function_set:
            return True
        # Substring-Treffer: Aemter-Eintrag in Personen-Funktion
        for func in function_set:
            if part in func:
                return True
    return False

# 5. Zeilen filtern
df_filtered = df_aemter[df_aemter.apply(row_matches_functions, axis=1)]

# 6. Gefilterte Zeilen inklusive Header in neue Excel-Datei schreiben
df_filtered.to_excel(output_path, index=False, header=True)

print(f"Gefilterte Zeilen mit Original-Header wurden nach '{output_path}' geschrieben.")
