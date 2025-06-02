import pandas as pd
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage

# =============================================================================
# Teil 1: Extraktion kursiv formatierter Segmente aus der PDF
# =============================================================================
def extract_italic_segments(pdf_path, start_page=563, debug=False):
    """
    Extrahiert ab der Seite 'start_page' aus der PDF alle Textsegmente,
    die ausschließlich aus kursiv formatierten Zeichen bestehen.
    """
    italic_segments = []  # Liste der extrahierten kursiven Segmente
    with open(pdf_path, 'rb') as fp:
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        
        current_page = 1
        for page in PDFPage.get_pages(fp):
            if current_page < start_page:
                current_page += 1
                continue

            interpreter.process_page(page)
            layout = device.get_result()
            page_segments = []  # Segmente der aktuellen Seite

            # Durchlaufe alle Elemente des Seitenlayouts
            for element in layout:
                if isinstance(element, LTTextBox):
                    for text_line in element:
                        current_segment = ""
                        # Gehe alle Objekte in der Textzeile durch
                        for obj in text_line:
                            if isinstance(obj, LTChar):
                                # Bestimme den Fontnamen in Kleinbuchstaben
                                fontname = obj.fontname.lower() if hasattr(obj, "fontname") else ""
                                # Prüfe, ob das Zeichen kursiv ist (enthält "italic" oder "oblique")
                                if "italic" in fontname or "oblique" in fontname:
                                    current_segment += obj.get_text()
                                else:
                                    # Sobald ein nicht-kursives Zeichen kommt, wird das Segment abgeschlossen.
                                    if current_segment:
                                        page_segments.append(current_segment.strip())
                                        current_segment = ""
                            elif isinstance(obj, LTAnno):
                                # LTAnno liefert meist Leerzeichen oder Zeilenumbrüche.
                                if obj.get_text().isspace():
                                    if current_segment and not current_segment.endswith(" "):
                                        current_segment += " "
                                else:
                                    if current_segment:
                                        page_segments.append(current_segment.strip())
                                        current_segment = ""
                        if current_segment:
                            page_segments.append(current_segment.strip())
            
            if debug:
                print(f"Seite {current_page} - erkannte kursiv geschriebene Segmente: {page_segments}")
            italic_segments.extend(page_segments)
            current_page += 1
    return italic_segments

# PDF-Dateipfad anpassen
pdf_file = "datasheets/Jaros-Iterationen-2021-klein.pdf"
# Extrahiere kursiv formatierte Segmente ab Seite 539 (Debug-Output aktiviert)
italic_segments = extract_italic_segments(pdf_file, start_page=539, debug=True)
# Für den Vergleich: Alle Segmente in Kleinbuchstaben umwandeln
italic_segments_lower = {seg.lower() for seg in italic_segments}

# =============================================================================
# Teil 2: Excel-Verarbeitung & Aggregation der Ortsdaten
# =============================================================================
def unique_concat(series):
    """
    Vereinigt alle nicht-leeren, eindeutigen Werte einer Series zu einem einzigen String,
    getrennt durch ein Semikolon.
    """
    values = series.dropna().unique()
    return ";".join(map(str, values))

# Excel-Datei einlesen (Ortsdaten zur Aggregation)
excel_input = "datasheets/Ortsdaten_Repertorium-aufbereitet.xlsx"
df = pd.read_excel(excel_input)

print("\nVerfügbare Spalten in der Excel-Datei:")
print(df.columns.tolist())

# Definiere die Spalten, die aggregiert werden sollen
spalten_agg = ["Region", "alternative/historische Schreibweisen", "Identifikation"]

# Gruppiere nach "Ort" und aggregiere die angegebenen Spalten
df_aggregiert = df.groupby("Ort", as_index=False).agg({spalte: unique_concat for spalte in spalten_agg})
df_aggregiert.sort_values("Ort", inplace=True)

print("\nErgebnis der Aggregation (erste Zeilen):")
print(df_aggregiert.head())

# Erstelle eine Hilfsspalte mit getrimmten, in Kleinbuchstaben umgewandelten Ortsnamen
df_aggregiert["Vergleichswert"] = df_aggregiert["Ort"].astype(str).str.strip().str.lower()

# Setze in "Status Identifikation" den Wert "nicht möglich", falls der Ortsname in den extrahierten Segmenten vorkommt
df_aggregiert["Status Identifikation"] = df_aggregiert["Vergleichswert"].apply(
    lambda x: "nicht möglich" if x in italic_segments_lower else ""
)
df_aggregiert.drop("Vergleichswert", axis=1, inplace=True)

# Speichere das aggregierte Ergebnis in eine Excel-Datei (wird später überschrieben)
output_file = "datasheets/orte/Orte_Identifikation.xlsx"
df_aggregiert.to_excel(output_file, index=False)
print(f"\nDie aggregierte Liste mit Status wurde erfolgreich in '{output_file}' gespeichert.")

# =============================================================================
# Teil 3: Merge mit zusätzlichen Identifikationsdaten (mehrere Mappen) & erneute Aggregation
# =============================================================================
# Lese die gerade erstellte Excel-Datei (Orte_Identifikation) ein
df_orte = pd.read_excel(output_file)

# Lese die Identifikationsdaten aus der Datei, welche mehrere Mappen enthält.
ident_file = "datasheets/180322_Identifikation_Objektorte3.xlsx"
sheets = ["KIII", "WO", "WII"]
df_list = []

for sheet in sheets:
    temp_df = pd.read_excel(ident_file, sheet_name=sheet, header=2)
    # Bereinige die Spaltennamen (entfernt führende und nachfolgende Leerzeichen)
    temp_df.columns = temp_df.columns.str.strip()
    df_list.append(temp_df)

# Fasse alle Identifikationsdaten zusammen
df_ident = pd.concat(df_list, ignore_index=True)

# Merge: Vergleiche 'Ort' aus df_orte mit 'Objektort' aus df_ident.
# Hier werden nun zusätzlich die Spalten "Heutiger Name", "Geodaten" und "Lage" übernommen.
df_merged = pd.merge(
    df_orte, 
    df_ident[['Objektort', 'Heutiger Name', 'Geodaten', 'Lage']],
    left_on='Ort', 
    right_on='Objektort', 
    how='left'
)

# Entferne die Hilfsspalte 'Objektort'
df_merged = df_merged.drop(columns='Objektort')

# Nach dem Merge können für einen Ort mehrere Zeilen vorhanden sein.
# Gruppiere deshalb nach 'Ort' und aggregiere alle übrigen Spalten.
agg_dict = {col: unique_concat for col in df_merged.columns if col != 'Ort'}
df_final = df_merged.groupby("Ort", as_index=False).agg(agg_dict)

# Umbenennen der Spalte 'Ort' in 'Schreibweise Ortsregister'
df_final.rename(columns={'Ort': 'Schreibweise Ortsregister'}, inplace=True)

# Überschreibe dieselbe Excel-Datei mit dem finalen, aggregierten Ergebnis
df_final.to_excel(output_file, index=False)
print(f"\nDie finale aktualisierte und aggregierte Liste wurde in '{output_file}' gespeichert.")

# =============================================================================
# Statistiken aus der finalen Excel-Datei
# =============================================================================
# Anzahl der Zeilen mit "nicht möglich" in der Spalte Status Identifikation
count_status_nicht = df_final[df_final["Status Identifikation"] == "nicht möglich"].shape[0]

# Anzahl der Zeilen mit einem Eintrag in der Spalte Geodaten (ignoriere leere Strings)
count_geodaten = df_final["Geodaten"].astype(str).str.strip().replace("", pd.NA).notna().sum()

# Anzahl der Zeilen, in denen in der Spalte "Heutiger Name" ein Eintrag vorhanden ist, aber in der Spalte Geodaten kein Eintrag
count_heutiger_name_no_geodaten = df_final[
    (df_final["Heutiger Name"].astype(str).str.strip() != "") &
    (df_final["Geodaten"].astype(str).str.strip() == "")
].shape[0]

print("\n--- Statistiken der finalen Excel-Datei ---")
print(f"Zeilen mit 'nicht möglich' in Status Identifikation: {count_status_nicht}")
print(f"Zeilen mit Eintrag in Geodaten: {count_geodaten}")
print(f"Zeilen mit Eintrag in Heutiger Name und ohne Geodaten: {count_heutiger_name_no_geodaten}")
