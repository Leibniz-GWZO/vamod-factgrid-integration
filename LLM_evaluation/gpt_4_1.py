import os
import pandas as pd
from openai import OpenAI
import re

# -----------------------------------------------------------------------------
# 1. API-Konfiguration
# -----------------------------------------------------------------------------
# (a) API-Key direkt gesetzt
api_key = "API KEY"
model = "gpt-4.1" 

# (b) OpenAI-Client initialisieren
client = OpenAI(
    api_key=api_key
)

# -----------------------------------------------------------------------------
# 2. Excel-Datei einlesen aus der Ground_Truth Mappe
# -----------------------------------------------------------------------------
excel_path = "LLM_Evaluierung.xlsx"

# Lade die Ground_Truth Mappe
df = pd.read_excel(excel_path, sheet_name="Ground_Truth")

# (a) Prüfen, ob die Spalten "Nr. Rep" und "Regest" existieren
required_cols = ["Nr. Rep", "Regest"]
for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"Spalte '{col}' nicht gefunden in Ground_Truth Mappe von {excel_path}")

# (b) Nur Zeilen mit nicht-leeren Regesten verwenden
df_nonempty = df.dropna(subset=["Regest"])
print(f"Verarbeite {len(df_nonempty)} Regesten...")

# -----------------------------------------------------------------------------
# 3. Funktion zum Parsen der LLM-Ausgabe
# -----------------------------------------------------------------------------
def parse_llm_output(llm_output):
    """
    Parst die LLM-Ausgabe und extrahiert Objektorte, Empfängersitze und Genannte Personen
    """
    result = {
        'objektorte': [],
        'empfaengersitze': [],
        'genannte_personen': []
    }
    
    lines = llm_output.split('\n')
    
    for line in lines:
        line = line.strip()
        
        # Objektorte extrahieren
        if line.startswith('- Objektorte:'):
            content = line.replace('- Objektorte:', '').strip()
            if content and content.lower() != 'none':
                result['objektorte'] = [item.strip() for item in content.split(';') if item.strip()]
        
        # Empfängersitze extrahieren
        elif line.startswith('- Empfängersitze:'):
            content = line.replace('- Empfängersitze:', '').strip()
            if content and content.lower() != 'none':
                result['empfaengersitze'] = [item.strip() for item in content.split(';') if item.strip()]
        
        # Genannte Personen extrahieren
        elif line.startswith('- Genannte Personen:'):
            content = line.replace('- Genannte Personen:', '').strip()
            if content and content.lower() != 'none':
                result['genannte_personen'] = [item.strip() for item in content.split(';') if item.strip()]
    
    return result

# -----------------------------------------------------------------------------
# 4. Für jede Regest: Prompt bauen, API-Aufruf, Parsing
# -----------------------------------------------------------------------------
results_data = []

for idx, row in enumerate(df_nonempty.itertuples(index=False), start=1):
    nr_rep = getattr(row, 'Nr. Rep', row[0])  # Sichere Extraktion der Nr. Rep
    regest_text = row.Regest

    # (a) Prompt mit zusätzlichen Hinweisen und expliziten Newlines
    prompt = (
    f"Bitte extrahiere aus dem folgenden Regest \"{regest_text}\" die folgenden Informationen:\n"
    "- Objektorte: Konkrete Ortsnamen.\n"
    "- Empfängersitze: Ortsnamen, die den Sitz oder Aufenthaltsort von Empfängern oder beteiligten Parteien angeben.\n"
    "- Genannte Personen: Vollständige Namen von Personen, die im Text erwähnt werden.\n"
    "\n"
    "Hinweise:\n"
    "- Der Aussteller, der immer am Anfang des Regests steht, darf nicht unter 'Genannte Personen' auftauchen. Auch seine im Namen enthaltene Ortsinformationen sollen ignoriert werden.\n"
    "- Ignoriere in Klammern stehende Ortsnamen.\n"
    "- Falls \"von\" oder \"de\" Teil des Namens ist (z. B. \"Pelka von Służów\"), extrahiere den gesamten Ausdruck \"Pelka von Służów\" als genannte Person und erkenne zusätzlich hinter \"von\"/\"de\" den Empfängersitz (z. B. \"Służów\"). 'Służów' darf nicht als Objektort erscheinen.\n"
    "- Der Empfängersitz muss aus dem Kontext abgeleitet werden. Ein Hinweis ist \"von XXX\", wobei XXX für eine Ortsbezeichnung steht.\n"
    "- Falls eine Person mit einem Zusatz davor wie \"Sohn des XXX\" oder \"Tochter des YYY\" oder \"Bruder des ZZZZ\" oder \"Vater des WWW\" erwähnt wird dann extrahiere den Namen nicht.\n"
    "- Wenn mehrere Personen derselben Familie aufgezählt werden und das Suffix (\"von\"/\"de\") nur nach dem letzten Namen erscheint (z. B. 'Johannes, Clemens und Spytek von Siemuszowa'), hänge das Suffix an **jeden** zuvor genannten Namen an, sodass 'Johannes von Siemuszowa', 'Clemens von Siemuszowa' und 'Spytek von Siemuszowa' extrahiert werden.\n"
    "- Wenn eine Person mit lateinischen Funktionsbezeichnungen oder beschreibenden Ausdrücken versehen ist (z. B. 'dictus Drewientha', 'heres de Borek'), extrahiere nur den Eigennamen (z. B. nur 'Stanislaus').\n"
    "- Wenn eine Person mit einem beschreibenden Zusatz wie \"dem Älteren\", \"dem Jüngeren\", \"dem Großen\" o. Ä. versehen ist, extrahiere nur den Eigennamen ohne diesen Zusatz (z. B. nur 'Vaško Teptukovyč')."
    "- Die Ortsbezeichnung hinter dem Aussteller, welcher der Name ist, der am Anfang des Regestes auftauch, darf weder als Objektort noch Empfängersitz extrahiert werden.\n"
    "- Ignoriere Ortsbezeichnungen, die Teil von Amtsbezeichnungen sind, insbesondere wenn sie in Latein verfasst sind.\n"
     "- Falls eine Kategorie keine Treffer enthält, gebe \"None\" zurück.\n"
    "\n"
    "Ausgabeformat (ohne zusätzliche Kommentare):\n"
    "- Objektorte: Extrahierte Objektorte durch \";\" getrennt\n"
    "- Empfängersitze: Extrahierte Empfängersitze durch \";\" getrennt\n"
    "- Genannte Personen: Extrahierte Genannte Personen durch \";\" getrennt\n"
    )

    # (b) API-Aufruf an /chat/completions mit deterministischen Parametern
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "Du bist ein hilfreicher Assistent."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.0,   
            max_tokens=512
        )

        # (c) Extrahiere den reinen Antworttext des Modells
        llm_output = response.choices[0].message.content.strip()

        # (d) Parse die LLM-Ausgabe
        parsed_result = parse_llm_output(llm_output)

        # (e) Ausgabe zur Kontrolle
        print(f"\n----- Regest #{idx} (Nr. Rep: {nr_rep}) -----")
        print(f"{regest_text}\n")
        print(">>> LLM-Antwort:")
        print(llm_output)
        print(f">>> Geparste Ergebnisse:")
        print(f"Objektorte: {parsed_result['objektorte']}")
        print(f"Empfängersitze: {parsed_result['empfaengersitze']}")
        print(f"Genannte Personen: {parsed_result['genannte_personen']}")
        print("----------------------------\n")

        # (f) Bereite Daten für die Tabelle vor
        row_data = {'Nr. Rep': nr_rep}
        
        # Objektorte als separate Spalten
        max_objektorte = max(len(parsed_result['objektorte']), 1)
        for i in range(max_objektorte):
            if i < len(parsed_result['objektorte']):
                row_data[f'Objektort {i+1}'] = parsed_result['objektorte'][i]
            else:
                row_data[f'Objektort {i+1}'] = ''
        
        # Empfängersitze als separate Spalten
        max_empfaengersitze = max(len(parsed_result['empfaengersitze']), 1)
        for i in range(max_empfaengersitze):
            if i < len(parsed_result['empfaengersitze']):
                row_data[f'Empfängersitz {i+1}'] = parsed_result['empfaengersitze'][i]
            else:
                row_data[f'Empfängersitz {i+1}'] = ''
        
        # Genannte Personen als separate Spalten
        max_personen = max(len(parsed_result['genannte_personen']), 1)
        for i in range(max_personen):
            if i < len(parsed_result['genannte_personen']):
                row_data[f'Genannte Person {i+1}'] = parsed_result['genannte_personen'][i]
            else:
                row_data[f'Genannte Person {i+1}'] = ''
        
        results_data.append(row_data)

    except Exception as e:
        print(f"Fehler bei Regest {nr_rep}: {e}")
        # Füge leere Zeile hinzu bei Fehler
        results_data.append({'Nr. Rep': nr_rep})

# -----------------------------------------------------------------------------
# 5. Ergebnisse in neue Excel-Mappe "LLM Gpt_4.1" speichern
# -----------------------------------------------------------------------------
if results_data:
    # Erstelle DataFrame aus den Ergebnissen
    results_df = pd.DataFrame(results_data)
    
    # Lade die bestehende Excel-Datei
    with pd.ExcelWriter(excel_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        results_df.to_excel(writer, sheet_name='LLM Gpt_4.1', index=False)
    
    print(f"\nErgebnisse wurden in die Mappe 'LLM Gpt_4.1' der Datei {excel_path} gespeichert.")
    print(f"Anzahl verarbeiteter Regesten: {len(results_data)}")
    print(f"Spalten in der Ergebnistabelle: {list(results_df.columns)}")
else:
    print("Keine Ergebnisse zum Speichern vorhanden.")