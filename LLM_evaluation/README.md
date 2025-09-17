# LLM Evaluation - Entitätserkennung in historischen Urkunden

Dieses Verzeichnis enthält Skripte zur Evaluierung verschiedener Ansätze zur automatischen Entitätserkennung (Named Entity Recognition) in mittelalterlichen Urkundenregesten. Es werden sowohl Large Language Models (LLMs) als auch regelbasierte Ansätze getestet und verglichen.

## 1. extract_samples.py

Dieses Skript extrahiert zufällige Samples aus der Urkunden-Tabelle (Urkunden_Repertorium_v2.xlsx) für die Erstellung einer Ground Truth. Es wählt eine festgelegte Anzahl von Urkunden aus (standardmäßig 50) und extrahiert relevante Spalten wie "Nr. Rep", "Datum", "Regest", "Angaben zur Überlieferung" sowie alle Spalten mit "Genannte Person", "Empfängersitz" und "Objektort". Die Samples werden in einer neuen Excel-Datei (LLM_Evaluierung.xlsx) gespeichert und dienen als Referenzdatensatz für die Evaluierung der verschiedenen Entitätserkennungsansätze.

## 2. add_regex_entity_recognition.py

Dieses Skript fügt eine neue "Regex"-Mappe zur LLM_Evaluierung.xlsx hinzu und kopiert spezifische Spalten aus der Urkunden_Repertorium_v2.xlsx basierend auf übereinstimmenden "Nr. Rep"-Werten. Es identifiziert automatisch alle relevanten Spalten (Genannte Person, Empfängersitz, Objektort mit Indizes) und erstellt damit einen Baseline-Datensatz für den Vergleich mit den LLM-Ergebnissen. Diese Daten repräsentieren die regelbasierte Entitätserkennungen aus der Repertoriums PDF Datei und sollen zum Vergleich mit den LLMs dienen.

## 3. detect_entities.py

Dieses Skript führt die Entitätserkennung mit einem lokalen Large Language Model (llama-3.3-70b-instruct) über die SAIA Academic Cloud API durch (GWDG). Es verarbeitet die Regesten aus der Ground Truth und extrahiert systematisch drei Entitätstypen: Objektorte (konkrete Ortsnamen), Empfängersitze (Sitze von Empfängern oder beteiligten Parteien) und Genannte Personen (vollständige Personennamen). Das Skript verwendet einen detaillierten Prompt mit spezifischen Regeln zur Behandlung von Ausstellern, Klammern, lateinischen Bezeichnungen und Namensvarianten. Die Ergebnisse werden in einer neuen "LLM"-Mappe der Excel-Datei gespeichert.

## 4. gpt_4o.py

Dieses Skript funktioniert analog zu detect_entities.py, verwendet jedoch das OpenAI GPT-4o Modell für die Entitätserkennung. Es nutzt denselben systematischen Prompt und dieselben Extraktionsregeln wie das Llama-Skript, um eine faire Vergleichbarkeit der Modelle zu gewährleisten. Die API-Aufrufe erfolgen über die offizielle OpenAI API mit deterministischen Parametern (temperature=0.0) für reproduzierbare Ergebnisse. Die extrahierten Entitäten werden in der Mappe "LLM Gpt_4o" gespeichert.

## 5. gpt_4_1.py

Dieses Skript testet das ältere GPT-4.1 Modell von OpenAI unter denselben Bedingungen wie die anderen LLM-Skripte. Es verwendet identische Prompts und Extraktionsregeln, um eine konsistente Evaluierung über verschiedene Modellgenerationen hinweg zu ermöglichen. Die Ergebnisse werden in der Mappe "LLM Gpt_4.1" der Excel-Datei gespeichert und ermöglichen einen direkten Vergleich der Leistungsfähigkeit verschiedener GPT-Modellversionen bei der historischen Entitätserkennung.

## 6. evaluation.py

Dieses Skript führt eine comprehensive Evaluierung aller getesteten Ansätze durch. Es berechnet Precision, Recall und F1-Score für jeden Entitätstyp (Genannte Person, Empfängersitz, Objektort) und jeden getesteten Ansatz (Regex/manuelle Annotation, llama-3.3, gpt-4o, gpt-4.1). Das Skript führt einen direkten Vergleich der extrahierten Entitäten mit der Ground Truth durch und erstellt detaillierte Metriken sowohl auf Ebene einzelner Entitätstypen als auch als macro-averaged Gesamtbewertung. Die Ausgabe umfasst sowohl eine Zusammenfassung als auch detaillierte Ergebnisse für jeden Ansatz.

## Evaluierungsworkflow

1. **Datenextraktion**: `extract_samples.py` erstellt den Ground Truth Datensatz
2. **Baseline-Erstellung**: `add_regex_entity_recognition.py` fügt Ergebnisse der Entitätserkennung durch reguläre Ausdrücke hinzu
3. **LLM-Evaluierung**: `detect_entities.py`, `gpt_4o.py`, `gpt_4_1.py` führen Entitätserkennung durch
4. **Vergleichsanalyse**: `evaluation.py` berechnet und vergleicht die Leistungsmetriken

## Datenstruktur

- **LLM_Evaluierung.xlsx**: Zentrale Excel-Datei mit allen Daten und Ergebnissen
  - Ground_Truth: Referenzdatensatz mit 50 zufälligen Urkundensamples
  - Regex: extahierte Entitys durch reguläre Ausdrücke als Baseline
  - LLM llama-3.3: Ergebnisse des Llama-Modells
  - LLM Gpt_4o: Ergebnisse des GPT-4o Modells  
  - LLM Gpt_4.1: Ergebnisse des GPT-4.1 Modells

## Entitätstypen

- **Objektorte**: Konkrete geografische Orte, die in den Urkunden erwähnt werden
- **Empfängersitze**: Sitze oder Aufenthaltsorte von Empfängern und beteiligten Parteien
- **Genannte Personen**: Vollständige Namen aller in den Regesten erwähnten Personen (außer Ausstellern)