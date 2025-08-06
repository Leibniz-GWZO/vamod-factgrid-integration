# Material Charters - FactGrid Aufbereitung

Diese Skripte bereiten Urkunden-Metadaten für den Import in FactGrid vor. Die Verarbeitung erfolgt schrittweise in der folgenden Reihenfolge:

## 1. add_label.py
Erstellt Labels und mappt Grunddaten aus dem Repertorium. Verknüpft Aussteller, Empfänger und Datum zu strukturierten Labels in vier Sprachen (DE, EN, PL, UKR). Mappt Repertoriumsdaten auf FactGrid-Felder.

**Input:** `Urkunden_Metadaten.xlsx`, `Urkunden_Repertorium_v4.xlsx`  
**Output:** `Urkunden_Metadatenliste_v2.xlsx`

## 2. add_descr_different_lang.py
Korrigiert Formatierungsprobleme in Labels und übersetzt deutsche Beschreibungen (Dde) in Englisch, Polnisch und Ukrainisch mittels Google Translate API.

**Input:** `Urkunden_Metadatenliste_v2.xlsx`  
**Output:** `Urkunden_Metadatenliste_v3.xlsx`

## 3. add_archive.py
Ergänzt Archiv- und Editionsinformationen aus der FactGrid-Archivliste. Verknüpft Überlieferungsangaben mit entsprechenden FactGrid-QIDs.

**Input:** `Urkunden_Metadatenliste_v3.xlsx`, `Archive_Editionen_FactGrid.xlsx`  
**Output:** `Urkunden_Metadatenliste_v4.xlsx`

## 4. add_place.py
Ergänzt Ortsdaten für Ausstellungsorte. Mappt Ortsnamen auf FactGrid-QIDs basierend auf der Ortsidentifikationsliste.

**Input:** `Urkunden_Metadatenliste_v4.xlsx`, `Orte_Identifikation_factgrid.xlsx`  
**Output:** `Urkunden_Metadatenliste_v5.xlsx`

## 5. prepare_factgrid.py
Konvertiert Datumsformate in das FactGrid-Format (+YYYY-MM-DDTHH:MM:SSZ/precision) für den finalen Import.

**Input:** `Urkunden_Metadatenliste_v5.xlsx`  
**Output:** `Urkunden_Metadatenliste_v6.xlsx`

