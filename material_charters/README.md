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

## 3. add_prints.py
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

## 6. add_more_archives.py
Ergänzt weitere Archiv- und Editionsinformationen aus einer separaten FactGrid-Archivliste. Verknüpft Überlieferungsangaben mit entsprechenden FactGrid-QIDs, basierend auf den ersten 17 Zeilen der Archivliste.

**Input:** `Urkunden_Metadatenliste_v6.xlsx`, `Archive_Editionen_FactGrid.xlsx`  
**Output:** `Urkunden_Metadatenliste_v7.xlsx`

## 7. add_more_prints.py
Verarbeitet Editionen (Drucke/Editionen) und trägt Werte in die Metadatenliste ein. Sucht in der Repertoriumstabelle nach Einträgen und aktualisiert Spalten wie 'Publiziert in (P64)' und 'Nr. (qal90)' usw. Teilt Angaben an '=' auf und füllt zusätzliche Spalten.

**Input:** `Urkunden_Metadatenliste_v7.xlsx`, `Archive_Editionen_FactGrid.xlsx`  
**Output:** `Urkunden_Metadatenliste_v8.xlsx`

## 8. cleaning.py
Bereinigt Datumsangaben und Labels in der Metadatenliste. Parst deutsche Datumsstrings in ein standardisiertes Format und formatiert Labels in Deutsch, Englisch, Polnisch und Ukrainisch neu.

**Input:** `Urkunden_Metadatenliste_v8.xlsx`  
**Output:** `Urkunden_Metadatenliste_v9.xlsx`

## 9. remaining_labels.py
Füllt fehlende Labels (Lde, Len, Lpl, Lukr) aus dem Repertorium und übersetzt sie mit GPT-4o in Englisch, Polnisch und Ukrainisch. Extrahiert Aussteller, Empfänger und Datum für fehlende Einträge.

**Input:** `Urkunden_Metadatenliste_v9.xlsx`, `Urkunden_Repertorium_v4.xlsx`  
**Output:** `Urkunden_Metadatenliste_v10.xlsx`

## 10. add_comment.py
Fügt Kommentare aus dem Repertorium zur Metadatenliste hinzu. Merged Kommentare in die Notiz-Spalte basierend auf 'Nr. Rep'.

**Input:** `Urkunden_Metadatenliste_v10.xlsx`, `Urkunden_Repertorium_v4.xlsx`  
**Output:** `Urkunden_Metadatenliste_v10.xlsx` (aktualisiert)

## 11. add_archives_final.py
Finale Verknüpfung von Archiv-Editionen mit Urkunden-Metadaten. Sucht nach Matches in der Repertoriumstabelle, aktualisiert Standort- und Inventarpositionen, und füllt Notizen für leere Standorte. Löscht unnötige Spalten.

**Input:** `Urkunden_Metadatenliste_v10.xlsx`, `Archive_Editionen_FactGrid_SJ.xlsx`, `Urkunden_Repertorium_v4.xlsx`  
**Output:** `Urkunden_Metadatenliste_v11.xlsx`

## 12. add_region.py
Fügt administrative Zugehörigkeit basierend auf Mustern in der Zusammenfassung hinzu. Setzt QIDs für Regionen wie (TH), (TS), (TL), (TP).

**Input:** `Urkunden_Metadatenliste_v11.xlsx`  
**Output:** `Urkunden_Metadatenliste_v12.xlsx`

Erg