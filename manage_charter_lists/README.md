Für die Erstellung der Urkundenliste zum Import in Factgrid sind zunächst zwei Dateien wichtig: die Excel-Tabelle *180809_Urkunden_Personen_korrekt_Sortiert2.xlsx* und die PDF-Datei *Jaros_Repertorium_kommentiert.pdf*.  

Beide Dateien enthalten wichtige, sich ergänzende Informationen. Die PDF-Datei dient als strukturierter Ausgangspunkt, da sie normierte Namens- und Ortsschreibweisen sowie Informationen zur Überlieferung, zu Kommentaren und zur Originalität der Urkunde enthält. Die Excel-Datei liefert ergänzend Angaben zu den Rollen der genannten Personen sowie zu ihren Ämtern zum Ausstellungszeitpunkt.

# 1. normalize_charter_sheets.py

Dieses Skript sortiert und normalisiert die Informationen aus der Excel-Datei *180809_Urkunden_Personen_korrekt_Sortiert2.xlsx*. Da die Datei mehrere Mappen mit unterschiedlichen Spaltenbezeichnungen enthält, wird eine einheitliche Struktur geschaffen und in eine neue Liste mit nur einer Mappe überführt. Zusätzlich wird die Datierungssicherheit überprüft: Ist z. B. das Datum als *1402-02-06/07?* angegeben, wird es in *1402-01-01* umgewandelt, und in einer zusätzlichen Spalte wird die ursprüngliche Angabe als „unsicher“ markiert. Das Ergebnis ist die neue Datei *Urkunden_gesamt_neu.xlsx*.

# 2. extract_repertorium_entries.py

Dieses Skript extrahiert strukturierte Informationen aus dem Repertorium historischer Urkunden im PDF-Format (*Jaros_Repertorium_kommentiert.pdf*). Die Daten werden gelesen, geparst und in eine Excel-Tabelle überführt. Pro Urkunde werden folgende Informationen ausgelesen:

- Nr. Rep  
- Datum  
- Datumssicherheit  
- Sicherheit Status Urkunde  
- Ausstellungsort  
- Regest  
- Angaben zur Überlieferung  
- Angaben zu Drucken/Editionen  
- Kommentar  
- Begünstigte  
- Objektorte  
- Aussteller  

Die Trennung der Informationen erfolgt mithilfe regulärer Ausdrücke.

# 3. parse_place_person_register_pdf.py

Dieses Skript extrahiert aus einem zweispaltigen PDF (*orts_namen_register.pdf*) strukturierte Informationen zu Orts- und Personennamen und speichert sie in eine Excel-Datei (*orts_namen_register.xlsx*). Diese Liste wird in Schritt 4 verwendet.

# 4. enrich_charters_with_places_and_persons.py

Dieses Skript ergänzt eine Excel-Tabelle historischer Urkunden (*Urkunden_Repertorium.xlsx*) durch automatisch erkannte:

- Aussteller  
- Begünstigte Personen  
- Objektorte  

→ Die Zuordnung erfolgt mithilfe des Referenzregisters mit Orts- und Personennamen (*orts_namen_register.xlsx*, siehe Schritt 3).

Die Idee ist, mithilfe des im Repertorium enthaltenen Orts- und Namensregisters sowie der Ortsliste aus dem Regest alle Entitäten zu extrahieren. Durch den Abgleich mit der Ortsliste werden Orte identifiziert, während verbleibende Entitäten als Namen bzw. Begünstigte im Regest erfasst werden. Dieser Schritt ist fehleranfällig und soll perspektivisch durch den Einsatz eines LLMs verbessert werden.

