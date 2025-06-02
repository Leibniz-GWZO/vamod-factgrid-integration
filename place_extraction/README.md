# 1. build_place_dataset.py

In diesem Schritt werden die Ortsdaten aus der Datei Ortsdaten_Repertorium-aufbereitet.xlsx geladen, wobei insbesondere die Spalten „Region“, „alternative/historische Schreibweisen“ und „Identifikation“ berücksichtigt werden. Die Datei enthält alle Orte, die im Urkundenrepertorium vorkommen. Zusätzlich werden Informationen aus der Datei 180322_Identifikation_Objektorte3.xlsx ergänzt, insbesondere die Spalten „heutiger Name“, „Geodaten“ und „Lage“. Darüber hinaus werden aus der Orts- und Personenliste der Dissertation alle kursiv geschriebenen Wörter extrahiert – in der Regel nicht identifizierte Orte. Tauchen diese im neu erstellten Datensatz auf, wird die Spalte „Status Identifikation“ mit „nicht möglich“ ergänzt.

# 2. merge_duplicate_places_by_variants.py

Dieses Skript führt eine Zusammenführung von Ortsnennungen durch, bei denen Ortsnamen in unterschiedlichen Schreibweisen – beispielsweise historisch oder alternativ – auftreten. Ziel ist es, Redundanzen in der Datei Orte_Identifikation.xlsx zu erkennen und zu aggregieren, um eine konsolidierte Version in Orte_Identifikation_merged.xlsx zu speichern.

# 3. enrich_place_data_with_edition_info.py

Dieses Python-Skript ergänzt den Datensatz Orte_Identifikation_merged.xlsx um zwei zusätzliche Informationen – Edition und Nr. Rep. Druck – durch Abgleich mit dem Referenzdatensatz Ortsdaten_Repertorium-aufbereitet.xlsx. Die Ergänzung basiert auf Ortsnamen, die entweder als Standardnamen oder als alternative/historische Schreibweisen angegeben sind.

# 4. OpenRefine-Reconciling und manuelle Ergänzungen

Nun wird auf die neu erstelle liste "Orte_Identifikation_merged.xlsx" ein Reconciling mit OpenRefine durchgeführt. Es wurde nach polnischen und ukrainischen Dröfern, sowie Städten geschaut und wikidata Informationen angereichert. Orte die nicht mit Reconciling gefunden wurden, per Chatgpt ins ukrainische translitiert und gegoogelt. Dann Wikidataeintrag über Instrumente der Wikipedia page manuell hinzugefügt.

# 5. resolve_factgrid_qids.py

Hier werden die Factgrid QIDs von der Wikidata QID resolved und in die neue Ortstabelle eingefügt. Nun ist die Tabelle nach noch einigen manuellen Anreicherungen erstmal in einem fertigen Zustand. Die Factgrid QIDs der Orte werden dann genutzt um sie bei dem Import der Urkunden über Quickstatements identifizieren zu können. 

# 6. Alte Skripte: orte.py und sum_orte.py

Die Skripte "orte.py" und "sum_orte.py" sind noch alte Versionen in denen alle Ortsnamen die in allen Listen der Forschungsdaten auftauchen zusammenführen. Letztlich wurde sich aber erstmal auf die Orte im Repertorium konzentriert. Die restliche Orte können bei Gelegenheit noch angereichert und eingefügt werden.

