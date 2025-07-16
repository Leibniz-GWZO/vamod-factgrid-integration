# 1. extract_relevant_persons_from_documents.py

Dieses Skript extrahiert aus einer umfangreichen Excel-Tabelle (Urkunden_gesamt_neu.xlsx) Personen mit relevanten Rollen (z. B. Aussteller, Empfänger) und speichert diese strukturiert in einer neuen Datei. Die Rollen die berücksichtigt werden beinhalten diese Strings:

patterns = ["aussteller", "begünst", "empfänger", "käufer", "vorgänger", "mitaussteller"] 
	
Es werden folgende Informationen bzw. Spalten aus der Urkunden_gesamt_neu.xlsx übertragen:
"Genannte Personen", "Rolle in Urkunde", "Funktion", "Geschlecht", "Edition", "Datum", "Aussteller". Mehrere Personen in einer Zelle (Genannte Personen) werden anhand von Semikolon (;) getrennt.

# 2. cluster_person_roles_by_similarity.py

Dieses Skript führt eine fuzzy-basierte Aggregation von Personen mit ihren Rollen in Urkunden durch. Es versucht unterschiedliche Namensvarianten durch ein fuzzy Matching zu gruppieren um die selbe Person zu identifizieren und so die Gesamtanzahl der Personen zu verringern. Das fuzzy matching basiert auf der Namensähnlichkeit und der Funktionenähnlichkeit. 

# 3. Alte Skripte: person_matching.py und list_persons.py

Die Skripte "person_matching.py" und "list_persons.py" sind noch alte Skripte in denen alle Personen extrahiert wurden und mit einem Cluster Verfahren gematcht wurden. Jedoch ist die Liste so groß das wir uns zunächst auf Personen mit bestimmten Rollen beschränkt haben.

# 4. cluster_name_abbreviations.py

Dieses Skript dient der Aggregation von Namensvarianten und Abkürzungen. Es liest eine Excel-Datei und fasst Zeilen zusammen, die vermutlich zur selben Person gehören. Der Prozess wird durch manuelle Markierungen (gelb hinterlegte Zellen) gesteuert, um verschiedene Schreibweisen eines Namens zu gruppieren.

# 5. cluster_names_after_manual_lookup.py

Dieses Skript finalisiert den Prozess der Namens-Clusterings nach einer manuellen Überprüfung. Es verarbeitet eine Excel-Datei, in der korrekte oder normalisierte Namen ("Normierte deutsche Schreibweise") grün markiert wurden. Anhand dieser Markierungen werden die Personendaten entsprechend gruppiert und zusammengefasst.

# 6. add_charter_identifier_to_person.py

Dieses Skript reichert die finale Personenliste mit der "Nr. Rep" (Nummer aus dem Repertorium) aus der Urkundenliste an. Es iteriert durch die lateinischen Schreibweisen der gennantnen Personen in den Urkunden und fügt die entsprechende "Nr. Rep" zu den Personen in der Personenliste hinzu, um eine direkte Verknüpfung zwischen Person und Urkunde herzustellen.

