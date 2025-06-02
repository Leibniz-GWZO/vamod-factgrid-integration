# 1. extract_relevant_persons_from_documents.py

Dieses Skript extrahiert aus einer umfangreichen Excel-Tabelle (Urkunden_gesamt_neu.xlsx) Personen mit relevanten Rollen (z. B. Aussteller, Empfänger) und speichert diese strukturiert in einer neuen Datei. Die Rollen die berücksichtigt werden beinhalten diese Strings:

patterns = ["aussteller", "begünst", "empfänger", "käufer", "vorgänger", "mitaussteller"] 
	
Es werden folgende Informationen bzw. Spalten aus der Urkunden_gesamt_neu.xlsx übertragen:
"Genannte Personen", "Rolle in Urkunde", "Funktion", "Geschlecht", "Edition", "Datum", "Aussteller". Mehrere Personen in einer Zelle (Genannte Personen) werden anhand von Semikolon (;) getrennt.

# 2. cluster_person_roles_by_similarity.py

Dieses Skript führt eine fuzzy-basierte Aggregation von Personen mit ihren Rollen in Urkunden durch. Es versucht unterschiedliche Namensvarianten durch ein fuzzy Matching zu gruppieren um die selbe Person zu identifizieren und so die Gesamtanzahl der Personen zu verringern. Das fuzzy matching basiert auf der Namensähnlichkeit und der Funktionenähnlichkeit. 

# 3. Alte Skripte: person_matching.py und list_persons.py

Die Skripte "person_matching.py" und "list_persons.py" sind noch alte Skripte in denen alle Personen extrahiert wurden und mit einem Cluster Verfahren gematcht wurden. Jedoch ist die Liste so groß das wir uns zunächst auf Personen mit bestimmten Rollen beschränkt haben.

