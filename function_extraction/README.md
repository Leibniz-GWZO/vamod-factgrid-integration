# 1. aemter_receiver_issuer.py

Dieses Skript zeigt die Ämter oder Funktionen von Personen mit bestimmten Rollen in den Urkunden. Aus der Liste "Personen_Empfaenger_Aussteller_aggregiert.xlsx", die aus der Liste "Urkunden_gesamt_neu.xlsx" erstellt wurde, werden nun die Ämter von Personen mit bestimmten Rollen extrahiert. Wir haben entschieden, uns zuerst um die wichtigsten Personen zu kümmern und die Zeugen hinten anzustellen.  
Wir haben uns auf folgende Rollen konzentriert:

patterns = ["aussteller", "begünst", "empfänger", "käufer", "vorgänger", "mitaussteller"]

# 2. cluster_similar_aemter_names.py

Dieses Skript gruppiert ähnlich lautende Ämternamen (z. B. durch Tippfehler oder Variationen entstanden) mithilfe von Ähnlichkeitsvergleichen (fuzzy Matching). Es erstellt ein Matching-Verzeichnis, in dem zusammengehörige Namen in einer gemeinsamen Zeile gruppiert sind. Dies erleichtert eine spätere Normalisierung oder Vereinheitlichung der Datensätze.

**Clustering basierend auf String-Ähnlichkeit:**

- Jeder Ämtername wird mit bestehenden Clustern verglichen.
- Die Namen werden in zwei Teile gesplittet:
  - Erstes Wort (z. B. „Kanzlei“)
  - Rest (z. B. „des Bischofs“)
- Zwei Schwellenwerte entscheiden über eine Gruppierung:
  - 0.89 für das erste Wort
  - 0.85 für den Rest
- Bei Erreichen beider Schwellen wird der neue Name zum existierenden Cluster hinzugefügt.
- Ansonsten wird ein neuer Cluster eröffnet.

**Ergebnisstruktur:**

- Die Ausgabe ist eine Excel-Tabelle *Aemter_Matching.xlsx*, in der:
  - Spalte „Ämtername“ den Hauptnamen des Clusters enthält,
  - Spalten „Name 2“, „Name 3“, … weitere ähnliche Namen enthalten.

Das Beibehalten der einzelnen Schreibweisen der Ämter ist wichtig, um am Ende wieder eine korrekte Zuordnung der Ämter zu den Personen machen zu können.

