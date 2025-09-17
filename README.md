# Projektseite und Erklärungen zum Projekt

[https://database.factgrid.de/wiki/FactGrid:VAMOD](https://database.factgrid.de/wiki/FactGrid:VAMOD)

Die Forschungsdaten liegen in unterschiedlichen Excel-Tabellen und PDF-Dokumenten vor. Ziel ist es, diese Informationen so zu bereinigen und aufzubereiten, dass daraus klar strukturierte CSV-Dateien entstehen, die sich über QuickStatements in FactGrid importieren lassen.

Geplant ist ein umfassender Import von rund 600 Urkunden. Damit dieser reibungslos erfolgen kann, müssen zunächst alle in den Urkunden enthaltenen Informationen – wie Orte, Personen, Ämter oder Rechtsakte – bereits in FactGrid vorhanden sein. Denn im Unterschied zu klassischen relationalen Datenbanken basiert FactGrid auf einem graphbasierten Modell: Daten werden in Form von Tripeln gespeichert, bestehend aus Knoten (Entitäten) und Kanten (Beziehungen). Neue Informationen müssen daher in ein bestehendes, dynamisch wachsendes Netzwerk eingebunden werden.

Bevor die Urkunden importiert werden können, ist daher zu prüfen, ob alle benötigten Entitäten bereits existieren oder noch angelegt werden müssen. So müssen etwa Ämter bereits vorhanden sein, bevor Personen mit diesen Ämtern erfasst werden können. Erst danach ist es möglich, die Urkunden samt ihrer Bezüge korrekt anzulegen.

Die Importlogik folgt dabei einer hierarchischen Struktur, die sich in der Ordnerstruktur widerspiegelt. In der zugehörigen Readme wird erläutert, wie aus den Rohdaten mit Hilfe von Python-Skripten, OpenRefine und manuellen Ergänzungen Schritt für Schritt importfähige CSV-Dateien erstellt werden, mit denen sich die jeweiligen Informationsschichten in FactGrid einpflegen lassen.


---

## Ordnerstruktur

Die Skripte innerhalb der Ordner sind in einer sinnvollen Reihenfolge abgelegt, da bestimmte Verarbeitungsschritte aufeinander aufbauen und daher in einer festen Abfolge ausgeführt werden sollten. 

### `material_charters`
Dieser Ordner enthält Skripte zur schrittweisen Aufbereitung der materiellen Urkundenmetadaten für den FactGrid-Import. Die Skripte verarbeiten die Grundinformationen der physischen Urkunden (Aussteller, Empfänger, Datum, Ausstellungsort, Archiv) und bereiten sie mit mehrsprachigen Labels, Übersetzungen und FactGrid-QID-Verknüpfungen für den Import vor. Die Verarbeitung erfolgt in mehreren aufeinander aufbauenden Schritten von der Label-Erstellung über Übersetzungen bis zur finalen Formatierung.

### `person_extraction`
Dieser Ordner enthält Skripte und Dokumentationen zur Extraktion, Filterung und Gruppierung historischer Personen mit relevanten Rollen aus den Urkunden. Ziel ist es, Personen mit den wichtigsten Rollen wie Aussteller, Empfänger oder Käufer strukturiert aufzubereiten und Varianten desselben Namens zusammenzuführen. Die extrahierten Personen müssen anschließend nochmals manuell gruppiert und mit Informationen angereichert werden. Dies ist der vorletzte Schritt des Schichtmodells.

### `funciton_extraction`
Dieser Ordner dient der Extraktion, Filterung und Vereinheitlichung historischer Funktionsbezeichnungen (Ämter), die Personen in den vormodernen Urkunden zugewiesen werden. Ziel ist es, Ämter der relevanten Personen zu extrahieren und Varianten (z. B. Tippfehler, Schreibvarianten) zu clustern.

### `place_extraction`
Dieser Ordner enthält die Skripte zur Erstellung und Anreicherung eines konsolidierten Ortsverzeichnisses für alle im *Repertorium diplomatum terrae Russie Regni Poloniae* erwähnten Orte. Grundlage ist eine kuratierte Tabelle mit Ortsnennungen sowie eine manuell bereinigte Identifikationstabelle. Ziel ist es, Mehrfachnennungen durch historische Varianten zu erkennen, die Orte mit Editionseinträgen zu verknüpfen und externe QIDs zur Nachnutzung einzubinden.

### `legal_types`
Dieser Ordner enthält Skripte und Textdateien, die zur Klassifizierung der Rechtsnatur der Urkunden dienen. Ziel ist es, aus verschiedenen Quellen wie den Urkundenbetreffen oder externen Webseiten eine Taxonomie von Transaktionstypen zu extrahieren und zu erstellen. Diese Taxonomie hilft dabei, die Urkunden systematisch zu kategorisieren.

### `manage_charte_lists`
Der Ordner enthält Skripte und Daten zur automatisierten Extraktion, Normalisierung und Anreicherung historischer Urkundeneinträge aus PDF- und Excel-Dokumenten. Ziel ist die strukturierte Aufbereitung für den CSV-Import. Dafür müssen Informationen aus PDF-Dateien sowie aus Excel-Tabellen zusammengeführt werden. Hier entsteht die CSV-Datei, mit der alle Urkunden dann per Quickstatement-Import auf einmal importiert werden. Das ist der letzte Schritt des Schichtmodells.

### `LLM_evaluation`
Dieser Ordner enthält Skripte zur systematischen Evaluierung verschiedener Ansätze zur automatischen Entitätserkennung (Named Entity Recognition) in historischen Urkundenregesten. Es werden Large Language Models (LLMs) wie GPT-4o, GPT-4.1 und llama-3.3 sowie regelbasierte Ansätze getestet und verglichen. Die Evaluierung erfolgt anhand eines Ground Truth-Datensatzes von 50 zufällig ausgewählten Urkunden und misst die Genauigkeit bei der Extraktion von drei Entitätstypen: Objektorte, Empfängersitze und genannte Personen. Ziel ist es, die Leistungsfähigkeit automatisierter Verfahren für die historische Textanalyse zu bewerten.

