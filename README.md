In diesem Repository soll dokumentiert werden, wie die Forschungsdaten des Projekts „VAMOD” am Leibniz-Institut für Geschichte und Kultur des östlichen Europa (GWZO) für einen Datenimport in die Graphdatenbank „Factgrid” vorbereitet und aufbereitet wurden.

## Projektseite und Erklärungen zum Projekt

[https://database.factgrid.de/wiki/FactGrid:VAMOD](https://database.factgrid.de/wiki/FactGrid:VAMOD)

Die Forschungsdaten liegen in Form von verschiedenen Excel-Tabellen und PDF-Dateien vor. Ziel ist es, die Informationen so aufzuarbeiten und zu bereinigen, dass einzelne, gut strukturierte CSV-Dateien erstellt werden können, die anschließend über Quickstatements bei Factgrid eingepflegt werden können.

Das Ziel besteht darin, am Ende ca. 600 Urkunden in einem großen Import in die Datenbank zu bringen. Für den Import der Urkundenliste müssen zunächst jedoch alle anderen Informationen, die in den Urkunden stehen, in FactGrid vorhanden sein. Denn anders als in einer klassischen relationalen Datenbank stehen in der FactGrid-Datenbank die Daten in Form von Knoten (Entitäten) und Kanten (Beziehungen) zueinander. Das heißt, Informationen sind bereits mit anderen Informationen verknüpft und müssen in ein bestehendes Netzwerk eingepflegt werden, das sich dynamisch verändert.

Damit die Urkunden eingepflegt werden können, muss zunächst kontrolliert werden, ob Informationen aus den Urkunden, wie etwa Orte, Personen, innegehabte Ämter und Rechtsakte, bereits in der Factgrid-Datenbank vorhanden sind oder ob sie ggf. noch angelegt werden müssen.

Zuerst legt man die Informationen an, auf die andere Informationen aufbauen.  
Damit Personen mit Ämtern eingepflegt werden können, müssen zuerst die Ämter angelegt werden. Danach können die Urkunden mit den genannten Personen angelegt werden. Die einzelnen Schichten sind durch die Ordner dargestellt. In der Readme wird erklärt, wie man aus den Forschungsdaten mithilfe von Python-Skripten, Open Refine und manuellen Anreicherungen ein importfähiges CSV-File erstellt, um die einzelne Schicht in Factgrid zu importieren.


---

## Ordnerstruktur

Die Skripte innerhalb der Ordner sind in einer sinnvollen Reihenfolge abgelegt, da bestimmte Verarbeitungsschritte aufeinander aufbauen und daher in einer festen Abfolge ausgeführt werden sollten. 

### `manage_charte_lists`
Der Ordner enthält Skripte und Daten zur automatisierten Extraktion, Normalisierung und Anreicherung historischer Urkundeneinträge aus PDF- und Excel-Dokumenten. Ziel ist die strukturierte Aufbereitung für den CSV-Import. Dafür müssen Informationen aus PDF-Dateien sowie aus Excel-Tabellen zusammengeführt werden. Hier entsteht die CSV-Datei, mit der alle Urkunden dann per Quickstatement-Import auf einmal importiert werden. Das ist der letzte Schritt des Schichtmodells.

### `person_extraction`
Dieser Ordner enthält Skripte und Dokumentationen zur Extraktion, Filterung und Gruppierung historischer Personen mit relevanten Rollen aus den Urkunden. Ziel ist es, Personen mit den wichtigsten Rollen wie Aussteller, Empfänger oder Käufer strukturiert aufzubereiten und Varianten desselben Namens zusammenzuführen. Die extrahierten Personen müssen anschließend nochmals manuell gruppiert und mit Informationen angereichert werden. Dies ist der vorletzte Schritt des Schichtmodells.

### `funciton_extraction`
Dieser Ordner dient der Extraktion, Filterung und Vereinheitlichung historischer Funktionsbezeichnungen (Ämter), die Personen in den vormodernen Urkunden zugewiesen werden. Ziel ist es, Ämter der relevanten Personen zu extrahieren und Varianten (z. B. Tippfehler, Schreibvarianten) zu clustern.

### `place_extraction`
Dieser Ordner enthält die Skripte zur Erstellung und Anreicherung eines konsolidierten Ortsverzeichnisses für alle im *Repertorium diplomatum terrae Russie Regni Poloniae* erwähnten Orte. Grundlage ist eine kuratierte Tabelle mit Ortsnennungen sowie eine manuell bereinigte Identifikationstabelle. Ziel ist es, Mehrfachnennungen durch historische Varianten zu erkennen, die Orte mit Editionseinträgen zu verknüpfen und externe QIDs zur Nachnutzung einzubinden.

---

## gesammelte Erfahrungen:

- Informationen, die in mehreren Tabellen zusammenhängen, sollten mit eindeutigen Identifikatoren verknüpft werden  
- einheitliche Schreibweisen und Trennzeichen verwenden
- keine Abkürzungen verwenden

