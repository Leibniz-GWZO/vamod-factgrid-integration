# 1. extract_category_from_big_charterlist.py

Dieses Skript extrahiert alle einzigartigen Werte aus der Spalte "Betreff" der Excel-Datei `Urkunden_gesamt_neu.xlsx`. Die extrahierten Werte werden sortiert und in der Textdatei `Betreffe_große_Urkundenliste.txt` gespeichert. Dies dient dazu, eine Übersicht über alle vorkommenden Betreff-Kategorien zu erhalten.

# 2. scrape_transaction_types.py

Dieses Skript extrahiert Transaktionstypen aus einer lokalen HTML-Datei (`Search _ People of Medieval Scotland.html`). Es parst die HTML-Struktur, um eine Liste von Transaktionstypen zu finden, bereinigt diese von überflüssigen Zeichen (wie z.B. Zahlen in Klammern) und speichert das Ergebnis in der Textdatei `Poms_transcation_types.txt`.
