import pandas as pd
import difflib
from tqdm import tqdm

# Ähnlichkeitsschwellen für Vergleich
FIRST_WORD_THRESHOLD = 0.89
REST_THRESHOLD = 0.85

# Excel-Datei einlesen (bitte den Dateinamen und Pfad ggf. anpassen)
df = pd.read_excel("datasheets/aemter/Funktionen_Urkunden.xlsx")

# Annahme: Die erste Spalte enthält die relevanten Ämternamen.
# Konvertiere die Werte in Kleinbuchstaben und entferne fehlende Werte.
names = df.iloc[:, 0].dropna().str.lower().tolist()

# Cluster initialisieren: Jeder Cluster ist eine Liste mit ähnlich klingenden Ämternamen.
clusters = []

for name in tqdm(names):
    found_cluster = False
    # Teile den aktuellen Namen in ersten Wortteil und Rest
    candidate_split = name.split(maxsplit=1)
    candidate_first = candidate_split[0]
    candidate_rest = candidate_split[1] if len(candidate_split) > 1 else ""
    
    for cluster in clusters:
        # Teile den repräsentativen Namen des Clusters in ersten Wortteil und Rest
        cluster_split = cluster[0].split(maxsplit=1)
        cluster_first = cluster_split[0]
        cluster_rest = cluster_split[1] if len(cluster_split) > 1 else ""
        
        similarity_first = difflib.SequenceMatcher(None, candidate_first, cluster_first).ratio()
        similarity_rest = difflib.SequenceMatcher(None, candidate_rest, cluster_rest).ratio()
        
        if similarity_first >= FIRST_WORD_THRESHOLD and similarity_rest >= REST_THRESHOLD:
            cluster.append(name)
            found_cluster = True
            break
            
    if not found_cluster:
        clusters.append([name])

# Erstelle eine neue DataFrame-Tabelle:
# - Spalte "Ämtername" enthält den ersten Namen des Clusters.
# - Weitere Spalten "Name 2", "Name 3" usw. enthalten die weiteren ähnlichen Namen, jedoch ohne Duplikate.
clustered_data = []
for cluster in clusters:
    # Entferne Duplikate unter Beibehaltung der Reihenfolge
    unique_cluster = []
    for n in cluster:
        if n not in unique_cluster:
            unique_cluster.append(n)
    
    row = {}
    row["Ämtername"] = unique_cluster[0]
    for idx, similar_name in enumerate(unique_cluster[1:], start=2):
        row[f"Name {idx}"] = similar_name
    clustered_data.append(row)

clustered_df = pd.DataFrame(clustered_data)

# Speichere die resultierende Tabelle in eine neue Excel-Datei im gewünschten Verzeichnis.
clustered_df.to_excel("datasheets/aemter/Aemter_Matching.xlsx", index=False)

print("Clustering abgeschlossen. Die Datei 'Aemter_Matching.xlsx' wurde erstellt.")
