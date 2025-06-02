import pandas as pd
import requests
from tqdm import tqdm  # tqdm importieren

# 1. Excel-Datei einlesen
df = pd.read_excel("datasheets/orte/Orte_Identifikation_merged_SJ-neu(1).xlsx")

# 2. Spalte umbenennen: "Wikidata" → "Wikidata URL"
df = df.rename(columns={"Wikidata": "Wikidata URL"})

# 3. Wikidata-QID extrahieren
df["Wikidata QID"] = (
    df["Wikidata URL"]
    .dropna()
    .astype(str)
    .str.rstrip("/")
    .str.split("/")
    .str[-1]
)

# 4. Funktion zum Abrufen von FactGrid-URL und -QID
def get_factgrid_entry(qid: str):
    base = "https://database.factgrid.de/wiki/Special:ItemByTitle/wikidatawiki/{}"
    url = base.format(qid)
    resp = requests.get(url, allow_redirects=True)
    if resp.history:
        final_url = resp.url
        factgrid_qid = final_url.rstrip("/").split("/")[-1].split(":")[-1]
        return final_url, factgrid_qid
    else:
        return None, None

# 5. FactGrid-Infos abfragen mit tqdm für den Fortschritt
factgrid_urls = []
factgrid_qids = []

for q in tqdm(df["Wikidata QID"], desc="FactGrid-Abfragen"):
    if pd.isna(q) or not str(q).startswith("Q"):
        factgrid_urls.append(None)
        factgrid_qids.append(None)
    else:
        url, fg_qid = get_factgrid_entry(q)
        factgrid_urls.append(url)
        factgrid_qids.append(fg_qid)

# 6. Neue Spalten übernehmen
df["Factgrid URL"] = factgrid_urls
df["Factgrid QID"] = factgrid_qids

# 7. In eine neue Excel speichern
df.to_excel("datasheets/orte/Orte_Identifikation_factgrid.xlsx", index=False)

# 8. Ausgabe der ersten Zeilen (optional)
print(df[["Wikidata URL", "Wikidata QID", "Factgrid URL", "Factgrid QID"]].head())
