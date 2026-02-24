import requests
import pandas as pd
from datetime import date

WDQS_URL = "https://urldefense.com/v3/__https://query.wikidata.org/sparql__;!!Nyu6ZXf5!pwFrgayqkbAGMWey_XE2Jkkv0azIsMlTcNU517EV35CbUfCO7K75iTXDsy5PeShDMgZqPwhUSDKGaZwbigRv8mBZEVhd$ "
DAX_ITEM = "Q155718"  # DAX

SPARQL = """
SELECT
  ?company
  ?companyLabel
  ?isin
  ?person
  ?personLabel
  ?start
WHERE {

  # DAX-Mitglieder (ohne Enddatum)
  ?company p:P361 ?daxStmt .
  ?daxStmt ps:P361 wd:Q155718 .
  FILTER NOT EXISTS { ?daxStmt pq:P582 ?daxEnd . }

  OPTIONAL { ?company wdt:P946 ?isin . }

  # Aufsichtsratsmitglied (ohne Enddatum)
  ?company p:P5052 ?stmt .
  ?stmt ps:P5052 ?person .
  FILTER NOT EXISTS { ?stmt pq:P582 ?end . }

  OPTIONAL { ?stmt pq:P580 ?start . }

  SERVICE wikibase:label { bd:serviceParam wikibase:language "de,en". }
}
ORDER BY ?companyLabel ?personLabel
"""

def wdqs_query(query: str) -> dict:
    headers = {
        "Accept": "application/sparql-results+json",
        "User-Agent": "dax-aufsichtsrat-tool/1.0 (github actions)"
    }
    r = requests.get(WDQS_URL, params={"query": query}, headers=headers, timeout=60)
    r.raise_for_status()
    return r.json()

def main():
    today = date.today().isoformat()

    data = wdqs_query(SPARQL)
    rows = []
    for b in data["results"]["bindings"]:
       company = b["companyLabel"]["value"] if "companyLabel" in b else ""
        isin = b["isin"]["value"] if "isin" in b else ""
        person = b["personLabel"]["value"] if "personLabel" in b else ""
         start = b.get("start", {}).get("value", "")

        rows.append({
            "Unternehmen": company,
            "ISIN": isin,
            "Person": person,
            "Rolle": "Mitglied Aufsichtsrat",
            "Startdatum (falls vorhanden)": start,
            "Stand": today,
            "Quelle": "Wikidata (P361=DAX, P5052=Aufsichtsrat)"
        })

    df = pd.DataFrame(rows)

    # Qualitätscheck: Firmen ohne Treffer identifizieren
    dax_companies = sorted(df["Unternehmen"].dropna().unique().tolist())
    # Hinweis: wenn Wikidata unvollständig ist, fehlen Firmen komplett. Dann sieht man es hier nicht.
    # Deshalb zusätzlich: Liste der DAX-Firmen separat abfragen und vergleichen.
    SPARQL_DAX_ONLY = f"""
    SELECT ?companyLabel WHERE {{
      ?company p:P361 ?daxStmt .
      ?daxStmt ps:P361 wd:{DAX_ITEM} .
      FILTER NOT EXISTS {{ ?daxStmt pq:P582 ?daxEnd . }}
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "de,en". }}
    }}
    ORDER BY ?companyLabel
    """
    dax_data = wdqs_query(SPARQL_DAX_ONLY)
    dax_all = []
    for x in dax_data["results"]["bindings"]:
        if "companyLabel" in x:
            dax_all.append(x["companyLabel"]["value"])

    have = set(df["Unternehmen"].dropna().unique())
    missing = sorted([c for c in dax_all if c not in have])

    df_missing = pd.DataFrame([{"Unternehmen": c, "Hinweis": "Kein Aufsichtsrats-Eintrag in Wikidata gefunden"} for c in missing])

    filename = f"aufsichtsraete_{today}.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Aufsichtsräte")
        pd.DataFrame({"DAX40 (laut Wikidata)": dax_all}).to_excel(writer, index=False, sheet_name="DAX40")
        df_missing.to_excel(writer, index=False, sheet_name="Fehlende Firmen")

    print("Excel erstellt:", filename)
    print("Fehlende Firmen:", len(missing))

if __name__ == "__main__":
    main()

