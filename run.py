import requests
import pandas as pd
from datetime import date
import time
import random

WDQS_URL = "https://urldefense.com/v3/__https://query.wikidata.org/sparql__;!!Nyu6ZXf5!o6CalnBzomk5ghgzAUKByy6vZagedUy39uO3NcIOtRBRkASroJtvCPTdF_zUzsgXdWVzy5rjxQJ1zpdK0uE8UQVJtgmO$ "

# DAX (Wikidata Item)
DAX_ITEM = "Q155718"

SPARQL_AUFSICHTSRAETE = f"""
SELECT
  ?company ?companyLabel
  ?isin
  ?person ?personLabel
  ?start
WHERE {{
  # Aktuelle DAX-Mitglieder (ohne Enddatum)
  ?company p:P361 ?daxStmt .
  ?daxStmt ps:P361 wd:{DAX_ITEM} .
  FILTER NOT EXISTS {{ ?daxStmt pq:P582 ?daxEnd . }}

  OPTIONAL {{ ?company wdt:P946 ?isin . }}

  # Aktuelle Aufsichtsratsmitglieder (P5052) ohne Enddatum
  ?company p:P5052 ?stmt .
  ?stmt ps:P5052 ?person .
  FILTER NOT EXISTS {{ ?stmt pq:P582 ?end . }}

  OPTIONAL {{ ?stmt pq:P580 ?start . }}

  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "de,en". }}
}}
ORDER BY ?companyLabel ?personLabel
"""

SPARQL_DAX_ONLY = f"""
SELECT ?company ?companyLabel WHERE {{
  ?company p:P361 ?daxStmt .
  ?daxStmt ps:P361 wd:{DAX_ITEM} .
  FILTER NOT EXISTS {{ ?daxStmt pq:P582 ?daxEnd . }}

  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "de,en". }}
}}
ORDER BY ?companyLabel
"""

def wdqs_query(query: str) -> dict:
    headers = {
        "Accept": "application/sparql-results+json",
        "User-Agent": "dax-aufsichtsrat-tool/1.0 (github actions)"
    }

    # Retries (WDQS liefert manchmal 429/503)
    for attempt in range(1, 8):
        r = requests.get(WDQS_URL, params={"query": query}, headers=headers, timeout=60)

        if r.status_code == 200:
            return r.json()

        if r.status_code in (429, 500, 502, 503, 504):
            wait = min(60, (2 ** attempt) + random.random())
            print(f"WDQS {r.status_code}, retry in {wait:.1f}s (attempt {attempt}/7)")
            time.sleep(wait)
            continue

        r.raise_for_status()

    raise RuntimeError("WDQS mehrfach fehlgeschlagen (429/5xx).")

def safe_value(binding: dict, key: str) -> str:
    if key in binding and "value" in binding[key]:
        return binding[key]["value"]
    return ""

def main():
    today = date.today().isoformat()

    data = wdqs_query(SPARQL_AUFSICHTSRAETE)
    rows = []
    for b in data.get("results", {}).get("bindings", []):
        company = safe_value(b, "companyLabel")
        isin = safe_value(b, "isin")
        person = safe_value(b, "personLabel")
        start = safe_value(b, "start")

        # wenn Label fehlt, Zeile überspringen
        if not company or not person:
            continue

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

    dax_data = wdqs_query(SPARQL_DAX_ONLY)
    dax_all = []
    for x in dax_data.get("results", {}).get("bindings", []):
        label = safe_value(x, "companyLabel")
        if label:
            dax_all.append(label)

    have = set(df["Unternehmen"].dropna().unique()) if not df.empty else set()
    missing = sorted([c for c in dax_all if c not in have])

    df_missing = pd.DataFrame(
        [{"Unternehmen": c, "Hinweis": "Kein Aufsichtsrats-Eintrag in Wikidata gefunden"} for c in missing]
    )

    filename = f"aufsichtsraete_{today}.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Aufsichtsräte")
        pd.DataFrame({"DAX40 (laut Wikidata)": dax_all}).to_excel(writer, index=False, sheet_name="DAX40")
        df_missing.to_excel(writer, index=False, sheet_name="Fehlende Firmen")

    print("Excel erstellt:", filename)
    print("Rows:", len(df))
    print("Missing companies:", len(missing))

if __name__ == "__main__":
    main()
