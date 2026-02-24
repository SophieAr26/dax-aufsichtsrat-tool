import requests
import pandas as pd
from datetime import date
import time
import random

WDQS_URL = "https://urldefense.com/v3/__https://query.wikidata.org/sparql__;!!Nyu6ZXf5!p-3L7vq--9HP0WDqvc7PtKLpkRtvPqoX8tvcY1EqBHnzq-Mx-7e9AO-4cZaQoTzfQ313VD-lKqwULNACWjCZOmUXeJ3r$ "
DAX_ITEM = "Q155718"

SPARQL = f"""
SELECT
  ?company
  ?companyName
  ?isin
  ?person
  ?personName
  ?start
WHERE {{

  # DAX-Mitglieder
  {{
    ?company wdt:P361 wd:{DAX_ITEM} .
  }}
  UNION
  {{
    ?company p:P361 ?daxStmt .
    ?daxStmt ps:P361 wd:{DAX_ITEM} .
    FILTER NOT EXISTS {{ ?daxStmt pq:P582 ?daxEnd . }}
  }}

  OPTIONAL {{ ?company wdt:P946 ?isin . }}

  # Aufsichtsratsmitglieder
  ?company p:P5052 ?stmt .
  ?stmt ps:P5052 ?person .
  FILTER NOT EXISTS {{ ?stmt pq:P582 ?end . }}

  OPTIONAL {{ ?stmt pq:P580 ?start . }}

  # Labels explizit abfragen
  OPTIONAL {{ ?company rdfs:label ?companyName .
             FILTER (lang(?companyName) = "de" || lang(?companyName) = "en") }}

  OPTIONAL {{ ?person rdfs:label ?personName .
             FILTER (lang(?personName) = "de" || lang(?personName) = "en") }}
}}
ORDER BY ?companyName ?personName
"""

def wdqs_query(query: str) -> dict:
    headers = {
        "Accept": "application/sparql-results+json",
        "User-Agent": "dax-aufsichtsrat-tool/1.0"
    }

    for attempt in range(1, 6):
        r = requests.get(WDQS_URL, params={"query": query}, headers=headers, timeout=60)

        if r.status_code == 200:
            return r.json()

        if r.status_code in (429, 500, 502, 503, 504):
            wait = 2 ** attempt
            print(f"Retry in {wait}s")
            time.sleep(wait)
            continue

        r.raise_for_status()

    raise RuntimeError("Wikidata Service nicht erreichbar.")

def get_value(binding, key):
    if key in binding and "value" in binding[key]:
        return binding[key]["value"]
    return ""

def main():
    today = date.today().isoformat()

    data = wdqs_query(SPARQL)
    bindings = data.get("results", {}).get("bindings", [])
    print("Treffer:", len(bindings))

    rows = []

    for b in bindings:
        company = get_value(b, "companyName")
        isin = get_value(b, "isin")
        person = get_value(b, "personName")
        start = get_value(b, "start")

        rows.append({
            "Unternehmen": company,
            "ISIN": isin,
            "Person": person,
            "Rolle": "Mitglied Aufsichtsrat",
            "Startdatum (falls vorhanden)": start,
            "Stand": today,
            "Quelle": "Wikidata"
        })

    df = pd.DataFrame(rows)

    filename = f"aufsichtsraete_{today}.xlsx"
    df.to_excel(filename, index=False)

    print("Excel erstellt:", filename)
    print("Zeilen:", len(df))

if __name__ == "__main__":
    main()
