import time
import re
import yaml
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import date


def load_config(path="sources.yml") -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def normalize_text(s: str) -> str:
    s = re.sub(r"\s+", " ", s or "").strip()
    # typische Trennzeichen/Artefakte etwas glätten
    s = s.replace("\u00a0", " ")
    return s


def looks_like_name(s: str) -> bool:
    """
    Heuristik: kein perfekter Namensdetektor, aber gut genug für Review-Workflow.
    """
    if not s:
        return False
    if len(s) < 5 or len(s) > 80:
        return False

    lower = s.lower()
    # typische Navigations-/Footer-Wörter raus
    blacklist = [
        "privacy", "cookie", "terms", "imprint", "kontakt", "contact",
        "datenschutz", "rechtliche", "legal", "karriere", "career",
        "presse", "news", "login", "sign in", "digital id"
    ]
    if any(w in lower for w in blacklist):
        return False

    # sollte aus mindestens 2 "Wörtern" bestehen
    parts = [p for p in s.split(" ") if p]
    if len(parts) < 2:
        return False

    # keine reinen Rollen-/Überschriften
    role_words = ["aufsichtsrat", "supervisory", "board", "mitglied", "vorsitz"]
    if any(w in lower for w in role_words) and len(parts) <= 3:
        return False

    return True


def fetch_html(url: str, user_agent: str, timeout: int = 45) -> str:
    r = requests.get(url, headers={"User-Agent": user_agent}, timeout=timeout)
    r.raise_for_status()
    return r.text


def extract_members_from_html(html: str, selector: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    nodes = soup.select(selector)
    texts = [normalize_text(n.get_text(" ", strip=True)) for n in nodes]
    # dedupe, preserve order
    seen = set()
    out = []
    for t in texts:
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def main():
    cfg = load_config("sources.yml")
    meta = cfg.get("meta", {})
    companies = cfg.get("companies", {})

    user_agent = meta.get("user_agent", "dax-aufsichtsrat-tool/1.0")
    delay = float(meta.get("request_delay_seconds", 1.0))
    min_members = int(meta.get("min_members", 6))
    max_members = int(meta.get("max_members", 30))

    today = date.today().isoformat()

    rows = []
    review_rows = []
    source_rows = []

    for company, info in companies.items():
        url = info.get("url", "").strip()
        selector = info.get("selector", "").strip()

        source_rows.append({
            "Unternehmen": company,
            "URL": url,
            "Selector": selector
        })

        if not url or not selector:
            review_rows.append({
                "Unternehmen": company,
                "Problem": "URL oder Selector fehlt",
                "Treffer": 0,
                "Hinweis": "In sources.yml ergänzen"
            })
            continue

        try:
            html = fetch_html(url, user_agent=user_agent)
            raw_members = extract_members_from_html(html, selector)

            # heuristisch filtern
            members = [m for m in raw_members if looks_like_name(m)]

            for m in members:
                rows.append({
                    "Unternehmen": company,
                    "Person": m,
                    "Rolle": "Mitglied Aufsichtsrat",
                    "Stand": today,
                    "Quelle": url
                })

            # Qualitätschecks → Review
            if len(members) < min_members:
                review_rows.append({
                    "Unternehmen": company,
                    "Problem": "Zu wenige Treffer",
                    "Treffer": len(members),
                    "Hinweis": "Selector prüfen oder Quelle wechseln"
                })
            elif len(members) > max_members:
                review_rows.append({
                    "Unternehmen": company,
                    "Problem": "Zu viele Treffer (wahrscheinlich Menü/Noise)",
                    "Treffer": len(members),
                    "Hinweis": "Selector enger machen"
                })

        except Exception as e:
            review_rows.append({
                "Unternehmen": company,
                "Problem": "Abruf/Parsing fehlgeschlagen",
                "Treffer": 0,
                "Hinweis": str(e)[:180]
            })

        time.sleep(delay)

    df = pd.DataFrame(rows, columns=["Unternehmen", "Person", "Rolle", "Stand", "Quelle"])
    df_review = pd.DataFrame(review_rows, columns=["Unternehmen", "Problem", "Treffer", "Hinweis"])
    df_sources = pd.DataFrame(source_rows, columns=["Unternehmen", "URL", "Selector"])

    filename = f"aufsichtsraete_{today}.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Aufsichtsräte")
        df_review.to_excel(writer, index=False, sheet_name="Review")
        df_sources.to_excel(writer, index=False, sheet_name="Quellen")

    print("Excel erstellt:", filename)
    print("Rows:", len(df))
    print("Review cases:", len(df_review))


if __name__ == "__main__":
    main()
