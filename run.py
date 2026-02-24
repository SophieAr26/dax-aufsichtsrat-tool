import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import time

# Hier trägst du die Links zu den Aufsichtsrats-Seiten ein
companies = {
    "SAP SE": "https://urldefense.com/v3/__https://www.sap.com/about/company/leadership/supervisory-board.html__;!!Nyu6ZXf5!reHhZxVTUykvn9a3f716S_km8VvUVjCIxRavH-UbdZftAo95TEy-Knxx9Mb6dIKfxwfYKfeaK7g55vRQtDqXAZSiWWD9$ ",
    "Siemens AG": "https://urldefense.com/v3/__https://www.siemens.com/global/en/company/about/governance/supervisory-board.html__;!!Nyu6ZXf5!reHhZxVTUykvn9a3f716S_km8VvUVjCIxRavH-UbdZftAo95TEy-Knxx9Mb6dIKfxwfYKfeaK7g55vRQtDqXAR2lSsSR$ "
}

def get_board_members(company, url):
    print(f"Lade {company}")
    members = []
   
    try:
        response = requests.get(url, timeout=30)
        soup = BeautifulSoup(response.text, "lxml")
       
        for li in soup.find_all("li"):
            text = li.get_text(strip=True)
            if len(text) > 5 and len(text) < 100:
                members.append(text)
               
    except Exception as e:
        print(f"Fehler bei {company}: {e}")
   
    return members

def main():
    today = date.today()
    all_rows = []
   
    for company, url in companies.items():
        members = get_board_members(company, url)
        time.sleep(2)
       
        for member in members:
            all_rows.append({
                "Unternehmen": company,
                "Name/Rolle": member,
                "Datum": today
            })
   
    df = pd.DataFrame(all_rows)
    filename = f"aufsichtsraete_{today}.xlsx"
    df.to_excel(filename, index=False)
    print("Excel-Datei erstellt:", filename)

if __name__ == "__main__":
    main()
