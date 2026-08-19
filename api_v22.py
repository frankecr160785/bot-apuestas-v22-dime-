import os, sys, time, argparse, requests
from datetime import datetime
from openpyxl import load_workbook

BASE_URL = "https://v3.football.api-sports.io"

ALLOWED_COUNTRIES = [
    "England","Spain","Italy","Germany","France","Netherlands","Portugal",
    "Switzerland","Sweden","Turkey","Finland","Iceland",
    "Brazil","Argentina","Colombia","Mexico","USA",
    "China","Japan","South-Korea"
]

GOAL_FOCUS_COUNTRIES = {"Finland","Iceland"}

PREFERRED_LEAGUES = {
    "England":["Premier League","Championship"],
    "Spain":["La Liga","Segunda División"],
    "Italy":["Serie A","Serie B"],
    "Germany":["Bundesliga","2. Bundesliga"],
    "France":["Ligue 1","Ligue 2"],
    "Netherlands":["Eredivisie","Eerste Divisie"],
    "Portugal":["Primeira Liga"],
    "Switzerland":["Super League","Challenge League"],
    "Sweden":["Allsvenskan","Superettan"],
    "Turkey":["Süper Lig","1. Lig"],
    "Finland":["Veikkausliiga","Ykkösliiga","Ykkönen","Kakkonen - Lohko A","Kakkonen - Lohko B","Kakkonen - Lohko C"],
    "Iceland":["Úrvalsdeild","1. Deild","2. Deild"],
    "Brazil":["Serie A","Serie B"],
    "Argentina":["Liga Profesional Argentina","Primera Nacional"],
    "Colombia":["Primera A","Primera B"],
    "Mexico":["Liga MX","Liga de Expansión MX"],
    "USA":["Major League Soccer","USL Championship"],
    "China":["Super League","League One"],
    "Japan":["J1 League","J2 League","J3 League"],
    "South-Korea":["K League 1","K League 2","K3 League"],
}

class RateLimiter:
    def __init__(self, interval=6.5):
        self.interval = interval
        self.last = 0
    def wait(self):
        delay = self.interval - (time.time() - self.last)
        if delay > 0:
            time.sleep(delay)
        self.last = time.time()

def api_get(session, limiter, path, params, key, retries=3):
    for attempt in range(retries):
        limiter.wait()
        r = session.get(
            BASE_URL + path,
            headers={"x-apisports-key": key},
            params=params,
            timeout=30
        )
        if r.status_code == 429:
            wait = 15 * (attempt + 1)
            print(f"429: esperando {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        data = r.json()
        if data.get("errors"):
            raise RuntimeError(str(data["errors"]))
        return data.get("response", [])
    raise RuntimeError("Límite de API alcanzado después de varios reintentos")

def get_leagues(session, limiter, key, country, season):
    return api_get(session, limiter, "/leagues",
                   {"country": country, "season": season}, key)

def get_fixtures(session, limiter, key, league_id, season, date_s):
    return api_get(session, limiter, "/fixtures",
                   {"league": league_id, "season": season,
                    "date": date_s, "timezone": "America/Lima"}, key)

def get_stats(session, limiter, key, fixture_id):
    return api_get(session, limiter, "/fixtures/statistics",
                   {"fixture": fixture_id}, key)

def get_odds(session, limiter, key, fixture_id):
    return api_get(session, limiter, "/odds",
                   {"fixture": fixture_id}, key)

def fill_excel(xlsx, fixtures):
    wb = load_workbook(xlsx)
    bd = wb["BASE_DATOS"]
    existing = set()
    for row in bd.iter_rows(min_row=2, values_only=True):
        if row[0] and row[2] and row[4]:
            existing.add((str(row[0]), str(row[2]), str(row[4])))

    for f in fixtures:
        fx, teams, league = f["fixture"], f["teams"], f["league"]
        date_s = fx.get("date", "")[:10]
        home, away = teams["home"]["name"], teams["away"]["name"]
        key = (date_s, home, away)
        if key in existing:
            continue
        bd.append([
            date_s, league.get("name",""), home, "LOCAL", away, "",
            "", "", "", "", "", "", "", "", "",
            f"fixture_id={fx.get('id')}; country={league.get('country','')}"
        ])
    wb.save(xlsx)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=datetime.now().strftime("%Y-%m-%d"))
    ap.add_argument("--season", type=int, default=datetime.now().year)
    ap.add_argument("--xlsx", default="Excel_V2_2_Motor_Automatico.xlsx")
    ap.add_argument("--details", action="store_true")
    ap.add_argument("--max-fixtures", type=int, default=8)
    args = ap.parse_args()

    key = os.getenv("API_FOOTBALL_KEY")
    if not key:
        print("ERROR: falta API_FOOTBALL_KEY")
        sys.exit(2)
    if not os.path.exists(args.xlsx):
        print(f"ERROR: no existe {args.xlsx}")
        sys.exit(3)

    session = requests.Session()
    limiter = RateLimiter()

    selected_leagues = []
    for country in ALLOWED_COUNTRIES:
        try:
            leagues = get_leagues(session, limiter, key, country, args.season)
            for item in leagues:
                lg = item.get("league", {})
                if lg.get("type") != "League":
                    continue
                name = lg.get("name","")
                preferred = name in PREFERRED_LEAGUES.get(country, [])
                selected_leagues.append({
                    "id": lg.get("id"),
                    "name": name,
                    "country": country,
                    "priority": 0 if preferred else 1
                })
        except Exception as e:
            print(f"Ligas {country}: {e}")

    selected_leagues.sort(key=lambda x: (x["priority"], x["country"], x["name"]))
    print(f"Ligas objetivo encontradas: {len(selected_leagues)}")

    fixtures = []
    for lg in selected_leagues:
        try:
            fs = get_fixtures(session, limiter, key, lg["id"], args.season, args.date)
            for f in fs:
                f["_priority"] = lg["priority"]
                f["_country"] = lg["country"]
                fixtures.append(f)
        except Exception as e:
            print(f"Fixture {lg['country']} {lg['name']}: {e}")

    fixtures.sort(key=lambda f: (f.get("_priority",1), f["fixture"].get("date","")))
    print(f"Partidos en ligas permitidas: {len(fixtures)}")

    # Detailed calls are deliberately capped to avoid the free-plan rate limit.
    details = fixtures[:args.max_fixtures] if args.details else []
    for i, f in enumerate(details, 1):
        fid = f["fixture"]["id"]
        try:
            get_stats(session, limiter, key, fid)
        except Exception as e:
            print("stats", fid, e)
        try:
            get_odds(session, limiter, key, fid)
        except Exception as e:
            print("odds", fid, e)
        print(f"Detalle {i}/{len(details)} fixture={fid}")

    fill_excel(args.xlsx, fixtures)
    print(f"Excel actualizado: {args.xlsx}")
    print("V2.3 OK")

if __name__ == "__main__":
    main()
