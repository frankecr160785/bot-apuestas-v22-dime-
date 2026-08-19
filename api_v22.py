import os, sys, time, argparse, math
from datetime import datetime, timedelta
import requests
from openpyxl import load_workbook

BASE_URL = "https://v3.football.api-sports.io"

def api_get(path, params, key):
    r = requests.get(
        BASE_URL + path,
        headers={"x-apisports-key": key},
        params=params,
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    if data.get("errors"):
        raise RuntimeError(f"API-Football error: {data['errors']}")
    return data

def as_num(x):
    if x is None or x == "":
        return ""
    if isinstance(x, str):
        x=x.replace("%","").strip()
    try:
        return float(x)
    except:
        return ""

def pct(x):
    x=as_num(x)
    return "" if x=="" else x/100 if x>1 else x

def parse_date(date_s):
    return datetime.strptime(date_s, "%Y-%m-%d").date()

def get_fixtures(key, date_s, timezone="America/Lima", league=None, season=None):
    params={"date":date_s, "timezone":timezone}
    if league: params["league"]=league
    if season: params["season"]=season
    return api_get("/fixtures", params, key).get("response", [])

def get_statistics(key, fixture_id):
    return api_get("/fixtures/statistics", {"fixture":fixture_id}, key).get("response", [])

def get_odds(key, fixture_id):
    return api_get("/odds", {"fixture":fixture_id}, key).get("response", [])

def stat_value(team_stats, name):
    for item in team_stats.get("statistics", []):
        if str(item.get("type","")).lower() == name.lower():
            return item.get("value")
    return ""

def bookmaker_first_market(odds_response):
    # Returns a compact list of common markets from the first bookmaker/bet found.
    out={}
    if not odds_response:
        return out
    for bookmaker in odds_response[0].get("bookmakers", []):
        for bet in bookmaker.get("bets", []):
            name=str(bet.get("name","")).lower()
            values=bet.get("values", [])
            if "match winner" in name:
                out["1X2"]=values
            elif "double chance" in name:
                out["double_chance"]=values
            elif "goals over/under" in name:
                out["goals"]=values
            elif "corners" in name:
                out.setdefault("corners", values)
            elif "cards" in name:
                out.setdefault("cards", values)
        if out:
            break
    return out

def fill_workbook(xlsx, fixtures, stats_by_fixture, odds_by_fixture):
    wb=load_workbook(xlsx)
    bd=wb["BASE_DATOS"]
    ev=wb["EVALUADOR"]

    # Append raw fixture-level data to BASE_DATOS.
    existing=set()
    for row in bd.iter_rows(min_row=2, values_only=True):
        # date + team + rival
        if row[0] and row[2] and row[4]:
            existing.add((str(row[0]), str(row[2]), str(row[4])))

    for f in fixtures:
        fixture=f["fixture"]; teams=f["teams"]; league=f["league"]
        dt=fixture.get("date","")
        date_s=dt[:10] if dt else ""
        home=teams["home"]["name"]; away=teams["away"]["name"]
        key=(date_s,home,away)
        if key in existing: continue

        stats=stats_by_fixture.get(fixture["id"], [])
        home_stats=next((x for x in stats if x.get("team",{}).get("id")==teams["home"]["id"]),{})
        away_stats=next((x for x in stats if x.get("team",{}).get("id")==teams["away"]["id"]),{})
        vals=[
            date_s, league.get("name",""), home, "LOCAL", away, "",
            "", "", stat_value(home_stats,"Total Shots"), stat_value(away_stats,"Total Shots"),
            stat_value(home_stats,"Total Shots"), stat_value(home_stats,"Shots on Goal"),
            stat_value(home_stats,"Fouls"), stat_value(home_stats,"Yellow Cards"),
            stat_value(home_stats,"Passes"), f"fixture_id={fixture['id']}"
        ]
        bd.append(vals)

    # Put the first fixtures into EVALUADOR as candidate rows, without inventing probabilities.
    next_row=2
    while next_row <= ev.max_row and ev[f"E{next_row}"].value:
        next_row += 1
    for f in fixtures:
        if next_row>301: break
        fixture=f["fixture"]; teams=f["teams"]; league=f["league"]
        ev[f"B{next_row}"]=fixture.get("date","")[:10]
        ev[f"C{next_row}"]=fixture.get("date","")[11:16]
        ev[f"D{next_row}"]=league.get("name","")
        ev[f"E{next_row}"]=f"{teams['home']['name']} vs {teams['away']['name']}"
        ev[f"F{next_row}"]="PENDIENTE"
        ev[f"G{next_row}"]="SELECCIONAR MERCADO"
        # Quote/statistics are intentionally left blank until a specific market is selected.
        next_row += 1

    wb.save(xlsx)

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--date", default=datetime.now().strftime("%Y-%m-%d"))
    ap.add_argument("--league", default=None, help="API-Football league ID, optional")
    ap.add_argument("--season", default=None, help="Season, optional")
    ap.add_argument("--xlsx", default="Excel_V2_2_Bot_Apuestas.xlsx")
    ap.add_argument("--timezone", default="America/Lima")
    ap.add_argument("--details", action="store_true", help="Also fetch fixture statistics and odds")
    args=ap.parse_args()

    key=os.getenv("API_FOOTBALL_KEY")
    if not key:
        print("Falta API_FOOTBALL_KEY. Configúrala como variable de entorno.")
        sys.exit(2)
    if not os.path.exists(args.xlsx):
        print(f"No existe el Excel: {args.xlsx}")
        sys.exit(3)

    fixtures=get_fixtures(key,args.date,args.timezone,args.league,args.season)
    print(f"Partidos encontrados: {len(fixtures)}")

    stats_by={}; odds_by={}
    if args.details:
        for i,f in enumerate(fixtures,1):
            fid=f["fixture"]["id"]
            try:
                stats_by[fid]=get_statistics(key,fid)
            except Exception as e:
                print("stats",fid,e)
            try:
                odds_by[fid]=get_odds(key,fid)
            except Exception as e:
                print("odds",fid,e)
            # Keep a small pause to be gentle with the API.
            time.sleep(0.15)

    fill_workbook(args.xlsx,fixtures,stats_by,odds_by)
    print(f"Excel actualizado: {args.xlsx}")

if __name__=="__main__":
    main()
