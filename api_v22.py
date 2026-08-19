import os
import sys
import time
import argparse
import requests
from datetime import datetime
from openpyxl import load_workbook

BASE_URL = "https://v3.football.api-sports.io"

ALLOWED_COUNTRIES = [
    "England", "Spain", "Italy", "Germany", "France",
    "Netherlands", "Portugal", "Switzerland", "Sweden",
    "Turkey", "Finland", "Iceland",
    "Brazil", "Argentina", "Colombia", "Mexico", "USA",
    "China", "Japan", "South-Korea"
]

GOAL_FOCUS_COUNTRIES = {"Finland", "Iceland"}

PREFERRED_LEAGUES = {
    "England": ["Premier League", "Championship"],
    "Spain": ["La Liga", "Segunda División"],
    "Italy": ["Serie A", "Serie B"],
    "Germany": ["Bundesliga", "2. Bundesliga"],
    "France": ["Ligue 1", "Ligue 2"],
    "Netherlands": ["Eredivisie", "Eerste Divisie"],
    "Portugal": ["Primeira Liga"],
    "Switzerland": ["Super League", "Challenge League"],
    "Sweden": ["Allsvenskan", "Superettan"],
    "Turkey": ["Süper Lig", "1. Lig"],
    "Finland": ["Veikkausliiga", "Ykkösliiga", "Ykkönen"],
    "Iceland": ["Úrvalsdeild", "1. Deild", "2. Deild"],
    "Brazil": ["Serie A", "Serie B"],
    "Argentina": ["Liga Profesional Argentina", "Primera Nacional"],
    "Colombia": ["Primera A", "Primera B"],
    "Mexico": ["Liga MX", "Liga de Expansión MX"],
    "USA": ["Major League Soccer", "USL Championship"],
    "China": ["Super League", "League One"],
    "Japan": ["J1 League", "J2 League", "J3 League"],
    "South-Korea": ["K League 1", "K League 2", "K3 League"]
}

LEAGUE_ALIASES = {
    "Iceland": {
        "Úrvalsdeild": {"Úrvalsdeild", "Besta deild karla", "Besta-deild karla"},
        "1. Deild": {"1. Deild", "1. deild karla", "1. deild"},
        "2. Deild": {"2. Deild", "2. deild karla", "2. deild"},
    },
    "Finland": {
        "Veikkausliiga": {"Veikkausliiga"},
        "Ykkösliiga": {"Ykkösliiga", "Ykkosliiga"},
        "Ykkönen": {"Ykkönen", "Ykkonen", "Ykkönen -"},
    }
}


class RateLimiter:
    def __init__(self, interval=6.5):
        self.interval = interval
        self.last = 0.0

    def wait(self):
        delay = self.interval - (time.time() - self.last)
        if delay > 0:
            time.sleep(delay)
        self.last = time.time()


def api_get(session, limiter, path, params, key, retries=3):
    for attempt in range(retries):
        limiter.wait()

        response = session.get(
            BASE_URL + path,
            headers={"x-apisports-key": key},
            params=params,
            timeout=45
        )

        if response.status_code == 429:
            wait = 15 * (attempt + 1)
            print(f"429: esperando {wait}s...")
            time.sleep(wait)
            continue

        response.raise_for_status()
        data = response.json()

        if data.get("errors"):
            raise RuntimeError(str(data["errors"]))

        return data.get("response", [])

    raise RuntimeError("Límite de API alcanzado después de varios intentos")


def get_stats(session, limiter, key, fixture_id):
    return api_get(
        session,
        limiter,
        "/fixtures/statistics",
        {"fixture": fixture_id},
        key
    )


def get_odds(session, limiter, key, fixture_id):
    return api_get(
        session,
        limiter,
        "/odds",
        {"fixture": fixture_id},
        key
    )


def normalize_name(value):
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().split())


def league_is_preferred(country, league_name):
    name = normalize_name(league_name)

    for preferred in PREFERRED_LEAGUES.get(country, []):
        if name == normalize_name(preferred):
            return True

    for aliases in LEAGUE_ALIASES.get(country, {}).values():
        if name in {normalize_name(x) for x in aliases}:
            return True

    return False


def prepare_stats_rows(fixture, stats_response):
    teams = fixture.get("teams", {})
    home = teams.get("home", {}).get("name", "")
    away = teams.get("away", {}).get("name", "")
    home_id = teams.get("home", {}).get("id")
    away_id = teams.get("away", {}).get("id")

    fixture_id = fixture.get("fixture", {}).get("id")
    date_s = str(fixture.get("fixture", {}).get("date", "") or "")[:10]

    rows = []

    for block in stats_response:
        team_id = block.get("team", {}).get("id")

        if team_id == home_id:
            side = "home"
        elif team_id == away_id:
            side = "away"
        else:
            side = ""

        values = {}

        for item in block.get("statistics", []):
            stat_type = normalize_name(item.get("type"))
            values[stat_type] = item.get("value")

        rows.append({
            "fixture_id": fixture_id,
            "date": date_s,
            "league": fixture.get("league", {}).get("name", ""),
            "country": fixture.get("_country", ""),
            "home": home,
            "away": away,
            "side": side,
            "team": block.get("team", {}).get("name", ""),
            "shots_on_goal": values.get("shots on goal"),
            "shots_off_goal": values.get("shots off goal"),
            "total_shots": values.get("total shots"),
            "blocked_shots": values.get("blocked shots"),
            "corners": values.get("corner kicks"),
            "fouls": values.get("fouls"),
            "offsides": values.get("offsides"),
            "ball_possession": values.get("ball possession"),
            "yellow_cards": values.get("yellow cards"),
            "red_cards": values.get("red cards"),
            "passes": values.get("total passes"),
            "passes_accurate": values.get("passes accurate"),
            "passes_pct": values.get("passes %"),
            "expected_goals": values.get("expected goals"),
            "goals_prevented": values.get("goals prevented")
        })

    return rows


def flatten_odds(fixture, odds_response):
    rows = []

    teams = fixture.get("teams", {})
    home = teams.get("home", {}).get("name", "")
    away = teams.get("away", {}).get("name", "")
    fixture_id = fixture.get("fixture", {}).get("id")
    date_s = str(fixture.get("fixture", {}).get("date", "") or "")[:10]

    for bookmaker_block in odds_response:
        bookmaker = bookmaker_block.get("bookmaker", {})
        bookmaker_name = bookmaker.get("name", "")

        for bet in bookmaker_block.get("bets", []):
            market = bet.get("name", "")

            for value in bet.get("values", []):
                rows.append({
                    "fixture_id": fixture_id,
                    "date": date_s,
                    "league": fixture.get("league", {}).get("name", ""),
                    "country": fixture.get("_country", ""),
                    "home": home,
                    "away": away,
                    "bookmaker": bookmaker_name,
                    "market": market,
                    "selection": value.get("value", ""),
                    "odd": value.get("odd")
                })

    return rows


def ensure_sheet(wb, title, headers):
    if title in wb.sheetnames:
        ws = wb[title]

        if ws.max_row == 1 and all(
            ws.cell(1, c).value is None
            for c in range(1, len(headers) + 1)
        ):
            for c, header in enumerate(headers, 1):
                ws.cell(1, c).value = header

        return ws

    ws = wb.create_sheet(title)

    for c, header in enumerate(headers, 1):
        ws.cell(1, c).value = header

    return ws


def existing_keys(ws, key_cols):
    keys = set()

    if ws.max_row < 2:
        return keys

    header = {
        ws.cell(1, c).value: c
        for c in range(1, ws.max_column + 1)
    }

    if not all(k in header for k in key_cols):
        return keys

    for row in ws.iter_rows(min_row=2, values_only=True):
        key = tuple(
            row[header[k] - 1]
            for k in key_cols
        )

        if any(v not in (None, "") for v in key):
            keys.add(
                tuple(
                    str(v) if v is not None else ""
                    for v in key
                )
            )

    return keys


def append_dict_rows(ws, rows):
    if not rows:
        return 0

    headers = [
        ws.cell(1, c).value
        for c in range(1, ws.max_column + 1)
    ]

    if not any(headers):
        headers = list(rows[0].keys())

        for c, header in enumerate(headers, 1):
            ws.cell(1, c).value = header

    added = 0

    for item in rows:
        ws.append([
            item.get(header, "") if header else ""
            for header in headers
        ])
        added += 1

    return added


def fill_excel(xlsx, fixtures, stats_rows, odds_rows, execution_date):
    wb = load_workbook(xlsx)

    if "BASE_DATOS" not in wb.sheetnames:
        ws_base = wb.create_sheet("BASE_DATOS")
        base_headers = [
            "Fecha",
            "Liga",
            "Local",
            "Condición",
            "Visitante",
            "fixture_id",
            "País"
        ]

        for c, header in enumerate(base_headers, 1):
            ws_base.cell(1, c).value = header
    else:
        ws_base = wb["BASE_DATOS"]

    # Detecta las columnas existentes del Excel y las llena sin destruirlas.
    headers = [
        ws_base.cell(1, c).value
        for c in range(1, ws_base.max_column + 1)
    ]

    header_map = {
        str(h).strip().lower(): i
        for i, h in enumerate(headers)
        if h
    }

    existing = existing_keys(
        ws_base,
        ["Fecha", "Local", "Visitante"]
    )

    added_base = 0

    for fixture in fixtures:
        fx = fixture.get("fixture", {})
        teams = fixture.get("teams", {})
        league = fixture.get("league", {})

        date_s = str(fx.get("date", "") or "")[:10]
        home = teams.get("home", {}).get("name", "")
        away = teams.get("away", {}).get("name", "")
        league_name = league.get("name", "")
        country = fixture.get("_country", "")
        fixture_id = fx.get("id")

        key = (
            str(date_s),
            str(home),
            str(away)
        )

        if key in existing:
            continue

        row = [None] * len(headers)

        values = {
            "fecha": date_s,
            "date": date_s,
            "liga": league_name,
            "league": league_name,
            "local": home,
            "home": home,
            "equipo": home,
            "condición": "LOCAL",
            "condicion": "LOCAL",
            "visitante": away,
            "away": away,
            "rival": away,
            "fixture_id": fixture_id,
            "país": country,
            "pais": country,
            "country": country
        }

        for name, value in values.items():
            if name in header_map:
                row[header_map[name]] = value

        # Si el Excel tiene las columnas originales del motor,
        # dejamos los demás campos vacíos para que el motor los calcule.
        ws_base.append(row)

        existing.add(key)
        added_base += 1

    stats_headers = [
        "fixture_id",
        "date",
        "league",
        "country",
        "home",
        "away",
        "side",
        "team",
        "shots_on_goal",
        "shots_off_goal",
        "total_shots",
        "blocked_shots",
        "corners",
        "fouls",
        "offsides",
        "ball_possession",
        "yellow_cards",
        "red_cards",
        "passes",
        "passes_accurate",
        "passes_pct",
        "expected_goals",
        "goals_prevented"
    ]

    ws_stats = ensure_sheet(
        wb,
        "API_STATS",
        stats_headers
    )

    stats_existing = existing_keys(
        ws_stats,
        ["fixture_id", "team"]
    )

    new_stats = []

    for row in stats_rows:
        key = (
            str(row.get("fixture_id", "")),
            str(row.get("team", ""))
        )

        if key not in stats_existing:
            new_stats.append(row)
            stats_existing.add(key)

    added_stats = append_dict_rows(
        ws_stats,
        new_stats
    )

    odds_headers = [
        "fixture_id",
        "date",
        "league",
        "country",
        "home",
        "away",
        "bookmaker",
        "market",
        "selection",
        "odd"
    ]

    ws_odds = ensure_sheet(
        wb,
        "API_ODDS",
        odds_headers
    )

    odds_existing = existing_keys(
        ws_odds,
        [
            "fixture_id",
            "bookmaker",
            "market",
            "selection"
        ]
    )

    new_odds = []

    for row in odds_rows:
        key = (
            str(row.get("fixture_id", "")),
            str(row.get("bookmaker", "")),
            str(row.get("market", "")),
            str(row.get("selection", ""))
        )

        if key not in odds_existing:
            new_odds.append(row)
            odds_existing.add(key)

    added_odds = append_dict_rows(
        ws_odds,
        new_odds
    )

    control_headers = [
        "Ejecución",
        "Fecha consultada",
        "Partidos encontrados",
        "BASE_DATOS nuevos",
        "API_STATS nuevos",
        "API_ODDS nuevos",
        "Países prioritarios"
    ]

    ws_control = ensure_sheet(
        wb,
        "CONTROL_API",
        control_headers
    )

    ws_control.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        execution_date,
        len(fixtures),
        added_base,
        added_stats,
        added_odds,
        ", ".join(sorted(GOAL_FOCUS_COUNTRIES))
    ])

    wb.save(xlsx)

    print(f"BASE_DATOS nuevos: {added_base}")
    print(f"API_STATS nuevos: {added_stats}")
    print(f"API_ODDS nuevos: {added_odds}")


def main():
    parser = argparse.ArgumentParser(
        description="Bot Apuestas V2.3.2"
    )

    parser.add_argument(
        "--date",
        default=datetime.now().strftime("%Y-%m-%d")
    )

    parser.add_argument(
        "--season",
        type=int,
        default=datetime.now().year
    )

    parser.add_argument(
        "--xlsx",
        default="Excel_V2_2_Motor_Automatico.xlsx"
    )

    # IMPORTANTE:
    # Los detalles quedan ACTIVADOS por defecto.
    parser.add_argument(
        "--details",
        action="store_true",
        default=True
    )

    parser.add_argument(
        "--max-fixtures",
        type=int,
        default=30
    )

    args = parser.parse_args()

    key = os.getenv("API_FOOTBALL_KEY")

    if not key:
        print("ERROR: falta API_FOOTBALL_KEY")
        sys.exit(2)

    if not os.path.exists(args.xlsx):
        print(f"ERROR: no existe {args.xlsx}")
        sys.exit(3)

    session = requests.Session()
    limiter = RateLimiter()

    print(
        f"Consultando partidos del día "
        f"{args.date} / zona America/Lima..."
    )

    fixtures = []

    try:
        all_fixtures = api_get(
            session,
            limiter,
            "/fixtures",
            {
                "date": args.date,
                "timezone": "America/Lima"
            },
            key
        )

        print(
            f"Partidos devueltos por API: "
            f"{len(all_fixtures)}"
        )

        allowed = {
            normalize_name(country): country
            for country in ALLOWED_COUNTRIES
        }

        for fixture in all_fixtures:
            league = fixture.get("league", {})

            country_name = league.get(
                "country",
                ""
            )

            league_name = league.get(
                "name",
                ""
            )

            canonical_country = allowed.get(
                normalize_name(country_name)
            )

            if not canonical_country:
                continue

            if not league_is_preferred(
                canonical_country,
                league_name
            ):
                continue

            fixture["_country"] = canonical_country

            fixture["_priority"] = (
                0
                if canonical_country
                in GOAL_FOCUS_COUNTRIES
                else 1
            )

            fixtures.append(fixture)

            home = fixture.get(
                "teams",
                {}
            ).get(
                "home",
                {}
            ).get(
                "name",
                ""
            )

            away = fixture.get(
                "teams",
                {}
            ).get(
                "away",
                {}
            ).get(
                "name",
                ""
            )

            print(
                f"OK: {canonical_country} | "
                f"{league_name} | "
                f"{home} vs {away}"
            )

    except Exception as error:
        print(
            f"ERROR consulta global fixtures: {error}"
        )
        fixtures = []

    fixtures.sort(
        key=lambda fixture: (
            fixture.get("_priority", 1),
            fixture.get(
                "league",
                {}
            ).get(
                "name",
                ""
            ),
            fixture.get(
                "fixture",
                {}
            ).get(
                "id",
                0
            )
        )
    )

    unique_fixtures = {}

    for fixture in fixtures:
        fixture_id = fixture.get(
            "fixture",
            {}
        ).get(
            "id"
        )

        if fixture_id:
            unique_fixtures[fixture_id] = fixture

    fixtures = list(
        unique_fixtures.values()
    )

    print(
        f"Partidos objetivo después de filtros: "
        f"{len(fixtures)}"
    )

    # Siempre obtiene estadísticas y cuotas de hasta 30 partidos.
    details = fixtures[:args.max_fixtures]

    stats_rows = []
    odds_rows = []

    print(
        f"Consultando detalles de "
        f"{len(details)} partidos..."
    )

    for index, fixture in enumerate(
        details,
        1
    ):
        fixture_id = fixture.get(
            "fixture",
            {}
        ).get(
            "id"
        )

        country = fixture.get(
            "_country",
            ""
        )

        league_name = fixture.get(
            "league",
            {}
        ).get(
            "name",
            ""
        )

        home = fixture.get(
            "teams",
            {}
        ).get(
            "home",
            {}
        ).get(
            "name",
            ""
        )

        away = fixture.get(
            "teams",
            {}
        ).get(
            "away",
            {}
        ).get(
            "name",
            ""
        )

        print(
            f"Detalle {index}/{len(details)} | "
            f"{country} | {league_name} | "
            f"{home} vs {away} | "
            f"{fixture_id}"
        )

        try:
            stats = get_stats(
                session,
                limiter,
                key,
                fixture_id
            )

            stats_rows.extend(
                prepare_stats_rows(
                    fixture,
                    stats
                )
            )

            print(
                f" Stats OK: "
                f"{len(stats)} bloques"
            )

        except Exception as error:
            print(
                f" Stats ERROR {fixture_id}: "
                f"{error}"
            )

        try:
            odds = get_odds(
                session,
                limiter,
                key,
                fixture_id
            )

            odds_rows.extend(
                flatten_odds(
                    fixture,
                    odds
                )
            )

            print(
                f" Odds OK: "
                f"{len(odds)} bookmakers"
            )

        except Exception as error:
            print(
                f" Odds ERROR {fixture_id}: "
                f"{error}"
            )

    fill_excel(
        args.xlsx,
        fixtures,
        stats_
