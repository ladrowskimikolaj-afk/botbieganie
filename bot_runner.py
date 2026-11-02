#!/usr/bin/env python3
"""Inteligentny Dziennik Biegowy: Telegram + Garmin Connect + Google Sheets + Claude.

Tryby (python bot_runner.py <tryb>):
  setup         tworzy zakładki i nagłówki w arkuszu
  garmin-login  jednorazowe logowanie do Garmina lokalnie (obsługuje MFA), zapisuje tokeny w arkuszu
  poll          nasłuchuje wiadomości z Telegrama (~4 min)
  morning       dane z Garmina + krótki poranny komunikat   (--at 07:00)
  evening       ocena dnia -> Kalendarz + Telegram           (--at 21:30)
  monthly       raport za poprzedni miesiąc                  (--at 08:00 --day 1)
  import-runs   [data] ściąga biegi z Garmin Connect do Dziennika
  test-garmin   wypisuje dane z Garmina z dzisiaj
"""
import json
import os
import re
import statistics
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import requests
from gspread.utils import rowcol_to_a1

TZ = ZoneInfo("Europe/Warsaw")
ENV = os.environ
MODEL = ENV.get("CLAUDE_MODEL", "claude-sonnet-5-5")
GEMINI_MODEL = ENV.get("GEMINI_MODEL", "gemini-3.8-flash")
POLL_SECONDS = int(ENV.get("POLL_SECONDS", "240"))

GARMIN, DZ, KAL, GEAR, STAN = "Dane Garmin", "Dziennik Treningów", "Kalendarz", "Sprzęt", "Stan"
SCHEMA = {
    GARMIN: ["Data", "Sen (h)", "Wynik snu", "Body Battery rano", "Body Battery min",
             "Tetno spoczynkowe", "Stres sr.", "Stres max", "Kroki"],
    DZ: ["Data", "Godzina", "Status", "Dystans km", "Czas min", "Tempo min/km", "Typ",
         "Samopoczucie 1-10", "Bol / uwagi", "Buty", "Przelozone na", "Notatka", "Wiadomosc oryginalna", "Tetno sr.", "Zrodlo"],
    KAL: ["Data", "Ocena 1-10", "Ocena slowna", "Sen (h)", "Body Battery rano", "Stres sr.",
          "Kroki", "Trening", "Komentarz trenera", "Notatki dnia"],
    GEAR: ["Nazwa", "Typ", "Data zakupu", "Limit km", "Przebieg startowy km",
           "Przebieg km (auto)", "Zuzycie %", "Status"],
    STAN: ["Klucz", "Wartosc"],
}
DEFAULT_SHOE_LIMIT = 700


def log(*a):
    print(datetime.now(TZ).strftime("%H:%M:%S"), *a, flush=True)


def now():
    return datetime.now(TZ)


def today():
    return now().date()


# ───────────────────────── Google Sheets ─────────────────────────
_sh = None
ARG_SEP = ","


def _loc(f):
    """Formuły wysyłane przez API są parsowane wg ustawień regionalnych arkusza (PL = średniki)."""
    if ARG_SEP == "," or not isinstance(f, str) or not f.startswith("="):
        return f
    out, q = [], False
    for ch in f:
        if ch == '"':
            q = not q
        out.append(";" if (ch == "," and not q) else ch)
    return "".join(out)


def sheet():
    global _sh, ARG_SEP
    if _sh is None:
        raw = ENV.get("GOOGLE_CREDS_JSON")
        if not raw:  # lokalnie: plik google_creds.json obok skryptu
            with open("google_creds.json", encoding="utf-8") as f:
                raw = f.read()
        creds = json.loads(raw)
        _sh = gspread.service_account_from_dict(creds).open_by_key(ENV["SHEET_ID"])
        try:
            loc = getattr(_sh, "locale", "") or ""
        except Exception:
            loc = ""
        ARG_SEP = "," if (not loc or loc[:2] in ("en", "ja", "zh", "ko", "he", "th")) else ";"
    return _sh


def ws(name):
    return sheet().worksheet(name)


def records(name):
    return ws(name).get_all_records(value_render_option="UNFORMATTED_VALUE")


def num(x):
    try:
        return float(x) if x not in ("", None) else None
    except (TypeError, ValueError):
        return None


def blank(v):
    return "" if v is None else v


def state_get(key):
    for row in ws(STAN).get_all_values()[1:]:
        if row and row[0] == key:
            return row[1] if len(row) > 1 else ""
    return None


def state_set(key, value):
    w = ws(STAN)
    keys = w.col_values(1)
    if key in keys:
        w.update(range_name=f"B{keys.index(key) + 1}", values=[[value]], value_input_option="RAW")
    else:
        w.append_row([key, value], value_input_option="RAW")


def setup():
    sh = sheet()
    existing = {w.title for w in sh.worksheets()}
    for name, headers in SCHEMA.items():
        w = sh.worksheet(name) if name in existing else sh.add_worksheet(name, 2000, len(headers) + 2)
        if w.row_values(1) != headers:
            w.update(range_name="A1", values=[headers])
            w.freeze(rows=1)
    # formuły przebiegu sprzętu (wiersze 2-51)
    g = ws(GEAR)
    rows = []
    for r in range(2, 52):
        f = (f"=IF(A{r}=\"\",\"\",E{r}+SUMIF('{DZ}'!J:J,A{r},'{DZ}'!D:D))")
        z = f"=IF(OR(A{r}=\"\",D{r}=\"\"),\"\",ROUND(F{r}/D{r}*100,0))"
        rows.append([_loc(f), _loc(z)])
    g.update(range_name="F2:G51", values=rows, value_input_option="USER_ENTERED")
    if "Arkusz1" in existing and len(sh.worksheets()) > 1:
        try:
            sh.del_worksheet(sh.worksheet("Arkusz1"))
        except Exception:
            pass
    try:
        style_sheets()
    except Exception as e:
        log("Formatowanie arkusza nie wyszło (dane są OK):", repr(e))
    try:
        build_dashboard()
    except Exception as e:
        log("Pulpit nie wyszedł (dane są OK):", repr(e))
    log("Arkusz gotowy.")


def _rgb(r, g, b):
    return {"red": r, "green": g, "blue": b}


WIDTHS = {
    GARMIN: [100, 80, 85, 130, 130, 140, 80, 80, 80],
    DZ: [100, 70, 105, 95, 85, 110, 100, 135, 230, 170, 115, 270, 330, 85, 150],
    KAL: [100, 90, 120, 80, 130, 85, 80, 270, 430, 300],
    GEAR: [190, 80, 110, 85, 160, 150, 100, 100],
    STAN: [160, 420],
}
TEXT_COLS = {GARMIN: set(), DZ: {8, 9, 11, 12, 14}, KAL: {7, 8, 9}, GEAR: {0}, STAN: {0, 1}}
NUM_FMT = {GARMIN: {0: "@", 1: "0.0"}, DZ: {0: "@", 1: "@", 3: "0.0", 10: "@"}, KAL: {0: "@", 3: "0.0"},
           GEAR: {5: "0.0", 6: '0"%"'}}


def style_sheets():
    sh = sheet()
    meta = sh.fetch_sheet_metadata()
    reqs = []
    for s in meta["sheets"]:
        pr = s["properties"]
        title, sid = pr["title"], pr["sheetId"]
        if title not in SCHEMA:
            continue
        n, rows = len(SCHEMA[title]), pr["gridProperties"]["rowCount"]
        for b in s.get("bandedRanges", []):
            reqs.append({"deleteBanding": {"bandedRangeId": b["bandedRangeId"]}})
        for i in reversed(range(len(s.get("conditionalFormats", [])))):
            reqs.append({"deleteConditionalFormatRule": {"sheetId": sid, "index": i}})

        def rng(r0, r1, c0, c1):
            return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                    "startColumnIndex": c0, "endColumnIndex": c1}

        # nagłówek
        reqs.append({"repeatCell": {
            "range": rng(0, 1, 0, n),
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(0.13, 0.24, 0.38),
                "textFormat": {"bold": True, "foregroundColor": _rgb(1, 1, 1), "fontSize": 10},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"}})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 38}, "fields": "pixelSize"}})
        # szerokości kolumn
        for c, w in enumerate(WIDTHS[title]):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1},
                "properties": {"pixelSize": w}, "fields": "pixelSize"}})
        # treść: wyrównanie i formaty liczb
        for c in range(n):
            fmt = {"verticalAlignment": "MIDDLE",
                   "horizontalAlignment": "LEFT" if c in TEXT_COLS[title] else "CENTER"}
            fields = "userEnteredFormat(verticalAlignment,horizontalAlignment"
            if c in NUM_FMT.get(title, {}):
                pat = NUM_FMT[title][c]
                fmt["numberFormat"] = {"type": "TEXT"} if pat == "@" else {"type": "NUMBER", "pattern": pat}
                fields += ",numberFormat"
            reqs.append({"repeatCell": {"range": rng(1, rows, c, c + 1),
                                        "cell": {"userEnteredFormat": fmt}, "fields": fields + ")"}})
        # naprzemienne kolory wierszy
        reqs.append({"addBanding": {"bandedRange": {
            "range": rng(0, rows, 0, n),
            "rowProperties": {"headerColor": _rgb(0.13, 0.24, 0.38),
                              "firstBandColor": _rgb(1, 1, 1), "secondBandColor": _rgb(0.94, 0.96, 0.98)}}}})
        # kolorowe skale
        def scale(col, lo, mid, hi, vlo, vmid, vhi):
            reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
                "ranges": [rng(1, rows, col, col + 1)],
                "gradientRule": {
                    "minpoint": {"color": lo, "type": "NUMBER", "value": str(vlo)},
                    "midpoint": {"color": mid, "type": "NUMBER", "value": str(vmid)},
                    "maxpoint": {"color": hi, "type": "NUMBER", "value": str(vhi)}}}}})
        red, yel, grn = _rgb(0.96, 0.62, 0.60), _rgb(1, 0.95, 0.70), _rgb(0.62, 0.84, 0.62)
        if title == KAL:
            scale(1, red, yel, grn, 1, 5, 10)
        if title == DZ:
            scale(7, red, yel, grn, 1, 5, 10)
        if title == GEAR:
            scale(6, grn, yel, red, 0, 50, 100)
        if title == STAN:
            reqs.append({"updateSheetProperties": {"properties": {"sheetId": sid, "hidden": True},
                                                   "fields": "hidden"}})
    sh.batch_update({"requests": reqs})


def set_cells(w, row, values):
    hdr = SCHEMA[w.title]
    w.batch_update([{"range": rowcol_to_a1(row, hdr.index(h) + 1), "values": [[v]]}
                    for h, v in values.items()], value_input_option="RAW")


def diary_row(v):
    return [blank(v.get(h)) for h in SCHEMA[DZ]]


def match_row(rows, ds, km, garmin_only):
    """Szuka wiersza biegu z tego dnia: garmin_only=True -> importu z Garmina bez wiadomości;
    False -> wpisu z wiadomości, jeszcze bez danych z Garmina."""
    best = None
    for i, r in enumerate(rows):
        if str(r.get("Data")) != ds or r.get("Status") != "wykonany":
            continue
        z = str(r.get("Zrodlo", ""))
        is_g = "Garmin:" in z
        if garmin_only and (not is_g or "wiadomosc" in z):
            continue
        if not garmin_only and is_g:
            continue
        rk = num(r.get("Dystans km"))
        if km and rk and abs(rk - km) / max(km, rk) > 0.25:
            continue
        score = abs(rk - km) if (km and rk) else 99
        if best is None or score < best[0]:
            best = (score, i)
    return best[1] if best else None


NEW_RUNS = []


def import_garmin_runs(g, d):
    """Pobiera biegi z Garmin Connect z danego dnia i scala je z wpisami z wiadomości."""
    ds = d.isoformat()
    try:
        acts = g.get_activities_by_date(ds, ds) or []
    except Exception as e:
        log("Garmin aktywności:", repr(e))
        return 0
    runs = [a for a in acts if "run" in str((a.get("activityType") or {}).get("typeKey", "")).lower()]
    if not runs:
        return 0
    w = ws(DZ)
    rows = w.get_all_records(value_render_option="UNFORMATTED_VALUE")
    added = 0
    for a in runs:
        aid = str(a.get("activityId"))
        dist, dur = num(a.get("distance")), num(a.get("duration"))
        if not dist or not dur or any(aid in str(r.get("Zrodlo", "")) for r in rows):
            continue
        km, mins = round(dist / 1000, 2), round(dur / 60, 1)
        pace = sec_pace(dur / (dist / 1000))
        hr = round(num(a.get("averageHR"))) if num(a.get("averageHR")) else None
        start = str(a.get("startTimeLocal") or "")[11:16] or None
        i = match_row(rows, ds, km, garmin_only=False)
        if i is not None:
            upd = {"Godzina": start, "Dystans km": km, "Czas min": mins, "Tempo min/km": pace,
                   "Tetno sr.": hr, "Zrodlo": f"Garmin:{aid}+wiadomosc"}
            set_cells(w, i + 2, {k: v for k, v in upd.items() if v is not None})
            rows[i].update(upd)
        else:
            v = {"Data": ds, "Godzina": start, "Status": "wykonany", "Dystans km": km, "Czas min": mins,
                 "Tempo min/km": pace, "Notatka": "Import z Garmin Connect", "Tetno sr.": hr,
                 "Zrodlo": f"Garmin:{aid}"}
            w.append_row(diary_row(v), value_input_option="RAW")
            rows.append(v)
            NEW_RUNS.append((km, pace, hr, start))
        added += 1
    return added


PULPIT, DANE = "Pulpit", "Wykresy_dane"


def build_dashboard():
    """Zakładka Pulpit: kafelki + 8 wykresów (źródło: ukryta zakładka Wykresy_dane). Odtwarzana przy każdym setup."""
    sh = sheet()
    for name in (PULPIT, DANE):
        try:
            sh.del_worksheet(sh.worksheet(name))
        except Exception:
            pass
    p = sh.add_worksheet(PULPIT, 90, 12, index=0)
    d = sh.add_worksheet(DANE, 40, 20)
    G, Z, K = f"'{GARMIN}'", f"'{DZ}'", f"'{KAL}'"
    UE = "USER_ENTERED"

    def dp(**k):
        k["values"] = [[_loc(c) for c in row] for row in k["values"]]
        d.update(**k)

    # --- dane pod wykresy (ostatnie 30 wpisów, rosnąco po dacie)
    dp(range_name="A1", values=[["Data", "Sen (h)", "Wynik snu", "Body Battery rano",
                                       "Tetno spoczynkowe", "Stres sr.", "Kroki"]], value_input_option=UE)
    dp(range_name="I1", values=[["Data", "Ocena dnia"]], value_input_option=UE)
    dp(range_name="L1", values=[["Data", "Dystans km", "Czas min", "Tetno sr.",
                                       "Tempo (min/km, ulamek dziesietny)"]], value_input_option=UE)
    dp(range_name="A2", value_input_option=UE, values=[[
        f'=IFERROR(SORT(QUERY({G}!A2:I,"select A,B,C,D,F,G,I where A is not null order by A desc limit 30",0),1,TRUE),"")']])
    dp(range_name="I2", value_input_option=UE, values=[[
        f'=IFERROR(SORT(QUERY({K}!A2:B,"select A,B where A is not null and B is not null order by A desc limit 30",0),1,TRUE),"")']])
    dp(range_name="L2", value_input_option=UE, values=[[
        f'=IFERROR(SORT(QUERY({Z}!A2:N,"select A,D,E,N where C=\'wykonany\' and D is not null order by A desc limit 30",0),1,TRUE),"")']])
    dp(range_name="P2", value_input_option=UE, values=[[
        '=ARRAYFORMULA(IF(M2:M31="","",ROUND(N2:N31/M2:M31,2)))']])

    # --- kafelki
    last = lambda col: (f'=IFERROR(LOOKUP(2,1/({G}!{col}2:{col}<>""),{G}!{col}2:{col}),"–")')
    pref = '(YEAR(TODAY())&"-"&TEXT(MONTH(TODAY()),"00")&"*")'
    x = f'(SUMIFS({Z}!E:E,{Z}!A:A,{pref},{Z}!C:C,"wykonany")/SUMIFS({Z}!D:D,{Z}!A:A,{pref},{Z}!C:C,"wykonany"))'
    grid = [[""] * 12 for _ in range(10)]

    def put(r, c, v):
        grid[r - 1][c] = v
    put(1, 0, "🏃 Pulpit biegacza")
    put(2, 0, '="Dane z Garmina odświeżone: "&IFERROR(INDEX(Stan!B:B,MATCH("garmin_last_update",Stan!A:A,0)),"jeszcze nie odświeżano")')
    put(4, 0, "DZIŚ (ostatnie dostępne dane z Garmina)")
    for i, (lab, col) in enumerate([("Sen (h)", "B"), ("Wynik snu", "C"), ("Body Battery rano", "D"),
                                    ("Tętno spoczynkowe", "F"), ("Stres średni", "G"), ("Kroki", "I")]):
        put(5, i * 2, lab)
        put(6, i * 2, last(col))
    put(8, 0, "TEN MIESIĄC")
    month = [
        ("Kilometry", f'=ROUND(SUMIFS({Z}!D:D,{Z}!A:A,{pref},{Z}!C:C,"wykonany"),1)'),
        ("Biegi", f'=COUNTIFS({Z}!A:A,{pref},{Z}!C:C,"wykonany")'),
        ("Średnie tempo", f'=IFERROR(INT{x}&":"&TEXT(ROUND(({x}-INT{x})*60,0),"00"),"–")'),
        ("Średnia ocena dnia", f'=IFERROR(ROUND(AVERAGEIFS({K}!B:B,{K}!A:A,{pref}),1),"–")'),
        ("Przełożone / odpuszczone", f'=COUNTIFS({Z}!A:A,{pref},{Z}!C:C,"przełożony")+COUNTIFS({Z}!A:A,{pref},{Z}!C:C,"odpuszczony")'),
        ("Średni wynik snu", f'=IFERROR(ROUND(AVERAGE({G}!C2:C),0),"–")'),
    ]
    for i, (lab, f) in enumerate(month):
        put(9, i * 2, lab)
        put(10, i * 2, f)
    p.update(range_name="A1", values=[[_loc(c) for c in row] for row in grid], value_input_option=UE)

    sid, did = p.id, d.id
    rgb = _rgb
    navy, card = rgb(0.13, 0.24, 0.38), rgb(0.92, 0.95, 0.99)

    def rg(r0, r1, c0, c1, s=sid):
        return {"sheetId": s, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}

    def fmt(r0, r1, c0, c1, f, fields):
        return {"repeatCell": {"range": rg(r0, r1, c0, c1), "cell": {"userEnteredFormat": f},
                               "fields": "userEnteredFormat(" + fields + ")"}}
    reqs = []
    for r in (0, 1, 3, 7):
        reqs.append({"mergeCells": {"range": rg(r, r + 1, 0, 12), "mergeType": "MERGE_ALL"}})
    for r in (4, 5, 8, 9):
        for c in range(0, 12, 2):
            reqs.append({"mergeCells": {"range": rg(r, r + 1, c, c + 2), "mergeType": "MERGE_ALL"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 0,
                 "endIndex": 12}, "properties": {"pixelSize": 105}, "fields": "pixelSize"}})
    for r, h in ((0, 48), (1, 26), (3, 30), (4, 26), (5, 52), (7, 30), (8, 26), (9, 52)):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": r,
                     "endIndex": r + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})
    mid = {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}
    reqs.append(fmt(0, 1, 0, 12, {**mid, "backgroundColor": navy,
                                  "textFormat": {"bold": True, "fontSize": 20, "foregroundColor": rgb(1, 1, 1)}},
                    "backgroundColor,textFormat,horizontalAlignment,verticalAlignment"))
    reqs.append(fmt(1, 2, 0, 12, {**mid, "textFormat": {"italic": True, "fontSize": 9, "foregroundColor": rgb(.4, .4, .4)}},
                    "textFormat,horizontalAlignment,verticalAlignment"))
    for r in (3, 7):
        reqs.append(fmt(r, r + 1, 0, 12, {"verticalAlignment": "MIDDLE", "backgroundColor": rgb(.82, .88, .96),
                        "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": navy}},
                        "backgroundColor,textFormat,verticalAlignment"))
    for r in (4, 8):
        reqs.append(fmt(r, r + 1, 0, 12, {**mid, "backgroundColor": card,
                        "textFormat": {"fontSize": 10, "foregroundColor": rgb(.35, .4, .5)}},
                        "backgroundColor,textFormat,horizontalAlignment,verticalAlignment"))
    for r in (5, 9):
        reqs.append(fmt(r, r + 1, 0, 12, {**mid, "backgroundColor": card,
                        "textFormat": {"bold": True, "fontSize": 22, "foregroundColor": navy}},
                        "backgroundColor,textFormat,horizontalAlignment,verticalAlignment"))
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"hideGridlines": True},
                 "tabColorStyle": {"rgbColor": navy}}, "fields": "gridProperties.hideGridlines,tabColorStyle"}})
    sh.batch_update({"requests": reqs})

    # --- wykresy (każdy osobno, żeby jeden błąd nie zablokował reszty)
    def src(col):
        return {"sourceRange": {"sources": [rg(0, 31, col, col + 1, did)]}}
    colors = [rgb(.13, .45, .85), rgb(.95, .55, .15), rgb(.2, .65, .35), rgb(.8, .25, .3)]
    charts = [
        ("Regeneracja: Body Battery rano i wynik snu", "LINE", 0, [3, 2], 12, 0),
        ("Sen (godziny)", "COLUMN", 0, [1], 12, 6),
        ("Tętno spoczynkowe", "LINE", 0, [4], 28, 0),
        ("Stres średni", "LINE", 0, [5], 28, 6),
        ("Kroki dziennie", "COLUMN", 0, [6], 44, 0),
        ("Ocena dnia (1-10)", "COLUMN", 8, [9], 44, 6),
        ("Dystans biegów (km)", "COLUMN", 11, [12], 60, 0),
        ("Tempo biegów (min/km, ułamek dziesiętny)", "LINE", 11, [15], 60, 6),
    ]
    for title, kind, dom, ser, r, c in charts:
        spec = {"title": title, "hiddenDimensionStrategy": "SHOW_ALL", "basicChart": {
            "chartType": kind, "headerCount": 1,
            "legendPosition": "BOTTOM_LEGEND" if len(ser) > 1 else "NO_LEGEND",
            "axis": [{"position": "BOTTOM_AXIS"}, {"position": "LEFT_AXIS"}],
            "domains": [{"domain": src(dom)}],
            "series": [{"series": src(sc), "targetAxis": "LEFT_AXIS", "colorStyle": {"rgbColor": colors[i % 4]}}
                       for i, sc in enumerate(ser)]}}
        pos = {"overlayPosition": {"anchorCell": {"sheetId": sid, "rowIndex": r, "columnIndex": c},
                                   "widthPixels": 625, "heightPixels": 295, "offsetXPixels": 4, "offsetYPixels": 4}}
        try:
            sh.batch_update({"requests": [{"addChart": {"chart": {"spec": spec, "position": pos}}}]})
        except Exception as e:
            log("Wykres nie wyszedł:", title, repr(e)[:200])
    sh.batch_update({"requests": [{"updateSheetProperties": {"properties": {"sheetId": did, "hidden": True},
                                                             "fields": "hidden"}}]})


def ensure_gear(name):
    w = ws(GEAR)
    names = w.col_values(1)
    if any(n.strip().lower() == name.strip().lower() for n in names[1:]):
        return
    r = len(names) + 1
    w.update(range_name=f"A{r}:E{r}",
             values=[[name, "buty", today().isoformat(), DEFAULT_SHOE_LIMIT, 0]],
             value_input_option="RAW")
    w.update(range_name=f"H{r}", values=[["aktywny"]], value_input_option="RAW")


# ───────────────────────── Telegram ─────────────────────────
API = f"https://api.telegram.org/bot{ENV.get('TELEGRAM_BOT_TOKEN', '')}"


def tg_send(text):
    chat = ENV["TELEGRAM_CHAT_ID"]
    for i in range(0, len(text), 3900):
        requests.post(f"{API}/sendMessage", json={"chat_id": chat, "text": text[i:i + 3900]}, timeout=30)


# ───────────────────────── Claude ─────────────────────────
def ask(system, user, max_tokens=900, json_mode=False):
    """Gemini (darmowy klucz z AI Studio), a gdy go brak - Claude API."""
    key = ENV.get("GEMINI_API_KEY")
    if key:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
        cfg = {"maxOutputTokens": max(max_tokens * 3, 2500)}  # zapas na "myślenie" modelu
        if json_mode:
            cfg["responseMimeType"] = "application/json"
        body = {"systemInstruction": {"parts": [{"text": system}]},
                "contents": [{"role": "user", "parts": [{"text": user}]}],
                "generationConfig": cfg}
        for attempt in range(3):
            r = requests.post(url, headers={"x-goog-api-key": key}, json=body, timeout=90)
            if r.status_code in (429, 500, 503) and attempt < 2:
                time.sleep(8 * (attempt + 1))
                continue
            r.raise_for_status()
            break
        parts = r.json()["candidates"][0]["content"]["parts"]
        return "".join(p.get("text", "") for p in parts if not p.get("thought")).strip()
    import anthropic
    client = anthropic.Anthropic()
    r = client.messages.create(model=MODEL, max_tokens=max_tokens, system=system,
                               messages=[{"role": "user", "content": user}])
    return "".join(b.text for b in r.content if b.type == "text").strip()


def ask_json(system, user, max_tokens=900):
    t = ask(system, user, max_tokens, json_mode=True)
    m = re.search(r"\{.*\}", t, re.S)
    return json.loads(m.group(0))


COACH = (
    "Jesteś osobistym trenerem biegowym i asystentem regeneracji użytkownika. Piszesz po polsku, "
    "ciepło i konkretnie, jak wyrozumiały, ale profesjonalny trener: bez sztywniactwa, bez moralizowania, "
    "bez wpędzania w poczucie winy. Opierasz się wyłącznie na danych, które dostajesz (nie zmyślaj liczb). "
    "Nie jesteś lekarzem: przy bólu, który się utrzymuje lub nasila, delikatnie zasugeruj fizjoterapeutę/lekarza."
)

MSG_SYSTEM = COACH + """

Dostajesz JSON z kontekstem (data, dzień tygodnia, dzisiejsze dane z Garmina, ostatnie 7 dni dziennika,
lista sprzętu) oraz wiadomość użytkownika. Zwróć WYŁĄCZNIE obiekt JSON:
{
 "intent": "bieg" | "przelozenie" | "odpuszczenie" | "bol" | "notatka_dnia" | "rozmowa",
 "dystans_km": liczba|null, "czas_min": liczba|null, "tempo": "m:ss"|null,
 "typ": "spokojny"|"tempo"|"interwały"|"długi"|"inny"|null,
 "samopoczucie": 1-10|null, "bol": tekst|null,
 "buty": dokładna nazwa z listy sprzętu, a jeśli to nowy sprzęt - nazwa z wypowiedzi|null,
 "data": "YYYY-MM-DD" (data biegu/zdarzenia; domyślnie dzisiejsza)|null,
 "godzina_biegu": "HH:MM"|null, "tetno_sr": liczba|null,
 "przelozone_na": "YYYY-MM-DD"|null, "notatka": krótki tekst|null,
 "odpowiedz": "MAKSYMALNIE 2-3 krótkie zdania po polsku"
}
Zasady odpowiedzi:
- bieg: zestaw wrażenia z biegu z porannymi danymi Garmina (np. wolniejsze tempo przy niskim Body Battery)
  i daj jedną wskazówkę na kolejny dzień. Jeśli brak danych z Garmina, nie wspominaj o nich.
- przelozenie/odpuszczenie: sprawdź stres, Body Battery i sen. Jeśli dane uzasadniają odpoczynek, powiedz to ze
  zrozumieniem; jeśli nie, nie oceniaj, po prostu wesprzyj decyzję. Zaproponuj konkretny termin (przelozone_na,
  domyślnie jutro).
- bol: potraktuj poważnie, ale spokojnie; zapytaj o jedną rzecz lub poradź lżejszy dzień.
- Data biegu: jeśli użytkownik poda godzinę biegu późniejszą niż aktualna godzina (np. jest 01:00, a bieg był
  o 19:30) albo napisze "wczoraj", użyj poprzedniego dnia. Godzinę biegu wpisz tylko, jeśli ją podał.
- Nie kopiuj całej wiadomości użytkownika do odpowiedzi. Nie używaj list ani nagłówków.
"""


# ───────────────────────── Garmin ─────────────────────────
def save_tokens(tdir):
    files = {}
    for n in os.listdir(tdir):
        with open(os.path.join(tdir, n), encoding="utf-8") as f:
            files[n] = f.read()
    state_set("garmin_tokens", json.dumps(files))


def garmin_client(interactive=False):
    from garminconnect import Garmin
    tdir = tempfile.mkdtemp()
    saved = state_get("garmin_tokens")
    if saved and not interactive:
        for n, c in json.loads(saved).items():
            with open(os.path.join(tdir, n), "w", encoding="utf-8") as f:
                f.write(c)
    mfa = (lambda: input("Kod MFA z Garmina: ")) if interactive else None
    g = Garmin(ENV["GARMIN_EMAIL"], ENV["GARMIN_PASSWORD"], prompt_mfa=mfa) if mfa else \
        Garmin(ENV["GARMIN_EMAIL"], ENV["GARMIN_PASSWORD"])
    g.login(tdir)
    save_tokens(tdir)  # tokeny mogły się odświeżyć
    return g


def fetch_garmin(g, d):
    ds = d.isoformat()
    row = {"Data": ds}
    try:
        s = g.get_stats(ds) or {}
        row["Body Battery rano"] = s.get("bodyBatteryAtWakeTime") or s.get("bodyBatteryHighestValue")
        row["Body Battery min"] = s.get("bodyBatteryLowestValue")
        row["Tetno spoczynkowe"] = s.get("restingHeartRate")
        row["Stres sr."] = s.get("averageStressLevel")
        row["Stres max"] = s.get("maxStressLevel")
        row["Kroki"] = s.get("totalSteps")
    except Exception as e:
        log("Garmin stats:", e)
    try:
        sl = (g.get_sleep_data(ds) or {}).get("dailySleepDTO") or {}
        secs = sl.get("sleepTimeSeconds")
        row["Sen (h)"] = round(secs / 3600, 2) if secs else None
        row["Wynik snu"] = ((sl.get("sleepScores") or {}).get("overall") or {}).get("value")
    except Exception as e:
        log("Garmin sleep:", e)
    return row


def upsert_garmin(g, d):
    new = fetch_garmin(g, d)
    w = ws(GARMIN)
    dates = w.col_values(1)
    headers = SCHEMA[GARMIN]
    if new["Data"] in dates:
        idx = dates.index(new["Data"]) + 1
        cells = [{"range": rowcol_to_a1(idx, c + 1), "values": [[new[h]]]}
                 for c, h in enumerate(headers) if new.get(h) is not None]
        if cells:
            w.batch_update(cells, value_input_option="RAW")
    else:
        w.append_row([blank(new.get(h)) for h in headers], value_input_option="RAW")
    mark_garmin_update()


def mark_garmin_update():
    try:
        state_set("garmin_last_update", now().strftime("%d.%m %H:%M"))
    except Exception as e:
        log("Znacznik aktualizacji:", repr(e))


def garmin_row(ds):
    for r in records(GARMIN):
        if str(r.get("Data")) == ds:
            return r
    return None


def fmt_garmin(r):
    if not r:
        return None
    return {k: r[k] for k in SCHEMA[GARMIN][1:] if r.get(k) not in ("", None)}


# ───────────────────────── Pomocnicze ─────────────────────────
def pace_sec(s):
    try:
        m, sec = str(s).split(":")
        return int(m) * 60 + int(sec)
    except Exception:
        return None


def sec_pace(x):
    x = int(round(x))
    return f"{x // 60}:{x % 60:02d}"


def diary_compact(rows):
    keep = ["Data", "Status", "Dystans km", "Tempo min/km", "Typ", "Samopoczucie 1-10", "Bol / uwagi", "Buty",
            "Przelozone na", "Notatka"]
    return [{k: r[k] for k in keep if r.get(k) not in ("", None)} for r in rows]


def corr(pairs):
    pairs = [(a, b) for a, b in pairs if a is not None and b is not None]
    if len(pairs) < 4:
        return None
    xs, ys = zip(*pairs)
    try:
        return {"r": round(statistics.correlation(xs, ys), 2), "n": len(pairs)}
    except Exception:
        return None


# ───────────────────────── Obsługa wiadomości ─────────────────────────
DAYS = ["poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota", "niedziela"]


def handle_text(text):
    n = now()
    ds = n.date().isoformat()
    since = (n.date() - timedelta(days=7)).isoformat()
    recent = [r for r in records(DZ) if str(r.get("Data")) >= since]
    gear = [r["Nazwa"] for r in records(GEAR) if r.get("Nazwa")]
    ctx = {
        "dzis": ds, "dzien_tygodnia": DAYS[n.weekday()], "godzina": n.strftime("%H:%M"),
        "garmin_dzis": fmt_garmin(garmin_row(ds)),
        "dziennik_7_dni": diary_compact(recent),
        "sprzet": gear,
        "wiadomosc": text,
    }
    d = ask_json(MSG_SYSTEM, json.dumps(ctx, ensure_ascii=False))
    intent = d.get("intent")
    status = {"bieg": "wykonany", "przelozenie": "przełożony", "odpuszczenie": "odpuszczony",
              "bol": "notatka", "notatka_dnia": "notatka"}.get(intent)
    if status:
        km, mins, pace = num(d.get("dystans_km")), num(d.get("czas_min")), d.get("tempo")
        if not pace and km and mins:
            pace = sec_pace(mins * 60 / km)
        shoe = d.get("buty")
        if shoe and intent == "bieg":
            ensure_gear(shoe)
        rd = d.get("data") if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(d.get("data") or "")) else ds
        hhmm = d.get("godzina_biegu") if re.fullmatch(r"\d{1,2}:\d{2}", str(d.get("godzina_biegu") or "")) \
            else n.strftime("%H:%M")
        v = {"Data": rd, "Godzina": hhmm, "Status": status, "Dystans km": km, "Czas min": mins,
             "Tempo min/km": pace, "Typ": d.get("typ"), "Samopoczucie 1-10": d.get("samopoczucie"),
             "Bol / uwagi": d.get("bol"), "Buty": shoe, "Przelozone na": d.get("przelozone_na"),
             "Notatka": d.get("notatka"), "Wiadomosc oryginalna": text, "Tetno sr.": d.get("tetno_sr"),
             "Zrodlo": "wiadomosc"}
        merged = False
        if intent == "bieg":
            try:  # najpierw spróbuj ściągnąć bieg z Garmina i połączyć go z tą wiadomością
                import_garmin_runs(garmin_client(), date.fromisoformat(rd))
            except Exception as e:
                log("Import biegu z Garmina:", repr(e))
            rows = records(DZ)
            i = match_row(rows, rd, km, garmin_only=True)
            if i is not None:
                upd = {k: v[k] for k in ["Typ", "Samopoczucie 1-10", "Bol / uwagi", "Buty", "Notatka",
                                         "Wiadomosc oryginalna"] if v.get(k) not in (None, "")}
                upd["Zrodlo"] = str(rows[i].get("Zrodlo", "")) + "+wiadomosc"
                set_cells(ws(DZ), i + 2, upd)
                merged = True
        if not merged:
            ws(DZ).append_row(diary_row(v), value_input_option="RAW")
    return d.get("odpowiedz") or "Zapisane."


def handle_update(u):
    msg = u.get("message") or {}
    if str((msg.get("chat") or {}).get("id")) != str(ENV["TELEGRAM_CHAT_ID"]):
        return  # ignoruj obcych
    text = msg.get("text")
    if not text:
        if msg.get("voice") or msg.get("audio"):
            tg_send("Głosówek nie obsługuję, napisz mi proszę tekstem.")
        return
    if text.startswith("/start"):
        tg_send("Cześć! Pisz, jak poszedł bieg albo jak wygląda dzień. Resztą się zajmę. "
                "Komendy: /ocena (ocena dnia teraz), /raport (raport za poprzedni miesiąc).")
        return
    if text.startswith("/ocena"):
        evening(force=True)
        return
    if text.startswith("/raport"):
        monthly(force=True)
        return
    tg_send(handle_text(text))


def check_new_runs():
    """Co ~10 min (5:00-24:00): odświeża dzisiejsze dane z Garmina (sen, Body Battery, stres, kroki,
    tętno) i sprawdza nowe biegi; po nowym biegu bot sam pyta o wrażenia."""
    if now().hour < 5:
        return
    last = state_get("last_run_check")
    if last and time.time() - float(last) < 10 * 60:
        return
    state_set("last_run_check", str(time.time()))
    g = garmin_client()
    upsert_garmin(g, today())
    import_garmin_runs(g, today())
    for km, pace, hr, start in NEW_RUNS:
        hr_txt = f", tętno {hr}" if hr else ""
        tg_send(f"🏃 Widzę Twój bieg z Garmina: {km:.1f} km, tempo {pace}{hr_txt}. "
                "Jak się biegło? Napisz 1-2 zdania (samopoczucie, ból, buty).")
    NEW_RUNS.clear()


def run_scheduled_tasks():
    """Poranek / wieczór / raport uruchamiane też z pętli poll, bo cron GitHuba bywa opóźniany o godziny."""
    n = now()
    ds = n.date().isoformat()
    jobs = [
        (7 <= n.hour < 12, "morning_done", ds, morning),
        ((n.hour, n.minute) >= (21, 30), "evening_done", ds, evening),
        (n.day == 1 and n.hour >= 8, "monthly_done", ds[:7], monthly),
    ]
    for due, flag, val, fn in jobs:
        try:
            if due and state_get(flag) != val:
                fn()
        except Exception as e:
            log("Zadanie cykliczne:", fn.__name__, repr(e))


def poll():
    run_scheduled_tasks()
    try:
        check_new_runs()
    except Exception as e:
        log("Sprawdzanie nowych biegów:", repr(e))
    offset = int(state_get("tg_offset") or 0)
    deadline = time.time() + POLL_SECONDS
    while time.time() < deadline:
        try:
            r = requests.get(f"{API}/getUpdates", params={
                "offset": offset, "timeout": 20, "allowed_updates": json.dumps(["message"])}, timeout=40)
            data = r.json()
        except Exception as e:
            log("getUpdates:", e)
            time.sleep(3)
            continue
        if not data.get("ok"):
            log("Telegram:", data)
            time.sleep(5)
            continue
        for u in data.get("result", []):
            offset = u["update_id"] + 1
            try:
                handle_update(u)
            except Exception as e:
                log("Błąd obsługi:", repr(e))
                tg_send(f"⚠️ Nie udało się przetworzyć wiadomości ({type(e).__name__}). Spróbuj jeszcze raz za chwilę.")
            state_set("tg_offset", str(offset))


# ───────────────────────── Poranek ─────────────────────────
def morning(force=False):
    ds0 = today().isoformat()
    if not force and state_get("morning_done") == ds0:
        return
    state_set("morning_done", ds0)
    d = today()
    try:
        g = garmin_client()
        upsert_garmin(g, d - timedelta(days=1))  # domknięcie stresu/kroków wczoraj
        upsert_garmin(g, d)
        import_garmin_runs(g, d - timedelta(days=1))
    except Exception as e:
        log("Garmin rano:", repr(e))
        tg_send("☀️ Dzień dobry! Dziś nie udało mi się połączyć z Garminem. Spróbuję uzupełnić dane wieczorem.")
        return
    row = fmt_garmin(garmin_row(d.isoformat()))
    if not row:
        tg_send("Dzień dobry! Garmin nie oddał dziś danych (może zegarek jeszcze się nie zsynchronizował).")
        return
    txt = ask(COACH + " Napisz poranny komunikat: MAKSYMALNIE 2 krótkie zdania - jak wygląda regeneracja "
                      "i co to znaczy dla dzisiejszego treningu.",
              json.dumps({"dzien": DAYS[now().weekday()], "garmin": row}, ensure_ascii=False), 300)
    tg_send("☀️ " + txt)


# ───────────────────────── Wieczór ─────────────────────────
def evening(force=False):
    ds0 = today().isoformat()
    if not force and state_get("evening_done") == ds0:
        return
    state_set("evening_done", ds0)
    ds = today().isoformat()
    if not force and any(str(r.get("Data")) == ds for r in records(KAL)):
        log("Ocena dnia już istnieje.")
        return
    try:
        g = garmin_client()
        upsert_garmin(g, today())
        import_garmin_runs(g, today())
    except Exception as e:
        log("Garmin wieczorem:", repr(e))
    garm = garmin_row(ds) or {}
    rows = [r for r in records(DZ) if str(r.get("Data")) == ds]
    runs = []
    for r in rows:
        if r["Status"] == "wykonany":
            runs.append(f"{r.get('Dystans km', '')} km @ {r.get('Tempo min/km', '')} (wykonany)")
        elif r["Status"] in ("przełożony", "odpuszczony"):
            runs.append(r["Status"])
    notes = " | ".join(str(r["Notatka"] or r["Bol / uwagi"]) for r in rows
                       if r.get("Notatka") or r.get("Bol / uwagi"))
    d = ask_json(
        COACH + """ Oceń dzień jako całość (sen, regeneracja, stres, aktywność, trening, notatki).
Zwróć WYŁĄCZNIE JSON: {"ocena": 1-10, "slowo": "jedno-dwa słowa", "komentarz": "MAKSYMALNIE 2 krótkie zdania, w tym wskazówka na jutro"}.
Brak treningu przy wysokim stresie lub niskiej regeneracji to nie jest minus. Jeśli brakuje danych, oceń ostrożnie.""",
        json.dumps({"data": ds, "garmin": fmt_garmin(garm), "treningi": runs, "notatki": notes},
                   ensure_ascii=False), 400)
    score = int(d.get("ocena", 5))
    ws(KAL).append_row([
        ds, score, d.get("slowo", ""), blank(garm.get("Sen (h)")), blank(garm.get("Body Battery rano")),
        blank(garm.get("Stres sr.")), blank(garm.get("Kroki")), "; ".join(runs), d.get("komentarz", ""), notes,
    ], value_input_option="RAW")
    tg_send(f"🌙 Ocena dnia: {score}/10, {d.get('slowo', '')}\n{d.get('komentarz', '')}")


# ───────────────────────── Raport miesięczny ─────────────────────────
def monthly(force=False):
    key = today().isoformat()[:7]
    if not force and state_get("monthly_done") == key:
        return
    state_set("monthly_done", key)
    first_this = today().replace(day=1)
    last_prev = first_this - timedelta(days=1)
    start, end = last_prev.replace(day=1).isoformat(), last_prev.isoformat()
    all_diary = records(DZ)
    diary = [r for r in all_diary if start <= str(r.get("Data")) <= end]
    garmin = {str(r["Data"]): r for r in records(GARMIN) if start <= str(r["Data"]) <= end}
    cal = [r for r in records(KAL) if start <= str(r.get("Data")) <= end]

    done = [r for r in diary if r["Status"] == "wykonany"]
    km = sum(num(r["Dystans km"]) or 0 for r in done)
    mins = sum(num(r["Czas min"]) or 0 for r in done if num(r["Dystans km"]))
    stats = {
        "okres": f"{start} - {end}", "suma_km": round(km, 1), "treningi_wykonane": len(done),
        "przelozone": sum(r["Status"] == "przełożony" for r in diary),
        "odpuszczone": sum(r["Status"] == "odpuszczony" for r in diary),
        "srednie_tempo": sec_pace(mins * 60 / km) if km and mins else None,
        "srednia_ocena_dnia": round(statistics.mean(num(r["Ocena 1-10"]) for r in cal if num(r["Ocena 1-10"])), 1)
        if any(num(r["Ocena 1-10"]) for r in cal) else None,
    }

    def avg(col):
        v = [num(r.get(col)) for r in garmin.values() if num(r.get(col)) is not None]
        return round(statistics.mean(v), 1) if v else None

    stats["garmin_srednie"] = {c: avg(c) for c in ["Sen (h)", "Wynik snu", "Body Battery rano",
                                                  "Tetno spoczynkowe", "Stres sr."]}
    runs, sl, bb, st, feel, pace = [], [], [], [], [], []
    for r in done:
        g = garmin.get(str(r["Data"]), {})
        p, f = pace_sec(r.get("Tempo min/km")), num(r.get("Samopoczucie 1-10"))
        runs.append({"data": r["Data"], "km": r["Dystans km"], "tempo": r.get("Tempo min/km"), "typ": r.get("Typ"),
                     "tetno_sr": r.get("Tetno sr.") or None, "samopoczucie": f, "bol": r.get("Bol / uwagi") or None, "buty": r.get("Buty") or None,
                     "sen_h": g.get("Sen (h)"), "wynik_snu": g.get("Wynik snu"), "bb_rano": g.get("Body Battery rano"),
                     "stres": g.get("Stres sr.")})
        sl.append(num(g.get("Wynik snu")))
        bb.append(num(g.get("Body Battery rano")))
        st.append(num(g.get("Stres sr.")))
        feel.append(f)
        pace.append(p)
    day_rating_vs_sleep = [(num(r["Ocena 1-10"]), num(garmin.get(str(r["Data"]), {}).get("Wynik snu")))
                           for r in cal]
    stats["korelacje"] = {
        "wynik_snu_vs_samopoczucie": corr(list(zip(sl, feel))),
        "wynik_snu_vs_tempo_s_na_km": corr(list(zip(sl, pace))),
        "body_battery_vs_samopoczucie": corr(list(zip(bb, feel))),
        "body_battery_vs_tempo_s_na_km": corr(list(zip(bb, pace))),
        "stres_vs_samopoczucie": corr(list(zip(st, feel))),
        "ocena_dnia_vs_wynik_snu": corr(day_rating_vs_sleep),
    }
    # sprzęt: przebieg z całej historii
    gear = []
    for g in records(GEAR):
        if not g.get("Nazwa"):
            continue
        life = (num(g.get("Przebieg startowy km")) or 0) + sum(
            num(r["Dystans km"]) or 0 for r in all_diary if r.get("Buty") == g["Nazwa"] and r["Status"] == "wykonany")
        month = sum(num(r["Dystans km"]) or 0 for r in done if r.get("Buty") == g["Nazwa"])
        gear.append({"nazwa": g["Nazwa"], "km_w_miesiacu": round(month, 1), "km_lacznie": round(life, 1),
                     "limit_km": g.get("Limit km"), "status": g.get("Status")})
    notes = [{"data": r["Data"], "notatka": r.get("Notatka") or r.get("Bol / uwagi")}
             for r in diary if r["Status"] == "notatka"][:25]

    report = ask(
        COACH + """ Napisz rozbudowany raport miesięczny (zwykły tekst bez markdown, ok. 2000-2800 znaków)
w 4 sekcjach z nagłówkami w stylu "1) Liczby":
1) Liczby: km, treningi wykonane vs przełożone/odpuszczone, średnie tempo.
2) Zdrowie a forma: jak sen, Body Battery, stres i oceny dni wiązały się z samopoczuciem i tempem. Korelacje
   to małe próby (pole n) - nazwij je wskazówkami, nie dowodami; jeśli n < 4 albo null, napisz że za mało danych.
3) Sprzęt: przebieg butów, co się zbliża do limitu.
4) Wnioski i rekomendacje na kolejny miesiąc: 3-4 konkretne, realistyczne punkty.
Ton: wspierający trener. Nie wymyślaj danych, których nie ma.""",
        json.dumps({"statystyki": stats, "biegi": runs, "sprzet": gear, "notatki_i_bole": notes},
                   ensure_ascii=False), 2200)
    tg_send(f"📊 Raport miesięczny ({start[:7]})\n\n{report}")


# ───────────────────────── Start ─────────────────────────
def load_env():
    if os.path.exists(".env"):
        for line in open(".env", encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                ENV.setdefault(k.strip(), v.strip().strip('"'))
    global API
    API = f"https://api.telegram.org/bot{ENV.get('TELEGRAM_BOT_TOKEN', '')}"


def in_window(args):
    """--at HH:MM  (okno -5..+50 min, bo GitHub Actions opóźnia crony), --day N"""
    if "--day" in args and now().day != int(args[args.index("--day") + 1]):
        return False
    if "--at" in args:
        h, m = map(int, args[args.index("--at") + 1].split(":"))
        target = now().replace(hour=h, minute=m, second=0, microsecond=0)
        return -5 <= (now() - target).total_seconds() / 60 <= 50
    return True


def main():
    load_env()
    args = sys.argv[1:]
    mode = args[0] if args else "poll"
    if not in_window(args):
        log("Poza oknem czasowym - pomijam.")
        return
    if mode == "setup":
        setup()
    elif mode == "garmin-login":
        garmin_client(interactive=True)
        log("Tokeny Garmina zapisane w arkuszu (zakładka Stan).")
    elif mode == "test-garmin":
        d = today()
        if len(args) > 1 and not args[1].startswith("--"):
            d = date.fromisoformat(args[1])
        g = garmin_client()
        print(fetch_garmin(g, d))
        raw = g.get_stats(d.isoformat()) or {}
        print("Surowe get_stats: pól =", len(raw), list(raw.keys())[:12])
    elif mode == "import-runs":
        d = date.fromisoformat(args[1]) if len(args) > 1 and not args[1].startswith("--") else today()
        log("Zaimportowano biegów:", import_garmin_runs(garmin_client(), d))
    elif mode == "poll":
        poll()
    elif mode == "morning":
        morning(force="--at" not in args)
    elif mode == "evening":
        evening(force="--at" not in args)
    elif mode == "monthly":
        monthly(force="--at" not in args)
    else:
        sys.exit(f"Nieznany tryb: {mode}")


if __name__ == "__main__":
    main()
