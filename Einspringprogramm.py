# Einspringprogramm.py

from collections import defaultdict
from datetime import datetime

SP_WOCHENTAG = 0
SP_DATUM = 1
SP_VORDERGRUND = 3
SP_RUFDIENST = 4

PUNKTE = {"ruf": 0.5, "vouw": 1, "vowo": 2}
WOCHENENDE_TAGE = {4, 5, 6}
ERLAUBTE_FARBEN = {"FFFF0000", "FF0000FF"}

def ist_eingesprungen(zelle):
    if not zelle or not zelle.font or not zelle.font.color:
        return False
    color = zelle.font.color
    return color.type == "rgb" and color.rgb.upper() in ERLAUBTE_FARBEN

def ist_gueltiger_name(name):
    if not name:
        return False
    name = str(name).strip().lower()
    return name not in {"rufdienst", "vordergrund", "gesamt", "vowo", "vouw", ""}

def get_wochentag(row):
    wert = row[SP_WOCHENTAG].value
    if isinstance(wert, str):
        name = wert.strip().lower()
        mapping = {"montag": 0, "dienstag": 1, "mittwoch": 2, "donnerstag": 3,
                   "freitag": 4, "samstag": 5, "sonntag": 6, "mo": 0, "di": 1, "mi": 2,
                   "do": 3, "fr": 4, "sa": 5, "so": 6}
        return mapping.get(name, None)
    elif isinstance(row[SP_DATUM].value, datetime):
        return row[SP_DATUM].value.weekday()
    return None

def berechne_einspringer_statistik(wb):
    daten = defaultdict(lambda: {"ruf": 0, "vouw": 0, "vowo": 0})

    for sheetname in wb.sheetnames:
        if sheetname == "Einspringer-Statistik":
            continue
        ws = wb[sheetname]
        for row in ws.iter_rows(min_row=2):
            wotag = get_wochentag(row)
            ist_we = wotag in WOCHENENDE_TAGE if wotag is not None else False

            zelle_vg = row[SP_VORDERGRUND]
            name_vg = str(zelle_vg.value).strip() if zelle_vg.value else ""
            if ist_gueltiger_name(name_vg) and ist_eingesprungen(zelle_vg):
                if ist_we:
                    daten[name_vg]["vowo"] += 1
                else:
                    daten[name_vg]["vouw"] += 1

            zelle_ruf = row[SP_RUFDIENST]
            name_ruf = str(zelle_ruf.value).strip() if zelle_ruf.value else ""
            if ist_gueltiger_name(name_ruf) and ist_eingesprungen(zelle_ruf):
                daten[name_ruf]["ruf"] += 1

    # Statistik-Blatt neu anlegen/ersetzen
    if "Einspringer-Statistik" in wb.sheetnames:
        del wb["Einspringer-Statistik"]
    ws_stat = wb.create_sheet("Einspringer-Statistik")
    ws_stat.append(["Name", "Ruf", "VoUW", "VoWE", "Gesamt"])

    for name, werte in sorted(daten.items(), key=lambda x: (
        x[1]["ruf"] * PUNKTE["ruf"] +
        x[1]["vouw"] * PUNKTE["vouw"] +
        x[1]["vowo"] * PUNKTE["vowo"]
    ), reverse=True):
        gesamt = (
            werte["ruf"] * PUNKTE["ruf"] +
            werte["vouw"] * PUNKTE["vouw"] +
            werte["vowo"] * PUNKTE["vowo"]
        )
        ws_stat.append([name, werte["ruf"], werte["vouw"], werte["vowo"], round(gesamt, 1)])

    return wb
