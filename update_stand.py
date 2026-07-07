#!/usr/bin/env python3
"""Werk index.html bij met de stand uit een nieuwe etappe-sheet.

Gebruik:
    python update_stand.py Etappe_3.xlsx
    python update_stand.py Etappe_3.xlsx --etappe 3   (als het nummer niet uit
                                                       bestandsnaam/tabblad blijkt)

Het script leest het algemeen klassement uit de sheet, filtert de gevolgde
teams (de lijst staat in het datablok van index.html) en voegt de etappe toe
aan het JSON-datablok. Bestaat de etappe al, dan wordt die overschreven.
"""
import argparse
import json
import re
import sys
from pathlib import Path

import pandas as pd

HTML = Path(__file__).parent / "index.html"
BLOK = re.compile(
    r'(<script id="tourspel-data" type="application/json">)(.*?)(</script>)',
    re.DOTALL,
)


def etappe_nummer(pad: Path, sheetnaam: str, expliciet: int | None) -> int:
    if expliciet:
        return expliciet
    for bron in (pad.stem, sheetnaam):
        m = re.search(r"(\d+)", bron)
        if m:
            return int(m.group(1))
    sys.exit("Kon het etappenummer niet afleiden; geef het mee met --etappe N")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet", type=Path, help="de dagelijkse Excel-sheet")
    ap.add_argument("--etappe", type=int, default=None)
    args = ap.parse_args()

    html = HTML.read_text(encoding="utf-8")
    m = BLOK.search(html)
    if not m:
        sys.exit("Datablok niet gevonden in index.html")
    data = json.loads(m.group(2))
    teams = data["teams"]

    xls = pd.ExcelFile(args.sheet)
    df, sheetnaam = None, None
    for naam in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=naam, header=None, nrows=8)
        for rij in range(len(raw)):
            waarden = [str(v).strip() for v in raw.iloc[rij].tolist()]
            if "Teamnaam" in waarden and "Plaats" in waarden:
                df = pd.read_excel(xls, sheet_name=naam, header=rij)
                sheetnaam = naam
                break
        if df is not None:
            break
    if df is None:
        sys.exit("Geen tabblad gevonden met kolommen 'Plaats' en 'Teamnaam'")
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(subset=["Plaats", "Teamnaam"])
    df["Teamnaam"] = df["Teamnaam"].astype(str).str.strip()

    stand = {}
    for t in teams:
        rij = df[df["Teamnaam"] == t.strip()]
        if rij.empty:
            print(f"  LET OP: team niet gevonden in de sheet: {t}")
            continue
        r = rij.iloc[0]
        stand[t] = {"pos": int(r["Plaats"]), "pts": round(float(r["Punten"]), 1)}

    nr = etappe_nummer(args.sheet, sheetnaam, args.etappe)
    entry = {"etappe": nr, "totaal": int(df["Plaats"].max()), "stand": stand}

    data["historie"] = [h for h in data["historie"] if h["etappe"] != nr]
    data["historie"].append(entry)
    data["historie"].sort(key=lambda h: h["etappe"])

    nieuw = json.dumps(data, ensure_ascii=False, indent=2)
    html = BLOK.sub(lambda mm: mm.group(1) + "\n" + nieuw + "\n" + mm.group(3), html, count=1)
    HTML.write_text(html, encoding="utf-8")
    print(f"index.html bijgewerkt: etappe {nr}, {len(stand)} teams, veld van {entry['totaal']}.")


if __name__ == "__main__":
    main()
