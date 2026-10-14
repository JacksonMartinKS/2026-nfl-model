#!/usr/bin/env python3
"""Rebuild forecast-data.js for the NFL forecast site from your two workbooks.

    python update_data.py "NFL SIMULATOR RESULTS.xlsx" NFL_SPREADSHEET_v2.xlsx

Options:
    --season 2026            season to read (default 2026)
    --out forecast-data.js   output path (default: next to this script)

What it reads
  Simulator workbook
    * "<season>_sim - pre_summary", "<season>_sim - wk1_summary", ... one sheet per
      week, with the columns your sim writes: team, mean_wins, stdev_wins, poff,
      p_div, p_conf_win, p_sb, exp_seed, p_seed1 ... p_seed7
    * "<yy> Personal" for team names, conference and division
  Elo workbook
    * "<yy>" season sheet: game rows under "<season> Week N" labels with
      away, home, away score, home score in columns B-E (used for W-L-T records)

Every week sheet that exists is included, so after Week 6 runs and its summary sheet
is in the workbook, re-run this and push index.html + forecast-data.js.
Needs openpyxl (pip install openpyxl). Save the workbooks in Excel first so cached
formula values are present.
"""
import argparse
import datetime as dt
import json
import re
from pathlib import Path

from openpyxl import load_workbook

FIELDS = ["mean_wins", "stdev_wins", "poff", "p_div", "p_conf_win", "p_sb"]


def sheet_rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def team_meta(rows):
    teams, seen = [], set()
    for r in rows:
        if len(r) < 5:
            continue
        name, abbr, conf, div = r[0], r[1], r[3], r[4]
        if isinstance(abbr, str) and re.fullmatch(r"[A-Z]{2,3}", abbr) and conf in ("AFC", "NFC") and abbr not in seen:
            teams.append({"name": str(name).strip(), "abbr": abbr, "conf": conf, "div": str(div).strip()})
            seen.add(abbr)
    if len(teams) != 32:
        raise SystemExit(f"Expected 32 teams in the Personal sheet, found {len(teams)}")
    return teams


def week_summary(rows, abbrs, sheet_name):
    h = next((i for i, r in enumerate(rows) if r and str(r[0]).strip().lower() == "team"), None)
    if h is None:
        raise SystemExit(f'No "team" header row in {sheet_name}')
    hdr = [str(c).strip() if c is not None else "" for c in rows[h]]
    out = {}
    for r in rows[h + 1:]:
        if not r or r[0] not in abbrs:
            continue
        d = dict(zip(hdr, r))
        out[r[0]] = [round(float(d[k]), 5) for k in FIELDS] + [[round(float(d[f"p_seed{s}"]), 5) for s in range(1, 8)]]
    missing = abbrs - set(out)
    if missing:
        raise SystemExit(f"{sheet_name} is missing teams: {sorted(missing)}")
    return out


def records_by_week(rows, abbrs, season, last_week):
    """Cumulative [W, L, T] after each week 0..last_week from final scores."""
    games, wk = {}, None
    for r in rows:
        if not r:
            continue
        m = re.match(rf"{season} Week (\d+)", str(r[0] or ""))
        if m:
            wk = int(m.group(1))
        if wk is None or wk > last_week or len(r) < 5:
            continue
        away, home = r[1], r[2]
        if away not in abbrs or home not in abbrs:
            continue
        try:
            sa, sh = float(r[3]), float(r[4])
        except (TypeError, ValueError):
            continue  # unplayed game
        games.setdefault(wk, []).append((away, home, sa, sh))
    rec = {a: [0, 0, 0] for a in abbrs}
    out = {0: {a: v[:] for a, v in rec.items()}}
    for w in range(1, last_week + 1):
        for away, home, sa, sh in games.get(w, []):
            if sa > sh:
                rec[away][0] += 1; rec[home][1] += 1
            elif sh > sa:
                rec[home][0] += 1; rec[away][1] += 1
            else:
                rec[away][2] += 1; rec[home][2] += 1
        out[w] = {a: v[:] for a, v in rec.items()}
    return out, {w: len(g) for w, g in games.items()}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("sim_xlsx")
    ap.add_argument("elo_xlsx")
    ap.add_argument("--season", type=int, default=2026)
    ap.add_argument("--out", default=str(Path(__file__).with_name("forecast-data.js")))
    args = ap.parse_args()
    season, yy = args.season, str(args.season)[2:]

    sim = load_workbook(args.sim_xlsx, read_only=True, data_only=True)
    teams = team_meta(sheet_rows(sim[f"{yy} Personal"]))
    abbrs = {t["abbr"] for t in teams}

    weeks = []
    for key in ["pre"] + [f"wk{n}" for n in range(1, 19)]:
        name = f"{season}_sim - {key}_summary"
        if name not in sim.sheetnames:
            continue
        rows = week_summary(sheet_rows(sim[name]), abbrs, name)
        label = "Preseason" if key == "pre" else f"Week {key[2:]}"
        weeks.append({"key": key, "label": label, "teams": rows})
        print(f"{name}: wins sum {sum(v[0] for v in rows.values()):.1f}, playoff sum {sum(v[2] for v in rows.values()):.2f}")
    if not weeks:
        raise SystemExit(f'No "{season}_sim - ..._summary" sheets found')

    last_week = 0 if weeks[-1]["key"] == "pre" else int(weeks[-1]["key"][2:])
    elo = load_workbook(args.elo_xlsx, read_only=True, data_only=True)
    recs, per_week = records_by_week(sheet_rows(elo[yy]), abbrs, season, last_week)
    print("games with scores per week:", per_week)

    for wk in weeks:
        n = 0 if wk["key"] == "pre" else int(wk["key"][2:])
        for a, row in wk["teams"].items():
            row.append(recs.get(n, recs[0])[a])

    data = {
        "season": season,
        "updated": dt.date.today().isoformat(),
        "fields": FIELDS + ["p_seed1..7", "record_wlt"],
        "teams": teams,
        "weeks": weeks,
    }
    Path(args.out).write_text(
        "/* Jackson Martin Forecasts · NFL forecast data. Generated by update_data.py; do not edit by hand. */\n"
        "window.NFL_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    print(f"Wrote {args.out} ({len(weeks)} weeks, through {weeks[-1]['label']})")


if __name__ == "__main__":
    main()
