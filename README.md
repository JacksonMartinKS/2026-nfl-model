# NFL 2026 Forecast · Jackson Martin Forecasts

Static site for GitHub Pages: `index.html` + `forecast-data.js` + the brand icons.

## Weekly update
1. Run the sim and get the new `2026_sim - wkN_summary` sheet into the simulator workbook; enter the week's scores in the `26` sheet of the Elo workbook.
2. `python update_data.py "NFL SIMULATOR RESULTS.xlsx" NFL_SPREADSHEET_v2.xlsx`
3. Commit and push `forecast-data.js`. The page picks up every week sheet automatically (table "As of" menu, team trend charts, header stamp).

Deep links: add a team abbreviation as the hash, e.g. `index.html#KC`.
