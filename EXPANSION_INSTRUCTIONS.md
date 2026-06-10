# Knicks Minutes Ledger — Expansion Instructions

How this project is built and exactly how to extend it to **more seasons** or **another team**.
Written after the June 2026 expansion that added 2010-11 → 2013-14 (project now spans **2010-11 → 2025-26**, 16 seasons).

---

## 1. What the project is

A single self-contained page (`index.html`) plus a spreadsheet. Both are **generated** — never edit `index.html` by hand.

| File | Role |
|------|------|
| `build_html.py` | Generates `index.html` (the interactive page: minutes board + acquisition ledger + 2025-26 salary/cap view). |
| `build_xlsx.py` | Generates `knicks_top10_minutes_2010-2026.xlsx` (two sheets: "By Season" and "Flat Data"). |
| `index.html` | The published page. **Build output — do not hand-edit.** |
| `knicks_top10_minutes_2010-2026.xlsx` | The spreadsheet. **Build output.** |
| `EXPANSION_INSTRUCTIONS.md` | This file. |

The core dataset = **top-10 New York Knicks players by TOTAL regular-season minutes** for each season, with name, GP, total MIN, MPG, a 5th-element flag, and (for rookies) draft slot; plus the head coach per season.

The same season/coach/draft data is **duplicated** in both scripts. When you change one, change the other identically — **except the 5th-element flag, which differs by script** (`build_html.py` = first-Knicks-year flag for the ★; `build_xlsx.py` = rookie flag). See §3. (A future refactor could move it to a shared `data.py`; until then, edit both.)

---

## 2. Build & verify (run after any data change)

```bash
pip install openpyxl -q            # only needed once, for the xlsx
python build_html.py              # writes index.html next to the script
python build_xlsx.py              # writes the .xlsx next to the script

# validate the workbook has no formula errors (skill script, if available):
python /mnt/skills/public/xlsx/scripts/recalc.py knicks_top10_minutes_2010-2026.xlsx
```

Then sanity-check the HTML (this is the verification that makes the page trustworthy):

```bash
# 1. No template tokens survived (must print 0):
grep -c "__SEASONS__\|__COACHES__\|__LEDGER__\|__SALARY__\|__CAP__\|__SEASONCOLS__\|__GLOSSARY__" index.html
# 2. Each new season string is present:
grep -o "2013-14" index.html | head -1
```

For a deeper check, parse the embedded JSON and confirm `MIN/GP ≈ MPG` for every row and the season count is what you expect (the June 2026 expansion used exactly this — see §6).

---

## 3. Data structures (same numbers, two key formats — and the 5th flag differs!)

### `build_html.py`
```python
seasons = {
  "YYYY-YY": [ (player_name, GP, total_minutes, MPG, first_knicks_year_flag), ... 10 tuples in rank order ],
}
coaches = { "YYYY-YY": "Coach Name" }                 # mid-season change -> "Name A / Name B (int.)"
draft   = { "YYYY-YY|Exact Player Name": "#NN 'YY" }  # rookies ONLY; "Undrafted 'YY" if undrafted
```
> **The 5th element here is the "first season on the Knicks" flag** (drives the ★ + row highlight), **not** the rookie flag. Set it to `1` when that season is the first year of a Knicks stint — i.e. the player was *not* on the roster the prior season (newcomers AND players returning for a new stint, e.g. Felton 2012-13, D. Rose 2020-21). The `draft` dict (rookies only) still controls the draft-slot chip independently.

### `build_xlsx.py` (same numbers, different key/label format — and still a true rookie flag!)
```python
data    = { "YYYY-YY": [ (name, GP, MIN, MPG, rookie_flag), ... ] }   # 5th element = ROOKIE flag here
coaches = { "YYYY-YY": "Coach Name" }                                 # same
draft   = { ("YYYY-YY", "Exact Player Name"): "#NN (YYYY)" }          # TUPLE key, full-year label
```

**Watch the draft format difference:** HTML uses `"2013-14|Tim Hardaway Jr." : "#24 '13"`; xlsx uses `("2013-14","Tim Hardaway Jr.") : "#24 (2013)"`.
The player name in a `draft` key **must exactly match** the name string in the tuple, including punctuation (`Tim Hardaway Jr.`, `Amar'e Stoudemire`, `JR Smith`).

---

## 4. Adding more SEASONS

### Step A — gather the data (the internet-dependent part)
1. **Source:** Fox Sports advanced team stats — `https://www.foxsports.com/nba/new-york-knicks-team-stats?season=YYYY&category=advanced`.
   **Fox uses the season's STARTING year:** `season=2013` → the 2013-14 season.
   **Fallback / cross-check:** Basketball-Reference team page `https://www.basketball-reference.com/teams/NYK/YYYY.html` — note B-Ref uses the **ENDING** year (`/2014.html` = 2013-14); its "Totals" table `MP` column is total minutes.
2. Take the **top 10 by total minutes** (not MPG — a durable high-GP role player can outrank a higher-MPG player who missed games; that's intended).
3. **Two different 5th-element flags — set each per its own rule:**
   - **`build_html.py` → first-Knicks-year flag = 1** if the player was **not on the Knicks the prior season** (true newcomers and players returning for a new stint). Continuous holdovers = 0.
   - **`build_xlsx.py` → rookie flag = 1** only if that season was the player's **first NBA season** — verify each one, don't assume. (Trap: a player's "breakout" year is often his 2nd year, e.g. Jeremy Lin in 2011-12 was *not* a rookie — his rookie year was 2010-11 in Golden State.)
4. For each rookie, get overall pick + draft year (`"#NN 'YY"` / `"#NN (YYYY)"`, or `"Undrafted 'YY"` / `"Undrafted (YYYY)"`).
5. Head coach per season; mid-season change → `"Name A / Name B (int.)"`.

> ⚠️ **Cached-page trap:** stats fetchers sometimes return a stale snapshot of the wrong season. Always confirm the page's season label and that the player names fit the era before trusting numbers. Cross-check across two sources when anything looks off. **Never invent a number** — if a source is unavailable, flag it rather than guess.
>
> ⚠️ **Network-restricted environments:** this repo's web sessions run under a host allowlist — `WebFetch`/`curl` to Fox, B-Ref, Wikipedia, StatMuse all return 403 ("Host not in allowlist"), and `WebSearch` snippets rarely carry exact totals. If you hit this, the data must be gathered in an internet-connected session and handed over as a paste-ready "data pack" (that's how the 2010-14 expansion was done). Sanity-check every handed-over row with `MIN/GP ≈ MPG` before pasting.

### Step B — insert into BOTH scripts
- **Ordering matters.** The page renders seasons in the dict's insertion order, **newest-first**. Keep that.
  - Adding **older** seasons → **append** after the current last entry, in descending order (e.g. 2013-14, then 2012-13, …).
  - Adding a **newer** season → insert it at the **top** of the dict, above the current first entry.
- Add `coaches` entries for the new seasons (both scripts).
- Add `draft` entries for rookies only (both scripts, each in its own key format).

### Step C — update the displayed range (HTML only)
In `build_html.py`:
- `<title>The Knicks Minutes Ledger · 2010–2026</title>`
- the `.sub` paragraph: `Top-10 regular-season minutes leaders for every year, <b>2010–11 → 2025–26</b>`

(These are just display strings; update the years to match the new span. The xlsx filename also encodes the range — see §3 / the `out_path` line in `build_xlsx.py`.)

### Step D — build & verify (§2).

---

## 5. Scope notes — what was intentionally NOT extended

- **Acquisition ledger** ("How They Were Built" cards in `build_html.py`, the `ledger` list): maps how players arrived (trade package / draft / FA). This is **optional** and was **not** extended for the 2010-14 era, because writing trade cards safely requires verifying each transaction package against a live source, which the network-restricted session couldn't do. The Carmelo Anthony 2011 blockbuster is already carded.
  - To add era headliners later (internet-connected session only): **Amar'e Stoudemire** (FA, Jul 2010), **Raymond Felton** (FA, 2010), **Tyson Chandler** (trade from Dallas, 2011), **Landry Fields** (draft #39, 2010), **Iman Shumpert** (draft #17, 2011), **Jeremy Lin** (claimed off waivers, Dec 2011). Verify each exact package before writing a card; principal assets only. Cards auto cross-link from the minutes board by player name (see `NAME2` in `build_html.py`).
- **Salary & Cap view** is intentionally **2025-26 only** — the apron/hard-cap rules it visualizes didn't exist before the 2023 CBA. Do **not** retro-apply it to older seasons.

---

## 6. Adapting to ANOTHER TEAM

The scripts are Knicks-specific in data and styling, but the structure is generic. To fork for another team:

1. **Data:** replace every `seasons`/`data`, `coaches`, and `draft` entry with the new team's (same Fox/B-Ref method, same top-10-by-total-minutes rule, same verification discipline from §4-A).
2. **Source URLs:** swap the team slug — Fox `.../<team-slug>-team-stats?season=...`; B-Ref `/teams/<TLA>/YYYY.html` (e.g. `BOS`, `LAL`).
3. **Branding in `build_html.py`:** the `.kicker`, `<h1>`, `<title>`, and `.sub` text say "Knicks"; the CSS `:root` palette (`--navy`, `--orange`) is Knicks colors — change both.
4. **Ledger** (`ledger` list + `dates` dict) and the **Salary & Cap** block (`salary`, `cap`, `glossary`) are entirely Knicks-specific — rebuild or remove them for a new team. If you remove the ledger, the name-chip cross-linking on the board simply shows nothing extra (harmless).
5. **Output filenames:** `index.html` is generic; rename the `.xlsx` (`out_path` in `build_xlsx.py`) to the new team/range.
6. Build & verify (§2).

---

## 7. Verification discipline (why the data is trustworthy)

1. Confirm every fetched page is the season you asked for (cached-page trap).
2. Cross-check minutes across the page you used and one other source when anything looks off.
3. Treat coach / rookie / draft facts as hypotheses to verify, not givens.
4. Never invent a number. If a source is unavailable, say so and flag the cell rather than guessing.
5. After building: 0 surviving `__PLACEHOLDER__` tokens, expected season count, and `MIN/GP ≈ MPG` on every row.
