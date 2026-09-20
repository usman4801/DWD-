# DWD Tool

Minimal Streamlit app: upload the raw daily roster CSV, get back the
finished DWD Excel report, built from your GitHub-hosted template with all
formulas/formatting intact.

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Before it works on your real data

Everything you need to change lives at the top of `app.py`, under
`CONFIGURATION`, plus two functions marked `# ADAPT:`. Specifically:

1. **`TEMPLATE_GITHUB_URL`** — point this at the raw URL of `Blank file.xlsx`
   in your GitHub repo, e.g.
   `https://raw.githubusercontent.com/<org>/<repo>/<branch>/Blank%20file.xlsx`.
   Easiest way to set this without editing code: create
   `.streamlit/secrets.toml` with:
   ```toml
   TEMPLATE_GITHUB_URL = "https://raw.githubusercontent.com/..."
   GITHUB_TOKEN = "ghp_..."   # only needed if the repo is private
   ```

2. **`COL_*` constants** — rename these to match the actual column headers
   in your raw CSV export (employee ID, name, shift status, attendance
   status, date).

3. **`compute_attendance_code()`** — this is where the PL / ABWI / SL /
   Present / OT rules live. It currently implements exactly what you
   described (present on a scheduled Week Off → OT, otherwise present on a
   scheduled shift → P, PL/ABWI/SL pass through as-is, anything else →
   Absent). Adjust the string matching if your raw file spells these
   differently.

4. **`ROSTER_*` layout constants** — these say where on the `Roster` sheet
   the date/day headers live, which column holds employee IDs, which row
   the data starts on, and which column to write the day's attendance code
   into. Update them to match your template's actual layout.

The app never touches cell styles — it only sets `.value` on the specific
cells it's told to update, so every formula, color, border and the whole
`Dashboard` sheet stay exactly as they are in the template.

## Note on scope

I don't have your actual `Blank file.xlsx` or a sample raw CSV, so the
column names, cell layout, and exact attendance-code strings above are
reasonable assumptions based on your spec, not verified against real
files. Once you can share a sample of each (or just the header rows), I
can tighten the constants and mapping logic to match exactly.
