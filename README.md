# Stock Report Tool (Streamlit)

A simple web app with four options — **Inputting**, **Reconciliation**,
**Variance Report**, and **Ann. B3** — for non-technical users to upload a
raw stock export and download a finished Excel report.

This replaces an earlier Vercel/TypeScript version, which hit a hard
platform limit: Vercel serverless functions cap request/response payloads at
~4.5 MB, and this tool's raw files can be 12–24+ MB. Streamlit runs as a
normal, persistent Python process (no per-request payload ceiling), and
pandas handles large spreadsheets far faster than the JS libraries available
on Vercel. Tested end-to-end on a synthetic 212,000-row (~12 MB) file:

| Report          | Time    | Output size |
|-----------------|---------|-------------|
| Inputting       | ~25s    | ~1.6 MB     |
| Reconciliation  | ~23s    | ~1.6 MB     |
| Variance Report | ~51s    | ~13.4 MB    |

Variance Report is slower because it also copies the full raw sheet
untouched into the output. Expect roughly double those numbers for a 24 MB
file. There's no hard timeout here — worst case, the user just waits a
minute or two with a spinner on screen.

## What each option does

- **Inputting** — pulls every article with stock on hand (SOH ≠ 0) from the
  raw system export, so the manager has a clean list to fill in physical
  counts against.
- **Reconciliation** — collapses the raw report down to one row per article
  code (summing size-level rows), dropping articles with nothing on hand and
  nothing counted.
- **Variance Report** — builds a 3-sheet workbook:
  1. **Stock** — the raw upload, completely untouched.
  2. **Variance** — one row per code, with Variance computed as
     `Counted − SOH` and Total Retail as `Variance × Retail Price`, filtered
     to non-zero variance.
  3. **Variance Report** — the same list with offsetting pairs removed
     (rows whose Total Retail values are exact opposites, e.g. 500 and
     -500), a blank "Reasons for Variance" column for the manager, borders,
     and a bold totals row with live `SUM()` formulas.
- **Ann. B3** — builds the per-store Ann B3 stocktake reconciliation
  workbook: one row per article (Code), grouped by Category (sorted 1-56),
  with letter-suffixed size/variant codes (e.g. "10140410L") merged into
  their base article. Only non-zero Counted articles are kept. A yellow
  subtotal row per category shows Auditor Pairs/Value (auto-summed) with
  Manager mirroring Auditor exactly (so Reconciliation always nets to "-"),
  plus a grand total row and a final AUDITOR/MANAGER/RECONCILIATION summary
  split into Footwear (FW, category < 50) and Non-Footwear (NFW, category
  50-56). The app shows a "processing details" panel after each run listing
  exactly what was dropped and why, since raw exports often mix junk rows
  (gift cards, vouchers, freight) in with real stock data.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

This opens the app at http://localhost:8501. Pick an option, upload a
`.xlsx` file, click **Generate report**, and download the result.

## Deploying

### Option A — Streamlit Community Cloud (free, easiest)

1. Push this folder to a GitHub repo.
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in, click
   **New app**, point it at your repo and `app.py`.
3. Deploy. You get a public URL to share with your team.

Note: Community Cloud's free tier has a default upload limit of 200 MB
(configurable via `.streamlit/config.toml` → `server.maxUploadSize`, already
well above what you need) and modest CPU/RAM. For 24 MB files this should
still be fine based on the benchmarks above, but if it ever feels slow,
that's the tier to look at upgrading.

### Option B — Your own server / VM / Docker

Streamlit is just a Python process, so it runs anywhere Python does:

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

Put it behind a reverse proxy (nginx, Caddy) with HTTPS if it's exposed to
the internet.

## Project structure

```
app.py            -- Streamlit UI (3 options, file upload, download button)
report_logic.py   -- Core report-building logic (pandas + openpyxl), no
                      Streamlit dependency, so it can be tested standalone
requirements.txt
```

## Expected input format

The raw `.xlsx` export should have a sheet with (at least) these columns:
`Code`, `SOH`, `Counted`, `Retail Price`, and ideally `Category` (used for
grouping). Reconciliation additionally expects `Variance` and
`Total Retail` columns (it trusts those as-is; Variance Report ignores
them and computes its own). The tool automatically finds the right sheet
if your workbook has more than one tab, and will tell you which sheets it
found and their headers if none match.
