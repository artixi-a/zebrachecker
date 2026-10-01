# zebrachecker

A web interface that replaces the pile of one-off Python report scripts. Upload a report,
run a check, preview the result, download the output. Everything the old scripts did is
now one reusable operation.

## What it does

Single file:

- **Full audit (all checks)** - runs GS presence, GS placement, prefix, length and duplicate checks in one go and returns a multi-sheet Excel workbook (`Summary`, `GS Missing`, `GS Placement`, `Prefix Offenders`, `Length Offenders`, `Duplicates`, optional `Passed`). You choose where GS is expected: *anywhere*, *on both sides of a marker* (e.g. `91EE12`), or *directly before a token* (e.g. `93`). The marker/prefix are auto-detected from the file and prefilled.
- **Cleaned export** - Full audit also produces a cleaned file with every failing row removed and duplicates collapsed (keep first / remove all / leave alone), downloadable as Excel, CSV or TSV.
- **GS separator check** - is the ASCII 29 (`\x1d` / `_x001D_`) present, on both sides of a marker (e.g. `91EE12`), or directly before a prefix (e.g. `93`)?
- **Prefix validation** - which codes do NOT start with the required prefix, and optionally remove them.
- **Length audit** - which codes are not the exact expected length?
- **Duplicate check** - report duplicate groups as exact vs corrupted/variation matches, then report-only, remove all, or remove corrupted (keep the row where GS precedes the prefix).
- **Deduplicate (remove)** - by first 31 chars (GTIN + serial) or by full normalized string.
- **Randomize / sample** - sample N, shuffle all, shuffle the middle, scatter the last N, insert dummy codes.
- **Filter / clean** - minimum length, split jammed codes on a prefix, remove rows by prefix, optional dedupe.

Two files:

- **Compare** - only-in-file-1 / only-in-file-2 / common / unique-to-either, with an option to export only unique codes (remove duplicates within the result) and choose first/last when a code repeats.
- **Merge validated** - filter valid codes from file 2 and append the ones not already in file 1.
- **Combine + sample** - append a random sample from file 2 into file 1.

All the hardcoded values from the old scripts (prefixes, `91EE12`, `93`, lengths, sample
counts, seeds) are now inputs in the UI.

## Interface features

- **Product presets** - a sidebar dropdown prefills the code column, required prefix, GS
  location, marker/token, expected length and duplicate basis for each product (Milk/Mleko,
  Nelly/Zebra, Cokolada...). Save your own from the Full audit screen; download/upload
  `presets.json` to share or move them. Built-in presets live in `presets.json`.
- **Pass/Fail verdict** - Full audit shows a green **ALL CHECKS PASSED** or a red
  **ISSUES FOUND** banner listing every failing check, and marks checks you left off as
  `skipped` instead of a misleading `0`.
- **Invisible-character view** - toggle in the sidebar renders GS as a visible token
  (`[GS]` by default) in every table, plus a colour-highlighted preview of the first
  offenders, so `\x1d` problems are obvious at a glance.
- **Auto-detected code column** - the app guesses the code column from the header and
  content; you can override it.
- **Auto-detected GS markers/prefixes** - the app scans the codes and prefills the GS
  marker (e.g. `91EE12`), the GS-adjacent token (e.g. `93`) and the common code prefix
  (e.g. `0104680679601959215`), so different products work without hand-typing.
- **Downloads** - every result offers Excel (`.xlsx`) plus CSV and TSV. Full audit offers a
  multi-sheet `.xlsx` and CSV/TSV zip files (one per sheet).
- **Input formats** - `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.ods`, `.csv`, `.tsv`, `.txt`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501. No login is required locally (auth is only enforced when
secrets are configured).

## Project layout

```
app.py                              Streamlit UI
barcode_tools/core.py               normalization + GS helpers (the shared logic)
barcode_tools/io.py                 robust Excel/CSV readers + safe Excel writer
barcode_tools/ops.py                operations A-H
barcode_tools/presets.py            product preset loading/export
barcode_tools/reports.py            report/table helpers
presets.json                        built-in product presets
tests/test_parity.py                parity tests vs. the original scripts
.streamlit/config.toml              upload size + theme
.streamlit/secrets.toml.example     Google-login template
.github/workflows/tests.yml         CI: run tests on every push
runtime.txt                         Python version for Streamlit Cloud
Dockerfile                          for container hosting
```

## Run the tests

```bash
python -m unittest discover -s tests -v
```

## Performance

Large files use the `python-calamine` engine when available. A 297,000-row / 19 MB report
loads in ~2 s (vs ~37 s with openpyxl), a core-ID dedupe runs in <1 s, and a GS check in
~0.4 s.

## Deploy to Streamlit Community Cloud (free)

1. **GitHub**

   Already set up: code lives at https://github.com/artixi-a/zebrachecker (data files are
   excluded by `.gitignore`, so only code is committed). Tests also run automatically on
   every push via GitHub Actions.

2. **Create the app**

   Go to https://share.streamlit.io, sign in with GitHub, "New app", pick the repo and
   `app.py`, choose a custom subdomain. Leave secrets empty for now to confirm it runs.

3. **Google login (restrict to you and your father)**

   a. In Google Cloud Console (https://console.cloud.google.com): create a project, then
      "APIs & Services" -> "OAuth consent screen" -> External -> add your and your father's
      emails as **Test users**.
   b. "Credentials" -> "Create credentials" -> "OAuth client ID" -> Web application.
      Add the redirect URI shown by Streamlit
      (in the app's "Settings" -> "Secrets" page Streamlit displays the exact callback URL,
      typically `https://<your-subdomain>.streamlit.app/oauth2callback`).
   c. Copy the Client ID and Client secret, and add these secrets in Streamlit Cloud:
      ```toml
      [auth]
      redirect_uri = "https://<your-subdomain>.streamlit.app/oauth2callback"
      cookie_secret = "<a-long-random-string>"
      client_id = "<google-client-id>"
      client_secret = "<google-client-secret>"
      server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
      allowed_emails = ["you@gmail.com", "dad@gmail.com"]
      ```
   d. Reboot the app. Visitors must now sign in with Google, and only the emails in
      `allowed_emails` get in.

   Tip: generate `cookie_secret` with
   `python -c "import secrets; print(secrets.token_hex(32))"`.

4. **Use it**

   Share the app URL with your father. He signs in with Google once per browser session,
   uploads a report, runs an operation, and downloads the result.

## Notes / limitations

- Files are processed in memory per session and never stored on disk.
- Streamlit Community Cloud apps sleep when idle and wake on the next visit (a few seconds).
- For very large files or heavier traffic, the included `Dockerfile` runs the same app on
  Render / Railway / Fly.io unchanged.
