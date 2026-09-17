# Inventory Tool Web

Excel input, Python/openpyxl calculation, static Excel reports. Monthly AVCO;
Moving Average is not implemented. User workbooks stay on their device.

- Staff instructions: [README_AZ.txt](README_AZ.txt)
- Detailed documentation: [README_RU.txt](README_RU.txt)
- Web application: https://zaurww.github.io/Inventory_Tool_Web/

## Development and publication

This is the single source repository. `dist/` contains authored web assets;
`python build_web.py` generates the browser Python copies in ignored `dist/py/`.
Do not edit generated copies. Local preview: `Run_Web_Pilot.bat`.
The build also generates a blank AZ template and a fictional demo with expected
results in ignored `dist/downloads/`. `sample_workbooks.py` is their only source;
no local accounting workbook is read. Both downloads work before Pyodide loads.
Publication validates their XLSX contents against freshly generated samples,
in addition to checking the public-file allowlist. Never commit these XLSX files.

Run `python -m unittest -v test_inventory test_migration test_web` after creating
the local `.work` directory. CI tests Python 3.12 and 3.14.

Work on branch `main`. Review changes, commit, then run `git push`.
GitHub Actions runs the tests, validates the exact public-file allowlist, and
deploys only `dist/` directly to GitHub Pages. Publication uses the built-in
short-lived GitHub Actions token; no deploy key, personal token, or second
repository is needed. Failed tests prevent publication.

The repository contains public source code, instructions and synthetic tests.
Real Excel files, backups, outputs, local secrets and environments are ignored
and must never be committed. Historical private Git commits were not imported.
The independent local checkout is `D:\Inventory_Tool_Web`.

Local-only directories: `.venv`, `.work`, `backups`, `outputs`, `examples`.
Old workbook-design utilities under `tools/` remain useful for reproducing legacy
templates; they are not part of the published application.

## Browser verification

Create synthetic fixtures with `python tools/browser_fixtures.py .work/browser-qa`.
With Playwright available, run `node tools/browser_smoke.mjs SITE_URL .work/browser-qa`.
Use `BROWSER_CHANNEL=msedge` (default) or `chrome`. `PLAYWRIGHT_MODULE` can point
to an existing installed module URL. Finally run
`python tools/browser_fixtures.py .work/browser-qa --verify` to compare the
downloaded results with native Python. Tests include repeat calculation, stable
codes, invalid data, corrupted files, downloads, mobile layout and HTTP traffic.
Only synthetic data is used, and all artifacts stay under ignored `.work/`.
Starter downloads are also checked with JavaScript disabled. The downloaded demo
is calculated in the browser and its report is compared with 14 fixed controls.

Every user-facing release increments `inventory.VERSION` and the matching
footer / JS / CSS versions in `dist/index.html` and staff documentation. The
build rejects mismatched versions; versioned asset URLs also refresh browser caches.
