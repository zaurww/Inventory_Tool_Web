# Inventory Tool Web

- Communicate with the user in Russian. Staff-facing UI and workbooks use Azerbaijani.
- This is the only active repository: `zaurww/Inventory_Tool_Web`, branch `main`.
- Never overwrite `Inventory_Input.xlsx` as a development or test action.
  Use synthetic fixtures under `.work/` and preserve local accounting files.
- Never commit Excel files, backups, outputs, real business data, environments or secrets.
- Canonical Python sources live in the repository root. `dist/py/` is generated
  by `python build_web.py`; do not edit or commit those generated copies.
- Browser UI sources are `dist/index.html`, `dist/app.js`, `dist/worker.js`, `dist/styles.css`.
- Monthly AVCO is implemented. Moving Average is not. Import expenses are
  allocated by purchase value; late expenses restate the original receipt and history.
- Local tests: create `.work/`, then run
  `python -m unittest -v test_inventory test_migration test_web`.
- Build: `python build_web.py`; check JavaScript with `node --check`.
- `.github/workflows/web.yml` tests Python 3.12/3.14 then deploys only validated
  `dist/` to Pages using its built-in token. Do not reintroduce cross-repository sync.
- Commit/push only when requested. Browser QA uses synthetic fixtures and the
  scripts in `tools/browser_fixtures.py` and `tools/browser_smoke.mjs`.
- Bump the application version for every user-facing update. Keep `VERSION` in
  `inventory.py`, the footer and JS/CSS/XLSX query versions in `dist/index.html`, and
  the versions in `README_RU.txt` / `README_AZ.txt` in sync. Build validates UI
  and asset versions. Group related changes in one release.
- `sample_workbooks.py` generates public template/demo downloads during build.
  Never source them from local accounting files or commit generated XLSX files.
