# Workbook designer

The installed application uses Python and openpyxl only.
These development utilities reproduce the Russian input layout.

1. Export canonical data with `export_template_data.py INPUT.xlsx data.json`.
   Use `--demo` only for the bundled demo: it supplies fictitious counterparties.
2. With `@oai/artifact-tool` available, run
   `node tools/build_template.mjs data.json NEW.xlsx previews`.
3. Run `python tools/finalize_template.py NEW.xlsx` to disable native row banding.
   The design library exports banding enabled and does not expose a documented
   setter in this environment. This utility changes only that XML flag.
4. Read the exported workbook through `inventory.read_input` and verify the
   calculation before replacing a user's file. Back up the old file first.

Do not commit exported accounting data or workbook previews.
