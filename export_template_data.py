"""Read input data for the development-only workbook designer."""
import json
from pathlib import Path
from inventory import read_input
from input_layout import SHEETS, FIELDS, INPUT_COLUMNS, SETTINGS, TYPES


def export(path, output, demo=False):
    raw, _ = read_input(path)
    if demo:
        for p in raw['Purchases']:
            p.update({'Supplier': 'Demo Supplier', 'Purchase Type': 'Import'})
        for e in raw['Expenses']:
            e.update({'Counterparty': 'Demo Logistics', 'Document': e['Expense ID']})
        raw['Counterparties'] = [{'Name': n} for n in ('Demo Supplier', 'Demo Logistics', 'Customer One', 'Customer Two')]
    raw.setdefault('Counterparties', [])
    payload = dict(data=raw, sheets=SHEETS, fields=FIELDS, columns=INPUT_COLUMNS, settings=SETTINGS, types=TYPES)
    Path(output).write_text(json.dumps(payload, ensure_ascii=False, default=str), encoding='utf-8')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('input')
    parser.add_argument('output')
    parser.add_argument('--demo', action='store_true')
    args = parser.parse_args()
    export(args.input, args.output, args.demo)
