"""Build and validate the static application and its public sample workbooks."""
from pathlib import Path
from html.parser import HTMLParser
from urllib.parse import urlsplit, parse_qs
import shutil
from inventory import VERSION
from sample_workbooks import public_workbooks, workbook_parts, TEMPLATE_FILE, DEMO_FILE


ROOT = Path(__file__).resolve().parent
TARGET = ROOT / "dist" / "py"
SOURCES = ("browser_api.py", "inventory.py", "input_layout.py", "prepare_input.py",
           "migrate_input.py", "report_locale.py")
PUBLIC_FILES = frozenset(('index.html', 'app.js', 'worker.js', 'styles.css', '.nojekyll',
                          TEMPLATE_FILE, DEMO_FILE,
                          *(f'py/{name}' for name in SOURCES)))


def validate_site(directory):
    """Fail closed: no accounting files, secrets, or arbitrary extra assets."""
    directory = Path(directory)
    entries = list(directory.rglob('*'))
    if any(p.is_symlink() for p in entries):
        raise ValueError('Symlinks are forbidden in the public site')
    actual = {p.relative_to(directory).as_posix() for p in entries if p.is_file()}
    if actual != PUBLIC_FILES:
        raise ValueError(f'Unexpected public files: {sorted(actual - PUBLIC_FILES)}; '
                         f'missing: {sorted(PUBLIC_FILES - actual)}')

    class Assets(HTMLParser):
        def handle_starttag(self, tag, attrs):
            for key, value in attrs:
                if key in ('src', 'href') and value and value.startswith('./'):
                    asset = urlsplit(value)
                    if asset.path[2:] not in PUBLIC_FILES:
                        raise ValueError(f'Unknown local asset: {value}')
                    if asset.path.endswith(('.js', '.css')) and parse_qs(asset.query) != {'v': [VERSION]}:
                        raise ValueError(f'Outdated asset version: {value}')

    html = (directory / 'index.html').read_text(encoding='utf-8')
    Assets().feed(html)
    if f'<span id="app-version">Inventory Tool {VERSION}</span>' not in html:
        raise ValueError('The web interface version must match inventory.VERSION')
    for name, expected in public_workbooks().items():
        try:
            matches = workbook_parts((directory / name).read_bytes()) == workbook_parts(expected)
        except Exception as exc:
            raise ValueError(f'Invalid public workbook: {name}') from exc
        if not matches:
            raise ValueError(f'Public workbook differs from the generated sample: {name}')
    return sorted(actual)


def main():
    TARGET.mkdir(parents=True, exist_ok=True)
    for name in SOURCES:
        shutil.copy2(ROOT / name, TARGET / name)
    for name, content in public_workbooks().items():
        destination = TARGET.parent / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(content)
    validate_site(TARGET.parent)
    print(f"Web sources and public sample workbooks built in {TARGET.parent}")


if __name__ == "__main__":
    main()
