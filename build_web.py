"""Copy the current calculation sources into the static Pyodide pilot."""
from pathlib import Path
from html.parser import HTMLParser
import shutil


ROOT = Path(__file__).resolve().parent
TARGET = ROOT / "dist" / "py"
SOURCES = ("browser_api.py", "inventory.py", "input_layout.py", "prepare_input.py",
           "migrate_input.py", "report_locale.py")
PUBLIC_FILES = frozenset(('index.html', 'app.js', 'worker.js', 'styles.css', '.nojekyll',
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
                    if value[2:] not in PUBLIC_FILES:
                        raise ValueError(f'Unknown local asset: {value}')

    Assets().feed((directory / 'index.html').read_text(encoding='utf-8'))
    return sorted(actual)


def main():
    TARGET.mkdir(parents=True, exist_ok=True)
    for name in SOURCES:
        shutil.copy2(ROOT / name, TARGET / name)
    validate_site(TARGET.parent)
    print(f"Web Python sources updated in {TARGET}")


if __name__ == "__main__":
    main()
