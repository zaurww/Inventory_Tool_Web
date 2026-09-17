"""Public deployment boundary: only the explicit application allowlist."""
from pathlib import Path
import tempfile
import unittest
from build_web import PUBLIC_FILES, validate_site


class WebBoundaryTests(unittest.TestCase):
    def setUp(self):
        root = Path(__file__).parent / '.work'
        root.mkdir(exist_ok=True)
        self.tmp = tempfile.TemporaryDirectory(dir=root)
        self.addCleanup(self.tmp.cleanup)
        self.site = Path(self.tmp.name)
        for name in PUBLIC_FILES:
            target = self.site / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text('', encoding='utf-8')

    def test_complete_application_is_allowed(self):
        self.assertEqual(set(validate_site(self.site)), PUBLIC_FILES)

    def test_accounting_or_unexpected_file_blocks_publication(self):
        for name in ('Inventory_Input.xlsx', '.env', 'backups/source.xlsx'):
            with self.subTest(name=name):
                target = self.site / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(b'private data')
                with self.assertRaises(ValueError):
                    validate_site(self.site)
                target.unlink()

    def test_missing_module_blocks_publication(self):
        (self.site / 'py/inventory.py').unlink()
        with self.assertRaises(ValueError):
            validate_site(self.site)


if __name__ == '__main__':
    unittest.main()
