"""Finish native table flags unsupported by the workbook design library.

Run after build_template.mjs. Changes only showRowStripes in table XML.
"""
import argparse
from pathlib import Path
import os
import tempfile
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def finish(path):
    path = Path(path)
    fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=path.parent)
    os.close(fd)
    try:
        with ZipFile(path) as src, ZipFile(temp, 'w') as dst:
            for info in src.infolist():
                content = src.read(info.filename)
                if info.filename.startswith('xl/tables/') and info.filename.endswith('.xml'):
                    root = ET.fromstring(content)
                    style = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tableStyleInfo')
                    if style is not None:
                        style.set('showRowStripes', '0')
                        content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                dst.writestr(info, content)
        os.replace(temp, path)
    finally:
        Path(temp).unlink(missing_ok=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('files', nargs='+')
    for filename in parser.parse_args().files:
        finish(filename)
