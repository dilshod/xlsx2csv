import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from xlsx2csv import SheetNotFoundException, Xlsx2csv, XlsxValueError


class InvalidSheetOutputTests(unittest.TestCase):
    def test_invalid_sheet_preserves_output(self):
        source = Path(__file__).parent / 'sheets.xlsx'
        with tempfile.TemporaryDirectory() as directory:
            for as_path in (False, True):
                existing = Path(directory) / 'existing.csv'
                existing.write_text('keep existing report', encoding='utf-8')
                missing = Path(directory) / 'missing.csv'
                for target in (existing, missing):
                    with Xlsx2csv(str(source)) as converter:
                        with self.assertRaises(XlsxValueError):
                            converter.convert(target if as_path else str(target), sheetid=9999)
                self.assertEqual(existing.read_text(encoding='utf-8'), 'keep existing report')
                self.assertFalse(missing.exists())

    def test_missing_worksheet_part_preserves_output(self):
        source = Path(__file__).parent / 'sheets.xlsx'
        with tempfile.TemporaryDirectory() as directory:
            broken = Path(directory) / 'broken.xlsx'
            with zipfile.ZipFile(str(source)) as src, zipfile.ZipFile(str(broken), 'w') as dst:
                for item in src.infolist():
                    if item.filename != 'xl/worksheets/sheet2.xml':
                        dst.writestr(item, src.read(item.filename))
            existing = Path(directory) / 'existing.csv'
            existing.write_text('keep existing report', encoding='utf-8')
            missing = Path(directory) / 'missing.csv'
            for target in (existing, missing):
                with Xlsx2csv(str(broken)) as converter:
                    with self.assertRaises(SheetNotFoundException):
                        converter.convert(str(target), sheetid=2)
            self.assertEqual(existing.read_text(encoding='utf-8'), 'keep existing report')
            self.assertFalse(missing.exists())


if __name__ == '__main__':
    unittest.main()
