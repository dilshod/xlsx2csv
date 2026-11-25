import csv
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from xlsx2csv import Sheet, SharedStrings, Styles, Workbook


class InlinePhoneticTests(unittest.TestCase):
    def test_inline_phonetics_do_not_change_cell_text(self):
        for prefix in ('', 'x:'):
            xml = '''<worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">
<c r="A1" t="inlineStr"><is><t>Tokyo</t><rPh sb="0" eb="5"><t>phonetic</t></rPh></is></c>
<c r="B1" t="inlineStr"><is><r><t>first</t></r><rPh sb="0" eb="5"><t>reading</t></rPh><r><t> last</t></r></is></c>
<c r="C1" t="inlineStr"><is><t>next cell</t></is></c>
</row></sheetData></worksheet>'''
            if prefix:
                import re
                xml = re.sub(r'<(/?)([A-Za-z]+)', r'<\1x:\2', xml)
            output = io.StringIO()
            sheet = Sheet(Workbook(), SharedStrings(), Styles(), io.BytesIO(xml.encode('utf-8')))
            sheet.to_csv(csv.writer(output, lineterminator='\n'))
            self.assertEqual(output.getvalue(), 'Tokyo,first last,next cell\n')


if __name__ == '__main__':
    unittest.main()
