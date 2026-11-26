import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from xlsx2csv import Styles


class StylesTests(unittest.TestCase):
    def test_unknown_numfmt_with_and_without_apply_number_format(self):
        # numFmtId 5 (built-in currency) is not in STANDARD_FORMATS
        xml = b'''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cellXfs count="2"><xf numFmtId="5" applyNumberFormat="1"/><xf numFmtId="5"/></cellXfs>
</styleSheet>'''
        styles = Styles()
        styles.parse(io.BytesIO(xml))
        self.assertEqual(styles.cellXfs, [1, 0])


if __name__ == '__main__':
    unittest.main()
