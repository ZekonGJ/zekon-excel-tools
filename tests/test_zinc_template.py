import hashlib
import json
import re
import struct
import unittest
from pathlib import Path

import xlrd

ROOT = Path(__file__).resolve().parents[1]


def records(data, p=0):
    while p + 4 <= len(data):
        k, n = struct.unpack_from('<HH', data, p)
        yield p, k, data[p + 4:p + 4 + n]
        p += 4 + n
        if k == 0x0A:
            break


class ZincTemplate(unittest.TestCase):
    def setUp(self):
        self.path = ROOT / 'addin/assets/zinc-template.xls'
        self.proof = json.loads(self.path.with_suffix('.json').read_text())
        self.book = xlrd.open_workbook(str(self.path), formatting_info=True, on_demand=True)

    def tearDown(self):
        self.book.release_resources()

    def test_headers_and_no_sample_data(self):
        b = self.book
        self.assertEqual(b.biff_version, 80)
        self.assertEqual(b.sheet_names(), ['Arkusz1', 'Arkusz2', 'Arkusz3'])
        s = b.sheet_by_index(0)
        self.assertEqual((s.nrows, s.ncols), (1, 8))
        self.assertEqual(s.row_values(0), self.proof['Headers'])
        for i in (1, 2):
            self.assertEqual(b.sheet_by_index(i).nrows, 0)
        self.assertEqual(b._sharedstrings, self.proof['Headers'])
        self.assertEqual(b.name_obj_list, [])

    def test_reference_formatting(self):
        b = self.book
        s = b.sheet_by_index(0)
        for c in range(8):
            xf = b.xf_list[s.cell_xf_index(0, c)]
            font = b.font_list[xf.font_index]
            self.assertEqual((font.name, font.height, font.bold), ('Calibri', 220, 0))
            self.assertEqual(b.format_map[xf.format_key].format_str, 'General')
            self.assertEqual(xf.background.fill_pattern, 0)
        self.assertEqual(s.default_row_height, 290)
        self.assertEqual(s.rowinfo_map[0].height, 290)
        self.assertEqual({i: col.width for i, col in s.colinfo_map.items()}, {2: 3490, 3: 3374, 9: 2327})

    def test_biff_offsets_and_exact_layout_records(self):
        b = self.book
        data = bytes(b.mem[b.base:b.base + b.stream_len])
        layouts = {0x7D, 0x55, 0x225, 0x81, 0x14, 0x15, 0x26, 0x27, 0x28, 0x29, 0xA1}
        offsets = [struct.unpack_from('<I', v)[0] for _, k, v in records(data) if k == 0x85]
        self.assertEqual(offsets, [p - b.base for p in b._sh_abs_posn])
        for i, start in enumerate(offsets):
            rs = list(records(data, start))
            by_kind = {k: (p, v) for p, k, v in rs}
            index = by_kind[0x20B][1]
            self.assertEqual(struct.unpack_from('<I', index, 12)[0], by_kind[0x55][0])
            if i == 0:
                dbpos, db = by_kind[0xD7]
                self.assertEqual(struct.unpack_from('<I', index, 16)[0], dbpos)
                self.assertEqual(dbpos - struct.unpack_from('<I', db)[0], by_kind[0x208][0])
            encoded = b''.join(struct.pack('<HH', k, len(v)) + v for _, k, v in rs if k in layouts)
            self.assertEqual(hashlib.sha256(encoded).hexdigest(), self.proof['SheetLayoutRecordSha256'][i])
            self.assertFalse({k for _, k, _ in rs} & {0x6, 0x5D, 0x1B8})

    def test_embedded_bytes_match_reviewed_template(self):
        text = (ROOT / 'addin/src/modZincTemplate.bas').read_text()
        embedded = bytes.fromhex(''.join(re.findall(r'ZincHex f, "([0-9A-F]+)"', text)))
        self.assertEqual(embedded, self.path.read_bytes())
        self.assertEqual(hashlib.sha256(embedded).hexdigest(), self.proof['TemplateSha256'])


if __name__ == '__main__':
    unittest.main()
