"""Remove sample rows from the supplied BIFF8 workbook; preserve formatting records.

Developer tool only: pip install xlrd xlwt. No production data is stored in Git.
MS-XLS sections 2.4.144 (Index), 2.4.78 (DBCell), 2.1.7.20.5 (worksheet).
xlwt is used only to wrap the reviewed Workbook stream in a new OLE container.
"""
import argparse
import hashlib
import json
import struct
from pathlib import Path

import xlrd
from xlwt.CompoundDoc import XlsDoc

ROOT = Path(__file__).resolve().parents[1]
HEADERS = ['Auf. Name', 'Auftr.', 'Pos.', 'Profil', 'Lange', 'Gewicht', 'Lfn nr.', 'Zekon Unterlieferanten']


def records(data, start=0):
    p = start
    while p + 4 <= len(data):
        kind, size = struct.unpack_from('<HH', data, p)
        if not kind and not size:
            break
        yield kind, data[p + 4:p + 4 + size]
        p += 4 + size
        if kind == 0x0A:
            break


def encode(rs):
    return b''.join(struct.pack('<HH', k, len(v)) + v for k, v in rs)


def prepare(source, output):
    book = xlrd.open_workbook(str(source), formatting_info=True, on_demand=True)
    assert book.biff_version == 80
    assert book.sheet_names() == ['Arkusz1', 'Arkusz2', 'Arkusz3']
    assert book.sheet_by_index(0).row_values(0) == HEADERS
    assert all(book.sheet_by_index(i).nrows == 0 for i in (1, 2))
    data = bytes(book.mem[book.base:book.base + book.stream_len])
    global_records = list(records(data))
    # Reject files with external links, names, macros or embedded objects.
    forbidden = {0x18, 0x17, 0x1AE, 0x2F, 0xD3, 0x1C, 0x5D, 0x1B8, 0x6}
    assert not any(k in forbidden for k, _ in global_records)
    headers_sst = struct.pack('<II', 8, 8) + b''.join(
        struct.pack('<HB', len(s), 0) + s.encode('ascii') for s in HEADERS)
    cleaned = []
    for k, v in global_records:
        if k == 0xFC:
            v = headers_sst
        elif k == 0xFF:  # ExtSST offsets are invalid after removing data.
            continue
        elif k == 0x5C:  # WriteAccess contains an author, not formatting.
            v = struct.pack('<HB', 5, 0) + b'ZEKON' + b' ' * (len(v) - 8)
        elif k == 0x3C:
            raise ValueError('Unexpected CONTINUE record; review this reference manually')
        cleaned.append((k, v))
    sheets = []
    format_proof = []
    for i, absolute in enumerate(book._sh_abs_posn):
        original = list(records(data, absolute - book.base))
        assert not any(k in forbidden for k, _ in original)
        rows = [v for k, v in original if k == 0x208 and struct.unpack_from('<H', v)[0] == 0]
        first_cells = [(k, v) for k, v in original if k == 0xFD and struct.unpack_from('<H', v)[0] == 0]
        if i == 0:
            assert len(rows) == 1 and len(first_cells) == 8
            assert [struct.unpack_from('<I', v, 6)[0] for _, v in first_cells] == list(range(8))
        out = []
        for k, v in original:
            if k in {0x208, 0xFD, 0xBD, 0x203, 0xD7}:
                continue
            if k == 0x20B:
                v = bytes(20 if i == 0 else 16)  # Patched after absolute offsets are known.
            if k == 0x200:
                v = struct.pack('<IIHHH', 0, 1 if i == 0 else 0, 0, 8 if i == 0 else 0, 0)
            # Retain the reference's fonts, XFs, widths, row defaults and print setup.
            # Reset selection/scroll only so a newly exported file opens at A1.
            if k == 0x23E:
                v = v[:2] + bytes(4) + v[6:]
            if k == 0x1D:
                v = struct.pack('<BHHHHHHBB', 3, 0, 0, 0, 1, 0, 0, 0, 0)
            out.append((k, v))
            if k == 0x200 and i == 0:
                block = [(0x208, rows[0])] + first_cells
                out.extend(block)
                out.append((0xD7, struct.pack('<IH', len(encode(block)), 0)))
        sheets.append(out)
        fmt = {0x7D, 0x55, 0x225, 0x81, 0x14, 0x15, 0x26, 0x27, 0x28, 0x29, 0xA1}
        old_fmt = encode([(k, v) for k, v in original if k in fmt])
        assert old_fmt == encode([(k, v) for k, v in out if k in fmt])
        format_proof.append(hashlib.sha256(old_fmt).hexdigest())
    cursor = len(encode(cleaned))
    sheet_offsets = []
    for i, rs in enumerate(sheets):
        sheet_offsets.append(cursor)
        position = cursor
        lookup = {}
        for k, v in rs:
            lookup[k] = position
            position += 4 + len(v)
        index = struct.pack('<IIII', 0, 0, 1 if i == 0 else 0, lookup[0x55])
        if i == 0:
            index += struct.pack('<I', lookup[0xD7])
        sheets[i] = [(k, index if k == 0x20B else v) for k, v in rs]
        cursor = position
    offset_iter = iter(sheet_offsets)
    cleaned = [(k, struct.pack('<I', next(offset_iter)) + v[4:] if k == 0x85 else v) for k, v in cleaned]
    output.parent.mkdir(parents=True, exist_ok=True)
    XlsDoc().save(str(output), encode(cleaned) + b''.join(encode(s) for s in sheets))
    check = xlrd.open_workbook(str(output), formatting_info=True)
    assert check.sheet_by_index(0).nrows == 1
    assert check.sheet_by_index(0).row_values(0) == HEADERS
    assert all(check.sheet_by_index(i).nrows == 0 for i in (1, 2))
    assert check.font_list[0].__dict__ == book.font_list[0].__dict__
    original_formats = encode([(k, v) for k, v in global_records if k in {0x31, 0x41E, 0xE0, 0x87D, 0x293}])
    new_formats = encode([(k, v) for k, v in cleaned if k in {0x31, 0x41E, 0xE0, 0x87D, 0x293}])
    assert original_formats == new_formats
    proof = {'ReferenceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
             'TemplateSha256': hashlib.sha256(output.read_bytes()).hexdigest(),
             'SheetLayoutRecordSha256': format_proof, 'Headers': HEADERS,
             'Status': 'Static comparison passed; native Excel open/save test still required.'}
    output.with_suffix('.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n')
    print('PASS: no sample rows or original document metadata; original formatting and print records preserved')


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('source', type=Path)
    p.add_argument('--output', type=Path, default=ROOT / 'addin/assets/zinc-template.xls')
    a = p.parse_args()
    prepare(a.source, a.output)
