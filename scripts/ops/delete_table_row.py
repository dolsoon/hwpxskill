#!/usr/bin/env python3
"""Delete a row from an existing <hp:tbl>.

Removes the <hp:tr>...</hp:tr> at the given row index, decrements the
table's rowCnt, and reduces the table's height by the deleted row's max
cellSz height.

Refuses to operate on rows that contain rowSpan>1 cells (deleting those
breaks span accounting). Use --force to bypass the check at your own risk.

Usage:
    # Inspect tables
    python ops/delete_table_row.py input.hwpx --list

    # Delete the third row (rowAddr=2) from table 5
    python ops/delete_table_row.py input.hwpx -o out.hwpx --table 5 --row 2

    # Delete the LAST row
    python ops/delete_table_row.py input.hwpx -o out.hwpx --table 5 --row -1
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, read_zip_entry, run_pitfall_check  # noqa: E402

from lxml import etree  # noqa: E402

NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
SECTION_ENTRY = "Contents/section0.xml"
TBL_OPEN = re.compile(rb"<hp:tbl[\s>]")


def _find_table_byte_range(data: bytes, idx: int) -> tuple[int, int] | None:
    matches = list(TBL_OPEN.finditer(data))
    if idx >= len(matches):
        return None
    start = matches[idx].start()
    cursor = start
    depth = 0
    while True:
        next_open = TBL_OPEN.search(data, cursor + 1)
        next_close = data.find(b"</hp:tbl>", cursor + 1)
        if next_close < 0:
            return None
        if next_open is not None and next_open.start() < next_close:
            depth += 1
            cursor = next_open.start()
            continue
        if depth == 0:
            return (start, next_close + len(b"</hp:tbl>"))
        depth -= 1
        cursor = next_close


def _list_tables(section: bytes) -> str:
    out: list[str] = []
    root = etree.fromstring(section)
    for ti, tbl in enumerate(root.xpath(".//hp:tbl", namespaces=NS)):
        rc = tbl.get("rowCnt")
        cc = tbl.get("colCnt")
        out.append(f"Table {ti}  id={tbl.get('id')}  rowCnt={rc} colCnt={cc}")
    return "\n".join(out)


def _tr_byte_ranges(table_bytes: bytes) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    pos = 0
    while True:
        i = table_bytes.find(b"<hp:tr", pos)
        if i < 0:
            return out
        # ensure it's a tag, not a longer name
        boundary = table_bytes[i + 6 : i + 7]
        if boundary not in (b" ", b">", b"\t", b"\n"):
            pos = i + 1
            continue
        e = table_bytes.find(b"</hp:tr>", i)
        if e < 0:
            return out
        out.append((i, e + len(b"</hp:tr>")))
        pos = e + len(b"</hp:tr>")


def _row_max_height(row_xml: bytes) -> int:
    root = etree.fromstring(b"<hp:wrap xmlns:hp='" + NS["hp"].encode() + b"'>" + row_xml + b"</hp:wrap>")
    heights = []
    for tc in root.xpath(".//hp:tc", namespaces=NS):
        sz = tc.find("hp:cellSz", namespaces=NS)
        if sz is not None:
            try:
                heights.append(int(sz.get("height", "0")))
            except ValueError:
                pass
    return max(heights) if heights else 0


def _row_has_rowspan(row_xml: bytes) -> bool:
    root = etree.fromstring(b"<hp:wrap xmlns:hp='" + NS["hp"].encode() + b"'>" + row_xml + b"</hp:wrap>")
    for tc in root.xpath(".//hp:tc", namespaces=NS):
        cs = tc.find("hp:cellSpan", namespaces=NS)
        if cs is not None and int(cs.get("rowSpan", 1)) != 1:
            return True
    return False


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx")
    p.add_argument("-o", "--output", help="Output .hwpx")
    p.add_argument("--table", type=int, default=0, help="Table index (0-based)")
    p.add_argument("--row", type=int, help="Row index in the table (0-based, -1=last)")
    p.add_argument(
        "--force",
        action="store_true",
        help="Allow deleting a row containing rowSpan>1 cells (unsafe).",
    )
    p.add_argument("--list", action="store_true", help="List tables and exit.")
    p.add_argument("--baseline", help="Optional baseline for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    src = Path(args.input)
    section = read_zip_entry(src, SECTION_ENTRY)

    if args.list:
        print(_list_tables(section))
        return 0
    if args.row is None:
        print("ERROR: --row required when patching.", file=sys.stderr)
        return 2
    if not args.output:
        print("ERROR: --output required when patching.", file=sys.stderr)
        return 2

    tbl_rng = _find_table_byte_range(section, args.table)
    if tbl_rng is None:
        print(f"ERROR: table {args.table} not found.", file=sys.stderr)
        return 2

    tbl_bytes = section[tbl_rng[0] : tbl_rng[1]]
    rows = _tr_byte_ranges(tbl_bytes)
    if not rows:
        print("ERROR: table has no rows.", file=sys.stderr)
        return 2

    idx = args.row
    if idx < 0:
        idx += len(rows)
    if not 0 <= idx < len(rows):
        print(
            f"ERROR: row index {args.row} out of range (0..{len(rows)-1}).",
            file=sys.stderr,
        )
        return 2

    s, e = rows[idx]
    row_xml = tbl_bytes[s:e]

    if not args.force and _row_has_rowspan(row_xml):
        print(
            f"ERROR: row {idx} contains a rowSpan>1 cell. Deleting it would "
            "break span accounting. Pass --force to override.",
            file=sys.stderr,
        )
        return 2

    deleted_height = _row_max_height(row_xml)

    new_table = tbl_bytes[:s] + tbl_bytes[e:]
    # decrement rowCnt
    new_table = re.sub(
        rb'(<hp:tbl[^>]*?\srowCnt=")(\d+)(")',
        lambda m: m.group(1) + str(int(m.group(2)) - 1).encode() + m.group(3),
        new_table,
        count=1,
    )
    if deleted_height > 0:
        new_table = re.sub(
            rb'(<hp:sz[^>]*?\sheight=")(\d+)(")',
            lambda m: m.group(1) + str(max(0, int(m.group(2)) - deleted_height)).encode() + m.group(3),
            new_table,
            count=1,
        )

    def transform(data: bytes) -> bytes:
        return data[: tbl_rng[0]] + new_table + data[tbl_rng[1] :]

    dst = Path(args.output)
    delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    print(
        f"delete_table_row: removed row {idx} of table {args.table}, "
        f"byte delta {delta:+d}, height -{deleted_height}, output={dst}",
        file=sys.stderr,
    )

    if args.no_check:
        return 0
    return run_pitfall_check(
        dst, baseline=Path(args.baseline) if args.baseline else None
    )


if __name__ == "__main__":
    sys.exit(main())
