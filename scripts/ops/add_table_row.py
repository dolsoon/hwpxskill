#!/usr/bin/env python3
"""Append a new row to an existing <hp:tbl> by cloning the LAST row.

Cloning the last row is the safe path: cellSz widths, paragraph styles,
and borderFillIDRefs are inherited verbatim. The script then:
- Re-numbers cellAddr rowAddr in the new row to last_row + 1
- Replaces text in the new row's cells per --cell COL=text (each cell's
  first <hp:t> only; other runs left as cloned)
- Bumps the table's rowCnt by 1
- Bumps the table's <hp:sz height> by the cloned row's max cellSz height

Usage:
    # Inspect first to find the table index
    python ops/add_table_row.py input.hwpx --list

    # Append a blank row to table 5 (clone last row, leave text as-is)
    python ops/add_table_row.py input.hwpx -o out.hwpx --table 5

    # Append a row with specific text in selected cells
    python ops/add_table_row.py input.hwpx -o out.hwpx --table 5 \\
        --cell 0=2024 --cell 1=A안 --cell 2=B안 --cell 3="확정 예정"

Limitations:
- Only works on tables whose last row has NO rowSpan>1 cells (else the
  spanned rows would be broken). Reports an error in that case.
- Paragraph ids inside the cloned row are made unique by appending a
  numeric suffix (last_row_id, last_row_id+1, ...). This avoids P1
  duplicate-id violations.
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
    # find matching </hp:tbl>
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


def _extract_last_tr(table_bytes: bytes) -> tuple[int, int] | None:
    last_tr_open = table_bytes.rfind(b"<hp:tr")
    if last_tr_open < 0:
        return None
    last_tr_close = table_bytes.find(b"</hp:tr>", last_tr_open)
    if last_tr_close < 0:
        return None
    return (last_tr_open, last_tr_close + len(b"</hp:tr>"))


def _list_tables(section: bytes) -> str:
    out: list[str] = []
    root = etree.fromstring(section)
    for ti, tbl in enumerate(root.xpath(".//hp:tbl", namespaces=NS)):
        rc = tbl.get("rowCnt")
        cc = tbl.get("colCnt")
        out.append(f"Table {ti}  id={tbl.get('id')}  rowCnt={rc} colCnt={cc}")
    return "\n".join(out)


def _last_row_height_from_xml(table_xml: bytes) -> int:
    root = etree.fromstring(
        b'<wrap xmlns:hp="' + NS["hp"].encode() + b'">' + table_xml + b"</wrap>"
    ).find("hp:tbl", namespaces=NS)
    trs = root.xpath("./hp:tr", namespaces=NS)
    if not trs:
        return 0
    last = trs[-1]
    heights = []
    for tc in last.xpath("./hp:tc", namespaces=NS):
        sz = tc.find("hp:cellSz", namespaces=NS)
        if sz is not None:
            try:
                heights.append(int(sz.get("height", "0")))
            except ValueError:
                pass
    return max(heights) if heights else 0


def _max_para_id(data: bytes) -> int:
    used = re.findall(rb'\sid="(\d+)"', data)
    if not used:
        return 9_999_900_000
    return max(int(x) for x in used)


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx")
    p.add_argument("-o", "--output", help="Output .hwpx")
    p.add_argument("--table", type=int, default=0, help="Table index (0-based)")
    p.add_argument(
        "--cell",
        action="append",
        default=[],
        help="Repeatable: COL=text — replace first <hp:t> in cell at column COL.",
    )
    p.add_argument(
        "--list",
        action="store_true",
        help="List all tables in the document, do not patch.",
    )
    p.add_argument("--baseline", help="Optional baseline for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    src = Path(args.input)
    section = read_zip_entry(src, SECTION_ENTRY)

    if args.list:
        print(_list_tables(section))
        return 0
    if not args.output:
        print("ERROR: --output required when patching.", file=sys.stderr)
        return 2

    tbl_rng = _find_table_byte_range(section, args.table)
    if tbl_rng is None:
        print(f"ERROR: table index {args.table} not found.", file=sys.stderr)
        return 2

    tbl_bytes = section[tbl_rng[0] : tbl_rng[1]]

    # Verify last row has no rowSpan>1
    root = etree.fromstring(
        b'<wrap xmlns:hp="' + NS["hp"].encode() + b'">' + tbl_bytes + b"</wrap>"
    ).find("hp:tbl", namespaces=NS)
    trs = root.xpath("./hp:tr", namespaces=NS)
    if not trs:
        print("ERROR: table has no rows to clone.", file=sys.stderr)
        return 2
    last = trs[-1]
    for tc in last.xpath("./hp:tc", namespaces=NS):
        cs = tc.find("hp:cellSpan", namespaces=NS)
        if cs is not None and int(cs.get("rowSpan", 1)) != 1:
            print(
                "ERROR: last row contains a cell with rowSpan>1; cloning would "
                "violate OWPML. Use replace_section for complex cases.",
                file=sys.stderr,
            )
            return 2

    # Find the last <hp:tr>...</hp:tr> in raw bytes
    tr_rng = _extract_last_tr(tbl_bytes)
    if tr_rng is None:
        print("ERROR: could not locate last <hp:tr> in table.", file=sys.stderr)
        return 2
    tr_bytes = tbl_bytes[tr_rng[0] : tr_rng[1]]

    # Determine new rowAddr
    last_row_addr = -1
    for tc in last.xpath("./hp:tc", namespaces=NS):
        addr = tc.find("hp:cellAddr", namespaces=NS)
        if addr is not None:
            last_row_addr = max(last_row_addr, int(addr.get("rowAddr", "0")))
    new_row_addr = last_row_addr + 1

    # Re-number rowAddr in cloned row
    cloned = re.sub(
        rb'(<hp:cellAddr[^/>]*?\srowAddr=")\d+(")',
        rb"\g<1>" + str(new_row_addr).encode() + rb"\2",
        tr_bytes,
    )

    # Make paragraph ids unique by adding offset
    base_id = _max_para_id(section) + 1
    counter = [0]

    def bump_id(m: re.Match) -> bytes:
        new_id = str(base_id + counter[0]).encode()
        counter[0] += 1
        return m.group(1) + new_id + m.group(2)

    cloned = re.sub(rb'(<hp:p[^>]*?\sid=")\d+(")', bump_id, cloned)

    # Replace text in selected cells per --cell COL=text
    if args.cell:
        col_map: dict[int, bytes] = {}
        for raw in args.cell:
            if "=" not in raw:
                print(f"ERROR: --cell expects COL=text, got {raw!r}", file=sys.stderr)
                return 2
            col, txt = raw.split("=", 1)
            col_map[int(col)] = txt.encode("utf-8")
        # walk cloned row, find each <hp:tc> and its colAddr
        out = bytearray()
        cursor = 0
        for tc_match in re.finditer(rb"<hp:tc[\s>]", cloned):
            tc_open = tc_match.start()
            tc_close = cloned.find(b"</hp:tc>", tc_open)
            if tc_close < 0:
                continue
            tc_close_end = tc_close + len(b"</hp:tc>")
            cell = cloned[tc_open:tc_close_end]
            addr = re.search(rb'<hp:cellAddr[^/>]*?\scolAddr="(\d+)"', cell)
            if addr is None:
                continue
            col = int(addr.group(1))
            if col in col_map:
                # Replace first <hp:t>...</hp:t> inner text
                cell = re.sub(
                    rb"<hp:t>[^<]*</hp:t>",
                    b"<hp:t>" + col_map[col] + b"</hp:t>",
                    cell,
                    count=1,
                )
            out.extend(cloned[cursor:tc_open])
            out.extend(cell)
            cursor = tc_close_end
        out.extend(cloned[cursor:])
        cloned = bytes(out)

    # Insert cloned row right after the existing last <hp:tr>
    new_table_inner = (
        tbl_bytes[: tr_rng[1]] + cloned + tbl_bytes[tr_rng[1] :]
    )

    # Bump rowCnt and tbl height
    new_table_inner = re.sub(
        rb'(<hp:tbl[^>]*?\srowCnt=")(\d+)(")',
        lambda m: m.group(1) + str(int(m.group(2)) + 1).encode() + m.group(3),
        new_table_inner,
        count=1,
    )

    cloned_height = _last_row_height_from_xml(tbl_bytes)
    if cloned_height > 0:
        new_table_inner = re.sub(
            rb'(<hp:sz[^>]*?\sheight=")(\d+)(")',
            lambda m: m.group(1) + str(int(m.group(2)) + cloned_height).encode() + m.group(3),
            new_table_inner,
            count=1,
        )

    def transform(data: bytes) -> bytes:
        return data[: tbl_rng[0]] + new_table_inner + data[tbl_rng[1] :]

    dst = Path(args.output)
    delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    print(
        f"add_table_row: cloned last row of table {args.table}, "
        f"new rowAddr={new_row_addr}, byte delta {delta:+d}, output={dst}",
        file=sys.stderr,
    )

    if args.no_check:
        return 0
    return run_pitfall_check(
        dst, baseline=Path(args.baseline) if args.baseline else None
    )


if __name__ == "__main__":
    sys.exit(main())
