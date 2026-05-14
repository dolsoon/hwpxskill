#!/usr/bin/env python3
"""Replace text inside a specific table cell of section0.xml.

Locates a cell by table index + (col, row) cellAddr, finds the first <hp:t>
node inside it, and replaces its inner text with the new value. The
surrounding cell structure (cellSz, cellSpan, borderFill, paragraph id,
linesegarray) is preserved bit-for-bit.

Usage:
    # Inspect tables first to find indices and cell addresses:
    python ops/swap_table_cells.py input.hwpx --list

    # Replace one cell:
    python ops/swap_table_cells.py input.hwpx -o out.hwpx \\
        --table 5 --col 1 --row 0 --text "새 헤더"

    # Multiple cells in one run:
    python ops/swap_table_cells.py input.hwpx -o out.hwpx \\
        --table 5 --cell 0,0=AAA --cell 1,0=BBB --cell 2,0=CCC

Limitations:
- Replaces the FIRST <hp:t> in the matching cell. Cells with multiple runs
  keep the trailing runs unchanged. Use --run N to target a different run.
- The cell must already contain at least one <hp:t> node (even empty).
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, read_zip_entry, run_pitfall_check  # noqa: E402

from lxml import etree  # noqa: E402

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}

SECTION_ENTRY = "Contents/section0.xml"


def list_tables(section_bytes: bytes) -> str:
    out: list[str] = []
    root = etree.fromstring(section_bytes)
    for ti, tbl in enumerate(root.xpath(".//hp:tbl", namespaces=NS)):
        rc = tbl.get("rowCnt")
        cc = tbl.get("colCnt")
        out.append(f"\nTable {ti}  id={tbl.get('id')}  rowCnt={rc} colCnt={cc}")
        for tr in tbl.xpath("./hp:tr", namespaces=NS):
            for tc in tr.xpath("./hp:tc", namespaces=NS):
                addr = tc.find("hp:cellAddr", namespaces=NS)
                if addr is None:
                    continue
                col = addr.get("colAddr")
                row = addr.get("rowAddr")
                # gather the first ~40 chars of cell text
                txt = "".join(tc.itertext()).strip()
                if len(txt) > 40:
                    txt = txt[:40] + "..."
                out.append(f"  ({col},{row})  {txt!r}")
    return "\n".join(out)


def find_cell_byte_range(
    section_bytes: bytes, table_idx: int, col: int, row: int
) -> tuple[int, int] | None:
    """Locate the byte range of the matching <hp:tc>...</hp:tc> in raw bytes.

    Strategy: find the cellAddr signature `colAddr="C" rowAddr="R"` that lives
    inside the Nth table, then walk backward to the enclosing <hp:tc and
    forward to the matching </hp:tc>.
    """

    # Find table opening tags to count which one is the Nth.
    table_re = re.compile(rb"<hp:tbl[\s>]")
    matches = list(table_re.finditer(section_bytes))
    if table_idx >= len(matches):
        return None
    tbl_start = matches[table_idx].start()
    if table_idx + 1 < len(matches):
        tbl_end = matches[table_idx + 1].start()
    else:
        tbl_end = len(section_bytes)
    region = section_bytes[tbl_start:tbl_end]

    # cellAddr signature inside this table only
    addr_sig = (
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"'.encode("utf-8")
    )
    addr_pos_local = region.find(addr_sig)
    if addr_pos_local < 0:
        return None
    addr_pos = tbl_start + addr_pos_local

    # walk backward to enclosing <hp:tc
    open_tag = b"<hp:tc"
    tc_open = section_bytes.rfind(open_tag, tbl_start, addr_pos)
    if tc_open < 0:
        return None

    # walk forward to closing </hp:tc> using nesting counter (in case of
    # nested objects, though hp:tc shouldn't nest under hp:tc)
    depth = 0
    open_re = re.compile(rb"<hp:tc[\s>]")
    close_tag = b"</hp:tc>"
    cursor = tc_open
    while True:
        next_open = open_re.search(section_bytes, cursor + 1, tbl_end)
        next_close = section_bytes.find(close_tag, cursor + 1, tbl_end)
        if next_close < 0:
            return None
        if next_open is not None and next_open.start() < next_close:
            depth += 1
            cursor = next_open.start()
            continue
        if depth == 0:
            tc_close = next_close + len(close_tag)
            return (tc_open, tc_close)
        depth -= 1
        cursor = next_close


def replace_first_t_in_cell(cell_bytes: bytes, new_text: str, run_idx: int) -> bytes:
    """Replace the inner text of the Nth <hp:t> inside a cell."""

    # Find <hp:t>...</hp:t> spans (tolerating <hp:t/> empty placeholder)
    occurrences: list[tuple[int, int]] = []
    cursor = 0
    while True:
        # First try the empty self-closing form
        empty = cell_bytes.find(b"<hp:t/>", cursor)
        full_open = cell_bytes.find(b"<hp:t>", cursor)
        # choose nearest
        candidates = [c for c in (empty, full_open) if c >= 0]
        if not candidates:
            break
        nearest = min(candidates)
        if nearest == empty and (full_open < 0 or empty < full_open):
            occurrences.append((empty, empty + len(b"<hp:t/>")))
            cursor = empty + len(b"<hp:t/>")
        else:
            close = cell_bytes.find(b"</hp:t>", full_open)
            if close < 0:
                break
            occurrences.append((full_open, close + len(b"</hp:t>")))
            cursor = close + len(b"</hp:t>")

    if run_idx >= len(occurrences):
        raise ValueError(
            f"Cell has only {len(occurrences)} <hp:t> node(s); cannot target run {run_idx}."
        )

    start, end = occurrences[run_idx]
    new_chunk = b"<hp:t>" + new_text.encode("utf-8") + b"</hp:t>"
    return cell_bytes[:start] + new_chunk + cell_bytes[end:]


def parse_cells_arg(values: list[str]) -> list[tuple[int, int, str]]:
    """Parse repeated --cell COL,ROW=text args."""

    out: list[tuple[int, int, str]] = []
    for raw in values:
        if "=" not in raw:
            raise ValueError(f"--cell expects COL,ROW=text, got: {raw!r}")
        coords, text = raw.split("=", 1)
        if "," not in coords:
            raise ValueError(f"--cell coordinates expect COL,ROW: {coords!r}")
        col_s, row_s = coords.split(",", 1)
        out.append((int(col_s), int(row_s), text))
    return out


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx file")
    p.add_argument("-o", "--output", help="Output .hwpx path")
    p.add_argument("--table", type=int, default=0, help="Table index (0-based)")
    p.add_argument("--col", type=int, help="Cell column address")
    p.add_argument("--row", type=int, help="Cell row address")
    p.add_argument("--text", help="Replacement text")
    p.add_argument(
        "--cell",
        action="append",
        default=[],
        help="Repeatable: COL,ROW=text (e.g., --cell 0,0=AAA --cell 1,0=BBB)",
    )
    p.add_argument(
        "--run",
        type=int,
        default=0,
        help="Which <hp:t> in the cell to replace (default 0 = first).",
    )
    p.add_argument(
        "--list",
        action="store_true",
        help="List all tables and their cells, do not patch.",
    )
    p.add_argument("--baseline", help="Optional baseline HWPX for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    src = Path(args.input)

    if args.list:
        section = read_zip_entry(src, SECTION_ENTRY)
        print(list_tables(section))
        return 0

    # Build edits list
    edits: list[tuple[int, int, str]] = []
    if args.col is not None and args.row is not None and args.text is not None:
        edits.append((args.col, args.row, args.text))
    try:
        edits.extend(parse_cells_arg(args.cell))
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    if not edits:
        print(
            "ERROR: nothing to do — pass --col/--row/--text or --cell COL,ROW=text "
            "(or --list to inspect).",
            file=sys.stderr,
        )
        return 2

    if not args.output:
        print("ERROR: --output is required when patching.", file=sys.stderr)
        return 2
    dst = Path(args.output)

    # Apply each edit by recomputing byte ranges from the freshly-edited bytes.
    def transform(data: bytes) -> bytes:
        cur = data
        for col, row, text in edits:
            rng = find_cell_byte_range(cur, args.table, col, row)
            if rng is None:
                raise ValueError(
                    f"Cell ({col},{row}) not found in table {args.table}"
                )
            start, end = rng
            cell = cur[start:end]
            new_cell = replace_first_t_in_cell(cell, text, args.run)
            cur = cur[:start] + new_cell + cur[end:]
        return cur

    try:
        delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    except (FileNotFoundError, KeyError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print(
        f"swap_table_cells: {len(edits)} edit(s), byte delta {delta:+d}, output={dst}",
        file=sys.stderr,
    )
    if args.no_check:
        return 0

    rc = run_pitfall_check(
        dst, baseline=Path(args.baseline) if args.baseline else None
    )
    if rc == 0:
        print("[pitfall_check] PASS", file=sys.stderr)
    elif rc == 2:
        print("[pitfall_check] WARNINGS (not blocking)", file=sys.stderr)
    else:
        print("[pitfall_check] FAIL — review output before opening.", file=sys.stderr)
    return rc


if __name__ == "__main__":
    sys.exit(main())
