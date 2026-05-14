#!/usr/bin/env python3
"""Insert a new <hp:p> paragraph at a marker position in section0.xml.

The new paragraph can be either:
- a simple text body (script wraps it into a minimal <hp:p>...</hp:p> with
  the chosen paraPrIDRef / charPrIDRef and a placeholder linesegarray), or
- a fully-formed <hp:p>...</hp:p> XML literal (use --xml or --xml-file).

The insertion position is anchored to a marker paragraph found by either
a text needle inside an <hp:t> or a paragraph id. By default the new
paragraph is inserted AFTER the marker paragraph; use --before to insert
before.

Usage:
    # Add a one-line body paragraph after the paragraph that contains "결론"
    python ops/add_paragraph.py input.hwpx -o out.hwpx \\
        --after-text "결론" --text "추가 분석 결과는 다음 절에서 다룬다."

    # Insert a fully-formed XML block before paragraph id 12345
    python ops/add_paragraph.py input.hwpx -o out.hwpx \\
        --before-id 12345 --xml-file new_paragraph.xml

    # Use a custom paraPr / charPr style for the simple-text path
    python ops/add_paragraph.py input.hwpx -o out.hwpx \\
        --after-text "서론" --text "본 연구는..." \\
        --para-pr 24 --char-pr 0 --new-id 9999900100

Each new paragraph gets `--new-id` (default 9999900001 + auto-increment
when the id collides). To avoid colliding with future inserts, pass an
explicit --new-id from a high range (e.g., 5000000000+).
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, read_zip_entry, run_pitfall_check  # noqa: E402

SECTION_ENTRY = "Contents/section0.xml"
HP_P_OPEN = re.compile(rb"<hp:p[\s>]")
HP_P_CLOSE = b"</hp:p>"


def _find_paragraph_by_text(data: bytes, needle: bytes) -> tuple[int, int] | None:
    """Locate <hp:p>...</hp:p> containing `needle` inside an <hp:t> node."""

    pos = 0
    while True:
        i = data.find(needle, pos)
        if i < 0:
            return None
        # confirm it's inside text (between <hp:t>...</hp:t>)
        last_t_open = data.rfind(b"<hp:t>", 0, i)
        last_t_close = data.rfind(b"</hp:t>", 0, i)
        if last_t_open > last_t_close:
            # we're inside a text node; now find the enclosing <hp:p>
            opens = list(HP_P_OPEN.finditer(data, 0, i))
            if not opens:
                pos = i + 1
                continue
            p_open = opens[-1].start()
            p_close = data.find(HP_P_CLOSE, i)
            if p_close < 0:
                return None
            return (p_open, p_close + len(HP_P_CLOSE))
        pos = i + 1


def _find_paragraph_by_id(data: bytes, pid: str) -> tuple[int, int] | None:
    sig = f' id="{pid}"'.encode("utf-8")
    pos = data.find(sig)
    while pos >= 0:
        # ensure this is on an <hp:p ...> tag
        opens = list(HP_P_OPEN.finditer(data, 0, pos))
        if not opens:
            pos = data.find(sig, pos + 1)
            continue
        p_open = opens[-1].start()
        # require there's no <hp:p in between p_open and pos
        # i.e., pos must lie within the <hp:p ...> opening tag
        first_gt = data.find(b">", p_open)
        if first_gt < 0:
            return None
        if pos < first_gt:
            p_close = data.find(HP_P_CLOSE, first_gt)
            if p_close < 0:
                return None
            return (p_open, p_close + len(HP_P_CLOSE))
        pos = data.find(sig, pos + 1)
    return None


def _next_unique_id(data: bytes, start: int = 9_999_900_001) -> int:
    """Find the smallest id >= start not currently used as <hp:p id="...">."""

    used = set(re.findall(rb'<hp:p[^>]*?\sid="(\d+)"', data))
    used_ints = {int(x) for x in used}
    cur = start
    while cur in used_ints:
        cur += 1
    return cur


def _build_simple_paragraph(
    text: str, *, pid: int, para_pr: int, char_pr: int, horz: int = 42520
) -> bytes:
    return (
        f'<hp:p id="{pid}" paraPrIDRef="{para_pr}" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_pr}"><hp:t>{text}</hp:t></hp:run>'
        f"<hp:linesegarray>"
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1000" '
        f'textheight="1000" baseline="850" spacing="600" horzpos="0" '
        f'horzsize="{horz}" flags="393216"/>'
        f"</hp:linesegarray>"
        f"</hp:p>"
    ).encode("utf-8")


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx")
    p.add_argument("-o", "--output", required=True, help="Output .hwpx")

    anchor = p.add_mutually_exclusive_group(required=True)
    anchor.add_argument("--after-text", help="Anchor: text inside an <hp:t>")
    anchor.add_argument("--before-text", help="Anchor: text inside an <hp:t>")
    anchor.add_argument("--after-id", help="Anchor: paragraph id")
    anchor.add_argument("--before-id", help="Anchor: paragraph id")

    body = p.add_mutually_exclusive_group(required=True)
    body.add_argument("--text", help="Simple text body for a one-paragraph insert")
    body.add_argument("--xml", help="Inline XML literal of <hp:p>...</hp:p>")
    body.add_argument("--xml-file", help="Read XML literal from this file")

    p.add_argument(
        "--para-pr", type=int, default=0, help="paraPrIDRef for --text mode (default 0)"
    )
    p.add_argument(
        "--char-pr", type=int, default=0, help="charPrIDRef for --text mode (default 0)"
    )
    p.add_argument(
        "--horzsize",
        type=int,
        default=42520,
        help="lineseg horzsize for --text mode (default A4 body width 42520).",
    )
    p.add_argument(
        "--new-id",
        type=int,
        help="Explicit id for the inserted paragraph (else auto-allocated).",
    )
    p.add_argument("--baseline", help="Optional baseline for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    src = Path(args.input)
    dst = Path(args.output)

    section = read_zip_entry(src, SECTION_ENTRY)
    if args.after_text or args.before_text:
        marker = (args.after_text or args.before_text).encode("utf-8")
        rng = _find_paragraph_by_text(section, marker)
    else:
        pid = args.after_id or args.before_id
        rng = _find_paragraph_by_id(section, pid)

    if rng is None:
        print("ERROR: anchor paragraph not found.", file=sys.stderr)
        return 2

    insert_after = bool(args.after_text or args.after_id)

    # Resolve replacement bytes
    if args.text is not None:
        new_id = args.new_id or _next_unique_id(section)
        new_block = _build_simple_paragraph(
            args.text,
            pid=new_id,
            para_pr=args.para_pr,
            char_pr=args.char_pr,
            horz=args.horzsize,
        )
    elif args.xml_file:
        new_block = Path(args.xml_file).read_bytes()
    else:
        new_block = (args.xml or "").encode("utf-8")

    def transform(data: bytes) -> bytes:
        # re-resolve range on the fresh bytes
        if args.after_text or args.before_text:
            marker = (args.after_text or args.before_text).encode("utf-8")
            r = _find_paragraph_by_text(data, marker)
        else:
            pid = args.after_id or args.before_id
            r = _find_paragraph_by_id(data, pid)
        if r is None:
            raise ValueError("anchor not found in transform pass")
        if insert_after:
            return data[: r[1]] + new_block + data[r[1] :]
        return data[: r[0]] + new_block + data[r[0] :]

    try:
        delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    except (FileNotFoundError, KeyError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print(
        f"add_paragraph: inserted {len(new_block)} bytes "
        f"({'after' if insert_after else 'before'} anchor), "
        f"byte delta {delta:+d}, output={dst}",
        file=sys.stderr,
    )

    if args.no_check:
        return 0
    rc = run_pitfall_check(
        dst, baseline=Path(args.baseline) if args.baseline else None
    )
    return rc


if __name__ == "__main__":
    sys.exit(main())
