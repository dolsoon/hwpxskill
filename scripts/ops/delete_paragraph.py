#!/usr/bin/env python3
"""Delete one or more <hp:p> paragraphs from section0.xml.

Targets are specified by paragraph id or by exact text inside an <hp:t>
node. The matching paragraph (the smallest enclosing <hp:p>...</hp:p>) is
removed bit-exact; surrounding paragraphs are untouched.

Usage:
    # Delete by paragraph id
    python ops/delete_paragraph.py input.hwpx -o out.hwpx --id 12345

    # Delete the paragraph containing this text
    python ops/delete_paragraph.py input.hwpx -o out.hwpx \\
        --text "이 문장이 있는 단락을 제거"

    # Multiple deletions in one shot
    python ops/delete_paragraph.py input.hwpx -o out.hwpx \\
        --id 12345 --id 23456 --text "삭제할 단락"

    # See which paragraphs would be deleted without writing
    python ops/delete_paragraph.py input.hwpx \\
        --id 12345 --text "확인 대상" --dry-run

The script refuses to delete the FIRST paragraph (which carries the
required <hp:secPr>) — that would corrupt section structure. Use
replace_section to swap it instead.
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


def _find_paragraph_by_id(data: bytes, pid: str) -> tuple[int, int] | None:
    sig = f' id="{pid}"'.encode("utf-8")
    pos = data.find(sig)
    while pos >= 0:
        opens = list(HP_P_OPEN.finditer(data, 0, pos))
        if opens:
            p_open = opens[-1].start()
            first_gt = data.find(b">", p_open)
            if 0 <= first_gt and pos < first_gt:
                p_close = data.find(HP_P_CLOSE, first_gt)
                if p_close >= 0:
                    return (p_open, p_close + len(HP_P_CLOSE))
        pos = data.find(sig, pos + 1)
    return None


def _find_paragraph_by_text(data: bytes, needle: bytes) -> tuple[int, int] | None:
    pos = 0
    while True:
        i = data.find(needle, pos)
        if i < 0:
            return None
        last_t_open = data.rfind(b"<hp:t>", 0, i)
        last_t_close = data.rfind(b"</hp:t>", 0, i)
        if last_t_open > last_t_close:
            opens = list(HP_P_OPEN.finditer(data, 0, i))
            if opens:
                p_open = opens[-1].start()
                p_close = data.find(HP_P_CLOSE, i)
                if p_close >= 0:
                    return (p_open, p_close + len(HP_P_CLOSE))
        pos = i + 1


def _first_paragraph_open(data: bytes) -> int:
    m = HP_P_OPEN.search(data)
    return m.start() if m else -1


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx")
    p.add_argument("-o", "--output", help="Output .hwpx (omit with --dry-run)")
    p.add_argument(
        "--id",
        action="append",
        default=[],
        help="Repeatable: delete paragraph with this id.",
    )
    p.add_argument(
        "--text",
        action="append",
        default=[],
        help="Repeatable: delete the paragraph containing this text node.",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Resolve targets and print byte ranges; do not write output.",
    )
    p.add_argument("--baseline", help="Optional baseline for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    if not (args.id or args.text):
        print("ERROR: at least one --id or --text required.", file=sys.stderr)
        return 2

    src = Path(args.input)
    section = read_zip_entry(src, SECTION_ENTRY)
    first_p = _first_paragraph_open(section)

    targets: list[tuple[int, int, str]] = []
    for pid in args.id:
        r = _find_paragraph_by_id(section, pid)
        if r is None:
            print(f"WARN: id={pid} not found, skipping.", file=sys.stderr)
            continue
        targets.append((*r, f"id={pid}"))
    for txt in args.text:
        r = _find_paragraph_by_text(section, txt.encode("utf-8"))
        if r is None:
            print(f"WARN: text={txt!r} not found, skipping.", file=sys.stderr)
            continue
        targets.append((*r, f"text={txt!r}"))

    if not targets:
        print("ERROR: no resolvable targets.", file=sys.stderr)
        return 2

    # de-dup by start position (in case id and text picked the same)
    seen = set()
    unique: list[tuple[int, int, str]] = []
    for s, e, label in sorted(targets):
        if s in seen:
            continue
        seen.add(s)
        if s == first_p:
            print(
                f"ERROR: refusing to delete first paragraph (carries secPr): "
                f"{label}",
                file=sys.stderr,
            )
            return 2
        unique.append((s, e, label))

    if args.dry_run:
        for s, e, label in unique:
            print(f"would delete bytes [{s}:{e}] ({label}, length={e-s})")
        return 0

    if not args.output:
        print("ERROR: --output required (or use --dry-run).", file=sys.stderr)
        return 2

    # Build deletion ranges; apply from highest offset to lowest so earlier
    # offsets stay valid.
    sorted_ranges = sorted(unique, key=lambda t: t[0], reverse=True)

    def transform(data: bytes) -> bytes:
        cur = data
        for s, e, _ in sorted_ranges:
            cur = cur[:s] + cur[e:]
        return cur

    dst = Path(args.output)
    delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    print(
        f"delete_paragraph: removed {len(unique)} paragraph(s), "
        f"byte delta {delta:+d}, output={dst}",
        file=sys.stderr,
    )

    if args.no_check:
        return 0
    return run_pitfall_check(
        dst, baseline=Path(args.baseline) if args.baseline else None
    )


if __name__ == "__main__":
    sys.exit(main())
