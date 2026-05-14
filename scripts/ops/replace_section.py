#!/usr/bin/env python3
"""Replace a contiguous block of section0.xml between two markers with new XML.

Useful for swapping a whole chapter, table, or paragraph block in one shot
without unpacking/repacking the whole HWPX.

The markers can be:
- Plain text that appears inside an <hp:t>...</hp:t> (the script wraps them
  automatically unless --raw is used).
- A raw substring (use --raw) — useful for matching paragraph ids like
  `id="3138635038"`.

The block boundary is the byte range from the START of the paragraph that
contains the start-marker to the END of the paragraph that contains the
end-marker (i.e., aligned to <hp:p ...>...</hp:p> boundaries) so you don't
accidentally split a paragraph mid-tag.

Usage:
    # Find boundaries first
    python ops/replace_section.py input.hwpx --probe \\
        --start "Ⅲ. 분석 결과" --end "Ⅳ. 결론"

    # Replace with new XML loaded from a file
    python ops/replace_section.py input.hwpx -o out.hwpx \\
        --start "Ⅲ. 분석 결과" --end "Ⅳ. 결론" \\
        --xml-file new_chapter3.xml

    # Inline replacement (multiline allowed)
    python ops/replace_section.py input.hwpx -o out.hwpx \\
        --start "Ⅲ. 분석 결과" --end "Ⅳ. 결론" \\
        --xml '<hp:p ... ></hp:p><hp:p ...></hp:p>'

The replacement XML must consist of complete <hp:p>...</hp:p> elements with
unique ids and proper linesegarray (run pitfall_check after).
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


def find_paragraph_containing(data: bytes, marker: bytes) -> tuple[int, int] | None:
    """Return the byte range of the <hp:p>...</hp:p> that contains marker."""

    pos = data.find(marker)
    if pos < 0:
        return None
    # walk back to last <hp:p
    open_m = None
    for m in HP_P_OPEN.finditer(data, 0, pos):
        open_m = m
    if open_m is None:
        return None
    # find matching </hp:p> after pos. Account for nested <hp:p> (uncommon
    # but possible in tables).
    depth = 0
    cursor = open_m.start()
    while True:
        next_open = HP_P_OPEN.search(data, cursor + 1)
        next_close = data.find(HP_P_CLOSE, cursor + 1)
        if next_close < 0:
            return None
        if next_open is not None and next_open.start() < next_close:
            depth += 1
            cursor = next_open.start()
            continue
        if depth == 0:
            return (open_m.start(), next_close + len(HP_P_CLOSE))
        depth -= 1
        cursor = next_close


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx")
    p.add_argument("-o", "--output", help="Output .hwpx")
    p.add_argument("--start", required=True, help="Start marker text")
    p.add_argument("--end", required=True, help="End marker text")
    p.add_argument(
        "--raw",
        action="store_true",
        help="Match markers as raw bytes (default: wrap in <hp:t>...</hp:t>).",
    )
    p.add_argument("--xml", help="Replacement XML inline.")
    p.add_argument("--xml-file", help="Replacement XML loaded from file path.")
    p.add_argument(
        "--probe",
        action="store_true",
        help="Print resolved byte range and snippet, do not patch.",
    )
    p.add_argument("--baseline", help="Optional baseline HWPX for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    p.add_argument(
        "--include-end-paragraph",
        action="store_true",
        help=(
            "Include the paragraph that contains the END marker in the "
            "replacement (default: stop just before it)."
        ),
    )
    args = p.parse_args()

    src = Path(args.input)
    section = read_zip_entry(src, SECTION_ENTRY)

    start_marker = args.start.encode("utf-8")
    end_marker = args.end.encode("utf-8")
    if not args.raw:
        start_marker = b"<hp:t>" + start_marker + b"</hp:t>"
        # END marker is matched as a substring too. For headings we usually
        # want the marker that opens the NEXT section, so wrapping helps.
        end_marker = b"<hp:t>" + end_marker + b"</hp:t>"

    start_p = find_paragraph_containing(section, start_marker)
    end_p = find_paragraph_containing(section, end_marker)
    if start_p is None:
        print(f"ERROR: start marker not found: {args.start!r}", file=sys.stderr)
        return 2
    if end_p is None:
        print(f"ERROR: end marker not found: {args.end!r}", file=sys.stderr)
        return 2

    if args.include_end_paragraph:
        block_start, block_end = start_p[0], end_p[1]
    else:
        block_start, block_end = start_p[0], end_p[0]

    if block_end < block_start:
        print(
            f"ERROR: end marker appears before start marker "
            f"(start_byte={block_start}, end_byte={block_end}).",
            file=sys.stderr,
        )
        return 2

    if args.probe:
        snippet = section[block_start : min(block_end, block_start + 400)]
        print(f"start_byte={block_start} end_byte={block_end} length={block_end-block_start}")
        print(f"first_400_bytes:\n{snippet.decode('utf-8', errors='replace')}")
        return 0

    if not args.output:
        print("ERROR: --output required when patching.", file=sys.stderr)
        return 2
    if not (args.xml or args.xml_file):
        print("ERROR: --xml or --xml-file required when patching.", file=sys.stderr)
        return 2

    if args.xml_file:
        replacement = Path(args.xml_file).read_bytes()
    else:
        replacement = (args.xml or "").encode("utf-8")

    def transform(data: bytes) -> bytes:
        # Resolve again on the freshly-passed bytes (should be identical to
        # `section` here since we patched no other entry).
        sp = find_paragraph_containing(data, start_marker)
        ep = find_paragraph_containing(data, end_marker)
        if sp is None or ep is None:
            raise ValueError("markers no longer resolvable in transform pass")
        if args.include_end_paragraph:
            return data[: sp[0]] + replacement + data[ep[1] :]
        return data[: sp[0]] + replacement + data[ep[0] :]

    dst = Path(args.output)
    try:
        delta = patch_zip_entry(src, dst, SECTION_ENTRY, transform)
    except (FileNotFoundError, KeyError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print(
        f"replace_section: replaced {block_end-block_start} bytes with "
        f"{len(replacement)} bytes (delta {delta:+d}), output={dst}",
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
