#!/usr/bin/env python3
"""Replace exact text occurrences inside <hp:t> elements in section0.xml.

Pure raw-byte patch: the surrounding XML structure (paragraph ids, lineseg,
tables, etc.) is left bit-identical. Useful for typo fixes, value updates,
and small body edits where you do NOT want to touch any structure.

Limitations:
- Only matches text that lives entirely inside ONE <hp:t> node. Text spanning
  multiple runs / styles must be edited via swap_table_cells.py or by writing
  a custom patch.
- The match is byte-exact (UTF-8). XML entities (&amp; &lt; etc.) must be
  written in the same form as they appear in the file.
- All occurrences are replaced unless --first is set.

Usage:
    python ops/replace_text.py input.hwpx --find "구버전" --replace "신버전" \
        -o output.hwpx
    python ops/replace_text.py input.hwpx \
        --find "<hp:t>구</hp:t>" --replace "<hp:t>신</hp:t>" \
        --raw -o output.hwpx --first
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, run_pitfall_check  # noqa: E402


SECTION_ENTRY = "Contents/section0.xml"


def make_transform(
    find: bytes,
    replace: bytes,
    *,
    first_only: bool,
    raw: bool,
) -> tuple:
    """Return (transform_fn, lambda counting matches)."""

    found = {"count": 0}

    if raw:
        needle = find
        repl = replace
    else:
        # Wrap in <hp:t>...</hp:t> markers to constrain matches to actual text
        # nodes. This avoids touching attribute values or comments.
        needle = b"<hp:t>" + find + b"</hp:t>"
        repl = b"<hp:t>" + replace + b"</hp:t>"

    def fn(data: bytes) -> bytes:
        c = data.count(needle)
        found["count"] = c
        if c == 0:
            return data
        if first_only:
            return data.replace(needle, repl, 1)
        return data.replace(needle, repl)

    return fn, found


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx file")
    p.add_argument("--find", required=True, help="Text to find (UTF-8)")
    p.add_argument("--replace", required=True, help="Replacement text (UTF-8)")
    p.add_argument("-o", "--output", required=True, help="Output .hwpx path")
    p.add_argument(
        "--raw",
        action="store_true",
        help=(
            "Match the bytes verbatim instead of wrapping in <hp:t>...</hp:t>. "
            "Use for advanced patches; may corrupt structure if misused."
        ),
    )
    p.add_argument(
        "--first",
        action="store_true",
        help="Replace only the first occurrence (default: all).",
    )
    p.add_argument(
        "--baseline",
        help="Optional baseline HWPX for pitfall_check --baseline diff.",
    )
    p.add_argument(
        "--no-check",
        action="store_true",
        help="Skip the auto pitfall_check after patch (NOT recommended).",
    )
    args = p.parse_args()

    src = Path(args.input)
    dst = Path(args.output)
    if dst.exists() and dst.samefile(src):
        print("ERROR: --output must differ from input.", file=sys.stderr)
        return 2

    fn, found = make_transform(
        args.find.encode("utf-8"),
        args.replace.encode("utf-8"),
        first_only=args.first,
        raw=args.raw,
    )

    try:
        delta = patch_zip_entry(src, dst, SECTION_ENTRY, fn)
    except (FileNotFoundError, KeyError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    n = found["count"]
    if n == 0:
        print(f"WARN: 0 occurrences of {args.find!r} found.", file=sys.stderr)
        # Still emit the file (it's an exact copy at this point) so the user
        # can chain commands deterministically.
    print(
        f"replace_text: {n} match(es), byte delta {delta:+d}, "
        f"output={dst}",
        file=sys.stderr,
    )

    if args.no_check:
        return 0

    rc = run_pitfall_check(
        dst,
        baseline=Path(args.baseline) if args.baseline else None,
    )
    if rc == 0:
        print("[pitfall_check] PASS", file=sys.stderr)
    elif rc == 2:
        print("[pitfall_check] WARNINGS (not blocking)", file=sys.stderr)
    else:
        print(
            "[pitfall_check] FAIL — output may be unsafe for Mac 한글.",
            file=sys.stderr,
        )
    return rc


if __name__ == "__main__":
    sys.exit(main())
