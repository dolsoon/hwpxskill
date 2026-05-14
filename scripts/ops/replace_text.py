#!/usr/bin/env python3
"""Replace text occurrences inside <hp:t> elements in section0.xml.

Pure raw-byte patch: the surrounding XML structure (paragraph ids, lineseg,
tables, etc.) is left bit-identical. Attribute values and tag names are NEVER
touched — only the character data inside <hp:t>...</hp:t> nodes.

Three matching modes:
- DEFAULT (text-node substring): find every <hp:t>...</hp:t> node and replace
  occurrences of `--find` inside the node's text. Safe against accidentally
  hitting attribute values, ids, etc. Handles cases like <hp:t>회의 결과</hp:t>
  where you want to swap "회의" for "검토".
- --whole-node: only match when --find equals the ENTIRE text of an <hp:t>
  node (the previous default). Use when you need exact-equality semantics.
- --raw: byte-exact match anywhere in section0.xml. Power-user mode; can
  corrupt XML if misused.

Other behavior:
- The match is UTF-8 byte-exact within the chosen mode. XML entities
  (&amp; &lt; etc.) must be written in the same form as they appear in the
  file.
- All occurrences are replaced unless --first is set.
- Exits with code 1 if `--require-match` is set and zero matches are found.

Usage:
    # default substring inside text nodes
    python ops/replace_text.py input.hwpx --find "회의" --replace "검토" \
        -o output.hwpx

    # exact <hp:t> equality
    python ops/replace_text.py input.hwpx --find "총괄" --replace "주관" \
        --whole-node -o output.hwpx

    # raw byte patch (advanced)
    python ops/replace_text.py input.hwpx \
        --find '<hp:t>구</hp:t>' --replace '<hp:t>신</hp:t>' \
        --raw -o output.hwpx --first
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, run_pitfall_check  # noqa: E402


SECTION_ENTRY = "Contents/section0.xml"


_HP_T_RE = __import__("re").compile(rb"<hp:t>([^<]*)</hp:t>")


def make_transform(
    find: bytes,
    replace: bytes,
    *,
    first_only: bool,
    mode: str,
) -> tuple:
    """Return (transform_fn, found-counter dict).

    mode is one of: "substring" (default), "whole-node", "raw".
    """

    found = {"count": 0}

    if mode == "raw":
        needle = find
        repl = replace

        def fn_raw(data: bytes) -> bytes:
            c = data.count(needle)
            found["count"] = c
            if c == 0:
                return data
            if first_only:
                return data.replace(needle, repl, 1)
            return data.replace(needle, repl)

        return fn_raw, found

    if mode == "whole-node":
        needle = b"<hp:t>" + find + b"</hp:t>"
        repl = b"<hp:t>" + replace + b"</hp:t>"

        def fn_whole(data: bytes) -> bytes:
            c = data.count(needle)
            found["count"] = c
            if c == 0:
                return data
            if first_only:
                return data.replace(needle, repl, 1)
            return data.replace(needle, repl)

        return fn_whole, found

    # substring mode (default): rewrite each <hp:t> node text individually
    def fn_sub(data: bytes) -> bytes:
        total = 0

        def repl_one(m):
            nonlocal total
            inner = m.group(1)
            if find not in inner:
                return m.group(0)
            n = inner.count(find)
            if first_only and total + n > 1:
                # only allow the first replacement total
                if total >= 1:
                    return m.group(0)
                # consume just one occurrence
                new_inner = inner.replace(find, replace, 1)
                total += 1
            else:
                new_inner = inner.replace(find, replace)
                total += n
            return b"<hp:t>" + new_inner + b"</hp:t>"

        new_data = _HP_T_RE.sub(repl_one, data)
        found["count"] = total
        return new_data

    return fn_sub, found


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx file")
    p.add_argument("--find", required=True, help="Text to find (UTF-8)")
    p.add_argument("--replace", required=True, help="Replacement text (UTF-8)")
    p.add_argument("-o", "--output", required=True, help="Output .hwpx path")
    p.add_argument(
        "--whole-node",
        action="store_true",
        help=(
            "Match only when --find equals the entire text of an <hp:t> node "
            "(safer; previous default)."
        ),
    )
    p.add_argument(
        "--raw",
        action="store_true",
        help=(
            "Match the bytes verbatim anywhere in section0.xml. Power-user; "
            "can corrupt XML if misused."
        ),
    )
    p.add_argument(
        "--first",
        action="store_true",
        help="Replace only the first occurrence (default: all).",
    )
    p.add_argument(
        "--require-match",
        action="store_true",
        help="Exit with code 1 if zero matches were found (default: 0).",
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

    if args.raw and args.whole_node:
        print("ERROR: --raw and --whole-node are mutually exclusive.", file=sys.stderr)
        return 2
    mode = "raw" if args.raw else ("whole-node" if args.whole_node else "substring")

    fn, found = make_transform(
        args.find.encode("utf-8"),
        args.replace.encode("utf-8"),
        first_only=args.first,
        mode=mode,
    )

    try:
        delta = patch_zip_entry(src, dst, SECTION_ENTRY, fn)
    except (FileNotFoundError, KeyError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    n = found["count"]
    if n == 0:
        msg = (
            f"WARN: 0 occurrences of {args.find!r} found "
            f"(mode={mode}). Output is a byte-identical copy of input."
        )
        print(msg, file=sys.stderr)
        if args.require_match:
            return 1
    print(
        f"replace_text: {n} match(es) [mode={mode}], byte delta {delta:+d}, "
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
