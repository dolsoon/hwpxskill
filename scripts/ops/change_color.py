#!/usr/bin/env python3
"""Bulk hex-color replacement across header.xml and/or section0.xml.

Replaces every occurrence of `#RRGGBB` (case-insensitive) with the new color.
Supports multiple --map FROM=TO pairs in one run, plus optional named-attribute
scoping so you only touch e.g. textColor or borderFill background.

Usage:
    # Swap a single color in both header & section
    python ops/change_color.py input.hwpx -o out.hwpx \\
        --map "#7B8B3D=#1F4E79"

    # Swap multiple, header only, in attribute textColor only
    python ops/change_color.py input.hwpx -o out.hwpx \\
        --scope header --attr textColor \\
        --map "#000000=#1A1A1A" --map "#FF0000=#C00000"

    # List all unique colors currently in the document
    python ops/change_color.py input.hwpx --list

Notes:
- Hex match is case-insensitive but PRESERVES the case of the FROM literal so
  you can target only the lower or only the upper case form if you want.
  Disable with --case-sensitive.
- The replacement preserves the leading `#` and emits exactly 6 hex chars.
- If --attr is set, only attributes whose name matches will be patched
  (`attr="#XXX"` form). Without --attr, every `#RRGGBB` occurrence is matched.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _zip_patch import patch_zip_entry, read_zip_entry, run_pitfall_check  # noqa: E402

HEADER_ENTRY = "Contents/header.xml"
SECTION_ENTRY = "Contents/section0.xml"

HEX_RE = re.compile(rb'#[0-9A-Fa-f]{6}')


def list_colors(*payloads: bytes) -> dict[str, int]:
    counter: dict[str, int] = {}
    for data in payloads:
        for m in HEX_RE.finditer(data):
            key = m.group(0).decode("ascii").upper()
            counter[key] = counter.get(key, 0) + 1
    return dict(sorted(counter.items(), key=lambda kv: -kv[1]))


def parse_map(values: list[str]) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for raw in values:
        if "=" not in raw:
            raise ValueError(f"--map expects FROM=TO, got: {raw!r}")
        a, b = raw.split("=", 1)
        a = a.strip()
        b = b.strip()
        if not (a.startswith("#") and len(a) == 7 and HEX_RE.fullmatch(a.encode("ascii"))):
            raise ValueError(f"Invalid hex color in --map FROM: {a!r}")
        if not (b.startswith("#") and len(b) == 7 and HEX_RE.fullmatch(b.encode("ascii"))):
            raise ValueError(f"Invalid hex color in --map TO: {b!r}")
        out.append((a, b))
    return out


def make_transform(
    mappings: list[tuple[str, str]],
    *,
    attr: str | None,
    case_sensitive: bool,
) -> tuple:
    counts = {fr: 0 for fr, _ in mappings}

    def fn(data: bytes) -> bytes:
        cur = data
        for fr, to in mappings:
            fr_b = fr.encode("ascii")
            to_b = to.encode("ascii")
            if attr:
                # Match only attr="#HEX" form (XML uses double quotes everywhere
                # in OWPML output).
                pattern = (
                    re.escape(attr.encode("ascii"))
                    + b'="'
                    + (
                        re.escape(fr_b)
                        if case_sensitive
                        else b"".join(
                            (
                                b"[" + bytes([c, c ^ 0x20]) + b"]"
                                if chr(c).isalpha()
                                else bytes([c]).replace(b".", rb"\.")
                            )
                            for c in fr_b
                        )
                    )
                    + b'"'
                )
                regex = re.compile(pattern)
                replacement = attr.encode("ascii") + b'="' + to_b + b'"'
                new_cur, n = regex.subn(replacement, cur)
            else:
                if case_sensitive:
                    n = cur.count(fr_b)
                    new_cur = cur.replace(fr_b, to_b)
                else:
                    regex = re.compile(re.escape(fr_b), re.IGNORECASE)
                    new_cur, n = regex.subn(to_b, cur)
            counts[fr] += n
            cur = new_cur
        return cur

    return fn, counts


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument("input", help="Input .hwpx file")
    p.add_argument("-o", "--output", help="Output .hwpx path")
    p.add_argument(
        "--map",
        action="append",
        default=[],
        help="Repeatable: FROM=TO (#RRGGBB=#RRGGBB).",
    )
    p.add_argument(
        "--scope",
        choices=("header", "section", "both"),
        default="both",
        help="Which XML entry to patch (default: both).",
    )
    p.add_argument(
        "--attr",
        help=(
            "Restrict matches to attribute=\"#HEX\" form (e.g., --attr textColor)."
        ),
    )
    p.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Disable case-insensitive hex matching.",
    )
    p.add_argument("--list", action="store_true", help="List unique colors and exit.")
    p.add_argument("--baseline", help="Optional baseline HWPX for pitfall_check.")
    p.add_argument("--no-check", action="store_true", help="Skip pitfall_check.")
    args = p.parse_args()

    src = Path(args.input)

    if args.list:
        h = read_zip_entry(src, HEADER_ENTRY) if args.scope in ("header", "both") else b""
        s = read_zip_entry(src, SECTION_ENTRY) if args.scope in ("section", "both") else b""
        colors = list_colors(h, s)
        print(f"Unique colors in {args.scope}:")
        for k, v in colors.items():
            print(f"  {k}  x{v}")
        return 0

    if not args.output:
        print("ERROR: --output required when patching.", file=sys.stderr)
        return 2
    if not args.map:
        print("ERROR: at least one --map FROM=TO required.", file=sys.stderr)
        return 2

    try:
        mappings = parse_map(args.map)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    fn, counts = make_transform(
        mappings, attr=args.attr, case_sensitive=args.case_sensitive
    )

    dst = Path(args.output)

    # Apply per scope. We patch header first, then section, copying through.
    intermediate = dst.with_suffix(dst.suffix + ".tmp")
    try:
        if args.scope in ("header", "both"):
            patch_zip_entry(src, intermediate, HEADER_ENTRY, fn)
            inp = intermediate
        else:
            inp = src

        if args.scope in ("section", "both"):
            patch_zip_entry(inp, dst, SECTION_ENTRY, fn)
            if intermediate.exists() and intermediate != dst:
                intermediate.unlink()
        else:
            # only header was patched → move intermediate to dst
            intermediate.replace(dst)
    except (FileNotFoundError, KeyError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    total = sum(counts.values())
    print(
        f"change_color: {total} replacement(s) across mappings={counts}, "
        f"output={dst}",
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
