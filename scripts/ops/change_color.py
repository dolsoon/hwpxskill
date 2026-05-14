#!/usr/bin/env python3
"""Bulk hex-color replacement across header.xml and/or section0.xml.

Two ways to specify the replacement:
1. Direct hex pairs:  --map "#FROM=#TO"
2. HSL transforms:    --darken/--lighten/--saturate/--shift-hue "#FROM=AMOUNT"
   The script converts the FROM color to HSL, applies the transform, and
   uses the resulting hex as the TO value. Useful when the user describes
   the change qualitatively ("진하게", "톤다운", "보색"). Multiple transform
   pairs can be combined with each other and with --map.

Examples:
    # Direct hex swap
    python ops/change_color.py input.hwpx -o out.hwpx --map "#7B8B3D=#1F4E79"

    # Make a red 20% darker (HSL Lightness -20pp)
    python ops/change_color.py input.hwpx -o out.hwpx --darken "#FF0000=20"

    # Lighten one color, desaturate another
    python ops/change_color.py input.hwpx -o out.hwpx \\
        --lighten "#000000=15" --saturate "#FF0000=-50"

    # Rotate hue by 180° (complementary color)
    python ops/change_color.py input.hwpx -o out.hwpx --shift-hue "#1F4E79=180"

    # Header only, restricted to textColor attribute
    python ops/change_color.py input.hwpx -o out.hwpx \\
        --scope header --attr textColor --map "#000000=#1A1A1A"

    # List all unique colors currently in the document
    python ops/change_color.py input.hwpx --list

HSL units:
- --lighten / --darken / --saturate : percentage points added to L or S
  (clamped to [0, 100]). Negative --saturate values reduce saturation.
- --shift-hue : degrees added to H (mod 360).

Notes:
- Hex match is case-insensitive by default. Disable with --case-sensitive.
- The replacement preserves the leading `#` and emits exactly 6 uppercase hex.
- If --attr is set, only attributes whose name matches will be patched
  (`attr="#XXX"` form). Without --attr, every `#RRGGBB` occurrence is matched.
"""

from __future__ import annotations

import argparse
import colorsys
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


# ---------------------------------------------------------------------------
# HSL transforms — convert FROM hex via HSL, return new hex
# ---------------------------------------------------------------------------


def _hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{r:02X}{g:02X}{b:02X}"


def _clamp01(x: float) -> float:
    return 0.0 if x < 0.0 else (1.0 if x > 1.0 else x)


def _apply_hsl(hex_in: str, *, dh: float = 0.0, dl: float = 0.0, ds: float = 0.0) -> str:
    """Apply HSL deltas to hex_in and return new hex.

    dh: degrees to add to hue (mod 360)
    dl: percentage points to add to lightness (clamped 0..100)
    ds: percentage points to add to saturation (clamped 0..100)
    """

    r, g, b = _hex_to_rgb(hex_in)
    h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
    # h in [0, 1] (fractions of 360°); l/s in [0, 1]
    h = ((h * 360.0 + dh) % 360.0) / 360.0
    l = _clamp01(l + dl / 100.0)
    s = _clamp01(s + ds / 100.0)
    nr, ng, nb = colorsys.hls_to_rgb(h, l, s)
    return _rgb_to_hex(round(nr * 255), round(ng * 255), round(nb * 255))


def _parse_hex_amount(values: list[str], op_name: str) -> list[tuple[str, float]]:
    out: list[tuple[str, float]] = []
    for raw in values:
        if "=" not in raw:
            raise ValueError(f"--{op_name} expects #HEX=AMOUNT, got: {raw!r}")
        h, amt = raw.split("=", 1)
        h = h.strip()
        if not (h.startswith("#") and len(h) == 7 and HEX_RE.fullmatch(h.encode("ascii"))):
            raise ValueError(f"Invalid hex in --{op_name}: {h!r}")
        try:
            f = float(amt.strip())
        except ValueError as e:
            raise ValueError(f"Invalid AMOUNT in --{op_name}: {amt!r}") from e
        out.append((h.upper(), f))
    return out


def expand_hsl_transforms(
    *,
    lighten: list[str],
    darken: list[str],
    saturate: list[str],
    shift_hue: list[str],
) -> tuple[list[tuple[str, str]], list[str]]:
    """Convert all HSL transforms into concrete (FROM, TO) hex pairs.

    Returns (mappings, log_lines). Multiple transforms targeting the same
    FROM hex are composed in argv order (lighten -> darken -> saturate ->
    shift-hue).
    """

    accumulated: dict[str, tuple[float, float, float]] = {}  # hex -> (dh, dl, ds)

    for h, amt in _parse_hex_amount(lighten, "lighten"):
        dh, dl, ds = accumulated.get(h, (0.0, 0.0, 0.0))
        accumulated[h] = (dh, dl + amt, ds)
    for h, amt in _parse_hex_amount(darken, "darken"):
        dh, dl, ds = accumulated.get(h, (0.0, 0.0, 0.0))
        accumulated[h] = (dh, dl - amt, ds)
    for h, amt in _parse_hex_amount(saturate, "saturate"):
        dh, dl, ds = accumulated.get(h, (0.0, 0.0, 0.0))
        accumulated[h] = (dh, dl, ds + amt)
    for h, amt in _parse_hex_amount(shift_hue, "shift-hue"):
        dh, dl, ds = accumulated.get(h, (0.0, 0.0, 0.0))
        accumulated[h] = (dh + amt, dl, ds)

    mappings: list[tuple[str, str]] = []
    log: list[str] = []
    for h, (dh, dl, ds) in accumulated.items():
        new = _apply_hsl(h, dh=dh, dl=dl, ds=ds)
        mappings.append((h, new))
        log.append(
            f"  {h} -> {new}  (dh={dh:+.0f}°, dl={dl:+.0f}pp, ds={ds:+.0f}pp)"
        )
    return mappings, log


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
        "--lighten",
        action="append",
        default=[],
        help='Repeatable: "#HEX=N" → HSL Lightness +N pp (clamped 0..100).',
    )
    p.add_argument(
        "--darken",
        action="append",
        default=[],
        help='Repeatable: "#HEX=N" → HSL Lightness -N pp (clamped 0..100).',
    )
    p.add_argument(
        "--saturate",
        action="append",
        default=[],
        help='Repeatable: "#HEX=N" → HSL Saturation +N pp (negative reduces).',
    )
    p.add_argument(
        "--shift-hue",
        action="append",
        default=[],
        help='Repeatable: "#HEX=N" → HSL Hue +N degrees (mod 360).',
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

    has_hsl = any([args.lighten, args.darken, args.saturate, args.shift_hue])
    if not args.map and not has_hsl:
        print(
            "ERROR: provide at least one of --map / --lighten / --darken / "
            "--saturate / --shift-hue.",
            file=sys.stderr,
        )
        return 2

    try:
        mappings = parse_map(args.map)
        hsl_mappings, hsl_log = expand_hsl_transforms(
            lighten=args.lighten,
            darken=args.darken,
            saturate=args.saturate,
            shift_hue=args.shift_hue,
        )
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    if hsl_mappings:
        print("HSL transforms resolved:", file=sys.stderr)
        for line in hsl_log:
            print(line, file=sys.stderr)
        mappings.extend(hsl_mappings)

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
