#!/usr/bin/env python3
"""HWPX OWPML pitfall checker (Mac 한글 무한루프 방지 게이트).

`validate.py` and `page_guard.py` only check structural integrity and page
drift. They do NOT catch the OWPML violations that cause Mac Hancom Office
to refuse the document or spin at 100% CPU. This script implements the seven
pitfalls documented in `references/hwpx-pitfalls.md` and exits non-zero when
any are detected.

Pitfalls checked:
  1. Duplicate <hp:p> ids
  2. Missing <hp:linesegarray> in non-empty paragraphs
  3. rowSpan/colSpan occupied positions containing extra empty <hp:tc>
  4. JUSTIFY body paragraph + single placeholder lineseg + text > 30 chars
     (causes letter-spacing blow-up)
  5. cellSz width sum != table sz width (per row)
  6. Undefined IDRef (charPrIDRef / paraPrIDRef / borderFillIDRef) in section
  7. lxml-serialization advisory (cannot auto-detect; emitted as note)

Usage:
    python pitfall_check.py document.hwpx
    python pitfall_check.py document.hwpx --json
    python pitfall_check.py document.hwpx --baseline reference.hwpx
    python pitfall_check.py document.hwpx --strict

Exit codes:
  0 = clean (or only baseline-equal warnings)
  1 = pitfall detected (default)
  2 = warnings only and --strict not set
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from zipfile import BadZipFile, ZipFile

from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
}

# 한 줄 추정 글자 수 (A4 본문폭 horzsize=37420 기준 보수치).
# 60자/줄로 잡으면 양쪽정렬 본문에서 lineseg 부족 진단이 안전 측에 위치.
CHARS_PER_LINE_DEFAULT = 60

# 본문(표 외부) 단락에서 함정 4 진단 시 무시할 텍스트 길이 임계.
JUSTIFY_TEXT_MIN = 30

# linesegarray 누락 단락이 이 값을 넘으면 한글이 첫 열기에서 무한루프.
LINESEG_MISSING_THRESHOLD = 80


@dataclass
class PitfallReport:
    """One pitfall finding."""

    code: str
    severity: str  # "error" | "warn" | "note"
    message: str
    detail: dict[str, Any] = field(default_factory=dict)


def _read_xml(zf: ZipFile, entry: str) -> etree._Element | None:
    if entry not in zf.namelist():
        return None
    return etree.fromstring(zf.read(entry))


def _is_in_table_cell(p: etree._Element) -> bool:
    """Return True if the paragraph lives inside an <hp:tc> (table cell)."""

    a = p.getparent()
    while a is not None:
        tag = etree.QName(a.tag).localname
        if tag == "tc":
            return True
        a = a.getparent()
    return False


def _text_length(p: etree._Element) -> int:
    return sum(
        len("".join(t.itertext())) for t in p.xpath(".//hp:t", namespaces=NS)
    )


# ---------------------------------------------------------------------------
# Pitfall 1: duplicate hp:p ids
# ---------------------------------------------------------------------------


def check_duplicate_ids(section: etree._Element) -> list[PitfallReport]:
    out: list[PitfallReport] = []
    paragraphs = section.xpath(".//hp:p", namespaces=NS)
    ids = [p.get("id") for p in paragraphs if p.get("id") is not None]
    counter = Counter(ids)
    dups = {i: c for i, c in counter.items() if c > 1}
    if dups:
        # placeholder ids that hwpx tooling tends to leave behind
        suspicious = {"0", "2147483648"}
        suspicious_hits = {i: c for i, c in dups.items() if i in suspicious}
        out.append(
            PitfallReport(
                code="P1_DUPLICATE_PARA_ID",
                severity="error",
                message=(
                    f"<hp:p> id duplicated across {len(dups)} unique value(s); "
                    f"{sum(dups.values())} paragraphs share a non-unique id."
                ),
                detail={
                    "total_paragraphs": len(paragraphs),
                    "unique_ids": len(set(ids)),
                    "duplicate_groups": dict(list(dups.items())[:20]),
                    "suspicious_placeholder_ids": suspicious_hits,
                },
            )
        )
    return out


# ---------------------------------------------------------------------------
# Pitfall 2: missing hp:linesegarray
# ---------------------------------------------------------------------------


def check_missing_lineseg(section: etree._Element) -> list[PitfallReport]:
    out: list[PitfallReport] = []
    paragraphs = section.xpath(".//hp:p", namespaces=NS)
    missing: list[str] = []
    for p in paragraphs:
        if _text_length(p) == 0:
            continue
        if p.find("hp:linesegarray", namespaces=NS) is None:
            missing.append(p.get("id") or "?")

    n = len(missing)
    if n == 0:
        return out

    if n >= LINESEG_MISSING_THRESHOLD:
        sev = "error"
        msg = (
            f"{n} non-empty paragraphs missing <hp:linesegarray> "
            f"(>= {LINESEG_MISSING_THRESHOLD} threshold for Mac 한글 freeze)."
        )
    else:
        sev = "warn"
        msg = (
            f"{n} non-empty paragraphs missing <hp:linesegarray> "
            f"(below {LINESEG_MISSING_THRESHOLD} threshold; safe but advisory)."
        )

    out.append(
        PitfallReport(
            code="P2_MISSING_LINESEG",
            severity=sev,
            message=msg,
            detail={
                "missing_paragraph_ids_sample": missing[:30],
                "missing_count": n,
                "threshold": LINESEG_MISSING_THRESHOLD,
            },
        )
    )
    return out


# ---------------------------------------------------------------------------
# Pitfall 3: rowSpan/colSpan occupied positions with extra empty cells
# ---------------------------------------------------------------------------


def check_span_occupation(section: etree._Element) -> list[PitfallReport]:
    out: list[PitfallReport] = []
    violations: list[dict[str, Any]] = []

    for ti, tbl in enumerate(section.xpath(".//hp:tbl", namespaces=NS)):
        occupied: dict[tuple[int, int], tuple[int, int]] = {}
        for tr in tbl.xpath("./hp:tr", namespaces=NS):
            for tc in tr.xpath("./hp:tc", namespaces=NS):
                addr = tc.find("hp:cellAddr", namespaces=NS)
                cs = tc.find("hp:cellSpan", namespaces=NS)
                if addr is None or cs is None:
                    continue
                col = int(addr.get("colAddr", 0))
                row = int(addr.get("rowAddr", 0))
                rspan = int(cs.get("rowSpan", 1))
                cspan = int(cs.get("colSpan", 1))
                if (col, row) in occupied:
                    src = occupied[(col, row)]
                    violations.append(
                        {
                            "table_index": ti,
                            "table_id": tbl.get("id"),
                            "duplicate_cell": {"col": col, "row": row},
                            "occupier_cell": {"col": src[0], "row": src[1]},
                        }
                    )
                for dr in range(rspan):
                    for dc in range(cspan):
                        if (dr, dc) == (0, 0):
                            continue
                        # Mark the position as occupied but DO NOT overwrite if
                        # already marked (so we attribute to the first owner).
                        occupied.setdefault((col + dc, row + dr), (col, row))
                # Also mark the originating cell so duplicate tc on same
                # (col,row) is detected.
                occupied.setdefault((col, row), (col, row))

    if violations:
        out.append(
            PitfallReport(
                code="P3_SPAN_OCCUPATION",
                severity="error",
                message=(
                    f"{len(violations)} cell position(s) occupied by "
                    "rowSpan/colSpan but contain an additional <hp:tc>."
                ),
                detail={"violations_sample": violations[:20]},
            )
        )
    return out


# ---------------------------------------------------------------------------
# Pitfall 4: JUSTIFY body paragraph + single placeholder lineseg + long text
# ---------------------------------------------------------------------------


def _justify_paraprop_ids(header: etree._Element | None) -> set[str]:
    out: set[str] = set()
    if header is None:
        return out
    for ppr in header.xpath(".//hh:paraPr", namespaces=NS):
        al = ppr.find("hh:align", namespaces=NS)
        if al is not None and al.get("horizontal") == "JUSTIFY":
            ppr_id = ppr.get("id")
            if ppr_id is not None:
                out.add(ppr_id)
    return out


def check_justify_lineseg_blowup(
    section: etree._Element,
    header: etree._Element | None,
    chars_per_line: int = CHARS_PER_LINE_DEFAULT,
) -> list[PitfallReport]:
    out: list[PitfallReport] = []
    justify_ids = _justify_paraprop_ids(header)
    if not justify_ids:
        # cannot evaluate without header context — emit advisory note
        return [
            PitfallReport(
                code="P4_JUSTIFY_LINESEG_BLOWUP",
                severity="note",
                message=(
                    "Header paraPr information unavailable; cannot detect "
                    "JUSTIFY letter-spacing blow-up."
                ),
            )
        ]

    violations: list[dict[str, Any]] = []
    for p in section.xpath(".//hp:p", namespaces=NS):
        if _is_in_table_cell(p):
            continue
        ppr_id = p.get("paraPrIDRef")
        if ppr_id not in justify_ids:
            continue
        t_len = _text_length(p)
        if t_len < JUSTIFY_TEXT_MIN:
            continue
        lsa = p.find("hp:linesegarray", namespaces=NS)
        if lsa is None:
            # Already removed → 한글 재계산 → safe.
            continue
        n_lseg = len(lsa.findall("hp:lineseg", namespaces=NS))
        expected = max(1, math.ceil(t_len / chars_per_line))
        if n_lseg < expected:
            violations.append(
                {
                    "paragraph_id": p.get("id"),
                    "text_length": t_len,
                    "lineseg_count": n_lseg,
                    "expected_min": expected,
                }
            )

    if violations:
        out.append(
            PitfallReport(
                code="P4_JUSTIFY_LINESEG_BLOWUP",
                severity="error",
                message=(
                    f"{len(violations)} JUSTIFY body paragraph(s) have fewer "
                    "lineseg than text length implies — letter-spacing will "
                    "blow up in 한글."
                ),
                detail={
                    "chars_per_line_assumed": chars_per_line,
                    "violations_sample": violations[:20],
                },
            )
        )
    return out


# ---------------------------------------------------------------------------
# Pitfall 5: cellSz width sum != table width
# ---------------------------------------------------------------------------


def check_cellsz_width_sum(section: etree._Element) -> list[PitfallReport]:
    out: list[PitfallReport] = []
    violations: list[dict[str, Any]] = []

    for ti, tbl in enumerate(section.xpath(".//hp:tbl", namespaces=NS)):
        sz = tbl.find("hp:sz", namespaces=NS)
        if sz is None:
            continue
        try:
            tbl_w = int(sz.get("width", 0))
        except (TypeError, ValueError):
            continue
        for ri, tr in enumerate(tbl.xpath("./hp:tr", namespaces=NS)):
            row_sum = 0
            for tc in tr.xpath("./hp:tc", namespaces=NS):
                cellSz = tc.find("hp:cellSz", namespaces=NS)
                cs = tc.find("hp:cellSpan", namespaces=NS)
                if cellSz is None or cs is None:
                    continue
                try:
                    w = int(cellSz.get("width", 0))
                except (TypeError, ValueError):
                    continue
                # Only count cells that occupy column 0..N once.
                # rowSpan does not affect width sum; colSpan does (sum of spanned cols).
                # Since cellSz already gives the spanned width, just add it.
                row_sum += w
            if row_sum != tbl_w:
                violations.append(
                    {
                        "table_index": ti,
                        "table_id": tbl.get("id"),
                        "row_index": ri,
                        "row_width_sum": row_sum,
                        "table_width": tbl_w,
                        "diff": row_sum - tbl_w,
                    }
                )

    if violations:
        out.append(
            PitfallReport(
                code="P5_CELLSZ_WIDTH_MISMATCH",
                severity="error",
                message=(
                    f"{len(violations)} table row(s) where cellSz width sum "
                    "does not match <hp:tbl> width."
                ),
                detail={"violations_sample": violations[:20]},
            )
        )
    return out


# ---------------------------------------------------------------------------
# Pitfall 6: undefined IDRef in section pointing nowhere in header
# ---------------------------------------------------------------------------


def _collect_header_ids(header: etree._Element | None) -> dict[str, set[str]]:
    """Return {kind -> set of defined ids} from header.xml."""

    out = {
        "charPr": set(),
        "paraPr": set(),
        "borderFill": set(),
        "style": set(),
    }
    if header is None:
        return out
    for el in header.xpath(".//hh:charPr", namespaces=NS):
        out["charPr"].add(el.get("id"))
    for el in header.xpath(".//hh:paraPr", namespaces=NS):
        out["paraPr"].add(el.get("id"))
    for el in header.xpath(".//hh:borderFill", namespaces=NS):
        out["borderFill"].add(el.get("id"))
    for el in header.xpath(".//hh:style", namespaces=NS):
        out["style"].add(el.get("id"))
    return out


def check_undefined_idref(
    section: etree._Element, header: etree._Element | None
) -> list[PitfallReport]:
    if header is None:
        return [
            PitfallReport(
                code="P6_UNDEFINED_IDREF",
                severity="note",
                message="header.xml unavailable; skipping IDRef cross-check.",
            )
        ]

    defined = _collect_header_ids(header)
    bad: list[dict[str, Any]] = []

    pairs = [
        ("charPrIDRef", "charPr"),
        ("paraPrIDRef", "paraPr"),
        ("borderFillIDRef", "borderFill"),
        ("styleIDRef", "style"),
    ]

    for el in section.iter():
        for attr, kind in pairs:
            v = el.get(attr)
            if v is None:
                continue
            if v not in defined[kind]:
                bad.append(
                    {
                        "element": etree.QName(el.tag).localname,
                        "attribute": attr,
                        "value": v,
                        "defined_count": len(defined[kind]),
                    }
                )

    if bad:
        return [
            PitfallReport(
                code="P6_UNDEFINED_IDREF",
                severity="error",
                message=(
                    f"{len(bad)} IDRef value(s) reference ids not defined in "
                    "header.xml."
                ),
                detail={"violations_sample": bad[:20]},
            )
        ]
    return []


# ---------------------------------------------------------------------------
# Pitfall 7: lxml serialization advisory
# ---------------------------------------------------------------------------


def check_lxml_advisory(zf: ZipFile) -> list[PitfallReport]:
    """We cannot deterministically detect lxml-mangled XML, but we can flag a
    common smell: section0.xml lacking the canonical first line that the
    skeleton template emits. This is an *advisory*, never an error."""

    if "Contents/section0.xml" not in zf.namelist():
        return []
    raw = zf.read("Contents/section0.xml")
    head = raw[:200]
    note = None
    if not head.startswith(b"<?xml "):
        note = "section0.xml missing XML declaration (lxml may have stripped it)."
    elif b' standalone=' not in head and b' standalone="' not in head:
        # Hancom skeleton always emits standalone="yes"; absence is a hint.
        note = (
            "section0.xml XML declaration lacks standalone attribute "
            "(may indicate lxml re-serialization)."
        )
    if note is None:
        return []
    return [
        PitfallReport(
            code="P7_LXML_SERIALIZATION_ADVISORY",
            severity="note",
            message=note,
        )
    ]


# ---------------------------------------------------------------------------
# Top-level runner
# ---------------------------------------------------------------------------


def run_checks(
    hwpx_path: Path, chars_per_line: int = CHARS_PER_LINE_DEFAULT
) -> list[PitfallReport]:
    if not hwpx_path.is_file():
        return [
            PitfallReport(
                code="E_FILE",
                severity="error",
                message=f"File not found: {hwpx_path}",
            )
        ]
    try:
        zf = ZipFile(hwpx_path, "r")
    except BadZipFile:
        return [
            PitfallReport(
                code="E_ZIP",
                severity="error",
                message=f"Not a valid ZIP archive: {hwpx_path}",
            )
        ]

    reports: list[PitfallReport] = []
    with zf:
        section = _read_xml(zf, "Contents/section0.xml")
        header = _read_xml(zf, "Contents/header.xml")
        if section is None:
            return [
                PitfallReport(
                    code="E_SECTION",
                    severity="error",
                    message="Contents/section0.xml not found in HWPX.",
                )
            ]

        reports.extend(check_duplicate_ids(section))
        reports.extend(check_missing_lineseg(section))
        reports.extend(check_span_occupation(section))
        reports.extend(check_justify_lineseg_blowup(section, header, chars_per_line))
        reports.extend(check_cellsz_width_sum(section))
        reports.extend(check_undefined_idref(section, header))
        reports.extend(check_lxml_advisory(zf))

    return reports


def diff_against_baseline(
    current: list[PitfallReport], baseline: list[PitfallReport]
) -> list[PitfallReport]:
    """Return only reports whose (code, identity-set signature) is not present
    in baseline.

    Identity-set signature picks ONLY the location/identity of each violation
    (paragraph id, table+row, duplicate-id values, etc.). Volatile noise like
    text length or lineseg count is intentionally excluded so that re-running
    the same edit twice yields the same signature, while a NEW violation
    location DOES create a fresh signature.
    """

    def sig(r: PitfallReport) -> tuple:
        d = r.detail or {}

        # P1: identity = the SET of duplicate id values
        if r.code == "P1_DUPLICATE_PARA_ID":
            ids = tuple(sorted((d.get("duplicate_groups") or {}).keys()))
            return (r.code, ("dup_ids", ids))

        # P2: identity = the SET of paragraph ids missing lineseg
        if r.code == "P2_MISSING_LINESEG":
            ids = tuple(sorted(d.get("missing_paragraph_ids_sample") or []))
            return (r.code, ("missing_ids_sample", ids))

        # P3: identity = (table_id, duplicate_cell, occupier_cell) tuples
        if r.code == "P3_SPAN_OCCUPATION":
            rows = tuple(
                sorted(
                    (
                        v.get("table_id"),
                        tuple(sorted((v.get("duplicate_cell") or {}).items())),
                        tuple(sorted((v.get("occupier_cell") or {}).items())),
                    )
                    for v in (d.get("violations_sample") or [])
                )
            )
            return (r.code, ("span_violations", rows))

        # P4: identity = the SET of paragraph ids with blow-up risk
        if r.code == "P4_JUSTIFY_LINESEG_BLOWUP":
            ids = tuple(
                sorted(
                    v.get("paragraph_id")
                    for v in (d.get("violations_sample") or [])
                )
            )
            return (r.code, ("blowup_pids", ids))

        # P5: identity = (table_id, row_index) tuples
        if r.code == "P5_CELLSZ_WIDTH_MISMATCH":
            rows = tuple(
                sorted(
                    (v.get("table_id"), v.get("row_index"))
                    for v in (d.get("violations_sample") or [])
                )
            )
            return (r.code, ("cellsz_rows", rows))

        # P6: identity = (element, attribute, value) tuples
        if r.code == "P6_UNDEFINED_IDREF":
            rows = tuple(
                sorted(
                    (
                        v.get("element"),
                        v.get("attribute"),
                        v.get("value"),
                    )
                    for v in (d.get("violations_sample") or [])
                )
            )
            return (r.code, ("idref_rows", rows))

        # P7 and notes: code alone (advisory; baseline-equal if both emit it)
        return (r.code, ())

    base_sigs = {sig(r) for r in baseline}
    return [r for r in current if sig(r) not in base_sigs]


def report_to_dict(r: PitfallReport) -> dict[str, Any]:
    return {
        "code": r.code,
        "severity": r.severity,
        "message": r.message,
        "detail": r.detail,
    }


def check_in_process(
    hwpx_path: Path | str,
    *,
    baseline: Path | str | None = None,
    chars_per_line: int = CHARS_PER_LINE_DEFAULT,
    strict: bool = False,
    verbose: bool = True,
    out: Any = None,
) -> int:
    """Run the full pitfall check in-process (no subprocess) and return an
    exit-code-compatible int.

    Designed for ops/ scripts that want to skip the ~70 ms of cold Python
    startup + lxml import incurred by spawning pitfall_check as a subprocess.
    Output is printed to `out` (default sys.stderr) only when verbose=True.
    """

    out = out if out is not None else sys.stderr
    cur = run_checks(Path(hwpx_path), chars_per_line=chars_per_line)
    if baseline is not None:
        base = run_checks(Path(baseline), chars_per_line=chars_per_line)
        cur = diff_against_baseline(cur, base)

    errors = [r for r in cur if r.severity == "error"]
    warns = [r for r in cur if r.severity == "warn"]
    notes = [r for r in cur if r.severity == "note"]

    if verbose:
        if not cur:
            print(f"PITFALL_CHECK PASS: {hwpx_path}", file=out)
        else:
            print(
                f"PITFALL_CHECK: {hwpx_path} "
                f"({len(errors)} error, {len(warns)} warn, {len(notes)} note)",
                file=out,
            )
            for r in errors:
                print(f"  [ERROR] {r.code}: {r.message}", file=out)
            for r in warns:
                print(f"  [WARN]  {r.code}: {r.message}", file=out)
            for r in notes:
                print(f"  [NOTE]  {r.code}: {r.message}", file=out)

    if errors:
        return 1
    if warns and strict:
        return 1
    if warns:
        return 2
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="HWPX OWPML pitfall checker (Mac 한글 무한루프 방지 게이트).",
    )
    parser.add_argument("input", help="Path to .hwpx file")
    parser.add_argument(
        "--baseline",
        help=(
            "Optional reference HWPX whose pitfalls are pre-existing and "
            "should be ignored. Only NEW pitfalls relative to baseline cause "
            "non-zero exit."
        ),
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit structured JSON instead of human-readable text.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat warnings (missing-lineseg below threshold, etc.) as errors.",
    )
    parser.add_argument(
        "--chars-per-line",
        type=int,
        default=CHARS_PER_LINE_DEFAULT,
        help=(
            "Estimated characters per line for JUSTIFY blow-up detection "
            f"(default {CHARS_PER_LINE_DEFAULT})."
        ),
    )
    args = parser.parse_args()

    cur = run_checks(Path(args.input), chars_per_line=args.chars_per_line)
    if args.baseline:
        base = run_checks(Path(args.baseline), chars_per_line=args.chars_per_line)
        cur = diff_against_baseline(cur, base)

    errors = [r for r in cur if r.severity == "error"]
    warns = [r for r in cur if r.severity == "warn"]
    notes = [r for r in cur if r.severity == "note"]

    if args.json:
        payload = {
            "input": args.input,
            "baseline": args.baseline,
            "errors": [report_to_dict(r) for r in errors],
            "warnings": [report_to_dict(r) for r in warns],
            "notes": [report_to_dict(r) for r in notes],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        if not cur:
            print(f"PITFALL_CHECK PASS: {args.input}")
            print("  All 7 OWPML pitfalls clear.")
        else:
            print(
                f"PITFALL_CHECK: {args.input} "
                f"({len(errors)} error, {len(warns)} warn, {len(notes)} note)"
            )
            for r in errors:
                print(f"  [ERROR] {r.code}: {r.message}")
                _print_detail(r.detail, indent=4)
            for r in warns:
                print(f"  [WARN]  {r.code}: {r.message}")
                _print_detail(r.detail, indent=4)
            for r in notes:
                print(f"  [NOTE]  {r.code}: {r.message}")

    if errors:
        return 1
    if warns and args.strict:
        return 1
    if warns:
        return 2
    return 0


def _print_detail(detail: dict[str, Any], indent: int = 4) -> None:
    if not detail:
        return
    pad = " " * indent
    for k, v in detail.items():
        if isinstance(v, (list, dict)) and len(str(v)) > 120:
            head = json.dumps(v, ensure_ascii=False)[:200]
            print(f"{pad}{k}: {head} ...")
        else:
            print(f"{pad}{k}: {v}")


if __name__ == "__main__":
    sys.exit(main())
