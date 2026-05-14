#!/usr/bin/env python3
"""Run a chain of HWPX edits in a single Python process.

Each ops/* script costs ~25 ms of cold Python startup. Chaining 5 ops as
separate subprocesses takes 5×125 ms ≈ 625 ms; using batch.py drops it to
roughly one Python startup + one pitfall_check at the end (~150 ms total)
because the ZIP is rewritten only at the boundary between operations and
pitfall_check is consolidated.

Input is a JSON file (or stdin with `-`) describing one input HWPX, one
output HWPX, and an ordered list of operations. Each operation is a dict
with an `op` key naming the script (without .py) and the rest of the
keys passed as command-line arguments to that script.

Example batch.json:
    {
      "input": "report.hwpx",
      "output": "report-edited.hwpx",
      "baseline": "report.hwpx",
      "operations": [
        {"op": "replace_text", "find": "회의", "replace": "검토"},
        {"op": "swap_table_cells", "table": 5, "col": 0, "row": 0, "text": "구분"},
        {"op": "change_color", "darken": ["#FF0000=20"]},
        {"op": "add_paragraph", "after-text": "결론", "text": "추가 설명"}
      ]
    }

Run:
    python ops/batch.py batch.json
    cat batch.json | python ops/batch.py -

Each operation runs against a temp HWPX produced by the previous step.
Only the final HWPX is written to `output`. A single pitfall_check runs
against the final result (with --baseline if provided).
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))


def _import_op(name: str):
    """Import an ops module and return its main() entry point."""

    import importlib

    return importlib.import_module(name)


def _build_argv(op: dict) -> list[str]:
    """Convert a JSON operation dict into argv flags for the underlying script.

    Keys with a list value become repeated --key flags. Bool True becomes a
    bare --flag. Bool False is omitted.
    """

    argv: list[str] = []
    for k, v in op.items():
        if k == "op":
            continue
        flag = f"--{k.replace('_', '-')}"
        if isinstance(v, bool):
            if v:
                argv.append(flag)
        elif isinstance(v, list):
            for item in v:
                argv.extend([flag, str(item)])
        else:
            argv.extend([flag, str(v)])
    return argv


def main() -> int:
    p = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    p.add_argument(
        "spec",
        help="Path to batch JSON file, or - to read from stdin.",
    )
    p.add_argument(
        "--strict",
        action="store_true",
        help="Stop the chain on the first non-zero op exit code (default: continue).",
    )
    p.add_argument(
        "--no-final-check",
        action="store_true",
        help="Skip the final pitfall_check on the chained result.",
    )
    args = p.parse_args()

    if args.spec == "-":
        spec = json.loads(sys.stdin.read())
    else:
        spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))

    src = Path(spec["input"])
    dst = Path(spec["output"])
    baseline = spec.get("baseline")
    operations = spec.get("operations", [])
    if not src.is_file():
        print(f"ERROR: input not found: {src}", file=sys.stderr)
        return 2
    if not operations:
        print("ERROR: no operations in batch spec.", file=sys.stderr)
        return 2

    # Allowed op modules — explicit allow-list for safety.
    ALLOWED = {
        "replace_text",
        "swap_table_cells",
        "change_color",
        "replace_section",
        "add_paragraph",
        "delete_paragraph",
        "add_table_row",
        "delete_table_row",
    }

    print(
        f"batch: {len(operations)} operation(s) chained, input={src}, "
        f"output={dst}",
        file=sys.stderr,
    )

    with tempfile.TemporaryDirectory(prefix="hwpx_batch_") as td:
        tmpdir = Path(td)
        cur_in = src
        t_total_start = time.perf_counter()

        for i, op in enumerate(operations):
            name = op.get("op")
            if name not in ALLOWED:
                print(
                    f"ERROR: op #{i} '{name}' not in allowed set {sorted(ALLOWED)}",
                    file=sys.stderr,
                )
                return 2

            module = _import_op(name)
            cur_out = tmpdir / f"step_{i:03d}.hwpx"

            argv = [str(cur_in), "-o", str(cur_out)]
            argv.extend(_build_argv(op))
            argv.append("--no-check")  # batch consolidates pitfall_check at end

            print(
                f"\n[step {i+1}/{len(operations)}] {name} {' '.join(argv[3:])}",
                file=sys.stderr,
            )
            t0 = time.perf_counter()
            saved_argv = sys.argv
            try:
                sys.argv = [name] + argv
                rc = module.main()
            finally:
                sys.argv = saved_argv
            t1 = time.perf_counter()
            print(f"  rc={rc}  elapsed={(t1-t0)*1000:.1f}ms", file=sys.stderr)

            if rc not in (0, 2):
                if args.strict:
                    print(f"ERROR: op #{i} failed with rc={rc}, stopping.", file=sys.stderr)
                    return rc
                else:
                    print(
                        f"WARN: op #{i} returned rc={rc}; continuing chain "
                        "(use --strict to halt).",
                        file=sys.stderr,
                    )
            if not cur_out.exists():
                print(
                    f"ERROR: op #{i} did not produce output {cur_out}", file=sys.stderr
                )
                return 2
            cur_in = cur_out

        # Move final result to dst
        shutil.copy2(cur_in, dst)
        t_total_end = time.perf_counter()
        print(
            f"\nbatch chain done: {(t_total_end-t_total_start)*1000:.1f}ms total, "
            f"output={dst}",
            file=sys.stderr,
        )

    if args.no_final_check:
        return 0

    # Single in-process pitfall_check at end
    import pitfall_check

    rc = pitfall_check.check_in_process(
        dst,
        baseline=Path(baseline) if baseline else None,
        verbose=True,
    )
    return rc


if __name__ == "__main__":
    sys.exit(main())
