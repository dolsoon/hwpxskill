"""Shared helpers for raw-byte HWPX edits.

These helpers preserve ZIP entry order, compression, and timestamps. Only the
target XML byte payload is replaced. lxml is intentionally NOT used to
re-serialize trees (see hwpx-pitfalls.md pitfall 7).
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Callable
from zipfile import ZipFile, ZipInfo

SCRIPT_DIR = Path(__file__).resolve().parent.parent  # .../scripts
PITFALL_CHECK = SCRIPT_DIR / "pitfall_check.py"

# Import pitfall_check as a module so ops can call it in-process and avoid
# ~70 ms of Python cold start + lxml import per invocation. Falls back to
# subprocess silently if the module fails to import (e.g., name conflict).
sys.path.insert(0, str(SCRIPT_DIR))
try:
    import pitfall_check as _pitfall_mod  # noqa: E402
except Exception:  # pragma: no cover - defensive
    _pitfall_mod = None


def patch_zip_entry(
    src: Path,
    dst: Path,
    entry: str,
    transform: Callable[[bytes], bytes],
) -> int:
    """Copy `src` to `dst`, applying `transform(bytes)->bytes` to one entry.

    Returns the number of byte-level changes (len(after) - len(before)).
    Preserves entry order, compression mode, and timestamps for ALL entries.
    """

    if not src.is_file():
        raise FileNotFoundError(src)

    delta = 0
    matched = False
    with ZipFile(src, "r") as zin:
        infos = zin.infolist()
        # Mac 한글 호환을 위해 mimetype은 ZIP_STORED 첫 엔트리로 유지.
        with ZipFile(dst, "w") as zout:
            for info in infos:
                data = zin.read(info.filename)
                if info.filename == entry:
                    new_data = transform(data)
                    matched = True
                    delta = len(new_data) - len(data)
                    data = new_data
                # Preserve stored vs deflated mode per entry
                new_info = ZipInfo(filename=info.filename, date_time=info.date_time)
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                new_info.create_system = info.create_system
                zout.writestr(new_info, data)

    if not matched:
        raise KeyError(f"Entry not found in {src}: {entry}")
    return delta


def read_zip_entry(src: Path, entry: str) -> bytes:
    with ZipFile(src, "r") as zf:
        return zf.read(entry)


def run_pitfall_check(
    hwpx: Path,
    baseline: Path | None = None,
    strict: bool = False,
    *,
    in_process: bool = True,
) -> int:
    """Run pitfall_check; return its exit code.

    Default `in_process=True` calls the imported module directly (saves
    ~70 ms vs subprocess). Set False to force subprocess (fallback or
    isolation). Falls back automatically if module import failed.
    """

    if in_process and _pitfall_mod is not None:
        return _pitfall_mod.check_in_process(
            hwpx, baseline=baseline, strict=strict, verbose=True
        )

    cmd = [sys.executable, str(PITFALL_CHECK), str(hwpx)]
    if baseline is not None:
        cmd.extend(["--baseline", str(baseline)])
    if strict:
        cmd.append("--strict")
    print(f"\n[pitfall_check] {' '.join(cmd)}", file=sys.stderr)
    proc = subprocess.run(cmd)
    return proc.returncode


def safe_overwrite(src: Path, dst: Path) -> None:
    """Move src to dst atomically; if dst exists keep it as .bak."""

    if dst.exists():
        bak = dst.with_suffix(dst.suffix + ".bak")
        shutil.copy2(dst, bak)
    shutil.move(str(src), str(dst))
