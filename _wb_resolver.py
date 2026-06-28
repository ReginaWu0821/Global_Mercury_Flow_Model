"""Resolve workbook paths across different workbook roots.

This module also implements the project's I/O policy:

* Inputs are read-only. The original workbooks under the input directory
  (``input_v2`` by default, or ``$GMF_INPUT_DIR``) are never modified.
* When a script asks for a workbook via :func:`resolve_workbook_path`, the
  input file is lazily copied into the output directory
  (``$GMF_OUTPUT_DIR``, default ``<repo>/output``) and the *output* path is
  returned. All subsequent reads/writes (including ``wb.save(file)``) target
  the copy, so the original input stays untouched.
"""

from __future__ import annotations

import os
import re
import shutil
from pathlib import Path


# Environment variable used to override the default workbook input directory.
# When set (typically by run_all_workbooks.py via a CLI argument), it takes
# precedence over the built-in defaults (input_v2, input, Code-2022, Global_Hg_Flow).
INPUT_DIR_ENV = "GMF_INPUT_DIR"

# Environment variable controlling where workbook outputs are written.
# When set, resolve_workbook_path() returns a path inside this directory
# (after copying the source workbook there once per run).
OUTPUT_DIR_ENV = "GMF_OUTPUT_DIR"


def _canonical_book_name(name: str) -> str:
    s = name.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-2022", "")
    return s.lower()


def _iter_workbook_candidates(root: Path):
    for p in sorted(root.glob("*.xlsx")):
        if ":Zone.Identifier" in p.name:
            continue
        yield p


def _candidate_roots(repo_root: Path) -> list[Path]:
    """Return ordered candidate roots for resolving workbook files.

    If the environment variable ``GMF_INPUT_DIR`` is set, that directory is
    tried first. Relative paths are resolved against ``repo_root``. The
    built-in defaults are always appended as a fallback so existing data
    layouts keep working.
    """

    roots: list[Path] = []
    env_dir = os.environ.get(INPUT_DIR_ENV)
    if env_dir:
        p = Path(env_dir).expanduser()
        if not p.is_absolute():
            p = (repo_root / p).resolve()
        roots.append(p)

    roots.append(repo_root / "input_v2")
    roots.append(repo_root / "input")
    roots.append(repo_root / "Code-2022")
    roots.append(repo_root / "Global_Hg_Flow")

    # Deduplicate while preserving order.
    seen: set[Path] = set()
    unique: list[Path] = []
    for r in roots:
        key = r.resolve() if r.exists() else r
        if key in seen:
            continue
        seen.add(key)
        unique.append(r)
    return unique


def resolve_workbook_path(repo_root: Path, declared_name: str) -> Path:
    """Return the workbook path that scripts should read from and write to.

    Resolution order for the *source* file:

    1) ``$GMF_INPUT_DIR`` (if set)
    2) ``<repo>/input_v2``
    3) ``<repo>/input``
    4) ``<repo>/Code-2022``
    5) ``<repo>/Global_Hg_Flow``

    Matching within each root:

    1) exact file name
    2) canonical name (collapse spaces, strip ``-2022``)

    If an output directory is configured via ``$GMF_OUTPUT_DIR``, the source
    file is copied there once (when the destination is missing or older than
    the source) and the destination path is returned. This guarantees that
    ``wb.save(file)`` calls inside individual scripts never touch the
    original input workbook.
    """

    roots = _candidate_roots(repo_root)
    declared_canon = _canonical_book_name(declared_name)

    source: Path | None = None
    for root in roots:
        if not root.exists():
            continue
        exact = root / declared_name
        if exact.exists():
            source = exact
            break
        for cand in _iter_workbook_candidates(root):
            if _canonical_book_name(cand.name) == declared_canon:
                source = cand
                break
        if source is not None:
            break

    if source is None:
        tried = ", ".join(str(r) for r in roots)
        raise FileNotFoundError(
            f"Cannot resolve workbook '{declared_name}' under: {tried}"
        )

    out_dir = _output_dir(repo_root)
    if out_dir is None:
        # No output redirection configured: legacy in-place behavior.
        return source

    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / source.name

    # Copy the source into the output directory the first time it is needed
    # (or when the source is newer than an existing copy). Subsequent calls
    # within the same run reuse the copy so progressive writes accumulate.
    try:
        need_copy = (
            not dest.exists()
            or source.stat().st_mtime > dest.stat().st_mtime
        )
    except OSError:
        need_copy = True

    if need_copy:
        # Avoid copying onto itself if input and output happen to coincide.
        try:
            same = dest.exists() and source.samefile(dest)
        except OSError:
            same = False
        if not same:
            shutil.copy2(source, dest)

    return dest


def _output_dir(repo_root: Path) -> Path | None:
    env_dir = os.environ.get(OUTPUT_DIR_ENV)
    if not env_dir:
        return None
    p = Path(env_dir).expanduser()
    if not p.is_absolute():
        p = (repo_root / p).resolve()
    return p
