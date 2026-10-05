"""Time and peak memory of python-calamine-plus's reading modes, side by side.

Native streaming (plain and ``rich=True``) and ``compatibility="openpyxl"`` with its
flags, each reading every row of every sheet of one workbook, plus a cost breakdown of
the ``rich=True``-equivalent compatibility mode (one kind of work added at a time).

Each scenario runs in its own process, so peak memory is not shared. Scenarios that
"read" an attribute access it on every cell, so producing it is included. (cProfile is
not used: the work happens in Rust, inside the row iterator and attribute getters,
which it cannot see.)

usage: python benchmarks/profile_modes.py BOOK.xlsx [-o OUT.txt] [--repeat N] [--only KEY ...]
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import platform
import subprocess
import sys
import time
from typing import Callable

# ---- peak memory -------------------------------------------------------------------


def peak_mb() -> float:
    """Peak resident memory of this process, in MB."""
    if sys.platform == "win32":
        import ctypes
        import ctypes.wintypes as wt

        class PMC(ctypes.Structure):
            _fields_ = [
                ("cb", wt.DWORD),
                ("PageFaultCount", wt.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        k32 = ctypes.WinDLL("kernel32")
        k32.GetCurrentProcess.restype = wt.HANDLE
        k32.K32GetProcessMemoryInfo.argtypes = [
            wt.HANDLE,
            ctypes.POINTER(PMC),
            wt.DWORD,
        ]
        pmc = PMC()
        pmc.cb = ctypes.sizeof(PMC)
        if not k32.K32GetProcessMemoryInfo(
            k32.GetCurrentProcess(), ctypes.byref(pmc), pmc.cb
        ):
            raise OSError("GetProcessMemoryInfo failed")
        return pmc.PeakWorkingSetSize / 2**20
    import resource

    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return rss / 2**20 if sys.platform == "darwin" else rss / 2**10


# ---- scenarios ---------------------------------------------------------------------


def native(path: str, rich: bool, read: bool) -> tuple[int, int]:
    from python_calamine_plus import CalamineWorkbook

    wb = CalamineWorkbook.from_path(path)
    rows = cells = 0
    for name in wb.sheet_names:
        for row in wb.stream_sheet_by_name(name, rich=rich):
            rows += 1
            cells += len(row)
            if not read:
                continue
            for c in row:
                c.value, c.formula, c.comment, c.hyperlink, c.merged_range
                style = c.style
                if style is not None:
                    style["font"]["bold"], style["fill"], style["border"], style[
                        "number_format"
                    ]
    return rows, cells


def compat(path: str, reads: frozenset[str], **options: object) -> tuple[int, int]:
    """Reads every cell; ``reads`` picks what is accessed on each (``values`` means
    ``values_only=True`` rows)."""
    from python_calamine_plus import load_workbook
    from python_calamine_plus._python_calamine import EmptyCell, MergedCell

    values_only = "values" in reads
    value, styles = "value" in reads, "styles" in reads
    formula, meta = "formula" in reads, "meta" in reads
    read_only = bool(options.get("read_only"))
    wb = load_workbook(path, compatibility="openpyxl", **options)  # type: ignore[call-overload]
    rows = cells = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=values_only):
            rows += 1
            cells += len(row)
            if values_only:
                continue
            for c in row:
                if type(c) is EmptyCell:
                    continue
                if value:
                    c.value
                if styles:
                    c.font.b, c.fill.fgColor, c.border.left, c.number_format
                if formula:
                    c.formula, c.cached_value
                if meta:
                    if read_only:
                        c.comment, c.hyperlink, c.merged_range
                    elif type(c) is not MergedCell:
                        c.comment, c.hyperlink
    return rows, cells


RO = dict(read_only=True)
META = dict(read_comments=True, read_hyperlinks=True, read_merged_cells=True)
F = frozenset
EVERYTHING = F({"value", "styles", "formula", "meta"})

# (key, group, description, function)
Scenario = tuple[str, str, str, Callable[[str], tuple[int, int]]]
SCENARIOS: list[Scenario] = [
    (
        "native",
        "native",
        "stream_sheet_by_name(): value rows",
        lambda p: native(p, False, False),
    ),
    (
        "native_rich_iter",
        "native",
        "stream_sheet_by_name(rich=True): rows of CalamineCell, nothing read",
        lambda p: native(p, True, False),
    ),
    (
        "native_rich",
        "native",
        "rich=True, every cell: value, formula, style, comment, hyperlink, merged range",
        lambda p: native(p, True, True),
    ),
    (
        "ro_values",
        "compat read_only=True",
        "iter_rows(values_only=True), formula view (data_only=False)",
        lambda p: compat(p, F({"values"}), **RO),
    ),
    (
        "ro_values_data",
        "compat read_only=True",
        "iter_rows(values_only=True), data_only=True",
        lambda p: compat(p, F({"values"}), data_only=True, **RO),
    ),
    (
        "ro_cells",
        "compat read_only=True",
        "iter_rows(): ReadOnlyCell objects, value read",
        lambda p: compat(p, F({"value"}), **RO),
    ),
    (
        "ro_everything",
        "compat read_only=True",
        "formula_and_value + read_comments/hyperlinks/merged_cells, everything read "
        "(the rich=True equivalent)",
        lambda p: compat(p, EVERYTHING, formula_and_value=True, **META, **RO),
    ),
    (
        "normal_values",
        "compat normal mode",
        "iter_rows(values_only=True)",
        lambda p: compat(p, F({"values"})),
    ),
    (
        "normal_everything",
        "compat normal mode",
        "formula_and_value, everything read (value, formula, cached value, styles, "
        "comment, hyperlink)",
        lambda p: compat(p, EVERYTHING, formula_and_value=True),
    ),
]

# Cost breakdown of ro_everything: each step adds one kind of work to the previous.
BREAKDOWN: list[Scenario] = [
    (
        "step1_iterate",
        "breakdown",
        "iter_rows(): ReadOnlyCell objects created, nothing read",
        lambda p: compat(p, F(), **RO),
    ),
    (
        "step2_value",
        "breakdown",
        "+ value read",
        lambda p: compat(p, F({"value"}), **RO),
    ),
    (
        "step3_styles",
        "breakdown",
        "+ font, fill, border, number_format read",
        lambda p: compat(p, F({"value", "styles"}), **RO),
    ),
    (
        "step4_formulas",
        "breakdown",
        "+ formula_and_value=True: formula and cached_value read",
        lambda p: compat(
            p, F({"value", "styles", "formula"}), formula_and_value=True, **RO
        ),
    ),
    (
        "step5_metadata",
        "breakdown",
        "+ read_comments/hyperlinks/merged_cells: comment, hyperlink, merged_range read",
        lambda p: compat(p, EVERYTHING, formula_and_value=True, **META, **RO),
    ),
]
BY_KEY = {s[0]: s for s in SCENARIOS + BREAKDOWN}


def run_one(key: str, path: str) -> dict[str, float]:
    import python_calamine_plus  # noqa: F401  (import cost is not measured)

    base = peak_mb()
    t = time.perf_counter()
    rows, cells = BY_KEY[key][3](path)
    secs = time.perf_counter() - t
    return {"rows": rows, "cells": cells, "secs": secs, "peak_mb": peak_mb() - base}


# ---- driver ------------------------------------------------------------------------


def measure(keys: list[str], book: str, repeat: int) -> dict[str, dict[str, float]]:
    results = {}
    for key in keys:
        runs = []
        for _ in range(repeat):
            out = subprocess.run(
                [sys.executable, os.path.abspath(__file__), book, "--run", key],
                capture_output=True,
                text=True,
                check=True,
            )
            runs.append(json.loads(out.stdout))
        results[key] = min(runs, key=lambda r: r["secs"])
        print(f"  {key}: {results[key]['secs']:.1f}s", file=sys.stderr, flush=True)
    return results


def main() -> None:
    ap = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    ap.add_argument("book")
    default_out = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "profile_results.txt"
    )
    ap.add_argument("-o", "--output", default=default_out)
    ap.add_argument(
        "--repeat", type=int, default=1, help="runs per scenario; the fastest is kept"
    )
    ap.add_argument("--only", nargs="*", help="scenario keys to run")
    ap.add_argument("--run", help=argparse.SUPPRESS)
    args = ap.parse_args()

    if args.run:
        print(json.dumps(run_one(args.run, args.book)))
        return

    import python_calamine_plus

    lines: list[str] = []

    def emit(s: str = "") -> None:
        lines.append(s)

    size = os.path.getsize(args.book) / 1e6
    emit("python-calamine-plus: native vs compatibility='openpyxl' reading modes")
    emit("=" * 96)
    emit(f"date     : {datetime.datetime.now():%Y-%m-%d %H:%M}")
    emit(f"workbook : {os.path.basename(args.book)} ({size:.1f} MB)")
    emit(
        f"python   : {platform.python_version()} ({platform.system()} {platform.machine()}), "
        f"{os.cpu_count()} CPUs"
    )
    emit(f"library  : {os.path.dirname(python_calamine_plus.__file__)}")
    emit(
        f"method   : each scenario in a fresh process, every row of every sheet, best of {args.repeat}"
    )
    emit(
        "           peak = increase of the process's peak resident memory while reading"
    )
    emit()

    def table(scenarios: list[Scenario], baseline: str | None, diff: bool) -> None:
        keys = [s[0] for s in scenarios if not args.only or s[0] in args.only]
        res = measure(keys, args.book, args.repeat)
        head = f"{'scenario':18s} {'time':>8s} {'peak':>7s} {'rows':>9s} {'cells':>11s} {'cells/s':>9s}"
        if diff:
            head += f" {'added':>8s}"
        group, prev = None, None
        for key, grp, desc, _ in scenarios:
            if key not in res:
                continue
            if grp != group:
                group = grp
                emit(f"--- {grp} " + "-" * (len(head) - len(grp) - 5))
                emit(head)
            r = res[key]
            rate = r["cells"] / r["secs"] / 1e6 if r["secs"] else 0.0
            line = (
                f"{key:18s} {r['secs']:7.1f}s {r['peak_mb']:5.0f}MB {int(r['rows']):9,d} "
                f"{int(r['cells']):11,d} {rate:7.2f}M/s"
            )
            if diff:
                added = r["secs"] - prev if prev is not None else r["secs"]
                line += f" {added:+7.1f}s"
                prev = r["secs"]
            emit(line)
            note = desc
            if baseline and baseline in res and key != baseline:
                note += f"  [{r['secs'] / res[baseline]['secs']:.2f}x {baseline}]"
            emit(f"{'':18s} {note}")
        emit()

    table(SCENARIOS, "native_rich", diff=False)
    emit(
        "Cost breakdown of ro_everything (read_only=True, every cell; 'added' = this step's cost)"
    )
    table(BREAKDOWN, None, diff=True)
    emit("Notes")
    emit(
        "- 'cells' counts every position yielded: read-only rows are padded to the sheet's"
    )
    emit(
        "  declared dimension (like openpyxl), native rows to the widest row seen so far."
    )
    emit(
        "- The formula view (data_only=False) builds the text of every formula; data_only=True"
    )
    emit(
        "  does not, so ro_values_data is faster than ro_values on formula-heavy files."
    )
    emit(
        "- Read-only metadata flags, normal mode and rich=True read the sheet's metadata (merged"
    )
    emit(
        "  ranges, hyperlinks, comments) with one extra pass over the sheet when iteration starts."
    )
    emit(
        "- Style objects are created once per distinct cell format and shared by its cells."
    )

    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))
    print(f"written to {args.output}")


if __name__ == "__main__":
    main()
