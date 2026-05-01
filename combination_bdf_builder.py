#!/usr/bin/env python3
"""
Combination BDF Builder
-----------------------
Reads a load combination Excel file, recursively finds unit case BDF files
under a given root directory, and writes a single output BDF that references
them via INCLUDE statements together with LOAD combination cards.

Excel formats supported
  Wide  : First column = combination name/ID, remaining column headers = unit
          case IDs, cell values = scale factors (0 or empty → skip).
  Long  : Three columns: combination name, unit case ID, scale factor.

Usage (CLI)
  python combination_bdf_builder.py \\
      --excel  combinations.xlsx \\
      --bdf-root /path/to/unit/cases \\
      --output  combined.bdf \\
      [--format wide|long] \\
      [--sheet  0] \\
      [--combo-col  "Combo"] \\
      [--case-col   "UnitCase"] \\
      [--factor-col "Factor"] \\
      [--relative-paths] \\
      [--combo-filter COMBO1 COMBO2 ...]

Usage (GUI)
  python combination_bdf_builder.py --gui
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is required. Run: pip install pandas openpyxl")


# ---------------------------------------------------------------------------
# BDF file discovery
# ---------------------------------------------------------------------------

def find_bdf_files(root_dir: str, unit_case_ids: List[str]) -> Dict[str, Optional[str]]:
    """
    Recursively search *root_dir* for BDF files whose stem (filename without
    extension) contains a unit case ID string.

    Matching priority
      1. Exact stem match (case-insensitive)
      2. Stem contains the ID as a substring (case-insensitive)
         - If multiple files match, the one with the shortest stem wins.

    Returns {case_id: absolute_path_or_None}.
    """
    root = Path(root_dir)
    if not root.is_dir():
        raise NotADirectoryError(f"BDF root directory not found: {root_dir}")

    # Index all .bdf files once
    stem_to_path: Dict[str, str] = {}
    for bdf_file in root.rglob("*.bdf"):
        stem_lower = bdf_file.stem.lower()
        # Keep the first occurrence when two files share an identical lower-case stem
        if stem_lower not in stem_to_path:
            stem_to_path[stem_lower] = str(bdf_file.resolve())

    result: Dict[str, Optional[str]] = {}
    for case_id in unit_case_ids:
        key = str(case_id).strip()
        key_lower = key.lower()
        found: Optional[str] = None

        # 1. Exact stem match
        if key_lower in stem_to_path:
            found = stem_to_path[key_lower]
        else:
            # 2. Substring match
            matches = [
                path for stem, path in stem_to_path.items()
                if key_lower in stem
            ]
            if len(matches) == 1:
                found = matches[0]
            elif len(matches) > 1:
                # Prefer shortest stem (most specific / least decorated name)
                found = min(matches, key=lambda p: len(Path(p).stem))

        result[key] = found

    return result


# ---------------------------------------------------------------------------
# Excel parsing
# ---------------------------------------------------------------------------

def parse_excel_wide(df: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    """
    Wide format
      Column 0        → combination name
      Column 1..N     → unit case IDs (headers)
      Cell value      → scale factor (0 or blank → excluded)
    """
    combinations: Dict[str, Dict[str, float]] = {}
    combo_col = df.columns[0]
    unit_cols = df.columns[1:]

    for _, row in df.iterrows():
        combo_name = str(row[combo_col]).strip()
        if not combo_name or combo_name.lower() == "nan":
            continue
        cases: Dict[str, float] = {}
        for col in unit_cols:
            raw = row[col]
            try:
                factor = float(raw)
            except (ValueError, TypeError):
                continue
            if factor != 0.0:
                cases[str(col).strip()] = factor
        if cases:
            combinations[combo_name] = cases

    return combinations


def parse_excel_long(
    df: pd.DataFrame,
    combo_col: str,
    case_col: str,
    factor_col: str,
) -> Dict[str, Dict[str, float]]:
    """
    Long format: each row is (combination_name, unit_case_id, scale_factor).
    """
    combinations: Dict[str, Dict[str, float]] = {}
    for _, row in df.iterrows():
        combo_name = str(row[combo_col]).strip()
        case_id = str(row[case_col]).strip()
        if not combo_name or combo_name.lower() == "nan":
            continue
        try:
            factor = float(row[factor_col])
        except (ValueError, TypeError):
            continue
        if factor == 0.0:
            continue
        combinations.setdefault(combo_name, {})[case_id] = factor

    return combinations


def load_combinations(
    excel_path: str,
    fmt: str = "wide",
    sheet=0,
    combo_col: str = "",
    case_col: str = "",
    factor_col: str = "",
) -> Dict[str, Dict[str, float]]:
    """Load and parse the Excel combination file."""
    df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
    # Drop fully-empty rows/columns
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    if fmt == "wide":
        return parse_excel_wide(df)
    elif fmt == "long":
        missing = [c for c in [combo_col, case_col, factor_col] if not c]
        if missing:
            raise ValueError(
                "Long format requires --combo-col, --case-col, --factor-col."
            )
        for col in [combo_col, case_col, factor_col]:
            if col not in df.columns:
                raise ValueError(
                    f"Column '{col}' not found in Excel. "
                    f"Available: {list(df.columns)}"
                )
        return parse_excel_long(df, combo_col, case_col, factor_col)
    else:
        raise ValueError(f"Unknown format: {fmt!r}. Use 'wide' or 'long'.")


# ---------------------------------------------------------------------------
# BDF writer
# ---------------------------------------------------------------------------

_BDF_LINE_WIDTH = 72  # Nastran free-field / comment width


def _nastran_load_card(sid: int, cases: Dict[str, float], case_sid_map: Dict[str, int]) -> List[str]:
    """
    Build LOAD bulk data card lines for a combination.
    LOAD SID S  S1 L1 S2 L2 ... (8 fields per continuation line).

    Only includes unit cases that appear in case_sid_map.
    Returns empty list if no cases are available.
    """
    pairs = [
        (factor, case_sid_map[cid])
        for cid, factor in cases.items()
        if cid in case_sid_map
    ]
    if not pairs:
        return []

    # Nastran small-field: 8 characters each field, 10 fields per line
    # LOAD    SID     S       S1      L1      S2      L2      S3      L3
    # Field widths: 8 chars each
    all_fields = []
    for factor, load_sid in pairs:
        all_fields.append(f"{factor:8g}")
        all_fields.append(f"{load_sid:8d}")

    lines = []
    # First line: LOAD + SID + overall_scale (1.0) + up to 3 pairs (6 fields)
    chunk_size = 6  # S1 L1 S2 L2 S3 L3
    first_chunk = all_fields[:chunk_size]
    rest = all_fields[chunk_size:]
    line = f"{'LOAD':<8}{sid:8d}{'1.0':>8}" + "".join(first_chunk)
    lines.append(line.rstrip() + "\n")

    while rest:
        chunk = rest[:8]
        rest = rest[8:]
        line = "+" + " " * 7 + "".join(chunk)
        lines.append(line.rstrip() + "\n")

    return lines


def write_combination_bdf(
    output_path: str,
    combinations: Dict[str, Dict[str, float]],
    bdf_map: Dict[str, Optional[str]],
    use_relative_paths: bool = False,
    combo_filter: Optional[List[str]] = None,
    start_load_sid: int = 10000,
) -> Tuple[int, List[str]]:
    """
    Write the output BDF file.

    Structure
      SOL 101  (placeholder – user should adjust to their solution)
      CEND
        SUBCASE per combination with LOAD = <auto-SID>
      BEGIN BULK
        INCLUDE statements for all unique unit case BDF files
        LOAD cards for each combination
      ENDDATA

    Returns (combos_written, missing_case_ids).
    """
    output_dir = Path(output_path).parent
    missing: List[str] = []

    if combo_filter:
        filter_set = {c.strip() for c in combo_filter}
        combinations = {k: v for k, v in combinations.items() if k in filter_set}

    # Assign a unique SID to each unit case found
    unique_cases: List[str] = sorted(
        {cid for cases in combinations.values() for cid in cases}
    )
    unit_sid_map: Dict[str, int] = {}  # case_id → assumed SID inside unit BDF

    # Collect BDF paths that were actually found
    found_case_ids = [cid for cid in unique_cases if bdf_map.get(cid)]
    not_found = [cid for cid in unique_cases if not bdf_map.get(cid)]
    missing.extend(not_found)

    def _include_path(abs_path: str) -> str:
        if use_relative_paths:
            try:
                return os.path.relpath(abs_path, output_dir)
            except ValueError:
                return abs_path
        return abs_path

    # Assign auto SIDs for combinations (used in CASE CONTROL and LOAD cards).
    # We cannot know the SIDs inside the unit case BDFs without parsing them,
    # so we skip LOAD card generation by default and rely on pure INCLUDE.
    # A comment block documents each combination's scale factors.
    combo_sids: Dict[str, int] = {
        name: start_load_sid + i
        for i, name in enumerate(combinations)
    }

    # ---- Build the file ----
    lines: List[str] = []

    def c(text: str = ""):
        lines.append(f"$ {text}\n" if text else "$\n")

    c("=" * (_BDF_LINE_WIDTH - 2))
    c("  Combination Load Case BDF")
    c("  Generated by: combination_bdf_builder.py")
    c("=" * (_BDF_LINE_WIDTH - 2))
    c()
    lines.append("SOL 101\n")
    lines.append("CEND\n")
    c()
    c("CASE CONTROL")
    c()

    for combo_name, cases in combinations.items():
        sid = combo_sids[combo_name]
        c(f"  Combination : {combo_name}")
        for cid, factor in cases.items():
            status = "OK" if bdf_map.get(cid) else "NOT FOUND"
            c(f"    {cid:>12}  x {factor:g}  [{status}]")
        lines.append(f"SUBCASE {sid}\n")
        lines.append(f"  LABEL = {combo_name}\n")
        lines.append(f"$ LOAD = {sid}  $ Uncomment after verifying load SIDs\n")
        c()

    lines.append("BEGIN BULK\n")
    c()
    c("-" * (_BDF_LINE_WIDTH - 2))
    c("  INCLUDE statements for unit case BDF files")
    c("-" * (_BDF_LINE_WIDTH - 2))
    c()

    for cid in unique_cases:
        bdf_path = bdf_map.get(cid)
        if bdf_path:
            ip = _include_path(bdf_path)
            c(f"Unit Case: {cid}")
            lines.append(f"INCLUDE '{ip}'\n")
        else:
            c(f"WARNING: BDF file NOT FOUND for unit case: {cid}")
        c()

    c()
    c("-" * (_BDF_LINE_WIDTH - 2))
    c("  Combination scale-factor reference (informational)")
    c("  To activate LOAD cards you must know the SID of each")
    c("  unit case and uncomment the LOAD entries below.")
    c("-" * (_BDF_LINE_WIDTH - 2))
    c()

    for combo_name, cases in combinations.items():
        sid = combo_sids[combo_name]
        c(f"Combination : {combo_name}  (SID={sid})")
        for cid, factor in cases.items():
            c(f"  LOAD {sid}  1.0  {factor:g}  <SID of {cid}>")
        c()

    lines.append("ENDDATA\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    return len(combinations), missing


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def print_report(
    excel_path: str,
    bdf_root: str,
    output_path: str,
    combinations: Dict[str, Dict[str, float]],
    bdf_map: Dict[str, Optional[str]],
    missing: List[str],
    combos_written: int,
):
    total_cases = len(bdf_map)
    found_cases = sum(1 for v in bdf_map.values() if v)

    print("\n" + "=" * 60)
    print("  Combination BDF Builder – Report")
    print("=" * 60)
    print(f"  Excel file   : {excel_path}")
    print(f"  BDF root     : {bdf_root}")
    print(f"  Output       : {output_path}")
    print(f"  Combinations : {combos_written}")
    print(f"  Unit cases   : {found_cases} / {total_cases} found")

    if missing:
        print(f"\n  [!] {len(missing)} unit case BDF(s) NOT FOUND:")
        for cid in missing:
            print(f"      - {cid}")
    else:
        print("\n  All unit case BDF files located successfully.")

    print("=" * 60)

    if not missing:
        print(f"\n  Output written to: {output_path}\n")


# ---------------------------------------------------------------------------
# GUI (tkinter)
# ---------------------------------------------------------------------------

def run_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk, scrolledtext
    except ImportError:
        sys.exit("tkinter is not available on this system.")

    root = tk.Tk()
    root.title("Combination BDF Builder")
    root.resizable(True, True)

    pad = {"padx": 6, "pady": 4}

    # ── Variables ──────────────────────────────────────────────────────────
    var_excel = tk.StringVar()
    var_bdf_root = tk.StringVar()
    var_output = tk.StringVar()
    var_fmt = tk.StringVar(value="wide")
    var_sheet = tk.StringVar(value="0")
    var_combo_col = tk.StringVar()
    var_case_col = tk.StringVar()
    var_factor_col = tk.StringVar()
    var_rel_paths = tk.BooleanVar(value=False)
    var_filter = tk.StringVar()

    # ── Layout helpers ─────────────────────────────────────────────────────
    def _row(parent, label, variable, browse_cmd=None, row=0):
        tk.Label(parent, text=label, anchor="w").grid(
            row=row, column=0, sticky="w", **pad
        )
        e = tk.Entry(parent, textvariable=variable, width=60)
        e.grid(row=row, column=1, sticky="ew", **pad)
        if browse_cmd:
            tk.Button(parent, text="Browse…", command=browse_cmd).grid(
                row=row, column=2, **pad
            )
        return row + 1

    def browse_excel():
        p = filedialog.askopenfilename(
            title="Select Combination Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if p:
            var_excel.set(p)

    def browse_bdf_root():
        p = filedialog.askdirectory(title="Select BDF Root Directory")
        if p:
            var_bdf_root.set(p)

    def browse_output():
        p = filedialog.asksaveasfilename(
            title="Save Output BDF As",
            defaultextension=".bdf",
            filetypes=[("BDF files", "*.bdf"), ("All files", "*.*")],
        )
        if p:
            var_output.set(p)

    # ── Main frame ─────────────────────────────────────────────────────────
    main = tk.Frame(root)
    main.pack(fill="both", expand=True, padx=10, pady=10)
    main.columnconfigure(1, weight=1)

    r = 0
    r = _row(main, "Combination Excel:", var_excel, browse_excel, r)
    r = _row(main, "Unit BDF Root Dir:", var_bdf_root, browse_bdf_root, r)
    r = _row(main, "Output BDF File:", var_output, browse_output, r)

    # Format selector
    tk.Label(main, text="Excel Format:", anchor="w").grid(
        row=r, column=0, sticky="w", **pad
    )
    fmt_frame = tk.Frame(main)
    fmt_frame.grid(row=r, column=1, sticky="w", **pad)
    tk.Radiobutton(fmt_frame, text="Wide", variable=var_fmt, value="wide").pack(side="left")
    tk.Radiobutton(fmt_frame, text="Long", variable=var_fmt, value="long").pack(side="left", padx=10)
    r += 1

    r = _row(main, "Sheet (name or index):", var_sheet, None, r)

    tk.Label(main, text="[Long format only]", fg="gray", anchor="w").grid(
        row=r, column=0, columnspan=3, sticky="w", padx=6
    )
    r += 1
    r = _row(main, "  Combo column:", var_combo_col, None, r)
    r = _row(main, "  Case ID column:", var_case_col, None, r)
    r = _row(main, "  Factor column:", var_factor_col, None, r)

    tk.Checkbutton(main, text="Use relative paths in INCLUDE", variable=var_rel_paths).grid(
        row=r, column=0, columnspan=2, sticky="w", **pad
    )
    r += 1

    r = _row(main, "Filter combos (comma-separated, blank=all):", var_filter, None, r)

    # ── Log area ───────────────────────────────────────────────────────────
    log = scrolledtext.ScrolledText(main, height=12, state="disabled", wrap="word")
    log.grid(row=r, column=0, columnspan=3, sticky="nsew", **pad)
    main.rowconfigure(r, weight=1)
    r += 1

    def _log(msg: str):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    # ── Run button ─────────────────────────────────────────────────────────
    def run():
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")

        excel = var_excel.get().strip()
        bdf_root = var_bdf_root.get().strip()
        output = var_output.get().strip()

        if not excel or not bdf_root or not output:
            messagebox.showerror("Error", "Excel, BDF root, and output path are required.")
            return

        sheet_raw = var_sheet.get().strip()
        try:
            sheet = int(sheet_raw)
        except ValueError:
            sheet = sheet_raw

        fmt = var_fmt.get()
        combo_col = var_combo_col.get().strip()
        case_col = var_case_col.get().strip()
        factor_col = var_factor_col.get().strip()
        rel = var_rel_paths.get()
        filter_raw = var_filter.get().strip()
        combo_filter = [x.strip() for x in filter_raw.split(",")] if filter_raw else None

        try:
            _log("Reading Excel…")
            combinations = load_combinations(
                excel, fmt, sheet, combo_col, case_col, factor_col
            )
            _log(f"  → {len(combinations)} combination(s) loaded.")

            unique_cases = sorted(
                {cid for cases in combinations.values() for cid in cases}
            )
            _log(f"  → {len(unique_cases)} unique unit case ID(s) found.")

            _log("Searching for BDF files…")
            bdf_map = find_bdf_files(bdf_root, unique_cases)
            found = sum(1 for v in bdf_map.values() if v)
            _log(f"  → {found} / {len(unique_cases)} BDF file(s) located.")

            _log("Writing output BDF…")
            combos_written, missing = write_combination_bdf(
                output, combinations, bdf_map, rel, combo_filter
            )
            _log(f"  → {combos_written} combination(s) written.")

            if missing:
                _log(f"\n[!] {len(missing)} BDF file(s) NOT FOUND:")
                for cid in missing:
                    _log(f"    - {cid}")
            else:
                _log("\nAll unit case BDF files found.")

            _log(f"\nDone!  Output: {output}")
            messagebox.showinfo("Done", f"Output written to:\n{output}")

        except Exception as exc:
            _log(f"\nERROR: {exc}")
            messagebox.showerror("Error", str(exc))

    btn_frame = tk.Frame(main)
    btn_frame.grid(row=r, column=0, columnspan=3, pady=8)
    tk.Button(btn_frame, text="Run", command=run, width=16, bg="#2a7ae2", fg="white").pack(side="left", padx=4)
    tk.Button(btn_frame, text="Quit", command=root.destroy, width=10).pack(side="left", padx=4)

    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--gui", action="store_true", help="Open graphical interface")
    p.add_argument("--excel", help="Path to the combination Excel file")
    p.add_argument("--bdf-root", help="Root directory to search for unit case BDF files")
    p.add_argument("--output", help="Output BDF file path")
    p.add_argument(
        "--format", choices=["wide", "long"], default="wide",
        help="Excel layout: 'wide' (default) or 'long'"
    )
    p.add_argument("--sheet", default="0",
        help="Excel sheet name or 0-based index (default: 0)")
    p.add_argument("--combo-col", default="", help="[Long] Combination name column header")
    p.add_argument("--case-col", default="", help="[Long] Unit case ID column header")
    p.add_argument("--factor-col", default="", help="[Long] Scale factor column header")
    p.add_argument("--relative-paths", action="store_true",
        help="Write INCLUDE paths relative to the output BDF directory")
    p.add_argument("--combo-filter", nargs="*",
        help="Only process these combination names (space-separated)")
    p.add_argument("--start-load-sid", type=int, default=10000,
        help="Starting SID for combination LOAD entries (default: 10000)")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.gui:
        run_gui()
        return

    # CLI mode – all three paths are required
    errors = []
    if not args.excel:
        errors.append("--excel is required")
    if not args.bdf_root:
        errors.append("--bdf-root is required")
    if not args.output:
        errors.append("--output is required")
    if errors:
        for e in errors:
            print(f"[ERROR] {e}", file=sys.stderr)
        parser.print_usage(sys.stderr)
        sys.exit(1)

    sheet_raw = args.sheet
    try:
        sheet = int(sheet_raw)
    except ValueError:
        sheet = sheet_raw

    print(f"Reading Excel: {args.excel}")
    combinations = load_combinations(
        args.excel,
        args.format,
        sheet,
        args.combo_col,
        args.case_col,
        args.factor_col,
    )
    print(f"  → {len(combinations)} combination(s) loaded.")

    unique_cases = sorted(
        {cid for cases in combinations.values() for cid in cases}
    )
    print(f"  → {len(unique_cases)} unique unit case ID(s): {unique_cases[:10]}{'…' if len(unique_cases) > 10 else ''}")

    print(f"\nSearching for BDF files under: {args.bdf_root}")
    bdf_map = find_bdf_files(args.bdf_root, unique_cases)
    found = sum(1 for v in bdf_map.values() if v)
    print(f"  → {found} / {len(unique_cases)} found.")

    print(f"\nWriting output BDF: {args.output}")
    combos_written, missing = write_combination_bdf(
        args.output,
        combinations,
        bdf_map,
        args.relative_paths,
        args.combo_filter,
        args.start_load_sid,
    )

    print_report(
        args.excel,
        args.bdf_root,
        args.output,
        combinations,
        bdf_map,
        missing,
        combos_written,
    )

    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
