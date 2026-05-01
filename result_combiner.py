#!/usr/bin/env python3
"""
Result Combiner
---------------
İki GFEM'den üretilmiş unit case .op2 sonuç dosyalarını, kombinasyon Excel
tablosundaki katsayılara göre dışarıda (solver dışı) toplar.

Varsayımlar
  • Her unit case → tek bir .op2 dosyası (subcase 1 kullanılır)
  • İki GFEM aynı (ya da neredeyse aynı) node yapısına sahip; ortak
    node ID'lerinde superposition uygulanır
  • Structural unit case'ler (MAINCASE, CABINCASE, vs.) → --structural-root
  • Thermal unit case'ler (THERMALCASE)               → --thermal-root

Çıktı (--output-dir)
  combo_<SID>.csv   – her kombinasyon için node bazlı birleşik deplasman
  summary.csv       – tüm kombinasyonlarda max toplam deplasman büyüklüğü

Kullanım (CLI)
  python result_combiner.py \\
      --excel  kombinasyonlar.xlsx \\
      --structural-root /gfem1/results \\
      --thermal-root    /gfem2/results \\
      --output-dir      ./combined_results \\
      [--format paired|wide|long] \\
      [--sheet 0] \\
      [--subcase 1] \\
      [--components T1 T2 T3] \\
      [--combo-filter COMBO1 COMBO2 ...]

Kullanım (GUI)
  python result_combiner.py --gui
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas gerekli. Kur: pip install pandas openpyxl")

try:
    from pyNastran.op2.op2 import OP2
except ImportError:
    sys.exit("pyNastran gerekli. Kur: pip install pyNastran")

# combination_bdf_builder'dan Excel ayrıştırıcıları ve tip yardımcılarını al
sys.path.insert(0, str(Path(__file__).parent))
from combination_bdf_builder import (
    CaseEntry,
    Combinations,
    _is_thermal,
    load_combinations,
)

# Deplasman bileşen indeksleri (pyNastran: tx ty tz rx ry rz → 0..5)
_COMP_MAP = {"T1": 0, "T2": 1, "T3": 2, "R1": 3, "R2": 4, "R3": 5}
_ALL_COMPS = list(_COMP_MAP.keys())


# ---------------------------------------------------------------------------
# OP2 dosyası arama
# ---------------------------------------------------------------------------

def find_op2_files(root_dir: str, unit_case_ids: List[str]) -> Dict[str, Optional[str]]:
    """
    *root_dir* altında .op2 dosyalarını özyinelemeli arar.
    find_bdf_files ile aynı eşleşme mantığı kullanılır:
      1. Tam stem eşleşmesi (büyük/küçük harf duyarsız)
      2. Stem içinde ID barındırıyorsa → en kısa stem kazanır
    """
    root = Path(root_dir)
    if not root.is_dir():
        raise NotADirectoryError(f"OP2 kök dizini bulunamadı: {root_dir}")

    stem_to_path: Dict[str, str] = {}
    for op2_file in root.rglob("*.op2"):
        stem_lower = op2_file.stem.lower()
        if stem_lower not in stem_to_path:
            stem_to_path[stem_lower] = str(op2_file.resolve())

    result: Dict[str, Optional[str]] = {}
    for case_id in unit_case_ids:
        key = str(case_id).strip()
        key_lower = key.lower()
        found: Optional[str] = None

        if key_lower in stem_to_path:
            found = stem_to_path[key_lower]
        else:
            matches = [p for s, p in stem_to_path.items() if key_lower in s]
            if len(matches) == 1:
                found = matches[0]
            elif len(matches) > 1:
                found = min(matches, key=lambda p: len(Path(p).stem))

        result[key] = found

    return result


def find_op2_files_by_type(
    combinations: Combinations,
    structural_root: str,
    thermal_root: str,
) -> Dict[str, Optional[str]]:
    """
    Unit case ID'lerini tipine göre ayırır:
      THERMALCASE → thermal_root
      Diğerleri   → structural_root
    """
    structural_ids: List[str] = []
    thermal_ids: List[str] = []
    seen: set = set()

    for entries in combinations.values():
        for type_name, case_id, _ in entries:
            if case_id in seen:
                continue
            seen.add(case_id)
            if _is_thermal(type_name):
                thermal_ids.append(case_id)
            else:
                structural_ids.append(case_id)

    op2_map: Dict[str, Optional[str]] = {}

    if structural_ids:
        op2_map.update(find_op2_files(structural_root, structural_ids))

    if thermal_ids:
        if not thermal_root:
            for cid in thermal_ids:
                op2_map[cid] = None
        else:
            op2_map.update(find_op2_files(thermal_root, thermal_ids))

    return op2_map


# ---------------------------------------------------------------------------
# OP2 okuma
# ---------------------------------------------------------------------------

def read_displacements(
    op2_path: str,
    subcase: int = 1,
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Belirtilen OP2 dosyasından deplasman verilerini okur.

    Döndürür
      node_ids : shape (N,)          – node ID'leri
      data     : shape (N, 6)        – T1 T2 T3 R1 R2 R3
    """
    op2 = OP2(debug=False, log=None)
    op2.read_op2(op2_path)

    if subcase not in op2.displacements:
        available = list(op2.displacements.keys())
        raise KeyError(
            f"Subcase {subcase} bulunamadı: {op2_path}\n"
            f"Mevcut subcaseler: {available}"
        )

    disp_obj = op2.displacements[subcase]
    # data shape: (ntimes, nnodes, 6) – ilk zaman adımı alınır
    node_ids = disp_obj.node_gridtype[:, 0].astype(int)
    data = disp_obj.data[0]  # (nnodes, 6)

    return node_ids, data


# ---------------------------------------------------------------------------
# Kombinasyon hesabı
# ---------------------------------------------------------------------------

def combine_results(
    combinations: Combinations,
    op2_map: Dict[str, Optional[str]],
    subcase: int = 1,
    components: Optional[List[str]] = None,
    combo_filter: Optional[List[str]] = None,
    log_fn=print,
) -> Dict[str, Tuple[np.ndarray, np.ndarray]]:
    """
    Her kombinasyon için deplasman sonuçlarını katsayıyla çarpıp toplar.

    Döndürür
      {combo_sid: (node_ids, combined_data)}
      combined_data shape: (N, len(components))
    """
    if components is None:
        components = _ALL_COMPS
    comp_indices = [_COMP_MAP[c] for c in components]

    if combo_filter:
        fs = {c.strip() for c in combo_filter}
        combinations = {k: v for k, v in combinations.items() if k in fs}

    # OP2 önbelleği: her dosyayı yalnızca bir kez oku
    cache: Dict[str, Tuple[np.ndarray, np.ndarray]] = {}

    def _load(case_id: str) -> Optional[Tuple[np.ndarray, np.ndarray]]:
        if case_id in cache:
            return cache[case_id]
        path = op2_map.get(case_id)
        if not path:
            log_fn(f"  [UYARI] OP2 bulunamadı: {case_id} – bu case atlanıyor")
            return None
        log_fn(f"  Okunuyor: {Path(path).name}  (case {case_id})")
        nids, d = read_displacements(path, subcase)
        cache[case_id] = (nids, d)
        return nids, d

    results: Dict[str, Tuple[np.ndarray, np.ndarray]] = {}

    for combo_sid, entries in combinations.items():
        log_fn(f"\nKombinasyon: {combo_sid}")

        # Referans node kümesini bul (ortak node'lar)
        ref_nids: Optional[np.ndarray] = None
        unit_data: List[Tuple[float, np.ndarray]] = []  # [(factor, data_Nx6)]

        for type_name, case_id, factor in entries:
            loaded = _load(case_id)
            if loaded is None:
                continue
            nids, d = loaded
            if ref_nids is None:
                ref_nids = nids
            else:
                # Ortak node ID'lerini bul
                common = np.intersect1d(ref_nids, nids)
                if len(common) < len(ref_nids):
                    log_fn(f"  [BİLGİ] {case_id}: {len(ref_nids) - len(common)} node eşleşmedi, ortak {len(common)} node kullanılıyor")
                ref_nids = common
            unit_data.append((factor, nids, d))

        if ref_nids is None or len(unit_data) == 0:
            log_fn(f"  [HATA] {combo_sid}: hiç geçerli unit case yok, atlanıyor")
            continue

        # Her unit case'i ref_nids sırasına hizala ve topla
        n = len(ref_nids)
        combined = np.zeros((n, len(comp_indices)), dtype=float)

        for factor, nids, d in unit_data:
            # ref_nids içindeki her node'un bu case'deki indeksini bul
            sorter = np.argsort(nids)
            idx_in_case = sorter[np.searchsorted(nids, ref_nids, sorter=sorter)]
            contrib = d[np.ix_(idx_in_case, comp_indices)]
            combined += factor * contrib

        results[combo_sid] = (ref_nids, combined)
        log_fn(f"  → Tamamlandı: {n} node, bileşenler={components}")

    return results


# ---------------------------------------------------------------------------
# CSV çıktısı
# ---------------------------------------------------------------------------

def write_results(
    results: Dict[str, Tuple[np.ndarray, np.ndarray]],
    output_dir: str,
    components: Optional[List[str]] = None,
) -> str:
    """
    Her kombinasyon için ayrı CSV + özet CSV yazar.
    Döndürür: output_dir
    """
    if components is None:
        components = _ALL_COMPS

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    summary_rows = []

    for combo_sid, (node_ids, data) in results.items():
        df = pd.DataFrame(data, columns=components)
        df.insert(0, "NODE_ID", node_ids)

        csv_path = out / f"combo_{combo_sid}.csv"
        df.to_csv(csv_path, index=False)

        # Toplam deplasman büyüklüğü (T1²+T2²+T3²)^0.5
        t_cols = [c for c in ["T1", "T2", "T3"] if c in components]
        if t_cols:
            mag = np.linalg.norm(data[:, [components.index(c) for c in t_cols]], axis=1)
            summary_rows.append({
                "Combo_SID":   combo_sid,
                "N_Nodes":     len(node_ids),
                "Max_Mag":     float(mag.max()),
                "Node_MaxMag": int(node_ids[mag.argmax()]),
                "Min_Mag":     float(mag.min()),
                "Node_MinMag": int(node_ids[mag.argmin()]),
            })

    if summary_rows:
        pd.DataFrame(summary_rows).to_csv(out / "summary.csv", index=False)

    return str(out)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def run_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext
    except ImportError:
        sys.exit("tkinter bu sistemde mevcut değil.")

    root = tk.Tk()
    root.title("Result Combiner – Skin Deflection")
    root.resizable(True, True)

    pad = {"padx": 6, "pady": 3}

    var_excel           = tk.StringVar()
    var_structural_root = tk.StringVar()
    var_thermal_root    = tk.StringVar()
    var_output_dir      = tk.StringVar()
    var_fmt             = tk.StringVar(value="paired")
    var_sheet           = tk.StringVar(value="0")
    var_subcase         = tk.StringVar(value="1")
    var_combo_col       = tk.StringVar()
    var_case_col        = tk.StringVar()
    var_factor_col      = tk.StringVar()
    var_filter          = tk.StringVar()

    # Bileşen checkboxları
    comp_vars = {c: tk.BooleanVar(value=True) for c in _ALL_COMPS}

    def _row(parent, label, variable, browse_cmd=None, row=0):
        tk.Label(parent, text=label, anchor="w").grid(row=row, column=0, sticky="w", **pad)
        tk.Entry(parent, textvariable=variable, width=56).grid(row=row, column=1, sticky="ew", **pad)
        if browse_cmd:
            tk.Button(parent, text="Gözat…", command=browse_cmd, width=8).grid(row=row, column=2, **pad)
        return row + 1

    def browse_excel():
        p = filedialog.askopenfilename(title="Kombinasyon Excel'ini Seç",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Tümü", "*.*")])
        if p: var_excel.set(p)

    def browse_structural():
        p = filedialog.askdirectory(title="Structural OP2 Kök Dizini (GFEM1)")
        if p: var_structural_root.set(p)

    def browse_thermal():
        p = filedialog.askdirectory(title="Thermal OP2 Kök Dizini (GFEM2)")
        if p: var_thermal_root.set(p)

    def browse_output():
        p = filedialog.askdirectory(title="Çıktı Klasörü")
        if p: var_output_dir.set(p)

    main = tk.Frame(root)
    main.pack(fill="both", expand=True, padx=10, pady=8)
    main.columnconfigure(1, weight=1)

    r = 0
    r = _row(main, "Kombinasyon Excel:", var_excel, browse_excel, r)
    r = _row(main, "Structural OP2 Kök (GFEM1):", var_structural_root, browse_structural, r)
    r = _row(main, "Thermal OP2 Kök (GFEM2):", var_thermal_root, browse_thermal, r)
    r = _row(main, "Çıktı Klasörü:", var_output_dir, browse_output, r)

    tk.Label(main, text="Excel Formatı:", anchor="w").grid(row=r, column=0, sticky="w", **pad)
    ff = tk.Frame(main); ff.grid(row=r, column=1, sticky="w", **pad)
    tk.Radiobutton(ff, text="Paired (önerilen)", variable=var_fmt, value="paired").pack(side="left")
    tk.Radiobutton(ff, text="Wide",  variable=var_fmt, value="wide").pack(side="left", padx=6)
    tk.Radiobutton(ff, text="Long",  variable=var_fmt, value="long").pack(side="left", padx=6)
    r += 1

    r = _row(main, "Sayfa (ad veya indeks):", var_sheet, None, r)
    r = _row(main, "OP2 Subcase no:", var_subcase, None, r)

    tk.Label(main, text="[Yalnızca Long format]", fg="gray", anchor="w").grid(
        row=r, column=0, columnspan=3, sticky="w", padx=6)
    r += 1
    r = _row(main, "  Kombo sütunu:",   var_combo_col,  None, r)
    r = _row(main, "  Case ID sütunu:", var_case_col,   None, r)
    r = _row(main, "  Katsayı sütunu:", var_factor_col, None, r)

    r = _row(main, "Filtrele (virgülle, boş=hepsi):", var_filter, None, r)

    # Bileşen seçimi
    tk.Label(main, text="Bileşenler:", anchor="w").grid(row=r, column=0, sticky="w", **pad)
    cf = tk.Frame(main); cf.grid(row=r, column=1, sticky="w", **pad)
    for c in _ALL_COMPS:
        tk.Checkbutton(cf, text=c, variable=comp_vars[c]).pack(side="left", padx=2)
    r += 1

    log = scrolledtext.ScrolledText(main, height=14, state="disabled", wrap="word")
    log.grid(row=r, column=0, columnspan=3, sticky="nsew", **pad)
    main.rowconfigure(r, weight=1)
    r += 1

    def _log(msg):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    def run():
        log.configure(state="normal"); log.delete("1.0", "end"); log.configure(state="disabled")
        excel           = var_excel.get().strip()
        structural_root = var_structural_root.get().strip()
        thermal_root    = var_thermal_root.get().strip()
        output_dir      = var_output_dir.get().strip()

        if not excel or not structural_root or not output_dir:
            messagebox.showerror("Hata", "Excel, Structural OP2 dizini ve çıktı klasörü zorunludur.")
            return

        sheet_raw = var_sheet.get().strip()
        try:
            sheet = int(sheet_raw)
        except ValueError:
            sheet = sheet_raw

        try:
            subcase = int(var_subcase.get().strip())
        except ValueError:
            subcase = 1

        components = [c for c in _ALL_COMPS if comp_vars[c].get()]
        if not components:
            messagebox.showerror("Hata", "En az bir bileşen seçilmeli.")
            return

        filter_raw = var_filter.get().strip()
        combo_filter = [x.strip() for x in filter_raw.split(",")] if filter_raw else None
        eff_thermal = thermal_root or structural_root

        try:
            _log("Excel okunuyor…")
            combs = load_combinations(
                excel, var_fmt.get(), sheet,
                var_combo_col.get().strip(),
                var_case_col.get().strip(),
                var_factor_col.get().strip(),
            )
            _log(f"  → {len(combs)} kombinasyon yüklendi.")

            _log("OP2 dosyaları aranıyor…")
            op2_map = find_op2_files_by_type(combs, structural_root, eff_thermal)
            found = sum(1 for v in op2_map.values() if v)
            _log(f"  → {found} / {len(op2_map)} OP2 bulundu.")

            _log("\nSonuçlar birleştiriliyor…")
            results = combine_results(combs, op2_map, subcase, components, combo_filter, _log)

            _log("\nCSV dosyaları yazılıyor…")
            out = write_results(results, output_dir, components)
            _log(f"  → {len(results)} kombinasyon yazıldı: {out}")

            messagebox.showinfo("Tamamlandı", f"{len(results)} kombinasyon CSV olarak yazıldı:\n{out}")

        except Exception as exc:
            import traceback
            _log(f"\nHATA: {exc}\n{traceback.format_exc()}")
            messagebox.showerror("Hata", str(exc))

    btn = tk.Frame(main)
    btn.grid(row=r, column=0, columnspan=3, pady=8)
    tk.Button(btn, text="Çalıştır", command=run, width=14, bg="#2a7ae2", fg="white").pack(side="left", padx=4)
    tk.Button(btn, text="Çıkış",   command=root.destroy, width=10).pack(side="left", padx=4)

    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--gui",              action="store_true", help="Grafik arayüzü aç")
    p.add_argument("--excel",            help="Kombinasyon Excel dosyası")
    p.add_argument("--structural-root",  help="Structural OP2 dosyalarının kök dizini (GFEM1)")
    p.add_argument("--thermal-root",     help="Thermal OP2 dosyalarının kök dizini (GFEM2); belirtilmezse --structural-root kullanılır")
    p.add_argument("--bdf-root",         help="Her iki tip için tek kök (geriye uyumluluk)")
    p.add_argument("--output-dir",       help="Çıktı CSV klasörü")
    p.add_argument("--format",  choices=["paired", "wide", "long"], default="paired",
        help="Excel formatı (varsayılan: paired)")
    p.add_argument("--sheet",   default="0", help="Excel sayfa adı veya indeks (varsayılan: 0)")
    p.add_argument("--subcase", type=int, default=1,
        help="OP2 içindeki subcase numarası (varsayılan: 1)")
    p.add_argument("--components", nargs="+", default=_ALL_COMPS,
        choices=_ALL_COMPS, metavar="COMP",
        help=f"Bileşenler: {_ALL_COMPS} (varsayılan: hepsi)")
    p.add_argument("--combo-col",  default="", help="[Long] Kombinasyon sütun başlığı")
    p.add_argument("--case-col",   default="", help="[Long] Unit case ID sütun başlığı")
    p.add_argument("--factor-col", default="", help="[Long] Katsayı sütun başlığı")
    p.add_argument("--combo-filter", nargs="*",
        help="Yalnızca bu kombinasyonları işle (boşlukla ayrılmış)")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.gui:
        run_gui()
        return

    structural_root = args.structural_root or args.bdf_root
    thermal_root    = args.thermal_root    or args.bdf_root

    errors = []
    if not args.excel:       errors.append("--excel gerekli")
    if not structural_root:  errors.append("--structural-root (veya --bdf-root) gerekli")
    if not args.output_dir:  errors.append("--output-dir gerekli")
    if errors:
        for e in errors:
            print(f"[HATA] {e}", file=sys.stderr)
        parser.print_usage(sys.stderr)
        sys.exit(1)

    sheet_raw = args.sheet
    try:
        sheet = int(sheet_raw)
    except ValueError:
        sheet = sheet_raw

    eff_thermal = thermal_root or structural_root

    print(f"Excel okunuyor: {args.excel}")
    combinations = load_combinations(
        args.excel, args.format, sheet,
        args.combo_col, args.case_col, args.factor_col,
    )
    print(f"  → {len(combinations)} kombinasyon yüklendi.")

    print(f"\nOP2 dosyaları aranıyor…")
    print(f"  Structural kök : {structural_root}")
    print(f"  Thermal kök    : {eff_thermal}")
    op2_map = find_op2_files_by_type(combinations, structural_root, eff_thermal)
    found = sum(1 for v in op2_map.values() if v)
    print(f"  → {found} / {len(op2_map)} bulundu.")

    missing = [cid for cid, p in op2_map.items() if not p]
    if missing:
        print(f"\n  [!] Bulunamayan OP2'ler: {missing}")

    print(f"\nSonuçlar birleştiriliyor…  (subcase={args.subcase})")
    results = combine_results(
        combinations, op2_map,
        args.subcase, args.components, args.combo_filter,
    )

    print(f"\nCSV dosyaları yazılıyor → {args.output_dir}")
    out = write_results(results, args.output_dir, args.components)

    print(f"\n{'=' * 56}")
    print(f"  Tamamlandı: {len(results)} kombinasyon → {out}")
    print(f"{'=' * 56}\n")

    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
