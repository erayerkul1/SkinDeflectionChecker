#!/usr/bin/env python3
"""
Combination BDF Builder
-----------------------
Gerçek Excel formatını okur, unit case BDF dosyalarını bulur ve
solver tipine göre (MSC/NX Nastran) birleşik bir BDF dosyası üretir.

Excel formatı – Paired Columns (varsayılan)
  Sütun A : Combined Load Case SID (çıktı BDF'teki subcase/combo SID'i)
  Kalan   : (TIP_ID, Multiplier) sütun çiftleri
  Örnek başlık: Combined Load Case | MAINCASE ID | Multiplier | THERMALCASE ID | Multiplier | ...

MSC Nastran çıktısı (--solver msc)
  Her unique unit case → ayrı SUBCASE
  Her kombinasyon     → SUBCOM + SUBSEQ (result-level superposition)

NX Nastran çıktısı (--solver nx)
  Her kombinasyon     → tek SUBCASE + LOAD bulk kart (load-level combination)
  THERMALCASE         → TEMP(LOAD) Case Control komutu ile eklenir

Kullanım (CLI)
  python combination_bdf_builder.py \\
      --excel  kombinasyonlar.xlsx \\
      --bdf-root /birim/case/dizini \\
      --output  combined.bdf \\
      --solver  nx|msc \\
      [--format paired|wide|long] \\
      [--sheet 0] \\
      [--relative-paths] \\
      [--combo-filter COMBO1 COMBO2 ...]

Kullanım (GUI)
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
    sys.exit("pandas gerekli. Kur: pip install pandas openpyxl")

# ---------------------------------------------------------------------------
# Tip tanımlamaları
# ---------------------------------------------------------------------------

# (tip_adı, unit_case_id, katsayı)
CaseEntry = Tuple[str, str, float]
# {combo_sid_str: [CaseEntry, ...]}
Combinations = Dict[str, List[CaseEntry]]

_THERMAL_KEYWORD = "THERMAL"  # tip adında bu kelime varsa → termal yük


# ---------------------------------------------------------------------------
# BDF dosyası arama
# ---------------------------------------------------------------------------

def find_bdf_files(root_dir: str, unit_case_ids: List[str]) -> Dict[str, Optional[str]]:
    """
    *root_dir* altında .bdf uzantılı dosyaları özyinelemeli arar.
    Her unit case ID için eşleşen dosyayı döndürür.

    Eşleşme önceliği
      1. Dosya steminin (uzantısız adı) tam eşleşmesi (büyük/küçük harf duyarsız)
      2. Stem içinde ID'yi barındıran dosyalar → stem en kısa olanı seç
    """
    root = Path(root_dir)
    if not root.is_dir():
        raise NotADirectoryError(f"BDF kök dizini bulunamadı: {root_dir}")

    stem_to_path: Dict[str, str] = {}
    for bdf_file in root.rglob("*.bdf"):
        stem_lower = bdf_file.stem.lower()
        if stem_lower not in stem_to_path:
            stem_to_path[stem_lower] = str(bdf_file.resolve())

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


# ---------------------------------------------------------------------------
# Excel ayrıştırıcılar
# ---------------------------------------------------------------------------

def _cell_to_case_id(raw) -> Optional[str]:
    """Ham Excel hücre değerini case ID string'ine dönüştür; boşsa None."""
    if raw is None:
        return None
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return None
    # 12001.0 → "12001"
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except ValueError:
        return s


def _extract_type_name(header: str) -> str:
    """'MAINCASE ID' → 'MAINCASE',  'THERMALCASE ID' → 'THERMALCASE'"""
    name = header.strip()
    if name.upper().endswith(" ID"):
        name = name[:-3].strip()
    return name.upper()


def parse_excel_paired_columns(df: pd.DataFrame) -> Combinations:
    """
    Paired-column (eşli sütun) formatını ayrıştırır.

    Sütun 0: Combined Load Case SID
    Kalan sütunlar soldan sağa taranır:
      - Başlık "multiplier" içeriyorsa → önceki ID sütununun katsayı sütunu
      - Aksi hâlde → yeni tip ID sütunu (başlıktan tip adı çıkarılır)
    """
    combo_col = df.columns[0]

    # (tip_adı, id_sütun_indeksi, katsayı_sütun_indeksi) üçlülerini bul
    pairs: List[Tuple[str, int, int]] = []
    pending: Optional[Tuple[str, int]] = None  # (tip_adı, sütun_indeksi)

    for i, header in enumerate(df.columns[1:], start=1):
        if "multiplier" in str(header).lower():
            if pending is not None:
                pairs.append((pending[0], pending[1], i))
                pending = None
        else:
            pending = (_extract_type_name(str(header)), i)

    combinations: Combinations = {}
    for _, row in df.iterrows():
        combo_sid = _cell_to_case_id(row[combo_col])
        if combo_sid is None:
            continue

        entries: List[CaseEntry] = []
        for type_name, id_idx, mult_idx in pairs:
            case_id = _cell_to_case_id(row[df.columns[id_idx]])
            if case_id is None:
                continue
            try:
                factor = float(row[df.columns[mult_idx]])
            except (ValueError, TypeError):
                factor = 1.0
            entries.append((type_name, case_id, factor))

        if entries:
            combinations[combo_sid] = entries

    return combinations


def parse_excel_wide(df: pd.DataFrame) -> Combinations:
    """
    Geniş format (geriye dönük uyumluluk).
    Sütun 0: kombinasyon adı, kalan sütun başlıkları: unit case ID'ler, değerler: katsayılar.
    """
    combo_col = df.columns[0]
    unit_cols = df.columns[1:]
    combinations: Combinations = {}

    for _, row in df.iterrows():
        combo_name = str(row[combo_col]).strip()
        if not combo_name or combo_name.lower() == "nan":
            continue
        entries: List[CaseEntry] = []
        for col in unit_cols:
            try:
                factor = float(row[col])
            except (ValueError, TypeError):
                continue
            if factor != 0.0:
                entries.append(("LOAD", str(col).strip(), factor))
        if entries:
            combinations[combo_name] = entries

    return combinations


def parse_excel_long(
    df: pd.DataFrame,
    combo_col: str,
    case_col: str,
    factor_col: str,
) -> Combinations:
    """
    Uzun format (geriye dönük uyumluluk).
    Her satır: (kombinasyon_adı, unit_case_id, katsayı).
    """
    combinations: Combinations = {}
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
        combinations.setdefault(combo_name, []).append(("LOAD", case_id, factor))
    return combinations


def load_combinations(
    excel_path: str,
    fmt: str = "paired",
    sheet=0,
    combo_col: str = "",
    case_col: str = "",
    factor_col: str = "",
) -> Combinations:
    """Excel kombinasyon dosyasını okuyup ayrıştırır."""
    df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    if fmt == "paired":
        return parse_excel_paired_columns(df)
    elif fmt == "wide":
        return parse_excel_wide(df)
    elif fmt == "long":
        for col in [combo_col, case_col, factor_col]:
            if not col:
                raise ValueError("Long format --combo-col, --case-col, --factor-col gerektirir.")
            if col not in df.columns:
                raise ValueError(
                    f"Sütun '{col}' Excel'de bulunamadı. Mevcut: {list(df.columns)}"
                )
        return parse_excel_long(df, combo_col, case_col, factor_col)
    else:
        raise ValueError(f"Bilinmeyen format: {fmt!r}. 'paired', 'wide' veya 'long' kullanın.")


# ---------------------------------------------------------------------------
# BDF yazma yardımcıları
# ---------------------------------------------------------------------------

_FW = 8  # Nastran small-field genişliği


def _f8i(val: int) -> str:
    return f"{val:>{_FW}d}"


def _f8f(val: float) -> str:
    s = f"{val:g}"
    if len(s) > _FW:
        s = f"{val:.5g}"
    return f"{s:>{_FW}}"


def _include_path(abs_path: str, use_relative: bool, output_dir: Path) -> str:
    if use_relative:
        try:
            return os.path.relpath(abs_path, output_dir)
        except ValueError:
            return abs_path
    return abs_path


def _is_thermal(type_name: str) -> bool:
    return _THERMAL_KEYWORD in type_name.upper()


# ---------------------------------------------------------------------------
# MSC Nastran BDF yazıcı (SUBCOM / SUBSEQ)
# ---------------------------------------------------------------------------

def write_bdf_msc(
    output_path: str,
    combinations: Combinations,
    bdf_map: Dict[str, Optional[str]],
    use_relative_paths: bool = False,
    combo_filter: Optional[List[str]] = None,
    start_subcase_sid: int = 1,
) -> Tuple[int, List[str]]:
    """
    MSC Nastran: result-level superposition.

    Yapı
      • Her unique unit case → ayrı SUBCASE
          - Termal case: TEMP(LOAD) = <case_id>
          - Diğerleri  : LOAD = <case_id>
      • Her kombinasyon → SUBCOM + SUBSEQ
          - SUBSEQ: tüm SUBCASE sırasına göre katsayı listesi (kullanılmayan → 0.0)
      • BEGIN BULK: INCLUDE satırları
    """
    output_dir = Path(output_path).parent
    missing: List[str] = []

    if combo_filter:
        fs = {c.strip() for c in combo_filter}
        combinations = {k: v for k, v in combinations.items() if k in fs}

    # Benzersiz unit case'leri sırayla topla: (tip, case_id) → subcase_id
    unit_sc: Dict[Tuple[str, str], int] = {}
    ordered_units: List[Tuple[str, str]] = []
    for entries in combinations.values():
        for type_name, case_id, _ in entries:
            key = (type_name, case_id)
            if key not in unit_sc:
                unit_sc[key] = start_subcase_sid + len(unit_sc)
                ordered_units.append(key)

    for _, case_id in ordered_units:
        if not bdf_map.get(case_id) and case_id not in missing:
            missing.append(case_id)

    lines: List[str] = []

    def c(text: str = ""):
        lines.append(f"$ {text}\n" if text else "$\n")

    # ── Başlık ──────────────────────────────────────────────────────────────
    c("=" * 70)
    c("  Combination BDF  –  MSC Nastran (SUBCOM/SUBSEQ)")
    c("  Üretildi: combination_bdf_builder.py")
    c("=" * 70)
    c()
    c("VARSAYIM: Her unit case BDF dosyasındaki yük seti SID'i = unit case ID'sidir.")
    c("  (Örn. 12001.bdf dosyası SID=12001 ile tanımlanmış FORCE/MOMENT/PLOAD içerir.)")
    c()
    lines.append("SOL 101\n")
    lines.append("CEND\n")
    c()

    # ── Unit Case SUBCASEs ───────────────────────────────────────────────────
    c("-" * 70)
    c("  Unit Case SUBCASEs")
    c("-" * 70)
    c()
    for (type_name, case_id), sc_id in unit_sc.items():
        try:
            cid_int = int(case_id)
        except ValueError:
            cid_int = case_id
        lines.append(f"SUBCASE {sc_id}\n")
        lines.append(f"  LABEL = {type_name}_{case_id}\n")
        if _is_thermal(type_name):
            lines.append(f"  TEMP(LOAD) = {cid_int}\n")
        else:
            lines.append(f"  LOAD = {cid_int}\n")
        c()

    # ── Combination SUBCOMs ─────────────────────────────────────────────────
    c("-" * 70)
    c("  Combination SUBCOMs")
    c("-" * 70)
    c()
    for combo_sid, entries in combinations.items():
        factor_map: Dict[Tuple[str, str], float] = {
            (t, cid): f for t, cid, f in entries
        }
        seq_vals = [factor_map.get(key, 0.0) for key in ordered_units]
        seq_strs = [f"{v:g}" for v in seq_vals]

        lines.append(f"SUBCOM {combo_sid}\n")
        lines.append(f"  LABEL = {combo_sid}\n")

        # SUBSEQ satırı – 72 karakter sınırı, virgüllü devam
        prefix = "  SUBSEQ = "
        cont   = "           "
        cur = prefix
        first = True
        for sv in seq_strs:
            candidate = cur + ("" if first else ", ") + sv
            if len(candidate) > 72 and not first:
                lines.append(cur + ",\n")
                cur = cont + sv
            else:
                cur = candidate
            first = False
        lines.append(cur + "\n")
        c()

    # ── BEGIN BULK ───────────────────────────────────────────────────────────
    lines.append("BEGIN BULK\n")
    c()
    c("-" * 70)
    c("  INCLUDE – Unit Case BDF Dosyaları")
    c("-" * 70)
    c()
    for type_name, case_id in ordered_units:
        bdf_path = bdf_map.get(case_id)
        if bdf_path:
            ip = _include_path(bdf_path, use_relative_paths, output_dir)
            c(f"[{type_name}] Case ID: {case_id}")
            lines.append(f"INCLUDE '{ip}'\n")
        else:
            c(f"WARNING: BDF bulunamadı  [{type_name}] Case ID: {case_id}")
        c()

    lines.append("ENDDATA\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    return len(combinations), missing


# ---------------------------------------------------------------------------
# NX Nastran BDF yazıcı (LOAD bulk kart)
# ---------------------------------------------------------------------------

def write_bdf_nx(
    output_path: str,
    combinations: Combinations,
    bdf_map: Dict[str, Optional[str]],
    use_relative_paths: bool = False,
    combo_filter: Optional[List[str]] = None,
) -> Tuple[int, List[str]]:
    """
    NX Nastran: load-level combination.

    Yapı
      • Case Control: her kombinasyon için tek SUBCASE
          - Mekanik case'ler: LOAD = <combo_sid>
          - İlk termal case : TEMP(LOAD) = <termal_case_id>
      • BEGIN BULK:
          - Tüm unique unit case'ler için INCLUDE
          - Her kombinasyon için LOAD bulk kart (yalnızca mekanik case'ler)

    LOAD kart formatı (Nastran small-field, 8 karakter/alan):
      LOAD    SID     S       S1      L1      S2      L2      S3      L3
      +               S4      L4      ...
    """
    output_dir = Path(output_path).parent
    missing: List[str] = []

    if combo_filter:
        fs = {c.strip() for c in combo_filter}
        combinations = {k: v for k, v in combinations.items() if k in fs}

    # Benzersiz unit case'ler: case_id → ilk görülen tip
    unique_cases: Dict[str, str] = {}
    for entries in combinations.values():
        for type_name, case_id, _ in entries:
            if case_id not in unique_cases:
                unique_cases[case_id] = type_name

    for case_id in unique_cases:
        if not bdf_map.get(case_id) and case_id not in missing:
            missing.append(case_id)

    lines: List[str] = []

    def c(text: str = ""):
        lines.append(f"$ {text}\n" if text else "$\n")

    # ── Başlık ──────────────────────────────────────────────────────────────
    c("=" * 70)
    c("  Combination BDF  –  NX Nastran (LOAD kart kombinasyonu)")
    c("  Üretildi: combination_bdf_builder.py")
    c("=" * 70)
    c()
    c("VARSAYIM: Her unit case BDF dosyasındaki yük seti SID'i = unit case ID'sidir.")
    c("  (Örn. 12001.bdf dosyası SID=12001 ile tanımlanmış FORCE/MOMENT/PLOAD içerir.)")
    c()
    lines.append("SOL 101\n")
    lines.append("CEND\n")
    c()

    # ── Case Control ─────────────────────────────────────────────────────────
    c("-" * 70)
    c("  Combination SUBCASEs")
    c("-" * 70)
    c()
    for combo_sid, entries in combinations.items():
        mech = [(t, cid, f) for t, cid, f in entries if not _is_thermal(t)]
        therm = [(t, cid, f) for t, cid, f in entries if _is_thermal(t)]

        try:
            sid_int = int(combo_sid)
        except ValueError:
            sid_int = combo_sid

        lines.append(f"SUBCASE {sid_int}\n")
        lines.append(f"  LABEL = {combo_sid}\n")
        if mech:
            lines.append(f"  LOAD = {sid_int}\n")
        if therm:
            t_sid = therm[0][1]
            try:
                t_sid_int = int(t_sid)
            except ValueError:
                t_sid_int = t_sid
            lines.append(f"  TEMP(LOAD) = {t_sid_int}\n")
            if len(therm) > 1:
                c(f"  UYARI: Birden fazla termal case var; yalnızca ilki ({t_sid}) TEMP(LOAD) olarak eklendi.")
        c()

    # ── BEGIN BULK ───────────────────────────────────────────────────────────
    lines.append("BEGIN BULK\n")
    c()
    c("-" * 70)
    c("  INCLUDE – Unit Case BDF Dosyaları")
    c("-" * 70)
    c()
    for case_id, type_name in unique_cases.items():
        bdf_path = bdf_map.get(case_id)
        if bdf_path:
            ip = _include_path(bdf_path, use_relative_paths, output_dir)
            c(f"[{type_name}] Case ID: {case_id}")
            lines.append(f"INCLUDE '{ip}'\n")
        else:
            c(f"WARNING: BDF bulunamadı  [{type_name}] Case ID: {case_id}")
        c()

    # ── LOAD kartları ────────────────────────────────────────────────────────
    c()
    c("-" * 70)
    c("  LOAD Kart Kombinasyonları (yalnızca mekanik case'ler)")
    c("-" * 70)
    c()
    for combo_sid, entries in combinations.items():
        mech = [(t, cid, f) for t, cid, f in entries if not _is_thermal(t)]
        if not mech:
            continue

        try:
            sid_int = int(combo_sid)
        except ValueError:
            sid_int = 0

        c(f"Kombinasyon: {combo_sid}")
        for t, cid, f in mech:
            c(f"  [{t}] {cid} x {f:g}")

        # Scale-load çiftlerini oluştur
        pair_fields: List[str] = []
        for _, case_id, factor in mech:
            try:
                cid_int = int(case_id)
            except ValueError:
                cid_int = 0
            pair_fields.append(_f8f(factor))
            pair_fields.append(_f8i(cid_int))

        # Satır 1: LOAD + SID + S(=1.0) + ilk 3 çift (6 alan)
        line1_data = pair_fields[:6]
        line1 = f"{'LOAD':<{_FW}}{_f8i(sid_int)}{_f8f(1.0)}" + "".join(line1_data)
        lines.append(line1.rstrip() + "\n")

        # Devam satırları: her satırda 4 çift (8 alan)
        rest = pair_fields[6:]
        while rest:
            chunk = rest[:8]
            rest = rest[8:]
            cont = f"{'+':<{_FW}}{'':>{_FW}}" + "".join(chunk)
            lines.append(cont.rstrip() + "\n")

        c()

    lines.append("ENDDATA\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    return len(combinations), missing


# ---------------------------------------------------------------------------
# Yönlendirici
# ---------------------------------------------------------------------------

def write_combination_bdf(
    output_path: str,
    combinations: Combinations,
    bdf_map: Dict[str, Optional[str]],
    solver: str = "nx",
    use_relative_paths: bool = False,
    combo_filter: Optional[List[str]] = None,
    start_subcase_sid: int = 1,
) -> Tuple[int, List[str]]:
    """Seçilen solver'a göre MSC veya NX BDF yazıcısına yönlendirir."""
    if solver == "msc":
        return write_bdf_msc(
            output_path, combinations, bdf_map,
            use_relative_paths, combo_filter, start_subcase_sid,
        )
    else:
        return write_bdf_nx(
            output_path, combinations, bdf_map,
            use_relative_paths, combo_filter,
        )


# ---------------------------------------------------------------------------
# Rapor
# ---------------------------------------------------------------------------

def print_report(
    excel_path: str,
    bdf_root: str,
    output_path: str,
    solver: str,
    combinations: Combinations,
    bdf_map: Dict[str, Optional[str]],
    missing: List[str],
    combos_written: int,
):
    total = len(bdf_map)
    found = sum(1 for v in bdf_map.values() if v)
    solver_label = "MSC Nastran (SUBCOM/SUBSEQ)" if solver == "msc" else "NX Nastran (LOAD kart)"

    print("\n" + "=" * 62)
    print("  Combination BDF Builder – Rapor")
    print("=" * 62)
    print(f"  Excel       : {excel_path}")
    print(f"  BDF kök     : {bdf_root}")
    print(f"  Çıktı       : {output_path}")
    print(f"  Solver      : {solver_label}")
    print(f"  Kombinasyon : {combos_written}")
    print(f"  Unit case   : {found} / {total} bulundu")

    if missing:
        print(f"\n  [!] {len(missing)} BDF dosyası BULUNAMADI:")
        for cid in missing:
            print(f"      - {cid}")
    else:
        print("\n  Tüm unit case BDF dosyaları başarıyla bulundu.")

    print("=" * 62)
    if not missing:
        print(f"\n  Çıktı yazıldı: {output_path}\n")


# ---------------------------------------------------------------------------
# GUI (tkinter)
# ---------------------------------------------------------------------------

def run_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext
    except ImportError:
        sys.exit("tkinter bu sistemde mevcut değil.")

    root = tk.Tk()
    root.title("Combination BDF Builder")
    root.resizable(True, True)

    pad = {"padx": 6, "pady": 3}

    # ── Değişkenler ──────────────────────────────────────────────────────────
    var_excel    = tk.StringVar()
    var_bdf_root = tk.StringVar()
    var_output   = tk.StringVar()
    var_solver   = tk.StringVar(value="nx")
    var_fmt      = tk.StringVar(value="paired")
    var_sheet    = tk.StringVar(value="0")
    var_combo_col  = tk.StringVar()
    var_case_col   = tk.StringVar()
    var_factor_col = tk.StringVar()
    var_rel_paths  = tk.BooleanVar(value=False)
    var_filter     = tk.StringVar()
    var_sc_sid     = tk.StringVar(value="1")

    # ── Yardımcılar ──────────────────────────────────────────────────────────
    def _row(parent, label, variable, browse_cmd=None, row=0):
        tk.Label(parent, text=label, anchor="w").grid(
            row=row, column=0, sticky="w", **pad
        )
        tk.Entry(parent, textvariable=variable, width=58).grid(
            row=row, column=1, sticky="ew", **pad
        )
        if browse_cmd:
            tk.Button(parent, text="Gözat…", command=browse_cmd, width=8).grid(
                row=row, column=2, **pad
            )
        return row + 1

    def browse_excel():
        p = filedialog.askopenfilename(
            title="Kombinasyon Excel'ini Seç",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Tümü", "*.*")],
        )
        if p:
            var_excel.set(p)

    def browse_root():
        p = filedialog.askdirectory(title="BDF Kök Dizinini Seç")
        if p:
            var_bdf_root.set(p)

    def browse_output():
        p = filedialog.asksaveasfilename(
            title="Çıktı BDF Dosyasını Kaydet",
            defaultextension=".bdf",
            filetypes=[("BDF", "*.bdf"), ("Tümü", "*.*")],
        )
        if p:
            var_output.set(p)

    # ── Ana çerçeve ──────────────────────────────────────────────────────────
    main = tk.Frame(root)
    main.pack(fill="both", expand=True, padx=10, pady=8)
    main.columnconfigure(1, weight=1)

    r = 0
    r = _row(main, "Kombinasyon Excel:", var_excel, browse_excel, r)
    r = _row(main, "Unit BDF Kök Dizin:", var_bdf_root, browse_root, r)
    r = _row(main, "Çıktı BDF Dosyası:", var_output, browse_output, r)

    # Solver
    tk.Label(main, text="Solver:", anchor="w").grid(row=r, column=0, sticky="w", **pad)
    sf = tk.Frame(main); sf.grid(row=r, column=1, sticky="w", **pad)
    tk.Radiobutton(sf, text="NX Nastran  (LOAD kart)", variable=var_solver, value="nx").pack(side="left")
    tk.Radiobutton(sf, text="MSC Nastran  (SUBCOM/SUBSEQ)", variable=var_solver, value="msc").pack(side="left", padx=14)
    r += 1

    # Format
    tk.Label(main, text="Excel Formatı:", anchor="w").grid(row=r, column=0, sticky="w", **pad)
    ff = tk.Frame(main); ff.grid(row=r, column=1, sticky="w", **pad)
    tk.Radiobutton(ff, text="Paired (önerilen)", variable=var_fmt, value="paired").pack(side="left")
    tk.Radiobutton(ff, text="Wide", variable=var_fmt, value="wide").pack(side="left", padx=6)
    tk.Radiobutton(ff, text="Long", variable=var_fmt, value="long").pack(side="left", padx=6)
    r += 1

    r = _row(main, "Sayfa (ad veya indeks):", var_sheet, None, r)

    tk.Label(main, text="[Yalnızca Long format]", fg="gray", anchor="w").grid(
        row=r, column=0, columnspan=3, sticky="w", padx=6
    )
    r += 1
    r = _row(main, "  Kombo sütunu:", var_combo_col, None, r)
    r = _row(main, "  Case ID sütunu:", var_case_col, None, r)
    r = _row(main, "  Katsayı sütunu:", var_factor_col, None, r)

    r = _row(main, "[MSC] Başlangıç SUBCASE SID:", var_sc_sid, None, r)

    tk.Checkbutton(
        main, text="INCLUDE yollarını göreli yaz", variable=var_rel_paths
    ).grid(row=r, column=0, columnspan=2, sticky="w", **pad)
    r += 1

    r = _row(main, "Filtrele (virgülle ayrılmış, boş=hepsi):", var_filter, None, r)

    # Log alanı
    log = scrolledtext.ScrolledText(main, height=14, state="disabled", wrap="word")
    log.grid(row=r, column=0, columnspan=3, sticky="nsew", **pad)
    main.rowconfigure(r, weight=1)
    r += 1

    def _log(msg: str):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    # Çalıştır
    def run():
        log.configure(state="normal"); log.delete("1.0", "end"); log.configure(state="disabled")
        excel   = var_excel.get().strip()
        bdf_root = var_bdf_root.get().strip()
        output  = var_output.get().strip()
        solver  = var_solver.get()

        if not excel or not bdf_root or not output:
            messagebox.showerror("Hata", "Excel, BDF kök dizini ve çıktı yolu zorunludur.")
            return

        sheet_raw = var_sheet.get().strip()
        try:
            sheet = int(sheet_raw)
        except ValueError:
            sheet = sheet_raw

        fmt    = var_fmt.get()
        filter_raw = var_filter.get().strip()
        combo_filter = [x.strip() for x in filter_raw.split(",")] if filter_raw else None
        try:
            sc_sid = int(var_sc_sid.get().strip())
        except ValueError:
            sc_sid = 1

        try:
            _log("Excel okunuyor…")
            combs = load_combinations(
                excel, fmt, sheet,
                var_combo_col.get().strip(),
                var_case_col.get().strip(),
                var_factor_col.get().strip(),
            )
            _log(f"  → {len(combs)} kombinasyon yüklendi.")

            unique_ids = sorted({cid for entries in combs.values() for _, cid, _ in entries})
            _log(f"  → {len(unique_ids)} benzersiz unit case ID'si bulundu.")

            _log("BDF dosyaları aranıyor…")
            bdf_map = find_bdf_files(bdf_root, unique_ids)
            found = sum(1 for v in bdf_map.values() if v)
            _log(f"  → {found} / {len(unique_ids)} BDF bulundu.")

            _log(f"Çıktı BDF yazılıyor ({solver.upper()})…")
            n, missing = write_combination_bdf(
                output, combs, bdf_map, solver, var_rel_paths.get(), combo_filter, sc_sid
            )
            _log(f"  → {n} kombinasyon yazıldı.")

            if missing:
                _log(f"\n[!] {len(missing)} BDF bulunamadı:")
                for cid in missing:
                    _log(f"    - {cid}")
            else:
                _log("\nTüm unit case BDF dosyaları bulundu.")

            _log(f"\nTamamlandı!  Çıktı: {output}")
            messagebox.showinfo("Tamamlandı", f"Çıktı yazıldı:\n{output}")

        except Exception as exc:
            _log(f"\nHATA: {exc}")
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
    p.add_argument("--gui", action="store_true", help="Grafik arayüzü aç")
    p.add_argument("--excel",    help="Kombinasyon Excel dosyası yolu")
    p.add_argument("--bdf-root", help="Unit case BDF dosyalarının kök dizini")
    p.add_argument("--output",   help="Çıktı BDF dosyası yolu")
    p.add_argument(
        "--solver", choices=["nx", "msc"], default="nx",
        help="Nastran solver: 'nx' (LOAD kart, varsayılan) veya 'msc' (SUBCOM/SUBSEQ)",
    )
    p.add_argument(
        "--format", choices=["paired", "wide", "long"], default="paired",
        help="Excel düzeni: 'paired' (varsayılan), 'wide' veya 'long'",
    )
    p.add_argument("--sheet", default="0",
        help="Excel sayfa adı veya 0 tabanlı indeks (varsayılan: 0)")
    p.add_argument("--combo-col",  default="", help="[Long] Kombinasyon adı sütun başlığı")
    p.add_argument("--case-col",   default="", help="[Long] Unit case ID sütun başlığı")
    p.add_argument("--factor-col", default="", help="[Long] Katsayı sütun başlığı")
    p.add_argument("--relative-paths", action="store_true",
        help="INCLUDE yollarını çıktı BDF dizinine göreli yaz")
    p.add_argument("--combo-filter", nargs="*",
        help="Yalnızca bu kombinasyonları işle (boşlukla ayrılmış)")
    p.add_argument("--start-subcase-sid", type=int, default=1,
        help="[MSC] Unit SUBCASE başlangıç SID'i (varsayılan: 1)")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.gui:
        run_gui()
        return

    errors = []
    if not args.excel:    errors.append("--excel gerekli")
    if not args.bdf_root: errors.append("--bdf-root gerekli")
    if not args.output:   errors.append("--output gerekli")
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

    print(f"Excel okunuyor: {args.excel}")
    combinations = load_combinations(
        args.excel, args.format, sheet,
        args.combo_col, args.case_col, args.factor_col,
    )
    print(f"  → {len(combinations)} kombinasyon yüklendi.")

    unique_ids = sorted({cid for entries in combinations.values() for _, cid, _ in entries})
    preview = unique_ids[:8]
    print(f"  → {len(unique_ids)} unique unit case ID: {preview}{'…' if len(unique_ids) > 8 else ''}")

    print(f"\nBDF dosyaları aranıyor: {args.bdf_root}")
    bdf_map = find_bdf_files(args.bdf_root, unique_ids)
    found = sum(1 for v in bdf_map.values() if v)
    print(f"  → {found} / {len(unique_ids)} bulundu.")

    solver_label = "MSC Nastran (SUBCOM/SUBSEQ)" if args.solver == "msc" else "NX Nastran (LOAD kart)"
    print(f"\nÇıktı BDF yazılıyor [{solver_label}]: {args.output}")
    combos_written, missing = write_combination_bdf(
        args.output, combinations, bdf_map,
        args.solver, args.relative_paths,
        args.combo_filter, args.start_subcase_sid,
    )

    print_report(
        args.excel, args.bdf_root, args.output,
        args.solver, combinations, bdf_map, missing, combos_written,
    )

    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
