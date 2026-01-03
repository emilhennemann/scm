from __future__ import annotations

import os
import re
import time
import json
import random
import string
import builtins
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Dict, Tuple, List

import requests
import pandas as pd
import win32com.client as win32
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

@dataclass(frozen=True)
class Config:
    base_dir: Path = field(default_factory=lambda: Path(__file__).resolve().parent)
    root: Path = field(init=False)

    folder_in: Path = field(init=False)
    folder_final: Path = field(init=False)
    folder_nutz: Path = field(init=False)

    file_suppliers: Path = field(init=False)
    file_erp: Path = field(init=False)
    file_scale: Path = field(init=False)
    file_nutz_template: Path = field(init=False)

    form_server: str = field(default_factory=lambda: os.getenv("SCM_FORM_SERVER", "http://localhost:8000"))
    send_to_final: str = "anouar97@gmx.de"

    keep_original_text_fields: frozenset[str] = frozenset({"co2-emissionen", "zahlungsbedingungen"})

    def __post_init__(self):
        object.__setattr__(self, "root", self.base_dir / "ROOT")
        object.__setattr__(self, "folder_in", self.root / "Antworten_Erhalt")
        object.__setattr__(self, "folder_final", self.root / "Einzelberichte_Lieferanten")
        object.__setattr__(self, "folder_nutz", self.root / "Nutzwertanalyse")

        object.__setattr__(self, "file_suppliers", self.root / "1. SCM-Anwendungen(MA)_Lieferantenuebersicht.xlsx")
        object.__setattr__(self, "file_erp", self.root / "4. SCM-Anwendungen(MA)_ERP-System.xlsx")
        object.__setattr__(self, "file_scale", self.root / "3. SCM-Anwendungen(MA)_Gesamtbewertung.xlsx")
        object.__setattr__(self, "file_nutz_template", self.root / "5. SCM-Nutzwertanalyse.xlsx")

    def ensure_dirs(self) -> None:
        for d in (self.folder_in, self.folder_final, self.folder_nutz):
            d.mkdir(parents=True, exist_ok=True)


# ============================================================
# STATE
# ============================================================

@dataclass
class RoundState:
    round_id: str
    state_file: Path
    sent: Dict[str, dict] = field(default_factory=dict)       # supplier_id -> meta
    responses: Dict[str, dict] = field(default_factory=dict)  # supplier_id -> meta
    nutzwert_done: bool = False
    final_mail_sent: bool = False

    @staticmethod
    def new(config: Config) -> "RoundState":
        rid = "".join(random.choices(string.digits, k=8))
        return RoundState(round_id=rid, state_file=config.root / f"round_state_{rid}.json")

    @staticmethod
    def load_or_new(config: Config) -> "RoundState":
        tmp = RoundState.new(config)
        if tmp.state_file.exists():
            try:
                raw = json.loads(tmp.state_file.read_text(encoding="utf-8"))
                st = RoundState(
                    round_id=raw.get("round_id", tmp.round_id),
                    state_file=tmp.state_file,
                    sent=raw.get("sent", {}),
                    responses=raw.get("responses", {}),
                    nutzwert_done=bool(raw.get("nutzwert_done", False)),
                    final_mail_sent=bool(raw.get("final_mail_sent", False)),
                )
                return st
            except Exception:
                pass
        return tmp

    def save(self) -> None:
        payload = {
            "round_id": self.round_id,
            "sent": self.sent,
            "responses": self.responses,
            "nutzwert_done": self.nutzwert_done,
            "final_mail_sent": self.final_mail_sent,
        }
        self.state_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def status_line(self) -> str:
        total = len(self.sent)
        got = len(self.responses)
        return f"[STATUS] {got} von {total} Antworten"

    def all_done(self) -> bool:
        total = len(self.sent)
        got = len(self.responses)
        return total > 0 and got >= total


# ============================================================
# TEXT / PARSING / SCORING (pure-ish functions)
# ============================================================

def norm_text(s: Any) -> str:
    s = "" if s is None else builtins.str(s)
    s = s.replace("\u00A0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def safe_str(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return builtins.str(x).strip()

def extract_first_number(val: Any) -> Optional[float]:
    """Extract first numeric token from text (handles 11 000, 11.000, 98,5, kg CO2e, % etc.)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = builtins.str(val)
    s = s.replace("\u00A0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ").lower()
    s = s.replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "").replace("%", "").strip()

    s_compact = s.replace(" ", "")
    m_th = re.search(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", s_compact)
    if m_th:
        token = m_th.group().replace(".", "").replace(",", ".")
        try:
            return float(token)
        except Exception:
            pass

    s2 = s.replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.\d+|\d+", s2)
    if not m:
        return None
    try:
        return float(m.group())
    except Exception:
        return None

def normalize_percent_if_needed(x: float) -> float:
    if 0 <= x <= 1:
        return round(x * 100, 2)
    return round(float(x), 2)

def parse_scale_condition(scale_text: str, erp_numeric: Optional[float]) -> bool:
    if erp_numeric is None:
        return False
    t = norm_text(scale_text)
    if not t or t in ("nan", "none"):
        return False

    raw = (
        t.replace("%", "")
         .replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "")
         .replace(" ", "")
         .replace(",", ".")
    )

    try:
        ops = [("<=", lambda a, b: a <= b), (">=", lambda a, b: a >= b),
               ("≤",  lambda a, b: a <= b), ("≥",  lambda a, b: a >= b),
               ("<",  lambda a, b: a < b),  (">",  lambda a, b: a > b)]
        for op, fn in ops:
            if op in raw:
                n = extract_first_number(raw.split(op, 1)[1])
                return n is not None and fn(erp_numeric, n)

        if "–" in t or "-" in t:
            parts = re.split(r"[–-]", t)
            if len(parts) >= 2:
                low = extract_first_number(parts[0])
                high = extract_first_number(parts[1])
                return low is not None and high is not None and low <= erp_numeric <= high
    except Exception:
        return False

    return False

def is_negative_presence_value(val: Any) -> bool:
    s = norm_text(val)
    if s in ("", "nan", "none"):
        return True
    negatives = ("nicht vorhanden", "nichtvorhanden", "nein", "no", "false", "0", "keine", "kein")
    return any(n in s for n in negatives)

def find_matching_criterion(crit_name: str, scale_data: dict) -> Optional[str]:
    c_clean = builtins.str(crit_name).strip().lower()
    if crit_name in scale_data:
        return crit_name
    for k in scale_data:
        kk = builtins.str(k).strip().lower()
        if kk == c_clean or c_clean in kk or kk in c_clean:
            return k
    return None

def map_erp_to_points(kriterium: str, erp_value: Any, scale_data: dict) -> int:
    actual = find_matching_criterion(kriterium, scale_data)
    if not actual:
        return 0

    crit_l = norm_text(actual)

    # ISO Fix
    if crit_l.startswith("iso ") or "iso " in crit_l:
        if is_negative_presence_value(erp_value):
            return 0

    # numeric path
    num = extract_first_number(erp_value)
    if num is not None:
        num = normalize_percent_if_needed(num)
        for pts in (100, 80, 60, 40, 20, 0):
            st = scale_data[actual].get(pts, "")
            if parse_scale_condition(st, num):
                return pts

    # text path
    val_str = norm_text(erp_value)
    for pts in (100, 80, 60, 40, 20, 0):
        st = norm_text(scale_data[actual].get(pts, ""))
        if not st or st in ("nan", "none"):
            continue

        if val_str and (val_str in st or st in val_str):
            return pts

        # CoC Spezial
        if "code of conduct" in crit_l or crit_l in ("coc",) or "coc" in crit_l:
            if pts == 100 and (("bme" in val_str and "bme" in st) or ("kb" in val_str and "kb" in st)):
                return 100

    return 0


# ============================================================
# SCALE LOADER
# ============================================================

def get_comprehensive_scale(file_scale: Path) -> dict:
    df = pd.read_excel(file_scale, sheet_name="Skala", header=None)
    scale_map = {}
    for i in range(4, len(df)):
        crit = builtins.str(df.iloc[i, 0]).strip()
        if not crit or crit.lower() in ("nan", "none"):
            continue
        scale_map[crit] = {
            0: builtins.str(df.iloc[i, 4]),
            20: builtins.str(df.iloc[i, 5]),
            40: builtins.str(df.iloc[i, 6]),
            60: builtins.str(df.iloc[i, 7]),
            80: builtins.str(df.iloc[i, 8]),
            100: builtins.str(df.iloc[i, 9]),
        }
    return scale_map


# ============================================================
# SERVER API
# ============================================================

@dataclass(frozen=True)
class FormAPI:
    base_url: str

    def get_form_link(self, supplier_id: str, round_id: str) -> str:
        r = requests.get(f"{self.base_url}/issue-link", params={"supplier_id": supplier_id, "round_id": round_id}, timeout=10)
        r.raise_for_status()
        return r.json()["url"]

    def list_submissions(self, round_id: str) -> list[dict]:
        r = requests.get(f"{self.base_url}/api/submissions", params={"round_id": round_id}, timeout=10)
        r.raise_for_status()
        return r.json()

    def download_submission_xlsx(self, round_id: str, supplier_id: str) -> bytes:
        r = requests.get(f"{self.base_url}/api/xlsx", params={"round_id": round_id, "supplier_id": supplier_id}, timeout=30)
        r.raise_for_status()
        return r.content


# ============================================================
# OUTLOOK UI (Playwright)
# ============================================================

class OutlookUI:
    def __init__(self, page):
        self.page = page

    def open_mail(self) -> None:
        self.page.goto("https://outlook.office.com/mail/")
        self.page.wait_for_selector('button[aria-label*="Neue"]', timeout=60000)

    def new_mail(self, to_email: str, subject: str, body: str) -> None:
        self.page.click('button[aria-label*="Neue"]')
        self.page.wait_for_timeout(500)
        self.page.fill('div[aria-label="An"]', to_email)
        self.page.fill('input[placeholder*="Betreff"]', subject)
        self.page.locator('div[role="textbox"]').first.click()
        self.page.keyboard.type(body)
        self.page.click('button[aria-label*="Senden"]')
        self.page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=30000)

    def new_mail_with_attachment(self, to_email: str, subject: str, body: str, attachment_path: Path) -> None:
        self.page.click('button[aria-label*="Neue"]')
        self.page.wait_for_timeout(500)
        self.page.fill('div[aria-label="An"]', to_email)
        self.page.fill('input[placeholder*="Betreff"]', subject)
        self.page.locator('div[role="textbox"]').first.click()
        self.page.keyboard.type(body)

        self.page.locator('button[aria-label*="Datei anfügen"]').first.click()
        with self.page.expect_file_chooser() as fc:
            self.page.locator('button[aria-label*="Diesen Computer durchsuchen"]').first.click()
        fc.value.set_files(builtins.str(attachment_path.resolve()))

        self.page.wait_for_timeout(1200)
        self.page.click('button[aria-label*="Senden"]')
        self.page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=30000)


# ============================================================
# EXCEL COM (Nutzwertanalyse)
# ============================================================

def excel_set_value_safe(ws, row: int, col: int, value: Any) -> None:
    cell = ws.Cells(row, col)
    try:
        if cell.MergeCells:
            cell = cell.MergeArea.Cells(1, 1)
    except Exception:
        pass
    cell.Value = value

def excel_find_rows(ws, start_row=3, max_scan_rows=400, template_nutzwert_col=4) -> Tuple[List[int], Optional[int]]:
    criteria_rows: List[int] = []
    sum_row = None
    for r in range(start_row, start_row + max_scan_rows):
        a_val = ws.Cells(r, 1).Value
        if a_val and "summe nutzwerte" in builtins.str(a_val).strip().lower():
            sum_row = r

        tmpl = ws.Cells(r, template_nutzwert_col).Formula
        if tmpl and isinstance(tmpl, builtins.str) and tmpl.startswith("="):
            criteria_rows.append(r)

    return criteria_rows, sum_row

def read_supplier_report(report_path: Path) -> dict:
    df = pd.read_excel(report_path)
    out: Dict[str, int] = {}
    for _, r in df.iterrows():
        crit = safe_str(r.get("Kriterium"))
        if not crit or norm_text(crit) in ("nan", "none"):
            continue
        try:
            out[crit] = int(r.get("Skalapunkte", 0))
        except Exception:
            out[crit] = 0
    return out

def match_points(rep_map: dict, template_crit: Any) -> int:
    if template_crit is None:
        return 0
    template_crit_s = builtins.str(template_crit).strip()
    if template_crit_s in rep_map:
        return int(rep_map[template_crit_s])

    t = template_crit_s.lower()
    for k, v in rep_map.items():
        kk = builtins.str(k).strip().lower()
        if kk == t or t in kk or kk in t:
            return int(v)
    return 0

def excel_find_supplier_column(ws, supplier_name: str, header_row=1, start_col=3, max_cols=400) -> Tuple[Optional[int], Optional[int]]:
    target = builtins.str(supplier_name).strip().lower()
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v and builtins.str(v).strip().lower() == target:
            return c, c + 1
        c += 2
    return None, None

def excel_next_free_supplier_column(ws, header_row=1, start_col=3, max_cols=400) -> Tuple[int, int]:
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v is None or builtins.str(v).strip() == "":
            return c, c + 1
        c += 2
    raise RuntimeError("Keine freie Spalte mehr in Nutzwertanalyse (max_cols erreicht).")

def excel_last_used_row(ws, min_row: int = 2, max_row: int = 800) -> int:
    """Heuristik: letzte relevante Zeile (Kriterium/Section/Summe) finden."""
    last = min_row
    for r in range(min_row, max_row + 1):
        a = ws.Cells(r, 1).Value
        d_formula = ws.Cells(r, 4).Formula  # Template hat i.d.R. hier Formeln
        if (a is not None and str(a).strip() != "") or (d_formula is not None and str(d_formula).strip().startswith("=")):
            last = r
    return last


def excel_apply_template_pair_formats(
    ws,
    dest_bew_col: int,
    dest_nutz_col: int,
    template_bew_col: int = 3,
    template_nutz_col: int = 4,
    header_row_1: int = 1,
    header_row_2: int = 2,
) -> None:
    """
    Kopiert exakt das Layout/Design aus dem Template-Lieferantenpaar (C:D) auf ein neues Paar.
    Dadurch: graue Zeilen + dicke Rahmen + Kopfzeilen-Merge wie im Screenshot.
    """
    last_row = excel_last_used_row(ws)

    src = ws.Range(ws.Cells(1, template_bew_col), ws.Cells(last_row, template_nutz_col))
    dst = ws.Range(ws.Cells(1, dest_bew_col), ws.Cells(last_row, dest_nutz_col))

    # Optional: Fokus hilft manchmal bei PasteSpecial
    try:
        ws.Activate()
    except Exception:
        pass

    # Nur Formate kopieren
    src.Copy()
    dst.PasteSpecial(Paste=-4122)  # xlPasteFormats

    # Kopfzeile: Lieferant über 2 Spalten (Merge)
    try:
        hdr = ws.Range(ws.Cells(header_row_1, dest_bew_col), ws.Cells(header_row_1, dest_nutz_col))
        if not hdr.MergeCells:
            hdr.Merge()
        hdr.Font.Bold = True
        hdr.HorizontalAlignment = -4108  # xlCenter
        hdr.VerticalAlignment = -4108
        hdr.WrapText = True
    except Exception:
        pass

    # Subheader in Zeile 2
    try:
        r2 = ws.Range(ws.Cells(header_row_2, dest_bew_col), ws.Cells(header_row_2, dest_nutz_col))
        r2.Font.Bold = True
        r2.HorizontalAlignment = -4108
        r2.VerticalAlignment = -4108
        r2.WrapText = True
    except Exception:
        pass

    # Clipboard leeren
    try:
        ws.Application.CutCopyMode = False
    except Exception:
        pass

def excel_create_nutzwert_chart(ws_nutzwert, wb):
    """
    Erstellt ein zweites Blatt mit Balkendiagramm (Nutzwerte je Lieferant)
    """

    # --- Diagrammblatt neu anlegen ---
    chart_name = "Diagramm"
    try:
        wb.Worksheets(chart_name).Delete()
    except Exception:
        pass

    ws_chart = wb.Worksheets.Add()
    ws_chart.Name = chart_name

    # --- Lieferanten + Nutzwerte sammeln ---
    header_row = 1
    sum_row = None

    # Summe-Nutzwerte-Zeile finden
    for r in range(3, 600):
        val = ws_nutzwert.Cells(r, 1).Value
        if val and "summe nutzwerte" in str(val).lower():
            sum_row = r
            break

    if not sum_row:
        raise RuntimeError("Summe Nutzwerte nicht gefunden")

    data = []
    col = 3  # erste Lieferantenspalte
    while True:
        supplier = ws_nutzwert.Cells(header_row, col).Value
        if not supplier:
            break

        nutzwert = ws_nutzwert.Cells(sum_row, col + 1).Value
        data.append((supplier, nutzwert))
        col += 2

    # --- Hilfstabelle schreiben ---
    ws_chart.Cells(1, 1).Value = "Lieferant"
    ws_chart.Cells(1, 2).Value = "Gesamtnutzwert"

    for i, (name, value) in enumerate(data, start=2):
        ws_chart.Cells(i, 1).Value = name
        ws_chart.Cells(i, 2).Value = value

    last_row = len(data) + 1

    # --- Diagramm erzeugen ---
    chart = wb.Charts.Add()
    chart.ChartType = 51  # xlColumnClustered (Balkendiagramm)

    chart.SetSourceData(
        ws_chart.Range(ws_chart.Cells(1, 1), ws_chart.Cells(last_row, 2))
    )

    chart.HasTitle = True
    chart.ChartTitle.Text = "Gesamtnutzwerte je Lieferant"

    chart.Axes(1).HasTitle = True
    chart.Axes(1).AxisTitle.Text = "Lieferanten"

    chart.Axes(2).HasTitle = True
    chart.Axes(2).AxisTitle.Text = "Nutzwert"

    chart.Location(Where=2, Name=chart_name)  # xlLocationAsObject

class ExcelApp:
    def __enter__(self):
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        try:
            self.excel.ScreenUpdating = False
        except Exception:
            pass
        try:
            self.excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass
        return self.excel

    def __exit__(self, exc_type, exc, tb):
        try:
            self.excel.Quit()
        except Exception:
            pass

def open_or_create_nutzwert(excel, out_path: Path, template_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        wb = excel.Workbooks.Open(builtins.str(out_path.resolve()))
    else:
        wb = excel.Workbooks.Open(builtins.str(template_path.resolve()))
        wb.SaveAs(builtins.str(out_path.resolve()))
    return wb, wb.Worksheets(1)

def upsert_supplier_into_nutzwert(
    excel,
    wb,
    ws,
    report_path: Path,
    supplier_name: str,
    start_col: int = 3,
) -> None:
    HEADER_ROW_1, HEADER_ROW_2 = 1, 2
    KRIT_COL, W_COL = 1, 2

    criteria_rows, sum_row = excel_find_rows(ws, start_row=3, max_scan_rows=600, template_nutzwert_col=4)
    if not criteria_rows:
        raise RuntimeError("Keine Kriterienzeilen im Template erkannt (Formeln fehlen?).")

    col_bew, col_nutz = excel_find_supplier_column(ws, supplier_name, header_row=HEADER_ROW_1, start_col=start_col)
    if col_bew is None:
        col_bew, col_nutz = excel_next_free_supplier_column(ws, header_row=HEADER_ROW_1, start_col=start_col)

        # Exakt den Stil wie im Template (Lieferant 1) übernehmen
        excel_apply_template_pair_formats(ws, col_bew, col_nutz)

        # Header setzen
        excel_set_value_safe(ws, HEADER_ROW_1, col_bew, supplier_name)
        excel_set_value_safe(ws, HEADER_ROW_2, col_bew, "Bewertung")
        excel_set_value_safe(ws, HEADER_ROW_2, col_nutz, "Nutzwert")

    rep_map = read_supplier_report(report_path)

    for r in criteria_rows:
        crit_txt = ws.Cells(r, KRIT_COL).Value
        pts = match_points(rep_map, crit_txt)
        ws.Cells(r, col_bew).Value = int(pts)

        w_cell = builtins.str(ws.Cells(r, W_COL).Address).replace("$", "")
        b_cell = builtins.str(ws.Cells(r, col_bew).Address).replace("$", "")
        ws.Cells(r, col_nutz).Formula = f"={w_cell}*{b_cell}"

    if sum_row:
        first = builtins.str(ws.Cells(criteria_rows[0], col_nutz).Address).replace("$", "")
        last = builtins.str(ws.Cells(criteria_rows[-1], col_nutz).Address).replace("$", "")
        ws.Cells(sum_row, col_nutz).Formula = f"=SUM({first}:{last})"

    try:
        wb.RefreshAll()
    except Exception:
        pass
    try:
        excel.CalculateFull()
    except Exception:
        pass


# ============================================================
# DOMAIN: Reports + Pipeline
# ============================================================

def load_suppliers(config: Config) -> pd.DataFrame:
    return pd.read_excel(config.file_suppliers, sheet_name="Lieferanten", header=2).dropna(subset=["Lieferant_Name"])

def supplier_name_for_id(config: Config, supplier_id: str) -> Optional[str]:
    df = load_suppliers(config).copy()
    df["Lieferant_ID_norm"] = df["Lieferant_ID"].astype(builtins.str).str.strip()
    match = df[df["Lieferant_ID_norm"] == supplier_id]
    if match.empty:
        return None
    return match["Lieferant_Name"].values[0]

def load_erp_dict(config: Config, supplier_name: str) -> dict:
    df_erp = pd.read_excel(config.file_erp, sheet_name=supplier_name, header=None)
    return dict(zip(df_erp[0][1:], df_erp[1][1:]))

def make_display_value(val: Any, crit_key_norm: str, keep_original: frozenset[str]) -> str:
    # Default: original
    display_val: Any = val

    num = extract_first_number(val)
    if num is not None:
        num2 = normalize_percent_if_needed(num)
        if isinstance(val, (int, float)) and 0 <= float(val) <= 1:
            display_val = f"{num2:.2f}%"
        elif "%" in safe_str(val):
            display_val = f"{num2:.2f}%"
        else:
            display_val = f"{num2:.2f}"

    if crit_key_norm in keep_original:
        return safe_str(val)

    return safe_str(display_val)

def process_supplier_from_xlsx(
    config: Config,
    scale_data: dict,
    round_id: str,
    supplier_id: str,
    file_path: Path,
) -> Optional[Path]:
    try:
        df_man = pd.read_excel(file_path)
        if df_man is None or df_man.empty:
            print(" [!] Manuelle Antwortdatei leer.")
            return None

        supplier_name = supplier_name_for_id(config, supplier_id)
        if not supplier_name:
            print(f" [!] Lieferant_ID {supplier_id} nicht in Lieferantenliste gefunden.")
            return None

        erp_dict = load_erp_dict(config, supplier_name)

        # manuelle Kriterien
        val_col_candidates = [c for c in df_man.columns if "bewertung" in builtins.str(c).lower()]
        if not val_col_candidates:
            print(" [!] Konnte Bewertungsspalte im Formular-Export nicht finden.")
            return None
        val_col = val_col_candidates[0]

        final_rows: list[dict] = []

        for _, row in df_man.iterrows():
            crit = safe_str(row.get("Kriterium"))
            if not crit or norm_text(crit) in ("nan", "none"):
                continue

            try:
                pts = int(row.get(val_col))
            except Exception:
                pts = 0

            actual = find_matching_criterion(crit, scale_data)
            desc = scale_data.get(actual, {}).get(pts, builtins.str(pts))
            final_rows.append({"Kriterium": crit, "Wert": desc, "Skalapunkte": pts})

        # ERP Kriterien
        for crit, val in erp_dict.items():
            crit_s = safe_str(crit)
            if not crit_s or norm_text(crit_s) in ("nan", "none"):
                continue

            pts = map_erp_to_points(crit_s, val, scale_data)
            display_val = make_display_value(val, norm_text(crit_s), config.keep_original_text_fields)
            final_rows.append({"Kriterium": crit_s, "Wert": display_val, "Skalapunkte": pts})

        out = config.folder_final / f"Bericht_{supplier_name}_R{round_id}.xlsx"
        pd.DataFrame(final_rows, columns=["Kriterium", "Wert", "Skalapunkte"]).to_excel(out, index=False)
        print(f" [FINISH] Bericht erstellt/aktualisiert: {out.name}")
        return out

    except Exception as e:
        print(f" [!] Fehler bei Bericht: {e}")
        traceback.print_exc()
        return None


# ============================================================
# PHASES
# ============================================================

def phase_dispatch_links(config: Config, api: FormAPI, outlook: OutlookUI, state: RoundState) -> None:
    print(f"\n[PHASE 1] Versand Runde {state.round_id} gestartet...")
    df_supp = load_suppliers(config)

    outlook.open_mail()

    for _, row in df_supp.iterrows():
        try:
            s_id = safe_str(row.get("Lieferant_ID"))
            email = safe_str(row.get("Email"))
            name = safe_str(row.get("Name"))
            lname = safe_str(row.get("Lieferant_Name"))

            if not s_id or not email:
                continue

            form_url = api.get_form_link(s_id, state.round_id)
            subject = f"SCM-Bewertung | Runde {state.round_id} | {s_id}"
            body = (
                f"Hallo {name},\n\n"
                f"im Rahmen unseres aktuellen Lieferantenbewertungsprozesses bitten wir Sie, die Bewertung über den folgenden Link auszufüllen:\n\n"
                f"{form_url}\n\n"
                f"Bitte füllen Sie die Bewertung vollständig und sorgfältig aus. Die Bearbeitung dauert in der Regel nur wenige Minuten.\n\n"
                f"Vielen Dank für Ihre Unterstützung.\n\n"
                f"Freundliche Grüße\n"
                f"Ihr SCM-Team\n\n"
                f"(Runde {state.round_id})"
            )
            outlook.new_mail(email, subject, body)

            state.sent[s_id] = {"name": lname, "email": email, "sent_at": time.time()}
            state.save()

            print(f" [OK] Link an {lname} ({s_id})")
            print(state.status_line())
        except Exception:
            # UI Recovery: Compose schließen
            try:
                outlook.page.keyboard.press("Escape")
            except Exception:
                pass

def phase_poll_and_process(
    config: Config,
    api: FormAPI,
    scale_data: dict,
    state: RoundState,
) -> None:
    print(f"\n[PHASE 2] Server-Polling aktiv (Runde {state.round_id})...")

    while True:
        try:
            if state.all_done():
                print(f"\n[OK] Alle Antworten verarbeitet. {state.status_line()}")
                break

            subs = api.list_submissions(state.round_id)
            for item in subs:
                sid = safe_str(item.get("supplier_id"))
                submitted_at = float(item.get("submitted_at", 0))

                if sid not in state.sent:
                    continue

                prev = state.responses.get(sid)
                if prev and submitted_at <= float(prev.get("submitted_at", 0)):
                    continue

                xlsx_bytes = api.download_submission_xlsx(state.round_id, sid)
                filename = f"Antwort_{sid}_R{state.round_id}_{int(submitted_at)}.xlsx"
                file_path = config.folder_in / filename
                file_path.write_bytes(xlsx_bytes)

                report_path = process_supplier_from_xlsx(config, scale_data, state.round_id, sid, file_path)
                if report_path:
                    # direkt inkrementell in Nutzwertanalyse updaten
                    with ExcelApp() as excel:
                        out_path = (config.folder_nutz / f"Nutzwertanalyse_R{state.round_id}.xlsx").resolve()
                        wb, ws = open_or_create_nutzwert(excel, out_path, config.file_nutz_template)
                        try:
                            supplier_name = supplier_name_for_id(config, sid)
                            if supplier_name:
                                upsert_supplier_into_nutzwert(excel, wb, ws, report_path, supplier_name)
                                wb.Save()
                                print(f"[OK] Nutzwertanalyse inkrementell aktualisiert: {out_path}")
                        finally:
                            try:
                                wb.Close(SaveChanges=True)
                            except Exception:
                                pass

                    state.responses[sid] = {"submitted_at": submitted_at, "filename": filename}
                    state.save()
                    print(state.status_line())

            time.sleep(6)
        except Exception as e:
            print(f"[WARN] Polling-Fehler: {e}")
            time.sleep(6)

def phase_finalize_nutzwert(config: Config, state: RoundState) -> Optional[Path]:
    # Wenn wir inkrementell updaten, ist "build" optional.
    out_path = (config.folder_nutz / f"Nutzwertanalyse_R{state.round_id}.xlsx").resolve()
    if out_path.exists():
        state.nutzwert_done = True
        state.save()
        return out_path

    # Fallback: falls noch nicht existiert, aus Einzelberichten bauen
    report_files = sorted(config.folder_final.glob(f"Bericht_*_R{state.round_id}.xlsx"))
    if not report_files:
        print("[!] Keine Einzelberichte gefunden – Nutzwertanalyse wird nicht erstellt.")
        return None

    if not config.file_nutz_template.exists():
        print(f"[!] Nutzwert-Template fehlt: {config.file_nutz_template}")
        return None

    suppliers: list[Tuple[str, Path]] = []
    for p in report_files:
        m = re.match(r"Bericht_(.*)_R(\d+)\.xlsx", p.name, flags=re.IGNORECASE)
        suppliers.append((m.group(1) if m else p.stem, p))

    with ExcelApp() as excel:
        wb = excel.Workbooks.Open(builtins.str(config.file_nutz_template.resolve()))
        ws = wb.Worksheets(1)
        try:
            for s_name, report_path in suppliers:
                upsert_supplier_into_nutzwert(excel, wb, ws, report_path, s_name)
            excel_create_nutzwert_chart(ws, wb)
            wb.SaveAs(builtins.str(out_path))
            state.nutzwert_done = True
            state.save()
            print(f"\n[FINISH] Nutzwertanalyse erstellt: {out_path}")
            return Path(out_path)
        finally:
            try:
                wb.Close(SaveChanges=True)
            except Exception:
                pass

def phase_send_final_mail(outlook: OutlookUI, state: RoundState, to_email: str, attachment: Path) -> None:
    if state.final_mail_sent:
        return
    subject = f"SCM Nutzwertanalyse Runde {state.round_id}"
    body = (
        f"Hallo,\n\n"
        f"anbei erhalten Sie die aktuelle Nutzwertanalyse zur Lieferantenbewertung (Runde {state.round_id}).\n\n"
        f"Die Datei enthält die konsolidierten Ergebnisse aus den eingegangenen Bewertungen sowie den ERP-Daten.\n"
        f"Bitte nutzen Sie die Auswertung als Entscheidungsgrundlage; die finale Lieferantenauswahl verbleibt bei Ihnen.\n\n"
        f"Freundliche Grüße\n"
        f"SCM Bot\n\n"
        f"(Runde {state.round_id})"
    )
    outlook.open_mail()
    outlook.new_mail_with_attachment(to_email, subject, body, attachment)
    state.final_mail_sent = True
    state.save()
    print(f"[FINISH] Abschlussmail mit Nutzwertanalyse gesendet an {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    config = Config()
    config.ensure_dirs()

    state = RoundState.load_or_new(config)
    api = FormAPI(config.form_server)
    scale_data = get_comprehensive_scale(config.file_scale)

    with sync_playwright() as p:
        user_data_dir = config.root / "Playwright_SCM_Profile"
        browser = p.chromium.launch_persistent_context(builtins.str(user_data_dir), headless=False, slow_mo=600)
        page = browser.new_page()
        outlook = OutlookUI(page)

        phase_dispatch_links(config, api, outlook, state)
        phase_poll_and_process(config, api, scale_data, state)

        out_path = phase_finalize_nutzwert(config, state)
        if out_path:
            try:
                phase_send_final_mail(outlook, state, config.send_to_final, out_path)
            except Exception as e:
                print(f"[!] Konnte Abschlussmail nicht senden: {e}")

if __name__ == "__main__":
    main()
