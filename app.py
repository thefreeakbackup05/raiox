from __future__ import annotations

import json
import os
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Any, Optional

import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Gerador de Raio-X"
VERSION = "0.3.2"

DEFAULT_CONFIG = {
    "input_dir": "XLSM",
    "output_dir": "PDF",
    "link_url": "",
    "link_trigger_phrases": [
        "nesse vídeo - aqui",
        "neste vídeo, aqui",
        "nesse vídeo, aqui",
        "neste vídeo - aqui"
    ],
    "click_text": "aqui",
    "plan_sheet": "LAMINA 2",
    "plan_cell": "A8",
    "month_sheet": "LAMINA 2",
    "month_cell": "A4",
    "output_pattern": "[NOME DO PLANO NO EXCEL] - Raio X de agosto",
    "strip_plan_prefix": "",
    "expected_pages": 3,
    "trim_trailing_junk": True,
    "overwrite": True,
}

MONTHS_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

CONFIG_PATH = app_dir() / "config.json"
LOG_PATH = app_dir() / "gerador_raio_x.log"

def load_config() -> dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        try:
            with CONFIG_PATH.open("r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg

def save_config(cfg: dict[str, Any]) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def log(msg: str) -> None:
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n"
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().casefold()

def clean_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "-", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(".")
    return name or "Raio X"

def excel_date_to_datetime(value: Any) -> Optional[datetime]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(value))
        except Exception:
            return None
    s = str(value).strip()
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%B %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def safe_cell_value(sheet, address: str) -> Any:
    return sheet.Range(address).Value

def _xl_col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s

def _sheet_bounds(sheet):
    used = sheet.UsedRange
    first_row = max(1, used.Row)
    first_col = max(1, used.Column)
    last_row = max(1, used.Row + used.Rows.Count - 1)
    last_col = max(1, used.Column + used.Columns.Count - 1)
    try:
        first = used.Find("*", None, -4163, 1, 1, 1, False, False, False)
        last = used.Find("*", None, -4163, 1, 1, 2, False, False, False)
        if first is not None:
            first_row, first_col = first.Row, first.Column
        if last is not None:
            last_row, last_col = last.Row, last.Column
    except Exception:
        pass
    return first_row, first_col, last_row, last_col

def read_layout_metadata(input_path: Path) -> dict[str, Any]:
    # Read merges/manual page breaks/chart anchors directly from the XLSX/XLSM
    # package. This avoids guessing from visual output and is fast enough for
    # a monthly batch of files.
    meta = {"merges": [], "manual_breaks": [], "charts": []}
    try:
        ns = {
            "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        }
        with zipfile.ZipFile(input_path, "r") as z:
            wb_root = ET.fromstring(z.read("xl/workbook.xml"))
            rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rels = {e.attrib["Id"]: e.attrib["Target"] for e in rel_root}

            target = None
            for sh in wb_root.findall("m:sheets/m:sheet", ns):
                if sh.attrib.get("name") == "LAMINA 2":
                    rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rels.get(rid)
                    break
            if not target:
                return meta

            sheet_path = os.path.normpath(os.path.join("xl", target))
            root = ET.fromstring(z.read(sheet_path))

            merge_node = root.find("m:mergeCells", ns)
            if merge_node is not None:
                meta["merges"] = [x.attrib["ref"] for x in merge_node if x.attrib.get("ref")]

            rb = root.find("m:rowBreaks", ns)
            if rb is not None:
                for br in rb:
                    if br.attrib.get("man") == "1":
                        try:
                            meta["manual_breaks"].append(int(br.attrib.get("id", "0")))
                        except Exception:
                            pass

            drawing = root.find("m:drawing", ns)
            if drawing is not None:
                sheet_rel_path = os.path.dirname(sheet_path) + "/_rels/" + os.path.basename(sheet_path) + ".rels"
                srel = ET.fromstring(z.read(sheet_rel_path))
                srela = {e.attrib["Id"]: e.attrib["Target"] for e in srel}
                drid = drawing.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                drtarget = srela.get(drid)
                if drtarget:
                    dpath = os.path.normpath(os.path.join(os.path.dirname(sheet_path), drtarget))
                    droot = ET.fromstring(z.read(dpath))
                    for anchor in list(droot):
                        if anchor.tag.split("}")[-1] not in ("twoCellAnchor", "oneCellAnchor"):
                            continue
                        fr = anchor.find("xdr:from", ns)
                        to = anchor.find("xdr:to", ns)
                        if fr is None:
                            continue
                        fr_row = fr.find("xdr:row", ns)
                        to_row = to.find("xdr:row", ns)
                        if fr_row is None:
                            continue
                        top = int(fr_row.text) + 1
                        bottom = int(to_row.text) + 1 if to_row is not None else top
                        if anchor.find(".//c:chart", ns) is not None:
                            meta["charts"].append({"top": top, "bottom": bottom})
    except Exception:
        pass
    return meta

def _expand_merge_ref(ref: str):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", str(ref))
    if not m:
        return None
    return m.group(1), int(m.group(2)), m.group(3), int(m.group(4))

def _adjust_manual_page_breaks(sheet, meta: dict[str, Any]):
    original = sorted(set(x for x in meta.get("manual_breaks", []) if x > 0))
    if not original:
        return
    merges = [_expand_merge_ref(x) for x in meta.get("merges", [])]
    merges = [x for x in merges if x]
    adjusted = []
    for after_row in original:
        target = after_row
        for _, start_row, _, end_row in merges:
            if start_row <= after_row < end_row:
                target = max(target, end_row)
        for chart in meta.get("charts", []):
            if chart["top"] <= after_row < chart["bottom"]:
                target = max(target, chart["bottom"])
        adjusted.append(target)

    try:
        for i in range(sheet.HPageBreaks.Count, 0, -1):
            pb = sheet.HPageBreaks.Item(i)
            try:
                if int(pb.Type) == -4135:
                    pb.Delete()
            except Exception:
                pass
        for after_row in sorted(set(adjusted)):
            sheet.HPageBreaks.Add(sheet.Cells(after_row + 1, 1))
    except Exception:
        pass

def _content_bottom(sheet, meta: dict[str, Any]) -> int:
    _, _, last_row, _ = _sheet_bounds(sheet)
    for ref in meta.get("merges", []):
        x = _expand_merge_ref(ref)
        if x:
            last_row = max(last_row, x[3])
    for chart in meta.get("charts", []):
        last_row = max(last_row, int(chart["bottom"]))
    return last_row

def _parse_print_area(area_text: str):
    if not area_text:
        return None
    m = re.search(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", str(area_text))
    if not m:
        return None
    return m.group(1), int(m.group(2)), m.group(3), int(m.group(4))

def prepare_excel_for_pdf(workbook, layout_meta: Optional[dict[str, Any]] = None):
    layout_meta = layout_meta or {}
    sheet = workbook.Worksheets("LAMINA 2")
    state = {
        "active_sheet": None,
        "print_area": None,
        "had_print_area": False,
        "zoom": None,
        "fit_wide": None,
        "fit_tall": None,
    }
    try:
        state["active_sheet"] = workbook.ActiveSheet.Name
    except Exception:
        pass
    try:
        state["print_area"] = sheet.PageSetup.PrintArea
        state["had_print_area"] = bool(state["print_area"])
        state["zoom"] = sheet.PageSetup.Zoom
        state["fit_wide"] = sheet.PageSetup.FitToPagesWide
        state["fit_tall"] = sheet.PageSetup.FitToPagesTall
    except Exception:
        pass

    content_bottom = _content_bottom(sheet, layout_meta)

    try:
        area = _parse_print_area(state["print_area"])
        if area:
            first_col, first_row, last_col, print_last_row = area
            final_row = max(print_last_row, content_bottom)
            sheet.PageSetup.PrintArea = (
                "$" + first_col + "$" + str(first_row) +
                ":$" + last_col + "$" + str(final_row)
            )
        else:
            first_row, first_col, last_row, last_col_num = _sheet_bounds(sheet)
            sheet.PageSetup.PrintArea = (
                "$" + _xl_col(first_col) + "$" + str(first_row) +
                ":$" + _xl_col(last_col_num) + "$" + str(max(last_row, content_bottom))
            )
            # Ford-like files have no native print area: fit the width only.
            # Height is automatic so nothing is vertically squeezed/cut.
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False
    except Exception:
        pass

    _adjust_manual_page_breaks(sheet, layout_meta)
    try:
        sheet.Activate()
    except Exception:
        pass
    return state

def restore_excel_state(workbook, state):
    try:
        sheet = workbook.Worksheets("LAMINA 2")
        if state.get("had_print_area"):
            sheet.PageSetup.PrintArea = state.get("print_area")
        else:
            sheet.PageSetup.PrintArea = ""
            if state.get("zoom") is not None:
                sheet.PageSetup.Zoom = state.get("zoom")
            if state.get("fit_wide") is not None:
                sheet.PageSetup.FitToPagesWide = state.get("fit_wide")
            if state.get("fit_tall") is not None:
                sheet.PageSetup.FitToPagesTall = state.get("fit_tall")
        if state.get("active_sheet"):
            workbook.Worksheets(state["active_sheet"]).Activate()
    except Exception:
        pass

def _remove_broken_defined_names(workbook) -> int:
    removed = 0
    try:
        for name in list(workbook.Names):
            try:
                if "#REF!" in str(name.RefersTo):
                    name.Delete()
                    removed += 1
            except Exception:
                pass
    except Exception:
        pass
    return removed

def export_pdf(excel, workbook, output: Path, layout_meta: Optional[dict[str, Any]] = None):
    state = prepare_excel_for_pdf(workbook, layout_meta)
    last_error = None
    try:
        sheet = workbook.Worksheets("LAMINA 2")
        for attempt in range(1, 4):
            try:
                if Path(output).exists():
                    Path(output).unlink()
                sheet.ExportAsFixedFormat(0, str(output))
                if Path(output).exists() and Path(output).stat().st_size > 0:
                    return state.get("had_print_area", False)
            except Exception as exc:
                last_error = exc
                time.sleep(1.0)

        try:
            removed = _remove_broken_defined_names(workbook)
            if removed:
                time.sleep(0.5)
                if Path(output).exists():
                    Path(output).unlink()
                sheet.ExportAsFixedFormat(0, str(output))
                if Path(output).exists() and Path(output).stat().st_size > 0:
                    return state.get("had_print_area", False)
        except Exception as exc:
            last_error = exc

        visible = {}
        try:
            for ws in workbook.Worksheets:
                try:
                    visible[ws.Name] = ws.Visible
                except Exception:
                    continue
                if ws.Name == "LAMINA 2":
                    ws.Visible = -1
                else:
                    try:
                        ws.Visible = 0
                    except Exception:
                        pass
            sheet.Activate()
            if Path(output).exists():
                Path(output).unlink()
            workbook.ExportAsFixedFormat(0, str(output))
            if Path(output).exists() and Path(output).stat().st_size > 0:
                return state.get("had_print_area", False)
        except Exception as exc:
            last_error = exc
        finally:
            for name, vis in visible.items():
                try:
                    workbook.Worksheets(name).Visible = vis
                except Exception:
                    pass
        raise last_error or RuntimeError("Falha na exportação para PDF.")
    finally:
        restore_excel_state(workbook, state)


def page_ink_ratio(page: fitz.Page) -> float:
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.10, 0.10), alpha=False)
        samples = pix.samples
        n = pix.width * pix.height
        if n <= 0:
            return 0.0
        ink = sum(
            1 for i in range(0, len(samples), 3)
            if samples[i] < 245 or samples[i + 1] < 245 or samples[i + 2] < 245
        )
        return ink / n
    except Exception:
        return 1.0

def trailing_junk_page(page: fitz.Page) -> bool:
    text = re.sub(r"\s+", " ", page.get_text("text") or "").strip()
    if len(text) >= 80:
        return False
    try:
        rect = page.rect
        page_area = max(rect.width * rect.height, 1)
        image_area = 0.0
        image_count = 0
        for info in page.get_image_info(xrefs=True):
            image_count += 1
            r = fitz.Rect(info["bbox"])
            image_area += max(0, r.width) * max(0, r.height)
        coverage = min(image_area / page_area, 1.0)
    except Exception:
        image_count = 0
        coverage = 0.0
    ink = page_ink_ratio(page)
    if not text and coverage < 0.10 and ink < 0.02:
        return True
    if len(text) < 25 and coverage < 0.08 and ink < 0.015:
        return True
    if len(text) < 25 and image_count >= 1 and coverage < 0.05 and ink < 0.03:
        return True
    return False

def trim_extra_pages(pdf_path: Path, expected_pages: int) -> tuple[int, int]:
    doc = fitz.open(pdf_path)
    before = len(doc)
    try:
        while len(doc) > expected_pages:
            last = doc[-1]
            text = (last.get_text("text") or "").strip()
            try:
                images = len(last.get_images(full=True))
            except Exception:
                images = 0
            if text or images:
                break
            doc.delete_page(len(doc) - 1)
        after = len(doc)
        if after != before:
            temp = str(pdf_path) + ".trimmed"
            doc.save(temp, garbage=4, deflate=True)
            doc.close()
            os.replace(temp, pdf_path)
        return before, after
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _word_token(text: str) -> str:
    return re.sub(r"^[^\wÀ-ÿ]+|[^\wÀ-ÿ]+$", "", normalize_text(text))

def find_trigger_link_word(page: fitz.Page, trigger_phrases: list[str], click_word: str = "aqui") -> Optional[fitz.Rect]:
    raw_words = page.get_text("words") or []
    words = []
    for w in raw_words:
        token = _word_token(w[4])
        if token:
            words.append((token, fitz.Rect(w[0], w[1], w[2], w[3])))

    click_norm = _word_token(click_word)
    for phrase in trigger_phrases:
        wanted = [_word_token(x) for x in re.findall(r"\S+", phrase) if _word_token(x)]
        if not wanted or wanted[-1] != click_norm:
            continue
        for start in range(0, len(words) - len(wanted) + 1):
            if [x[0] for x in words[start:start + len(wanted)]] == wanted:
                return words[start + len(wanted) - 1][1]
    return None

def add_video_hyperlink(pdf_path: Path, link_url: str, trigger_phrases: list[str], click_word: str) -> tuple[bool, bool, str]:
    doc = fitz.open(pdf_path)
    try:
        for page_index in range(len(doc) - 1, -1, -1):
            page = doc[page_index]
            rect = find_trigger_link_word(page, trigger_phrases, click_word)
            if rect is None:
                continue
            if not link_url.strip():
                return True, False, "vídeo detectado, mas URL não configurada"

            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": rect,
                "uri": link_url.strip(),
            })
            page.draw_line(
                fitz.Point(rect.x0, rect.y1 + 0.8),
                fitz.Point(rect.x1, rect.y1 + 0.8),
                color=(0, 0, 0),
                width=0.7,
                overlay=True,
            )
            doc.saveIncr()
            return True, True, f"hyperlink inserido na página {page_index + 1}"
        return False, False, "nenhuma frase de vídeo configurada foi encontrada"
    finally:
        try:
            doc.close()
        except Exception:
            pass


@dataclass
class ProcessResult:
    file: str
    output: str = ""
    pages_before: int = 0
    pages_after: int = 0
    video: bool = False
    link: bool = False
    status: str = ""
    detail: str = ""

def infer_month_from_filename(path: Path) -> Optional[str]:
    name = normalize_text(path.stem)
    months = {
        "jan":1,"janeiro":1,"fev":2,"fevereiro":2,"mar":3,"marco":3,"março":3,
        "abr":4,"abril":4,"mai":5,"maio":5,"jun":6,"junho":6,"jul":7,"julho":7,
        "ago":8,"agosto":8,"set":9,"setembro":9,"out":10,"outubro":10,
        "nov":11,"novembro":11,"dez":12,"dezembro":12
    }
    for token, num in months.items():
        if re.search(rf"(?<![a-z]){re.escape(token)}(?![a-z])", name):
            return MONTHS_PT[num]
    return None

def output_name(workbook, input_path: Path, config: dict[str, Any]) -> tuple[str, str, str]:
    ws_plan = workbook.Worksheets(config.get("plan_sheet", "LAMINA 2"))
    plan = str(safe_cell_value(ws_plan, config.get("plan_cell", "A8")) or "").strip()
    strip_prefix = str(config.get("strip_plan_prefix", ""))
    if strip_prefix and plan.casefold().startswith(strip_prefix.casefold()):
        plan = plan[len(strip_prefix):].strip()
    if not plan:
        raise ValueError(f"Não foi possível determinar o nome do plano em {input_path.name}")
    pattern = str(config.get("output_pattern", DEFAULT_CONFIG["output_pattern"]))
    if "[MÊS DO RAIO X]" in pattern:
        raise ValueError("A regra de nome ainda contém [MÊS DO RAIO X]. Digite o mês diretamente na regra, por exemplo: [NOME DO PLANO NO EXCEL] - Raio X de agosto")
    name = pattern.replace("[NOME DO PLANO NO EXCEL]", plan)
    return clean_filename(name) + ".pdf", plan, ""

def create_excel_instance():
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        excel.EnableEvents = False
    except Exception:
        pass
    try:
        excel.AutomationSecurity = 3
    except Exception:
        pass
    return excel

def process_one(excel, input_path: Path, output_dir: Path, config: dict[str, Any]) -> ProcessResult:
    result = ProcessResult(file=input_path.name)
    workbook = None
    working_pdf = None
    try:
        import tempfile
        import shutil

        layout_meta = read_layout_metadata(input_path)
        workbook = excel.Workbooks.Open(
            str(input_path), UpdateLinks=0, ReadOnly=True,
            IgnoreReadOnlyRecommended=True, AddToMru=False
        )
        try:
            excel.CalculateFull()
        except Exception:
            pass

        name, plan, month = output_name(workbook, input_path, config)
        output_pdf = output_dir / name

        if output_pdf.exists() and not config.get("overwrite", True):
            raise FileExistsError(f"Arquivo já existe: {output_pdf.name}")

        working_pdf = Path(tempfile.gettempdir()) / f"raiox_work_{os.getpid()}_{time.time_ns()}.pdf"
        had_native_print_area = export_pdf(excel, workbook, working_pdf, layout_meta)
        time.sleep(0.25)

        with fitz.open(working_pdf) as doc:
            result.pages_before = len(doc)

        configured_expected = int(config.get("expected_pages", 3))
        expected = configured_expected if had_native_print_area else 0
        if config.get("trim_trailing_junk", True):
            result.pages_before, result.pages_after = trim_extra_pages(working_pdf, expected)
        else:
            result.pages_after = result.pages_before

        video_found, link_added, detail = add_video_hyperlink(
            working_pdf,
            str(config.get("link_url", "")),
            [str(x) for x in config.get("link_trigger_phrases", []) if str(x).strip()],
            str(config.get("click_text", "aqui"))
        )
        result.video = video_found
        result.link = link_added

        output_dir.mkdir(parents=True, exist_ok=True)
        if output_pdf.exists() and config.get("overwrite", True):
            try:
                output_pdf.unlink()
            except PermissionError as exc:
                raise PermissionError(
                    f"O PDF de destino está aberto ou bloqueado pelo Windows/OneDrive: {output_pdf.name}"
                ) from exc

        shutil.copy2(working_pdf, output_pdf)
        result.output = output_pdf.name

        if result.pages_after != expected:
            result.status = "ATENÇÃO"
            result.detail = f"{result.pages_after} páginas (esperadas {expected})"
        elif video_found and not link_added:
            result.status = "ATENÇÃO"
            result.detail = detail
        else:
            result.status = "OK"
            result.detail = detail if video_found else ""

        log(f"{input_path.name} -> {result.output} [{result.status}] {result.detail}")
        return result
    except Exception as exc:
        result.status = "ERRO"
        result.detail = str(exc)
        log(f"ERRO {input_path.name}: {exc}\n{traceback.format_exc()}")
        return result
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if working_pdf is not None:
            try:
                if Path(working_pdf).exists():
                    Path(working_pdf).unlink()
            except Exception:
                pass


def worker(config: dict[str, Any], callback) -> None:
    input_dir = Path(config["input_dir"]).expanduser()
    output_dir = Path(config["output_dir"]).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(
        [p for p in input_dir.iterdir()
         if p.is_file() and p.suffix.lower() in {".xlsm", ".xlsx"}],
        key=lambda p: p.name.casefold()
    )
    if not files:
        callback("finished", [], "Nenhum .xlsm encontrado na pasta de entrada.")
        return
    excel = None
    results = []
    try:
        excel = create_excel_instance()
        total = len(files)
        for idx, path in enumerate(files, start=1):
            callback("progress", (idx - 1, total, path.name))
            res = process_one(excel, path, output_dir, config)
            results.append(res)
            callback("result", res)
            callback("progress", (idx, total, path.name))
    finally:
        if excel is not None:
            try: excel.Quit()
            except Exception: pass
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
    callback("finished", results, f"Processamento concluído: {len(results)} arquivo(s).")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("980x720")
        self.minsize(860, 620)
        self.config_data = load_config()
        self.results = []
        self._build_ui()
        self._load_form()

    def _build_ui(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="GERADOR DE RAIO-X DE INVESTIMENTOS", font=("Segoe UI", 17, "bold")).pack(anchor="w")
        ttk.Label(root, text="XLSX/XLSM → PDF + preservação do conteúdo + hyperlink").pack(anchor="w", pady=(2,14))

        frm = ttk.LabelFrame(root, text="Configurações", padding=12)
        frm.pack(fill="x")
        for i in range(2): frm.columnconfigure(i, weight=1)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.link_var = tk.StringVar()
        self.trigger_var = tk.StringVar()
        self.click_var = tk.StringVar()
        self.pattern_var = tk.StringVar()
        self.strip_prefix_var = tk.StringVar()
        self.expected_pages_var = tk.StringVar()

        self._field(frm,0,"Pasta dos arquivos Excel",self.input_var,True)
        self._field(frm,1,"Pasta dos PDFs",self.output_var,True)
        self._field(frm,2,"URL do vídeo",self.link_var)
        self._field(frm,3,"Frases exatas que ativam o link (;)",self.trigger_var)
        self._field(frm,4,"Palavra clicável",self.click_var)
        self._field(frm,5,"Regra de nome",self.pattern_var)
        self._field(frm,6,"Prefixo a remover (opcional)",self.strip_prefix_var)
        self._field(frm,7,"Páginas esperadas (0 = automático)",self.expected_pages_var)

        btns=ttk.Frame(root); btns.pack(fill="x",pady=14)
        self.generate_btn=ttk.Button(btns,text="GERAR TODOS",command=self.start_generation)
        self.generate_btn.pack(side="left")
        ttk.Button(btns,text="Abrir entrada",command=lambda:self.open_folder(self.input_var.get())).pack(side="left",padx=8)
        ttk.Button(btns,text="Abrir PDFs",command=lambda:self.open_folder(self.output_var.get())).pack(side="left")
        ttk.Button(btns,text="Salvar configurações",command=self.save_form).pack(side="right")

        self.progress=ttk.Progressbar(root,orient="horizontal",mode="determinate"); self.progress.pack(fill="x")
        self.status_var=tk.StringVar(value="Pronto."); ttk.Label(root,textvariable=self.status_var).pack(anchor="w",pady=(4,10))

        lf=ttk.LabelFrame(root,text="Resultado",padding=8); lf.pack(fill="both",expand=True)
        cols=("arquivo","pag","video","link","status","detalhe")
        self.tree=ttk.Treeview(lf,columns=cols,show="headings",height=16)
        heads={"arquivo":"Arquivo","pag":"Páginas","video":"Vídeo","link":"Link","status":"Status","detalhe":"Detalhe"}
        widths={"arquivo":270,"pag":90,"video":70,"link":70,"status":90,"detalhe":280}
        for c in cols:
            self.tree.heading(c,text=heads[c])
            self.tree.column(c,width=widths[c],anchor="center" if c in ("pag","video","link","status") else "w")
        vsb=ttk.Scrollbar(lf,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(lf,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        lf.rowconfigure(0,weight=1); lf.columnconfigure(0,weight=1)

    def _field(self,parent,row,label,var,directory=False):
        col=row%2; r=row//2
        base=ttk.Frame(parent); base.grid(row=r,column=col,sticky="ew",padx=6,pady=5); base.columnconfigure(0,weight=1)
        ttk.Label(base,text=label).grid(row=0,column=0,sticky="w")
        ttk.Entry(base,textvariable=var).grid(row=1,column=0,sticky="ew",pady=(2,0))
        if directory: ttk.Button(base,text="...",width=4,command=lambda v=var:self.choose_dir(v)).grid(row=1,column=1,padx=(5,0))

    def _load_form(self):
        base=app_dir()
        def resolve(v):
            p=Path(str(v))
            return str((base/p) if not p.is_absolute() else p)
        self.input_var.set(resolve(self.config_data.get("input_dir","XLSM")))
        self.output_var.set(resolve(self.config_data.get("output_dir","PDF")))
        self.link_var.set(self.config_data.get("link_url",""))
        self.trigger_var.set("; ".join(self.config_data.get("link_trigger_phrases",[])))
        self.click_var.set(self.config_data.get("click_text","aqui"))
        self.pattern_var.set(self.config_data.get("output_pattern",DEFAULT_CONFIG["output_pattern"]))
        self.strip_prefix_var.set(self.config_data.get("strip_plan_prefix",""))
        self.expected_pages_var.set(str(self.config_data.get("expected_pages",3)))

    def choose_dir(self,var):
        p=filedialog.askdirectory()
        if p: var.set(p)

    def save_form(self):
        try:
            cfg={
                "input_dir":self.input_var.get(),
                "output_dir":self.output_var.get(),
                "link_url":self.link_var.get(),
                "link_trigger_phrases":[x.strip() for x in self.trigger_var.get().split(";") if x.strip()],
                "click_text":self.click_var.get().strip(),
                "output_pattern":self.pattern_var.get(),
                "strip_plan_prefix":self.strip_prefix_var.get(),
                "expected_pages":int(self.expected_pages_var.get()),
                "trim_trailing_junk":True,
                "overwrite":True,
                "plan_sheet":"LAMINA 2","plan_cell":"A8","month_sheet":"LAMINA 2","month_cell":"A4",
            }
            save_config(cfg); self.config_data=cfg; self.status_var.set("Configurações salvas.")
            return cfg
        except Exception as exc:
            messagebox.showerror("Configuração",str(exc)); return None

    def start_generation(self):
        cfg=self.save_form()
        if not cfg: return
        input_dir=Path(cfg["input_dir"])
        if not input_dir.exists():
            messagebox.showerror("Entrada",f"Pasta não encontrada:\n{input_dir}"); return
        self.results=[]; 
        for item in self.tree.get_children(): self.tree.delete(item)
        self.generate_btn.config(state="disabled")
        self.progress["value"]=0
        self.status_var.set("Iniciando...")
        threading.Thread(target=worker,args=(cfg,self.thread_callback),daemon=True).start()

    def thread_callback(self,kind,data,message=None):
        self.after(0,lambda:self.handle_callback(kind,data,message))

    def handle_callback(self,kind,data,message):
        if kind=="progress":
            idx,total,name=data
            self.progress["maximum"]=total
            self.progress["value"]=idx
            self.status_var.set(f"{idx}/{total} — {name}")
        elif kind=="result":
            r=data
            self.results.append(r)
            self.tree.insert("", "end", values=(r.file,f"{r.pages_before} → {r.pages_after}","SIM" if r.video else "NÃO","OK" if r.link else "—",r.status,r.detail))
        elif kind=="finished":
            self.progress["value"]=self.progress["maximum"]
            self.generate_btn.config(state="normal")
            self.status_var.set(message or "Concluído.")
            if data:
                ok=sum(1 for x in data if x.status=="OK")
                errors=sum(1 for x in data if x.status=="ERRO")
                warnings=sum(1 for x in data if x.status=="ATENÇÃO")
                links=sum(1 for x in data if x.link)
                messagebox.showinfo("Geração concluída",f"Arquivos: {len(data)}\nOK: {ok}\nAtenção: {warnings}\nErros: {errors}\nLinks inseridos: {links}")

    def open_folder(self,p):
        import os
        p=str(p)
        if Path(p).exists(): os.startfile(p)
        else: messagebox.showwarning("Pasta",f"Pasta não encontrada:\n{p}")


# Log any startup exception to a file even when running as a windowed EXE.
def _startup_exception_hook(exc_type, exc_value, exc_tb):
    try:
        with (app_dir() / "startup_error.log").open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Startup error\\n")
            traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"{exc_value}", "Gerador de Raio-X - erro", 0x10)
    except Exception:
        pass

sys.excepthook = _startup_exception_hook

if __name__=="__main__":
    App().mainloop()
