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
VERSION = "0.1.4"

DEFAULT_CONFIG = {
    "input_dir": "XLSM",
    "output_dir": "PDF",
    "link_url": "",
    "video_marker": "vídeo",
    "link_texts": [
        "nesse vídeo, aqui",
        "neste vídeo, aqui",
        "nesse vídeo - aqui",
        "neste vídeo - aqui",
        "aqui",
    ],
    "plan_sheet": "LAMINA 2",
    "plan_cell": "A8",
    "month_sheet": "LAMINA 2",
    "month_cell": "A4",
    "output_pattern": "[NOME DO PLANO NO EXCEL] - Raio X de [MÊS DO RAIO X]",
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

def prepare_excel_for_pdf(workbook):
    s1 = workbook.Worksheets("LAMINA")
    s2 = workbook.Worksheets("LAMINA 2")
    state = {"visible": {ws.Name: ws.Visible for ws in workbook.Worksheets}, "active_sheet": None}
    try:
        state["active_sheet"] = workbook.ActiveSheet.Name
    except Exception:
        pass

    # Do not rewrite the source workbook's print settings. The received XLSM
    # already contains the intended print areas/layout. We only select the two
    # publication sheets so Excel exports them as one PDF.
    for ws in (s1, s2):
        if ws.Visible != -1:
            ws.Visible = -1
    try:
        s1.Select()
        s2.Select(False)
    except Exception:
        pass
    return state

def restore_excel_state(workbook, state):
    try:
        if state.get("active_sheet"):
            workbook.Worksheets(state["active_sheet"]).Activate()
    except Exception:
        pass
    for name, vis in state.get("visible", {}).items():
        try:
            workbook.Worksheets(name).Visible = vis
        except Exception:
            pass

def export_pdf(excel, workbook, output: Path):
    # Export to the user's temp directory first. This separates Excel's PDF
    # generation from permissions/sync issues in the destination folder.
    import tempfile
    temp_pdf = Path(tempfile.gettempdir()) / f"raiox_{os.getpid()}_{int(time.time()*1000)}.pdf"
    state = prepare_excel_for_pdf(workbook)
    try:
        # With the two sheets selected, ActiveSheet.ExportAsFixedFormat exports
        # the selected worksheets together. Only the two required positional
        # arguments are supplied to avoid COM named-argument quirks.
        workbook.ActiveSheet.ExportAsFixedFormat(0, str(temp_pdf))
        if not temp_pdf.exists() or temp_pdf.stat().st_size == 0:
            raise RuntimeError("O Excel não criou o PDF temporário.")
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            output.unlink()
        import shutil
        shutil.copy2(temp_pdf, output)
    finally:
        restore_excel_state(workbook, state)
        try:
            if temp_pdf.exists():
                temp_pdf.unlink()
        except Exception:
            pass

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
        while len(doc) > expected_pages and len(doc) > 1:
            if not trailing_junk_page(doc[-1]):
                break
            doc.delete_page(len(doc) - 1)
        after = len(doc)
        if after != before:
            temp = str(pdf_path) + ".trimmed"
            doc.save(temp, garbage=4, deflate=True)
            doc.close()
            os.replace(temp, pdf_path)
            return before, after
        return before, after
    finally:
        try:
            doc.close()
        except Exception:
            pass

def find_link_rect(page: fitz.Page, candidates: list[str]) -> Optional[fitz.Rect]:
    for phrase in candidates:
        phrase = phrase.strip()
        if not phrase:
            continue
        try:
            rects = page.search_for(phrase, quads=False)
            if rects:
                return rects[-1]
        except Exception:
            continue
    return None

def add_video_hyperlink(pdf_path: Path, link_url: str, video_marker: str, link_texts: list[str]) -> tuple[bool, str]:
    doc = fitz.open(pdf_path)
    try:
        marker_norm = normalize_text(video_marker)
        target_page = None
        for i in range(len(doc) - 1, -1, -1):
            txt = normalize_text(doc[i].get_text("text") or "")
            if marker_norm and marker_norm in txt:
                target_page = doc[i]
                break
        if target_page is None:
            return False, "nenhuma chamada de vídeo encontrada"
        if not link_url.strip():
            return False, "vídeo encontrado, mas URL do vídeo não configurada"

        rect = find_link_rect(target_page, link_texts)
        if rect is None:
            rects = target_page.search_for("aqui", quads=False)
            if rects:
                rect = rects[-1]
        if rect is None:
            return False, "vídeo encontrado, mas trecho clicável não localizado"

        target_page.insert_link({
            "kind": fitz.LINK_URI,
            "from": rect,
            "uri": link_url.strip(),
        })
        temp = str(pdf_path) + ".link"
        doc.save(temp, garbage=4, deflate=True)
        doc.close()
        os.replace(temp, pdf_path)
        return True, "hyperlink inserido"
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
    ws_month = workbook.Worksheets(config.get("month_sheet", "LAMINA 2"))
    plan = str(safe_cell_value(ws_plan, config.get("plan_cell", "A8")) or "").strip()
    strip_prefix = str(config.get("strip_plan_prefix", ""))
    if strip_prefix and plan.casefold().startswith(strip_prefix.casefold()):
        plan = plan[len(strip_prefix):].strip()
    d = excel_date_to_datetime(safe_cell_value(ws_month, config.get("month_cell", "A4")))
    month = MONTHS_PT[d.month] if d else infer_month_from_filename(input_path)
    if not month:
        raise ValueError(f"Não foi possível determinar o mês do Raio-X em {input_path.name}")
    if not plan:
        raise ValueError(f"Não foi possível determinar o nome do plano em {input_path.name}")
    pattern = str(config.get("output_pattern", DEFAULT_CONFIG["output_pattern"]))
    name = pattern.replace("[NOME DO PLANO NO EXCEL]", plan).replace("[MÊS DO RAIO X]", month)
    return clean_filename(name) + ".pdf", plan, month

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
    output_pdf = None
    try:
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
        if output_pdf.exists():
            output_pdf.unlink()

        export_pdf(excel, workbook, output_pdf)
        time.sleep(0.25)

        with fitz.open(output_pdf) as doc:
            result.pages_before = len(doc)
        expected = int(config.get("expected_pages", 3))
        if config.get("trim_trailing_junk", True):
            result.pages_before, result.pages_after = trim_extra_pages(output_pdf, expected)
        else:
            result.pages_after = result.pages_before

        link_added, detail = add_video_hyperlink(
            output_pdf,
            str(config.get("link_url", "")),
            str(config.get("video_marker", "vídeo")),
            [str(x) for x in config.get("link_texts", []) if str(x).strip()]
        )
        result.link = link_added
        result.video = "encontrada" in detail or "encontrado" in detail or link_added
        if result.pages_after == expected and (not detail.startswith("vídeo encontrado") or link_added):
            result.status = "OK"
        else:
            result.status = "ATENÇÃO"
            result.detail = detail if "vídeo encontrado" in detail else f"PDF ficou com {result.pages_after} páginas (esperadas {expected})"
        result.output = output_pdf.name
        log(f"Concluído: {input_path.name} -> {result.output} [{result.status}] {result.detail}")
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

def worker(config: dict[str, Any], callback) -> None:
    input_dir = Path(config["input_dir"]).expanduser()
    output_dir = Path(config["output_dir"]).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)
    files = sorted([p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".xlsm"], key=lambda p: p.name.casefold())
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
        ttk.Label(root, text="XLSM → PDF + hyperlink + correção de páginas + renomeação automática").pack(anchor="w", pady=(2,14))

        frm = ttk.LabelFrame(root, text="Configurações", padding=12)
        frm.pack(fill="x")
        for i in range(2): frm.columnconfigure(i, weight=1)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.link_var = tk.StringVar()
        self.marker_var = tk.StringVar()
        self.click_var = tk.StringVar()
        self.pattern_var = tk.StringVar()
        self.strip_prefix_var = tk.StringVar()
        self.expected_pages_var = tk.StringVar()

        self._field(frm,0,"Pasta dos XLSM",self.input_var,True)
        self._field(frm,1,"Pasta dos PDFs",self.output_var,True)
        self._field(frm,2,"URL do vídeo",self.link_var)
        self._field(frm,3,"Texto que indica vídeo",self.marker_var)
        self._field(frm,4,"Trecho clicável / variações (;) ",self.click_var)
        self._field(frm,5,"Regra de nome",self.pattern_var)
        self._field(frm,6,"Prefixo a remover (opcional)",self.strip_prefix_var)
        self._field(frm,7,"Páginas esperadas",self.expected_pages_var)

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
        self.marker_var.set(self.config_data.get("video_marker","vídeo"))
        self.click_var.set("; ".join(self.config_data.get("link_texts",[])))
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
                "video_marker":self.marker_var.get(),
                "link_texts":[x.strip() for x in self.click_var.get().split(";") if x.strip()],
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
