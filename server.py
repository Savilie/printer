from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import win32print
import win32ui
from typing import Literal, Optional

app = FastAPI(title="Принтер-сервер для ПВЗ", version="1.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def private_network_headers(request, call_next):
    """Разрешает браузеру с https-сайта ходить на http://localhost (Private Network Access)."""
    response = await call_next(request)
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


# --- Модели данных для запросов ---
class PrintRequest(BaseModel):
    id: Optional[str] = None
    qr: Optional[str] = None
    title: Optional[str] = None      # человекопонятная первая строка (артикул и т.п.)
    subtitle: Optional[str] = None   # вторая строка (заказ / откуда-куда)
    lines: Optional[list] = None     # до 3 строк текста (id, бренд, артикул / «Коробка», код)
    printer: Optional[Literal["tlp100", "lp58", "both"]] = "both"
    printer_name: Optional[str] = None  # явное имя принтера (если нужно)
    copies: Optional[int] = 1


def _all_printers() -> list:
    names = []
    try:
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            names.append(p[2])
    except Exception as e:
        print("EnumPrinters error:", e)
    return names


def resolve_printer(kind: str, explicit: Optional[str] = None) -> str:
    """Найти подходящий принтер: TLP100 если есть, иначе другой физический.

    Возвращает '' если физических принтеров нет (только PDF/XPS и т.п.).
    """
    names = _all_printers()
    if explicit and explicit in names:
        return explicit
    virtual = ("XPS", "PDF", "Fax", "OneNote", "Generic", "ABBYY", "FineReader")
    physical = [n for n in names if not any(v.lower() in n.lower() for v in virtual)]

    def has(needle):
        return [n for n in physical if needle.lower() in n.lower()]

    if kind == "tlp100":
        prefs = (has("TLP100") and [n for n in has("TLP100") if "- ZPL" not in n]) or has("TLP100") or has("LP58") or physical
    else:  # lp58
        prefs = has("LP58") or has("TLP100") or physical
    return prefs[0] if prefs else ""


def _raw_printer(printer_name: str, data: bytes) -> bool:
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hprinter, 1, ("Label", None, "RAW"))
        try:
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, data)
            win32print.EndPagePrinter(hprinter)
        finally:
            win32print.EndDocPrinter(hprinter)
        return True
    except Exception as e:
        print(f"❌ {printer_name} ошибка: {e}")
        return False
    finally:
        win32print.ClosePrinter(hprinter)


def _epl_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")[:40]


def _label_lines(qr_text: str, title: str, subtitle: str, lines) -> list:
    """Нормализовать строки текста этикетки (без служебных префиксов).

    Приоритет: явный lines → title+subtitle → сам qr (без *- / ID: / CHZ-BOX).
    Возвращает максимум 3 строки.
    """
    if lines:
        raw = [str(x).strip() for x in lines if str(x).strip()]
    else:
        raw = [title or "", subtitle or ""]
        raw = [x.strip() for x in raw if x.strip()]
        if not raw:
            raw = [qr_text]

    out = []
    for s in raw:
        s = s.replace("*-", "").replace("ID:", "").strip()
        if s:
            out.append(s)
    if not out:
        out = [qr_text.replace("*-", "").replace("CHZ-BOX", "Коробка").strip() or qr_text]
    return out[:3]

def _qr_cell_size(data_len: int) -> int:
    if data_len <= 8:
        return 7
    elif data_len <= 14:
        return 5
    elif data_len <= 20:
        return 4
    else:
        return 3


def _qr_ecc_level(data_len: int) -> str:
    return "H" if data_len <= 10 else "M"

def build_tlp100(qr_text: str, title: str = "", subtitle: str = "", lines=None) -> bytes:
    texts = _label_lines(qr_text, title, subtitle, lines)
    limit = 17
    col_step = 18
    x_start = 330
    y_base = 15      # было 292 — теперь якорь у верхнего края, текст растёт ВНИЗ

    rows = []
    for i, t in enumerate(texts):
        x = x_start + i * col_step
        rows.append(f'A{x},{y_base},1,2,1,1,N,"{_epl_escape(t[:limit])}"')

    epl = ["N", "q400", "Q300,24",
           f'b20,15,Q,m2,s13,eL,"{_epl_escape(qr_text)}"']
    epl.extend(rows)
    epl.append("P1")
    return ("\r\n".join(epl) + "\r\n").encode("cp1251")


def build_lp58(qr_text: str, title: str = "", subtitle: str = "", lines=None) -> bytes:
    texts = _label_lines(qr_text, title, subtitle, lines)
    lines_cmd = ["SIZE 58 mm, 40 mm", "GAP 2 mm, 0 mm", "CODEPAGE 866", "CLS"]

    x_start = 380
    col_step = 32
    y_base = 5

    for i, t in enumerate(texts):
        x = x_start + i * col_step
        maxlen = 18
        lines_cmd.append(f'TEXT {x},{y_base},"2",90,1,1,"{t[:maxlen]}"')

    cell = _qr_cell_size(len(qr_text))
    ecc = _qr_ecc_level(len(qr_text))
    lines_cmd.append(f'QRCODE 30,20,{ecc},13,A,0,M2,S{cell},"{qr_text[:80]}"')
    lines_cmd.append("PRINT 1")
    return ("\r\n".join(lines_cmd) + "\r\n").encode("cp866", errors="replace")


def print_tlp100(qr_text: str, title: str = "", subtitle: str = "", printer_name: str = "MPRINT Terra Nova TLP100", copies: int = 1, lines=None) -> bool:
    ok = True
    for _ in range(max(1, copies)):
        if not _raw_printer(printer_name, build_tlp100(qr_text, title, subtitle, lines)):
            ok = False
    print(f"📨 TLP100 ×{copies}: {qr_text}")
    return ok


def print_lp58(qr_text: str, title: str = "", subtitle: str = "", printer_name: str = "MPRINT LP58 EVA", copies: int = 1, lines=None) -> bool:
    ok = True
    for _ in range(max(1, copies)):
        if not _raw_printer(printer_name, build_lp58(qr_text, title, subtitle, lines)):
            ok = False
    print(f"📨 LP58 ×{copies}: {qr_text}")
    return ok


@app.get("/")
def root():
    import os
    import sys
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "index.html")
    if os.path.exists(path):
        return __import__("fastapi.responses", fromlist=["HTMLResponse"]).HTMLResponse(
            open(path, encoding="utf-8").read())
    return {"ok": True, "hint": "index.html не найден рядом с printer-server"}


@app.get("/ping")
def ping():
    return {"ok": True}


@app.get("/printers")
def printers():
    names = []
    try:
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            names.append(p[2])
    except Exception as e:
        return {"printers": [], "error": str(e)}
    return {"printers": names}


@app.post("/print")
def print_label(req: PrintRequest):
    qr_text = req.qr or req.id or ""
    if not qr_text:
        raise HTTPException(400, "Нет данных для печати (qr/id)")

    # Надпись на этикетке без служебных префиксов (теперь можно слать готовые lines)
    title = (req.title or "").replace("*-", "").replace("ID:", "")
    subtitle = req.subtitle or ""
    lines = req.lines or None

    results = {}
    tlp_name = resolve_printer("tlp100", req.printer_name)
    lp_name = resolve_printer("lp58", req.printer_name)
    if req.printer in ("tlp100", "both") and tlp_name:
        results["tlp100"] = print_tlp100(qr_text, title, subtitle, printer_name=tlp_name, copies=req.copies or 1, lines=lines)
    if req.printer in ("lp58", "both") and lp_name:
        results["lp58"] = print_lp58(qr_text, title, subtitle, printer_name=lp_name, copies=req.copies or 1, lines=lines)
    if not results:
        results["error"] = "Физический принтер не найден"
    ok = bool(results) and all(results.values())
    return {"ok": ok, "results": results, "qr": qr_text, "printer_tlp100": tlp_name, "printer_lp58": lp_name}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=5000)
