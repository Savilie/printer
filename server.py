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
    printer: Optional[Literal["tlp100", "lp58", "both"]] = "both"
    copies: Optional[int] = 1


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
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")[:60]


def build_tlp100(qr_text: str, title: str, subtitle: str = "") -> bytes:
    """EPL-этикетка для TLP100 (50×30): QR слева, текст справа."""
    lines = []
    lines.append("N")
    lines.append("q400")
    lines.append("Q300,24")
    # QR-код (слева)
    lines.append(f'b20,15,Q,m2,s13,eL,"{_epl_escape(qr_text)}"')
    # Правая колонка: первая строка крупно
    lines.append(f'A320,150,1,2,1,1,N,"{_epl_escape(title)}"')
    if subtitle:
        lines.append(f'A320,90,1,1,1,1,N,"{_epl_escape(subtitle)}"')
    lines.append("P1")
    return "\n".join(lines).encode("cp1251")


def build_lp58(qr_text: str, title: str, subtitle: str = "") -> bytes:
    """TSPL для LP58 EVA (58×40): QR слева, текст вертикально справа."""
    lines = []
    lines.append("SIZE 58 mm, 40 mm")
    lines.append("GAP 2 mm, 0 mm")
    lines.append("CLS")
    t = (title or qr_text).replace("ID:", "").replace("*-", "")[:24]
    lines.append(f'TEXT 380,180,"3",90,1,1,"{t}"')
    lines.append(f'QRCODE 30,16,H,13,A,0,M2,S7,"{qr_text[:80]}"')
    lines.append("PRINT 1")
    return "\n".join(lines).encode("ascii", errors="replace")


def print_tlp100(qr_text: str, title: str, subtitle: str = "", printer_name: str = "MPRINT Terra Nova TLP100", copies: int = 1) -> bool:
    ok = True
    for _ in range(max(1, copies)):
        if not _raw_printer(printer_name, build_tlp100(qr_text, title, subtitle)):
            ok = False
    print(f"📨 TLP100 ×{copies}: {qr_text}")
    return ok


def print_lp58(qr_text: str, title: str, subtitle: str = "", printer_name: str = "MPRINT LP58 EVA", copies: int = 1) -> bool:
    ok = True
    for _ in range(max(1, copies)):
        if not _raw_printer(printer_name, build_lp58(qr_text, title, subtitle)):
            ok = False
    print(f"📨 LP58 ×{copies}: {qr_text}")
    return ok


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

    # Надпись на этикетке без служебных префиксов
    raw = req.title or qr_text
    title = raw.replace("*-", "").replace("ID:", "").replace("CHZ-BOX", "Коробка " + raw.replace("CHZ-BOX", "")) if raw == qr_text and raw.startswith("CHZ-BOX") else raw.replace("*-", "").replace("ID:", "")
    subtitle = req.subtitle or ""

    results = {}
    if req.printer in ("tlp100", "both"):
        results["tlp100"] = print_tlp100(qr_text, title, subtitle, copies=req.copies or 1)
    if req.printer in ("lp58", "both"):
        results["lp58"] = print_lp58(qr_text, title, subtitle, copies=req.copies or 1)
    ok = bool(results) and all(results.values())
    return {"ok": ok, "results": results, "qr": qr_text}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=5000)
