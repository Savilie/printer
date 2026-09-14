import json
import os
import sys

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


# Порты «печати в файл»: задание уходит в .prn/.txt, а не на принтер.
_FILE_PORTS = ("file:", "portprompt:", "nul", "null")


def _is_file_port(port: str) -> bool:
    p = (port or "").strip().lower()
    return p.startswith("file") or p.startswith("portprompt") or p in ("nul", "null")


# Признаки виртуальных принтеров: они «печатают» в PDF/XPS/файл, а не на бумагу.
# RAW-задание (EPL/TSPL) на такой драйвер = битый PDF в «Документах»
# (имя документа у нас «Label» → Label.pdf) — печать не идёт, файл не открывается.
_VIRTUAL_HINTS = (
    "pdf", "xps", "onenote", "fax", "one note", "print to", "microsoft print",
    "generic", "abbyy", "finereader", "uniflow", "cutepdf", "foxit", "nitro",
    "dopdf", "bullzip", "novapdf", "print to file", "документ xps",
)


def _is_virtual_printer(name: str, port: str = "") -> bool:
    """Виртуальный принтер (PDF/XPS/Fax/OneNote/…) или порт «печать в файл»."""
    low = (name or "").lower()
    if any(h in low for h in _VIRTUAL_HINTS):
        return True
    return _is_file_port(port)


PRINTER_FAMILIES = {
    "tlp100": ("tlp100", "tlp-100", "tlp 100", "terra nova"),
    "lp58": ("lp58", "lp-58", "lp 58", "eva"),
}


def _config_path() -> str:
    """Файл настроек рядом с exe (не внутри распакованного _MEIPASS)."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "printer-config.json")


def _load_prefs() -> dict:
    try:
        with open(_config_path(), encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save_pref(kind: str, name: str) -> None:
    """Запомнить принтер, выбранный вручную (его же использует магазин).

    Виртуальные принтеры (PDF/XPS/Fax/файл) НЕ запоминаем: печатать на них
    нельзя, иначе магазин будет молча «печатать» в Label.pdf в «Документах».
    """
    if _is_virtual_printer(name, _printer_details(name).get("port", "")):
        print(f"⚠ «{name}» — виртуальный принтер (PDF/XPS/файл), в настройки не сохраняю")
        return
    path = _config_path()
    d = _load_prefs()
    if d.get(kind) == name:
        return
    d[kind] = name
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        print(f"💾 Запомнил принтер для «{kind}»: {name}\n   (файл {path})")
    except Exception as e:
        print(f"⚠ не смог сохранить настройку: {e}")


def _forget_pref(kind: str) -> None:
    """Убрать сохранённый принтер для типа (например, оказался виртуальным)."""
    path = _config_path()
    d = _load_prefs()
    if kind not in d:
        return
    d.pop(kind, None)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        print(f"🧹 Убрал негодный сохранённый принтер для «{kind}»")
    except Exception as e:
        print(f"⚠ не смог обновить настройку: {e}")


def _family_match(name: str, kind: str) -> bool:
    """Относится ли принтер к семейству kind (и не к чужому).

    ВАЖНО: EPL и TSPL — разные языки. Нельзя подсунуть TSPL-задание (LP58)
    EPL-принтеру (TLP100) — принтер получает мусор и только мигает.
    """
    low = name.lower()
    other = "lp58" if kind == "tlp100" else "tlp100"
    if any(o in low for o in PRINTER_FAMILIES[other]):
        return False
    return any(m in low for m in PRINTER_FAMILIES[kind])


def _driver_label_mm(printer_name: str) -> tuple:
    """Размер этикетки из настроек драйвера (мм): (ширина, длина). (0, 0) если неизвестно.

    DEVMODE хранит PaperWidth/PaperLength в десятых долях миллиметра.
    Нужно, чтобы файл этикетки и раскладка совпадали с РЕАЛЬНОЙ этикеткой в принтере.
    """
    try:
        h = win32print.OpenPrinter(printer_name)
    except Exception:
        return (0.0, 0.0)
    try:
        info = win32print.GetPrinter(h, 2) or {}
        dm = info.get("pDevMode")
        if not dm:
            return (0.0, 0.0)
        w = float(getattr(dm, "PaperWidth", 0) or 0) / 10.0
        length = float(getattr(dm, "PaperLength", 0) or 0) / 10.0
        if not (5 <= w <= 1000) or not (5 <= length <= 1000):
            return (0.0, 0.0)
        return (round(w, 1), round(length, 1))
    except Exception:
        return (0.0, 0.0)
    finally:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass


def _driver_dpi(printer_name: str) -> int:
    """Разрешение драйвера по X (точек/дюйм). 0 если неизвестно.

    Нужно, чтобы файл этикетки рисовался 1:1 с драйвером: если драйвер печатает
    300 dpi, а картинка сделана на 203 dpi, Windows её растянет — текст и QR
    получаются «волнистыми» и нечитаемыми.
    """
    try:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
    except Exception:
        return 0
    try:
        return int(hdc.GetDeviceCaps(88))  # LOGPIXELSX
    except Exception:
        return 0
    finally:
        try:
            hdc.DeleteDC()
        except Exception:
            pass


def _printer_details(name: str) -> dict:
    """Имя/порт/драйвер принтера (win32print.GetPrinter level 2). {} если недоступно."""
    try:
        h = win32print.OpenPrinter(name)
    except Exception:
        return {}
    try:
        info = win32print.GetPrinter(h, 2) or {}
        return {
            "name": info.get("pPrinterName", name) or name,
            "port": info.get("pPortName", "") or "",
            "driver": info.get("pDriverName", "") or "",
        }
    except Exception:
        return {}
    finally:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass


def _printers_report() -> list:
    """Список принтеров с портом/драйвером + флагом «печать в файл»."""
    out = []
    for n in _all_printers():
        det = _printer_details(n)
        port = det.get("port", "")
        out.append({
            "name": n,
            "port": port,
            "driver": det.get("driver", ""),
            "file_port": _is_file_port(port),
            "virtual": _is_virtual_printer(n, port),
        })
    return out


def _port_rank(port: str) -> int:
    """Чем меньше — тем «физичнее» порт. USB в приоритете, сеть позже."""
    p = (port or "").strip().lower()
    if p.startswith("usb"):
        return 0
    if p.startswith(("lpt", "com", "dot4", "printer")):
        return 1
    if p.startswith("ip_") or p.startswith("wsd") or p.startswith("\\\\") or "network" in p:
        return 2
    return 3


def _default_printer() -> str:
    """Принтер Windows по умолчанию (им обычно и печатают). '' если нет."""
    try:
        return win32print.GetDefaultPrinter() or ""
    except Exception:
        return ""


def resolve_printer(kind: str, explicit: Optional[str] = None) -> str:
    """Найти подходящий физический принтер: TLP100 если есть, иначе LP58, иначе любой.

    Исключаются виртуальные принтеры (XPS/PDF/Fax/OneNote/...) и принтеры с портом
    печати В ФАЙЛ (FILE:/PORTPROMPT:) — иначе задание «печатается» в label.prn
    и на бумагу ничего не выходит.

    Из подходящих сначала берётся принтер Windows по умолчанию, затем лучший по
    типу порта (USB → LPT/COM → сеть), потом по порядку.

    Возвращает '' если физических принтеров нет.
    """
    names = _all_printers()
    if explicit and explicit in names:
        det = _printer_details(explicit)
        if _is_virtual_printer(explicit, det.get("port", "")):
            print(f"⚠ Выбранный «{explicit}» [{det.get('port') or '?'}] — виртуальный принтер "
                  f"(PDF/XPS/файл). Игнорирую и выбираю физический: RAW-задание на виртуальный "
                  f"даёт битый PDF в «Документах», а не этикетку")
        else:
            return explicit
    # принтер, выбранный вручную ранее (в тестовой форме) — он же для магазина
    saved = (_load_prefs() or {}).get(kind, "")
    if saved and saved in names:
        if _is_virtual_printer(saved, _printer_details(saved).get("port", "")):
            print(f"🧹 Сохранённый принтер «{saved}» — виртуальный/печать в файл: "
                  f"выбрасываю из настроек и беру физический")
            _forget_pref(kind)
        else:
            print(f"📌 Сохранённый принтер для «{kind}»: {saved}")
            return saved
    default = _default_printer()

    physical, skipped = [], []
    for idx, n in enumerate(names):
        det = _printer_details(n)
        port = det.get("port", "")
        if _is_virtual_printer(n, port):
            skipped.append(f"{n} [{port or '?'}]{' — виртуальный' if not _is_file_port(port) else ' — печать в файл'}")
            continue
        physical.append((n, port, _port_rank(port), idx))

    if skipped:
        print("⚠ Пропущены (виртуальные / печать в файл): " + "; ".join(skipped))

    # физические — от лучшего порта к худшему
    physical.sort(key=lambda x: (x[2], x[3]))

    def pick() -> list:
        cands = [x for x in physical if _family_match(x[0], kind)]
        # принтер по умолчанию Windows — самый надёжный сигнал: им печатают
        cands.sort(key=lambda x: (0 if x[0] == default else 1, x[2], x[3]))
        return cands

    if kind == "tlp100":
        # EPL-принтер предпочитаем без «- ZPL»
        cands = [x for x in pick() if "- zpl" not in x[0].lower()] or pick()
    else:  # lp58 — TSPL
        cands = pick()

    if not cands:
        print(f"⚠ Принтер «{kind}» не найден. Доступные: " + ("; ".join(
            f"{p['name']} [{p['port'] or '?'}]" + (" (ФАЙЛ)" if p["file_port"] else "")
            for p in _printers_report()) or "нет"))
        print(f"⚠ Задание формата {'EPL (TLP100)' if kind == 'tlp100' else 'TSPL (LP58)'} "
              f"не будет отправлено: нет принтера этого семейства (чужой язык печати не подставляем)")
        return ""

    chosen = cands[0][0]
    if len(cands) > 1:
        print("🖨 Кандидаты: " + "; ".join(
            f"{n} [{p}]{' ← ВЫБРАН' if n == chosen else ''}" for n, p, _, _ in cands))
    return chosen


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
    print(f"📨 TLP100 [{printer_name}] ×{copies}: {qr_text}")
    return ok


def print_lp58(qr_text: str, title: str = "", subtitle: str = "", printer_name: str = "MPRINT LP58 EVA", copies: int = 1, lines=None) -> bool:
    ok = True
    for _ in range(max(1, copies)):
        if not _raw_printer(printer_name, build_lp58(qr_text, title, subtitle, lines)):
            ok = False
    print(f"📨 LP58 [{printer_name}] ×{copies}: {qr_text}")
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
    try:
        det = _printers_report()
        resolved = {"tlp100": resolve_printer("tlp100"), "lp58": resolve_printer("lp58")}
    except Exception as e:
        return {"printers": [], "error": str(e)}
    label_mm = {}
    label_dpi = {}
    for kind, name in resolved.items():
        if not name:
            continue
        w, h = _driver_label_mm(name)
        if w and h:
            label_mm[kind] = [w, h]
        d = _driver_dpi(name)
        if d:
            label_dpi[kind] = d
    return {"printers": [d["name"] for d in det], "details": det,
            "resolved": resolved, "saved": _load_prefs(),
            "label_mm": label_mm, "label_dpi": label_dpi}


@app.post("/printer-config/reset")
def reset_printer_config():
    """Сбросить запомненный вручную принтер (например, был выбран виртуальный PDF)."""
    path = _config_path()
    try:
        if os.path.exists(path):
            os.remove(path)
        print("🧹 Сохранённый выбор принтера сброшен (printer-config.json удалён)")
        return {"ok": True, "message": "Сохранённый выбор принтера сброшен — снова работает автоподбор"}
    except Exception as e:
        return {"ok": False, "message": f"не удалось сбросить: {e}"}


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
    explicit = (req.printer_name or "").strip()
    ignored_explicit = ""
    if explicit:
        det = _printer_details(explicit)
        if _is_virtual_printer(explicit, det.get("port", "")):
            ignored_explicit = explicit
            print(f"⚠ Ручной выбор «{explicit}» [{det.get('port') or '?'}] — виртуальный принтер "
                  f"(PDF/XPS/файл): игнорирую, печатаю на физический, иначе вышел бы битый PDF в «Документах»")
            explicit = ""
    if explicit:
        # ручной выбор принтера: формат по имени, ровно одно задание, без дублей
        kind = "lp58" if ("lp58" in explicit.lower() or "eva" in explicit.lower()) else "tlp100"
        _save_pref(kind, explicit)
        if kind == "tlp100":
            tlp_name, lp_name = explicit, ""
            results["tlp100"] = print_tlp100(qr_text, title, subtitle, printer_name=explicit,
                                             copies=req.copies or 1, lines=lines)
        else:
            tlp_name, lp_name = "", explicit
            results["lp58"] = print_lp58(qr_text, title, subtitle, printer_name=explicit,
                                         copies=req.copies or 1, lines=lines)
    else:
        tlp_name = resolve_printer("tlp100")
        lp_name = resolve_printer("lp58")
        if req.printer in ("tlp100", "both") and tlp_name:
            results["tlp100"] = print_tlp100(qr_text, title, subtitle, printer_name=tlp_name, copies=req.copies or 1, lines=lines)
        if req.printer in ("lp58", "both") and lp_name:
            results["lp58"] = print_lp58(qr_text, title, subtitle, printer_name=lp_name, copies=req.copies or 1, lines=lines)
    if not results:
        results["error"] = "Физический принтер не найден (есть только виртуальные или печать в файл)"
        print("⚠ Нечего печатать. Принтеры: " + ("; ".join(
            f"{d['name']} [{d['port'] or '?'}]" + (" (ФАЙЛ)" if d["file_port"] else "")
            for d in _printers_report()) or "нет"))
    ok = bool(results) and "error" not in results and all(results.values())
    print(f"📥 Задание {qr_text!r} → ok={ok}, TLP100={tlp_name or '—'}, LP58={lp_name or '—'}"
          + (f", проигнорирован «{ignored_explicit}» (виртуальный)" if ignored_explicit else ""))
    if ok:
        msg = "Напечатано на: " + ", ".join(x for x in (tlp_name, lp_name) if x)
        if ignored_explicit:
            msg += f". Выбранный «{ignored_explicit}» — виртуальный (PDF/XPS), печатать на него нельзя"
    else:
        msg = str(results.get("error") or "Не удалось отправить на принтер")
    return {"ok": ok, "results": results, "qr": qr_text,
            "printer_tlp100": tlp_name, "printer_lp58": lp_name,
            "ignored_printer": ignored_explicit,
            "message": msg,
            "printers": _printers_report()}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=5000)
