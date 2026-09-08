from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import win32print
from typing import Literal, Optional

app = FastAPI(title="Принтер-сервер для ПВЗ", version="1.0")

# Разрешаем запросы с сайта
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # В продакшене замени на свой домен
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# --- Модели данных для запросов ---
class PrintRequest(BaseModel):
    id: str
    printer: Optional[Literal["tlp100", "lp58", "both"]] = "both"


# --- Функции печати ---
def print_tlp100(product_id: str, printer_name: str = "MPRINT Terra Nova TLP100") -> bool:
    """
    Печать на TLP100 (EPL)
    """
    epl_data = f"""
N
q400
Q300,24
A320,180,1,2,1,1,N,"ID:{product_id}"
b20,15,Q,m2,s13,eL,"{product_id}"
P1
"""
    print(f"📨 Отправляю на TLP100: {product_id}")
    print(epl_data)

    try:
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            win32print.StartDocPrinter(hprinter, 1, ("TLP100_Final", None, "RAW"))
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, epl_data.encode('cp1251'))
            win32print.EndPagePrinter(hprinter)
            win32print.EndDocPrinter(hprinter)
            print("✅ TLP100: QR-код отправлен!")
            return True
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        print(f"❌ TLP100 ошибка: {e}")
        return False


def print_lp58(product_id: str, printer_name: str = "MPRINT LP58 EVA") -> bool:
    """
    Печать на LP58 EVA (TSPL)
    """
    tspl_data = f"""
SIZE 58 mm, 40 mm
GAP 2 mm, 0 mm
CLS
TEXT 380,180,"3",90,1,1,"ID:{product_id}"
QRCODE 30,16,H,13,A,0,M2,S7,"{product_id}"
PRINT 1
"""
    print(f"📨 Отправляю на LP58 EVA: {product_id}")
    print(tspl_data)

    try:
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            win32print.StartDocPrinter(hprinter, 1, ("LP58_Final", None, "RAW"))
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, tspl_data.encode('ascii'))
            win32print.EndPagePrinter(hprinter)
            win32print.EndDocPrinter(hprinter)
            print("✅ LP58 EVA: QR-код отправлен!")
            return True
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        print(f"❌ LP58 EVA ошибка: {e}")
        return False


def print_both(product_id: str) -> bool:
    """Печатает на ОБОИХ принтерах"""
    print(f"\n🖨️🖨️ Печатаю на двух принтерах для товара: {product_id}")
    print("=" * 50)

    result1 = print_tlp100(product_id)
    result2 = print_lp58(product_id)

    if result1 and result2:
        print("✅✅ Оба принтера напечатали успешно!")
    else:
        print("⚠️ Один из принтеров не сработал")

    return result1 and result2


# --- API Эндпоинты ---
@app.get("/")
async def root():
    return {
        "service": "Принтер-сервер для ПВЗ",
        "status": "running",
        "endpoints": {
            "/ping": "GET - проверка статуса",
            "/printers": "GET - список принтеров",
            "/print": "POST - печать этикетки"
        }
    }


@app.get("/ping")
async def ping():
    return {"status": "ok", "message": "Принтер-сервер работает"}


@app.get("/printers")
async def list_printers():
    """Список доступных принтеров в системе"""
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    return {
        "printers": [p[2] for p in printers],
        "count": len(printers)
    }


@app.post("/print")
async def print_label(request: PrintRequest):
    """
    Печать этикетки с QR-кодом

    - **id**: ID товара (обязательно)
    - **printer**: tlp100, lp58 или both (по умолчанию both)
    """
    product_id = request.id
    printer = request.printer

    print(f"\n📨 Получен запрос на печать: ID={product_id}, принтер={printer}")

    try:
        if printer == "tlp100":
            success = print_tlp100(product_id)
        elif printer == "lp58":
            success = print_lp58(product_id)
        else:  # both
            success = print_both(product_id)

        if success:
            return {
                "status": "ok",
                "message": f"Печать для товара {product_id} выполнена",
                "printer": printer
            }
        else:
            raise HTTPException(status_code=500, detail="Ошибка печати")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# --- Запуск ---
if __name__ == "__main__":
    import uvicorn

    print("🖨️ Принтер-сервер запущен на http://localhost:5000")
    print("📋 Документация API: http://localhost:5000/docs")
    print("📍 POST /print - отправить { 'id': '12345', 'printer': 'both' }")
    print("📍 GET /ping - проверка статуса")
    print("📍 GET /printers - список принтеров")
    print("=" * 50)

    uvicorn.run(
        "server:app",
        host="0.0.0.0",
        port=5000,
        reload=False
    )