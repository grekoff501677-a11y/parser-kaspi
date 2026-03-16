"""
Kaspi Parser — Cloud API
FastAPI backend for CRM integration.

Usage from CRM:
    1. POST /api/parse/pickup   → {task_id}
    2. GET  /api/tasks/{id}     → poll status/progress
    3. GET  /api/tasks/{id}/excel  → download result
"""

import uuid
from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
from fastapi import BackgroundTasks, FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from kaspi_client import KaspiClient

app = FastAPI(title="Kaspi Parser API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # lock down to your CRM domain in production
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory task store (swap for Redis in production)
_tasks: dict[str, dict] = {}

CITIES = {
    "astana": "710000000",
    "pavlodar": "551010000",
}


# ── Pydantic models ────────────────────────────────────────────────────────────

class PickupRequest(BaseModel):
    category: str               # "rims" | "wheels" | "auto%20parts"
    city: str                   # "astana" | "pavlodar"
    merchant_filter: Optional[str] = None   # e.g. "1 ALTRA AUTO"
    extra_filters: Optional[dict] = None    # e.g. {"Rims*Type": "литые"}


class OurShopRequest(BaseModel):
    category: str
    merchant_id: str = "Kama"


# ── Task helpers ───────────────────────────────────────────────────────────────

def _create_task() -> str:
    tid = uuid.uuid4().hex[:8]
    _tasks[tid] = {
        "id": tid,
        "status": "pending",   # pending | running | done | error
        "progress": 0,
        "total": 0,
        "message": "Ожидание...",
        "results": [],
        "created_at": datetime.now().isoformat(),
    }
    return tid


def _t(tid: str) -> dict:
    return _tasks[tid]


# ── Background workers ─────────────────────────────────────────────────────────

async def _worker_pickup(tid: str, req: PickupRequest):
    t = _t(tid)
    city_id = CITIES.get(req.city.lower(), req.city)

    try:
        t["status"] = "running"
        t["message"] = "Собираю список товаров из категории..."

        async with KaspiClient(concurrency=15) as client:

            async def on_page(page: int, total: int):
                t["message"] = f"Страница {page}/{total or '?'} — собираю ссылки"

            items = await client.get_category_skus(
                req.category,
                city_id,
                merchant_id=req.merchant_filter,
                extra_filters=req.extra_filters,
                on_page=on_page,
            )

            t["total"] = len(items)
            t["message"] = f"Найдено {len(items)} товаров. Проверяю самовывоз..."

            async def on_progress(checked: int):
                t["progress"] = checked
                found = sum(1 for _ in t["results"])  # updated inside bulk_check_pickup
                t["message"] = f"Проверено {checked}/{t['total']}"

            results = await client.bulk_check_pickup(
                items, city_id, on_progress=on_progress
            )

            t["results"] = results
            t["status"] = "done"
            t["progress"] = len(items)
            t["message"] = f"Готово: {len(results)} товаров с самовывозом из {len(items)}"

    except Exception as e:
        t["status"] = "error"
        t["message"] = f"Ошибка: {e}"


async def _worker_our_shop(tid: str, req: OurShopRequest):
    t = _t(tid)

    try:
        t["status"] = "running"
        t["message"] = "Собираю товары нашего магазина..."

        async with KaspiClient(concurrency=15) as client:

            async def on_progress(processed: int, total: int):
                t["progress"] = processed
                t["total"] = total
                t["message"] = f"Определяю город: {processed}/{total}"

            results = await client.get_merchant_products_with_city(
                req.category,
                req.merchant_id,
                on_progress=on_progress,
            )

            t["results"] = results
            t["status"] = "done"
            pvl = sum(1 for r in results if r["Город"] == "Павлодар")
            ast = len(results) - pvl
            t["message"] = f"Готово: {len(results)} товаров (Павлодар: {pvl}, Астана: {ast})"

    except Exception as e:
        t["status"] = "error"
        t["message"] = f"Ошибка: {e}"


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.post("/api/parse/pickup", summary="Парсинг самовывоза конкурентов")
async def parse_pickup(req: PickupRequest, bg: BackgroundTasks):
    """
    Запускает фоновую задачу: собирает товары из категории,
    проверяет наличие самовывоза в указанном городе.

    Пример запроса из CRM:
        POST /api/parse/pickup
        {"category": "rims", "city": "pavlodar"}
    """
    tid = _create_task()
    bg.add_task(_worker_pickup, tid, req)
    return {"task_id": tid}


@app.post("/api/parse/our-shop", summary="Парсинг нашего магазина")
async def parse_our_shop(req: OurShopRequest, bg: BackgroundTasks):
    """
    Запускает фоновую задачу: собирает все товары нашего магазина
    и классифицирует по городу (Павлодар / Астана).
    """
    tid = _create_task()
    bg.add_task(_worker_our_shop, tid, req)
    return {"task_id": tid}


@app.get("/api/tasks/{tid}", summary="Статус задачи (polling)")
async def get_task_status(tid: str):
    """
    Используйте этот endpoint для polling из CRM каждые 2-3 секунды.

    Статусы: pending → running → done | error
    Когда status == "done" — запрашивайте /excel или /json.
    """
    t = _tasks.get(tid)
    if not t:
        raise HTTPException(404, "Задача не найдена")
    return {
        "id": t["id"],
        "status": t["status"],
        "progress": t["progress"],
        "total": t["total"],
        "message": t["message"],
        "count": len(t["results"]),
        "created_at": t["created_at"],
    }


@app.get("/api/tasks/{tid}/json", summary="Результаты в JSON")
async def get_results_json(tid: str):
    t = _tasks.get(tid)
    if not t:
        raise HTTPException(404, "Задача не найдена")
    if t["status"] != "done":
        raise HTTPException(400, f"Задача не завершена (статус: {t['status']})")
    return {"count": len(t["results"]), "results": t["results"]}


@app.get("/api/tasks/{tid}/excel", summary="Скачать Excel")
async def download_excel(tid: str):
    t = _tasks.get(tid)
    if not t:
        raise HTTPException(404, "Задача не найдена")
    if t["status"] != "done":
        raise HTTPException(400, "Задача не завершена")

    df = pd.DataFrame(t["results"])
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Товары")
    buf.seek(0)

    fname = f"kaspi_{tid}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"},
    )


@app.get("/api/tasks", summary="Список последних задач")
async def list_tasks():
    return sorted(
        [
            {
                "id": t["id"],
                "status": t["status"],
                "message": t["message"],
                "count": len(t["results"]),
                "created_at": t["created_at"],
            }
            for t in _tasks.values()
        ],
        key=lambda x: x["created_at"],
        reverse=True,
    )


@app.get("/health")
async def health():
    return {"ok": True}
