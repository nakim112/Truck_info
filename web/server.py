# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템 - 웹 버전
- FastAPI 서버 (REST API + 정적 프론트엔드)

실행:
    uvicorn web.server:app --host 0.0.0.0 --port 8000
또는:
    python run_web.py
"""
from __future__ import annotations

import os
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import logic, store

_HERE = os.path.dirname(os.path.abspath(__file__))
_STATIC = os.path.join(_HERE, "static")

app = FastAPI(title="INC4 HUB 차량 정보 시스템 (Web)", version="1.0")


class RecordIn(BaseModel):
    delivery_type: str
    camp: str = ""
    dock: str = ""
    hocha: str = ""
    car_num: str = ""
    phone: str = ""
    time: str = ""


class StatusIn(BaseModel):
    status: str


class RecordPatch(BaseModel):
    delivery_type: Optional[str] = None
    camp: Optional[str] = None
    dock: Optional[str] = None
    hocha: Optional[str] = None
    car_num: Optional[str] = None
    phone: Optional[str] = None
    time: Optional[str] = None


def _decorate(rec: dict) -> dict:
    out = dict(rec)
    out["emoji"] = logic.status_emoji(rec.get("status", ""))
    return out


@app.get("/api/config")
def get_config() -> dict:
    return logic.config_payload()


@app.get("/api/records")
def get_records(board: bool = False) -> dict:
    records = [_decorate(r) for r in store.list_records(board=board)]
    return {"records": records, "count": len(records)}


@app.post("/api/records")
def create_record(payload: RecordIn) -> dict:
    try:
        rec = store.add_record(
            payload.delivery_type,
            payload.camp,
            payload.dock,
            payload.hocha,
            payload.car_num,
            payload.phone,
            payload.time,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return _decorate(rec)


@app.patch("/api/records/{record_id}/status")
def patch_status(record_id: str, payload: StatusIn) -> dict:
    if payload.status not in logic.STATUS_LIST:
        raise HTTPException(status_code=400, detail="알 수 없는 상태값입니다.")
    rec = store.update_status(record_id, payload.status)
    if rec is None:
        raise HTTPException(status_code=404, detail="레코드를 찾을 수 없습니다.")
    return _decorate(rec)


@app.patch("/api/records/{record_id}")
def patch_record(record_id: str, payload: RecordPatch) -> dict:
    fields = {k: v for k, v in payload.dict().items() if v is not None}
    rec = store.update_record(record_id, fields)
    if rec is None:
        raise HTTPException(status_code=404, detail="레코드를 찾을 수 없습니다.")
    return _decorate(rec)


@app.delete("/api/records/{record_id}")
def remove_record(record_id: str) -> dict:
    if not store.delete_record(record_id):
        raise HTTPException(status_code=404, detail="레코드를 찾을 수 없습니다.")
    return {"deleted": record_id}


@app.delete("/api/records")
def remove_all() -> dict:
    return {"cleared": store.clear_all()}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(os.path.join(_STATIC, "index.html"))


@app.get("/board")
def board() -> FileResponse:
    return FileResponse(os.path.join(_STATIC, "board.html"))


@app.get("/healthz")
def healthz() -> JSONResponse:
    return JSONResponse({"status": "ok"})


# 정적 자산 (JS/CSS). 마지막에 마운트해 API 라우트를 가리지 않게 합니다.
app.mount("/static", StaticFiles(directory=_STATIC), name="static")
