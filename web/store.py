# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템 - 웹 버전
- 레코드 영속화 (JSON 파일) 및 CRUD

데이터는 TRUCK_INFO_DATA 환경변수로 지정한 JSON 파일(기본: web/data.json)에
저장됩니다. 동시 접근을 막기 위해 간단한 스레드 락을 사용합니다.
"""
from __future__ import annotations

import json
import os
import threading
import time
import uuid
from typing import Dict, List, Optional

from . import logic

_DEFAULT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
DATA_PATH = os.environ.get("TRUCK_INFO_DATA", _DEFAULT_PATH)

_lock = threading.RLock()


def _read() -> List[dict]:
    if not os.path.exists(DATA_PATH):
        return []
    try:
        with open(DATA_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _write(records: List[dict]) -> None:
    os.makedirs(os.path.dirname(DATA_PATH), exist_ok=True)
    tmp = f"{DATA_PATH}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    os.replace(tmp, DATA_PATH)


def list_records(board: bool = False) -> List[dict]:
    """정렬된 레코드 목록. board=True이면 표시 시간대 필터를 적용."""
    with _lock:
        records = _read()
    if board:
        records = [r for r in records if logic.is_visible(r)]
    return logic.sort_records(records)


def add_record(
    delivery_type: str,
    camp: str,
    dock: str,
    hocha: str,
    car_num: str = "",
    phone: str = "",
    time_val: str = "",
) -> dict:
    """검증 후 레코드를 추가하고 저장. 실패 시 ValueError."""
    with _lock:
        records = _read()
        ok, msg, norm_time = logic.validate_record(
            delivery_type, camp, dock, hocha, car_num, phone, time_val, existing=records
        )
        if not ok:
            raise ValueError(msg)

        if delivery_type == "KPP":
            camp, dock = "-", "-"
        mapped = logic.override_delivery_type(delivery_type, camp, hocha)

        rec = {
            "id": uuid.uuid4().hex,
            "status": "미접안",
            "delivery_type": mapped,
            "camp": camp.strip(),
            "dock": dock.strip(),
            "hocha": hocha.strip(),
            "car_num": car_num.strip(),
            "phone": phone.strip(),
            "time": norm_time,
            "created_ms": int(time.time() * 1000),
        }
        records.append(rec)
        _write(records)
        return rec


def update_status(record_id: str, status: str) -> Optional[dict]:
    """상태 변경. '접안'이면 접안시간을 현재 시각으로 설정(데스크톱 동작과 동일)."""
    with _lock:
        records = _read()
        target = None
        for rec in records:
            if rec.get("id") == record_id:
                rec["status"] = status
                if status == "접안":
                    rec["time"] = time.strftime("%H:%M")
                target = rec
                break
        if target is not None:
            _write(records)
        return target


def update_record(record_id: str, fields: Dict[str, str]) -> Optional[dict]:
    """편집 가능한 필드 수정 (캠프/도크/호차/차량번호/전화번호/시간/타입)."""
    allowed = {"delivery_type", "camp", "dock", "hocha", "car_num", "phone", "time"}
    with _lock:
        records = _read()
        target = None
        for rec in records:
            if rec.get("id") == record_id:
                for k, v in fields.items():
                    if k in allowed:
                        rec[k] = (v or "").strip()
                if "time" in fields:
                    rec["time"] = logic.normalize_time(rec.get("time", ""))
                target = rec
                break
        if target is not None:
            _write(records)
        return target


def delete_record(record_id: str) -> bool:
    with _lock:
        records = _read()
        new_records = [r for r in records if r.get("id") != record_id]
        changed = len(new_records) != len(records)
        if changed:
            _write(new_records)
        return changed


def clear_all() -> int:
    with _lock:
        count = len(_read())
        _write([])
        return count
