# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템 - 웹 버전
- 비즈니스 로직 (데스크톱 앱의 규칙을 그대로 이식)

데스크톱 앱(truck_info_ui_airport.py)의 배차 타입 보정, 입력 검증,
표시 시간대 판정, 정렬 규칙을 서버 측에서 동일하게 재현합니다.
"""
from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Optional, Tuple

# 리포지토리 루트를 import 경로에 추가하여 기존 config.py를 재사용합니다.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from config import (  # noqa: E402
    DEFAULT_CAMPS,
    DEFAULT_DOCKS,
    DEFAULT_HOCHAS,
    STATUS_LIST,
    DISPATCH_TABS,
    PHONE_PATTERN,
    TIME_PATTERN,
    DEFAULT_SETTINGS,
)

VISIBILITY_SCHEDULE: Dict[str, List[str]] = DEFAULT_SETTINGS["schedules"]["visibility"]

# 상태 → 표시용 이모지 (데스크톱 apply_status 규칙과 동일)
STATUS_EMOJI: Dict[str, str] = {
    "접안": "🟢",
    "미접안": "🟡",
    "수배중": "⏳",
    "재수배": "🔴",
    "차량변경": "🔴",
}


def status_emoji(status: str) -> str:
    return STATUS_EMOJI.get(status, "🟡")


def delivery_types() -> List[str]:
    """등록 가능한 Delivery Type 목록 (캠프 매핑 키 순서 유지)."""
    return list(DEFAULT_CAMPS.keys())


def camps_for_type(d_type: str) -> List[str]:
    return list(DEFAULT_CAMPS.get(d_type, []))


def override_delivery_type(d_type: str, camp: str, hocha: str) -> str:
    """데스크톱 _override_delivery_type 이식: 캠프/호차에 따른 타입 자동 보정."""
    camp_clean = str(camp).replace(" ", "").upper()
    hocha_clean = str(hocha).replace(" ", "")

    if d_type == "SD2" and "인천3" in camp_clean and hocha_clean in ["2호차", "3호차", "2", "3"]:
        return "SD3"
    if d_type == "W2":
        ganseon_targets = ["김해4HUB", "남김해1(SH)", "동대구화물터미널", "서대구1(SH)", "전라광주7HUB"]
        for target in ganseon_targets:
            if target.replace(" ", "").upper() in camp_clean:
                return "간선"
    return d_type


def normalize_time(time_val: str) -> str:
    """'0830' → '08:30' 형태로 정규화."""
    time_val = (time_val or "").strip()
    if len(time_val) == 4 and time_val.isdigit():
        return f"{time_val[:2]}:{time_val[2:]}"
    return time_val


def validate_record(
    d_type: str,
    camp: str,
    dock: str,
    hocha: str,
    car_num: str,
    phone: str,
    time_val: str,
    existing: Optional[List[dict]] = None,
) -> Tuple[bool, str, str]:
    """데스크톱 _validate_inputs 이식. (ok, 메시지, 정규화된 시간) 반환."""
    d_type = (d_type or "").strip()
    hocha = (hocha or "").strip()
    car_num = (car_num or "").strip()
    phone = (phone or "").strip()
    time_val = (time_val or "").strip()

    if d_type == "KPP":
        camp, dock = "-", "-"
    else:
        camp, dock = (camp or "").strip(), (dock or "").strip()

    if not all([d_type, camp, dock, hocha]):
        return False, "필수 항목(Type, 캠프, 도크, 호차)을 모두 입력해주세요.", time_val

    if phone and not PHONE_PATTERN.match(phone):
        return False, "전화번호 형식이 올바르지 않습니다.", time_val

    if time_val:
        time_val = normalize_time(time_val)
        if not TIME_PATTERN.match(time_val):
            return False, "시간은 HH:MM 형식으로 입력하세요.", time_val

    mapped = override_delivery_type(d_type, camp, hocha)
    car_norm = car_num.replace(" ", "")
    for rec in existing or []:
        if (
            rec.get("delivery_type") == mapped
            and rec.get("camp") == camp
            and rec.get("car_num", "").replace(" ", "") == car_norm
            and car_norm
        ):
            return False, "동일 차량이 이미 등록되어 있습니다.", time_val

    return True, "", time_val


def _continuous_mins(time_str: str) -> int:
    try:
        h, m = map(int, time_str.split(":"))
        return (h * 60 + m) + (24 * 60 if h < 6 else 0)
    except Exception:
        return 0


def is_time_in_range(start_str: str, end_str: str, now: Optional[datetime] = None) -> bool:
    """데스크톱 _is_time_in_range 이식 (새벽 구간을 익일로 환산)."""
    now = now or datetime.now()
    try:
        now_mins = now.hour * 60 + now.minute + (24 * 60 if now.hour < 6 else 0)
        return _continuous_mins(start_str) <= now_mins < _continuous_mins(end_str)
    except Exception:
        return True


def is_visible(rec: dict, now: Optional[datetime] = None) -> bool:
    """표시 시간대 스케줄에 따라 전광판 노출 여부 판정."""
    d_type = rec.get("delivery_type", "")
    if d_type in VISIBILITY_SCHEDULE:
        return is_time_in_range(*VISIBILITY_SCHEDULE[d_type], now=now)
    return True


def _status_priority(status: str) -> int:
    s = str(status).strip()
    if "차량변경" in s:
        return 0
    if "미접안" in s:
        return 1
    if "접안" in s:
        return 2
    return 3


def _hocha_num(hocha: str) -> int:
    m = re.search(r"\d+", str(hocha))
    return int(m.group()) if m else 999


def sort_records(records: List[dict]) -> List[dict]:
    """데스크톱 _sort_treeview 이식: 상태 → 캠프 → 호차 순."""
    return sorted(
        records,
        key=lambda r: (
            _status_priority(r.get("status", "")),
            r.get("camp", ""),
            _hocha_num(r.get("hocha", "")),
        ),
    )


def config_payload() -> dict:
    """프론트엔드 드롭다운 구성을 위한 설정 묶음."""
    return {
        "delivery_types": delivery_types(),
        "camps_by_type": {k: list(v) for k, v in DEFAULT_CAMPS.items()},
        "docks": list(DEFAULT_DOCKS),
        "hochas": list(DEFAULT_HOCHAS),
        "statuses": list(STATUS_LIST),
        "dispatch_tabs": list(DISPATCH_TABS),
        "visibility": {k: list(v) for k, v in VISIBILITY_SCHEDULE.items()},
    }
