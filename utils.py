# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템
- 공통 유틸리티 함수 (Utils)
"""
import os
import sys
import json
from typing import Dict, Any

# 작성해둔 config 모듈에서 기본 설정값을 가져옵니다.
from config import DEFAULT_SETTINGS

# 이미지 로드에 필요한 라이브러리 체크 (Pillow)
try:
    from PIL import Image, ImageTk
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    print("경고: Pillow 라이브러리가 설치되지 않았습니다. 텍스트 로고로 대체됩니다.")

def resource_path(relative_path: str) -> str:
    """PyInstaller 빌드 시 임시 폴더(sys._MEIPASS) 또는 로컬 경로를 반환합니다."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_hq_image(path: str, target_height: int) -> Any:
    """고품질 이미지를 로드하고 비율에 맞춰 리사이징합니다."""
    if not PILLOW_AVAILABLE or not os.path.exists(path):
        return None
    try:
        img = Image.open(path)
        h_percent = (target_height / float(img.size[1]))
        w_size = int((float(img.size[0]) * float(h_percent)))
        resized_img = img.resize((w_size, target_height), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(resized_img)
    except Exception:
        return None

def deep_copy_default_settings() -> Dict[str, Any]:
    """기본 설정값의 깊은 복사본을 반환하여 원본 훼손을 방지합니다."""
    return json.loads(json.dumps(DEFAULT_SETTINGS, ensure_ascii=False))