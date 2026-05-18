# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템
- 엑셀 데이터 처리 및 PDF/프린트 전담 (Excel Service)
"""
import os
import re
import tempfile
import shutil
import subprocess
import platform
from typing import Dict, Any, Optional

# config에서 마스터 설정 컬럼을 가져옵니다.
from config import MASTER_COLS

# --- [라이브러리 체크] ---
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import pythoncom
    import win32com.client as win32
    WIN32COM_AVAILABLE = True
except ImportError:
    pythoncom = None
    win32 = None
    WIN32COM_AVAILABLE = False

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


class ExcelManager:
    def __init__(self, app):
        # 메인 앱(DeliveryApp)의 기능이 필요할 때를 대비해 연결해 둡니다.
        self.app = app

    # ==========================================
    # 1. Pandas 데이터 파싱 (LS 엑셀 불러오기용)
    # ==========================================
    def normalize_dataframe_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = df.columns.astype(str).str.replace(r'\s+', '', regex=True)
        return df

    def read_ls_excel(self, filepath: str) -> pd.DataFrame:
        read_kwargs = {"dtype": object}
        try:
            df = pd.read_excel(filepath, engine="openpyxl", **read_kwargs)
        except TypeError:
            df = pd.read_excel(filepath, **read_kwargs)
        return self.normalize_dataframe_columns(df)

    def map_delivery_type(self, val: Any) -> str:
        if pd.isna(val):
            return "부자재"
        raw = str(val).strip()
        compact = raw.replace(" ", "")
        if compact.lower() in ["nan", "none", ""]:
            return "부자재"
        upper_val = compact.upper()

        if "QUICK" in upper_val or "퀵" in compact: return "퀵"
        if "SAMEDAY_WAVE2" in upper_val: return "SD2"
        if "DAWN_WAVE1" in upper_val: return "D1"
        if "DAWN_WAVE2" in upper_val: return "D2"
        if "DAWN_WAVE3" in upper_val: return "D3"
        if "DAWN_WAVE4" in upper_val: return "D4"
        if compact in ["NORMAL", "동대구화물터미널1(SH)", "동대구2(SH)", "서대구1(SH)", "남김해1(SH)", "대전2(SH)", "전라광주7HUB", "안성5HUB", "북김해1(SH)"]: return "간선"
        if "WAVE2" in upper_val or "타이완" in compact: return "W2"
        return compact

    def format_time_value(self, val: Any) -> str:
        if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat", "<na>"]:
            return ""
        try:
            return pd.to_datetime(str(val)).strftime("%H:%M")
        except Exception:
            if hasattr(val, 'strftime'):
                return val.strftime("%H:%M")
            m = re.search(r'(\d{1,2}):(\d{2})', str(val).strip())
            return f"{int(m.group(1)):02d}:{m.group(2)}" if m else str(val).strip()

    # ==========================================
    # 2. Openpyxl 엑셀 조작 (마스터_양식설정 불러오기)
    # ==========================================
    def load_master_config(self, master_path: str) -> Dict[str, Dict[str, str]]:
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("openpyxl 라이브러리가 필요합니다.")
        wb = load_workbook(master_path, data_only=True)
        ws = wb.active
        result: Dict[str, Dict[str, str]] = {}

        header_row_idx = None
        header_col_map: Dict[str, int] = {}
        for ri in range(1, min(ws.max_row, 20) + 1):
            row_vals = [str(ws.cell(ri, ci).value or "").strip() for ci in range(1, ws.max_column + 1)]
            matched = {col: row_vals.index(col) + 1 for col in MASTER_COLS if col in row_vals}
            if len(matched) >= 3:
                header_row_idx = ri
                header_col_map = matched
                break

        if header_row_idx is None:
            wb.close()
            raise ValueError(f"마스터_양식설정.xlsx 에서 헤더를 찾지 못했습니다.\n첫 행에 다음 컬럼들이 있어야 합니다:\n{', '.join(MASTER_COLS)}")

        camp_col = 1
        for ci in range(1, ws.max_column + 1):
            v = str(ws.cell(header_row_idx, ci).value or "").strip()
            if "캠프" in v:
                camp_col = ci
                break

        for ri in range(header_row_idx + 1, ws.max_row + 1):
            camp_name = str(ws.cell(ri, camp_col).value or "").strip()
            if not camp_name: continue
            cfg: Dict[str, str] = {}
            for col_name, col_idx in header_col_map.items():
                val = ws.cell(ri, col_idx).value
                cfg[col_name] = str(val).strip() if val is not None else ""
            result[camp_name] = cfg

        wb.close()
        return result

    # ==========================================
    # 3. PDF 변환 및 인쇄 기능
    # ==========================================
    def convert_xlsx_to_pdf_lo(self, xlsx_path: str) -> str:
        """LibreOffice를 사용해 엑셀을 PDF로 변환합니다."""
        soffice = shutil.which("libreoffice") or shutil.which("soffice")
        if not soffice:
            raise RuntimeError("LibreOffice를 찾지 못했습니다.")
        out_dir = tempfile.mkdtemp(prefix="dispatch_pdf_")
        cmd = [soffice, "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", out_dir, os.path.abspath(xlsx_path)]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if proc.returncode != 0:
            raise RuntimeError((proc.stderr or proc.stdout or "PDF 변환 실패").strip())
        base = os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf"
        expected = os.path.join(out_dir, base)
        if os.path.exists(expected):
            return expected
        candidates = [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.endswith(".pdf")]
        if candidates: return candidates[0]
        raise RuntimeError("변환된 PDF 파일을 찾지 못했습니다.")

    def save_image_as_pdf(self, image, output_pdf_path: Optional[str] = None) -> str:
        """Pillow 이미지를 PDF로 저장합니다."""
        if not PILLOW_AVAILABLE:
            raise RuntimeError("Pillow 라이브러리가 필요합니다.")
        if output_pdf_path is None:
            temp_fd, output_pdf_path = tempfile.mkstemp(prefix="dispatch_", suffix=".pdf")
            os.close(temp_fd)
        rgb_image = image.convert("RGB")
        rgb_image.save(output_pdf_path, "PDF", resolution=150.0)
        return output_pdf_path

    def send_pdf_to_printer(self, pdf_path: str) -> None:
        """PDF 파일을 OS 기본 프린터로 인쇄합니다."""
        if platform.system() == "Windows" and hasattr(os, "startfile"):
            try:
                os.startfile(pdf_path, "print")
                return
            except OSError:
                from tkinter import messagebox
                messagebox.showwarning("수동 인쇄 안내", "PC에 자동 인쇄 프로그램이 설정되어 있지 않아 뷰어를 엽니다.\n직접 인쇄(Ctrl+P)를 눌러주세요.")
                os.startfile(pdf_path)
                return

        for cmd_name in ["lp", "lpr"]:
            if shutil.which(cmd_name):
                proc = subprocess.run([cmd_name, pdf_path], capture_output=True, text=True, timeout=30)
                if proc.returncode == 0: return
        raise RuntimeError("PDF 인쇄 명령을 사용할 수 없습니다.")
