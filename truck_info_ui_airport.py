# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템 v2.0
- 메인 애플리케이션 (Main UI)
"""
import json
from copy import copy
import os
import random
import sys
import re
import ctypes
import platform
import shutil
import subprocess
import tempfile
import io
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
from tkinter import font as tkfont
from typing import Dict, List, Any, Tuple, Optional

# --- 데이터 수집 및 엑셀, PDF 제어용 라이브러리 체크 ---
try:
    import pandas as pd

    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("경고: pandas 라이브러리가 설치되지 않았습니다. 'LS불러오기' 기능이 제한될 수 있습니다.")

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, range_boundaries
    from openpyxl.styles import Font

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("경고: openpyxl 라이브러리가 설치되지 않았습니다. 출차서 미리보기/출력이 제한될 수 있습니다.")

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont

    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    print("경고: Pillow 라이브러리가 설치되지 않았습니다. 텍스트 로고로 대체됩니다.")

try:
    import fitz

    PYMUPDF_AVAILABLE = True
except ImportError:
    fitz = None
    PYMUPDF_AVAILABLE = False

try:
    import pythoncom
    import win32com.client as win32

    WIN32COM_AVAILABLE = True
except ImportError:
    pythoncom = None
    win32 = None
    WIN32COM_AVAILABLE = False

# ==========================================
# 🚀 분리된 전용 모듈 불러오기
# ==========================================
from config import (
    LOGO_FILENAME, SETTINGS_FILE, DISPATCH_TEMPLATE_FILENAME,
    MASTER_CONFIG_FILENAME, GOIMOK_TEMPLATE_FILENAME,
    MASTER_COLS, COLUMNS, STATUS_LIST, PHONE_PATTERN, TIME_PATTERN, SEAL_PATTERN,
    DISPATCH_TABS, DISPATCH_HEADERS,
    DEFAULT_CAMPS, DEFAULT_DOCKS, DEFAULT_HOCHAS, DEFAULT_SETTINGS
)
from utils import resource_path, load_hq_image, deep_copy_default_settings
from ui_palette import (
    DARK_PRESET_NAME,
    LEGACY_PRESET_NAME,
    MONO_PRESET_NAME,
    PRESET_LABELS,
    PALETTE_PRESETS,
    WIN11_PRESET_NAME,
    build_main_palette,
    get_active_preset_name,
    set_active_preset_name,
)
from widgets import SplitFlapCell
from api_service import ApiManager, SELENIUM_AVAILABLE, REQUESTS_AVAILABLE
from excel_service import ExcelManager


# ==========================================
# [메인 애플리케이션 클래스]
# ==========================================
class DeliveryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._init_settings()

        self.root.title(self.app_title)
        self.root.geometry("1240x780")
        self.root.minsize(1040, 680)
        self.root.configure(bg="#E7EEF4")

        # ── 디자인 시스템  ──
        self.C = {
            "bg": "#E7EEF4",  # 전체 배경
            "white": "#FBFDFF",  # 카드/패널 배경
            "panel": "#FBFDFF",
            "toolbar": "#F6FAFD",
            "header": "#15344D",  # 딥 네이비
            "header_txt": "#F7FBFF",
            "header_muted": "#AFC3D4",
            "accent": "#5D91AE",  # 에어 블루
            "accent_soft": "#D9E8F1",
            "border": "#C8D8E3",
            "border_soft": "#D9E5ED",
            "border_focus": "#5E8FAF",
            "label": "#506474",
            "text": "#132838",
            "placeholder": "#7E95A7",
            # 버튼
            "btn_primary": "#2F6F8F",  # 메인 블루
            "btn_success": "#2D7C72",  # 차분한 청록
            "btn_danger": "#BE5C57",  # 저채도 레드
            "btn_warning": "#C69754",  # 샌드 골드
            "btn_neutral": "#758796",  # 스틸 그레이
            "btn_dark": "#244B67",  # 전광판/보조 네이비
            "btn_teal": "#3D8DA5",  # 출차서 버튼
            "btn_secondary": "#31506C",  # 설정/보조 액션
            # 상태 색상
            "status_ok": "#2B7E73",
            "status_warn": "#9C762E",
            "status_err": "#B65A55",
            # 테이블
            "row_odd": "#FBFDFF",
            "row_even": "#F3F8FC",
            "row_sel": "#DCEAF3",
            "heading": "#EAF2F7",
            "heading_active": "#DDE8F0",
        }
        self.C = build_main_palette()
        self.root.configure(bg=self.C["bg"])
        self._setup_styles()

        self.blink_state = False
        self.fs_rows = []
        self.fs_current_page = 0
        self.MAX_FS_ROWS = 15
        self.fs_page_timer_id = None
        self.flap_refresh_timer_id = None
        self._resize_timer = None

        self.detached_items = {}
        self.last_ls_filepath: Optional[str] = None

        self.dispatch_source_rows: Dict[str, List[Dict[str, Any]]] = {k: [] for k in DISPATCH_TABS}
        self.dispatch_tab_frames: Dict[str, tk.Frame] = {}
        self.dispatch_tab_meta: Dict[str, Dict[str, Any]] = {}
        self.dispatch_template_path = resource_path(DISPATCH_TEMPLATE_FILENAME)

        self.dispatch_registry: Dict[str, Dict[str, str]] = {}
        self.api_data_dict: Dict[str, Dict[str, str]] = {}

        # ── 외부 서비스 매니저 장착 ──
        self.api_manager = ApiManager()
        self.excel_manager = ExcelManager(self)

        self.api_auto_refresh_enabled = False
        self.api_refresh_timer_id = None
        self.api_current_url = ""
        self.api_run_date = ""  # 운행일자 (출차서 기입용)

        self.create_widgets()

        self._blink_timer()
        self._periodic_flap_refresh()
        self._visibility_scheduler()

        self.root.after(500, self.open_api_fetch_window)

    # ──────────────────────────────────────────────────────────────────────────
    # 스타일 설정
    # ──────────────────────────────────────────────────────────────────────────
    def _setup_styles(self):
        C = self.C
        s = ttk.Style()
        s.theme_use("clam")
        tree_font = self._get_treeview_font()

        s.configure("Treeview", background=C["white"], foreground=C["text"], fieldbackground=C["white"], borderwidth=0,
                    rowheight=36, font=tree_font)
        s.configure("Treeview.Heading", background=C["heading"], foreground=C["label"], relief="flat", borderwidth=0,
                    font=("맑은 고딕", 9, "bold"), padding=[8, 9])
        s.map("Treeview", background=[("selected", C["row_sel"])], foreground=[("selected", C["text"])])
        s.map("Treeview.Heading", background=[("active", C["heading_active"])], foreground=[("active", C["text"])])

        s.configure("TCombobox", fieldbackground=C["white"], background=C["white"], foreground=C["text"],
                    selectbackground=C["white"], selectforeground=C["text"], bordercolor=C["border"],
                    lightcolor=C["white"], darkcolor=C["border"], arrowcolor=C["accent"], relief="flat", padding=4,
                    font=("맑은 고딕", 10))
        s.map("TCombobox", fieldbackground=[("readonly", C["white"])], foreground=[("readonly", C["text"])],
              bordercolor=[("focus", C["border_focus"]), ("readonly", C["border"])],
              lightcolor=[("focus", C["accent_soft"])], darkcolor=[("focus", C["border_focus"])])

        s.configure("Vertical.TScrollbar", background=C["accent_soft"], troughcolor=C["bg"], arrowcolor=C["accent"],
                    relief="flat", borderwidth=0, width=11)
        s.map("Vertical.TScrollbar", background=[("active", C["accent"])])

        s.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab", background=C["bg"], foreground=C["label"], padding=[20, 10],
                    font=("맑은 고딕", 10, "bold"), borderwidth=0)
        s.map("TNotebook.Tab", background=[("selected", C["white"]), ("active", C["heading"])],
              foreground=[("selected", C["btn_primary"]), ("active", C["text"])])

    def _entry(self, parent, width=18, **kw):
        C = self.C
        return tk.Entry(parent, width=width, bg=C["white"], fg=C["text"], insertbackground=C["text"], relief="solid",
                        bd=1, disabledbackground=C["heading"], readonlybackground=C["heading"], highlightthickness=2,
                        highlightbackground=C["border_soft"], highlightcolor=C["border_focus"], font=("맑은 고딕", 10),
                        **kw)

    def _btn(self, parent, text, cmd, color=None, width=None, **kw):
        C = self.C
        if color is None: color = C["btn_primary"]
        btn = tk.Button(parent, text=text, command=cmd, bg=color, fg="white", activebackground=color,
                        activeforeground="white", relief="flat", cursor="hand2", font=("맑은 고딕", 10, "bold"), bd=0,
                        highlightthickness=0, padx=16, pady=7, **kw)
        if width: btn.config(width=width)

        def _shade(hex_color: str, delta: float = -0.10) -> str:
            import colorsys
            r, g, b = int(hex_color[1:3], 16) / 255, int(hex_color[3:5], 16) / 255, int(hex_color[5:7], 16) / 255
            h, s, v = colorsys.rgb_to_hsv(r, g, b)
            r2, g2, b2 = colorsys.hsv_to_rgb(h, s, max(0, min(1, v + delta)))
            return "#{:02x}{:02x}{:02x}".format(int(r2 * 255), int(g2 * 255), int(b2 * 255))

        hover_color = _shade(color, -0.10)
        btn.bind("<Enter>", lambda _e, b=btn, c=hover_color: b.config(bg=c, activebackground=c))
        btn.bind("<Leave>", lambda _e, b=btn, c=color: b.config(bg=c, activebackground=c))
        return btn

    def _init_settings(self):
        self.settings = self.load_settings()
        self.camp_data = self.settings.get("camps", DEFAULT_CAMPS)
        self.dock_list = self.settings.get("docks", DEFAULT_DOCKS)
        self.hocha_list = self.settings.get("hochas", DEFAULT_HOCHAS)
        self.dock_presets = self.settings.get("dock_presets", {})
        self.dispatch_manual_order = self.settings.get("dispatch_manual_order", {})
        self.app_title = self.settings.get("app_title", DEFAULT_SETTINGS["app_title"])
        self.fs_title = self.settings.get("fs_title", DEFAULT_SETTINGS["fs_title"])
        self.colors = self.settings.get("colors", DEFAULT_SETTINGS["colors"])

        display_timing = self.settings.get("display_timing", DEFAULT_SETTINGS["display_timing"])
        self.flap_seconds = self._safe_positive_int(display_timing.get("flap_seconds", 5), 5)
        self.page_flip_seconds = self._safe_positive_int(display_timing.get("page_flip_seconds", 7), 7)
        self.use_split_flap = display_timing.get("use_split_flap", True)
        self.api_refresh_minutes = self._safe_positive_int(display_timing.get("api_refresh_minutes", 3), 1)

        schedules = self.settings.get("schedules", DEFAULT_SETTINGS["schedules"])
        self.VISIBILITY_SCHEDULE = schedules.get("visibility", DEFAULT_SETTINGS["schedules"]["visibility"])
        self.logo_path = resource_path(LOGO_FILENAME)

    def _safe_positive_int(self, value: Any, default: int) -> int:
        try:
            parsed = int(value)
            return parsed if parsed > 0 else default
        except Exception:
            return default

    def prevent_display_sleep(self, prevent: bool = True) -> None:
        if platform.system() == "Windows":
            try:
                ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED = 0x80000000, 0x00000001, 0x00000002
                flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED if prevent else ES_CONTINUOUS
                ctypes.windll.kernel32.SetThreadExecutionState(flags)
            except Exception as e:
                print(f"화면 꺼짐 방지 제어 오류: {e}")

    def _clean_status(self, status_str: str) -> str:
        for token in ["🟢", "🟡", "🔴", "⏳"]:
            status_str = str(status_str).replace(token, "")
        return status_str.strip()

    def _get_treeview_font(self) -> Tuple[str, int, str]:
        available_fonts = set(tkfont.families())
        for family in ["Segoe UI Emoji", "Segoe UI Symbol", "Malgun Gothic", "맑은 고딕"]:
            if family in available_fonts:
                return family, 10, "bold"
        return "TkDefaultFont", 10, "bold"

    def _ordered_ui_mode_keys(self) -> List[str]:
        preferred_order = [
            LEGACY_PRESET_NAME,
            MONO_PRESET_NAME,
            WIN11_PRESET_NAME,
            DARK_PRESET_NAME,
        ]
        ordered = [key for key in preferred_order if key in PALETTE_PRESETS]
        remainder = [key for key in PALETTE_PRESETS if key not in ordered]
        return ordered + remainder

    def _normalize_text(self, value: Any) -> str:
        return re.sub(r"\s+", "", str(value or "")).strip().upper()

    def _find_best_match_key(self, target: str, candidates: List[str]) -> Optional[str]:
        norm_target = self._normalize_text(target)
        if not norm_target: return None
        for candidate in candidates:
            if self._normalize_text(candidate) == norm_target: return candidate
            
        base_target = re.sub(r'\(.*?\)', '', norm_target)
        for candidate in candidates:
            if self._normalize_text(candidate) == base_target: return candidate
            
        sorted_candidates = sorted(candidates, key=lambda x: len(self._normalize_text(x)), reverse=True)
        for candidate in sorted_candidates:
            norm_candidate = self._normalize_text(candidate)
            if norm_target in norm_candidate or norm_candidate in norm_target: return candidate
            if base_target in norm_candidate or norm_candidate in base_target: return candidate
        return None

    def _get_continuous_mins(self, time_str: str) -> int:
        try:
            h, m = map(int, time_str.split(':'))
            return (h * 60 + m) + (24 * 60 if h < 6 else 0)
        except Exception:
            return 0

    def _is_time_in_range(self, start_str: str, end_str: str, now_time: datetime) -> bool:
        try:
            now_mins = now_time.hour * 60 + now_time.minute + (24 * 60 if now_time.hour < 6 else 0)
            start_mins = self._get_continuous_mins(start_str)
            end_mins = self._get_continuous_mins(end_str)
            return start_mins <= now_mins < end_mins
        except Exception:
            return True

    def _check_visibility(self) -> None:
        now = datetime.now()
        for item_id in list(self.tree.get_children()):
            v = self.tree.item(item_id, "values")
            if v and v[1] in self.VISIBILITY_SCHEDULE:
                if not self._is_time_in_range(*self.VISIBILITY_SCHEDULE[v[1]], now):
                    self.detached_items[item_id] = {'values': v, 'tags': self.tree.item(item_id, "tags")}
                    self.tree.detach(item_id)

        for item_id, info in list(self.detached_items.items()):
            v = info['values']
            if v[1] in self.VISIBILITY_SCHEDULE:
                if self._is_time_in_range(*self.VISIBILITY_SCHEDULE[v[1]], now):
                    self.tree.reattach(item_id, "", "end")
                    del self.detached_items[item_id]

    def _auto_assign_hocha(self) -> None:
        pass

    def _visibility_scheduler(self) -> None:
        if not self.root.winfo_exists(): return
        self._check_visibility()
        self._auto_assign_hocha()
        self._sync_full_screen()
        self.root.after(30000, self._visibility_scheduler)

    def _periodic_flap_refresh(self) -> None:
        if self.root.winfo_exists() and hasattr(self, "fs_win") and self.fs_win.winfo_exists():
            for r_idx, row_cells in enumerate(self.fs_rows):
                self.root.after(r_idx * 100,
                                lambda cells=row_cells: [c.force_animate() for c in cells if c.winfo_exists()])
        delay = max(1, self.flap_seconds) * 1000
        self.flap_refresh_timer_id = self.root.after(delay, self._periodic_flap_refresh)

    def _blink_timer(self) -> None:
        if not self.root.winfo_exists(): return
        self.blink_state = not self.blink_state
        try:
            self.tree.tag_configure("blink", foreground="#C0392B" if self.blink_state else "#7B1A13")
        except Exception:
            pass

        if hasattr(self, "fs_win") and self.fs_win.winfo_exists():
            if hasattr(self, "fs_time_label") and self.fs_time_label.winfo_exists():
                self.fs_time_label.config(text=datetime.now().strftime("%H:%M:%S"))

            start_idx = self.fs_current_page * self.MAX_FS_ROWS
            page_items = self.tree.get_children()[start_idx:start_idx + self.MAX_FS_ROWS]

            if len(page_items) <= len(self.fs_rows):
                now = datetime.now()
                now_min = now.hour * 60 + now.minute
                for i, item_id in enumerate(page_items):
                    v = self.tree.item(item_id, "values")
                    if not v: continue

                    status_clean, time_str = self._clean_status(v[0]), v[7]
                    time_cell, status_cell = self.fs_rows[i][5], self.fs_rows[i][6]

                    if status_clean in ["재수배", "차량변경"]:
                        status_cell.set_color("#FF4444" if self.blink_state else "#994444")

                    time_color = self.colors["font"]
                    if time_str and ":" in time_str and status_clean != "접안":
                        try:
                            h, m = map(int, time_str.split(":"))
                            arrive_min = h * 60 + m  # Arrive 시간 (분)

                            # 근무일 기준 판단:
                            # - Arrive time 07:00~23:59 → 당일 시간
                            # - Arrive time 00:00~06:59 → 다음날 시간 (자정 이후)
                            # 현재 시각이 07:00~23:59(now_min >= 420)이면
                            # 다음날(00:00~06:59) arrive는 아직 미래 → 절대 깜박이지 않음
                            is_next_day_arrive = (h < 7)  # 자정 이후 시간
                            if is_next_day_arrive:
                                # 다음날 arrive: 현재가 00:00~06:59일 때만 비교
                                if now_min < 420 and now_min >= arrive_min:
                                    time_color = "#FF4444" if self.blink_state else self.colors["font"]
                                # 현재가 07:00~23:59이면 아직 미래 → 깜박이지 않음
                            else:
                                # 당일 arrive: 현재가 같은 시간대(07:00~23:59)일 때만 비교
                                if now_min >= 420 and now_min >= arrive_min:
                                    time_color = "#FF4444" if self.blink_state else self.colors["font"]
                        except ValueError:
                            pass
                    time_cell.set_color(time_color)

        self.root.after(500, self._blink_timer)

    def _schedule_fs_page_flip(self) -> None:
        if getattr(self, "fs_page_timer_id", None):
            self.root.after_cancel(self.fs_page_timer_id)
            self.fs_page_timer_id = None
        if hasattr(self, "fs_win") and self.fs_win.winfo_exists():
            self.fs_page_timer_id = self.root.after(max(1, self.page_flip_seconds) * 1000, self._fs_page_flip_loop)

    def _fs_page_flip_loop(self):
        if not hasattr(self, "fs_win") or not self.fs_win.winfo_exists():
            self.fs_page_timer_id = None
            return
        total_pages = max(1, (len(self.tree.get_children()) + self.MAX_FS_ROWS - 1) // self.MAX_FS_ROWS)
        if total_pages > 1:
            self.fs_current_page = (self.fs_current_page + 1) % total_pages
            self._sync_full_screen()
        self._schedule_fs_page_flip()

    def _on_board_configure(self, event) -> None:
        if getattr(self, "_resize_timer", None):
            self.root.after_cancel(self._resize_timer)
        self._resize_timer = self.root.after(200, self._apply_new_size)

    def _apply_new_size(self) -> None:
        if not hasattr(self, "fs_board") or not self.fs_board.winfo_exists(): return
        new_max = max(5, (self.fs_board.winfo_height() - 60) // 55)
        if self.MAX_FS_ROWS != new_max:
            self.MAX_FS_ROWS, self.fs_current_page = new_max, 0
            self._sync_full_screen(force_rebuild=True)

    def show_full_screen(self) -> None:
        if hasattr(self, "fs_win") and self.fs_win.winfo_exists():
            self.fs_win.lift()
            return

        self.fs_win = tk.Toplevel(self.root)
        self.fs_win.title(self.fs_title)
        self.fs_win.geometry("1280x720")
        self.fs_win.configure(bg=self.colors["bg"])
        self.prevent_display_sleep(True)
        self.fs_is_fullscreen = False
        self.normal_geometry = "1280x720+0+0"
        self.fs_current_page = 0

        def close_fs():
            if getattr(self, "fs_page_timer_id", None):
                self.root.after_cancel(self.fs_page_timer_id)
                self.fs_page_timer_id = None
            if getattr(self, "_resize_timer", None):
                self.root.after_cancel(self._resize_timer)
                self._resize_timer = None
            self.prevent_display_sleep(False)
            self.fs_win.destroy()

        self.fs_win.protocol("WM_DELETE_WINDOW", close_fs)

        def toggle_fullscreen(event=None):
            self.fs_is_fullscreen = not self.fs_is_fullscreen
            if self.fs_is_fullscreen:
                self.normal_geometry = self.fs_win.geometry()
                try:
                    if platform.system() == "Windows":
                        hwnd = ctypes.windll.user32.GetParent(self.fs_win.winfo_id())
                        hmon = ctypes.windll.user32.MonitorFromWindow(hwnd, 2)

                        class RECT(ctypes.Structure):
                            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long),
                                        ("bottom", ctypes.c_long)]

                        class MONITORINFO(ctypes.Structure):
                            _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", RECT), ("rcWork", RECT),
                                        ("dwFlags", ctypes.c_ulong)]

                        mi = MONITORINFO()
                        mi.cbSize = ctypes.sizeof(MONITORINFO)
                        ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(mi))
                        w, h = mi.rcMonitor.right - mi.rcMonitor.left, mi.rcMonitor.bottom - mi.rcMonitor.top
                        self.fs_win.overrideredirect(True)
                        self.fs_win.geometry(f"{w}x{h}+{mi.rcMonitor.left}+{mi.rcMonitor.top}")
                    else:
                        self.fs_win.attributes("-fullscreen", True)
                except Exception:
                    self.fs_win.attributes("-fullscreen", True)
            else:
                if platform.system() == "Windows":
                    self.fs_win.overrideredirect(False)
                self.fs_win.attributes("-fullscreen", False)
                self.fs_win.geometry(self.normal_geometry)
                self.fs_win.lift()
            return "break"

        self.fs_win.bind("<F11>", toggle_fullscreen)
        self.fs_win.bind("<Escape>", lambda e: toggle_fullscreen() if self.fs_is_fullscreen else close_fs())

        header_frame = tk.Frame(self.fs_win, bg=self.colors["bg"], pady=16)
        header_frame.pack(fill="x")

        tk.Label(header_frame, text=self.fs_title, font=("맑은 고딕", 36, "bold"), fg="#f4e0aa", bg=self.colors["bg"]).pack(
            side="left", padx=(40, 10))

        fs_logo_label = tk.Label(header_frame, bg=self.colors["bg"])
        fs_logo_label.pack(side="right", padx=(0, 40))

        fs_hq_img = load_hq_image(self.logo_path, target_height=50)
        if fs_hq_img:
            self.fs_logo_img = fs_hq_img
            fs_logo_label.config(image=self.fs_logo_img)
        else:
            fs_logo_label.config(text="COUPANG", font=("Impact", 28, "bold"), fg="#FF5C00")

        self.fs_time_label = tk.Label(header_frame, text="", font=("Courier New", 20, "bold"), fg=self.colors["clock"],
                                      bg=self.colors["bg"])
        self.fs_time_label.pack(side="right", padx=(0, 30))

        self.fs_page_label = tk.Label(header_frame, text="", font=("맑은 고딕", 16, "bold"), fg="#00BCD4",
                                      bg=self.colors["bg"])
        self.fs_page_label.pack(side="right", padx=20)

        self.fs_board = tk.Frame(self.fs_win, bg=self.colors["bg"])
        self.fs_board.pack(fill="both", expand=True, padx=40, pady=10)
        self.fs_board.bind("<Configure>", self._on_board_configure)

        for i in range(7):
            self.fs_board.columnconfigure(i, weight=1 if i in [0, 3, 5] else 2, uniform="col")

        for i, text in enumerate(("TYPE", "CAMP", "VEHICLE NO.", "DOCK", "CAR NUMBER", "Arrive", "STATUS")):
            tk.Label(self.fs_board, text=text, font=("Courier New", 18, "bold"), fg="#aaaaaa",
                     bg=self.colors["bg"]).grid(row=0, column=i, sticky="nsew", pady=(0, 10))

        self.fs_rows = []
        self._sync_full_screen(force_rebuild=True)
        self._schedule_fs_page_flip()

    def _sync_full_screen(self, force_rebuild=False) -> None:
        if not hasattr(self, "fs_win") or not self.fs_win.winfo_exists(): return

        all_items = self.tree.get_children()
        total_pages = max(1, (len(all_items) + self.MAX_FS_ROWS - 1) // self.MAX_FS_ROWS)

        if self.fs_current_page >= total_pages: self.fs_current_page = 0
        page_items = all_items[self.fs_current_page * self.MAX_FS_ROWS: (self.fs_current_page + 1) * self.MAX_FS_ROWS]

        if hasattr(self, "fs_page_label"):
            self.fs_page_label.config(
                text=f"페이지: {self.fs_current_page + 1} / {total_pages}" if total_pages > 1 else "")

        if force_rebuild or len(self.fs_rows) != self.MAX_FS_ROWS:
            for widget in self.fs_board.winfo_children():
                try:
                    if int(widget.grid_info().get("row", 0)) > 0: widget.destroy()
                except Exception as e:
                    print(f"[출력 디버그] 고임목 LibreOffice PDF 변환 실패: {e}")

            self.fs_rows = []
            for row_idx in range(1, self.MAX_FS_ROWS + 1):
                cells = []
                bg_color = self.colors["even_bg"] if row_idx % 2 == 0 else self.colors["odd_bg"]
                for col_idx in range(7):
                    cell = SplitFlapCell(self.fs_board, bg_color=bg_color)
                    cell.use_anim = getattr(self, "use_split_flap", True)
                    cell.grid(row=row_idx, column=col_idx, sticky="nsew", padx=5, pady=5)
                    cells.append(cell)
                self.fs_rows.append(cells)

        for i in range(self.MAX_FS_ROWS):
            if i < len(page_items):
                v = self.tree.item(page_items[i], "values")
                if not v: continue

                status_clean = self._clean_status(v[0])
                display_data = (v[1], v[2], v[4], "", "", "", status_clean) if status_clean == "재수배" else (v[1], v[2],
                                                                                                           v[4], v[3],
                                                                                                           v[5], v[7],
                                                                                                           status_clean)

                for j, val in enumerate(display_data):
                    color = self.colors["font"]
                    if j == 6:
                        if status_clean == "접안":
                            color = "#00FF00"
                        elif status_clean in ["재수배", "차량변경"]:
                            color = "#FF4444" if getattr(self, "blink_state", True) else "#994444"
                        elif status_clean in ["수배중", "미접안"]:
                            color = "#FFFF00"
                    self.fs_rows[i][j].set_text(val, fg_color=color)
            else:
                for j in range(7):
                    self.fs_rows[i][j].set_text("")

    def _extract_oid(self, tags) -> str:
        for tag in tags:
            if str(tag).startswith("OID_"): return str(tag)[4:]
        return ""

    def _override_delivery_type(self, d_type: str, camp: str, hocha: str) -> str:
        camp_clean = str(camp).replace(" ", "").upper()
        hocha_clean = str(hocha).replace(" ", "")

        if d_type == "SD2" and "인천3" in camp_clean and hocha_clean in ["2호차", "3호차", "2", "3"]: return "SD3"
        if d_type == "W2":
            ganseon_targets = ["김해4HUB", "남김해1(SH)", "동대구화물터미널", "서대구1(SH)", "전라광주7HUB"]
            for target in ganseon_targets:
                if target.replace(" ", "").upper() in camp_clean: return "간선"
        return d_type

    def _validate_inputs(self) -> Tuple[bool, str]:
        d_type, hocha = self.type_combo.get().strip(), self.hocha_combo.get().strip()
        car_num, phone, time_val = self.car_num_entry.get().strip(), self.phone_entry.get().strip(), self.time_entry.get().strip()
        camp, dock = ("-", "-") if d_type == "KPP" else (self.camp_combo.get().strip(), self.dock_combo.get().strip())

        if not all([d_type, camp, dock, hocha]): return False, "필수 항목을 모두 입력해주세요."
        if phone and not PHONE_PATTERN.match(phone): return False, "전화번호 형식이 올바르지 않습니다."

        if time_val:
            if len(time_val) == 4 and time_val.isdigit():
                time_val = f"{time_val[:2]}:{time_val[2:]}"
                self.time_entry.delete(0, tk.END)
                self.time_entry.insert(0, time_val)
            if not TIME_PATTERN.match(time_val): return False, "시간은 HH:MM 형식으로 입력하세요."

        car_num_normalized = car_num.replace(" ", "")
        mapped_d_type = self._override_delivery_type(d_type, camp, hocha)

        for iid in self.tree.get_children():
            v = self.tree.item(iid, "values")
            if v and v[1] == mapped_d_type and v[2] == camp and v[5].replace(" ",
                                                                             "") == car_num_normalized and car_num_normalized:
                return False, "동일 차량이 이미 등록되어 있습니다."
        return True, ""

    def add_record(self) -> None:
        valid, msg = self._validate_inputs()
        if not valid:
            messagebox.showwarning("입력 오류", msg)
            return

        raw_d_type = self.type_combo.get().strip()
        camp = self.camp_combo.get().strip() if raw_d_type != "KPP" else "-"
        dock = self.dock_combo.get().strip() if raw_d_type != "KPP" else "-"
        hocha = self.hocha_combo.get().strip()

        d_type = self._override_delivery_type(raw_d_type, camp, hocha)
        car_num = self.car_num_entry.get().strip()
        phone = self.phone_entry.get().strip()
        time_val = self.time_entry.get().strip()

        v = ("🟡 미접안", d_type, camp, dock, hocha, car_num, phone, time_val)
        oid = f"MAN_{int(time.time() * 1000)}"
        self.tree.insert("", "end", values=v, tags=(f"OID_{oid}",))

        self.dispatch_registry[oid] = {"d_type": d_type, "camp": camp, "car_num": car_num, "phone": phone,
                                       "driver_name": "", "ls_id": ""}

        for entry in (self.car_num_entry, self.phone_entry, self.time_entry): entry.delete(0, tk.END)
        self.car_num_entry.focus_set()

        self._sort_treeview()
        self._check_visibility()
        self._auto_assign_dock_from_preset()
        self._sync_full_screen()

        now = datetime.now()
        if d_type in self.VISIBILITY_SCHEDULE and not self._is_time_in_range(*self.VISIBILITY_SCHEDULE[d_type], now):
            messagebox.showinfo("알림", f"'{d_type}' 차량이 시스템에 정상 추가되었으나,\n현재 해당 Type의 화면 표시 시간대가 아니어 숨김 처리되었습니다.")

    def apply_status(self) -> None:
        selected_items = self.tree.selection()
        new_status = self.status_action_combo.get()
        if not selected_items: return

        for iid in selected_items:
            v = list(self.tree.item(iid, "values"))
            old_tags = list(self.tree.item(iid, "tags"))
            if new_status == "접안":
                v[0], v[7] = f"🟢 {new_status}", datetime.now().strftime("%H:%M")
                if "docked" not in old_tags: old_tags.append("docked")
            elif new_status == "미접안":
                v[0] = f"🟡 {new_status}"
                if "docked" in old_tags: old_tags.remove("docked")
            elif new_status == "수배중":
                v[0] = f"⏳ {new_status}"
            else:
                v[0] = f"🔴 {new_status}"
                if "blink" not in old_tags: old_tags.append("blink")

            self.tree.item(iid, values=tuple(v), tags=tuple(old_tags))
        self._sort_treeview()
        self._sync_full_screen()

    def delete_record(self) -> None:
        selected = self.tree.selection()
        if selected and messagebox.askyesno("확인", f"선택한 {len(selected)}개 항목을 삭제하시겠습니까? (출차서 데이터에서도 삭제됩니다)"):
            for item in selected:
                oid = self._extract_oid(self.tree.item(item, "tags"))
                if oid in self.dispatch_registry: del self.dispatch_registry[oid]
                self.tree.delete(item)
            self._sync_full_screen()

    def delete_all_records(self) -> None:
        all_items = self.tree.get_children()
        total_count = len(all_items) + len(self.detached_items)
        if total_count > 0 and messagebox.askyesno("확인", f"숨겨진 리스트를 포함한 전체 {total_count}개 항목을 모두 삭제하시겠습니까? (출차서 초기화)"):
            for item in all_items: self.tree.delete(item)
            for item in list(self.detached_items.keys()): self.tree.delete(item)
            self.detached_items.clear()
            self.dispatch_registry.clear()
            self._sync_full_screen()

    def on_tree_double_click(self, event) -> None:
        if selected := self.tree.selection(): self.open_edit_window(selected[0])

    def open_edit_window(self, item_id: str) -> None:
        v, original_tags = self.tree.item(item_id, "values"), self.tree.item(item_id, "tags")
        C = self.C

        edit_win = tk.Toplevel(self.root)
        edit_win.title("차량 정보 수정")
        edit_win.geometry("680x230")
        edit_win.configure(bg=C["bg"])
        # resizable constraint removed for UX
        edit_win.transient(self.root)
        edit_win.grab_set()

        hdr = tk.Frame(edit_win, bg=C["header"], height=40)
        hdr.pack(fill="x");
        hdr.pack_propagate(False)
        tk.Label(hdr, text="차량 정보 수정", bg=C["header"], fg="white", font=("맑은 고딕", 10, "bold")).pack(side="left",
                                                                                                    padx=16, pady=10)

        body = tk.Frame(edit_win, bg=C["white"], padx=20, pady=16)
        body.pack(fill="both", expand=True)

        def lbl(text):
            return tk.Label(body, text=text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9))

        r0 = tk.Frame(body, bg=C["white"])
        r0.pack(fill="x", pady=(0, 8))
        for i, t in enumerate(["Type", "캠프명", "도크", "호차"]): lbl(t).grid(row=0, column=i * 3, in_=r0,
                                                                        padx=(0 if i == 0 else 14, 4), sticky="w")

        e_type = ttk.Combobox(r0, values=list(self.camp_data.keys()), state="readonly", width=10)
        e_camp = ttk.Combobox(r0, state="readonly", width=13)
        e_dock = ttk.Combobox(r0, values=self.dock_list, state="readonly", width=8)
        e_hocha = ttk.Combobox(r0, values=self.hocha_list, state="readonly", width=8)
        e_type.grid(row=0, column=1);
        e_camp.grid(row=0, column=4);
        e_dock.grid(row=0, column=7);
        e_hocha.grid(row=0, column=10)

        r1 = tk.Frame(body, bg=C["white"])
        r1.pack(fill="x", pady=(0, 14))
        for i, t in enumerate(["차량번호", "전화번호", "접안시간"]): lbl(t).grid(row=0, column=i * 3, in_=r1,
                                                                     padx=(0 if i == 0 else 14, 4), sticky="w")

        e_car = self._entry(r1, width=16)
        e_phone = self._entry(r1, width=16)
        e_time = self._entry(r1, width=10)
        e_car.grid(row=0, column=1);
        e_phone.grid(row=0, column=4);
        e_time.grid(row=0, column=7)

        btn_row = tk.Frame(body, bg=C["white"])
        btn_row.pack(fill="x")
        save_btn = self._btn(btn_row, "저장", lambda: None, C["btn_success"])
        save_btn.pack(side="right")
        self._btn(btn_row, "취소", edit_win.destroy, C["btn_neutral"]).pack(side="right", padx=(0, 8))

        e_type.set(v[1]);
        e_hocha.set(v[4]);
        e_car.insert(0, v[5]);
        e_phone.insert(0, v[6]);
        e_time.insert(0, v[7])

        def update_combos(e=None):
            if e_type.get() == "KPP":
                e_camp.set("-");
                e_camp.config(state="disabled")
                e_dock.set("-");
                e_dock.config(state="disabled")
            else:
                e_camp.config(state="readonly");
                e_dock.config(state="readonly")
                camps = self.camp_data.get(e_type.get(), [])
                e_camp["values"] = camps
                if e_camp.get() not in camps: e_camp.set(camps[0] if camps else "")
                if e_dock.get() == "-": e_dock.set(self.dock_list[0] if self.dock_list else "")

        e_type.bind("<<ComboboxSelected>>", update_combos)
        update_combos()
        if v[1] != "KPP":
            e_camp.set(v[2]);
            e_dock.set(v[3])

        def save_edit():
            time_val = e_time.get().strip()
            if time_val and not TIME_PATTERN.match(time_val):
                messagebox.showerror("오류", "시간은 HH:MM 형식으로 입력하세요.")
                return
            raw_nt = e_type.get().strip()
            nc = e_camp.get().strip() if raw_nt != "KPP" else "-"
            nd = e_dock.get().strip() if raw_nt != "KPP" else "-"
            n_hocha = e_hocha.get().strip()
            nt = self._override_delivery_type(raw_nt, nc, n_hocha)
            n_car = e_car.get().strip();
            n_phone = e_phone.get().strip()
            self.tree.item(item_id, values=(v[0], nt, nc, nd, n_hocha, n_car, n_phone, time_val), tags=original_tags)
            oid = self._extract_oid(original_tags)
            if oid and oid in self.dispatch_registry:
                self.dispatch_registry[oid].update({"d_type": nt, "camp": nc, "car_num": n_car, "phone": n_phone})
            self._sort_treeview()
            self._check_visibility();
            self._sync_full_screen();
            edit_win.destroy()

        save_btn.config(command=save_edit)
        edit_win.bind("<Return>", lambda _e: save_edit())

    def update_camp_combobox(self, event=None) -> None:
        if self.type_combo.get() == "KPP":
            self.camp_combo.set("-");
            self.camp_combo.config(state="disabled")
            self.dock_combo.set("-");
            self.dock_combo.config(state="disabled")
        else:
            self.camp_combo.config(state="readonly");
            self.dock_combo.config(state="readonly")
            camps = self.camp_data.get(self.type_combo.get(), [])
            self.camp_combo["values"] = camps
            if self.camp_combo.get() not in camps: self.camp_combo.set(camps[0] if camps else "")
            if self.dock_combo.get() == "-": self.dock_combo.set(self.dock_list[0] if self.dock_list else "")

    def create_widgets(self) -> None:
        C = self.C
        self.root.configure(bg=C["bg"])

        header = tk.Frame(self.root, bg=C["header"], height=64)
        header.pack(fill="x");
        header.pack_propagate(False)

        logo_frame = tk.Frame(header, bg=C["header"])
        logo_frame.pack(side="left", padx=(20, 0))
        hq_img = load_hq_image(self.logo_path, target_height=30)
        if hq_img:
            self.logo_img = hq_img
            tk.Label(logo_frame, image=self.logo_img, bg=C["header"]).pack(side="left", pady=16)
        else:
            tk.Label(logo_frame, text="COUPANG", bg=C["header"], font=("Impact", 17, "bold"), fg="#FF6600").pack(
                side="left", pady=16)

        tk.Frame(header, bg=C["accent"], width=1, height=34).pack(side="left", padx=16, pady=15)

        title_wrap = tk.Frame(header, bg=C["header"])
        title_wrap.pack(side="left")
        tk.Label(title_wrap, text=self.app_title, bg=C["header"], fg=C["header_txt"], font=("맑은 고딕", 11, "bold")).pack(
            anchor="w")

        self._btn(header, "⚙  설정", self.open_presets, C["btn_secondary"]).pack(side="right", padx=16, pady=14)

        inp_card = tk.Frame(self.root, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        inp_card.pack(fill="x", padx=12, pady=(12, 0))

        inp_inner = tk.Frame(inp_card, bg=C["white"], padx=16, pady=12)
        inp_inner.pack(fill="x")

        tk.Label(inp_inner, text="차량 등록", bg=C["white"], fg=C["btn_primary"], font=("맑은 고딕", 10, "bold")).grid(row=0,
                                                                                                               column=0,
                                                                                                               columnspan=10,
                                                                                                               sticky="w",
                                                                                                               pady=(0,
                                                                                                                     8))

        def lbl(text):
            return tk.Label(inp_inner, text=text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9))

        for i, t in enumerate(["Type", "캠프명", "접안 도크", "호차"]): lbl(t).grid(row=1, column=i * 3,
                                                                           padx=(0 if i == 0 else 14, 4), sticky="w")

        self.type_combo = ttk.Combobox(inp_inner, values=list(self.camp_data.keys()), state="readonly", width=10)
        self.camp_combo = ttk.Combobox(inp_inner, state="readonly", width=13)
        self.dock_combo = ttk.Combobox(inp_inner, values=self.dock_list, state="readonly", width=9)
        self.hocha_combo = ttk.Combobox(inp_inner, values=self.hocha_list, state="readonly", width=9)

        for i, cb in enumerate([self.type_combo, self.camp_combo, self.dock_combo, self.hocha_combo]):
            cb.grid(row=1, column=i * 3 + 1, padx=(0, 0), sticky="w")
        self.type_combo.bind("<<ComboboxSelected>>", self.update_camp_combobox)

        inp_inner.columnconfigure(13, weight=1)
        self._btn(inp_inner, "  +  등록", self.add_record, C["btn_success"]).grid(row=1, column=14, rowspan=2,
                                                                                padx=(14, 0), ipadx=14, ipady=4,
                                                                                sticky="e")

        for i, t in enumerate(["차량번호", "전화번호", "접안시간"]): lbl(t).grid(row=2, column=i * 3, padx=(0 if i == 0 else 14, 4),
                                                                     sticky="w", pady=(6, 0))

        self.car_num_entry = self._entry(inp_inner, width=18)
        self.phone_entry = self._entry(inp_inner, width=18)
        self.time_entry = self._entry(inp_inner, width=10)
        self.car_num_entry.grid(row=2, column=1, sticky="w", pady=(6, 0))
        self.phone_entry.grid(row=2, column=4, sticky="w", pady=(6, 0))
        self.time_entry.grid(row=2, column=7, sticky="w", pady=(6, 0))

        self.root.bind("<Return>", lambda _e: self.add_record())

        toolbar = tk.Frame(self.root, bg=C["toolbar"], highlightthickness=1, highlightbackground=C["border_soft"],
                           height=54)
        toolbar.pack(fill="x", padx=12, pady=(8, 0))
        toolbar.pack_propagate(False)

        tb_left = tk.Frame(toolbar, bg=C["toolbar"])
        tb_left.pack(side="left", padx=12, pady=8)

        tk.Label(tb_left, text="상태 변경", bg=C["toolbar"], fg=C["label"], font=("맑은 고딕", 9, "bold")).pack(side="left",
                                                                                                        padx=(0, 6))
        self.status_action_combo = ttk.Combobox(tb_left, values=STATUS_LIST, state="readonly", width=9)
        self.status_action_combo.current(0)
        self.status_action_combo.pack(side="left", padx=(0, 6))
        self._btn(tb_left, "적용", self.apply_status, C["btn_primary"]).pack(side="left", padx=(0, 4))
        self._btn(tb_left, "삭제", self.delete_record, C["btn_danger"]).pack(side="left")

        tk.Frame(toolbar, bg=C["border"], width=1).pack(side="left", fill="y", padx=10, pady=8)

        btns = [
            ("📺 전광판", self.show_full_screen, C["btn_dark"]),
            ("🌐 LS연결", self.open_api_fetch_window, C["btn_primary"]),
            #("📥 LS엑셀", self.load_ls_data, "#5e4b8a"),
            ("📄 공유", self.show_print_view, C["btn_warning"]),
            ("🚚 출차서", self.open_dispatch_window, C["btn_teal"]),
        ]
        for text, cmd, color in btns:
            self._btn(toolbar, text, cmd, color).pack(side="left", padx=3, pady=8)

        tb_right = tk.Frame(toolbar, bg=C["toolbar"])
        tb_right.pack(side="right", padx=12, pady=8)
        tk.Label(tb_right, text="더블클릭: 수정", bg=C["toolbar"], fg=C["placeholder"], font=("맑은 고딕", 8)).pack(side="right",
                                                                                                          padx=10)
        self._btn(tb_right, "🗑 전체 삭제", self.delete_all_records, C["btn_danger"]).pack(side="right")

        tree_card = tk.Frame(self.root, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        tree_card.pack(fill="both", expand=True, padx=12, pady=8)

        self.tree = ttk.Treeview(tree_card, columns=COLUMNS, show="headings")
        col_w = {"상태": 108, "Delivery Type": 105, "캠프명": 110, "접안 도크": 88, "호차": 78, "차량번호": 122, "전화번호": 140,
                 "접안시간": 82}
        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_w.get(col, 110), anchor="center", minwidth=60)

        self.tree.tag_configure("docked", foreground=C["status_ok"])
        self.tree.tag_configure("blink", foreground=C["status_err"])
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.tag_configure("odd", background=C["row_odd"])
        self.tree.tag_configure("even", background=C["row_even"])

        sb = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        statusbar = tk.Frame(self.root, bg=C["header"], height=24)
        statusbar.pack(fill="x", side="bottom")
        statusbar.pack_propagate(False)
        tk.Label(statusbar, text="INC4 HUB  |  v1.5  |  Producer by Miles", bg=C["header"], fg=C["header_muted"],
                 font=("맑은 고딕", 8)).pack(side="left", padx=14)
                 
        self.api_status_label = tk.Label(statusbar, text="🔴 시스템 끊김", bg=C["header"], fg="#FF4444", font=("맑은 고딕", 9, "bold"))
        self.api_status_label.pack(side="right", padx=14)

        if self.camp_data:
            self.type_combo.current(0)
            self.update_camp_combobox()
        if self.dock_list: self.dock_combo.current(0)
        if self.hocha_list: self.hocha_combo.current(0)

    # ------------------------------------------
    # API 데이터 수집 (ApiManager 연동)
    # ------------------------------------------
    def _find_existing_record(self, order_id: str, d_type: str, camp: str, car_num: str, hocha: str) -> Tuple[
        Optional[str], str, str]:
        if order_id and order_id in self.dispatch_registry:
            for iid in self.tree.get_children():
                if f"OID_{order_id}" in self.tree.item(iid, "tags"): return iid, "tree", order_id
            for iid, info in self.detached_items.items():
                if f"OID_{order_id}" in info["tags"]: return iid, "detached", order_id
            return None, "registry", order_id

        car_num_normalized = car_num.replace(" ", "") if car_num else ""
        for iid in self.tree.get_children():
            v = self.tree.item(iid, "values")
            if v and v[1] == d_type and v[2] == camp:
                if car_num_normalized and v[5].replace(" ",
                                                       "") == car_num_normalized: return iid, "tree", self._extract_oid(
                    self.tree.item(iid, "tags"))
                if hocha and v[4] == hocha: return iid, "tree", self._extract_oid(self.tree.item(iid, "tags"))

        for iid, info in self.detached_items.items():
            v = info["values"]
            if v and v[1] == d_type and v[2] == camp:
                if car_num_normalized and v[5].replace(" ",
                                                       "") == car_num_normalized: return iid, "detached", self._extract_oid(
                    info["tags"])
                if hocha and v[4] == hocha: return iid, "detached", self._extract_oid(info["tags"])
        return None, "none", ""

    def _get_item_values_and_tags(self, item_id: str, location: str) -> Tuple[List[Any], List[str]]:
        if location == "tree": return list(self.tree.item(item_id, "values")), list(self.tree.item(item_id, "tags"))
        if location == "detached":
            info = self.detached_items[item_id]
            return list(info["values"]), list(info["tags"])
        raise ValueError(f"지원하지 않는 location 입니다: {location}")

    def _set_item_values_and_tags(self, item_id: str, location: str, values: List[Any], tags: List[str]) -> None:
        if location == "tree":
            self.tree.item(item_id, values=tuple(values), tags=tuple(tags))
            return
        if location == "detached":
            self.detached_items[item_id]["values"] = tuple(values)
            self.detached_items[item_id]["tags"] = tuple(tags)
            return
        raise ValueError(f"지원하지 않는 location 입니다: {location}")

    def _update_existing_item_from_api(self, item_id: str, location: str, final_oid: str, dock_arrive_time: str,
                                       status_val: str, final_time: str, raw_car_num: str, car_num: str,
                                       phone: str, truck_seq: str = "") -> bool:
        old_v, old_tags = self._get_item_values_and_tags(item_id, location)
        needs_update = False

        if "미접안" in old_v[0] and dock_arrive_time:
            old_v[0] = status_val
            old_v[7] = final_time
            if "docked" not in old_tags: old_tags.append("docked")
            needs_update = True

        if not old_v[5] and car_num:
            old_v[5] = raw_car_num
            needs_update = True
        if not old_v[6] and phone:
            old_v[6] = phone
            needs_update = True
            
        if truck_seq and old_v[4] != truck_seq:
            old_v[4] = truck_seq
            needs_update = True

        oid_tag = f"OID_{final_oid}"
        if oid_tag not in old_tags:
            old_tags.append(oid_tag)
            needs_update = True

        if needs_update: self._set_item_values_and_tags(item_id, location, old_v, old_tags)
        return needs_update

    def open_api_fetch_window(self) -> None:
        if not REQUESTS_AVAILABLE:
            messagebox.showwarning("라이브러리 필요",
                                   "API 연동 기능을 사용하려면 requests 라이브러리가 필요합니다.\n명령어: pip install requests")
            return
        C = self.C
        self.api_win = tk.Toplevel(self.root)
        self.api_win.title("API 데이터 수집")
        self.api_win.geometry("460x240")
        self.api_win.configure(bg=C["bg"])
        # resizable constraint removed
        self.api_win.transient(self.root)
        self.api_win.grab_set()

        outer = tk.Frame(self.api_win, bg=C["bg"])
        outer.pack(fill="both", expand=True, padx=20, pady=16)

        card1 = tk.Frame(outer, bg=C["white"], relief="solid", bd=1, highlightthickness=1,
                         highlightbackground=C["border"])
        card1.pack(fill="x", pady=(0, 12))

        for row_idx, (icon, label_text, default_val, is_date) in enumerate(
                [("📅", "조회 날짜 (YYYY-MM-DD):", datetime.now().strftime("%Y-%m-%d"), True),
                 ("🏢", "근무지 코드 (LS조회코드):", "5", False)]):
            row = tk.Frame(card1, bg=C["white"])
            row.pack(fill="x", padx=16, pady=10)
            tk.Label(row, text=f"{icon}  {label_text}", bg=C["white"], fg=C["label"], font=("맑은 고딕", 10)).pack(
                side="left")
            ent = tk.Entry(row, width=16, justify="center", bg=C["white"], fg=C["text"], relief="solid", bd=1,
                           highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["border_focus"],
                           font=("맑은 고딕", 10))
            ent.insert(0, default_val)
            ent.pack(side="right")
            if is_date:
                self.api_date_entry = ent
            else:
                self.api_loc_entry = ent
            if row_idx == 0: tk.Frame(card1, bg=C["border"], height=1).pack(fill="x")

        self.auto_refresh_var = tk.BooleanVar(value=True)
        chk_row = tk.Frame(outer, bg=C["bg"])
        chk_row.pack(fill="x", pady=(0, 14))
        tk.Checkbutton(chk_row, text="  백그라운드 자동 갱신 활성화", variable=self.auto_refresh_var, bg=C["bg"], fg=C["text"],
                       activebackground=C["bg"], selectcolor=C["white"], font=("맑은 고딕", 10), cursor="hand2").pack(
            side="left")
        
        # 주기 설정 (Spinbox) 추가
        tk.Label(chk_row, text="주기(분):", bg=C["bg"], fg=C["label"], font=("맑은 고딕", 10)).pack(side="left", padx=(15, 5))
        self.refresh_spin = tk.Spinbox(chk_row, from_=1, to=60, width=5, justify="center", bg=C["white"], fg=C["text"], 
                                      relief="solid", bd=1, highlightthickness=1, highlightbackground=C["border"],
                                      highlightcolor=C["border_focus"], font=("맑은 고딕", 10))
        self.refresh_spin.delete(0, "end")
        self.refresh_spin.insert(0, str(self.api_refresh_minutes))
        self.refresh_spin.pack(side="left")

        btn_row = tk.Frame(outer, bg=C["bg"])
        btn_row.pack()
        selenium_btn = self._btn(btn_row, "LS OKTA 로그인 및 수집", lambda: self.start_api_fetch(use_selenium=True),
                                 C["btn_teal"])
        selenium_btn.pack(side="left", padx=(0, 8), ipadx=6, ipady=2)
        if not SELENIUM_AVAILABLE:
            selenium_btn.config(state="disabled", cursor="arrow", bg=C["btn_neutral"], activebackground=C["btn_neutral"])
        self._btn(btn_row, "Cookie연결", self.open_manual_cookie_window, C["btn_primary"]).pack(side="left", ipadx=6,
                                                                                               ipady=2)

    def open_manual_cookie_window(self) -> None:
        C = self.C
        cookie_win = tk.Toplevel(self.api_win)
        cookie_win.title("수동 쿠키 입력")
        cookie_win.geometry("420x180")
        cookie_win.configure(bg=C["bg"])
        # resizable constraint removed
        cookie_win.transient(self.api_win)
        cookie_win.grab_set()

        tk.Label(cookie_win, text="웹 브라우저의 쿠키(Cookie) 값을 입력해 주세요:", bg=C["bg"], fg=C["label"],
                 font=("맑은 고딕", 10, "bold")).pack(pady=(20, 10), padx=20, anchor="w")
        cookie_entry = tk.Entry(cookie_win, bg=C["white"], fg=C["text"], relief="solid", bd=1, font=("맑은 고딕", 10))
        cookie_entry.pack(fill="x", padx=20, ipady=6)

        btn_frame = tk.Frame(cookie_win, bg=C["bg"])
        btn_frame.pack(pady=20)

        def on_confirm():
            cookie_val = cookie_entry.get().strip()
            if not cookie_val:
                messagebox.showwarning("입력 오류", "쿠키값이 입력되지 않았습니다.", parent=cookie_win)
                return
            cookie_win.destroy()
            self.start_api_fetch(use_selenium=False, manual_cookie=cookie_val)

        self._btn(btn_frame, "확인 및 수집", on_confirm, C["btn_primary"]).pack(side="left", padx=5)
        self._btn(btn_frame, "취소", cookie_win.destroy, C["btn_neutral"]).pack(side="left", padx=5)
        cookie_entry.focus_set()

    def start_api_fetch(self, use_selenium: bool, manual_cookie: str = "") -> None:
        date_val = self.api_date_entry.get().strip()
        loc_val = self.api_loc_entry.get().strip()
        self.api_run_date = date_val
        self.api_current_url = f"https://ls.coupang.net/data/order?statuses=&usageTypes=&orderStart={date_val}&orderEnd={date_val}&locationStart={loc_val}&page=0&pageSize=300&isShuttle=false"
        self.api_auto_refresh_enabled = self.auto_refresh_var.get()
        try:
            self.api_refresh_minutes = int(self.refresh_spin.get())
            if self.api_refresh_minutes < 1: self.api_refresh_minutes = 1
        except:
            self.api_refresh_minutes = 3

        if self.api_refresh_timer_id:
            try:
                self.root.after_cancel(self.api_refresh_timer_id)
            except Exception:
                pass
            self.api_refresh_timer_id = None

        # 설정 저장
        if "display_timing" not in self.settings: self.settings["display_timing"] = {}
        self.settings["display_timing"]["api_refresh_minutes"] = self.api_refresh_minutes
        self.save_settings(self.settings)

        if not REQUESTS_AVAILABLE:
            messagebox.showerror("오류", "requests 라이브러리가 설치되어 있지 않습니다.")
            return

        if not use_selenium:
            if not manual_cookie:
                messagebox.showerror("오류", "수동 쿠키값이 전달되지 않았습니다.")
                return
            self.api_manager.setup_manual_session(manual_cookie)
            self.api_win.destroy()
            self._fetch_and_apply_api_data()
            return

        if not SELENIUM_AVAILABLE:
            messagebox.showerror("오류", "Selenium 라이브러리가 설치되어 있지 않습니다.")
            return

        try:
            self.api_manager.start_selenium_login()
            C = self.C
            self.login_wait_win = tk.Toplevel(self.api_win)
            self.login_wait_win.title("로그인 대기")
            self.login_wait_win.geometry("400x160")
            self.login_wait_win.configure(bg=C["bg"])
            self.login_wait_win.transient(self.api_win)

            hdr2 = tk.Frame(self.login_wait_win, bg=C["header"], height=36)
            hdr2.pack(fill="x");
            hdr2.pack_propagate(False)
            tk.Label(hdr2, text="OKTA 로그인 대기", bg=C["header"], fg="white", font=("맑은 고딕", 10, "bold")).pack(side="left",
                                                                                                            padx=14,
                                                                                                            pady=8)

            body2 = tk.Frame(self.login_wait_win, bg=C["white"], padx=20, pady=16)
            body2.pack(fill="both", expand=True)
            tk.Label(body2, text="크롬 브라우저에서 OKTA 로그인을 완료해주세요.\n완료 후 아래 버튼을 클릭하면 데이터를 수집합니다.", justify="center",
                     bg=C["white"], fg=C["label"], font=("맑은 고딕", 10)).pack(pady=(0, 12))
            self._btn(body2, "로그인 완료 — 데이터 수집 실행", self._finalize_selenium_login, C["btn_primary"]).pack()
        except Exception as e:
            messagebox.showerror("브라우저 실행 오류", f"크롬 브라우저 구동 실패 (chromedriver 확인 필요):\n{e}")

    def _finalize_selenium_login(self) -> None:
        try:
            self.api_manager.finalize_selenium_login()
            self.login_wait_win.destroy()
            self.api_win.destroy()
            self._fetch_and_apply_api_data()
        except Exception as e:
            messagebox.showerror("오류", f"세션 처리 중 오류가 발생했습니다:\n{e}")

    def _sort_treeview(self) -> None:
        def get_status_priority(status_str):
            s = str(status_str).strip()
            if "차량변경" in s: return 0
            if "미접안" in s: return 1
            if "접안" in s: return 2
            return 3
            
        def extract_hocha_num(hocha_str):
            import re
            m = re.search(r'\d+', str(hocha_str))
            return int(m.group()) if m else 999
            
        children = list(self.tree.get_children(""))
        
        def sort_key(iid):
            vals = self.tree.item(iid, "values")
            if not vals or len(vals) < 5: return (3, "", 999)
            status = vals[0]
            camp = vals[2]
            hocha = vals[4]
            return (get_status_priority(status), camp, extract_hocha_num(hocha))
            
        children.sort(key=sort_key)
        for index, iid in enumerate(children):
            self.tree.move(iid, "", index)

    def _fetch_and_apply_api_data(self, is_auto_refresh=False) -> None:
        if not self.api_manager.session:
            return

        try:
            data = self.api_manager.fetch_data(self.api_current_url)

            contents = data.get('data', {}).get('content', [])
            if not contents:
                contents = data.get("content", data) if isinstance(data, dict) else data

            import urllib.parse as uparse
            seq_mapping = {}
            try:
                parsed_vp = uparse.urlparse(self.api_current_url)
                params_vp = uparse.parse_qs(parsed_vp.query)
                run_date_vp = params_vp.get('orderStart', [self.api_run_date if hasattr(self, 'api_run_date') else ''])[0]
                loc_val_vp = params_vp.get('locationStart', [''])[0]
                
                if run_date_vp and loc_val_vp:
                    clean_date = run_date_vp.split('T')[0].split(' ')[0]
                    # 사용자 요청에 따라 currentWorkplace=5 사용
                    wp_val = loc_val_vp if loc_val_vp else "5"
                    seq_url = f"https://ls.coupang.net/data/truck/fill-rate/loading?orderDate={clean_date}&currentWorkplace={wp_val}"
                    seq_data = self.api_manager.fetch_data(seq_url)

                    def find_trucks_recursive(obj, mapping):
                        if isinstance(obj, dict):
                            # 직접 트럭 정보인 경우 (오직 'sequence' 필드만 확인)
                            if 'truckRequestId' in obj:
                                tr_id = str(obj.get('truckRequestId', '')).strip()
                                # 사용자 요청: "seq"는 무시하고 오직 정확한 "sequence"만 사용
                                seq = obj.get('sequence', None)
                                if tr_id and seq is not None and str(seq).strip() != '':
                                    mapping[tr_id] = str(seq)
                            
                            # 자식 노드들 탐색
                            for val in obj.values():
                                find_trucks_recursive(val, mapping)
                        elif isinstance(obj, list):
                            for item in obj:
                                find_trucks_recursive(item, mapping)
                    
                    req_id_seq_mapping = {}
                    if seq_data:
                        find_trucks_recursive(seq_data, req_id_seq_mapping)
                    
                    if req_id_seq_mapping:
                        print(f"[{datetime.now().strftime('%H:%M:%S')}] Secondary mapping complete: {len(req_id_seq_mapping)} items found.")
                        if "25716693" in req_id_seq_mapping:
                            print(f"  Found targeted ID 25716693 -> {req_id_seq_mapping['25716693']}호차")
                    
                    if req_id_seq_mapping:
                        print(f"[{datetime.now().strftime('%H:%M:%S')}] Mapping successful: {len(req_id_seq_mapping)} truck IDs mapped.")
            except Exception as e:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] 서브 API(sequence) 수집 실패: {e}")

            added_count = 0
            updated_count = 0

            for item in contents:
                api_status = str(item.get("status", "")).strip().upper()
                if api_status != "CONFIRMED":
                    continue

                dest_name = item.get('dest', {}).get('name', '')
                if not dest_name:
                    dest_name = item.get("locationEnd", item.get("destLocation", "알수없음"))

                delivery_types = item.get('deliveryTypes', [])
                d_type = delivery_types[0] if delivery_types else "부자재"

                truck_info = item.get('truckInfo', {})
                tr_req_id = str(item.get("truckRequestId") or truck_info.get("truckRequestId") or "").strip()
                
                truck_seq = ""
                # 오직 truckRequestId 로만 매칭
                if tr_req_id and tr_req_id in req_id_seq_mapping:
                    truck_seq = req_id_seq_mapping[tr_req_id] + "호차"
                
                if not truck_seq:    
                    truck_seq = item.get("truckSeq", "")
                    if str(truck_seq).isdigit():
                        truck_seq = f"{truck_seq}호차"
                    else:
                        truck_seq = str(truck_seq)

                mapped_type = self._override_delivery_type(self.excel_manager.map_delivery_type(d_type), dest_name,
                                                           truck_seq)

                req_time = item.get('requestTime', item.get("dockReqTime", ""))

                trackings = item.get('trackings', [])
                dock_arrive_time = None
                dock_departure_time = None

                if trackings and isinstance(trackings, list) and len(trackings) > 0:
                    if trackings[0].get('dockArriveTime'):
                        dt = datetime.fromtimestamp(trackings[0]['dockArriveTime'] / 1000.0)
                        dock_arrive_time = dt.strftime('%H:%M')
                    if trackings[0].get('dockDepartureTime'):
                        dock_departure_time = trackings[0]['dockDepartureTime']

                if not dock_arrive_time and item.get("dockArriveTime"):
                    dock_arrive_time = self.excel_manager.format_time_value(item.get("dockArriveTime"))
                if not dock_departure_time and item.get("dockDepartureTime"):
                    dock_departure_time = item.get("dockDepartureTime")

                is_departed = bool(dock_departure_time)

                raw_car_num = str(truck_info.get('plateNumber') or item.get("vehicleNo") or item.get("carNumber") or "")
                car_num = raw_car_num.replace(" ", "")
                driver_name = str(
                    truck_info.get('driverName') or item.get("driverName") or item.get("기사 성함") or "").strip()
                phone = str(
                    truck_info.get('driverPhone') or item.get("driverPhone") or item.get("driverContact") or "").strip()
                
                truck_vendor_raw = item.get("truckVendor") or truck_info.get("truckVendor") or truck_info.get("vendor") or ""
                truck_vendor_val = truck_vendor_raw.get("name", "").strip() if isinstance(truck_vendor_raw, dict) else str(truck_vendor_raw).strip()

                order_id = str(item.get("id", item.get("orderId", ""))).strip()
                if not order_id or order_id.lower() in ['nan', 'none']:
                    order_id = f"CAR_{car_num}"
                
                # 목적지가 여러 개인 멀티드랍 차량이 동일한 order_id를 가질 경우를 대비해 도착지를 뒤에 붙여 고유 식별자로 만듭니다.
                # (예: 동일 차량이 인천2(SH)와 동인천2(SH)를 둘 다 방문할 때 덮어쓰기/합쳐짐 방지)
                if dest_name:
                    order_id = f"{order_id}_{dest_name}"

                status_val = "🟡 미접안"
                final_time = self.excel_manager.format_time_value(req_time)
                item_tags = []

                if dock_arrive_time:
                    status_val = "🟢 접안"
                    final_time = dock_arrive_time
                    item_tags.append("docked")

                existing_item, location, found_oid = self._find_existing_record(order_id, mapped_type, dest_name,
                                                                                car_num, truck_seq)
                final_oid = found_oid if found_oid else order_id

                if final_oid not in self.dispatch_registry:
                    self.dispatch_registry[final_oid] = {}

                self.dispatch_registry[final_oid].update({
                    "d_type": mapped_type,
                    "camp": dest_name,
                    "car_num": raw_car_num if raw_car_num else self.dispatch_registry[final_oid].get("car_num", ""),
                    "phone": phone if phone else self.dispatch_registry[final_oid].get("phone", ""),
                    "driver_name": driver_name if driver_name else self.dispatch_registry[final_oid].get("driver_name",
                                                                                                         ""),
                    "truckVendor": truck_vendor_val if truck_vendor_val else self.dispatch_registry[final_oid].get("truckVendor", ""),
                    "ls_id": tr_req_id,
                    "dep_id": item.get("departureDetailId", ""),
                    "arr_id": item.get("arrivalDetailId", ""),
                    "hocha": truck_seq if truck_seq else self.dispatch_registry[final_oid].get("hocha", ""),
                    "_raw": item
                })

                if is_departed:
                    if location == "tree" and existing_item:
                        self.tree.delete(existing_item)
                    elif location == "detached" and existing_item:
                        del self.detached_items[existing_item]
                    continue

                if location in ["tree", "detached"] and existing_item:
                    if self._update_existing_item_from_api(existing_item, location, final_oid, dock_arrive_time,
                                                           status_val, final_time, raw_car_num, car_num, phone, truck_seq):
                        updated_count += 1
                else:
                    # 인천11허브는 목록에 표시하지 않음
                    if str(dest_name).replace(" ", "") == "인천11허브":
                        continue
                    item_tags.append(f"OID_{final_oid}")
                    v = (status_val, mapped_type, dest_name, "-", truck_seq, raw_car_num, phone, final_time)
                    self.tree.insert("", "end", values=v, tags=tuple(item_tags))
                    added_count += 1

            self._sort_treeview()
            self._check_visibility()
            self._auto_assign_dock_from_preset()
            self._sync_full_screen()

            if not is_auto_refresh:
                messagebox.showinfo("수집 완료", f"신규 {added_count}개 추가, {updated_count}개 갱신되었습니다.")
            else:
                print(
                    f"[{datetime.now().strftime('%H:%M:%S')}] API 자동 갱신 완료 - 신규 추가: {added_count} / 기존 갱신: {updated_count}")
                    
            if hasattr(self, "api_status_label"):
                self.api_status_label.config(text="🟢 시스템 연결", fg="#00FF00")

            if self.api_auto_refresh_enabled:
                if self.api_refresh_timer_id:
                    self.root.after_cancel(self.api_refresh_timer_id)
                interval_ms = max(60000, self.api_refresh_minutes * 60 * 1000) # 최소 1분
                self.api_refresh_timer_id = self.root.after(interval_ms, lambda: self._fetch_and_apply_api_data(
                    is_auto_refresh=True))

        except (PermissionError, ValueError) as pe:
            if not is_auto_refresh:
                messagebox.showerror("데이터 파싱/인증 오류", str(pe))
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] {pe}")
            self.api_auto_refresh_enabled = False
            if hasattr(self, "api_status_label"):
                self.api_status_label.config(text="🔴 시스템 끊김", fg="#FF4444")
        except Exception as e:
            msg = f"데이터 수집 및 처리 중 오류가 발생했습니다:\n{e}"
            if not is_auto_refresh:
                messagebox.showerror("오류", msg)
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            if hasattr(self, "api_status_label"):
                self.api_status_label.config(text="🔴 시스템 끊김", fg="#FF4444")

    # ------------------------------------------
    # LS 데이터 로드 (ExcelManager 연동)
    # ------------------------------------------
    def load_ls_data(self) -> None:
        if not PANDAS_AVAILABLE:
            messagebox.showerror("오류", "pandas 라이브러리가 설치되어 있지 않습니다.\n(명령어: pip install pandas openpyxl)")
            return

        filepath = filedialog.askopenfilename(title="LS 엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath: return

        try:
            df = self.excel_manager.read_ls_excel(filepath)
            self.last_ls_filepath = filepath
            orig_cols = df.columns.tolist()
            required_cols = ["요청시간", "출발지", "도착지", "배송유형", "차량번호", "상태", "도크접안시각"]

            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                messagebox.showerror("양식 오류",
                                     f"필수 항목을 찾을 수 없습니다: {', '.join(missing)}\n\n[현재 엑셀의 컬럼 구조]:\n{', '.join(orig_cols[:10])}...")
                return

            df = df[(df['출발지'] == '인천4HUB') & (df['상태'].astype(str).str.strip().str.upper() == 'CONFIRMED') & (df['도착지'].astype(str).str.replace(' ', '') != '인천11허브')].copy()
            if df.empty:
                messagebox.showinfo("알림", "조건에 맞는 데이터가 없습니다.")
                return

            phone_col = next(
                (c for c in df.columns if any(k in str(c).replace(" ", "") for k in ["전화번호", "연락처", "핸드폰", "폰"])), None)
            name_col = next((c for c in df.columns if
                             any(k in str(c).replace(" ", "") for k in ["이름", "기사명", "기사성함", "차주명", "성함"])), None)
            vendor_col = next((c for c in df.columns if
                               any(k in str(c).replace(" ", "").lower() for k in ["간선사", "운송사", "truckvendor", "vendor"])), None)

            df['Mapped_Type'] = df['배송유형'].apply(self.excel_manager.map_delivery_type)
            df['도착지'] = df['도착지'].fillna("").astype(str)
            df = df.sort_values(by=['Mapped_Type', '도착지'])
            df['호차번호'] = df.groupby(['Mapped_Type', '도착지']).cumcount() + 1
            df['호차'] = df['호차번호'].astype(str) + "호차"

            df['Mapped_Type'] = df.apply(
                lambda row: self._override_delivery_type(row['Mapped_Type'], row['도착지'], row['호차']), axis=1)
            df['요청시간_포맷'] = df['요청시간'].apply(self.excel_manager.format_time_value)
            df['도크접안시각_포맷'] = df['도크접안시각'].apply(self.excel_manager.format_time_value)

            added_count = 0
            for _, row in df.iterrows():
                raw_car_num = str(row['차량번호']).strip() if not pd.isna(row['차량번호']) else ""
                car_num = raw_car_num.replace(" ", "")

                req_t = row['요청시간_포맷']
                dock_t = row['도크접안시각_포맷']

                if dock_t:
                    final_time, status_val, item_tags = dock_t, "🟢 접안", ["docked"]
                else:
                    final_time, status_val, item_tags = req_t, "🟡 미접안", []

                order_id = str(row.get("요청ID", "")).strip()
                if not order_id or order_id.lower() in ["nan", "none"]:
                    order_id = f"EXCEL_{int(time.time() * 1000)}_{random.randint(100, 999)}"

                item_tags.append(f"OID_{order_id}")
                phone_val = str(row[phone_col]).strip() if phone_col and not pd.isna(row[phone_col]) else ""
                name_val = str(row[name_col]).strip() if name_col and not pd.isna(row[name_col]) else ""
                vendor_val = str(row[vendor_col]).strip() if vendor_col and not pd.isna(row[vendor_col]) else ""

                if order_id not in self.dispatch_registry:
                    self.dispatch_registry[order_id] = {}

                self.dispatch_registry[order_id].update({
                    "d_type": row['Mapped_Type'], "camp": row['도착지'], "car_num": raw_car_num,
                    "phone": phone_val, "driver_name": name_val,
                    "truckVendor": vendor_val,
                    "ls_id": order_id if not order_id.startswith("EXCEL_") else "",
                    "hocha": row['호차']
                })

                existing_item, loc, _ = self._find_existing_record(order_id, row['Mapped_Type'], row['도착지'], car_num,
                                                                   row['호차'])
                if existing_item: continue

                v = (status_val, row['Mapped_Type'], row['도착지'], "-", row['호차'], raw_car_num, phone_val, final_time)
                self.tree.insert("", "end", values=v, tags=tuple(item_tags))
                added_count += 1
                if added_count % 500 == 0: self.tree.update_idletasks()

            self._sort_treeview()
            self._check_visibility()
            self._auto_assign_dock_from_preset()
            self._sync_full_screen()

            now = datetime.now()
            hidden_types = set()
            for _, row in df.iterrows():
                d_type = row['Mapped_Type']
                if d_type in self.VISIBILITY_SCHEDULE and not self._is_time_in_range(*self.VISIBILITY_SCHEDULE[d_type],
                                                                                     now):
                    hidden_types.add(d_type)

            msg = f"{added_count}개의 새로운 차량 정보를 성공적으로 불러왔습니다."
            if hidden_types: msg += f"\n\n※ 안내: 현재 화면 표시 시간대에 의해 아래 Type은 자동 숨김 처리되었습니다.\n숨김 처리된 타입: {', '.join(hidden_types)}"
            messagebox.showinfo("완료", msg)

        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일을 불러오는 중 알 수 없는 오류가 발생했습니다:\n{e}")

    # ------------------------------------------
    # 리스트 출력 보기
    # ------------------------------------------


    def show_print_view(self) -> None:
        if not (txt := self.get_formatted_list()): return
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("출력 미리보기")
        win.geometry("460x560")
        win.configure(bg=C["bg"])
        win.transient(self.root)

        hdr = tk.Frame(win, bg=C["header"], height=40)
        hdr.pack(fill="x");
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📄  출력 미리보기", bg=C["header"], fg="white", font=("맑은 고딕", 10, "bold")).pack(side="left",
                                                                                                      padx=16, pady=10)

        card = tk.Frame(win, bg=C["white"], padx=4, pady=4)
        card.pack(fill="both", expand=True, padx=12, pady=12)
        ta = tk.Text(card, font=("맑은 고딕", 11), bg=C["white"], fg=C["text"], relief="flat", borderwidth=0, padx=12,
                     pady=10, selectbackground=C["row_sel"])
        ta.pack(fill="both", expand=True)
        ta.insert("1.0", txt)

        tk.Frame(win, bg=C["border"], height=1).pack(fill="x")
        footer = tk.Frame(win, bg=C["bg"], pady=8)
        footer.pack(fill="x")
        self._btn(footer, "닫기", win.destroy, C["btn_neutral"], width=10).pack(side="right", padx=14)

    # ------------------------------------------
    # 출차서 기능
    # ------------------------------------------
    def open_dispatch_window(self, keep_manual=False) -> None:
        if not keep_manual:
            self.dispatch_source_rows = {k: [] for k in DISPATCH_TABS}
            
            # Build mapping from OID to its time_val and dock from tree
            time_lookup = {}
            dock_lookup = {}
            for iid in self.tree.get_children():
                v = self.tree.item(iid, "values")
                tags = self.tree.item(iid, "tags")
                oid = self._extract_oid(tags)
                if oid:
                    time_lookup[oid] = v[7] if len(v) > 7 and v[7] else "99:99"
                    dock_lookup[oid] = v[3] if len(v) > 3 else "-"
                
            for info in self.detached_items.values():
                v = info['values']
                tags = info['tags']
                oid = self._extract_oid(tags)
                if oid:
                    time_lookup[oid] = v[7] if len(v) > 7 and v[7] else "99:99"
                    dock_lookup[oid] = v[3] if len(v) > 3 else "-"
                
            for oid, info in self.dispatch_registry.items():
                d_type = info.get("d_type", "")
                if d_type not in self.dispatch_source_rows: continue
                
                t_val = time_lookup.get(oid, "99:99")
                d_val = dock_lookup.get(oid, info.get("dock", "-"))
                hocha_val = info.get("hocha", "")
                
                self.dispatch_source_rows[d_type].append({
                    "캠프명": info.get("camp", ""), "차량번호": info.get("car_num", ""),
                    "전화번호": info.get("phone", ""), "이름": info.get("driver_name", ""),
                    "간선사": info.get("truckVendor", "") or d_type,
                    "좌측윙": "", "우측윙": "", "후면": "",
                    "_time": t_val,
                    "도크": d_val,
                    "호차": hocha_val
                })
                
            def extract_hocha_num(h_str):
                import re
                m = re.search(r'\d+', str(h_str))
                return int(m.group()) if m else 9999
                
            def get_dock_index(dock_name):
                try:
                    return DEFAULT_DOCKS.index(dock_name)
                except ValueError:
                    return 9999

            # Sort based on user preference (Default: Dock)
            for d_type in self.dispatch_source_rows:
                self.dispatch_source_rows[d_type].sort(key=lambda x: (
                    get_dock_index(x.get("도크", "-")),
                    x["캠프명"],
                    extract_hocha_num(x.get("호차", "")),
                    x["_time"]
                ))
        
        # Capture current selection if exists
        old_tab_idx = 0
        if hasattr(self, "dispatch_nb") and self.dispatch_nb.winfo_exists():
            try: old_tab_idx = self.dispatch_nb.index(self.dispatch_nb.select())
            except: pass

        if hasattr(self, "dispatch_win") and self.dispatch_win.winfo_exists(): self.dispatch_win.destroy()

        C = self.C
        self.dispatch_win = tk.Toplevel(self.root)
        self.dispatch_win.title("출차서 출력")
        self.dispatch_win.geometry("1650x850") # Slightly wider for new column
        self.dispatch_win.configure(bg=C["bg"])
        self.dispatch_win.transient(self.root)

        self._dispatch_active_tab = tk.StringVar(value=DISPATCH_TABS[old_tab_idx])

        hdr = tk.Frame(self.dispatch_win, bg=C["header"], height=48)
        hdr.pack(fill="x");
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🚚  출차서 출력", bg=C["header"], fg="white", font=("맑은 고딕", 12, "bold")).pack(side="left",
                                                                                                     padx=16, pady=12)
        tk.Label(hdr, text="전체 데이터 연동 (출차건 포함)", bg=C["header"], fg=C["header_muted"], font=("맑은 고딕", 9)).pack(
            side="left", padx=8)

        self._btn(hdr, "새로고침", self.open_dispatch_window, C["btn_secondary"]).pack(side="right", padx=(6, 16), pady=10)
        self._btn(hdr, "🖨  일괄 프린트", self.batch_print_dispatch, C["btn_danger"]).pack(side="right", pady=10)

        style = ttk.Style(self.dispatch_win)
        style.configure("Disp.TNotebook", background=C["bg"], borderwidth=0)
        style.configure("Disp.TNotebook.Tab", background=C["bg"], foreground=C["label"], padding=[16, 8],
                        font=("맑은 고딕", 10, "bold"), borderwidth=0)
        style.map("Disp.TNotebook.Tab", background=[("selected", C["white"])],
                  foreground=[("selected", C["btn_primary"])])

        self.dispatch_nb = ttk.Notebook(self.dispatch_win, style="Disp.TNotebook")
        self.dispatch_nb.pack(fill="both", expand=True, padx=0, pady=0)

        def _on_tab_changed(event):
            try:
                idx = self.dispatch_nb.index(self.dispatch_nb.select())
                self._dispatch_active_tab.set(DISPATCH_TABS[idx])
            except Exception:
                pass

        self.dispatch_nb.bind("<<NotebookTabChanged>>", _on_tab_changed)

        self.dispatch_tab_frames.clear()
        self.dispatch_tab_meta.clear()
        for d_type in DISPATCH_TABS:
            tab = tk.Frame(self.dispatch_nb, bg=C["bg"])
            self.dispatch_nb.add(tab, text=d_type)
            self._build_dispatch_tab(tab, d_type)
            
        # Restore tab selection
        self.dispatch_nb.select(old_tab_idx)

    def _build_dispatch_tab(self, parent: tk.Frame, d_type: str) -> None:
        C = self.C
        outer = tk.Frame(parent, bg=C["bg"])
        outer.pack(fill="both", expand=True)

        bottom_bar = tk.Frame(parent, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        bottom_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 10))
        self._btn(bottom_bar, "🪵 고임목 출력", lambda dt=d_type: self.print_goimok(dt), C["btn_warning"]).pack(side="right",
                                                                                                          padx=12,
                                                                                                          pady=8)

        # ── 우측윙 자동 입력 토글 체크박스 ──────────────────────────────────
        auto_right_var = tk.BooleanVar(value=True)

        def _on_auto_toggle():
            pass

        toggle_frame = tk.Frame(bottom_bar, bg=C["white"])
        toggle_frame.pack(side="left", padx=14, pady=6)
        tk.Checkbutton(
            toggle_frame,
            text="봉인씰 자동 입력",
            variable=auto_right_var,
            command=_on_auto_toggle,
            bg=C["white"], fg=C["label"],
            activebackground=C["white"],
            selectcolor=C["white"],
            font=("맑은 고딕", 9),
        ).pack()

        canvas = tk.Canvas(outer, highlightthickness=0, bg=C["bg"])
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=C["bg"])
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        col_px = {"순서": 64, "캠프명": 150, "차량번호": 158, "전화번호": 158, "이름": 118, "간선사": 118, "좌측윙": 138, "우측윙": 138, "후면": 138,
                  "출력하기": 178}
        hdrs = ["순서"] + DISPATCH_HEADERS + ["출력하기"]

        wrap = tk.Frame(scroll_frame, bg=C["bg"])
        wrap.pack(fill="x", padx=14, pady=(12, 8), anchor="n")

        self.dispatch_tab_meta[d_type] = {
            "canvas": canvas,
            "wrap": wrap,
            "auto_right_wing": auto_right_var,
            "col_px": col_px,
            "hdrs": hdrs,
            "rows": []
        }
        
        self._fill_dispatch_tab(d_type)

    def _fill_dispatch_tab(self, d_type: str) -> None:
        """출차서 탭 내의 테이블 내용을 채움 (수동 조정 시 부분 갱신 가능)"""
        C = self.C
        meta = self.dispatch_tab_meta.get(d_type)
        if not meta: return
        
        wrap = meta["wrap"]
        for child in wrap.winfo_children():
            child.destroy()
            
        rows = self.dispatch_source_rows.get(d_type, [])
        tbl = tk.Frame(wrap, bg=C["bg"])
        tbl.pack(anchor="n")
        
        col_px = meta["col_px"]
        hdrs = meta["hdrs"]
        auto_right_var = meta["auto_right_wing"]

        for ci, h in enumerate(hdrs):
            tbl.grid_columnconfigure(ci, minsize=col_px[h], weight=0)
            tk.Label(tbl, text=h, font=("맑은 고딕", 9, "bold"), bg=C["heading"], fg=C["label"], anchor="center",
                     relief="flat", padx=4, pady=8).grid(row=0, column=ci, padx=2, pady=(0, 3), sticky="nsew")

        row_widgets = []
        _entries_ref = []
        ro_keys = {"캠프명", "차량번호", "간선사"}

        for idx, row in enumerate(rows):
            row_vars = {k: tk.StringVar(value=str(row.get(k, ""))) for k in DISPATCH_HEADERS}
            entries = {}
            grid_row = idx + 1
            ro_bg = C["heading"]

            # 순서 조정 버튼 (첫 번째 컬럼)
            order_f = tk.Frame(tbl, bg=C["bg"])
            order_f.grid(row=grid_row, column=0, padx=2, pady=2, sticky="nsew")
            
            btn_up = tk.Button(order_f, text="▲", font=("Arial", 8), width=2, bg=C["white"], fg=C["btn_primary"],
                               command=lambda d=d_type, i=idx: self._move_dispatch_row(d, i, -1))
            btn_up.pack(side="left", padx=1)
            if idx == 0: btn_up.config(state="disabled")

            btn_dn = tk.Button(order_f, text="▼", font=("Arial", 8), width=2, bg=C["white"], fg=C["btn_primary"],
                               command=lambda d=d_type, i=idx: self._move_dispatch_row(d, i, 1))
            btn_dn.pack(side="left", padx=1)
            if idx == len(rows) - 1: btn_dn.config(state="disabled")

            for ci, h in enumerate(DISPATCH_HEADERS):
                is_ro = h in ro_keys
                ent = tk.Entry(tbl, textvariable=row_vars[h], justify="center", relief="solid", bd=1,
                               font=("맑은 고딕", 10), bg=ro_bg if is_ro else C["white"],
                               fg=C["label"] if is_ro else C["text"], readonlybackground=ro_bg, highlightthickness=1,
                               highlightbackground=C["border"], highlightcolor=C["border_focus"])
                if is_ro:
                    ent.config(state="readonly")
                ent.grid(row=grid_row, column=ci + 1, padx=2, pady=2, sticky="nsew", ipady=7)
                entries[h] = ent

            _entries_ref.append(entries)

            af = tk.Frame(tbl, bg=C["bg"])
            af.grid(row=grid_row, column=len(DISPATCH_HEADERS) + 1, padx=6, pady=2, sticky="nsew")
            
            _row_ls_id = row.get("_ls_id", "")
            _seal_btn_color = C.get("btn_success", "#2D7C72") if _row_ls_id else C.get("btn_disabled", "#AAAAAA")
            _seal_btn = self._btn(
                af,
                "봉인씰 입력",
                lambda dv=row_vars, lid=_row_ls_id, did=row.get("_dep_id", ""), aid=row.get("_arr_id", ""), raw_p=row.get("_raw", {}): self._upload_seal_to_ls(
                    ls_id=lid,
                    seal_left=dv["좌측윙"].get(),
                    seal_right=dv["우측윙"].get(),
                    seal_back=dv["후면"].get(),
                    camp_name=dv["캠프명"].get(),
                    dep_id=did,
                    arr_id=aid,
                    raw_payload=raw_p,
                ),
                _seal_btn_color,
            )
            _seal_btn.pack(side="left", padx=(0, 4))
            if not _row_ls_id:
                _seal_btn.config(state="disabled")

            self._btn(af, "프린트", lambda dv=row_vars, dt=d_type: self.print_dispatch_row(dt, dv), C["btn_primary"]).pack(side="left", padx=(0, 4))

            for si, sk in enumerate(["좌측윙", "우측윙", "후면"]):
                entries[sk].bind("<KeyRelease>",
                                 lambda e, t=d_type, r=idx, c=si: self._on_dispatch_seal_keyrelease(t, r, c))
                entries[sk].bind("<Return>", lambda e: "break")

            row_widgets.append({"vars": row_vars, "entries": entries})

        if not rows:
            tk.Label(wrap, text="해당 타입의 데이터가 없습니다.", bg=C["bg"], fg=C["placeholder"], font=("맑은 고딕", 11)).pack(
                pady=40)
        
        meta["rows"] = row_widgets


    def _move_dispatch_row(self, d_type: str, idx: int, delta: int) -> None:
        """출차서 리스트 순서 수동 조정 (Up/Down) - 창 꺼짐 없이 부분 갱신"""
        rows = self.dispatch_source_rows.get(d_type, [])
        if not rows: return
        
        target_idx = idx + delta
        if 0 <= target_idx < len(rows):
            # 스크롤 위치 저장
            meta = self.dispatch_tab_meta.get(d_type)
            canvas = meta.get("canvas") if meta else None
            scroll_pos = canvas.yview()[0] if canvas else 0
            
            # 행 이동 (Swap)
            rows[idx], rows[target_idx] = rows[target_idx], rows[idx]
            
            # 부분 갱신 (창 전체를 다시 그리지 않음)
            self._fill_dispatch_tab(d_type)
            
            # 스크롤 위치 복구
            if canvas:
                # 잠깐의 지연을 주어 렌더링 완료 후 스크롤 이동
                self.root.after(10, lambda: canvas.yview_moveto(scroll_pos))

    def open_dispatch_window(self, keep_manual=False) -> None:
        self.dispatch_source_rows = {k: [] for k in DISPATCH_TABS}

        time_lookup = {}
        for iid in self.tree.get_children():
            v = self.tree.item(iid, "values")
            tags = self.tree.item(iid, "tags")
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = v[7] if len(v) > 7 and v[7] else "99:99"

        for info in self.detached_items.values():
            v = info["values"]
            tags = info["tags"]
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = v[7] if len(v) > 7 and v[7] else "99:99"

        for oid, info in self.dispatch_registry.items():
            d_type = info.get("d_type", "")
            if d_type not in self.dispatch_source_rows:
                continue

            t_val = time_lookup.get(oid, "99:99")
            hocha_val = info.get("hocha", "")

            self.dispatch_source_rows[d_type].append({
                "캠프명": info.get("camp", ""),
                "차량번호": info.get("car_num", ""),
                "전화번호": info.get("phone", ""),
                "이름": info.get("driver_name", ""),
                "간선사": info.get("truckVendor", "") or d_type,
                "좌측윙": "",
                "우측윙": "",
                "후면": "",
                "_time": t_val,
                "호차": hocha_val,
                "_ls_id": info.get("ls_id", ""),
                "_dep_id": info.get("dep_id", ""),
                "_arr_id": info.get("arr_id", ""),
                "_raw": info.get("_raw", {}),
            })

        def extract_hocha_num(h_str):
            import re
            m = re.search(r"\d+", str(h_str))
            return int(m.group()) if m else 9999

        for d_type in self.dispatch_source_rows:
            self.dispatch_source_rows[d_type].sort(
                key=lambda x: (x["캠프명"], extract_hocha_num(x["호차"]), x["_time"])
            )

        if hasattr(self, "dispatch_win") and self.dispatch_win.winfo_exists():
            self.dispatch_win.destroy()

        C = self.C
        self.dispatch_win = tk.Toplevel(self.root)
        self.dispatch_win.title("출차서 출력")
        self.dispatch_win.geometry("1620x820")
        self.dispatch_win.configure(bg=C["bg"])
        self.dispatch_win.transient(self.root)

        self._dispatch_active_tab = tk.StringVar(value=DISPATCH_TABS[0])

        hdr = tk.Frame(self.dispatch_win, bg=C["header"], height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🚚  출차서 출력", bg=C["header"], fg="white", font=("맑은 고딕", 12, "bold")).pack(
            side="left", padx=16, pady=12
        )
        tk.Label(hdr, text="전체 데이터 연동 (출차건 포함)", bg=C["header"], fg=C["header_muted"], font=("맑은 고딕", 9)).pack(
            side="left", padx=8
        )
        self._btn(hdr, "새로고침", self.open_dispatch_window, C["btn_secondary"]).pack(side="right", padx=(6, 16), pady=10)
        self._btn(hdr, "🖨  일괄 프린트", self.batch_print_dispatch, C["btn_danger"]).pack(side="right", pady=10)

        style = ttk.Style(self.dispatch_win)
        style.configure("Disp.TNotebook", background=C["bg"], borderwidth=0)
        style.configure("Disp.TNotebook.Tab", background=C["bg"], foreground=C["label"], padding=[16, 8],
                        font=("맑은 고딕", 10, "bold"), borderwidth=0)
        style.map("Disp.TNotebook.Tab", background=[("selected", C["white"])],
                  foreground=[("selected", C["btn_primary"])])

        nb = ttk.Notebook(self.dispatch_win, style="Disp.TNotebook")
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        def _on_tab_changed(event):
            try:
                idx = nb.index(nb.select())
                self._dispatch_active_tab.set(DISPATCH_TABS[idx])
            except Exception:
                pass

        nb.bind("<<NotebookTabChanged>>", _on_tab_changed)

        self.dispatch_tab_frames.clear()
        self.dispatch_tab_meta.clear()
        for d_type in DISPATCH_TABS:
            tab = tk.Frame(nb, bg=C["bg"])
            nb.add(tab, text=d_type)
            self._build_dispatch_tab(tab, d_type)

    def _build_dispatch_tab(self, parent: tk.Frame, d_type: str) -> None:
        C = self.C
        rows = self.dispatch_source_rows.get(d_type, [])

        outer = tk.Frame(parent, bg=C["bg"])
        outer.pack(fill="both", expand=True)

        bottom_bar = tk.Frame(parent, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        bottom_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 10))
        self._btn(bottom_bar, "📦 고임목 출력", lambda dt=d_type: self.print_goimok(dt), C["btn_warning"]).pack(
            side="right", padx=12, pady=8
        )

        auto_right_var = tk.BooleanVar(value=True)

        def _apply_right_wing_state(is_auto: bool, entries_list: list) -> None:
            pass

        _entries_ref: list = []

        def _on_auto_toggle():
            is_auto = auto_right_var.get()
            _apply_right_wing_state(is_auto, _entries_ref)

        toggle_frame = tk.Frame(bottom_bar, bg=C["white"])
        toggle_frame.pack(side="left", padx=14, pady=6)
        tk.Checkbutton(
            toggle_frame,
            text="우측윙 자동 입력",
            variable=auto_right_var,
            command=_on_auto_toggle,
            bg=C["white"],
            fg=C["label"],
            activebackground=C["white"],
            selectcolor=C["white"],
            font=("맑은 고딕", 9),
        ).pack()

        canvas = tk.Canvas(outer, highlightthickness=0, bg=C["bg"])
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=C["bg"])
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _dispatch_mousewheel(event):
            if getattr(event, "delta", 0):
                canvas.yview_scroll(int(-event.delta / 120), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        def _bind_dispatch_wheel(_event=None):
            canvas.bind_all("<MouseWheel>", _dispatch_mousewheel)
            canvas.bind_all("<Button-4>", _dispatch_mousewheel)
            canvas.bind_all("<Button-5>", _dispatch_mousewheel)

        def _unbind_dispatch_wheel(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_dispatch_wheel)
        canvas.bind("<Leave>", _unbind_dispatch_wheel)
        scroll_frame.bind("<Enter>", _bind_dispatch_wheel)
        scroll_frame.bind("<Leave>", _unbind_dispatch_wheel)

        col_px = {"캠프명": 150, "차량번호": 158, "전화번호": 158, "이름": 118, "간선사": 118, "좌측윙": 138, "우측윙": 138, "후면": 138,
                  "출력하기": 178}
        hdrs = DISPATCH_HEADERS + ["출력하기"]

        wrap = tk.Frame(scroll_frame, bg=C["bg"])
        wrap.pack(fill="x", padx=14, pady=(12, 8), anchor="n")
        tbl = tk.Frame(wrap, bg=C["bg"])
        tbl.pack(anchor="n")

        for ci, h in enumerate(hdrs):
            tbl.grid_columnconfigure(ci, minsize=col_px[h], weight=0)
            tk.Label(tbl, text=h, font=("맑은 고딕", 9, "bold"), bg=C["heading"], fg=C["label"], anchor="center",
                     relief="flat", padx=4, pady=8).grid(row=0, column=ci, padx=2, pady=(0, 3), sticky="nsew")

        row_widgets = []
        ro_keys = {"캠프명", "차량번호", "간선사"}

        for idx, row in enumerate(rows):
            row_vars = {k: tk.StringVar(value=str(row.get(k, ""))) for k in DISPATCH_HEADERS}
            entries = {}
            grid_row = idx + 1
            ro_bg = C["heading"]

            for ci, h in enumerate(DISPATCH_HEADERS):
                is_ro = h in ro_keys
                ent = tk.Entry(tbl, textvariable=row_vars[h], justify="center", relief="solid", bd=1,
                               font=("맑은 고딕", 10), bg=ro_bg if is_ro else C["white"],
                               fg=C["label"] if is_ro else C["text"], readonlybackground=ro_bg, highlightthickness=1,
                               highlightbackground=C["border"], highlightcolor=C["border_focus"])
                if is_ro:
                    ent.config(state="readonly")
                ent.grid(row=grid_row, column=ci, padx=2, pady=2, sticky="nsew", ipady=7)
                entries[h] = ent

            _entries_ref.append(entries)

            af = tk.Frame(tbl, bg=C["bg"])
            af.grid(row=grid_row, column=len(DISPATCH_HEADERS), padx=6, pady=2, sticky="nsew")
            
            _row_ls_id = row.get("_ls_id", "")
            _seal_btn_color = C.get("btn_success", "#2D7C72") if _row_ls_id else C.get("btn_disabled", "#AAAAAA")
            _seal_btn = self._btn(
                af,
                "봉인씰 입력",
                lambda dv=row_vars, lid=_row_ls_id, did=row.get("_dep_id", ""), aid=row.get("_arr_id", ""), raw_p=row.get("_raw", {}): self._upload_seal_to_ls(
                    ls_id=lid,
                    seal_left=dv["좌측윙"].get(),
                    seal_right=dv["우측윙"].get(),
                    seal_back=dv["후면"].get(),
                    camp_name=dv["캠프명"].get(),
                    dep_id=did,
                    arr_id=aid,
                    raw_payload=raw_p,
                ),
                _seal_btn_color,
            )
            _seal_btn.pack(side="left", padx=(0, 4))
            if not _row_ls_id:
                _seal_btn.config(state="disabled")

            self._btn(af, "프린트", lambda dv=row_vars, dt=d_type, lid=row.get("_ls_id", ""): self.print_dispatch_row(dt, dv, lid), C["btn_primary"]).pack(
                side="left", padx=(0, 4)
            )

            for si, sk in enumerate(["좌측윙", "우측윙", "후면"]):
                entries[sk].bind("<KeyRelease>",
                                 lambda e, t=d_type, r=idx, c=si: self._on_dispatch_seal_keyrelease(t, r, c))
                entries[sk].bind("<Return>", lambda e: "break")

            row_widgets.append({"vars": row_vars, "entries": entries, "_ls_id": row.get("_ls_id", ""), "_dep_id": row.get("_dep_id", ""), "_arr_id": row.get("_arr_id", "")})

        if not rows:
            tk.Label(scroll_frame, text="해당 타입의 데이터가 없습니다.", bg=C["bg"], fg=C["placeholder"], font=("맑은 고딕", 11)).pack(
                pady=40
            )
        self.dispatch_tab_meta[d_type] = {"rows": row_widgets, "canvas": canvas, "auto_right_wing": auto_right_var}

    # ── 고임목 출력 (ExcelManager 연동) ──────────────────────────────────────────────────────────
    def _get_goimok_template_path(self) -> Optional[str]:
        path = resource_path(GOIMOK_TEMPLATE_FILENAME)
        if os.path.exists(path): return path
        path2 = os.path.join(os.path.abspath("."), GOIMOK_TEMPLATE_FILENAME)
        if os.path.exists(path2): return path2
        chosen = filedialog.askopenfilename(title="고임목.xlsx 파일 선택", filetypes=[("Excel files", "*.xlsx")])
        return chosen if chosen else None

    def print_goimok(self, d_type: str) -> None:
        tab_meta = self.dispatch_tab_meta.get(d_type)
        if not tab_meta: return
        rows = tab_meta.get("rows", [])
        data_rows = [(d_type, r["vars"]["캠프명"].get().strip(), r["vars"]["차량번호"].get().strip()) for r in rows if
                     r["vars"]["캠프명"].get().strip() or r["vars"]["차량번호"].get().strip()]
        if not data_rows:
            messagebox.showinfo("안내", f"[{d_type}] 탭에 기입된 데이터가 없습니다.")
            return

        goimok_path = self._get_goimok_template_path()
        if not goimok_path: return
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("오류", "openpyxl 라이브러리가 필요합니다.")
            return

        temp_fd, temp_path = tempfile.mkstemp(prefix="goimok_", suffix=".xlsx")
        os.close(temp_fd)
        shutil.copy2(goimok_path, temp_path)

        try:
            wb = load_workbook(temp_path)
            ws = wb.active

            for idx, (dt, camp, car_num) in enumerate(data_rows):
                row_num = 2 + idx
                cell_a, cell_b, cell_c = ws.cell(row=row_num, column=1), ws.cell(row=row_num, column=2), ws.cell(
                    row=row_num, column=3)
                font_a, font_b, font_c = copy(cell_a.font) if cell_a.font else None, copy(
                    cell_b.font) if cell_b.font else None, copy(cell_c.font) if cell_c.font else None
                cell_a.value, cell_b.value, cell_c.value = dt, camp, car_num
                if font_a: cell_a.font = font_a
                if font_b: cell_b.font = font_b
                if font_c: cell_c.font = font_c

            wb.save(temp_path)
            wb.close()

            printed = False
            if platform.system() == "Windows" and WIN32COM_AVAILABLE:
                try:
                    pythoncom.CoInitialize()
                    xl = win32.DispatchEx("Excel.Application")
                    xl.Visible = False;
                    xl.DisplayAlerts = False
                    wb2 = xl.Workbooks.Open(os.path.abspath(temp_path))
                    wb2.ActiveSheet.PrintOut()
                    wb2.Close(SaveChanges=False)
                    xl.Quit()
                    printed = True
                except Exception as e:
                    print(f"[디버그] 고임목 Excel COM 인쇄 실패: {e}")
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

            if not printed:
                try:
                    pdf_path = self.excel_manager.convert_xlsx_to_pdf_lo(temp_path)
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup([temp_path, pdf_path, os.path.dirname(pdf_path)])
                    messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
                    return
                except Exception as e:
                    print(f"[출력 디버그] 출차서 LibreOffice PDF 변환 실패: {e}")

                if OPENPYXL_AVAILABLE and PILLOW_AVAILABLE:
                    img = self._render_dispatch_sheet_image(temp_path, ws.title)
                    pdf_path = self.excel_manager.save_image_as_pdf(img)
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup([temp_path, pdf_path])
                    messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
                    return

                messagebox.showerror("오류", "인쇄 가능한 방법을 찾지 못했습니다.\nwin32com, LibreOffice, 또는 Pillow 중 하나가 필요합니다.")
                return

            messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
            self._schedule_temp_cleanup([temp_path])

        except Exception as e:
            self._safe_remove_paths(temp_path)
            messagebox.showerror("고임목 출력 오류", str(e))

    def _dispatch_seal_keys(self) -> List[str]:
        return ["좌측윙", "우측윙", "후면"]

    def _on_dispatch_seal_keyrelease(self, d_type: str, row_index: int, seal_col_index: int) -> None:
        tab_meta = self.dispatch_tab_meta.get(d_type)
        if not tab_meta: return
        rows = tab_meta["rows"]
        if row_index >= len(rows): return

        seal_keys = self._dispatch_seal_keys()
        key_name = seal_keys[seal_col_index]
        var = rows[row_index]["vars"][key_name]

        raw_value = re.sub(r"\D", "", var.get())[:8]
        if var.get() != raw_value: var.set(raw_value)
        if len(raw_value) == 8:
            auto_right = tab_meta.get("auto_right_wing")
            is_auto = auto_right.get() if auto_right is not None else True
            if is_auto:
                # 자동 모드: 좌측윙 기준으로 나머지 필드 자동 전파 후 커서 이동
                self._propagate_dispatch_seals_linear(d_type, row_index, seal_col_index, int(raw_value))
            # 수동/자동 모두 8자리 완성 시 커서 이동
            self._focus_next_dispatch_seal(d_type, row_index, seal_col_index)

    def _propagate_dispatch_seals_linear(self, d_type: str, start_row: int, start_col: int, start_val: int) -> None:
        """봉인번호 자동 전파 (자동 모드 전용).
        start_col 위치부터 모든 행의 좌측윙/우측윙/후면을 순서대로 +1씩 채운다.
        """
        tab_meta = self.dispatch_tab_meta.get(d_type)
        if not tab_meta: return
        rows = tab_meta["rows"]
        seal_keys = self._dispatch_seal_keys()  # ["좌측윙", "우측윙", "후면"]

        current_val = start_val
        for r in range(start_row, len(rows)):
            c_start = start_col if r == start_row else 0
            for c in range(c_start, len(seal_keys)):
                rows[r]["vars"][seal_keys[c]].set(f"{current_val:08d}")
                current_val += 1

    def _focus_next_dispatch_seal(self, d_type: str, row_index: int, seal_col_index: int) -> None:
        tab_meta = self.dispatch_tab_meta.get(d_type)
        if not tab_meta: return
        rows = tab_meta["rows"]
        seal_keys = self._dispatch_seal_keys()

        next_row_index, next_col_index = row_index, seal_col_index + 1
        if next_col_index >= len(seal_keys): next_row_index, next_col_index = next_row_index + 1, 0
        if next_row_index >= len(rows): return

        try:
            next_entry = rows[next_row_index]["entries"][seal_keys[next_col_index]]
            next_entry.focus_set();
            next_entry.icursor(tk.END);
            next_entry.selection_range(0, tk.END)
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════════
    # 출차서 엔진 — ExcelManager 연동
    # ══════════════════════════════════════════════════════════════════════════
    def _get_master_config_path(self) -> Optional[str]:
        path = resource_path(MASTER_CONFIG_FILENAME)
        if os.path.exists(path): return path
        path2 = os.path.join(os.path.abspath("."), MASTER_CONFIG_FILENAME)
        if os.path.exists(path2): return path2
        chosen = filedialog.askopenfilename(title="마스터_양식설정.xlsx 파일 선택", filetypes=[("Excel files", "*.xlsx")])
        return chosen if chosen else None

    def _get_dispatch_template_path(self) -> Optional[str]:
        path = resource_path(DISPATCH_TEMPLATE_FILENAME)
        if os.path.exists(path): return path
        path2 = os.path.join(os.path.abspath("."), DISPATCH_TEMPLATE_FILENAME)
        if os.path.exists(path2): return path2
        chosen = filedialog.askopenfilename(title="출차서.xlsx 파일 선택", filetypes=[("Excel files", "*.xlsx")])
        return chosen if chosen else None

    def _write_dispatch_and_print(self, template_path: str, sheet_name: str, cell_map: Dict[str, str],
                                  values: Dict[str, str], action: str) -> None:
        if not OPENPYXL_AVAILABLE: raise RuntimeError("openpyxl 라이브러리가 필요합니다.")

        temp_fd, temp_path = tempfile.mkstemp(prefix="dispatch_", suffix=".xlsx")
        os.close(temp_fd)
        shutil.copy2(template_path, temp_path)
        cleanup = [temp_path]

        try:
            wb = load_workbook(temp_path)
            target_ws = None
            norm_sheet = self._normalize_text(sheet_name)
            for sname in wb.sheetnames:
                if self._normalize_text(sname) == norm_sheet:
                    target_ws = wb[sname]
                    break
            if target_ws is None:
                for sname in wb.sheetnames:
                    if norm_sheet in self._normalize_text(sname) or self._normalize_text(sname) in norm_sheet:
                        target_ws = wb[sname]
                        break
            if target_ws is None:
                raise ValueError(f"출차서.xlsx 에서 시트 '{sheet_name}' 를 찾지 못했습니다.\n존재하는 시트: {', '.join(wb.sheetnames)}")

            for sn in wb.sheetnames:
                if sn != target_ws.title: wb[sn].sheet_state = "hidden"
            wb.active = wb.index(target_ws)

            def write_cell(cell_addr: str, val: str):
                if not cell_addr or not val: return
                try:
                    cell = target_ws[cell_addr.upper()]
                    original_font = copy(cell.font) if cell.font else None
                    cell.value = val
                    if original_font: cell.font = original_font
                except Exception as e:
                    print(f"[출력 디버그] 출차서 미리보기 PDF 변환 실패: {e}")

            for idx in range(1, 5):
                postfix = str(idx)
                write_cell(cell_map.get(f"운행일자{postfix}", ""), values.get(f"운행일자{postfix}", ""))
                write_cell(cell_map.get(f"차량번호{postfix}", ""), values.get(f"차량번호{postfix}", ""))
                write_cell(cell_map.get(f"이름{postfix}", ""), values.get(f"이름{postfix}", ""))
                write_cell(cell_map.get(f"연락처{postfix}", ""), values.get(f"전화번호{postfix}", ""))
                write_cell(cell_map.get(f"좌측윙{postfix}", ""), values.get(f"좌측윙{postfix}", ""))
                write_cell(cell_map.get(f"우측윙{postfix}", ""), values.get(f"우측윙{postfix}", ""))
                write_cell(cell_map.get(f"후면 도어{postfix}", ""), values.get(f"후면 도어{postfix}", ""))
                write_cell(cell_map.get(f"간선사{postfix}", ""), values.get(f"간선사{postfix}", ""))

            wb.save(temp_path)
            wb.close()

            if action == "print":
                if platform.system() == "Windows" and WIN32COM_AVAILABLE:
                    temp_fd2, temp_com_path = tempfile.mkstemp(prefix="dispatch_com_", suffix=".xlsx")
                    os.close(temp_fd2)
                    shutil.copy2(template_path, temp_com_path)
                    cleanup.append(temp_com_path)
                    try:
                        pythoncom.CoInitialize()
                        xl = win32.DispatchEx("Excel.Application")
                        xl.Visible = False;
                        xl.DisplayAlerts = False
                        wb2 = xl.Workbooks.Open(os.path.abspath(temp_com_path))

                        com_ws = None
                        for i in range(1, wb2.Worksheets.Count + 1):
                            if self._normalize_text(wb2.Worksheets(i).Name) == self._normalize_text(sheet_name):
                                com_ws = wb2.Worksheets(i);
                                break
                        if com_ws is None:
                            for i in range(1, wb2.Worksheets.Count + 1):
                                ns, nw = self._normalize_text(sheet_name), self._normalize_text(wb2.Worksheets(i).Name)
                                if ns in nw or nw in ns:
                                    com_ws = wb2.Worksheets(i);
                                    break
                        if com_ws is None: com_ws = wb2.ActiveSheet

                        def com_write(addr, val):
                            if addr and val:
                                try:
                                    com_ws.Range(addr.upper()).Value = val
                                except Exception:
                                    pass

                        for idx in range(1, 5):
                            postfix = str(idx)
                            com_write(cell_map.get(f"운행일자{postfix}", ""), values.get(f"운행일자{postfix}", ""))
                            com_write(cell_map.get(f"차량번호{postfix}", ""), values.get(f"차량번호{postfix}", ""))
                            com_write(cell_map.get(f"이름{postfix}", ""), values.get(f"이름{postfix}", ""))
                            com_write(cell_map.get(f"연락처{postfix}", ""), values.get(f"전화번호{postfix}", ""))
                            com_write(cell_map.get(f"좌측윙{postfix}", ""), values.get(f"좌측윙{postfix}", ""))
                            com_write(cell_map.get(f"우측윙{postfix}", ""), values.get(f"우측윙{postfix}", ""))
                            com_write(cell_map.get(f"후면 도어{postfix}", ""), values.get(f"후면 도어{postfix}", ""))
                            com_write(cell_map.get(f"바코드{postfix}", ""), values.get(f"바코드{postfix}", ""))
                            com_write(cell_map.get(f"간선사{postfix}", ""), values.get(f"간선사{postfix}", ""))

                        com_ws.PrintOut()
                        wb2.Close(SaveChanges=False)
                        xl.Quit()
                        return
                    except Exception as e:
                        print(f"[디버그] Excel COM 인쇄 실패 원인: {e}")
                    finally:
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass

                try:
                    pdf_path = self.excel_manager.convert_xlsx_to_pdf_lo(temp_path)
                    cleanup.extend([pdf_path, os.path.dirname(pdf_path)])
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup(cleanup)
                    return
                except Exception:
                    pass

                if OPENPYXL_AVAILABLE and PILLOW_AVAILABLE:
                    img = self._render_dispatch_sheet_image(temp_path, sheet_name)
                    pdf_path = self.excel_manager.save_image_as_pdf(img)
                    cleanup.append(pdf_path)
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup(cleanup)
                    return

                raise RuntimeError("인쇄 가능한 방법을 찾지 못했습니다.\nwin32com, LibreOffice, 또는 Pillow 중 하나가 필요합니다.")

            elif action == "preview":
                try:
                    pdf_path = self.excel_manager.convert_xlsx_to_pdf_lo(temp_path)
                    cleanup.extend([pdf_path, os.path.dirname(pdf_path)])
                    self._show_dispatch_pdf_preview(pdf_path, cleanup)
                    return
                except Exception:
                    pass

                if OPENPYXL_AVAILABLE and PILLOW_AVAILABLE:
                    img = self._render_dispatch_sheet_image(temp_path, sheet_name)
                    pdf_path = self.excel_manager.save_image_as_pdf(img)
                    cleanup.append(pdf_path)
                    self._show_dispatch_pdf_preview(pdf_path, cleanup)
                    return

                raise RuntimeError("미리보기를 위해 LibreOffice 또는 Pillow 라이브러리가 필요합니다.")

        except Exception:
            self._safe_remove_paths(*cleanup)
            raise

    def _run_dispatch_action(self, payloads: List[Dict[str, str]], action: str) -> None:
        if not payloads: return
        camp_name = payloads[0].get("캠프명", "").strip()
        if not camp_name: raise ValueError("캠프명이 비어 있습니다.")

        master_path = self._get_master_config_path()
        if not master_path: raise RuntimeError("마스터_양식설정.xlsx 파일을 선택하지 않았습니다.")

        template_path = self._get_dispatch_template_path()
        if not template_path: raise RuntimeError("출차서.xlsx 파일을 선택하지 않았습니다.")

        master_cfg = self.excel_manager.load_master_config(master_path)
        matched_key = self._find_best_match_key(camp_name, list(master_cfg.keys()))
        matched_cfg = master_cfg.get(matched_key) if matched_key else None
        if matched_cfg is None: raise ValueError(
            f"마스터_양식설정.xlsx 에서 캠프 '{camp_name}' 을 찾지 못했습니다.\n등록된 캠프: {', '.join(list(master_cfg.keys())[:10])}...")

        sheet_name = matched_cfg.get("출차서_시트명", "").strip()
        if not sheet_name: raise ValueError(f"캠프 '{camp_name}' 의 출차서_시트명이 비어 있습니다.")

        run_date = getattr(self, "api_run_date", "")
        values = {}
        
        # 4개의 슬롯(1~4)에 데이터를 채웁니다.
        # 데이터가 4개 미만일 경우, 나머지 빈 슬롯은 모두 첫 번째 데이터(payloads[0])로 채워 동일하게 출력되도록 합니다.
        for i in range(4):
            postfix = str(i + 1)
            payload = payloads[i] if i < len(payloads) else payloads[0]
            
            values[f"운행일자{postfix}"] = run_date
            values[f"차량번호{postfix}"] = payload.get("차량번호", "")
            values[f"이름{postfix}"] = payload.get("이름", "")
            values[f"전화번호{postfix}"] = payload.get("전화번호", "")
            values[f"좌측윙{postfix}"] = payload.get("좌측윙", "")
            values[f"우측윙{postfix}"] = payload.get("우측윙", "")
            values[f"후면 도어{postfix}"] = payload.get("후면", "")
            values[f"바코드{postfix}"] = payload.get("바코드", "")
            values[f"간선사{postfix}"] = payload.get("간선사", "")

        print(f"[디버그] {camp_name} 출차서 작성 - 매핑변수: {matched_cfg}")
        print(f"[디버그] {camp_name} 출차서 작성 - 작성값: {values}")

        self._write_dispatch_and_print(template_path, sheet_name, matched_cfg, values, action)

    def preview_dispatch_row(self, d_type: str, row_vars: Dict[str, Any]) -> None:
        try:
            payload = self._row_vars_to_payload(row_vars)
            self._run_dispatch_action(self._expand_dispatch_payloads(d_type, [payload]), "preview")
        except Exception as e:
            messagebox.showerror("미리보기 오류", str(e))

    def _upload_seal_to_ls(self, ls_id: str, seal_left: str, seal_right: str, seal_back: str,
                            camp_name: str = "", dep_id: str = "", arr_id: str = "", raw_payload: dict = None) -> None:
        """truckRequestId(ls_id)에 봉인씰 번호 3종을 LS 포털 API로 업로드합니다."""
        # ── 세션 확인 ──────────────────────────────────────────────────────────
        if not hasattr(self, "api_manager") or getattr(self.api_manager, "session", None) is None:
            messagebox.showerror(
                "LS 씰 등록 오류",
                "API 세션이 없습니다.\n메인 화면에서 'LS불러오기'를 통해 먼저 로그인(인증)을 진행해주세요."
            )
            return

        if not ls_id:
            messagebox.showerror(
                "LS 씰 등록 오류",
                f"[{camp_name}] truckRequestId(LS 요청 ID)가 없습니다.\n'LS불러오기' 후 데이터를 다시 불러와 주세요."
            )
            return

        # ── 봉인씰 입력 여부 확인 ──────────────────────────────────────────────
        filled = [v for v in [seal_left, seal_right, seal_back] if v.strip()]
        if not filled:
            messagebox.showwarning(
                "LS 씰 등록",
                f"[{camp_name}] 좌측윙, 우측윙, 후면 봉인씰 번호가 모두 비어 있습니다.\n봉인씰 번호를 먼저 입력한 후 등록해 주세요."
            )
            return

        if len(filled) < 3:
            missing = []
            if not seal_left.strip():  missing.append("좌측윙")
            if not seal_right.strip(): missing.append("우측윙")
            if not seal_back.strip():  missing.append("후면")
            if not messagebox.askyesno(
                "LS 씰 등록 확인",
                f"[{camp_name}] 다음 봉인씰이 비어 있습니다: {', '.join(missing)}\n빈 값도 포함하여 등록하시겠습니까?"
            ):
                return

        # ── PUT 요청 ─────────────────────────────────────────────────────────
        url = f"https://ls.coupang.net/gateway-api/v1/truckOrderTracking/{ls_id}"
        
        rp = raw_payload or {}
        
        final_dep_id = dep_id if str(dep_id).strip() else rp.get("departureDetailId")
        if str(final_dep_id).isdigit(): final_dep_id = int(final_dep_id)
        elif not str(final_dep_id).strip(): final_dep_id = None
        
        final_arr_id = arr_id if str(arr_id).strip() else rp.get("arrivalDetailId")
        if str(final_arr_id).isdigit(): final_arr_id = int(final_arr_id)
        elif not str(final_arr_id).strip(): final_arr_id = None

        payload = {
            "truckRequestId": int(ls_id) if str(ls_id).isdigit() else ls_id,
            "departureDetailId": final_dep_id,
            "arrivalDetailId": final_arr_id,
            "dockArriveTime": rp.get("dockArriveTime"),
            "loadingTime": rp.get("loadingTime"),
            "dockDepartureTime": rp.get("dockDepartureTime"),
            "parcelSent": rp.get("parcelSent"),
            "departureWaitingTime": rp.get("departureWaitingTime"),
            "sealLeftNo": seal_left.strip() or None,
            "sealRightNo": seal_right.strip() or None,
            "sealBackNo": seal_back.strip() or None,
            "arrivalTime": rp.get("arrivalTime"),
            "unloadingTime": rp.get("unloadingTime"),
            "unloadCompletedTime": rp.get("unloadCompletedTime", rp.get("unloadComplete")),
            "arrivalWaitingTime": rp.get("arrivalWaitingTime"),
            "parcelReceived": rp.get("parcelReceived")
        }
        
        print(f"[LS 씰 등록] PUT {url} | payload={payload}")

        try:
            result = self.api_manager.put_data(url, payload)
            print(f"[LS 씰 등록] 응답: {result}")
            messagebox.showinfo(
                "LS 씰 등록 완료",
                f"[{camp_name}] 봉인씰 번호가 LS 포털에 성공적으로 등록되었습니다.\n"
                f"좌측윙: {seal_left.strip()}\n우측윙: {seal_right.strip()}\n후면: {seal_back.strip()}"
            )
        except PermissionError as e:
            messagebox.showerror("인증 오류", str(e))
        except Exception as e:
            messagebox.showerror("LS 씰 등록 오류", f"[{camp_name}] 봉인씰 등록 중 오류가 발생했습니다.\n\n{e}")

    def _print_ls(self, payloads: List[Dict[str, Any]]) -> None:

        if not hasattr(self, "api_manager") or getattr(self.api_manager, "session", None) is None:
            messagebox.showerror("오류", "API 세션이 없습니다. 메인 화면에서 'LS불러오기'를 통해 먼저 로그인(인증)을 진행해주세요.")
            return

        success = 0
        errors = []
        cleanup_files = []
        
        for payload in payloads:
            ls_id = payload.get("_ls_id")
            if not ls_id:
                errors.append(f"[{payload.get('캠프명', '')}] LS 요청ID(truckRequestId)가 없어 출력할 수 없습니다.")
                continue
                
            url = f"https://ls.coupang.net/data/linehaul/slip/generate?truckRequestId={ls_id}&locale=ko_KR"
            try:
                resp = self.api_manager.session.get(url, timeout=30)
                resp.raise_for_status()
                
                temp_fd, temp_pdf = tempfile.mkstemp(prefix=f"ls_print_{ls_id}_", suffix=".pdf")
                with os.fdopen(temp_fd, 'wb') as f:
                    f.write(resp.content)
                cleanup_files.append(temp_pdf)
                
                self.excel_manager.send_pdf_to_printer(temp_pdf)
                success += 1
            except Exception as e:
                errors.append(f"[{payload.get('캠프명', '')}] 다운로드/출력 오류: {e}")
                
        if cleanup_files:
            self._schedule_temp_cleanup(cleanup_files)
            
        if errors:
            err_msg = "\n".join(errors[:5]) + ("\n..." if len(errors) > 5 else "")
            messagebox.showwarning("완료 (일부 오류)", f"{success}건 프린트 요청 완료, {len(errors)}건 실패.\n\n[오류 내역]\n{err_msg}")
        else:
            messagebox.showinfo("완료", f"총 {success}건 LS 방식 프린트 요청이 전송되었습니다.")

    def print_dispatch_row(self, d_type: str, row_vars: Dict[str, Any], ls_id: str = "") -> None:
        try:
            payload = self._row_vars_to_payload(row_vars)
            payload["_ls_id"] = ls_id

            use_ls_print = False
            if hasattr(self, "ls_print_var"):
                use_ls_print = self.ls_print_var.get()

            if use_ls_print:
                if not messagebox.askyesno("확인", f"[{payload.get('캠프명', '')}] LS 방식으로 프린트하시겠습니까?"):
                    return
                self._print_ls([payload])
                return

            self._run_dispatch_action(self._expand_dispatch_payloads(d_type, [payload]), "print")
            messagebox.showinfo("완료", "출차서 프린트 요청이 전송되었습니다.")
        except Exception as e:
            messagebox.showerror("프린트 오류", str(e))

    def batch_print_dispatch(self) -> None:
        if not hasattr(self, "dispatch_tab_meta") or not self.dispatch_tab_meta: return
        active_tab = getattr(self, "_dispatch_active_tab", None)
        d_type = active_tab.get() if active_tab else DISPATCH_TABS[0]
        rows = self.dispatch_tab_meta.get(d_type, {}).get("rows", [])
        payloads = [self._row_vars_to_payload(r["vars"]) for r in rows if
                    self._row_vars_to_payload(r["vars"]).get("캠프명")]

        expanded_payloads = []
        for p in payloads:
            expanded_payloads.append(p)
            camp_name = p.get("캠프명", "")
            if d_type in ["SD2", "D1"] and self._normalize_text(camp_name) in ["동인천2", "동인천2(SH)"]:
                p2 = p.copy()
                p2["캠프명"] = "인천3"
                expanded_payloads.append(p2)
        payloads = expanded_payloads

        if not payloads:
            messagebox.showinfo("안내", f"[{d_type}] 탭에 프린트할 데이터가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"[{d_type}] 탭의 총 {len(payloads)}건을 일괄 프린트하시겠습니까?"): return

        if not (platform.system() == "Windows" and WIN32COM_AVAILABLE):
            messagebox.showerror("오류", "일괄 프린트 고속 모드는 Windows 환경에서 win32com 을 통해서만 지원됩니다.")
            return

        master_path = self._get_master_config_path()
        if not master_path: return
        template_path = self._get_dispatch_template_path()
        if not template_path: return
        
        try:
            master_cfg = self.excel_manager.load_master_config(master_path)
        except Exception as e:
            messagebox.showerror("양식 설정 오류", str(e))
            return

        run_date = getattr(self, "api_run_date", "")
        success, errors = 0, []
        
        from collections import defaultdict
        
        # Determine logical hocha indices and resolve template sheet for each payload
        camp_counts = defaultdict(int)
        for p in payloads:
            if "간선사" not in p or not p["간선사"]:
                p["간선사"] = d_type
                
            camp_name = p.get("캠프명", "")
            camp_counts[camp_name] += 1
            p["_hocha_idx"] = camp_counts[camp_name]
            
            matched_key = self._find_best_match_key(camp_name, list(master_cfg.keys()))
            p["_cfg"] = master_cfg.get(matched_key) if matched_key else None
            p["_sheet"] = p["_cfg"].get("출차서_시트명", "").strip() if p["_cfg"] else ""

        # Sort all payloads by (sheet name, hocha index, camp name)
        # This achieves: all 1st cars of the same template print first, then all 2nd cars of the same template, etc.
        payloads.sort(key=lambda x: (x["_sheet"], x["_hocha_idx"], x["캠프명"]))

        try:
            pythoncom.CoInitialize()
            xl = win32.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
        except Exception as e:
            messagebox.showerror("엑셀 실행 오류", f"엑셀(COM)을 실행할 수 없습니다:\n{e}")
            try: pythoncom.CoUninitialize()
            except Exception: pass
            return

        temp_fd, temp_com_path = tempfile.mkstemp(prefix="batch_dispatch_master_", suffix=".xlsx")
        os.close(temp_fd)
        shutil.copy2(template_path, temp_com_path)
        cleanup_files = [temp_com_path]

        try:
            wb2 = xl.Workbooks.Open(os.path.abspath(temp_com_path))
        except Exception as e:
            messagebox.showerror("템플릿 로드 오류", f"출차서 템플릿을 열 수 없습니다:\n{e}")
            try: xl.Quit()
            except Exception: pass
            try: pythoncom.CoUninitialize()
            except Exception: pass
            self._schedule_temp_cleanup(cleanup_files)
            return

        try:
            last_sheet_name = None
            com_ws = None
            
            for payload in payloads:
                camp_name = payload.get("캠프명", "")
                matched_cfg = payload.get("_cfg")
                sheet_name = payload.get("_sheet", "")
                
                if not matched_cfg:
                    errors.append(f"[{camp_name}] 마스터_양식설정에서 이 캠프를 찾을 수 없습니다.")
                    continue
                    
                if not sheet_name:
                    errors.append(f"[{camp_name}] 출차서_시트명이 지정되지 않았습니다.")
                    continue

                if sheet_name != last_sheet_name or com_ws is None:
                    last_sheet_name = sheet_name
                    com_ws = None
                    for si in range(1, wb2.Worksheets.Count + 1):
                        ns, nw = self._normalize_text(sheet_name), self._normalize_text(wb2.Worksheets(si).Name)
                        if ns == nw or ns in nw or nw in ns:
                            com_ws = wb2.Worksheets(si)
                            break
                    if com_ws is None:
                        com_ws = wb2.ActiveSheet

                def _current_com_write(addr, val, ws=com_ws):
                    if addr:
                        try: ws.Range(addr.upper()).Value = val if val else ""
                        except Exception: pass

                try:
                    for idx in range(4):
                        postfix = str(idx + 1)
                        
                        _current_com_write(matched_cfg.get(f"운행일자{postfix}", ""), run_date)
                        _current_com_write(matched_cfg.get(f"차량번호{postfix}", ""), payload.get("차량번호", ""))
                        _current_com_write(matched_cfg.get(f"이름{postfix}", ""), payload.get("이름", ""))
                        _current_com_write(matched_cfg.get(f"연락처{postfix}", ""), payload.get("전화번호", ""))
                        _current_com_write(matched_cfg.get(f"좌측윙{postfix}", ""), payload.get("좌측윙", ""))
                        _current_com_write(matched_cfg.get(f"우측윙{postfix}", ""), payload.get("우측윙", ""))
                        _current_com_write(matched_cfg.get(f"후면 도어{postfix}", ""), payload.get("후면", ""))
                        _current_com_write(matched_cfg.get(f"바코드{postfix}", ""), payload.get("바코드", ""))
                        _current_com_write(matched_cfg.get(f"간선사{postfix}", ""), payload.get("간선사", ""))
                        
                    com_ws.PrintOut()
                    success += 1
                except Exception as e:
                    errors.append(f"[{camp_name}] 출력 중 오류: {e}")

        finally:
            try: wb2.Close(SaveChanges=False)
            except Exception: pass
            try: xl.Quit()
            except Exception: pass
            try: pythoncom.CoUninitialize()
            except Exception: pass
            
        self._schedule_temp_cleanup(cleanup_files)

        if errors:
            messagebox.showwarning("일괄 프린트 완료", f"성공: {success}건\n실패: {len(errors)}건\n\n" + "\n".join(errors[:10]))
        else:
            messagebox.showinfo("완료", f"[{d_type}] 총 {success}건 프린트 완료.")

    def _safe_remove_paths(self, *paths: str) -> None:
        for path in paths:
            if not path: continue
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
                elif os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass

    def _schedule_temp_cleanup(self, paths: List[str], delay_ms: int = 120000) -> None:
        if not paths or not hasattr(self, "root") or not self.root.winfo_exists(): return
        self.root.after(delay_ms, lambda cleanup_paths=list(paths): self._safe_remove_paths(*cleanup_paths))

    def _find_best_korean_font_path(self) -> Optional[str]:
        if hasattr(self, "_dispatch_preview_font_path"): return self._dispatch_preview_font_path
        candidates = []
        if platform.system() == "Windows":
            win_dir = os.environ.get("WINDIR", r"C:\Windows")
            candidates.extend(
                [os.path.join(win_dir, "Fonts", "malgun.ttf"), os.path.join(win_dir, "Fonts", "malgunbd.ttf"),
                 os.path.join(win_dir, "Fonts", "gulim.ttc")])
        else:
            for font_name in ["Noto Sans CJK KR", "Noto Sans CJK", "NanumGothic", "Malgun Gothic", "DejaVu Sans"]:
                try:
                    proc = subprocess.run(["fc-match", "-f", "%{file}\n", font_name], capture_output=True, text=True,
                                          timeout=3)
                    font_path = proc.stdout.strip().splitlines()[0].strip() if proc.stdout.strip() else ""
                    if font_path and os.path.exists(font_path): candidates.append(font_path)
                except Exception:
                    continue
        font_path = next((path for path in candidates if path and os.path.exists(path)), None)
        self._dispatch_preview_font_path = font_path
        return font_path

    def _get_preview_font(self, size: int = 12, bold: bool = False):
        if not PILLOW_AVAILABLE: return None
        font_path = self._find_best_korean_font_path()
        font_size = max(10, int(size))
        if font_path:
            try:
                return ImageFont.truetype(font_path, font_size)
            except Exception:
                pass
        try:
            return ImageFont.load_default()
        except Exception:
            return None

    def _resolve_color_hex(self, color_obj, default: str = "#FFFFFF") -> str:
        rgb = getattr(color_obj, "rgb", None)
        if rgb:
            rgb = str(rgb).replace("#", "")
            if len(rgb) == 8: rgb = rgb[2:]
            if len(rgb) == 6: return f"#{rgb}"
        return default

    def _column_width_to_pixels(self, ws, col_idx: int) -> int:
        width = ws.column_dimensions[get_column_letter(col_idx)].width
        return max(24, int((width or 8.43) * 7 + 8))

    def _row_height_to_pixels(self, ws, row_idx: int) -> int:
        height = ws.row_dimensions[row_idx].height
        return max(20, int(float(height or 20) * 96 / 72))

    def _determine_dispatch_render_range(self, ws) -> Tuple[int, int, int, int]:
        if ws.print_area:
            try:
                area = str(ws.print_area).split(",")[0].split("!")[-1].replace("$", "")
                min_col, min_row, max_col, max_row = range_boundaries(area)
                return min_row, max_row, min_col, max_col
            except Exception:
                pass
        return 1, min(ws.max_row or 60, 80), 1, min(ws.max_column or 12, 30)

    def _build_merged_cell_maps(self, ws, min_row: int, max_row: int, min_col: int, max_col: int):
        merged_lookup, merged_bounds = {}, {}
        for merged_range in ws.merged_cells.ranges:
            start_col, start_row, end_col, end_row = merged_range.bounds
            if end_row < min_row or start_row > max_row or end_col < min_col or start_col > max_col: continue
            top_left = (start_row, start_col)
            merged_bounds[top_left] = (end_row, end_col)
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    merged_lookup[(r, c)] = top_left
        return merged_lookup, merged_bounds

    def _render_dispatch_sheet_image(self, workbook_path: str, camp_name: str, scale: float = 1.4):
        if not (OPENPYXL_AVAILABLE and PILLOW_AVAILABLE): raise RuntimeError("openpyxl 과 Pillow 라이브러리가 필요합니다.")
        wb = load_workbook(workbook_path)
        try:
            matched_sheet = self._find_best_match_key(camp_name, list(wb.sheetnames))
            ws = wb[matched_sheet] if matched_sheet else wb.active
            min_row, max_row, min_col, max_col = self._determine_dispatch_render_range(ws)
            merged_lookup, merged_bounds = self._build_merged_cell_maps(ws, min_row, max_row, min_col, max_col)

            col_widths = {col: max(24, int(self._column_width_to_pixels(ws, col) * scale)) for col in
                          range(min_col, max_col + 1)}
            row_heights = {row: max(20, int(self._row_height_to_pixels(ws, row) * scale)) for row in
                           range(min_row, max_row + 1)}

            x_positions, current_x = {}, 20
            for col in range(min_col, max_col + 1): x_positions[col] = current_x; current_x += col_widths[col]
            y_positions, current_y = {}, 20
            for row in range(min_row, max_row + 1): y_positions[row] = current_y; current_y += row_heights[row]

            image = Image.new("RGB", (max(current_x + 20, 800), max(current_y + 20, 500)), "white")
            draw = ImageDraw.Draw(image)
            default_font = self._get_preview_font(max(11, int(12 * scale)))
            bold_font = self._get_preview_font(max(11, int(13 * scale)), bold=True)

            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    if merged_lookup.get((row, col), (row, col)) != (row, col): continue

                    cell = ws.cell(row=row, column=col)
                    end_row, end_col = merged_bounds.get((row, col), (row, col))
                    x1, y1 = x_positions[col], y_positions[row]
                    x2, y2 = x_positions[end_col] + col_widths[end_col], y_positions[end_row] + row_heights[end_row]

                    fill_color = self._resolve_color_hex(getattr(cell.fill, "fgColor", None), "#FFFFFF")
                    draw.rectangle([x1, y1, x2, y2],
                                   fill=fill_color if getattr(cell.fill, "fill_type", None) == "solid" else "#FFFFFF",
                                   outline="#CFCFCF")

                    border = cell.border
                    if getattr(border.left, "style", None): draw.line([(x1, y1), (x1, y2)],
                                                                      fill=self._resolve_color_hex(border.left.color,
                                                                                                   "#666666"), width=2)
                    if getattr(border.top, "style", None): draw.line([(x1, y1), (x2, y1)],
                                                                     fill=self._resolve_color_hex(border.top.color,
                                                                                                  "#666666"), width=2)
                    if getattr(border.right, "style", None): draw.line([(x2, y1), (x2, y2)],
                                                                       fill=self._resolve_color_hex(border.right.color,
                                                                                                    "#666666"), width=2)
                    if getattr(border.bottom, "style", None): draw.line([(x1, y2), (x2, y2)],
                                                                        fill=self._resolve_color_hex(
                                                                            border.bottom.color, "#666666"), width=2)

                    text = "" if cell.value is None else str(cell.value)
                    if not text: continue

                    font_size = max(11, int(((cell.font.sz or 10) * scale)))
                    font_obj = self._get_preview_font(font_size, bold=bool(getattr(cell.font, "bold", False)))
                    if font_obj is None: font_obj = bold_font if getattr(cell.font, "bold", False) else default_font

                    text_color = self._resolve_color_hex(getattr(cell.font, "color", None), "#000000")
                    try:
                        bbox = draw.multiline_textbbox((0, 0), text, font=font_obj, spacing=4, align="left")
                        text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                    except Exception:
                        text_w, text_h = int(len(text) * font_size * 0.6), int(font_size * 1.4)

                    align = cell.alignment
                    horizontal, vertical = (align.horizontal or "center").lower(), (align.vertical or "center").lower()

                    if horizontal == "left":
                        tx = x1 + 8
                    elif horizontal == "right":
                        tx = max(x1 + 8, x2 - text_w - 8)
                    else:
                        tx = max(x1 + 8, x1 + (x2 - x1 - text_w) / 2)

                    if vertical == "top":
                        ty = y1 + 6
                    elif vertical == "bottom":
                        ty = max(y1 + 4, y2 - text_h - 6)
                    else:
                        ty = max(y1 + 4, y1 + (y2 - y1 - text_h) / 2)

                    draw.multiline_text((tx, ty), text, fill=text_color, font=font_obj, spacing=4, align="left")

            for img_obj in getattr(ws, "_images", []):
                try:
                    anchor = getattr(img_obj, "anchor", None)
                    if anchor is None or not hasattr(anchor, "_from"): continue
                    start_col, start_row = anchor._from.col + 1, anchor._from.row + 1
                    if start_col < min_col or start_col > max_col or start_row < min_row or start_row > max_row: continue
                    x = x_positions.get(start_col, 20) + int(getattr(anchor._from, "colOff", 0) / 9525 * scale)
                    y = y_positions.get(start_row, 20) + int(getattr(anchor._from, "rowOff", 0) / 9525 * scale)
                    img_bytes = img_obj._data() if callable(getattr(img_obj, "_data", None)) else None
                    if not img_bytes: continue
                    pasted = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                    target_w, target_h = max(1, int(getattr(img_obj, "width", pasted.width) * scale)), max(1,
                                                                                                           int(getattr(
                                                                                                               img_obj,
                                                                                                               "height",
                                                                                                               pasted.height) * scale))
                    pasted = pasted.resize((target_w, target_h), Image.Resampling.LANCZOS)
                    image.paste(pasted, (x, y), pasted)
                except Exception:
                    continue

            return image
        finally:
            wb.close()

    def _open_file_with_default_app(self, path: str) -> None:
        if platform.system() == "Windows" and hasattr(os, "startfile"):
            os.startfile(path);
            return
        if platform.system() == "Darwin":
            subprocess.Popen(["open", path]);
            return
        if opener := shutil.which("xdg-open"):
            subprocess.Popen([opener, path]);
            return
        raise RuntimeError("파일을 열 수 있는 기본 프로그램을 찾지 못했습니다.")

    def _show_dispatch_pdf_preview(self, pdf_path: str, cleanup_paths: Optional[List[str]] = None) -> None:
        cleanup_paths = cleanup_paths or []
        if not (PYMUPDF_AVAILABLE and PILLOW_AVAILABLE):
            self._open_file_with_default_app(pdf_path)
            self._schedule_temp_cleanup(cleanup_paths, delay_ms=300000)
            return

        C = self.C
        doc = fitz.open(pdf_path)
        preview_win = tk.Toplevel(
            self.dispatch_win if hasattr(self, "dispatch_win") and self.dispatch_win.winfo_exists() else self.root)
        preview_win.title("출차서 미리보기")
        preview_win.geometry("1100x850")
        preview_win.configure(bg=C["bg"])

        top_bar = tk.Frame(preview_win, bg=C["header"], height=42)
        top_bar.pack(fill="x");
        top_bar.pack_propagate(False)
        tk.Label(top_bar, text=f"📄  PDF 미리보기  ({doc.page_count} page)", bg=C["header"], fg="white",
                 font=("맑은 고딕", 10, "bold")).pack(side="left", padx=16, pady=10)
        self._btn(top_bar, "외부 뷰어로 열기", lambda: self._open_file_with_default_app(pdf_path), C["btn_secondary"]).pack(
            side="right", padx=10, pady=8)

        canvas = tk.Canvas(preview_win, bg="#E8EAED", highlightthickness=0)
        vbar = ttk.Scrollbar(preview_win, orient="vertical", command=canvas.yview)
        hbar = ttk.Scrollbar(preview_win, orient="horizontal", command=canvas.xview)
        scroll_frame = tk.Frame(canvas, bg="#E8EAED")
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=max(e.width - 4, 100)))
        canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")

        photo_refs = []
        try:
            for page_index in range(doc.page_count):
                pix = doc.load_page(page_index).get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
                photo = ImageTk.PhotoImage(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
                photo_refs.append(photo)
                wrapper = tk.Frame(scroll_frame, bg="#E8EAED")
                wrapper.pack(pady=12)
                tk.Label(wrapper, text=f"Page {page_index + 1}", bg="#E8EAED", fg=C["label"],
                         font=("맑은 고딕", 10, "bold")).pack(pady=(0, 6))
                tk.Label(wrapper, image=photo, bd=1, relief="solid").pack()
        finally:
            doc.close()

        preview_win._dispatch_preview_refs = photo_refs

        def _close_preview():
            self._safe_remove_paths(*cleanup_paths)
            preview_win.destroy()

        preview_win.protocol("WM_DELETE_WINDOW", _close_preview)

    def _row_vars_to_payload(self, row_vars: Dict[str, Any]) -> Dict[str, str]:
        result = {}
        for key in DISPATCH_HEADERS:
            value_obj = row_vars.get(key, "")
            result[key] = value_obj.get().strip() if hasattr(value_obj, "get") else str(value_obj).strip()
        return result

    # ------------------------------------------
    # 프리셋 설정 파일 로드 및 저장
    # ------------------------------------------
    def load_settings(self) -> Dict[str, Any]:
        default_data = deep_copy_default_settings()
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                settings_changed = False
                for key in ["app_title", "fs_title", "colors", "schedules", "camps", "docks", "hochas",
                            "display_timing", "dock_presets", "dispatch_manual_order", "ui_mode"]:
                    if key not in data:
                        data[key] = default_data[key]
                        settings_changed = True
                if data.get("app_title") == "INC4 HUB 차량 정보 시스템":
                    data["app_title"] = "HUB 차량 정보 시스템"
                    settings_changed = True
                if "ui_mode" not in data or data.get("ui_mode") not in PALETTE_PRESETS:
                    try:
                        with open("ui_palette_state.json", "r", encoding="utf-8") as f:
                            legacy_ui_state = json.load(f)
                        legacy_mode = legacy_ui_state.get("active_preset", "")
                        if legacy_mode in PALETTE_PRESETS:
                            data["ui_mode"] = legacy_mode
                            settings_changed = True
                    except Exception:
                        data["ui_mode"] = default_data.get("ui_mode", LEGACY_PRESET_NAME)
                        settings_changed = True
                for key, value in default_data["colors"].items():
                    if key not in data.setdefault("colors", {}):
                        data["colors"][key] = value
                        settings_changed = True
                for key, value in default_data["display_timing"].items():
                    if key not in data.setdefault("display_timing", {}):
                        data["display_timing"][key] = value
                        settings_changed = True
                visibility = data.setdefault("schedules", {}).setdefault("visibility", {})
                for key, value in default_data["schedules"]["visibility"].items():
                    if key not in visibility:
                        visibility[key] = value
                        settings_changed = True
                if "dispatch_manual_order" not in data:
                    data["dispatch_manual_order"] = {}
                    settings_changed = True
                if settings_changed:
                    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=4)
                return data
            except Exception as e:
                backup_path = f"{SETTINGS_FILE}.broken_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                try:
                    shutil.copy2(SETTINGS_FILE, backup_path)
                    print(f"[설정 복구] 손상된 설정 파일을 '{backup_path}'로 백업했습니다: {e}")
                except Exception as backup_error:
                    print(f"[설정 복구] 설정 백업에 실패했습니다: {backup_error}")
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(default_data, f, ensure_ascii=False, indent=4)
        return default_data

    def save_settings(self, data: Dict[str, Any]) -> None:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)
        self.settings = data

    # ------------------------------------------
    # 설정 (Settings / Presets) 탭 구성 관리
    # ------------------------------------------
    def open_presets(self) -> None:
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("설정")
        win.geometry("900x640")
        win.configure(bg=C["bg"])
        win.transient(self.root)

        hdr = tk.Frame(win, bg=C["header"], height=40)
        hdr.pack(fill="x");
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚙  시스템 설정", bg=C["header"], fg="white", font=("맑은 고딕", 10, "bold")).pack(side="left",
                                                                                                     padx=16, pady=10)

        style = ttk.Style()
        style.configure("Cfg.TNotebook", background=C["bg"], borderwidth=0)
        style.configure("Cfg.TNotebook.Tab", background=C["bg"], foreground=C["label"], padding=[18, 9],
                        font=("맑은 고딕", 10, "bold"), borderwidth=0)
        style.map("Cfg.TNotebook.Tab", background=[("selected", C["white"])],
                  foreground=[("selected", C["btn_primary"])])

        nb = ttk.Notebook(win, style="Cfg.TNotebook")
        nb.pack(fill="both", expand=True)

        def mktab(text):
            f = tk.Frame(nb, bg=C["bg"]);
            nb.add(f, text=text)
            return f

        t1, t2, t3, t4, t5 = mktab("캠프"), mktab("도크"), mktab("호차"), mktab("시간 설정"), mktab("기타")
        self._build_camp_tab(t1)
        self._build_dock_preset_tab(t2)
        self._build_simple_list_tab(t3, "호차", self.hocha_list.copy(), lambda i: self._save_preset("hochas", i))
        self._build_time_tab(t4)
        self._build_others_tab(t5)

        tk.Frame(win, bg=C["border"], height=1).pack(fill="x")
        footer = tk.Frame(win, bg=C["bg"], pady=8)
        footer.pack(fill="x")
        self._btn(footer, "닫기", win.destroy, C["btn_neutral"], width=10).pack(side="right", padx=14)

    # ------------------------------------------
    # 도크 프리셋 탭 (Delivery Type별 도크-캠프 매핑 + 도크번호 관리)
    # ------------------------------------------
    def _build_dock_preset_tab(self, parent: tk.Frame) -> None:
        C = self.C
        parent.configure(bg=C["bg"])

        DOCK_TYPES = ["SD2", "SD3", "D1", "D2", "D3", "D4", "간선", "W2"]

        left = tk.Frame(parent, bg=C["white"], width=160, highlightthickness=1, highlightbackground=C["border"])
        left.pack(side="left", fill="y", padx=(12, 0), pady=12)
        left.pack_propagate(False)

        right_outer = tk.Frame(parent, bg=C["bg"])
        right_outer.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        tk.Label(left, text="Delivery Type", bg=C["white"], fg=C["btn_primary"],
                 font=("맑은 고딕", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
        type_lb = tk.Listbox(left, height=14, bg=C["white"], fg=C["text"],
                             selectbackground=C["row_sel"], selectforeground=C["btn_primary"],
                             borderwidth=0, highlightthickness=1, highlightbackground=C["border"],
                             relief="flat", font=("맑은 고딕", 10), activestyle="none")
        type_lb.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for t in DOCK_TYPES:
            type_lb.insert(tk.END, t)

        # 우측: 스크롤 가능한 도크-캠프 매핑 영역
        right_card = tk.Frame(right_outer, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        right_card.pack(fill="both", expand=True)

        # ── 헤더 (3컬럼: 도크 번호 | 캠프명 | 삭제)
        header_row = tk.Frame(right_card, bg=C["heading"])
        header_row.pack(fill="x")
        tk.Label(header_row, text="도크 번호", bg=C["heading"], fg=C["label"],
                 font=("맑은 고딕", 9, "bold"), width=12, anchor="center", pady=8).pack(side="left", padx=(8, 4))
        tk.Label(header_row, text="캠프명 (목록에서 선택 또는 직접 입력)", bg=C["heading"], fg=C["label"],
                 font=("맑은 고딕", 9, "bold"), anchor="w", pady=8).pack(side="left", padx=8, fill="x", expand=True)
        tk.Label(header_row, text="삭제", bg=C["heading"], fg=C["label"],
                 font=("맑은 고딕", 9, "bold"), width=5, anchor="center", pady=8).pack(side="right", padx=8)

        scroll_canvas = tk.Canvas(right_card, highlightthickness=0, bg=C["white"])
        vsb = ttk.Scrollbar(right_card, orient="vertical", command=scroll_canvas.yview)
        scroll_frame = tk.Frame(scroll_canvas, bg=C["white"])
        scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_frame.bind("<Configure>", lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
        scroll_canvas.configure(yscrollcommand=vsb.set)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 행 리스트: List[{"dock_var": StringVar, "camp_var": StringVar, "frame": Frame}]
        # d_type별로 관리
        self._dock_preset_rows: Dict[str, List[Dict]] = {}
        current_type_var = tk.StringVar(value="")

        def _refresh_scroll():
            scroll_canvas.update_idletasks()
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

        def _make_row(d_type: str, dock_val: str, camp_val: str, bg: str) -> Dict:
            """도크번호/캠프명 한 행 위젯을 생성하고 dict 반환"""
            camps_for_type = [""] + self.camp_data.get(d_type, [])

            row_f = tk.Frame(scroll_frame, bg=bg)
            row_f.pack(fill="x")

            dock_var = tk.StringVar(value=dock_val)
            camp_var = tk.StringVar(value=camp_val)

            # 도크 번호 – Combobox (드롭다운 + 직접 입력 가능)
            dock_cb = ttk.Combobox(row_f, textvariable=dock_var,
                                   values=self.dock_list, state="normal",
                                   width=10, font=("맑은 고딕", 10))
            dock_cb.pack(side="left", padx=(12, 4), pady=5)

            # 캠프명 – Combobox
            camp_cb = ttk.Combobox(row_f, textvariable=camp_var,
                                   values=camps_for_type, state="normal",
                                   width=20, font=("맑은 고딕", 10))
            camp_cb.pack(side="left", padx=4, pady=5)

            row_data: Dict = {"dock_var": dock_var, "camp_var": camp_var, "frame": row_f}

            # ❌ 삭제 버튼
            def _delete_row(rd=row_data, dt=d_type):
                rd["frame"].destroy()
                if dt in self._dock_preset_rows and rd in self._dock_preset_rows[dt]:
                    self._dock_preset_rows[dt].remove(rd)
                _recolor_rows(dt)
                _refresh_scroll()

            del_btn = tk.Button(row_f, text="✕", command=_delete_row,
                                bg=C["btn_danger"], fg="white",
                                activebackground=C["btn_danger"], activeforeground="white",
                                relief="flat", cursor="hand2", font=("맑은 고딕", 9, "bold"),
                                width=3, bd=0, highlightthickness=0)
            del_btn.pack(side="right", padx=(4, 10), pady=4)

            return row_data

        def _recolor_rows(d_type: str):
            """삭제 후 행 배경색을 홀/짝 교대로 재적용"""
            rows = self._dock_preset_rows.get(d_type, [])
            for idx, rd in enumerate(rows):
                bg = C["white"] if idx % 2 == 0 else C["row_even"]
                try:
                    rd["frame"].configure(bg=bg)
                    for child in rd["frame"].winfo_children():
                        if isinstance(child, tk.Button):
                            child.configure(bg=C["btn_danger"])
                except tk.TclError:
                    pass

        def _render_dock_rows(d_type: str):
            """선택된 Delivery Type의 도크-캠프 매핑 행을 렌더링"""
            for w in scroll_frame.winfo_children():
                w.destroy()

            saved_mapping = self.dock_presets.get(d_type, {})

            ordered_docks = list(self.dock_list)
            for dock_num in saved_mapping:
                if dock_num not in ordered_docks:
                    ordered_docks.append(dock_num)
            source_items = [(dock, saved_mapping.get(dock, "")) for dock in ordered_docks]

            rows: List[Dict] = []
            for idx, (dock_num, camp_val) in enumerate(source_items):
                bg = C["white"] if idx % 2 == 0 else C["row_even"]
                rd = _make_row(d_type, dock_num, camp_val, bg)
                rows.append(rd)

            self._dock_preset_rows[d_type] = rows
            _refresh_scroll()

        def _on_type_select(e=None):
            if sel := type_lb.curselection():
                d_type = type_lb.get(sel[0])
                current_type_var.set(d_type)
                _render_dock_rows(d_type)

        type_lb.bind("<<ListboxSelect>>", _on_type_select)

        # ── 도크 번호 추가 영역
        add_frame = tk.Frame(right_outer, bg=C["bg"], pady=4)
        add_frame.pack(fill="x")
        tk.Label(add_frame, text="새 도크번호:", bg=C["bg"], fg=C["label"],
                 font=("맑은 고딕", 9)).pack(side="left")
        new_dock_entry = self._entry(add_frame, width=10)
        new_dock_entry.pack(side="left", padx=(4, 6))

        def _add_dock_row():
            d_type = current_type_var.get()
            if not d_type:
                return
            new_dock = new_dock_entry.get().strip()
            if not new_dock:
                messagebox.showwarning("입력 오류", "도크 번호를 입력해주세요.")
                return

            # dock_list에 없으면 전역 목록에도 추가
            if new_dock not in self.dock_list:
                self.dock_list.append(new_dock)
                # 현재 열린 모든 도크 Combobox의 드롭다운 값 갱신
                for dt_rows in self._dock_preset_rows.values():
                    for rd in dt_rows:
                        try:
                            for child in rd["frame"].winfo_children():
                                if isinstance(child, ttk.Combobox) and child.cget("width") == 10:
                                    child["values"] = self.dock_list
                        except tk.TclError:
                            pass

            rows = self._dock_preset_rows.setdefault(d_type, [])
            idx = len(rows)
            bg = C["white"] if idx % 2 == 0 else C["row_even"]
            rd = _make_row(d_type, new_dock, "", bg)
            rows.append(rd)
            new_dock_entry.delete(0, tk.END)
            _refresh_scroll()
            # 스크롤을 맨 아래로
            scroll_canvas.after(50, lambda: scroll_canvas.yview_moveto(1.0))

        self._btn(add_frame, "추가", _add_dock_row, C["btn_primary"]).pack(side="left")

        # ── 저장 버튼 영역
        btn_frame = tk.Frame(right_outer, bg=C["bg"])
        btn_frame.pack(fill="x", pady=(4, 0))
        tk.Label(btn_frame, text="※ 저장하면 기존 데이터의 도크가 즉시 자동 배정됩니다.",
                 bg=C["bg"], fg=C["placeholder"], font=("맑은 고딕", 8)).pack(side="left")
        self._btn(btn_frame, "저장",
                  lambda: self._save_dock_presets(current_type_var.get()),
                  C["btn_success"], width=10).pack(side="right")

        # 초기 선택
        type_lb.selection_set(0)
        _on_type_select()

    def _save_dock_presets(self, current_type: str = "") -> None:
        """현재 UI의 도크-캠프 매핑을 dock_presets에 저장 후 자동 배정 실행"""
        if not hasattr(self, "_dock_preset_rows"):
            return

        # 새로 추가된 도크 번호도 포함하여 dock_list 갱신 후 저장
        new_dock_nums: List[str] = []
        for d_type, rows in self._dock_preset_rows.items():
            type_mapping: Dict[str, str] = {}
            for rd in rows:
                dock_num = rd["dock_var"].get().strip()
                camp_val = rd["camp_var"].get().strip()
                if dock_num:
                    if camp_val:
                        type_mapping[dock_num] = camp_val
                    if dock_num not in self.dock_list:
                        new_dock_nums.append(dock_num)
            self.dock_presets[d_type] = type_mapping

        # dock_list 업데이트
        for d in new_dock_nums:
            if d not in self.dock_list:
                self.dock_list.append(d)
        self.settings["docks"] = self.dock_list
        self.settings["dock_presets"] = self.dock_presets
        self.save_settings(self.settings)

        # 메인 도크 콤보박스 갱신
        if hasattr(self, "dock_combo"):
            self.dock_combo["values"] = self.dock_list

        # 저장 후 즉시 기존 데이터에 자동 배정 적용
        assigned_count = self._auto_assign_dock_from_preset()
        msg = "도크 프리셋이 저장되었습니다."
        if assigned_count > 0:
            msg += f"\n{assigned_count}개 차량에 도크가 자동 배정되었습니다."
        messagebox.showinfo("저장 완료", msg)

    def _apply_dock_preset_btn(self) -> None:
        """툴바 '도크 배정' 버튼: 현재 데이터에 프리셋을 즉시 재적용"""
        if not self.dock_presets:
            messagebox.showinfo("안내", "도크 프리셋이 설정되어 있지 않습니다.\n설정 → 도크 탭에서 먼저 설정해 주세요.")
            return
        count = self._auto_assign_dock_from_preset()
        if count > 0:
            messagebox.showinfo("도크 배정 완료", f"{count}개 차량에 도크가 배정되었습니다.")
        else:
            messagebox.showinfo("도크 배정", "새로 배정된 도크가 없습니다.\n(이미 배정되었거나 매칭되는 항목이 없습니다.)")

    def _auto_assign_dock_from_preset(self) -> int:
        """
        현재 tree의 각 항목에 대해 (d_type, camp) 기준으로 dock_presets를 역방향 조회하여
        프리셋과 다른 도크가 있으면 최신 매핑으로 갱신한다.
        반환값: 배정/갱신된 항목 수
        """
        if not self.dock_presets:
            return 0

        assigned_count = 0
        # tree 아이템
        for iid in self.tree.get_children():
            v = list(self.tree.item(iid, "values"))
            if not v or len(v) < 4:
                continue
            d_type = str(v[1]).strip()
            camp = str(v[2]).strip()
            dock = self._lookup_dock_for_camp(d_type, camp)
            if dock and str(v[3]).strip() != dock:
                v[3] = dock
                self.tree.item(iid, values=tuple(v))
                assigned_count += 1

        # detached 아이템
        for iid, info in self.detached_items.items():
            v = list(info["values"])
            if not v or len(v) < 4:
                continue
            d_type = str(v[1]).strip()
            camp = str(v[2]).strip()
            dock = self._lookup_dock_for_camp(d_type, camp)
            if dock and str(v[3]).strip() != dock:
                v[3] = dock
                self.detached_items[iid]["values"] = tuple(v)
                assigned_count += 1

        if assigned_count > 0:
            self._sync_full_screen()

        return assigned_count

    def _lookup_dock_for_camp(self, d_type: str, camp: str) -> str:
        """
        dock_presets[d_type]에서 camp와 매핑된 도크 번호를 반환.
        없으면 빈 문자열 반환.
        매칭 우선순위:
          1) 정확 매칭 (공백·대소문자 무시)
          2) 괄호 이하 제거 후 매칭 ("김포1(SH)" → "김포1")
          3) 어느 쪽이 다른 쪽을 포함하는 부분 매칭
        """
        type_mapping = self.dock_presets.get(d_type, {})
        if not type_mapping:
            return ""

        import re as _re

        def _norm(s: str) -> str:
            return str(s).replace(" ", "").upper()

        def _strip_paren(s: str) -> str:
            """괄호(소/중/대) 및 그 내용 제거"""
            return _re.sub(r"[\(\（\[（【].*?[\)\）\]）】]", "", str(s)).replace(" ", "").upper()

        camp_norm    = _norm(camp)
        camp_stripped = _strip_paren(camp)

        # 1순위: 정확 매칭
        for dock_num, mapped_camp in type_mapping.items():
            if _norm(mapped_camp) == camp_norm:
                return dock_num

        # 2순위: 괄호 제거 후 매칭
        for dock_num, mapped_camp in type_mapping.items():
            if _strip_paren(mapped_camp) == camp_stripped:
                return dock_num

        # 3순위: 어느 쪽이 다른 쪽을 포함 (prefix/contains)
        for dock_num, mapped_camp in type_mapping.items():
            m = _norm(mapped_camp)
            if m and (camp_norm.startswith(m) or m.startswith(camp_norm)):
                return dock_num

        return ""

    def _build_camp_tab(self, parent: tk.Frame) -> None:
        C = self.C
        parent.configure(bg=C["bg"])
        left = tk.Frame(parent, bg=C["white"], width=190, highlightthickness=1, highlightbackground=C["border"])
        left.pack(side="left", fill="y", padx=(12, 0), pady=12);
        left.pack_propagate(False)
        right = tk.Frame(parent, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        right.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        tk.Label(left, text="Delivery Type", bg=C["white"], fg=C["btn_primary"], font=("맑은 고딕", 9, "bold")).pack(
            anchor="w", padx=12, pady=(10, 4))
        self._type_lb = tk.Listbox(left, height=14, bg=C["white"], fg=C["text"], selectbackground=C["row_sel"],
                                   selectforeground=C["btn_primary"], borderwidth=0, highlightthickness=1,
                                   highlightbackground=C["border"], relief="flat", font=("맑은 고딕", 10),
                                   activestyle="none")
        self._type_lb.pack(fill="both", expand=True, padx=10)
        for t in self.camp_data.keys(): self._type_lb.insert(tk.END, t)

        add_row = tk.Frame(left, bg=C["white"])
        add_row.pack(fill="x", padx=10, pady=6)
        self._new_type_entry = self._entry(add_row, width=9)
        self._new_type_entry.pack(side="left", padx=(0, 4))
        self._btn(add_row, "추가", self._add_type, C["btn_primary"]).pack(side="left")
        self._btn(left, "선택 삭제", self._delete_type, C["btn_danger"]).pack(fill="x", padx=10, pady=(0, 10))

        hdr_row = tk.Frame(right, bg=C["white"])
        hdr_row.pack(fill="x", padx=12, pady=(10, 6))
        self._current_type_var = tk.StringVar(value="")
        tk.Label(hdr_row, text="선택된 Type:", bg=C["white"], fg=C["label"], font=("맑은 고딕", 9)).pack(side="left")
        tk.Label(hdr_row, textvariable=self._current_type_var, bg=C["white"], fg=C["btn_primary"],
                 font=("맑은 고딕", 10, "bold")).pack(side="left", padx=6)

        self._camp_lb = tk.Listbox(right, bg=C["white"], fg=C["text"], selectbackground=C["row_sel"],
                                   selectforeground=C["btn_primary"], borderwidth=0, highlightthickness=1,
                                   highlightbackground=C["border"], relief="flat", font=("맑은 고딕", 10),
                                   activestyle="none")
        self._camp_lb.pack(fill="both", expand=True, padx=10)

        ctrl_row = tk.Frame(right, bg=C["white"])
        ctrl_row.pack(fill="x", padx=10, pady=6)
        self._new_camp_entry = self._entry(ctrl_row)
        self._new_camp_entry.pack(side="left", padx=(0, 6))
        self._btn(ctrl_row, "추가", self._add_camp, C["btn_primary"]).pack(side="left", padx=(0, 4))
        self._btn(ctrl_row, "삭제", self._delete_camp, C["btn_danger"]).pack(side="left")
        self._btn(right, "저장", self._save_camps, C["btn_success"], width=10).pack(side="right", padx=10, pady=(0, 10))

        def _on_type_select(e):
            if sel := self._type_lb.curselection():
                t = self._type_lb.get(sel[0])
                self._current_type_var.set(t)
                self._camp_lb.delete(0, tk.END)
                for c in self.camp_data.get(t, []): self._camp_lb.insert(tk.END, c)

        self._type_lb.bind("<<ListboxSelect>>", _on_type_select)
        if self._type_lb.size(): self._type_lb.selection_set(0); _on_type_select(None)

    def _add_type(self):
        t = self._new_type_entry.get().strip()
        if t and t not in self.camp_data:
            self.camp_data[t] = []
            self._type_lb.insert(tk.END, t)
            self._new_type_entry.delete(0, tk.END)
            self.type_combo["values"] = list(self.camp_data.keys())

    def _delete_type(self):
        if sel := self._type_lb.curselection():
            t = self._type_lb.get(sel[0])
            self.camp_data.pop(t, None)
            self._type_lb.delete(sel[0])
            self._camp_lb.delete(0, tk.END)
            self._current_type_var.set("")
            self.type_combo["values"] = list(self.camp_data.keys())

    def _add_camp(self):
        t, c = self._current_type_var.get().strip(), self._new_camp_entry.get().strip()
        if t and c and c not in self.camp_data.get(t, []):
            self.camp_data[t].append(c)
            self._camp_lb.insert(tk.END, c)
            self._new_camp_entry.delete(0, tk.END)

    def _delete_camp(self):
        t, sel = self._current_type_var.get().strip(), self._camp_lb.curselection()
        if t and sel:
            for i in reversed(sel):
                val = self._camp_lb.get(i)
                self._camp_lb.delete(i)
                try:
                    self.camp_data[t].remove(val)
                except ValueError:
                    pass

    def _save_camps(self):
        if t := self._current_type_var.get().strip():
            self.camp_data[t] = [self._camp_lb.get(i) for i in range(self._camp_lb.size())]
        self.settings["camps"] = self.camp_data
        self.save_settings(self.settings)
        messagebox.showinfo("저장 완료", "캠프 프리셋이 저장되었습니다.")

    def _build_simple_list_tab(self, parent, title, items, on_save):
        C = self.C
        parent.configure(bg=C["bg"])
        tk.Label(parent, text=f"{title} 프리셋", bg=C["bg"], fg=C["btn_primary"], font=("맑은 고딕", 10, "bold")).pack(
            anchor="w", padx=14, pady=(12, 6))

        card = tk.Frame(parent, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        card.pack(fill="both", expand=True, padx=14)

        lb = tk.Listbox(card, bg=C["white"], fg=C["text"], selectbackground=C["row_sel"],
                        selectforeground=C["btn_primary"], borderwidth=0, highlightthickness=0, relief="flat",
                        font=("맑은 고딕", 10), activestyle="none")
        lb.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        for it in items: lb.insert(tk.END, it)

        side = tk.Frame(card, bg=C["white"])
        side.pack(side="left", fill="y", padx=10, pady=10)
        ent = self._entry(side, width=14)
        ent.pack(fill="x", pady=(0, 6))
        self._btn(side, "추가",
                  lambda: (lb.insert(tk.END, ent.get().strip()), ent.delete(0, tk.END)) if ent.get().strip() else None,
                  C["btn_primary"]).pack(fill="x", pady=(0, 4))
        self._btn(side, "삭제", lambda: [lb.delete(i) for i in reversed(lb.curselection())], C["btn_danger"]).pack(
            fill="x")

        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", padx=14, pady=8)
        self._btn(parent, "저장", lambda: on_save([lb.get(i) for i in range(lb.size())]), C["btn_success"],
                  width=10).pack(side="right", padx=14, pady=(0, 10))

    def _save_preset(self, key, items):
        items = [s for s in items if str(s).strip()]
        self.settings[key] = items
        self.save_settings(self.settings)
        if key == "docks":
            self.dock_list = items
            self.dock_combo["values"] = items
        if key == "hochas":
            self.hocha_list = items
            self.hocha_combo["values"] = items
        messagebox.showinfo("저장 완료", f"{key} 설정이 저장되었습니다.")

    def _build_time_tab(self, parent: tk.Frame) -> None:
        C = self.C
        parent.configure(bg=C["bg"])
        outer = tk.Frame(parent, bg=C["bg"], padx=16, pady=14)
        outer.pack(fill="both", expand=True)

        card = tk.Frame(outer, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        card.pack(fill="x")

        hdr_row = tk.Frame(card, bg=C["heading"])
        hdr_row.pack(fill="x")
        for text, w in [("Type", 100), ("표시 시작", 140), ("", 30), ("표시 종료", 140)]:
            tk.Label(hdr_row, text=text, bg=C["heading"], fg=C["label"], font=("맑은 고딕", 9, "bold"), width=w // 8,
                     anchor="center", pady=8).pack(side="left", padx=8)

        self.time_vars = {}
        for idx, t_type in enumerate(["SD2", "SD3", "D1", "D2", "D3", "D4", "간선", "W2"]):
            sv = self.VISIBILITY_SCHEDULE.get(t_type, ["00:00", "00:00"])
            start_var, end_var = tk.StringVar(value=sv[0]), tk.StringVar(value=sv[1])
            self.time_vars[t_type] = (start_var, end_var)

            row = tk.Frame(card, bg=C["white"] if idx % 2 == 0 else C["row_even"])
            row.pack(fill="x")
            tk.Label(row, text=t_type, bg=row["bg"], fg=C["text"], font=("맑은 고딕", 10, "bold"), width=12,
                     anchor="center", pady=8).pack(side="left", padx=8)

            for var in [start_var, end_var]:
                e = tk.Entry(row, textvariable=var, width=10, justify="center", bg=C["white"], fg=C["text"],
                             relief="solid", bd=1, highlightthickness=1, highlightbackground=C["border"],
                             highlightcolor=C["border_focus"], font=("맑은 고딕", 10))
                e.pack(side="left", padx=8, pady=4)
                if var == start_var: tk.Label(row, text="~", bg=row["bg"], fg=C["placeholder"],
                                              font=("맑은 고딕", 10)).pack(side="left")

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=10)
        self._btn(outer, "저장", self._save_time_settings, C["btn_success"], width=10).pack(side="right")

    def _save_time_settings(self):
        for t_type, (s_var, e_var) in self.time_vars.items():
            s_val, e_val = s_var.get().strip(), e_var.get().strip()
            if not (TIME_PATTERN.match(s_val) and TIME_PATTERN.match(e_val)):
                messagebox.showerror("형식 오류", f"[{t_type}]의 시간 형식이 잘못되었습니다.\n(올바른 예시: 08:30, 22:40)")
                return
            self.VISIBILITY_SCHEDULE[t_type] = [s_val, e_val]

        self.settings["schedules"]["visibility"] = self.VISIBILITY_SCHEDULE
        self.save_settings(self.settings)
        self._check_visibility()
        self._sync_full_screen()
        messagebox.showinfo("저장 완료", "시간 설정이 성공적으로 저장되었습니다.")

    def _build_others_tab(self, parent: tk.Frame) -> None:
        C = self.C
        parent.configure(bg=C["bg"])
        outer = tk.Frame(parent, bg=C["bg"], padx=16, pady=14)
        outer.pack(fill="both", expand=True)

        self.app_title_var = tk.StringVar(value=self.app_title)
        self.fs_title_var = tk.StringVar(value=self.fs_title)
        self.color_bg_var = tk.StringVar(value=self.colors["bg"])
        self.color_odd_var = tk.StringVar(value=self.colors["odd_bg"])
        self.color_even_var = tk.StringVar(value=self.colors["even_bg"])
        self.color_font_var = tk.StringVar(value=self.colors["font"])
        self.color_clock_var = tk.StringVar(value=self.colors["clock"])
        self.flap_seconds_var = tk.StringVar(value=str(self.flap_seconds))
        self.page_flip_seconds_var = tk.StringVar(value=str(self.page_flip_seconds))
        self.use_split_flap_var = tk.BooleanVar(value=getattr(self, "use_split_flap", True))

        def sec(text):
            tk.Label(outer, text=text, bg=C["bg"], fg=C["btn_primary"], font=("맑은 고딕", 9, "bold")).pack(anchor="w",
                                                                                                        pady=(10, 4))

        def card_frame():
            f = tk.Frame(outer, bg=C["white"], highlightthickness=1, highlightbackground=C["border"], padx=16, pady=10)
            f.pack(fill="x");
            return f

        def txt_row(parent_f, label_text, var):
            row = tk.Frame(parent_f, bg=C["white"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label_text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=16, anchor="w").pack(
                side="left")
            tk.Entry(row, textvariable=var, width=36, bg=C["white"], fg=C["text"], relief="solid", bd=1,
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["border_focus"],
                     font=("맑은 고딕", 10)).pack(side="left")

        sec("▎ 문구 설정");
        tf = card_frame()
        txt_row(tf, "프로그램 타이틀", self.app_title_var)
        txt_row(tf, "전광판 타이틀", self.fs_title_var)

        sec("▎ 전광판 색상 설정");
        cf = card_frame()

        def pick_color(var_obj, btn_obj):
            code = colorchooser.askcolor(title="색상 선택", initialcolor=var_obj.get())[1]
            if code: var_obj.set(code); btn_obj.config(bg=code)

        color_defs = [("바탕화면 색상", self.color_bg_var, 0, 0), ("홀수열 배경", self.color_odd_var, 1, 0),
                      ("짝수열 배경", self.color_even_var, 2, 0), ("폰트 색상", self.color_font_var, 0, 3),
                      ("시계 색상", self.color_clock_var, 1, 3)]
        for ltext, var, r, c in color_defs:
            tk.Label(cf, text=ltext, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=14, anchor="w").grid(row=r,
                                                                                                                 column=c,
                                                                                                                 padx=(
                                                                                                                     14,
                                                                                                                     4),
                                                                                                                 pady=6)
            btn = tk.Button(cf, width=8, bg=var.get(), relief="solid", bd=1, cursor="hand2")
            btn.config(command=lambda v=var, b=btn: pick_color(v, b))
            btn.grid(row=r, column=c + 1, pady=6, padx=(0, 20))

        sec("▎ 전광판 타이밍");
        tmf = card_frame()
        for r, (lbl_text, var) in enumerate(
                [("스플릿플랩 작동 시간 (초)", self.flap_seconds_var), ("다음 페이지 대기 시간 (초)", self.page_flip_seconds_var)]):
            tk.Label(tmf, text=lbl_text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=24, anchor="w").grid(
                row=r, column=0, padx=4, pady=6)
            tk.Entry(tmf, textvariable=var, width=10, justify="center", bg=C["white"], fg=C["text"], relief="solid",
                     bd=1, highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["border_focus"],
                     font=("맑은 고딕", 10)).grid(row=r, column=1, padx=8, pady=6)
        tk.Checkbutton(tmf, text="스플릿 플랩 애니메이션 사용 (저사양 PC용 정지 가능)", variable=self.use_split_flap_var, bg=C["white"], fg=C["text"]).grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=4)
        tk.Label(tmf, text="기본값: 스플릿플랩 5초 / 페이지 전환 7초", bg=C["white"], fg=C["placeholder"], font=("맑은 고딕", 8)).grid(
            row=3, column=0, columnspan=3, sticky="w", padx=4, pady=(0, 4))

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=12)
        btn_row = tk.Frame(outer, bg=C["bg"])
        btn_row.pack(fill="x")
        self._btn(btn_row, "기본 설정으로", self._reset_others_settings, C["btn_warning"]).pack(side="left")
        self._btn(btn_row, "저장", self._save_others_settings, C["btn_success"], width=10).pack(side="right")

    def _reset_others_settings(self):
        if messagebox.askyesno("확인", "타이틀/색상/전광판 시간 설정을 기본값으로 되돌리시겠습니까?"):
            self.app_title_var.set(DEFAULT_SETTINGS["app_title"])
            self.fs_title_var.set(DEFAULT_SETTINGS["fs_title"])
            self.color_bg_var.set(DEFAULT_SETTINGS["colors"]["bg"])
            self.color_odd_var.set(DEFAULT_SETTINGS["colors"]["odd_bg"])
            self.color_even_var.set(DEFAULT_SETTINGS["colors"]["even_bg"])
            self.color_font_var.set(DEFAULT_SETTINGS["colors"]["font"])
            self.color_clock_var.set(DEFAULT_SETTINGS["colors"]["clock"])
            self.flap_seconds_var.set(str(DEFAULT_SETTINGS["display_timing"]["flap_seconds"]))
            self.page_flip_seconds_var.set(str(DEFAULT_SETTINGS["display_timing"]["page_flip_seconds"]))
            self.use_split_flap_var.set(DEFAULT_SETTINGS["display_timing"].get("use_split_flap", True))
            self._save_others_settings()
            messagebox.showinfo("안내", "기본 설정으로 복구되었습니다. 변경사항을 반영하기 위해 설정창을 닫았다 다시 열어주세요.")

    def _save_others_settings(self):
        flap_seconds = self._safe_positive_int(self.flap_seconds_var.get().strip(),
                                               DEFAULT_SETTINGS["display_timing"]["flap_seconds"])
        page_flip_seconds = self._safe_positive_int(self.page_flip_seconds_var.get().strip(),
                                                    DEFAULT_SETTINGS["display_timing"]["page_flip_seconds"])

        if flap_seconds <= 0 or page_flip_seconds <= 0:
            messagebox.showerror("입력 오류", "전광판 시간은 1초 이상의 숫자로 입력해주세요.")
            return

        self.app_title, self.fs_title = self.app_title_var.get().strip(), self.fs_title_var.get().strip()
        self.colors["bg"], self.colors["odd_bg"], self.colors[
            "even_bg"] = self.color_bg_var.get(), self.color_odd_var.get(), self.color_even_var.get()
        self.colors["font"], self.colors["clock"] = self.color_font_var.get(), self.color_clock_var.get()
        self.flap_seconds, self.page_flip_seconds = flap_seconds, page_flip_seconds
        self.use_split_flap = self.use_split_flap_var.get()

        self.settings["app_title"], self.settings["fs_title"] = self.app_title, self.fs_title
        self.settings["colors"] = self.colors
        self.settings["display_timing"] = {"flap_seconds": self.flap_seconds,
                                           "page_flip_seconds": self.page_flip_seconds,
                                           "use_split_flap": self.use_split_flap}

        self.save_settings(self.settings)
        self.root.title(self.app_title)

        if getattr(self, "flap_refresh_timer_id", None):
            self.root.after_cancel(self.flap_refresh_timer_id)
            self.flap_refresh_timer_id = None
        self._periodic_flap_refresh()
        self._schedule_fs_page_flip()
        self._sync_full_screen(force_rebuild=True)
        messagebox.showinfo("저장 완료", "기타 설정이 성공적으로 저장되었습니다.")

    def _build_others_tab(self, parent: tk.Frame) -> None:
        C = self.C
        parent.configure(bg=C["bg"])
        scroll_wrap = tk.Frame(parent, bg=C["bg"])
        scroll_wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(scroll_wrap, bg=C["bg"], highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(scroll_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        outer = tk.Frame(canvas, bg=C["bg"], padx=16, pady=14)
        outer_window = canvas.create_window((0, 0), window=outer, anchor="nw")

        def _sync_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _fit_outer_width(event):
            canvas.itemconfigure(outer_window, width=event.width)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")

        outer.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _fit_outer_width)
        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        self.app_title_var = tk.StringVar(value=self.app_title)
        self.fs_title_var = tk.StringVar(value=self.fs_title)
        self.color_bg_var = tk.StringVar(value=self.colors["bg"])
        self.color_odd_var = tk.StringVar(value=self.colors["odd_bg"])
        self.color_even_var = tk.StringVar(value=self.colors["even_bg"])
        self.color_font_var = tk.StringVar(value=self.colors["font"])
        self.color_clock_var = tk.StringVar(value=self.colors["clock"])
        self.flap_seconds_var = tk.StringVar(value=str(self.flap_seconds))
        self.page_flip_seconds_var = tk.StringVar(value=str(self.page_flip_seconds))
        self.use_split_flap_var = tk.BooleanVar(value=getattr(self, "use_split_flap", True))
        self._ui_mode_key_to_label = {
            mode_key: PRESET_LABELS.get(mode_key, mode_key)
            for mode_key in self._ordered_ui_mode_keys()
        }
        self._ui_mode_label_to_key = {
            label: mode_key for mode_key, label in self._ui_mode_key_to_label.items()
        }
        active_mode_key = get_active_preset_name()
        active_mode_label = self._ui_mode_key_to_label.get(active_mode_key, active_mode_key)
        self.ui_mode_label_var = tk.StringVar(value=active_mode_label)

        def sec(text):
            tk.Label(outer, text=text, bg=C["bg"], fg=C["btn_primary"], font=("맑은 고딕", 9, "bold")).pack(
                anchor="w", pady=(10, 4)
            )

        def card_frame():
            frame = tk.Frame(outer, bg=C["white"], highlightthickness=1, highlightbackground=C["border"], padx=16, pady=10)
            frame.pack(fill="x")
            return frame

        def txt_row(parent_f, label_text, var):
            row = tk.Frame(parent_f, bg=C["white"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label_text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=16, anchor="w").pack(
                side="left"
            )
            tk.Entry(
                row,
                textvariable=var,
                width=36,
                bg=C["white"],
                fg=C["text"],
                relief="solid",
                bd=1,
                highlightthickness=1,
                highlightbackground=C["border"],
                highlightcolor=C["border_focus"],
                font=("맑은 고딕", 10),
            ).pack(side="left")

        sec("문구 설정")
        tf = card_frame()
        txt_row(tf, "프로그램 타이틀", self.app_title_var)
        txt_row(tf, "전광판 타이틀", self.fs_title_var)

        sec("메인 UI 모드")
        mf = card_frame()
        mode_row = tk.Frame(mf, bg=C["white"])
        mode_row.pack(fill="x", pady=3)
        tk.Label(mode_row, text="UI 모드", bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=16,
                 anchor="w").pack(side="left")
        mode_values = [self._ui_mode_key_to_label[key] for key in self._ordered_ui_mode_keys()]
        ttk.Combobox(
            mode_row,
            state="readonly",
            width=24,
            values=mode_values,
            textvariable=self.ui_mode_label_var,
        ).pack(side="left")
        tk.Label(
            mf,
            text="저장 후 앱 재실행 시 전체 화면에 반영됩니다.",
            bg=C["white"],
            fg=C["placeholder"],
            font=("맑은 고딕", 8),
        ).pack(anchor="w", padx=4, pady=(2, 0))

        sec("전광판 색상 설정")
        cf = card_frame()

        def pick_color(var_obj, btn_obj):
            code = colorchooser.askcolor(title="색상 선택", initialcolor=var_obj.get())[1]
            if code:
                var_obj.set(code)
                btn_obj.config(bg=code)

        color_defs = [
            ("바탕화면 색상", self.color_bg_var, 0, 0),
            ("홀수열 배경", self.color_odd_var, 1, 0),
            ("짝수열 배경", self.color_even_var, 2, 0),
            ("폰트 색상", self.color_font_var, 0, 3),
            ("시계 색상", self.color_clock_var, 1, 3),
        ]
        for ltext, var, r, c in color_defs:
            tk.Label(cf, text=ltext, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=14, anchor="w").grid(
                row=r, column=c, padx=(14, 4), pady=6
            )
            btn = tk.Button(cf, width=8, bg=var.get(), relief="solid", bd=1, cursor="hand2")
            btn.config(command=lambda v=var, b=btn: pick_color(v, b))
            btn.grid(row=r, column=c + 1, pady=6, padx=(0, 20))

        sec("전광판 시간 설정")
        tmf = card_frame()
        for r, (lbl_text, var) in enumerate(
            [
                ("플립 애니메이션 시간 (초)", self.flap_seconds_var),
                ("다음 페이지 대기 시간 (초)", self.page_flip_seconds_var),
            ]
        ):
            tk.Label(tmf, text=lbl_text, bg=C["white"], fg=C["label"], font=("맑은 고딕", 9), width=24, anchor="w").grid(
                row=r, column=0, padx=4, pady=6
            )
            tk.Entry(
                tmf,
                textvariable=var,
                width=10,
                justify="center",
                bg=C["white"],
                fg=C["text"],
                relief="solid",
                bd=1,
                highlightthickness=1,
                highlightbackground=C["border"],
                highlightcolor=C["border_focus"],
                font=("맑은 고딕", 10),
            ).grid(row=r, column=1, padx=8, pady=6)
        tk.Checkbutton(
            tmf,
            text="플립보드 애니메이션 사용 (저사양 PC에서 해제 가능)",
            variable=self.use_split_flap_var,
            bg=C["white"],
            fg=C["text"],
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=4)
        tk.Label(
            tmf,
            text="기본값: 플립 애니메이션 5초 / 페이지 전환 7초",
            bg=C["white"],
            fg=C["placeholder"],
            font=("맑은 고딕", 8),
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=4, pady=(0, 4))

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=12)
        btn_row = tk.Frame(outer, bg=C["bg"])
        btn_row.pack(fill="x")
        self._btn(btn_row, "기본 설정으로", self._reset_others_settings, C["btn_warning"]).pack(side="left")
        self._btn(btn_row, "저장", self._save_others_settings, C["btn_success"], width=10).pack(side="right")

    def _reset_others_settings(self):
        if messagebox.askyesno("확인", "타이틀/색상/전광판 시간 설정을 기본값으로 되돌리시겠습니까?"):
            self.app_title_var.set(DEFAULT_SETTINGS["app_title"])
            self.fs_title_var.set(DEFAULT_SETTINGS["fs_title"])
            self.color_bg_var.set(DEFAULT_SETTINGS["colors"]["bg"])
            self.color_odd_var.set(DEFAULT_SETTINGS["colors"]["odd_bg"])
            self.color_even_var.set(DEFAULT_SETTINGS["colors"]["even_bg"])
            self.color_font_var.set(DEFAULT_SETTINGS["colors"]["font"])
            self.color_clock_var.set(DEFAULT_SETTINGS["colors"]["clock"])
            self.flap_seconds_var.set(str(DEFAULT_SETTINGS["display_timing"]["flap_seconds"]))
            self.page_flip_seconds_var.set(str(DEFAULT_SETTINGS["display_timing"]["page_flip_seconds"]))
            self.use_split_flap_var.set(DEFAULT_SETTINGS["display_timing"].get("use_split_flap", True))
            self.ui_mode_label_var.set(self._ui_mode_key_to_label.get(LEGACY_PRESET_NAME, LEGACY_PRESET_NAME))
            self._save_others_settings()
            messagebox.showinfo("안내", "기본 설정으로 복구되었습니다. 변경사항을 반영하기 위해 설정창을 닫았다 다시 열어주세요.")

    def _save_others_settings(self):
        flap_seconds = self._safe_positive_int(
            self.flap_seconds_var.get().strip(),
            DEFAULT_SETTINGS["display_timing"]["flap_seconds"],
        )
        page_flip_seconds = self._safe_positive_int(
            self.page_flip_seconds_var.get().strip(),
            DEFAULT_SETTINGS["display_timing"]["page_flip_seconds"],
        )

        if flap_seconds <= 0 or page_flip_seconds <= 0:
            messagebox.showerror("입력 오류", "전광판 시간은 1초 이상의 숫자로 입력해 주세요.")
            return

        self.app_title = self.app_title_var.get().strip()
        self.fs_title = self.fs_title_var.get().strip()
        self.colors["bg"] = self.color_bg_var.get()
        self.colors["odd_bg"] = self.color_odd_var.get()
        self.colors["even_bg"] = self.color_even_var.get()
        self.colors["font"] = self.color_font_var.get()
        self.colors["clock"] = self.color_clock_var.get()
        self.flap_seconds = flap_seconds
        self.page_flip_seconds = page_flip_seconds
        self.use_split_flap = self.use_split_flap_var.get()

        selected_ui_mode_label = self.ui_mode_label_var.get().strip()
        selected_ui_mode_key = self._ui_mode_label_to_key.get(selected_ui_mode_label, LEGACY_PRESET_NAME)
        current_ui_mode_key = get_active_preset_name()
        ui_mode_changed = selected_ui_mode_key != current_ui_mode_key
        try:
            set_active_preset_name(selected_ui_mode_key)
        except Exception as e:
            messagebox.showerror("저장 오류", f"UI 모드 저장 중 오류가 발생했습니다.\n{e}")
            return

        self.settings["app_title"] = self.app_title
        self.settings["fs_title"] = self.fs_title
        self.settings["colors"] = self.colors
        self.settings["ui_mode"] = selected_ui_mode_key
        self.settings["display_timing"] = {
            "flap_seconds": self.flap_seconds,
            "page_flip_seconds": self.page_flip_seconds,
            "use_split_flap": self.use_split_flap,
        }

        self.save_settings(self.settings)
        self.root.title(self.app_title)
        if ui_mode_changed:
            self.C = build_main_palette(selected_ui_mode_key)
            self._setup_styles()
            self.root.configure(bg=self.C["bg"])

        if getattr(self, "flap_refresh_timer_id", None):
            self.root.after_cancel(self.flap_refresh_timer_id)
            self.flap_refresh_timer_id = None
        self._periodic_flap_refresh()
        self._schedule_fs_page_flip()
        self._sync_full_screen(force_rebuild=True)
        if ui_mode_changed:
            messagebox.showinfo("저장 완료", "기타 설정이 저장되었습니다.\nUI 모드 변경은 앱 재실행 시 전체 화면에 반영됩니다.")
        else:
            messagebox.showinfo("저장 완료", "기타 설정이 성공적으로 저장되었습니다.")

    def _make_dispatch_order_key(self, d_type: str, source: Dict[str, Any]) -> str:
        ls_id = str(source.get("ls_id", source.get("_ls_id", "")) or "").strip()
        if ls_id:
            return f"{d_type}|LS|{ls_id}"
        camp = self._normalize_text(source.get("camp", source.get("캠프명", "")))
        car_num = self._normalize_text(source.get("car_num", source.get("차량번호", "")))
        phone = re.sub(r"\D", "", str(source.get("phone", source.get("전화번호", "")) or ""))
        hocha = self._normalize_text(source.get("hocha", source.get("호차", "")))
        return f"{d_type}|{camp}|{car_num}|{phone}|{hocha}"

    def _dispatch_dock_sort_key(self, dock_name: str) -> Tuple[int, int]:
        dock = str(dock_name or "").strip()
        if dock:
            if dock in self.dock_list:
                return 0, self.dock_list.index(dock)
            if dock in DEFAULT_DOCKS:
                return 0, DEFAULT_DOCKS.index(dock)
            match = re.match(r"^(\d+)\s*([A-Za-z])$", dock)
            if match:
                suffix_order = {"A": 0, "B": 1, "C": 2}
                numeric = int(match.group(1))
                suffix = suffix_order.get(match.group(2).upper(), 9)
                return 0, numeric * 10 + suffix
        return 1, 99999

    def _apply_saved_dispatch_order(self, d_type: str, rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        saved_keys = list(self.dispatch_manual_order.get(d_type, []) or [])
        if not saved_keys:
            return rows
        keyed_rows: Dict[str, List[Dict[str, Any]]] = {}
        for row in rows:
            order_key = row.get("_order_key") or self._make_dispatch_order_key(d_type, row)
            keyed_rows.setdefault(order_key, []).append(row)

        ordered_rows = []
        used_keys = set()
        used_row_ids = set()
        for order_key in saved_keys:
            row_bucket = keyed_rows.get(order_key, [])
            if row_bucket and order_key not in used_keys:
                ordered_rows.extend(row_bucket)
                used_keys.add(order_key)
                used_row_ids.update(id(row) for row in row_bucket)

        for row in rows:
            if id(row) not in used_row_ids:
                ordered_rows.append(row)
                used_row_ids.add(id(row))
        return ordered_rows

    def _make_dispatch_order_key_from_tree_item(self, values: Tuple[Any, ...], tags: Tuple[str, ...]) -> str:
        d_type = values[1] if len(values) > 1 else ""
        oid = self._extract_oid(tags)
        registry_info = self.dispatch_registry.get(oid, {}) if oid else {}
        ls_id = registry_info.get("ls_id", "")
        if ls_id:
            return f"{d_type}|LS|{ls_id}"
        return self._make_dispatch_order_key(
            d_type,
            {
                "캠프명": values[2] if len(values) > 2 else "",
                "차량번호": values[5] if len(values) > 5 else "",
                "전화번호": values[6] if len(values) > 6 else "",
                "호차": values[4] if len(values) > 4 else "",
            },
        )

    def _save_dispatch_manual_order(self, d_type: str) -> None:
        rows = self.dispatch_source_rows.get(d_type, [])
        self.dispatch_manual_order[d_type] = [row.get("_order_key") or self._make_dispatch_order_key(d_type, row) for row in rows]
        self.settings["dispatch_manual_order"] = self.dispatch_manual_order
        self.save_settings(self.settings)

    def _fill_dispatch_tab(self, d_type: str) -> None:
        C = self.C
        meta = self.dispatch_tab_meta.get(d_type)
        if not meta:
            return

        wrap = meta["wrap"]
        for child in wrap.winfo_children():
            child.destroy()

        rows = self.dispatch_source_rows.get(d_type, [])
        tbl = tk.Frame(wrap, bg=C["bg"])
        tbl.pack(anchor="n")

        col_px = meta["col_px"]
        hdrs = meta["hdrs"]
        row_widgets = []
        ro_keys = {"캠프명", "차량번호", "간선사"}

        for ci, header in enumerate(hdrs):
            tbl.grid_columnconfigure(ci, minsize=col_px[header], weight=0)
            tk.Label(tbl, text=header, font=("맑은 고딕", 9, "bold"), bg=C["heading"], fg=C["label"], anchor="center",
                     relief="flat", padx=4, pady=8).grid(row=0, column=ci, padx=2, pady=(0, 3), sticky="nsew")

        for idx, row in enumerate(rows):
            row_vars = {k: tk.StringVar(value=str(row.get(k, ""))) for k in DISPATCH_HEADERS}
            entries = {}
            grid_row = idx + 1
            ro_bg = C["heading"]

            order_frame = tk.Frame(tbl, bg=C["bg"])
            order_frame.grid(row=grid_row, column=0, padx=2, pady=2, sticky="nsew")

            btn_up = tk.Button(order_frame, text="▲", font=("Arial", 8), width=3, bg=C["white"], fg=C["btn_primary"],
                               command=lambda current=idx, dt=d_type: self._move_dispatch_row(dt, current, -1))
            btn_up.pack(side="left", padx=(0, 2))
            if idx == 0:
                btn_up.config(state="disabled")

            btn_down = tk.Button(order_frame, text="▼", font=("Arial", 8), width=3, bg=C["white"], fg=C["btn_primary"],
                                 command=lambda current=idx, dt=d_type: self._move_dispatch_row(dt, current, 1))
            btn_down.pack(side="left")
            if idx == len(rows) - 1:
                btn_down.config(state="disabled")

            for ci, header in enumerate(DISPATCH_HEADERS, start=1):
                is_ro = header in ro_keys
                entry = tk.Entry(tbl, textvariable=row_vars[header], justify="center", relief="solid", bd=1,
                                 font=("맑은 고딕", 10), bg=ro_bg if is_ro else C["white"],
                                 fg=C["label"] if is_ro else C["text"], readonlybackground=ro_bg, highlightthickness=1,
                                 highlightbackground=C["border"], highlightcolor=C["border_focus"])
                if is_ro:
                    entry.config(state="readonly")
                entry.grid(row=grid_row, column=ci, padx=2, pady=2, sticky="nsew", ipady=7)
                entries[header] = entry

            action_frame = tk.Frame(tbl, bg=C["bg"])
            action_frame.grid(row=grid_row, column=len(DISPATCH_HEADERS) + 1, padx=6, pady=2, sticky="nsew")
            self._btn(action_frame, "미리보기", lambda dv=row_vars, dt=d_type: self.preview_dispatch_row(dt, dv), C["btn_success"]).pack(
                side="left", padx=(0, 4)
            )
            self._btn(action_frame, "프린트", lambda dv=row_vars, dt=d_type, ls_id=row.get("_ls_id", ""): self.print_dispatch_row(dt, dv, ls_id), C["btn_primary"]).pack(
                side="left", padx=(0, 4)
            )

            # ── 봉인씰 등록 버튼 ────────────────────────────────────────────────
            _row_ls_id = row.get("_ls_id", "")
            _seal_btn_color = C.get("btn_warning", "#E67E22") if _row_ls_id else C.get("btn_disabled", "#AAAAAA")
            _seal_btn = self._btn(
                action_frame,
                "봉인씰등록",
                lambda dv=row_vars, lid=_row_ls_id, did=row.get("_dep_id", ""), aid=row.get("_arr_id", ""), raw_p=row.get("_raw", {}): self._upload_seal_to_ls(
                    ls_id=lid,
                    seal_left=dv["좌측윙"].get(),
                    seal_right=dv["우측윙"].get(),
                    seal_back=dv["후면"].get(),
                    camp_name=dv["캠프명"].get(),
                    dep_id=did,
                    arr_id=aid,
                    raw_payload=raw_p,
                ),
                _seal_btn_color,
            )
            _seal_btn.pack(side="left")
            if not _row_ls_id:
                _seal_btn.config(state="disabled")

            for seal_index, seal_key in enumerate(["좌측윙", "우측윙", "후면"]):
                entries[seal_key].bind(
                    "<KeyRelease>",
                    lambda event, t=d_type, r=idx, c=seal_index: self._on_dispatch_seal_keyrelease(t, r, c),
                )
                entries[seal_key].bind("<Return>", lambda event: "break")

            row_widgets.append({
                "vars": row_vars, 
                "entries": entries, 
                "order_key": row.get("_order_key", ""), 
                "_ls_id": row.get("_ls_id", ""), 
                "_dep_id": row.get("_dep_id", ""), 
                "_arr_id": row.get("_arr_id", "")
            })

        if not rows:
            tk.Label(wrap, text="해당 타입의 데이터가 없습니다.", bg=C["bg"], fg=C["placeholder"], font=("맑은 고딕", 11)).pack(
                pady=40
            )

        meta["rows"] = row_widgets

    def _move_dispatch_row(self, d_type: str, idx: int, delta: int) -> None:
        rows = self.dispatch_source_rows.get(d_type, [])
        if not rows:
            return
        target_idx = idx + delta
        if not (0 <= target_idx < len(rows)):
            return

        meta = self.dispatch_tab_meta.get(d_type, {})
        canvas = meta.get("canvas")
        scroll_pos = canvas.yview()[0] if canvas else 0

        rows[idx], rows[target_idx] = rows[target_idx], rows[idx]
        self._save_dispatch_manual_order(d_type)
        self._fill_dispatch_tab(d_type)

        if canvas:
            self.root.after(10, lambda: canvas.yview_moveto(scroll_pos))

    def open_dispatch_window(self, keep_manual=False) -> None:
        time_lookup: Dict[str, str] = {}
        dock_lookup: Dict[str, str] = {}
        for iid in self.tree.get_children():
            values = self.tree.item(iid, "values")
            tags = self.tree.item(iid, "tags")
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = values[7] if len(values) > 7 and values[7] else "99:99"
                dock_lookup[oid] = values[3] if len(values) > 3 and values[3] else ""

        for info in self.detached_items.values():
            values = info["values"]
            tags = info["tags"]
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = values[7] if len(values) > 7 and values[7] else "99:99"
                dock_lookup[oid] = values[3] if len(values) > 3 and values[3] else ""

        previous_tab = 0
        if hasattr(self, "dispatch_nb") and self.dispatch_nb.winfo_exists():
            try:
                previous_tab = self.dispatch_nb.index(self.dispatch_nb.select())
            except Exception:
                previous_tab = 0

        self.dispatch_source_rows = {k: [] for k in DISPATCH_TABS}
        for oid, info in self.dispatch_registry.items():
            d_type = info.get("d_type", "")
            if d_type not in self.dispatch_source_rows:
                continue

            camp_name = info.get("camp", "")
            dock_name = dock_lookup.get(oid) or self._lookup_dock_for_camp(d_type, camp_name) or "-"
            row = {
                "캠프명": camp_name,
                "차량번호": info.get("car_num", ""),
                "전화번호": info.get("phone", ""),
                "이름": info.get("driver_name", ""),
                "간선사": info.get("truckVendor", "") or d_type,
                "좌측윙": "", "우측윙": "", "후면": "",
                "_time": time_lookup.get(oid, "99:99"),
                "호차": info.get("hocha", ""),
                "_dock": dock_name,
                "_ls_id": info.get("ls_id", ""),
                "_dep_id": info.get("dep_id", ""),
                "_arr_id": info.get("arr_id", ""),
                "_raw": info.get("_raw", {}),
            }
            row["_order_key"] = self._make_dispatch_order_key(d_type, row)
            self.dispatch_source_rows[d_type].append(row)

        def extract_hocha_num(hocha_str: str) -> int:
            match = re.search(r"\d+", str(hocha_str))
            return int(match.group()) if match else 9999

        for d_type, rows in self.dispatch_source_rows.items():
            rows.sort(
                key=lambda row: (
                    self._dispatch_dock_sort_key(row.get("_dock", "")),
                    self._normalize_text(row.get("캠프명", "")),
                    extract_hocha_num(row.get("호차", "")),
                    str(row.get("_time", "99:99")),
                    self._make_dispatch_order_key(d_type, row),
                )
            )
            self.dispatch_source_rows[d_type] = self._apply_saved_dispatch_order(d_type, rows)

        if hasattr(self, "dispatch_win") and self.dispatch_win.winfo_exists():
            self.dispatch_win.destroy()

        C = self.C
        self.dispatch_win = tk.Toplevel(self.root)
        self.dispatch_win.title("출차서 출력")
        self.dispatch_win.geometry("1680x850")
        self.dispatch_win.configure(bg=C["bg"])
        self.dispatch_win.transient(self.root)

        self._dispatch_active_tab = tk.StringVar(value=DISPATCH_TABS[min(previous_tab, len(DISPATCH_TABS) - 1)])

        hdr = tk.Frame(self.dispatch_win, bg=C["header"], height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="출차서 출력", bg=C["header"], fg="white", font=("맑은 고딕", 12, "bold")).pack(
            side="left", padx=16, pady=12
        )
        tk.Label(hdr, text="도크 순서 정렬 + 수동 위치 저장", bg=C["header"], fg=C["header_muted"], font=("맑은 고딕", 9)).pack(
            side="left", padx=8
        )
        self._btn(hdr, "새로고침", self.open_dispatch_window, C["btn_secondary"]).pack(side="right", padx=(6, 16), pady=10)
        self._btn(hdr, "일괄 프린트", self.batch_print_dispatch, C["btn_danger"]).pack(side="right", pady=10)

        self.ls_print_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            hdr,
            text="LS 프린트",
            variable=self.ls_print_var,
            bg=C["header"],
            fg=C["header_txt"],
            activebackground=C["header"],
            activeforeground=C["header_txt"],
            selectcolor=C["header"],
            font=("맑은 고딕", 10, "bold")
        ).pack(side="right", padx=(0, 16))

        style = ttk.Style(self.dispatch_win)
        style.configure("Disp.TNotebook", background=C["bg"], borderwidth=0)
        style.configure("Disp.TNotebook.Tab", background=C["bg"], foreground=C["label"], padding=[16, 8],
                        font=("맑은 고딕", 10, "bold"), borderwidth=0)
        style.map("Disp.TNotebook.Tab", background=[("selected", C["white"])],
                  foreground=[("selected", C["btn_primary"])])

        self.dispatch_nb = ttk.Notebook(self.dispatch_win, style="Disp.TNotebook")
        self.dispatch_nb.pack(fill="both", expand=True, padx=0, pady=0)

        def _on_tab_changed(event):
            try:
                idx = self.dispatch_nb.index(self.dispatch_nb.select())
                self._dispatch_active_tab.set(DISPATCH_TABS[idx])
            except Exception:
                pass

        self.dispatch_nb.bind("<<NotebookTabChanged>>", _on_tab_changed)

        self.dispatch_tab_frames.clear()
        self.dispatch_tab_meta.clear()
        for d_type in DISPATCH_TABS:
            tab = tk.Frame(self.dispatch_nb, bg=C["bg"])
            self.dispatch_nb.add(tab, text=d_type)
            self._build_dispatch_tab(tab, d_type)

        self.dispatch_nb.select(min(previous_tab, len(DISPATCH_TABS) - 1))

    def _build_dispatch_tab(self, parent: tk.Frame, d_type: str) -> None:
        C = self.C
        outer = tk.Frame(parent, bg=C["bg"])
        outer.pack(fill="both", expand=True)

        bottom_bar = tk.Frame(parent, bg=C["white"], highlightthickness=1, highlightbackground=C["border"])
        bottom_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 10))
        if d_type not in {"SD3", "D2", "D4", "간선"}:
            self._btn(bottom_bar, "고임목 출력", lambda dt=d_type: self.print_goimok(dt), C["btn_warning"]).pack(
                side="right", padx=12, pady=8
            )

        auto_right_var = tk.BooleanVar(value=True)
        toggle_frame = tk.Frame(bottom_bar, bg=C["white"])
        toggle_frame.pack(side="left", padx=14, pady=6)
        tk.Checkbutton(
            toggle_frame,
            text="봉인씰 자동 입력",
            variable=auto_right_var,
            command=lambda: None,
            bg=C["white"],
            fg=C["label"],
            activebackground=C["white"],
            selectcolor=C["white"],
            font=("맑은 고딕", 9),
        ).pack()

        canvas = tk.Canvas(outer, highlightthickness=0, bg=C["bg"])
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=C["bg"])
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        scroll_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(canvas_window, width=event.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _dispatch_mousewheel(event):
            if getattr(event, "delta", 0):
                canvas.yview_scroll(int(-event.delta / 120), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        def _bind_dispatch_wheel(_event=None):
            canvas.bind_all("<MouseWheel>", _dispatch_mousewheel)
            canvas.bind_all("<Button-4>", _dispatch_mousewheel)
            canvas.bind_all("<Button-5>", _dispatch_mousewheel)

        def _unbind_dispatch_wheel(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_dispatch_wheel)
        canvas.bind("<Leave>", _unbind_dispatch_wheel)
        scroll_frame.bind("<Enter>", _bind_dispatch_wheel)
        scroll_frame.bind("<Leave>", _unbind_dispatch_wheel)

        wrap = tk.Frame(scroll_frame, bg=C["bg"])
        wrap.pack(fill="x", padx=14, pady=(12, 8), anchor="n")
        self.dispatch_tab_meta[d_type] = {
            "rows": [],
            "canvas": canvas,
            "wrap": wrap,
            "auto_right_wing": auto_right_var,
            "col_px": {
                "순서": 78,
                "캠프명": 150,
                "차량번호": 158,
                "전화번호": 158,
                "이름": 118,
                "간선사": 118,
                "좌측윙": 138,
                "우측윙": 138,
                "후면": 138,
                "출력하기": 178,
            },
            "hdrs": ["순서"] + list(DISPATCH_HEADERS) + ["출력하기"],
        }
        self._fill_dispatch_tab(d_type)

    def batch_print_dispatch(self) -> None:
        if not hasattr(self, "dispatch_tab_meta") or not self.dispatch_tab_meta:
            return
        active_tab = getattr(self, "_dispatch_active_tab", None)
        d_type = active_tab.get() if active_tab else DISPATCH_TABS[0]
        rows = self.dispatch_tab_meta.get(d_type, {}).get("rows", [])
        payloads = []
        for row in rows:
            payload = self._row_vars_to_payload(row["vars"])
            if payload.get("캠프명"):
                payload["_ls_id"] = row.get("_ls_id", "")
                payloads.append(payload)

        payloads = self._expand_dispatch_payloads(d_type, payloads)

        if not payloads:
            messagebox.showinfo("안내", f"[{d_type}] 탭에 프린트할 데이터가 없습니다.")
            return

        use_ls_print = False
        if hasattr(self, "ls_print_var"):
            use_ls_print = self.ls_print_var.get()
            
        if use_ls_print:
            if not messagebox.askyesno("확인", f"[{d_type}] 탭의 총 {len(payloads)}건을 LS 방식으로 일괄 프린트하시겠습니까?"):
                return
            self._print_ls(payloads)
            return
            
        if not messagebox.askyesno("확인", f"[{d_type}] 탭의 총 {len(payloads)}건을 마스터양식(엑셀)으로 일괄 프린트하시겠습니까?"):
            return

        if not (platform.system() == "Windows" and WIN32COM_AVAILABLE):
            messagebox.showerror("오류", "일괄 프린트 고속 모드는 Windows 환경에서 win32com 사용 시에만 지원됩니다.")
            return

        master_path = self._get_master_config_path()
        if not master_path:
            return
        template_path = self._get_dispatch_template_path()
        if not template_path:
            return

        try:
            master_cfg = self.excel_manager.load_master_config(master_path)
        except Exception as e:
            messagebox.showerror("양식 설정 오류", str(e))
            return

        run_date = getattr(self, "api_run_date", "")
        success, errors = 0, []

        for payload in payloads:
            if "간선사" not in payload or not payload["간선사"]:
                payload["간선사"] = d_type
            camp_name = payload.get("캠프명", "")
            matched_key = self._find_best_match_key(camp_name, list(master_cfg.keys()))
            payload["_cfg"] = master_cfg.get(matched_key) if matched_key else None
            payload["_sheet"] = payload["_cfg"].get("출차서_시트명", "").strip() if payload["_cfg"] else ""

        try:
            pythoncom.CoInitialize()
            xl = win32.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
        except Exception as e:
            messagebox.showerror("엑셀 실행 오류", f"엑셀(COM)을 실행할 수 없습니다:\n{e}")
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return

        cleanup_files = []
        wb2 = None
        try:
            wb2 = xl.Workbooks.Open(os.path.abspath(template_path))
            last_sheet_name = None
            com_ws = None

            for payload in payloads:
                camp_name = payload.get("캠프명", "")
                matched_cfg = payload.get("_cfg")
                sheet_name = payload.get("_sheet", "")

                if not matched_cfg:
                    errors.append(f"[{camp_name}] 마스터 양식설정에서 이 캠프를 찾을 수 없습니다.")
                    continue

                if not sheet_name:
                    errors.append(f"[{camp_name}] 출차서 시트명이 지정되지 않았습니다.")
                    continue

                if sheet_name != last_sheet_name or com_ws is None:
                    last_sheet_name = sheet_name
                    com_ws = None
                    for si in range(1, wb2.Worksheets.Count + 1):
                        ns = self._normalize_text(sheet_name)
                        nw = self._normalize_text(wb2.Worksheets(si).Name)
                        if ns == nw or ns in nw or nw in ns:
                            com_ws = wb2.Worksheets(si)
                            break
                    if com_ws is None:
                        com_ws = wb2.ActiveSheet

                def _current_com_write(addr, val, ws=com_ws):
                    if addr:
                        try:
                            ws.Range(addr.upper()).Value = val if val else ""
                        except Exception:
                            pass

                try:
                    for idx in range(4):
                        postfix = str(idx + 1)
                        _current_com_write(matched_cfg.get(f"운행일자{postfix}", ""), run_date)
                        _current_com_write(matched_cfg.get(f"차량번호{postfix}", ""), payload.get("차량번호", ""))
                        _current_com_write(matched_cfg.get(f"이름{postfix}", ""), payload.get("이름", ""))
                        _current_com_write(matched_cfg.get(f"연락처{postfix}", ""), payload.get("전화번호", ""))
                        _current_com_write(matched_cfg.get(f"좌측윙{postfix}", ""), payload.get("좌측윙", ""))
                        _current_com_write(matched_cfg.get(f"우측윙{postfix}", ""), payload.get("우측윙", ""))
                        _current_com_write(matched_cfg.get(f"후면도어{postfix}", ""), payload.get("후면", ""))
                        _current_com_write(matched_cfg.get(f"바코드{postfix}", ""), payload.get("바코드", ""))
                        _current_com_write(matched_cfg.get(f"간선사{postfix}", ""), payload.get("간선사", ""))
                    com_ws.PrintOut()
                    success += 1
                except Exception as e:
                    errors.append(f"[{camp_name}] 출력 중 오류: {e}")

        finally:
            try:
                if wb2 is not None:
                    wb2.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                xl.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        self._schedule_temp_cleanup(cleanup_files)
        if errors:
            messagebox.showwarning("일괄 프린트 완료", f"성공: {success}건\n실패: {len(errors)}건\n\n" + "\n".join(errors[:10]))
        else:
            messagebox.showinfo("완료", f"[{d_type}] 총 {success}건 프린트 완료.")

    def _expand_dispatch_payloads(self, d_type: str, payloads: List[Dict[str, str]]) -> List[Dict[str, str]]:
        expanded_payloads = []
        for source_payload in payloads:
            payload = dict(source_payload)
            if "간선사" not in payload or not payload["간선사"]:
                payload["간선사"] = d_type
            expanded_payloads.append(payload)

            normalized_camp = self._normalize_text(payload.get("캠프명", ""))
            if d_type in ["SD2", "D1"] and normalized_camp in ["동인천2", "동인천2(SH)"]:
                copied = payload.copy()
                copied["캠프명"] = "인천3"
                expanded_payloads.append(copied)
            elif d_type == "W2" and normalized_camp in ["인천3", "인천3(SH)"]:
                copied = payload.copy()
                copied["캠프명"] = "TW"
                expanded_payloads.append(copied)
        return expanded_payloads

    def _build_dispatch_rows_by_type(self) -> Dict[str, List[Dict[str, Any]]]:
        time_lookup: Dict[str, str] = {}
        dock_lookup: Dict[str, str] = {}

        for iid in self.tree.get_children():
            values = self.tree.item(iid, "values")
            tags = self.tree.item(iid, "tags")
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = values[7] if len(values) > 7 and values[7] else "99:99"
                dock_lookup[oid] = values[3] if len(values) > 3 and values[3] else ""

        for info in self.detached_items.values():
            values = info["values"]
            tags = info["tags"]
            oid = self._extract_oid(tags)
            if oid:
                time_lookup[oid] = values[7] if len(values) > 7 and values[7] else "99:99"
                dock_lookup[oid] = values[3] if len(values) > 3 and values[3] else ""

        rows_by_type: Dict[str, List[Dict[str, Any]]] = {key: [] for key in DISPATCH_TABS}
        for oid, info in self.dispatch_registry.items():
            d_type = info.get("d_type", "")
            if d_type not in rows_by_type:
                continue

            camp_name = info.get("camp", "")
            row = {
                "캠프명": camp_name,
                "차량번호": info.get("car_num", ""),
                "전화번호": info.get("phone", ""),
                "이름": info.get("driver_name", ""),
                "간선사": info.get("truckVendor", "") or d_type,
                "좌측윙": "",
                "우측윙": "",
                "후면": "",
                "_time": time_lookup.get(oid, "99:99"),
                "호차": info.get("hocha", ""),
                "_dock": dock_lookup.get(oid) or self._lookup_dock_for_camp(d_type, camp_name) or "-",
                "_ls_id": info.get("ls_id", ""),
            }
            row["_order_key"] = self._make_dispatch_order_key(d_type, row)
            rows_by_type[d_type].append(row)

        def extract_hocha_num(hocha_str: str) -> int:
            match = re.search(r"\d+", str(hocha_str))
            return int(match.group()) if match else 9999

        for d_type, rows in rows_by_type.items():
            rows.sort(
                key=lambda row: (
                    self._dispatch_dock_sort_key(row.get("_dock", "")),
                    self._normalize_text(row.get("캠프명", "")),
                    extract_hocha_num(row.get("호차", "")),
                    str(row.get("_time", "99:99")),
                    self._make_dispatch_order_key(d_type, row),
                )
            )
            rows_by_type[d_type] = self._apply_saved_dispatch_order(d_type, rows)
        return rows_by_type

    def get_formatted_list(self) -> str:
        import re as _re

        def extract_hocha_num(hocha_str):
            match = _re.search(r"\d+", str(hocha_str))
            return int(match.group()) if match else 999

        all_items_data = []
        for iid in self.tree.get_children():
            all_items_data.append((self.tree.item(iid, "values"), self.tree.item(iid, "tags")))
        for info in self.detached_items.values():
            all_items_data.append((info["values"], info["tags"]))

        dt_groups = {}
        for values, tags in all_items_data:
            dt = values[1]
            dt_groups.setdefault(dt, []).append((values, tags))

        dispatch_rows_by_type = self._build_dispatch_rows_by_type()
        dispatch_order_maps = {}
        for d_type, rows in dispatch_rows_by_type.items():
            dispatch_order_maps[d_type] = {
                row.get("_order_key") or self._make_dispatch_order_key(d_type, row): idx
                for idx, row in enumerate(rows)
            }

        ordered_types = [dt for dt in DISPATCH_TABS if dt in dt_groups]
        ordered_types.extend(sorted(dt for dt in dt_groups.keys() if dt not in ordered_types))

        result = []
        for dt in ordered_types:
            items = dt_groups[dt]
            dispatch_order_map = dispatch_order_maps.get(dt, {})

            def get_dock_index(dock_name):
                try:
                    return self.dock_list.index(dock_name)
                except ValueError:
                    return 999

            camp_counts = {}
            for values, _tags in items:
                camp = values[2]
                camp_counts[camp] = camp_counts.get(camp, 0) + 1

            # 출차서 창에서 수동 조정한 순서가 있으면 우선 반영하고,
            # 저장된 순서가 없는 항목은 도크/캠프/호차 순으로 안정 정렬한다.
            items.sort(
                key=lambda item: (
                    dispatch_order_map.get(self._make_dispatch_order_key_from_tree_item(item[0], item[1]), 99999),
                    get_dock_index(item[0][3]),
                    item[0][2],
                    extract_hocha_num(item[0][4]),
                    item[0][5],
                    item[0][6],
                )
            )

            lines = []
            for values, _tags in items:
                camp = values[2]
                car_num = values[5]
                phone = values[6]
                hocha = values[4]

                if dt == "KPP":
                    line = f"{car_num} {phone}".strip()
                else:
                    line = f"{camp} - {car_num} - {phone}"
                    if camp_counts.get(camp, 0) >= 2:
                        line += f" - [{hocha}]"

                lines.append(line)
                lines.append("")

            result.append(f"=============[{dt}]=============")
            result.extend(lines)

        return "\n".join(result).strip()

    def _is_goimok_ansung5_first(self, row: Dict[str, Any]) -> bool:
        camp_name = self._normalize_text(row.get("캠프명", ""))
        hocha = self._normalize_text(row.get("호차", ""))
        return camp_name.startswith("안성5HUB") and hocha in {"1호차", "1"}

    def _collect_goimok_rows(self, target_d_type: str) -> List[Tuple[str, str, str]]:
        rules = {
            "SD2": ["SD2", "SD3"],
            "D1": ["D1", "D2"],
            "D3": ["D3", "D4"],
            "W2": ["간선", "W2"],
        }
        source_types = rules.get(target_d_type, [target_d_type])
        collected_rows: List[Tuple[str, str, str]] = []

        for source_type in source_types:
            for row in self.dispatch_source_rows.get(source_type, []):
                if target_d_type == "W2" and self._is_goimok_ansung5_first(row):
                    continue
                camp = str(row.get("캠프명", "")).strip()
                car_num = str(row.get("차량번호", "")).strip()
                if not (camp or car_num):
                    continue
                collected_rows.append((source_type, camp, car_num))

        if target_d_type == "SD2":
            for row in self.dispatch_source_rows.get("간선", []):
                if self._is_goimok_ansung5_first(row):
                    camp = str(row.get("캠프명", "")).strip()
                    car_num = str(row.get("차량번호", "")).strip()
                    if camp or car_num:
                        collected_rows.append(("간선", camp, car_num))
                    break

        return collected_rows

    def print_goimok(self, d_type: str) -> None:
        if not hasattr(self, "dispatch_source_rows"):
            return

        data_rows = self._collect_goimok_rows(d_type)
        if not data_rows:
            messagebox.showinfo("안내", f"[{d_type}] 탭에 기입된 데이터가 없습니다.")
            return

        goimok_path = self._get_goimok_template_path()
        if not goimok_path:
            return
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("오류", "openpyxl 라이브러리가 필요합니다.")
            return

        temp_fd, temp_path = tempfile.mkstemp(prefix="goimok_", suffix=".xlsx")
        os.close(temp_fd)
        shutil.copy2(goimok_path, temp_path)

        try:
            wb = load_workbook(temp_path)
            ws = wb.active

            for idx, (row_type, camp, car_num) in enumerate(data_rows):
                row_num = 2 + idx
                cell_a = ws.cell(row=row_num, column=1)
                cell_b = ws.cell(row=row_num, column=2)
                cell_c = ws.cell(row=row_num, column=3)
                font_a = copy(cell_a.font) if cell_a.font else None
                font_b = copy(cell_b.font) if cell_b.font else None
                font_c = copy(cell_c.font) if cell_c.font else None
                cell_a.value, cell_b.value, cell_c.value = row_type, camp, car_num
                if font_a:
                    cell_a.font = font_a
                if font_b:
                    cell_b.font = font_b
                if font_c:
                    cell_c.font = font_c

            wb.save(temp_path)
            wb.close()

            printed = False
            if platform.system() == "Windows" and WIN32COM_AVAILABLE:
                try:
                    pythoncom.CoInitialize()
                    xl = win32.DispatchEx("Excel.Application")
                    xl.Visible = False
                    xl.DisplayAlerts = False
                    wb2 = xl.Workbooks.Open(os.path.abspath(temp_path))
                    wb2.ActiveSheet.PrintOut()
                    wb2.Close(SaveChanges=False)
                    xl.Quit()
                    printed = True
                except Exception as e:
                    print(f"[디버그] 고임목 Excel COM 인쇄 실패: {e}")
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

            if not printed:
                try:
                    pdf_path = self.excel_manager.convert_xlsx_to_pdf_lo(temp_path)
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup([temp_path, pdf_path, os.path.dirname(pdf_path)])
                    messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
                    return
                except Exception as e:
                    print(f"[출력 디버그] 고임목 LibreOffice PDF 변환 실패: {e}")

                if OPENPYXL_AVAILABLE and PILLOW_AVAILABLE:
                    img = self._render_dispatch_sheet_image(temp_path, ws.title)
                    pdf_path = self.excel_manager.save_image_as_pdf(img)
                    self.excel_manager.send_pdf_to_printer(pdf_path)
                    self._schedule_temp_cleanup([temp_path, pdf_path])
                    messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
                    return

                messagebox.showerror("오류", "인쇄 가능한 방법을 찾지 못했습니다.\nwin32com, LibreOffice, 또는 Pillow 중 하나가 필요합니다.")
                return

            messagebox.showinfo("완료", f"[{d_type}] 고임목 {len(data_rows)}건 프린트 요청 완료.")
            self._schedule_temp_cleanup([temp_path])

        except Exception as e:
            self._safe_remove_paths(temp_path)
            messagebox.showerror("고임목 출력 오류", str(e))

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DeliveryApp(root)
        root.mainloop()
    except Exception as e:
        print(f"시스템 오류 발생: {e}")
