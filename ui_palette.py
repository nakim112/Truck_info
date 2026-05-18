import json
import sys
from pathlib import Path
from typing import Dict


LEGACY_PRESET_NAME = "legacy_default"
WIN11_PRESET_NAME = "windows11_default"
MONO_PRESET_NAME = "mono_mode"
DARK_PRESET_NAME = "dark_mode"
DEFAULT_PRESET_NAME = LEGACY_PRESET_NAME
FALLBACK_PRESET_NAME = LEGACY_PRESET_NAME


def _resolve_active_preset_file() -> Path:
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent
    return base_dir / "ui_palette_state.json"


def _resolve_settings_file() -> Path:
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent
    return base_dir / "camp_settings.json"


ACTIVE_PRESET_FILE = _resolve_active_preset_file()
SETTINGS_FILE = _resolve_settings_file()

PRESET_LABELS: Dict[str, str] = {
    LEGACY_PRESET_NAME: "원래 색상",
    MONO_PRESET_NAME: "흑백모드",
    WIN11_PRESET_NAME: "Windows 11",
    DARK_PRESET_NAME: "다크모드",
}

PALETTE_PRESETS: Dict[str, Dict[str, str]] = {
    LEGACY_PRESET_NAME: {
        "primary_100": "#15344D",
        "primary_200": "#2F6F8F",
        "primary_300": "#DCEAF3",
        "accent_100": "#5D91AE",
        "accent_200": "#5E8FAF",
        "text_100": "#132838",
        "text_200": "#506474",
        "bg_100": "#FBFDFF",
        "bg_200": "#E7EEF4",
        "bg_300": "#C8D8E3",
    },
    MONO_PRESET_NAME: {
        "primary_100": "#2A2E33",
        "primary_200": "#3B4148",
        "primary_300": "#E8EEF3",
        "accent_100": "#66707A",
        "accent_200": "#4D5966",
        "text_100": "#1E2022",
        "text_200": "#495864",
        "bg_100": "#F2F6F9",
        "bg_200": "#D0DAE3",
        "bg_300": "#B8C3CE",
    },
    WIN11_PRESET_NAME: {
        "primary_100": "#0078D4",
        "primary_200": "#005A9E",
        "primary_300": "#C7E0F4",
        "accent_100": "#0078D4",
        "accent_200": "#005A9E",
        "text_100": "#1F1F1F",
        "text_200": "#616161",
        "bg_100": "#FFFFFF",
        "bg_200": "#F3F3F3",
        "bg_300": "#D9D9D9",
    },
    DARK_PRESET_NAME: {
        "primary_100": "#3B4A63",
        "primary_200": "#4E607C",
        "primary_300": "#2A3950",
        "accent_100": "#5E718E",
        "accent_200": "#7186A6",
        "text_100": "#FFFFFF",
        "text_200": "#B8C3D1",
        "bg_100": "#1A2433",
        "bg_200": "#0F1724",
        "bg_300": "#2B3648",
    },
}


def get_active_preset_name() -> str:
    try:
        settings_payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        preset_name = settings_payload.get("ui_mode", FALLBACK_PRESET_NAME)
        if preset_name in PALETTE_PRESETS:
            return preset_name
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass

    try:
        payload = json.loads(ACTIVE_PRESET_FILE.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return FALLBACK_PRESET_NAME
    except json.JSONDecodeError:
        return FALLBACK_PRESET_NAME

    preset_name = payload.get("active_preset", FALLBACK_PRESET_NAME)
    return preset_name if preset_name in PALETTE_PRESETS else FALLBACK_PRESET_NAME


def set_active_preset_name(preset_name: str) -> None:
    if preset_name not in PALETTE_PRESETS:
        available = ", ".join(sorted(PALETTE_PRESETS))
        raise ValueError(f"Unknown preset '{preset_name}'. Available presets: {available}")
    try:
        settings_payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        settings_payload = {}
    settings_payload["ui_mode"] = preset_name
    SETTINGS_FILE.write_text(
        json.dumps(settings_payload, ensure_ascii=False, indent=4) + "\n",
        encoding="utf-8",
    )


def get_palette_tokens(preset_name: str | None = None) -> Dict[str, str]:
    selected = preset_name or get_active_preset_name()
    if selected not in PALETTE_PRESETS:
        selected = FALLBACK_PRESET_NAME
    return dict(PALETTE_PRESETS[selected])


def build_main_palette(preset_name: str | None = None) -> Dict[str, str]:
    selected = preset_name or get_active_preset_name()
    tokens = get_palette_tokens(selected)
    if selected == LEGACY_PRESET_NAME:
        return {
            "bg": "#E7EEF4",
            "white": "#FBFDFF",
            "panel": "#FBFDFF",
            "toolbar": "#F6FAFD",
            "header": "#15344D",
            "header_txt": "#F7FBFF",
            "header_muted": "#AFC3D4",
            "accent": "#5D91AE",
            "accent_soft": "#D9E8F1",
            "border": "#C8D8E3",
            "border_soft": "#D9E5ED",
            "border_focus": "#5E8FAF",
            "label": "#506474",
            "text": "#132838",
            "placeholder": "#7E95A7",
            "btn_primary": "#2F6F8F",
            "btn_success": "#2D7C72",
            "btn_danger": "#BE5C57",
            "btn_warning": "#C69754",
            "btn_neutral": "#758796",
            "btn_dark": "#244B67",
            "btn_teal": "#3D8DA5",
            "btn_secondary": "#31506C",
            "status_ok": "#2B7E73",
            "status_warn": "#9C762E",
            "status_err": "#B65A55",
            "row_odd": "#FBFDFF",
            "row_even": "#F3F8FC",
            "row_sel": "#DCEAF3",
            "heading": "#EAF2F7",
            "heading_active": "#DDE8F0",
        }
    return {
        "bg": tokens["bg_200"],
        "white": tokens["bg_100"],
        "panel": tokens["bg_100"],
        "toolbar": tokens["bg_100"],
        "header": tokens["primary_100"],
        "header_txt": tokens["bg_300"],
        "header_muted": tokens["primary_300"],
        "accent": tokens["accent_100"],
        "accent_soft": tokens["primary_300"],
        "border": tokens["bg_300"],
        "border_soft": tokens["bg_300"],
        "border_focus": tokens["accent_200"],
        "label": tokens["text_200"],
        "text": tokens["text_100"],
        "placeholder": tokens["text_200"],
        "btn_primary": tokens["primary_100"],
        "btn_success": tokens["primary_200"],
        "btn_danger": tokens["accent_200"],
        "btn_warning": tokens["accent_100"],
        "btn_neutral": tokens["text_200"],
        "btn_dark": "#244B67",
        "btn_teal": tokens["accent_100"],
        "btn_secondary": tokens["accent_200"],
        "status_ok": tokens["primary_100"],
        "status_warn": tokens["accent_100"],
        "status_err": tokens["accent_200"],
        "row_odd": tokens["bg_100"],
        "row_even": tokens["bg_200"],
        "row_sel": tokens["primary_300"],
        "heading": tokens["bg_200"],
        "heading_active": tokens["primary_300"],
    }
