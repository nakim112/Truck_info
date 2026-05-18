# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템
- 커스텀 UI 위젯 (Widgets)
"""
import random
import tkinter as tk
from tkinter import font as tkfont

# ==========================================
# [Split Flap 전광판 셀 클래스]
# ==========================================
class SplitFlapCell(tk.Frame):
    def __init__(self, parent, font=("HY헤드라인M", 24), bg_color="#2b2622"):
        super().__init__(parent, bg="#1a1a1a", highlightbackground="#3a3a3a", highlightthickness=2)
        self.target_text = ""
        self.target_fg = "#ffffff"

        available_fonts = tkfont.families()
        if font[0] not in available_fonts:
            font = ("맑은 고딕", 18, "bold")

        self.label = tk.Label(self, text="", font=font, bg=bg_color, fg=self.target_fg, anchor="c")
        self.label.pack(padx=2, pady=2, fill="both", expand=True)
        self.chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ가나다라마바사아자차카타파하0123456789- :"
        self._anim_id = None
        self.use_anim = True

    def set_text(self, text: str, fg_color: str = "#ffffff") -> None:
        if self.target_text == text:
            self.set_color(fg_color)
            return
        self.target_text = text
        self.target_fg = fg_color
        self.force_animate()

    def set_color(self, fg_color: str) -> None:
        self.target_fg = fg_color
        self.label.config(fg=fg_color)

    def force_animate(self) -> None:
        if not str(self.target_text).strip():
            self.label.config(text="")
            return
        if self._anim_id:
            self.after_cancel(self._anim_id)
        if hasattr(self, 'use_anim') and not self.use_anim:
            self.label.config(text=self.target_text, fg=self.target_fg)
            return
        self._animate(0)

    def _animate(self, step: int) -> None:
        if step < 6:
            rand_str = "".join(random.choice(self.chars) for _ in range(len(self.target_text)))
            self.label.config(text=rand_str, fg="#aaaaaa")
            self._anim_id = self.after(50, self._animate, step + 1)
        else:
            self.label.config(text=self.target_text, fg=self.target_fg)
            self._anim_id = None

    def destroy(self) -> None:
        if self._anim_id:
            self.after_cancel(self._anim_id)
        super().destroy()