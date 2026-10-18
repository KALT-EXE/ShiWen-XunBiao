"""识文寻表 —— 框选屏幕文字，在 Excel 中自动查找。"""

from __future__ import annotations

import argparse
import ctypes
import json
import locale
import logging
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageTk

import excel_index as core
from ocr_service import OcrError, OcrResult, OcrService
from win_support import (
    GlobalHotkey,
    SelectionOverlay,
    capture_screen,
    enable_dpi_awareness,
    set_app_user_model_id,
    virtual_screen_rect,
)

APP_NAME = "识文寻表"
APP_VERSION = "1.1.0"
APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
CODE_DIR = Path(__file__).resolve().parent

# 颜色
BG = "#F3F6F8"
CARD = "#FFFFFF"
INK = "#0F172A"
INK_SOFT = "#334155"
MUTED = "#64748B"
MUTED_LIGHT = "#94A3B8"
BORDER = "#E2E8F0"
BORDER_SOFT = "#EEF2F6"
NAVY = "#0B1220"
NAVY_SOFT = "#111C2E"
NAVY_CARD = "#172033"
NAVY_LINE = "#263449"
TEAL = "#14B8A6"
TEAL_DARK = "#0F766E"
TEAL_LIGHT = "#CCFBF1"
TEAL_PALE = "#F0FDFA"
AMBER = "#F59E0B"
RED = "#EF4444"
GREEN = "#10B981"

# 字体
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_SMALL = ("Microsoft YaHei UI", 9)
FONT_TINY = ("Microsoft YaHei UI", 8)
FONT_BODY = ("Microsoft YaHei UI", 10)
FONT_TITLE = ("Microsoft YaHei UI", 23, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 12, "bold")
FONT_BRAND = ("Microsoft YaHei UI", 18, "bold")
FONT_MONO = ("Cascadia Mono", 9)

MODE_LABELS = ["智能匹配", "精确匹配", "包含匹配"]
MODE_MAP = {"智能匹配": "smart", "精确匹配": "exact", "包含匹配": "contains"}

DEFAULT_HOTKEY_MODIFIERS = 0x0002 | 0x0004  # Ctrl + Shift
DEFAULT_HOTKEY_VK = 0x51  # Q
DEFAULT_HOTKEY_LABEL = "Ctrl + Shift + Q"
HOTKEY_MOD_MASK = 0x0001 | 0x0002 | 0x0004 | 0x0008

_KEY_NAMES = {
    0x08: "Backspace",
    0x09: "Tab",
    0x0D: "Enter",
    0x13: "Pause",
    0x14: "CapsLock",
    0x20: "Space",
    0x21: "PageUp",
    0x22: "PageDown",
    0x23: "End",
    0x24: "Home",
    0x25: "Left",
    0x26: "Up",
    0x27: "Right",
    0x28: "Down",
    0x2D: "Insert",
    0x2E: "Delete",
    0x5B: "Win",
    0x5C: "Win",
    0x90: "NumLock",
    0x91: "ScrollLock",
}


def hotkey_label(modifiers: int, vk: int, key_name: str | None = None) -> str:
    parts: list[str] = []
    if modifiers & 0x0002:
        parts.append("Ctrl")
    if modifiers & 0x0001:
        parts.append("Alt")
    if modifiers & 0x0004:
        parts.append("Shift")
    if modifiers & 0x0008:
        parts.append("Win")
    name = key_name or _key_name_from_vk(vk)
    parts.append(name)
    return " + ".join(parts)


def _key_name_from_vk(vk: int) -> str:
    if 0x41 <= vk <= 0x5A:
        return chr(vk)
    if 0x30 <= vk <= 0x39:
        return chr(vk)
    if 0x70 <= vk <= 0x87:
        return f"F{vk - 0x6F}"
    if 0x60 <= vk <= 0x69:
        return f"Num{vk - 0x60}"
    return _KEY_NAMES.get(vk, f"VK{vk:02X}")


# ---------------------------------------------------------------------------
# 资源 / 配置 / 日志
# ---------------------------------------------------------------------------


def resource_path(*parts: str) -> Path:
    return APP_DIR.joinpath(*parts)


def _config_root() -> Path:
    candidates = []
    appdata = os.environ.get("APPDATA")
    if appdata:
        candidates.append(Path(appdata) / "ShiWenXunBiao")
    candidates.append(CODE_DIR / ".config")
    candidates.append(Path.home() / ".shiwen_xunbiao")
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / ".write_test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return candidate
        except OSError:
            continue
    return CODE_DIR / ".config"


CONFIG_DIR = _config_root()
CONFIG_PATH = CONFIG_DIR / "config.json"
LOG_PATH = CONFIG_DIR / "app.log"


def setup_logging() -> None:
    try:
        logging.basicConfig(
            filename=str(LOG_PATH),
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            encoding="utf-8",
        )
    except OSError:
        logging.basicConfig(level=logging.INFO)


def load_config() -> dict:
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_config(data: dict) -> None:
    try:
        CONFIG_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        logging.exception("保存配置失败")


# ---------------------------------------------------------------------------
# 小部件
# ---------------------------------------------------------------------------


class PlaceholderEntry(tk.Entry):
    def __init__(self, master, placeholder: str, **kwargs):
        super().__init__(
            master,
            bd=0,
            relief="flat",
            bg=CARD,
            fg=MUTED,
            insertbackground=INK,
            font=("Microsoft YaHei UI", 11),
            **kwargs,
        )
        self.placeholder = placeholder
        self._placeholder_active = True
        self._show_placeholder()
        self.bind("<FocusIn>", self._on_focus_in, add="+")
        self.bind("<FocusOut>", self._on_focus_out, add="+")

    def _show_placeholder(self) -> None:
        self._placeholder_active = True
        self.delete(0, "end")
        self.insert(0, self.placeholder)
        self.configure(fg=MUTED_LIGHT)

    def _on_focus_in(self, _event=None) -> None:
        if self._placeholder_active:
            self._placeholder_active = False
            self.delete(0, "end")
            self.configure(fg=INK)

    def _on_focus_out(self, _event=None) -> None:
        if not self.get().strip():
            self._show_placeholder()

    def get_value(self) -> str:
        if self._placeholder_active:
            return ""
        return self.get().strip()

    def set_value(self, text: str) -> None:
        self._placeholder_active = False
        self.delete(0, "end")
        self.insert(0, text)
        self.configure(fg=INK)
        if not text:
            self._on_focus_out()

    def prepare_for_input(self) -> None:
        self.focus_set()
        self._on_focus_in()
        self.icursor("end")


def _bind_hover(widget: tk.Widget, normal: str, hover: str) -> None:
    widget.bind("<Enter>", lambda _e: widget.configure(bg=hover), add="+")
    widget.bind("<Leave>", lambda _e: widget.configure(bg=normal), add="+")


def make_button(
    parent,
    text: str,
    command,
    kind: str = "primary",
    font=FONT_SMALL,
    padx: int = 14,
    pady: int = 7,
    width: int | None = None,
) -> tk.Button:
    colors = {
        "primary": (TEAL, "#2DD4BF", "#052E2B"),
        "dark": (NAVY_CARD, "#223047", "#F8FAFC"),
        "ghost": (CARD, "#F1F5F9", INK_SOFT),
        "soft": (TEAL_PALE, TEAL_LIGHT, TEAL_DARK),
        "link": (CARD, CARD, TEAL_DARK),
    }
    normal, hover, fg = colors.get(kind, colors["ghost"])
    options = dict(
        text=text,
        command=command,
        bg=normal,
        fg=fg,
        activebackground=hover,
        activeforeground=fg,
        relief="flat",
        bd=0,
        cursor="hand2",
        font=font,
        padx=padx,
        pady=pady,
        highlightthickness=1 if kind == "ghost" else 0,
        highlightbackground=BORDER,
        highlightcolor=BORDER,
    )
    if width is not None:
        options["width"] = width
    button = tk.Button(parent, **options)
    _bind_hover(button, normal, hover)
    return button


def card_frame(parent, bg=CARD, border=BORDER) -> tk.Frame:
    return tk.Frame(
        parent,
        bg=bg,
        highlightbackground=border,
        highlightthickness=1,
        bd=0,
    )


def _truncate(text: str, limit: int) -> str:
    text = " ".join(str(text).split())
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


# ---------------------------------------------------------------------------
# 主应用
# ---------------------------------------------------------------------------


class ShiWenApp:
    def __init__(self, root: tk.Tk, initial_file: Path | None = None):
        self.root = root
        self.config = load_config()
        self.initial_file = initial_file
        self.workbook: core.WorkbookData | None = None
        self.excel_path: Path | None = None
        self.last_query = ""
        self.last_result: core.SearchResult | None = None
        self.ocr_lines: list[str] = []
        self.preview_photo = None
        self._row_hits: list[core.SearchHit] = []
        self._busy: str | None = None
        self._overlay: SelectionOverlay | None = None
        self._capturing = False
        self._pending_query = ""
        self._closing = False
        self._ui_queue: queue.Queue = queue.Queue()
        self._recent_queries = [
            str(item)
            for item in self.config.get("recent_queries", [])
            if str(item).strip()
        ][:5]
        self._mode_label = str(self.config.get("mode", "智能匹配"))
        if self._mode_label not in MODE_LABELS:
            self._mode_label = "智能匹配"
        self._icon_photo = None
        self._help_window = None
        self._hotkey_dialog = None
        self._topmost_enabled = bool(self.config.get("topmost", True))
        try:
            configured_mods = int(
                self.config.get("hotkey_modifiers", DEFAULT_HOTKEY_MODIFIERS)
            )
        except (TypeError, ValueError):
            configured_mods = DEFAULT_HOTKEY_MODIFIERS
        try:
            configured_vk = int(self.config.get("hotkey_vk", DEFAULT_HOTKEY_VK))
        except (TypeError, ValueError):
            configured_vk = DEFAULT_HOTKEY_VK
        configured_mods &= HOTKEY_MOD_MASK
        if not 1 <= configured_vk <= 255:
            configured_mods = DEFAULT_HOTKEY_MODIFIERS
            configured_vk = DEFAULT_HOTKEY_VK
        configured_label = str(self.config.get("hotkey_label", "")).strip()
        if not configured_label:
            configured_label = hotkey_label(configured_mods, configured_vk)
        self._hotkey_modifiers = configured_mods
        self._hotkey_vk = configured_vk
        self._hotkey_label = configured_label
        self._shown_hotkey_label = configured_label

        self.ocr_service = OcrService(resource_path("assets", "ocr_windows.ps1"))
        self.hotkey = GlobalHotkey(
            configured_mods, configured_vk, configured_label
        )

        self.root.title(f"{APP_NAME} · 框选识字，自动查表")
        self.root.configure(bg=BG)
        self.root.minsize(1100, 700)
        self.root.option_add("*Font", FONT_UI)
        self._set_window_geometry()
        self._set_window_icon()
        try:
            self.root.attributes("-topmost", self._topmost_enabled)
        except tk.TclError:
            pass
        self._configure_styles()
        self._build_ui()
        self._bind_shortcuts()
        self._start_hotkey()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.report_callback_exception = self._report_callback_exception
        self._poll_queue()
        self._poll_hotkey()
        self.root.after(260, self._startup_load)

    # ------------------------------------------------------------------
    # 初始化
    # ------------------------------------------------------------------

    def _set_window_geometry(self) -> None:
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(1360, max(1060, screen_w - 180))
        height = min(860, max(700, screen_h - 120))
        saved = self.config.get("geometry", "")
        m = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)([+-]\d+)", str(saved))
        if m:
            w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            if 980 <= w <= screen_w and 640 <= h <= screen_h:
                if -w + 100 < x < screen_w - 100 and -h + 80 < y < screen_h - 80:
                    self.root.geometry(f"{w}x{h}{x:+d}{y:+d}")
                    return
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _set_window_icon(self) -> None:
        try:
            icon_path = resource_path("assets", "app.ico")
            if icon_path.exists():
                self.root.iconbitmap(default=str(icon_path))
        except tk.TclError:
            pass
        try:
            image = self._draw_app_icon(64)
            self._icon_photo = ImageTk.PhotoImage(image, master=self.root)
            self.root.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    @staticmethod
    def _draw_app_icon(size: int = 64) -> Image.Image:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        margin = max(1, size // 16)
        draw.rounded_rectangle(
            (margin, margin, size - margin, size - margin),
            radius=size // 5,
            fill=TEAL,
        )
        dark = "#042F2E"
        pad = size // 5
        arm = size // 6
        width = max(2, size // 16)
        for points in (
            (pad, pad + arm, pad, pad, pad + arm, pad),
            (size - pad - arm, pad, size - pad, pad, size - pad, pad + arm),
            (pad, size - pad - arm, pad, size - pad, pad + arm, size - pad),
            (
                size - pad - arm,
                size - pad,
                size - pad,
                size - pad,
                size - pad,
                size - pad - arm,
            ),
        ):
            draw.line(points, fill=dark, width=width)
        for idx, y in enumerate((size // 2 - 5, size // 2 + 1)):
            draw.rounded_rectangle(
                (size // 2 - size // 7, y, size // 2 + size // 7, y + 2),
                radius=1,
                fill=dark,
            )
        return img

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Results.Treeview",
            background=CARD,
            fieldbackground=CARD,
            foreground=INK_SOFT,
            rowheight=35,
            borderwidth=0,
            font=FONT_SMALL,
        )
        style.configure(
            "Results.Treeview.Heading",
            background="#F8FAFC",
            foreground=MUTED,
            relief="flat",
            font=("Microsoft YaHei UI", 9, "bold"),
            padding=(8, 9),
        )
        style.map(
            "Results.Treeview.Heading",
            background=[("active", "#F1F5F9")],
        )
        style.map(
            "Results.Treeview",
            background=[("selected", TEAL_LIGHT)],
            foreground=[("selected", INK)],
        )
        style.configure(
            "App.TCombobox",
            fieldbackground=CARD,
            background=CARD,
            foreground=INK_SOFT,
            arrowcolor=MUTED,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=(8, 6),
            font=FONT_SMALL,
        )
        style.map(
            "App.TCombobox",
            fieldbackground=[("readonly", CARD)],
            selectbackground=[("readonly", CARD)],
            selectforeground=[("readonly", INK_SOFT)],
        )
        style.configure(
            "App.Vertical.TScrollbar",
            background="#CBD5E1",
            troughcolor=CARD,
            bordercolor=CARD,
            arrowcolor=MUTED,
            width=11,
        )
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#EDF2F5",
            background=TEAL,
            bordercolor="#EDF2F5",
            lightcolor=TEAL,
            darkcolor=TEAL,
            thickness=3,
        )
        self.root.option_add("*TCombobox*Listbox.font", FONT_SMALL)
        self.root.option_add("*TCombobox*Listbox.background", CARD)
        self.root.option_add("*TCombobox*Listbox.foreground", INK_SOFT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", TEAL_LIGHT)

    # ------------------------------------------------------------------
    # 界面
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=0, minsize=288)
        self.root.grid_columnconfigure(1, weight=1)
        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self) -> None:
        sidebar = tk.Frame(self.root, bg=NAVY, width=288)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.pack_propagate(False)
        self.sidebar = sidebar

        footer = tk.Frame(sidebar, bg=NAVY)
        footer.pack(side="bottom", fill="x", padx=20, pady=(0, 12))
        tk.Frame(footer, bg=NAVY_LINE, height=1).pack(fill="x", pady=(0, 10))
        bottom_row = tk.Frame(footer, bg=NAVY)
        bottom_row.pack(fill="x")
        privacy_dot = tk.Canvas(
            bottom_row, width=8, height=8, bg=NAVY, highlightthickness=0
        )
        privacy_dot.pack(side="left")
        privacy_dot.create_oval(1, 1, 7, 7, fill=GREEN, outline="")
        tk.Label(
            bottom_row,
            text="本地运行",
            bg=NAVY,
            fg="#64748B",
            font=FONT_TINY,
        ).pack(side="left", padx=(5, 0))
        tk.Label(
            bottom_row,
            text=f"v{APP_VERSION}",
            bg=NAVY,
            fg="#475569",
            font=("Segoe UI", 7),
        ).pack(side="right")

        body = tk.Frame(sidebar, bg=NAVY)
        body.pack(side="top", fill="both", expand=True)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.side_canvas = tk.Canvas(
            body,
            bg=NAVY,
            highlightthickness=0,
            bd=0,
        )
        self.side_canvas.grid(row=0, column=0, sticky="nsew")
        self.side_scroll = tk.Scrollbar(
            body,
            orient="vertical",
            command=self.side_canvas.yview,
            bg=NAVY,
            troughcolor=NAVY,
            activebackground=NAVY_LINE,
            highlightthickness=0,
            bd=0,
            relief="flat",
            width=8,
        )
        self.side_canvas.configure(yscrollcommand=self.side_scroll.set)
        inner = tk.Frame(self.side_canvas, bg=NAVY)
        self.side_inner = inner
        self._side_window = self.side_canvas.create_window(
            (20, 18), window=inner, anchor="nw", width=248
        )
        inner.bind("<Configure>", lambda _e: self._sync_sidebar_scroll())
        self.side_canvas.bind("<Configure>", lambda _e: self._sync_sidebar_scroll())

        brand = tk.Frame(inner, bg=NAVY)
        brand.pack(fill="x")
        logo = tk.Canvas(brand, width=44, height=44, bg=NAVY, highlightthickness=0)
        logo.pack(side="left", padx=(0, 11))
        self._draw_logo(logo)
        brand_text = tk.Frame(brand, bg=NAVY)
        brand_text.pack(side="left", fill="x", expand=True)
        tk.Label(
            brand_text,
            text=APP_NAME,
            bg=NAVY,
            fg="#F8FAFC",
            font=FONT_BRAND,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand_text,
            text="SCAN  ·  MATCH  ·  LOCATE",
            bg=NAVY,
            fg="#5EEAD4",
            font=("Segoe UI", 7, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(1, 0))

        tk.Frame(inner, bg=NAVY_LINE, height=1).pack(fill="x", pady=(13, 12))

        self._sidebar_step(inner, "01", "数据表格")
        file_card = tk.Frame(
            inner,
            bg=NAVY_SOFT,
            highlightbackground=NAVY_LINE,
            highlightthickness=1,
            bd=0,
        )
        file_card.pack(fill="x", pady=(7, 0))
        file_inner = tk.Frame(file_card, bg=NAVY_SOFT)
        file_inner.pack(fill="x", padx=12, pady=10)

        top_row = tk.Frame(file_inner, bg=NAVY_SOFT)
        top_row.pack(fill="x")
        tk.Label(
            top_row,
            text="当前工作簿",
            bg=NAVY_SOFT,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
            anchor="w",
        ).pack(side="left")
        self.file_badge = tk.Label(
            top_row,
            text="XLSX",
            bg="#134E4A",
            fg="#99F6E4",
            font=("Segoe UI", 7, "bold"),
            padx=6,
            pady=2,
        )
        self.file_badge.pack(side="right")

        self.file_name_var = tk.StringVar(value="尚未选择表格")
        self.file_name_label = tk.Label(
            file_inner,
            textvariable=self.file_name_var,
            bg=NAVY_SOFT,
            fg="#F1F5F9",
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
            justify="left",
            wraplength=220,
        )
        self.file_name_label.pack(fill="x", pady=(7, 2))
        self.file_stats_var = tk.StringVar(value="选择 .xlsx / .xlsm / .csv")
        tk.Label(
            file_inner,
            textvariable=self.file_stats_var,
            bg=NAVY_SOFT,
            fg="#94A3B8",
            font=FONT_TINY,
            anchor="w",
            justify="left",
            wraplength=220,
        ).pack(fill="x")
        file_buttons = tk.Frame(file_inner, bg=NAVY_SOFT)
        file_buttons.pack(fill="x", pady=(9, 0))
        choose_btn = tk.Button(
            file_buttons,
            text="选择表格",
            command=self.choose_file,
            bg="#1E293B",
            fg="#E2E8F0",
            activebackground="#273449",
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
            padx=10,
            pady=5,
        )
        _bind_hover(choose_btn, "#1E293B", "#273449")
        choose_btn.pack(side="left")
        reload_btn = tk.Button(
            file_buttons,
            text="重新载入",
            command=self.reload_file,
            bg=NAVY_SOFT,
            fg="#94A3B8",
            activebackground=NAVY_SOFT,
            activeforeground="#F8FAFC",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
            padx=4,
            pady=5,
        )
        reload_btn.pack(side="right")

        self._sidebar_step(inner, "02", "框选识别", pady=(13, 0))
        self.scan_icon_photo = self._make_scan_icon()
        self.scan_button = tk.Button(
            inner,
            text="  框选屏幕文字",
            image=self.scan_icon_photo,
            compound="left",
            command=self.start_capture,
            bg=TEAL,
            fg="#042F2E",
            activebackground="#2DD4BF",
            activeforeground="#042F2E",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Microsoft YaHei UI", 11, "bold"),
            pady=11,
        )
        _bind_hover(self.scan_button, TEAL, "#2DD4BF")
        self.scan_button.pack(fill="x", pady=(8, 7))

        hotkey_row = tk.Frame(inner, bg=NAVY)
        hotkey_row.pack(fill="x")
        self.hotkey_dot = tk.Canvas(
            hotkey_row, width=8, height=8, bg=NAVY, highlightthickness=0
        )
        self.hotkey_dot.pack(side="left", pady=(2, 0))
        self.hotkey_dot.create_oval(1, 1, 7, 7, fill=MUTED, outline="")
        tk.Label(
            hotkey_row,
            text="快捷键",
            bg=NAVY,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="left", padx=(6, 7))
        self.hotkey_label_var = tk.StringVar(value=self.hotkey.label)
        self.hotkey_chip = tk.Button(
            hotkey_row,
            textvariable=self.hotkey_label_var,
            command=self.show_hotkey_dialog,
            bg=NAVY_CARD,
            fg="#CBD5E1",
            activebackground=NAVY_LINE,
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI", 8),
            padx=7,
            pady=2,
        )
        self.hotkey_chip.pack(side="left")
        self.hotkey_change_button = tk.Button(
            hotkey_row,
            text="更改",
            command=self.show_hotkey_dialog,
            bg=NAVY,
            fg=TEAL,
            activebackground=NAVY,
            activeforeground="#5EEAD4",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
        )
        self.hotkey_change_button.pack(side="right")

        tk.Frame(inner, bg=NAVY_LINE, height=1).pack(fill="x", pady=(12, 10))
        self._sidebar_info_row(inner, "识别语言", "简体中文 · 英文")
        self._sidebar_info_row(inner, "查找范围", "全部工作表")
        self._sidebar_info_row(inner, "数据安全", "本地处理 · 不上传", value_color="#86EFAC")

        recent_header = tk.Frame(inner, bg=NAVY)
        recent_header.pack(fill="x", pady=(12, 5))
        tk.Label(
            recent_header,
            text="最近查找",
            bg=NAVY,
            fg="#E2E8F0",
            font=("Microsoft YaHei UI", 9, "bold"),
            anchor="w",
        ).pack(side="left")
        clear_recent = tk.Button(
            recent_header,
            text="清空",
            command=self.clear_recent,
            bg=NAVY,
            fg="#64748B",
            activebackground=NAVY,
            activeforeground="#94A3B8",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
        )
        clear_recent.pack(side="right")
        self.recent_frame = tk.Frame(inner, bg=NAVY)
        self.recent_frame.pack(fill="x")

        self._rebuild_recent()
        self._bind_sidebar_wheel(self.side_canvas)
        self.root.after(80, self._sync_sidebar_scroll)

    def _sync_sidebar_scroll(self) -> None:
        try:
            content_height = self.side_inner.winfo_reqheight() + 30
            view_height = self.side_canvas.winfo_height()
            self.side_canvas.configure(scrollregion=(0, 0, 288, content_height))
            if content_height > view_height + 2:
                if not self.side_scroll.winfo_ismapped():
                    self.side_scroll.grid(row=0, column=1, sticky="ns")
            else:
                self.side_scroll.grid_remove()
                self.side_canvas.yview_moveto(0)
        except tk.TclError:
            pass

    def _on_sidebar_wheel(self, event) -> str:
        try:
            if not self.side_scroll.winfo_ismapped():
                return "break"
            step = -1 if event.delta > 0 else 1
            self.side_canvas.yview_scroll(step * 3, "units")
        except tk.TclError:
            pass
        return "break"

    def _bind_sidebar_wheel(self, widget) -> None:
        try:
            widget.bind("<MouseWheel>", self._on_sidebar_wheel)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self._bind_sidebar_wheel(child)

    def _make_scan_icon(self, size: int = 20):
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        dark = "#042F2E"
        pad = 2
        arm = 6
        width = 2
        for points in (
            (pad, pad + arm, pad, pad, pad + arm, pad),
            (size - pad - arm, pad, size - pad, pad, size - pad, pad + arm),
            (pad, size - pad - arm, pad, size - pad, pad + arm, size - pad),
            (
                size - pad - arm,
                size - pad,
                size - pad,
                size - pad,
                size - pad,
                size - pad - arm,
            ),
        ):
            draw.line(points, fill=dark, width=width)
        draw.line(
            (size // 2 - 4, size // 2, size // 2 + 4, size // 2),
            fill=dark,
            width=2,
        )
        return ImageTk.PhotoImage(img, master=self.root)

    def _draw_logo(self, canvas: tk.Canvas) -> None:
        canvas.create_rectangle(0, 0, 44, 44, fill=TEAL, outline="")
        dark = "#042F2E"
        for points in (
            (9, 18, 9, 9, 18, 9),
            (26, 9, 35, 9, 35, 18),
            (9, 26, 9, 35, 18, 35),
            (26, 35, 35, 35, 35, 26),
        ):
            canvas.create_line(*points, fill=dark, width=3)
        canvas.create_line(17, 22, 27, 22, fill=dark, width=2)
        canvas.create_line(17, 26, 24, 26, fill=dark, width=2)

    def _sidebar_step(self, parent, number: str, title: str, pady=(0, 0)) -> None:
        row = tk.Frame(parent, bg=NAVY)
        row.pack(fill="x", pady=pady)
        tk.Label(
            row,
            text=number,
            bg=NAVY,
            fg=TEAL,
            font=("Segoe UI", 8, "bold"),
            anchor="w",
        ).pack(side="left")
        tk.Label(
            row,
            text=title,
            bg=NAVY,
            fg="#E2E8F0",
            font=("Microsoft YaHei UI", 9, "bold"),
            anchor="w",
        ).pack(side="left", padx=(7, 0))

    def _sidebar_info_row(
        self,
        parent,
        label: str,
        value: str,
        value_color: str = "#CBD5E1",
    ) -> None:
        row = tk.Frame(parent, bg=NAVY)
        row.pack(fill="x", pady=2)
        tk.Label(
            row,
            text=label,
            bg=NAVY,
            fg="#64748B",
            font=FONT_TINY,
            anchor="w",
        ).pack(side="left")
        tk.Label(
            row,
            text=value,
            bg=NAVY,
            fg=value_color,
            font=FONT_TINY,
            anchor="e",
        ).pack(side="right")

    def _build_main(self) -> None:
        main = tk.Frame(self.root, bg=BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_rowconfigure(1, weight=1)
        main.grid_columnconfigure(0, weight=1)

        header = tk.Frame(main, bg=BG)
        header.grid(row=0, column=0, sticky="ew", padx=30, pady=(22, 15))
        header.grid_columnconfigure(0, weight=1)
        header_left = tk.Frame(header, bg=BG)
        header_left.grid(row=0, column=0, sticky="w")
        tk.Label(
            header_left,
            text="框选文字，自动查表",
            bg=BG,
            fg=INK,
            font=FONT_TITLE,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header_left,
            text="框选屏幕文字 · Excel 全表智能查找",
            bg=BG,
            fg=MUTED,
            font=("Microsoft YaHei UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(4, 0))

        header_right = tk.Frame(header, bg=BG)
        header_right.grid(row=0, column=1, sticky="e")
        status_pill = tk.Frame(
            header_right,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1,
            bd=0,
        )
        status_pill.pack(side="left", padx=(0, 9), ipady=2)
        self.status_dot = tk.Canvas(
            status_pill, width=9, height=9, bg=CARD, highlightthickness=0
        )
        self.status_dot.pack(side="left", padx=(10, 0))
        self.status_dot_item = self.status_dot.create_oval(
            1, 1, 8, 8, fill=GREEN, outline=""
        )
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(
            status_pill,
            textvariable=self.status_var,
            bg=CARD,
            fg=INK_SOFT,
            font=("Microsoft YaHei UI", 8),
            padx=6,
            pady=4,
            anchor="w",
        )
        self.status_label.pack(side="left")
        self.topmost_button = make_button(
            header_right,
            "置顶：开" if self._topmost_enabled else "置顶：关",
            self.toggle_topmost,
            kind="ghost",
            font=FONT_TINY,
            padx=10,
            pady=8,
        )
        self.topmost_button.configure(
            fg=TEAL_DARK if self._topmost_enabled else MUTED
        )
        self.topmost_button.pack(side="left", padx=(0, 7))
        help_btn = make_button(
            header_right,
            "使用说明",
            self.show_help,
            kind="ghost",
            font=FONT_TINY,
            padx=11,
            pady=8,
        )
        help_btn.pack(side="left")

        content = tk.Frame(main, bg=BG)
        content.grid(row=1, column=0, sticky="nsew", padx=30, pady=(0, 20))
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)
        self.content = content
        self._build_ocr_card(content)
        self._build_results_card(content)

    def _build_ocr_card(self, parent) -> None:
        card = card_frame(parent)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        card.grid_columnconfigure(0, weight=0, minsize=238)
        card.grid_columnconfigure(1, weight=1)

        preview_box = tk.Frame(
            card,
            bg="#F8FAFC",
            width=238,
            height=212,
            highlightbackground=BORDER,
            highlightthickness=1,
            bd=0,
        )
        preview_box.grid(row=0, column=0, padx=16, pady=16, sticky="nw")
        preview_box.grid_propagate(False)
        preview_box.pack_propagate(False)
        self.preview_box = preview_box
        self.preview_label = tk.Label(
            preview_box,
            text="框选后\n在此预览",
            bg="#F8FAFC",
            fg="#94A3B8",
            font=("Microsoft YaHei UI", 9),
            justify="center",
        )
        self.preview_label.pack(fill="both", expand=True)
        self.preview_badge = tk.Label(
            preview_box,
            text="截图预览",
            bg="#0F172A",
            fg="#E2E8F0",
            font=("Microsoft YaHei UI", 7),
            padx=6,
            pady=2,
        )
        self.preview_badge.place(relx=1.0, rely=1.0, x=-1, y=-1, anchor="se")

        detail = tk.Frame(card, bg=CARD)
        detail.grid(row=0, column=1, sticky="nsew", padx=(0, 16), pady=14)
        detail.grid_columnconfigure(0, weight=1)
        detail.grid_rowconfigure(1, weight=1)

        text_header = tk.Frame(detail, bg=CARD)
        text_header.grid(row=0, column=0, sticky="ew")
        tk.Label(
            text_header,
            text="识别文字",
            bg=CARD,
            fg=INK,
            font=("Microsoft YaHei UI", 10, "bold"),
        ).pack(side="left")
        self.ocr_meta_var = tk.StringVar(value="Windows OCR · 中英文")
        tk.Label(
            text_header,
            textvariable=self.ocr_meta_var,
            bg=CARD,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="right")

        text_border = tk.Frame(detail, bg=BORDER)
        text_border.grid(row=1, column=0, sticky="nsew", pady=(6, 9))
        text_border.grid_rowconfigure(0, weight=1)
        text_border.grid_columnconfigure(0, weight=1)
        text_inner = tk.Frame(text_border, bg="#F8FAFC")
        text_inner.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        text_inner.grid_rowconfigure(0, weight=1)
        text_inner.grid_columnconfigure(0, weight=1)
        self.ocr_text = tk.Text(
            text_inner,
            height=3,
            width=1,
            wrap="word",
            bg="#F8FAFC",
            fg=MUTED_LIGHT,
            relief="flat",
            bd=0,
            padx=10,
            pady=7,
            font=("Microsoft YaHei UI", 10),
            insertbackground=INK,
            selectbackground=TEAL_LIGHT,
            selectforeground=INK,
            undo=True,
        )
        self.ocr_text.grid(row=0, column=0, sticky="nsew")
        text_scroll = ttk.Scrollbar(
            text_inner,
            orient="vertical",
            command=self.ocr_text.yview,
            style="App.Vertical.TScrollbar",
        )
        text_scroll.grid(row=0, column=1, sticky="ns")
        self.ocr_text.configure(yscrollcommand=text_scroll.set)
        self.ocr_text.insert(
            "1.0", "框选后识别结果会显示在这里，也可双击词句直接查找。"
        )
        self.ocr_text.bind("<Double-Button-1>", self._on_ocr_double_click, add="+")

        query_row = tk.Frame(detail, bg=CARD)
        query_row.grid(row=2, column=0, sticky="ew")
        query_row.grid_columnconfigure(1, weight=1)
        tk.Label(
            query_row,
            text="查找",
            bg=CARD,
            fg=MUTED,
            font=("Microsoft YaHei UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        entry_border = tk.Frame(
            query_row,
            bg=BORDER,
            highlightbackground=BORDER,
            highlightthickness=0,
        )
        entry_border.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.query_entry = PlaceholderEntry(
            entry_border, placeholder="输入关键词…"
        )
        self.query_entry.pack(fill="both", expand=True, padx=1, pady=1, ipady=4)
        self.query_entry.bind("<Return>", lambda _e: self.run_search())
        self.mode_var = tk.StringVar(value=self._mode_label)
        self.mode_combo = ttk.Combobox(
            query_row,
            textvariable=self.mode_var,
            values=MODE_LABELS,
            state="readonly",
            width=8,
            style="App.TCombobox",
            font=FONT_TINY,
        )
        self.mode_combo.grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_changed)
        self.search_button = make_button(
            query_row,
            "查找",
            self.run_search,
            kind="primary",
            font=("Microsoft YaHei UI", 9, "bold"),
            padx=15,
            pady=6,
        )
        self.search_button.grid(row=0, column=3, sticky="e")

        chips_row = tk.Frame(detail, bg=CARD)
        chips_row.grid(row=3, column=0, sticky="ew", pady=(7, 0))
        tk.Label(
            chips_row,
            text="快捷",
            bg=CARD,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="left")
        self.chips_frame = tk.Frame(chips_row, bg=CARD)
        self.chips_frame.pack(side="left", fill="x", expand=True, padx=(7, 0))
        self._set_chips([])

    def _build_results_card(self, parent) -> None:
        card = card_frame(parent)
        card.grid(row=1, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        header = tk.Frame(card, bg=CARD)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 10))
        header.grid_columnconfigure(1, weight=1)
        title_row = tk.Frame(header, bg=CARD)
        title_row.grid(row=0, column=0, sticky="w")
        tk.Label(
            title_row,
            text="查找结果",
            bg=CARD,
            fg=INK,
            font=("Microsoft YaHei UI", 12, "bold"),
        ).pack(side="left")
        self.result_count_var = tk.StringVar(value="尚未查找")
        self.result_count_label = tk.Label(
            title_row,
            textvariable=self.result_count_var,
            bg=CARD,
            fg=MUTED,
            font=FONT_TINY,
            padx=8,
        )
        self.result_count_label.pack(side="left")

        actions = tk.Frame(header, bg=CARD)
        actions.grid(row=0, column=2, sticky="e")
        self.copy_button = make_button(
            actions,
            "复制地址",
            self.copy_selected_address,
            kind="ghost",
            font=FONT_TINY,
            padx=10,
            pady=6,
        )
        self.copy_button.pack(side="left", padx=(0, 7))
        self.locate_button = make_button(
            actions,
            "在 Excel 中定位",
            self.locate_selected,
            kind="soft",
            font=("Microsoft YaHei UI", 9, "bold"),
            padx=12,
            pady=6,
        )
        self.locate_button.pack(side="left")
        clear_btn = tk.Button(
            actions,
            text="清空",
            command=self.clear_results,
            bg=CARD,
            fg=MUTED_LIGHT,
            activebackground=CARD,
            activeforeground=INK_SOFT,
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
            padx=6,
            pady=6,
        )
        clear_btn.pack(side="left", padx=(6, 0))

        separator = tk.Frame(card, bg=BORDER_SOFT, height=1)
        separator.grid(row=1, column=0, sticky="ew")

        table_wrap = tk.Frame(card, bg=CARD)
        table_wrap.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 0))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)
        self.table_wrap = table_wrap

        columns = ("score", "kind", "sheet", "cell", "value", "rowtext")
        self.tree = ttk.Treeview(
            table_wrap,
            columns=columns,
            show="headings",
            selectmode="browse",
            style="Results.Treeview",
        )
        headings = {
            "score": "匹配度",
            "kind": "方式",
            "sheet": "工作表",
            "cell": "单元格",
            "value": "匹配内容",
            "rowtext": "所在行",
        }
        widths = {
            "score": 72,
            "kind": 84,
            "sheet": 96,
            "cell": 68,
            "value": 150,
            "rowtext": 260,
        }
        anchors = {
            "score": "center",
            "kind": "center",
            "sheet": "w",
            "cell": "center",
            "value": "w",
            "rowtext": "w",
        }
        for column in columns:
            self.tree.heading(column, text=headings[column], anchor=anchors[column])
            self.tree.column(
                column,
                width=widths[column],
                minwidth=widths[column] if column != "rowtext" else 180,
                stretch=(column == "rowtext"),
                anchor=anchors[column],
            )
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll = ttk.Scrollbar(
            table_wrap,
            orient="vertical",
            command=self.tree.yview,
            style="App.Vertical.TScrollbar",
        )
        tree_scroll.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.tag_configure("exact", background=TEAL_PALE)
        self.tree.tag_configure("loose", background="#FFFBEB")
        self.tree.bind("<<TreeviewSelect>>", self._on_result_select)
        self.tree.bind("<Double-1>", lambda _e: self.locate_selected())
        self.tree.bind("<Button-3>", self._show_result_menu)
        self.tree.bind("<Return>", lambda _e: self.locate_selected())

        self.results_progress = ttk.Progressbar(
            table_wrap,
            mode="indeterminate",
            style="Accent.Horizontal.TProgressbar",
        )
        self.results_progress.place(relx=0, rely=0, relwidth=1)
        self.results_progress.place_forget()

        self.empty_frame = tk.Frame(table_wrap, bg=CARD)
        self.empty_frame.grid_columnconfigure(0, weight=1)
        self.empty_icon = tk.Canvas(
            self.empty_frame, width=56, height=56, bg=CARD, highlightthickness=0
        )
        self.empty_icon.grid(row=0, column=0, pady=(0, 8))
        self.empty_icon.create_oval(4, 4, 52, 52, fill=TEAL_PALE, outline="")
        self.empty_icon.create_oval(17, 16, 38, 37, outline=TEAL_DARK, width=3)
        self.empty_icon.create_line(35, 35, 45, 45, fill=TEAL_DARK, width=4)
        self.empty_icon.create_line(23, 27, 33, 27, fill=TEAL, width=2)
        self.empty_title_var = tk.StringVar(value="等待框选文字")
        tk.Label(
            self.empty_frame,
            textvariable=self.empty_title_var,
            bg=CARD,
            fg=INK,
            font=("Microsoft YaHei UI", 12, "bold"),
        ).grid(row=1, column=0)
        self.empty_subtitle_var = tk.StringVar(
            value=f"按 {self.hotkey.label} 框选屏幕上的文字\n识别结果会自动在整本工作簿中查找"
        )
        tk.Label(
            self.empty_frame,
            textvariable=self.empty_subtitle_var,
            bg=CARD,
            fg=MUTED,
            font=FONT_SMALL,
            justify="center",
            wraplength=560,
        ).grid(row=2, column=0, pady=(6, 10))
        hotkey_hint = tk.Frame(self.empty_frame, bg="#F8FAFC")
        hotkey_hint.grid(row=3, column=0)
        self.empty_hint = hotkey_hint
        self.empty_hotkey_var = tk.StringVar(value=self.hotkey.label)
        tk.Label(
            hotkey_hint,
            textvariable=self.empty_hotkey_var,
            bg="#F8FAFC",
            fg=TEAL_DARK,
            font=("Segoe UI", 8, "bold"),
            padx=8,
            pady=4,
        ).pack(side="left")
        tk.Label(
            hotkey_hint,
            text="框选",
            bg="#F8FAFC",
            fg=MUTED_LIGHT,
            font=FONT_TINY,
            padx=8,
            pady=4,
        ).pack(side="left")
        self._empty_show_hint = True
        self._empty_compact = False
        self.empty_frame.place(relx=0.5, rely=0.5, y=18, anchor="center")
        table_wrap.bind("<Configure>", self._on_results_table_resize, add="+")

        self.row_detail_frame = tk.Frame(card, bg=CARD)
        self.row_detail_frame.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 0))
        detail_header = tk.Frame(self.row_detail_frame, bg=CARD)
        detail_header.pack(fill="x", pady=(2, 4))
        tk.Label(
            detail_header,
            text="整行内容",
            bg=CARD,
            fg=INK,
            font=("Microsoft YaHei UI", 9, "bold"),
        ).pack(side="left")
        self.row_detail_meta_var = tk.StringVar(value="")
        tk.Label(
            detail_header,
            textvariable=self.row_detail_meta_var,
            bg=CARD,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="left", padx=(7, 0))
        self.row_detail_toggle_button = tk.Button(
            detail_header,
            text="收起",
            command=self.toggle_row_detail,
            bg=CARD,
            fg=MUTED,
            activebackground=CARD,
            activeforeground=INK_SOFT,
            relief="flat",
            bd=0,
            cursor="hand2",
            font=FONT_TINY,
            padx=5,
            pady=2,
        )
        self.row_detail_toggle_button.pack(side="right")
        self.row_detail_copy_button = make_button(
            detail_header,
            "复制整行",
            self.copy_selected_row,
            kind="ghost",
            font=FONT_TINY,
            padx=8,
            pady=3,
        )
        self.row_detail_copy_button.pack(side="right", padx=(0, 6))
        self.row_detail_copy_button.configure(state="disabled")

        self.row_detail_body = tk.Frame(self.row_detail_frame, bg=CARD)
        self.row_detail_body.pack(fill="x", pady=(0, 6))
        detail_border = tk.Frame(self.row_detail_body, bg=BORDER)
        detail_border.pack(fill="x")
        detail_inner = tk.Frame(detail_border, bg="#F8FAFC")
        detail_inner.pack(fill="both", expand=True, padx=1, pady=1)
        self.row_detail_text = tk.Text(
            detail_inner,
            height=2,
            width=1,
            wrap="word",
            bg="#F8FAFC",
            fg=INK_SOFT,
            relief="flat",
            bd=0,
            padx=9,
            pady=6,
            font=("Microsoft YaHei UI", 9),
            selectbackground=TEAL_LIGHT,
            selectforeground=INK,
        )
        self.row_detail_text.pack(side="left", fill="both", expand=True)
        detail_scroll = ttk.Scrollbar(
            detail_inner,
            orient="vertical",
            command=self.row_detail_text.yview,
            style="App.Vertical.TScrollbar",
        )
        detail_scroll.pack(side="right", fill="y")
        self.row_detail_text.configure(yscrollcommand=detail_scroll.set)
        self.row_detail_text.insert(
            "1.0", "选择一条结果后，这里会显示完整的一行内容，可直接复制。"
        )
        self.row_detail_text.configure(state="disabled")
        self._detail_open = True
        self.row_detail_frame.grid_remove()

        footer = tk.Frame(card, bg=CARD)
        footer.grid(row=4, column=0, sticky="ew", padx=16, pady=(4, 8))
        tk.Label(
            footer,
            text="单击查看整行  ·  双击在 Excel 中定位  ·  右键更多",
            bg=CARD,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
            anchor="w",
        ).pack(side="left")
        self.footer_status_var = tk.StringVar(value="")
        tk.Label(
            footer,
            textvariable=self.footer_status_var,
            bg=CARD,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
            anchor="e",
        ).pack(side="right")

        self.result_menu = tk.Menu(self.root, tearoff=0, font=FONT_SMALL)
        self.result_menu.add_command(label="在 Excel 中定位", command=self.locate_selected)
        self.result_menu.add_command(
            label="复制单元格地址", command=self.copy_selected_address
        )
        self.result_menu.add_command(
            label="复制整行内容", command=self.copy_selected_row
        )

    def _set_chips(self, lines: list[str]) -> None:
        for child in self.chips_frame.winfo_children():
            child.destroy()
        if not lines:
            tk.Label(
                self.chips_frame,
                text="框选后自动提取可查找内容",
                bg=CARD,
                fg=MUTED_LIGHT,
                font=FONT_TINY,
                anchor="w",
            ).pack(side="left")
            return
        for line in lines[:3]:
            label = _truncate(line, 14)
            btn = tk.Button(
                self.chips_frame,
                text=label,
                command=lambda value=line: self._use_candidate(value),
                bg="#F1F5F9",
                fg=INK_SOFT,
                activebackground=TEAL_LIGHT,
                activeforeground=TEAL_DARK,
                relief="flat",
                bd=0,
                cursor="hand2",
                font=FONT_TINY,
                padx=8,
                pady=3,
            )
            _bind_hover(btn, "#F1F5F9", TEAL_LIGHT)
            btn.pack(side="left", padx=(0, 6))
        if len(lines) > 3:
            more = tk.Button(
                self.chips_frame,
                text=f"+{len(lines) - 3} 行",
                command=lambda: self._show_more_lines(lines),
                bg=CARD,
                fg=TEAL_DARK,
                activebackground=CARD,
                activeforeground=TEAL_DARK,
                relief="flat",
                bd=0,
                cursor="hand2",
                font=FONT_TINY,
                padx=4,
                pady=3,
            )
            more.pack(side="left")

    # ------------------------------------------------------------------
    # 快捷键 / 热键
    # ------------------------------------------------------------------

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.choose_file())
        self.root.bind("<Control-O>", lambda _e: self.choose_file())
        self.root.bind("<F5>", lambda _e: self.reload_file())
        self.root.bind("<Control-Return>", lambda _e: self.run_search())
        self.root.bind("<Control-l>", lambda _e: self.query_entry.prepare_for_input())

    def _start_hotkey(self) -> None:
        ok = self.hotkey.start()
        self._set_hotkey_state(ok)
        if not ok:
            logging.warning("全局快捷键注册失败：%s", self.hotkey.error)

    def _set_hotkey_state(self, ok: bool) -> None:
        try:
            self.hotkey_dot.itemconfigure(1, fill=GREEN if ok else AMBER)
            self.hotkey_chip.configure(fg="#CBD5E1" if ok else "#FCD34D")
            self.hotkey_change_button.configure(
                text="更改" if ok else "冲突",
                fg=TEAL if ok else "#FCD34D",
            )
        except tk.TclError:
            pass

    def _update_hotkey_ui(self, previous_label: str = "") -> None:
        label = self.hotkey.label
        self._hotkey_label = label
        try:
            self.hotkey_label_var.set(label)
            self.empty_hotkey_var.set(label)
            subtitle = self.empty_subtitle_var.get()
            if previous_label and previous_label in subtitle:
                self.empty_subtitle_var.set(
                    subtitle.replace(previous_label, label)
                )
            elif not previous_label:
                self.empty_subtitle_var.set(
                    subtitle.replace(self._shown_hotkey_label, label)
                )
        except tk.TclError:
            pass
        self._shown_hotkey_label = label

    def _apply_hotkey(self, modifiers: int, vk: int, label: str) -> tuple[bool, str]:
        modifiers &= HOTKEY_MOD_MASK
        if (
            modifiers == self.hotkey.modifiers
            and vk == self.hotkey.vk
            and self.hotkey.registered
        ):
            return True, ""

        old_modifiers = self.hotkey.modifiers
        old_vk = self.hotkey.vk
        old_label = self.hotkey.label
        old_hotkey = self.hotkey
        old_hotkey.stop()

        new_hotkey = GlobalHotkey(modifiers, vk, label)
        if new_hotkey.start():
            self.hotkey = new_hotkey
            self._hotkey_modifiers = modifiers
            self._hotkey_vk = vk
            self._hotkey_label = label
            self.config["hotkey_modifiers"] = modifiers
            self.config["hotkey_vk"] = vk
            self.config["hotkey_label"] = label
            save_config(self.config)
            self._set_hotkey_state(True)
            self._update_hotkey_ui(old_label)
            self._set_status(f"快捷键已设为 {label}", "ok")
            return True, ""

        error = new_hotkey.error or f"无法注册快捷键 {label}"
        try:
            new_hotkey.stop()
        except Exception:
            pass

        restored = GlobalHotkey(old_modifiers, old_vk, old_label)
        restored_ok = restored.start()
        self.hotkey = restored
        self._set_hotkey_state(restored_ok)
        self._update_hotkey_ui(old_label)
        if not restored_ok:
            error += f"\n原快捷键 {old_label} 也未能恢复，请重新设置。"
        self._set_status("快捷键未更改", "error")
        return False, error

    def show_hotkey_dialog(self) -> None:
        if (
            self._hotkey_dialog is not None
            and self._hotkey_dialog.winfo_exists()
        ):
            self._hotkey_dialog.lift()
            self._hotkey_dialog.focus_force()
            return

        win = tk.Toplevel(self.root)
        self._hotkey_dialog = win
        win.title("设置快捷键 · 识文寻表")
        win.configure(bg=BG)
        win.transient(self.root)
        win.resizable(False, False)
        win.geometry("680x560")
        x = self.root.winfo_rootx() + max(20, (self.root.winfo_width() - 680) // 2)
        y = self.root.winfo_rooty() + max(20, (self.root.winfo_height() - 560) // 2)
        win.geometry(f"+{x}+{y}")
        if self._topmost_enabled:
            try:
                win.attributes("-topmost", True)
            except tk.TclError:
                pass

        header = tk.Frame(win, bg=NAVY, height=96)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="设置全局快捷键",
            bg=NAVY,
            fg="#F8FAFC",
            font=("Microsoft YaHei UI", 16, "bold"),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(15, 1))
        tk.Label(
            header,
            text="选择一套不冲突的组合键，用于任何界面下快速框选",
            bg=NAVY,
            fg="#5EEAD4",
            font=FONT_TINY,
            anchor="w",
        ).pack(fill="x", padx=24)

        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=16)

        current_card = card_frame(body)
        current_card.pack(fill="x", pady=(0, 10))
        current_inner = tk.Frame(current_card, bg=CARD)
        current_inner.pack(fill="x", padx=14, pady=10)
        tk.Label(
            current_inner,
            text="当前快捷键",
            bg=CARD,
            fg=MUTED,
            font=FONT_TINY,
            anchor="w",
        ).pack(side="left")
        tk.Label(
            current_inner,
            text=self.hotkey.label,
            bg=CARD,
            fg=TEAL_DARK,
            font=("Segoe UI", 11, "bold"),
            anchor="e",
        ).pack(side="right")

        tk.Label(
            body,
            text="点击下面的输入框，然后按下新的组合键",
            bg=BG,
            fg=INK,
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            body,
            text="建议至少搭配 Ctrl、Alt 或 Shift；单独使用 F1–F12 也可以。Esc 取消。",
            bg=BG,
            fg=MUTED,
            font=FONT_TINY,
            anchor="w",
        ).pack(fill="x", pady=(2, 8))

        capture_var = tk.StringVar(value="点击这里，然后按下组合键")
        capture_border = tk.Frame(body, bg=BORDER)
        capture_border.pack(fill="x")
        capture_entry = tk.Entry(
            capture_border,
            textvariable=capture_var,
            state="readonly",
            readonlybackground="#F8FAFC",
            fg=INK,
            justify="center",
            relief="flat",
            bd=0,
            font=("Segoe UI", 15, "bold"),
            cursor="hand2",
            insertwidth=0,
        )
        capture_entry.pack(fill="x", padx=1, pady=1, ipady=12)
        capture_entry.bind("<Button-1>", lambda _e: capture_entry.focus_force())

        presets_row = tk.Frame(body, bg=BG)
        presets_row.pack(fill="x", pady=(9, 3))
        tk.Label(
            presets_row,
            text="快速选择",
            bg=BG,
            fg=MUTED,
            font=FONT_TINY,
        ).pack(side="left", padx=(0, 7))

        pending: dict[str, int | str | None] = {
            "modifiers": None,
            "vk": None,
            "label": "",
        }
        held_mods: set[str] = set()
        mod_keys = {
            "Control_L": "Ctrl",
            "Control_R": "Ctrl",
            "Shift_L": "Shift",
            "Shift_R": "Shift",
            "Alt_L": "Alt",
            "Alt_R": "Alt",
        }

        info_var = tk.StringVar(
            value=f"当前使用 {self.hotkey.label}。更改后立即生效。"
        )
        info_label = tk.Label(
            body,
            textvariable=info_var,
            bg=BG,
            fg=MUTED,
            font=FONT_SMALL,
            anchor="w",
            justify="left",
            wraplength=620,
        )
        info_label.pack(fill="x", pady=(8, 0))

        footer = tk.Frame(win, bg=BG)
        footer.pack(fill="x", padx=24, pady=(0, 16))
        tk.Label(
            footer,
            text="若新快捷键仍被占用，程序会保留原快捷键。",
            bg=BG,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="left")

        save_button = make_button(
            footer,
            "保存并启用",
            lambda: save_hotkey(),
            kind="primary",
            font=FONT_SMALL,
            padx=18,
            pady=7,
        )
        save_button.configure(
            state="disabled", disabledforeground="#CCFBF1"
        )
        save_button.pack(side="right")
        cancel_button = make_button(
            footer,
            "取消",
            lambda: close_dialog(),
            kind="ghost",
            font=FONT_SMALL,
            padx=15,
            pady=7,
        )
        cancel_button.pack(side="right", padx=(0, 8))

        def update_pending(modifiers: int, vk: int, label: str) -> None:
            pending["modifiers"] = modifiers
            pending["vk"] = vk
            pending["label"] = label
            capture_var.set(label)
            save_button.configure(state="normal")
            info_var.set("按“保存并启用”应用这个组合键。")
            info_label.configure(fg=MUTED)

        def preview_held() -> None:
            parts: list[str] = []
            if "Ctrl" in held_mods:
                parts.append("Ctrl")
            if "Alt" in held_mods:
                parts.append("Alt")
            if "Shift" in held_mods:
                parts.append("Shift")
            if parts:
                capture_var.set(" + ".join(parts) + " + …")
            else:
                capture_var.set("点击这里，然后按下组合键")

        def preset(modifiers: int, vk: int) -> None:
            update_pending(modifiers, vk, hotkey_label(modifiers, vk))

        for preset_label, preset_mods, preset_vk in (
            ("Ctrl + Shift + Q", 0x0002 | 0x0004, 0x51),
            ("Alt + Shift + Q", 0x0001 | 0x0004, 0x51),
            ("Ctrl + Alt + S", 0x0001 | 0x0002, 0x53),
        ):
            b = tk.Button(
                presets_row,
                text=preset_label,
                command=lambda m=preset_mods, v=preset_vk: preset(m, v),
                bg=CARD,
                fg=TEAL_DARK,
                activebackground=TEAL_LIGHT,
                activeforeground=TEAL_DARK,
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=BORDER,
                cursor="hand2",
                font=FONT_TINY,
                padx=9,
                pady=4,
            )
            _bind_hover(b, CARD, TEAL_LIGHT)
            b.pack(side="left", padx=(0, 6))

        def on_key(event) -> str:
            keysym = event.keysym
            if keysym in mod_keys:
                held_mods.add(mod_keys[keysym])
                if pending["vk"] is None:
                    preview_held()
                return "break"
            if keysym == "Escape":
                close_dialog()
                return "break"
            if keysym in ("Return", "KP_Enter"):
                if pending["vk"] is not None:
                    save_hotkey()
                return "break"
            if keysym in ("Tab", "CapsLock", "NumLock", "ScrollLock"):
                return "break"

            modifiers = 0
            if "Ctrl" in held_mods or (event.state & 0x0004):
                modifiers |= 0x0002
            if "Shift" in held_mods or (event.state & 0x0001):
                modifiers |= 0x0004
            if (
                "Alt" in held_mods
                or (event.state & 0x0008)
                or (event.state & 0x20000)
            ):
                modifiers |= 0x0001

            vk = int(event.keycode)
            if not 1 <= vk <= 255:
                info_var.set("没有识别到这个按键，请再试一次。")
                info_label.configure(fg=RED)
                return "break"
            key_name = _key_name_from_vk(vk)
            if not modifiers and not key_name.startswith("F"):
                info_var.set("请至少搭配 Ctrl、Alt 或 Shift 之一；也可以使用 F1–F12。")
                info_label.configure(fg=RED)
                return "break"
            update_pending(modifiers, vk, hotkey_label(modifiers, vk, key_name))
            return "break"

        def on_release(event) -> str:
            if event.keysym in mod_keys:
                held_mods.discard(mod_keys[event.keysym])
                if pending["vk"] is None:
                    preview_held()
                return "break"
            return ""

        def save_hotkey() -> None:
            modifiers = pending["modifiers"]
            vk = pending["vk"]
            label = str(pending["label"] or "")
            if modifiers is None or vk is None or not label:
                info_var.set("请先按下或选择一个快捷键组合。")
                info_label.configure(fg=RED)
                return
            save_button.configure(state="disabled")
            info_var.set("正在尝试注册快捷键…")
            info_label.configure(fg=MUTED)
            win.update_idletasks()
            ok, message = self._apply_hotkey(int(modifiers), int(vk), label)
            if ok:
                close_dialog()
                return
            info_var.set(message)
            info_label.configure(fg=RED)
            save_button.configure(state="normal")

        def close_dialog() -> None:
            self._hotkey_dialog = None
            try:
                win.grab_release()
            except tk.TclError:
                pass
            try:
                win.destroy()
            except tk.TclError:
                pass

        capture_entry.bind("<KeyPress>", on_key)
        capture_entry.bind("<KeyRelease>", on_release)
        win.protocol("WM_DELETE_WINDOW", close_dialog)
        win.bind("<Escape>", lambda _e: close_dialog())
        try:
            win.grab_set()
        except tk.TclError:
            pass
        def focus_capture() -> None:
            try:
                win.focus_force()
                capture_entry.focus_force()
            except tk.TclError:
                pass

        win.after(60, focus_capture)

    def _poll_hotkey(self) -> None:
        if self._closing:
            return
        if self.hotkey.drain():
            self._on_hotkey()
        self.root.after(120, self._poll_hotkey)

    def _on_hotkey(self) -> None:
        if self._hotkey_dialog is not None:
            return
        if self._overlay is not None:
            return
        if self._busy is not None:
            self._set_status("正在处理，请稍候…", "busy")
            return
        self.start_capture()

    # ------------------------------------------------------------------
    # 启动加载 / 文件
    # ------------------------------------------------------------------

    def _startup_load(self) -> None:
        candidate: Path | None = None
        if self.initial_file:
            candidate = Path(self.initial_file)
        elif self.config.get("last_file"):
            candidate = Path(str(self.config["last_file"]))
        elif (CODE_DIR / "示例数据.xlsx").exists():
            candidate = CODE_DIR / "示例数据.xlsx"
        if candidate is not None and candidate.exists():
            self.load_file(candidate, auto=True, fallback_to_sample=True)
            return
        self._set_status("请选择 Excel 表格开始", "idle")
        self._set_empty_state(
            "先选择一份 Excel 表格",
            "支持 .xlsx、.xlsm 和 .csv；选中后会建立本地搜索索引\n"
            f"随后按 {self.hotkey.label} 框选屏幕上的文字",
        )

    def choose_file(self) -> None:
        if self._busy:
            self._set_status("正在处理，请稍候…", "busy")
            return
        current_dir = str(self.excel_path.parent) if self.excel_path else str(CODE_DIR)
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择要查找的 Excel 表格",
            initialdir=current_dir,
            filetypes=[
                ("Excel 工作簿", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.load_file(Path(path), auto=False)

    def reload_file(self) -> None:
        if not self.excel_path:
            self.choose_file()
            return
        if self._busy:
            self._set_status("正在处理，请稍候…", "busy")
            return
        self.load_file(
            self.excel_path,
            auto=False,
            then_search=self.last_query if self.last_query else None,
        )

    def load_file(
        self,
        path: Path,
        auto: bool = False,
        fallback_to_sample: bool = False,
        then_search: str | None = None,
    ) -> None:
        if self._busy:
            self._set_status("正在处理，请稍候…", "busy")
            return
        path = Path(path)
        self._busy = "load"
        self._pending_query = then_search or ""
        self.file_name_var.set(path.name)
        self.file_stats_var.set("正在读取并建立索引…")
        try:
            self.file_badge.configure(
                text=path.suffix.lstrip(".").upper()[:4] or "XLSX"
            )
        except tk.TclError:
            pass
        self._set_status("正在读取表格…", "busy")
        self._show_progress(True)
        self._set_empty_state(
            "正在读取表格",
            "首次打开会建立本地索引，较大的工作簿可能需要几秒。",
            show_hint=False,
        )
        self.search_button.configure(state="disabled")

        def worker() -> None:
            try:
                wb = core.load_workbook(
                    path,
                    progress=lambda message: self._ui_queue.put(
                        ("load_progress", message)
                    ),
                )
                self._ui_queue.put(("load_ok", (path, wb, auto, fallback_to_sample)))
            except Exception as exc:
                logging.exception("读取表格失败")
                self._ui_queue.put(
                    ("load_error", (path, exc, auto, fallback_to_sample))
                )

        threading.Thread(target=worker, name="load-workbook", daemon=True).start()

    def _handle_load_ok(
        self,
        path: Path,
        wb: core.WorkbookData,
        auto: bool,
        fallback_to_sample: bool,
    ) -> None:
        self.workbook = wb
        self.excel_path = path
        self._busy = None
        self.search_button.configure(state="normal")
        self._show_progress(False)
        self.file_name_var.set(path.name)
        self.file_badge.configure(text=path.suffix.lstrip(".").upper()[:4] or "XLSX")
        stats = (
            f"{wb.sheet_count} 个工作表 · {wb.data_row_count:,} 行 · "
            f"{core.format_file_size(wb.file_size)}"
        )
        self.file_stats_var.set(stats)
        self.tree.delete(*self.tree.get_children())
        self._row_hits = []
        self.last_result = None
        self._hide_row_detail()
        self.result_count_var.set("尚未查找")
        self.footer_status_var.set("")
        self._set_empty_state(
            "表格已就绪",
            f"按 {self.hotkey.label} 框选屏幕上的文字\n识别后会自动在整本工作簿中查找",
        )
        warning_text = f" · {len(wb.warnings)} 个提示" if wb.warnings else ""
        self._set_status(f"已载入 {wb.data_row_count:,} 行{warning_text}", "ok")
        if not self.hotkey.registered:
            self._set_status("快捷键冲突", "error")
        self.config["last_file"] = str(path)
        save_config(self.config)
        if then_search := self._pending_query:
            self._pending_query = ""
            self.root.after(80, lambda: self.run_search(query=then_search))
        logging.info(
            "Loaded %s: %s sheets, %s rows, %.0f ms",
            path,
            wb.sheet_count,
            wb.data_row_count,
            wb.load_ms,
        )

    def _handle_load_error(
        self,
        path: Path,
        exc: Exception,
        auto: bool,
        fallback_to_sample: bool,
    ) -> None:
        logging.error(
            "Load error for %s: %s",
            path,
            exc,
            exc_info=(type(exc), exc, exc.__traceback__),
        )
        self._pending_query = ""
        self._busy = None
        self.search_button.configure(state="normal")
        self._show_progress(False)
        message = str(exc) if isinstance(exc, core.ExcelReadError) else f"读取失败：{exc}"
        if self.workbook is None:
            self.file_name_var.set("尚未选择表格")
            self.file_stats_var.set("选择 .xlsx / .xlsm / .csv")
        else:
            self.file_name_var.set(self.excel_path.name if self.excel_path else "表格")
            self.file_stats_var.set(
                f"{self.workbook.sheet_count} 个工作表 · "
                f"{self.workbook.data_row_count:,} 行"
            )
        self._set_status("表格读取失败", "error")
        self._set_empty_state("无法读取表格", message, show_hint=False)
        if auto:
            sample = CODE_DIR / "示例数据.xlsx"
            if fallback_to_sample and sample.exists() and path != sample:
                self._set_status("已载入示例表格", "error")
                self.root.after(100, lambda: self.load_file(sample, auto=True))
                return
            self._set_status("表格读取失败", "error")
        messagebox.showerror("无法读取表格", message, parent=self.root)

    # ------------------------------------------------------------------
    # 截图 / OCR
    # ------------------------------------------------------------------

    def start_capture(self) -> None:
        if self._overlay is not None or self._capturing:
            return
        if self._busy:
            self._set_status("正在处理，请稍候…", "busy")
            return
        self._capturing = True
        self._set_status("准备框选…", "busy")
        self.root.withdraw()
        self.root.after(190, self._capture_after_hide)

    def _capture_after_hide(self) -> None:
        if self._closing:
            return
        try:
            bounds = virtual_screen_rect()
            shot = capture_screen()
        except Exception as exc:
            logging.exception("Screen capture failed")
            self._capturing = False
            self._restore_window()
            self._set_status("截屏失败", "error")
            messagebox.showerror("无法截屏", f"截取屏幕失败：\n{exc}", parent=self.root)
            return
        try:
            self._overlay = SelectionOverlay(
                self.root,
                shot,
                bounds,
                self._on_region_selected,
                self._on_capture_cancelled,
            )
        except Exception as exc:
            logging.exception("Selection overlay failed")
            self._capturing = False
            self._overlay = None
            self._restore_window()
            self._set_status("框选窗口打开失败", "error")
            messagebox.showerror("无法框选", f"打开框选窗口失败：\n{exc}", parent=self.root)

    def _on_region_selected(
        self, image: Image.Image, _rect: tuple[int, int, int, int]
    ) -> None:
        self._overlay = None
        self._capturing = False
        self._restore_window()
        if image.width < 8 or image.height < 6:
            self._set_status("选区太小，请重新框选", "error")
            return
        self._set_preview(image)
        self._run_ocr(image)

    def _on_capture_cancelled(self) -> None:
        self._overlay = None
        self._capturing = False
        self._restore_window()
        self._set_status("已取消框选", "idle")

    def _restore_window(self) -> None:
        try:
            self.root.deiconify()
            self.root.state("normal")
            self.root.attributes("-topmost", self._topmost_enabled)
            self.root.lift()
            self.root.focus_force()
        except tk.TclError:
            pass

    def _set_preview(self, image: Image.Image) -> None:
        thumb = image.convert("RGB").copy()
        thumb.thumbnail((236, 210), Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(thumb, master=self.root)
        self.preview_label.configure(
            image=self.preview_photo,
            text="",
            bg="#F8FAFC",
        )
        self.preview_badge.configure(
            text=f"截图预览  {image.width}×{image.height}"
        )

    def _run_ocr(self, image: Image.Image) -> None:
        self._busy = "ocr"
        self._set_status("正在识别文字…", "busy")
        self._show_progress(True)
        self.search_button.configure(state="disabled")
        self.ocr_meta_var.set("正在调用 Windows OCR…")
        self._set_empty_state(
            "正在识别文字",
            "正在分析截图中的中英文内容，请稍候…",
            show_hint=False,
        )

        def worker() -> None:
            try:
                result = self.ocr_service.recognize(image)
                self._ui_queue.put(("ocr_ok", result))
            except Exception as exc:
                logging.exception("OCR failed")
                self._ui_queue.put(("ocr_error", exc))

        threading.Thread(target=worker, name="ocr", daemon=True).start()

    def _handle_ocr_ok(self, result: OcrResult) -> None:
        self._busy = None
        self._show_progress(False)
        self.search_button.configure(state="normal")
        self.ocr_lines = result.lines
        self.ocr_text.configure(fg=INK_SOFT)
        self.ocr_text.delete("1.0", "end")
        self.ocr_text.insert("1.0", result.text)
        language = result.language or "系统 OCR"
        self.ocr_meta_var.set(
            f"Windows OCR · {language} · {result.elapsed_ms / 1000:.1f}s"
        )
        self._set_chips(result.lines)
        query = core.suggest_query(result.lines)
        self.query_entry.set_value(query)
        self._set_status(f"识别完成 · {len(result.lines)} 行", "ok")
        self.root.after(30, lambda: self.run_search(query=query, auto=True))

    def _handle_ocr_error(self, exc: Exception) -> None:
        self._busy = None
        self._show_progress(False)
        self.search_button.configure(state="normal")
        self.ocr_meta_var.set("识别失败")
        message = str(exc) if isinstance(exc, OcrError) else f"识别失败：{exc}"
        self._set_status("未识别到文字", "error")
        self._set_empty_state(
            "这次没有识别出文字",
            "请框选更大、更清晰、对比度更高的文字区域\n也可以尝试放大页面后再框选",
            show_hint=False,
        )
        messagebox.showwarning("未能识别文字", message, parent=self.root)

    def _on_ocr_double_click(self, _event=None) -> None:
        try:
            selected = self.ocr_text.get("sel.first", "sel.last").strip()
        except tk.TclError:
            return
        if selected:
            selected = selected.replace("\n", " ").strip()
            self.query_entry.set_value(selected)
            self.run_search(query=selected)

    def _use_candidate(self, line: str) -> None:
        self.query_entry.set_value(line)
        self.run_search(query=line)

    def _show_more_lines(self, lines: list[str]) -> None:
        menu = tk.Menu(self.root, tearoff=0, font=FONT_SMALL)
        for line in lines[3:]:
            menu.add_command(
                label=_truncate(line, 50),
                command=lambda value=line: self._use_candidate(value),
            )
        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------
    # 查找
    # ------------------------------------------------------------------

    def request_search(self) -> None:
        self.run_search()

    def run_search(self, query: str | None = None, auto: bool = False) -> None:
        if self._busy:
            self._set_status("正在处理，请稍候…", "busy")
            return
        if self.workbook is None:
            self._set_status("请先选择 Excel 表格", "error")
            if messagebox.askyesno(
                "还没有表格",
                "需要先选择一份 Excel 表格才能查找。\n现在去选择吗？",
                parent=self.root,
            ):
                self.choose_file()
            return
        query = (query if query is not None else self.query_entry.get_value()).strip()
        if not query:
            self._set_status("请输入要查找的内容", "error")
            self.query_entry.prepare_for_input()
            return
        if self._workbook_changed():
            self._set_status("表格已更新，正在刷新…", "busy")
            self.load_file(
                self.excel_path or self.workbook.path,
                auto=False,
                then_search=query,
            )
            return
        mode = MODE_MAP.get(self.mode_var.get(), "smart")
        self._busy = "search"
        self.last_query = query
        self._pending_query = ""
        self._set_status("正在表格中查找…", "busy")
        self.result_count_var.set("查找中…")
        self.footer_status_var.set("")
        self._hide_row_detail()
        self._show_progress(True)
        self.search_button.configure(state="disabled")
        self._set_empty_state(
            "正在查找",
            "正在检索全部工作表，请稍候…",
            show_hint=False,
        )
        workbook = self.workbook

        def worker() -> None:
            try:
                result = core.search_workbook(workbook, query, mode=mode)
                self._ui_queue.put(("search_ok", (result, query, mode, auto)))
            except Exception as exc:
                logging.exception("Search failed")
                self._ui_queue.put(("search_error", (exc, query)))

        threading.Thread(target=worker, name="search", daemon=True).start()

    def _workbook_changed(self) -> bool:
        if not self.workbook:
            return False
        try:
            return abs(self.workbook.path.stat().st_mtime - self.workbook.mtime) > 0.01
        except OSError:
            return False

    def _handle_search_ok(
        self,
        result: core.SearchResult,
        query: str,
        mode: str,
        auto: bool,
    ) -> None:
        self._busy = None
        self._show_progress(False)
        self.search_button.configure(state="normal")
        self.last_result = result
        self.last_query = query
        self._render_results(result, query, mode)
        if result.total:
            self._add_recent(query)
            self._set_status(f"找到 {result.total} 条", "ok")
        else:
            self._set_status("没有找到匹配项", "error")
        logging.info(
            "Search %r mode=%s: %s hits in %.0f ms (displayed %s)",
            query,
            mode,
            result.total,
            result.elapsed_ms,
            len(result.hits),
        )

    def _handle_search_error(self, exc: Exception, query: str) -> None:
        self._busy = None
        self._show_progress(False)
        self.search_button.configure(state="normal")
        self._set_status("查找失败", "error")
        message = str(exc) if isinstance(exc, ValueError) else f"查找时发生错误：{exc}"
        self._set_empty_state("查找失败", message, show_hint=False)
        messagebox.showerror("查找失败", message, parent=self.root)

    def _render_results(
        self, result: core.SearchResult, query: str, mode: str
    ) -> None:
        self.tree.delete(*self.tree.get_children())
        self._row_hits = []
        for index, hit in enumerate(result.hits):
            tag = "exact" if hit.score >= 0.95 else "loose"
            values = (
                f"{round(hit.score * 100)}%",
                hit.match_type,
                hit.sheet,
                hit.address,
                _truncate(hit.cell_value, 30),
                _truncate(hit.row_text, 160),
            )
            self.tree.insert("", "end", iid=str(index), values=values, tags=(tag,))
            self._row_hits.append(hit)

        count_text = f"{result.total} 条"
        if result.collect_truncated:
            count_text = f"至少 {result.total} 条"
        if result.truncated:
            count_text += f" · 显示前 {len(result.hits)} 条"
        self.result_count_var.set(
            f"{count_text}   ·   {result.elapsed_ms:.0f} ms   ·   {core.mode_label(mode)}"
        )
        self.footer_status_var.set(f"检索词：{_truncate(query.replace(chr(10), ' / '), 42)}")

        if result.hits:
            self.empty_frame.place_forget()
            self._set_row_detail_open(self.root.winfo_height() >= 780)
            self.row_detail_frame.grid()
            first = self.tree.get_children()
            if first:
                self.tree.selection_set(first[0])
                self.tree.focus(first[0])
                if self._row_hits:
                    self._show_row_detail(self._row_hits[0])
        else:
            self._hide_row_detail()
            self._set_empty_state(
                f"没有找到“{_truncate(query, 22)}”",
                "可以试试缩短关键词、切换为“包含匹配”\n或检查表格中是否有空格、错别字、全半角差异",
                show_hint=False,
            )

    def _show_progress(self, visible: bool) -> None:
        try:
            if visible:
                self.results_progress.place(relx=0, rely=0, relwidth=1)
                self.results_progress.start(12)
            else:
                self.results_progress.stop()
                self.results_progress.place_forget()
        except tk.TclError:
            pass

    def _set_empty_state(
        self, title: str, subtitle: str, show_hint: bool = True
    ) -> None:
        self.empty_title_var.set(title)
        self.empty_subtitle_var.set(subtitle)
        self._empty_show_hint = show_hint
        if show_hint and not self._empty_compact:
            self.empty_hint.grid()
        else:
            self.empty_hint.grid_remove()
        self.empty_frame.place(relx=0.5, rely=0.5, y=18, anchor="center")

    def _on_results_table_resize(self, event) -> None:
        compact = event.height < 230
        if compact == self._empty_compact:
            return
        self._empty_compact = compact
        if compact:
            self.empty_icon.grid_remove()
            self.empty_hint.grid_remove()
        else:
            self.empty_icon.grid()
            if self._empty_show_hint:
                self.empty_hint.grid()

    def _selected_hit(self) -> core.SearchHit | None:
        selection = self.tree.selection()
        if not selection:
            return None
        try:
            index = int(selection[0])
            return self._row_hits[index]
        except (ValueError, IndexError):
            return None

    def _on_result_select(self, _event=None) -> None:
        hit = self._selected_hit()
        if hit:
            self._show_row_detail(hit)
            self.footer_status_var.set(
                f"{hit.sheet} · {hit.address} · {_truncate(hit.cell_value, 32)}"
            )

    def toggle_row_detail(self) -> None:
        self._set_row_detail_open(not self._detail_open)

    def _set_row_detail_open(self, open_state: bool) -> None:
        self._detail_open = bool(open_state)
        try:
            if self._detail_open:
                self.row_detail_body.pack(fill="x", pady=(0, 6))
                self.row_detail_toggle_button.configure(text="收起")
            else:
                self.row_detail_body.pack_forget()
                self.row_detail_toggle_button.configure(text="展开")
        except tk.TclError:
            pass

    def _show_row_detail(self, hit: core.SearchHit) -> None:
        self.row_detail_meta_var.set(
            f"{hit.sheet} · 第 {hit.row} 行 · {hit.address}"
        )
        self.row_detail_text.configure(state="normal")
        self.row_detail_text.delete("1.0", "end")
        self.row_detail_text.insert(
            "1.0", hit.row_text or "（这一行没有可显示的内容）"
        )
        self.row_detail_text.configure(state="disabled")
        self.row_detail_text.yview_moveto(0)
        self.row_detail_copy_button.configure(state="normal")
        try:
            self.row_detail_frame.grid()
        except tk.TclError:
            pass

    def _hide_row_detail(self) -> None:
        self.row_detail_meta_var.set("")
        self.row_detail_text.configure(state="normal")
        self.row_detail_text.delete("1.0", "end")
        self.row_detail_text.insert(
            "1.0", "选择一条结果后，这里会显示完整的一行内容，可直接复制。"
        )
        self.row_detail_text.configure(state="disabled")
        self.row_detail_copy_button.configure(state="disabled")
        try:
            self.row_detail_frame.grid_remove()
        except tk.TclError:
            pass

    def _show_result_menu(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if not row:
            return
        self.tree.selection_set(row)
        self.tree.focus(row)
        try:
            self.result_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.result_menu.grab_release()

    def copy_selected_address(self) -> None:
        hit = self._selected_hit()
        if not hit:
            self._set_status("请先选择一条结果", "error")
            return
        text = f"{hit.sheet}!{hit.address}"
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status("已复制单元格地址", "ok")

    def copy_selected_row(self) -> None:
        hit = self._selected_hit()
        if not hit:
            self._set_status("请先选择一条结果", "error")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(hit.row_text)
        self._set_status("已复制整行内容", "ok")

    def locate_selected(self) -> None:
        hit = self._selected_hit()
        if not hit:
            self._set_status("请先选择一条结果", "error")
            return
        if not self.excel_path:
            return
        self._set_status("正在打开 Excel…", "busy")
        workbook_path = str(self.excel_path)
        sheet = hit.sheet
        address = hit.address

        def worker() -> None:
            script = resource_path("assets", "open_excel.ps1")
            powershell = (
                Path(os.environ.get("SystemRoot", r"C:\Windows"))
                / "System32"
                / "WindowsPowerShell"
                / "v1.0"
                / "powershell.exe"
            )
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            try:
                completed = subprocess.run(
                    [
                        str(powershell),
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        str(script),
                        "-WorkbookPath",
                        workbook_path,
                        "-SheetName",
                        sheet,
                        "-Address",
                        address,
                    ],
                    capture_output=True,
                    timeout=25,
                    creationflags=creationflags,
                    startupinfo=startupinfo,
                )
                ok = completed.returncode == 0
                detail = (completed.stderr or b"").decode(
                    locale.getpreferredencoding(False), errors="replace"
                ).strip()
                self._ui_queue.put(
                    ("locate_done", (ok, sheet, address, workbook_path, detail))
                )
            except Exception as exc:
                self._ui_queue.put(
                    ("locate_done", (False, sheet, address, workbook_path, str(exc)))
                )

        threading.Thread(target=worker, name="locate-in-excel", daemon=True).start()

    def _handle_locate_done(
        self,
        ok: bool,
        sheet: str,
        address: str,
        workbook_path: str,
        detail: str,
    ) -> None:
        if ok:
            self._set_status("已在 Excel 中定位", "ok")
            return
        self._set_status("未能定位，地址已复制", "error")
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(f"{sheet}!{address}")
        except tk.TclError:
            pass
        messagebox.showinfo(
            "无法定位到 Excel",
            "未能调用 Excel 定位单元格。\n\n"
            f"已复制地址：{sheet}!{address}\n\n"
            "可能原因：未安装 Microsoft Excel、文件被加密，或 Excel 正在显示对话框。\n"
            + (f"\n{detail[-500:]}" if detail else ""),
            parent=self.root,
        )

    def clear_results(self) -> None:
        if self._busy:
            return
        self.tree.delete(*self.tree.get_children())
        self._row_hits = []
        self.last_result = None
        self._hide_row_detail()
        self.result_count_var.set("尚未查找")
        self.footer_status_var.set("")
        self._set_empty_state(
            "等待框选文字",
            f"按 {self.hotkey.label} 框选屏幕上的文字\n识别结果会自动在整本工作簿中查找",
        )
        self._set_status("结果已清空", "idle")

    # ------------------------------------------------------------------
    # 最近查找
    # ------------------------------------------------------------------

    def _add_recent(self, query: str) -> None:
        value = query.strip().replace("\n", " / ")
        if not value:
            return
        value = _truncate(value, 36)
        self._recent_queries = [item for item in self._recent_queries if item != value]
        self._recent_queries.insert(0, value)
        self._recent_queries = self._recent_queries[:5]
        self.config["recent_queries"] = self._recent_queries
        save_config(self.config)
        self._rebuild_recent()

    def _rebuild_recent(self) -> None:
        for child in self.recent_frame.winfo_children():
            child.destroy()
        if not self._recent_queries:
            tk.Label(
                self.recent_frame,
                text="还没有查找记录",
                bg=NAVY,
                fg="#475569",
                font=FONT_TINY,
                anchor="w",
            ).pack(fill="x", pady=2)
            for child in self.recent_frame.winfo_children():
                self._bind_sidebar_wheel(child)
            self.root.after_idle(self._sync_sidebar_scroll)
            return
        for query in self._recent_queries[:4]:
            btn = tk.Button(
                self.recent_frame,
                text=f"↗  {_truncate(query, 22)}",
                command=lambda value=query: self._use_recent(value),
                bg=NAVY,
                fg="#94A3B8",
                activebackground=NAVY_CARD,
                activeforeground="#F8FAFC",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=FONT_TINY,
                anchor="w",
                padx=2,
                pady=3,
            )
            _bind_hover(btn, NAVY, NAVY_CARD)
            btn.pack(fill="x")
            self._bind_sidebar_wheel(btn)
        self.root.after_idle(self._sync_sidebar_scroll)

    def _use_recent(self, query: str) -> None:
        self.query_entry.set_value(query)
        self.run_search(query=query)

    def clear_recent(self) -> None:
        self._recent_queries = []
        self.config["recent_queries"] = []
        save_config(self.config)
        self._rebuild_recent()

    # ------------------------------------------------------------------
    # 状态 / 帮助
    # ------------------------------------------------------------------

    def _set_status(self, text: str, tone: str = "ok") -> None:
        colors = {
            "ok": GREEN,
            "busy": AMBER,
            "error": RED,
            "idle": MUTED_LIGHT,
        }
        color = colors.get(tone, GREEN)
        try:
            self.status_var.set(text)
            self.status_dot.itemconfigure(self.status_dot_item, fill=color)
            self.status_label.configure(
                fg=INK_SOFT if tone != "error" else "#B91C1C"
            )
        except tk.TclError:
            pass

    def toggle_topmost(self) -> None:
        self._topmost_enabled = not self._topmost_enabled
        try:
            self.root.attributes("-topmost", self._topmost_enabled)
            if self._topmost_enabled:
                self.root.lift()
        except tk.TclError:
            pass
        self.config["topmost"] = self._topmost_enabled
        save_config(self.config)
        try:
            self.topmost_button.configure(
                text="置顶：开" if self._topmost_enabled else "置顶：关",
                fg=TEAL_DARK if self._topmost_enabled else MUTED,
            )
        except tk.TclError:
            pass
        self._set_status(
            "窗口已置顶" if self._topmost_enabled else "窗口已取消置顶",
            "ok" if self._topmost_enabled else "idle",
        )

    def _on_mode_changed(self, _event=None) -> None:
        self._mode_label = self.mode_var.get()
        self.config["mode"] = self._mode_label
        save_config(self.config)

    def show_help(self) -> None:
        if self._help_window is not None and self._help_window.winfo_exists():
            self._help_window.lift()
            self._help_window.focus_force()
            return
        win = tk.Toplevel(self.root)
        self._help_window = win
        win.title("使用说明 · 识文寻表")
        win.configure(bg=BG)
        win.transient(self.root)
        win.resizable(False, False)
        win.geometry("740x740")
        win.grab_set()
        x = self.root.winfo_rootx() + max(20, (self.root.winfo_width() - 740) // 2)
        y = self.root.winfo_rooty() + max(20, (self.root.winfo_height() - 740) // 2)
        win.geometry(f"+{x}+{y}")
        if self._topmost_enabled:
            try:
                win.attributes("-topmost", True)
            except tk.TclError:
                pass

        header = tk.Frame(win, bg=NAVY, height=86)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="三步完成查找",
            bg=NAVY,
            fg="#F8FAFC",
            font=("Microsoft YaHei UI", 17, "bold"),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(17, 1))
        tk.Label(
            header,
            text="OCR 文字识别  ×  Excel 全表检索",
            bg=NAVY,
            fg="#5EEAD4",
            font=FONT_TINY,
            anchor="w",
        ).pack(fill="x", padx=24)

        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=16)
        steps = [
            ("01", "选择表格", "点击左侧“选择表格”，支持 .xlsx、.xlsm、.csv；打开后建立本地索引，查找无需联网。"),
            ("02", "框选文字", f"按 {self.hotkey.label.replace(' + ', '+')}，或点击左侧绿色按钮；拖选文字后松开鼠标即可。"),
            ("03", "查看整行", "识别后自动查找；结果下方直接显示完整的整行内容，可一键复制，双击结果可在 Excel 中定位。"),
        ]
        for number, title, desc in steps:
            row = card_frame(body)
            row.pack(fill="x", pady=(0, 9))
            badge = tk.Label(
                row,
                text=number,
                bg=TEAL_PALE,
                fg=TEAL_DARK,
                font=("Segoe UI", 10, "bold"),
                width=4,
                pady=7,
            )
            badge.pack(side="left", padx=10, pady=10)
            text_box = tk.Frame(row, bg=CARD)
            text_box.pack(side="left", fill="x", expand=True, pady=9, padx=(0, 12))
            tk.Label(
                text_box,
                text=title,
                bg=CARD,
                fg=INK,
                font=("Microsoft YaHei UI", 10, "bold"),
                anchor="w",
            ).pack(fill="x")
            tk.Label(
                text_box,
                text=desc,
                bg=CARD,
                fg=MUTED,
                font=FONT_SMALL,
                anchor="w",
                justify="left",
                wraplength=570,
            ).pack(fill="x", pady=(2, 0))

        tips = tk.Frame(body, bg="#ECFDF5")
        tips.pack(fill="x", pady=(2, 0))
        tk.Label(
            tips,
            text="小技巧",
            bg="#ECFDF5",
            fg=TEAL_DARK,
            font=("Microsoft YaHei UI", 9, "bold"),
            anchor="w",
        ).pack(fill="x", padx=12, pady=(8, 2))
        tk.Label(
            tips,
            text=(
                "· 识别多行时，可点击“快捷”标签切换要查找的内容\n"
                "· 找不到结果时，试试“包含匹配”或只保留关键词\n"
                "· 结果下方可直接查看并复制整行完整内容\n"
                "· 快捷键冲突时，点左侧“更改”重新设置；快捷键和置顶状态都会记住"
            ),
            bg="#ECFDF5",
            fg="#065F46",
            font=FONT_SMALL,
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=12, pady=(0, 9))

        footer = tk.Frame(win, bg=BG)
        footer.pack(fill="x", padx=24, pady=(0, 16))
        tk.Label(
            footer,
            text=f"{APP_NAME} v{APP_VERSION} · Windows 系统 OCR",
            bg=BG,
            fg=MUTED_LIGHT,
            font=FONT_TINY,
        ).pack(side="left")
        make_button(
            footer,
            "知道了",
            win.destroy,
            kind="primary",
            font=FONT_SMALL,
            padx=22,
            pady=7,
        ).pack(side="right")

        def close_help() -> None:
            self._help_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close_help)
        win.bind("<Escape>", lambda _e: close_help())
        win.after(50, win.focus_force)

    # ------------------------------------------------------------------
    # 队列 / 关闭
    # ------------------------------------------------------------------

    def _poll_queue(self) -> None:
        if self._closing:
            return
        while True:
            try:
                kind, payload = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                if kind == "load_progress":
                    self.file_stats_var.set(_truncate(str(payload), 38))
                elif kind == "load_ok":
                    self._handle_load_ok(*payload)
                elif kind == "load_error":
                    self._handle_load_error(*payload)
                elif kind == "ocr_ok":
                    self._handle_ocr_ok(payload)
                elif kind == "ocr_error":
                    self._handle_ocr_error(payload)
                elif kind == "search_ok":
                    self._handle_search_ok(*payload)
                elif kind == "search_error":
                    self._handle_search_error(*payload)
                elif kind == "locate_done":
                    self._handle_locate_done(*payload)
            except Exception:
                logging.exception("UI queue handler failed: %s", kind)
        self.root.after(55, self._poll_queue)

    def _report_callback_exception(self, exc, val, tb) -> None:
        logging.error("Tk callback exception: %s", "".join(traceback.format_exception(exc, val, tb)))
        self._set_status("发生错误", "error")
        messagebox.showerror(
            "发生错误",
            f"操作失败：{val}\n\n详细信息已记录到：\n{LOG_PATH}",
            parent=self.root,
        )

    def _on_close(self) -> None:
        if self._closing:
            return
        self._closing = True
        try:
            self.hotkey.stop()
        except Exception:
            pass
        try:
            if self.root.state() == "normal":
                self.config["geometry"] = self.root.winfo_geometry()
        except tk.TclError:
            pass
        self.config["mode"] = self.mode_var.get()
        self.config["recent_queries"] = self._recent_queries
        self.config["topmost"] = self._topmost_enabled
        self.config["hotkey_modifiers"] = self.hotkey.modifiers
        self.config["hotkey_vk"] = self.hotkey.vk
        self.config["hotkey_label"] = self.hotkey.label
        save_config(self.config)
        try:
            self.root.destroy()
        except tk.TclError:
            pass


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------


_SINGLE_INSTANCE_MUTEX = None


def _acquire_single_instance() -> bool:
    """避免重复启动导致全局快捷键互相抢占。"""

    global _SINGLE_INSTANCE_MUTEX
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        ctypes.set_last_error(0)
        _SINGLE_INSTANCE_MUTEX = kernel32.CreateMutexW(
            None, False, "Local\\ShiWenXunBiao.SingleInstance"
        )
        if ctypes.get_last_error() == 183:
            title = f"{APP_NAME} · 框选识字，自动查表"
            try:
                hwnd = ctypes.windll.user32.FindWindowW(None, title)
                if hwnd:
                    ctypes.windll.user32.ShowWindow(hwnd, 9)
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                pass
            ctypes.windll.user32.MessageBoxW(
                None,
                "识文寻表已经在运行。\n请查看任务栏中的窗口。",
                APP_NAME,
                0x00000040,
            )
            return False
    except Exception:
        logging.exception("单实例检测失败，继续启动")
    return True


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--file", type=str, default=None, help="启动时打开的表格")
    return parser.parse_args()


def main() -> None:
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Starting %s %s", APP_NAME, APP_VERSION)
    enable_dpi_awareness()
    set_app_user_model_id("ShiWenXunBiao.OCRFinder")
    if not _acquire_single_instance():
        return
    args = parse_args()
    root = tk.Tk()
    root.withdraw()
    initial = Path(args.file) if args.file else None
    app = ShiWenApp(root, initial_file=initial)
    root.deiconify()
    root.mainloop()
    logger.info("Exit")


if __name__ == "__main__":
    main()
