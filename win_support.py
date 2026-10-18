"""Windows 屏幕捕获、框选浮层与全局快捷键。"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import queue
import threading
from typing import Callable

from PIL import Image, ImageEnhance, ImageTk


# ---------------------------------------------------------------------------
# DPI / 屏幕
# ---------------------------------------------------------------------------


def enable_dpi_awareness() -> None:
    """让 Tk 坐标和屏幕截图使用一致的物理像素。"""

    try:
        user32 = ctypes.windll.user32
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetProcessDpiAwarenessContext.restype = wt.BOOL
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def set_app_user_model_id(app_id: str) -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def virtual_screen_rect() -> tuple[int, int, int, int]:
    user32 = ctypes.windll.user32
    user32.GetSystemMetrics.argtypes = [ctypes.c_int]
    user32.GetSystemMetrics.restype = ctypes.c_int
    left = user32.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
    top = user32.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN
    width = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
    height = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN
    if width <= 0 or height <= 0:
        width = user32.GetSystemMetrics(0)
        height = user32.GetSystemMetrics(1)
        left = top = 0
    return left, top, width, height


def capture_screen() -> Image.Image:
    from PIL import ImageGrab

    return ImageGrab.grab(all_screens=True)


def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    rect = wt.RECT()
    ctypes.windll.user32.GetWindowRect(wt.HWND(hwnd), ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


# ---------------------------------------------------------------------------
# 全局快捷键
# ---------------------------------------------------------------------------


class GlobalHotkey:
    """用 RegisterHotKey 注册一个可配置的全局快捷键。"""

    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012
    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_WIN = 0x0008
    MOD_NOREPEAT = 0x4000
    HOTKEY_ID = 0xB0C1

    def __init__(
        self,
        modifiers: int = MOD_CONTROL | MOD_SHIFT,
        vk: int = 0x51,
        label: str = "Ctrl + Shift + Q",
    ) -> None:
        self.modifiers = modifiers
        self.vk = vk
        self.label = label
        self.events: queue.Queue[None] = queue.Queue()
        self.registered = False
        self.error: str | None = None
        self._thread: threading.Thread | None = None
        self._thread_id: int | None = None
        self._ready = threading.Event()

    @property
    def description(self) -> str:
        return self.label

    def start(self, timeout: float = 1.2) -> bool:
        if self._thread is not None:
            return self.registered
        self._thread = threading.Thread(
            target=self._run, name="global-hotkey", daemon=True
        )
        self._thread.start()
        self._ready.wait(timeout)
        return self.registered

    def _run(self) -> None:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.RegisterHotKey.argtypes = [wt.HWND, ctypes.c_int, wt.UINT, wt.UINT]
        user32.RegisterHotKey.restype = wt.BOOL
        user32.UnregisterHotKey.argtypes = [wt.HWND, ctypes.c_int]
        user32.UnregisterHotKey.restype = wt.BOOL
        user32.GetMessageW.argtypes = [
            ctypes.POINTER(wt.MSG),
            wt.HWND,
            wt.UINT,
            wt.UINT,
        ]
        user32.GetMessageW.restype = wt.BOOL
        user32.PostThreadMessageW.argtypes = [
            wt.DWORD,
            wt.UINT,
            wt.WPARAM,
            wt.LPARAM,
        ]
        user32.PostThreadMessageW.restype = wt.BOOL
        kernel32.GetCurrentThreadId.restype = wt.DWORD

        self._thread_id = int(kernel32.GetCurrentThreadId())
        ok = bool(
            user32.RegisterHotKey(
                None,
                self.HOTKEY_ID,
                self.modifiers | self.MOD_NOREPEAT,
                self.vk,
            )
        )
        self.registered = ok
        if not ok:
            self.error = f"快捷键 {self.label} 被其他程序占用"
            self._ready.set()
            return
        self._ready.set()

        msg = wt.MSG()
        try:
            while True:
                result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if result == 0 or result == -1:
                    break
                if msg.message == self.WM_HOTKEY and msg.wParam == self.HOTKEY_ID:
                    self.events.put(None)
        finally:
            user32.UnregisterHotKey(None, self.HOTKEY_ID)
            self.registered = False

    def stop(self) -> None:
        if self._thread_id:
            try:
                ctypes.windll.user32.PostThreadMessageW(
                    self._thread_id, self.WM_QUIT, 0, 0
                )
            except Exception:
                pass
        if self._thread is not None and self._thread.is_alive():
            self._thread.join(timeout=0.5)
        self._thread = None

    def drain(self) -> bool:
        fired = False
        while True:
            try:
                self.events.get_nowait()
                fired = True
            except queue.Empty:
                return fired


# ---------------------------------------------------------------------------
# 框选浮层
# ---------------------------------------------------------------------------


class SelectionOverlay:
    """覆盖虚拟桌面的截图，让用户拖拽框选。"""

    def __init__(
        self,
        root,
        screenshot: Image.Image,
        bounds: tuple[int, int, int, int],
        on_done: Callable[[Image.Image, tuple[int, int, int, int]], None],
        on_cancel: Callable[[], None],
    ) -> None:
        import tkinter as tk

        self.tk = tk
        self.root = root
        self.screenshot = screenshot
        self.bounds = bounds
        self.on_done = on_done
        self.on_cancel = on_cancel
        self.closed = False
        self.start: tuple[int, int] | None = None
        self.end: tuple[int, int] | None = None
        self._dim_photo = None
        self._size_rect = None
        self._size_text = None
        self._tip_text = None

        left, top, width, height = bounds
        self.left = left
        self.top = top
        self.width = max(1, width)
        self.height = max(1, height)

        base = screenshot.convert("RGB")
        if base.size != (self.width, self.height):
            base = base.resize((self.width, self.height), Image.Resampling.LANCZOS)
        self._base = base
        dim = Image.blend(
            ImageEnhance.Brightness(base).enhance(0.58),
            Image.new("RGB", base.size, (6, 16, 30)),
            0.18,
        )

        window = tk.Toplevel(root)
        self.window = window
        window.overrideredirect(True)
        window.attributes("-topmost", True)
        try:
            window.attributes("-alpha", 0.99)
        except tk.TclError:
            pass
        window.geometry(f"{self.width}x{self.height}{self._geometry_pos(left, top)}")
        window.configure(bg="#0B1220", cursor="crosshair")

        self.canvas = tk.Canvas(
            window,
            width=self.width,
            height=self.height,
            bg="#0B1220",
            highlightthickness=0,
            cursor="crosshair",
        )
        self.canvas.pack(fill="both", expand=True)
        self._dim_photo = ImageTk.PhotoImage(dim, master=window)
        self.canvas.create_image(0, 0, image=self._dim_photo, anchor="nw")

        self._draw_instruction()
        self._draw_crosshair(0, 0)

        self.selection_rect = self.canvas.create_rectangle(
            0,
            0,
            0,
            0,
            outline="#2DD4BF",
            width=3,
            state="hidden",
        )
        self.selection_inner = self.canvas.create_rectangle(
            0,
            0,
            0,
            0,
            outline="#FFFFFF",
            width=1,
            state="hidden",
        )

        window.bind("<Escape>", self._cancel_event)
        window.bind("<ButtonPress-3>", self._cancel_event)
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Motion>", self._on_motion)
        window.protocol("WM_DELETE_WINDOW", self._cancel_event)
        window.focus_force()
        window.lift()
        try:
            window.grab_set()
        except tk.TclError:
            pass

    @staticmethod
    def _geometry_pos(left: int, top: int) -> str:
        x = f"+{left}" if left >= 0 else str(left)
        y = f"+{top}" if top >= 0 else str(top)
        return f"{x}{y}"

    def _rounded_rect(self, x1, y1, x2, y2, radius=10, **kwargs):
        points = [
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]
        return self.canvas.create_polygon(points, smooth=True, **kwargs)

    def _draw_instruction(self) -> None:
        width = 460
        x1 = max(16, self.width / 2 - width / 2)
        y1 = 26
        x2 = x1 + width
        y2 = y1 + 46
        self._rounded_rect(
            x1,
            y1,
            x2,
            y2,
            radius=12,
            fill="#0F172A",
            outline="#334155",
            width=1,
        )
        self.canvas.create_text(
            x1 + 20,
            (y1 + y2) / 2,
            text="⌗",
            fill="#2DD4BF",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        )
        self.canvas.create_text(
            x1 + 52,
            (y1 + y2) / 2,
            text="拖动框选要识别的文字",
            fill="#F8FAFC",
            font=("Microsoft YaHei UI", 11, "bold"),
            anchor="w",
        )
        self.canvas.create_text(
            x2 - 18,
            (y1 + y2) / 2,
            text="Esc 取消",
            fill="#94A3B8",
            font=("Microsoft YaHei UI", 9),
            anchor="e",
        )

    def _draw_crosshair(self, x: int, y: int) -> None:
        self.canvas.delete("crosshair")
        self.canvas.create_line(
            x, 0, x, self.height, fill="#5EEAD4", dash=(4, 6), tags="crosshair"
        )
        self.canvas.create_line(
            0, y, self.width, y, fill="#5EEAD4", dash=(4, 6), tags="crosshair"
        )
        self.canvas.tag_lower("crosshair")

    def _on_motion(self, event) -> None:
        if self.start is None:
            self._draw_crosshair(event.x, event.y)

    def _on_press(self, event) -> None:
        if event.y < 90 and abs(event.x - self.width / 2) < 260:
            # 顶部提示区不启动选区，避免误拖。
            return
        self.start = (max(0, min(event.x, self.width - 1)), max(0, min(event.y, self.height - 1)))
        self.end = self.start
        self.canvas.itemconfigure(self.selection_rect, state="normal")
        self.canvas.itemconfigure(self.selection_inner, state="normal")
        self.canvas.coords(self.selection_rect, *self.start, *self.start)
        self.canvas.coords(self.selection_inner, *self.start, *self.start)
        if self._tip_text is not None:
            self.canvas.delete(self._tip_text)
        self._tip_text = self.canvas.create_text(
            self.width / 2,
            self.height - 36,
            text="松开鼠标，开始识别",
            fill="#E2E8F0",
            font=("Microsoft YaHei UI", 10),
            tags="tip",
        )
        self.canvas.create_rectangle(
            self.width / 2 - 100,
            self.height - 52,
            self.width / 2 + 100,
            self.height - 20,
            fill="#0F172A",
            outline="",
            tags="tipbg",
        )
        self.canvas.tag_raise("tip")
        self.canvas.tag_raise("tipbg")
        self.canvas.tag_raise("tip")
        self.canvas.tag_raise(self.selection_rect)
        self.canvas.tag_raise(self.selection_inner)

    def _on_drag(self, event) -> None:
        if self.start is None:
            return
        x = max(0, min(event.x, self.width - 1))
        y = max(0, min(event.y, self.height - 1))
        self.end = (x, y)
        self._draw_crosshair(x, y)
        self._update_selection()

    def _update_selection(self) -> None:
        if self.start is None or self.end is None:
            return
        x1, y1 = self.start
        x2, y2 = self.end
        left, right = sorted((x1, x2))
        top, bottom = sorted((y1, y2))
        self.canvas.coords(self.selection_rect, left, top, right, bottom)
        self.canvas.coords(self.selection_inner, left + 2, top + 2, right - 2, bottom - 2)
        w = abs(x2 - x1)
        h = abs(y2 - y1)
        text = f"{w} × {h} px"

        if self._size_rect is not None:
            self.canvas.delete(self._size_rect)
        if self._size_text is not None:
            self.canvas.delete(self._size_text)
        if w < 5 or h < 5:
            return
        label_w = 108
        label_h = 26
        label_x = right - label_w
        if label_x < left:
            label_x = left
        label_y = bottom + 8
        if label_y + label_h > self.height:
            label_y = max(4, top - label_h - 8)
        self._size_rect = self.canvas.create_rectangle(
            label_x,
            label_y,
            label_x + label_w,
            label_y + label_h,
            fill="#0F172A",
            outline="#2DD4BF",
            width=1,
            tags="size",
        )
        self._size_text = self.canvas.create_text(
            label_x + label_w / 2,
            label_y + label_h / 2,
            text=text,
            fill="#F8FAFC",
            font=("Segoe UI", 9, "bold"),
            tags="size",
        )
        self.canvas.tag_raise("size")
        self.canvas.tag_raise(self.selection_rect)
        self.canvas.tag_raise(self.selection_inner)

    def _on_release(self, event) -> None:
        if self.start is None:
            return
        self._on_drag(event)
        if self.end is None:
            return
        x1, y1 = self.start
        x2, y2 = self.end
        left, right = sorted((x1, x2))
        top, bottom = sorted((y1, y2))
        if right - left < 12 or bottom - top < 8:
            self.start = None
            self.end = None
            self.canvas.itemconfigure(self.selection_rect, state="hidden")
            self.canvas.itemconfigure(self.selection_inner, state="hidden")
            self.canvas.delete("size")
            self.canvas.delete("tip")
            self.canvas.delete("tipbg")
            self._size_rect = None
            self._size_text = None
            self._show_too_small()
            return

        # 截图与原图尺寸可能因 DPI 有细微差异，按比例映射。
        scale_x = self.screenshot.width / self.width
        scale_y = self.screenshot.height / self.height
        margin = 8
        crop_box = (
            max(0, int(round(left * scale_x)) - margin),
            max(0, int(round(top * scale_y)) - margin),
            min(self.screenshot.width, int(round(right * scale_x)) + margin),
            min(self.screenshot.height, int(round(bottom * scale_y)) + margin),
        )
        if crop_box[2] - crop_box[0] < 4 or crop_box[3] - crop_box[1] < 4:
            self._show_too_small()
            return
        crop = self.screenshot.crop(crop_box).copy()
        global_rect = (self.left + left, self.top + top, self.left + right, self.top + bottom)
        self._close()
        self.root.after(30, lambda: self.on_done(crop, global_rect))

    def _show_too_small(self) -> None:
        self.canvas.delete("small")
        self.canvas.create_rectangle(
            self.width / 2 - 190,
            self.height / 2 - 28,
            self.width / 2 + 190,
            self.height / 2 + 28,
            fill="#7F1D1D",
            outline="#FCA5A5",
            tags="small",
        )
        self.canvas.create_text(
            self.width / 2,
            self.height / 2,
            text="选区太小，请拖动框选一段文字",
            fill="#FEE2E2",
            font=("Microsoft YaHei UI", 11, "bold"),
            tags="small",
        )
        self.canvas.after(1300, lambda: self.canvas.delete("small"))

    def _cancel_event(self, _event=None) -> None:
        self._close()
        self.root.after(30, self.on_cancel)

    def _close(self) -> None:
        if self.closed:
            return
        self.closed = True
        try:
            self.window.grab_release()
        except Exception:
            pass
        try:
            self.window.destroy()
        except Exception:
            pass
