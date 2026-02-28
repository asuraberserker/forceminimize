import ctypes
import sys
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Dict, List, Optional

import psutil
import win32con
import win32gui
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


@dataclass
class WindowInfo:
    hwnd: int
    title: str
    pid: int
    process_name: str

    @property
    def label(self) -> str:
        return f"{self.title}  ({self.process_name}, PID={self.pid}, HWND={self.hwnd})"


class WindowGuardApp:
    CHECK_INTERVAL_SECONDS = 0.5
    BG_COLOR = "#1f1f1f"
    PANEL_COLOR = "#252526"
    FG_COLOR = "#f3f3f3"
    INPUT_BG_COLOR = "#2d2d30"
    INPUT_FG_COLOR = "#f3f3f3"
    ACCENT_COLOR = "#3a3d41"
    CONFIG_FILE = Path(__file__).with_name("config.txt")

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("窗口守护（最小化 + 静音）")
        self.root.geometry("550x250")

        self.windows: List[WindowInfo] = []
        self.label_to_window: Dict[str, WindowInfo] = {}
        self.selected_window: Optional[WindowInfo] = None
        self.running = False
        self.last_muted_state: Optional[bool] = None
        self.saved_process_name = self._load_saved_process_name()

        self._apply_dark_theme()
        self._build_ui()
        self._refresh_windows()

    def _apply_dark_theme(self) -> None:
        self.root.configure(bg=self.BG_COLOR)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.configure("TFrame", background=self.BG_COLOR)
        style.configure("TLabel", background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.configure(
            "TButton",
            background=self.ACCENT_COLOR,
            foreground=self.FG_COLOR,
            borderwidth=0,
            focusthickness=0,
            padding=(10, 6),
        )
        style.map(
            "TButton",
            background=[("active", "#4a4d52"), ("disabled", "#2c2c2c")],
            foreground=[("disabled", "#8b8b8b")],
        )
        style.configure(
            "TCombobox",
            fieldbackground=self.INPUT_BG_COLOR,
            background=self.INPUT_BG_COLOR,
            foreground=self.INPUT_FG_COLOR,
            arrowcolor=self.FG_COLOR,
            selectbackground="#3c3c3c",
            selectforeground=self.FG_COLOR,
        )
        style.map("TCombobox", fieldbackground=[("readonly", self.INPUT_BG_COLOR)])

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self.root, padding=12, style="Card.TFrame")
        top_frame.pack(fill="both", expand=True)

        ttk.Style(self.root).configure("Card.TFrame", background=self.BG_COLOR)

        instruction = (
            "1) 点击“刷新列表”选择要守护的窗口。\n"
            "2) 点击“开始守护”。\n"
            "3) 当该窗口失去焦点时：程序会将其最小化并静音。\n"
            "4) 当切回该窗口时：程序会尝试恢复并取消静音。"
        )
        ttk.Label(top_frame, text=instruction, justify="left").pack(anchor="w", pady=(0, 10))

        select_row = ttk.Frame(top_frame, style="Card.TFrame")
        select_row.pack(fill="x", pady=6)

        self.window_combobox = ttk.Combobox(select_row, state="readonly")
        self.window_combobox.pack(side="left", fill="x", expand=True, padx=(0, 8))

        button_row = ttk.Frame(top_frame, style="Card.TFrame")
        button_row.pack(fill="x", pady=(8, 10))

        self.refresh_btn = ttk.Button(button_row, text="刷新列表", command=self._refresh_windows)
        self.refresh_btn.pack(side="left")

        self.start_btn = ttk.Button(button_row, text="开始守护", command=self._start_guard)
        self.start_btn.pack(side="left", padx=(8, 0))

        self.stop_btn = ttk.Button(button_row, text="停止守护", command=self._stop_guard, state="disabled")
        self.stop_btn.pack(side="left", padx=8)

        self.status_var = tk.StringVar(value="状态：未启动")
        ttk.Label(top_frame, textvariable=self.status_var).pack(anchor="w", pady=(4, 0))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _refresh_windows(self) -> None:
        windows = self._enumerate_windows()
        self.windows = windows

        labels = [w.label for w in windows]
        self.label_to_window = {w.label: w for w in windows}
        self.window_combobox["values"] = labels

        if labels:
            saved_index = self._find_saved_process_index(windows)
            self.window_combobox.current(saved_index if saved_index is not None else 0)
            self.status_var.set(f"状态：已加载 {len(labels)} 个窗口")
        else:
            self.window_combobox.set("")
            self.status_var.set("状态：未找到可选择窗口")

    def _find_saved_process_index(self, windows: List[WindowInfo]) -> Optional[int]:
        if not self.saved_process_name:
            return None

        for index, window in enumerate(windows):
            if window.process_name == self.saved_process_name:
                return index
        return None

    def _load_saved_process_name(self) -> Optional[str]:
        try:
            value = self.CONFIG_FILE.read_text(encoding="utf-8").strip()
            return value or None
        except OSError:
            return None

    def _save_process_name(self, process_name: str) -> None:
        try:
            self.CONFIG_FILE.write_text(process_name, encoding="utf-8")
        except OSError:
            messagebox.showwarning("提示", "保存 config.txt 失败，已忽略。")

    def _start_guard(self) -> None:
        label = self.window_combobox.get().strip()
        if not label:
            messagebox.showwarning("提示", "请先选择一个窗口")
            return

        selected = self.label_to_window.get(label)
        if not selected:
            messagebox.showerror("错误", "当前选择无效，请刷新后重试")
            return

        if not win32gui.IsWindow(selected.hwnd):
            messagebox.showerror("错误", "该窗口已不存在，请刷新后重试")
            return

        if self.saved_process_name != selected.process_name:
            self.saved_process_name = selected.process_name
            self._save_process_name(selected.process_name)

        self.selected_window = selected
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set(f"状态：守护中 -> {selected.title}")
        self._schedule_tick()

    def _stop_guard(self) -> None:
        if self.selected_window:
            self._set_process_mute(self.selected_window.pid, mute=False)

        self.running = False
        self.last_muted_state = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_var.set("状态：已停止")

    def _schedule_tick(self) -> None:
        if not self.running:
            return
        self._guard_tick()
        self.root.after(int(self.CHECK_INTERVAL_SECONDS * 1000), self._schedule_tick)

    def _guard_tick(self) -> None:
        target = self.selected_window
        if not target:
            return

        if not win32gui.IsWindow(target.hwnd):
            self.status_var.set("状态：目标窗口已关闭，已停止守护")
            self._stop_guard()
            return

        foreground = win32gui.GetForegroundWindow()
        is_target_active = foreground == target.hwnd

        if is_target_active:
            self._restore_window(target.hwnd)
            self._set_process_mute(target.pid, mute=False)
            self.last_muted_state = False
            self.status_var.set(f"状态：目标窗口激活（已取消静音） -> {target.title}")
        else:
            self._minimize_window(target.hwnd)
            self._set_process_mute(target.pid, mute=True)
            self.last_muted_state = True
            self.status_var.set(f"状态：目标窗口非激活（已最小化+静音） -> {target.title}")

    def _set_process_mute(self, pid: int, mute: bool) -> None:
        # 只在状态改变时执行，避免重复调用
        if self.last_muted_state is mute:
            return

        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            process = session.Process
            if process is None:
                continue
            if process.pid != pid:
                continue

            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(1 if mute else 0, None)

    @staticmethod
    def _restore_window(hwnd: int) -> None:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    @staticmethod
    def _minimize_window(hwnd: int) -> None:
        if not win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

    def _on_close(self) -> None:
        self.running = False
        self.root.destroy()

    @staticmethod
    def _enumerate_windows() -> List[WindowInfo]:
        windows: List[WindowInfo] = []

        def callback(hwnd: int, _: int) -> bool:
            if not win32gui.IsWindowVisible(hwnd):
                return True

            title = win32gui.GetWindowText(hwnd).strip()
            if not title:
                return True

            if win32gui.GetWindow(hwnd, win32con.GW_OWNER) != 0:
                return True

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == 0:
                return True

            try:
                process_name = psutil.Process(pid).name()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                process_name = "unknown"

            windows.append(WindowInfo(hwnd=hwnd, title=title, pid=pid, process_name=process_name))
            return True

        win32gui.EnumWindows(callback, 0)

        # 按窗口标题排序，便于查找
        windows.sort(key=lambda item: item.title.lower())
        return windows


def ensure_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("该程序仅支持 Windows 运行。")

    # 确保高 DPI 下窗口显示正常
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def main() -> None:
    ensure_windows()
    root = tk.Tk()
    app = WindowGuardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
