import ctypes
import json
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


class SilenceModeApp:
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
        self.root.title("程序自动静默")
        self.root.geometry("550x200")

        self.windows: List[WindowInfo] = []
        self.label_to_window: Dict[str, WindowInfo] = {}
        self.selected_window: Optional[WindowInfo] = None
        self.running = False
        self.last_muted_state: Optional[bool] = None
        config = self._load_config()
        self.saved_process_name = config.get("process_name")
        self.minimize_enabled = config.get("minimize_enabled", True)
        self.mute_enabled = config.get("mute_enabled", True)

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


        select_row = ttk.Frame(top_frame, style="Card.TFrame")
        select_row.pack(fill="x", pady=6)

        self.window_combobox = ttk.Combobox(select_row, state="readonly")
        self.window_combobox.pack(side="left", fill="x", expand=True, padx=(0, 8))

        button_row = ttk.Frame(top_frame, style="Card.TFrame")
        button_row.pack(fill="x", pady=(8, 4))

        self.refresh_btn = ttk.Button(button_row, text="刷新列表", command=self._refresh_windows)
        self.refresh_btn.pack(side="left")

        self.start_btn = ttk.Button(button_row, text="开始静默", command=self._start_silence)
        self.start_btn.pack(side="left", padx=(8, 0))

        self.stop_btn = ttk.Button(button_row, text="停止静默", command=self._stop_silence, state="disabled")
        self.stop_btn.pack(side="left", padx=8)

        # 开关行
        switch_row = ttk.Frame(top_frame, style="Card.TFrame")
        switch_row.pack(fill="x", pady=(0, 6))

        self.minimize_var = tk.BooleanVar(value=self.minimize_enabled)
        self.mute_var = tk.BooleanVar(value=self.mute_enabled)

        self.minimize_cb = ttk.Checkbutton(
            switch_row,
            text="最小化",
            variable=self.minimize_var,
            command=self._on_minimize_toggle
        )
        self.minimize_cb.pack(side="left", padx=(0, 12))

        self.mute_cb = ttk.Checkbutton(
            switch_row,
            text="静音",
            variable=self.mute_var,
            command=self._on_mute_toggle
        )
        self.mute_cb.pack(side="left")

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

    def _load_config(self) -> dict:
        """加载配置文件，返回包含 process_name, minimize_enabled, mute_enabled 的字典。
        兼容旧版纯文本格式（仅包含进程名）。"""
        try:
            content = self.CONFIG_FILE.read_text(encoding="utf-8").strip()
            if not content:
                return {}
            # 尝试解析为 JSON
            try:
                config = json.loads(content)
                # 确保必需的字段存在
                result = {}
                if "process_name" in config:
                    result["process_name"] = config["process_name"]
                result["minimize_enabled"] = config.get("minimize_enabled", True)
                result["mute_enabled"] = config.get("mute_enabled", True)
                return result
            except json.JSONDecodeError:
                # 旧版格式：纯文本进程名
                return {"process_name": content, "minimize_enabled": True, "mute_enabled": True}
        except OSError:
            return {}

    def _save_config(self, process_name: Optional[str] = None) -> None:
        """保存配置到文件。如果提供了 process_name 则更新，否则保留现有值。"""
        config = self._load_config()
        if process_name is not None:
            config["process_name"] = process_name
        # 更新当前开关状态
        config["minimize_enabled"] = self.minimize_enabled
        config["mute_enabled"] = self.mute_enabled

        try:
            self.CONFIG_FILE.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError:
            messagebox.showwarning("提示", "保存 config.txt 失败，已忽略。")

    def _start_silence(self) -> None:
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
            self._save_config(selected.process_name)

        self.selected_window = selected
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set(f"状态：静默中 -> {selected.title}")
        self._schedule_tick()

    def _stop_silence(self) -> None:
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
        self._silence_tick()
        self.root.after(int(self.CHECK_INTERVAL_SECONDS * 1000), self._schedule_tick)

    def _silence_tick(self) -> None:
        target = self.selected_window
        if not target:
            return

        if not win32gui.IsWindow(target.hwnd):
            self.status_var.set("状态：目标窗口已关闭，已停止静默")
            self._stop_silence()
            return

        foreground = win32gui.GetForegroundWindow()
        is_target_active = foreground == target.hwnd

        if is_target_active:
            if self.minimize_enabled:
                self._restore_window(target.hwnd)
            if self.mute_enabled:
                self._set_process_mute(target.pid, mute=False)
                self.last_muted_state = False
            # 更新状态文本
            actions = []
            if self.minimize_enabled:
                actions.append("已恢复窗口")
            if self.mute_enabled:
                actions.append("已取消静音")
            status_action = "、".join(actions) if actions else "已忽略操作"
            self.status_var.set(f"状态：目标窗口激活（{status_action}） -> {target.title}")
        else:
            if self.minimize_enabled:
                self._minimize_window(target.hwnd)
            if self.mute_enabled:
                self._set_process_mute(target.pid, mute=True)
                self.last_muted_state = True
            # 更新状态文本
            actions = []
            if self.minimize_enabled:
                actions.append("已最小化")
            if self.mute_enabled:
                actions.append("已静音")
            status_action = "、".join(actions) if actions else "已忽略操作"
            self.status_var.set(f"状态：目标窗口非激活（{status_action}） -> {target.title}")

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

    def _on_minimize_toggle(self) -> None:
        self.minimize_enabled = self.minimize_var.get()
        self._save_config()

    def _on_mute_toggle(self) -> None:
        self.mute_enabled = self.mute_var.get()
        self._save_config()

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
    app = SilenceModeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
