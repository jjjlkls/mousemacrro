"""
마우스 매크로 프로
- 단일 파일 실행 가능
- PyInstaller로 exe 배포 가능
- 설정 자동 저장 지원
- 좌표/대기시간 기반 매크로 실행
"""

import json
import os
import threading
import time
import tkinter as tk
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from pynput import mouse
import pyautogui

try:
    import keyboard
    KEYBOARD_IMPORT_ERROR = None
except Exception as exc:
    keyboard = None
    KEYBOARD_IMPORT_ERROR = exc


APP_NAME = "마우스 매크로 프로"
APP_VERSION = "1.1.0"
DEFAULT_SETTINGS = {
    "always_on_top": False,
    "countdown": "2",
    "repeat": "1",
    "default_wait": "1.0",
    "move_duration": "0.2",
    "minimize_during_play": True,
    "stop_on_mouse_move": True,
    "ignore_window_clicks": True,
}

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.05


@dataclass
class MacroStep:
    x: int
    y: int
    wait: float = 1.0


class CoordinateDialog(simpledialog.Dialog):
    def __init__(self, parent, title, initial_value=None):
        self.initial_value = initial_value or {"x": "", "y": "", "wait": "1.0"}
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        ttk.Label(master, text="X 좌표:").grid(row=0, column=0, sticky=tk.W, padx=6, pady=6)
        ttk.Label(master, text="Y 좌표:").grid(row=1, column=0, sticky=tk.W, padx=6, pady=6)
        ttk.Label(master, text="대기 시간(초):").grid(row=2, column=0, sticky=tk.W, padx=6, pady=6)

        self.x_entry = ttk.Entry(master, width=20)
        self.y_entry = ttk.Entry(master, width=20)
        self.wait_entry = ttk.Entry(master, width=20)

        self.x_entry.grid(row=0, column=1, padx=6, pady=6)
        self.y_entry.grid(row=1, column=1, padx=6, pady=6)
        self.wait_entry.grid(row=2, column=1, padx=6, pady=6)

        self.x_entry.insert(0, str(self.initial_value.get("x", "")))
        self.y_entry.insert(0, str(self.initial_value.get("y", "")))
        self.wait_entry.insert(0, str(self.initial_value.get("wait", "1.0")))

        ttk.Button(master, text="현재 마우스 위치 가져오기", command=self.fill_current_mouse_position).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 0)
        )
        return self.x_entry

    def fill_current_mouse_position(self):
        x, y = pyautogui.position()
        self.x_entry.delete(0, tk.END)
        self.y_entry.delete(0, tk.END)
        self.x_entry.insert(0, str(x))
        self.y_entry.insert(0, str(y))

    def validate(self):
        try:
            int(self.x_entry.get().strip())
            int(self.y_entry.get().strip())
            wait = float(self.wait_entry.get().strip())
            if wait < 0:
                raise ValueError
            return True
        except ValueError:
            messagebox.showerror("입력 오류", "X/Y는 정수, 대기 시간은 0 이상의 숫자로 입력해주세요.", parent=self)
            return False

    def apply(self):
        self.result = MacroStep(
            x=int(self.x_entry.get().strip()),
            y=int(self.y_entry.get().strip()),
            wait=float(self.wait_entry.get().strip()),
        )


class BulkEditWindow(tk.Toplevel):
    def __init__(self, parent, initial_steps, callback):
        super().__init__(parent)
        self.title("일괄 편집")
        self.geometry("520x450")
        self.callback = callback
        self.original_text = "\n".join([f"{s.x},{s.y},{s.wait}" for s in initial_steps])

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        wrapper = ttk.Frame(self, padding=10)
        wrapper.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            wrapper,
            text="한 줄에 하나씩 입력하세요. 형식: x,y,대기시간\n예: 100,200,1.5",
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 8))

        self.text_widget = tk.Text(wrapper, wrap=tk.NONE, font=("Consolas", 10))
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        self.text_widget.insert("1.0", self.original_text)

        button_frame = ttk.Frame(wrapper)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(button_frame, text="저장", command=self.save_and_close).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(button_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    def parse_steps(self):
        steps = []
        lines = self.text_widget.get("1.0", tk.END).strip().splitlines()
        for index, line in enumerate(lines, start=1):
            line = line.strip()
            if not line:
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) != 3:
                raise ValueError(f"{index}번째 줄: 'x,y,대기시간' 형식이어야 합니다.")
            x, y = int(parts[0]), int(parts[1])
            wait = float(parts[2])
            if wait < 0:
                raise ValueError(f"{index}번째 줄: 대기 시간은 0 이상이어야 합니다.")
            steps.append(MacroStep(x=x, y=y, wait=wait))
        return steps

    def save_and_close(self):
        try:
            steps = self.parse_steps()
        except Exception as e:
            messagebox.showerror("파싱 오류", str(e), parent=self)
            return
        self.callback(steps)
        self.destroy()

    def on_close(self):
        current_text = self.text_widget.get("1.0", tk.END).strip()
        if current_text != self.original_text:
            response = messagebox.askyesnocancel("변경사항 저장", "변경사항을 저장하시겠습니까?", parent=self)
            if response is True:
                self.save_and_close()
            elif response is False:
                self.destroy()
        else:
            self.destroy()


class MouseMacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("760x760")
        self.root.minsize(760, 760)

        self.is_recording = False
        self.is_playing = False
        self.is_countdown = False
        self.is_tracking_coords = False
        self.mouse_listener = None
        self.stop_reason = ""
        self.recorded_steps = []

        self.hotkeys_available = keyboard is not None
        self.f8_hotkey_id = None
        self.f9_hotkey_id = None
        self.esc_hotkey_id = None

        self.settings_path = self.get_settings_path()
        settings = self.load_settings()

        self.always_on_top_var = tk.BooleanVar(value=settings["always_on_top"])
        self.countdown_var = tk.StringVar(value=settings["countdown"])
        self.repeat_var = tk.StringVar(value=settings["repeat"])
        self.default_wait_var = tk.StringVar(value=settings["default_wait"])
        self.move_duration_var = tk.StringVar(value=settings["move_duration"])
        self.minimize_during_play_var = tk.BooleanVar(value=settings["minimize_during_play"])
        self.stop_on_mouse_move_var = tk.BooleanVar(value=settings["stop_on_mouse_move"])
        self.ignore_window_clicks_var = tk.BooleanVar(value=settings["ignore_window_clicks"])
        self.status_var = tk.StringVar(value="대기 중...")
        self.coord_var = tk.StringVar(value="X: -, Y: -")
        self.summary_var = tk.StringVar(value="기록된 동작 없음")

        self._build_ui()
        self.toggle_always_on_top()
        self.register_global_hotkeys()
        self.update_button_states()
        self.refresh_step_list()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        if not self.hotkeys_available:
            self.status_var.set("대기 중... (전역 단축키 비활성화: keyboard 라이브러리 또는 권한 문제)")

    def get_settings_path(self):
        base_dir = Path(os.getenv("APPDATA", str(Path.home()))) / "MouseMacroPro"
        base_dir.mkdir(parents=True, exist_ok=True)
        return base_dir / "settings.json"

    def load_settings(self):
        settings = DEFAULT_SETTINGS.copy()
        try:
            if self.settings_path.exists():
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                if isinstance(saved, dict):
                    settings.update({k: saved.get(k, v) for k, v in settings.items()})
        except Exception:
            pass
        return settings

    def save_settings(self):
        data = {
            "always_on_top": self.always_on_top_var.get(),
            "countdown": self.countdown_var.get(),
            "repeat": self.repeat_var.get(),
            "default_wait": self.default_wait_var.get(),
            "move_duration": self.move_duration_var.get(),
            "minimize_during_play": self.minimize_during_play_var.get(),
            "stop_on_mouse_move": self.stop_on_mouse_move_var.get(),
            "ignore_window_clicks": self.ignore_window_clicks_var.get(),
        }
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Header.TLabel", font=("맑은 고딕", 15, "bold"))
        style.configure("Section.TLabelframe.Label", font=("맑은 고딕", 10, "bold"))

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text=APP_NAME, style="Header.TLabel").pack(anchor="w", pady=(0, 10))

        top_frame = ttk.Frame(main)
        top_frame.pack(fill=tk.X, pady=(0, 8))

        settings_frame = ttk.LabelFrame(top_frame, text="기본 설정", style="Section.TLabelframe")
        settings_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        controls_frame = ttk.LabelFrame(top_frame, text="실행 제어", style="Section.TLabelframe")
        controls_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))

        self._build_settings_ui(settings_frame)
        self._build_controls_ui(controls_frame)

        status_frame = ttk.Frame(main)
        status_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(status_frame, textvariable=self.status_var, relief="sunken", anchor="w", padding=8).pack(fill=tk.X)

        tracking_frame = ttk.LabelFrame(main, text="좌표 추적", style="Section.TLabelframe")
        tracking_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(tracking_frame, text="실시간 좌표 추적 시작/중지", command=self.toggle_coord_tracking).pack(side=tk.LEFT, padx=8, pady=8)
        ttk.Label(tracking_frame, textvariable=self.coord_var, font=("Consolas", 11)).pack(side=tk.LEFT, padx=8)
        ttk.Button(tracking_frame, text="현재 좌표 추가", command=self.add_current_mouse_position).pack(side=tk.RIGHT, padx=8, pady=8)

        list_frame = ttk.LabelFrame(main, text="기록된 동작", style="Section.TLabelframe")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ("idx", "x", "y", "wait")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=14)
        self.tree.heading("idx", text="#")
        self.tree.heading("x", text="X")
        self.tree.heading("y", text="Y")
        self.tree.heading("wait", text="다음 동작까지 대기(초)")
        self.tree.column("idx", width=50, anchor="center")
        self.tree.column("x", width=120, anchor="center")
        self.tree.column("y", width=120, anchor="center")
        self.tree.column("wait", width=180, anchor="center")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=8)
        self.tree.bind("<<TreeviewSelect>>", lambda event: self.update_button_states())
        self.tree.bind("<Double-1>", lambda event: self.edit_action())

        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 8), pady=8)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        bottom = ttk.Frame(main)
        bottom.pack(fill=tk.X)

        left_buttons = ttk.Frame(bottom)
        left_buttons.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.add_btn = ttk.Button(left_buttons, text="좌표 추가", command=self.add_action)
        self.edit_btn = ttk.Button(left_buttons, text="선택 편집", command=self.edit_action)
        self.delete_btn = ttk.Button(left_buttons, text="선택 삭제", command=self.delete_action)
        self.duplicate_btn = ttk.Button(left_buttons, text="선택 복제", command=self.duplicate_action)
        self.up_btn = ttk.Button(left_buttons, text="위로 이동", command=lambda: self.move_selected(-1))
        self.down_btn = ttk.Button(left_buttons, text="아래로 이동", command=lambda: self.move_selected(1))
        self.bulk_btn = ttk.Button(left_buttons, text="일괄 편집", command=self.open_bulk_edit)

        for widget in [self.add_btn, self.edit_btn, self.delete_btn, self.duplicate_btn, self.up_btn, self.down_btn, self.bulk_btn]:
            widget.pack(side=tk.LEFT, padx=3)

        ttk.Label(bottom, textvariable=self.summary_var).pack(side=tk.RIGHT)

    def _build_settings_ui(self, parent):
        grid = ttk.Frame(parent)
        grid.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        ttk.Label(grid, text="기록 전 딜레이(초)").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(grid, textvariable=self.countdown_var, width=10).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(grid, text="반복 횟수").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(grid, textvariable=self.repeat_var, width=10).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(grid, text="새 좌표 기본 대기(초)").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(grid, textvariable=self.default_wait_var, width=10).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(grid, text="마우스 이동 시간(초)").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(grid, textvariable=self.move_duration_var, width=10).grid(row=3, column=1, sticky="w", pady=4)

        ttk.Checkbutton(grid, text="창 항상 위에 표시", variable=self.always_on_top_var, command=self.toggle_always_on_top).grid(row=4, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(grid, text="재생 시 창 최소화", variable=self.minimize_during_play_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(grid, text="재생 중 마우스 움직임 감지 시 중지", variable=self.stop_on_mouse_move_var).grid(row=6, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(grid, text="프로그램 창 내부 클릭은 기록하지 않기", variable=self.ignore_window_clicks_var).grid(row=7, column=0, columnspan=2, sticky="w", pady=4)

    def _build_controls_ui(self, parent):
        body = ttk.Frame(parent)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        start_label = "기록 시작 (F8)" if self.hotkeys_available else "기록 시작"
        play_label = "재생 (F9)" if self.hotkeys_available else "재생"

        self.start_btn = ttk.Button(body, text=start_label, command=self.start_recording)
        self.stop_btn = ttk.Button(body, text="기록 중지", command=self.stop_recording)
        self.play_btn = ttk.Button(body, text=play_label, command=self.play_actions)
        self.reset_btn = ttk.Button(body, text="초기화", command=self.reset)
        self.save_btn = ttk.Button(body, text="저장", command=self.save_records)
        self.load_btn = ttk.Button(body, text="불러오기", command=self.load_records)

        self.start_btn.grid(row=0, column=0, sticky="ew", padx=3, pady=3)
        self.stop_btn.grid(row=0, column=1, sticky="ew", padx=3, pady=3)
        self.play_btn.grid(row=1, column=0, sticky="ew", padx=3, pady=3)
        self.reset_btn.grid(row=1, column=1, sticky="ew", padx=3, pady=3)
        self.save_btn.grid(row=2, column=0, sticky="ew", padx=3, pady=3)
        self.load_btn.grid(row=2, column=1, sticky="ew", padx=3, pady=3)

        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        if self.hotkeys_available:
            help_text = "단축키\n- F8: 기록 시작/중지\n- F9: 재생 시작/중지\n- ESC: 재생 즉시 중지"
        else:
            help_text = "단축키 비활성화\n- keyboard 라이브러리 권한 문제 또는 로드 실패\n- 버튼으로는 정상 사용 가능"
        ttk.Label(body, text=help_text, justify=tk.LEFT).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def register_global_hotkeys(self):
        if keyboard is None:
            self.hotkeys_available = False
            return
        try:
            self.f8_hotkey_id = keyboard.add_hotkey("f8", self.toggle_recording_hotkey)
            self.f9_hotkey_id = keyboard.add_hotkey("f9", self.toggle_play_hotkey)
            self.hotkeys_available = True
        except Exception:
            self.hotkeys_available = False
            self.f8_hotkey_id = None
            self.f9_hotkey_id = None

    def remove_hotkey(self, hotkey_id):
        if keyboard is None or hotkey_id is None:
            return
        try:
            keyboard.remove_hotkey(hotkey_id)
        except Exception:
            pass

    def unregister_global_hotkeys(self):
        self.remove_hotkey(self.f8_hotkey_id)
        self.remove_hotkey(self.f9_hotkey_id)
        self.remove_hotkey(self.esc_hotkey_id)
        self.f8_hotkey_id = None
        self.f9_hotkey_id = None
        self.esc_hotkey_id = None

    def toggle_recording_hotkey(self):
        if self.is_recording:
            self.root.after(0, self.stop_recording)
        elif not self.is_playing and not self.is_countdown:
            self.root.after(0, self.start_recording)

    def toggle_play_hotkey(self):
        if self.is_playing:
            self.root.after(0, self.stop_playing, "F9 키 입력")
        elif self.recorded_steps and not self.is_recording and not self.is_countdown:
            self.root.after(0, self.play_actions)

    def toggle_always_on_top(self):
        self.root.attributes("-topmost", self.always_on_top_var.get())

    def set_status(self, text):
        self.root.after(0, self.status_var.set, text)

    def get_selected_index(self):
        selected = self.tree.selection()
        if not selected:
            return None
        return int(selected[0])

    def update_button_states(self):
        busy = self.is_recording or self.is_countdown or self.is_playing
        has_steps = bool(self.recorded_steps)
        selected_index = self.get_selected_index()

        self.start_btn.config(state=tk.DISABLED if busy else tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL if self.is_recording else tk.DISABLED)
        self.play_btn.config(state=tk.NORMAL if has_steps else tk.DISABLED)
        self.reset_btn.config(state=tk.DISABLED if busy else tk.NORMAL)
        self.save_btn.config(state=tk.DISABLED if busy or not has_steps else tk.NORMAL)
        self.load_btn.config(state=tk.DISABLED if busy else tk.NORMAL)

        action_state = tk.DISABLED if busy else tk.NORMAL
        self.add_btn.config(state=action_state)
        self.bulk_btn.config(state=action_state)
        self.edit_btn.config(state=tk.NORMAL if selected_index is not None and not busy else tk.DISABLED)
        self.delete_btn.config(state=tk.NORMAL if selected_index is not None and not busy else tk.DISABLED)
        self.duplicate_btn.config(state=tk.NORMAL if selected_index is not None and not busy else tk.DISABLED)
        self.up_btn.config(state=tk.NORMAL if selected_index not in (None, 0) and not busy else tk.DISABLED)
        self.down_btn.config(state=tk.NORMAL if selected_index is not None and selected_index < len(self.recorded_steps) - 1 and not busy else tk.DISABLED)

        if self.is_playing:
            self.play_btn.config(text="재생 중지 (ESC/F9)" if self.hotkeys_available else "재생 중지", command=lambda: self.stop_playing("사용자 요청"))
        else:
            self.play_btn.config(text="재생 (F9)" if self.hotkeys_available else "재생", command=self.play_actions)

    def start_recording(self):
        if self.recorded_steps:
            ok = messagebox.askokcancel("기록 시작", "기존 기록을 지우고 새로 기록할까요?")
            if not ok:
                return
            self.recorded_steps.clear()
            self.refresh_step_list()

        try:
            delay = max(0, int(float(self.countdown_var.get().strip())))
        except ValueError:
            messagebox.showerror("입력 오류", "기록 전 딜레이는 정수로 입력해주세요.")
            return

        self.is_countdown = True
        self.update_button_states()
        if delay > 0:
            self._countdown(delay)
        else:
            self.is_countdown = False
            self._start_actual_recording()

    def _countdown(self, seconds_left):
        if not self.is_countdown:
            return
        if seconds_left > 0:
            self.status_var.set(f"기록 시작까지 {seconds_left}초...")
            self.root.after(1000, self._countdown, seconds_left - 1)
        else:
            self.is_countdown = False
            self._start_actual_recording()

    def _start_actual_recording(self):
        self.is_recording = True
        self.update_button_states()
        self.status_var.set("기록 중... 좌클릭으로 좌표가 추가됩니다.")
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()

    def stop_recording(self):
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
            except Exception:
                pass
            self.mouse_listener = None
        self.is_recording = False
        self.refresh_step_list()
        self.status_var.set(f"기록 완료. 총 {len(self.recorded_steps)}개 동작이 저장되었습니다.")

    def play_actions(self):
        if not self.recorded_steps:
            messagebox.showwarning("알림", "재생할 동작이 없습니다.")
            return

        try:
            repeats = int(self.repeat_var.get().strip())
            if repeats < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "반복 횟수는 1 이상의 정수여야 합니다.")
            return

        try:
            move_duration = float(self.move_duration_var.get().strip())
            if move_duration < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "마우스 이동 시간은 0 이상의 숫자여야 합니다.")
            return

        self.is_playing = True
        self.stop_reason = ""
        self.update_button_states()

        if self.minimize_during_play_var.get():
            self.root.iconify()

        threading.Thread(target=self._input_watcher, daemon=True).start()
        threading.Thread(target=self._playback_worker, args=(repeats, move_duration), daemon=True).start()

    def stop_playing(self, reason="사용자 요청"):
        if self.is_playing:
            self.is_playing = False
            self.stop_reason = reason

    def _input_watcher(self):
        self.remove_hotkey(self.esc_hotkey_id)
        self.esc_hotkey_id = None
        if keyboard is not None and self.hotkeys_available:
            try:
                self.esc_hotkey_id = keyboard.add_hotkey("esc", lambda: self.stop_playing("ESC 키 입력"))
            except Exception:
                self.esc_hotkey_id = None

        last_pos = pyautogui.position()
        movement_start_time = None

        while self.is_playing:
            if self.stop_on_mouse_move_var.get():
                current_pos = pyautogui.position()
                if current_pos != last_pos:
                    if movement_start_time is None:
                        movement_start_time = time.time()
                    elif time.time() - movement_start_time >= 1.0:
                        self.stop_playing("마우스 움직임 감지")
                        break
                else:
                    movement_start_time = None
                last_pos = current_pos
            time.sleep(0.1)

        self.remove_hotkey(self.esc_hotkey_id)
        self.esc_hotkey_id = None

    def _playback_worker(self, repeats, move_duration):
        try:
            for repeat_index in range(repeats):
                if not self.is_playing:
                    break
                for step_index, step in enumerate(self.recorded_steps, start=1):
                    if not self.is_playing:
                        break
                    self.set_status(f"재생 중... {repeat_index + 1}/{repeats}회, {step_index}/{len(self.recorded_steps)}번째 동작")
                    pyautogui.moveTo(step.x, step.y, duration=move_duration)
                    pyautogui.click(step.x, step.y)

                    wait_until = time.time() + step.wait
                    while self.is_playing and time.time() < wait_until:
                        time.sleep(0.05)
        except Exception as e:
            self.stop_reason = f"오류: {e}"
        finally:
            self.is_playing = False
            self.root.after(0, self.on_playback_finished)

    def on_playback_finished(self):
        if self.stop_reason:
            self.status_var.set(f"재생이 중지되었습니다. ({self.stop_reason})")
            self.stop_reason = ""
        else:
            self.status_var.set("재생 완료!")
        self.update_button_states()

    def reset(self):
        self.recorded_steps.clear()
        self.refresh_step_list()
        self.status_var.set("대기 중...")

    def add_current_mouse_position(self):
        try:
            wait = float(self.default_wait_var.get().strip())
            if wait < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "기본 대기 시간은 0 이상의 숫자여야 합니다.")
            return

        x, y = pyautogui.position()
        self.recorded_steps.append(MacroStep(x=x, y=y, wait=wait))
        self.refresh_step_list(select_index=len(self.recorded_steps) - 1)
        self.status_var.set(f"현재 좌표 ({x}, {y})가 추가되었습니다.")

    def add_action(self):
        try:
            default_wait = float(self.default_wait_var.get().strip())
            if default_wait < 0:
                raise ValueError
        except ValueError:
            default_wait = 1.0

        dialog = CoordinateDialog(self.root, "새 동작 추가", initial_value={"x": "", "y": "", "wait": str(default_wait)})
        if dialog.result:
            self.recorded_steps.append(dialog.result)
            self.refresh_step_list(select_index=len(self.recorded_steps) - 1)

    def edit_action(self):
        index = self.get_selected_index()
        if index is None:
            return
        step = self.recorded_steps[index]
        dialog = CoordinateDialog(self.root, "동작 편집", initial_value={"x": step.x, "y": step.y, "wait": step.wait})
        if dialog.result:
            self.recorded_steps[index] = dialog.result
            self.refresh_step_list(select_index=index)

    def duplicate_action(self):
        index = self.get_selected_index()
        if index is None:
            return
        step = self.recorded_steps[index]
        self.recorded_steps.insert(index + 1, MacroStep(step.x, step.y, step.wait))
        self.refresh_step_list(select_index=index + 1)
        self.status_var.set("선택한 동작을 복제했습니다.")

    def delete_action(self):
        index = self.get_selected_index()
        if index is None:
            return
        step = self.recorded_steps[index]
        if messagebox.askyesno("삭제 확인", f"({step.x}, {step.y}) 동작을 삭제하시겠습니까?"):
            self.recorded_steps.pop(index)
            new_index = min(index, len(self.recorded_steps) - 1) if self.recorded_steps else None
            self.refresh_step_list(select_index=new_index)

    def move_selected(self, direction):
        index = self.get_selected_index()
        if index is None:
            return
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.recorded_steps):
            return
        self.recorded_steps[index], self.recorded_steps[new_index] = self.recorded_steps[new_index], self.recorded_steps[index]
        self.refresh_step_list(select_index=new_index)

    def open_bulk_edit(self):
        BulkEditWindow(self.root, self.recorded_steps, self.update_steps_from_bulk)

    def update_steps_from_bulk(self, steps):
        self.recorded_steps = steps
        self.refresh_step_list()
        self.status_var.set(f"일괄 편집 완료. 총 {len(self.recorded_steps)}개 동작")

    def save_records(self):
        if not self.recorded_steps:
            messagebox.showwarning("알림", "저장할 동작이 없습니다.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
            title="매크로 저장",
        )
        if not filepath:
            return

        try:
            if filepath.lower().endswith(".txt"):
                with open(filepath, "w", encoding="utf-8") as f:
                    for step in self.recorded_steps:
                        f.write(f"{step.x},{step.y},{step.wait}\n")
            else:
                payload = {"version": 2, "steps": [asdict(step) for step in self.recorded_steps]}
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"{os.path.basename(filepath)} 파일로 저장했습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"저장 중 오류가 발생했습니다.\n\n{e}")

    def load_records(self):
        if self.recorded_steps:
            ok = messagebox.askokcancel("불러오기", "현재 목록을 지우고 새 파일을 불러올까요?")
            if not ok:
                return

        filepath = filedialog.askopenfilename(
            filetypes=[("지원 파일", "*.json;*.txt"), ("JSON 파일", "*.json"), ("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
            title="매크로 불러오기",
        )
        if not filepath:
            return

        try:
            steps = []
            if filepath.lower().endswith(".json"):
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for item in data.get("steps", []):
                    steps.append(MacroStep(x=int(item["x"]), y=int(item["y"]), wait=float(item.get("wait", 1.0))))
            else:
                with open(filepath, "r", encoding="utf-8") as f:
                    for line_no, line in enumerate(f, start=1):
                        line = line.strip()
                        if not line:
                            continue
                        parts = [p.strip() for p in line.split(",")]
                        if len(parts) == 2:
                            x, y = int(parts[0]), int(parts[1])
                            wait = 1.0
                        elif len(parts) == 3:
                            x, y = int(parts[0]), int(parts[1])
                            wait = float(parts[2])
                        else:
                            raise ValueError(f"{line_no}번째 줄 형식이 잘못되었습니다.")
                        steps.append(MacroStep(x=x, y=y, wait=wait))

            self.recorded_steps = steps
            self.refresh_step_list()
            self.status_var.set(f"{os.path.basename(filepath)} 파일에서 {len(steps)}개 동작을 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("불러오기 오류", f"불러오는 중 오류가 발생했습니다.\n\n{e}")

    def toggle_coord_tracking(self):
        self.is_tracking_coords = not self.is_tracking_coords
        if self.is_tracking_coords:
            self.status_var.set("실시간 좌표 추적 시작")
            self.update_coord_label()
        else:
            self.status_var.set("실시간 좌표 추적 중지")

    def update_coord_label(self):
        if not self.is_tracking_coords:
            return
        try:
            x, y = pyautogui.position()
            self.coord_var.set(f"X: {x}, Y: {y}")
        except Exception as e:
            self.coord_var.set(f"좌표 읽기 실패: {e}")
            self.is_tracking_coords = False
            return
        self.root.after(100, self.update_coord_label)

    def on_click(self, x, y, button, pressed):
        if not (self.is_recording and pressed and button == mouse.Button.left):
            return

        if self.ignore_window_clicks_var.get() and self.is_point_inside_window(x, y):
            return

        try:
            wait = float(self.default_wait_var.get().strip())
            if wait < 0:
                raise ValueError
        except ValueError:
            wait = 1.0

        self.recorded_steps.append(MacroStep(x=x, y=y, wait=wait))
        self.root.after(0, self.refresh_step_list, len(self.recorded_steps) - 1)

    def is_point_inside_window(self, x, y):
        try:
            left = self.root.winfo_rootx()
            top = self.root.winfo_rooty()
            right = left + self.root.winfo_width()
            bottom = top + self.root.winfo_height()
            return left <= x <= right and top <= y <= bottom
        except Exception:
            return False

    def refresh_step_list(self, select_index=None):
        self.tree.delete(*self.tree.get_children())
        total_wait = 0.0
        for index, step in enumerate(self.recorded_steps):
            total_wait += step.wait
            self.tree.insert("", tk.END, iid=str(index), values=(index + 1, step.x, step.y, f"{step.wait:.2f}"))
        if select_index is not None and 0 <= select_index < len(self.recorded_steps):
            self.tree.selection_set(str(select_index))
            self.tree.focus(str(select_index))
        self.summary_var.set(f"총 {len(self.recorded_steps)}개 / 누적 대기 {total_wait:.1f}초")
        self.update_button_states()

    def on_close(self):
        self.save_settings()
        if self.is_playing:
            self.stop_playing("프로그램 종료")
        if self.is_recording:
            self.stop_recording()
        self.unregister_global_hotkeys()
        if keyboard is not None:
            try:
                keyboard.unhook_all()
            except Exception:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    MouseMacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            detail = f"프로그램 실행 중 오류가 발생했습니다.\n\n{e}"
            if KEYBOARD_IMPORT_ERROR is not None:
                detail += f"\n\nkeyboard 로드 오류: {KEYBOARD_IMPORT_ERROR}"
            messagebox.showerror("오류", detail)
        except Exception:
            print(f"프로그램 실행 중 오류가 발생했습니다: {e}")
