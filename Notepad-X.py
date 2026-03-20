import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import sys
import ctypes
import json
import bisect
import re
import hashlib
import secrets
import tempfile
import traceback
import time
from datetime import datetime
from ctypes import wintypes

try:
    from idlelib.colorizer import ColorDelegator
    from idlelib.percolator import Percolator
except ImportError:
    ColorDelegator = None
    Percolator = None

class NotepadX:
    def __init__(self, isolated_session=False):
        self.root = tk.Tk()
        self.root.title("Notepad-X")
        self.resource_dir = self.get_resource_dir()
        self.app_dir = self.get_app_dir()
        self.isolated_session = isolated_session
        self.icon_path = self.resolve_gfx_path("Notepad-X.ico")
        self.splash_path = self.resolve_gfx_path("splash.png")
        self.splash_max_width = 430
        self.splash_max_height = 645
        self.note_sound_path = self.resolve_audio_path("note.mp3")
        self.delete_note_sound_path = self.resolve_audio_path("delete_note.mp3")
        
        # Try to set window icon
        if os.path.exists(self.icon_path):
            self.root.iconbitmap(self.icon_path)
        
        self.root.geometry("1500x700")
        self.root.configure(bg='#1e1e1e')
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)

        # Grid layout
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Dark theme colors
        self.bg_color     = '#1e1e1e'
        self.fg_color     = '#d4d4d4'
        self.text_bg      = '#0d1117'
        self.text_fg      = '#c9d1d9'
        self.cursor_color = '#58a6ff'
        self.select_bg    = '#264f78'
        self.match_bg     = '#e3b505'    # search highlight
        self.note_bg      = '#f2cc60'
        self.panel_bg     = '#252525'

        self.current_file = None
        self.find_matches_tag = 'find_match'
        self.find_current_tag = 'find_current'
        self.documents = {}
        self.memory_used_mb = 0
        self.syntax_highlighting_available = ColorDelegator is not None and Percolator is not None
        self.large_file_threshold_bytes = 5 * 1024 * 1024
        self.file_load_chunk_size = 256 * 1024
        self.huge_file_preview_threshold_bytes = 100 * 1024 * 1024
        self.huge_file_preview_bytes = 2 * 1024 * 1024
        self.virtual_file_window_lines = 5000
        self.virtual_file_margin_lines = 800
        self.session_path = os.path.join(self.app_dir, "Notepad-X.session.json")
        if self.isolated_session:
            self.session_path = os.path.join(self.app_dir, f"Notepad-X.{os.getpid()}.session.json")
        self.editor_identity_path = os.path.join(self.app_dir, "Notepad-X.editor.json")
        if self.isolated_session:
            self.editor_identity_path = os.path.join(self.app_dir, f"Notepad-X.{os.getpid()}.editor.json")
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.crash_log_path = os.path.join(self.app_dir, "Notepad-X.crash.log")
        self.help_path = os.path.join(self.resource_dir, "Notepad-X-help.txt")
        self.max_recent_files = 10
        self.recent_files = []
        self.closed_session_files = set()
        self.note_sync_interval_ms = 2000
        self.kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        self.psapi = ctypes.WinDLL('psapi', use_last_error=True)
        self.configure_memory_api()
        self.configure_sound_api()
        self.known_editor_ids = self.load_known_editor_ids()
        self.editor_id = self.generate_editor_id()
        self.editor_aliases = set(self.known_editor_ids)
        self.editor_aliases.add(self.editor_id)
        self.persist_editor_identity()

        # Font size control
        self.base_font_size = 13
        self.current_font_size = self.base_font_size
        self.font_family = 'Courier New'
        self.min_font_size = 6
        self.max_font_size = 32
        self.word_wrap_enabled = tk.BooleanVar(value=True)
        self.sound_enabled = tk.BooleanVar(value=True)
        self.status_bar_enabled = tk.BooleanVar(value=True)
        self.numbered_lines_enabled = tk.BooleanVar(value=True)
        self.autocomplete_enabled = tk.BooleanVar(value=True)
        self.search_all_tabs = tk.BooleanVar(value=False)
        self.note_filter = tk.StringVar(value='all')
        self.syntax_theme = tk.StringVar(value='Default')
        self.syntax_mode_selection = tk.StringVar(value='auto')
        self.recovery_job = None
        self.compare_active = False
        self.compare_source_tab = None
        self.compare_refresh_job = None
        self.compare_view = None
        self.last_active_editor_widget = None
        self.toast_popup = None
        self.toast_after_id = None
        self._last_gutter_copy = None
        self.autocomplete_popup = None
        self.autocomplete_listbox = None
        self.autocomplete_target = None
        self.autocomplete_target_doc = None
        self.autocomplete_prefix_start = None

        # Panel visibility flags
        self.find_panel_visible = False
        self.replace_panel_visible = False
        self.fullscreen = False
        self.fullscreen_panel_restore = False

        self.setup_exception_handling()

        # Create UI in logical order
        self.create_text_widget()
        self.create_bottom_panels()
        self.create_menu()
        self.create_status_bar()
        self.restore_session()
        self.restore_recovery_state()

        self.bind_keys()
        self.update_font()  # initial font
        self.update_memory_usage()
        self.poll_shared_notes()
        self.center_window(self.root)

        self.root.mainloop()

    class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
        _fields_ = [
            ('cb', wintypes.DWORD),
            ('PageFaultCount', wintypes.DWORD),
            ('PeakWorkingSetSize', ctypes.c_size_t),
            ('WorkingSetSize', ctypes.c_size_t),
            ('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
            ('QuotaPagedPoolUsage', ctypes.c_size_t),
            ('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
            ('QuotaNonPagedPoolUsage', ctypes.c_size_t),
            ('PagefileUsage', ctypes.c_size_t),
            ('PeakPagefileUsage', ctypes.c_size_t),
        ]

    def configure_memory_api(self):
        self.kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        self.kernel32.GetCurrentProcess.argtypes = []
        self.psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
        self.psapi.GetProcessMemoryInfo.argtypes = [
            wintypes.HANDLE,
            ctypes.c_void_p,
            wintypes.DWORD
        ]
        self.kernel32.GetFileAttributesW.restype = wintypes.DWORD
        self.kernel32.GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
        self.kernel32.SetFileAttributesW.restype = wintypes.BOOL
        self.kernel32.SetFileAttributesW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]

    def configure_sound_api(self):
        try:
            self.winmm = ctypes.WinDLL('winmm')
            self.winmm.mciSendStringW.restype = wintypes.UINT
            self.winmm.mciSendStringW.argtypes = [
                wintypes.LPCWSTR,
                wintypes.LPWSTR,
                wintypes.UINT,
                wintypes.HANDLE
            ]
        except Exception:
            self.winmm = None

    def get_resource_dir(self):
        if getattr(sys, 'frozen', False):
            return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
        return os.path.dirname(__file__)

    def get_user_support_dir(self):
        base_dir = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        return os.path.join(base_dir, 'Notepad-X')

    def directory_is_writable(self, directory):
        try:
            os.makedirs(directory, exist_ok=True)
            fd, temp_path = tempfile.mkstemp(prefix='notepadx-write-test-', suffix='.tmp', dir=directory)
            os.close(fd)
            os.remove(temp_path)
            return True
        except OSError:
            return False

    def get_app_dir(self):
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            if self.directory_is_writable(exe_dir):
                return exe_dir
            fallback_dir = self.get_user_support_dir()
            os.makedirs(fallback_dir, exist_ok=True)
            return fallback_dir
        return os.path.dirname(__file__)

    def move_support_paths_to_user_dir(self):
        fallback_dir = self.get_user_support_dir()
        os.makedirs(fallback_dir, exist_ok=True)
        self.app_dir = fallback_dir
        self.session_path = os.path.join(self.app_dir, "Notepad-X.session.json")
        if self.isolated_session:
            self.session_path = os.path.join(self.app_dir, f"Notepad-X.{os.getpid()}.session.json")
        self.editor_identity_path = os.path.join(self.app_dir, "Notepad-X.editor.json")
        if self.isolated_session:
            self.editor_identity_path = os.path.join(self.app_dir, f"Notepad-X.{os.getpid()}.editor.json")
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.crash_log_path = os.path.join(self.app_dir, "Notepad-X.crash.log")

    def resolve_gfx_path(self, filename):
        base_dir = self.resource_dir
        candidates = [
            os.path.join(base_dir, "gfx", filename),
            os.path.join(base_dir, filename),
        ]
        for candidate in candidates:
            if os.path.exists(candidate):
                return candidate
        return candidates[0]

    def resolve_audio_path(self, filename):
        base_dir = self.resource_dir
        candidates = [
            os.path.join(base_dir, "audio", filename),
            os.path.join(base_dir, "appdir", "audio", filename),
        ]
        for candidate in candidates:
            if os.path.exists(candidate):
                return candidate
        return candidates[0]

    def play_sound(self, sound_path):
        if not self.sound_enabled.get() or not self.winmm or not os.path.exists(sound_path):
            return
        sound_path = sound_path.replace('"', '""')
        try:
            self.winmm.mciSendStringW('close notepadx_note', None, 0, None)
            self.winmm.mciSendStringW(f'open "{sound_path}" type mpegvideo alias notepadx_note', None, 0, None)
            self.winmm.mciSendStringW('play notepadx_note from 0', None, 0, None)
        except Exception:
            pass

    def play_unread_note_sound(self):
        self.play_sound(self.note_sound_path)

    def play_delete_note_sound(self):
        self.play_sound(self.delete_note_sound_path)

    def hide_support_file(self, file_path):
        if os.name != 'nt' or not file_path or not os.path.exists(file_path):
            return
        try:
            attributes = self.kernel32.GetFileAttributesW(file_path)
            if attributes == 0xFFFFFFFF:
                return
            hidden_attributes = attributes | 0x2
            self.kernel32.SetFileAttributesW(file_path, hidden_attributes)
        except Exception:
            pass

    def log_exception(self, where, exc=None):
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            trace = traceback.format_exc()
            if trace.strip() == 'NoneType: None':
                trace = ''.join(traceback.format_stack(limit=12))
            with open(self.crash_log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {where}\n")
                if exc is not None:
                    f.write(f"{type(exc).__name__}: {exc}\n")
                f.write(trace)
                f.write("\n" + ("-" * 80) + "\n")
        except Exception:
            pass

    def handle_unhandled_exception(self, exc_type, exc_value, exc_traceback):
        try:
            with open(self.crash_log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Unhandled exception\n")
                f.write(''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
                f.write("\n" + ("-" * 80) + "\n")
        except Exception:
            pass
        self.persist_recovery_state()
        if exc_type is KeyboardInterrupt:
            return
        try:
            messagebox.showerror("Notepad-X Crash", f"An unexpected error occurred.\nA crash log was written to:\n{self.crash_log_path}", parent=self.root)
        except Exception:
            pass

    def setup_exception_handling(self):
        sys.excepthook = self.handle_unhandled_exception

        def tk_exception_handler(exc_type, exc_value, exc_traceback):
            self.handle_unhandled_exception(exc_type, exc_value, exc_traceback)

        tk.Tk.report_callback_exception = staticmethod(tk_exception_handler)

    def load_known_editor_ids(self):
        try:
            if os.path.exists(self.editor_identity_path):
                with open(self.editor_identity_path, 'r', encoding='utf-8') as f:
                    identity = json.load(f)
                self.hide_support_file(self.editor_identity_path)
                if isinstance(identity, dict):
                    known_ids = identity.get('known_editor_ids', [])
                    if isinstance(known_ids, list):
                        return [str(editor_id).strip() for editor_id in known_ids if str(editor_id).strip()]
                    legacy_id = str(identity.get('editor_id', '')).strip()
                    if legacy_id:
                        return [legacy_id]
        except (OSError, json.JSONDecodeError, AttributeError):
            pass
        return []

    def generate_editor_id(self):
        seed = f"notepad-x-{os.getpid()}-{datetime.now().timestamp()}-{secrets.token_hex(16)}"
        return hashlib.md5(seed.encode('utf-8')).hexdigest()

    def persist_editor_identity(self):
        known_ids = list(dict.fromkeys(self.known_editor_ids + [self.editor_id]))[-32:]
        for attempt in range(2):
            try:
                os.makedirs(os.path.dirname(self.editor_identity_path), exist_ok=True)
                with open(self.editor_identity_path, 'w', encoding='utf-8') as f:
                    json.dump({
                        'editor_id': self.editor_id,
                        'known_editor_ids': known_ids
                    }, f, indent=2)
                self.hide_support_file(self.editor_identity_path)
                return
            except PermissionError as exc:
                if attempt == 0:
                    self.move_support_paths_to_user_dir()
                    continue
                self.log_exception("persist editor identity", exc)
                return
            except Exception as exc:
                self.log_exception("persist editor identity", exc)
                return

    def center_window(self, window, parent=None):
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()

        if parent is not None and parent.winfo_exists():
            parent.update_idletasks()
            x = parent.winfo_x() + max(0, (parent.winfo_width() - width) // 2)
            y = parent.winfo_y() + max(0, (parent.winfo_height() - height) // 2)
        else:
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = max(0, (screen_width - width) // 2)
            y = max(0, (screen_height - height) // 2)

        window.geometry(f"{width}x{height}+{x}+{y}")

    def center_window_after_show(self, window, parent=None):
        if not window.winfo_exists():
            return
        self.center_window(window, parent)
        window.lift()

    # ─── Font Zoom ───────────────────────────────────────────────
    def change_font_size(self, delta):
        old_size = self.current_font_size
        self.current_font_size += delta
        self.current_font_size = max(self.min_font_size, min(self.max_font_size, self.current_font_size))
        if self.current_font_size != old_size:
            self.update_font()
            self.update_status()

    def update_font(self):
        font_tuple = (self.font_family, self.current_font_size)
        for doc in self.documents.values():
            doc['text'].configure(font=font_tuple)
            self.update_line_number_gutter(doc)
        if self.compare_view:
            self.compare_view['text'].configure(font=font_tuple)
            self.update_line_number_gutter(self.compare_view)

    def toggle_numbered_lines(self):
        show_gutters = self.numbered_lines_enabled.get()
        for doc in self.documents.values():
            gutter = doc.get('line_numbers')
            if not gutter:
                continue
            if show_gutters:
                gutter.grid()
            else:
                gutter.grid_remove()
            self.update_line_number_gutter(doc)
        if self.compare_view:
            gutter = self.compare_view.get('line_numbers')
            if gutter:
                if show_gutters:
                    gutter.grid()
                else:
                    gutter.grid_remove()
                self.update_line_number_gutter(self.compare_view)
        self.save_session()
        return "break"

    def update_word_wrap(self):
        wrap_mode = tk.WORD if self.word_wrap_enabled.get() else tk.NONE
        for doc in self.documents.values():
            doc['text'].configure(wrap=wrap_mode)
        if self.compare_view:
            self.compare_view['text'].configure(wrap=wrap_mode)

    def on_ctrl_mousewheel(self, event):
        if event.state & 0x4:  # Ctrl held
            if event.delta > 0 or event.num == 4:
                self.change_font_size(+1)
            elif event.delta < 0 or event.num == 5:
                self.change_font_size(-1)
            return "break"

    def zoom_in(self, event=None):
        self.change_font_size(+1)
        return "break"

    def zoom_out(self, event=None):
        self.change_font_size(-1)
        return "break"

    def get_memory_usage_mb(self):
        try:
            counters = self.PROCESS_MEMORY_COUNTERS()
            counters.cb = ctypes.sizeof(counters)
            handle = self.kernel32.GetCurrentProcess()
            success = self.psapi.GetProcessMemoryInfo(
                handle,
                ctypes.byref(counters),
                counters.cb
            )
            if success:
                return max(1, round(counters.PagefileUsage / (1024 * 1024)))
        except Exception:
            pass
        return 0

    def update_memory_usage(self):
        self.memory_used_mb = self.get_memory_usage_mb()
        self.update_status()
        self.root.after(1000, self.update_memory_usage)

    # ─── Status Bar ──────────────────────────────────────────────
    def create_status_bar(self):
        self.status_frame = tk.Frame(self.root, bg='#2d2d2d')
        self.status_frame.grid(row=3, column=0, sticky='ew')
        self.status_frame.grid_columnconfigure(0, weight=1)

        self.status_left = tk.Frame(self.status_frame, bg='#2d2d2d')
        self.status_left.grid(row=0, column=0, sticky='ew')

        self.status = tk.Label(
            self.status_left,
            text="Ln 1 of 1, Col 1 | 0 characters | UTF-8 | Normal",
            anchor='w',
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=0, pady=4
        )
        self.status.pack(side='left')

        self.status_sync = tk.Label(
            self.status_left,
            text="",
            anchor='w',
            bg='#2d2d2d',
            fg='#4ecb71',
            font=('Segoe UI', 9, 'bold'),
            padx=0, pady=4
        )
        self.status_sync.pack(side='left')

        self.status_tail = tk.Label(
            self.status_left,
            text=" | Memory used: 0MB",
            anchor='w',
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=0, pady=4
        )
        self.status_tail.pack(side='left')

        self.status_clock = tk.Label(
            self.status_frame,
            text="",
            anchor='e',
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=8, pady=4
        )
        self.status_clock.grid(row=0, column=1, sticky='e')

    def get_zoom_text(self):
        if self.current_font_size == self.base_font_size:
            return "Normal"
        percent = round((self.current_font_size / self.base_font_size) * 100)
        return f"+{percent-100}%" if percent > 100 else f"{percent-100}%"

    def update_status(self):
        if not self.text or not hasattr(self, 'status'):
            return
        current_doc = self.get_current_doc()
        row, col = self.text.index(tk.INSERT).split('.')
        row = int(row)
        col = int(col) + 1
        if current_doc and current_doc.get('virtual_mode'):
            row = current_doc['window_start_line'] + row - 1
            total_lines = current_doc['total_file_lines']
            total_chars = current_doc['file_size_bytes']
            char_info = f"{total_chars:,} bytes"
        else:
            full_content = self.text.get('1.0', 'end-1c')
            total_lines = int(self.text.index('end-1c').split('.')[0])
            total_chars = len(full_content)
            try:
                sel_start = self.text.index('sel.first')
                sel_end = self.text.index('sel.last')
                selected_count = len(self.text.get(sel_start, sel_end))
            except tk.TclError:
                selected_count = 0
            char_info = f"{selected_count:,} of {total_chars:,} characters" if selected_count > 0 else f"{total_chars:,} characters"

        try:
            if current_doc and current_doc.get('virtual_mode'):
                selected_count = 0
            else:
                sel_start = self.text.index('sel.first')
                sel_end = self.text.index('sel.last')
                selected_count = len(self.text.get(sel_start, sel_end))
        except tk.TclError:
            selected_count = 0
        zoom_text = self.get_zoom_text()
        mode_suffix = ""
        if current_doc and current_doc.get('virtual_mode'):
            mode_suffix = " | Virtual"
        elif current_doc and current_doc.get('preview_mode'):
            mode_suffix = " | Preview"
        editor_label_text = ""
        if current_doc and current_doc.get('file_path'):
            editor_label_text = f" | ID: Notepad-X-{self.editor_id}"
        shared_notes_tail = ""
        if current_doc and current_doc.get('file_path'):
            unread_count = self.get_unread_note_count(current_doc)
            shared_notes_tail = (
                f" | {unread_count} unread (F3 to view) | "
                f"({current_doc.get('note_active_editors', 0)} editing)"
            )
        status_main_text = (
            f"Ln {row} of {total_lines}, Col {col} | "
            f"{char_info} | UTF-8 | {zoom_text}{mode_suffix}{editor_label_text}"
        )
        status_tail_text = f"{shared_notes_tail} | Memory used: {self.memory_used_mb}MB"
        self.status.config(text=status_main_text)
        self.status_sync.config(text="| Notes Synced" if current_doc and current_doc.get('file_path') else "")
        self.status_tail.config(text=status_tail_text)

    def hide_toast(self):
        popup = getattr(self, 'toast_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.toast_popup = None
        self.toast_after_id = None

    def show_toast(self, message, x=None, y=None):
        if self.toast_after_id:
            try:
                self.root.after_cancel(self.toast_after_id)
            except tk.TclError:
                pass
            self.toast_after_id = None
        self.hide_toast()

        popup = tk.Label(
            self.root,
            text=message,
            bg='#1f6feb',
            fg='white',
            font=('Segoe UI', 9, 'bold'),
            padx=12,
            pady=7,
            bd=0,
            highlightthickness=0
        )
        popup.update_idletasks()

        if x is None or y is None:
            x = self.root.winfo_pointerx()
            y = self.root.winfo_pointery()

        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        local_x = int(x) - root_x + 16
        local_y = int(y) - root_y + 14
        max_x = max(8, self.root.winfo_width() - popup.winfo_reqwidth() - 8)
        max_y = max(8, self.root.winfo_height() - popup.winfo_reqheight() - 8)
        local_x = max(8, min(local_x, max_x))
        local_y = max(8, min(local_y, max_y))
        popup.place(x=local_x, y=local_y)
        popup.lift()

        self.toast_popup = popup
        self.toast_after_id = self.root.after(1400, self.hide_toast)

    def create_line_number_gutter(self, parent, tab_id=None, doc=None):
        gutter = tk.Canvas(
            parent,
            width=56,
            bg='#0d1117',
            highlightthickness=0,
            borderwidth=0,
            relief='flat'
        )
        if tab_id is not None:
            gutter.bind('<Button-1>', lambda e, frame=tab_id: self.copy_line_from_gutter(e, frame))
        elif doc is not None:
            gutter.bind('<Button-1>', lambda e, target=doc: self.copy_line_from_gutter(e, target_doc=target))
        return gutter

    def get_display_line_number(self, doc, local_line):
        if doc.get('virtual_mode'):
            return doc.get('window_start_line', 1) + local_line - 1
        return local_line

    def update_line_number_gutter(self, doc):
        if not doc:
            return
        gutter = doc.get('line_numbers')
        text = doc.get('text')
        if not gutter or not text:
            return
        try:
            if not gutter.winfo_exists() or not text.winfo_exists():
                return
        except tk.TclError:
            return

        gutter.delete('all')
        gutter_height = max(gutter.winfo_height(), text.winfo_height(), 1)
        gutter_width = int(gutter.cget('width'))
        current_line = 1
        try:
            current_line = int(text.index(tk.INSERT).split('.')[0])
        except tk.TclError:
            pass

        gutter.create_rectangle(gutter_width - 1, 0, gutter_width, gutter_height, fill='#30363d', outline='')

        try:
            index = text.index('@0,0')
            while True:
                info = text.dlineinfo(index)
                if info is None:
                    break
                local_line = int(index.split('.')[0])
                display_line = self.get_display_line_number(doc, local_line)
                y = info[1]
                line_height = info[3]
                if local_line == current_line:
                    gutter.create_rectangle(0, y, gutter_width - 1, y + line_height, fill='#161b22', outline='')
                    line_fg = '#c9d1d9'
                else:
                    line_fg = '#8b949e'
                gutter.create_text(
                    gutter_width - 10,
                    y + (line_height / 2),
                    anchor='e',
                    text=str(display_line),
                    fill=line_fg,
                    font=(self.font_family, max(9, self.current_font_size - 1))
                )
                index = text.index(f"{local_line + 1}.0")
        except tk.TclError:
            return

    def copy_line_from_gutter(self, event, tab_id=None, target_doc=None):
        doc = target_doc
        if doc is None and tab_id is not None:
            doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"

        text = doc.get('text')
        if not text:
            return "break"

        try:
            local_index = text.index(f"@0,{event.y}")
            local_line = int(local_index.split('.')[0])
            line_text = text.get(f"{local_line}.0", f"{local_line}.end")
        except tk.TclError:
            return "break"

        copy_signature = (id(event.widget), local_line, line_text)
        now = time.monotonic()
        if self._last_gutter_copy:
            last_signature, last_time = self._last_gutter_copy
            if copy_signature == last_signature and (now - last_time) < 0.35:
                return "break"
        self._last_gutter_copy = (copy_signature, now)

        self.root.clipboard_clear()
        self.root.clipboard_append(line_text)
        self.root.update_idletasks()
        display_line = self.get_display_line_number(doc, local_line)
        toast_x = event.widget.winfo_rootx() + event.x
        toast_y = event.widget.winfo_rooty() + event.y
        self.show_toast(
            f"Copied line {display_line} to clipboard",
            x=toast_x,
            y=toast_y
        )
        return "break"
        self.status_clock.config(text=datetime.now().strftime("%A | %I:%M:%S %p | %m/%d/%Y").lstrip('0').replace('/0', '/'))

    # ─── Bottom Find / Replace Panels ────────────────────────────
    def create_bottom_panels(self):
        self.bottom_frame = tk.Frame(self.root, bg=self.panel_bg)
        self.bottom_frame.grid(row=2, column=0, sticky='ew')
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        # Find panel
        self.find_frame = tk.Frame(self.bottom_frame, bg=self.panel_bg)
        tk.Label(self.find_frame, text="Find:", bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(8,4), pady=6)
        self.find_entry = tk.Entry(self.find_frame, width=40)
        self.find_entry.pack(side='left', padx=4, pady=6)
        tk.Button(self.find_frame, text="Find Next", command=self.find_next)\
            .pack(side='left', padx=4, pady=6)
        tk.Checkbutton(
            self.find_frame,
            text="Search across all tabs",
            variable=self.search_all_tabs,
            bg=self.panel_bg,
            fg=self.fg_color,
            activebackground=self.panel_bg,
            activeforeground='white',
            selectcolor=self.panel_bg
        ).pack(side='left', padx=(10, 4), pady=6)
        self.find_entry.bind('<Return>', lambda e: self.find_next())
        self.find_entry.bind('<KeyRelease>', self.on_find_entry_change)
        self.find_entry.bind('<Escape>', lambda e: self.show_find_panel())   # ← added

        # Replace panel
        self.replace_frame = tk.Frame(self.bottom_frame, bg=self.panel_bg)
        tk.Label(self.replace_frame, text="Find:", bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(8,4), pady=6)
        self.replace_find_entry = tk.Entry(self.replace_frame, width=30)
        self.replace_find_entry.pack(side='left', padx=4, pady=6)
        tk.Label(self.replace_frame, text="Replace with:", bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(12,4), pady=6)
        self.replace_entry = tk.Entry(self.replace_frame, width=30)
        self.replace_entry.pack(side='left', padx=4, pady=6)
        tk.Button(self.replace_frame, text="Replace All", command=self.replace_all)\
            .pack(side='left', padx=8, pady=6)
        tk.Checkbutton(
            self.replace_frame,
            text="Search across all tabs",
            variable=self.search_all_tabs,
            bg=self.panel_bg,
            fg=self.fg_color,
            activebackground=self.panel_bg,
            activeforeground='white',
            selectcolor=self.panel_bg
        ).pack(side='left', padx=(10, 4), pady=6)

        self.replace_find_entry.bind('<Return>', lambda e: self.find_next())     # ← added
        self.replace_find_entry.bind('<KeyRelease>', self.on_find_entry_change)
        self.replace_find_entry.bind('<Escape>', lambda e: self.show_replace_panel())  # ← added
        self.replace_entry.bind('<Escape>', lambda e: self.show_replace_panel())       # ← added

        # Hide both initially
        self.find_frame.grid_remove()
        self.replace_frame.grid_remove()
        self.bottom_frame.grid_remove()

    def update_bottom_panel_visibility(self):
        if self.find_panel_visible or self.replace_panel_visible:
            self.bottom_frame.grid()
        else:
            self.bottom_frame.grid_remove()

    def show_find_panel(self):
        current_doc = self.get_current_doc()
        if current_doc and current_doc.get('virtual_mode'):
            messagebox.showinfo("Large File Mode", "Find is not available in buffered large-file mode yet.", parent=self.root)
            return "break"

        if self.replace_panel_visible:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.clear_find_highlights()

        if not self.find_panel_visible:
            self.bottom_frame.grid()
            self.find_frame.grid(sticky='ew')
            self.find_panel_visible = True
            self.find_entry.focus_set()
            self.on_find_entry_change()  # highlight if text already present
        else:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.find_entry.delete(0, tk.END)
            self.clear_find_highlights()
            self.focus_last_active_editor()

        self.update_bottom_panel_visibility()

        return "break"

    def show_replace_panel(self):
        current_doc = self.get_current_doc()
        if current_doc and current_doc.get('virtual_mode'):
            messagebox.showinfo("Large File Mode", "Replace is not available in buffered large-file mode.", parent=self.root)
            return "break"

        if self.find_panel_visible:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.clear_find_highlights()

        if not self.replace_panel_visible:
            self.bottom_frame.grid()
            self.replace_frame.grid(sticky='ew')
            self.replace_panel_visible = True
            self.replace_find_entry.focus_set()
        else:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.replace_find_entry.delete(0, tk.END)
            self.clear_find_highlights()
            self.focus_last_active_editor()

        self.update_bottom_panel_visibility()

        return "break"

    # ─── Find / Replace Logic ────────────────────────────────────
    def set_last_active_editor_widget(self, widget):
        if widget is None:
            return
        try:
            if widget.winfo_exists():
                self.last_active_editor_widget = widget
        except tk.TclError:
            pass

    def get_compare_text_widget(self):
        if not self.compare_active or not self.compare_view:
            return None
        compare_widget = self.compare_view.get('text')
        if not compare_widget:
            return None
        try:
            return compare_widget if compare_widget.winfo_exists() else None
        except tk.TclError:
            return None

    def get_active_search_widget(self):
        valid_widgets = [doc['text'] for doc in self.documents.values() if doc.get('text')]
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None:
            valid_widgets.append(compare_widget)

        focus_widget = self.root.focus_get()
        if focus_widget in valid_widgets:
            self.set_last_active_editor_widget(focus_widget)
            return focus_widget

        last_widget = getattr(self, 'last_active_editor_widget', None)
        if last_widget in valid_widgets:
            return last_widget

        if self.text in valid_widgets:
            return self.text
        return compare_widget

    def focus_last_active_editor(self):
        widget = self.get_active_search_widget()
        if widget is None:
            return
        try:
            widget.focus_set()
        except tk.TclError:
            pass

    def remember_compare_focus(self, event=None):
        self.hide_autocomplete_popup()
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None:
            self.root.after_idle(lambda widget=compare_widget: self.set_last_active_editor_widget(widget))
        self.update_status()
        return None

    def get_doc_for_text_widget(self, widget):
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and widget == compare_widget:
            return self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        for doc in self.documents.values():
            if doc.get('text') == widget:
                return doc
        return None

    def clear_current_find_marker(self):
        widgets = self.get_find_target_widgets()
        seen = set()
        for widget in widgets:
            widget_key = str(widget)
            if widget_key in seen:
                continue
            seen.add(widget_key)
            widget.tag_remove(self.find_current_tag, '1.0', tk.END)

    def set_current_find_match(self, widget, pos, end):
        if widget is None:
            return
        self.clear_current_find_marker()
        widget.tag_remove('sel', '1.0', tk.END)
        widget.tag_add('sel', pos, end)
        widget.tag_add(self.find_current_tag, pos, end)
        widget.mark_set(tk.INSERT, end)
        widget.see(pos)
        self.set_last_active_editor_widget(widget)

        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and widget == compare_widget:
            compare_doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
            if compare_doc:
                compare_doc['text'].tag_remove(self.find_current_tag, '1.0', tk.END)
                compare_doc['text'].tag_add(self.find_current_tag, pos, end)
        else:
            doc = self.get_doc_for_text_widget(widget)
            if doc:
                self.mirror_compare_current_match(doc, pos, end)

    def search_next_in_widget(self, widget, query, start_index=None, wrap=True):
        if widget is None or not query:
            return None

        start = start_index or widget.index(tk.INSERT)
        pos = widget.search(query, start, stopindex=tk.END, nocase=True, exact=False)
        if not pos and wrap:
            pos = widget.search(query, '1.0', stopindex=start, nocase=True, exact=False)
        if not pos:
            return None

        end = f"{pos}+{len(query)}c"
        self.set_current_find_match(widget, pos, end)
        self.update_status()
        return pos, end

    def find_next(self, event=None):
        if self.find_panel_visible:
            query = self.find_entry.get().strip()
        elif self.replace_panel_visible:
            query = self.replace_find_entry.get().strip()
        else:
            return self.goto_next_unread_note()

        if not query:
            return

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_next_in_widget(target_widget, query, wrap=True)
            return "break"

        if self.search_all_tabs.get():
            return self.find_next_across_tabs(query, target_widget)

        self.search_next_in_widget(target_widget, query, wrap=True)
        return "break"

    def find_next_across_tabs(self, query, start_widget=None):
        tab_ids = list(self.notebook.tabs())
        if not tab_ids:
            return "break"

        start_doc = self.get_doc_for_text_widget(start_widget) if start_widget is not None else self.get_current_doc()
        current_tab = str(start_doc['frame']) if start_doc else self.notebook.select()
        try:
            start_index = tab_ids.index(current_tab)
        except ValueError:
            start_index = 0
        search_order = tab_ids[start_index:] + tab_ids[:start_index]
        first_pass = True
        for tab_id in search_order:
            doc = self.documents.get(str(tab_id))
            if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
                continue
            self.notebook.select(doc['frame'])
            self.set_active_document(doc['frame'])
            search_widget = doc['text']
            start = search_widget.index(tk.INSERT) if first_pass and search_widget == start_widget else '1.0'
            pos = search_widget.search(query, start, stopindex=tk.END, nocase=True, exact=False)
            if pos:
                end = f"{pos}+{len(query)}c"
                self.set_current_find_match(search_widget, pos, end)
                return "break"
            first_pass = False
        return "break"

    def goto_next_unread_note(self):
        doc = self.get_current_doc()
        if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
            return "break"

        unread_tags = self.get_unread_note_tags(doc)
        if not unread_tags:
            return "break"

        note_tag = unread_tags[0]
        ranges = self.text.tag_ranges(note_tag)
        if len(ranges) < 2:
            return "break"

        start = str(ranges[0])
        end = str(ranges[1])
        self.text.tag_remove('sel', '1.0', tk.END)
        self.text.tag_add('sel', start, end)
        self.text.mark_set(tk.INSERT, end)
        self.text.see(start)
        self.mark_note_as_read(doc, note_tag)
        note_data = doc['notes'].get(note_tag)
        if note_data:
            bbox = self.text.bbox(start)
            if bbox:
                x, y, width, height = bbox
                self.show_note_popup(
                    doc,
                    note_data,
                    self.text.winfo_rootx() + x + width,
                    self.text.winfo_rooty() + y + height
                )
        self.update_status()
        return "break"

    def get_ordered_note_tags(self, doc):
        ordered = []
        for note_tag, note_data in doc['notes'].items():
            if not self.note_matches_filter(doc, note_tag, note_data):
                continue
            ranges = doc['text'].tag_ranges(note_tag)
            if len(ranges) < 2:
                continue
            start_index = str(ranges[0])
            note_id = str(note_data.get('id', ''))
            ordered.append((doc['text'].count('1.0', start_index, 'chars')[0], int(note_id) if note_id.isdigit() else 0, note_tag))
        ordered.sort(key=lambda item: (item[0], item[1]))
        return [tag for _, _, tag in ordered]

    def note_matches_filter(self, doc, note_tag, note_data):
        filter_mode = self.note_filter.get()
        if filter_mode == 'all':
            return True
        if filter_mode == 'unread':
            return note_tag in self.get_unread_note_tags(doc)
        if filter_mode == 'allowed':
            return bool(note_data.get('approved_by'))
        if filter_mode == 'denied':
            return bool(note_data.get('dissapproved_by'))
        return True

    def goto_next_note(self, event=None):
        doc = self.get_current_doc()
        if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
            return "break"

        ordered_tags = self.get_ordered_note_tags(doc)
        if not ordered_tags:
            return "break"

        current_offset = self.text.count('1.0', self.text.index(tk.INSERT), 'chars')[0]
        note_tag = ordered_tags[0]
        for candidate_tag in ordered_tags:
            ranges = self.text.tag_ranges(candidate_tag)
            if len(ranges) < 2:
                continue
            note_offset = self.text.count('1.0', str(ranges[0]), 'chars')[0]
            if note_offset > current_offset:
                note_tag = candidate_tag
                break

        ranges = self.text.tag_ranges(note_tag)
        if len(ranges) < 2:
            return "break"

        start = str(ranges[0])
        end = str(ranges[1])
        self.text.tag_remove('sel', '1.0', tk.END)
        self.text.tag_add('sel', start, end)
        self.text.mark_set(tk.INSERT, end)
        self.text.see(start)
        self.mark_note_as_read(doc, note_tag)
        note_data = doc['notes'].get(note_tag)
        if note_data:
            bbox = self.text.bbox(start)
            if bbox:
                x, y, width, height = bbox
                self.show_note_popup(
                    doc,
                    note_data,
                    self.text.winfo_rootx() + x + width,
                    self.text.winfo_rooty() + y + height
                )
        self.update_status()
        return "break"

    def get_find_target_widgets(self):
        widgets = []
        if self.search_all_tabs.get():
            for doc in self.documents.values():
                if doc.get('virtual_mode') or doc.get('preview_mode'):
                    continue
                widgets.append(doc['text'])
        elif self.text:
            widgets.append(self.text)

        if self.compare_active and self.compare_view and self.compare_view.get('text'):
            widgets.append(self.compare_view['text'])
        return widgets

    def clear_find_highlights(self):
        widgets = [doc['text'] for doc in self.documents.values()]
        if self.compare_view and self.compare_view.get('text'):
            widgets.append(self.compare_view['text'])
        for widget in widgets:
            widget.tag_remove(self.find_matches_tag, '1.0', tk.END)
            widget.tag_remove(self.find_current_tag, '1.0', tk.END)

    def highlight_matches_in_widget(self, text_widget, query):
        text_widget.tag_remove(self.find_matches_tag, '1.0', tk.END)
        text_widget.tag_remove(self.find_current_tag, '1.0', tk.END)
        if not query:
            return
        start = "1.0"
        while True:
            pos = text_widget.search(query, start, stopindex=tk.END, nocase=True, exact=False)
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            text_widget.tag_add(self.find_matches_tag, pos, end)
            start = end

    def mirror_compare_current_match(self, doc, pos, end):
        if not self.compare_active or not self.compare_view:
            return
        if self.compare_source_tab != str(doc['frame']):
            self.compare_view['text'].tag_remove(self.find_current_tag, '1.0', tk.END)
            return
        compare_text = self.compare_view['text']
        compare_text.tag_remove(self.find_current_tag, '1.0', tk.END)
        compare_text.tag_add(self.find_current_tag, pos, end)
        compare_text.see(pos)

    def highlight_all_matches(self, query):
        widgets = self.get_find_target_widgets()
        for widget in widgets:
            self.highlight_matches_in_widget(widget, query)

    def on_find_entry_change(self, event=None):
        if self.find_panel_visible:
            query = self.find_entry.get().strip()
        elif self.replace_panel_visible:
            query = self.replace_find_entry.get().strip()
        else:
            return

        self.highlight_all_matches(query)
        if not query:
            return
        if self.search_all_tabs.get():
            self.find_next_across_tabs(query)
            return

        pos = self.text.search(query, '1.0', stopindex=tk.END, nocase=True, exact=False)
        if not pos:
            return

        end = f"{pos}+{len(query)}c"
        self.text.tag_remove('sel', '1.0', tk.END)
        self.text.tag_add('sel', pos, end)
        self.text.tag_add(self.find_current_tag, pos, end)
        self.text.mark_set(tk.INSERT, end)
        self.text.see(pos)
        self.update_status()

    def replace_all(self):
        query = self.replace_find_entry.get().strip()
        replace_text = self.replace_entry.get()
        if not query:
            return
        content = self.text.get('1.0', tk.END)
        new_content = content.replace(query, replace_text)
        self.text.delete('1.0', tk.END)
        self.text.insert('1.0', new_content.rstrip('\n'))
        self.text.edit_modified(True)
        messagebox.showinfo("Replace All", f"Replaced {content.count(query)} occurrence(s).", parent=self.root)
        self.update_status()

    # ─── Key Bindings ────────────────────────────────────────────
    def bind_keys(self):
        # File
        self.root.bind('<Control-t>', self.new_tab)
        self.root.bind('<Control-T>', self.new_tab)
        self.root.bind('<Control-w>', self.open_file)
        self.root.bind('<Control-W>', self.open_file)
        self.root.bind('<Control-p>', self.print_file)
        self.root.bind('<Control-P>', self.print_file)
        self.root.bind('<Control-Shift-W>', self.open_project)
        self.root.bind('<Control-Shift-w>', self.open_project)
        self.root.bind('<Control-s>', self.save)
        self.root.bind('<Control-S>', self.save)
        self.root.bind('<Control-Shift-s>', self.save_all)
        self.root.bind('<Control-Shift-S>', self.save_all)
        self.root.bind('<Control-Shift-q>', lambda e: self.save_copy_as())
        self.root.bind('<Control-Shift-Q>', lambda e: self.save_copy_as())
        self.root.bind('<Control-Shift-X>', self.ctrl_shift_x)
        self.root.bind('<Control-Shift-x>', self.ctrl_shift_x)
        self.root.bind_all('<Control-Shift-X>', self.ctrl_shift_x)
        self.root.bind_all('<Control-Shift-x>', self.ctrl_shift_x)

        # Edit
        self.root.bind('<Control-z>', self.undo)
        self.root.bind('<Control-Z>', self.undo)
        self.root.bind('<Control-Shift-Z>', self.redo)
        self.root.bind('<Control-Shift-z>', self.redo)
        self.root.bind('<Control-x>', lambda e: self.text.event_generate("<<Cut>>"))
        self.root.bind('<Control-X>', lambda e: self.text.event_generate("<<Cut>>"))
        self.root.bind('<Control-a>', self.select_all)
        self.root.bind('<Control-A>', self.select_all)
        self.root.bind_all('<Control-b>', self.toggle_status_bar)
        self.root.bind_all('<Control-B>', self.toggle_status_bar)
        self.root.bind('<Control-d>', self.insert_date)
        self.root.bind('<Control-D>', self.insert_date)
        self.root.bind('<Control-Shift-D>', self.insert_time_date)
        self.root.bind('<Control-Shift-d>', self.insert_time_date)
        self.root.bind('<Control-e>', lambda e: self.export_notes_report())
        self.root.bind('<Control-E>', lambda e: self.export_notes_report())
        self.root.bind('<Control-Shift-T>', self.close_current_tab)
        self.root.bind('<Control-Shift-t>', self.close_current_tab)
        self.root.bind('<Control-Tab>', self.switch_tab_right)
        self.root.bind('<Control-Left>', self.switch_tab_left)
        self.root.bind('<Control-Right>', self.switch_tab_right)
        self.root.bind('<Control-Shift-F>', self.show_font_dialog)
        self.root.bind('<Control-Shift-f>', self.show_font_dialog)
        # Search / Navigation
        self.root.bind('<Control-f>', lambda e: self.show_find_panel())
        self.root.bind('<Control-F>', lambda e: self.show_find_panel())
        self.root.bind('<Control-r>', lambda e: self.show_replace_panel())
        self.root.bind('<Control-R>', lambda e: self.show_replace_panel())
        self.root.bind('<F3>', self.find_next)
        self.root.bind('<F4>', self.goto_next_note)
        self.root.bind('<Control-g>', self.goto_line_dialog)
        self.root.bind('<Control-G>', self.goto_line_dialog)

        # Zoom
        self.root.bind('<Control-MouseWheel>', self.on_ctrl_mousewheel)
        self.root.bind('<Control-Button-4>', self.on_ctrl_mousewheel)
        self.root.bind('<Control-Button-5>', self.on_ctrl_mousewheel)
        self.root.bind('<Control-plus>', self.zoom_in)
        self.root.bind('<Control-equal>', self.zoom_in)
        self.root.bind('<Control-KP_Add>', self.zoom_in)
        self.root.bind('<Control-minus>', self.zoom_out)
        self.root.bind('<Control-KP_Subtract>', self.zoom_out)
        self.root.bind('<Control-q>', lambda e: self.show_split_compare())
        self.root.bind('<Control-Q>', lambda e: self.show_split_compare())
        self.root.bind('<F11>', self.toggle_fullscreen)

    def cut_or_close_panel(self, event=None):
        if self.find_panel_visible or self.replace_panel_visible:
            if self.find_panel_visible:
                self.show_find_panel()
            if self.replace_panel_visible:
                self.show_replace_panel()
            self.focus_last_active_editor()
            return "break"
        else:
            if self.current_doc_is_large_readonly():
                return "break"
            self.text.event_generate("<<Cut>>")
            return "break"

    def ctrl_shift_x(self, event=None):
        if self.compare_active:
            return self.close_compare_panel()
        if self.find_panel_visible or self.replace_panel_visible:
            if self.find_panel_visible:
                self.show_find_panel()
            if self.replace_panel_visible:
                self.show_replace_panel()
            self.focus_last_active_editor()
            return "break"
        return self.exit_app(event)

    def toggle_fullscreen(self, event=None):
        self.fullscreen = not self.fullscreen
        self.root.attributes('-fullscreen', self.fullscreen)

        if self.fullscreen:
            self.fullscreen_panel_restore = self.find_panel_visible or self.replace_panel_visible
            if self.find_panel_visible:
                self.find_frame.grid_remove()
                self.find_panel_visible = False
            if self.replace_panel_visible:
                self.replace_frame.grid_remove()
                self.replace_panel_visible = False
            self.update_bottom_panel_visibility()
            self.root.config(menu='')
        else:
            self.root.config(menu=self.menu)
            if self.fullscreen_panel_restore:
                self.update_bottom_panel_visibility()
            self.fullscreen_panel_restore = False

        self.update_status()
        return "break"

    def toggle_word_wrap(self):
        self.update_word_wrap()
        self.update_status()
        self.save_session()

    def toggle_autocomplete(self):
        if not self.autocomplete_enabled.get():
            self.hide_autocomplete_popup()
        self.save_session()

    def toggle_sound(self):
        self.save_session()

    def toggle_status_bar(self, event=None):
        if event is not None:
            self.status_bar_enabled.set(not self.status_bar_enabled.get())
        if self.status_bar_enabled.get():
            self.status_frame.grid()
        else:
            self.status_frame.grid_remove()
        self.save_session()
        return "break"

    # ─── Text Widget ─────────────────────────────────────────────
    def create_text_widget(self):
        frame = tk.Frame(self.root, bg=self.bg_color)
        frame.grid(row=0, column=0, sticky='nsew')
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use('default')
        style.configure('EditorTabs.TNotebook', background=self.bg_color, borderwidth=0)
        style.configure(
            'EditorTabs.TNotebook.Tab',
            background='#2d2d2d',
            foreground=self.fg_color,
            padding=(16, 8),
            borderwidth=0
        )
        style.map(
            'EditorTabs.TNotebook.Tab',
            background=[('selected', '#0d1117')],
            foreground=[('selected', 'white')]
        )

        self.editor_frame = frame
        self.editor_paned = tk.PanedWindow(
            frame,
            orient='horizontal',
            sashrelief='flat',
            sashwidth=4,
            bd=0,
            relief='flat',
            bg='#2d2d2d',
            showhandle=False
        )
        self.editor_paned.grid(row=0, column=0, sticky='nsew')

        self.primary_editor_container = tk.Frame(self.editor_paned, bg=self.bg_color)
        self.primary_editor_container.grid_rowconfigure(0, weight=1)
        self.primary_editor_container.grid_columnconfigure(0, weight=1)
        self.editor_paned.add(self.primary_editor_container, stretch='always')

        self.notebook = ttk.Notebook(self.primary_editor_container, style='EditorTabs.TNotebook')
        self.notebook.grid(row=0, column=0, sticky='nsew')
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
        self.notebook.bind('<ButtonPress-1>', self.on_tab_drag_start, add='+')
        self.notebook.bind('<B1-Motion>', self.on_tab_drag_motion, add='+')
        self.notebook.bind('<ButtonRelease-1>', self.on_tab_drag_end, add='+')

        self.compare_container = tk.Frame(self.editor_paned, bg=self.bg_color)
        self.compare_container.grid_rowconfigure(1, weight=1)
        self.compare_container.grid_columnconfigure(0, weight=1)

        compare_header = tk.Frame(self.compare_container, bg='#161b22')
        compare_header.grid(row=0, column=0, sticky='ew')
        compare_header.grid_columnconfigure(0, weight=1)

        self.compare_title = tk.Label(
            compare_header,
            text="",
            bg='#161b22',
            fg=self.fg_color,
            font=('Segoe UI', 10, 'bold'),
            anchor='w',
            padx=10,
            pady=8
        )
        self.compare_title.grid(row=0, column=0, sticky='ew')

        compare_body = tk.Frame(self.compare_container, bg=self.bg_color)
        compare_body.grid(row=1, column=0, sticky='nsew')
        compare_body.grid_rowconfigure(0, weight=1)
        compare_body.grid_columnconfigure(1, weight=1)

        self.compare_line_numbers = self.create_line_number_gutter(compare_body)
        self.compare_line_numbers.grid(row=0, column=0, sticky='ns')
        if not self.numbered_lines_enabled.get():
            self.compare_line_numbers.grid_remove()

        self.compare_text = tk.Text(
            compare_body,
            undo=False,
            bg=self.text_bg,
            fg=self.text_fg,
            insertbackground=self.cursor_color,
            selectbackground=self.select_bg,
            selectforeground='white',
            wrap=tk.WORD if self.word_wrap_enabled.get() else tk.NONE,
            padx=8,
            pady=6,
            borderwidth=0,
            highlightthickness=0,
            relief='flat',
            font=(self.font_family, self.base_font_size),
            spacing1=2,
            spacing2=1,
            spacing3=2
        )
        self.compare_text.grid(row=0, column=1, sticky='nsew')

        compare_v_scroll = ttk.Scrollbar(compare_body, orient='vertical', command=self.on_compare_vertical_scroll)
        compare_v_scroll.grid(row=0, column=2, sticky='ns')
        compare_h_scroll = ttk.Scrollbar(compare_body, orient='horizontal', command=self.compare_text.xview)
        compare_h_scroll.grid(row=1, column=1, sticky='ew')
        self.compare_text.config(
            yscrollcommand=lambda first, last, scroll=compare_v_scroll: self.update_compare_vertical_scrollbar(scroll, first, last),
            xscrollcommand=compare_h_scroll.set
        )
        self.compare_text.bind('<KeyPress>', self.handle_compare_keypress)
        self.compare_text.bind('<<Paste>>', lambda e: "break")
        self.compare_text.bind('<<Cut>>', lambda e: "break")
        self.compare_text.bind('<Control-b>', self.toggle_status_bar)
        self.compare_text.bind('<Control-B>', self.toggle_status_bar)
        self.compare_text.bind('<FocusIn>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<Button-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<ButtonRelease-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<F3>', self.find_next)
        self.compare_text.bind('<Control-Shift-X>', self.ctrl_shift_x)
        self.compare_text.bind('<Control-Shift-x>', self.ctrl_shift_x)

        self.compare_view = {
            'frame': self.compare_container,
            'text': self.compare_text,
            'line_numbers': self.compare_line_numbers,
            'percolator': None,
            'colorizer': None,
            'large_file_mode': False,
            'preview_mode': False,
            'virtual_mode': False,
            'syntax_job': None,
            'syntax_mode': None,
            'syntax_override': None,
            'file_path': None,
            'window_start_line': 1,
            'window_end_line': 1,
            'total_file_lines': 1,
        }
        self.compare_line_numbers.bind('<Button-1>', lambda e: self.copy_line_from_gutter(e, target_doc=self.compare_view))
        self.compare_text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        self.compare_text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')
        self.apply_syntax_tag_colors(self.compare_text)

        self.text = None
        self.create_tab()

    def create_tab(self, file_path=None, content="", select=True):
        tab_frame = tk.Frame(self.notebook, bg=self.bg_color)
        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(1, weight=1)

        line_numbers = self.create_line_number_gutter(tab_frame, tab_id=tab_frame)
        line_numbers.grid(row=0, column=0, sticky='ns')
        if not self.numbered_lines_enabled.get():
            line_numbers.grid_remove()

        text = tk.Text(
            tab_frame, undo=True,
            bg=self.text_bg, fg=self.text_fg,
            insertbackground=self.cursor_color,
            selectbackground=self.select_bg,
            selectforeground='white',
            wrap=tk.WORD if self.word_wrap_enabled.get() else tk.NONE, padx=8, pady=6,
            borderwidth=0, highlightthickness=0, relief='flat',
            font=(self.font_family, self.base_font_size),
            spacing1=2, spacing2=1, spacing3=2
        )
        text.grid(row=0, column=1, sticky='nsew')

        v_scroll = ttk.Scrollbar(tab_frame, orient='vertical')
        v_scroll.grid(row=0, column=2, sticky='ns')
        h_scroll = ttk.Scrollbar(tab_frame, orient='horizontal', command=text.xview)
        h_scroll.grid(row=1, column=1, sticky='ew')
        v_scroll.configure(command=lambda *args, frame=tab_frame: self.on_vertical_scroll(frame, *args))
        text.config(
            yscrollcommand=lambda first, last, frame=tab_frame, scroll=v_scroll: self.update_vertical_scrollbar(frame, scroll, first, last),
            xscrollcommand=h_scroll.set
        )

        for evt in ('<KeyRelease>', '<MouseWheel>', '<Button-4>', '<Button-5>',
                    '<Button-1>', '<B1-Motion>', '<ButtonRelease-1>', '<Double-Button-1>',
                    '<<Selection>>'):
            text.bind(evt, lambda e, frame=tab_frame: self.handle_text_activity(e, frame))

        text.bind('<MouseWheel>', lambda e, frame=tab_frame: self.on_text_mousewheel(e, frame))
        text.bind('<Button-4>', lambda e, frame=tab_frame: self.on_text_mousewheel(e, frame))
        text.bind('<Button-5>', lambda e, frame=tab_frame: self.on_text_mousewheel(e, frame))
        text.bind('<Control-b>', self.toggle_status_bar)
        text.bind('<Control-B>', self.toggle_status_bar)
        text.bind('<FocusIn>', lambda e, frame=tab_frame: self.remember_doc_focus(frame), add='+')
        text.bind('<KeyPress>', lambda e, frame=tab_frame: self.handle_text_keypress(e, frame))
        text.bind('<KeyRelease>', lambda e, frame=tab_frame: self.handle_text_keyrelease(e, frame), add='+')
        text.bind('<ButtonRelease-1>', lambda e, frame=tab_frame: self.on_text_click_release(e, frame), add='+')
        text.bind('<Button-3>', lambda e, frame=tab_frame: self.show_text_context_menu(e, frame))

        text.bind('<<Modified>>', lambda e, frame=tab_frame: self.on_text_modified(frame))
        text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')

        if content:
            text.insert('1.0', content)
        text.edit_modified(False)

        self.documents[str(tab_frame)] = {
            'frame': tab_frame,
            'text': text,
            'line_numbers': line_numbers,
            'v_scroll': v_scroll,
            'file_path': file_path,
            'untitled_name': self.next_untitled_name(),
            'percolator': None,
            'colorizer': None,
            'large_file_mode': False,
            'preview_mode': False,
            'file_size_bytes': 0,
            'virtual_mode': False,
            'line_starts': None,
            'total_file_lines': 1,
            'window_start_line': 1,
            'window_end_line': 1,
            'suspend_modified_events': False,
            'notes': {},
            'note_counter': 1,
            'note_popup': None,
            'context_menu': None,
            'context_note_tag': None,
            'note_sync_mtime': None,
            'note_sync_signature': None,
            'note_active_editors': 0,
            'note_editors': [],
            'note_editor_label': None,
            'last_unread_count': 0,
            'notes_registered': False,
            'syntax_job': None,
            'syntax_mode': None,
            'syntax_override': None,
        }
        self.apply_syntax_tag_colors(text)
        self.configure_syntax_highlighting(tab_frame)
        self.create_text_context_menu(tab_frame)

        self.notebook.add(tab_frame, text=self.get_doc_title(tab_frame))
        if select:
            self.notebook.select(tab_frame)
            self.set_active_document(tab_frame)
        self.root.after_idle(lambda frame=tab_frame: self.update_line_number_gutter(self.documents.get(str(frame))))
        return tab_frame

    def get_syntax_mode(self, doc):
        if doc['large_file_mode'] or doc['virtual_mode'] or doc['preview_mode']:
            return None

        syntax_override = doc.get('syntax_override')
        if syntax_override == 'plain':
            return None
        if syntax_override and syntax_override != 'auto':
            return syntax_override

        if doc['file_path']:
            file_name = os.path.basename(doc['file_path']).lower()
            ext = os.path.splitext(file_name)[1].lower()
            syntax_extensions = {
                '.as': 'actionscript',
                '.mx': 'actionscript',
                '.asp': 'asp',
                '.aspx': 'aspx',
                '.au3': 'autoit',
                '.sh': 'bash',
                '.bsh': 'bash',
                '.bat': 'batch',
                '.cmd': 'batch',
                '.ml': 'caml',
                '.mli': 'caml',
                '.sml': 'caml',
                '.thy': 'caml',
                '.c': 'c',
                '.cpp': 'cpp',
                '.cxx': 'cpp',
                '.cc': 'cpp',
                '.h': 'cpp',
                '.hxx': 'cpp',
                '.hpp': 'cpp',
                '.hh': 'cpp',
                '.rgs': 'cpp',
                '.cs': 'csharp',
                '.css': 'css',
                '.diff': 'diff',
                '.patch': 'diff',
                '.f': 'fortran',
                '.for': 'fortran',
                '.f90': 'fortran',
                '.f95': 'fortran',
                '.f2k': 'fortran',
                '.ini': 'ini',
                '.inf': 'ini',
                '.reg': 'ini',
                '.url': 'ini',
                '.iss': 'inno',
                '.java': 'java',
                '.js': 'javascript',
                '.lsp': 'lisp',
                '.lisp': 'lisp',
                '.mak': 'makefile',
                '.m': 'matlab',
                '.nfo': 'nfo',
                '.nsi': 'nsis',
                '.nsh': 'nsis',
                '.pas': 'pascal',
                '.inc': 'pascal',
                '.pl': 'perl',
                '.pm': 'perl',
                '.plx': 'perl',
                '.php': 'php',
                '.php3': 'php',
                '.phtml': 'php',
                '.py': 'python',
                '.pyw': 'python',
                '.rc': 'resource',
                '.rs': 'rust',
                '.st': 'smalltalk',
                '.tex': 'tex',
                '.sql': 'sql',
                '.vb': 'vb',
                '.vbs': 'vbscript',
                '.xml': 'xml',
                '.xsd': 'xml',
                '.xsml': 'xml',
                '.xsl': 'xml',
                '.kml': 'xml',
                '.asm': 'assembly',
                '.s': 'assembly',
                '.html': 'html',
                '.htm': 'html',
            }

            if file_name in {'makefile', 'gnumakefile'}:
                return 'makefile'

            mode = syntax_extensions.get(ext)
            if mode == 'python' and self.syntax_highlighting_available:
                return 'python'
            if mode:
                return mode

        return self.detect_syntax_mode_from_content(doc['text'].get('1.0', 'end-1c'))

    def detect_syntax_mode_from_content(self, content):
        sample = content[:12000]
        lowered = sample.lower()

        if self.syntax_highlighting_available and ('def ' in sample or 'import ' in sample or 'class ' in sample):
            if ':' in sample and 'self' in sample:
                return 'python'

        if '<?php' in lowered:
            return 'php'
        if '<html' in lowered or '<div' in lowered or '<body' in lowered or '<head' in lowered:
            return 'html'
        if 'fn main' in sample or 'impl ' in sample or 'let mut ' in sample or '::' in sample or 'use ' in sample:
            return 'rust'
        if '#include' in sample or 'printf(' in sample or 'scanf(' in sample:
            return 'c'
        if 'std::' in sample or 'cout <<' in sample or 'cin >>' in sample or 'class ' in sample:
            return 'cpp'
        if 'using system;' in lowered or 'namespace ' in sample and 'class ' in sample:
            return 'csharp'
        if 'public static void main' in sample or 'system.out.' in lowered:
            return 'java'
        if 'function ' in sample and 'var ' in sample:
            return 'javascript'
        if re.search(r'^\s*select\b|^\s*update\b|^\s*insert\b|^\s*delete\b|^\s*create\b', lowered, re.MULTILINE):
            return 'sql'
        if '<?xml' in lowered:
            return 'xml'
        if re.search(r'^\s*\$[A-Za-z_]\w*\s*=', sample, re.MULTILINE):
            return 'bash'
        if re.search(r'^\s*@echo\b|^\s*set\s+\w+=', lowered, re.MULTILINE):
            return 'batch'
        if re.search(r'^\s*(mov|jmp|cmp|push|pop|call|ret)\b', lowered, re.MULTILINE):
            return 'assembly'
        return None

    def set_large_file_mode(self, doc, enabled):
        doc['large_file_mode'] = enabled
        doc['text'].configure(undo=not enabled, autoseparators=not enabled)
        if enabled:
            try:
                doc['text'].edit_reset()
            except tk.TclError:
                pass

    def configure_syntax_highlighting(self, tab_id):
        doc = self.documents.get(str(tab_id))
        self.configure_syntax_for_doc(doc)

    def configure_syntax_for_doc(self, doc):
        if not doc:
            return

        text = doc['text']
        palette = self.get_syntax_palette()
        self.apply_syntax_tag_colors(text)
        mode = self.get_syntax_mode(doc)
        doc['syntax_mode'] = mode

        if mode == 'python' and doc['colorizer'] is None:
            colorizer = ColorDelegator()
            colorizer.tagdefs.update({
                'COMMENT': {'foreground': palette['comment']},
                'KEYWORD': {'foreground': palette['keyword']},
                'BUILTIN': {'foreground': palette['preprocessor']},
                'STRING': {'foreground': palette['string']},
                'DEFINITION': {'foreground': palette['type']},
                'SYNC': {'background': None, 'foreground': None},
                'TODO': {'background': None, 'foreground': None},
                'ERROR': {'foreground': '#ff6b6b'},
                'hit': {'background': self.match_bg, 'foreground': 'black'},
            })
            percolator = Percolator(text)
            percolator.insertfilter(colorizer)
            doc['percolator'] = percolator
            doc['colorizer'] = colorizer
            colorizer.notify_range('1.0', tk.END)
        elif mode != 'python' and doc['colorizer'] is not None:
            try:
                doc['percolator'].removefilter(doc['colorizer'])
            except Exception:
                pass
            doc['percolator'] = None
            doc['colorizer'] = None

        if mode and mode != 'python':
            self.schedule_syntax_highlight(doc)
        else:
            self.clear_custom_syntax_tags(doc)

    def get_syntax_palette(self):
        palettes = {
            'Default': {
                'keyword': '#ff7b72',
                'type': '#79c0ff',
                'string': '#a5d6ff',
                'comment': '#6a9955',
                'number': '#f2cc60',
                'preprocessor': '#d2a8ff',
                'tag': '#7ee787',
            },
            'Soft': {
                'keyword': '#f38ba8',
                'type': '#89b4fa',
                'string': '#94e2d5',
                'comment': '#a6adc8',
                'number': '#f9e2af',
                'preprocessor': '#cba6f7',
                'tag': '#a6e3a1',
            },
            'Vivid': {
                'keyword': '#ff4d6d',
                'type': '#4cc9f0',
                'string': '#72efdd',
                'comment': '#80ed99',
                'number': '#ffd166',
                'preprocessor': '#b388ff',
                'tag': '#06d6a0',
            },
        }
        return palettes.get(self.syntax_theme.get(), palettes['Default'])

    def apply_syntax_tag_colors(self, text_widget):
        palette = self.get_syntax_palette()
        text_widget.tag_config('syntax_keyword', foreground=palette['keyword'])
        text_widget.tag_config('syntax_type', foreground=palette['type'])
        text_widget.tag_config('syntax_string', foreground=palette['string'])
        text_widget.tag_config('syntax_comment', foreground=palette['comment'])
        text_widget.tag_config('syntax_number', foreground=palette['number'])
        text_widget.tag_config('syntax_preprocessor', foreground=palette['preprocessor'])
        text_widget.tag_config('syntax_tag', foreground=palette['tag'])

    def set_syntax_theme(self, theme_name):
        self.syntax_theme.set(theme_name)
        palette = self.get_syntax_palette()
        for doc in self.documents.values():
            self.apply_syntax_tag_colors(doc['text'])
            if doc.get('colorizer') is not None:
                doc['colorizer'].tagdefs.update({
                    'COMMENT': {'foreground': palette['comment']},
                    'KEYWORD': {'foreground': palette['keyword']},
                    'BUILTIN': {'foreground': palette['preprocessor']},
                    'STRING': {'foreground': palette['string']},
                    'DEFINITION': {'foreground': palette['type']},
                })
                try:
                    doc['colorizer'].notify_range('1.0', tk.END)
                except Exception:
                    pass
            self.configure_syntax_highlighting(doc['frame'])
        if self.compare_active:
            self.refresh_compare_panel()
        self.save_session()

    def set_current_syntax_override(self, syntax_mode):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        doc['syntax_override'] = syntax_mode
        self.syntax_mode_selection.set(syntax_mode)
        self.configure_syntax_highlighting(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        self.update_status()
        self.save_session()
        return "break"

    def clear_custom_syntax_tags(self, doc):
        for tag_name in ('syntax_keyword', 'syntax_type', 'syntax_string', 'syntax_comment', 'syntax_number', 'syntax_preprocessor', 'syntax_tag'):
            doc['text'].tag_remove(tag_name, '1.0', tk.END)

    def schedule_syntax_highlight(self, doc):
        if doc.get('syntax_job'):
            try:
                self.root.after_cancel(doc['syntax_job'])
            except tk.TclError:
                pass
        doc['syntax_job'] = self.root.after(120, lambda current=doc: self.apply_custom_syntax_highlighting(current))

    def apply_pattern_tag(self, doc, pattern, tag_name, flags=0):
        content = doc['text'].get('1.0', 'end-1c')
        for match in re.finditer(pattern, content, flags):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            doc['text'].tag_add(tag_name, start, end)

    def apply_custom_syntax_highlighting(self, doc):
        doc['syntax_job'] = None
        mode = doc.get('syntax_mode')
        if not mode or mode == 'python' or doc.get('virtual_mode') or doc.get('preview_mode'):
            return

        self.clear_custom_syntax_tags(doc)

        language_sets = {
            'actionscript': {
                'keywords': r'\b(as|break|case|catch|class|const|continue|default|delete|do|else|extends|false|finally|for|function|if|implements|import|in|instanceof|interface|is|new|null|package|private|protected|public|return|super|switch|this|throw|true|try|typeof|use|var|void|while|with)\b',
                'types': r'\b(Array|Boolean|Date|Function|int|Number|Object|RegExp|String|uint|void|XML|XMLList)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': None,
            },
            'asp': {
                'keywords': r'\b(If|Then|Else|End|For|Each|Next|Do|Loop|While|Function|Sub|Dim|Set|Select|Case|Class|Const)\b|<%|%>',
                'types': r'\b(String|Integer|Boolean|Object|Variant|Date)\b',
                'comment': r'\'.*?$|<!--.*?-->',
                'preprocessor': None,
            },
            'aspx': {
                'keywords': r'\b(Page|Control|Import|Master|Register|if|else|for|foreach|public|private|protected|class|return)\b|<%@.*?%>',
                'types': r'\b(string|int|bool|object|var|void)\b',
                'comment': r'//.*?$|/\*.*?\*/|<!--.*?-->',
                'preprocessor': None,
            },
            'autoit': {
                'keywords': r'\b(And|ByRef|Case|Const|ContinueCase|ContinueLoop|Default|Dim|Do|Else|ElseIf|EndFunc|EndIf|EndSelect|EndSwitch|EndWith|Enum|Exit|ExitLoop|False|For|Func|Global|If|In|Local|Next|Not|Null|Or|Return|Select|Step|Switch|Then|To|True|Until|Volatile|WEnd|While|With)\b',
                'types': r'\b(String|Int|Bool|Binary|Ptr|HWnd)\b',
                'comment': r';.*?$|#cs.*?#ce',
                'preprocessor': r'^\s*#\w+.*?$',
            },
            'bash': {
                'keywords': r'\b(case|do|done|elif|else|esac|fi|for|function|if|in|local|return|select|then|until|while)\b',
                'types': r'\b(true|false)\b',
                'comment': r'#.*?$',
                'preprocessor': r'^#!.*?$',
            },
            'batch': {
                'keywords': r'\b(call|echo|else|exist|exit|for|goto|if|in|not|pause|set|shift|start)\b',
                'types': r'\b(errorlevel|defined|exist)\b',
                'comment': r'^\s*(::|rem\b).*?$',
                'preprocessor': None,
            },
            'caml': {
                'keywords': r'\b(and|as|begin|class|do|done|downto|else|end|exception|external|for|fun|function|if|in|include|inherit|initializer|lazy|let|match|method|module|mutable|new|object|of|open|or|private|rec|sig|struct|then|to|try|type|val|virtual|when|while|with)\b',
                'types': r'\b(array|bool|char|exn|float|int|list|option|string|unit)\b',
                'comment': r'\(\*.*?\*\)',
                'preprocessor': None,
            },
            'c': {
                'keywords': r'\b(auto|break|case|const|continue|default|do|else|enum|extern|for|goto|if|register|return|sizeof|static|struct|switch|typedef|union|volatile|while)\b',
                'types': r'\b(char|double|float|int|long|short|signed|unsigned|void|bool|size_t)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': r'^\s*#.*?$',
            },
            'cpp': {
                'keywords': r'\b(alignas|alignof|auto|break|case|catch|class|const|constexpr|continue|default|delete|do|else|enum|explicit|export|for|friend|goto|if|namespace|new|noexcept|operator|private|protected|public|return|switch|template|this|throw|try|typename|using|virtual|while)\b',
                'types': r'\b(bool|char|double|float|int|long|short|signed|string|unsigned|void|wchar_t|size_t|std)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': r'^\s*#.*?$',
            },
            'csharp': {
                'keywords': r'\b(abstract|as|base|break|case|catch|checked|class|const|continue|default|delegate|do|else|enum|event|explicit|extern|false|finally|fixed|for|foreach|goto|if|implicit|in|interface|internal|is|lock|namespace|new|null|operator|out|override|params|private|protected|public|readonly|ref|return|sealed|sizeof|stackalloc|static|struct|switch|this|throw|true|try|typeof|unchecked|unsafe|using|virtual|void|while)\b',
                'types': r'\b(bool|byte|char|decimal|double|dynamic|float|int|long|object|sbyte|short|string|uint|ulong|ushort|var|void)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': r'^\s*#.*?$',
            },
            'css': {
                'keywords': r'(?<=\{|;|\s)(align-items|background|border|color|display|flex|font|grid|height|justify-content|margin|padding|position|width|z-index)\b',
                'types': r'\b(auto|block|flex|grid|inline|none|relative|absolute|fixed|sticky|solid|dashed)\b',
                'comment': r'/\*.*?\*/',
                'preprocessor': None,
            },
            'diff': {
                'keywords': r'^(---|\+\+\+|@@.*?@@)$',
                'types': r'^(\+.*)$|^(-.*)$',
                'comment': r'^\s*diff\b.*?$|^\s*index\b.*?$',
                'preprocessor': None,
            },
            'fortran': {
                'keywords': r'\b(allocate|call|case|contains|cycle|data|deallocate|do|else|elseif|end|endif|entry|equivalence|exit|format|function|goto|if|implicit|include|inquire|interface|module|namelist|open|parameter|pause|print|program|read|return|save|select|stop|subroutine|then|use|where|while|write)\b',
                'types': r'\b(character|complex|double|integer|logical|real)\b',
                'comment': r'!.*?$',
                'preprocessor': r'^\s*#.*?$',
            },
            'ini': {
                'keywords': r'^\s*\[[^\]]+\]',
                'types': r'^\s*[^=;\r\n]+(?==)',
                'comment': r'^\s*[;#].*$',
                'preprocessor': None,
            },
            'inno': {
                'keywords': r'\b(AppId|AppName|AppVersion|ArchitecturesInstallIn64BitMode|DefaultDirName|DefaultGroupName|OutputBaseFilename|PrivilegesRequired|SetupIconFile|Source|DestDir|Flags|Filename|Name|Description|Types|Tasks)\b',
                'types': r'^\s*\[[^\]]+\]',
                'comment': r';.*?$',
                'preprocessor': None,
            },
            'rust': {
                'keywords': r'\b(as|async|await|break|const|continue|crate|else|enum|extern|fn|for|if|impl|in|let|loop|match|mod|move|mut|pub|ref|return|self|Self|static|struct|super|trait|type|unsafe|use|where|while)\b',
                'types': r'\b(bool|char|f32|f64|i8|i16|i32|i64|i128|isize|str|u8|u16|u32|u64|u128|usize|String)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': None,
            },
            'java': {
                'keywords': r'\b(abstract|assert|break|case|catch|class|const|continue|default|do|else|enum|extends|final|finally|for|if|implements|import|instanceof|interface|native|new|package|private|protected|public|return|static|strictfp|super|switch|synchronized|this|throw|throws|transient|try|volatile|while)\b',
                'types': r'\b(boolean|byte|char|double|float|int|long|short|String|void)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': None,
            },
            'javascript': {
                'keywords': r'\b(await|break|case|catch|class|const|continue|debugger|default|delete|do|else|export|extends|false|finally|for|from|function|if|import|in|instanceof|let|new|null|return|super|switch|this|throw|true|try|typeof|var|void|while|yield)\b',
                'types': r'\b(Array|Boolean|Date|Map|Number|Object|Promise|RegExp|Set|String)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': None,
            },
            'lisp': {
                'keywords': r'\b(and|cond|defmacro|defparameter|defun|dolist|dotimes|if|lambda|let|let\*|loop|nil|or|progn|quote|setq|t|unless|when)\b',
                'types': r'\b(cons|list|string|symbol|vector)\b',
                'comment': r';.*?$',
                'preprocessor': None,
            },
            'makefile': {
                'keywords': r'^\s*[A-Za-z0-9_.-]+\s*:',
                'types': r'\$\([A-Za-z0-9_]+\)|\${[A-Za-z0-9_]+}',
                'comment': r'#.*?$',
                'preprocessor': None,
            },
            'matlab': {
                'keywords': r'\b(break|case|catch|classdef|continue|else|elseif|end|for|function|global|if|otherwise|parfor|persistent|return|spmd|switch|try|while)\b',
                'types': r'\b(cell|char|double|function_handle|int32|logical|single|string|struct|uint32)\b',
                'comment': r'%.*?$',
                'preprocessor': None,
            },
            'nfo': {
                'keywords': None,
                'types': None,
                'comment': None,
                'preprocessor': None,
            },
            'nsis': {
                'keywords': r'\b(Function|FunctionEnd|Section|SectionEnd|IfFileExists|Goto|IntCmp|MessageBox|Return|StrCmp|Var)\b',
                'types': r'\$\w+',
                'comment': r';.*?$|#.*?$',
                'preprocessor': r'^\s*!\w+.*?$',
            },
            'pascal': {
                'keywords': r'\b(and|array|begin|case|class|const|constructor|destructor|div|do|downto|else|end|except|exports|file|finalization|finally|for|function|goto|if|implementation|in|inherited|initialization|inline|interface|label|mod|nil|not|object|of|or|packed|procedure|program|record|repeat|set|shl|shr|then|threadvar|to|try|type|unit|until|uses|var|while|with|xor)\b',
                'types': r'\b(Boolean|Byte|Char|Double|Integer|LongInt|Real|String|Word)\b',
                'comment': r'\{.*?\}|\(\*.*?\*\)|//.*?$',
                'preprocessor': None,
            },
            'perl': {
                'keywords': r'\b(and|cmp|continue|do|else|elsif|eq|for|foreach|ge|goto|gt|if|last|le|local|lt|my|ne|next|not|or|our|package|redo|require|return|sub|unless|until|use|while|xor)\b',
                'types': r'\b(scalar|array|hash)\b',
                'comment': r'#.*?$',
                'preprocessor': None,
            },
            'php': {
                'keywords': r'\b(abstract|as|break|case|catch|class|const|continue|declare|default|do|echo|else|elseif|extends|final|finally|for|foreach|function|global|if|implements|include|include_once|interface|namespace|new|private|protected|public|require|require_once|return|static|switch|throw|trait|try|use|while)\b',
                'types': r'\b(array|bool|callable|float|int|iterable|object|string|void)\b',
                'comment': r'//.*?$|/\*.*?\*/|#.*?$',
                'preprocessor': None,
            },
            'resource': {
                'keywords': r'\b(BEGIN|BITMAP|CURSOR|DIALOG|DLGINIT|END|FONT|ICON|MENU|POPUP|STRINGTABLE|STYLE|VALUE|VERSIONINFO)\b',
                'types': r'\b(ACCELERATORS|AUTO3STATE|CAPTION|CHECKBOX|COMBOBOX|CONTROL|CTEXT|DEFPUSHBUTTON|EDITTEXT|GROUPBOX|LTEXT|PUSHBUTTON|RADIOBUTTON|RTEXT|SCROLLBAR)\b',
                'comment': r'//.*?$|/\*.*?\*/',
                'preprocessor': r'^\s*#.*?$',
            },
            'assembly': {
                'keywords': r'\b(mov|add|sub|mul|div|jmp|je|jne|jg|jge|jl|jle|cmp|call|ret|push|pop|lea|and|or|xor|not|shl|shr|inc|dec|nop)\b',
                'types': r'\b(eax|ebx|ecx|edx|esi|edi|esp|ebp|rax|rbx|rcx|rdx|rsi|rdi|rsp|rbp)\b',
                'comment': r';.*?$',
                'preprocessor': r'^\s*\.[A-Za-z_][\w.]*',
            },
            'smalltalk': {
                'keywords': r'\b(self|super|true|false|nil|thisContext)\b',
                'types': r'\b(Array|Boolean|Character|Dictionary|Integer|Object|String|Symbol)\b',
                'comment': r'"[^"]*"',
                'preprocessor': None,
            },
            'tex': {
                'keywords': r'\\[A-Za-z@]+',
                'types': r'\{[^{}]*\}',
                'comment': r'%.*?$',
                'preprocessor': None,
            },
            'sql': {
                'keywords': r'\b(add|alter|and|as|asc|begin|between|by|case|commit|create|delete|desc|distinct|drop|else|end|exec|exists|from|group|having|if|in|insert|into|is|join|left|like|not|null|on|or|order|primary|procedure|right|rollback|select|set|table|then|top|truncate|union|update|values|view|when|where)\b',
                'types': r'\b(bigint|bit|char|date|datetime|decimal|float|int|money|numeric|nvarchar|text|time|timestamp|tinyint|varchar)\b',
                'comment': r'--.*?$|/\*.*?\*/',
                'preprocessor': None,
            },
            'vb': {
                'keywords': r'\b(AddHandler|And|As|Boolean|ByRef|ByVal|Case|Catch|Class|Const|Continue|Do|Each|Else|ElseIf|End|Enum|Event|Exit|False|Finally|For|Friend|Function|Get|GoTo|Handles|If|Implements|Imports|In|Inherits|Interface|Is|Let|Loop|Me|Module|MustInherit|MyBase|MyClass|Namespace|New|Next|Not|Nothing|On|Or|Overloads|Overrides|Private|Property|Protected|Public|RaiseEvent|ReadOnly|ReDim|Return|Select|Set|Shared|Static|Step|Structure|Sub|SyncLock|Then|Throw|To|True|Try|Using|While|With|WriteOnly|Xor)\b',
                'types': r'\b(Boolean|Byte|Char|Date|Decimal|Double|Integer|Long|Object|Short|Single|String)\b',
                'comment': r"'.*$",
                'preprocessor': r'^\s*#.*?$',
            },
            'vbscript': {
                'keywords': r'\b(Call|Class|Const|Dim|Do|Each|Else|ElseIf|End|Erase|Execute|Exit|False|For|Function|Get|If|In|Is|Loop|Mod|Next|Nothing|Not|Null|On|Option|Private|Public|Randomize|Redim|Rem|Resume|Return|Select|Set|Sub|Then|To|True|Until|While|With)\b',
                'types': r'\b(Boolean|Byte|Date|Double|Empty|Integer|Long|Nothing|Null|Object|Single|String|Variant)\b',
                'comment': r"'.*$|Rem.*$",
                'preprocessor': None,
            },
        }

        self.apply_pattern_tag(doc, r'"([^"\\]|\\.)*"|\'([^\'\\]|\\.)*\'', 'syntax_string', re.MULTILINE)
        self.apply_pattern_tag(doc, r'\b\d+(\.\d+)?\b', 'syntax_number', re.MULTILINE)

        if mode in {'html', 'xml'}:
            self.apply_pattern_tag(doc, r'<!--.*?-->', 'syntax_comment', re.DOTALL)
            self.apply_pattern_tag(doc, r'</?[A-Za-z][^>]*?>', 'syntax_tag', re.MULTILINE)
            self.apply_pattern_tag(doc, r'<\?.*?\?>|<%@.*?%>', 'syntax_preprocessor', re.DOTALL)
        else:
            syntax = language_sets.get(mode)
            if syntax:
                if syntax['comment']:
                    self.apply_pattern_tag(doc, syntax['comment'], 'syntax_comment', re.MULTILINE | re.DOTALL)
                if syntax['keywords']:
                    self.apply_pattern_tag(doc, syntax['keywords'], 'syntax_keyword', re.MULTILINE)
                if syntax['types']:
                    self.apply_pattern_tag(doc, syntax['types'], 'syntax_type', re.MULTILINE)
                if syntax['preprocessor']:
                    self.apply_pattern_tag(doc, syntax['preprocessor'], 'syntax_preprocessor', re.MULTILINE)

    def build_line_index(self, file_path):
        line_starts = [0]
        position = 0
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(self.file_load_chunk_size)
                if not chunk:
                    break
                search_from = 0
                while True:
                    newline_index = chunk.find(b'\n', search_from)
                    if newline_index == -1:
                        break
                    line_starts.append(position + newline_index + 1)
                    search_from = newline_index + 1
                position += len(chunk)
                self.root.update_idletasks()
        return line_starts, position

    def read_virtual_line_window(self, doc, start_line, end_line):
        start_byte = doc['line_starts'][start_line - 1]
        if end_line < doc['total_file_lines']:
            end_byte = doc['line_starts'][end_line]
        else:
            end_byte = doc['file_size_bytes']

        with open(doc['file_path'], 'rb') as f:
            f.seek(start_byte)
            data = f.read(end_byte - start_byte)
        return data.decode('utf-8', errors='replace')

    def load_virtual_window(self, doc, target_line=1):
        total_lines = max(1, doc['total_file_lines'])
        target_line = max(1, min(total_lines, target_line))
        start_line = max(1, target_line - (self.virtual_file_window_lines // 2))
        end_line = min(total_lines, start_line + self.virtual_file_window_lines - 1)
        start_line = max(1, end_line - self.virtual_file_window_lines + 1)

        content = self.read_virtual_line_window(doc, start_line, end_line)
        text = doc['text']
        try:
            current_col = int(text.index(tk.INSERT).split('.')[1])
        except tk.TclError:
            current_col = 0

        doc['suspend_modified_events'] = True
        text.delete('1.0', tk.END)
        text.insert('1.0', content)
        text.edit_modified(False)
        doc['suspend_modified_events'] = False

        doc['window_start_line'] = start_line
        doc['window_end_line'] = end_line

        local_line = max(1, target_line - start_line + 1)
        try:
            line_length = len(text.get(f"{local_line}.0", f"{local_line}.end"))
            text.mark_set(tk.INSERT, f"{local_line}.{min(current_col, line_length)}")
            text.see(f"{local_line}.0")
        except tk.TclError:
            text.mark_set(tk.INSERT, '1.0')

        self.update_vertical_scrollbar(doc['frame'], None, None, None)
        self.update_line_number_gutter(doc)

    def ensure_virtual_line_visible(self, doc, target_line=None):
        if not doc.get('virtual_mode'):
            return

        if target_line is None:
            try:
                local_line = int(doc['text'].index(tk.INSERT).split('.')[0])
            except tk.TclError:
                local_line = 1
            target_line = doc['window_start_line'] + local_line - 1

        target_line = max(1, min(doc['total_file_lines'], target_line))
        local_line = target_line - doc['window_start_line'] + 1
        visible_lines = max(1, doc['window_end_line'] - doc['window_start_line'] + 1)

        if local_line <= self.virtual_file_margin_lines or local_line >= visible_lines - self.virtual_file_margin_lines:
            self.load_virtual_window(doc, target_line)

    def handle_text_activity(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        self.set_last_active_editor_widget(doc['text'])
        if doc.get('virtual_mode'):
            # Don't rebuffer while the user is actively selecting text.
            event_type = str(getattr(event, 'type', ''))
            if event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease'}:
                self.update_line_number_gutter(doc)
                self.update_status()
                return
            try:
                doc['text'].index('sel.first')
                self.update_line_number_gutter(doc)
                self.update_status()
                return
            except tk.TclError:
                pass
            self.ensure_virtual_line_visible(doc)
        self.update_line_number_gutter(doc)
        self.update_status()

    def remember_doc_focus(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if doc:
            self.set_last_active_editor_widget(doc['text'])
        return None

    def handle_text_keypress(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return

        if self.autocomplete_enabled.get() and self.autocomplete_popup and self.autocomplete_target == doc.get('text'):
            if event.keysym == 'Up':
                self.move_autocomplete_selection(-1)
                return "break"
            if event.keysym == 'Down':
                self.move_autocomplete_selection(1)
                return "break"
            if event.keysym in {'Tab', 'Return', 'KP_Enter'}:
                return self.accept_autocomplete_selection()
            if event.keysym == 'Escape':
                self.hide_autocomplete_popup()
                return "break"

        if not doc.get('virtual_mode'):
            return

        navigation_keys = {
            'Up', 'Down', 'Left', 'Right', 'Prior', 'Next', 'Home', 'End',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Escape'
        }
        if event.keysym in navigation_keys or (event.state & 0x4):
            return
        return "break"

    def handle_text_keyrelease(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        if (not self.autocomplete_enabled.get()) or doc.get('virtual_mode') or doc.get('preview_mode') or doc.get('large_file_mode'):
            self.hide_autocomplete_popup()
            return

        ignored_keys = {
            'Up', 'Down', 'Return', 'KP_Enter', 'Tab', 'Escape',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Prior', 'Next', 'Caps_Lock'
        }
        if event.keysym in ignored_keys or (event.state & 0x4):
            return
        self.root.after_idle(lambda current_doc=doc: self.update_autocomplete_for_doc(current_doc))

    def handle_compare_keypress(self, event):
        self.hide_autocomplete_popup()
        navigation_keys = {
            'Up', 'Down', 'Left', 'Right', 'Prior', 'Next', 'Home', 'End',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Escape', 'F3', 'F4'
        }
        if event.keysym in navigation_keys or (event.state & 0x4):
            return
        return "break"

    def get_autocomplete_keywords(self, syntax_mode):
        keyword_map = {
            'python': ['and', 'as', 'assert', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else',
                       'except', 'False', 'finally', 'for', 'from', 'global', 'if', 'import', 'in', 'is',
                       'lambda', 'None', 'nonlocal', 'not', 'or', 'pass', 'raise', 'return', 'self',
                       'True', 'try', 'while', 'with', 'yield'],
            'c': ['auto', 'break', 'case', 'char', 'const', 'continue', 'default', 'do', 'double', 'else',
                  'enum', 'extern', 'float', 'for', 'goto', 'if', 'inline', 'int', 'long', 'register',
                  'return', 'short', 'signed', 'sizeof', 'static', 'struct', 'switch', 'typedef',
                  'union', 'unsigned', 'void', 'volatile', 'while'],
            'cpp': ['auto', 'bool', 'break', 'case', 'catch', 'char', 'class', 'const', 'constexpr',
                    'continue', 'default', 'delete', 'double', 'else', 'enum', 'explicit', 'extern',
                    'false', 'float', 'for', 'friend', 'if', 'inline', 'int', 'namespace', 'new',
                    'nullptr', 'operator', 'private', 'protected', 'public', 'return', 'static',
                    'struct', 'template', 'this', 'throw', 'true', 'try', 'typename', 'using', 'virtual',
                    'void', 'while'],
            'rust': ['as', 'break', 'const', 'continue', 'crate', 'else', 'enum', 'extern', 'false', 'fn',
                     'for', 'if', 'impl', 'in', 'let', 'loop', 'match', 'mod', 'move', 'mut', 'pub',
                     'ref', 'return', 'Self', 'self', 'static', 'struct', 'trait', 'true', 'type',
                     'unsafe', 'use', 'where', 'while'],
            'java': ['abstract', 'boolean', 'break', 'byte', 'case', 'catch', 'char', 'class', 'const',
                     'continue', 'default', 'do', 'double', 'else', 'enum', 'extends', 'final', 'finally',
                     'float', 'for', 'if', 'implements', 'import', 'int', 'interface', 'long', 'new',
                     'null', 'package', 'private', 'protected', 'public', 'return', 'short', 'static',
                     'super', 'switch', 'this', 'throw', 'throws', 'true', 'try', 'void', 'while'],
            'javascript': ['async', 'await', 'break', 'case', 'catch', 'class', 'const', 'continue',
                           'default', 'delete', 'else', 'export', 'extends', 'false', 'finally', 'for',
                           'function', 'if', 'import', 'in', 'let', 'new', 'null', 'return', 'switch',
                           'this', 'throw', 'true', 'try', 'typeof', 'var', 'while', 'yield'],
            'html': ['body', 'button', 'class', 'div', 'form', 'head', 'header', 'html', 'id', 'img',
                     'input', 'label', 'link', 'main', 'meta', 'script', 'section', 'span', 'style', 'title'],
            'php': ['abstract', 'array', 'as', 'break', 'case', 'catch', 'class', 'const', 'continue',
                    'echo', 'else', 'elseif', 'extends', 'false', 'final', 'for', 'foreach', 'function',
                    'if', 'implements', 'include', 'interface', 'namespace', 'new', 'null', 'private',
                    'protected', 'public', 'require', 'return', 'static', 'switch', 'this', 'throw',
                    'trait', 'true', 'try', 'use', 'while'],
            'xml': ['CDATA', 'encoding', 'version', 'xmlns'],
            'sql': ['AND', 'AS', 'BY', 'CREATE', 'DELETE', 'DROP', 'FROM', 'GROUP', 'HAVING', 'INSERT',
                    'INTO', 'JOIN', 'LEFT', 'LIKE', 'NOT', 'NULL', 'ON', 'OR', 'ORDER', 'RIGHT',
                    'SELECT', 'SET', 'TABLE', 'UPDATE', 'VALUES', 'WHERE'],
        }
        return keyword_map.get(syntax_mode or '', [])

    def get_autocomplete_prefix(self, text_widget):
        try:
            insert_index = text_widget.index(tk.INSERT)
            line, _ = map(int, insert_index.split('.'))
            before_cursor = text_widget.get(f"{line}.0", insert_index)
        except tk.TclError:
            return None, None

        match = re.search(r'([A-Za-z_][A-Za-z0-9_]*)$', before_cursor)
        if not match:
            return None, None

        prefix = match.group(1)
        start_col = len(before_cursor) - len(prefix)
        return prefix, f"{line}.{start_col}"

    def get_autocomplete_source_text(self, doc):
        text_widget = doc.get('text')
        if not text_widget:
            return ""
        try:
            total_chars = int(text_widget.count('1.0', 'end-1c', 'chars')[0])
            if total_chars <= 250000:
                return text_widget.get('1.0', 'end-1c')
            current_line = int(text_widget.index(tk.INSERT).split('.')[0])
            start_line = max(1, current_line - 500)
            end_line = current_line + 500
            return text_widget.get(f"{start_line}.0", f"{end_line}.end")
        except tk.TclError:
            return ""

    def get_autocomplete_suggestions(self, doc, prefix):
        if not prefix or not doc:
            return []

        mode = doc.get('syntax_mode') or self.get_syntax_mode(doc)
        prefix_lower = prefix.lower()
        candidates = {}

        for keyword in self.get_autocomplete_keywords(mode):
            if keyword.lower().startswith(prefix_lower) and keyword != prefix:
                candidates[keyword] = True

        for word in re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', self.get_autocomplete_source_text(doc)):
            if word.lower().startswith(prefix_lower) and word != prefix:
                candidates[word] = True

        suggestions = sorted(
            candidates.keys(),
            key=lambda value: (
                0 if value.startswith(prefix) else 1,
                len(value),
                value.lower()
            )
        )
        return suggestions[:12]

    def hide_autocomplete_popup(self):
        popup = self.autocomplete_popup
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.autocomplete_popup = None
        self.autocomplete_listbox = None
        self.autocomplete_target = None
        self.autocomplete_target_doc = None
        self.autocomplete_prefix_start = None

    def move_autocomplete_selection(self, direction):
        listbox = self.autocomplete_listbox
        if listbox is None:
            return
        try:
            size = listbox.size()
            if size <= 0:
                return
            selection = listbox.curselection()
            current_index = selection[0] if selection else 0
            next_index = (current_index + direction) % size
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(next_index)
            listbox.activate(next_index)
            listbox.see(next_index)
        except tk.TclError:
            pass

    def position_autocomplete_popup(self):
        popup = self.autocomplete_popup
        target = self.autocomplete_target
        if popup is None or target is None:
            return
        try:
            bbox = target.bbox(tk.INSERT)
            if bbox is None:
                x = target.winfo_rootx() + 12
                y = target.winfo_rooty() + 24
            else:
                x = target.winfo_rootx() + bbox[0]
                y = target.winfo_rooty() + bbox[1] + bbox[3] + 2
            popup.update_idletasks()
            width = popup.winfo_reqwidth()
            height = popup.winfo_reqheight()
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            x = min(max(8, x), max(8, screen_width - width - 8))
            y = min(max(8, y), max(8, screen_height - height - 8))
            popup.geometry(f"+{x}+{y}")
        except tk.TclError:
            self.hide_autocomplete_popup()

    def show_autocomplete_popup(self, doc, suggestions, prefix_start):
        text_widget = doc.get('text')
        if not text_widget or not suggestions:
            self.hide_autocomplete_popup()
            return

        if self.autocomplete_popup is None or not self.autocomplete_popup.winfo_exists():
            popup = tk.Toplevel(self.root)
            popup.wm_overrideredirect(True)
            popup.configure(bg='#30363d')

            listbox = tk.Listbox(
                popup,
                bg='#161b22',
                fg=self.fg_color,
                selectbackground=self.select_bg,
                selectforeground='white',
                activestyle='none',
                relief='flat',
                borderwidth=0,
                highlightthickness=1,
                highlightbackground='#30363d',
                exportselection=False
            )
            listbox.pack(fill='both', expand=True)
            listbox.bind('<ButtonRelease-1>', lambda e: self.accept_autocomplete_selection())
            listbox.bind('<Double-Button-1>', lambda e: self.accept_autocomplete_selection())

            self.autocomplete_popup = popup
            self.autocomplete_listbox = listbox

        self.autocomplete_target = text_widget
        self.autocomplete_target_doc = doc
        self.autocomplete_prefix_start = prefix_start

        listbox = self.autocomplete_listbox
        listbox.delete(0, tk.END)
        for suggestion in suggestions:
            listbox.insert(tk.END, suggestion)
        if suggestions:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.see(0)

        self.position_autocomplete_popup()
        self.autocomplete_popup.lift()

    def accept_autocomplete_selection(self, event=None):
        if not self.autocomplete_listbox or not self.autocomplete_target or not self.autocomplete_prefix_start:
            return "break"
        try:
            selection = self.autocomplete_listbox.curselection()
            if not selection:
                return "break"
            suggestion = self.autocomplete_listbox.get(selection[0])
            target = self.autocomplete_target
            target.delete(self.autocomplete_prefix_start, tk.INSERT)
            target.insert(self.autocomplete_prefix_start, suggestion)
            target.mark_set(tk.INSERT, f"{self.autocomplete_prefix_start}+{len(suggestion)}c")
            target.see(tk.INSERT)
            target.focus_set()
            try:
                target.edit_modified(True)
            except tk.TclError:
                pass
            doc = self.autocomplete_target_doc
            self.hide_autocomplete_popup()
            if doc:
                self.on_text_modified(doc['frame'])
            else:
                self.update_status()
        except tk.TclError:
            self.hide_autocomplete_popup()
        return "break"

    def update_autocomplete_for_doc(self, doc):
        if not self.autocomplete_enabled.get():
            self.hide_autocomplete_popup()
            return
        if not doc or self.get_current_doc() != doc:
            self.hide_autocomplete_popup()
            return
        text_widget = doc.get('text')
        if text_widget is None or self.root.focus_get() != text_widget:
            self.hide_autocomplete_popup()
            return

        prefix, prefix_start = self.get_autocomplete_prefix(text_widget)
        if not prefix:
            self.hide_autocomplete_popup()
            return

        suggestions = self.get_autocomplete_suggestions(doc, prefix)
        if not suggestions:
            self.hide_autocomplete_popup()
            return

        self.show_autocomplete_popup(doc, suggestions, prefix_start)

    def on_text_mousewheel(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc or not doc.get('virtual_mode'):
            self.update_status()
            return

        text = doc['text']
        if hasattr(event, 'delta') and event.delta:
            delta = -1 if event.delta > 0 else 1
        elif getattr(event, 'num', None) == 4:
            delta = -1
        else:
            delta = 1

        visible_lines = max(1, doc['window_end_line'] - doc['window_start_line'] + 1)
        try:
            top_local_line = int(text.index('@0,0').split('.')[0])
        except tk.TclError:
            top_local_line = 1
        try:
            bottom_local_line = int(text.index(f"@0,{max(0, text.winfo_height() - 1)}").split('.')[0])
        except tk.TclError:
            bottom_local_line = visible_lines

        if delta < 0:
            if top_local_line > 1:
                text.yview_scroll(-3, 'units')
            elif doc['window_start_line'] > 1:
                self.load_virtual_window(doc, max(1, doc['window_start_line'] - 20))
        else:
            if bottom_local_line < visible_lines:
                text.yview_scroll(3, 'units')
            elif doc['window_end_line'] < doc['total_file_lines']:
                self.load_virtual_window(doc, min(doc['total_file_lines'], doc['window_end_line'] + 20))
        self.update_line_number_gutter(doc)
        self.update_status()
        return "break"

    def on_vertical_scroll(self, tab_id, *args):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return

        if not doc.get('virtual_mode'):
            doc['text'].yview(*args)
            self.update_line_number_gutter(doc)
            return

        if not args:
            return

        if args[0] == 'moveto':
            fraction = float(args[1])
            target_line = int(fraction * max(0, doc['total_file_lines'] - 1)) + 1
        elif args[0] == 'scroll':
            count = int(args[1])
            unit = args[2]
            step = 1 if unit == 'units' else max(1, self.virtual_file_window_lines // 2)
            target_line = doc['window_start_line'] + (count * step)
        else:
            return

        self.load_virtual_window(doc, target_line)
        self.update_status()

    def update_vertical_scrollbar(self, tab_id, scrollbar, first, last):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return

        if scrollbar is None:
            scrollbar = doc['v_scroll']

        if not doc.get('virtual_mode'):
            if first is not None and last is not None:
                scrollbar.set(first, last)
            self.update_line_number_gutter(doc)
            return

        total_lines = max(1, doc['total_file_lines'])
        first_fraction = (doc['window_start_line'] - 1) / total_lines
        last_fraction = doc['window_end_line'] / total_lines
        scrollbar.set(first_fraction, min(1.0, last_fraction))
        self.update_line_number_gutter(doc)

    def on_compare_vertical_scroll(self, *args):
        if not self.compare_view or not self.compare_view.get('text'):
            return
        self.compare_view['text'].yview(*args)
        self.update_line_number_gutter(self.compare_view)

    def update_compare_vertical_scrollbar(self, scrollbar, first, last):
        scrollbar.set(first, last)
        self.update_line_number_gutter(self.compare_view)

    def create_text_context_menu(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return

        menu = tk.Menu(self.root, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        menu.add_command(label="Cut", command=self.cut_or_close_panel)
        menu.add_command(label="Copy", command=self.copy)
        menu.add_command(label="Paste", command=self.paste)
        menu.add_separator()
        menu.add_command(label="Select All", command=self.select_all)
        menu.add_command(label="Add note", command=lambda frame=tab_id: self.add_note_to_selection(frame))
        menu.add_command(label="Allow change", command=lambda frame=tab_id: self.approve_note(frame))
        menu.add_command(label="Deny change", command=lambda frame=tab_id: self.dissapprove_note(frame))
        menu.add_command(label="Remove note", command=lambda frame=tab_id: self.remove_note(frame))
        doc['context_menu'] = menu

    def show_text_context_menu(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        self.hide_autocomplete_popup()

        self.notebook.select(doc['frame'])
        self.set_active_document(doc['frame'])

        selection_ranges = self.text.tag_ranges('sel')
        note_state = 'normal' if len(selection_ranges) >= 2 and str(selection_ranges[0]) != str(selection_ranges[1]) else 'disabled'

        index = self.text.index(f"@{event.x},{event.y}")
        clicked_note_tags = [tag for tag in self.text.tag_names(index) if tag in doc['notes']]
        doc['context_note_tag'] = clicked_note_tags[-1] if clicked_note_tags else None

        if self.current_doc_is_large_readonly():
            doc['context_menu'].entryconfig("Cut", state='disabled')
            doc['context_menu'].entryconfig("Paste", state='disabled')
            note_state = 'disabled'
        else:
            doc['context_menu'].entryconfig("Cut", state='normal')
            doc['context_menu'].entryconfig("Paste", state='normal')

        doc['context_menu'].entryconfig("Add note", state=note_state)
        note_action_state = 'normal' if doc.get('context_note_tag') else 'disabled'
        approve_state = note_action_state
        dissapprove_state = note_action_state
        if doc.get('context_note_tag'):
            note_data = doc['notes'].get(doc['context_note_tag'], {})
            if note_data.get('approved_by'):
                approve_state = 'disabled'
            if note_data.get('dissapproved_by'):
                dissapprove_state = 'disabled'
        doc['context_menu'].entryconfig("Allow change", state=approve_state)
        doc['context_menu'].entryconfig("Deny change", state=dissapprove_state)
        doc['context_menu'].entryconfig("Remove note", state=note_action_state)
        try:
            doc['context_menu'].tk_popup(event.x_root, event.y_root)
        finally:
            doc['context_menu'].grab_release()
        return "break"

    def hide_note_popup(self, doc):
        popup = doc.get('note_popup')
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            doc['note_popup'] = None

    def show_note_popup(self, doc, note_data, x, y):
        self.hide_note_popup(doc)

        popup = tk.Toplevel(self.root)
        popup.wm_overrideredirect(True)
        popup.configure(bg='#1f2430')

        if note_data.get('dissapproved_by'):
            highlight_color = '#f85149'
            status_text = f"Denied by: {note_data['dissapproved_by']}"
            status_fg = '#ffb3ad'
            reply_text = note_data.get('dissapproved_note')
        elif note_data.get('approved_by'):
            highlight_color = '#3fb950'
            status_text = f"Allowed by: {note_data['approved_by']}"
            status_fg = '#86efac'
            reply_text = note_data.get('approved_note')
        else:
            highlight_color = self.note_bg
            status_text = None
            status_fg = None
            reply_text = None
        frame = tk.Frame(popup, bg='#1f2430', highlightthickness=1, highlightbackground=highlight_color)
        frame.pack(fill='both', expand=True)

        tk.Label(
            frame,
            text="Code note",
            bg='#1f2430',
            fg=highlight_color,
            font=('Segoe UI', 9, 'bold'),
            anchor='w'
        ).pack(fill='x', padx=10, pady=(8, 2))

        meta_parts = []
        if note_data.get('author_label'):
            meta_parts.append(f"Author: {note_data['author_label']}")
        elif note_data.get('author_id'):
            meta_parts.append(f"Author ID: {note_data['author_id']}")
        if note_data.get('created_at'):
            meta_parts.append(f"Created: {note_data['created_at'].replace('T', ' ')}")
        if meta_parts:
            tk.Label(
                frame,
                text=" | ".join(meta_parts),
                bg='#1f2430',
                fg='#9aa0a6',
                font=('Segoe UI', 8),
                justify='left',
                wraplength=280,
                anchor='w'
            ).pack(fill='x', padx=10, pady=(0, 6))

        tk.Label(
            frame,
            text=note_data['text'],
            bg='#1f2430',
            fg='#f5f5f5',
            font=('Segoe UI', 9),
            justify='left',
            wraplength=280,
            anchor='w'
        ).pack(fill='both', expand=True, padx=10, pady=(0, 8))

        if status_text:
            tk.Label(
                frame,
                text=status_text,
                bg='#1f2430',
                fg=status_fg,
                font=('Segoe UI', 9, 'bold'),
                justify='left',
                anchor='w'
            ).pack(fill='x', padx=10, pady=(0, 8))

        if reply_text:
            tk.Label(
                frame,
                text="Code note replay",
                bg='#1f2430',
                fg=highlight_color,
                font=('Segoe UI', 9, 'bold'),
                anchor='w'
            ).pack(fill='x', padx=10, pady=(0, 2))

            tk.Label(
                frame,
                text=reply_text,
                bg='#1f2430',
                fg='#f5f5f5',
                font=('Segoe UI', 9),
                justify='left',
                wraplength=280,
                anchor='w'
            ).pack(fill='both', expand=True, padx=10, pady=(0, 8))

        popup.update_idletasks()
        popup.geometry(f"+{x + 12}+{y + 12}")
        popup.bind('<Escape>', lambda e, current=doc: self.hide_note_popup(current))
        doc['note_popup'] = popup

    def prompt_note_input(self, title, prompt, initialvalue="", parent=None):
        parent = parent or self.root
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)

        result = {'value': None}
        value_var = tk.StringVar(value=initialvalue)

        tk.Label(
            dialog,
            text=prompt,
            bg='#f0f0f0',
            fg='black',
            font=('Segoe UI', 9)
        ).pack(anchor='w', pady=(0, 6))

        entry = tk.Entry(dialog, textvariable=value_var, width=34)
        entry.pack(fill='x')

        button_row = tk.Frame(dialog, bg='#f0f0f0')
        button_row.pack(fill='x', pady=(10, 0))

        def submit(event=None):
            result['value'] = value_var.get()
            dialog.destroy()

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()

        tk.Button(button_row, text="OK", width=10, command=submit).pack(side='left')
        tk.Button(button_row, text="Cancel", width=10, command=cancel).pack(side='right')

        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol("WM_DELETE_WINDOW", cancel)

        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        entry.focus_force()
        entry.selection_range(0, tk.END)
        dialog.wait_visibility()
        dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
        parent.wait_window(dialog)
        return result['value']

    def export_notes_report(self):
        doc = self.get_current_doc()
        if not doc or not doc.get('notes'):
            messagebox.showinfo("Export Notes", "There are no notes to export in this tab.", parent=self.root)
            return "break"
        initial_name = f"{self.get_doc_name(doc['frame'])}-notes.json"
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".json",
            initialfile=initial_name,
            filetypes=[("JSON", "*.json"), ("Markdown", "*.md"), ("All Files", "*.*")]
        )
        if not output_path:
            return "break"
        ordered_tags = self.get_ordered_note_tags(doc)
        note_rows = []
        for note_tag in ordered_tags:
            note_data = doc['notes'][note_tag]
            ranges = doc['text'].tag_ranges(note_tag)
            if len(ranges) < 2:
                continue
            row = {
                'id': note_data.get('id'),
                'selection_start': str(ranges[0]),
                'selection_end': str(ranges[1]),
                'selected_text': doc['text'].get(str(ranges[0]), str(ranges[1])),
                'note': note_data.get('text'),
                'author': note_data.get('author_label') or note_data.get('author_id'),
                'created_at': note_data.get('created_at'),
                'allowed_by': note_data.get('approved_by'),
                'allowed_note': note_data.get('approved_note'),
                'denied_by': note_data.get('dissapproved_by'),
                'denied_note': note_data.get('dissapproved_note'),
            }
            note_rows.append(row)
        try:
            if output_path.lower().endswith('.md'):
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(f"# Notes Export for {self.get_doc_name(doc['frame'])}\n\n")
                    for row in note_rows:
                        f.write(f"## Note {row['id']}\n\n")
                        f.write(f"- Range: `{row['selection_start']}` to `{row['selection_end']}`\n")
                        if row['author']:
                            f.write(f"- Author: {row['author']}\n")
                        if row['created_at']:
                            f.write(f"- Created: {row['created_at']}\n")
                        if row['allowed_by']:
                            f.write(f"- Allowed by: {row['allowed_by']}\n")
                        if row['denied_by']:
                            f.write(f"- Denied by: {row['denied_by']}\n")
                        f.write("\n### Selected Text\n\n```\n")
                        f.write(row['selected_text'])
                        f.write("\n```\n\n### Code Note\n\n")
                        f.write(row['note'] or '')
                        if row['allowed_note']:
                            f.write("\n\n### Allow Reply\n\n")
                            f.write(row['allowed_note'])
                        if row['denied_note']:
                            f.write("\n\n### Deny Reply\n\n")
                            f.write(row['denied_note'])
                        f.write("\n\n")
            else:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(note_rows, f, indent=2)
            messagebox.showinfo("Export Notes", f"Notes exported to:\n{output_path}", parent=self.root)
        except OSError as exc:
            self.log_exception("export notes report", exc)
            messagebox.showerror("Export Notes", str(exc), parent=self.root)
        return "break"

    def add_note_to_selection(self, tab_id=None, event=None):
        doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or self.current_doc_is_large_readonly():
            return "break"

        self.notebook.select(doc['frame'])
        self.set_active_document(doc['frame'])

        try:
            start = self.text.index('sel.first')
            end = self.text.index('sel.last')
        except tk.TclError:
            messagebox.showinfo("Add Note", "Select some text first.", parent=self.root)
            return "break"

        self.hide_note_popup(doc)
        note_input = self.prompt_note_input("Add Note", "Note:", parent=self.root)
        if not note_input:
            return "break"

        note_text = note_input.strip()
        if note_text.lower().startswith("# note:"):
            note_text = note_text[7:].strip()
        if not note_text:
            return "break"

        note_tag = self.create_note_tag(
            doc,
            start,
            end,
            note_text,
            author_id=self.editor_id,
            author_label=self.get_doc_editor_label(doc),
            read_by=[self.editor_id]
        )
        self.persist_doc_notes(doc)
        return "break"

    def create_note_tag(
        self,
        doc,
        start,
        end,
        note_text,
        approved_by=None,
        dissapproved_by=None,
        note_id=None,
        author_id=None,
        author_label=None,
        read_by=None,
        author_unread=False,
        created_at=None,
        anchor_text=None,
        anchor_line=None
    ):
        if note_id is None:
            note_id = doc['note_counter']
        doc['note_counter'] = max(doc['note_counter'], int(note_id) + 1)
        note_tag = f"note_{note_id}"
        normalized_read_by = []
        for editor_id in read_by or []:
            editor_text = str(editor_id).strip()
            if editor_text and editor_text not in normalized_read_by:
                normalized_read_by.append(editor_text)
        doc['notes'][note_tag] = {
            'id': str(note_id),
            'text': note_text,
            'approved_by': approved_by,
            'dissapproved_by': dissapproved_by,
            'approved_note': None,
            'dissapproved_note': None,
            'author_id': str(author_id).strip() if author_id else None,
            'author_label': str(author_label).strip() if author_label else None,
            'read_by': normalized_read_by,
            'author_unread': bool(author_unread),
            'created_at': created_at or datetime.now().isoformat(timespec='seconds'),
            'anchor_text': anchor_text if anchor_text is not None else doc['text'].get(start, end),
            'anchor_line': anchor_line if anchor_line is not None else int(str(start).split('.')[0]),
        }
        self.apply_note_tag(doc, note_tag, start, end)
        return note_tag

    def apply_note_tag(self, doc, note_tag, start, end):
        note_data = doc['notes'][note_tag]
        if note_data.get('dissapproved_by'):
            highlight_color = '#f85149'
        elif note_data.get('approved_by'):
            highlight_color = '#7ee787'
        else:
            highlight_color = self.note_bg
        doc['text'].tag_add(note_tag, start, end)
        doc['text'].tag_config(note_tag, background=highlight_color, foreground='black')
        doc['text'].tag_bind(note_tag, '<Button-1>', lambda e, frame=doc['frame'], tag=note_tag: self.open_note_from_tag(e, frame, tag))
        doc['text'].tag_bind(note_tag, '<Enter>', lambda e, text=doc['text']: text.config(cursor='hand2'))
        doc['text'].tag_bind(note_tag, '<Leave>', lambda e, text=doc['text']: text.config(cursor='xterm'))

    def approve_note(self, tab_id=None):
        doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        self.hide_note_popup(doc)
        approver = self.prompt_note_input("Allow Change", "Name:", parent=self.root)
        if not approver:
            return "break"
        review_note = self.prompt_note_input("Allow Change", "Why:", parent=self.root)
        if review_note is None:
            return "break"

        note_tag = doc['context_note_tag']
        note_data = doc['notes'].get(note_tag)
        if not note_data:
            return "break"

        note_data['approved_by'] = approver.strip()
        note_data['dissapproved_by'] = None
        note_data['approved_note'] = review_note.strip()
        note_data['dissapproved_note'] = None
        if note_data.get('author_id') and note_data.get('author_id') not in self.editor_aliases:
            note_data['author_unread'] = True
            note_data['read_by'] = [
                str(editor_id) for editor_id in note_data.get('read_by', [])
                if str(editor_id).strip() and str(editor_id).strip() != note_data.get('author_id')
            ]
        ranges = doc['text'].tag_ranges(note_tag)
        if len(ranges) >= 2:
            doc['text'].tag_remove(note_tag, '1.0', tk.END)
            self.apply_note_tag(doc, note_tag, str(ranges[0]), str(ranges[1]))
        self.persist_doc_notes(doc)
        return "break"

    def dissapprove_note(self, tab_id=None):
        doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        self.hide_note_popup(doc)
        reviewer = self.prompt_note_input("Deny Change", "Name:", parent=self.root)
        if not reviewer:
            return "break"
        review_note = self.prompt_note_input("Deny Change", "Why:", parent=self.root)
        if review_note is None:
            return "break"

        note_tag = doc['context_note_tag']
        note_data = doc['notes'].get(note_tag)
        if not note_data:
            return "break"

        note_data['dissapproved_by'] = reviewer.strip()
        note_data['approved_by'] = None
        note_data['dissapproved_note'] = review_note.strip()
        note_data['approved_note'] = None
        if note_data.get('author_id') and note_data.get('author_id') not in self.editor_aliases:
            note_data['author_unread'] = True
            note_data['read_by'] = [
                str(editor_id) for editor_id in note_data.get('read_by', [])
                if str(editor_id).strip() and str(editor_id).strip() != note_data.get('author_id')
            ]
        ranges = doc['text'].tag_ranges(note_tag)
        if len(ranges) >= 2:
            doc['text'].tag_remove(note_tag, '1.0', tk.END)
            self.apply_note_tag(doc, note_tag, str(ranges[0]), str(ranges[1]))
        self.persist_doc_notes(doc)
        return "break"

    def remove_note(self, tab_id=None):
        doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        note_tag = doc['context_note_tag']
        doc['text'].tag_delete(note_tag)
        doc['text'].config(cursor='xterm')
        doc['notes'].pop(note_tag, None)
        doc['context_note_tag'] = None
        self.hide_note_popup(doc)
        self.persist_doc_notes(doc)
        self.play_delete_note_sound()
        return "break"

    def open_note_from_tag(self, event, tab_id, note_tag):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        note_data = doc['notes'].get(note_tag)
        if not note_data:
            return
        self.mark_note_as_read(doc, note_tag)
        self.show_note_popup(doc, note_data, event.x_root, event.y_root)
        self.update_status()
        return "break"

    def on_text_click_release(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        self.hide_autocomplete_popup()

        index = doc['text'].index(f"@{event.x},{event.y}")
        tags = doc['text'].tag_names(index)
        note_tags = [tag for tag in tags if tag in doc['notes']]
        if note_tags:
            self.mark_note_as_read(doc, note_tags[-1])
            note_data = doc['notes'][note_tags[-1]]
            self.show_note_popup(doc, note_data, event.x_root, event.y_root)
            self.update_status()
        else:
            self.hide_note_popup(doc)

    def get_unread_note_tags(self, doc):
        unread = []
        for note_tag, note_data in doc['notes'].items():
            note_id = str(note_data.get('id', ''))
            if note_data.get('author_id') in self.editor_aliases:
                if note_data.get('author_unread'):
                    unread.append((int(note_id) if note_id.isdigit() else 0, note_tag))
                continue
            read_by = {str(editor_id) for editor_id in note_data.get('read_by', []) if str(editor_id).strip()}
            if note_id and not (self.editor_aliases & read_by):
                unread.append((int(note_id) if note_id.isdigit() else 0, note_tag))
        unread.sort(key=lambda item: item[0])
        return [tag for _, tag in unread]

    def get_unread_note_count(self, doc):
        return len(self.get_unread_note_tags(doc))

    def mark_note_as_read(self, doc, note_tag):
        note_data = doc['notes'].get(note_tag)
        if not note_data or not note_data.get('id'):
            return
        changed = False
        if note_data.get('author_id') in self.editor_aliases and note_data.get('author_unread'):
            note_data['author_unread'] = False
            changed = True
        read_by = [str(editor_id) for editor_id in note_data.get('read_by', []) if str(editor_id).strip()]
        if not (self.editor_aliases & set(read_by)):
            read_by.append(self.editor_id)
            note_data['read_by'] = read_by
            changed = True
        if changed and doc.get('file_path') and not doc.get('virtual_mode') and not doc.get('preview_mode'):
            self.persist_doc_notes(doc)

    def get_notes_sidecar_path(self, file_path):
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        return os.path.join(directory, f"{filename}.notepadx.notes.json")

    def get_notes_sidecar_signature(self, sidecar_path):
        try:
            stat = os.stat(sidecar_path)
        except OSError:
            return None
        return (stat.st_mtime_ns, stat.st_size)

    def sanitize_shared_editors(self, editors):
        sanitized = []
        seen_ids = set()
        for entry in editors or []:
            if not isinstance(entry, dict):
                continue
            editor_id = str(entry.get('id', '')).strip()
            if not editor_id or editor_id in seen_ids:
                continue
            label = str(entry.get('label', '')).strip() or None
            pid = entry.get('pid')
            if isinstance(pid, str) and pid.isdigit():
                pid = int(pid)
            if not isinstance(pid, int) or pid <= 0:
                pid = None
            sanitized.append({'id': editor_id, 'label': label, 'pid': pid})
            seen_ids.add(editor_id)
        return sanitized

    def is_editor_process_alive(self, pid):
        if not isinstance(pid, int) or pid <= 0:
            return False
        if pid == os.getpid():
            return True
        try:
            os.kill(pid, 0)
        except Exception:
            return False
        return True

    def prune_inactive_shared_editors(self, editors):
        return [
            entry for entry in self.sanitize_shared_editors(editors)
            if self.is_editor_process_alive(entry.get('pid'))
        ]

    def allocate_editor_label(self, editors):
        used_numbers = set()
        for entry in editors:
            label = entry.get('label') or ''
            match = re.fullmatch(r'Notepad-X-(\d+)', label)
            if match:
                used_numbers.add(int(match.group(1)))

        label_number = 1
        while label_number in used_numbers:
            label_number += 1
        return f"Notepad-X-{label_number}"

    def get_doc_editor_label(self, doc):
        return doc.get('note_editor_label') or 'Notepad-X'

    def export_doc_notes(self, doc):
        exported = []
        for note_tag, note_data in doc['notes'].items():
            ranges = doc['text'].tag_ranges(note_tag)
            if len(ranges) >= 2:
                exported.append({
                    'id': note_data.get('id'),
                    'start': str(ranges[0]),
                    'end': str(ranges[1]),
                    'text': note_data['text'],
                    'approved_by': note_data.get('approved_by'),
                    'dissapproved_by': note_data.get('dissapproved_by'),
                    'approved_note': note_data.get('approved_note'),
                    'dissapproved_note': note_data.get('dissapproved_note'),
                    'author_id': note_data.get('author_id'),
                    'author_label': note_data.get('author_label'),
                    'read_by': [str(editor_id) for editor_id in note_data.get('read_by', []) if str(editor_id).strip()],
                    'author_unread': bool(note_data.get('author_unread')),
                    'created_at': note_data.get('created_at'),
                    'anchor_text': note_data.get('anchor_text'),
                    'anchor_line': note_data.get('anchor_line'),
                })
        return exported

    def normalize_optional_metadata(self, value):
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text.lower() in {'none', 'null'}:
            return None
        return text

    def sanitize_note_payload(self, saved_note):
        if not isinstance(saved_note, dict):
            return None
        note_text = str(saved_note.get('text', '')).strip()
        start = str(saved_note.get('start', '')).strip()
        end = str(saved_note.get('end', '')).strip()
        note_id = str(saved_note.get('id', '')).strip()
        if not note_text or not start or not end:
            return None
        read_by = saved_note.get('read_by', [])
        if not isinstance(read_by, list):
            read_by = []
        anchor_line = saved_note.get('anchor_line')
        try:
            anchor_line = int(anchor_line) if anchor_line is not None else None
        except (TypeError, ValueError):
            anchor_line = None
        return {
            'id': note_id or None,
            'start': start,
            'end': end,
            'text': note_text,
            'approved_by': self.normalize_optional_metadata(saved_note.get('approved_by')),
            'dissapproved_by': self.normalize_optional_metadata(saved_note.get('dissapproved_by')),
            'approved_note': self.normalize_optional_metadata(saved_note.get('approved_note')),
            'dissapproved_note': self.normalize_optional_metadata(saved_note.get('dissapproved_note')),
            'author_id': self.normalize_optional_metadata(saved_note.get('author_id')),
            'author_label': self.normalize_optional_metadata(saved_note.get('author_label')),
            'read_by': [str(editor_id).strip() for editor_id in read_by if str(editor_id).strip()],
            'author_unread': bool(saved_note.get('author_unread', False)),
            'created_at': self.normalize_optional_metadata(saved_note.get('created_at')),
            'anchor_text': str(saved_note.get('anchor_text', '')),
            'anchor_line': anchor_line,
        }

    def resolve_note_range(self, doc, saved_note):
        start = saved_note.get('start')
        end = saved_note.get('end')
        anchor_text = saved_note.get('anchor_text') or ''
        if start and end:
            try:
                if doc['text'].get(start, end) == anchor_text or not anchor_text:
                    return start, end
            except tk.TclError:
                pass
        if not anchor_text:
            return start, end
        content = doc['text'].get('1.0', 'end-1c')
        positions = []
        search_from = 0
        while True:
            found = content.find(anchor_text, search_from)
            if found == -1:
                break
            positions.append(found)
            search_from = found + max(1, len(anchor_text))
            if len(positions) >= 50:
                break
        if not positions:
            return start, end
        if saved_note.get('anchor_line'):
            target_offset = doc['text'].count('1.0', f"{saved_note['anchor_line']}.0", 'chars')[0]
            best_pos = min(positions, key=lambda pos: abs(pos - target_offset))
        else:
            best_pos = positions[0]
        return f"1.0+{best_pos}c", f"1.0+{best_pos + len(anchor_text)}c"

    def clear_doc_notes(self, doc):
        for note_tag in list(doc['notes'].keys()):
            try:
                doc['text'].tag_delete(note_tag)
            except tk.TclError:
                pass
        doc['notes'] = {}
        doc['note_counter'] = 1
        doc['context_note_tag'] = None
        self.hide_note_popup(doc)

    def write_shared_notes(self, sidecar_path, notes_payload, active_editors=0, editors=None):
        os.makedirs(os.path.dirname(sidecar_path), exist_ok=True)
        fd, temp_path = tempfile.mkstemp(prefix='notepadx-notes-', suffix='.tmp', dir=os.path.dirname(sidecar_path))
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as temp_file:
                sanitized_editors = self.prune_inactive_shared_editors(editors)
                json.dump({
                    'active_editors': len(sanitized_editors),
                    'editors': sanitized_editors,
                    'notes': notes_payload
                }, temp_file, indent=2)
            os.replace(temp_path, sidecar_path)
            self.hide_support_file(sidecar_path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def load_shared_notes(self, sidecar_path):
        if not os.path.exists(sidecar_path):
            return {'active_editors': 0, 'editors': [], 'notes': []}
        try:
            with open(sidecar_path, 'r', encoding='utf-8') as f:
                payload = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            self.log_exception("load shared notes", exc)
            return {'active_editors': 0, 'editors': [], 'notes': []}
        notes = payload.get('notes', [])
        editors = self.prune_inactive_shared_editors(payload.get('editors', []))
        sanitized_notes = []
        dirty_payload = False
        for note in notes if isinstance(notes, list) else []:
            sanitized_note = self.sanitize_note_payload(note)
            if sanitized_note is not None:
                sanitized_notes.append(sanitized_note)
                if sanitized_note != note:
                    dirty_payload = True
            else:
                dirty_payload = True
        if dirty_payload:
            try:
                self.write_shared_notes(sidecar_path, sanitized_notes, len(editors), editors)
            except OSError as exc:
                self.log_exception("rewrite sanitized shared notes", exc)
        return {
            'active_editors': len(editors),
            'editors': editors,
            'notes': sanitized_notes
        }

    def persist_doc_notes(self, doc):
        if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return

        exported = self.export_doc_notes(doc)
        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        try:
            if exported or doc.get('note_active_editors', 0) > 0 or os.path.exists(sidecar_path):
                self.write_shared_notes(sidecar_path, exported, doc.get('note_active_editors', 0), doc.get('note_editors', []))
                doc['note_sync_mtime'] = os.path.getmtime(sidecar_path)
                doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
        except OSError:
            pass
        self.save_session()

    def restore_doc_notes(self, doc):
        if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        shared_payload = self.load_shared_notes(sidecar_path)
        saved_notes = shared_payload.get('notes', [])
        doc['note_editors'] = self.sanitize_shared_editors(shared_payload.get('editors', []))
        doc['note_active_editors'] = max(shared_payload.get('active_editors', 0), len(doc['note_editors']))
        existing_editor = next((entry for entry in doc['note_editors'] if entry.get('id') == self.editor_id), None)
        doc['note_editor_label'] = existing_editor.get('label') if existing_editor else None
        self.clear_doc_notes(doc)
        if not saved_notes:
            doc['note_sync_mtime'] = os.path.getmtime(sidecar_path) if os.path.exists(sidecar_path) else None
            doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
            return

        for saved_note in saved_notes:
            resolved_range = self.resolve_note_range(doc, saved_note)
            start, end = resolved_range
            note_id = saved_note.get('id')
            note_text = saved_note.get('text', '').strip()
            approved_by = saved_note.get('approved_by')
            dissapproved_by = saved_note.get('dissapproved_by')
            approved_note = saved_note.get('approved_note')
            dissapproved_note = saved_note.get('dissapproved_note')
            author_id = saved_note.get('author_id')
            author_label = saved_note.get('author_label')
            read_by = saved_note.get('read_by', [])
            author_unread = saved_note.get('author_unread', False)
            created_at = saved_note.get('created_at')
            anchor_text = saved_note.get('anchor_text')
            anchor_line = saved_note.get('anchor_line')
            if not start or not end or not note_text:
                continue
            try:
                note_tag = self.create_note_tag(
                    doc, start, end, note_text,
                    approved_by=approved_by,
                    dissapproved_by=dissapproved_by,
                    note_id=note_id,
                    author_id=author_id,
                    author_label=author_label,
                    read_by=read_by,
                    author_unread=author_unread,
                    created_at=created_at,
                    anchor_text=anchor_text,
                    anchor_line=anchor_line
                )
            except tk.TclError as exc:
                self.log_exception("restore note tag", exc)
                continue
            restored_note = doc['notes'].get(note_tag)
            if restored_note is not None:
                restored_note['approved_note'] = approved_note
                restored_note['dissapproved_note'] = dissapproved_note
        doc['note_sync_mtime'] = os.path.getmtime(sidecar_path) if os.path.exists(sidecar_path) else None
        doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
        doc['last_unread_count'] = self.get_unread_note_count(doc)

    def register_doc_for_shared_notes(self, doc):
        if not doc.get('file_path') or doc.get('notes_registered'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        payload = self.load_shared_notes(sidecar_path)
        editors = self.sanitize_shared_editors(payload.get('editors', []))
        existing_editor = next((entry for entry in editors if entry.get('id') == self.editor_id), None)
        if existing_editor is None:
            existing_editor = {
                'id': self.editor_id,
                'label': self.allocate_editor_label(editors),
                'pid': os.getpid()
            }
            editors.append(existing_editor)
        else:
            existing_editor['pid'] = os.getpid()
        doc['note_editors'] = editors
        doc['note_editor_label'] = existing_editor.get('label')
        doc['note_active_editors'] = len(editors)
        doc['notes_registered'] = True
        try:
            self.write_shared_notes(sidecar_path, payload.get('notes', []), doc['note_active_editors'], editors)
            doc['note_sync_mtime'] = os.path.getmtime(sidecar_path)
            doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
        except OSError:
            doc['notes_registered'] = False
        self.update_status()

    def unregister_doc_from_shared_notes(self, doc):
        if not doc.get('file_path') or not doc.get('notes_registered'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        payload = self.load_shared_notes(sidecar_path)
        editors = [
            entry for entry in self.sanitize_shared_editors(payload.get('editors', []))
            if entry.get('id') != self.editor_id
        ]
        active_editors = len(editors)
        doc['note_editors'] = editors
        doc['note_active_editors'] = active_editors
        doc['note_editor_label'] = None
        doc['notes_registered'] = False
        try:
            self.write_shared_notes(sidecar_path, payload.get('notes', []), active_editors, editors)
            doc['note_sync_mtime'] = os.path.getmtime(sidecar_path)
            doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
        except OSError:
            pass
        self.update_status()

    def poll_shared_notes(self):
        for doc in self.documents.values():
            if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
                continue

            sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
            current_signature = self.get_notes_sidecar_signature(sidecar_path)
            current_mtime = current_signature[0] / 1_000_000_000 if current_signature else None
            if current_signature != doc.get('note_sync_signature') or current_mtime != doc.get('note_sync_mtime'):
                previous_unread_count = self.get_unread_note_count(doc)
                self.restore_doc_notes(doc)
                current_unread_count = self.get_unread_note_count(doc)
                if current_unread_count > previous_unread_count:
                    self.play_unread_note_sound()
                doc['last_unread_count'] = current_unread_count
        self.update_status()

        self.root.after(self.note_sync_interval_ms, self.poll_shared_notes)

    def load_content_into_doc(self, doc, file_path):
        file_size = os.path.getsize(file_path)
        doc['file_size_bytes'] = file_size
        self.set_large_file_mode(doc, file_size >= self.large_file_threshold_bytes)
        doc['preview_mode'] = False
        doc['virtual_mode'] = file_size >= self.huge_file_preview_threshold_bytes

        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)

        if doc['virtual_mode']:
            text.insert(tk.END, "Indexing large file...\n")
            self.root.update_idletasks()
            doc['line_starts'], doc['file_size_bytes'] = self.build_line_index(file_path)
            doc['total_file_lines'] = max(1, len(doc['line_starts']))
            self.load_virtual_window(doc, 1)
        else:
            doc['line_starts'] = None
            doc['total_file_lines'] = 1
            doc['window_start_line'] = 1
            doc['window_end_line'] = 1
            text.configure(state='normal')
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                while True:
                    chunk = f.read(self.file_load_chunk_size)
                    if not chunk:
                        break
                    text.insert(tk.END, chunk)
                    if doc['large_file_mode']:
                        self.root.update_idletasks()

        text.edit_modified(False)
        text.mark_set(tk.INSERT, '1.0')
        text.tag_remove('sel', '1.0', tk.END)
        text.see('1.0')
        self.configure_syntax_highlighting(doc['frame'])
        self.restore_doc_notes(doc)
        self.register_doc_for_shared_notes(doc)
        self.refresh_tab_title(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        if str(doc['frame']) == self.notebook.select():
            self.update_status()

    def get_session_state(self):
        current_doc = self.get_current_doc()
        selected_file = current_doc['file_path'] if current_doc and current_doc['file_path'] else None
        open_files = []

        for tab_id in self.notebook.tabs():
            doc = self.documents.get(str(tab_id))
            if not doc:
                continue
            if doc['file_path'] and os.path.exists(doc['file_path']):
                open_files.append(doc['file_path'])

        open_files = list(dict.fromkeys(open_files))
        open_files = [path for path in open_files if path not in self.closed_session_files]
        if selected_file not in open_files:
            selected_file = None

        return {
            'open_files': open_files,
            'selected_file': selected_file,
            'recent_files': [
                path for path in self.recent_files
                if isinstance(path, str) and os.path.exists(path)
            ][:self.max_recent_files],
            'closed_session_files': [
                path for path in self.closed_session_files
                if isinstance(path, str)
            ],
            'sound_enabled': bool(self.sound_enabled.get()),
            'status_bar_enabled': bool(self.status_bar_enabled.get()),
            'numbered_lines_enabled': bool(self.numbered_lines_enabled.get()),
            'autocomplete_enabled': bool(self.autocomplete_enabled.get()),
            'syntax_theme': self.syntax_theme.get(),
        }

    def schedule_recovery_save(self):
        if self.isolated_session:
            return
        if self.recovery_job:
            try:
                self.root.after_cancel(self.recovery_job)
            except tk.TclError:
                pass
        self.recovery_job = self.root.after(1500, self.persist_recovery_state)

    def get_recovery_state(self):
        unsaved_tabs = []
        selected_untitled = None
        for doc in self.documents.values():
            if doc.get('file_path'):
                continue
            content = doc['text'].get('1.0', 'end-1c')
            if not content.strip() and not doc['text'].edit_modified():
                continue
            unsaved_tabs.append({
                'untitled_name': doc.get('untitled_name') or self.next_untitled_name(),
                'content': content,
                'modified': bool(doc['text'].edit_modified())
            })
            if self.get_current_doc() == doc:
                selected_untitled = doc.get('untitled_name')
        return {
            'unsaved_tabs': unsaved_tabs,
            'selected_untitled': selected_untitled,
            'timestamp': datetime.now().isoformat(timespec='seconds')
        }

    def persist_recovery_state(self):
        self.recovery_job = None
        if self.isolated_session:
            return
        recovery = self.get_recovery_state()
        if not recovery['unsaved_tabs']:
            if os.path.exists(self.recovery_path):
                try:
                    os.remove(self.recovery_path)
                except OSError as exc:
                    self.log_exception("remove recovery state", exc)
            return
        for attempt in range(2):
            try:
                recovery_dir = os.path.dirname(self.recovery_path)
                os.makedirs(recovery_dir, exist_ok=True)
                fd, temp_path = tempfile.mkstemp(prefix='notepadx-recovery-', suffix='.tmp', dir=recovery_dir)
                try:
                    with os.fdopen(fd, 'w', encoding='utf-8') as f:
                        json.dump(recovery, f, indent=2)
                    os.replace(temp_path, self.recovery_path)
                finally:
                    if os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except OSError:
                            pass
                self.hide_support_file(self.recovery_path)
                return
            except PermissionError as exc:
                if attempt == 0:
                    self.move_support_paths_to_user_dir()
                    continue
                self.log_exception("persist recovery state", exc)
                return
            except Exception as exc:
                self.log_exception("persist recovery state", exc)
                return

    def restore_recovery_state(self):
        if self.isolated_session or not os.path.exists(self.recovery_path):
            return
        try:
            with open(self.recovery_path, 'r', encoding='utf-8') as f:
                recovery = json.load(f)
        except Exception as exc:
            self.log_exception("restore recovery state", exc)
            return
        unsaved_tabs = recovery.get('unsaved_tabs', [])
        if not isinstance(unsaved_tabs, list) or not unsaved_tabs:
            return
        if not messagebox.askyesno("Recover Tabs", "Notepad-X found unsaved tabs from a previous crash. Restore them?", parent=self.root):
            return
        current_doc = self.get_current_doc()
        if current_doc and not current_doc.get('file_path') and not current_doc['text'].edit_modified() and not current_doc['text'].get('1.0', 'end-1c').strip():
            self.notebook.forget(current_doc['frame'])
            self.documents.pop(str(current_doc['frame']), None)
        selected_name = recovery.get('selected_untitled')
        selected_tab = None
        for recovered_tab in unsaved_tabs:
            if not isinstance(recovered_tab, dict):
                continue
            content = str(recovered_tab.get('content', ''))
            tab_id = self.create_tab(content=content, select=False)
            doc = self.documents[str(tab_id)]
            doc['untitled_name'] = str(recovered_tab.get('untitled_name') or doc['untitled_name'])
            doc['text'].edit_modified(bool(recovered_tab.get('modified', True)))
            self.refresh_tab_title(doc['frame'])
            if doc['untitled_name'] == selected_name:
                selected_tab = doc['frame']
        if selected_tab is None and self.documents:
            selected_tab = next(iter(self.documents.values()))['frame']
        if selected_tab is not None:
            self.notebook.select(selected_tab)
            self.set_active_document(selected_tab)

    def save_session(self):
        if self.isolated_session:
            return
        session = self.get_session_state()
        has_recent_files = bool(session.get('recent_files'))
        if not session['open_files'] and not has_recent_files:
            if os.path.exists(self.session_path):
                try:
                    os.remove(self.session_path)
                except OSError:
                    pass
            return

        for attempt in range(2):
            try:
                session_dir = os.path.dirname(self.session_path)
                os.makedirs(session_dir, exist_ok=True)
                fd, temp_path = tempfile.mkstemp(prefix='notepadx-session-', suffix='.tmp', dir=session_dir)
                try:
                    with os.fdopen(fd, 'w', encoding='utf-8') as f:
                        json.dump(session, f, indent=2)
                    os.replace(temp_path, self.session_path)
                finally:
                    if os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except OSError:
                            pass
                self.hide_support_file(self.session_path)
                return
            except PermissionError as exc:
                if attempt == 0:
                    self.move_support_paths_to_user_dir()
                    continue
                self.log_exception("save session", exc)
                return
            except Exception as exc:
                self.log_exception("save session", exc)
                return

    def restore_session(self):
        if self.isolated_session:
            return
        if not os.path.exists(self.session_path):
            return

        try:
            with open(self.session_path, 'r', encoding='utf-8') as f:
                session = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            self.log_exception("restore session", exc)
            return

        open_files = [
            path for path in session.get('open_files', [])
            if isinstance(path, str) and os.path.exists(path)
        ]
        self.closed_session_files = {
            path for path in session.get('closed_session_files', [])
            if isinstance(path, str)
        }
        open_files = [path for path in open_files if path not in self.closed_session_files]
        self.recent_files = [
            path for path in session.get('recent_files', [])
            if isinstance(path, str) and os.path.exists(path)
        ][:self.max_recent_files]
        self.sound_enabled.set(bool(session.get('sound_enabled', True)))
        self.status_bar_enabled.set(bool(session.get('status_bar_enabled', True)))
        self.numbered_lines_enabled.set(bool(session.get('numbered_lines_enabled', True)))
        self.autocomplete_enabled.set(bool(session.get('autocomplete_enabled', True)))
        self.syntax_theme.set(str(session.get('syntax_theme', 'Default')))
        if self.status_bar_enabled.get():
            self.status_frame.grid()
        else:
            self.status_frame.grid_remove()
        self.toggle_numbered_lines()
        self.refresh_recent_files_menu()
        if not open_files:
            return

        current_doc = self.get_current_doc()
        if current_doc and not current_doc['file_path'] and not current_doc['text'].edit_modified():
            self.notebook.forget(current_doc['frame'])
            self.documents.pop(str(current_doc['frame']), None)

        restored_tabs = {}
        for file_path in open_files:
            tab_id = self.create_tab(file_path=file_path, select=False)
            doc = self.documents[str(tab_id)]
            self.load_content_into_doc(doc, file_path)
            restored_tabs[file_path] = doc['frame']

        selected_file = session.get('selected_file')
        selected_tab = restored_tabs.get(selected_file)
        if selected_tab is None and restored_tabs:
            selected_tab = next(iter(restored_tabs.values()))

        if selected_tab is not None:
            self.notebook.select(selected_tab)
            self.set_active_document(selected_tab)

    def add_recent_file(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return
        self.closed_session_files.discard(file_path)
        self.recent_files = [path for path in self.recent_files if path != file_path]
        self.recent_files.insert(0, file_path)
        self.recent_files = self.recent_files[:self.max_recent_files]
        self.refresh_recent_files_menu()

    def clear_recent_files(self):
        self.recent_files = []
        self.refresh_recent_files_menu()
        self.save_session()

    def refresh_recent_files_menu(self):
        if not hasattr(self, 'recent_menu'):
            return

        self.recent_menu.delete(0, tk.END)
        valid_recent_files = [path for path in self.recent_files if os.path.exists(path)]
        self.recent_files = valid_recent_files[:self.max_recent_files]

        if not self.recent_files:
            self.recent_menu.add_command(label="(Empty)", state='disabled')
            return

        for file_path in self.recent_files:
            self.recent_menu.add_command(
                label=os.path.basename(file_path),
                command=lambda path=file_path: self.open_recent_file(path)
            )

        self.recent_menu.add_separator()
        self.recent_menu.add_command(label="Clear list", command=self.clear_recent_files)

    def next_untitled_name(self):
        used_numbers = set()
        for doc in self.documents.values():
            if doc['file_path']:
                continue
            name = doc.get('untitled_name', '')
            if name.startswith("Untitled "):
                try:
                    used_numbers.add(int(name.split(" ", 1)[1]))
                except (ValueError, IndexError):
                    pass

        next_number = 1
        while next_number in used_numbers:
            next_number += 1
        return f"Untitled {next_number}"

    def get_current_doc(self):
        current_tab = self.notebook.select()
        return self.documents.get(current_tab)

    def set_active_document(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            self.text = None
            self.current_file = None
            self.syntax_mode_selection.set('auto')
            self.root.title("Notepad-X")
            return
        self.text = doc['text']
        self.current_file = doc['file_path']
        self.syntax_mode_selection.set(doc.get('syntax_override') or 'auto')
        self.update_line_number_gutter(doc)
        self.update_window_title()
        self.update_status()

    def on_tab_changed(self, event=None):
        self.hide_autocomplete_popup()
        for doc in self.documents.values():
            self.hide_note_popup(doc)
        self.set_active_document(self.notebook.select())
        if self.text:
            self.set_last_active_editor_widget(self.text)
            self.text.focus_set()

    def get_doc_name(self, tab_id):
        doc = self.documents[str(tab_id)]
        if doc['file_path']:
            return os.path.basename(doc['file_path'])
        return doc['untitled_name']

    def get_doc_title(self, tab_id):
        doc = self.documents[str(tab_id)]
        title = self.get_doc_name(tab_id)
        if doc['text'].edit_modified():
            title += " *"
        return title

    def refresh_tab_title(self, tab_id):
        self.notebook.tab(tab_id, text=self.get_doc_title(tab_id))
        if str(tab_id) == self.notebook.select():
            self.update_window_title()
        if self.compare_active and self.compare_source_tab == str(tab_id):
            self.refresh_compare_header()
        self.save_session()

    def update_window_title(self):
        doc = self.get_current_doc()
        if not doc:
            self.root.title("Notepad-X")
            return
        title = self.get_doc_name(doc['frame'])
        if doc['text'].edit_modified():
            title += " *"
        self.root.title(f"Notepad-X - {title}")

    def on_text_modified(self, tab_id):
        if str(tab_id) not in self.documents:
            return
        doc = self.documents[str(tab_id)]
        if doc.get('suspend_modified_events'):
            doc['text'].edit_modified(False)
            return
        if not doc.get('file_path'):
            self.configure_syntax_highlighting(tab_id)
        if doc.get('syntax_mode') and doc.get('syntax_mode') != 'python':
            self.schedule_syntax_highlight(doc)
        self.update_line_number_gutter(doc)
        self.refresh_tab_title(tab_id)
        self.schedule_recovery_save()
        if self.compare_active and self.compare_source_tab == str(tab_id):
            self.schedule_compare_refresh()
        if str(tab_id) == self.notebook.select():
            self.update_status()

    def new_tab(self, event=None):
        self.create_tab()
        self.save_session()
        self.schedule_recovery_save()
        return "break"

    def on_tab_drag_start(self, event):
        try:
            self.dragging_tab_index = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            self.dragging_tab_index = None

    def on_tab_drag_motion(self, event):
        if getattr(self, 'dragging_tab_index', None) is None:
            return
        try:
            target_index = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return
        if target_index != self.dragging_tab_index:
            self.notebook.insert(target_index, self.notebook.tabs()[self.dragging_tab_index])
            self.dragging_tab_index = target_index

    def on_tab_drag_end(self, event):
        self.dragging_tab_index = None

    def switch_tab_by_offset(self, offset):
        tab_ids = self.notebook.tabs()
        if len(tab_ids) < 2:
            return "break"

        current_tab = self.notebook.select()
        if current_tab not in tab_ids:
            return "break"

        next_index = (tab_ids.index(current_tab) + offset) % len(tab_ids)
        next_tab = tab_ids[next_index]
        self.notebook.select(next_tab)
        self.set_active_document(next_tab)
        if self.text:
            self.text.focus_set()
        return "break"

    def switch_tab_left(self, event=None):
        return self.switch_tab_by_offset(-1)

    def switch_tab_right(self, event=None):
        return self.switch_tab_by_offset(1)

    def close_current_tab(self, event=None, recreate_if_empty=True):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        if (
            len(self.documents) == 1 and
            not doc.get('file_path') and
            doc.get('untitled_name') == 'Untitled 1'
        ):
            return "break"
        if not self.confirm_close_tab(doc):
            return "break"

        closed_file_path = doc.get('file_path')
        if closed_file_path:
            self.closed_session_files.add(closed_file_path)

        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.close_compare_panel()

        self.unregister_doc_from_shared_notes(doc)
        tab_id = str(doc['frame'])
        self.notebook.forget(doc['frame'])
        self.documents.pop(tab_id, None)

        if not self.documents:
            if recreate_if_empty:
                new_tab = self.create_tab()
                self.set_active_document(new_tab)
                self.current_file = None
            else:
                self.save_session()
                self.root.quit()
                return "break"
        else:
            self.set_active_document(self.notebook.select())
        if not any(existing_doc.get('file_path') for existing_doc in self.documents.values()):
            self.current_file = None
        self.save_session()
        return "break"

    # ─── Menu ────────────────────────────────────────────────────
    def create_menu(self):
        self.menu = tk.Menu(self.root, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a', activeforeground='white')
        self.root.config(menu=self.menu)

        file_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open", command=self.open_file, accelerator="Ctrl+W")
        file_menu.add_command(label="Open Project", command=self.open_project, accelerator="Ctrl+Shift+W")
        self.recent_menu = tk.Menu(file_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                                   activebackground='#3a3a3a')
        file_menu.add_cascade(label="Recent", menu=self.recent_menu)
        self.refresh_recent_files_menu()
        file_menu.add_command(label="New Tab", command=self.new_tab, accelerator="Ctrl+T")
        file_menu.add_command(label="Close Tab", command=self.close_current_tab, accelerator="Ctrl+Shift+T")
        file_menu.add_command(label="Save", command=self.save, accelerator="Ctrl+S")
        file_menu.add_command(label="Save all", command=self.save_all, accelerator="Ctrl+Shift+S")
        file_menu.add_command(label="Save Copy As", command=self.save_copy_as, accelerator="Ctrl+Shift+Q")
        file_menu.add_command(label="Print", command=self.print_file, accelerator="Ctrl+P")
        file_menu.add_command(label="Export Notes", command=self.export_notes_report, accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app, accelerator="Ctrl+Shift+X")

        edit_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self.undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=self.redo, accelerator="Ctrl+Shift+Z")
        edit_menu.add_separator()
        edit_menu.add_command(label="Cut",  command=lambda: self.text.event_generate("<<Cut>>"),  accelerator="Ctrl+X")
        edit_menu.add_command(label="Copy", command=self.copy, accelerator="Ctrl+C")
        edit_menu.add_command(label="Paste", command=self.paste, accelerator="Ctrl+V")
        edit_menu.add_command(label="Select All", command=self.select_all, accelerator="Ctrl+A")
        edit_menu.add_separator()
        edit_menu.add_command(label="Find", command=self.show_find_panel, accelerator="Ctrl+F")
        edit_menu.add_command(label="Find Next", command=self.find_next, accelerator="F3")
        edit_menu.add_command(label="Cycle Notes", command=self.goto_next_note, accelerator="F4")
        edit_menu.add_command(label="Replace", command=self.show_replace_panel, accelerator="Ctrl+R")
        note_filter_menu = tk.Menu(edit_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        edit_menu.add_cascade(label="Filter Notes", menu=note_filter_menu)
        note_filter_menu.add_radiobutton(label="All", variable=self.note_filter, value='all')
        note_filter_menu.add_radiobutton(label="Unread", variable=self.note_filter, value='unread')
        note_filter_menu.add_radiobutton(label="Allowed", variable=self.note_filter, value='allowed')
        note_filter_menu.add_radiobutton(label="Denied", variable=self.note_filter, value='denied')
        edit_menu.add_separator()
        edit_menu.add_command(label="Go To Line", command=self.goto_line_dialog, accelerator="Ctrl+G")
        edit_menu.add_command(label="Date", command=self.insert_date, accelerator="Ctrl+D")
        edit_menu.add_command(label="Time/Date", command=self.insert_time_date, accelerator="Ctrl+Shift+D")
        edit_menu.add_command(label="Font", command=self.show_font_dialog, accelerator="Ctrl+Shift+F")
        view_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Full Screen", command=self.toggle_fullscreen, accelerator="F11")
        view_menu.add_command(label="Switch Tab", command=self.switch_tab_right, accelerator="Ctrl+Tab")
        view_menu.add_command(label="Switch Tab Left", command=self.switch_tab_left, accelerator="Ctrl+Left")
        view_menu.add_command(label="Switch Tab Right", command=self.switch_tab_right, accelerator="Ctrl+Right")
        view_menu.add_checkbutton(label="Status Bar", variable=self.status_bar_enabled, command=self.toggle_status_bar, accelerator="Ctrl+B")
        view_menu.add_checkbutton(label="Numbered Lines", variable=self.numbered_lines_enabled, command=self.toggle_numbered_lines)
        view_menu.add_checkbutton(label="Autocomplete", variable=self.autocomplete_enabled, command=self.toggle_autocomplete)
        view_menu.add_checkbutton(label="Word Wrap", variable=self.word_wrap_enabled, command=self.toggle_word_wrap)
        view_menu.add_checkbutton(label="Sound", variable=self.sound_enabled, command=self.toggle_sound)
        syntax_theme_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label="Syntax Theme", menu=syntax_theme_menu)
        for theme_name in ('Default', 'Soft', 'Vivid'):
            syntax_theme_menu.add_radiobutton(label=theme_name, variable=self.syntax_theme, value=theme_name, command=lambda name=theme_name: self.set_syntax_theme(name))
        syntax_mode_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label="Syntax Mode", menu=syntax_mode_menu)
        for mode_label, mode_value in (
            ('Auto', 'auto'), ('Plain Text', 'plain'), ('Python', 'python'), ('C', 'c'),
            ('C++', 'cpp'), ('Rust', 'rust'), ('Java', 'java'), ('JavaScript', 'javascript'),
            ('HTML', 'html'), ('PHP', 'php'), ('XML', 'xml'), ('SQL', 'sql')
        ):
            syntax_mode_menu.add_radiobutton(
                label=mode_label,
                variable=self.syntax_mode_selection,
                value=mode_value,
                command=lambda value=mode_value: self.set_current_syntax_override(value)
            )
        view_menu.add_command(label="Compare Tabs", command=self.show_split_compare, accelerator="Ctrl+Q")
        view_menu.add_command(label="Close Compare Tabs", command=self.close_compare_panel, accelerator="Ctrl+Shift+X")

        help_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Help Contents", command=self.show_help_contents)
        help_menu.add_command(label="About Notepad-X", command=self.show_about_dialog)

    def show_help_contents(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Notepad-X Help")
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color)
        dialog.geometry("900x650")

        if os.path.exists(self.icon_path):
            try:
                dialog.iconbitmap(self.icon_path)
            except tk.TclError:
                pass

        container = tk.Frame(dialog, bg=self.bg_color)
        container.pack(fill='both', expand=True, padx=12, pady=12)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        help_text = tk.Text(
            container,
            wrap='word',
            bg=self.text_bg,
            fg=self.text_fg,
            insertbackground=self.cursor_color,
            selectbackground=self.select_bg,
            selectforeground='white',
            font=('Consolas', 11),
            padx=12,
            pady=12,
            borderwidth=0,
            highlightthickness=0,
            relief='flat'
        )
        help_text.grid(row=0, column=0, sticky='nsew')

        scroll = ttk.Scrollbar(container, orient='vertical', command=help_text.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        help_text.configure(yscrollcommand=scroll.set)

        if os.path.exists(self.help_path):
            try:
                with open(self.help_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except OSError:
                content = "Unable to open the Notepad-X help file."
        else:
            content = "Notepad-X help file not found."

        help_text.insert('1.0', content)
        help_text.configure(state='disabled')

        close_button = tk.Button(
            dialog,
            text="Close",
            command=dialog.destroy,
            bg='#2d2d2d',
            fg=self.fg_color,
            activebackground='#3a3a3a',
            activeforeground='white',
            relief='flat',
            borderwidth=0,
            padx=18,
            pady=6
        )
        close_button.pack(pady=(0, 12))

        dialog.bind('<Escape>', lambda e: dialog.destroy())
        self.center_window(dialog)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        dialog.focus_force()
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current))
        dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)

    def show_split_compare(self):
        current_doc = self.get_current_doc()
        if not current_doc or len(self.documents) < 2:
            messagebox.showinfo("Compare With Tab", "Open at least two tabs to compare.", parent=self.root)
            return "break"
        choices = []
        for tab_id, doc in self.documents.items():
            if doc is current_doc:
                continue
            choices.append((self.get_doc_name(doc['frame']), doc))
        dialog = tk.Toplevel(self.root)
        dialog.title("Compare With Tab")
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color, padx=12, pady=12)
        tk.Label(dialog, text="Choose a tab to compare with the current one:", bg=self.bg_color, fg=self.fg_color).pack(anchor='w', pady=(0, 8))
        listbox = tk.Listbox(dialog, width=50, height=min(10, len(choices)))
        for label, _ in choices:
            listbox.insert(tk.END, label)
        listbox.pack(fill='both', expand=True)

        def open_compare(event=None):
            selection = listbox.curselection()
            if not selection:
                return
            other_doc = choices[selection[0]][1]
            dialog.destroy()
            self.start_inline_compare(other_doc)

        tk.Button(dialog, text="Compare", command=open_compare).pack(pady=(10, 0))
        listbox.bind('<Double-Button-1>', open_compare)
        listbox.focus_set()
        self.center_window(dialog)

    def refresh_compare_header(self):
        if not self.compare_active:
            return
        doc = self.documents.get(self.compare_source_tab)
        if not doc:
            self.close_compare_panel()
            return
        self.compare_title.config(text=f"Comparing with: {self.get_doc_title(doc['frame'])}")

    def schedule_compare_refresh(self):
        if self.compare_refresh_job:
            try:
                self.root.after_cancel(self.compare_refresh_job)
            except tk.TclError:
                pass
        self.compare_refresh_job = self.root.after(120, self.refresh_compare_panel)

    def refresh_compare_panel(self):
        self.compare_refresh_job = None
        if not self.compare_active or not self.compare_view:
            return

        doc = self.documents.get(self.compare_source_tab)
        if not doc:
            self.close_compare_panel()
            return

        compare_doc = self.compare_view
        compare_doc['file_path'] = doc.get('file_path')
        compare_doc['syntax_override'] = doc.get('syntax_override')
        compare_doc['large_file_mode'] = bool(doc.get('large_file_mode'))
        compare_doc['preview_mode'] = bool(doc.get('preview_mode'))
        compare_doc['virtual_mode'] = bool(doc.get('virtual_mode'))
        compare_doc['window_start_line'] = doc.get('window_start_line', 1)
        compare_doc['window_end_line'] = doc.get('window_end_line', 1)
        compare_doc['total_file_lines'] = doc.get('total_file_lines', 1)

        self.refresh_compare_header()

        compare_text = compare_doc['text']
        compare_text.delete('1.0', tk.END)
        compare_text.insert('1.0', doc['text'].get('1.0', 'end-1c'))
        self.configure_syntax_for_doc(compare_doc)
        self.update_line_number_gutter(compare_doc)

    def set_compare_sash_position(self):
        if not self.compare_active:
            return
        try:
            width = self.editor_paned.winfo_width()
            if width > 0:
                self.editor_paned.sash_place(0, max(240, width // 2), 0)
        except tk.TclError:
            pass

    def start_inline_compare(self, source_doc):
        if not source_doc:
            return "break"
        self.compare_source_tab = str(source_doc['frame'])
        if str(self.compare_container) not in self.editor_paned.panes():
            self.editor_paned.add(self.compare_container, stretch='always')
        self.compare_active = True
        self.refresh_compare_panel()
        self.root.after_idle(self.set_compare_sash_position)
        return "break"

    def close_compare_panel(self, event=None):
        if self.compare_refresh_job:
            try:
                self.root.after_cancel(self.compare_refresh_job)
            except tk.TclError:
                pass
            self.compare_refresh_job = None

        if self.compare_view:
            if self.compare_view.get('colorizer') is not None and self.compare_view.get('percolator') is not None:
                try:
                    self.compare_view['percolator'].removefilter(self.compare_view['colorizer'])
                except Exception:
                    pass
            self.compare_view['percolator'] = None
            self.compare_view['colorizer'] = None
            self.compare_view['syntax_job'] = None
            self.compare_view['syntax_mode'] = None
            self.compare_view['file_path'] = None
            self.compare_view['syntax_override'] = None
            self.compare_view['large_file_mode'] = False
            self.compare_view['preview_mode'] = False
            self.compare_view['virtual_mode'] = False
            self.compare_view['text'].delete('1.0', tk.END)

        self.compare_active = False
        self.compare_source_tab = None
        self.compare_title.config(text="")
        try:
            self.editor_paned.forget(self.compare_container)
        except tk.TclError:
            pass
        if self.text and self.text.winfo_exists():
            self.text.focus_set()
        return "break"

    def show_about_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("About Notepad-X")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=24, pady=20)
        dialog.pong_after_id = None

        if os.path.exists(self.icon_path):
            try:
                dialog.iconbitmap(self.icon_path)
            except tk.TclError:
                pass

        content = tk.Frame(dialog, bg=self.bg_color)
        content.pack()

        image_loaded = False
        icon_widget = None
        if os.path.exists(self.splash_path):
            try:
                from PIL import Image, ImageTk
                splash_image = Image.open(self.splash_path)
                splash_image.thumbnail((self.splash_max_width, self.splash_max_height), Image.LANCZOS)
                dialog.icon_image = ImageTk.PhotoImage(splash_image)
                image_loaded = True
                icon_widget = tk.Label(content, image=dialog.icon_image, bg=self.bg_color, cursor='hand2')
                icon_widget.pack(pady=(0, 12))
            except Exception:
                image_loaded = False

        if not image_loaded and os.path.exists(self.icon_path):
            try:
                from PIL import Image, ImageTk
                icon_image = Image.open(self.icon_path)
                dialog.icon_image = ImageTk.PhotoImage(icon_image)
                image_loaded = True
                icon_widget = tk.Label(content, image=dialog.icon_image, bg=self.bg_color, cursor='hand2')
                icon_widget.pack(pady=(0, 12))
            except Exception:
                image_loaded = False

        if not image_loaded:
            icon_widget = tk.Label(
                content,
                text="[Icon]",
                bg=self.bg_color,
                fg=self.fg_color,
                font=('Segoe UI', 12),
                cursor='hand2'
            )
            icon_widget.pack(pady=(0, 12))

        icon_widget.bind('<Button-1>', lambda e, d=dialog: self.maybe_start_about_pong(d, e))

        tk.Label(
            content,
            text="Notepad-X",
            bg=self.bg_color,
            fg=self.fg_color,
            font=('Segoe UI', 16, 'bold')
        ).pack(pady=(0, 16))

        tk.Label(
            content,
            text="Built because Microsoft forgot what Notepad was supposed to be.",
            bg=self.bg_color,
            fg='#9aa0a6',
            font=('Segoe UI', 9)
        ).pack(pady=(0, 16))

        tk.Button(
            dialog,
            text="Close",
            command=dialog.destroy,
            bg='#2d2d2d',
            fg=self.fg_color,
            activebackground='#3a3a3a',
            activeforeground='white',
            relief='flat',
            borderwidth=0,
            padx=18,
            pady=6
        ).pack()

        dialog.bind('<Escape>', lambda e: dialog.destroy())
        self.center_window(dialog)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        dialog.focus_force()
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current))
        dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)

    def maybe_start_about_pong(self, dialog, event):
        if not (event.state & 0x4):
            return
        self.start_about_pong(dialog)

    def start_about_pong(self, dialog):
        for widget in dialog.winfo_children():
            widget.destroy()

        dialog.configure(padx=12, pady=12)
        dialog.geometry("520x400")
        dialog.minsize(520, 400)

        header = tk.Label(
            dialog,
            text="Pong-X",
            bg=self.bg_color,
            fg=self.fg_color,
            font=('Segoe UI', 14, 'bold')
        )
        header.pack(pady=(0, 8))

        score_label = tk.Label(
            dialog,
            text="User 0   Computer 0",
            bg=self.bg_color,
            fg='#9aa0a6',
            font=('Segoe UI', 10)
        )
        score_label.pack(pady=(0, 8))

        canvas = tk.Canvas(
            dialog,
            width=480,
            height=240,
            bg='#0b0f14',
            highlightthickness=1,
            highlightbackground='#2d2d2d'
        )
        canvas.pack()

        info_frame = tk.Frame(dialog, bg='#161b22', highlightthickness=1, highlightbackground='#2d2d2d')
        info_frame.pack(fill='x', pady=(8, 10))

        info_label = tk.Label(
            info_frame,
            text="Player 1 keys: W & S Player 2 keys: Up & Down, Press Up/Down once to start PVP. Press R to restart.",
            bg='#161b22',
            fg='#e6edf3',
            font=('Segoe UI', 10, 'bold'),
            anchor='w',
            justify='left',
            wraplength=456,
            padx=10,
            pady=8
        )
        info_label.pack(fill='x')

        tk.Button(
            dialog,
            text="Close",
            command=dialog.destroy,
            bg='#2d2d2d',
            fg=self.fg_color,
            activebackground='#3a3a3a',
            activeforeground='white',
            relief='flat',
            borderwidth=0,
            padx=18,
            pady=6
        ).pack()

        state = {
            'canvas': canvas,
            'score_label': score_label,
            'width': 480,
            'height': 240,
            'left_y': 100,
            'right_y': 100,
            'paddle_h': 52,
            'paddle_w': 10,
            'ball_x': 240,
            'ball_y': 120,
            'ball_size': 10,
            'ball_dx': 4,
            'ball_dy': 3,
            'left_score': 0,
            'right_score': 0,
            'left_move': 0,
            'right_move': 0,
            'mode': 'ai',
        }
        dialog.pong_state = state

        dialog.bind('<KeyPress-Up>', lambda e: self.handle_about_pong_right_input(dialog, -7))
        dialog.bind('<KeyPress-Down>', lambda e: self.handle_about_pong_right_input(dialog, 7))
        dialog.bind('<KeyRelease-Up>', lambda e: self.set_about_pong_move(dialog, 'right', 0))
        dialog.bind('<KeyRelease-Down>', lambda e: self.set_about_pong_move(dialog, 'right', 0))
        dialog.bind('<KeyPress-w>', lambda e: self.set_about_pong_move(dialog, 'left', -7))
        dialog.bind('<KeyPress-W>', lambda e: self.set_about_pong_move(dialog, 'left', -7))
        dialog.bind('<KeyRelease-w>', lambda e: self.set_about_pong_move(dialog, 'left', 0))
        dialog.bind('<KeyRelease-W>', lambda e: self.set_about_pong_move(dialog, 'left', 0))
        dialog.bind('<KeyPress-s>', lambda e: self.set_about_pong_move(dialog, 'left', 7))
        dialog.bind('<KeyPress-S>', lambda e: self.set_about_pong_move(dialog, 'left', 7))
        dialog.bind('<KeyRelease-s>', lambda e: self.set_about_pong_move(dialog, 'left', 0))
        dialog.bind('<KeyRelease-S>', lambda e: self.set_about_pong_move(dialog, 'left', 0))
        dialog.bind('<KeyPress-r>', lambda e: self.restart_about_pong_match(dialog))
        dialog.bind('<KeyPress-R>', lambda e: self.restart_about_pong_match(dialog))
        dialog.bind('<Destroy>', lambda e, d=dialog: self.stop_about_pong(d), add='+')
        dialog.focus_set()

        self.draw_about_pong(dialog)
        self.run_about_pong(dialog)

    def set_about_pong_move(self, dialog, side, amount):
        if hasattr(dialog, 'pong_state'):
            dialog.pong_state[f'{side}_move'] = amount

    def handle_about_pong_right_input(self, dialog, amount):
        state = getattr(dialog, 'pong_state', None)
        if not state:
            return
        if state['mode'] == 'ai':
            state['mode'] = 'human'
            state['left_score'] = 0
            state['right_score'] = 0
            state['left_y'] = 100
            state['right_y'] = 100
            state['left_move'] = 0
            self.reset_about_pong_ball(state, -1)
        self.set_about_pong_move(dialog, 'right', amount)

    def restart_about_pong_match(self, dialog):
        state = getattr(dialog, 'pong_state', None)
        if not state:
            return
        state['left_score'] = 0
        state['right_score'] = 0
        state['left_y'] = 100
        state['right_y'] = 100
        state['left_move'] = 0
        state['right_move'] = 0
        self.reset_about_pong_ball(state, -1)
        self.draw_about_pong(dialog)

    def stop_about_pong(self, dialog):
        after_id = getattr(dialog, 'pong_after_id', None)
        if after_id:
            try:
                dialog.after_cancel(after_id)
            except tk.TclError:
                pass
            dialog.pong_after_id = None

    def reset_about_pong_ball(self, state, direction):
        state['ball_x'] = state['width'] / 2
        state['ball_y'] = state['height'] / 2
        state['ball_dx'] = 4 * direction
        state['ball_dy'] = 3 if state['ball_dy'] >= 0 else -3

    def draw_about_pong(self, dialog):
        state = getattr(dialog, 'pong_state', None)
        if not state:
            return

        canvas = state['canvas']
        canvas.delete('all')
        canvas.create_line(
            state['width'] / 2, 0,
            state['width'] / 2, state['height'],
            fill='#2d2d2d',
            dash=(6, 6)
        )

        player_x = 18
        cpu_x = state['width'] - 18 - state['paddle_w']
        canvas.create_rectangle(
            player_x, state['left_y'],
            player_x + state['paddle_w'], state['left_y'] + state['paddle_h'],
            fill='#79c0ff', outline=''
        )
        canvas.create_rectangle(
            cpu_x, state['right_y'],
            cpu_x + state['paddle_w'], state['right_y'] + state['paddle_h'],
            fill='#ff7b72', outline=''
        )
        canvas.create_oval(
            state['ball_x'], state['ball_y'],
            state['ball_x'] + state['ball_size'], state['ball_y'] + state['ball_size'],
            fill='#d4d4d4', outline=''
        )
        state['score_label'].config(
            text=(
                f"{'User' if state['mode'] == 'ai' else 'Player 1'} {state['left_score']}   "
                f"{'Computer' if state['mode'] == 'ai' else 'Player 2'} {state['right_score']}"
            )
        )

    def run_about_pong(self, dialog):
        if not dialog.winfo_exists():
            return

        state = getattr(dialog, 'pong_state', None)
        if not state:
            return

        state['left_y'] += state['left_move']
        state['left_y'] = max(0, min(state['height'] - state['paddle_h'], state['left_y']))
        if state['mode'] == 'ai':
            cpu_center = state['right_y'] + state['paddle_h'] / 2
            ball_center = state['ball_y'] + state['ball_size'] / 2
            if cpu_center < ball_center - 4:
                state['right_y'] += 4
            elif cpu_center > ball_center + 4:
                state['right_y'] -= 4
        else:
            state['right_y'] += state['right_move']
        state['right_y'] = max(0, min(state['height'] - state['paddle_h'], state['right_y']))

        state['ball_x'] += state['ball_dx']
        state['ball_y'] += state['ball_dy']

        if state['ball_y'] <= 0 or state['ball_y'] + state['ball_size'] >= state['height']:
            state['ball_dy'] *= -1

        player_x = 18
        cpu_x = state['width'] - 18 - state['paddle_w']

        if (
            state['ball_dx'] < 0 and
            state['ball_x'] <= player_x + state['paddle_w'] and
            state['left_y'] <= state['ball_y'] + state['ball_size'] and
            state['ball_y'] <= state['left_y'] + state['paddle_h']
        ):
            state['ball_x'] = player_x + state['paddle_w']
            state['ball_dx'] *= -1

        if (
            state['ball_dx'] > 0 and
            state['ball_x'] + state['ball_size'] >= cpu_x and
            state['right_y'] <= state['ball_y'] + state['ball_size'] and
            state['ball_y'] <= state['right_y'] + state['paddle_h']
        ):
            state['ball_x'] = cpu_x - state['ball_size']
            state['ball_dx'] *= -1

        if state['ball_x'] < 0:
            state['right_score'] += 1
            self.reset_about_pong_ball(state, 1)
        elif state['ball_x'] > state['width']:
            state['left_score'] += 1
            self.reset_about_pong_ball(state, -1)

        self.draw_about_pong(dialog)
        dialog.pong_after_id = dialog.after(16, lambda d=dialog: self.run_about_pong(d))

    # ─── File Operations ─────────────────────────────────────────
    def get_project_source_files(self, file_path):
        source_extensions = {
            '.py', '.pyw', '.rs', '.c', '.h', '.cpp', '.cc', '.cxx',
            '.hpp', '.hh', '.hxx', '.java', '.html', '.htm', '.php',
            '.asm', '.s', '.js', '.ts', '.tsx', '.jsx', '.css', '.json',
            '.xml', '.toml', '.yaml', '.yml', '.ini', '.cfg', '.sh', '.bat'
        }
        excluded_project_files = {
            'notepad-x.session.json',
            'notepad-x.editor.json',
        }
        file_path = os.path.abspath(file_path)
        project_dir = os.path.dirname(file_path)
        selected_extension = os.path.splitext(file_path)[1].lower()
        related_files = [file_path]

        try:
            directory_entries = sorted(os.scandir(project_dir), key=lambda entry: entry.name.lower())
        except OSError:
            return related_files

        for entry in directory_entries:
            if not entry.is_file():
                continue
            candidate_path = os.path.abspath(entry.path)
            if candidate_path == file_path:
                continue
            candidate_name = entry.name.lower()
            if candidate_name in excluded_project_files or candidate_name.endswith('.notepadx.notes.json'):
                continue
            candidate_extension = os.path.splitext(entry.name)[1].lower()
            if candidate_extension in source_extensions:
                related_files.append(candidate_path)

        if selected_extension not in source_extensions:
            return [file_path]
        return related_files

    def open_file_path(self, file_path):
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("File Missing", "That file could not be found.", parent=self.root)
            self.recent_files = [path for path in self.recent_files if path != file_path]
            self.refresh_recent_files_menu()
            self.save_session()
            return False

        for tab_id, doc in self.documents.items():
            if doc['file_path'] == file_path:
                self.notebook.select(doc['frame'])
                self.set_active_document(tab_id)
                self.add_recent_file(file_path)
                self.save_session()
                return True

        current_doc = self.get_current_doc()
        if current_doc and not current_doc['file_path'] and not current_doc['text'].edit_modified():
            current_doc['file_path'] = file_path
            self.load_content_into_doc(current_doc, file_path)
            self.set_active_document(current_doc['frame'])
        else:
            tab_id = self.create_tab(file_path=file_path, select=True)
            new_doc = self.documents[str(tab_id)]
            self.load_content_into_doc(new_doc, file_path)
            self.set_active_document(tab_id)

        self.add_recent_file(file_path)
        self.save_session()
        return True

    def open_file(self, event=None):
        file_path = filedialog.askopenfilename(parent=self.root)
        if file_path:
            self.open_file_path(file_path)
        return "break"

    def open_project(self, event=None):
        file_path = filedialog.askopenfilename(parent=self.root)
        if not file_path:
            return "break"

        selected_project_path = os.path.normcase(os.path.abspath(file_path))
        project_files = self.get_project_source_files(file_path)
        for project_file in project_files:
            self.open_file_path(project_file)
        for tab_id, doc in self.documents.items():
            doc_file_path = doc.get('file_path')
            if doc_file_path and os.path.normcase(os.path.abspath(doc_file_path)) == selected_project_path:
                self.notebook.select(doc['frame'])
                self.set_active_document(tab_id)
                if self.text:
                    self.text.focus_set()
                break
        return "break"

    def open_recent_file(self, file_path):
        self.open_file_path(file_path)

    def write_file_atomically(self, file_path, content):
        directory = os.path.dirname(file_path) or '.'
        fd, temp_path = tempfile.mkstemp(prefix='notepadx-save-', suffix='.tmp', dir=directory)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as temp_file:
                temp_file.write(content)
            os.replace(temp_path, file_path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError as exc:
                    self.log_exception("cleanup temp save file", exc)

    def save(self, event=None):
        doc = self.get_current_doc()
        if not doc:
            return False
        if doc.get('preview_mode') or doc.get('virtual_mode'):
            messagebox.showinfo(
                "Large File Mode",
                "This tab is opened in buffered large-file mode. Editing and saving are disabled for the full file view.",
                parent=self.root
            )
            return False

        if doc['file_path']:
            self.write_file_atomically(doc['file_path'], doc['text'].get('1.0', tk.END).rstrip('\n'))
            doc['text'].edit_modified(False)
            self.add_recent_file(doc['file_path'])
            self.refresh_tab_title(doc['frame'])
            self.update_status()
            self.save_session()
            return True
        else:
            return self.save_as()

    def save_all(self, event=None):
        original_tab = self.notebook.select()
        any_saved = False

        for tab_id in list(self.notebook.tabs()):
            doc = self.documents.get(str(tab_id))
            if not doc or doc.get('preview_mode') or doc.get('virtual_mode'):
                continue

            self.notebook.select(tab_id)
            self.set_active_document(tab_id)

            if doc['file_path'] or doc['text'].edit_modified():
                if not self.save():
                    if original_tab in self.notebook.tabs():
                        self.notebook.select(original_tab)
                        self.set_active_document(original_tab)
                    return "break"
                any_saved = True

        if original_tab in self.notebook.tabs():
            self.notebook.select(original_tab)
            self.set_active_document(original_tab)

        if any_saved:
            self.update_status()
        return "break"

    def save_as(self):
        doc = self.get_current_doc()
        if not doc:
            return False
        if doc.get('preview_mode') or doc.get('virtual_mode'):
            messagebox.showinfo(
                "Large File Mode",
                "Save As is disabled for buffered large-file tabs.",
                parent=self.root
            )
            return False
        file_path = filedialog.asksaveasfilename(parent=self.root)
        if file_path:
            old_file_path = doc.get('file_path')
            if old_file_path and old_file_path != file_path:
                self.unregister_doc_from_shared_notes(doc)
            doc['file_path'] = file_path
            self.configure_syntax_highlighting(doc['frame'])
            self.set_active_document(doc['frame'])
            self.register_doc_for_shared_notes(doc)
            return self.save()
        return False

    def save_copy_as(self):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        suggested_name = self.get_doc_name(doc['frame'])
        output_path = filedialog.asksaveasfilename(parent=self.root, initialfile=suggested_name)
        if not output_path:
            return "break"
        try:
            if doc.get('virtual_mode') or doc.get('preview_mode'):
                with open(doc['file_path'], 'rb') as src, open(output_path, 'wb') as dst:
                    while True:
                        chunk = src.read(self.file_load_chunk_size)
                        if not chunk:
                            break
                        dst.write(chunk)
            else:
                self.write_file_atomically(output_path, doc['text'].get('1.0', tk.END).rstrip('\n'))
            messagebox.showinfo("Save Copy As", f"Copy saved to:\n{output_path}", parent=self.root)
        except OSError as exc:
            self.log_exception("save copy as", exc)
            messagebox.showerror("Save Copy As", str(exc), parent=self.root)
        return "break"

    def print_file(self, event=None):
        doc = self.get_current_doc()
        if not doc:
            return "break"

        if not doc['file_path']:
            saved = self.save_as()
            if not saved:
                return "break"
            doc = self.get_current_doc()

        try:
            os.startfile(doc['file_path'], 'print')
        except OSError as exc:
            messagebox.showerror("Print Failed", str(exc), parent=self.root)
        return "break"

    def exit_app(self, event=None):
        for doc in list(self.documents.values()):
            if not self.confirm_close_tab(doc):
                return "break"
        for doc in list(self.documents.values()):
            self.unregister_doc_from_shared_notes(doc)
        if self.recovery_job:
            try:
                self.root.after_cancel(self.recovery_job)
            except tk.TclError:
                pass
            self.recovery_job = None
        if os.path.exists(self.recovery_path):
            try:
                os.remove(self.recovery_path)
            except OSError as exc:
                self.log_exception("remove recovery file on exit", exc)
        self.persist_editor_identity()
        self.save_session()
        self.root.quit()
        return "break"

    def confirm_close_tab(self, doc):
        if not doc['text'].edit_modified():
            return True

        self.notebook.select(doc['frame'])
        self.set_active_document(doc['frame'])
        answer = messagebox.askyesnocancel(
            "Save",
            f"Save changes to {self.get_doc_name(doc['frame'])} before closing?",
            parent=self.root
        )
        if answer is True:
            return self.save()
        return answer is False

    def current_doc_is_large_readonly(self):
        doc = self.get_current_doc()
        return bool(doc and (doc.get('preview_mode') or doc.get('virtual_mode')))

    # ─── Undo / Clipboard wrappers ───────────────────────────────
    def undo(self, event=None):
        if self.current_doc_is_large_readonly():
            return "break"
        try:
            self.text.edit_undo()
        except tk.TclError:
            pass

    def redo(self, event=None):
        if self.current_doc_is_large_readonly():
            return "break"
        try:
            self.text.edit_redo()
        except tk.TclError:
            pass

    def select_all(self, event=None):
        if not self.text:
            return "break"
        self.text.tag_add('sel', '1.0', 'end-1c')
        self.text.mark_set(tk.INSERT, '1.0')
        self.text.see(tk.INSERT)
        self.update_status()
        return "break"

    def copy(self, event=None):
        target = None
        if event is not None and isinstance(getattr(event, 'widget', None), tk.Text):
            target = event.widget
        elif isinstance(self.root.focus_get(), tk.Text):
            target = self.root.focus_get()
        elif isinstance(self.text, tk.Text):
            target = self.text

        if target is None:
            return "break"

        try:
            selection = target.get('sel.first', 'sel.last')
        except tk.TclError:
            return "break"

        self.root.clipboard_clear()
        self.root.clipboard_append(selection)
        self.root.update_idletasks()
        return "break"

    def paste(self, event=None):
        target = None
        if event is not None and isinstance(getattr(event, 'widget', None), tk.Text):
            target = event.widget
        elif isinstance(self.root.focus_get(), tk.Text):
            target = self.root.focus_get()
        elif isinstance(self.text, tk.Text):
            target = self.text

        compare_widget = self.get_compare_text_widget()
        if target is None or (compare_widget is not None and target == compare_widget):
            return "break"

        doc = self.get_doc_for_text_widget(target)
        if doc and (doc.get('preview_mode') or doc.get('virtual_mode')):
            return "break"

        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            return "break"

        if clipboard_text is None:
            return "break"

        try:
            target.delete('sel.first', 'sel.last')
        except tk.TclError:
            pass

        target.insert(tk.INSERT, clipboard_text)
        try:
            target.edit_modified(True)
        except tk.TclError:
            pass
        target.see(tk.INSERT)
        self.set_last_active_editor_widget(target)
        if doc:
            self.on_text_modified(doc['frame'])
        else:
            self.update_status()
        return "break"

    def insert_date(self, event=None):
        if not self.text or self.current_doc_is_large_readonly():
            return "break"
        current_date = datetime.now().strftime("%m/%d/%Y").lstrip('0').replace('/0', '/')
        self.text.insert(tk.INSERT, current_date)
        self.text.edit_modified(True)
        self.update_status()
        return "break"

    def insert_time_date(self, event=None):
        if not self.text or self.current_doc_is_large_readonly():
            return "break"
        timestamp = datetime.now().strftime("%I:%M %p %m/%d/%Y").lstrip('0').replace('/0', '/')
        self.text.insert(tk.INSERT, timestamp)
        self.text.edit_modified(True)
        self.update_status()
        return "break"

    def show_font_dialog(self, event=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("Font")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=16, pady=16)

        tk.Label(dialog, text="Font:", bg=self.bg_color, fg=self.fg_color).grid(row=0, column=0, sticky='w', pady=(0, 8))
        tk.Label(dialog, text="Size:", bg=self.bg_color, fg=self.fg_color).grid(row=0, column=1, sticky='w', padx=(12, 0), pady=(0, 8))

        families = sorted(set(tkfont.families()))
        family_var = tk.StringVar(value=self.font_family)
        size_var = tk.StringVar(value=str(self.current_font_size))

        family_combo = ttk.Combobox(dialog, textvariable=family_var, values=families, state='readonly', width=28)
        family_combo.grid(row=1, column=0, sticky='ew')
        size_combo = ttk.Combobox(dialog, textvariable=size_var, values=[str(size) for size in range(6, 33)], state='readonly', width=8)
        size_combo.grid(row=1, column=1, sticky='ew', padx=(12, 0))

        preview = tk.Label(
            dialog,
            text="AaBbYyZz 0123456789",
            bg=self.text_bg,
            fg=self.text_fg,
            padx=12,
            pady=12,
            relief='flat'
        )
        preview.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(12, 12))

        def update_preview(*_):
            try:
                preview.configure(font=(family_var.get(), int(size_var.get())))
            except (ValueError, tk.TclError):
                pass

        family_combo.bind('<<ComboboxSelected>>', update_preview)
        size_combo.bind('<<ComboboxSelected>>', update_preview)
        update_preview()

        button_row = tk.Frame(dialog, bg=self.bg_color)
        button_row.grid(row=3, column=0, columnspan=2, sticky='e')

        def apply_font():
            try:
                new_size = int(size_var.get())
            except ValueError:
                messagebox.showwarning("Invalid Size", "Choose a valid font size.", parent=dialog)
                return

            self.font_family = family_var.get() or self.font_family
            self.current_font_size = max(self.min_font_size, min(self.max_font_size, new_size))
            self.update_font()
            self.update_status()
            dialog.destroy()

        tk.Button(
            button_row,
            text="OK",
            command=apply_font,
            bg='#2d2d2d',
            fg=self.fg_color,
            activebackground='#3a3a3a',
            activeforeground='white',
            relief='flat',
            borderwidth=0,
            padx=16,
            pady=6
        ).pack(side='left', padx=(0, 8))

        tk.Button(
            button_row,
            text="Cancel",
            command=dialog.destroy,
            bg='#2d2d2d',
            fg=self.fg_color,
            activebackground='#3a3a3a',
            activeforeground='white',
            relief='flat',
            borderwidth=0,
            padx=16,
            pady=6
        ).pack(side='left')

        family_combo.focus_set()
        dialog.bind('<Return>', lambda e: apply_font())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
        dialog.grab_set()
        return "break"

    # ─── Goto Line ───────────────────────────────────────────────
    def goto_line_dialog(self, event=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("Go To Line")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color)

        tk.Label(dialog, text="Line Number:", bg=self.bg_color, fg=self.fg_color)\
            .grid(row=0, column=0, padx=6, pady=8)

        entry = tk.Entry(dialog, width=15)
        entry.grid(row=0, column=1, padx=6, pady=8)
        entry.focus_set()

        def goto():
            try:
                line = int(entry.get())
                current_doc = self.get_current_doc()
                if current_doc and current_doc.get('virtual_mode'):
                    last_line = current_doc['total_file_lines']
                else:
                    last_line = int(self.text.index('end-1c').split('.')[0])
                if 1 <= line <= last_line:
                    if current_doc and current_doc.get('virtual_mode'):
                        self.load_virtual_window(current_doc, line)
                        pos = f"{line - current_doc['window_start_line'] + 1}.0"
                    else:
                        pos = f"{line}.0"
                    self.text.mark_set(tk.INSERT, pos)
                    self.text.see(pos)
                    self.text.tag_remove('sel', '1.0', tk.END)
                    self.text.tag_add('sel', pos, f"{pos} +1 line")
                    self.update_status()
                    dialog.destroy()
                else:
                    messagebox.showwarning("Invalid", "Line number out of range.", parent=dialog)
            except ValueError:
                messagebox.showwarning("Invalid", "Enter a valid number.", parent=dialog)

        tk.Button(dialog, text="Go", command=goto)\
            .grid(row=1, column=1, sticky='e', padx=6, pady=6)

        dialog.bind('<Return>', lambda e: goto())
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
        dialog.grab_set()

if __name__ == "__main__":
    isolated_mode = '--isolated' in {arg.lower() for arg in sys.argv[1:]}
    NotepadX(isolated_session=isolated_mode)
