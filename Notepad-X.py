import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk, simpledialog, colorchooser
import os
import sys
import ctypes
import json
import bisect
import re
import hashlib
import base64
import secrets
import subprocess
import tempfile
import traceback
import time
import shutil
import stat
import socket
import threading
import webbrowser
from datetime import datetime, timezone
from ctypes import wintypes
from pathlib import Path
from types import SimpleNamespace

try:
    import resource
except ImportError:
    resource = None

try:
    import winreg
except ImportError:
    winreg = None

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.exceptions import InvalidTag
except ImportError:
    AESGCM = None
    InvalidTag = Exception

try:
    from idlelib.colorizer import ColorDelegator
    from idlelib.percolator import Percolator
except ImportError:
    ColorDelegator = None
    Percolator = None

DEFAULT_LOCALE_STRINGS = {
    "app.name": "Notepad-X",
    "app.about_title": "About Notepad-X",
    "app.help_title": "Notepad-X Help",
    "app.compare_title": "Compare With Tab",
    "about.heading": "Notepad-X",
    "about.tagline": "Built because Microsoft forgot what Notepad was supposed to be.",
    "common.close": "Close",
    "common.compare": "Compare",
    "common.clear_list": "Clear list",
    "common.empty": "(Empty)",
    "common.ok": "OK",
    "context.cut": "Cut",
    "context.copy": "Copy",
    "context.paste": "Paste",
    "context.select_all": "Select All",
    "context.add_note": "Add note",
    "context.respond": "Respond",
    "context.remove_note": "Remove note",
    "menu.file": "File",
    "menu.file.open": "Open",
    "menu.file.open_project": "Open Project",
    "menu.file.recent": "Recent",
    "menu.file.new_tab": "New Tab",
    "menu.file.close_tab": "Close Tab",
    "menu.file.save": "Save",
    "menu.file.save_all": "Save All",
    "menu.file.save_as": "Save As",
    "menu.file.save_and_run": "Save and Run",
    "menu.file.save_as_encrypted": "Save As Encrypted",
    "menu.file.print": "Print",
    "menu.file.export_notes": "Export Notes",
    "menu.file.exit": "Exit",
    "menu.edit": "Edit",
    "menu.edit.undo": "Undo",
    "menu.edit.redo": "Redo",
    "menu.edit.cut": "Cut",
    "menu.edit.copy": "Copy",
    "menu.edit.paste": "Paste",
    "menu.edit.select_all": "Select All",
    "menu.edit.find": "Find",
    "menu.edit.find_next": "Find Next",
    "menu.edit.find_previous": "Find Previous",
    "menu.edit.replace": "Replace",
    "menu.edit.cycle_notes": "Cycle Notes",
    "menu.edit.filter_notes": "Filter Notes",
    "menu.edit.goto_line": "Go To Line",
    "menu.edit.top_of_document": "Top of Document",
    "menu.edit.bottom_of_document": "Bottom of Document",
    "menu.edit.sync_page_navigation": "Sync PgUp/PgDn in Compare",
    "menu.edit.date": "Date",
    "menu.edit.time_date": "Time/Date",
    "menu.edit.font": "Font",
    "menu.edit.language": "Language",
    "menu.view": "View",
    "menu.view.full_screen": "Full Screen",
    "menu.view.switch_tab": "Switch Tab",
    "menu.view.status_bar": "Status Bar",
    "menu.view.numbered_lines": "Numbered Lines",
    "menu.view.autocomplete": "Autocomplete",
    "menu.view.edit_with_notepadx": "Edit with Notepad-X",
    "menu.view.word_wrap": "Word Wrap",
    "menu.view.sound": "Sound",
    "menu.view.syntax_theme": "Syntax Theme",
    "menu.view.create_theme": "Create Theme",
    "menu.view.syntax_mode": "Syntax Mode",
    "menu.view.currently_editing": "Currently Editing",
    "menu.view.compare_tabs": "Compare Tabs",
    "menu.view.close_compare_tabs": "Close Compare Tabs",
    "menu.help": "Help",
    "menu.help.contents": "Help Contents",
    "menu.help.about": "About Notepad-X",
    "note.filter.all": "All",
    "note.filter.unread": "Unread",
    "note.filter.yellow": "Yellow",
    "note.filter.green": "Green",
    "note.filter.red": "Red",
    "note.filter.blue": "Light Blue",
    "note.popup.title": "Code note",
    "note.popup.author": "Author",
    "note.popup.author_id": "Author ID",
    "note.popup.created": "Created",
    "note.popup.responses": "Responses",
    "note.prompt.add_title": "Add Note",
    "note.prompt.author": "Author:",
    "note.prompt.note": "Note:",
    "note.prompt.select_text_first": "Select some text first.",
    "note.prompt.author_required": "Enter an author name first.",
    "note.prompt.respond_title": "Respond",
    "note.prompt.name": "Name:",
    "note.prompt.response": "Response:",
    "note.prompt.response_color": "Response Color",
    "note.prompt.name_required": "Enter a name first.",
    "panel.currently_editing.title": "Currently Editing",
    "panel.currently_editing.unsaved": "No active IDs found.",
    "panel.currently_editing.none": "No active IDs found.",
    "status.initial": "Ln 1 of 1, Col 1 | 0 characters | UTF-8 | Normal",
    "status.memory_initial": " | Memory used: 0MB",
    "status.synced": "| Notes Synced",
    "status.line": "Ln",
    "status.col": "Col",
    "status.of": "of",
    "status.characters": "characters",
    "status.bytes": "bytes",
    "status.encoding": "UTF-8",
    "status.mode.normal": "Normal",
    "status.mode.virtual": "Virtual",
    "status.mode.preview": "Preview",
    "status.main": "{line_label} {row} {of_label} {total_lines}, {col_label} {col} | {char_info} | {encoding} | {zoom_text}{mode_suffix}",
    "status.char_count": "{total_chars} {characters_label}",
    "status.selected_char_count": "{selected_count} {of_label} {total_chars} {characters_label}",
    "status.byte_count": "{total_bytes} {bytes_label}",
    "status.memory": " | Memory used: {memory_mb}MB",
    "status.editor_id": " | ID: {editor_id}",
    "status.unread_tail": " | {unread_count} unread (F3 to view) | ({active_editors} editing)",
    "compare.need_two_tabs": "Open at least two tabs to compare.",
    "compare.choose_prompt": "Choose a tab to compare with the current one:",
    "compare.header": "Comparing with: {title}",
    "help.open_failed": "Unable to open the Notepad-X help file.",
    "help.not_found": "Notepad-X help file not found.",
    "syntax.theme.default": "Default",
    "syntax.theme.soft": "Soft",
    "syntax.theme.vivid": "Vivid",
    "syntax.theme.base4tone": "Base4Tone",
    "syntax.theme.green_monochrome": "Green Monochrome",
    "syntax.theme.orange_monochrome": "Orange Monochrome",
    "theme.create.title": "Create Theme",
    "theme.create.name": "Theme Name:",
    "theme.create.save": "Save Theme",
    "theme.create.pick_color": "Choose...",
    "theme.create.name_required": "Enter a theme name first.",
    "theme.create.saved": "Theme saved.",
    "theme.create.invalid_color": "Choose all theme colors before saving.",
    "theme.create.exists_title": "Overwrite Theme",
    "theme.create.exists_message": "A theme file with that name already exists.\n\nOverwrite it?",
    "theme.field.text_bg": "Editor Background",
    "theme.field.text_fg": "Editor Text",
    "theme.field.cursor": "Cursor",
    "theme.field.selection": "Selection",
    "theme.field.gutter_bg": "Gutter Background",
    "theme.field.gutter_current_bg": "Current Line Gutter",
    "theme.field.gutter_fg": "Gutter Text",
    "theme.field.gutter_current_fg": "Current Gutter Text",
    "theme.field.gutter_divider": "Gutter Divider",
    "theme.field.keyword": "Keyword",
    "theme.field.type": "Type",
    "theme.field.string": "String",
    "theme.field.comment": "Comment",
    "theme.field.number": "Number",
    "theme.field.preprocessor": "Preprocessor",
    "theme.field.tag": "Tag",
    "syntax.mode.auto": "Auto",
    "syntax.mode.plain": "Plain Text",
    "syntax.mode.python": "Python",
    "syntax.mode.c": "C",
    "syntax.mode.cpp": "C++",
    "syntax.mode.rust": "Rust",
    "syntax.mode.java": "Java",
    "syntax.mode.javascript": "JavaScript",
    "syntax.mode.html": "HTML",
    "syntax.mode.php": "PHP",
    "syntax.mode.xml": "XML",
    "syntax.mode.sql": "SQL",
    "accel.open": "Ctrl+W",
    "accel.open_project": "Ctrl+Shift+W",
    "accel.new_tab": "Ctrl+T",
    "accel.close_tab": "Ctrl+Shift+T",
    "accel.save": "Ctrl+S",
    "accel.save_all": "Ctrl+Shift+S",
    "accel.save_as": "Ctrl+Shift+Q",
    "accel.save_and_run": "Ctrl+Shift+R",
    "accel.save_as_encrypted": "Ctrl+Shift+E",
    "accel.print": "Ctrl+P",
    "accel.export_notes": "Ctrl+E",
    "accel.exit": "Ctrl+Shift+X",
    "accel.undo": "Ctrl+Z",
    "accel.redo": "Ctrl+Shift+Z",
    "accel.cut": "Ctrl+X",
    "accel.copy": "Ctrl+C",
    "accel.paste": "Ctrl+V",
    "accel.select_all": "Ctrl+A",
    "accel.find": "Ctrl+F",
    "accel.find_next": "F3",
    "accel.find_previous": "Shift+F3",
    "accel.replace": "Ctrl+R",
    "accel.cycle_notes": "F4",
    "accel.goto_line": "Ctrl+G",
    "accel.top_of_document": "Ctrl+PgUp",
    "accel.bottom_of_document": "Ctrl+PgDn",
    "accel.date": "Ctrl+D",
    "accel.time_date": "Ctrl+Shift+D",
    "accel.font": "Ctrl+Shift+F",
    "accel.full_screen": "F11",
    "accel.switch_tab": "Ctrl+Tab",
    "accel.status_bar": "Ctrl+B",
    "accel.currently_editing": "Ctrl+Shift+C",
    "accel.compare_tabs": "Ctrl+Q",
    "accel.close_compare_tabs": "Ctrl+Shift+X",
    "run.title": "Save and Run",
    "run.unsupported": "Save and Run is not available for this file type yet.",
    "run.runtime_missing": "Notepad-X could not find a runtime for {language} on this system.",
    "run.open_browser_failed": "Notepad-X could not open this file in your browser."
}

RTL_LOCALE_CODES = {'ar', 'ar_sa', 'ar_ae', 'ar_eg', 'ar_ma'}

LOCALE_DISPLAY_NAMES = {
    'en_us': 'English (US)',
    'ar': 'العربية',
    'ar_sa': 'العربية (السعودية)',
    'ar_ae': 'العربية (الإمارات)',
    'ar_eg': 'العربية (مصر)',
    'ar_ma': 'العربية (المغرب)',
    'fr_ca': 'Français (Canada)',
    'hi_in': 'हिन्दी (भारत)',
    'ja_jp': '日本語 (日本)',
    'ru_ru': 'Русский (Россия)',
    'zh_cn': '简体中文 (中国)',
}

LANGUAGE_NATIVE_NAMES = {
    'en': 'English',
    'ar': 'العربية',
    'de': 'Deutsch',
    'es': 'Español',
    'fr': 'Français',
    'hi': 'हिन्दी',
    'it': 'Italiano',
    'ja': '日本語',
    'nl': 'Nederlands',
    'pt': 'Português',
    'ru': 'Русский',
    'uk': 'Українська',
    'zh': '简体中文',
}

REGION_DISPLAY_NAMES = {
    'us': 'US',
    'sa': 'السعودية',
    'ae': 'الإمارات',
    'eg': 'مصر',
    'ma': 'المغرب',
    'de': 'Deutschland',
    '419': 'Latinoamérica',
    'es': 'España',
    'ca': 'Canada',
    'in': 'भारत',
    'it': 'Italia',
    'jp': '日本',
    'nl': 'Nederland',
    'br': 'Brasil',
    'ru': 'Россия',
    'ua': 'Україна',
    'cn': '中国',
}


def get_notepadx_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def is_notepadx_support_file_path(file_path):
    try:
        candidate_name = os.path.basename(str(file_path)).lower()
    except Exception:
        return False
    return (
        candidate_name.endswith('.notepadx.notes.json')
        or candidate_name.endswith('.notepadx.editors.json')
        or candidate_name == 'notepad-x.recovery.json'
        or candidate_name == 'notepad-x.crash.log'
        or (candidate_name.startswith('notepad-x.') and candidate_name.endswith('.session.json'))
        or (candidate_name.startswith('notepad-x.') and candidate_name.endswith('.editor.json'))
    )


def get_notepadx_single_instance_port(app_dir):
    seed = (
        f"NotepadX::{os.path.normcase(os.path.abspath(app_dir))}::"
        f"{os.environ.get('USERNAME') or os.environ.get('USER') or 'user'}"
    )
    seed_hash = hashlib.sha256(seed.encode('utf-8')).hexdigest()
    return 43000 + (int(seed_hash[:8], 16) % 10000)


def send_files_to_running_notepadx(app_dir, startup_files):
    normalized_files = []
    for raw_path in startup_files or []:
        if not raw_path:
            continue
        candidate_path = os.path.abspath(raw_path)
        if not os.path.exists(candidate_path) or is_notepadx_support_file_path(candidate_path):
            continue
        normalized_files.append(candidate_path)

    if not normalized_files:
        return False

    payload = json.dumps({
        'command': 'open_files',
        'files': normalized_files,
    }).encode('utf-8')

    client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client.settimeout(0.6)
    try:
        client.connect(('127.0.0.1', get_notepadx_single_instance_port(app_dir)))
        client.sendall(payload)
        return True
    except OSError:
        return False
    finally:
        try:
            client.close()
        except OSError:
            pass

class NotepadX:
    def __init__(self, isolated_session=False, startup_files=None):
        self.root = tk.Tk()
        self.root.title("Notepad-X")
        self.isolated_session = isolated_session
        self.startup_files = list(startup_files or [])
        self.init_config()
        self.root.title(self.app_name)
        self.init_runtime()
        self.init_ui()
        self.root.mainloop()

    def init_config(self):
        self.is_windows = os.name == 'nt'
        self.is_linux = sys.platform.startswith('linux')
        self.app_version = "v0.9.7"
        self.resource_dir = self.get_resource_dir()
        self.app_dir = self.get_app_dir()
        self.machine_profile_slug = self.get_machine_profile_slug()
        self.repo_url = "https://github.com/OPCraniX/Notepad-X"
        self.icon_path = self.resolve_gfx_path("Notepad-X.ico")
        self.splash_path = self.resolve_gfx_path("splash.png")
        self.splash_max_width = 430
        self.splash_max_height = 645
        self.note_sound_path = self.resolve_audio_path("note.mp3")
        self.delete_note_sound_path = self.resolve_audio_path("delete_note.mp3")
        self.window_icon_image = None

        self.bg_color     = '#1e1e1e'
        self.fg_color     = '#d4d4d4'
        self.text_bg      = '#0d1117'
        self.text_fg      = '#c9d1d9'
        self.cursor_color = '#58a6ff'
        self.select_bg    = '#264f78'
        self.match_bg     = '#e3b505'    # search highlight
        self.note_colors  = {
            'yellow': '#f2cc60',
            'green': '#7ee787',
            'red': '#f85149',
            'blue': '#79c0ff',
        }
        self.bracket_match_tag = 'bracket_match'
        self.bracket_pairs = {'(': ')', '[': ']', '{': '}'}
        self.reverse_bracket_pairs = {value: key for key, value in self.bracket_pairs.items()}
        self.bracket_match_max_chars = 500000
        self.note_color_labels = {
            'yellow': 'Yellow',
            'green': 'Green',
            'red': 'Red',
            'blue': 'Light Blue',
        }
        self.note_color_order = ('yellow', 'green', 'red', 'blue')
        self.note_bg      = self.note_colors['yellow']
        self.panel_bg     = '#252525'

        self.current_file = None
        self.find_matches_tag = 'find_match'
        self.find_current_tag = 'find_current'
        self.documents = {}
        self.memory_used_mb = 0
        self.syntax_highlighting_available = ColorDelegator is not None and Percolator is not None
        self.large_file_threshold_bytes = 5 * 1024 * 1024
        self.max_editable_large_file_bytes = 64 * 1024 * 1024
        self.file_load_chunk_size = 256 * 1024
        self.huge_file_preview_threshold_bytes = 100 * 1024 * 1024
        self.huge_file_preview_bytes = 2 * 1024 * 1024
        self.virtual_file_window_lines = 5000
        self.virtual_file_margin_lines = 800
        config_dir = self.get_config_dir(self.app_dir)
        os.makedirs(config_dir, exist_ok=True)
        self.locale_dir = self.get_locale_dir(config_dir)
        os.makedirs(self.locale_dir, exist_ok=True)
        self.theme_dir = self.get_theme_dir(config_dir)
        os.makedirs(self.theme_dir, exist_ok=True)
        self.migrate_language_files(config_dir=config_dir, locale_dir=self.locale_dir)
        self.ensure_theme_files(self.theme_dir)
        self.theme_definitions = self.load_theme_definitions(self.theme_dir)
        self.locale_code = 'en_us'
        self.locale_path = self.get_locale_file_path(self.locale_code, locale_dir=self.locale_dir)
        self.locale_strings = self.load_locale_strings(self.locale_path)
        self.app_name = self.tr('app.name', 'Notepad-X')
        self.session_path = self.build_session_path(config_dir)
        self.editor_identity_path = self.build_editor_identity_path(config_dir)
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.crash_log_path = os.path.join(self.app_dir, "Notepad-X.crash.log")
        self.help_path = os.path.join(self.resource_dir, "Notepad-X-help.txt")
        self.note_color_labels = {
            'yellow': self.tr('note.filter.yellow', 'Yellow'),
            'green': self.tr('note.filter.green', 'Green'),
            'red': self.tr('note.filter.red', 'Red'),
            'blue': self.tr('note.filter.blue', 'Light Blue'),
        }
        self.max_recent_files = 10
        self.max_session_files = 100
        self.max_recovery_tabs = 20
        self.shared_editor_stale_seconds = 30
        self.max_note_text_length = 4000
        self.max_note_name_length = 120
        self.max_note_reply_length = 1000
        self.encryption_magic = b'NPXENC1'
        self.encryption_version = 1
        self.encryption_key_length = 32
        self.encryption_nonce_length = 12
        self.encryption_salt_length = 16
        self.encryption_scrypt_n = 1 << 15
        self.encryption_scrypt_r = 8
        self.encryption_scrypt_p = 1
        self.encryption_scrypt_maxmem = 128 * 1024 * 1024
        self.live_find_min_chars = 2
        self.live_find_max_matches_per_widget = 1000
        self.live_find_max_matches_typing = 150
        self.recent_files = []
        self.closed_session_files = set()
        self.note_sync_interval_ms = 100
        self.note_editor_heartbeat_interval_ms = 1500
        self.single_instance_host = '127.0.0.1'
        self.single_instance_port = get_notepadx_single_instance_port(self.app_dir)
        self.single_instance_server = None
        self.single_instance_listener_thread = None
        self.single_instance_running = False
        self.remote_open_files = []
        self.remote_open_lock = threading.Lock()
        self.background_file_results = []
        self.background_file_lock = threading.Lock()
        self.kernel32 = None
        self.psapi = None

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
        self.edit_with_shell_enabled = tk.BooleanVar(value=False)
        self.search_all_tabs = tk.BooleanVar(value=False)
        self.note_filter = tk.StringVar(value='all')
        self.syntax_theme = tk.StringVar(value='Default')
        self.syntax_mode_selection = tk.StringVar(value='auto')
        self.language_selection = tk.StringVar(value=self.locale_code)
        self.recovery_job = None
        self.find_change_job = None
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
        self.autocomplete_doc_id = None
        self.autocomplete_start_index = None
        self.autocomplete_prefix = ""
        self.autocomplete_suspended = 0
        self.active_context_menu = None
        self.context_menu_posted_at = 0.0
        self.hovered_editor_widget = None
        self.sync_page_navigation_enabled = tk.BooleanVar(value=False)
        self.find_panel_visible = False
        self.replace_panel_visible = False
        self.currently_editing_panel_visible = False
        self.fullscreen = False
        self.fullscreen_panel_restore = False

    def init_runtime(self):
        self.kernel32 = None
        self.psapi = None
        if self.is_windows:
            self.kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            self.psapi = ctypes.WinDLL('psapi', use_last_error=True)
        self.configure_memory_api()
        self.configure_sound_api()
        self.known_editor_ids = self.load_known_editor_ids()
        self.editor_id = self.generate_editor_id()
        self.editor_aliases = set(self.known_editor_ids)
        self.editor_aliases.add(self.editor_id)
        self.persist_editor_identity()

        self.setup_exception_handling()
        self.start_single_instance_server()

    def init_ui(self):
        self.apply_window_icon(self.root)
        self.root.geometry("1500x700")
        self.root.configure(bg='#1e1e1e')
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.create_text_widget()
        self.create_bottom_panels()
        self.create_menu()
        self.create_status_bar()
        self.restore_session()
        self.restore_recovery_state()
        self.open_startup_files(self.startup_files)

        self.bind_keys()
        self.update_font()
        self.update_clock()
        self.update_memory_usage()
        self.process_background_file_results()
        self.process_remote_open_requests()
        self.poll_shared_notes()
        self.center_window(self.root)

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
        if not self.kernel32 or not self.psapi:
            return
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
        self.winmm = None
        self.shell32 = None
        if not self.is_windows:
            return
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
        try:
            self.shell32 = ctypes.WinDLL('shell32', use_last_error=True)
            self.shell32.ShellExecuteW.restype = wintypes.HINSTANCE
            self.shell32.ShellExecuteW.argtypes = [
                wintypes.HWND,
                wintypes.LPCWSTR,
                wintypes.LPCWSTR,
                wintypes.LPCWSTR,
                wintypes.LPCWSTR,
                ctypes.c_int,
            ]
            self.shell32.SHChangeNotify.restype = None
            self.shell32.SHChangeNotify.argtypes = [
                wintypes.LONG,
                wintypes.UINT,
                ctypes.c_void_p,
                ctypes.c_void_p,
            ]
        except Exception:
            self.shell32 = None

    def get_resource_dir(self):
        if getattr(sys, 'frozen', False):
            return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
        return os.path.dirname(__file__)

    def apply_window_icon(self, window):
        if window is None:
            return
        if self.is_windows and os.path.exists(self.icon_path):
            try:
                window.iconbitmap(self.icon_path)
                return
            except tk.TclError:
                pass
        if os.path.exists(self.splash_path):
            try:
                self.window_icon_image = tk.PhotoImage(file=self.splash_path)
                window.iconphoto(True, self.window_icon_image)
                return
            except tk.TclError:
                pass

    def create_about_display_image(self):
        if os.path.exists(self.splash_path):
            try:
                from PIL import Image, ImageTk
                splash_image = Image.open(self.splash_path)
                splash_image.thumbnail((self.splash_max_width, self.splash_max_height), Image.LANCZOS)
                return ImageTk.PhotoImage(splash_image)
            except Exception:
                pass
            try:
                splash_image = tk.PhotoImage(file=self.splash_path)
                width = max(1, splash_image.width())
                height = max(1, splash_image.height())
                scale = max(
                    1,
                    (width + self.splash_max_width - 1) // self.splash_max_width,
                    (height + self.splash_max_height - 1) // self.splash_max_height,
                )
                if scale > 1:
                    splash_image = splash_image.subsample(scale, scale)
                return splash_image
            except tk.TclError:
                pass

        if os.path.exists(self.icon_path):
            try:
                from PIL import Image, ImageTk
                icon_image = Image.open(self.icon_path)
                return ImageTk.PhotoImage(icon_image)
            except Exception:
                pass

        return None

    def get_user_support_dir(self):
        base_dir = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        if self.is_linux:
            base_dir = os.environ.get('XDG_STATE_HOME') or os.path.join(os.path.expanduser('~'), '.local', 'state')
        return os.path.join(base_dir, 'Notepad-X')

    def get_emergency_support_dir(self):
        return os.path.join(tempfile.gettempdir(), 'Notepad-X')

    def get_config_dir(self, base_dir):
        return os.path.join(base_dir, 'cfg')

    def get_locale_dir(self, config_dir):
        return os.path.join(config_dir, 'language')

    def get_theme_dir(self, config_dir):
        return os.path.join(config_dir, 'themes')

    def migrate_language_files(self, config_dir=None, locale_dir=None):
        config_dir = config_dir or self.get_config_dir(self.app_dir)
        locale_dir = locale_dir or self.get_locale_dir(config_dir)
        try:
            os.makedirs(locale_dir, exist_ok=True)
        except OSError:
            return
        try:
            entries = os.listdir(config_dir)
        except OSError:
            return
        for entry in entries:
            if not entry.lower().endswith('.yml'):
                continue
            source_path = os.path.join(config_dir, entry)
            if not os.path.isfile(source_path):
                continue
            target_path = os.path.join(locale_dir, entry)
            if os.path.exists(target_path):
                continue
            try:
                shutil.move(source_path, target_path)
            except OSError:
                continue

    def get_builtin_theme_definitions(self):
        return {
            'Default': {
                'surface': {
                    'text_bg': self.text_bg,
                    'text_fg': self.text_fg,
                    'cursor': self.cursor_color,
                    'selection': self.select_bg,
                    'gutter_bg': '#0d1117',
                    'gutter_current_bg': '#161b22',
                    'gutter_fg': '#8b949e',
                    'gutter_current_fg': '#c9d1d9',
                    'gutter_divider': '#30363d',
                },
                'syntax': {
                    'keyword': '#ff7b72',
                    'type': '#79c0ff',
                    'string': '#a5d6ff',
                    'comment': '#6a9955',
                    'number': '#f2cc60',
                    'preprocessor': '#d2a8ff',
                    'tag': '#7ee787',
                },
            },
            'Soft': {
                'surface': {
                    'text_bg': '#11131a',
                    'text_fg': '#cad3f5',
                    'cursor': '#89b4fa',
                    'selection': '#313244',
                    'gutter_bg': '#11131a',
                    'gutter_current_bg': '#181b24',
                    'gutter_fg': '#7f849c',
                    'gutter_current_fg': '#cad3f5',
                    'gutter_divider': '#313244',
                },
                'syntax': {
                    'keyword': '#f38ba8',
                    'type': '#89b4fa',
                    'string': '#94e2d5',
                    'comment': '#a6adc8',
                    'number': '#f9e2af',
                    'preprocessor': '#cba6f7',
                    'tag': '#a6e3a1',
                },
            },
            'Vivid': {
                'surface': {
                    'text_bg': '#101217',
                    'text_fg': '#f8f9fa',
                    'cursor': '#4cc9f0',
                    'selection': '#3a0f2d',
                    'gutter_bg': '#101217',
                    'gutter_current_bg': '#1b1f2a',
                    'gutter_fg': '#9aa0a6',
                    'gutter_current_fg': '#f8f9fa',
                    'gutter_divider': '#343a40',
                },
                'syntax': {
                    'keyword': '#ff4d6d',
                    'type': '#4cc9f0',
                    'string': '#72efdd',
                    'comment': '#80ed99',
                    'number': '#ffd166',
                    'preprocessor': '#b388ff',
                    'tag': '#06d6a0',
                },
            },
            'Base4Tone': {
                'surface': {
                    'text_bg': '#231f20',
                    'text_fg': '#e6d6c4',
                    'cursor': '#f6c177',
                    'selection': '#3b2f32',
                    'gutter_bg': '#1d191a',
                    'gutter_current_bg': '#2a2325',
                    'gutter_fg': '#7c6f72',
                    'gutter_current_fg': '#eadfd2',
                    'gutter_divider': '#3a3234',
                },
                'syntax': {
                    'keyword': '#cf8a8a',
                    'type': '#7ab0c8',
                    'string': '#d6b28a',
                    'comment': '#8b7d78',
                    'number': '#f1c27d',
                    'preprocessor': '#b7a1d3',
                    'tag': '#8fc7b0',
                },
            },
            'Green Monochrome': {
                'surface': {
                    'text_bg': '#081108',
                    'text_fg': '#86f08a',
                    'cursor': '#b8ffb8',
                    'selection': '#173617',
                    'gutter_bg': '#050b05',
                    'gutter_current_bg': '#102510',
                    'gutter_fg': '#3d8f47',
                    'gutter_current_fg': '#b8ffb8',
                    'gutter_divider': '#24552b',
                },
                'syntax': {
                    'keyword': '#9af59f',
                    'type': '#74db7c',
                    'string': '#c8ffb2',
                    'comment': '#4a9252',
                    'number': '#b7ff91',
                    'preprocessor': '#8ae28f',
                    'tag': '#79c87f',
                },
            },
            'Orange Monochrome': {
                'surface': {
                    'text_bg': '#140c04',
                    'text_fg': '#ffb45a',
                    'cursor': '#ffd08a',
                    'selection': '#3a2108',
                    'gutter_bg': '#100802',
                    'gutter_current_bg': '#2a1806',
                    'gutter_fg': '#b56d26',
                    'gutter_current_fg': '#ffd08a',
                    'gutter_divider': '#5b3511',
                },
                'syntax': {
                    'keyword': '#ffc56d',
                    'type': '#ffaf4d',
                    'string': '#ffd89b',
                    'comment': '#b67a3a',
                    'number': '#ffe18c',
                    'preprocessor': '#ffbf74',
                    'tag': '#ffa33f',
                },
            },
        }

    def slugify_theme_name(self, theme_name):
        slug = re.sub(r'[^a-z0-9]+', '_', str(theme_name).strip().lower()).strip('_')
        return slug or 'theme'

    def sanitize_theme_palette_section(self, payload, required_keys):
        if not isinstance(payload, dict):
            return None
        sanitized = {}
        for key in required_keys:
            value = payload.get(key)
            if not isinstance(value, str) or not value.strip():
                return None
            sanitized[key] = value.strip()
        return sanitized

    def sanitize_theme_definition(self, theme_name, payload):
        if not isinstance(payload, dict):
            return None
        safe_name = str(payload.get('name') or theme_name).strip()
        if not safe_name:
            safe_name = str(theme_name).strip()
        if not safe_name:
            return None
        surface = self.sanitize_theme_palette_section(
            payload.get('surface'),
            (
                'text_bg',
                'text_fg',
                'cursor',
                'selection',
                'gutter_bg',
                'gutter_current_bg',
                'gutter_fg',
                'gutter_current_fg',
                'gutter_divider',
            ),
        )
        syntax = self.sanitize_theme_palette_section(
            payload.get('syntax'),
            (
                'keyword',
                'type',
                'string',
                'comment',
                'number',
                'preprocessor',
                'tag',
            ),
        )
        if surface is None or syntax is None:
            return None
        return {
            'name': safe_name,
            'surface': surface,
            'syntax': syntax,
        }

    def ensure_theme_files(self, theme_dir):
        builtins = self.get_builtin_theme_definitions()
        for theme_name, theme_payload in builtins.items():
            file_name = f"{self.slugify_theme_name(theme_name)}.json"
            file_path = os.path.join(theme_dir, file_name)
            if os.path.exists(file_path):
                continue
            payload = {
                'name': theme_name,
                'surface': dict(theme_payload['surface']),
                'syntax': dict(theme_payload['syntax']),
            }
            try:
                with open(file_path, 'w', encoding='utf-8') as theme_file:
                    json.dump(payload, theme_file, indent=2, ensure_ascii=False)
                    theme_file.write('\n')
            except OSError:
                continue

    def load_theme_definitions(self, theme_dir):
        builtins = self.get_builtin_theme_definitions()
        ordered_builtin_names = list(builtins.keys())
        loaded_themes = {}
        try:
            entries = sorted(
                [entry for entry in os.listdir(theme_dir) if entry.lower().endswith('.json')],
                key=lambda entry: entry.lower(),
            )
        except OSError:
            entries = []
        for entry in entries:
            file_path = os.path.join(theme_dir, entry)
            if not os.path.isfile(file_path):
                continue
            try:
                with open(file_path, 'r', encoding='utf-8') as theme_file:
                    payload = json.load(theme_file)
            except (OSError, json.JSONDecodeError):
                continue
            fallback_name = os.path.splitext(entry)[0].replace('_', ' ').strip().title()
            sanitized = self.sanitize_theme_definition(fallback_name, payload)
            if sanitized is None:
                continue
            loaded_themes[sanitized['name']] = sanitized

        ordered_themes = {}
        for theme_name in ordered_builtin_names:
            builtin_payload = builtins[theme_name]
            sanitized = loaded_themes.get(theme_name)
            if sanitized is None:
                sanitized = self.sanitize_theme_definition(theme_name, {'name': theme_name, **builtin_payload})
            ordered_themes[theme_name] = sanitized

        for theme_name in sorted(loaded_themes.keys(), key=lambda value: value.lower()):
            if theme_name not in ordered_themes:
                ordered_themes[theme_name] = loaded_themes[theme_name]

        return ordered_themes

    def get_available_syntax_theme_names(self):
        if not getattr(self, 'theme_definitions', None):
            return ['Default']
        return list(self.theme_definitions.keys())

    def get_syntax_theme_label(self, theme_name):
        theme_key_map = {
            'Default': 'syntax.theme.default',
            'Soft': 'syntax.theme.soft',
            'Vivid': 'syntax.theme.vivid',
            'Base4Tone': 'syntax.theme.base4tone',
            'Green Monochrome': 'syntax.theme.green_monochrome',
            'Orange Monochrome': 'syntax.theme.orange_monochrome',
        }
        locale_key = theme_key_map.get(theme_name)
        if locale_key:
            return self.tr(locale_key, theme_name)
        return theme_name

    def get_theme_creation_fields(self):
        return [
            ('surface', 'text_bg', 'theme.field.text_bg', 'Editor Background'),
            ('surface', 'text_fg', 'theme.field.text_fg', 'Editor Text'),
            ('surface', 'cursor', 'theme.field.cursor', 'Cursor'),
            ('surface', 'selection', 'theme.field.selection', 'Selection'),
            ('surface', 'gutter_bg', 'theme.field.gutter_bg', 'Gutter Background'),
            ('surface', 'gutter_current_bg', 'theme.field.gutter_current_bg', 'Current Line Gutter'),
            ('surface', 'gutter_fg', 'theme.field.gutter_fg', 'Gutter Text'),
            ('surface', 'gutter_current_fg', 'theme.field.gutter_current_fg', 'Current Gutter Text'),
            ('surface', 'gutter_divider', 'theme.field.gutter_divider', 'Gutter Divider'),
            ('syntax', 'keyword', 'theme.field.keyword', 'Keyword'),
            ('syntax', 'type', 'theme.field.type', 'Type'),
            ('syntax', 'string', 'theme.field.string', 'String'),
            ('syntax', 'comment', 'theme.field.comment', 'Comment'),
            ('syntax', 'number', 'theme.field.number', 'Number'),
            ('syntax', 'preprocessor', 'theme.field.preprocessor', 'Preprocessor'),
            ('syntax', 'tag', 'theme.field.tag', 'Tag'),
        ]

    def get_theme_file_path(self, theme_name):
        return os.path.join(self.theme_dir, f"{self.slugify_theme_name(theme_name)}.json")

    def build_theme_payload(self, theme_name, color_values):
        payload = {
            'name': theme_name,
            'surface': {},
            'syntax': {},
        }
        for section, field_name, _label_key, _label_default in self.get_theme_creation_fields():
            payload[section][field_name] = color_values[field_name]
        return payload

    def save_theme_payload(self, theme_name, payload):
        file_path = self.get_theme_file_path(theme_name)
        try:
            with open(file_path, 'w', encoding='utf-8') as theme_file:
                json.dump(payload, theme_file, indent=2, ensure_ascii=False)
                theme_file.write('\n')
        except OSError as exc:
            self.show_filesystem_error("Save Theme Failed", file_path, exc)
            return False
        self.theme_definitions = self.load_theme_definitions(self.theme_dir)
        self.create_menu()
        self.set_syntax_theme(theme_name)
        return True

    def show_create_theme_dialog(self):
        t = self.tr
        dialog = tk.Toplevel(self.root)
        dialog.title(t('theme.create.title', 'Create Theme'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color)
        self.apply_window_icon(dialog)

        current_theme = (getattr(self, 'theme_definitions', None) or {}).get(self.syntax_theme.get())
        if current_theme is None:
            current_theme = self.sanitize_theme_definition(
                'Default',
                {'name': 'Default', **self.get_builtin_theme_definitions()['Default']}
            )

        outer = tk.Frame(dialog, bg=self.bg_color, padx=16, pady=14)
        outer.pack(fill='both', expand=True)

        tk.Label(
            outer,
            text=t('theme.create.name', 'Theme Name:'),
            bg=self.bg_color,
            fg=self.fg_color,
            anchor='w'
        ).grid(row=0, column=0, columnspan=3, sticky='w')

        name_var = tk.StringVar(value="")
        name_entry = tk.Entry(
            outer,
            textvariable=name_var,
            bg=self.text_bg,
            fg=self.text_fg,
            insertbackground=self.cursor_color,
            relief='solid',
            width=32
        )
        name_entry.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(4, 12))

        color_values = {}
        preview_labels = {}

        def set_preview(field_name, value):
            preview = preview_labels[field_name]
            preview.configure(bg=value, text=value)

        def choose_color(field_name):
            chosen = colorchooser.askcolor(
                color=color_values[field_name],
                title=t('theme.create.pick_color', 'Choose...'),
                parent=dialog
            )[1]
            if chosen:
                color_values[field_name] = chosen
                set_preview(field_name, chosen)

        for row_index, (section, field_name, label_key, label_default) in enumerate(self.get_theme_creation_fields(), start=2):
            initial_value = current_theme[section][field_name]
            color_values[field_name] = initial_value
            tk.Label(
                outer,
                text=t(label_key, label_default),
                bg=self.bg_color,
                fg=self.fg_color,
                anchor='w'
            ).grid(row=row_index, column=0, sticky='w', padx=(0, 10), pady=3)

            preview = tk.Label(
                outer,
                text=initial_value,
                bg=initial_value,
                fg='#ffffff' if field_name not in {'text_bg', 'selection', 'gutter_bg', 'gutter_current_bg'} else '#000000',
                width=12,
                relief='solid'
            )
            preview.grid(row=row_index, column=1, sticky='ew', padx=(0, 10), pady=3)
            preview_labels[field_name] = preview

            tk.Button(
                outer,
                text=t('theme.create.pick_color', 'Choose...'),
                command=lambda current=field_name: choose_color(current)
            ).grid(row=row_index, column=2, sticky='ew', pady=3)

        outer.grid_columnconfigure(1, weight=1)

        button_row = tk.Frame(outer, bg=self.bg_color)
        button_row.grid(row=len(self.get_theme_creation_fields()) + 2, column=0, columnspan=3, pady=(14, 0))

        def save_theme():
            theme_name = name_var.get().strip()
            if not theme_name:
                messagebox.showerror(
                    t('theme.create.title', 'Create Theme'),
                    t('theme.create.name_required', 'Enter a theme name first.'),
                    parent=dialog
                )
                name_entry.focus_set()
                return
            for field_name, value in color_values.items():
                if not isinstance(value, str) or not value.strip():
                    messagebox.showerror(
                        t('theme.create.title', 'Create Theme'),
                        t('theme.create.invalid_color', 'Choose all theme colors before saving.'),
                        parent=dialog
                    )
                    return
            payload = self.build_theme_payload(theme_name, color_values)
            file_path = self.get_theme_file_path(theme_name)
            if os.path.exists(file_path):
                overwrite = messagebox.askyesno(
                    t('theme.create.exists_title', 'Overwrite Theme'),
                    t('theme.create.exists_message', 'A theme file with that name already exists.\n\nOverwrite it?'),
                    parent=dialog
                )
                if not overwrite:
                    return
            if self.save_theme_payload(theme_name, payload):
                dialog.destroy()

        tk.Button(button_row, text=t('theme.create.save', 'Save Theme'), command=save_theme).pack(side='left', padx=(0, 8))
        tk.Button(button_row, text=t('common.close', 'Close'), command=dialog.destroy).pack(side='left')

        dialog.bind('<Return>', lambda _event: save_theme())
        self.center_window(dialog, self.root)
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current, self.root))
        name_entry.focus_set()

    def serialize_locale_strings(self, strings):
        lines = []
        for key in sorted(strings):
            value = strings[key]
            if not isinstance(value, str):
                value = str(value)
            lines.append(f"{key}: {json.dumps(value, ensure_ascii=False)}")
        return "\n".join(lines) + "\n"

    def parse_locale_strings(self, content):
        try:
            payload = json.loads(content)
        except json.JSONDecodeError:
            payload = None
        if isinstance(payload, dict):
            parsed = {}
            for key, value in payload.items():
                if isinstance(key, str) and isinstance(value, str):
                    parsed[key] = value
            return parsed

        def decode_locale_scalar(value):
            if value is None:
                return None
            value = value.strip()
            if value == "":
                return ""
            try:
                decoded = json.loads(value)
            except json.JSONDecodeError:
                if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
                    decoded = value[1:-1]
                else:
                    decoded = value
            if isinstance(decoded, str):
                return decoded
            if decoded is None:
                return None
            return str(decoded)

        parsed = {}
        key_stack = []
        for raw_line in content.splitlines():
            stripped = raw_line.strip()
            if not stripped or stripped.startswith('#') or ':' not in raw_line:
                continue
            indent = len(raw_line) - len(raw_line.lstrip(' '))
            key, value = raw_line.split(':', 1)
            key = key.strip()
            if not key:
                continue
            while key_stack and indent <= key_stack[-1][0]:
                key_stack.pop()
            value = value.strip()
            full_key_parts = [entry[1] for entry in key_stack] + [key]
            full_key = ".".join(full_key_parts)
            if not value:
                key_stack.append((indent, key))
                continue
            decoded = decode_locale_scalar(value)
            if isinstance(decoded, str):
                parsed[full_key] = decoded
                if full_key.endswith('.label'):
                    parent_key = full_key[:-6]
                    if parent_key and parent_key not in parsed:
                        parsed[parent_key] = decoded
        return parsed

    def load_locale_strings(self, locale_path):
        strings = dict(DEFAULT_LOCALE_STRINGS)
        if not os.path.exists(locale_path):
            try:
                with open(locale_path, 'w', encoding='utf-8') as f:
                    f.write(self.serialize_locale_strings(strings))
            except OSError:
                return strings
        try:
            with open(locale_path, 'r', encoding='utf-8') as f:
                payload = self.parse_locale_strings(f.read())
            if isinstance(payload, dict):
                for key, value in payload.items():
                    if isinstance(key, str) and isinstance(value, str):
                        strings[key] = value
        except OSError:
            pass
        return strings

    def tr(self, key, default=None, **kwargs):
        value = self.locale_strings.get(key, default if default is not None else key)
        if not isinstance(value, str):
            value = str(value)
        if kwargs:
            try:
                return value.format(**kwargs)
            except Exception:
                return value
        return value

    def get_locale_file_path(self, locale_code, locale_dir=None):
        safe_code = str(locale_code or 'en_us').strip().lower()
        locale_dir = locale_dir or getattr(self, 'locale_dir', None) or (
            os.path.dirname(self.locale_path) if getattr(self, 'locale_path', None)
            else self.get_locale_dir(self.get_config_dir(self.app_dir))
        )
        candidate = os.path.join(locale_dir, f'{safe_code}.yml')
        try:
            for entry in os.listdir(locale_dir):
                if not entry.lower().endswith('.yml'):
                    continue
                if os.path.splitext(entry)[0].strip().lower() == safe_code:
                    return os.path.join(locale_dir, entry)
        except OSError:
            pass
        return candidate

    def get_available_language_codes(self):
        config_dir = getattr(self, 'locale_dir', None) or (
            os.path.dirname(self.locale_path) if getattr(self, 'locale_path', None)
            else self.get_locale_dir(self.get_config_dir(self.app_dir))
        )
        codes = []
        try:
            for entry in os.listdir(config_dir):
                if not entry.lower().endswith('.yml'):
                    continue
                code = os.path.splitext(entry)[0].strip().lower()
                if code:
                    codes.append(code)
        except OSError:
            pass
        if 'en_us' not in codes:
            codes.append('en_us')
        return sorted(set(codes), key=lambda value: (value != 'en_us', value))

    def get_language_display_name(self, locale_code):
        code = str(locale_code or '').strip().lower()
        if code in LOCALE_DISPLAY_NAMES:
            return LOCALE_DISPLAY_NAMES[code]
        parts = [part for part in code.split('_') if part]
        if parts:
            language = parts[0]
            region = parts[1] if len(parts) > 1 else ''
            if language in LANGUAGE_NATIVE_NAMES:
                language_name = LANGUAGE_NATIVE_NAMES[language]
                if region:
                    region_name = REGION_DISPLAY_NAMES.get(region, region.upper())
                    return f"{language_name} ({region_name})"
                return language_name
        parts = [part.upper() if len(part) <= 3 else part.title() for part in parts]
        return " / ".join(parts) if parts else 'Unknown'

    def is_rtl_locale(self, locale_code=None):
        code = str(locale_code or self.locale_code or '').strip().lower()
        return code in RTL_LOCALE_CODES

    def ui_anchor_start(self):
        return 'e' if self.is_rtl_locale() else 'w'

    def ui_anchor_end(self):
        return 'w' if self.is_rtl_locale() else 'e'

    def ui_justify(self):
        return 'right' if self.is_rtl_locale() else 'left'

    def get_available_font_families_map(self):
        try:
            families = tkfont.families(self.root)
        except Exception:
            families = ()
        font_map = {}
        for family in families:
            if not isinstance(family, str):
                continue
            normalized = family.strip().lower()
            if normalized and normalized not in font_map:
                font_map[normalized] = family
        return font_map

    def get_locale_font_candidates(self, locale_code=None):
        code = str(locale_code or self.locale_code or 'en_us').strip().lower()
        language = code.split('_', 1)[0]
        if language == 'ar':
            if self.is_windows:
                return ['Tahoma', 'Arial', 'Segoe UI', 'Courier New']
            return ['Noto Sans Arabic', 'Noto Naskh Arabic', 'DejaVu Sans', 'DejaVu Sans Mono', 'Liberation Sans', 'Monospace']
        if language == 'ja':
            if self.is_windows:
                return ['Yu Gothic UI', 'Yu Gothic', 'Meiryo', 'MS Gothic', 'Courier New']
            return ['Noto Sans CJK JP', 'Noto Sans JP', 'IPAGothic', 'VL Gothic', 'TakaoGothic', 'DejaVu Sans Mono']
        if language == 'zh':
            if self.is_windows:
                return ['Microsoft YaHei UI', 'Microsoft YaHei', 'SimSun', 'Courier New']
            return ['Noto Sans CJK SC', 'Noto Sans SC', 'WenQuanYi Zen Hei', 'AR PL UKai CN', 'DejaVu Sans Mono']
        if language == 'hi':
            if self.is_windows:
                return ['Nirmala UI', 'Mangal', 'Arial Unicode MS', 'Courier New']
            return ['Noto Sans Devanagari', 'Lohit Devanagari', 'Sahadeva', 'DejaVu Sans']
        if language == 'ru':
            return ['Consolas', 'Courier New', 'DejaVu Sans Mono', 'Liberation Mono', 'Monospace']
        return ['Courier New', 'Consolas', 'DejaVu Sans Mono', 'Liberation Mono', 'Monospace']

    def resolve_locale_font_family(self, locale_code=None):
        font_map = self.get_available_font_families_map()
        for candidate in self.get_locale_font_candidates(locale_code):
            normalized = candidate.strip().lower()
            if normalized in font_map:
                return font_map[normalized]
        return self.font_family or 'Courier New'

    def apply_locale(self, locale_code, persist=True):
        target_code = str(locale_code or 'en_us').strip().lower()
        target_path = self.get_locale_file_path(target_code)
        if not os.path.exists(target_path):
            target_code = 'en_us'
            target_path = self.get_locale_file_path(target_code)
        self.locale_code = target_code
        self.locale_path = target_path
        self.locale_strings = self.load_locale_strings(self.locale_path)
        self.app_name = self.tr('app.name', 'Notepad-X')
        self.font_family = self.resolve_locale_font_family(self.locale_code)
        self.note_color_labels = {
            'yellow': self.tr('note.filter.yellow', 'Yellow'),
            'green': self.tr('note.filter.green', 'Green'),
            'red': self.tr('note.filter.red', 'Red'),
            'blue': self.tr('note.filter.blue', 'Light Blue'),
        }
        if hasattr(self, 'language_selection'):
            self.language_selection.set(self.locale_code)
        if hasattr(self, 'menu'):
            self.create_menu()
        if hasattr(self, 'status'):
            self.status.config(text=self.tr('status.initial', "Ln 1 of 1, Col 1 | 0 characters | UTF-8 | Normal"))
            self.status.config(anchor=self.ui_anchor_start())
        if hasattr(self, 'status_tail'):
            self.status_tail.config(text=self.tr('status.memory_initial', " | Memory used: 0MB"))
            self.status_tail.config(anchor=self.ui_anchor_start())
        if hasattr(self, 'status_sync'):
            self.status_sync.config(text="")
            self.status_sync.config(anchor=self.ui_anchor_start())
        if hasattr(self, 'status_clock'):
            self.status_clock.config(anchor=self.ui_anchor_end())
        if hasattr(self, 'compare_status'):
            self.compare_status.config(anchor=self.ui_anchor_start())
        if hasattr(self, 'currently_editing_title_label'):
            self.currently_editing_title_label.config(
                text=self.tr('panel.currently_editing.title', 'Currently Editing'),
                anchor=self.ui_anchor_start(),
                justify=self.ui_justify(),
                wraplength=max(120, self.get_currently_editing_sidebar_width() - 20)
            )
        if hasattr(self, 'compare_title'):
            self.refresh_compare_header()
        if hasattr(self, 'root'):
            self.update_window_title()
        if hasattr(self, 'text'):
            self.update_font()
            self.update_status()
        if hasattr(self, 'currently_editing_content_label'):
            self.refresh_currently_editing_panel()
        if persist and hasattr(self, 'save_session'):
            self.save_session()

    def get_machine_profile_slug(self):
        user_name = os.environ.get('USERNAME') or os.environ.get('USER') or 'user'
        host_name = socket.gethostname() or 'host'
        raw_slug = f"{host_name}-{user_name}".lower()
        safe_slug = re.sub(r'[^a-z0-9._-]+', '-', raw_slug).strip('-')
        return safe_slug or 'default'

    def build_session_path(self, config_dir):
        base_name = f"Notepad-X.{self.machine_profile_slug}.session.json"
        if self.isolated_session:
            base_name = f"Notepad-X.{self.machine_profile_slug}.{os.getpid()}.session.json"
        return os.path.join(config_dir, base_name)

    def build_editor_identity_path(self, config_dir):
        base_name = f"Notepad-X.{self.machine_profile_slug}.editor.json"
        if self.isolated_session:
            base_name = f"Notepad-X.{self.machine_profile_slug}.{os.getpid()}.editor.json"
        return os.path.join(config_dir, base_name)

    def get_linux_desktop_entry_path(self):
        applications_dir = os.path.join(os.path.expanduser('~'), '.local', 'share', 'applications')
        return os.path.join(applications_dir, 'notepad-x.desktop')

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
        config_dir = self.get_config_dir(self.app_dir)
        os.makedirs(config_dir, exist_ok=True)
        self.locale_dir = self.get_locale_dir(config_dir)
        os.makedirs(self.locale_dir, exist_ok=True)
        self.migrate_language_files(config_dir=config_dir, locale_dir=self.locale_dir)
        self.session_path = self.build_session_path(config_dir)
        self.editor_identity_path = self.build_editor_identity_path(config_dir)
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.crash_log_path = os.path.join(self.app_dir, "Notepad-X.crash.log")

    def move_support_paths_to_emergency_dir(self):
        fallback_dir = self.get_emergency_support_dir()
        os.makedirs(fallback_dir, exist_ok=True)
        self.app_dir = fallback_dir
        config_dir = self.get_config_dir(self.app_dir)
        os.makedirs(config_dir, exist_ok=True)
        self.locale_dir = self.get_locale_dir(config_dir)
        os.makedirs(self.locale_dir, exist_ok=True)
        self.migrate_language_files(config_dir=config_dir, locale_dir=self.locale_dir)
        self.session_path = self.build_session_path(config_dir)
        self.editor_identity_path = self.build_editor_identity_path(config_dir)
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.crash_log_path = os.path.join(self.app_dir, "Notepad-X.crash.log")

    def utc_timestamp(self):
        return datetime.now(timezone.utc).isoformat(timespec='seconds')

    def trim_text(self, value, max_length):
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        return text[:max_length]

    def normalize_note_color(self, value):
        color_key = str(value or '').strip().lower()
        if color_key in self.note_colors:
            return color_key
        return 'yellow'

    def sanitize_note_responses(self, responses):
        sanitized = []
        if not isinstance(responses, list):
            return sanitized
        for response in responses:
            if not isinstance(response, dict):
                continue
            response_text = self.trim_text(response.get('text', ''), self.max_note_text_length)
            if not response_text:
                continue
            sanitized.append({
                'author_id': self.trim_text(self.normalize_optional_metadata(response.get('author_id')), 128),
                'author_label': self.trim_text(self.normalize_optional_metadata(response.get('author_label')), self.max_note_name_length),
                'author_host': self.trim_text(self.normalize_optional_metadata(response.get('author_host')), 128),
                'author_ip': self.trim_text(self.normalize_optional_metadata(response.get('author_ip')), 64),
                'text': response_text,
                'color': self.normalize_note_color(response.get('color')),
                'created_at': self.normalize_optional_metadata(response.get('created_at')),
            })
        return sanitized

    def get_local_machine_name(self):
        try:
            return self.trim_text(socket.gethostname() or None, 128)
        except OSError:
            return None

    def get_local_lan_ip(self):
        candidates = []
        try:
            probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            try:
                probe.connect(("8.8.8.8", 80))
                ip = probe.getsockname()[0]
                if ip:
                    candidates.append(ip)
            finally:
                probe.close()
        except OSError:
            pass
        try:
            hostname = socket.gethostname()
            for _, _, _, _, sockaddr in socket.getaddrinfo(hostname, None, socket.AF_INET):
                ip = sockaddr[0]
                if ip:
                    candidates.append(ip)
        except OSError:
            pass

        preferred = None
        fallback = None
        for ip in candidates:
            if not ip or ip.startswith("127."):
                continue
            fallback = fallback or ip
            if ip.startswith(("10.", "192.168.", "172.")):
                preferred = ip
                break
        return self.trim_text(preferred or fallback, 64)

    def get_note_color_hex(self, value):
        return self.note_colors[self.normalize_note_color(value)]

    def get_note_color_label(self, value):
        color_key = self.normalize_note_color(value)
        return self.note_color_labels.get(color_key, color_key.title())

    def parse_iso_datetime(self, value):
        if not isinstance(value, str):
            return None
        try:
            return datetime.fromisoformat(value.strip())
        except ValueError:
            return None

    def format_note_timestamp(self, value):
        parsed = self.parse_iso_datetime(value)
        if parsed is None:
            if not value:
                return None
            return str(value).replace('T', ' ')
        try:
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            parsed = parsed.astimezone()
        except Exception:
            pass
        return parsed.strftime('%Y-%m-%d %I:%M:%S %p')

    def read_json_file(self, file_path, context_name, default):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                payload = json.load(f)
        except (OSError, json.JSONDecodeError, UnicodeDecodeError, ValueError) as exc:
            self.log_exception(context_name, exc)
            return default
        return payload

    def write_json_atomically(self, file_path, payload, prefix, context_name):
        directory = os.path.dirname(file_path) or '.'
        os.makedirs(directory, exist_ok=True)
        target_mode = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode) | 0o664
            except OSError:
                target_mode = 0o664
        fd, temp_path = tempfile.mkstemp(prefix=prefix, suffix='.tmp', dir=directory)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as temp_file:
                json.dump(payload, temp_file, indent=2)
                temp_file.flush()
                os.fsync(temp_file.fileno())
            os.replace(temp_path, file_path)
            if target_mode is not None:
                os.chmod(file_path, target_mode)
            return True
        except OSError as exc:
            self.log_exception(context_name, exc)
            try:
                with open(file_path, 'w', encoding='utf-8') as direct_file:
                    json.dump(payload, direct_file, indent=2)
                    direct_file.flush()
                    os.fsync(direct_file.fileno())
                if target_mode is not None:
                    os.chmod(file_path, target_mode)
                return True
            except OSError as direct_exc:
                self.log_exception(f"{context_name} direct fallback", direct_exc)
                return False
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def write_binary_atomically(self, file_path, payload_bytes, prefix, context_name):
        directory = os.path.dirname(file_path) or '.'
        os.makedirs(directory, exist_ok=True)
        target_mode = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode)
            except OSError:
                target_mode = 0o664
        fd, temp_path = tempfile.mkstemp(prefix=prefix, suffix='.tmp', dir=directory)
        try:
            with os.fdopen(fd, 'wb') as temp_file:
                temp_file.write(payload_bytes)
                temp_file.flush()
                os.fsync(temp_file.fileno())
            os.replace(temp_path, file_path)
            if target_mode is not None:
                os.chmod(file_path, target_mode)
            return True
        except OSError as exc:
            self.log_exception(context_name, exc)
            return False
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def encryption_available(self):
        return AESGCM is not None

    def create_encryption_header(self, original_name=None, salt=None):
        return {
            'format': 'Notepad-X Encrypted',
            'version': self.encryption_version,
            'cipher': 'AES-256-GCM',
            'kdf': 'scrypt',
            'n': self.encryption_scrypt_n,
            'r': self.encryption_scrypt_r,
            'p': self.encryption_scrypt_p,
            'salt': base64.b64encode(salt or os.urandom(self.encryption_salt_length)).decode('ascii'),
            'original_name': os.path.basename(original_name) if original_name else None,
            'encoding': 'utf-8',
        }

    def derive_encryption_key(self, passphrase, header):
        if not self.encryption_available():
            raise RuntimeError("The cryptography package is required for encrypted files.")
        passphrase_bytes = str(passphrase or '').encode('utf-8')
        if not passphrase_bytes:
            raise ValueError("A passphrase is required.")
        salt_text = header.get('salt')
        if not isinstance(salt_text, str) or not salt_text.strip():
            raise ValueError("Encrypted file is missing its salt.")
        salt = base64.b64decode(salt_text.encode('ascii'))
        n = int(header.get('n', self.encryption_scrypt_n))
        r = int(header.get('r', self.encryption_scrypt_r))
        p = int(header.get('p', self.encryption_scrypt_p))
        required_memory = (128 * n * r) + (256 * r * p)
        maxmem = max(self.encryption_scrypt_maxmem, required_memory * 2)
        try:
            return hashlib.scrypt(
                passphrase_bytes,
                salt=salt,
                n=n,
                r=r,
                p=p,
                maxmem=maxmem,
                dklen=self.encryption_key_length
            )
        except ValueError as exc:
            if 'memory limit exceeded' in str(exc).lower():
                raise RuntimeError(
                    "Notepad-X could not derive the encryption key because the scrypt memory limit "
                    "was exceeded on this machine."
                ) from exc
            raise

    def build_encrypted_payload(self, header, nonce, ciphertext):
        header_bytes = json.dumps(header, separators=(',', ':')).encode('utf-8')
        return (
            self.encryption_magic
            + len(header_bytes).to_bytes(4, 'big')
            + header_bytes
            + nonce
            + ciphertext
        )

    def parse_encrypted_payload(self, payload_bytes):
        if not payload_bytes.startswith(self.encryption_magic):
            return None
        header_offset = len(self.encryption_magic)
        if len(payload_bytes) < header_offset + 4:
            raise ValueError("Encrypted file header is incomplete.")
        header_length = int.from_bytes(payload_bytes[header_offset:header_offset + 4], 'big')
        header_start = header_offset + 4
        header_end = header_start + header_length
        if header_length <= 0 or len(payload_bytes) < header_end + self.encryption_nonce_length:
            raise ValueError("Encrypted file header is invalid.")
        header = json.loads(payload_bytes[header_start:header_end].decode('utf-8'))
        nonce_start = header_end
        nonce_end = nonce_start + self.encryption_nonce_length
        nonce = payload_bytes[nonce_start:nonce_end]
        ciphertext = payload_bytes[nonce_end:]
        if not isinstance(header, dict) or header.get('format') != 'Notepad-X Encrypted':
            raise ValueError("Encrypted file header is invalid.")
        if header.get('cipher') != 'AES-256-GCM' or header.get('kdf') != 'scrypt':
            raise ValueError("Unsupported encrypted file settings.")
        if not ciphertext:
            raise ValueError("Encrypted file has no ciphertext.")
        return header, nonce, ciphertext

    def file_looks_encrypted(self, file_path):
        try:
            with open(file_path, 'rb') as encrypted_file:
                return encrypted_file.read(len(self.encryption_magic)) == self.encryption_magic
        except OSError:
            return False

    def show_encryption_unavailable(self, parent=None):
        messagebox.showerror(
            "Encryption Unavailable",
            "Encrypted save/open needs the 'cryptography' Python package.\n\n"
            "Install it on this machine to use Notepad-X encrypted files.",
            parent=parent or self.root
        )

    def prompt_encryption_options(self, default_encrypt=False, parent=None):
        parent = parent or self.root
        dialog = tk.Toplevel(parent)
        dialog.title("Save Encrypted Copy As")
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)
        self.apply_window_icon(dialog)

        result = {'value': None}
        encrypt_var = tk.BooleanVar(value=bool(default_encrypt))
        show_var = tk.BooleanVar(value=False)
        passphrase_var = tk.StringVar()
        confirm_var = tk.StringVar()

        tk.Checkbutton(
            dialog,
            text="Encrypt file",
            variable=encrypt_var,
            bg='#f0f0f0',
            anchor='w'
        ).pack(anchor='w', pady=(0, 8))

        tk.Label(dialog, text="Passphrase:", bg='#f0f0f0', fg='black', font=('Segoe UI', 9)).pack(anchor='w')
        passphrase_entry = tk.Entry(dialog, textvariable=passphrase_var, width=34, show='*')
        passphrase_entry.pack(fill='x', pady=(0, 6))

        tk.Label(dialog, text="Confirm passphrase:", bg='#f0f0f0', fg='black', font=('Segoe UI', 9)).pack(anchor='w')
        confirm_entry = tk.Entry(dialog, textvariable=confirm_var, width=34, show='*')
        confirm_entry.pack(fill='x', pady=(0, 6))

        def update_visibility(*_args):
            state = 'normal' if encrypt_var.get() else 'disabled'
            passphrase_entry.configure(state=state, show='' if show_var.get() else '*')
            confirm_entry.configure(state=state, show='' if show_var.get() else '*')

        tk.Checkbutton(
            dialog,
            text="Show passphrase",
            variable=show_var,
            bg='#f0f0f0',
            anchor='w',
            command=update_visibility
        ).pack(anchor='w')

        button_row = tk.Frame(dialog, bg='#f0f0f0')
        button_row.pack(fill='x', pady=(10, 0))

        def submit(event=None):
            if encrypt_var.get():
                if not self.encryption_available():
                    self.show_encryption_unavailable(dialog)
                    return "break"
                passphrase = passphrase_var.get()
                confirm = confirm_var.get()
                if not passphrase:
                    messagebox.showinfo("Save Encrypted Copy As", "Enter a passphrase first.", parent=dialog)
                    return "break"
                if passphrase != confirm:
                    messagebox.showinfo("Save Encrypted Copy As", "Passphrases do not match.", parent=dialog)
                    return "break"
                result['value'] = {'encrypt': True, 'passphrase': passphrase}
            else:
                result['value'] = {'encrypt': False, 'passphrase': None}
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()
            return "break"

        tk.Button(button_row, text="OK", width=10, command=submit).pack(side='left')
        tk.Button(button_row, text="Cancel", width=10, command=cancel).pack(side='right')

        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.update_idletasks()
        update_visibility()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        if encrypt_var.get():
            passphrase_entry.focus_force()
        else:
            dialog.focus_force()
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(1, lambda current=dialog: self.center_window_after_show(current, parent))
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.hide_autocomplete_popup()

    def prompt_open_passphrase(self, file_path, parent=None):
        parent = parent or self.root
        dialog = tk.Toplevel(parent)
        dialog.title("Open Encrypted File")
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)
        self.apply_window_icon(dialog)

        result = {'value': None}
        show_var = tk.BooleanVar(value=False)
        passphrase_var = tk.StringVar()

        tk.Label(
            dialog,
            text=f"Passphrase for:\n{os.path.basename(file_path)}",
            bg='#f0f0f0',
            fg='black',
            justify='left',
            font=('Segoe UI', 9)
        ).pack(anchor='w', pady=(0, 8))

        passphrase_entry = tk.Entry(dialog, textvariable=passphrase_var, width=34, show='*')
        passphrase_entry.pack(fill='x', pady=(0, 6))

        def update_visibility(*_args):
            passphrase_entry.configure(show='' if show_var.get() else '*')

        tk.Checkbutton(
            dialog,
            text="Show passphrase",
            variable=show_var,
            bg='#f0f0f0',
            anchor='w',
            command=update_visibility
        ).pack(anchor='w')

        button_row = tk.Frame(dialog, bg='#f0f0f0')
        button_row.pack(fill='x', pady=(10, 0))

        def submit(event=None):
            result['value'] = passphrase_var.get()
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()
            return "break"

        tk.Button(button_row, text="OK", width=10, command=submit).pack(side='left')
        tk.Button(button_row, text="Cancel", width=10, command=cancel).pack(side='right')

        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.update_idletasks()
        update_visibility()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        passphrase_entry.focus_force()
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(1, lambda current=dialog: self.center_window_after_show(current, parent))
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.hide_autocomplete_popup()

    def insert_text_content(self, doc, content):
        text = doc['text']
        doc['line_starts'] = None
        doc['total_file_lines'] = 1
        doc['window_start_line'] = 1
        doc['window_end_line'] = 1
        text.configure(state='normal')
        for offset in range(0, len(content), self.file_load_chunk_size):
            text.insert(tk.END, content[offset:offset + self.file_load_chunk_size])
            if doc.get('large_file_mode'):
                self.root.update_idletasks()
        doc['total_file_lines'] = max(1, int(text.index('end-1c').split('.')[0]))
        doc['window_end_line'] = doc['total_file_lines']

    def queue_background_file_result(self, result):
        with self.background_file_lock:
            self.background_file_results.append(result)

    def process_background_file_results(self):
        if not self.root.winfo_exists():
            return
        results = []
        with self.background_file_lock:
            if self.background_file_results:
                results = self.background_file_results[:]
                self.background_file_results.clear()
        for result in results:
            try:
                self.handle_background_file_result(result)
            except Exception as exc:
                self.log_exception("handle background file result", exc)
        self.root.after(60, self.process_background_file_results)

    def build_line_index_background(self, file_path):
        line_starts = [0]
        file_size = 0
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(self.file_load_chunk_size)
                if not chunk:
                    break
                base_offset = file_size
                file_size += len(chunk)
                line_starts.extend(base_offset + index + 1 for index, byte in enumerate(chunk) if byte == 10)
        return line_starts, file_size

    def start_background_text_load(self, doc, file_path):
        doc['background_loading'] = True
        doc['background_load_kind'] = 'text'
        doc['background_load_file_path'] = file_path
        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)
        text.insert('1.0', "Loading large file...\n")
        text.edit_modified(False)
        frame_id = str(doc['frame'])

        def worker():
            try:
                with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()
                self.queue_background_file_result({
                    'kind': 'text',
                    'tab_id': frame_id,
                    'file_path': file_path,
                    'content': content,
                })
            except Exception as exc:
                self.queue_background_file_result({
                    'kind': 'error',
                    'tab_id': frame_id,
                    'file_path': file_path,
                    'error': exc,
                })

        threading.Thread(target=worker, name='NotepadXLargeFileLoad', daemon=True).start()

    def start_background_virtual_index(self, doc, file_path):
        doc['background_loading'] = True
        doc['background_load_kind'] = 'virtual'
        doc['background_load_file_path'] = file_path
        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)
        text.insert('1.0', "Indexing large file...\n")
        text.edit_modified(False)
        frame_id = str(doc['frame'])

        def worker():
            try:
                line_starts, file_size = self.build_line_index_background(file_path)
                self.queue_background_file_result({
                    'kind': 'virtual',
                    'tab_id': frame_id,
                    'file_path': file_path,
                    'line_starts': line_starts,
                    'file_size_bytes': file_size,
                })
            except Exception as exc:
                self.queue_background_file_result({
                    'kind': 'error',
                    'tab_id': frame_id,
                    'file_path': file_path,
                    'error': exc,
                })

        threading.Thread(target=worker, name='NotepadXLargeFileIndex', daemon=True).start()

    def begin_background_text_insert(self, doc, content):
        doc['pending_insert_content'] = content
        doc['pending_insert_offset'] = 0
        doc['text'].configure(state='normal')
        doc['text'].delete('1.0', tk.END)
        self.continue_background_text_insert(doc)

    def continue_background_text_insert(self, doc):
        text = doc.get('text')
        content = doc.get('pending_insert_content')
        if not text or not isinstance(content, str):
            return
        offset = int(doc.get('pending_insert_offset', 0) or 0)
        batch_size = self.file_load_chunk_size * 2
        next_offset = min(len(content), offset + batch_size)
        if next_offset > offset:
            text.insert(tk.END, content[offset:next_offset])
            doc['pending_insert_offset'] = next_offset
        if next_offset < len(content):
            self.root.after(1, lambda current=doc: self.continue_background_text_insert(current))
            return
        doc.pop('pending_insert_content', None)
        doc.pop('pending_insert_offset', None)
        doc['total_file_lines'] = max(1, int(text.index('end-1c').split('.')[0]))
        doc['window_start_line'] = 1
        doc['window_end_line'] = doc['total_file_lines']
        text.edit_modified(False)
        text.mark_set(tk.INSERT, '1.0')
        text.tag_remove('sel', '1.0', tk.END)
        text.see('1.0')
        doc['last_insert_index'] = '1.0'
        doc['last_yview'] = 0.0
        doc['last_xview'] = 0.0
        doc['background_loading'] = False
        doc['background_load_kind'] = None
        doc['background_load_file_path'] = None
        self.update_doc_file_signature(doc)
        self.configure_syntax_highlighting(doc['frame'])
        self.restore_doc_notes(doc)
        self.register_doc_for_shared_notes(doc)
        self.refresh_tab_title(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        if str(doc['frame']) == self.notebook.select():
            self.update_status()

    def handle_background_file_error(self, doc, exc):
        doc['background_loading'] = False
        doc['background_load_kind'] = None
        doc['background_load_file_path'] = None
        doc.pop('pending_insert_content', None)
        doc.pop('pending_insert_offset', None)
        file_path = doc.get('file_path')
        messagebox.showerror("Open Failed", f"Notepad-X could not open:\n{file_path}\n\n{exc}", parent=self.root)
        if doc.get('background_open_new_tab'):
            try:
                self.notebook.forget(doc['frame'])
            except tk.TclError:
                pass
            self.documents.pop(str(doc['frame']), None)
        else:
            doc['file_path'] = None
            doc['text'].delete('1.0', tk.END)
        doc['background_open_new_tab'] = False

    def handle_background_file_result(self, result):
        tab_id = str(result.get('tab_id') or '')
        doc = self.documents.get(tab_id)
        if not doc:
            return
        file_path = os.path.abspath(str(result.get('file_path') or ''))
        if not file_path or os.path.abspath(str(doc.get('file_path') or '')) != file_path:
            return
        if result.get('kind') == 'error':
            self.handle_background_file_error(doc, result.get('error'))
            return
        if result.get('kind') == 'virtual':
            doc['line_starts'] = result.get('line_starts') or [0]
            doc['file_size_bytes'] = int(result.get('file_size_bytes') or 0)
            doc['total_file_lines'] = max(1, len(doc['line_starts']))
            doc['background_loading'] = False
            doc['background_load_kind'] = None
            doc['background_load_file_path'] = None
            doc['window_start_line'] = 1
            doc['window_end_line'] = 1
            self.load_virtual_window(doc, 1)
            self.refresh_tab_title(doc['frame'])
            if str(doc['frame']) == self.notebook.select():
                self.update_status()
            doc['background_open_new_tab'] = False
            return
        if result.get('kind') == 'text':
            self.begin_background_text_insert(doc, result.get('content') or '')
            doc['background_open_new_tab'] = False

    def write_encrypted_text_file(self, file_path, text_content, passphrase=None, header=None, key=None, original_name=None):
        if not self.encryption_available():
            raise RuntimeError("Encryption support is unavailable.")
        encryption_header = dict(header or {})
        if key is None:
            encryption_header = self.create_encryption_header(original_name=original_name or file_path)
            key = self.derive_encryption_key(passphrase, encryption_header)
        else:
            if not encryption_header:
                raise ValueError("Encrypted file metadata is missing.")
            encryption_header['original_name'] = encryption_header.get('original_name') or os.path.basename(original_name or file_path)
        plaintext_bytes = str(text_content).encode('utf-8')
        nonce = os.urandom(self.encryption_nonce_length)
        ciphertext = AESGCM(key).encrypt(nonce, plaintext_bytes, self.encryption_magic)
        payload_bytes = self.build_encrypted_payload(encryption_header, nonce, ciphertext)
        if not self.write_binary_atomically(file_path, payload_bytes, 'notepadx-encrypted-', 'write encrypted file'):
            raise OSError(f"Could not write encrypted file: {file_path}")
        return encryption_header, key

    def read_encrypted_text_file(self, file_path):
        if not self.file_looks_encrypted(file_path):
            return None
        if not self.encryption_available():
            self.show_encryption_unavailable(self.root)
            raise OSError("Encryption support is unavailable.")
        with open(file_path, 'rb') as encrypted_file:
            payload_bytes = encrypted_file.read()
        header, nonce, ciphertext = self.parse_encrypted_payload(payload_bytes)
        while True:
            passphrase = self.prompt_open_passphrase(file_path, parent=self.root)
            if passphrase is None:
                raise OSError("Encrypted file open cancelled.")
            try:
                key = self.derive_encryption_key(passphrase, header)
                plaintext_bytes = AESGCM(key).decrypt(nonce, ciphertext, self.encryption_magic)
                return plaintext_bytes.decode(header.get('encoding') or 'utf-8'), header, key
            except RuntimeError:
                raise
            except (InvalidTag, ValueError, UnicodeDecodeError):
                retry = messagebox.askretrycancel(
                    "Open Encrypted File",
                    "That passphrase did not unlock the encrypted file.",
                    parent=self.root
                )
                if not retry:
                    raise OSError("Encrypted file open cancelled.")

    def get_file_signature(self, file_path):
        try:
            stat = os.stat(file_path)
        except OSError:
            return None
        return (stat.st_mtime_ns, stat.st_size)

    def update_doc_file_signature(self, doc):
        file_path = doc.get('file_path') if doc else None
        signature = self.get_file_signature(file_path) if file_path else None
        if doc is not None:
            doc['file_signature'] = signature
        return signature

    def confirm_external_file_change(self, doc):
        file_path = doc.get('file_path') if doc else None
        if not file_path or not os.path.exists(file_path):
            return True
        current_signature = self.get_file_signature(file_path)
        known_signature = doc.get('file_signature')
        if known_signature is None or current_signature is None or current_signature == known_signature:
            return True
        answer = messagebox.askyesno(
            "File Changed on Disk",
            "This file changed on disk after it was opened.\n\n"
            "Overwrite the newer disk version with your current editor contents?",
            parent=self.root
        )
        if answer:
            return True
        refreshed = messagebox.askyesno(
            "Reload File",
            "Do you want Notepad-X to reload the file from disk instead?",
            parent=self.root
        )
        if refreshed:
            self.load_content_into_doc(doc, file_path)
            self.set_active_document(doc['frame'])
        return False

    def path_looks_safe_for_shell(self, file_path):
        if not isinstance(file_path, str) or not file_path.strip():
            return False
        file_path = os.path.abspath(file_path)
        if any(char in file_path for char in '\r\n'):
            return False
        return True

    def show_filesystem_error(self, title, file_path, exc):
        location = os.path.abspath(file_path) if file_path else "that path"
        messagebox.showerror(title, f"Notepad-X could not access:\n{location}\n\n{exc}", parent=self.root)

    def is_probably_binary_file(self, file_path, sample_size=8192):
        try:
            with open(file_path, 'rb') as f:
                sample = f.read(sample_size)
        except OSError:
            return False
        if not sample:
            return False
        if b'\x00' in sample:
            return True
        text_bytes = b'\t\n\r\b\f' + bytes(range(32, 127))
        non_text = sum(byte not in text_bytes for byte in sample)
        return (non_text / max(1, len(sample))) > 0.30

    def sanitize_session_payload(self, session):
        if not isinstance(session, dict):
            return None
        open_files = []
        for path in session.get('open_files', []):
            if isinstance(path, str) and os.path.exists(path):
                open_files.append(path)
        open_files = list(dict.fromkeys(open_files))[:self.max_session_files]

        recent_files = []
        for path in session.get('recent_files', []):
            if isinstance(path, str) and os.path.exists(path):
                recent_files.append(path)
        recent_files = list(dict.fromkeys(recent_files))[:self.max_recent_files]

        closed_files = []
        for path in session.get('closed_session_files', []):
            if isinstance(path, str):
                closed_files.append(path)

        selected_file = session.get('selected_file')
        if not isinstance(selected_file, str) or selected_file not in open_files:
            selected_file = None

        compare_file = session.get('compare_file')
        if not isinstance(compare_file, str) or compare_file not in open_files:
            compare_file = None

        compare_base_file = session.get('compare_base_file')
        if not isinstance(compare_base_file, str) or compare_base_file not in open_files:
            compare_base_file = None

        syntax_theme = str(session.get('syntax_theme', 'Default'))
        if syntax_theme not in self.get_available_syntax_theme_names():
            syntax_theme = 'Default'

        try:
            current_font_size = int(session.get('current_font_size', self.base_font_size))
        except (TypeError, ValueError):
            current_font_size = self.base_font_size
        current_font_size = max(self.min_font_size, min(self.max_font_size, current_font_size))
        locale_code = str(session.get('locale_code', self.locale_code)).strip().lower()
        if not os.path.exists(self.get_locale_file_path(locale_code)):
            locale_code = self.locale_code

        return {
            'open_files': open_files,
            'selected_file': selected_file,
            'recent_files': recent_files,
            'closed_session_files': closed_files,
            'sound_enabled': bool(session.get('sound_enabled', True)),
            'status_bar_enabled': bool(session.get('status_bar_enabled', True)),
            'numbered_lines_enabled': bool(session.get('numbered_lines_enabled', True)),
            'autocomplete_enabled': bool(session.get('autocomplete_enabled', True)),
            'sync_page_navigation_enabled': bool(session.get('sync_page_navigation_enabled', False)),
            'edit_with_shell_enabled': bool(session.get('edit_with_shell_enabled', False)),
            'current_font_size': current_font_size,
            'syntax_theme': syntax_theme,
            'locale_code': locale_code,
            'compare_file': compare_file,
            'compare_base_file': compare_base_file,
        }

    def sanitize_recovery_payload(self, recovery):
        if not isinstance(recovery, dict):
            return None
        recovery_tabs = []
        source_tabs = recovery.get('recovery_tabs', recovery.get('unsaved_tabs', []))
        for tab in source_tabs:
            if not isinstance(tab, dict):
                continue
            file_path = tab.get('file_path')
            if isinstance(file_path, str) and file_path.strip():
                file_path = os.path.abspath(file_path)
            else:
                file_path = None
            untitled_name = self.trim_text(tab.get('untitled_name'), 120)
            if not untitled_name and not file_path:
                untitled_name = self.next_untitled_name()
            content = tab.get('content', '')
            if not isinstance(content, str):
                content = str(content)
            recovery_tabs.append({
                'file_path': file_path,
                'untitled_name': untitled_name,
                'content': content,
                'modified': bool(tab.get('modified', True))
            })
            if len(recovery_tabs) >= self.max_recovery_tabs:
                break
        selected_recovery_key = self.trim_text(
            recovery.get('selected_recovery_key') or recovery.get('selected_untitled'),
            260
        )
        timestamp = self.trim_text(recovery.get('timestamp'), 64)
        return {
            'recovery_tabs': recovery_tabs,
            'selected_recovery_key': selected_recovery_key,
            'timestamp': timestamp,
        }

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
        if not self.sound_enabled.get() or not os.path.exists(sound_path):
            return
        if self.is_windows and self.winmm:
            sound_path = sound_path.replace('"', '""')
            try:
                self.winmm.mciSendStringW('close notepadx_note', None, 0, None)
                self.winmm.mciSendStringW(f'open "{sound_path}" type mpegvideo alias notepadx_note', None, 0, None)
                self.winmm.mciSendStringW('play notepadx_note from 0', None, 0, None)
            except Exception:
                pass
            return
        if self.is_linux:
            for player in (['paplay', sound_path], ['aplay', sound_path], ['ffplay', '-nodisp', '-autoexit', sound_path], ['xdg-open', sound_path]):
                if shutil.which(player[0]):
                    try:
                        subprocess.Popen(player, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return
                    except Exception:
                        continue

    def play_unread_note_sound(self):
        self.play_sound(self.note_sound_path)

    def play_delete_note_sound(self):
        self.play_sound(self.delete_note_sound_path)

    def hide_support_file(self, file_path):
        if not self.is_windows or not file_path or not os.path.exists(file_path):
            return
        try:
            attributes = self.kernel32.GetFileAttributesW(file_path)
            if attributes == 0xFFFFFFFF:
                return
            hidden_attributes = attributes | 0x2
            self.kernel32.SetFileAttributesW(file_path, hidden_attributes)
        except Exception:
            pass

    def show_support_file(self, file_path):
        if not self.is_windows or not file_path or not os.path.exists(file_path):
            return
        try:
            attributes = self.kernel32.GetFileAttributesW(file_path)
            if attributes == 0xFFFFFFFF:
                return
            visible_attributes = attributes & ~0x2
            self.kernel32.SetFileAttributesW(file_path, visible_attributes)
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
        if not os.path.exists(self.editor_identity_path):
            return []
        identity = self.read_json_file(self.editor_identity_path, "load editor identity", {})
        if isinstance(identity, dict):
            known_ids = identity.get('known_editor_ids', [])
            if isinstance(known_ids, list):
                return [str(editor_id).strip() for editor_id in known_ids if str(editor_id).strip()]
            legacy_id = str(identity.get('editor_id', '')).strip()
            if legacy_id:
                return [legacy_id]
        return []

    def generate_editor_id(self):
        seed = f"notepad-x-{os.getpid()}-{datetime.now().timestamp()}-{secrets.token_hex(16)}"
        return hashlib.md5(seed.encode('utf-8')).hexdigest()

    def persist_editor_identity(self):
        known_ids = list(dict.fromkeys(self.known_editor_ids + [self.editor_id]))[-32:]
        for attempt in range(3):
            try:
                if self.write_json_atomically(
                    self.editor_identity_path,
                    {
                        'editor_id': self.editor_id,
                        'known_editor_ids': known_ids
                    },
                    'notepadx-editor-',
                    'persist editor identity'
                ):
                    return
                raise PermissionError(self.editor_identity_path)
            except PermissionError as exc:
                if attempt == 0:
                    self.move_support_paths_to_user_dir()
                    continue
                if attempt == 1:
                    self.move_support_paths_to_emergency_dir()
                    continue
                self.log_exception("persist editor identity", exc)
                return
            except Exception as exc:
                self.log_exception("persist editor identity", exc)
                return

    def center_window(self, window, parent=None):
        try:
            window.update_idletasks()
        except tk.TclError:
            return

        width = max(window.winfo_reqwidth(), window.winfo_width())
        height = max(window.winfo_reqheight(), window.winfo_height())
        if width <= 1 or height <= 1:
            try:
                geometry = window.geometry().split('+', 1)[0]
                width_text, height_text = geometry.split('x', 1)
                width = max(width, int(width_text))
                height = max(height, int(height_text))
            except (ValueError, tk.TclError):
                pass

        if parent is not None and parent.winfo_exists():
            try:
                parent.update_idletasks()
                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                parent_width = max(parent.winfo_width(), parent.winfo_reqwidth())
                parent_height = max(parent.winfo_height(), parent.winfo_reqheight())
            except tk.TclError:
                parent = None
            else:
                x = parent_x + max(0, (parent_width - width) // 2)
                y = parent_y + max(0, (parent_height - height) // 2)
        if parent is None:
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

    def is_notepadx_support_file(self, file_path):
        return is_notepadx_support_file_path(file_path)

    def start_single_instance_server(self):
        if self.isolated_session:
            return
        try:
            server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            server.bind((self.single_instance_host, self.single_instance_port))
            server.listen(5)
            server.settimeout(0.5)
        except OSError:
            try:
                server.close()
            except Exception:
                pass
            self.single_instance_server = None
            return

        self.single_instance_server = server
        self.single_instance_running = True

        def listen():
            while self.single_instance_running:
                try:
                    connection, _ = server.accept()
                except socket.timeout:
                    continue
                except OSError:
                    break

                with connection:
                    data = b''
                    while True:
                        try:
                            chunk = connection.recv(65536)
                        except OSError:
                            break
                        if not chunk:
                            break
                        data += chunk
                    if not data:
                        continue
                    try:
                        payload = json.loads(data.decode('utf-8'))
                    except Exception:
                        continue
                    if payload.get('command') != 'open_files':
                        continue
                    incoming_files = []
                    for raw_path in payload.get('files', []):
                        if not raw_path:
                            continue
                        candidate_path = os.path.abspath(str(raw_path))
                        if self.is_notepadx_support_file(candidate_path):
                            continue
                        incoming_files.append(candidate_path)
                    if incoming_files:
                        with self.remote_open_lock:
                            self.remote_open_files.extend(incoming_files)

        self.single_instance_listener_thread = threading.Thread(
            target=listen,
            name='NotepadXSingleInstanceListener',
            daemon=True
        )
        self.single_instance_listener_thread.start()

    def stop_single_instance_server(self):
        self.single_instance_running = False
        server = self.single_instance_server
        self.single_instance_server = None
        if server is not None:
            try:
                server.close()
            except OSError:
                pass

    def process_remote_open_requests(self):
        if not self.root.winfo_exists():
            return
        pending_files = []
        with self.remote_open_lock:
            if self.remote_open_files:
                pending_files = self.remote_open_files[:]
                self.remote_open_files.clear()
        if pending_files:
            self.open_startup_files(pending_files)
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            except tk.TclError:
                pass
        self.root.after(150, self.process_remote_open_requests)

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
        if not self.kernel32 or not self.psapi:
            if self.is_linux:
                try:
                    with open('/proc/self/status', 'r', encoding='utf-8', errors='replace') as status_file:
                        for line in status_file:
                            if line.startswith('VmRSS:'):
                                parts = line.split()
                                if len(parts) >= 2:
                                    return max(1, round(int(parts[1]) / 1024))
                except Exception:
                    pass
            if resource is not None:
                try:
                    usage = resource.getrusage(resource.RUSAGE_SELF)
                    rss = int(getattr(usage, 'ru_maxrss', 0))
                    if rss > 0:
                        if self.is_linux:
                            return max(1, round(rss / 1024))
                        return max(1, round(rss / (1024 * 1024)))
                except Exception:
                    pass
            return 0
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
            text=self.tr('status.initial', "Ln 1 of 1, Col 1 | 0 characters | UTF-8 | Normal"),
            anchor=self.ui_anchor_start(),
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=0, pady=4
        )
        self.status.pack(side='left')

        self.status_sync = tk.Label(
            self.status_left,
            text="",
            anchor=self.ui_anchor_start(),
            bg='#2d2d2d',
            fg='#4ecb71',
            font=('Segoe UI', 9, 'bold'),
            padx=0, pady=4
        )
        self.status_sync.pack(side='left')

        self.status_tail = tk.Label(
            self.status_left,
            text=self.tr('status.memory_initial', " | Memory used: 0MB"),
            anchor=self.ui_anchor_start(),
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=0, pady=4
        )
        self.status_tail.pack(side='left')

        self.status_clock = tk.Label(
            self.status_frame,
            text="",
            anchor=self.ui_anchor_end(),
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=8, pady=4
        )
        self.status_clock.grid(row=0, column=1, sticky='e')

        self.compare_status = tk.Label(
            self.status_frame,
            text="",
            anchor=self.ui_anchor_start(),
            bg='#2d2d2d',
            fg='#d4d4d4',
            font=('Segoe UI', 9),
            padx=0,
            pady=4
        )
        self.compare_status.place_forget()
        self.status_frame.bind('<Configure>', self.position_compare_status, add='+')

    def position_compare_status(self, event=None):
        if not hasattr(self, 'compare_status') or not hasattr(self, 'status_frame'):
            return
        if not self.compare_active:
            self.compare_status.place_forget()
            return
        try:
            self.status_frame.update_idletasks()
            if not self.compare_container.winfo_ismapped():
                self.compare_status.place_forget()
                return
            status_root_x = self.status_frame.winfo_rootx()
            compare_root_x = self.compare_container.winfo_rootx()
            compare_root_right = compare_root_x + self.compare_container.winfo_width()
            start_x = max(0, compare_root_x - status_root_x)
            end_x = min(
                compare_root_right - status_root_x,
                self.status_clock.winfo_x()
            )
            available_width = max(120, end_x - start_x)
            self.compare_status.place(x=start_x, y=0, width=available_width, relheight=1.0)
            self.compare_status.lift()
        except tk.TclError:
            self.compare_status.place_forget()

    def update_clock(self):
        if hasattr(self, 'status_clock') and self.status_clock.winfo_exists():
            self.status_clock.config(
                text=datetime.now().strftime("%A | %I:%M:%S %p | %m/%d/%Y").lstrip('0').replace('/0', '/')
            )
            self.root.after(1000, self.update_clock)

    def get_zoom_text(self):
        if self.current_font_size == self.base_font_size:
            return self.tr('status.mode.normal', 'Normal')
        percent = round((self.current_font_size / self.base_font_size) * 100)
        return f"+{percent-100}%" if percent > 100 else f"{percent-100}%"

    def build_editor_status_text(self, doc, text_widget):
        row, col = text_widget.index(tk.INSERT).split('.')
        row = int(row)
        col = int(col) + 1

        if doc and doc.get('virtual_mode'):
            row = doc['window_start_line'] + row - 1
            total_lines = doc['total_file_lines']
            total_chars = doc['file_size_bytes']
            char_info = self.tr(
                'status.byte_count',
                '{total_bytes} {bytes_label}',
                total_bytes=f"{total_chars:,}",
                bytes_label=self.tr('status.bytes', 'bytes')
            )
        else:
            full_content = text_widget.get('1.0', 'end-1c')
            total_lines = int(text_widget.index('end-1c').split('.')[0])
            total_chars = len(full_content)
            try:
                sel_start = text_widget.index('sel.first')
                sel_end = text_widget.index('sel.last')
                selected_count = len(text_widget.get(sel_start, sel_end))
            except tk.TclError:
                selected_count = 0
            if selected_count > 0:
                char_info = self.tr(
                    'status.selected_char_count',
                    '{selected_count} {of_label} {total_chars} {characters_label}',
                    selected_count=f"{selected_count:,}",
                    of_label=self.tr('status.of', 'of'),
                    total_chars=f"{total_chars:,}",
                    characters_label=self.tr('status.characters', 'characters')
                )
            else:
                char_info = self.tr(
                    'status.char_count',
                    '{total_chars} {characters_label}',
                    total_chars=f"{total_chars:,}",
                    characters_label=self.tr('status.characters', 'characters')
                )

        zoom_text = self.get_zoom_text()
        mode_suffix = ""
        if doc and doc.get('virtual_mode'):
            mode_suffix = f" | {self.tr('status.mode.virtual', 'Virtual')}"
        elif doc and doc.get('preview_mode'):
            mode_suffix = f" | {self.tr('status.mode.preview', 'Preview')}"

        return self.tr(
            'status.main',
            '{line_label} {row} {of_label} {total_lines}, {col_label} {col} | {char_info} | {encoding} | {zoom_text}{mode_suffix}',
            line_label=self.tr('status.line', 'Ln'),
            row=row,
            of_label=self.tr('status.of', 'of'),
            total_lines=total_lines,
            col_label=self.tr('status.col', 'Col'),
            col=col,
            char_info=char_info,
            encoding=self.tr('status.encoding', 'UTF-8'),
            zoom_text=zoom_text,
            mode_suffix=mode_suffix
        )

    def update_status(self):
        if not self.text or not hasattr(self, 'status'):
            return
        current_doc = self.get_current_doc()
        status_main_text = self.build_editor_status_text(current_doc, self.text)
        editor_label_text = ""
        if current_doc and current_doc.get('file_path'):
            editor_label_text = self.tr('status.editor_id', ' | ID: {editor_id}', editor_id=f"Notepad-X-{self.editor_id}")
        shared_notes_tail = ""
        if current_doc and current_doc.get('file_path'):
            unread_count = self.get_unread_note_count(current_doc)
            shared_notes_tail = self.tr(
                'status.unread_tail',
                ' | {unread_count} unread (F3 to view) | ({active_editors} editing)',
                unread_count=unread_count,
                active_editors=current_doc.get('note_active_editors', 0)
            )
        status_main_text = f"{status_main_text}{editor_label_text}"
        status_tail_text = f"{shared_notes_tail}{self.tr('status.memory', ' | Memory used: {memory_mb}MB', memory_mb=self.memory_used_mb)}"
        self.status.config(text=status_main_text)
        self.status_sync.config(text=self.tr('status.synced', '| Notes Synced') if current_doc and current_doc.get('file_path') else "")
        self.status_tail.config(text=status_tail_text)

        if hasattr(self, 'compare_status') and self.compare_view and self.compare_active:
            try:
                compare_widget = self.compare_view.get('text')
                compare_doc = self.get_doc_for_text_widget(compare_widget) if compare_widget is not None else None
                if compare_widget is None or compare_doc is None:
                    raise tk.TclError
                self.compare_status.config(
                    text=self.build_editor_status_text(compare_doc, compare_widget)
                )
                self.position_compare_status()
            except tk.TclError:
                self.compare_status.config(text="")
                self.compare_status.place_forget()
        elif hasattr(self, 'compare_status'):
            self.compare_status.config(text="")
            self.compare_status.place_forget()

        if self.currently_editing_panel_visible:
            self.refresh_currently_editing_panel()

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
        surface = self.get_syntax_surface_palette()
        gutter = tk.Canvas(
            parent,
            width=56,
            bg=surface['gutter_bg'],
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

        try:
            if doc.get('virtual_mode'):
                max_line_number = max(1, int(doc.get('total_file_lines', 1)))
            else:
                max_line_number = max(1, int(text.index('end-1c').split('.')[0]))
        except (tk.TclError, TypeError, ValueError):
            max_line_number = 1

        gutter_font = tkfont.Font(family=self.font_family, size=max(9, self.current_font_size - 1))
        desired_gutter_width = max(56, gutter_font.measure('9' * len(str(max_line_number))) + 24)
        current_gutter_width = int(gutter.cget('width'))
        if current_gutter_width != desired_gutter_width:
            gutter.configure(width=desired_gutter_width)

        surface = self.get_syntax_surface_palette()
        gutter.configure(bg=surface['gutter_bg'])
        gutter.delete('all')
        gutter_height = max(gutter.winfo_height(), text.winfo_height(), 1)
        gutter_width = int(gutter.cget('width'))
        current_line = 1
        try:
            current_line = int(text.index(tk.INSERT).split('.')[0])
        except tk.TclError:
            pass

        gutter.create_rectangle(0, 0, gutter_width - 1, gutter_height, fill=surface['gutter_bg'], outline='')
        gutter.create_rectangle(gutter_width - 1, 0, gutter_width, gutter_height, fill=surface['gutter_divider'], outline='')

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
                    gutter.create_rectangle(0, y, gutter_width - 1, y + line_height, fill=surface['gutter_current_bg'], outline='')
                    line_fg = surface['gutter_current_fg']
                else:
                    line_fg = surface['gutter_fg']
                gutter.create_text(
                    gutter_width - 10,
                    y + (line_height / 2),
                    anchor='e',
                    text=str(display_line),
                    fill=line_fg,
                    font=gutter_font
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
        self.find_entry.bind('<Return>', self.find_from_input)
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

        self.replace_find_entry.bind('<Return>', self.find_from_input)     # ← added
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

    def refresh_currently_editing_panel(self):
        if not hasattr(self, 'currently_editing_content_label'):
            return

        current_doc = self.get_current_doc()
        if not current_doc or not current_doc.get('file_path'):
            lines = [self.tr('panel.currently_editing.unsaved', 'Save this tab to track currently editing IDs.')]
        else:
            editors = self.sanitize_shared_editors(current_doc.get('note_editors', []))
            if not editors:
                lines = [self.tr('panel.currently_editing.none', 'No active editing IDs for this file.')]
            else:
                lines = []
                for entry in editors:
                    editor_id = entry.get('id', '')
                    if not editor_id:
                        continue
                    header_parts = []
                    if entry.get('host'):
                        header_parts.append(entry['host'])
                    if entry.get('ip'):
                        header_parts.append(entry['ip'])
                    if header_parts:
                        lines.append(" | ".join(header_parts))
                    lines.append(editor_id)
                    lines.append("")
                if lines and not lines[-1].strip():
                    lines.pop()
                if not lines:
                    lines = [self.tr('panel.currently_editing.none', 'No active editing IDs for this file.')]

        try:
            self.currently_editing_content_label.config(
                text="\n".join(lines),
                justify=self.ui_justify(),
                anchor='nw'
            )
        except tk.TclError:
            pass

    def show_currently_editing_panel(self):
        if self.find_panel_visible:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.cancel_find_change_job()
            self.clear_find_highlights()
        if self.replace_panel_visible:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.cancel_find_change_job()
            self.clear_find_highlights()

        self.refresh_currently_editing_panel()
        self.currently_editing_panel_visible = True
        self.currently_editing_sidebar.configure(width=self.get_currently_editing_sidebar_width())
        self.currently_editing_sidebar.grid()
        self.rebuild_editor_panes()
        self.schedule_compare_layout_refresh()
        self.root.after_idle(self.position_compare_status)
        self.restore_editor_focus_after_panel_change()
        return "break"

    def close_currently_editing_panel(self, event=None):
        self.currently_editing_panel_visible = False
        self.currently_editing_sidebar.grid_remove()
        self.rebuild_editor_panes()
        self.schedule_compare_layout_refresh()
        self.root.after_idle(self.position_compare_status)
        self.restore_editor_focus_after_panel_change()
        return "break"

    def is_currently_editing_panel_open(self):
        if not hasattr(self, 'currently_editing_sidebar'):
            return False
        try:
            return self.currently_editing_panel_visible or self.currently_editing_sidebar.winfo_ismapped()
        except tk.TclError:
            return False

    def restore_editor_focus_after_panel_change(self):
        target = self.get_active_search_widget()
        if target is None:
            return

        def apply_focus(widget=target):
            try:
                self.root.focus_force()
            except tk.TclError:
                return
            try:
                widget.focus_force()
            except tk.TclError:
                pass

        apply_focus()
        self.root.after(10, apply_focus)
        self.root.after(50, apply_focus)

    def toggle_currently_editing_panel(self, event=None):
        if self.is_currently_editing_panel_open():
            return self.close_currently_editing_panel(event)
        return self.show_currently_editing_panel()

    def bind_currently_editing_panel_shortcuts(self, widget):
        return

    def handle_global_ctrl_shift_shortcuts(self, event):
        state = int(getattr(event, 'state', 0) or 0)
        keysym = str(getattr(event, 'keysym', '') or '')
        if not keysym:
            return None
        if not ((state & 0x4) and (state & 0x1)):
            return None

        key = keysym.lower()
        if key not in {'c', 'x'}:
            return None

        now = time.monotonic()
        last_key = getattr(self, '_last_ctrl_shift_shortcut_key', None)
        last_time = getattr(self, '_last_ctrl_shift_shortcut_time', 0.0)
        if last_key == key and (now - last_time) < 0.25:
            return "break"

        self._last_ctrl_shift_shortcut_key = key
        self._last_ctrl_shift_shortcut_time = now

        if key == 'c':
            return self.toggle_currently_editing_panel(event)
        if key == 'x':
            return self.ctrl_shift_x(event)
        return None

    def refocus_after_currently_editing_click(self, event=None):
        self.root.after_idle(self.focus_last_active_editor)
        return "break"

    def show_find_panel(self):
        current_doc = self.get_current_doc()
        if current_doc and current_doc.get('virtual_mode'):
            messagebox.showinfo("Large File Mode", "Find is not available in buffered large-file mode yet.", parent=self.root)
            return "break"

        self.cancel_find_change_job()
        if self.replace_panel_visible:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.clear_find_highlights()

        if not self.find_panel_visible:
            self.bottom_frame.grid()
            self.find_frame.grid(sticky='ew')
            self.find_panel_visible = True
            self.find_entry.focus_set()
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

        self.cancel_find_change_job()
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

    def safe_focus_get(self):
        try:
            return self.root.focus_get()
        except (tk.TclError, KeyError):
            return None

    def get_active_search_widget(self):
        valid_widgets = [doc['text'] for doc in self.documents.values() if doc.get('text')]
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None:
            valid_widgets.append(compare_widget)

        focus_widget = self.safe_focus_get()
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
            return
        try:
            if self.safe_focus_get() != widget:
                widget.focus_force()
        except tk.TclError:
            pass

    def goto_document_start(self, event=None):
        targets = self.get_page_navigation_targets()
        if not targets:
            return "break"
        for widget in targets:
            try:
                widget.mark_set(tk.INSERT, '1.0')
                widget.tag_remove('sel', '1.0', tk.END)
                widget.see('1.0')
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            doc = self.get_doc_for_text_widget(widget)
            if widget == self.get_compare_text_widget():
                self.update_line_number_gutter(self.compare_view)
            elif doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
        self.update_status()
        return "break"

    def goto_document_end(self, event=None):
        targets = self.get_page_navigation_targets()
        if not targets:
            return "break"
        for widget in targets:
            try:
                last_line = widget.index('end-1c').split('.')[0]
                target_index = f'{last_line}.0'
                widget.mark_set(tk.INSERT, target_index)
                widget.tag_remove('sel', '1.0', tk.END)
                widget.see(target_index)
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            doc = self.get_doc_for_text_widget(widget)
            if widget == self.get_compare_text_widget():
                self.update_line_number_gutter(self.compare_view)
            elif doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
        self.update_status()
        return "break"

    def remember_compare_focus(self, event=None):
        self.hide_autocomplete_popup()
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None:
            self.root.after_idle(lambda widget=compare_widget: self.set_last_active_editor_widget(widget))
        self.update_status()
        return None

    def remember_hovered_editor(self, event=None):
        widget = getattr(event, 'widget', None)
        if isinstance(widget, tk.Text):
            self.hovered_editor_widget = widget
        return None

    def get_hovered_editor_widget(self):
        widget = getattr(self, 'hovered_editor_widget', None)
        if not isinstance(widget, tk.Text):
            return None
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and widget == compare_widget:
            return widget
        for doc in self.documents.values():
            if doc.get('text') == widget:
                return widget
        return None

    def get_page_navigation_targets(self):
        if self.compare_active and self.sync_page_navigation_enabled.get():
            targets = []
            if isinstance(self.text, tk.Text):
                targets.append(self.text)
            compare_widget = self.get_compare_text_widget()
            if compare_widget is not None:
                targets.append(compare_widget)
            return targets

        hovered_widget = self.get_hovered_editor_widget()
        if hovered_widget is not None:
            return [hovered_widget]

        active_widget = self.get_active_search_widget()
        if active_widget is not None:
            return [active_widget]
        return []

    def sync_widget_insert_to_visible_line(self, widget):
        try:
            visible_index = widget.index("@0,0")
            visible_line = visible_index.split('.')[0]
            widget.mark_set(tk.INSERT, f"{visible_line}.0")
        except tk.TclError:
            pass

    def page_up(self, event=None):
        targets = self.get_page_navigation_targets()
        if not targets:
            return
        for widget in targets:
            try:
                widget.yview_scroll(-1, 'page')
                self.sync_widget_insert_to_visible_line(widget)
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            doc = self.get_doc_for_text_widget(widget)
            if widget == self.get_compare_text_widget():
                self.update_line_number_gutter(self.compare_view)
            elif doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
        self.update_status()
        return "break"

    def page_down(self, event=None):
        targets = self.get_page_navigation_targets()
        if not targets:
            return
        for widget in targets:
            try:
                widget.yview_scroll(1, 'page')
                self.sync_widget_insert_to_visible_line(widget)
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            doc = self.get_doc_for_text_widget(widget)
            if widget == self.get_compare_text_widget():
                self.update_line_number_gutter(self.compare_view)
            elif doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
        self.update_status()
        return "break"

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
        self.raise_find_tags(widget)
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
        start_offset = self.get_text_char_offset(widget, start)
        try:
            content = widget.get('1.0', 'end-1c')
        except tk.TclError:
            return None
        matches = self.find_query_offsets(widget, query, start_offset=start_offset, max_matches=1, nocase=True)
        if not matches and wrap:
            matches = self.find_query_offsets(widget, query, start_offset=0, stop_offset=start_offset, max_matches=1, nocase=True)
        pos = self.text_index_from_offset(widget, matches[0], content=content) if matches else None
        if not pos:
            return None

        end = self.text_index_from_offset(widget, matches[0] + len(query), content=content)
        if not end:
            return None
        self.set_current_find_match(widget, pos, end)
        self.update_status()
        return pos, end

    def search_previous_in_widget(self, widget, query, start_index=None, wrap=True):
        if widget is None or not query:
            return None

        start = start_index or widget.index(tk.INSERT)
        start_offset = self.get_text_char_offset(widget, start)
        try:
            content = widget.get('1.0', 'end-1c')
        except tk.TclError:
            return None

        previous_limit = max(0, start_offset - len(query))
        matches = self.find_query_offsets(widget, query, start_offset=0, stop_offset=previous_limit, nocase=True)
        if matches:
            match_offset = matches[-1]
        elif wrap:
            wrapped_matches = self.find_query_offsets(widget, query, start_offset=0, stop_offset=len(content), nocase=True)
            if not wrapped_matches:
                return None
            match_offset = wrapped_matches[-1]
        else:
            return None

        pos = self.text_index_from_offset(widget, match_offset, content=content)
        if not pos:
            return None
        end = self.text_index_from_offset(widget, match_offset + len(query), content=content)
        if not end:
            return None
        self.set_current_find_match(widget, pos, end)
        self.update_status()
        return pos, end

    def refresh_compare_find_highlights(self, query, max_matches=None, allow_short_query=False):
        if self.search_all_tabs.get() or not self.compare_active or not query:
            return False
        if not allow_short_query and len(query) < self.live_find_min_chars:
            return False

        widgets = []
        seen = set()
        for widget in (self.text, self.get_compare_text_widget()):
            if not widget:
                continue
            try:
                if not widget.winfo_exists():
                    continue
            except tk.TclError:
                continue
            widget_id = str(widget)
            if widget_id in seen:
                continue
            seen.add(widget_id)
            widgets.append(widget)

        if len(widgets) < 2:
            return False

        self.clear_find_highlights()
        for widget in widgets:
            self.highlight_matches_in_widget(
                widget,
                query,
                max_matches=max_matches,
                allow_short_query=allow_short_query
            )
        return True

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

        self.refresh_compare_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_next_in_widget(target_widget, query, wrap=True)
            return "break"

        if self.search_all_tabs.get():
            return self.find_next_across_tabs(query, target_widget)

        self.search_next_in_widget(target_widget, query, wrap=True)
        return "break"

    def find_previous(self, event=None):
        if self.find_panel_visible:
            query = self.find_entry.get().strip()
        elif self.replace_panel_visible:
            query = self.replace_find_entry.get().strip()
        else:
            return "break"

        if not query:
            return "break"

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        self.refresh_compare_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_previous_in_widget(target_widget, query, wrap=True)
            return "break"

        if self.search_all_tabs.get():
            return self.find_previous_across_tabs(query, target_widget)

        self.search_previous_in_widget(target_widget, query, wrap=True)
        return "break"

    def find_next_across_tabs(self, query, start_widget=None, start_from_top=False):
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
            start = '1.0' if start_from_top else (search_widget.index(tk.INSERT) if first_pass and search_widget == start_widget else '1.0')
            start_offset = self.get_text_char_offset(search_widget, start)
            try:
                content = search_widget.get('1.0', 'end-1c')
            except tk.TclError:
                first_pass = False
                continue
            matches = self.find_query_offsets(search_widget, query, start_offset=start_offset, max_matches=1, nocase=True)
            if matches:
                pos = self.text_index_from_offset(search_widget, matches[0], content=content)
                if not pos:
                    first_pass = False
                    continue
                end = self.text_index_from_offset(search_widget, matches[0] + len(query), content=content)
                if not end:
                    first_pass = False
                    continue
                self.set_current_find_match(search_widget, pos, end)
                return "break"
            first_pass = False
        return "break"

    def find_previous_across_tabs(self, query, start_widget=None):
        tab_ids = list(self.notebook.tabs())
        if not tab_ids:
            return "break"

        start_doc = self.get_doc_for_text_widget(start_widget) if start_widget is not None else self.get_current_doc()
        current_tab = str(start_doc['frame']) if start_doc else self.notebook.select()
        try:
            start_index = tab_ids.index(current_tab)
        except ValueError:
            start_index = 0
        search_order = list(reversed(tab_ids[:start_index + 1])) + list(reversed(tab_ids[start_index + 1:]))
        first_pass = True
        for tab_id in search_order:
            doc = self.documents.get(str(tab_id))
            if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
                continue
            self.notebook.select(doc['frame'])
            self.set_active_document(doc['frame'])
            search_widget = doc['text']
            start = search_widget.index(tk.INSERT) if first_pass and search_widget == start_widget else tk.END
            result = self.search_previous_in_widget(search_widget, query, start_index=start, wrap=not first_pass)
            if result:
                return "break"
            first_pass = False
        return "break"

    def find_from_input(self, event=None):
        if self.find_panel_visible:
            query = self.find_entry.get().strip()
        elif self.replace_panel_visible:
            query = self.replace_find_entry.get().strip()
        else:
            return "break"

        if not query:
            return "break"

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        self.refresh_compare_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_next_in_widget(target_widget, query, start_index='1.0', wrap=True)
            return "break"

        if self.search_all_tabs.get():
            return self.find_next_across_tabs(query, target_widget, start_from_top=True)

        self.search_next_in_widget(target_widget, query, start_index='1.0', wrap=True)
        return "break"

    def goto_next_unread_note(self):
        try:
            doc = self.get_current_doc()
            if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
                return "break"

            text_widget = doc.get('text')
            if not text_widget:
                return "break"

            unread_tags = self.get_unread_note_tags(doc)
            if not unread_tags:
                return "break"

            note_tag = unread_tags[0]
            ranges = text_widget.tag_ranges(note_tag)
            if len(ranges) < 2:
                return "break"

            start = str(ranges[0])
            end = str(ranges[1])
            text_widget.tag_remove('sel', '1.0', tk.END)
            text_widget.tag_add('sel', start, end)
            text_widget.mark_set(tk.INSERT, end)
            text_widget.see(start)
            self.set_last_active_editor_widget(text_widget)
            self.mark_note_as_read(doc, note_tag)
            note_data = doc['notes'].get(note_tag)
            if note_data:
                bbox = text_widget.bbox(start)
                if bbox:
                    x, y, width, height = bbox
                    self.show_note_popup(
                        doc,
                        note_data,
                        text_widget.winfo_rootx() + x + width,
                        text_widget.winfo_rooty() + y + height
                    )
            self.update_status()
        except Exception as exc:
            self.log_exception("goto next unread note", exc)
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
        if filter_mode in self.note_colors:
            return self.normalize_note_color(note_data.get('color')) == filter_mode
        return True

    def goto_next_note(self, event=None):
        try:
            doc = self.get_current_doc()
            if not doc or doc.get('virtual_mode') or doc.get('preview_mode'):
                return "break"

            text_widget = doc.get('text')
            if not text_widget:
                return "break"

            ordered_tags = self.get_ordered_note_tags(doc)
            if not ordered_tags:
                doc['last_note_cycle_tag'] = None
                return "break"

            last_tag = doc.get('last_note_cycle_tag')
            if last_tag in ordered_tags:
                note_tag = ordered_tags[(ordered_tags.index(last_tag) + 1) % len(ordered_tags)]
            else:
                note_tag = ordered_tags[0]

            ranges = text_widget.tag_ranges(note_tag)
            if len(ranges) < 2:
                doc['last_note_cycle_tag'] = None
                return "break"

            start = str(ranges[0])
            end = str(ranges[1])
            text_widget.tag_remove('sel', '1.0', tk.END)
            text_widget.tag_add('sel', start, end)
            text_widget.mark_set(tk.INSERT, end)
            text_widget.see(start)
            self.set_last_active_editor_widget(text_widget)
            doc['last_note_cycle_tag'] = note_tag
            self.mark_note_as_read(doc, note_tag)
            note_data = doc['notes'].get(note_tag)
            if note_data:
                bbox = text_widget.bbox(start)
                if bbox:
                    x, y, width, height = bbox
                    self.show_note_popup(
                        doc,
                        note_data,
                        text_widget.winfo_rootx() + x + width,
                        text_widget.winfo_rooty() + y + height
                    )
            self.update_status()
        except Exception as exc:
            self.log_exception("goto next note", exc)
        return "break"

    def get_find_target_widgets(self):
        widgets = []
        seen = set()

        def add_widget(widget):
            if not widget:
                return
            try:
                if not widget.winfo_exists():
                    return
            except tk.TclError:
                return
            widget_id = str(widget)
            if widget_id in seen:
                return
            seen.add(widget_id)
            widgets.append(widget)

        if self.search_all_tabs.get():
            for doc in self.documents.values():
                if doc.get('virtual_mode') or doc.get('preview_mode'):
                    continue
                add_widget(doc.get('text'))
        elif self.text:
            add_widget(self.text)

        if self.compare_active and self.compare_view and self.compare_view.get('text'):
            add_widget(self.compare_view.get('text'))
        return widgets

    def clear_find_highlights(self):
        widgets = []
        seen = set()
        for doc in self.documents.values():
            widget = doc.get('text')
            if not widget:
                continue
            widget_id = str(widget)
            if widget_id in seen:
                continue
            seen.add(widget_id)
            widgets.append(widget)
        if self.compare_view and self.compare_view.get('text'):
            widget = self.compare_view.get('text')
            if widget and str(widget) not in seen:
                widgets.append(widget)
        for widget in widgets:
            try:
                if not widget.winfo_exists():
                    continue
                widget.tag_remove(self.find_matches_tag, '1.0', tk.END)
                widget.tag_remove(self.find_current_tag, '1.0', tk.END)
            except tk.TclError:
                continue

    def highlight_matches_in_widget(self, text_widget, query, max_matches=None, allow_short_query=False):
        try:
            if not text_widget or not text_widget.winfo_exists():
                return
            self.raise_find_tags(text_widget)
            text_widget.tag_remove(self.find_matches_tag, '1.0', tk.END)
            text_widget.tag_remove(self.find_current_tag, '1.0', tk.END)
        except tk.TclError:
            return
        if not query or (not allow_short_query and len(query) < self.live_find_min_chars):
            return
        max_allowed_matches = self.live_find_max_matches_per_widget if max_matches is None else max_matches
        match_offsets = self.find_query_offsets(
            text_widget,
            query,
            start_offset=0,
            max_matches=max_allowed_matches,
            nocase=True
        )
        try:
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return
        for offset in match_offsets:
            try:
                pos = self.text_index_from_offset(text_widget, offset, content=content)
                if not pos:
                    continue
                end = self.text_index_from_offset(text_widget, offset + len(query), content=content)
                if not end:
                    continue
                text_widget.tag_add(self.find_matches_tag, pos, end)
            except tk.TclError:
                break

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

    def get_text_char_offset(self, text_widget, index):
        try:
            return max(0, int(text_widget.count('1.0', index, 'chars')[0]))
        except (tk.TclError, TypeError, ValueError, IndexError):
            return 0

    def text_index_from_offset(self, text_widget, offset, content=None):
        try:
            safe_offset = max(0, int(offset))
        except (TypeError, ValueError):
            safe_offset = 0
        try:
            if content is None:
                content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return None

        content_length = len(content)
        safe_offset = min(safe_offset, content_length)
        line = content.count('\n', 0, safe_offset) + 1
        last_newline = content.rfind('\n', 0, safe_offset)
        column = safe_offset if last_newline < 0 else safe_offset - last_newline - 1
        return f"{line}.{column}"

    def find_query_offsets(self, text_widget, query, start_offset=0, stop_offset=None, max_matches=None, nocase=True):
        try:
            if not text_widget or not text_widget.winfo_exists() or not query:
                return []
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return []

        content_length = len(content)
        try:
            search_start = max(0, min(int(start_offset), content_length))
        except (TypeError, ValueError):
            search_start = 0
        if stop_offset is None:
            search_stop = content_length
        else:
            try:
                search_stop = max(search_start, min(int(stop_offset), content_length))
            except (TypeError, ValueError):
                search_stop = content_length

        haystack = content.casefold() if nocase else content
        needle = query.casefold() if nocase else query
        if not needle:
            return []

        offsets = []
        cursor = search_start
        step = max(1, len(needle))
        while cursor <= search_stop:
            position = haystack.find(needle, cursor, search_stop)
            if position < 0:
                break
            offsets.append(position)
            if max_matches and len(offsets) >= max_matches:
                break
            cursor = position + step
        return offsets

    def highlight_all_matches(self, query, max_matches=None, allow_short_query=False):
        widgets = self.get_find_target_widgets()
        for widget in widgets:
            self.highlight_matches_in_widget(
                widget,
                query,
                max_matches=max_matches,
                allow_short_query=allow_short_query
            )

    def raise_find_tags(self, text_widget):
        try:
            if not text_widget or not text_widget.winfo_exists():
                return
            text_widget.tag_raise(self.find_matches_tag)
            text_widget.tag_raise(self.find_current_tag, self.find_matches_tag)
        except tk.TclError:
            pass

    def highlight_live_find_matches(self, query):
        if not query:
            self.clear_find_highlights()
            return
        if len(query) < self.live_find_min_chars:
            return
        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return
        if self.refresh_compare_find_highlights(
            query,
            max_matches=self.live_find_max_matches_typing
        ):
            return
        self.clear_find_highlights()
        self.highlight_matches_in_widget(
            target_widget,
            query,
            max_matches=self.live_find_max_matches_typing
        )

    def cancel_find_change_job(self):
        if self.find_change_job:
            try:
                self.root.after_cancel(self.find_change_job)
            except tk.TclError:
                pass
            self.find_change_job = None

    def apply_live_find_change(self):
        self.find_change_job = None
        try:
            if self.find_panel_visible:
                if not self.find_entry.winfo_exists():
                    return
                query = self.find_entry.get().strip()
            elif self.replace_panel_visible:
                if not self.replace_find_entry.winfo_exists():
                    return
                query = self.replace_find_entry.get().strip()
            else:
                return

            self.highlight_live_find_matches(query)
        except Exception as exc:
            self.log_exception("live find change", exc)

    def on_find_entry_change(self, event=None):
        try:
            if self.find_panel_visible:
                query = self.find_entry.get().strip()
            elif self.replace_panel_visible:
                query = self.replace_find_entry.get().strip()
            else:
                return
        except tk.TclError as exc:
            self.log_exception("read live find query", exc)
            return

        self.cancel_find_change_job()
        if query and len(query) < self.live_find_min_chars:
            return
        try:
            self.find_change_job = self.root.after(30, self.apply_live_find_change)
        except tk.TclError as exc:
            self.log_exception("schedule live find change", exc)

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
        self.root.bind('<Control-Shift-e>', lambda e: self.save_encrypted_copy())
        self.root.bind('<Control-Shift-E>', lambda e: self.save_encrypted_copy())
        self.root.bind('<Control-Shift-r>', self.save_and_run)
        self.root.bind('<Control-Shift-R>', self.save_and_run)
        self.root.bind_all('<KeyPress>', self.handle_global_ctrl_shift_shortcuts, add='+')
        self.root.bind_all('<Control-Shift-x>', self.ctrl_shift_x)
        self.root.bind_all('<Control-Shift-X>', self.ctrl_shift_x)

        # Edit
        self.root.bind('<Control-z>', self.undo)
        self.root.bind('<Control-Z>', self.undo)
        self.root.bind('<Control-Shift-Z>', self.redo)
        self.root.bind('<Control-Shift-z>', self.redo)
        self.root.bind('<Control-a>', self.select_all)
        self.root.bind('<Control-A>', self.select_all)
        self.root.bind_all('<Control-b>', self.toggle_status_bar)
        self.root.bind_all('<Control-B>', self.toggle_status_bar)
        self.root.bind_all('<ButtonRelease-1>', self.maybe_dismiss_transient_ui, add='+')
        self.root.bind_all('<ButtonRelease-3>', self.maybe_dismiss_transient_ui, add='+')
        self.root.bind('<Control-d>', self.insert_date)
        self.root.bind('<Control-D>', self.insert_date)
        self.root.bind('<Control-Shift-D>', self.insert_time_date)
        self.root.bind('<Control-Shift-d>', self.insert_time_date)
        self.root.bind('<Control-e>', lambda e: self.export_notes_report())
        self.root.bind('<Control-E>', lambda e: self.export_notes_report())
        self.root.bind('<Control-Shift-T>', self.close_current_tab)
        self.root.bind('<Control-Shift-t>', self.close_current_tab)
        self.root.bind('<Control-Tab>', self.switch_tab_right)
        self.root.bind('<Control-Shift-F>', self.show_font_dialog)
        self.root.bind('<Control-Shift-f>', self.show_font_dialog)
        # Search / Navigation
        self.root.bind('<Control-f>', lambda e: self.show_find_panel())
        self.root.bind('<Control-F>', lambda e: self.show_find_panel())
        self.root.bind('<Control-r>', lambda e: self.show_replace_panel())
        self.root.bind('<Control-R>', lambda e: self.show_replace_panel())
        self.root.bind('<F3>', self.find_next)
        self.root.bind('<Shift-F3>', self.find_previous)
        self.root.bind('<F4>', self.goto_next_note)
        self.root.bind('<Control-g>', self.goto_line_dialog)
        self.root.bind('<Control-G>', self.goto_line_dialog)
        self.root.bind_all('<Prior>', self.page_up)
        self.root.bind_all('<Next>', self.page_down)
        self.root.bind_all('<Control-Prior>', self.goto_document_start)
        self.root.bind_all('<Control-Next>', self.goto_document_end)

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
        if self.close_help_or_about_window():
            return "break"
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

    def close_help_or_about_window(self):
        for widget in reversed(self.root.winfo_children()):
            if not isinstance(widget, tk.Toplevel):
                continue
            try:
                title = widget.title()
            except tk.TclError:
                continue
            if title in {
                self.tr('app.help_title', 'Notepad-X Help'),
                self.tr('app.about_title', 'About Notepad-X')
            }:
                try:
                    widget.destroy()
                    return True
                except tk.TclError:
                    continue
        return False

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

    def toggle_sound(self):
        self.save_session()

    def toggle_autocomplete(self):
        if not self.autocomplete_enabled.get():
            self.hide_autocomplete_popup()
        self.save_session()
        return "break"

    def get_edit_with_shell_extensions(self):
        extensions = []
        seen = set()
        for label, pattern in self.get_save_filetypes():
            if label in ('All Supported', 'All Files'):
                continue
            for token in str(pattern).split():
                token = token.strip().lower()
                if token.startswith('*.') and len(token) > 2:
                    extension = f".{token[2:]}"
                elif token.startswith('.') and len(token) > 1:
                    extension = token
                else:
                    continue
                if extension not in seen:
                    seen.add(extension)
                    extensions.append(extension)
        return extensions

    def get_windows_open_command(self):
        if getattr(sys, 'frozen', False):
            executable_path = os.path.abspath(sys.executable)
            if not self.path_looks_safe_for_shell(executable_path):
                raise OSError("Unsafe executable path for Windows shell integration.")
            return f'"{executable_path}" "%1"'
        interpreter_path = os.path.abspath(sys.executable)
        script_path = os.path.abspath(__file__)
        if not self.path_looks_safe_for_shell(interpreter_path) or not self.path_looks_safe_for_shell(script_path):
            raise OSError("Unsafe interpreter or script path for Windows shell integration.")
        return f'"{interpreter_path}" "{script_path}" "%1"'

    def get_linux_open_command(self):
        if getattr(sys, 'frozen', False):
            executable_path = os.path.abspath(sys.executable)
            if not self.path_looks_safe_for_shell(executable_path):
                raise OSError("Unsafe executable path for Linux desktop integration.")
            return f'"{executable_path}" %F'
        interpreter_path = os.path.abspath(sys.executable)
        script_path = os.path.abspath(__file__)
        if not self.path_looks_safe_for_shell(interpreter_path) or not self.path_looks_safe_for_shell(script_path):
            raise OSError("Unsafe interpreter or script path for Linux desktop integration.")
        return f'"{interpreter_path}" "{script_path}" %F'

    def get_linux_mime_types(self):
        return [
            'text/plain',
            'text/markdown',
            'application/json',
            'text/x-python',
            'text/x-csrc',
            'text/x-chdr',
            'text/x-c++src',
            'text/x-c++hdr',
            'text/x-java',
            'application/javascript',
            'text/javascript',
            'text/html',
            'application/xhtml+xml',
            'text/css',
            'application/xml',
            'text/xml',
            'application/x-shellscript',
            'text/x-script.python',
            'text/x-script.sh',
            'text/x-php',
            'text/x-sql',
            'text/x-diff',
            'text/x-patch',
            'text/x-tex',
            'text/x-csharp',
        ]

    def set_edit_with_shell_linux(self, enabled):
        desktop_entry_path = self.get_linux_desktop_entry_path()
        applications_dir = os.path.dirname(desktop_entry_path)
        if enabled:
            os.makedirs(applications_dir, exist_ok=True)
            icon_path = self.splash_path if os.path.exists(self.splash_path) else self.icon_path
            desktop_entry = [
                '[Desktop Entry]',
                'Type=Application',
                'Version=1.0',
                'Name=Notepad-X',
                'GenericName=Text Editor',
                'Comment=Edit text files with Notepad-X',
                f'Exec={self.get_linux_open_command()}',
                'Terminal=false',
                'Categories=Utility;TextEditor;',
                f'Icon={icon_path}',
                f'MimeType={";".join(self.get_linux_mime_types())};',
                'Actions=EditWithNotepadX;',
                '',
                '[Desktop Action EditWithNotepadX]',
                'Name=Edit with Notepad-X',
                f'Exec={self.get_linux_open_command()}',
                'Terminal=false',
                '',
            ]
            if not self.write_file_atomically(desktop_entry_path, '\n'.join(desktop_entry).rstrip('\n')):
                raise OSError(f"Could not write Linux desktop entry to {desktop_entry_path}")
            update_db = shutil.which('update-desktop-database')
            if update_db:
                try:
                    subprocess.run([update_db, applications_dir], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                except Exception:
                    pass
        else:
            if os.path.exists(desktop_entry_path):
                os.remove(desktop_entry_path)
            update_db = shutil.which('update-desktop-database')
            if update_db:
                try:
                    subprocess.run([update_db, applications_dir], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                except Exception:
                    pass

    def delete_registry_tree(self, root, subkey_path):
        try:
            with winreg.OpenKey(root, subkey_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                while True:
                    try:
                        child_name = winreg.EnumKey(key, 0)
                    except OSError:
                        break
                    self.delete_registry_tree(root, rf"{subkey_path}\{child_name}")
            winreg.DeleteKey(root, subkey_path)
        except FileNotFoundError:
            pass

    def get_windows_application_registration_name(self):
        if getattr(sys, 'frozen', False):
            executable_name = os.path.basename(os.path.abspath(sys.executable))
        else:
            executable_name = 'Notepad-X.exe'
        executable_name = os.path.basename(str(executable_name).strip()) or 'Notepad-X.exe'
        if not executable_name.lower().endswith('.exe'):
            executable_name += '.exe'
        return executable_name

    def set_edit_with_shell_app_registration(self, enabled):
        if not self.is_windows or winreg is None:
            return
        app_key = rf"Software\Classes\Applications\{self.get_windows_application_registration_name()}"
        if enabled:
            icon_source = os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else self.icon_path)
            if not self.path_looks_safe_for_shell(icon_source):
                raise OSError("Unsafe icon path for Windows application registration.")
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, app_key) as key:
                winreg.SetValueEx(key, 'ApplicationName', 0, winreg.REG_SZ, 'Notepad-X')
                winreg.SetValueEx(key, 'FriendlyAppName', 0, winreg.REG_SZ, 'Notepad-X')
                winreg.SetValueEx(
                    key,
                    'ApplicationDescription',
                    0,
                    winreg.REG_SZ,
                    'Edit supported text and code files with Notepad-X.'
                )
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{app_key}\DefaultIcon") as icon_key:
                winreg.SetValueEx(icon_key, '', 0, winreg.REG_SZ, icon_source)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{app_key}\shell\open\command") as command_key:
                winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, self.get_windows_open_command())
            supported_types_key = rf"{app_key}\SupportedTypes"
            self.delete_registry_tree(winreg.HKEY_CURRENT_USER, supported_types_key)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, supported_types_key) as types_key:
                for extension in self.get_edit_with_shell_extensions():
                    winreg.SetValueEx(types_key, extension, 0, winreg.REG_SZ, '')
        else:
            self.delete_registry_tree(winreg.HKEY_CURRENT_USER, app_key)

    def set_edit_with_shell_for_extension(self, extension, enabled):
        if not self.is_windows or winreg is None or not extension:
            return
        menu_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\EditWithNotepadX"
        if enabled:
            icon_source = os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else self.icon_path)
            if not self.path_looks_safe_for_shell(icon_source):
                raise OSError("Unsafe icon path for Windows shell integration.")
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_key) as key:
                winreg.SetValueEx(key, 'MUIVerb', 0, winreg.REG_SZ, 'Edit with Notepad-X')
                winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, icon_source)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{menu_key}\command") as command_key:
                winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, self.get_windows_open_command())
        else:
            self.delete_registry_tree(winreg.HKEY_CURRENT_USER, menu_key)

    def is_edit_with_shell_registered(self):
        if self.is_linux:
            return os.path.exists(self.get_linux_desktop_entry_path())
        if not self.is_windows or winreg is None:
            return bool(self.edit_with_shell_enabled.get())
        app_key = rf"Software\Classes\Applications\{self.get_windows_application_registration_name()}"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, app_key, 0, winreg.KEY_READ):
                return True
        except FileNotFoundError:
            pass
        except OSError:
            pass
        for extension in self.get_edit_with_shell_extensions():
            menu_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\EditWithNotepadX"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, menu_key, 0, winreg.KEY_READ):
                    return True
            except FileNotFoundError:
                continue
            except OSError:
                continue
        return False

    def notify_windows_shell_change(self):
        if not self.is_windows or not self.shell32:
            return
        try:
            self.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        except Exception:
            pass

    def sync_edit_with_shell_menu(self, show_errors=False):
        try:
            enabled = bool(self.edit_with_shell_enabled.get())
            if self.is_windows:
                if winreg is None:
                    return True
                self.set_edit_with_shell_app_registration(enabled)
                for extension in self.get_edit_with_shell_extensions():
                    self.set_edit_with_shell_for_extension(extension, enabled)
                self.notify_windows_shell_change()
            elif self.is_linux:
                self.set_edit_with_shell_linux(enabled)
            return True
        except OSError as exc:
            self.log_exception("sync edit with shell menu", exc)
            if show_errors:
                messagebox.showerror(
                    "Edit with Notepad-X",
                    "Notepad-X could not update the OS shell integration.\n\n"
                    f"{exc}",
                    parent=self.root
                )
            return False

    def toggle_edit_with_shell(self):
        current_value = bool(self.edit_with_shell_enabled.get())
        previous_value = not current_value
        if not self.sync_edit_with_shell_menu(show_errors=True):
            self.edit_with_shell_enabled.set(previous_value)
            return "break"
        self.save_session()
        return "break"

    def open_startup_files(self, startup_files):
        normalized_files = []

        for raw_path in startup_files:
            if not raw_path:
                continue
            candidate_path = os.path.abspath(raw_path)
            if os.path.exists(candidate_path) and not self.is_notepadx_support_file(candidate_path):
                normalized_files.append(candidate_path)

        opened_frames = []
        for file_path in normalized_files:
            if self.open_file_path(file_path):
                current_doc = self.get_current_doc()
                if current_doc:
                    opened_frames.append(current_doc['frame'])

        if opened_frames:
            self.notebook.select(opened_frames[0])
            self.set_active_document(opened_frames[0])
            if self.text:
                self.text.focus_set()

    def toggle_status_bar(self, event=None):
        if event is not None:
            self.status_bar_enabled.set(not self.status_bar_enabled.get())
        if self.status_bar_enabled.get():
            self.status_frame.grid()
        else:
            self.status_frame.grid_remove()
        self.save_session()
        return "break"

    def hide_autocomplete_popup(self):
        popup = getattr(self, 'autocomplete_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.autocomplete_popup = None
        self.autocomplete_listbox = None
        self.autocomplete_doc_id = None
        self.autocomplete_start_index = None
        self.autocomplete_prefix = ""

    def autocomplete_popup_visible(self):
        popup = getattr(self, 'autocomplete_popup', None)
        if popup is None:
            return False
        try:
            return popup.winfo_exists()
        except tk.TclError:
            return False

    def get_autocomplete_keywords(self, syntax_mode):
        keyword_sets = {
            'python': ['and', 'as', 'assert', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else', 'except', 'False', 'finally', 'for', 'from', 'if', 'import', 'in', 'is', 'lambda', 'None', 'not', 'or', 'pass', 'raise', 'return', 'self', 'True', 'try', 'while', 'with', 'yield'],
            'c': ['auto', 'break', 'case', 'char', 'const', 'continue', 'default', 'do', 'double', 'else', 'enum', 'extern', 'float', 'for', 'goto', 'if', 'inline', 'int', 'long', 'register', 'return', 'short', 'signed', 'sizeof', 'static', 'struct', 'switch', 'typedef', 'union', 'unsigned', 'void', 'volatile', 'while'],
            'cpp': ['bool', 'break', 'case', 'catch', 'class', 'const', 'constexpr', 'continue', 'default', 'delete', 'do', 'double', 'else', 'enum', 'false', 'float', 'for', 'if', 'inline', 'int', 'namespace', 'new', 'nullptr', 'private', 'protected', 'public', 'return', 'static', 'std', 'struct', 'switch', 'template', 'this', 'throw', 'true', 'try', 'using', 'virtual', 'void', 'while'],
            'rust': ['as', 'break', 'const', 'continue', 'crate', 'else', 'enum', 'false', 'fn', 'for', 'if', 'impl', 'in', 'let', 'loop', 'match', 'mod', 'move', 'mut', 'pub', 'ref', 'return', 'self', 'Self', 'static', 'struct', 'trait', 'true', 'type', 'unsafe', 'use', 'where', 'while'],
            'java': ['abstract', 'boolean', 'break', 'byte', 'case', 'catch', 'class', 'continue', 'default', 'double', 'else', 'extends', 'false', 'final', 'finally', 'float', 'for', 'if', 'implements', 'import', 'int', 'interface', 'long', 'new', 'null', 'package', 'private', 'protected', 'public', 'return', 'short', 'static', 'super', 'switch', 'this', 'throw', 'true', 'try', 'void', 'while'],
            'javascript': ['async', 'await', 'break', 'case', 'catch', 'class', 'const', 'continue', 'default', 'else', 'export', 'extends', 'false', 'finally', 'for', 'function', 'if', 'import', 'let', 'new', 'null', 'return', 'super', 'switch', 'this', 'throw', 'true', 'try', 'undefined', 'var', 'while'],
            'html': ['body', 'div', 'footer', 'form', 'head', 'header', 'html', 'img', 'input', 'label', 'link', 'main', 'meta', 'nav', 'script', 'section', 'span', 'style', 'title'],
            'php': ['class', 'echo', 'else', 'elseif', 'false', 'foreach', 'function', 'if', 'namespace', 'new', 'null', 'private', 'protected', 'public', 'require', 'return', 'static', 'this', 'true', 'use', 'var', 'while'],
            'xml': ['attribute', 'comment', 'element', 'encoding', 'namespace', 'version'],
            'sql': ['alter', 'and', 'as', 'create', 'delete', 'drop', 'from', 'group', 'having', 'insert', 'into', 'join', 'not', 'null', 'or', 'order', 'select', 'set', 'table', 'update', 'values', 'where'],
        }
        return keyword_sets.get(syntax_mode or '', [])

    def text_has_selection(self, text_widget):
        if not text_widget:
            return False
        try:
            ranges = text_widget.tag_ranges('sel')
            return len(ranges) >= 2 and str(ranges[0]) != str(ranges[1])
        except tk.TclError:
            return False

    def update_autocomplete_popup(self, doc):
        if self.autocomplete_suspended > 0:
            self.hide_autocomplete_popup()
            return
        if not self.autocomplete_enabled.get() or not doc:
            self.hide_autocomplete_popup()
            return
        if doc.get('virtual_mode') or doc.get('preview_mode') or doc.get('large_file_mode'):
            self.hide_autocomplete_popup()
            return

        text = doc.get('text')
        if not text or not text.winfo_exists():
            self.hide_autocomplete_popup()
            return
        if self.text_has_selection(text):
            self.hide_autocomplete_popup()
            return

        try:
            line_start = text.index(f"{tk.INSERT} linestart")
            current_line = text.get(line_start, tk.INSERT)
        except tk.TclError:
            self.hide_autocomplete_popup()
            return

        match = re.search(r'([A-Za-z_][A-Za-z0-9_]*)$', current_line)
        if not match:
            self.hide_autocomplete_popup()
            return

        prefix = match.group(1)
        if len(prefix) < 2:
            self.hide_autocomplete_popup()
            return

        syntax_mode = self.get_syntax_mode(doc)
        suggestions = []
        seen = set()

        try:
            content = text.get('1.0', 'end-1c')
        except tk.TclError:
            content = ''

        for word in re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', content):
            if len(word) < len(prefix) or word == prefix or not word.lower().startswith(prefix.lower()):
                continue
            lowered = word.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            suggestions.append(word)
            if len(suggestions) >= 24:
                break

        for word in self.get_autocomplete_keywords(syntax_mode):
            if len(word) < len(prefix) or word == prefix or not word.lower().startswith(prefix.lower()):
                continue
            lowered = word.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            suggestions.append(word)
            if len(suggestions) >= 24:
                break

        if not suggestions:
            self.hide_autocomplete_popup()
            return

        try:
            bbox = text.bbox(tk.INSERT)
        except tk.TclError:
            bbox = None
        if not bbox:
            self.hide_autocomplete_popup()
            return

        x, y, width, height = bbox
        popup = self.autocomplete_popup
        if not self.autocomplete_popup_visible():
            popup = tk.Toplevel(self.root)
            popup.wm_overrideredirect(True)
            popup.configure(bg='#2d2d2d')
            listbox = tk.Listbox(
                popup,
                bg='#161b22',
                fg=self.fg_color,
                selectbackground='#264f78',
                selectforeground='white',
                activestyle='none',
                highlightthickness=1,
                highlightbackground='#30363d',
                relief='flat',
                borderwidth=0,
                font=('Segoe UI', 10),
                width=max(18, min(42, max(len(word) for word in suggestions) + 2)),
                height=min(8, len(suggestions))
            )
            listbox.pack(fill='both', expand=True)
            self.autocomplete_popup = popup
            self.autocomplete_listbox = listbox
        else:
            listbox = self.autocomplete_listbox
            listbox.configure(width=max(18, min(42, max(len(word) for word in suggestions) + 2)), height=min(8, len(suggestions)))

        listbox.delete(0, tk.END)
        for suggestion in suggestions:
            listbox.insert(tk.END, suggestion)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(0)
        listbox.activate(0)

        popup.geometry(f"+{text.winfo_rootx() + x}+{text.winfo_rooty() + y + height + 2}")
        popup.lift()
        self.autocomplete_doc_id = str(doc['frame'])
        self.autocomplete_start_index = f"{line_start}+{match.start(1)}c"
        self.autocomplete_prefix = prefix

    def move_autocomplete_selection(self, direction):
        if not self.autocomplete_popup_visible() or not self.autocomplete_listbox:
            return "break"
        listbox = self.autocomplete_listbox
        selection = listbox.curselection()
        current_index = selection[0] if selection else 0
        next_index = max(0, min(listbox.size() - 1, current_index + direction))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(next_index)
        listbox.activate(next_index)
        listbox.see(next_index)
        return "break"

    def accept_autocomplete_selection(self):
        if not self.autocomplete_popup_visible() or not self.autocomplete_listbox:
            return False
        doc = self.documents.get(self.autocomplete_doc_id)
        if not doc and self.compare_view and self.autocomplete_doc_id == str(self.compare_view['frame']):
            doc = self.compare_view
        if not doc:
            self.hide_autocomplete_popup()
            return False
        text = doc.get('text')
        if not text or not text.winfo_exists():
            self.hide_autocomplete_popup()
            return False
        selection = self.autocomplete_listbox.curselection()
        if not selection:
            return False
        suggestion = self.autocomplete_listbox.get(selection[0])
        try:
            text.delete(self.autocomplete_start_index, tk.INSERT)
            text.insert(self.autocomplete_start_index, suggestion)
            text.edit_modified(True)
        except tk.TclError:
            self.hide_autocomplete_popup()
            return False
        self.hide_autocomplete_popup()
        if doc is self.compare_view:
            self.on_compare_modified()
        else:
            self.on_text_modified(doc['frame'])
        return True

    # ─── Text Widget ─────────────────────────────────────────────
    def create_text_widget(self):
        frame = tk.Frame(self.root, bg=self.bg_color)
        frame.grid(row=0, column=0, sticky='nsew')
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=0)

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
        self.editor_paned.bind('<Configure>', lambda e: self.on_editor_paned_configure(), add='+')

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
            undo=True,
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
        for evt in ('<KeyRelease>', '<MouseWheel>', '<Button-4>', '<Button-5>',
                    '<Button-1>', '<B1-Motion>', '<ButtonRelease-1>', '<Double-Button-1>',
                    '<<Selection>>'):
            self.compare_text.bind(evt, self.handle_compare_activity, add='+')
        self.compare_text.bind('<KeyPress>', self.handle_compare_keypress)
        self.compare_text.bind('<Control-b>', self.toggle_status_bar)
        self.compare_text.bind('<Control-B>', self.toggle_status_bar)
        self.compare_text.bind('<Control-x>', self.cut)
        self.compare_text.bind('<Control-X>', self.cut)
        self.compare_text.bind('<Control-v>', self.paste)
        self.compare_text.bind('<Control-V>', self.paste)
        self.compare_text.bind('<Control-z>', self.undo)
        self.compare_text.bind('<Control-Z>', self.undo)
        self.compare_text.bind('<Control-Shift-z>', self.redo)
        self.compare_text.bind('<Control-Shift-Z>', self.redo)
        self.compare_text.bind('<FocusIn>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<Enter>', self.remember_hovered_editor, add='+')
        self.compare_text.bind('<Motion>', self.remember_hovered_editor, add='+')
        self.compare_text.bind('<Button-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<ButtonRelease-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<Button-3>', self.show_compare_context_menu)
        self.compare_text.bind('<F3>', self.find_next)
        self.compare_text.bind('<Shift-F3>', self.find_previous)
        self.compare_text.bind('<Control-Shift-r>', self.save_and_run)
        self.compare_text.bind('<Control-Shift-R>', self.save_and_run)
        self.compare_text.bind('<Control-Shift-x>', self.ctrl_shift_x)
        self.compare_text.bind('<Control-Shift-X>', self.ctrl_shift_x)
        self.compare_text.bind('<<Modified>>', self.on_compare_modified)

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
            'suspend_modified_events': False,
        }
        self.compare_line_numbers.bind('<Button-1>', lambda e: self.copy_line_from_gutter(e, target_doc=self.compare_view))
        self.compare_text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        self.compare_text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')
        self.compare_text.tag_config(self.bracket_match_tag, background='#2f81f7', foreground='white')
        self.apply_syntax_tag_colors(self.compare_text)
        self.raise_find_tags(self.compare_text)
        self.create_text_context_menu(self.compare_container, doc_override=self.compare_view, action_tab_id='__compare__')

        self.currently_editing_sidebar = tk.Frame(frame, bg=self.panel_bg)
        self.currently_editing_sidebar.grid_rowconfigure(1, weight=1)
        self.currently_editing_sidebar.grid_columnconfigure(0, weight=1)
        self.currently_editing_sidebar.configure(width=self.get_currently_editing_sidebar_width())
        self.currently_editing_sidebar.configure(takefocus=0)

        self.currently_editing_title_label = tk.Label(
            self.currently_editing_sidebar,
            text=self.tr('panel.currently_editing.title', 'Currently Editing'),
            bg=self.panel_bg,
            fg=self.fg_color,
            anchor=self.ui_anchor_start(),
            padx=10,
            pady=8,
            font=('Segoe UI', 10, 'bold'),
            takefocus=0
        )
        self.currently_editing_title_label.grid(row=0, column=0, sticky='ew')
        self.bind_currently_editing_panel_shortcuts(self.currently_editing_title_label)

        self.currently_editing_content_label = tk.Label(
            self.currently_editing_sidebar,
            width=33,
            text='',
            bg=self.text_bg,
            fg=self.text_fg,
            font=('Consolas', 10),
            padx=8,
            pady=6,
            borderwidth=0,
            highlightthickness=0,
            relief='flat',
            takefocus=0,
            cursor='arrow',
            justify='left',
            anchor='nw'
        )
        self.currently_editing_content_label.grid(row=1, column=0, sticky='nsew', padx=8, pady=(0, 8))
        self.bind_currently_editing_panel_shortcuts(self.currently_editing_sidebar)
        self.bind_currently_editing_panel_shortcuts(self.currently_editing_content_label)
        self.currently_editing_sidebar.bind('<Button-1>', self.refocus_after_currently_editing_click, add='+')
        self.currently_editing_title_label.bind('<Button-1>', self.refocus_after_currently_editing_click, add='+')
        self.currently_editing_content_label.bind('<Button-1>', self.refocus_after_currently_editing_click, add='+')
        self.currently_editing_title_label.configure(
            justify=self.ui_justify(),
            wraplength=max(120, self.get_currently_editing_sidebar_width() - 20)
        )
        self.currently_editing_sidebar.grid(row=0, column=1, sticky='ns')
        self.currently_editing_sidebar.grid_remove()

        self.text = None
        self.create_tab()

    def on_editor_paned_configure(self):
        if self.compare_active:
            self.set_compare_sash_position()

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
        text.bind('<Control-x>', self.cut)
        text.bind('<Control-X>', self.cut)
        text.bind('<Control-v>', self.paste)
        text.bind('<Control-V>', self.paste)
        text.bind('<Control-z>', self.undo)
        text.bind('<Control-Z>', self.undo)
        text.bind('<Control-Shift-z>', self.redo)
        text.bind('<Control-Shift-Z>', self.redo)
        text.bind('<Control-Shift-r>', self.save_and_run)
        text.bind('<Control-Shift-R>', self.save_and_run)
        text.bind('<Control-Shift-x>', self.ctrl_shift_x)
        text.bind('<Control-Shift-X>', self.ctrl_shift_x)
        text.bind('<FocusIn>', lambda e, frame=tab_frame: self.remember_doc_focus(frame), add='+')
        text.bind('<Enter>', self.remember_hovered_editor, add='+')
        text.bind('<Motion>', self.remember_hovered_editor, add='+')
        text.bind('<KeyPress>', lambda e, frame=tab_frame: self.handle_text_keypress(e, frame))
        text.bind('<ButtonRelease-1>', lambda e, frame=tab_frame: self.on_text_click_release(e, frame), add='+')
        text.bind('<Button-3>', lambda e, frame=tab_frame: self.show_text_context_menu(e, frame))

        text.bind('<<Modified>>', lambda e, frame=tab_frame: self.on_text_modified(frame))
        text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')
        text.tag_config(self.bracket_match_tag, background='#2f81f7', foreground='white')
        self.raise_find_tags(text)

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
            'note_editors_signature': None,
            'note_active_editors': 0,
            'note_editors': [],
            'note_editor_label': None,
            'note_last_heartbeat_at': 0.0,
            'last_unread_count': 0,
            'notes_registered': False,
            'last_note_cycle_tag': None,
            'syntax_job': None,
            'syntax_mode': None,
            'syntax_override': None,
            'last_insert_index': '1.0',
            'last_yview': 0.0,
            'last_xview': 0.0,
            'file_signature': None,
            'encrypted_file': False,
            'encryption_header': None,
            'encryption_key': None,
            'background_loading': False,
            'background_load_kind': None,
            'background_load_file_path': None,
            'background_open_new_tab': False,
            'pending_insert_content': None,
            'pending_insert_offset': 0,
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

    def remember_doc_view_state(self, doc):
        if not doc:
            return
        text = doc.get('text')
        if not text or not text.winfo_exists():
            return
        try:
            doc['last_insert_index'] = text.index(tk.INSERT)
        except tk.TclError:
            pass
        try:
            doc['last_yview'] = text.yview()[0]
        except (tk.TclError, IndexError):
            pass
        try:
            doc['last_xview'] = text.xview()[0]
        except (tk.TclError, IndexError):
            pass

    def restore_doc_view_state(self, doc):
        if not doc:
            return
        text = doc.get('text')
        if not text or not text.winfo_exists():
            return
        insert_index = doc.get('last_insert_index') or '1.0'
        try:
            text.mark_set(tk.INSERT, insert_index)
        except tk.TclError:
            text.mark_set(tk.INSERT, '1.0')

        if doc.get('virtual_mode'):
            try:
                global_line = doc['window_start_line'] + int(text.index(tk.INSERT).split('.')[0]) - 1
            except (tk.TclError, ValueError):
                global_line = doc.get('window_start_line', 1)
            self.ensure_virtual_line_visible(doc, global_line)
            text = doc.get('text')
            if not text or not text.winfo_exists():
                return

        try:
            text.xview_moveto(float(doc.get('last_xview', 0.0)))
        except (tk.TclError, TypeError, ValueError):
            pass
        try:
            text.yview_moveto(float(doc.get('last_yview', 0.0)))
        except (tk.TclError, TypeError, ValueError):
            pass

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

    def get_syntax_surface_palette(self):
        selected = (getattr(self, 'theme_definitions', None) or {}).get(self.syntax_theme.get())
        if selected is None:
            selected = self.sanitize_theme_definition(
                'Default',
                {'name': 'Default', **self.get_builtin_theme_definitions()['Default']}
            )
        return dict(selected['surface'])

    def get_syntax_palette(self):
        selected = (getattr(self, 'theme_definitions', None) or {}).get(self.syntax_theme.get())
        if selected is None:
            selected = self.sanitize_theme_definition(
                'Default',
                {'name': 'Default', **self.get_builtin_theme_definitions()['Default']}
            )
        return dict(selected['syntax'])

    def apply_syntax_tag_colors(self, text_widget):
        surface = self.get_syntax_surface_palette()
        palette = self.get_syntax_palette()
        text_widget.configure(
            bg=surface['text_bg'],
            fg=surface['text_fg'],
            insertbackground=surface['cursor'],
            selectbackground=surface['selection'],
            selectforeground='white',
        )
        text_widget.tag_config('syntax_keyword', foreground=palette['keyword'])
        text_widget.tag_config('syntax_type', foreground=palette['type'])
        text_widget.tag_config('syntax_string', foreground=palette['string'])
        text_widget.tag_config('syntax_comment', foreground=palette['comment'])
        text_widget.tag_config('syntax_number', foreground=palette['number'])
        text_widget.tag_config('syntax_preprocessor', foreground=palette['preprocessor'])
        text_widget.tag_config('syntax_tag', foreground=palette['tag'])

    def set_syntax_theme(self, theme_name):
        if theme_name not in self.get_available_syntax_theme_names():
            theme_name = 'Default'
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
            self.update_line_number_gutter(doc)
        if self.compare_active:
            self.refresh_compare_panel()
        elif getattr(self, 'compare_view', None):
            self.apply_syntax_tag_colors(self.compare_text)
            self.update_line_number_gutter(self.compare_view)
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
        self.remember_doc_view_state(doc)
        event_type = str(getattr(event, 'type', ''))
        if doc.get('virtual_mode'):
            # Don't rebuffer while the user is actively selecting text.
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
        if self.is_shift_selection_navigation(event):
            self.hide_autocomplete_popup()
        elif event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease'}:
            self.hide_autocomplete_popup()
        elif event_type not in {'35'}:
            self.update_autocomplete_popup(doc)
        self.update_bracket_match_highlight(doc)
        self.update_line_number_gutter(doc)
        self.update_status()

    def get_auto_indent_unit(self, line_prefix):
        if '\t' in (line_prefix or '') and ' ' not in (line_prefix or ''):
            return '\t'
        return '    '

    def build_auto_indent_prefix(self, text_widget, insert_index):
        line_start = f"{insert_index} linestart"
        before_cursor = text_widget.get(line_start, insert_index)
        current_line = text_widget.get(line_start, f"{insert_index} lineend")
        indent_match = re.match(r'[ \t]*', current_line)
        base_indent = indent_match.group(0) if indent_match else ''
        trimmed_prefix = before_cursor.rstrip()
        if trimmed_prefix.endswith(tuple(self.bracket_pairs.keys())) or trimmed_prefix.endswith(':'):
            return base_indent + self.get_auto_indent_unit(base_indent)
        return base_indent

    def insert_auto_indent_newline(self, text_widget):
        try:
            insert_index = text_widget.index(tk.INSERT)
        except tk.TclError:
            return None
        indent_prefix = self.build_auto_indent_prefix(text_widget, insert_index)
        if text_widget.tag_ranges('sel'):
            try:
                text_widget.delete('sel.first', 'sel.last')
            except tk.TclError:
                pass
        text_widget.insert(tk.INSERT, '\n' + indent_prefix)
        return "break"

    def clear_bracket_match_highlight(self, text_widget):
        if not isinstance(text_widget, tk.Text):
            return
        try:
            if text_widget.winfo_exists():
                text_widget.tag_remove(self.bracket_match_tag, '1.0', tk.END)
        except tk.TclError:
            pass

    def find_matching_bracket_offset(self, content, bracket_offset):
        if bracket_offset < 0 or bracket_offset >= len(content):
            return None
        current_char = content[bracket_offset]
        if current_char in self.bracket_pairs:
            target_char = self.bracket_pairs[current_char]
            depth = 1
            for offset in range(bracket_offset + 1, len(content)):
                char = content[offset]
                if char == current_char:
                    depth += 1
                elif char == target_char:
                    depth -= 1
                    if depth == 0:
                        return offset
        elif current_char in self.reverse_bracket_pairs:
            target_char = self.reverse_bracket_pairs[current_char]
            depth = 1
            for offset in range(bracket_offset - 1, -1, -1):
                char = content[offset]
                if char == current_char:
                    depth += 1
                elif char == target_char:
                    depth -= 1
                    if depth == 0:
                        return offset
        return None

    def get_bracket_offsets_near_insert(self, text_widget):
        try:
            insert_offset = self.get_text_char_offset(text_widget, tk.INSERT)
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return None, None, None
        if not content or len(content) > self.bracket_match_max_chars:
            return None, None, content
        candidate_offsets = []
        if 0 <= insert_offset - 1 < len(content):
            candidate_offsets.append(insert_offset - 1)
        if 0 <= insert_offset < len(content):
            candidate_offsets.append(insert_offset)
        for candidate_offset in candidate_offsets:
            char = content[candidate_offset]
            if char in self.bracket_pairs or char in self.reverse_bracket_pairs:
                match_offset = self.find_matching_bracket_offset(content, candidate_offset)
                if match_offset is not None:
                    return candidate_offset, match_offset, content
        return None, None, content

    def update_bracket_match_highlight(self, doc_or_widget):
        if isinstance(doc_or_widget, dict):
            doc = doc_or_widget
            text_widget = doc.get('text')
            if doc.get('virtual_mode') or doc.get('preview_mode'):
                self.clear_bracket_match_highlight(text_widget)
                return
        else:
            doc = None
            text_widget = doc_or_widget
        if not isinstance(text_widget, tk.Text):
            return
        self.clear_bracket_match_highlight(text_widget)
        bracket_offset, match_offset, content = self.get_bracket_offsets_near_insert(text_widget)
        if bracket_offset is None or match_offset is None:
            return
        start_index = self.text_index_from_offset(text_widget, bracket_offset, content=content)
        end_index = self.text_index_from_offset(text_widget, bracket_offset + 1, content=content)
        match_index = self.text_index_from_offset(text_widget, match_offset, content=content)
        match_end_index = self.text_index_from_offset(text_widget, match_offset + 1, content=content)
        try:
            text_widget.tag_add(self.bracket_match_tag, start_index, end_index)
            text_widget.tag_add(self.bracket_match_tag, match_index, match_end_index)
            text_widget.tag_raise(self.bracket_match_tag)
        except tk.TclError:
            pass

    def remember_doc_focus(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if doc:
            self.set_last_active_editor_widget(doc['text'])
        return None

    def is_shift_selection_navigation(self, event):
        keysym = str(getattr(event, 'keysym', ''))
        state = int(getattr(event, 'state', 0) or 0)
        navigation_keys = {'Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}
        return bool((state & 0x1) and keysym in navigation_keys)

    def handle_text_keypress(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        state = int(getattr(event, 'state', 0) or 0)
        keysym = str(getattr(event, 'keysym', '') or '')

        if (state & 0x4) and (state & 0x1) and keysym.lower() == 'x':
            return self.ctrl_shift_x(event)

        if self.is_shift_selection_navigation(event):
            self.hide_autocomplete_popup()
            return

        if (
            event.keysym in {'Return', 'KP_Enter'}
            and not (state & 0x4)
            and not (self.autocomplete_popup_visible() and self.autocomplete_doc_id == str(tab_id))
        ):
            if not doc.get('virtual_mode') and not doc.get('preview_mode'):
                self.hide_autocomplete_popup()
                return self.insert_auto_indent_newline(doc['text'])

        if not doc.get('virtual_mode') and not doc.get('preview_mode'):
            if self.autocomplete_popup_visible() and self.autocomplete_doc_id == str(tab_id):
                if event.keysym == 'Up':
                    return self.move_autocomplete_selection(-1)
                if event.keysym == 'Down':
                    return self.move_autocomplete_selection(1)
                if event.keysym in {'Tab', 'Return', 'KP_Enter'}:
                    if self.accept_autocomplete_selection():
                        return "break"
                if event.keysym == 'Escape':
                    self.hide_autocomplete_popup()
                    return "break"
                if event.keysym in {'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}:
                    self.hide_autocomplete_popup()
            return

        navigation_keys = {
            'Up', 'Down', 'Left', 'Right', 'Prior', 'Next', 'Home', 'End',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Escape'
        }
        if event.keysym in navigation_keys or (event.state & 0x4):
            return
        return "break"

    def handle_compare_keypress(self, event):
        compare_doc = self.compare_view
        source_doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        state = int(getattr(event, 'state', 0) or 0)
        keysym = str(getattr(event, 'keysym', '') or '')

        if (state & 0x4) and (state & 0x1) and keysym.lower() == 'x':
            return self.ctrl_shift_x(event)

        if self.is_shift_selection_navigation(event):
            self.hide_autocomplete_popup()
            return
        if (
            event.keysym in {'Return', 'KP_Enter'}
            and not (state & 0x4)
            and not (self.autocomplete_popup_visible() and self.autocomplete_doc_id == (str(compare_doc['frame']) if compare_doc else None))
        ):
            if source_doc and not source_doc.get('virtual_mode') and not source_doc.get('preview_mode'):
                self.hide_autocomplete_popup()
                return self.insert_auto_indent_newline(compare_doc['text'])
        if source_doc and not source_doc.get('virtual_mode') and not source_doc.get('preview_mode'):
            compare_doc_id = str(compare_doc['frame']) if compare_doc else None
            if self.autocomplete_popup_visible() and self.autocomplete_doc_id == compare_doc_id:
                if event.keysym == 'Up':
                    return self.move_autocomplete_selection(-1)
                if event.keysym == 'Down':
                    return self.move_autocomplete_selection(1)
                if event.keysym in {'Tab', 'Return', 'KP_Enter'}:
                    if self.accept_autocomplete_selection():
                        return "break"
                if event.keysym == 'Escape':
                    self.hide_autocomplete_popup()
                    return "break"
                if event.keysym in {'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}:
                    self.hide_autocomplete_popup()
            return

        navigation_keys = {
            'Up', 'Down', 'Left', 'Right', 'Prior', 'Next', 'Home', 'End',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Escape', 'F3', 'F4'
        }
        if event.keysym in navigation_keys or (event.state & 0x4):
            return
        return "break"

    def handle_compare_activity(self, event=None):
        compare_doc = self.compare_view
        if not compare_doc:
            return
        compare_text = compare_doc.get('text')
        if not compare_text:
            return
        self.set_last_active_editor_widget(compare_text)
        event_type = str(getattr(event, 'type', ''))
        if self.is_shift_selection_navigation(event):
            self.hide_autocomplete_popup()
        elif event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease'}:
            self.hide_autocomplete_popup()
        elif event_type not in {'35'}:
            self.update_autocomplete_popup(compare_doc)
        self.update_bracket_match_highlight(compare_doc)
        self.update_line_number_gutter(compare_doc)
        self.update_status()

    def on_compare_modified(self, event=None):
        compare_doc = self.compare_view
        if not compare_doc:
            return
        compare_text = compare_doc.get('text')
        if not compare_text:
            return
        if compare_doc.get('suspend_modified_events'):
            compare_text.edit_modified(False)
            return

        source_doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        if not source_doc or source_doc.get('virtual_mode') or source_doc.get('preview_mode'):
            compare_text.edit_modified(False)
            return

        self.set_last_active_editor_widget(compare_text)
        self.handle_compare_activity()

        try:
            compare_content = compare_text.get('1.0', 'end-1c')
            compare_insert = compare_text.index(tk.INSERT)
        except tk.TclError:
            compare_text.edit_modified(False)
            return

        source_text = source_doc['text']
        try:
            source_content = source_text.get('1.0', 'end-1c')
        except tk.TclError:
            compare_text.edit_modified(False)
            return
        if compare_content == source_content:
            compare_doc['last_insert_index'] = compare_insert
            compare_text.edit_modified(False)
            self.update_line_number_gutter(compare_doc)
            self.update_status()
            return

        self.remember_doc_view_state(source_doc)
        source_doc['suspend_modified_events'] = True
        compare_doc['pushing_to_source'] = True
        try:
            source_text.delete('1.0', tk.END)
            source_text.insert('1.0', compare_content)
            source_text.edit_modified(True)
            source_doc['last_insert_index'] = compare_insert
        finally:
            source_doc['suspend_modified_events'] = False

        self.on_text_modified(source_doc['frame'])
        compare_doc['pushing_to_source'] = False
        source_doc['last_insert_index'] = compare_insert
        compare_doc['last_insert_index'] = compare_insert
        if compare_doc.get('syntax_mode') and compare_doc.get('syntax_mode') != 'python':
            self.schedule_syntax_highlight(compare_doc)
        self.update_line_number_gutter(compare_doc)
        self.update_status()
        compare_text.edit_modified(False)

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
            self.remember_doc_view_state(doc)
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
        self.remember_doc_view_state(doc)
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
            self.remember_doc_view_state(doc)
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

    def sync_compare_note_tags(self, source_doc=None):
        if not self.compare_active or not self.compare_view:
            return
        source_doc = source_doc or (self.documents.get(self.compare_source_tab) if self.compare_source_tab else None)
        compare_text = self.compare_view.get('text')
        if not source_doc or not compare_text:
            return
        try:
            for tag_name in list(compare_text.tag_names()):
                if tag_name.startswith('note_'):
                    compare_text.tag_delete(tag_name)
        except tk.TclError:
            return
        for note_tag, note_data in source_doc.get('notes', {}).items():
            try:
                ranges = source_doc['text'].tag_ranges(note_tag)
                if len(ranges) < 2:
                    continue
                highlight_color = self.get_note_color_hex(note_data.get('color'))
                compare_text.tag_add(note_tag, str(ranges[0]), str(ranges[1]))
                compare_text.tag_config(note_tag, background=highlight_color, foreground='black')
                compare_text.tag_bind(note_tag, '<Button-1>', lambda e, frame=source_doc['frame'], tag=note_tag: self.open_note_from_tag(e, frame, tag))
                compare_text.tag_bind(note_tag, '<Enter>', lambda e, text=compare_text: text.config(cursor='hand2'))
                compare_text.tag_bind(note_tag, '<Leave>', lambda e, text=compare_text: text.config(cursor='xterm'))
            except tk.TclError:
                continue

    def create_text_context_menu(self, tab_id, doc_override=None, action_tab_id=None):
        doc = doc_override or self.documents.get(str(tab_id))
        if not doc:
            return

        action_target = action_tab_id if action_tab_id is not None else tab_id
        menu = tk.Menu(self.root, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        menu.add_command(label=self.tr('context.cut', 'Cut'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.cut))
        menu.add_command(label=self.tr('context.copy', 'Copy'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.copy))
        menu.add_command(label=self.tr('context.paste', 'Paste'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.paste))
        menu.add_separator()
        menu.add_command(label=self.tr('context.select_all', 'Select All'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.select_all))
        menu.add_command(label=self.tr('context.add_note', 'Add note'), command=lambda frame=action_target: self.run_context_menu_action(lambda: self.add_note_to_selection(frame)))
        menu.add_command(label=self.tr('context.respond', 'Respond'), command=lambda frame=action_target: self.run_context_menu_action(lambda: self.respond_to_note(frame)))
        menu.add_command(label=self.tr('context.remove_note', 'Remove note'), command=lambda frame=action_target: self.run_context_menu_action(lambda: self.remove_note(frame)))
        doc['context_menu'] = menu

    def run_context_menu_action(self, callback):
        self.dismiss_context_menu()
        try:
            self.root.after(1, callback)
        except tk.TclError:
            try:
                return callback()
            except Exception as exc:
                self.log_exception("context menu action", exc)
        return "break"

    def get_context_action_target_widget(self, action_target):
        if action_target == '__compare__':
            return self.get_compare_text_widget()

        doc = self.documents.get(str(action_target))
        if not doc:
            return None

        target_widget = doc.get('context_target_widget')
        if isinstance(target_widget, tk.Text):
            try:
                if target_widget.winfo_exists():
                    return target_widget
            except tk.TclError:
                pass
        return doc.get('text')

    def invoke_context_menu_widget_action(self, callback, target_widget):
        if target_widget is None:
            return "break"
        try:
            try:
                if target_widget.winfo_exists():
                    target_widget.focus_force()
            except tk.TclError:
                pass
            self.set_last_active_editor_widget(target_widget)
            return callback(SimpleNamespace(widget=target_widget))
        except Exception as exc:
            self.log_exception("context menu widget action", exc)
            return "break"

    def run_context_menu_widget_action(self, action_target, callback):
        target_widget = self.get_context_action_target_widget(action_target)
        self.dismiss_context_menu()
        if target_widget is None:
            return "break"
        try:
            self.root.after(
                1,
                lambda widget=target_widget, handler=callback: self.invoke_context_menu_widget_action(handler, widget)
            )
        except tk.TclError:
            return self.invoke_context_menu_widget_action(callback, target_widget)
        return "break"

    def dismiss_context_menu(self, event=None):
        menu = getattr(self, 'active_context_menu', None)
        if menu is None:
            return None
        try:
            menu.unpost()
        except tk.TclError:
            pass
        self.active_context_menu = None
        self.context_menu_posted_at = 0.0
        return None

    def is_note_popup_widget(self, widget):
        if widget is None:
            return False
        current = widget
        while current is not None:
            for doc in self.documents.values():
                if doc.get('note_popup') == current:
                    return True
            parent_name = getattr(current, 'master', None)
            if parent_name is None:
                break
            current = parent_name
        return False

    def dismiss_note_popups(self):
        for doc in self.documents.values():
            self.hide_note_popup(doc)
        return None

    def click_is_on_note(self, event=None):
        widget = getattr(event, 'widget', None)
        if not isinstance(widget, tk.Text):
            return False
        doc = self.get_doc_for_text_widget(widget)
        if not doc or not doc.get('notes'):
            return False
        try:
            index = widget.index(f"@{event.x},{event.y}")
            return any(tag in doc['notes'] for tag in widget.tag_names(index))
        except tk.TclError:
            return False

    def maybe_dismiss_transient_ui(self, event=None):
        menu = getattr(self, 'active_context_menu', None)
        widget = getattr(event, 'widget', None)
        if isinstance(widget, tk.Menu):
            return None
        if menu is not None and (time.monotonic() - float(getattr(self, 'context_menu_posted_at', 0.0) or 0.0)) < 0.25:
            return None
        if self.is_note_popup_widget(widget):
            if menu is not None:
                try:
                    self.root.after_idle(self.dismiss_context_menu)
                except tk.TclError:
                    self.dismiss_context_menu()
            return None
        if self.click_is_on_note(event):
            if menu is not None:
                try:
                    self.root.after_idle(self.dismiss_context_menu)
                except tk.TclError:
                    self.dismiss_context_menu()
            return None
        try:
            self.root.after_idle(self.dismiss_context_menu)
        except tk.TclError:
            self.dismiss_context_menu()
        try:
            self.root.after_idle(self.dismiss_note_popups)
        except tk.TclError:
            self.dismiss_note_popups()
        return None

    def show_text_context_menu(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        return self.show_context_menu_for_doc(event, doc, select_tab=True)

    def show_compare_context_menu(self, event):
        doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        return self.show_context_menu_for_doc(event, doc, select_tab=False)

    def show_context_menu_for_doc(self, event, doc, select_tab):
        if not doc:
            return "break"

        if select_tab:
            self.notebook.select(doc['frame'])
            self.set_active_document(doc['frame'])

        target_widget = event.widget if isinstance(getattr(event, 'widget', None), tk.Text) else doc['text']
        if not isinstance(target_widget, tk.Text):
            target_widget = doc['text']
        try:
            target_widget.focus_force()
        except tk.TclError:
            pass
        self.set_last_active_editor_widget(target_widget)

        selection_ranges = target_widget.tag_ranges('sel')
        note_state = 'normal' if len(selection_ranges) >= 2 and str(selection_ranges[0]) != str(selection_ranges[1]) else 'disabled'

        index = target_widget.index(f"@{event.x},{event.y}")
        clicked_note_tags = [tag for tag in target_widget.tag_names(index) if tag in doc['notes']]
        doc['context_note_tag'] = clicked_note_tags[-1] if clicked_note_tags else None
        doc['context_target_widget'] = target_widget

        is_readonly_target = bool(doc.get('preview_mode') or doc.get('virtual_mode'))
        if is_readonly_target:
            doc['context_menu'].entryconfig(self.tr('context.cut', 'Cut'), state='disabled')
            doc['context_menu'].entryconfig(self.tr('context.paste', 'Paste'), state='disabled')
            note_state = 'disabled'
        else:
            doc['context_menu'].entryconfig(self.tr('context.cut', 'Cut'), state='normal')
            doc['context_menu'].entryconfig(self.tr('context.paste', 'Paste'), state='normal')

        doc['context_menu'].entryconfig(self.tr('context.add_note', 'Add note'), state=note_state)
        note_action_state = 'normal' if doc.get('context_note_tag') else 'disabled'
        doc['context_menu'].entryconfig(self.tr('context.respond', 'Respond'), state=note_action_state)
        doc['context_menu'].entryconfig(self.tr('context.remove_note', 'Remove note'), state=note_action_state)
        try:
            self.dismiss_context_menu()
            self.active_context_menu = doc['context_menu']
            self.context_menu_posted_at = time.monotonic()
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

        highlight_color = self.get_note_color_hex(note_data.get('color'))
        frame = tk.Frame(popup, bg='#1f2430', highlightthickness=1, highlightbackground=highlight_color)
        frame.pack(fill='both', expand=True)

        tk.Label(
            frame,
            text=self.tr('note.popup.title', 'Code note'),
            bg='#1f2430',
            fg=highlight_color,
            font=('Segoe UI', 9, 'bold'),
            anchor=self.ui_anchor_start()
        ).pack(fill='x', padx=10, pady=(8, 2))

        author_line_parts = []
        if note_data.get('author_label'):
            author_line_parts.append(f"{self.tr('note.popup.author', 'Author')}: {note_data['author_label']}")
        elif note_data.get('author_id'):
            author_line_parts.append(f"{self.tr('note.popup.author_id', 'Author ID')}: {note_data['author_id']}")
        author_location_parts = []
        if note_data.get('author_host'):
            author_location_parts.append(note_data['author_host'])
        if note_data.get('author_ip'):
            author_location_parts.append(note_data['author_ip'])
        if author_location_parts:
            author_line_parts.append(" | ".join(author_location_parts))
        meta_lines = []
        if author_line_parts:
            meta_lines.append(" | ".join(author_line_parts))
        if note_data.get('created_at'):
            meta_lines.append(f"{self.tr('note.popup.created', 'Created')}: {self.format_note_timestamp(note_data.get('created_at'))}")
        if meta_lines:
            tk.Label(
                frame,
                text="\n".join(meta_lines),
                bg='#1f2430',
                fg='#9aa0a6',
                font=('Segoe UI', 8),
                justify=self.ui_justify(),
                wraplength=280,
                anchor=self.ui_anchor_start()
            ).pack(fill='x', padx=10, pady=(0, 6))

        tk.Label(
            frame,
            text=note_data['text'],
            bg='#1f2430',
            fg='#f5f5f5',
            font=('Segoe UI', 9),
            justify=self.ui_justify(),
            wraplength=280,
            anchor=self.ui_anchor_start()
        ).pack(fill='both', expand=True, padx=10, pady=(0, 8))

        responses = self.sanitize_note_responses(note_data.get('responses', []))
        if responses:
            tk.Frame(frame, bg=highlight_color, height=1).pack(fill='x', padx=10, pady=(0, 6))
            tk.Label(
                frame,
                text=self.tr('note.popup.responses', 'Responses'),
                bg='#1f2430',
                fg='#c9d1d9',
                font=('Segoe UI', 8, 'bold'),
                anchor=self.ui_anchor_start()
            ).pack(fill='x', padx=10, pady=(0, 4))
            for response in responses:
                response_color = self.get_note_color_hex(response.get('color'))
                response_line_parts = []
                if response.get('author_label'):
                    response_line_parts.append(response['author_label'])
                elif response.get('author_id'):
                    response_line_parts.append(response['author_id'])
                response_location_parts = []
                if response.get('author_host'):
                    response_location_parts.append(response['author_host'])
                if response.get('author_ip'):
                    response_location_parts.append(response['author_ip'])
                if response_location_parts:
                    response_line_parts.append(" | ".join(response_location_parts))
                response_meta_lines = []
                if response_line_parts:
                    response_meta_lines.append(" | ".join(response_line_parts))
                if response.get('created_at'):
                    response_meta_lines.append(
                        f"{self.tr('note.popup.created', 'Created')}: {self.format_note_timestamp(response.get('created_at'))}"
                    )
                if response_meta_lines:
                    tk.Label(
                        frame,
                        text="\n".join(response_meta_lines),
                        bg='#1f2430',
                        fg=response_color,
                        font=('Segoe UI', 8, 'bold'),
                        justify=self.ui_justify(),
                        wraplength=280,
                        anchor=self.ui_anchor_start()
                    ).pack(fill='x', padx=10, pady=(0, 2))
                tk.Label(
                    frame,
                    text=response.get('text', ''),
                    bg='#1f2430',
                    fg='#f5f5f5',
                    font=('Segoe UI', 9),
                    justify=self.ui_justify(),
                    wraplength=280,
                    anchor=self.ui_anchor_start()
                ).pack(fill='x', padx=10, pady=(0, 6))

        popup.update_idletasks()
        self.center_window(popup, self.root)
        popup.bind('<Escape>', lambda e, current=doc: self.hide_note_popup(current))
        doc['note_popup'] = popup

    def prompt_note_input(self, title, prompt, initialvalue="", parent=None):
        parent = parent or self.root
        self.hide_autocomplete_popup()
        self.autocomplete_suspended += 1
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)
        self.apply_window_icon(dialog)

        result = {'value': None}
        value_var = tk.StringVar(value=initialvalue)

        tk.Label(
            dialog,
            text=prompt,
            bg='#f0f0f0',
            fg='black',
            font=('Segoe UI', 9)
        ).pack(anchor=self.ui_anchor_start(), pady=(0, 6))

        entry = tk.Entry(dialog, textvariable=value_var, width=34)
        entry.pack(fill='x')

        button_row = tk.Frame(dialog, bg='#f0f0f0')
        button_row.pack(fill='x', pady=(10, 0))

        def submit(event=None):
            result['value'] = value_var.get()
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()
            return "break"

        tk.Button(button_row, text="OK", width=10, command=submit).pack(side='right' if self.is_rtl_locale() else 'left')
        tk.Button(button_row, text="Cancel", width=10, command=cancel).pack(side='left' if self.is_rtl_locale() else 'right')

        entry.bind('<Return>', submit)
        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol("WM_DELETE_WINDOW", cancel)

        dialog.update_idletasks()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        entry.focus_force()
        entry.selection_range(0, tk.END)
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.autocomplete_suspended = max(0, self.autocomplete_suspended - 1)
            self.hide_autocomplete_popup()

    def prompt_note_color(self, title="Note Color", initialvalue='yellow', parent=None):
        parent = parent or self.root
        self.hide_autocomplete_popup()
        self.autocomplete_suspended += 1
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)
        self.apply_window_icon(dialog)

        result = {'value': None}
        color_var = tk.StringVar(value=self.normalize_note_color(initialvalue))

        tk.Label(
            dialog,
            text="Color:",
            bg='#f0f0f0',
            fg='black',
            font=('Segoe UI', 9)
        ).pack(anchor=self.ui_anchor_start(), pady=(0, 6))

        choices_frame = tk.Frame(dialog, bg='#f0f0f0')
        choices_frame.pack(fill='x')
        for color_key in self.note_color_order:
            tk.Radiobutton(
                choices_frame,
                text=self.get_note_color_label(color_key),
                variable=color_var,
                value=color_key,
                bg='#f0f0f0',
                anchor=self.ui_anchor_start(),
                selectcolor=self.get_note_color_hex(color_key)
            ).pack(anchor=self.ui_anchor_start())

        button_row = tk.Frame(dialog, bg='#f0f0f0')
        button_row.pack(fill='x', pady=(10, 0))

        def submit(event=None):
            result['value'] = self.normalize_note_color(color_var.get())
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()
            return "break"

        tk.Button(button_row, text="OK", width=10, command=submit).pack(side='right' if self.is_rtl_locale() else 'left')
        tk.Button(button_row, text="Cancel", width=10, command=cancel).pack(side='left' if self.is_rtl_locale() else 'right')

        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol("WM_DELETE_WINDOW", cancel)

        dialog.update_idletasks()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.autocomplete_suspended = max(0, self.autocomplete_suspended - 1)
            self.hide_autocomplete_popup()

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
                'color': self.get_note_color_label(note_data.get('color')),
                'responses': self.sanitize_note_responses(note_data.get('responses', [])),
            }
            note_rows.append(row)
        try:
            if output_path.lower().endswith('.md'):
                markdown_parts = [f"# Notes Export for {self.get_doc_name(doc['frame'])}\n"]
                for row in note_rows:
                    markdown_parts.append(f"\n## Note {row['id']}\n")
                    markdown_parts.append(f"\n- Range: `{row['selection_start']}` to `{row['selection_end']}`\n")
                    if row['author']:
                        markdown_parts.append(f"- Author: {row['author']}\n")
                    if row['color']:
                        markdown_parts.append(f"- Color: {row['color']}\n")
                    if row['created_at']:
                        markdown_parts.append(f"- Created: {row['created_at']}\n")
                    markdown_parts.append("\n### Selected Text\n\n```\n")
                    markdown_parts.append(row['selected_text'])
                    markdown_parts.append("\n```\n\n### Code Note\n\n")
                    markdown_parts.append(row['note'] or '')
                    markdown_parts.append("\n")
                    responses = row.get('responses') or []
                    if responses:
                        markdown_parts.append("\n### Responses\n")
                        for response in responses:
                            author = response.get('author_label') or response.get('author_id') or 'Unknown'
                            color = self.get_note_color_label(response.get('color'))
                            created = response.get('created_at') or ''
                            markdown_parts.append(f"\n- {author} | {color}")
                            if created:
                                markdown_parts.append(f" | {created}")
                            markdown_parts.append(f"\n\n  {response.get('text', '')}\n")
                self.write_file_atomically(output_path, ''.join(markdown_parts))
            else:
                if not self.write_json_atomically(output_path, note_rows, 'notepadx-export-', 'export notes report'):
                    raise OSError(output_path)
            messagebox.showinfo("Export Notes", f"Notes exported to:\n{output_path}", parent=self.root)
        except PermissionError as exc:
            self.show_filesystem_error("Export Notes", output_path, exc)
        except OSError as exc:
            self.log_exception("export notes report", exc)
            self.show_filesystem_error("Export Notes", output_path, exc)
        return "break"

    def add_note_to_selection(self, tab_id=None, event=None):
        if tab_id == '__compare__':
            doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        else:
            doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or self.current_doc_is_large_readonly():
            return "break"

        selection_widget = doc.get('context_target_widget')
        compare_widget = self.get_compare_text_widget()
        if selection_widget is None or not isinstance(selection_widget, tk.Text):
            selection_widget = compare_widget if tab_id == '__compare__' else doc['text']

        try:
            start = selection_widget.index('sel.first')
            end = selection_widget.index('sel.last')
        except tk.TclError:
            messagebox.showinfo(
                self.tr('note.prompt.add_title', 'Add Note'),
                self.tr('note.prompt.select_text_first', 'Select some text first.'),
                parent=self.root
            )
            return "break"

        self.hide_note_popup(doc)
        suggested_author = getattr(self, 'note_author_name', '') or ''
        author_name = self.prompt_note_input(
            self.tr('note.prompt.add_title', 'Add Note'),
            self.tr('note.prompt.author', 'Author:'),
            initialvalue=suggested_author,
            parent=self.root
        )
        if author_name is None:
            return "break"
        author_name = self.trim_text(author_name, self.max_note_name_length)
        if not author_name:
            messagebox.showinfo(
                self.tr('note.prompt.add_title', 'Add Note'),
                self.tr('note.prompt.author_required', 'Enter an author name first.'),
                parent=self.root
            )
            return "break"
        self.note_author_name = author_name
        note_input = self.prompt_note_input(
            self.tr('note.prompt.add_title', 'Add Note'),
            self.tr('note.prompt.note', 'Note:'),
            parent=self.root
        )
        if not note_input:
            return "break"

        note_text = self.trim_text(note_input, self.max_note_text_length)
        if note_text.lower().startswith("# note:"):
            note_text = self.trim_text(note_text[7:], self.max_note_text_length)
        if not note_text:
            return "break"
        note_color = self.prompt_note_color(parent=self.root)
        if not note_color:
            return "break"

        note_tag = self.create_note_tag(
            doc,
            start,
            end,
            note_text,
            note_color=note_color,
            author_id=self.editor_id,
            author_label=author_name,
            author_host=self.get_local_machine_name(),
            author_ip=self.get_local_lan_ip(),
            read_by=[self.editor_id]
        )
        self.persist_doc_notes(doc)
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.sync_compare_note_tags(doc)
        return "break"

    def create_note_tag(
        self,
        doc,
        start,
        end,
        note_text,
        note_color='yellow',
        note_id=None,
        author_id=None,
        author_label=None,
        author_host=None,
        author_ip=None,
        read_by=None,
        author_unread=False,
        created_at=None,
        anchor_text=None,
        anchor_line=None,
        responses=None
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
        safe_note_text = self.trim_text(note_text, self.max_note_text_length)
        safe_note_color = self.normalize_note_color(note_color)
        safe_author_id = self.trim_text(author_id, 128)
        safe_author_label = self.trim_text(author_label, self.max_note_name_length)
        safe_author_host = self.trim_text(author_host, 128)
        safe_author_ip = self.trim_text(author_ip, 64)
        safe_anchor_text = anchor_text if anchor_text is not None else doc['text'].get(start, end)
        safe_anchor_text = str(safe_anchor_text)[:self.max_note_text_length]
        doc['notes'][note_tag] = {
            'id': str(note_id),
            'text': safe_note_text,
            'color': safe_note_color,
            'author_id': safe_author_id,
            'author_label': safe_author_label,
            'author_host': safe_author_host,
            'author_ip': safe_author_ip,
            'read_by': normalized_read_by[:64],
            'author_unread': bool(author_unread),
            'created_at': created_at or self.utc_timestamp(),
            'anchor_text': safe_anchor_text,
            'anchor_line': anchor_line if anchor_line is not None else int(str(start).split('.')[0]),
            'responses': self.sanitize_note_responses(responses),
        }
        self.apply_note_tag(doc, note_tag, start, end)
        return note_tag

    def respond_to_note(self, tab_id=None):
        if tab_id == '__compare__':
            doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        else:
            doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        note_tag = doc.get('context_note_tag')
        note_data = doc['notes'].get(note_tag)
        if not note_data:
            return "break"

        suggested_author = getattr(self, 'note_author_name', '') or ''
        author_name = self.prompt_note_input(
            self.tr('note.prompt.respond_title', 'Respond'),
            self.tr('note.prompt.name', 'Name:'),
            initialvalue=suggested_author,
            parent=self.root
        )
        if author_name is None:
            return "break"
        author_name = self.trim_text(author_name, self.max_note_name_length)
        if not author_name:
            messagebox.showinfo(
                self.tr('note.prompt.respond_title', 'Respond'),
                self.tr('note.prompt.name_required', 'Enter a name first.'),
                parent=self.root
            )
            return "break"
        self.note_author_name = author_name

        response_text = self.prompt_note_input(
            self.tr('note.prompt.respond_title', 'Respond'),
            self.tr('note.prompt.response', 'Response:'),
            parent=self.root
        )
        if response_text is None:
            return "break"
        response_text = self.trim_text(response_text, self.max_note_text_length)
        if not response_text:
            return "break"

        response_color = self.prompt_note_color(
            title=self.tr('note.prompt.response_color', 'Response Color'),
            initialvalue=note_data.get('color', 'yellow'),
            parent=self.root
        )
        if not response_color:
            return "break"

        responses = list(note_data.get('responses', []))
        responses.append({
            'author_id': self.editor_id,
            'author_label': author_name,
            'author_host': self.get_local_machine_name(),
            'author_ip': self.get_local_lan_ip(),
            'text': response_text,
            'color': self.normalize_note_color(response_color),
            'created_at': self.utc_timestamp(),
        })
        note_data['responses'] = self.sanitize_note_responses(responses)
        note_data['color'] = self.normalize_note_color(response_color)
        note_data['read_by'] = [self.editor_id]
        note_data['author_unread'] = note_data.get('author_id') not in self.editor_aliases

        ranges = doc['text'].tag_ranges(note_tag)
        if len(ranges) >= 2:
            doc['text'].tag_remove(note_tag, '1.0', tk.END)
            self.apply_note_tag(doc, note_tag, str(ranges[0]), str(ranges[1]))
        self.persist_doc_notes(doc)
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.sync_compare_note_tags(doc)
        return "break"

    def apply_note_tag(self, doc, note_tag, start, end):
        note_data = doc['notes'][note_tag]
        highlight_color = self.get_note_color_hex(note_data.get('color'))
        doc['text'].tag_add(note_tag, start, end)
        doc['text'].tag_config(note_tag, background=highlight_color, foreground='black')
        doc['text'].tag_bind(note_tag, '<Button-1>', lambda e, frame=doc['frame'], tag=note_tag: self.open_note_from_tag(e, frame, tag))
        doc['text'].tag_bind(note_tag, '<Enter>', lambda e, text=doc['text']: text.config(cursor='hand2'))
        doc['text'].tag_bind(note_tag, '<Leave>', lambda e, text=doc['text']: text.config(cursor='xterm'))

    def set_note_color(self, color_key, tab_id=None):
        if tab_id == '__compare__':
            doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        else:
            doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        note_tag = doc.get('context_note_tag')
        note_data = doc['notes'].get(note_tag)
        if not note_data:
            return "break"

        note_data['color'] = self.normalize_note_color(color_key)
        ranges = doc['text'].tag_ranges(note_tag)
        if len(ranges) >= 2:
            doc['text'].tag_remove(note_tag, '1.0', tk.END)
            self.apply_note_tag(doc, note_tag, str(ranges[0]), str(ranges[1]))
        self.persist_doc_notes(doc)
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.sync_compare_note_tags(doc)
        return "break"

    def remove_note(self, tab_id=None):
        if tab_id == '__compare__':
            doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        else:
            doc = self.documents.get(str(tab_id)) if tab_id is not None else self.get_current_doc()
        if not doc or not doc.get('context_note_tag'):
            return "break"

        note_tag = doc['context_note_tag']
        removed_note = doc['notes'].get(note_tag, {})
        removed_note_id = removed_note.get('id')
        doc['text'].tag_delete(note_tag)
        doc['text'].config(cursor='xterm')
        doc['notes'].pop(note_tag, None)
        doc['context_note_tag'] = None
        self.hide_note_popup(doc)
        self.persist_doc_notes(doc)
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.sync_compare_note_tags(doc)
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

    def get_canonical_document_filename(self, file_path):
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        try:
            if os.path.isdir(directory):
                directory_entries = os.listdir(directory)
                exact_match = next((entry for entry in directory_entries if entry == filename), None)
                if exact_match:
                    return exact_match
                casefold_match = next((entry for entry in directory_entries if entry.lower() == filename.lower()), None)
                if casefold_match:
                    return casefold_match
        except OSError:
            pass
        return filename

    def resolve_sidecar_path(self, file_path, suffix):
        directory = os.path.dirname(file_path)
        preferred_filename = self.get_canonical_document_filename(file_path)
        return os.path.join(directory, f"{preferred_filename}{suffix}")

    def get_sidecar_variants(self, sidecar_path):
        directory = os.path.dirname(sidecar_path)
        target_name = os.path.basename(sidecar_path)
        variants = []
        seen = set()
        try:
            if os.path.isdir(directory):
                for entry in os.listdir(directory):
                    if entry.lower() != target_name.lower():
                        continue
                    full_path = os.path.join(directory, entry)
                    normalized = os.path.normcase(os.path.abspath(full_path))
                    if normalized in seen:
                        continue
                    seen.add(normalized)
                    variants.append(full_path)
        except OSError:
            pass
        if not variants:
            return [sidecar_path]
        canonical_normalized = os.path.normcase(os.path.abspath(sidecar_path))
        variants.sort(key=lambda path: (
            0 if os.path.normcase(os.path.abspath(path)) == canonical_normalized else 1,
            path.lower()
        ))
        return variants

    def cleanup_duplicate_sidecar_variants(self, canonical_path):
        canonical_normalized = os.path.normcase(os.path.abspath(canonical_path))
        for variant_path in self.get_sidecar_variants(canonical_path):
            if os.path.normcase(os.path.abspath(variant_path)) == canonical_normalized:
                continue
            try:
                os.remove(variant_path)
            except OSError:
                pass

    def get_notes_sidecar_path(self, file_path):
        return self.resolve_sidecar_path(file_path, ".notepadx.notes.json")

    def get_editors_sidecar_path(self, file_path):
        return self.resolve_sidecar_path(file_path, ".notepadx.editors.json")

    def get_notes_sidecar_signature(self, sidecar_path):
        variant_paths = [path for path in self.get_sidecar_variants(sidecar_path) if os.path.exists(path)]
        if not variant_paths:
            return None
        notes = []
        found_payload = False
        try:
            variant_paths.sort(key=lambda path: (os.path.getmtime(path), path.lower()))
        except OSError:
            variant_paths.sort(key=lambda path: path.lower())
        for variant_path in variant_paths:
            payload = self.read_json_file(variant_path, "load shared notes signature", None)
            if not isinstance(payload, dict):
                continue
            found_payload = True
            raw_notes = payload.get('notes', [])
            for note in raw_notes if isinstance(raw_notes, list) else []:
                sanitized_note = self.sanitize_note_payload(note)
                if sanitized_note is not None:
                    notes.append(sanitized_note)
        if not found_payload:
            return None
        notes = self.dedupe_notes_payload(notes)
        serialized_notes = json.dumps(notes, sort_keys=True, separators=(',', ':'))
        return hashlib.md5(serialized_notes.encode('utf-8')).hexdigest()

    def get_editors_sidecar_signature(self, sidecar_path):
        variant_paths = [path for path in self.get_sidecar_variants(sidecar_path) if os.path.exists(path)]
        if not variant_paths:
            return None
        editors = []
        found_payload = False
        try:
            variant_paths.sort(key=lambda path: (os.path.getmtime(path), path.lower()))
        except OSError:
            variant_paths.sort(key=lambda path: path.lower())
        for variant_path in variant_paths:
            payload = self.read_json_file(variant_path, "load shared editors signature", None)
            if not isinstance(payload, dict):
                continue
            found_payload = True
            raw_editors = payload.get('editors', [])
            if isinstance(raw_editors, list):
                editors.extend(raw_editors)
        if not found_payload:
            return None
        editors = self.prune_inactive_shared_editors(editors)
        serialized_editors = json.dumps(editors, sort_keys=True, separators=(',', ':'))
        return hashlib.md5(serialized_editors.encode('utf-8')).hexdigest()

    def apply_shared_editors_to_doc(self, doc, shared_editors_payload):
        doc['note_editors'] = self.sanitize_shared_editors(shared_editors_payload.get('editors', []))
        doc['note_active_editors'] = max(shared_editors_payload.get('active_editors', 0), len(doc['note_editors']))
        existing_editor = next((entry for entry in doc['note_editors'] if entry.get('id') == self.editor_id), None)
        doc['note_editor_label'] = existing_editor.get('label') if existing_editor else None

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
            last_seen = self.normalize_optional_metadata(entry.get('last_seen'))
            host = self.trim_text(self.normalize_optional_metadata(entry.get('host')), 128)
            ip = self.trim_text(self.normalize_optional_metadata(entry.get('ip')), 64)
            sanitized.append({'id': editor_id, 'label': label, 'pid': pid, 'last_seen': last_seen, 'host': host, 'ip': ip})
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
        now = datetime.now(timezone.utc)
        active_editors = []
        for entry in self.sanitize_shared_editors(editors):
            last_seen = self.parse_iso_datetime(entry.get('last_seen'))
            if last_seen is not None and last_seen.tzinfo is None:
                last_seen = last_seen.replace(tzinfo=timezone.utc)
            last_seen_recent = bool(last_seen and abs((now - last_seen).total_seconds()) <= self.shared_editor_stale_seconds)
            if self.is_editor_process_alive(entry.get('pid')) or last_seen_recent:
                active_editors.append(entry)
        return active_editors

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

    def refresh_doc_shared_editor_state(self, doc, force_write=False):
        if not doc or not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return
        sidecar_path = self.get_editors_sidecar_path(doc['file_path'])
        payload = self.load_shared_editors(sidecar_path)
        editors = self.sanitize_shared_editors(payload.get('editors', []))
        current_timestamp = self.utc_timestamp()
        existing_editor = next((entry for entry in editors if entry.get('id') == self.editor_id), None)
        local_host = self.get_local_machine_name()
        local_ip = self.get_local_lan_ip()
        if existing_editor is None:
            existing_editor = {
                'id': self.editor_id,
                'label': self.allocate_editor_label(editors),
                'pid': os.getpid(),
                'last_seen': current_timestamp,
                'host': local_host,
                'ip': local_ip,
            }
            editors.append(existing_editor)
            force_write = True
        else:
            existing_editor['pid'] = os.getpid()
            existing_editor['last_seen'] = current_timestamp
            existing_editor['host'] = local_host
            existing_editor['ip'] = local_ip
        doc['note_editors'] = editors
        doc['note_editor_label'] = existing_editor.get('label')
        doc['note_active_editors'] = len(editors)
        self.write_shared_editors(sidecar_path, editors)
        doc['note_editors_signature'] = self.get_editors_sidecar_signature(sidecar_path)
        doc['note_last_heartbeat_at'] = time.monotonic()

    def export_doc_notes(self, doc):
        exported = []
        for note_tag, note_data in doc['notes'].items():
            exported_note = self.build_exported_note(doc, note_tag)
            if exported_note is not None:
                exported.append(exported_note)
        return exported

    def build_exported_note(self, doc, note_tag, fallback_start=None, fallback_end=None):
        if not doc or note_tag not in doc.get('notes', {}):
            return None
        note_data = doc['notes'][note_tag]
        ranges = doc['text'].tag_ranges(note_tag)
        if len(ranges) >= 2:
            start = str(ranges[0])
            end = str(ranges[1])
        else:
            start = str(fallback_start).strip() if fallback_start else None
            end = str(fallback_end).strip() if fallback_end else None
            if not start or not end:
                return None
        return {
            'id': note_data.get('id'),
            'start': start,
            'end': end,
            'text': note_data['text'],
            'color': self.normalize_note_color(note_data.get('color')),
            'author_id': note_data.get('author_id'),
            'author_label': note_data.get('author_label'),
            'author_host': note_data.get('author_host'),
            'author_ip': note_data.get('author_ip'),
            'read_by': [str(editor_id).strip()[:128] for editor_id in note_data.get('read_by', []) if str(editor_id).strip()][:64],
            'author_unread': bool(note_data.get('author_unread')),
            'created_at': note_data.get('created_at'),
            'anchor_text': (note_data.get('anchor_text') or '')[:self.max_note_text_length],
            'anchor_line': note_data.get('anchor_line'),
            'responses': self.sanitize_note_responses(note_data.get('responses', [])),
        }

    def refresh_doc_note_signatures(self, doc):
        if not doc or not doc.get('file_path'):
            return
        notes_sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        editors_sidecar_path = self.get_editors_sidecar_path(doc['file_path'])
        doc['note_sync_mtime'] = os.path.getmtime(notes_sidecar_path) if os.path.exists(notes_sidecar_path) else None
        doc['note_sync_signature'] = self.get_notes_sidecar_signature(notes_sidecar_path)
        doc['note_editors_signature'] = self.get_editors_sidecar_signature(editors_sidecar_path)
        doc['note_last_heartbeat_at'] = time.monotonic()
        doc['last_unread_count'] = self.get_unread_note_count(doc)

    def write_doc_notes_payload(self, doc, notes_payload):
        if not doc or not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return False
        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        try:
            self.write_shared_notes(sidecar_path, notes_payload, doc.get('note_active_editors', 0), doc.get('note_editors', []))
            self.refresh_doc_note_signatures(doc)
            return True
        except PermissionError as exc:
            self.show_filesystem_error("Code Notes", sidecar_path, exc)
        except OSError as exc:
            self.log_exception("write doc notes payload", exc)
            self.show_filesystem_error("Code Notes", sidecar_path, exc)
        return False

    def append_shared_note_to_sidecar(self, doc, exported_note):
        if not exported_note:
            return False
        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        payload = self.load_shared_notes(sidecar_path)
        notes = []
        target_note_id = str(exported_note.get('id', '')).strip()
        for note in payload.get('notes', []):
            if target_note_id and str(note.get('id', '')).strip() == target_note_id:
                continue
            sanitized_note = self.sanitize_note_payload(note)
            if sanitized_note is not None:
                notes.append(sanitized_note)
        notes.append(exported_note)
        return self.write_doc_notes_payload(doc, notes)

    def mutate_shared_note_in_sidecar(self, doc, note_id, mutate_callback):
        if not doc or not doc.get('file_path') or not note_id:
            return False
        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        payload = self.load_shared_notes(sidecar_path)
        notes = []
        found_note = False
        target_note_id = str(note_id).strip()
        for note in payload.get('notes', []):
            current_note_id = str(note.get('id', '')).strip()
            if current_note_id != target_note_id:
                sanitized_note = self.sanitize_note_payload(note)
                if sanitized_note is not None:
                    notes.append(sanitized_note)
                continue
            found_note = True
            updated_note = dict(note)
            callback_result = mutate_callback(updated_note)
            if callback_result is False:
                continue
            if isinstance(callback_result, dict):
                updated_note = callback_result
            sanitized_note = self.sanitize_note_payload(updated_note)
            if sanitized_note is not None:
                notes.append(sanitized_note)
        if not found_note:
            return False
        return self.write_doc_notes_payload(doc, notes)

    def sync_single_note_to_sidecar(self, doc, note_tag=None, remove_note_id=None, fallback_start=None, fallback_end=None):
        if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        try:
            payload = self.load_shared_notes(sidecar_path)
            notes = payload.get('notes', [])
            if not isinstance(notes, list):
                notes = []

            target_note_id = None
            if note_tag is not None:
                note_data = doc.get('notes', {}).get(note_tag)
                if note_data:
                    target_note_id = str(note_data.get('id', '')).strip() or None
            if remove_note_id is not None:
                target_note_id = str(remove_note_id).strip() or target_note_id

            existing_note = None
            if target_note_id:
                for note in notes:
                    if str(note.get('id', '')).strip() == target_note_id:
                        existing_note = dict(note)
                        break
                notes = [note for note in notes if str(note.get('id', '')).strip() != target_note_id]

            if note_tag is not None:
                exported_note = self.build_exported_note(doc, note_tag, fallback_start=fallback_start, fallback_end=fallback_end)
                if exported_note is None and existing_note is not None:
                    local_note = doc.get('notes', {}).get(note_tag)
                    if local_note:
                        exported_note = dict(existing_note)
                        exported_note.update({
                            'text': local_note.get('text'),
                            'color': self.normalize_note_color(local_note.get('color')),
                            'author_id': local_note.get('author_id'),
                            'author_label': local_note.get('author_label'),
                            'read_by': [str(editor_id).strip()[:128] for editor_id in local_note.get('read_by', []) if str(editor_id).strip()][:64],
                            'author_unread': bool(local_note.get('author_unread')),
                            'created_at': local_note.get('created_at'),
                            'anchor_text': (local_note.get('anchor_text') or '')[:self.max_note_text_length],
                            'anchor_line': local_note.get('anchor_line'),
                            'responses': self.sanitize_note_responses(local_note.get('responses', [])),
                        })
                if exported_note is not None:
                    notes.append(exported_note)

            self.write_shared_notes(sidecar_path, notes, doc.get('note_active_editors', 0), doc.get('note_editors', []))
            doc['note_sync_mtime'] = os.path.getmtime(sidecar_path)
            doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
            doc['note_last_heartbeat_at'] = time.monotonic()
            doc['last_unread_count'] = self.get_unread_note_count(doc)
        except PermissionError as exc:
            self.show_filesystem_error("Code Notes", sidecar_path, exc)
        except OSError as exc:
            self.log_exception("sync single note to sidecar", exc)
            self.show_filesystem_error("Code Notes", sidecar_path, exc)

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
        note_text = self.trim_text(saved_note.get('text', ''), self.max_note_text_length)
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
        note_color = self.normalize_optional_metadata(saved_note.get('color'))
        if not note_color:
            if self.normalize_optional_metadata(saved_note.get('dissapproved_by')):
                note_color = 'red'
            elif self.normalize_optional_metadata(saved_note.get('approved_by')):
                note_color = 'green'
            else:
                note_color = 'yellow'
        return {
            'id': note_id or None,
            'start': start,
            'end': end,
            'text': note_text,
            'color': self.normalize_note_color(note_color),
            'author_id': self.trim_text(self.normalize_optional_metadata(saved_note.get('author_id')), 128),
            'author_label': self.trim_text(self.normalize_optional_metadata(saved_note.get('author_label')), self.max_note_name_length),
            'author_host': self.trim_text(self.normalize_optional_metadata(saved_note.get('author_host')), 128),
            'author_ip': self.trim_text(self.normalize_optional_metadata(saved_note.get('author_ip')), 64),
            'read_by': [str(editor_id).strip()[:128] for editor_id in read_by if str(editor_id).strip()][:64],
            'author_unread': bool(saved_note.get('author_unread', False)),
            'created_at': self.normalize_optional_metadata(saved_note.get('created_at')),
            'anchor_text': str(saved_note.get('anchor_text', ''))[:self.max_note_text_length],
            'anchor_line': anchor_line,
            'responses': self.sanitize_note_responses(saved_note.get('responses', [])),
        }

    def dedupe_notes_payload(self, notes):
        if not isinstance(notes, list):
            return []
        deduped = []
        by_id = {}
        for note in notes:
            sanitized_note = self.sanitize_note_payload(note)
            if sanitized_note is None:
                continue
            note_id = str(sanitized_note.get('id', '')).strip()
            if note_id:
                if note_id in by_id:
                    deduped[by_id[note_id]] = sanitized_note
                else:
                    by_id[note_id] = len(deduped)
                    deduped.append(sanitized_note)
            else:
                deduped.append(sanitized_note)
        return deduped

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
        doc['last_note_cycle_tag'] = None
        doc['context_note_tag'] = None
        self.hide_note_popup(doc)

    def write_shared_notes(self, sidecar_path, notes_payload, active_editors=0, editors=None):
        payload = {
            'notes': self.dedupe_notes_payload(notes_payload)
        }
        if not self.write_json_atomically(sidecar_path, payload, 'notepadx-notes-', 'write shared notes'):
            raise OSError(f"Could not write note sidecar: {sidecar_path}")
        self.cleanup_duplicate_sidecar_variants(sidecar_path)
        self.show_support_file(sidecar_path)

    def load_shared_notes(self, sidecar_path):
        variant_paths = [path for path in self.get_sidecar_variants(sidecar_path) if os.path.exists(path)]
        if not variant_paths:
            return {'active_editors': 0, 'editors': [], 'notes': []}
        try:
            variant_paths.sort(key=lambda path: (os.path.getmtime(path), path.lower()))
        except OSError:
            variant_paths.sort(key=lambda path: path.lower())

        original_notes = []
        embedded_editors = []
        dirty_payload = len(variant_paths) > 1
        for variant_path in variant_paths:
            payload = self.read_json_file(variant_path, "load shared notes", {'active_editors': 0, 'editors': [], 'notes': []})
            if not isinstance(payload, dict):
                dirty_payload = True
                continue
            notes = payload.get('notes', [])
            if isinstance(notes, list):
                original_notes.extend(notes)
            else:
                dirty_payload = True
            embedded = payload.get('editors', [])
            if isinstance(embedded, list):
                embedded_editors.extend(embedded)

        sanitized_notes = self.dedupe_notes_payload(original_notes)
        if sanitized_notes != original_notes:
            dirty_payload = True
        editors = self.prune_inactive_shared_editors(embedded_editors)
        if dirty_payload:
            try:
                self.write_shared_notes(sidecar_path, sanitized_notes, len(editors), editors)
                self.cleanup_duplicate_sidecar_variants(sidecar_path)
            except OSError as exc:
                self.log_exception("rewrite sanitized shared notes", exc)
        return {
            'active_editors': len(editors),
            'editors': editors,
            'notes': sanitized_notes
        }

    def write_shared_editors(self, sidecar_path, editors):
        sanitized_editors = self.prune_inactive_shared_editors(editors)
        payload = {
            'active_editors': len(sanitized_editors),
            'editors': sanitized_editors
        }
        if not self.write_json_atomically(sidecar_path, payload, 'notepadx-editors-', 'write shared editors'):
            raise OSError(f"Could not write editor sidecar: {sidecar_path}")
        self.cleanup_duplicate_sidecar_variants(sidecar_path)
        self.show_support_file(sidecar_path)

    def load_shared_editors(self, sidecar_path, fallback_payload=None):
        variant_paths = [path for path in self.get_sidecar_variants(sidecar_path) if os.path.exists(path)]
        if variant_paths:
            try:
                variant_paths.sort(key=lambda path: (os.path.getmtime(path), path.lower()), reverse=True)
            except OSError:
                variant_paths.sort(key=lambda path: path.lower(), reverse=True)
            merged_editors = []
            dirty_payload = len(variant_paths) > 1
            for variant_path in variant_paths:
                payload = self.read_json_file(variant_path, "load shared editors", {'active_editors': 0, 'editors': []})
                if not isinstance(payload, dict):
                    dirty_payload = True
                    continue
                editors = payload.get('editors', [])
                if isinstance(editors, list):
                    merged_editors.extend(editors)
                else:
                    dirty_payload = True
            editors = self.prune_inactive_shared_editors(merged_editors)
            if dirty_payload:
                try:
                    self.write_shared_editors(sidecar_path, editors)
                    self.cleanup_duplicate_sidecar_variants(sidecar_path)
                except OSError as exc:
                    self.log_exception("rewrite sanitized shared editors", exc)
            return {
                'active_editors': len(editors),
                'editors': editors
            }
        fallback_payload = fallback_payload if isinstance(fallback_payload, dict) else {}
        editors = self.prune_inactive_shared_editors(fallback_payload.get('editors', []))
        return {
            'active_editors': len(editors),
            'editors': editors
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
                doc['note_last_heartbeat_at'] = time.monotonic()
        except PermissionError as exc:
            self.show_filesystem_error("Code Notes", sidecar_path, exc)
        except OSError as exc:
            self.log_exception("persist doc notes", exc)
            self.show_filesystem_error("Code Notes", sidecar_path, exc)
        self.save_session()

    def restore_doc_notes(self, doc):
        if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        editors_sidecar_path = self.get_editors_sidecar_path(doc['file_path'])
        shared_payload = self.load_shared_notes(sidecar_path)
        shared_editors_payload = self.load_shared_editors(editors_sidecar_path, fallback_payload=shared_payload)
        saved_notes = shared_payload.get('notes', [])
        self.apply_shared_editors_to_doc(doc, shared_editors_payload)
        self.clear_doc_notes(doc)
        if not saved_notes:
            doc['note_sync_mtime'] = os.path.getmtime(sidecar_path) if os.path.exists(sidecar_path) else None
            doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
            doc['note_editors_signature'] = self.get_editors_sidecar_signature(editors_sidecar_path)
            return

        for saved_note in saved_notes:
            resolved_range = self.resolve_note_range(doc, saved_note)
            start, end = resolved_range
            note_id = saved_note.get('id')
            note_text = saved_note.get('text', '').strip()
            note_color = saved_note.get('color')
            author_id = saved_note.get('author_id')
            author_label = saved_note.get('author_label')
            read_by = saved_note.get('read_by', [])
            author_unread = saved_note.get('author_unread', False)
            created_at = saved_note.get('created_at')
            anchor_text = saved_note.get('anchor_text')
            anchor_line = saved_note.get('anchor_line')
            responses = saved_note.get('responses', [])
            if not start or not end or not note_text:
                continue
            try:
                note_tag = self.create_note_tag(
                    doc, start, end, note_text,
                    note_color=note_color,
                    note_id=note_id,
                    author_id=author_id,
                    author_label=author_label,
                    read_by=read_by,
                    author_unread=author_unread,
                    created_at=created_at,
                    anchor_text=anchor_text,
                    anchor_line=anchor_line,
                    responses=responses
                )
            except tk.TclError as exc:
                self.log_exception("restore note tag", exc)
                continue
        doc['note_sync_mtime'] = os.path.getmtime(sidecar_path) if os.path.exists(sidecar_path) else None
        doc['note_sync_signature'] = self.get_notes_sidecar_signature(sidecar_path)
        doc['note_editors_signature'] = self.get_editors_sidecar_signature(editors_sidecar_path)
        doc['last_unread_count'] = self.get_unread_note_count(doc)

    def register_doc_for_shared_notes(self, doc):
        if not doc.get('file_path') or doc.get('notes_registered'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        doc['notes_registered'] = True
        try:
            self.refresh_doc_shared_editor_state(doc, force_write=True)
        except OSError as exc:
            doc['notes_registered'] = False
            self.log_exception("register doc for shared notes", exc)
        self.update_status()

    def unregister_doc_from_shared_notes(self, doc):
        if not doc.get('file_path') or not doc.get('notes_registered'):
            return

        sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
        editors_sidecar_path = self.get_editors_sidecar_path(doc['file_path'])
        payload = self.load_shared_editors(editors_sidecar_path)
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
            self.write_shared_editors(editors_sidecar_path, editors)
            doc['note_editors_signature'] = self.get_editors_sidecar_signature(editors_sidecar_path)
        except OSError as exc:
            self.log_exception("unregister doc from shared notes", exc)
        self.update_status()

    def poll_shared_notes(self):
        try:
            now = time.monotonic()
            heartbeat_interval = max(0.25, self.note_editor_heartbeat_interval_ms / 1000.0)
            status_dirty = False
            for doc in list(self.documents.values()):
                try:
                    if not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
                        continue

                    notes_sidecar_path = self.get_notes_sidecar_path(doc['file_path'])
                    editors_sidecar_path = self.get_editors_sidecar_path(doc['file_path'])
                    current_note_signature = self.get_notes_sidecar_signature(notes_sidecar_path)
                    current_editors_signature = self.get_editors_sidecar_signature(editors_sidecar_path)

                    if current_note_signature != doc.get('note_sync_signature'):
                        previous_unread_count = self.get_unread_note_count(doc)
                        self.restore_doc_notes(doc)
                        current_unread_count = self.get_unread_note_count(doc)
                        if current_unread_count > previous_unread_count:
                            self.play_unread_note_sound()
                        doc['last_unread_count'] = current_unread_count
                        status_dirty = True
                    elif current_editors_signature != doc.get('note_editors_signature'):
                        shared_editors_payload = self.load_shared_editors(editors_sidecar_path)
                        self.apply_shared_editors_to_doc(doc, shared_editors_payload)
                        doc['note_editors_signature'] = current_editors_signature
                        status_dirty = True
                    if doc.get('notes_registered'):
                        try:
                            last_heartbeat = float(doc.get('note_last_heartbeat_at', 0.0) or 0.0)
                            if (now - last_heartbeat) >= heartbeat_interval:
                                self.refresh_doc_shared_editor_state(doc)
                                status_dirty = True
                        except OSError as exc:
                            self.log_exception("refresh shared editor state", exc)
                except Exception as exc:
                    self.log_exception("poll shared notes doc", exc)
            try:
                if status_dirty:
                    self.update_status()
            except Exception as exc:
                self.log_exception("poll shared notes status", exc)
        except Exception as exc:
            self.log_exception("poll shared notes", exc)
        finally:
            try:
                self.root.after(self.note_sync_interval_ms, self.poll_shared_notes)
            except Exception as exc:
                self.log_exception("reschedule poll shared notes", exc)

    def load_content_into_doc(self, doc, file_path):
        try:
            file_size = os.path.getsize(file_path)
        except OSError as exc:
            self.log_exception("load content into doc", exc)
            messagebox.showerror("Open Failed", f"Notepad-X could not open:\n{file_path}\n\n{exc}", parent=self.root)
            return False
        doc['encrypted_file'] = False
        doc['encryption_header'] = None
        doc['encryption_key'] = None
        doc['file_size_bytes'] = file_size
        doc['background_loading'] = False
        doc['background_load_kind'] = None
        doc['background_load_file_path'] = None
        doc.pop('pending_insert_content', None)
        doc['pending_insert_offset'] = 0

        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)

        try:
            encrypted_result = self.read_encrypted_text_file(file_path) if self.file_looks_encrypted(file_path) else None
            if encrypted_result is not None:
                plaintext, encryption_header, encryption_key = encrypted_result
                plaintext_size = len(plaintext.encode('utf-8'))
                doc['file_size_bytes'] = plaintext_size
                self.set_large_file_mode(doc, plaintext_size >= self.large_file_threshold_bytes)
                doc['preview_mode'] = False
                doc['virtual_mode'] = False
                self.insert_text_content(doc, plaintext)
                doc['encrypted_file'] = True
                doc['encryption_header'] = encryption_header
                doc['encryption_key'] = encryption_key
            else:
                self.set_large_file_mode(doc, file_size >= self.large_file_threshold_bytes)
                doc['preview_mode'] = self.is_probably_binary_file(file_path)
                doc['virtual_mode'] = (
                    file_size > self.max_editable_large_file_bytes and
                    not doc['preview_mode']
                )

            if doc['preview_mode']:
                with open(file_path, 'rb') as f:
                    preview_bytes = f.read(self.huge_file_preview_bytes)
                preview_text = preview_bytes.decode('utf-8', errors='replace')
                text.insert(tk.END, preview_text)
                doc['line_starts'] = None
                doc['total_file_lines'] = max(1, int(text.index('end-1c').split('.')[0]))
                doc['window_start_line'] = 1
                doc['window_end_line'] = doc['total_file_lines']
            elif doc['virtual_mode']:
                self.start_background_virtual_index(doc, file_path)
                self.refresh_tab_title(doc['frame'])
                if str(doc['frame']) == self.notebook.select():
                    self.update_status()
                return True
            else:
                if not doc['encrypted_file'] and file_size >= self.large_file_threshold_bytes:
                    self.start_background_text_load(doc, file_path)
                    self.refresh_tab_title(doc['frame'])
                    if str(doc['frame']) == self.notebook.select():
                        self.update_status()
                    return True
                if not doc['encrypted_file']:
                    self.insert_text_content(doc, '')
                    text.delete('1.0', tk.END)
                    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                        while True:
                            chunk = f.read(self.file_load_chunk_size)
                            if not chunk:
                                break
                            text.insert(tk.END, chunk)
                            if doc['large_file_mode']:
                                self.root.update_idletasks()
                    doc['total_file_lines'] = max(1, int(text.index('end-1c').split('.')[0]))
                    doc['window_end_line'] = doc['total_file_lines']
        except RuntimeError as exc:
            self.log_exception("load content into doc", exc)
            messagebox.showerror("Open Failed", str(exc), parent=self.root)
            return False
        except (OSError, UnicodeDecodeError, ValueError) as exc:
            self.log_exception("load content into doc", exc)
            messagebox.showerror("Open Failed", f"Notepad-X could not open:\n{file_path}\n\n{exc}", parent=self.root)
            return False

        text.edit_modified(False)
        text.mark_set(tk.INSERT, '1.0')
        text.tag_remove('sel', '1.0', tk.END)
        text.see('1.0')
        doc['last_insert_index'] = '1.0'
        doc['last_yview'] = 0.0
        doc['last_xview'] = 0.0
        self.update_doc_file_signature(doc)
        self.configure_syntax_highlighting(doc['frame'])
        self.restore_doc_notes(doc)
        self.register_doc_for_shared_notes(doc)
        self.refresh_tab_title(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        if str(doc['frame']) == self.notebook.select():
            self.update_status()
        return True

    def get_session_state(self):
        current_doc = self.get_current_doc()
        selected_tab_id = self.notebook.select()
        selected_tab_doc = self.documents.get(str(selected_tab_id)) if selected_tab_id else None
        selected_file = current_doc['file_path'] if current_doc and current_doc['file_path'] else None
        compare_file = None
        compare_base_file = None
        if self.compare_active and self.compare_source_tab:
            compare_doc = self.documents.get(self.compare_source_tab)
            if compare_doc and compare_doc.get('file_path') and os.path.exists(compare_doc['file_path']):
                compare_file = compare_doc['file_path']
            if selected_tab_doc and selected_tab_doc.get('file_path') and os.path.exists(selected_tab_doc['file_path']):
                compare_base_file = selected_tab_doc['file_path']
        open_files = []

        for tab_id in self.notebook.tabs():
            doc = self.documents.get(str(tab_id))
            if not doc:
                continue
            if doc['file_path'] and os.path.exists(doc['file_path']):
                open_files.append(doc['file_path'])

        open_files = list(dict.fromkeys(open_files))[:self.max_session_files]
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
            ][:self.max_session_files],
            'sound_enabled': bool(self.sound_enabled.get()),
            'status_bar_enabled': bool(self.status_bar_enabled.get()),
            'numbered_lines_enabled': bool(self.numbered_lines_enabled.get()),
            'autocomplete_enabled': bool(self.autocomplete_enabled.get()),
            'sync_page_navigation_enabled': bool(self.sync_page_navigation_enabled.get()),
            'edit_with_shell_enabled': bool(self.edit_with_shell_enabled.get()),
            'current_font_size': int(self.current_font_size),
            'syntax_theme': self.syntax_theme.get(),
            'locale_code': self.locale_code,
            'compare_file': compare_file,
            'compare_base_file': compare_base_file,
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
        recovery_tabs = []
        selected_recovery_key = None
        current_doc = self.get_current_doc()
        for doc in self.documents.values():
            content = doc['text'].get('1.0', 'end-1c')
            modified = bool(doc['text'].edit_modified())
            file_path = doc.get('file_path')
            if file_path:
                if not modified:
                    continue
                recovery_tabs.append({
                    'file_path': os.path.abspath(file_path),
                    'untitled_name': None,
                    'content': content,
                    'modified': True,
                })
                recovery_key = f"file:{os.path.abspath(file_path)}"
            else:
                if not content.strip() and not modified:
                    continue
                untitled_name = self.trim_text(doc.get('untitled_name') or self.next_untitled_name(), 120)
                recovery_tabs.append({
                    'file_path': None,
                    'untitled_name': untitled_name,
                    'content': content,
                    'modified': modified
                })
                recovery_key = f"untitled:{untitled_name}"
            if len(recovery_tabs) >= self.max_recovery_tabs:
                break
            if current_doc == doc:
                selected_recovery_key = recovery_key
        return {
            'recovery_tabs': recovery_tabs,
            'selected_recovery_key': self.trim_text(selected_recovery_key, 260),
            'timestamp': self.utc_timestamp()
        }

    def apply_recovery_content_to_doc(self, doc, content, modified):
        if not doc:
            return
        text = doc.get('text')
        if not text:
            return
        doc['suspend_modified_events'] = True
        try:
            text.delete('1.0', tk.END)
            text.insert('1.0', content)
            text.edit_modified(bool(modified))
        finally:
            doc['suspend_modified_events'] = False
        self.refresh_tab_title(doc['frame'])
        self.update_doc_file_signature(doc)

    def persist_recovery_state(self):
        self.recovery_job = None
        if self.isolated_session:
            return
        recovery = self.get_recovery_state()
        if not recovery['recovery_tabs']:
            if os.path.exists(self.recovery_path):
                try:
                    os.remove(self.recovery_path)
                except OSError as exc:
                    self.log_exception("remove recovery state", exc)
            return
        for attempt in range(2):
            try:
                if self.write_json_atomically(self.recovery_path, recovery, 'notepadx-recovery-', 'persist recovery state'):
                    return
                raise PermissionError(self.recovery_path)
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
        recovery = self.read_json_file(self.recovery_path, "restore recovery state", None)
        recovery = self.sanitize_recovery_payload(recovery)
        if recovery is None:
            return
        recovery_tabs = recovery.get('recovery_tabs', [])
        if not isinstance(recovery_tabs, list) or not recovery_tabs:
            return
        if not messagebox.askyesno("Recover Tabs", "Notepad-X found unsaved tabs from a previous crash. Restore them?", parent=self.root):
            return
        current_doc = self.get_current_doc()
        if current_doc and not current_doc.get('file_path') and not current_doc['text'].edit_modified() and not current_doc['text'].get('1.0', 'end-1c').strip():
            self.notebook.forget(current_doc['frame'])
            self.documents.pop(str(current_doc['frame']), None)
        selected_recovery_key = recovery.get('selected_recovery_key')
        selected_tab = None
        for recovered_tab in recovery_tabs:
            if not isinstance(recovered_tab, dict):
                continue
            content = str(recovered_tab.get('content', ''))
            modified = bool(recovered_tab.get('modified', True))
            file_path = recovered_tab.get('file_path')
            doc = None
            recovery_key = None
            if file_path:
                normalized_path = os.path.abspath(file_path)
                recovery_key = f"file:{normalized_path}"
                for existing_doc in self.documents.values():
                    existing_path = existing_doc.get('file_path')
                    if existing_path and os.path.abspath(existing_path) == normalized_path:
                        doc = existing_doc
                        break
                if doc is None and os.path.exists(normalized_path):
                    tab_id = self.create_tab(file_path=normalized_path, select=False)
                    doc = self.documents[str(tab_id)]
                    if not self.load_content_into_doc(doc, normalized_path):
                        try:
                            self.notebook.forget(doc['frame'])
                        except tk.TclError:
                            pass
                        self.documents.pop(str(tab_id), None)
                        doc = None
                if doc is not None:
                    self.apply_recovery_content_to_doc(doc, content, modified)
                else:
                    recovered_name = self.trim_text(
                        recovered_tab.get('untitled_name') or f"Recovered {os.path.basename(normalized_path)}",
                        120
                    )
                    tab_id = self.create_tab(content=content, select=False)
                    doc = self.documents[str(tab_id)]
                    doc['untitled_name'] = recovered_name or doc['untitled_name']
                    self.apply_recovery_content_to_doc(doc, content, modified)
                    recovery_key = f"untitled:{doc['untitled_name']}"
            else:
                tab_id = self.create_tab(content=content, select=False)
                doc = self.documents[str(tab_id)]
                doc['untitled_name'] = str(recovered_tab.get('untitled_name') or doc['untitled_name'])
                self.apply_recovery_content_to_doc(doc, content, modified)
                recovery_key = f"untitled:{doc['untitled_name']}"
            self.refresh_tab_title(doc['frame'])
            if recovery_key == selected_recovery_key:
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

        for attempt in range(2):
            try:
                if self.write_json_atomically(self.session_path, session, 'notepadx-session-', 'save session'):
                    return
                raise PermissionError(self.session_path)
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

        session = self.read_json_file(self.session_path, "restore session", None)
        session = self.sanitize_session_payload(session)
        if session is None:
            return

        open_files = list(session.get('open_files', []))
        self.closed_session_files = set(session.get('closed_session_files', []))
        open_files = [path for path in open_files if path not in self.closed_session_files]
        self.recent_files = list(session.get('recent_files', []))[:self.max_recent_files]
        self.sound_enabled.set(bool(session.get('sound_enabled', True)))
        self.status_bar_enabled.set(bool(session.get('status_bar_enabled', True)))
        self.numbered_lines_enabled.set(bool(session.get('numbered_lines_enabled', True)))
        self.autocomplete_enabled.set(bool(session.get('autocomplete_enabled', True)))
        self.sync_page_navigation_enabled.set(bool(session.get('sync_page_navigation_enabled', False)))
        saved_edit_with_shell = bool(session.get('edit_with_shell_enabled', False))
        shell_registered = self.is_edit_with_shell_registered()
        self.edit_with_shell_enabled.set(saved_edit_with_shell or shell_registered)
        if saved_edit_with_shell and not shell_registered:
            self.sync_edit_with_shell_menu(show_errors=False)
        self.apply_locale(session.get('locale_code', self.locale_code), persist=False)
        saved_font_size = session.get('current_font_size', self.base_font_size)
        try:
            self.current_font_size = max(self.min_font_size, min(self.max_font_size, int(saved_font_size)))
        except (TypeError, ValueError):
            self.current_font_size = self.base_font_size
        saved_theme = str(session.get('syntax_theme', 'Default'))
        if saved_theme not in self.get_available_syntax_theme_names():
            saved_theme = 'Default'
        self.syntax_theme.set(saved_theme)
        for doc in self.documents.values():
            self.apply_syntax_tag_colors(doc['text'])
            self.update_line_number_gutter(doc)
        if getattr(self, 'compare_view', None):
            self.apply_syntax_tag_colors(self.compare_text)
            self.update_line_number_gutter(self.compare_view)
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
            if not self.load_content_into_doc(doc, file_path):
                try:
                    self.notebook.forget(doc['frame'])
                except tk.TclError:
                    pass
                self.documents.pop(str(tab_id), None)
                continue
            restored_tabs[file_path] = doc['frame']

        selected_file = session.get('selected_file')
        compare_base_file = session.get('compare_base_file')
        compare_file = session.get('compare_file')

        primary_file = selected_file
        if isinstance(compare_base_file, str) and compare_base_file in restored_tabs:
            primary_file = compare_base_file

        selected_tab = restored_tabs.get(primary_file)
        if selected_tab is None and restored_tabs:
            selected_tab = next(iter(restored_tabs.values()))

        if selected_tab is not None:
            self.notebook.select(selected_tab)
            self.set_active_document(selected_tab)

        compare_tab = restored_tabs.get(compare_file) if isinstance(compare_file, str) else None
        if compare_tab is not None and compare_tab != selected_tab:
            compare_doc = self.documents.get(str(compare_tab))
            if compare_doc:
                self.start_inline_compare(compare_doc)

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
            self.recent_menu.add_command(label=self.tr('common.empty', '(Empty)'), state='disabled')
            return

        for file_path in self.recent_files:
            self.recent_menu.add_command(
                label=os.path.basename(file_path),
                command=lambda path=file_path: self.open_recent_file(path)
            )

        self.recent_menu.add_separator()
        self.recent_menu.add_command(label=self.tr('common.clear_list', 'Clear list'), command=self.clear_recent_files)

    def refresh_language_menu(self):
        if not hasattr(self, 'language_menu') or self.language_menu is None:
            return
        self.language_menu.delete(0, tk.END)
        for language_code in self.get_available_language_codes():
            self.language_menu.add_radiobutton(
                label=self.get_language_display_name(language_code),
                variable=self.language_selection,
                value=language_code,
                command=lambda code=language_code: self.apply_locale(code)
            )

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
        compare_widget = self.get_compare_text_widget()
        focus_widget = self.safe_focus_get()
        last_widget = getattr(self, 'last_active_editor_widget', None)
        if compare_widget is not None and (focus_widget == compare_widget or last_widget == compare_widget):
            if self.compare_source_tab:
                compare_doc = self.documents.get(self.compare_source_tab)
                if compare_doc:
                    return compare_doc
        current_tab = self.notebook.select()
        return self.documents.get(current_tab)

    def set_active_document(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            self.text = None
            self.current_file = None
            self.syntax_mode_selection.set('auto')
            self.root.title(self.app_name)
            return
        self.text = doc['text']
        self.current_file = doc['file_path']
        self.syntax_mode_selection.set(doc.get('syntax_override') or 'auto')
        self.restore_doc_view_state(doc)
        self.update_line_number_gutter(doc)
        self.update_window_title()
        self.update_status()

    def on_tab_changed(self, event=None):
        self.hide_autocomplete_popup()
        previous_widget = getattr(self, 'last_active_editor_widget', None)
        if previous_widget is not None:
            previous_doc = self.get_doc_for_text_widget(previous_widget)
            if previous_doc:
                self.remember_doc_view_state(previous_doc)
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
            self.root.title(self.app_name)
            return
        title = self.get_doc_name(doc['frame'])
        if doc['text'].edit_modified():
            title += " *"
        self.root.title(f"{self.app_name} - {title}")

    def on_text_modified(self, tab_id):
        if str(tab_id) not in self.documents:
            return
        doc = self.documents[str(tab_id)]
        if doc.get('suspend_modified_events'):
            doc['text'].edit_modified(False)
            return
        self.remember_doc_view_state(doc)
        if not doc.get('file_path'):
            self.configure_syntax_highlighting(tab_id)
        if doc.get('syntax_mode') and doc.get('syntax_mode') != 'python':
            self.schedule_syntax_highlight(doc)
        self.update_line_number_gutter(doc)
        self.refresh_tab_title(tab_id)
        self.schedule_recovery_save()
        if self.compare_active and self.compare_source_tab == str(tab_id) and not self.compare_view.get('pushing_to_source'):
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

        current_doc = self.documents.get(str(current_tab))
        if current_doc:
            self.remember_doc_view_state(current_doc)

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
        t = self.tr
        self.menu = tk.Menu(self.root, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a', activeforeground='white')
        self.root.config(menu=self.menu)

        file_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.file', 'File'), menu=file_menu)
        file_menu.add_command(label=t('menu.file.open', 'Open'), command=self.open_file, accelerator=t('accel.open', 'Ctrl+W'))
        file_menu.add_command(label=t('menu.file.open_project', 'Open Project'), command=self.open_project, accelerator=t('accel.open_project', 'Ctrl+Shift+W'))
        self.recent_menu = tk.Menu(file_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                                   activebackground='#3a3a3a')
        file_menu.add_cascade(label=t('menu.file.recent', 'Recent'), menu=self.recent_menu)
        self.refresh_recent_files_menu()
        file_menu.add_command(label=t('menu.file.new_tab', 'New Tab'), command=self.new_tab, accelerator=t('accel.new_tab', 'Ctrl+T'))
        file_menu.add_command(label=t('menu.file.close_tab', 'Close Tab'), command=self.close_current_tab, accelerator=t('accel.close_tab', 'Ctrl+Shift+T'))
        file_menu.add_command(label=t('menu.file.save', 'Save'), command=self.save, accelerator=t('accel.save', 'Ctrl+S'))
        file_menu.add_command(label=t('menu.file.save_all', 'Save All'), command=self.save_all, accelerator=t('accel.save_all', 'Ctrl+Shift+S'))
        file_menu.add_command(label=t('menu.file.save_as', 'Save As'), command=self.save_copy_as, accelerator=t('accel.save_as', 'Ctrl+Shift+Q'))
        file_menu.add_command(label=t('menu.file.save_and_run', 'Save and Run'), command=self.save_and_run, accelerator=t('accel.save_and_run', 'Ctrl+Shift+R'))
        file_menu.add_command(label=t('menu.file.save_as_encrypted', 'Save As Encrypted'), command=self.save_encrypted_copy, accelerator=t('accel.save_as_encrypted', 'Ctrl+Shift+E'))
        file_menu.add_command(label=t('menu.file.print', 'Print'), command=self.print_file, accelerator=t('accel.print', 'Ctrl+P'))
        file_menu.add_command(label=t('menu.file.export_notes', 'Export Notes'), command=self.export_notes_report, accelerator=t('accel.export_notes', 'Ctrl+E'))
        file_menu.add_separator()
        file_menu.add_command(label=t('menu.file.exit', 'Exit'), command=self.exit_app, accelerator=t('accel.exit', 'Ctrl+Shift+X'))

        edit_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.edit', 'Edit'), menu=edit_menu)
        edit_menu.add_command(label=t('menu.edit.undo', 'Undo'), command=self.undo, accelerator=t('accel.undo', 'Ctrl+Z'))
        edit_menu.add_command(label=t('menu.edit.redo', 'Redo'), command=self.redo, accelerator=t('accel.redo', 'Ctrl+Shift+Z'))
        edit_menu.add_separator()
        edit_menu.add_command(label=t('menu.edit.cut', 'Cut'),  command=self.cut,  accelerator=t('accel.cut', 'Ctrl+X'))
        edit_menu.add_command(label=t('menu.edit.copy', 'Copy'), command=self.copy, accelerator=t('accel.copy', 'Ctrl+C'))
        edit_menu.add_command(label=t('menu.edit.paste', 'Paste'), command=self.paste, accelerator=t('accel.paste', 'Ctrl+V'))
        edit_menu.add_command(label=t('menu.edit.select_all', 'Select All'), command=self.select_all, accelerator=t('accel.select_all', 'Ctrl+A'))
        edit_menu.add_separator()
        edit_menu.add_command(label=t('menu.edit.find', 'Find'), command=self.show_find_panel, accelerator=t('accel.find', 'Ctrl+F'))
        edit_menu.add_command(label=t('menu.edit.find_next', 'Find Next'), command=self.find_next, accelerator=t('accel.find_next', 'F3'))
        edit_menu.add_command(label=t('menu.edit.find_previous', 'Find Previous'), command=self.find_previous, accelerator=t('accel.find_previous', 'Shift+F3'))
        edit_menu.add_command(label=t('menu.edit.replace', 'Replace'), command=self.show_replace_panel, accelerator=t('accel.replace', 'Ctrl+R'))
        edit_menu.add_separator()
        edit_menu.add_command(label=t('menu.edit.date', 'Date'), command=self.insert_date, accelerator=t('accel.date', 'Ctrl+D'))
        edit_menu.add_command(label=t('menu.edit.time_date', 'Time/Date'), command=self.insert_time_date, accelerator=t('accel.time_date', 'Ctrl+Shift+D'))
        edit_menu.add_command(label=t('menu.edit.font', 'Font'), command=self.show_font_dialog, accelerator=t('accel.font', 'Ctrl+Shift+F'))
        edit_menu.add_checkbutton(label=t('menu.view.edit_with_notepadx', 'Edit with Notepad-X'), variable=self.edit_with_shell_enabled, command=self.toggle_edit_with_shell)
        edit_menu.add_checkbutton(label=t('menu.view.sound', 'Sound'), variable=self.sound_enabled, command=self.toggle_sound)
        self.language_menu = tk.Menu(edit_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a', postcommand=self.refresh_language_menu)
        edit_menu.add_cascade(label=t('menu.edit.language', 'Language'), menu=self.language_menu)
        self.refresh_language_menu()
        view_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.view', 'View'), menu=view_menu)
        view_menu.add_command(label=t('menu.view.full_screen', 'Full Screen'), command=self.toggle_fullscreen, accelerator=t('accel.full_screen', 'F11'))
        view_menu.add_command(label=t('menu.view.switch_tab', 'Switch Tab'), command=self.switch_tab_right, accelerator=t('accel.switch_tab', 'Ctrl+Tab'))
        view_menu.add_checkbutton(label=t('menu.view.status_bar', 'Status Bar'), variable=self.status_bar_enabled, command=self.toggle_status_bar, accelerator=t('accel.status_bar', 'Ctrl+B'))
        view_menu.add_checkbutton(label=t('menu.view.numbered_lines', 'Numbered Lines'), variable=self.numbered_lines_enabled, command=self.toggle_numbered_lines)
        view_menu.add_checkbutton(label=t('menu.view.autocomplete', 'Autocomplete'), variable=self.autocomplete_enabled, command=self.toggle_autocomplete)
        view_menu.add_checkbutton(label=t('menu.view.word_wrap', 'Word Wrap'), variable=self.word_wrap_enabled, command=self.toggle_word_wrap)
        view_menu.add_command(label=t('menu.view.currently_editing', 'Currently Editing'), command=self.toggle_currently_editing_panel, accelerator=t('accel.currently_editing', 'Ctrl+Shift+C'))
        view_menu.add_command(label=t('menu.edit.cycle_notes', 'Cycle Notes'), command=self.goto_next_note, accelerator=t('accel.cycle_notes', 'F4'))
        note_filter_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label=t('menu.edit.filter_notes', 'Filter Notes'), menu=note_filter_menu)
        note_filter_menu.add_radiobutton(label=t('note.filter.all', 'All'), variable=self.note_filter, value='all')
        note_filter_menu.add_radiobutton(label=t('note.filter.unread', 'Unread'), variable=self.note_filter, value='unread')
        note_filter_menu.add_radiobutton(label=t('note.filter.yellow', 'Yellow'), variable=self.note_filter, value='yellow')
        note_filter_menu.add_radiobutton(label=t('note.filter.green', 'Green'), variable=self.note_filter, value='green')
        note_filter_menu.add_radiobutton(label=t('note.filter.red', 'Red'), variable=self.note_filter, value='red')
        note_filter_menu.add_radiobutton(label=t('note.filter.blue', 'Light Blue'), variable=self.note_filter, value='blue')
        view_menu.add_command(label=t('menu.edit.goto_line', 'Go To Line'), command=self.goto_line_dialog, accelerator=t('accel.goto_line', 'Ctrl+G'))
        view_menu.add_command(label=t('menu.edit.top_of_document', 'Top of Document'), command=self.goto_document_start, accelerator=t('accel.top_of_document', 'Ctrl+PgUp'))
        view_menu.add_command(label=t('menu.edit.bottom_of_document', 'Bottom of Document'), command=self.goto_document_end, accelerator=t('accel.bottom_of_document', 'Ctrl+PgDn'))
        view_menu.add_checkbutton(label=t('menu.edit.sync_page_navigation', 'Sync PgUp/PgDn in Compare'), variable=self.sync_page_navigation_enabled, command=self.save_session)
        syntax_theme_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label=t('menu.view.syntax_theme', 'Syntax Theme'), menu=syntax_theme_menu)
        syntax_theme_menu.add_command(label=t('menu.view.create_theme', 'Create Theme'), command=self.show_create_theme_dialog)
        syntax_theme_menu.add_separator()
        for theme_name in self.get_available_syntax_theme_names():
            syntax_theme_menu.add_radiobutton(
                label=self.get_syntax_theme_label(theme_name),
                variable=self.syntax_theme,
                value=theme_name,
                command=lambda name=theme_name: self.set_syntax_theme(name)
            )
        syntax_mode_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label=t('menu.view.syntax_mode', 'Syntax Mode'), menu=syntax_mode_menu)
        for mode_label, mode_value in (
            (t('syntax.mode.auto', 'Auto'), 'auto'), (t('syntax.mode.plain', 'Plain Text'), 'plain'), (t('syntax.mode.python', 'Python'), 'python'), (t('syntax.mode.c', 'C'), 'c'),
            (t('syntax.mode.cpp', 'C++'), 'cpp'), (t('syntax.mode.rust', 'Rust'), 'rust'), (t('syntax.mode.java', 'Java'), 'java'), (t('syntax.mode.javascript', 'JavaScript'), 'javascript'),
            (t('syntax.mode.html', 'HTML'), 'html'), (t('syntax.mode.php', 'PHP'), 'php'), (t('syntax.mode.xml', 'XML'), 'xml'), (t('syntax.mode.sql', 'SQL'), 'sql')
        ):
            syntax_mode_menu.add_radiobutton(
                label=mode_label,
                variable=self.syntax_mode_selection,
                value=mode_value,
                command=lambda value=mode_value: self.set_current_syntax_override(value)
            )
        view_menu.add_command(label=t('menu.view.compare_tabs', 'Compare Tabs'), command=self.show_split_compare, accelerator=t('accel.compare_tabs', 'Ctrl+Q'))
        view_menu.add_command(label=t('menu.view.close_compare_tabs', 'Close Compare Tabs'), command=self.close_compare_panel, accelerator=t('accel.close_compare_tabs', 'Ctrl+Shift+X'))

        help_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.help', 'Help'), menu=help_menu)
        help_menu.add_command(label=t('menu.help.contents', 'Help Contents'), command=self.show_help_contents)
        help_menu.add_command(label=t('menu.help.about', 'About Notepad-X'), command=self.show_about_dialog)

    def show_help_contents(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.tr('app.help_title', 'Notepad-X Help'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color)
        dialog.geometry("900x650")
        self.apply_window_icon(dialog)

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
                content = self.tr('help.open_failed', 'Unable to open the Notepad-X help file.')
        else:
            content = self.tr('help.not_found', 'Notepad-X help file not found.')

        help_text.insert('1.0', content)
        help_text.configure(state='disabled')

        close_button = tk.Button(
            dialog,
            text=self.tr('common.close', 'Close'),
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
            messagebox.showinfo(
                self.tr('app.compare_title', 'Compare With Tab'),
                self.tr('compare.need_two_tabs', 'Open at least two tabs to compare.'),
                parent=self.root
            )
            return "break"
        choices = []
        for tab_id, doc in self.documents.items():
            if doc is current_doc:
                continue
            choices.append((self.get_doc_name(doc['frame']), doc))
        dialog = tk.Toplevel(self.root)
        dialog.title(self.tr('app.compare_title', 'Compare With Tab'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color, padx=12, pady=12)
        self.apply_window_icon(dialog)
        tk.Label(
            dialog,
            text=self.tr('compare.choose_prompt', 'Choose a tab to compare with the current one:'),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor=self.ui_anchor_start(), pady=(0, 8))
        listbox = tk.Listbox(dialog, width=50, height=min(10, len(choices)))
        for label, _ in choices:
            listbox.insert(tk.END, label)
        listbox.pack(fill='both', expand=True)
        if choices:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.see(0)

        def open_compare(event=None):
            selection = listbox.curselection()
            if not selection:
                return
            other_doc = choices[selection[0]][1]
            dialog.destroy()
            self.start_inline_compare(other_doc)

        tk.Button(dialog, text=self.tr('common.compare', 'Compare'), command=open_compare).pack(pady=(10, 0))
        listbox.bind('<Double-Button-1>', open_compare)
        listbox.focus_set()
        self.center_window(dialog, self.root)
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current, self.root))

    def refresh_compare_header(self):
        if not self.compare_active:
            return
        doc = self.documents.get(self.compare_source_tab)
        if not doc:
            self.close_compare_panel()
            return
        self.compare_title.config(text=self.tr('compare.header', 'Comparing with: {title}', title=self.get_doc_title(doc['frame'])))

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
        try:
            compare_doc['last_insert_index'] = compare_text.index(tk.INSERT)
            compare_doc['last_yview'] = compare_text.yview()[0]
            compare_doc['last_xview'] = compare_text.xview()[0]
        except (tk.TclError, IndexError):
            pass

        compare_doc['suspend_modified_events'] = True
        compare_text.delete('1.0', tk.END)
        compare_text.insert('1.0', doc['text'].get('1.0', 'end-1c'))
        compare_text.edit_modified(False)
        compare_doc['suspend_modified_events'] = False

        insert_index = compare_doc.get('last_insert_index') or '1.0'
        try:
            compare_text.mark_set(tk.INSERT, insert_index)
        except tk.TclError:
            compare_text.mark_set(tk.INSERT, '1.0')
        try:
            compare_text.xview_moveto(float(compare_doc.get('last_xview', 0.0)))
        except (tk.TclError, TypeError, ValueError):
            pass
        try:
            compare_text.yview_moveto(float(compare_doc.get('last_yview', 0.0)))
        except (tk.TclError, TypeError, ValueError):
            pass
        self.configure_syntax_for_doc(compare_doc)
        self.sync_compare_note_tags(doc)
        self.update_line_number_gutter(compare_doc)
        self.update_status()

    def set_compare_sash_position(self):
        if not self.compare_active:
            return
        try:
            width = self.editor_paned.winfo_width()
            if width > 0:
                self.editor_paned.sash_place(0, max(240, width // 2), 0)
                self.position_compare_status()
        except tk.TclError:
            pass

    def get_currently_editing_sidebar_width(self):
        try:
            font_spec = self.currently_editing_content_label.cget('font')
            font = tkfont.Font(font=font_spec)
        except Exception:
            font = tkfont.Font(family='Consolas', size=10)
        # One full 32-char MD5 plus a trailing space, with tight sidebar padding.
        return max(160, font.measure(('0' * 32) + ' ') + 12)

    def ensure_currently_editing_sidebar_position(self):
        return

    def rebuild_editor_panes(self):
        try:
            self.editor_paned.forget(self.compare_container)
        except tk.TclError:
            pass
        try:
            if self.compare_active:
                self.editor_paned.add(self.compare_container, stretch='always')
        except tk.TclError:
            pass

    def set_currently_editing_sash_position(self):
        return

    def schedule_compare_layout_refresh(self):
        if not self.compare_active:
            return
        self.root.after_idle(self.set_compare_sash_position)
        self.root.after(10, self.set_compare_sash_position)
        self.root.after(50, self.set_compare_sash_position)
        self.root.after(100, self.set_compare_sash_position)
        self.root.after_idle(self.position_compare_status)
        self.root.after(10, self.position_compare_status)
        self.root.after(50, self.position_compare_status)

    def start_inline_compare(self, source_doc):
        if not source_doc:
            return "break"
        self.compare_source_tab = str(source_doc['frame'])
        self.compare_active = True
        self.rebuild_editor_panes()
        self.refresh_compare_panel()
        self.root.after_idle(self.set_compare_sash_position)
        self.root.after_idle(self.set_currently_editing_sash_position)
        self.update_status()
        self.save_session()
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
        self.rebuild_editor_panes()
        self.root.after_idle(self.set_currently_editing_sash_position)
        if hasattr(self, 'compare_status'):
            self.compare_status.place_forget()
        self.update_status()
        self.save_session()
        if self.text and self.text.winfo_exists():
            self.text.focus_set()
        return "break"

    def show_about_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.tr('app.about_title', 'About Notepad-X'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=24, pady=20)
        dialog.pong_after_id = None
        self.apply_window_icon(dialog)

        content = tk.Frame(dialog, bg=self.bg_color)
        content.pack()

        dialog.icon_image = self.create_about_display_image()
        icon_widget = None
        if dialog.icon_image is not None:
            icon_widget = tk.Label(content, image=dialog.icon_image, bg=self.bg_color, cursor='hand2')
            icon_widget.pack(pady=(0, 12))
        else:
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
            text=f"{self.tr('about.heading', 'Notepad-X')} {self.app_version}",
            bg=self.bg_color,
            fg=self.fg_color,
            font=('Segoe UI', 16, 'bold')
        ).pack(pady=(0, 16))

        repo_link = tk.Label(
            content,
            text=self.repo_url,
            bg=self.bg_color,
            fg='#58a6ff',
            font=('Segoe UI', 9, 'underline'),
            cursor='hand2'
        )
        repo_link.pack(pady=(0, 14))
        repo_link.bind('<Button-1>', lambda e: self.open_repo_link())

        tk.Label(
            content,
            text=self.tr('about.tagline', 'Built because Microsoft forgot what Notepad was supposed to be.'),
            bg=self.bg_color,
            fg='#9aa0a6',
            font=('Segoe UI', 9)
        ).pack(pady=(0, 16))

        tk.Button(
            dialog,
            text=self.tr('common.close', 'Close'),
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

    def open_repo_link(self):
        try:
            webbrowser.open(self.repo_url)
        except Exception as exc:
            self.log_exception("open repo link", exc)

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
            if self.is_notepadx_support_file(candidate_name):
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
        file_path = os.path.abspath(file_path)

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
            current_doc['background_open_new_tab'] = False
            if not self.load_content_into_doc(current_doc, file_path):
                current_doc['file_path'] = None
                return False
            current_doc['background_open_new_tab'] = False
            self.set_active_document(current_doc['frame'])
        else:
            tab_id = self.create_tab(file_path=file_path, select=True)
            new_doc = self.documents[str(tab_id)]
            new_doc['background_open_new_tab'] = True
            if not self.load_content_into_doc(new_doc, file_path):
                try:
                    self.notebook.forget(new_doc['frame'])
                except tk.TclError:
                    pass
                self.documents.pop(str(tab_id), None)
                return False
            if not new_doc.get('background_loading'):
                new_doc['background_open_new_tab'] = False
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
        target_mode = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode)
            except OSError:
                target_mode = 0o664
        fd, temp_path = tempfile.mkstemp(prefix='notepadx-save-', suffix='.tmp', dir=directory)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as temp_file:
                temp_file.write(content)
                temp_file.flush()
                os.fsync(temp_file.fileno())
            os.replace(temp_path, file_path)
            if target_mode is not None:
                os.chmod(file_path, target_mode)
            return True
        except OSError as exc:
            self.log_exception("write file atomically", exc)
            try:
                with open(file_path, 'w', encoding='utf-8') as direct_file:
                    direct_file.write(content)
                    direct_file.flush()
                    os.fsync(direct_file.fileno())
                if target_mode is not None:
                    os.chmod(file_path, target_mode)
                return True
            except OSError as direct_exc:
                self.log_exception("write file direct fallback", direct_exc)
                raise
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError as exc:
                    self.log_exception("cleanup temp save file", exc)

    def get_save_filetypes(self, include_encrypted=False):
        filetypes = [
            ("All Supported", "*.txt *.md .gitignore *.py *.pyw *.c *.cpp *.cxx *.cc *.h *.hpp *.hxx *.hh *.cs *.rs *.java *.js *.html *.htm *.php *.xml *.sql *.css *.json *.ini *.bat *.cmd *.sh *.asm *.s *.tex *.vb *.vbs *.pas *.pl *.pm *.diff *.patch *.nsi *.nsh *.iss *.rc *.as *.mx *.asp *.aspx *.au3 *.ml *.mli *.sml *.thy *.for *.f *.f90 *.f95 *.f2k *.lsp *.lisp *.mak *.m *.nfo *.st *.xsd *.xsml *.xsl *.kml"),
            ("Text Document", "*.txt"),
            ("Markdown", "*.md"),
            ("Git Ignore", ".gitignore"),
            ("Python", "*.py *.pyw"),
            ("C / Headers", "*.c *.h"),
            ("C++ / Headers", "*.cpp *.cxx *.cc *.hpp *.hxx *.hh"),
            ("C#", "*.cs"),
            ("Rust", "*.rs"),
            ("Java", "*.java"),
            ("JavaScript", "*.js"),
            ("HTML", "*.html *.htm"),
            ("PHP", "*.php *.php3 *.phtml"),
            ("XML", "*.xml *.xsd *.xsml *.xsl *.kml"),
            ("SQL", "*.sql"),
            ("CSS", "*.css"),
            ("JSON", "*.json"),
            ("INI / Config", "*.ini *.inf *.reg *.url"),
            ("Batch", "*.bat *.cmd"),
            ("Shell", "*.sh *.bsh"),
            ("Assembly", "*.asm *.s"),
            ("Pascal", "*.pas *.inc"),
            ("Perl", "*.pl *.pm *.plx"),
            ("Diff / Patch", "*.diff *.patch"),
            ("VB / VBScript", "*.vb *.vbs"),
            ("ActionScript", "*.as *.mx"),
            ("ASP / ASPX", "*.asp *.aspx"),
            ("AutoIt", "*.au3"),
            ("Caml", "*.ml *.mli *.sml *.thy"),
            ("Fortran", "*.f *.for *.f90 *.f95 *.f2k"),
            ("Inno Setup", "*.iss"),
            ("Lisp", "*.lsp *.lisp"),
            ("Makefile", "*.mak"),
            ("Matlab", "*.m"),
            ("NFO", "*.nfo"),
            ("NSIS", "*.nsi *.nsh"),
            ("Resource", "*.rc"),
            ("Smalltalk", "*.st"),
            ("TeX", "*.tex"),
            ("All Files", "*.*"),
        ]
        if include_encrypted:
            return [("Notepad-X Encrypted", "*.npxe"), *filetypes]
        return filetypes

    def get_save_and_run_language(self, doc, file_path):
        extension = os.path.splitext(file_path or '')[1].lower()
        syntax_mode = self.get_syntax_mode(doc) if doc else 'plain'

        if extension in {'.py', '.pyw'} or syntax_mode == 'python':
            return 'python'
        if extension in {'.js', '.mjs', '.cjs'} or syntax_mode == 'javascript':
            return 'javascript'
        if extension in {'.php', '.php3', '.phtml'} or syntax_mode == 'php':
            return 'php'
        if extension in {'.bat', '.cmd'}:
            return 'batch'
        if extension in {'.ps1'}:
            return 'powershell'
        if extension in {'.sh', '.bsh'}:
            return 'shell'
        if extension in {'.html', '.htm'} or syntax_mode == 'html':
            return 'html'
        return None

    def get_save_and_run_language_name(self, language):
        display_names = {
            'python': 'Python',
            'javascript': 'JavaScript',
            'php': 'PHP',
            'batch': 'Batch',
            'powershell': 'PowerShell',
            'shell': 'Shell',
            'html': 'HTML',
        }
        return display_names.get(language, str(language or '').title() or 'Unknown')

    def choose_available_command(self, candidate_commands):
        for command_args in candidate_commands:
            if not command_args:
                continue
            executable = command_args[0]
            if os.path.isabs(executable) and os.path.exists(executable):
                return command_args
            if shutil.which(executable):
                return command_args
        return None

    def get_save_and_run_command(self, language, file_path):
        if language == 'python':
            if self.is_windows:
                return self.choose_available_command((['py', '-3', file_path], ['python', file_path], ['python3', file_path]))
            return self.choose_available_command((['python3', file_path], ['python', file_path]))
        if language == 'javascript':
            return self.choose_available_command((['node', file_path],))
        if language == 'php':
            return self.choose_available_command((['php', file_path],))
        if language == 'powershell':
            if self.is_windows:
                return self.choose_available_command((
                    ['pwsh', '-NoExit', '-File', file_path],
                    ['powershell', '-NoExit', '-ExecutionPolicy', 'Bypass', '-File', file_path],
                ))
            return self.choose_available_command((['pwsh', '-NoExit', '-File', file_path],))
        if language == 'shell':
            if self.is_windows:
                return self.choose_available_command((['bash', file_path], ['sh', file_path]))
            return self.choose_available_command((['bash', file_path], ['sh', file_path]))
        return None

    def launch_terminal_command(self, command_args, cwd):
        if self.is_windows:
            command_line = subprocess.list2cmdline(command_args)
            creationflags = getattr(subprocess, 'CREATE_NEW_CONSOLE', 0)
            subprocess.Popen(['cmd.exe', '/k', command_line], cwd=cwd, creationflags=creationflags)
            return
        subprocess.Popen(command_args, cwd=cwd)

    def save_and_run(self, event=None):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        if doc.get('preview_mode') or doc.get('virtual_mode'):
            messagebox.showinfo(
                self.tr('run.title', 'Save and Run'),
                "Save and Run is not available for buffered large-file tabs.",
                parent=self.root
            )
            return "break"
        if not self.save():
            return "break"

        doc = self.get_current_doc()
        file_path = os.path.abspath(doc.get('file_path') or '')
        if not file_path:
            return "break"
        if not self.path_looks_safe_for_shell(file_path):
            messagebox.showerror(
                self.tr('run.title', 'Save and Run'),
                "That file path could not be sent to a runtime safely.",
                parent=self.root
            )
            return "break"

        language = self.get_save_and_run_language(doc, file_path)
        if not language:
            messagebox.showinfo(
                self.tr('run.title', 'Save and Run'),
                self.tr('run.unsupported', 'Save and Run is not available for this file type yet.'),
                parent=self.root
            )
            return "break"

        if language == 'html':
            try:
                webbrowser.open_new_tab(Path(file_path).resolve().as_uri())
            except Exception as exc:
                self.log_exception("save and run html", exc)
                messagebox.showerror(
                    self.tr('run.title', 'Save and Run'),
                    self.tr('run.open_browser_failed', 'Notepad-X could not open this file in your browser.'),
                    parent=self.root
                )
            return "break"

        if language == 'batch':
            creationflags = getattr(subprocess, 'CREATE_NEW_CONSOLE', 0)
            try:
                subprocess.Popen(
                    ['cmd.exe', '/k', f'call {subprocess.list2cmdline([file_path])}'],
                    cwd=os.path.dirname(file_path) or None,
                    creationflags=creationflags
                )
            except OSError as exc:
                self.log_exception("save and run batch", exc)
                messagebox.showerror(self.tr('run.title', 'Save and Run'), str(exc), parent=self.root)
            return "break"

        command_args = self.get_save_and_run_command(language, file_path)
        if not command_args:
            messagebox.showerror(
                self.tr('run.title', 'Save and Run'),
                self.tr(
                    'run.runtime_missing',
                    'Notepad-X could not find a runtime for {language} on this system.',
                    language=self.get_save_and_run_language_name(language)
                ),
                parent=self.root
            )
            return "break"

        try:
            self.launch_terminal_command(command_args, os.path.dirname(file_path) or None)
        except OSError as exc:
            self.log_exception("save and run", exc)
            messagebox.showerror(self.tr('run.title', 'Save and Run'), str(exc), parent=self.root)
        return "break"

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
            if not self.confirm_external_file_change(doc):
                return False
            try:
                text_content = doc['text'].get('1.0', tk.END).rstrip('\n')
                if doc.get('encrypted_file'):
                    self.write_encrypted_text_file(
                        doc['file_path'],
                        text_content,
                        header=doc.get('encryption_header'),
                        key=doc.get('encryption_key'),
                        original_name=doc.get('file_path')
                    )
                else:
                    self.write_file_atomically(doc['file_path'], text_content)
            except PermissionError as exc:
                self.show_filesystem_error("Save Failed", doc['file_path'], exc)
                return False
            except RuntimeError as exc:
                messagebox.showerror("Save Failed", str(exc), parent=self.root)
                return False
            except ValueError as exc:
                messagebox.showerror("Save Failed", str(exc), parent=self.root)
                return False
            except OSError as exc:
                self.log_exception("save file", exc)
                self.show_filesystem_error("Save Failed", doc['file_path'], exc)
                return False
            doc['text'].edit_modified(False)
            self.update_doc_file_signature(doc)
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
        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save As",
            filetypes=self.get_save_filetypes()
        )
        if file_path:
            old_file_path = doc.get('file_path')
            if old_file_path and old_file_path != file_path:
                self.unregister_doc_from_shared_notes(doc)
            doc['file_path'] = os.path.abspath(file_path)
            doc['untitled_name'] = None
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
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save As",
            initialfile=suggested_name,
            filetypes=self.get_save_filetypes()
        )
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
        except PermissionError as exc:
            self.show_filesystem_error("Save Copy As", output_path, exc)
        except RuntimeError as exc:
            messagebox.showerror("Save Copy As", str(exc), parent=self.root)
        except ValueError as exc:
            messagebox.showerror("Save Copy As", str(exc), parent=self.root)
        except OSError as exc:
            self.log_exception("save copy as", exc)
            self.show_filesystem_error("Save Copy As", output_path, exc)
        return "break"

    def save_encrypted_copy(self):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        encryption_options = self.prompt_encryption_options(default_encrypt=True, parent=self.root)
        if encryption_options is None or not encryption_options.get('encrypt'):
            return "break"
        suggested_name = self.get_doc_name(doc['frame'])
        suggested_encrypted_name = suggested_name if suggested_name.lower().endswith('.npxe') else f"{suggested_name}.npxe"
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save Encrypted Copy",
            initialfile=suggested_encrypted_name,
            defaultextension=".npxe",
            filetypes=self.get_save_filetypes(include_encrypted=True)
        )
        if not output_path:
            return "break"
        if not output_path.lower().endswith('.npxe'):
            output_path = f"{output_path}.npxe"
        try:
            if doc.get('virtual_mode') or doc.get('preview_mode'):
                messagebox.showinfo(
                    "Save Encrypted Copy",
                    "Encryption is not available for buffered large-file or preview tabs.",
                    parent=self.root
                )
                return "break"
            self.write_encrypted_text_file(
                output_path,
                doc['text'].get('1.0', tk.END).rstrip('\n'),
                passphrase=encryption_options.get('passphrase'),
                original_name=suggested_name
            )
            messagebox.showinfo("Save Encrypted Copy", f"Encrypted copy saved to:\n{output_path}", parent=self.root)
        except PermissionError as exc:
            self.show_filesystem_error("Save Encrypted Copy", output_path, exc)
        except RuntimeError as exc:
            messagebox.showerror("Save Encrypted Copy", str(exc), parent=self.root)
        except ValueError as exc:
            messagebox.showerror("Save Encrypted Copy", str(exc), parent=self.root)
        except OSError as exc:
            self.log_exception("save encrypted copy", exc)
            self.show_filesystem_error("Save Encrypted Copy", output_path, exc)
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

        if not self.path_looks_safe_for_shell(doc['file_path']):
            messagebox.showerror("Print Failed", "That file path could not be sent to the print command safely.", parent=self.root)
            return "break"

        print_error = None
        if self.is_windows:
            try:
                if hasattr(os, 'startfile'):
                    os.startfile(doc['file_path'], 'print')
                    return "break"
                elif self.shell32:
                    result = self.shell32.ShellExecuteW(None, 'print', doc['file_path'], None, None, 0)
                    if result <= 32:
                        raise OSError(f"Windows print action failed with code {result}")
                    return "break"
            except OSError as exc:
                print_error = exc

            try:
                subprocess.Popen(
                    ['notepad.exe', '/p', doc['file_path']],
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
                return "break"
            except OSError as exc:
                if print_error is None:
                    print_error = exc
        elif self.is_linux:
            for command in (['lp', doc['file_path']], ['lpr', doc['file_path']]):
                if shutil.which(command[0]):
                    try:
                        subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return "break"
                    except OSError as exc:
                        if print_error is None:
                            print_error = exc
            if shutil.which('xdg-open'):
                try:
                    subprocess.Popen(['xdg-open', doc['file_path']], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    return "break"
                except OSError as exc:
                    if print_error is None:
                        print_error = exc

        if print_error is None:
            print_error = OSError("Print is not available on this platform.")
        self.log_exception("print file", print_error)
        messagebox.showerror("Print Failed", str(print_error), parent=self.root)
        return "break"

    def exit_app(self, event=None):
        for doc in list(self.documents.values()):
            if not self.confirm_close_tab(doc):
                return "break"
        for doc in list(self.documents.values()):
            self.unregister_doc_from_shared_notes(doc)
        self.stop_single_instance_server()
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
        focused_widget = self.safe_focus_get()
        target = focused_widget if isinstance(focused_widget, tk.Text) else self.text
        try:
            target.edit_undo()
        except tk.TclError:
            pass
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and target == compare_widget:
            self.on_compare_modified()
        else:
            doc = self.get_doc_for_text_widget(target)
            if doc:
                self.on_text_modified(doc['frame'])
        return "break"

    def redo(self, event=None):
        if self.current_doc_is_large_readonly():
            return "break"
        focused_widget = self.safe_focus_get()
        target = focused_widget if isinstance(focused_widget, tk.Text) else self.text
        try:
            target.edit_redo()
        except tk.TclError:
            pass
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and target == compare_widget:
            self.on_compare_modified()
        else:
            doc = self.get_doc_for_text_widget(target)
            if doc:
                self.on_text_modified(doc['frame'])
        return "break"

    def select_all(self, event=None):
        focused_widget = self.safe_focus_get()
        target = focused_widget if isinstance(focused_widget, tk.Text) else self.text
        if not target:
            return "break"
        target.tag_add('sel', '1.0', 'end-1c')
        target.mark_set(tk.INSERT, '1.0')
        target.see(tk.INSERT)
        self.set_last_active_editor_widget(target)
        self.update_status()
        return "break"

    def cut(self, event=None):
        target = None
        if event is not None and isinstance(getattr(event, 'widget', None), tk.Text):
            target = event.widget
        else:
            focused_widget = self.safe_focus_get()
            if isinstance(focused_widget, tk.Text):
                target = focused_widget
        if target is None and isinstance(self.text, tk.Text):
            target = self.text

        if target is None:
            return "break"

        compare_widget = self.get_compare_text_widget()
        doc = self.get_doc_for_text_widget(target)
        if doc and (doc.get('preview_mode') or doc.get('virtual_mode')):
            return "break"

        try:
            selection = target.get('sel.first', 'sel.last')
        except tk.TclError:
            return "break"

        self.root.clipboard_clear()
        self.root.clipboard_append(selection)
        self.root.update_idletasks()

        try:
            target.delete('sel.first', 'sel.last')
            target.edit_modified(True)
        except tk.TclError:
            return "break"

        self.set_last_active_editor_widget(target)
        if compare_widget is not None and target == compare_widget:
            self.on_compare_modified()
        elif doc:
            self.on_text_modified(doc['frame'])
        else:
            self.update_status()
        return "break"

    def copy(self, event=None):
        target = None
        if event is not None and isinstance(getattr(event, 'widget', None), tk.Text):
            target = event.widget
        else:
            focused_widget = self.safe_focus_get()
            if isinstance(focused_widget, tk.Text):
                target = focused_widget
        if target is None and isinstance(self.text, tk.Text):
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
        else:
            focused_widget = self.safe_focus_get()
            if isinstance(focused_widget, tk.Text):
                target = focused_widget
        if target is None and isinstance(self.text, tk.Text):
            target = self.text

        compare_widget = self.get_compare_text_widget()
        if target is None:
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
        if compare_widget is not None and target == compare_widget:
            self.on_compare_modified()
        elif doc:
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
        self.apply_window_icon(dialog)

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
        self.apply_window_icon(dialog)

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
    raw_args = sys.argv[1:]
    isolated_mode = '--isolated' in {arg.lower() for arg in raw_args}
    startup_files = [arg for arg in raw_args if arg.lower() != '--isolated']
    app_dir = get_notepadx_app_dir()
    if not isolated_mode and send_files_to_running_notepadx(app_dir, startup_files):
        sys.exit(0)
    NotepadX(isolated_session=isolated_mode, startup_files=startup_files)
