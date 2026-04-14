import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk, colorchooser
from concurrent.futures import ProcessPoolExecutor, as_completed
from collections import OrderedDict
import colorsys
import os
import sys
import ctypes
import json
import ast
import bisect
import builtins
import glob
import keyword
import re
import hashlib
import base64
import codecs
import secrets
import subprocess
import tempfile
import traceback
import time
import shutil
import stat
import socket
import threading
import multiprocessing
import webbrowser
import gc
import zlib
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from ctypes import wintypes
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import unquote, urlparse
from array import array

_null_streams = []
for _stream_name in ('stdout', 'stderr'):
    if getattr(sys, _stream_name, None) is None:
        _stream = open(os.devnull, 'w', encoding='utf-8', buffering=1)
        setattr(sys, _stream_name, _stream)
        _null_streams.append(_stream)

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

try:
    from spellchecker import SpellChecker
except ImportError:
    SpellChecker = None


class EncryptedFileOpenCancelled(OSError):
    pass


DEFAULT_LOCALE_STRINGS = {
    "app.name": "Notepad-X",
    "app.about_title": "About Notepad-X",
    "app.help_title": "Notepad-X Help",
    "app.compare_title": "Compare With Tab",
    "about.heading": "Notepad-X",
    "about.tagline": "Built because Microsoft forgot what Notepad was supposed to be.",
    "about.icon_placeholder": "[Icon]",
    "about.pong.title": "Pong-X",
    "about.pong.info": "Player 1 keys: W & S Player 2 keys: Up & Down, Press Up/Down once to start PVP. Press R to restart.",
    "about.pong.user": "User",
    "about.pong.computer": "Computer",
    "about.pong.player1": "Player 1",
    "about.pong.player2": "Player 2",
    "about.pong.score": "{left_label} {left_score}   {right_label} {right_score}",
    "common.close": "Close",
    "common.compare": "Compare",
    "common.clear_list": "Clear list",
    "common.empty": "(Empty)",
    "common.ok": "OK",
    "common.cancel": "Cancel",
    "common.save": "Save",
    "common.unknown": "Unknown",
    "common.go": "Go",
    "common.color": "Color:",
    "clipboard.line_copied": "Copied line {line_number} to clipboard",
    "clipboard.path_copied": "Copied to clipboard",
    "context.cut": "Cut",
    "context.copy": "Copy",
    "context.paste": "Paste",
    "context.select_all": "Select All",
    "context.add_note": "Add note",
    "context.respond": "Respond",
    "context.remove_note": "Remove note",
    "context.add_to_dictionary": "Add to Dictionary",
    "context.no_suggestions": "No suggestions",
    "menu.file": "File",
    "menu.file.open": "Open",
    "menu.file.open_project": "Open Project",
    "menu.file.grab_git": "Grab Git",
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
    "menu.edit.sync_page_navigation": "Sync PgUp/PgDn in Compare/Preview",
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
    "menu.view.spell_check": "Spell Check",
    "menu.view.edit_with_notepadx": "Edit with Notepad-X",
    "menu.view.word_wrap": "Word Wrap",
    "menu.view.sound": "Sound",
    "menu.view.preview_markdown": "Preview Markdown",
    "menu.view.syntax_theme": "Syntax Theme",
    "menu.view.create_theme": "Create Theme",
    "menu.view.syntax_mode": "Syntax Mode",
    "menu.view.currently_editing": "Currently Editing",
    "menu.view.compare_tabs": "Compare Tabs",
    "menu.view.close_compare_tabs": "Close Compare Tabs",
    "menu.settings": "Settings",
    "menu.settings.hotkeys": "Hotkey Settings",
    "menu.help": "Help",
    "menu.help.contents": "Help Contents",
    "menu.help.about": "About Notepad-X",
    "hotkey.dialog.title": "Hotkey Settings",
    "hotkey.dialog.instructions": "Select an action, then press the shortcut you want to assign.",
    "hotkey.dialog.actions": "Actions",
    "hotkey.dialog.current": "Current Shortcut:",
    "hotkey.dialog.new": "New Shortcut:",
    "hotkey.dialog.capture": "Press a shortcut",
    "hotkey.dialog.assign": "Assign",
    "hotkey.dialog.clear": "Clear",
    "hotkey.dialog.reset_action": "Reset Action",
    "hotkey.dialog.reset_all": "Reset All",
    "hotkey.dialog.unassigned": "Unassigned",
    "hotkey.dialog.ready": "Ready.",
    "hotkey.dialog.select_action": "Select an action to edit.",
    "hotkey.dialog.invalid": "Use a function key or a shortcut with Ctrl or Alt.",
    "hotkey.dialog.conflict_title": "Shortcut Already In Use",
    "hotkey.dialog.conflict_message": "{shortcut} is already assigned to {action_label}. Reassign it?",
    "hotkey.dialog.reset_all_title": "Reset All Hotkeys",
    "hotkey.dialog.reset_all_message": "Reset every hotkey to its default value?",
    "hotkey.dialog.captured": "Captured {shortcut}.",
    "hotkey.dialog.assigned": "Assigned {shortcut} to {action_label}.",
    "hotkey.dialog.cleared": "Cleared {action_label}.",
    "hotkey.dialog.reset": "Reset {action_label} to its default shortcut.",
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
    "note.color.title": "Note Color",
    "note.color.prompt": "Color:",
    "panel.currently_editing.title": "Currently Editing",
    "panel.currently_editing.unsaved": "Save this tab to track currently editing IDs.",
    "panel.currently_editing.none": "No active editing IDs for this file.",
    "status.initial": "Ln 1 of 1, Col 1 | 0 characters | UTF-8 | Normal",
    "status.resource_initial": " | CPU: 0.0% Memory: 0MB",
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
    "status.compare": "{line_label} {row} {of_label} {total_lines}, {col_label} {col} | {char_info}{mode_suffix}",
    "status.char_count": "{total_chars} {characters_label}",
    "status.selected_char_count": "{selected_count} {of_label} {total_chars} {characters_label}",
    "status.byte_count": "{total_bytes} {bytes_label}",
    "status.resource_usage": " | CPU: {cpu_percent}% Memory: {memory_mb}MB",
    "status.editor_id": " | ID: {editor_id}",
    "status.unread_tail": " | {unread_count} unread (F3 to view) | ({active_editors} editing)",
    "compare.need_two_tabs": "Open at least two tabs to compare.",
    "compare.choose_prompt": "Choose a tab to compare with the current one:",
    "compare.header": "Comparing with: {title}",
    "markdown.preview.header": "Markdown Preview: {title}",
    "markdown.preview.empty": "Nothing to preview.",
    "find.panel.find": "Find:",
    "find.panel.find_next": "Find Next",
    "find.panel.browse": "Browse",
    "find.panel.find_in_label": "Find In:",
    "find.panel.found_summary": "| Found: {count} {instance_word} of \"{query}\"",
    "find.panel.instance_singular": "instance",
    "find.panel.instance_plural": "instances",
    "find.in.title": "Find In Files",
    "find.in.choose_directory": "Choose a folder to search",
    "find.in.query_required": "Enter some text in Find In first.",
    "find.in.directory_required": "Choose a folder with Browse first.",
    "find.in.searching_prompt": "Searching:\n{directory}\n\nPlease wait...",
    "find.in.results_summary": "Found {match_count} {instance_word} across {file_count} matching file(s).\nDirectory: {directory}\nScanned {scanned_count} supported files.",
    "find.in.no_matches": "No matching files were found for \"{query}\" in:\n{directory}",
    "find.in.search_failed": "Notepad-X could not search:\n{directory}\n\n{error_detail}",
    "find.in.column.instances": "Instances",
    "find.in.column.file": "File",
    "find.in.open_selected": "Open Selected",
    "find.in.select_results": "Select one or more results to open.",
    "replace.panel.replace_with": "Replace with:",
    "replace.panel.replace_all": "Replace All",
    "replace_all.title": "Replace All",
    "replace_all.completed": "Replaced {count} occurrence(s).",
    "large_file.title": "Large File Mode",
    "large_file.find_unavailable": "Find is not available in buffered large-file mode yet.",
    "large_file.loading": "Loading large file...",
    "large_file.indexing": "Indexing large file...",
    "large_file.buffering_title": "Buffering Large File",
    "large_file.buffering_prompt": "Loading:\n{file_path}",
    "large_file.buffering_status": "{percent}% ({loaded} / {total})",
    "large_file.buffering_lines": "Buffered {line_count} lines",
    "large_file.replace_unavailable": "Replace is not available in buffered large-file mode.",
    "large_file.save_run_unavailable": "Save and Run is not available for buffered large-file tabs.",
    "large_file.save_disabled": "This tab is opened in buffered large-file mode. Editing and saving are disabled for the full file view.",
    "large_file.save_as_disabled": "Save As is disabled for buffered large-file tabs.",
    "large_file.encryption_disabled": "Encryption is not available for buffered large-file or preview tabs.",
    "spellcheck.unavailable_title": "Spell Check Unavailable",
    "spellcheck.unavailable_message": "Spell check needs pyspellchecker and the bundled English dictionary. Rebuild Notepad-X if the menu shows enabled but no words are marked.",
    "file.open_failed_title": "Open Failed",
    "file.open_failed_message": "Notepad-X could not open:\n{file_path}\n\n{error_detail}",
    "file.missing_title": "File Missing",
    "file.missing_message": "That file could not be found.",
    "file.changed_title": "File Changed on Disk",
    "file.changed_message": "This file changed on disk after it was opened.\n\nOverwrite the newer disk version with your current editor contents?",
    "file.reload_title": "Reload File",
    "file.reload_message": "Do you want Notepad-X to reload the file from disk instead?",
    "filesystem.access_error": "Notepad-X could not access:\n{location}\n\n{error_detail}",
    "filesystem.unknown_path": "that path",
    "save.title": "Save",
    "save.close_prompt": "Save changes to {file_name} before closing?",
    "save.failed_title": "Save Failed",
    "save.as_title": "Save As",
    "save.copy_title": "Save Copy As",
    "save.copy_saved": "Copy saved to:\n{output_path}",
    "save.encrypted_copy_title": "Save Encrypted Copy",
    "save.encrypted_copy_saved": "Encrypted copy saved to:\n{output_path}",
    "print.failed_title": "Print Failed",
    "print.unsafe_path": "That file path could not be sent to the print command safely.",
    "print.unavailable": "Print is not available on this platform.",
    "print.windows_failed_code": "Windows print action failed with code {code}",
    "app.crash_title": "Notepad-X Crash",
    "app.crash_message": "An unexpected error occurred.\nA crash log was written to:\n{crash_log_path}",
    "recover.tabs_title": "Recover Tabs",
    "recover.tabs_message": "Notepad-X found unsaved tabs from a previous crash. Restore them?",
    "help.open_failed": "Unable to open the Notepad-X help file.",
    "help.not_found": "Notepad-X help file not found.",
    "theme.save_failed_title": "Save Theme Failed",
    "code_notes.title": "Code Notes",
    "shell_integration.update_failed": "Notepad-X could not update the OS shell integration.\n\n{error_detail}",
    "shell_integration.generic_name": "Text Editor",
    "shell_integration.desktop_comment": "Edit text files with Notepad-X",
    "shell_integration.app_description": "Edit supported text and code files with Notepad-X.",
    "shell_integration.error.unsafe_windows_executable": "Unsafe executable path for Windows shell integration.",
    "shell_integration.error.unsafe_windows_script": "Unsafe interpreter or script path for Windows shell integration.",
    "shell_integration.error.unsafe_linux_executable": "Unsafe executable path for Linux desktop integration.",
    "shell_integration.error.unsafe_linux_script": "Unsafe interpreter or script path for Linux desktop integration.",
    "shell_integration.error.write_linux_desktop_entry": "Could not write Linux desktop entry to {desktop_entry_path}",
    "shell_integration.error.unsafe_icon_shell": "Unsafe icon path for Windows shell integration.",
    "shell_integration.error.unsafe_icon_registration": "Unsafe icon path for Windows application registration.",
    "export.notes.none": "There are no notes to export in this tab.",
    "export.notes.saved": "Notes exported to:\n{output_path}",
    "export.notes.markdown.title": "Notes Export for {doc_name}",
    "export.notes.markdown.note_heading": "Note {note_id}",
    "export.notes.markdown.range": "Range",
    "export.notes.markdown.color": "Color",
    "export.notes.markdown.selected_text_heading": "Selected Text",
    "export.notes.markdown.code_note_heading": "Code Note",
    "filetype.all_supported": "All Supported",
    "filetype.text_document": "Text Document",
    "filetype.markdown": "Markdown",
    "filetype.git_ignore": "Git Ignore",
    "filetype.python": "Python",
    "filetype.c_headers": "C / Headers",
    "filetype.cpp_headers": "C++ / Headers",
    "filetype.csharp": "C#",
    "filetype.rust": "Rust",
    "filetype.java": "Java",
    "filetype.javascript": "JavaScript",
    "filetype.html": "HTML",
    "filetype.php": "PHP",
    "filetype.xml": "XML",
    "filetype.sql": "SQL",
    "filetype.css": "CSS",
    "filetype.json": "JSON",
    "filetype.ini_config": "INI / Config",
    "filetype.batch": "Batch",
    "filetype.powershell": "PowerShell",
    "filetype.shell": "Shell",
    "filetype.assembly": "Assembly",
    "filetype.pascal": "Pascal",
    "filetype.perl": "Perl",
    "filetype.diff_patch": "Diff / Patch",
    "filetype.vb_vbscript": "VB / VBScript",
    "filetype.actionscript": "ActionScript",
    "filetype.asp_aspx": "ASP / ASPX",
    "filetype.autoit": "AutoIt",
    "filetype.caml": "Caml",
    "filetype.fortran": "Fortran",
    "filetype.inno_setup": "Inno Setup",
    "filetype.lisp": "Lisp",
    "filetype.makefile": "Makefile",
    "filetype.matlab": "Matlab",
    "filetype.nfo": "NFO",
    "filetype.nsis": "NSIS",
    "filetype.resource": "Resource",
    "filetype.smalltalk": "Smalltalk",
    "filetype.tex": "TeX",
    "filetype.all_files": "All Files",
    "filetype.encrypted": "Notepad-X Encrypted",
    "font.title": "Font",
    "font.family": "Font:",
    "font.size": "Size:",
    "font.preview": "AaBbYyZz 0123456789",
    "font.invalid_size_title": "Invalid Size",
    "font.invalid_size_message": "Choose a valid font size.",
    "goto_line.title": "Go To Line",
    "goto_line.prompt": "Line Number:",
    "goto_line.invalid_title": "Invalid",
    "goto_line.invalid_range": "Line number out of range.",
    "goto_line.invalid_number": "Enter a valid number.",
    "encryption.unavailable_title": "Encryption Unavailable",
    "encryption.unavailable_message": "Encrypted save/open needs the 'cryptography' Python package.\n\nInstall it on this machine to use Notepad-X encrypted files.",
    "encryption.save_title": "Save Encrypted Copy As",
    "encryption.encrypt_file": "Encrypt file",
    "encryption.passphrase": "Passphrase:",
    "encryption.confirm_passphrase": "Confirm passphrase:",
    "encryption.show_passphrase": "Show passphrase",
    "encryption.passphrase_required": "Enter a passphrase first.",
    "encryption.passphrase_mismatch": "Passphrases do not match.",
    "encryption.open_title": "Open Encrypted File",
    "encryption.open_prompt": "Passphrase for:\n{file_name}",
    "encryption.unlock_failed": "That passphrase did not unlock the encrypted file.",
    "encryption.error.cryptography_required": "The cryptography package is required for encrypted files.",
    "encryption.error.passphrase_required": "A passphrase is required.",
    "encryption.error.missing_salt": "Encrypted file is missing its salt.",
    "encryption.error.scrypt_memory": "Notepad-X could not derive the encryption key because the scrypt memory limit was exceeded on this machine.",
    "encryption.error.header_incomplete": "Encrypted file header is incomplete.",
    "encryption.error.header_invalid": "Encrypted file header is invalid.",
    "encryption.error.unsupported_settings": "Unsupported encrypted file settings.",
    "encryption.error.no_ciphertext": "Encrypted file has no ciphertext.",
    "encryption.error.support_unavailable": "Encryption support is unavailable.",
    "encryption.error.metadata_missing": "Encrypted file metadata is missing.",
    "encryption.error.write_failed": "Could not write encrypted file: {file_path}",
    "encryption.error.open_cancelled": "Encrypted file open cancelled.",
    "syntax.theme.default": "Default",
    "syntax.theme.soft": "Soft",
    "syntax.theme.vivid": "Vivid",
    "syntax.theme.base4tone": "Base4Tone",
    "syntax.theme.green_monochrome": "Green Monochrome",
    "syntax.theme.orange_monochrome": "Orange Monochrome",
    "syntax.theme.lolcat": "Lolcat",
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
    "accel.grab_git": "Ctrl+Shift+G",
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
    "accel.spell_check": "F7",
    "accel.preview_markdown": "Ctrl+Shift+P",
    "accel.currently_editing": "Ctrl+Shift+C",
    "accel.compare_tabs": "Ctrl+Q",
    "accel.close_compare_tabs": "Ctrl+Shift+X",
    "grab_git.title": "Grab Git",
    "grab_git.prompt": "Enter the GitHub project as:\nusername/project",
    "grab_git.invalid": "Enter the project as username/project.\n\nExample:\nopenai/openai-python",
    "grab_git.git_missing": "Git could not be found on this system.",
    "grab_git.choose_folder": "Choose where to save the GitHub project",
    "grab_git.exists": "That destination folder already exists.\nChoose another location or remove the existing folder first.",
    "grab_git.downloading": "Downloading Project",
    "grab_git.downloading_prompt": "Downloading:\n{repo_identifier}\n\nPlease wait...",
    "grab_git.clone_failed": "Notepad-X could not download that GitHub project.\n\n{error_detail}",
    "grab_git.not_found": "GitHub could not find \"{repo_identifier}\".\n\nCheck the username/project and try again.",
    "grab_git.private_repo": "That GitHub project could not be downloaded.\n\nIt may be private or require authentication.",
    "grab_git.no_files": "The project downloaded successfully, but no matching files were found to open.",
    "grab_git.unknown_failure": "Unknown git clone failure.",
    "grab_git.open_title": "Open Downloaded Project Files",
    "grab_git.open_prompt": "Downloaded project root:\n{root_dir}\n\nSelect one or more files to open.",
    "grab_git.select_files": "Select one or more files to open.",
    "grab_git.open_selected": "Open Selected",
    "run.title": "Save and Run",
    "run.unsupported": "Save and Run is not available for this file type yet.",
    "run.large_file_unavailable": "Save and Run is not available for buffered large-file tabs.",
    "run.unsafe_path": "That file path could not be sent to a runtime safely.",
    "run.runtime_missing": "Notepad-X could not find a runtime for {language} on this system.",
    "run.open_browser_failed": "Notepad-X could not open this file in your browser.",
    "locale.language.en": "English",
    "locale.language.ar": "Arabic",
    "locale.language.bn": "Bengali",
    "locale.language.de": "German",
    "locale.language.es": "Spanish",
    "locale.language.fr": "French",
    "locale.language.he": "Hebrew",
    "locale.language.hi": "Hindi",
    "locale.language.id": "Indonesian",
    "locale.language.it": "Italian",
    "locale.language.ja": "Japanese",
    "locale.language.nl": "Dutch",
    "locale.language.pt": "Portuguese",
    "locale.language.ru": "Russian",
    "locale.language.uk": "Ukrainian",
    "locale.language.zh": "Chinese",
    "locale.region.us": "US",
    "locale.region.sa": "Saudi Arabia",
    "locale.region.ae": "United Arab Emirates",
    "locale.region.bd": "Bangladesh",
    "locale.region.eg": "Egypt",
    "locale.region.ma": "Morocco",
    "locale.region.de": "Germany",
    "locale.region.419": "Latin America",
    "locale.region.es": "Spain",
    "locale.region.ca": "Canada",
    "locale.region.in": "India",
    "locale.region.il": "Israel",
    "locale.region.id": "Indonesia",
    "locale.region.it": "Italy",
    "locale.region.jp": "Japan",
    "locale.region.nl": "Netherlands",
    "locale.region.br": "Brazil",
    "locale.region.ru": "Russia",
    "locale.region.ua": "Ukraine",
    "locale.region.cn": "China"
}

RTL_LOCALE_CODES = {'ar', 'ar_sa', 'ar_ae', 'ar_eg', 'ar_ma', 'he', 'he_il'}

LOCALE_DISPLAY_NAMES = {
    'en_us': 'English (US)',
    'ar': 'العربية',
    'ar_sa': 'العربية (السعودية)',
    'ar_ae': 'العربية (الإمارات)',
    'ar_eg': 'العربية (مصر)',
    'ar_ma': 'العربية (المغرب)',
    'bn_bd': 'বাংলা (বাংলাদেশ)',
    'fr_ca': 'Français (Canada)',
    'he_il': 'עברית (ישראל)',
    'hi_in': 'हिन्दी (भारत)',
    'id_id': 'Bahasa Indonesia (Indonesia)',
    'ja_jp': '日本語 (日本)',
    'ru_ru': 'Русский (Россия)',
    'zh_cn': '简体中文 (中国)',
}

LANGUAGE_NATIVE_NAMES = {
    'en': 'English',
    'ar': 'العربية',
    'bn': 'বাংলা',
    'de': 'Deutsch',
    'es': 'Español',
    'fr': 'Français',
    'he': 'עברית',
    'hi': 'हिन्दी',
    'id': 'Bahasa Indonesia',
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
    'bd': 'বাংলাদেশ',
    'eg': 'مصر',
    'ma': 'المغرب',
    'de': 'Deutschland',
    '419': 'Latinoamérica',
    'es': 'España',
    'ca': 'Canada',
    'in': 'भारत',
    'il': 'ישראל',
    'id': 'Indonesia',
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


def get_notepadx_instance_scope_dir(app_dir):
    try:
        candidate_dir = os.path.abspath(str(app_dir))
    except Exception:
        return str(app_dir)
    parent_dir = os.path.dirname(candidate_dir)
    if os.path.basename(candidate_dir).lower() == 'output':
        if os.path.exists(os.path.join(parent_dir, 'Notepad-X.py')):
            return parent_dir
    return candidate_dir


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
        or candidate_name == 'notepad-x.startup.trace.log'
        or (candidate_name.startswith('notepad-x.') and candidate_name.endswith('.session.json'))
        or (candidate_name.startswith('notepad-x.') and candidate_name.endswith('.editor.json'))
    )


def get_windows_command_line_args():
    if os.name != 'nt':
        return []
    try:
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        shell32 = ctypes.WinDLL('shell32', use_last_error=True)
        kernel32.GetCommandLineW.restype = wintypes.LPWSTR
        shell32.CommandLineToArgvW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_int)]
        shell32.CommandLineToArgvW.restype = ctypes.POINTER(wintypes.LPWSTR)
        kernel32.LocalFree.argtypes = [ctypes.c_void_p]
        kernel32.LocalFree.restype = ctypes.c_void_p
        command_line = kernel32.GetCommandLineW()
        if not command_line:
            return []
        argc = ctypes.c_int(0)
        argv = shell32.CommandLineToArgvW(command_line, ctypes.byref(argc))
        if not argv:
            return []
        try:
            return [argv[index] for index in range(max(0, argc.value))]
        finally:
            kernel32.LocalFree(argv)
    except Exception:
        return []


def get_process_launch_arguments():
    collected_args = []
    seen_args = set()
    arg_sources = [list(sys.argv[1:])]
    windows_args = get_windows_command_line_args()
    if len(windows_args) > 1:
        arg_sources.append(list(windows_args[1:]))
    for source in arg_sources:
        for raw_arg in source:
            if not isinstance(raw_arg, str):
                continue
            value = raw_arg.strip()
            if not value:
                continue
            lowered = value.lower() if os.name == 'nt' else value
            if lowered in seen_args:
                continue
            seen_args.add(lowered)
            collected_args.append(value)
    return collected_args


def normalize_startup_path_argument(raw_path, base_dir=None):
    if raw_path is None:
        return None
    value = str(raw_path).strip()
    if not value:
        return None
    for _ in range(2):
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1].strip()
    if not value:
        return None
    lowered = value.lower()
    if lowered in {'--isolated', '/dde', '-dde', '/embedding', '-embedding'} or lowered.startswith('/prefetch:'):
        return None
    parsed = urlparse(value)
    if parsed.scheme.lower() == 'file':
        candidate_value = unquote(parsed.path or '')
        if parsed.netloc:
            candidate_value = f"//{parsed.netloc}{candidate_value}"
        if os.name == 'nt' and re.match(r'^/[A-Za-z]:', candidate_value):
            candidate_value = candidate_value[1:]
        value = candidate_value
    value = os.path.expandvars(os.path.expanduser(value))
    if not value:
        return None
    candidate_paths = []
    if os.path.isabs(value):
        candidate_paths.append(os.path.abspath(value))
    else:
        active_base_dir = base_dir or os.getcwd()
        candidate_paths.append(os.path.abspath(os.path.join(active_base_dir, value)))
        if active_base_dir != os.getcwd():
            candidate_paths.append(os.path.abspath(os.path.join(os.getcwd(), value)))
    for candidate_path in candidate_paths:
        if os.path.exists(candidate_path):
            return candidate_path
    return None


def get_notepadx_single_instance_port(app_dir):
    scope_dir = get_notepadx_instance_scope_dir(app_dir)
    seed = (
        f"NotepadX::{os.path.normcase(os.path.abspath(scope_dir))}::"
        f"{os.environ.get('USERNAME') or os.environ.get('USER') or 'user'}"
    )
    seed_hash = hashlib.sha256(seed.encode('utf-8')).hexdigest()
    return 43000 + (int(seed_hash[:8], 16) % 10000)


def send_files_to_running_notepadx(app_dir, startup_files):
    normalized_files = []
    seen_files = set()
    for raw_path in startup_files or []:
        candidate_path = normalize_startup_path_argument(raw_path)
        if not candidate_path:
            candidate_path = normalize_startup_path_argument(raw_path, base_dir=app_dir)
        if not candidate_path:
            continue
        candidate_key = os.path.normcase(candidate_path)
        if candidate_key in seen_files:
            continue
        seen_files.add(candidate_key)
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


def scan_large_text_file_index_range_worker(file_path, nominal_start, nominal_end, chunk_size=4 * 1024 * 1024, range_index=0):
    safe_chunk_size = max(64 * 1024, int(chunk_size or (4 * 1024 * 1024)))
    file_size = os.path.getsize(file_path)
    start_offset = max(0, min(file_size, int(nominal_start or 0)))
    end_offset = max(start_offset, min(file_size, int(nominal_end if nominal_end is not None else file_size)))

    if start_offset >= file_size or start_offset >= end_offset:
        return {'range_index': int(range_index or 0), 'line_lengths': []}

    def align_start(source_file, offset):
        if offset <= 0:
            return 0
        source_file.seek(offset)
        current = offset
        while current < file_size:
            block = source_file.read(min(safe_chunk_size, file_size - current))
            if not block:
                return file_size
            newline_index = block.find(b'\n')
            if newline_index != -1:
                return current + newline_index + 1
            current += len(block)
        return file_size

    def align_end(source_file, offset):
        if offset >= file_size:
            return file_size
        current = offset
        source_file.seek(current)
        while current < file_size:
            block = source_file.read(min(safe_chunk_size, file_size - current))
            if not block:
                return file_size
            newline_index = block.find(b'\n')
            if newline_index != -1:
                return current + newline_index + 1
            current += len(block)
        return file_size

    with open(file_path, 'rb') as source_file:
        aligned_start = align_start(source_file, start_offset)
        aligned_end = align_end(source_file, end_offset)
        if aligned_start >= aligned_end:
            return {'range_index': int(range_index or 0), 'line_lengths': []}

        source_file.seek(aligned_start)
        remaining = aligned_end - aligned_start
        line_lengths = array('I')
        current_line_length = 0
        remainder = b''

        while remaining > 0:
            chunk = source_file.read(min(safe_chunk_size, remaining))
            if not chunk:
                break
            remaining -= len(chunk)
            data = remainder + chunk
            parts = data.split(b'\n')
            if remaining == 0:
                remainder = parts.pop() if parts else b''
            else:
                remainder = parts.pop() if parts else b''
            if not parts and remainder and remaining > 0:
                current_line_length += len(remainder)
                remainder = b''
                continue
            if parts:
                line_lengths.append(current_line_length + len(parts[0]))
                if len(parts) > 1:
                    line_lengths.extend(len(part) for part in parts[1:])
                current_line_length = len(remainder)
            else:
                current_line_length += len(remainder)
                remainder = b''

        final_length = current_line_length + len(remainder)
        if final_length > 0 or aligned_end >= file_size:
            line_lengths.append(final_length)
    return {
        'range_index': int(range_index or 0),
        'line_lengths': line_lengths.tolist(),
    }


def scan_newline_start_offsets_range_worker(file_path, start_offset, end_offset, chunk_size=4 * 1024 * 1024, range_index=0):
    safe_chunk_size = max(64 * 1024, int(chunk_size or (4 * 1024 * 1024)))
    file_size = os.path.getsize(file_path)
    safe_start = max(0, min(file_size, int(start_offset or 0)))
    safe_end = max(safe_start, min(file_size, int(end_offset if end_offset is not None else file_size)))
    newline_starts = array('Q')
    if safe_start >= safe_end:
        return {
            'range_index': int(range_index or 0),
            'line_starts': [],
            'bytes_scanned': 0,
        }

    with open(file_path, 'rb') as source_file:
        source_file.seek(safe_start)
        absolute_offset = safe_start
        remaining = safe_end - safe_start
        while remaining > 0:
            chunk = source_file.read(min(safe_chunk_size, remaining))
            if not chunk:
                break
            search_from = 0
            while True:
                newline_index = chunk.find(b'\n', search_from)
                if newline_index == -1:
                    break
                newline_starts.append(absolute_offset + newline_index + 1)
                search_from = newline_index + 1
            consumed = len(chunk)
            absolute_offset += consumed
            remaining -= consumed

    return {
        'range_index': int(range_index or 0),
        'line_starts': newline_starts.tolist(),
        'bytes_scanned': max(0, safe_end - safe_start),
    }


class NotepadX:
    def __init__(self, isolated_session=False, startup_files=None):
        self.root = tk.Tk()
        try:
            self.root.withdraw()
        except tk.TclError:
            pass
        self._shutdown_requested = False
        self.main_loop_started = False
        self.root.title(DEFAULT_LOCALE_STRINGS.get('app.name', 'Notepad-X'))
        self.isolated_session = isolated_session
        self.startup_files = list(startup_files or [])
        self.init_config()
        self.root.title(self.app_name)
        self.init_runtime()
        self.init_ui()
        if not self._shutdown_requested:
            try:
                if self.root.winfo_exists():
                    self.main_loop_started = True
                    self.root.mainloop()
            except tk.TclError:
                pass

    def init_config(self):
        self.is_windows = os.name == 'nt'
        self.is_linux = sys.platform.startswith('linux')
        self.app_version = "v1.0.7"
        self.resource_dir = self.get_resource_dir()
        self.app_dir = self.get_app_dir()
        self.machine_profile_slug = self.get_machine_profile_slug()
        self.repo_url = "https://github.com/OPCraniX/Notepad-X"
        self.app_user_model_id = "OPCraniX.Notepad-X"
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
        self.spellcheck_tag = 'spellcheck_misspelled'
        self.documents = {}
        self.cpu_used_percent = 0.0
        self.memory_used_mb = 0
        self.cpu_sample_count = max(1, int(os.cpu_count() or 1))
        self.last_cpu_sample = None
        self.syntax_highlighting_available = ColorDelegator is not None and Percolator is not None
        self.large_file_threshold_bytes = 5 * 1024 * 1024
        self.max_editable_large_file_bytes = 24 * 1024 * 1024
        self.file_load_chunk_size = 256 * 1024
        self.background_stream_read_chunk_size = 256 * 1024
        self.background_stream_max_pending_chunks = 4
        self.background_file_result_batch_size = 2
        self.virtual_index_chunk_size = 4 * 1024 * 1024
        self.huge_file_preview_threshold_bytes = 100 * 1024 * 1024
        self.huge_file_preview_bytes = 2 * 1024 * 1024
        self.virtual_file_window_lines = 5000
        self.virtual_file_margin_lines = 800
        self.virtual_file_window_max_bytes = 8 * 1024 * 1024
        self.huge_virtual_file_window_max_bytes = 4 * 1024 * 1024
        self.virtual_hot_chunk_cache_entries = 10
        self.virtual_cold_chunk_cache_entries = 20
        self.virtual_cold_chunk_cache_enabled = True
        config_dir = self.get_config_dir(self.app_dir)
        os.makedirs(config_dir, exist_ok=True)
        self.locale_dir = self.get_locale_dir(config_dir)
        os.makedirs(self.locale_dir, exist_ok=True)
        self.theme_dir = self.get_theme_dir(config_dir)
        os.makedirs(self.theme_dir, exist_ok=True)
        self.backup_dir = os.path.join(config_dir, "backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        self.remote_cache_dir = os.path.join(config_dir, "remote-cache")
        os.makedirs(self.remote_cache_dir, exist_ok=True)
        self.migrate_language_files(config_dir=config_dir, locale_dir=self.locale_dir)
        self.ensure_theme_files(self.theme_dir)
        self.theme_definitions = self.load_theme_definitions(self.theme_dir)
        self.theme_effect_delay_ms = 90
        self.rainbow_theme_tag_prefix = 'theme_rainbow_'
        self.rainbow_theme_target_ranges = 6000
        self.rainbow_theme_palette = self.build_rainbow_theme_palette()
        self.locale_code = 'en_us'
        self.locale_path = self.get_locale_file_path(self.locale_code, locale_dir=self.locale_dir)
        self.locale_strings = self.load_locale_strings(self.locale_path)
        self.app_name = self.tr('app.name', 'Notepad-X')
        self.session_path = self.build_session_path(config_dir)
        self.editor_identity_path = self.build_editor_identity_path(config_dir)
        self.recovery_path = os.path.join(self.app_dir, "Notepad-X.recovery.json")
        self.refresh_log_paths()
        self.help_path = os.path.join(self.resource_dir, "Notepad-X-help.txt")
        self.note_color_labels = {
            'yellow': self.tr('note.filter.yellow', 'Yellow'),
            'green': self.tr('note.filter.green', 'Green'),
            'red': self.tr('note.filter.red', 'Red'),
            'blue': self.tr('note.filter.blue', 'Light Blue'),
        }
        self.max_recent_files = 10
        self.max_search_history = 10
        self.max_command_history = 25
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
        self.intellisense_max_project_files = 24
        self.intellisense_max_suggestions = 28
        self.spellcheck_delay_ms = 260
        self.spellcheck_max_chars = 250000
        self.spellcheck_max_words = 8000
        self.spellcheck_token_pattern = re.compile(r"[A-Za-z]+(?:'[A-Za-z]+)?")
        self.spellcheck_skip_neighbor_chars = set('_/\\.@#:-')
        self.spellcheck_custom_words = {
            'notepad', 'notepadx', 'github', 'gitignore', 'gitattributes', 'grab',
            'lolcat', 'codex', 'pyspellchecker', 'plaintext', 'utf', 'npxe', 'npx',
            'json', 'yaml', 'toml', 'xml', 'html', 'css', 'javascript', 'typescript',
            'python', 'java', 'powershell', 'markdown', 'autocomplete'
        }
        self.spellcheck_supported_modes = {None, 'ini', 'nfo', 'tex', 'markdown'}
        self.spell_checker = None
        self.spell_checker_ready = False
        self.recent_files = []
        self.find_history = []
        self.find_in_history = []
        self.closed_session_files = set()
        self.note_sync_interval_ms = 100
        self.note_editor_heartbeat_interval_ms = 1500
        self.markdown_preview_delay_ms = 45
        self.command_output_max_chars = 50000
        self.minimap_width = 64
        self.minimap_max_segments = 240
        self.diagnostics_delay_ms = 500
        self.diagnostic_tooltip_delay_ms = 350
        self.autosave_delay_ms = 12000
        self.max_backup_versions_per_doc = 20
        self.markdown_link_pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
        self.markdown_inline_patterns = (
            ('md_code_inline', re.compile(r'`([^`]+)`')),
            ('md_bold', re.compile(r'(\*\*|__)(.+?)\1')),
            ('md_italic', re.compile(r'(?<!\*)\*([^*\n]+)\*(?!\*)|(?<!_)_([^_\n]+)_(?!_)')),
        )
        self.markdown_heading_pattern = re.compile(r'^(#{1,6})\s+(.*)$')
        self.markdown_quote_pattern = re.compile(r'^\s*>\s?(.*)$')
        self.markdown_list_pattern = re.compile(r'^(\s*)([-*+])\s+(.*)$')
        self.markdown_ordered_list_pattern = re.compile(r'^(\s*)(\d+)[.)]\s+(.*)$')
        self.log_traceback_header_pattern = re.compile(r'^Traceback \(most recent call last\):\s*$', re.IGNORECASE)
        self.log_traceback_frame_pattern = re.compile(r'^\s*File\s+"[^"]+",\s+line\s+\d+.*$')
        self.log_exception_line_pattern = re.compile(r'^\s*([A-Za-z_][\w.]*(Error|Exception|Warning))(?::|\b)(.*)$')
        self.log_error_level_pattern = re.compile(r'^\s*(?:\[[^\]]+\]\s*)?(?:error|fatal|critical)\b[:\s-]*(.*)$', re.IGNORECASE)
        self.log_warning_level_pattern = re.compile(r'^\s*(?:\[[^\]]+\]\s*)?(?:warn(?:ing)?)\b[:\s-]*(.*)$', re.IGNORECASE)
        self.single_instance_host = '127.0.0.1'
        self.single_instance_port = get_notepadx_single_instance_port(self.app_dir)
        self.single_instance_server = None
        self.single_instance_listener_thread = None
        self.single_instance_running = False
        self.remote_open_files = []
        self.remote_open_lock = threading.Lock()
        self.background_file_results = []
        self.background_file_lock = threading.Lock()
        self.index_process_executor = None
        self.index_process_executor_disabled = False
        self.index_process_chunk_size = 4 * 1024 * 1024
        self.index_process_target_workers = max(1, (int(os.cpu_count() or 1) + 1) // 2)
        self.index_process_min_bytes_per_worker = 24 * 1024 * 1024
        self.virtual_index_task_multiplier = 4
        self.kernel32 = None
        self.psapi = None

        self.base_font_size = 13
        self.current_font_size = self.base_font_size
        self.font_family = 'Courier New'
        self.gutter_font = None
        self.currently_editing_measure_font = None
        self.min_font_size = 6
        self.max_font_size = 32
        self.word_wrap_enabled = tk.BooleanVar(value=True)
        self.sound_enabled = tk.BooleanVar(value=True)
        self.status_bar_enabled = tk.BooleanVar(value=True)
        self.numbered_lines_enabled = tk.BooleanVar(value=True)
        self.autocomplete_enabled = tk.BooleanVar(value=True)
        self.spell_check_enabled = tk.BooleanVar(value=SpellChecker is not None)
        self.edit_with_shell_enabled = tk.BooleanVar(value=False)
        self.auto_pair_enabled = tk.BooleanVar(value=True)
        self.compare_multi_edit_enabled = tk.BooleanVar(value=False)
        self.command_panel_enabled = tk.BooleanVar(value=True)
        self.minimap_enabled = tk.BooleanVar(value=True)
        self.breadcrumbs_enabled = tk.BooleanVar(value=True)
        self.diagnostics_enabled = tk.BooleanVar(value=True)
        self.autosave_enabled = tk.BooleanVar(value=True)
        self.find_in_selected_directory = ""
        self.note_filter = tk.StringVar(value='all')
        self.syntax_theme = tk.StringVar(value='Default')
        self.syntax_mode_selection = tk.StringVar(value='auto')
        self.language_selection = tk.StringVar(value=self.locale_code)
        self.markdown_preview_enabled = tk.BooleanVar(value=False)
        self.recovery_job = None
        self.find_change_job = None
        self.compare_active = False
        self.compare_source_tab = None
        self.compare_refresh_job = None
        self.markdown_preview_source_tab = None
        self.markdown_preview_refresh_job = None
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
        self.search_history_popup = None
        self.search_history_listbox = None
        self.search_history_entry = None
        self.search_history_items = []
        self.search_history_hide_job = None
        self.command_suggestion_popup = None
        self.command_suggestion_listbox = None
        self.command_suggestion_items = []
        self.command_suggestion_hide_job = None
        self.command_panel_visible = False
        self.command_history = []
        self.command_history_index = None
        self.command_history_listbox = None
        self.command_panel_height = 8
        self.command_panel_default_height = 8
        self.command_panel_min_height = 5
        self.command_panel_max_height = 28
        self.command_panel_resize_active = False
        self.command_panel_resize_origin_y = 0
        self.command_panel_resize_origin_height = self.command_panel_height
        self.command_runner_thread = None
        self.command_runner_active = False
        self.command_output_buffer = []
        self.default_window_width = 1500
        self.default_window_height = 700
        self.window_width = self.default_window_width
        self.window_height = self.default_window_height
        self.window_state = 'normal'
        self.window_layout_job = None
        self.window_layout_restored = False
        self.window_layout_save_delay_ms = 650
        self.hotkey_overrides = {}
        self.hotkey_bound_sequences = set()
        self.hotkey_dialog = None
        self.active_context_menu = None
        self.context_menu_posted_at = 0.0
        self.hovered_editor_widget = None
        self.diagnostic_tooltip_popup = None
        self.diagnostic_tooltip_job = None
        self.diagnostic_tooltip_doc = None
        self.diagnostic_tooltip_signature = None
        self.tab_context_menu = None
        self.tab_context_menu_tab_id = None
        self.sync_page_navigation_enabled = tk.BooleanVar(value=False)
        self.find_panel_visible = False
        self.replace_panel_visible = False
        self.currently_editing_panel_visible = False
        self.fullscreen = False
        self.fullscreen_panel_restore = False
        self.startup_recovery_restore_scheduled = False

    def init_runtime(self):
        self.kernel32 = None
        self.psapi = None
        if self.is_windows:
            self.kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            self.psapi = ctypes.WinDLL('psapi', use_last_error=True)
        self.configure_memory_api()
        self.configure_sound_api()
        self.configure_process_identity()
        self.known_editor_ids = self.load_known_editor_ids()
        self.editor_id = self.generate_editor_id()
        self.editor_aliases = set(self.known_editor_ids)
        self.editor_aliases.add(self.editor_id)
        self.persist_editor_identity()

        self.setup_exception_handling()
        self.start_single_instance_server()

    def init_ui(self):
        self.apply_window_icon(self.root)
        self.root.geometry(f"{self.default_window_width}x{self.default_window_height}")
        self.root.configure(bg='#1e1e1e')
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.create_text_widget()
        self.create_bottom_panels()
        self.create_menu()
        self.create_status_bar()
        self.reset_startup_trace()
        self.trace_startup(f"init_ui session_path={self.session_path}")
        self.trace_startup(f"init_ui startup_files={self.startup_files}")
        self.restore_session()
        self.root.bind('<Configure>', self.on_root_configure, add='+')

        self.bind_keys()
        self.update_font()
        self.update_clock()
        self.update_memory_usage()
        self.process_background_file_results()
        self.process_remote_open_requests()
        self.poll_shared_notes()
        if not self.window_layout_restored or self.window_state != 'zoomed':
            self.center_window(self.root)
        self.reveal_main_window()
        self.schedule_startup_file_opening()
        self.schedule_blank_startup_memory_trim()

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
        try:
            self.psapi.EmptyWorkingSet.restype = wintypes.BOOL
            self.psapi.EmptyWorkingSet.argtypes = [wintypes.HANDLE]
        except AttributeError:
            pass

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

    def configure_process_identity(self):
        if not self.is_windows:
            return
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(self.app_user_model_id)
        except Exception:
            pass

    def request_app_shutdown(self):
        self._shutdown_requested = True
        try:
            if self.root.winfo_exists():
                self.root.destroy()
        except tk.TclError:
            pass

    def begin_doc_load(self, doc):
        if not doc:
            return
        doc['loading_file'] = True
        doc['suspend_modified_events'] = True

    def end_doc_load(self, doc):
        if not doc:
            return
        doc['loading_file'] = False
        doc['suspend_modified_events'] = False
        text_widget = doc.get('text')
        if text_widget:
            try:
                text_widget.edit_modified(False)
            except tk.TclError:
                pass

    def get_index_process_executor(self):
        if self.index_process_executor_disabled:
            return None
        if self.index_process_executor is not None:
            return self.index_process_executor
        try:
            self.index_process_executor = ProcessPoolExecutor(
                max_workers=max(1, int(getattr(self, 'index_process_target_workers', 1) or 1))
            )
        except Exception as exc:
            self.index_process_executor_disabled = True
            self.log_exception("create index process executor", exc)
            return None
        return self.index_process_executor

    def shutdown_index_process_executor(self):
        executor = self.index_process_executor
        self.index_process_executor = None
        if executor is None:
            return
        try:
            executor.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            executor.shutdown(wait=False)
        except Exception as exc:
            self.log_exception("shutdown index process executor", exc)

    def cancel_doc_background_index(self, doc):
        if not doc:
            return
        future = doc.get('background_index_future')
        doc['background_index_future'] = None
        doc['background_index_token'] = None
        doc['background_index_active'] = False
        if future is None:
            return
        futures = list(future) if isinstance(future, (list, tuple, set)) else [future]
        for pending_future in futures:
            try:
                pending_future.cancel()
            except Exception:
                pass

    def get_background_index_worker_count(self, file_size):
        try:
            safe_file_size = max(0, int(file_size or 0))
        except (TypeError, ValueError):
            safe_file_size = 0
        target_workers = max(1, int(getattr(self, 'index_process_target_workers', 1) or 1))
        bytes_per_worker = max(4 * 1024 * 1024, int(getattr(self, 'index_process_min_bytes_per_worker', 24 * 1024 * 1024) or (24 * 1024 * 1024)))
        size_limited_workers = max(1, (safe_file_size + bytes_per_worker - 1) // bytes_per_worker)
        return max(1, min(target_workers, size_limited_workers))

    def build_background_index_ranges(self, file_size, worker_count, task_multiplier=1):
        try:
            safe_file_size = max(0, int(file_size or 0))
        except (TypeError, ValueError):
            safe_file_size = 0
        safe_worker_count = max(1, int(worker_count or 1))
        safe_task_multiplier = max(1, int(task_multiplier or 1))
        if safe_file_size <= 0:
            return [(0, 0)]
        target_segments = max(1, safe_worker_count * safe_task_multiplier)
        if target_segments == 1:
            return [(0, safe_file_size)]
        step = max(1, (safe_file_size + target_segments - 1) // target_segments)
        ranges = []
        start_offset = 0
        while start_offset < safe_file_size:
            end_offset = min(safe_file_size, start_offset + step)
            ranges.append((start_offset, end_offset))
            start_offset = end_offset
        return ranges

    def start_background_text_index(self, doc, file_path):
        if not doc or not file_path:
            return False
        executor = self.get_index_process_executor()
        if executor is None:
            return False

        token = str(doc.get('background_load_token') or f"{doc['frame']}:index:{time.time_ns()}")
        doc['background_index_token'] = token
        doc['background_index_active'] = True
        file_size = max(0, int(doc.get('background_bytes_total', 0) or 0))
        worker_count = self.get_background_index_worker_count(file_size)
        index_ranges = self.build_background_index_ranges(file_size, worker_count)
        try:
            futures = [
                executor.submit(
                    scan_large_text_file_index_range_worker,
                    file_path,
                    start_offset,
                    end_offset,
                    self.index_process_chunk_size,
                    range_index
                )
                for range_index, (start_offset, end_offset) in enumerate(index_ranges)
            ]
        except Exception as exc:
            doc['background_index_token'] = None
            doc['background_index_active'] = False
            self.log_exception("submit background text index", exc)
            return False
        doc['background_index_future'] = list(futures)
        frame_id = str(doc['frame'])
        file_path_abs = os.path.abspath(file_path)

        def coordinator(current_futures=tuple(futures), current_tab_id=frame_id, current_token=token, current_file_path=file_path_abs, current_file_size=file_size):
            try:
                partial_results = [future.result() for future in current_futures]
            except Exception as exc:
                self.queue_background_file_result({
                    'kind': 'text_index_error',
                    'tab_id': current_tab_id,
                    'token': current_token,
                    'file_path': current_file_path,
                    'error': exc,
                })
                return
            partial_results = sorted(
                (dict(item or {}) for item in partial_results),
                key=lambda item: int(item.get('range_index', 0) or 0)
            )
            total_lines = 0
            for item in partial_results:
                try:
                    total_lines += len(item.get('line_lengths') or [])
                except TypeError:
                    continue
            total_lines = max(1, total_lines)
            sample_step = max(1, total_lines // self.minimap_max_segments)
            segment_count = max(1, (total_lines + sample_step - 1) // sample_step)
            segment_max_lengths = [0] * segment_count
            current_line_index = 0
            for item in partial_results:
                for raw_length in (item.get('line_lengths') or []):
                    safe_length = max(0, int(raw_length or 0))
                    segment_index = min(segment_count - 1, current_line_index // sample_step)
                    if safe_length > segment_max_lengths[segment_index]:
                        segment_max_lengths[segment_index] = safe_length
                    current_line_index += 1
            payload = {
                'total_lines': total_lines,
                'sample_step': sample_step,
                'segment_max_lengths': segment_max_lengths,
                'file_size_bytes': current_file_size,
                'kind': 'text_index',
                'tab_id': current_tab_id,
                'token': current_token,
                'file_path': current_file_path,
            }
            self.queue_background_file_result(payload)

        threading.Thread(target=coordinator, name='NotepadXIndexCoordinator', daemon=True).start()
        return True

    def format_byte_size(self, size_bytes):
        try:
            size_value = max(0, int(size_bytes or 0))
        except (TypeError, ValueError):
            size_value = 0
        if size_value == 1:
            return '1 byte'
        if size_value < 1024:
            return f'{size_value} bytes'
        units = ['KB', 'MB', 'GB', 'TB']
        size_float = float(size_value)
        for unit in units:
            size_float /= 1024.0
            if size_float < 1024.0 or unit == units[-1]:
                return f'{size_float:.1f} {unit}'
        return f'{size_value} bytes'

    def close_doc_load_progress(self, doc):
        if not doc:
            return
        dialog = doc.get('load_progress_dialog')
        doc['load_progress_dialog'] = None
        doc['load_progress_bar'] = None
        doc['load_progress_status_label'] = None
        doc['load_progress_detail_label'] = None
        if not dialog:
            return
        try:
            if dialog.winfo_exists():
                dialog.destroy()
        except tk.TclError:
            pass

    def ensure_doc_load_progress(self, doc, file_path=None, total_bytes=None):
        if not doc:
            return
        dialog = doc.get('load_progress_dialog')
        try:
            if dialog and dialog.winfo_exists():
                return
        except tk.TclError:
            pass

        display_path = str(file_path or doc.get('display_name') or doc.get('file_path') or self.get_doc_name(doc['frame']) or '')
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('large_file.buffering_title', 'Buffering Large File'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=14, pady=12)

        tk.Label(
            dialog,
            text=self.tr('large_file.buffering_prompt', 'Loading:\n{file_path}', file_path=display_path),
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w',
            wraplength=540
        ).pack(anchor='w', pady=(0, 10))

        progress_bar = ttk.Progressbar(dialog, orient='horizontal', mode='determinate', length=340)
        progress_bar.pack(fill='x')

        status_label = tk.Label(
            dialog,
            text='',
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        )
        status_label.pack(anchor='w', pady=(10, 0))

        detail_label = tk.Label(
            dialog,
            text='',
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        )
        detail_label.pack(anchor='w', pady=(4, 0))

        dialog.protocol('WM_DELETE_WINDOW', lambda: None)
        dialog.update_idletasks()
        self.center_window(dialog, self.root)
        dialog.lift()
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current, self.root))

        doc['load_progress_dialog'] = dialog
        doc['load_progress_bar'] = progress_bar
        doc['load_progress_status_label'] = status_label
        doc['load_progress_detail_label'] = detail_label
        self.update_doc_load_progress(
            doc,
            loaded_bytes=doc.get('background_bytes_loaded', 0),
            total_bytes=total_bytes if total_bytes is not None else doc.get('background_bytes_total', 0),
            line_count=doc.get('background_lines_loaded', 1)
        )

    def update_doc_load_progress(self, doc, loaded_bytes=None, total_bytes=None, line_count=None):
        if not doc:
            return
        dialog = doc.get('load_progress_dialog')
        if not dialog:
            return
        try:
            if not dialog.winfo_exists():
                self.close_doc_load_progress(doc)
                return
        except tk.TclError:
            self.close_doc_load_progress(doc)
            return

        if loaded_bytes is None:
            loaded_bytes = doc.get('background_bytes_loaded', 0)
        if total_bytes is None:
            total_bytes = doc.get('background_bytes_total', 0)
        if line_count is None:
            line_count = doc.get('background_lines_loaded', 1)

        try:
            loaded_value = max(0, int(loaded_bytes or 0))
        except (TypeError, ValueError):
            loaded_value = 0
        try:
            total_value = max(0, int(total_bytes or 0))
        except (TypeError, ValueError):
            total_value = 0
        try:
            line_value = max(1, int(line_count or 1))
        except (TypeError, ValueError):
            line_value = 1

        if total_value > 0:
            loaded_value = min(loaded_value, total_value)
            percent_value = int((loaded_value / total_value) * 100)
        else:
            percent_value = 0

        progress_bar = doc.get('load_progress_bar')
        status_label = doc.get('load_progress_status_label')
        detail_label = doc.get('load_progress_detail_label')
        try:
            if progress_bar is not None:
                progress_bar.configure(maximum=max(1, total_value or 1), value=max(0, loaded_value))
            if status_label is not None:
                status_label.configure(
                    text=self.tr(
                        'large_file.buffering_status',
                        '{percent}% ({loaded} / {total})',
                        percent=percent_value,
                        loaded=self.format_byte_size(loaded_value),
                        total=self.format_byte_size(total_value)
                    )
                )
            if detail_label is not None:
                detail_label.configure(
                    text=self.tr(
                        'large_file.buffering_lines',
                        'Buffered {line_count} lines',
                        line_count=f'{line_value:,}'
                    )
                )
        except tk.TclError:
            self.close_doc_load_progress(doc)

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

    def create_toplevel(self, parent=None):
        window = tk.Toplevel(parent or self.root)
        self.apply_window_icon(window)
        return window

    def create_popup_toplevel(self, parent=None, topmost=True):
        window = tk.Toplevel(parent or self.root)
        try:
            window.withdraw()
        except tk.TclError:
            pass
        try:
            window.wm_overrideredirect(True)
        except tk.TclError:
            pass
        if self.is_windows:
            try:
                window.attributes('-toolwindow', True)
            except tk.TclError:
                pass
        if topmost:
            try:
                window.attributes('-topmost', True)
            except tk.TclError:
                pass
        return window

    def show_popup_toplevel(self, window, x, y):
        if window is None:
            return
        try:
            window.geometry(f'+{int(x)}+{int(y)}')
            window.deiconify()
            window.lift()
        except tk.TclError:
            return

    def reveal_main_window(self):
        try:
            if not self.root.winfo_exists():
                return
            self.root.update_idletasks()
            self.root.deiconify()
            self.root.lift()
        except tk.TclError:
            return

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

    def get_logs_dir(self, base_dir):
        return os.path.join(base_dir, 'logs')

    def refresh_log_paths(self):
        self.logs_dir = self.get_logs_dir(self.app_dir)
        os.makedirs(self.logs_dir, exist_ok=True)
        self.crash_log_path = os.path.join(self.logs_dir, "Notepad-X.crash.log")
        self.startup_trace_path = os.path.join(self.logs_dir, "Notepad-X.startup.trace.log")

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
            'Lolcat': {
                'surface': {
                    'text_bg': '#101218',
                    'text_fg': '#f8f8ff',
                    'cursor': '#ffffff',
                    'selection': '#3c2a57',
                    'gutter_bg': '#0c0e13',
                    'gutter_current_bg': '#171b25',
                    'gutter_fg': '#8a90a3',
                    'gutter_current_fg': '#f8f8ff',
                    'gutter_divider': '#31384b',
                },
                'syntax': {
                    'keyword': '#ff4d4d',
                    'type': '#ff9f1c',
                    'string': '#ffe66d',
                    'comment': '#4cd964',
                    'number': '#34c3ff',
                    'preprocessor': '#5b5bff',
                    'tag': '#c155ff',
                },
                'text_effect': 'rainbow',
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
        sanitized = {
            'name': safe_name,
            'surface': surface,
            'syntax': syntax,
        }
        text_effect = str(payload.get('text_effect') or '').strip().lower()
        if text_effect == 'rainbow':
            sanitized['text_effect'] = text_effect
        return sanitized

    def ensure_theme_files(self, theme_dir):
        builtins = self.get_builtin_theme_definitions()
        for theme_name, theme_payload in builtins.items():
            file_name = f"{self.slugify_theme_name(theme_name)}.json"
            if file_name == 'lolcat.json':
                continue
            file_path = os.path.join(theme_dir, file_name)
            if os.path.exists(file_path):
                continue
            payload = {
                'name': theme_name,
                'surface': dict(theme_payload['surface']),
                'syntax': dict(theme_payload['syntax']),
            }
            if theme_payload.get('text_effect'):
                payload['text_effect'] = theme_payload['text_effect']
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
            'Lolcat': 'syntax.theme.lolcat',
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
            self.show_filesystem_error(self.tr('theme.save_failed_title', 'Save Theme Failed'), file_path, exc)
            return False
        self.theme_definitions = self.load_theme_definitions(self.theme_dir)
        self.create_menu()
        self.set_syntax_theme(theme_name)
        return True

    def show_create_theme_dialog(self):
        t = self.tr
        dialog = self.create_toplevel(self.root)
        dialog.title(t('theme.create.title', 'Create Theme'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color)

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
        payload = {}
        needs_write = False
        if not os.path.exists(locale_path):
            needs_write = True
        try:
            if os.path.exists(locale_path):
                with open(locale_path, 'r', encoding='utf-8') as f:
                    loaded_payload = self.parse_locale_strings(f.read())
                if isinstance(loaded_payload, dict):
                    for key, value in loaded_payload.items():
                        if isinstance(key, str) and isinstance(value, str):
                            payload[key] = value
        except OSError:
            pass
        if payload:
            strings.update(payload)
        if not payload or any(key not in payload for key in DEFAULT_LOCALE_STRINGS):
            needs_write = True
        if needs_write:
            try:
                os.makedirs(os.path.dirname(locale_path), exist_ok=True)
                with open(locale_path, 'w', encoding='utf-8') as f:
                    f.write(self.serialize_locale_strings(strings))
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
        if not code:
            return self.tr('common.unknown', 'Unknown')
        exact_key = f'locale.display.{code}'
        exact_translation = self.locale_strings.get(exact_key)
        if isinstance(exact_translation, str) and exact_translation.strip():
            return exact_translation
        parts = [part for part in code.split('_') if part]
        if parts:
            language = parts[0]
            region = parts[1] if len(parts) > 1 else ''
            if language in LANGUAGE_NATIVE_NAMES:
                language_name = self.tr(
                    f'locale.language.{language}',
                    LANGUAGE_NATIVE_NAMES[language]
                )
                if region:
                    region_name = self.tr(
                        f'locale.region.{region}',
                        REGION_DISPLAY_NAMES.get(region, region.upper())
                    )
                    return f"{language_name} ({region_name})"
                return language_name
        if code in LOCALE_DISPLAY_NAMES:
            return LOCALE_DISPLAY_NAMES[code]
        parts = [part.upper() if len(part) <= 3 else part.title() for part in parts]
        return " / ".join(parts) if parts else self.tr('common.unknown', 'Unknown')

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
        if language == 'he':
            if self.is_windows:
                return ['Segoe UI', 'Arial', 'Tahoma', 'Courier New']
            return ['Noto Sans Hebrew', 'Liberation Sans', 'DejaVu Sans', 'Monospace']
        if language == 'bn':
            if self.is_windows:
                return ['Nirmala UI', 'Vrinda', 'Arial Unicode MS', 'Courier New']
            return ['Noto Sans Bengali', 'Lohit Bengali', 'Mukti Narrow', 'DejaVu Sans']
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
            self.status_tail.config(text=self.tr('status.resource_initial', " | CPU: 0.0% Memory: 0MB"))
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
        if (
            persist
            and hasattr(self, 'edit_with_shell_enabled')
            and bool(self.edit_with_shell_enabled.get())
        ):
            self.sync_edit_with_shell_menu(show_errors=False)
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
        self.refresh_log_paths()

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
        self.refresh_log_paths()

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
        temp_path = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode) | 0o664
            except OSError:
                target_mode = 0o664
        try:
            fd, temp_path = tempfile.mkstemp(prefix=prefix, suffix='.tmp', dir=directory)
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
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def write_binary_atomically(self, file_path, payload_bytes, prefix, context_name):
        directory = os.path.dirname(file_path) or '.'
        os.makedirs(directory, exist_ok=True)
        target_mode = None
        temp_path = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode)
            except OSError:
                target_mode = 0o664
        try:
            fd, temp_path = tempfile.mkstemp(prefix=prefix, suffix='.tmp', dir=directory)
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
            if temp_path and os.path.exists(temp_path):
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
            raise RuntimeError(self.tr('encryption.error.cryptography_required', 'The cryptography package is required for encrypted files.'))
        passphrase_bytes = str(passphrase or '').encode('utf-8')
        if not passphrase_bytes:
            raise ValueError(self.tr('encryption.error.passphrase_required', 'A passphrase is required.'))
        salt_text = header.get('salt')
        if not isinstance(salt_text, str) or not salt_text.strip():
            raise ValueError(self.tr('encryption.error.missing_salt', 'Encrypted file is missing its salt.'))
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
                    self.tr(
                        'encryption.error.scrypt_memory',
                        'Notepad-X could not derive the encryption key because the scrypt memory limit was exceeded on this machine.'
                    )
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
            raise ValueError(self.tr('encryption.error.header_incomplete', 'Encrypted file header is incomplete.'))
        header_length = int.from_bytes(payload_bytes[header_offset:header_offset + 4], 'big')
        header_start = header_offset + 4
        header_end = header_start + header_length
        if header_length <= 0 or len(payload_bytes) < header_end + self.encryption_nonce_length:
            raise ValueError(self.tr('encryption.error.header_invalid', 'Encrypted file header is invalid.'))
        header = json.loads(payload_bytes[header_start:header_end].decode('utf-8'))
        nonce_start = header_end
        nonce_end = nonce_start + self.encryption_nonce_length
        nonce = payload_bytes[nonce_start:nonce_end]
        ciphertext = payload_bytes[nonce_end:]
        if not isinstance(header, dict) or header.get('format') != 'Notepad-X Encrypted':
            raise ValueError(self.tr('encryption.error.header_invalid', 'Encrypted file header is invalid.'))
        if header.get('cipher') != 'AES-256-GCM' or header.get('kdf') != 'scrypt':
            raise ValueError(self.tr('encryption.error.unsupported_settings', 'Unsupported encrypted file settings.'))
        if not ciphertext:
            raise ValueError(self.tr('encryption.error.no_ciphertext', 'Encrypted file has no ciphertext.'))
        return header, nonce, ciphertext

    def file_looks_encrypted(self, file_path):
        try:
            with open(file_path, 'rb') as encrypted_file:
                return encrypted_file.read(len(self.encryption_magic)) == self.encryption_magic
        except OSError:
            return False

    def show_encryption_unavailable(self, parent=None):
        messagebox.showerror(
            self.tr('encryption.unavailable_title', 'Encryption Unavailable'),
            self.tr(
                'encryption.unavailable_message',
                "Encrypted save/open needs the 'cryptography' Python package.\n\nInstall it on this machine to use Notepad-X encrypted files."
            ),
            parent=parent or self.root
        )

    def prompt_encryption_options(self, default_encrypt=False, parent=None):
        parent = parent or self.root
        dialog = self.create_toplevel(parent)
        dialog.title(self.tr('encryption.save_title', 'Save Encrypted Copy As'))
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)

        result = {'value': None}
        encrypt_var = tk.BooleanVar(value=bool(default_encrypt))
        show_var = tk.BooleanVar(value=False)
        passphrase_var = tk.StringVar()
        confirm_var = tk.StringVar()

        tk.Checkbutton(
            dialog,
            text=self.tr('encryption.encrypt_file', 'Encrypt file'),
            variable=encrypt_var,
            bg='#f0f0f0',
            anchor='w'
        ).pack(anchor='w', pady=(0, 8))

        tk.Label(dialog, text=self.tr('encryption.passphrase', 'Passphrase:'), bg='#f0f0f0', fg='black', font=('Segoe UI', 9)).pack(anchor='w')
        passphrase_entry = tk.Entry(dialog, textvariable=passphrase_var, width=34, show='*')
        passphrase_entry.pack(fill='x', pady=(0, 6))

        tk.Label(dialog, text=self.tr('encryption.confirm_passphrase', 'Confirm passphrase:'), bg='#f0f0f0', fg='black', font=('Segoe UI', 9)).pack(anchor='w')
        confirm_entry = tk.Entry(dialog, textvariable=confirm_var, width=34, show='*')
        confirm_entry.pack(fill='x', pady=(0, 6))

        def update_visibility(*_args):
            state = 'normal' if encrypt_var.get() else 'disabled'
            passphrase_entry.configure(state=state, show='' if show_var.get() else '*')
            confirm_entry.configure(state=state, show='' if show_var.get() else '*')

        tk.Checkbutton(
            dialog,
            text=self.tr('encryption.show_passphrase', 'Show passphrase'),
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
                    messagebox.showinfo(
                        self.tr('encryption.save_title', 'Save Encrypted Copy As'),
                        self.tr('encryption.passphrase_required', 'Enter a passphrase first.'),
                        parent=dialog
                    )
                    return "break"
                if passphrase != confirm:
                    messagebox.showinfo(
                        self.tr('encryption.save_title', 'Save Encrypted Copy As'),
                        self.tr('encryption.passphrase_mismatch', 'Passphrases do not match.'),
                        parent=dialog
                    )
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

        tk.Button(button_row, text=self.tr('common.ok', 'OK'), width=10, command=submit).pack(side='left')
        tk.Button(button_row, text=self.tr('common.cancel', 'Cancel'), width=10, command=cancel).pack(side='right')

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
        self.trace_startup(
            f"prompt_open_passphrase create file={file_path} "
            f"mainloop={self.main_loop_started} viewable={self.root_is_ready_for_dialogs()}"
        )
        dialog = self.create_toplevel(parent)
        dialog.title(self.tr('encryption.open_title', 'Open Encrypted File'))
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)

        result = {'value': None}
        show_var = tk.BooleanVar(value=False)
        passphrase_var = tk.StringVar()

        tk.Label(
            dialog,
            text=self.tr('encryption.open_prompt', 'Passphrase for:\n{file_name}', file_name=os.path.basename(file_path)),
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
            text=self.tr('encryption.show_passphrase', 'Show passphrase'),
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

        tk.Button(button_row, text=self.tr('common.ok', 'OK'), width=10, command=submit).pack(side='left')
        tk.Button(button_row, text=self.tr('common.cancel', 'Cancel'), width=10, command=cancel).pack(side='right')

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
        if self._shutdown_requested:
            return
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        results = []
        remaining_results = False
        with self.background_file_lock:
            if self.background_file_results:
                batch_size = max(1, int(getattr(self, 'background_file_result_batch_size', 1) or 1))
                results = self.background_file_results[:batch_size]
                del self.background_file_results[:batch_size]
                remaining_results = bool(self.background_file_results)
        for result in results:
            try:
                self.handle_background_file_result(result)
            except Exception as exc:
                self.log_exception("handle background file result", exc)
        if self._shutdown_requested:
            return
        try:
            self.root.after(8 if remaining_results else 30, self.process_background_file_results)
        except tk.TclError:
            pass

    def build_line_index_background(self, file_path):
        line_starts = [0]
        file_size = 0
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(self.virtual_index_chunk_size)
                if not chunk:
                    break
                base_offset = file_size
                file_size += len(chunk)
                search_from = 0
                while True:
                    newline_index = chunk.find(b'\n', search_from)
                    if newline_index == -1:
                        break
                    line_starts.append(base_offset + newline_index + 1)
                    search_from = newline_index + 1
        return line_starts, file_size

    def start_background_text_load(self, doc, file_path):
        self.cancel_doc_background_index(doc)
        doc['background_loading'] = True
        doc['background_load_kind'] = 'text'
        doc['background_load_file_path'] = file_path
        doc['background_load_token'] = f"{doc['frame']}:{time.time_ns()}"
        doc['background_bytes_loaded'] = 0
        try:
            doc['background_bytes_total'] = max(0, int(os.path.getsize(file_path)))
        except OSError:
            doc['background_bytes_total'] = 0
        doc['background_lines_loaded'] = 1
        doc['pending_insert_batch_count'] = 0
        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)
        text.edit_modified(False)
        text.mark_set(tk.INSERT, '1.0')
        text.tag_remove('sel', '1.0', tk.END)
        text.see('1.0')
        self.invalidate_minimap_cache(doc)
        process_index_started = self.start_background_text_index(doc, file_path)
        if process_index_started:
            doc['minimap_progressive_state'] = None
            doc['minimap_model'] = self.create_minimap_model(1, 1, [0], complete=False)
            doc['minimap_model_dirty'] = False
        else:
            self.start_progressive_minimap_build(doc, None)
        self.refresh_minimap(doc)
        self.ensure_doc_load_progress(doc, file_path=file_path, total_bytes=doc.get('background_bytes_total', 0))
        frame_id = str(doc['frame'])
        load_token = str(doc.get('background_load_token') or '')

        def worker():
            try:
                decoder = codecs.getincrementaldecoder('utf-8')('replace')
                file_size = int(doc.get('background_bytes_total', 0) or 0)
                bytes_loaded = 0
                total_lines = 1
                with open(file_path, 'rb') as f:
                    while True:
                        if doc.get('background_load_token') != load_token:
                            return
                        chunk = f.read(self.background_stream_read_chunk_size)
                        if not chunk:
                            break
                        bytes_loaded += len(chunk)
                        chunk_text = decoder.decode(chunk, final=False)
                        if chunk_text:
                            total_lines += chunk_text.count('\n')
                            self.queue_background_file_result({
                                'kind': 'text_chunk',
                                'tab_id': frame_id,
                                'token': load_token,
                                'file_path': file_path,
                                'chunk_text': chunk_text,
                                'bytes_loaded': bytes_loaded,
                                'file_size_bytes': file_size,
                                'total_lines_hint': total_lines,
                            })
                        else:
                            self.queue_background_file_result({
                                'kind': 'text_progress',
                                'tab_id': frame_id,
                                'token': load_token,
                                'file_path': file_path,
                                'bytes_loaded': bytes_loaded,
                                'file_size_bytes': file_size,
                                'total_lines_hint': total_lines,
                            })
                        while self.count_pending_background_stream_chunks(frame_id, load_token) >= self.background_stream_max_pending_chunks:
                            if doc.get('background_load_token') != load_token:
                                return
                            time.sleep(0.01)
                tail_text = decoder.decode(b'', final=True)
                if tail_text:
                    total_lines += tail_text.count('\n')
                    self.queue_background_file_result({
                        'kind': 'text_chunk',
                        'tab_id': frame_id,
                        'token': load_token,
                        'file_path': file_path,
                        'chunk_text': tail_text,
                        'bytes_loaded': bytes_loaded,
                        'file_size_bytes': file_size,
                        'total_lines_hint': total_lines,
                    })
                self.queue_background_file_result({
                    'kind': 'text_complete',
                    'tab_id': frame_id,
                    'token': load_token,
                    'file_path': file_path,
                    'bytes_loaded': bytes_loaded,
                    'file_size_bytes': file_size,
                    'total_lines_hint': max(1, total_lines),
                })
            except Exception as exc:
                self.queue_background_file_result({
                    'kind': 'error',
                    'tab_id': frame_id,
                    'token': load_token,
                    'file_path': file_path,
                    'error': exc,
                })

        threading.Thread(target=worker, name='NotepadXLargeFileLoad', daemon=True).start()

    def count_pending_background_stream_chunks(self, tab_id, load_token):
        target_tab_id = str(tab_id or '')
        target_token = str(load_token or '')
        with self.background_file_lock:
            return sum(
                1 for item in self.background_file_results
                if str(item.get('tab_id') or '') == target_tab_id
                and str(item.get('token') or '') == target_token
                and item.get('kind') in {'text_chunk', 'text_progress'}
            )

    def append_background_text_chunk(self, doc, chunk_text, bytes_loaded=None, total_bytes=None, total_lines_hint=None):
        text = doc.get('text')
        if not text:
            return
        safe_chunk = chunk_text if isinstance(chunk_text, str) else ''
        update_status_now = not safe_chunk
        if bytes_loaded is not None:
            try:
                doc['background_bytes_loaded'] = max(0, int(bytes_loaded))
            except (TypeError, ValueError):
                pass
        if total_bytes is not None:
            try:
                doc['background_bytes_total'] = max(0, int(total_bytes))
            except (TypeError, ValueError):
                pass
        if total_lines_hint is not None:
            try:
                doc['background_lines_loaded'] = max(1, int(total_lines_hint))
            except (TypeError, ValueError):
                pass

        if safe_chunk:
            text.insert(tk.END, safe_chunk)
            text.edit_modified(False)
            if not doc.get('background_index_active'):
                self.append_progressive_minimap_chunk(doc, safe_chunk, finalize=False)
            doc['pending_insert_batch_count'] = int(doc.get('pending_insert_batch_count', 0) or 0) + 1
            batch_count = int(doc.get('pending_insert_batch_count', 0) or 0)
            update_status_now = batch_count == 1 or batch_count % 8 == 0
            if not doc.get('background_index_active') and (batch_count == 1 or batch_count % 16 == 0):
                self.refresh_minimap(doc)
        self.update_doc_load_progress(doc)
        if update_status_now and str(doc['frame']) == self.notebook.select():
            self.update_status()

    def finalize_background_text_load(self, doc, bytes_loaded=None, total_bytes=None, total_lines_hint=None):
        text = doc.get('text')
        if not text:
            return
        if bytes_loaded is not None:
            try:
                doc['background_bytes_loaded'] = max(0, int(bytes_loaded))
            except (TypeError, ValueError):
                pass
        if total_bytes is not None:
            try:
                doc['background_bytes_total'] = max(0, int(total_bytes))
            except (TypeError, ValueError):
                pass
        if total_lines_hint is not None:
            try:
                doc['background_lines_loaded'] = max(1, int(total_lines_hint))
            except (TypeError, ValueError):
                pass

        if doc.get('background_index_active'):
            if doc.get('minimap_model') is None:
                doc['minimap_model'] = self.create_minimap_model(max(1, int(doc.get('background_lines_loaded', 1) or 1)), 1, [0], complete=False)
                doc['minimap_model_dirty'] = False
        else:
            self.append_progressive_minimap_chunk(doc, '', finalize=True)
            self.refresh_minimap(doc)
        try:
            doc['total_file_lines'] = max(1, int(text.index('end-1c').split('.')[0]))
        except (tk.TclError, TypeError, ValueError):
            doc['total_file_lines'] = max(1, int(doc.get('background_lines_loaded', 1) or 1))
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
        doc['background_load_token'] = None
        doc['pending_insert_batch_count'] = 0
        doc['symbol_cache_signature'] = None
        doc['symbol_cache'] = None
        self.update_doc_load_progress(doc)
        self.close_doc_load_progress(doc)
        self.end_doc_load(doc)
        self.update_doc_file_signature(doc)
        self.configure_syntax_highlighting(doc['frame'])
        self.restore_doc_notes(doc)
        self.register_doc_for_shared_notes(doc)
        self.invalidate_fold_regions(doc)
        self.schedule_diagnostics(doc)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        self.refresh_tab_title(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        if self.markdown_preview_enabled.get() and str(doc['frame']) == self.notebook.select():
            self.schedule_markdown_preview_refresh()
        if str(doc['frame']) == self.notebook.select():
            self.update_status()

    def start_background_virtual_index(self, doc, file_path):
        self.cancel_doc_background_index(doc)
        doc['background_loading'] = True
        doc['background_load_kind'] = 'virtual'
        doc['background_load_file_path'] = file_path
        doc['background_load_token'] = f"{doc['frame']}:virtual:{time.time_ns()}"
        try:
            doc['background_bytes_total'] = max(0, int(os.path.getsize(file_path)))
        except OSError:
            doc['background_bytes_total'] = 0
        doc['background_bytes_loaded'] = 0
        doc['background_lines_loaded'] = 1
        text = doc['text']
        text.configure(state='normal')
        text.delete('1.0', tk.END)
        text.edit_modified(False)
        text.mark_set(tk.INSERT, '1.0')
        text.tag_remove('sel', '1.0', tk.END)
        text.see('1.0')
        self.ensure_doc_load_progress(doc, file_path=file_path, total_bytes=doc.get('background_bytes_total', 0))
        frame_id = str(doc['frame'])
        load_token = str(doc.get('background_load_token') or '')
        file_size = max(0, int(doc.get('background_bytes_total', 0) or 0))
        executor = self.get_index_process_executor()
        if executor is None:
            def worker():
                try:
                    line_starts, indexed_file_size = self.build_line_index_background(file_path)
                    minimap_model = self.build_minimap_model_from_line_starts(line_starts, indexed_file_size)
                    self.queue_background_file_result({
                        'kind': 'virtual',
                        'tab_id': frame_id,
                        'token': load_token,
                        'file_path': file_path,
                        'line_starts': line_starts,
                        'file_size_bytes': indexed_file_size,
                        'minimap_model': minimap_model,
                    })
                except Exception as exc:
                    self.queue_background_file_result({
                        'kind': 'error',
                        'tab_id': frame_id,
                        'token': load_token,
                        'file_path': file_path,
                        'error': exc,
                    })

            threading.Thread(target=worker, name='NotepadXLargeFileIndex', daemon=True).start()
            return

        worker_count = self.get_background_index_worker_count(file_size)
        index_ranges = self.build_background_index_ranges(file_size, worker_count, self.virtual_index_task_multiplier)
        token = load_token
        doc['background_index_token'] = token
        doc['background_index_active'] = True
        try:
            futures = [
                executor.submit(
                    scan_newline_start_offsets_range_worker,
                    file_path,
                    start_offset,
                    end_offset,
                    self.index_process_chunk_size,
                    range_index
                )
                for range_index, (start_offset, end_offset) in enumerate(index_ranges)
            ]
        except Exception as exc:
            doc['background_index_token'] = None
            doc['background_index_active'] = False
            self.log_exception("submit background virtual index", exc)
            def fallback_worker():
                try:
                    line_starts, indexed_file_size = self.build_line_index_background(file_path)
                    minimap_model = self.build_minimap_model_from_line_starts(line_starts, indexed_file_size)
                    self.queue_background_file_result({
                        'kind': 'virtual',
                        'tab_id': frame_id,
                        'token': token,
                        'file_path': file_path,
                        'line_starts': line_starts,
                        'file_size_bytes': indexed_file_size,
                        'minimap_model': minimap_model,
                    })
                except Exception as inner_exc:
                    self.queue_background_file_result({
                        'kind': 'error',
                        'tab_id': frame_id,
                        'token': token,
                        'file_path': file_path,
                        'error': inner_exc,
                    })

            threading.Thread(target=fallback_worker, name='NotepadXLargeFileIndexFallback', daemon=True).start()
            return

        doc['background_index_future'] = list(futures)
        file_path_abs = os.path.abspath(file_path)
        total_ranges = max(1, len(index_ranges))

        def coordinator(current_futures=tuple(futures), current_tab_id=frame_id, current_token=token, current_file_path=file_path_abs, current_file_size=file_size, current_total_ranges=total_ranges):
            completed_results = {}
            bytes_scanned = 0
            completed_count = 0
            try:
                for completed_future in as_completed(current_futures):
                    payload = completed_future.result()
                    payload = dict(payload or {})
                    completed_results[int(payload.get('range_index', completed_count) or completed_count)] = list(payload.get('line_starts') or [])
                    bytes_scanned += max(0, int(payload.get('bytes_scanned', 0) or 0))
                    completed_count += 1
                    self.queue_background_file_result({
                        'kind': 'virtual_index_progress',
                        'tab_id': current_tab_id,
                        'token': current_token,
                        'file_path': current_file_path,
                        'bytes_loaded': min(current_file_size, bytes_scanned),
                        'file_size_bytes': current_file_size,
                        'completed_ranges': completed_count,
                        'total_ranges': current_total_ranges,
                    })
                line_starts = [0]
                for range_index in range(current_total_ranges):
                    line_starts.extend(completed_results.get(range_index, []))
                if not line_starts:
                    line_starts = [0]
                minimap_model = self.build_minimap_model_from_line_starts(line_starts, current_file_size)
                self.queue_background_file_result({
                    'kind': 'virtual',
                    'tab_id': current_tab_id,
                    'token': current_token,
                    'file_path': current_file_path,
                    'line_starts': line_starts,
                    'file_size_bytes': current_file_size,
                    'minimap_model': minimap_model,
                })
            except Exception as exc:
                self.queue_background_file_result({
                    'kind': 'error',
                    'tab_id': current_tab_id,
                    'token': current_token,
                    'file_path': current_file_path,
                    'error': exc,
                })

        threading.Thread(target=coordinator, name='NotepadXVirtualIndexCoordinator', daemon=True).start()

    def begin_background_text_insert(self, doc, content, total_lines_hint=None):
        doc['pending_insert_content'] = content
        doc['pending_insert_offset'] = 0
        doc['pending_insert_batch_count'] = 0
        self.invalidate_minimap_cache(doc)
        self.start_progressive_minimap_build(doc, total_lines_hint)
        doc['diagnostics'] = []
        doc['text'].configure(state='normal')
        doc['text'].delete('1.0', tk.END)
        self.refresh_minimap(doc)
        self.continue_background_text_insert(doc)

    def continue_background_text_insert(self, doc):
        text = doc.get('text')
        content = doc.get('pending_insert_content')
        if not text or not isinstance(content, str):
            return
        offset = int(doc.get('pending_insert_offset', 0) or 0)
        batch_size = max(64 * 1024, self.file_load_chunk_size // 4)
        next_offset = min(len(content), offset + batch_size)
        if next_offset > offset:
            chunk_text = content[offset:next_offset]
            text.insert(tk.END, chunk_text)
            self.append_progressive_minimap_chunk(doc, chunk_text, finalize=(next_offset >= len(content)))
            doc['pending_insert_offset'] = next_offset
            doc['pending_insert_batch_count'] = int(doc.get('pending_insert_batch_count', 0) or 0) + 1
            batch_count = int(doc.get('pending_insert_batch_count', 0) or 0)
            if batch_count == 1 or batch_count % 4 == 0 or next_offset >= len(content):
                self.refresh_minimap(doc)
        elif next_offset >= len(content):
            self.append_progressive_minimap_chunk(doc, '', finalize=True)
            self.refresh_minimap(doc)
        if next_offset < len(content):
            self.root.after(1, lambda current=doc: self.continue_background_text_insert(current))
            return
        doc.pop('pending_insert_content', None)
        doc.pop('pending_insert_offset', None)
        doc.pop('pending_insert_batch_count', None)
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
        doc['background_load_token'] = None
        self.close_doc_load_progress(doc)
        self.end_doc_load(doc)
        self.update_doc_file_signature(doc)
        self.configure_syntax_highlighting(doc['frame'])
        self.restore_doc_notes(doc)
        self.register_doc_for_shared_notes(doc)
        self.invalidate_fold_regions(doc)
        self.schedule_diagnostics(doc)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        self.refresh_tab_title(doc['frame'])
        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.refresh_compare_panel()
        if str(doc['frame']) == self.notebook.select():
            self.update_status()

    def handle_background_file_error(self, doc, exc):
        self.cancel_doc_background_index(doc)
        doc['background_loading'] = False
        doc['background_load_kind'] = None
        doc['background_load_file_path'] = None
        doc['background_load_token'] = None
        doc.pop('pending_insert_content', None)
        doc.pop('pending_insert_offset', None)
        doc.pop('pending_insert_batch_count', None)
        self.close_doc_load_progress(doc)
        self.end_doc_load(doc)
        file_path = doc.get('file_path')
        messagebox.showerror(
            self.tr('file.open_failed_title', 'Open Failed'),
            self.tr(
                'file.open_failed_message',
                'Notepad-X could not open:\n{file_path}\n\n{error_detail}',
                file_path=file_path,
                error_detail=exc
            ),
            parent=self.root
        )
        self.cleanup_failed_file_open(doc)

    def cleanup_failed_file_open(self, doc):
        self.trace_startup(
            f"cleanup_failed_file_open frame={doc.get('frame')} "
            f"file_path={doc.get('file_path')} new_tab={doc.get('background_open_new_tab')}"
        )
        self.cancel_doc_background_index(doc)
        self.reset_virtual_backing_store(doc, remove_files=True)
        self.invalidate_minimap_cache(doc)
        if doc.get('background_open_new_tab'):
            try:
                self.notebook.forget(doc['frame'])
            except tk.TclError:
                pass
            self.documents.pop(str(doc['frame']), None)
        else:
            doc['file_path'] = None
            self.clear_remote_metadata(doc)
            doc['encrypted_file'] = False
            doc['encryption_header'] = None
            doc['encryption_key'] = None
            doc['preview_mode'] = False
            doc['virtual_mode'] = False
            doc['large_file_mode'] = False
            doc['file_size_bytes'] = 0
            doc['line_starts'] = None
            doc['total_file_lines'] = 1
            doc['window_start_line'] = 1
            doc['window_end_line'] = 1
            doc['last_virtual_line'] = 1
            doc['last_virtual_col'] = 0
            doc['pending_virtual_target_line'] = None
            doc['pending_insert_batch_count'] = 0
            doc['background_load_token'] = None
            doc['background_index_future'] = None
            doc['background_index_token'] = None
            doc['background_index_active'] = False
            doc['background_bytes_loaded'] = 0
            doc['background_bytes_total'] = 0
            doc['background_lines_loaded'] = 1
            self.close_doc_load_progress(doc)
            try:
                doc['text'].configure(state='normal')
                doc['text'].delete('1.0', tk.END)
                doc['text'].edit_modified(False)
                doc['text'].mark_set(tk.INSERT, '1.0')
                doc['text'].tag_remove('sel', '1.0', tk.END)
                doc['text'].see('1.0')
            except tk.TclError:
                pass
            self.refresh_tab_title(doc['frame'])
            if str(doc['frame']) == self.notebook.select():
                self.set_active_document(doc['frame'])
        doc['background_open_new_tab'] = False

    def handle_background_file_result(self, result):
        tab_id = str(result.get('tab_id') or '')
        doc = self.documents.get(tab_id)
        if not doc:
            return
        file_path = os.path.abspath(str(result.get('file_path') or ''))
        if not file_path or os.path.abspath(str(doc.get('file_path') or '')) != file_path:
            return
        result_kind = str(result.get('kind') or '')
        result_token = str(result.get('token') or '')
        doc_token = str(doc.get('background_load_token') or '')
        if result_token and doc_token and result_token != doc_token:
            return
        if result_kind in {'text_index', 'text_index_error'}:
            index_token = str(doc.get('background_index_token') or '')
            if index_token and result_token and result_token != index_token:
                return
        if result_kind == 'virtual_index_progress':
            doc['background_bytes_loaded'] = max(0, int(result.get('bytes_loaded') or 0))
            doc['background_bytes_total'] = max(0, int(result.get('file_size_bytes') or doc.get('background_bytes_total', 0) or 0))
            completed_ranges = max(0, int(result.get('completed_ranges') or 0))
            total_ranges = max(1, int(result.get('total_ranges') or 1))
            estimated_lines = max(1, int(doc.get('background_lines_loaded', 1) or 1))
            if completed_ranges >= total_ranges and doc.get('background_bytes_total'):
                estimated_lines = max(estimated_lines, int(doc.get('total_file_lines', estimated_lines) or estimated_lines))
            self.update_doc_load_progress(
                doc,
                loaded_bytes=doc.get('background_bytes_loaded'),
                total_bytes=doc.get('background_bytes_total'),
                line_count=estimated_lines
            )
            if str(doc['frame']) == self.notebook.select():
                self.update_status()
            return
        if result_kind == 'error':
            self.handle_background_file_error(doc, result.get('error'))
            return
        if result_kind == 'text_index_error':
            if str(doc.get('background_index_token') or '') == result_token:
                doc['background_index_future'] = None
                doc['background_index_active'] = False
                doc['background_index_token'] = None
            self.log_exception("background text index", result.get('error'))
            return
        if result_kind == 'text_index':
            if str(doc.get('background_index_token') or '') == result_token:
                doc['background_index_future'] = None
                doc['background_index_active'] = False
                doc['background_index_token'] = None
            total_lines = max(1, int(result.get('total_lines') or doc.get('background_lines_loaded', 1) or 1))
            doc['background_lines_loaded'] = max(total_lines, int(doc.get('background_lines_loaded', 1) or 1))
            sample_step = max(1, int(result.get('sample_step') or max(1, total_lines // self.minimap_max_segments)))
            segment_max_lengths = list(result.get('segment_max_lengths') or [0])
            doc['minimap_progressive_state'] = None
            doc['minimap_model'] = self.create_minimap_model(total_lines, sample_step, segment_max_lengths, complete=True)
            doc['minimap_model_dirty'] = False
            if not doc.get('background_loading'):
                doc['total_file_lines'] = total_lines
                doc['window_end_line'] = max(int(doc.get('window_end_line', 1) or 1), total_lines)
            self.update_doc_load_progress(doc, line_count=doc.get('background_lines_loaded'))
            self.schedule_minimap_refresh(doc)
            if str(doc['frame']) == self.notebook.select():
                self.update_status()
            return
        if result_kind == 'virtual':
            if str(doc.get('background_index_token') or '') == result_token:
                doc['background_index_future'] = None
                doc['background_index_active'] = False
                doc['background_index_token'] = None
            doc['line_starts'] = result.get('line_starts') or [0]
            doc['file_size_bytes'] = int(result.get('file_size_bytes') or 0)
            doc['total_file_lines'] = max(1, len(doc['line_starts']))
            self.initialize_virtual_backing_store(doc)
            doc['minimap_model'] = result.get('minimap_model')
            doc['minimap_model_dirty'] = not bool(doc.get('minimap_model'))
            doc['minimap_progressive_state'] = None
            doc['background_loading'] = False
            doc['background_load_kind'] = None
            doc['background_load_file_path'] = None
            doc['background_load_token'] = None
            doc['background_bytes_loaded'] = doc['file_size_bytes']
            doc['background_bytes_total'] = doc['file_size_bytes']
            doc['background_lines_loaded'] = doc['total_file_lines']
            doc['window_start_line'] = 1
            doc['window_end_line'] = 1
            self.update_doc_load_progress(doc)
            self.close_doc_load_progress(doc)
            self.end_doc_load(doc)
            pending_target = doc.pop('pending_virtual_target_line', None)
            if pending_target == 'end':
                target_line = doc['total_file_lines']
            else:
                target_line = int(pending_target or doc.get('last_virtual_line', 1) or 1)
            self.load_virtual_window(doc, target_line)
            self.refresh_tab_title(doc['frame'])
            if str(doc['frame']) == self.notebook.select():
                self.update_status()
            doc['background_open_new_tab'] = False
            return
        if result_kind == 'text_progress':
            self.append_background_text_chunk(
                doc,
                '',
                bytes_loaded=result.get('bytes_loaded'),
                total_bytes=result.get('file_size_bytes'),
                total_lines_hint=result.get('total_lines_hint')
            )
            doc['background_open_new_tab'] = False
            return
        if result_kind == 'text_chunk':
            self.append_background_text_chunk(
                doc,
                result.get('chunk_text') or '',
                bytes_loaded=result.get('bytes_loaded'),
                total_bytes=result.get('file_size_bytes'),
                total_lines_hint=result.get('total_lines_hint')
            )
            doc['background_open_new_tab'] = False
            return
        if result_kind == 'text_complete':
            self.finalize_background_text_load(
                doc,
                bytes_loaded=result.get('bytes_loaded'),
                total_bytes=result.get('file_size_bytes'),
                total_lines_hint=result.get('total_lines_hint')
            )
            doc['background_open_new_tab'] = False
            return
        if result_kind == 'text':
            try:
                doc['file_size_bytes'] = int(result.get('file_size_bytes') or 0)
            except (TypeError, ValueError):
                doc['file_size_bytes'] = 0
            self.begin_background_text_insert(
                doc,
                result.get('content') or '',
                result.get('total_lines_hint')
            )
            doc['background_open_new_tab'] = False

    def write_encrypted_text_file(self, file_path, text_content, passphrase=None, header=None, key=None, original_name=None):
        if not self.encryption_available():
            raise RuntimeError(self.tr('encryption.error.support_unavailable', 'Encryption support is unavailable.'))
        encryption_header = dict(header or {})
        if key is None:
            encryption_header = self.create_encryption_header(original_name=original_name or file_path)
            key = self.derive_encryption_key(passphrase, encryption_header)
        else:
            if not encryption_header:
                raise ValueError(self.tr('encryption.error.metadata_missing', 'Encrypted file metadata is missing.'))
            encryption_header['original_name'] = encryption_header.get('original_name') or os.path.basename(original_name or file_path)
        plaintext_bytes = str(text_content).encode('utf-8')
        nonce = os.urandom(self.encryption_nonce_length)
        ciphertext = AESGCM(key).encrypt(nonce, plaintext_bytes, self.encryption_magic)
        payload_bytes = self.build_encrypted_payload(encryption_header, nonce, ciphertext)
        if not self.write_binary_atomically(file_path, payload_bytes, 'notepadx-encrypted-', 'write encrypted file'):
            raise OSError(
                self.tr(
                    'encryption.error.write_failed',
                    'Could not write encrypted file: {file_path}',
                    file_path=file_path
                )
            )
        return encryption_header, key

    def read_encrypted_text_file(self, file_path):
        if not self.file_looks_encrypted(file_path):
            return None
        self.trace_startup(f"read_encrypted_text_file detected={file_path}")
        if not self.encryption_available():
            self.show_encryption_unavailable(self.root)
            raise OSError(self.tr('encryption.error.support_unavailable', 'Encryption support is unavailable.'))
        with open(file_path, 'rb') as encrypted_file:
            payload_bytes = encrypted_file.read()
        header, nonce, ciphertext = self.parse_encrypted_payload(payload_bytes)
        while True:
            self.trace_startup(f"read_encrypted_text_file prompting={file_path}")
            passphrase = self.prompt_open_passphrase(file_path, parent=self.root)
            if passphrase is None:
                self.trace_startup(f"read_encrypted_text_file cancelled={file_path}")
                raise EncryptedFileOpenCancelled(
                    self.tr('encryption.error.open_cancelled', 'Encrypted file open cancelled.')
                )
            try:
                key = self.derive_encryption_key(passphrase, header)
                plaintext_bytes = AESGCM(key).decrypt(nonce, ciphertext, self.encryption_magic)
                self.trace_startup(f"read_encrypted_text_file unlocked={file_path} bytes={len(plaintext_bytes)}")
                return plaintext_bytes.decode(header.get('encoding') or 'utf-8'), header, key
            except RuntimeError:
                raise
            except (InvalidTag, ValueError, UnicodeDecodeError):
                self.trace_startup(f"read_encrypted_text_file retry={file_path}")
                retry = messagebox.askretrycancel(
                    self.tr('encryption.open_title', 'Open Encrypted File'),
                    self.tr('encryption.unlock_failed', 'That passphrase did not unlock the encrypted file.'),
                    parent=self.root
                )
                if not retry:
                    raise EncryptedFileOpenCancelled(
                        self.tr('encryption.error.open_cancelled', 'Encrypted file open cancelled.')
                    )

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
        if doc and doc.get('is_remote'):
            return True
        file_path = doc.get('file_path') if doc else None
        if not file_path or not os.path.exists(file_path):
            return True
        current_signature = self.get_file_signature(file_path)
        known_signature = doc.get('file_signature')
        if known_signature is None or current_signature is None or current_signature == known_signature:
            return True
        answer = messagebox.askyesno(
            self.tr('file.changed_title', 'File Changed on Disk'),
            self.tr(
                'file.changed_message',
                'This file changed on disk after it was opened.\n\nOverwrite the newer disk version with your current editor contents?'
            ),
            parent=self.root
        )
        if answer:
            return True
        refreshed = messagebox.askyesno(
            self.tr('file.reload_title', 'Reload File'),
            self.tr('file.reload_message', 'Do you want Notepad-X to reload the file from disk instead?'),
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
        location = os.path.abspath(file_path) if file_path else self.tr('filesystem.unknown_path', 'that path')
        messagebox.showerror(
            title,
            self.tr('filesystem.access_error', 'Notepad-X could not access:\n{location}\n\n{error_detail}', location=location, error_detail=exc),
            parent=self.root
        )

    def slugify_storage_name(self, value):
        slug = re.sub(r'[^a-z0-9._-]+', '-', str(value or '').strip().lower()).strip('-')
        return slug or 'doc'

    def get_doc_persistence_path(self, doc):
        if not doc:
            return None
        return doc.get('remote_shadow_path') or doc.get('file_path')

    def is_virtual_editable(self, doc):
        return bool(doc and doc.get('virtual_mode') and doc.get('virtual_editable'))

    def is_doc_text_readonly(self, doc):
        return bool(doc and (doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc))))

    def doc_has_unsaved_changes(self, doc):
        if not doc:
            return False
        modified = False
        text_widget = doc.get('text')
        if text_widget:
            try:
                modified = bool(text_widget.edit_modified())
            except tk.TclError:
                modified = False
        if self.is_virtual_editable(doc):
            return bool(modified or doc.get('virtual_doc_dirty'))
        return modified

    def reset_virtual_backing_store(self, doc, remove_files=True):
        if not doc:
            return
        add_buffer_path = doc.get('virtual_add_buffer_path')
        if remove_files and add_buffer_path and os.path.exists(add_buffer_path):
            try:
                os.remove(add_buffer_path)
            except OSError:
                pass
        doc['virtual_editable'] = False
        doc['virtual_doc_dirty'] = False
        doc['virtual_piece_table'] = []
        doc['virtual_add_buffer_path'] = None
        doc['virtual_add_buffer_size'] = 0
        doc['virtual_revision'] = 0
        doc['virtual_source_path'] = None
        doc['virtual_window_start_byte'] = 0
        doc['virtual_window_end_byte'] = 0
        doc['virtual_hot_chunk_cache'] = OrderedDict()
        doc['virtual_cold_chunk_cache'] = OrderedDict()

    def initialize_virtual_backing_store(self, doc):
        if not doc:
            return
        file_size = max(0, int(doc.get('file_size_bytes', 0) or 0))
        self.reset_virtual_backing_store(doc, remove_files=True)
        doc['virtual_editable'] = True
        doc['virtual_source_path'] = self.get_doc_persistence_path(doc)
        doc['virtual_piece_table'] = [{'source': 'file', 'start': 0, 'length': file_size}] if file_size > 0 else []
        doc['virtual_hot_chunk_cache'] = OrderedDict()
        doc['virtual_cold_chunk_cache'] = OrderedDict()

    def clear_virtual_window_caches(self, doc):
        if not doc:
            return
        doc['virtual_hot_chunk_cache'] = OrderedDict()
        doc['virtual_cold_chunk_cache'] = OrderedDict()

    def build_virtual_window_cache_key(self, doc, start_byte, end_byte):
        return (
            int(doc.get('virtual_revision', 0) or 0),
            max(0, int(start_byte or 0)),
            max(0, int(end_byte or 0)),
        )

    def store_virtual_window_cache_text(self, doc, cache_key, text_value):
        if not doc:
            return
        hot_cache = doc.get('virtual_hot_chunk_cache')
        if not isinstance(hot_cache, OrderedDict):
            hot_cache = OrderedDict()
            doc['virtual_hot_chunk_cache'] = hot_cache
        hot_cache[cache_key] = text_value
        hot_cache.move_to_end(cache_key)
        while len(hot_cache) > self.virtual_hot_chunk_cache_entries:
            evicted_key, evicted_text = hot_cache.popitem(last=False)
            if not self.virtual_cold_chunk_cache_enabled:
                continue
            cold_cache = doc.get('virtual_cold_chunk_cache')
            if not isinstance(cold_cache, OrderedDict):
                cold_cache = OrderedDict()
                doc['virtual_cold_chunk_cache'] = cold_cache
            try:
                cold_cache[evicted_key] = zlib.compress(evicted_text.encode('utf-8'), level=1)
                cold_cache.move_to_end(evicted_key)
            except Exception:
                continue
            while len(cold_cache) > self.virtual_cold_chunk_cache_entries:
                cold_cache.popitem(last=False)

    def get_virtual_window_cache_text(self, doc, cache_key):
        if not doc:
            return None
        hot_cache = doc.get('virtual_hot_chunk_cache')
        if isinstance(hot_cache, OrderedDict) and cache_key in hot_cache:
            hot_cache.move_to_end(cache_key)
            return hot_cache[cache_key]
        cold_cache = doc.get('virtual_cold_chunk_cache')
        if isinstance(cold_cache, OrderedDict) and cache_key in cold_cache:
            try:
                cached_text = zlib.decompress(cold_cache.pop(cache_key)).decode('utf-8', errors='replace')
                self.store_virtual_window_cache_text(doc, cache_key, cached_text)
                return cached_text
            except Exception:
                cold_cache.pop(cache_key, None)
        return None

    def ensure_virtual_add_buffer_path(self, doc):
        if not doc:
            return None
        existing_path = doc.get('virtual_add_buffer_path')
        if existing_path and os.path.exists(existing_path):
            return existing_path
        temp_dir = os.path.join(self.get_emergency_support_dir(), 'virtual-add')
        os.makedirs(temp_dir, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(prefix='notepadx-virtual-', suffix='.bin', dir=temp_dir)
        os.close(fd)
        doc['virtual_add_buffer_path'] = temp_path
        doc['virtual_add_buffer_size'] = 0
        return temp_path

    def append_virtual_add_bytes(self, doc, payload_bytes):
        data = bytes(payload_bytes or b'')
        if not data:
            return None
        add_buffer_path = self.ensure_virtual_add_buffer_path(doc)
        start_offset = max(0, int(doc.get('virtual_add_buffer_size', 0) or 0))
        with open(add_buffer_path, 'ab') as add_file:
            add_file.write(data)
            add_file.flush()
            os.fsync(add_file.fileno())
        doc['virtual_add_buffer_size'] = start_offset + len(data)
        return start_offset

    def iter_virtual_piece_segments(self, doc, start_byte, end_byte):
        logical_offset = 0
        safe_start = max(0, int(start_byte or 0))
        safe_end = max(safe_start, int(end_byte or safe_start))
        for piece in doc.get('virtual_piece_table') or []:
            piece_length = max(0, int(piece.get('length', 0) or 0))
            if piece_length <= 0:
                continue
            piece_end = logical_offset + piece_length
            if piece_end <= safe_start:
                logical_offset = piece_end
                continue
            if logical_offset >= safe_end:
                break
            local_start = max(0, safe_start - logical_offset)
            local_end = min(piece_length, safe_end - logical_offset)
            if local_end > local_start:
                yield piece, local_start, local_end - local_start
            logical_offset = piece_end

    def read_virtual_byte_range_text(self, doc, start_byte, end_byte):
        safe_start = max(0, int(start_byte or 0))
        safe_end = max(safe_start, int(end_byte or safe_start))
        cache_key = self.build_virtual_window_cache_key(doc, safe_start, safe_end)
        cached_text = self.get_virtual_window_cache_text(doc, cache_key)
        if cached_text is not None:
            return cached_text

        chunk_parts = []
        open_handles = {}
        try:
            for piece, local_offset, span_length in self.iter_virtual_piece_segments(doc, safe_start, safe_end):
                source_name = piece.get('source')
                source_path = (doc.get('virtual_source_path') or doc.get('file_path')) if source_name == 'file' else doc.get('virtual_add_buffer_path')
                if not source_path:
                    continue
                handle = open_handles.get(source_path)
                if handle is None:
                    handle = open(source_path, 'rb')
                    open_handles[source_path] = handle
                handle.seek(max(0, int(piece.get('start', 0) or 0)) + local_offset)
                remaining = span_length
                while remaining > 0:
                    chunk = handle.read(min(self.file_load_chunk_size, remaining))
                    if not chunk:
                        break
                    chunk_parts.append(chunk)
                    remaining -= len(chunk)
        finally:
            for handle in open_handles.values():
                try:
                    handle.close()
                except OSError:
                    pass

        decoded_text = b''.join(chunk_parts).decode('utf-8', errors='replace')
        self.store_virtual_window_cache_text(doc, cache_key, decoded_text)
        return decoded_text

    def merge_virtual_piece_table(self, pieces):
        merged = []
        for piece in pieces or []:
            piece_length = max(0, int(piece.get('length', 0) or 0))
            if piece_length <= 0:
                continue
            normalized_piece = {
                'source': piece.get('source'),
                'start': max(0, int(piece.get('start', 0) or 0)),
                'length': piece_length,
            }
            if (
                merged and
                merged[-1]['source'] == normalized_piece['source'] and
                merged[-1]['start'] + merged[-1]['length'] == normalized_piece['start']
            ):
                merged[-1]['length'] += normalized_piece['length']
            else:
                merged.append(normalized_piece)
        return merged

    def replace_virtual_byte_range(self, doc, start_byte, end_byte, replacement_piece=None):
        safe_start = max(0, int(start_byte or 0))
        safe_end = max(safe_start, int(end_byte or safe_start))
        current_pieces = list(doc.get('virtual_piece_table') or [])
        updated_pieces = []
        inserted = False
        logical_offset = 0
        for piece in current_pieces:
            piece_length = max(0, int(piece.get('length', 0) or 0))
            if piece_length <= 0:
                continue
            piece_start = logical_offset
            piece_end = piece_start + piece_length
            logical_offset = piece_end

            if piece_end <= safe_start or piece_start >= safe_end:
                if not inserted and piece_start >= safe_end and replacement_piece is not None:
                    updated_pieces.append(dict(replacement_piece))
                    inserted = True
                updated_pieces.append(dict(piece))
                continue

            if piece_start < safe_start:
                updated_pieces.append({
                    'source': piece.get('source'),
                    'start': int(piece.get('start', 0) or 0),
                    'length': safe_start - piece_start,
                })
            if not inserted and replacement_piece is not None:
                updated_pieces.append(dict(replacement_piece))
                inserted = True
            if piece_end > safe_end:
                suffix_offset = safe_end - piece_start
                updated_pieces.append({
                    'source': piece.get('source'),
                    'start': int(piece.get('start', 0) or 0) + suffix_offset,
                    'length': piece_end - safe_end,
                })

        if not inserted and replacement_piece is not None:
            updated_pieces.append(dict(replacement_piece))
        doc['virtual_piece_table'] = self.merge_virtual_piece_table(updated_pieces)

    def update_virtual_line_index_after_replace(self, doc, start_line, end_line, start_byte, old_end_byte, replacement_bytes):
        line_starts = list(doc.get('line_starts') or [0])
        total_lines = max(1, int(doc.get('total_file_lines', len(line_starts)) or len(line_starts)))
        safe_start_line = max(1, min(total_lines, int(start_line or 1)))
        safe_end_line = max(safe_start_line, min(total_lines, int(end_line or safe_start_line)))
        start_index = safe_start_line - 1
        next_line_index = safe_end_line

        prefix = line_starts[:start_index + 1]
        replacement_new_starts = []
        for newline_offset, byte_value in enumerate(replacement_bytes or b''):
            if byte_value == 10:
                replacement_new_starts.append(start_byte + newline_offset + 1)

        byte_delta = len(replacement_bytes or b'') - max(0, old_end_byte - start_byte)
        shifted_suffix = [int(line_start) + byte_delta for line_start in line_starts[next_line_index:]]

        doc['line_starts'] = prefix + replacement_new_starts + shifted_suffix
        if not doc['line_starts']:
            doc['line_starts'] = [0]
        doc['file_size_bytes'] = max(0, int(doc.get('file_size_bytes', 0) or 0) + byte_delta)
        doc['total_file_lines'] = max(1, len(doc['line_starts']))

    def flush_virtual_window_edits(self, doc, force=False):
        if not self.is_virtual_editable(doc):
            return True
        if not self.is_virtual_index_ready(doc):
            return False
        text_widget = doc.get('text')
        if not text_widget:
            return False
        try:
            window_dirty = bool(text_widget.edit_modified())
        except tk.TclError:
            return False
        if not force and not window_dirty:
            return True

        start_line = max(1, int(doc.get('window_start_line', 1) or 1))
        end_line = max(start_line, int(doc.get('window_end_line', start_line) or start_line))
        line_starts = doc.get('line_starts') or [0]
        if start_line - 1 >= len(line_starts):
            return False
        start_byte = int(line_starts[start_line - 1])
        old_end_byte = self.get_virtual_line_end_byte(doc, end_line)
        try:
            replacement_text = text_widget.get('1.0', 'end-1c')
            insert_line, insert_col = [int(part) for part in text_widget.index(tk.INSERT).split('.')]
        except (tk.TclError, ValueError):
            replacement_text = ''
            insert_line = 1
            insert_col = 0
        replacement_bytes = replacement_text.encode('utf-8')
        replacement_piece = None
        replacement_piece_start = self.append_virtual_add_bytes(doc, replacement_bytes)
        if replacement_piece_start is not None:
            replacement_piece = {
                'source': 'add',
                'start': replacement_piece_start,
                'length': len(replacement_bytes),
            }

        self.replace_virtual_byte_range(doc, start_byte, old_end_byte, replacement_piece=replacement_piece)
        self.update_virtual_line_index_after_replace(doc, start_line, end_line, start_byte, old_end_byte, replacement_bytes)
        doc['virtual_doc_dirty'] = True
        doc['virtual_revision'] = int(doc.get('virtual_revision', 0) or 0) + 1
        self.clear_virtual_window_caches(doc)
        new_end_byte = start_byte + len(replacement_bytes)
        doc['virtual_window_start_byte'] = start_byte
        doc['virtual_window_end_byte'] = new_end_byte
        new_window_line_count = max(1, replacement_bytes.count(b'\n') + 1)
        doc['window_start_line'] = start_line
        doc['window_end_line'] = min(doc['total_file_lines'], start_line + new_window_line_count - 1)
        doc['last_virtual_line'] = max(1, min(doc['total_file_lines'], start_line + insert_line - 1))
        doc['last_virtual_col'] = max(0, insert_col)
        self.store_virtual_window_cache_text(
            doc,
            self.build_virtual_window_cache_key(doc, start_byte, new_end_byte),
            replacement_text
        )
        doc['minimap_model'] = None
        doc['minimap_model_dirty'] = True
        doc['symbol_cache_signature'] = None
        doc['symbol_cache'] = None
        try:
            text_widget.edit_modified(False)
        except tk.TclError:
            pass
        self.refresh_tab_title(doc['frame'])
        if str(doc['frame']) == self.notebook.select():
            self.update_status()
        return True

    def iter_virtual_document_chunks(self, doc, chunk_size=None):
        if not doc:
            return
        active_chunk_size = max(64 * 1024, int(chunk_size or self.file_load_chunk_size or 0))
        open_handles = {}
        try:
            for piece in doc.get('virtual_piece_table') or []:
                piece_length = max(0, int(piece.get('length', 0) or 0))
                if piece_length <= 0:
                    continue
                source_name = piece.get('source')
                source_path = (doc.get('virtual_source_path') or doc.get('file_path')) if source_name == 'file' else doc.get('virtual_add_buffer_path')
                if not source_path:
                    continue
                handle = open_handles.get(source_path)
                if handle is None:
                    handle = open(source_path, 'rb')
                    open_handles[source_path] = handle
                handle.seek(max(0, int(piece.get('start', 0) or 0)))
                remaining = piece_length
                while remaining > 0:
                    chunk = handle.read(min(active_chunk_size, remaining))
                    if not chunk:
                        break
                    remaining -= len(chunk)
                    yield chunk
        finally:
            for handle in open_handles.values():
                try:
                    handle.close()
                except OSError:
                    pass

    def write_virtual_document_atomically(self, doc, file_path):
        directory = os.path.dirname(file_path) or '.'
        os.makedirs(directory, exist_ok=True)
        target_mode = None
        temp_path = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode)
            except OSError:
                target_mode = 0o664
        try:
            fd, temp_path = tempfile.mkstemp(prefix='notepadx-save-', suffix='.tmp', dir=directory)
            with os.fdopen(fd, 'wb') as temp_file:
                for chunk in self.iter_virtual_document_chunks(doc):
                    temp_file.write(chunk)
                temp_file.flush()
                os.fsync(temp_file.fileno())
            os.replace(temp_path, file_path)
            if target_mode is not None:
                os.chmod(file_path, target_mode)
            return True
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def rebase_virtual_document_after_save(self, doc, persisted_path=None):
        if not self.is_virtual_editable(doc):
            return
        base_path = persisted_path or doc.get('file_path')
        file_size = 0
        if base_path:
            try:
                file_size = os.path.getsize(base_path)
            except OSError:
                file_size = max(0, int(doc.get('file_size_bytes', 0) or 0))
        add_buffer_path = doc.get('virtual_add_buffer_path')
        if add_buffer_path and os.path.exists(add_buffer_path):
            try:
                os.remove(add_buffer_path)
            except OSError:
                pass
        doc['virtual_add_buffer_path'] = None
        doc['virtual_add_buffer_size'] = 0
        doc['virtual_piece_table'] = [{'source': 'file', 'start': 0, 'length': file_size}] if file_size > 0 else []
        doc['virtual_doc_dirty'] = False
        doc['virtual_revision'] = int(doc.get('virtual_revision', 0) or 0) + 1
        doc['virtual_source_path'] = base_path
        self.clear_virtual_window_caches(doc)
        doc['file_size_bytes'] = file_size
        if self.is_virtual_index_ready(doc):
            start_line = max(1, min(doc['total_file_lines'], int(doc.get('window_start_line', 1) or 1)))
            end_line = max(start_line, min(doc['total_file_lines'], int(doc.get('window_end_line', start_line) or start_line)))
            doc['virtual_window_start_byte'] = int((doc.get('line_starts') or [0])[start_line - 1])
            doc['virtual_window_end_byte'] = self.get_virtual_line_end_byte(doc, end_line)
        try:
            doc['text'].edit_modified(False)
        except tk.TclError:
            pass

    def save_virtual_document(self, doc, autosave=False, show_errors=True, update_recent=True):
        if not doc:
            return False
        if not doc.get('file_path'):
            if autosave:
                return False
            return self.save_as()
        if not self.flush_virtual_window_edits(doc, force=True):
            return False
        if not autosave and not doc.get('is_remote'):
            if not self.confirm_external_file_change(doc):
                return False
        try:
            if not autosave:
                self.create_backup_snapshot(doc)
            if doc.get('is_remote'):
                shadow_path = doc.get('remote_shadow_path') or doc.get('file_path')
                if not shadow_path:
                    raise OSError('Remote document metadata is incomplete')
                self.write_virtual_document_atomically(doc, shadow_path)
                remote_spec = doc.get('remote_spec')
                scp_path = self.get_scp_executable()
                if not remote_spec or not scp_path:
                    raise OSError('scp is unavailable')
                completed = subprocess.run(
                    [scp_path, '-q', shadow_path, remote_spec],
                    capture_output=True,
                    text=True,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
                if completed.returncode != 0:
                    error_detail = (completed.stderr or completed.stdout or 'Unknown scp failure').strip()
                    raise OSError(error_detail)
                self.rebase_virtual_document_after_save(doc, shadow_path)
            else:
                self.write_virtual_document_atomically(doc, doc['file_path'])
                self.rebase_virtual_document_after_save(doc, doc['file_path'])
        except (PermissionError, RuntimeError, ValueError, OSError) as exc:
            if show_errors:
                if isinstance(exc, PermissionError):
                    self.show_filesystem_error(self.tr('save.failed_title', 'Save Failed'), doc.get('file_path'), exc)
                elif isinstance(exc, (RuntimeError, ValueError)):
                    messagebox.showerror(self.tr('save.failed_title', 'Save Failed'), str(exc), parent=self.root)
                else:
                    self.log_exception('save virtual document', exc)
                    self.show_filesystem_error(self.tr('save.failed_title', 'Save Failed'), doc.get('file_path'), exc)
            return False

        self.update_doc_file_signature(doc)
        if update_recent and not doc.get('is_remote'):
            self.add_recent_file(doc['file_path'])
        self.refresh_tab_title(doc['frame'])
        self.update_status()
        self.save_session()
        return True

    def clear_remote_metadata(self, doc):
        if not doc:
            return
        previous_spec = doc.get('remote_spec')
        doc['is_remote'] = False
        doc['remote_spec'] = None
        doc['remote_host'] = None
        doc['remote_path'] = None
        doc['remote_shadow_path'] = None
        if doc.get('display_name') and doc.get('display_name') == previous_spec:
            doc['display_name'] = None

    def remote_tools_available(self, notify=False):
        if self.get_scp_executable():
            return True
        if notify:
            messagebox.showerror(
                self.tr('menu.file.open', 'Open'),
                'OpenSSH scp was not found on this system. Remote editing needs scp in PATH.',
                parent=self.root
            )
        return False

    def get_scp_executable(self):
        scp_path = shutil.which('scp')
        if not scp_path:
            return None
        normalized_path = os.path.abspath(str(scp_path).strip())
        if not normalized_path or any(char in normalized_path for char in '\r\n\0'):
            return None
        return normalized_path

    def parse_remote_spec(self, spec_text):
        spec = str(spec_text or '').strip()
        if any(char in spec for char in '\r\n\0'):
            raise ValueError('Remote file paths cannot contain control characters')
        match = re.match(r'^(?P<host>[^:\s/][^:\s]*):(?P<path>/.*)$', spec)
        if not match:
            raise ValueError('Remote files must use the format user@host:/absolute/path/to/file')
        host = match.group('host').strip()
        path = match.group('path').strip()
        if not host or not path or path == '/':
            raise ValueError('Remote files must include both a host and an absolute file path')
        if host.startswith('-'):
            raise ValueError('Remote host names cannot start with a dash')
        return {'spec': f'{host}:{path}', 'host': host, 'path': path}

    def build_remote_shadow_path(self, remote_spec):
        parsed = self.parse_remote_spec(remote_spec)
        remote_name = os.path.basename(parsed['path']) or 'remote.txt'
        host_slug = self.slugify_storage_name(parsed['host'])
        path_hash = hashlib.sha1(parsed['spec'].encode('utf-8', errors='replace')).hexdigest()[:12]
        shadow_dir = os.path.join(self.remote_cache_dir, host_slug)
        os.makedirs(shadow_dir, exist_ok=True)
        return os.path.join(shadow_dir, f'{path_hash}-{remote_name}')

    def fetch_remote_file_to_shadow(self, remote_spec, shadow_path):
        scp_path = self.get_scp_executable()
        if not scp_path:
            self.remote_tools_available(notify=True)
            raise OSError('scp is unavailable')
        os.makedirs(os.path.dirname(shadow_path), exist_ok=True)
        completed = subprocess.run(
            [scp_path, '-q', remote_spec, shadow_path],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        )
        if completed.returncode != 0:
            error_detail = (completed.stderr or completed.stdout or 'Unknown scp failure').strip()
            raise OSError(error_detail)
        return shadow_path

    def open_remote_file_dialog(self, event=None):
        remote_spec = self.prompt_text_input(
            self.tr('menu.file.open', 'Open'),
            'Remote file (user@host:/absolute/path):',
            parent=self.root
        )
        if remote_spec:
            self.open_remote_file(remote_spec)
        return "break"

    def open_remote_file(self, remote_spec):
        try:
            parsed = self.parse_remote_spec(remote_spec)
        except ValueError as exc:
            messagebox.showerror(self.tr('menu.file.open', 'Open'), str(exc), parent=self.root)
            return False

        for doc in self.documents.values():
            if doc.get('is_remote') and doc.get('remote_spec') == parsed['spec']:
                self.notebook.select(doc['frame'])
                self.set_active_document(doc['frame'])
                return True

        shadow_path = self.build_remote_shadow_path(parsed['spec'])
        try:
            self.fetch_remote_file_to_shadow(parsed['spec'], shadow_path)
        except OSError as exc:
            self.log_exception('open remote file', exc)
            messagebox.showerror(
                self.tr('menu.file.open', 'Open'),
                f'Notepad-X could not fetch the remote file.\n\n{exc}',
                parent=self.root
            )
            return False

        current_doc = self.get_current_doc()
        if current_doc and not current_doc['file_path'] and not current_doc['text'].edit_modified():
            target_doc = current_doc
            target_doc['file_path'] = shadow_path
            target_doc['background_open_new_tab'] = False
            if not self.load_content_into_doc(target_doc, shadow_path):
                self.cleanup_failed_file_open(target_doc)
                return False
        else:
            tab_id = self.create_tab(file_path=shadow_path, select=True)
            target_doc = self.documents[str(tab_id)]
            target_doc['background_open_new_tab'] = False
            if not self.load_content_into_doc(target_doc, shadow_path):
                try:
                    self.notebook.forget(target_doc['frame'])
                except tk.TclError:
                    pass
                self.documents.pop(str(target_doc['frame']), None)
                return False

        target_doc['is_remote'] = True
        target_doc['remote_spec'] = parsed['spec']
        target_doc['remote_host'] = parsed['host']
        target_doc['remote_path'] = parsed['path']
        target_doc['remote_shadow_path'] = shadow_path
        target_doc['display_name'] = parsed['spec']
        target_doc['untitled_name'] = None
        self.refresh_tab_title(target_doc['frame'])
        self.set_active_document(target_doc['frame'])
        self.save_session()
        return True

    def get_doc_backup_prefix(self, doc):
        source_name = self.get_doc_display_path(doc) or doc.get('untitled_name') or 'untitled'
        return self.slugify_storage_name(source_name)

    def trim_backup_history(self, backup_prefix):
        pattern = os.path.join(self.backup_dir, f'{backup_prefix}-*.bak')
        backups = sorted(glob.glob(pattern), key=lambda path: os.path.getmtime(path))
        while len(backups) > self.max_backup_versions_per_doc:
            old_path = backups.pop(0)
            try:
                os.remove(old_path)
            except OSError:
                continue

    def create_backup_snapshot(self, doc):
        if not doc:
            return None
        source_path = self.get_doc_persistence_path(doc)
        if not source_path or not os.path.exists(source_path):
            return None
        backup_prefix = self.get_doc_backup_prefix(doc)
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        backup_path = os.path.join(self.backup_dir, f'{backup_prefix}-{timestamp}.bak')
        try:
            shutil.copy2(source_path, backup_path)
        except OSError as exc:
            self.log_exception('create backup snapshot', exc)
            return None
        self.trim_backup_history(backup_prefix)
        return backup_path

    def save_remote_document(self, doc, text_content):
        remote_spec = doc.get('remote_spec')
        shadow_path = doc.get('remote_shadow_path') or doc.get('file_path')
        if not remote_spec or not shadow_path:
            raise OSError('Remote document metadata is incomplete')
        scp_path = self.get_scp_executable()
        if not scp_path:
            raise OSError('scp is unavailable')
        self.write_file_atomically(shadow_path, text_content)
        completed = subprocess.run(
            [scp_path, '-q', shadow_path, remote_spec],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        )
        if completed.returncode != 0:
            error_detail = (completed.stderr or completed.stdout or 'Unknown scp failure').strip()
            raise OSError(error_detail)
        return True

    def save_document_content(self, doc, autosave=False, show_errors=True, update_recent=True):
        if not doc:
            return False
        if doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
            if show_errors:
                messagebox.showinfo(
                    self.tr('large_file.title', 'Large File Mode'),
                    self.tr(
                        'large_file.save_disabled',
                        'This tab is opened in buffered large-file mode. Editing and saving are disabled for the full file view.'
                    ),
                    parent=self.root
                )
            return False
        if self.is_virtual_editable(doc):
            return self.save_virtual_document(doc, autosave=autosave, show_errors=show_errors, update_recent=update_recent)

        if not doc.get('file_path'):
            if autosave:
                return False
            return self.save_as()

        if not autosave and not doc.get('is_remote'):
            if not self.confirm_external_file_change(doc):
                return False

        try:
            text_content = doc['text'].get('1.0', tk.END).rstrip('\n')
            if doc.get('is_remote'):
                if not autosave:
                    self.create_backup_snapshot(doc)
                self.save_remote_document(doc, text_content)
            elif doc.get('encrypted_file'):
                if not autosave:
                    self.create_backup_snapshot(doc)
                self.write_encrypted_text_file(
                    doc['file_path'],
                    text_content,
                    header=doc.get('encryption_header'),
                    key=doc.get('encryption_key'),
                    original_name=doc.get('file_path')
                )
            else:
                if not autosave:
                    self.create_backup_snapshot(doc)
                self.write_file_atomically(doc['file_path'], text_content)
        except (PermissionError, RuntimeError, ValueError, OSError) as exc:
            if show_errors:
                if isinstance(exc, PermissionError):
                    self.show_filesystem_error(self.tr('save.failed_title', 'Save Failed'), doc.get('file_path'), exc)
                elif isinstance(exc, (RuntimeError, ValueError)):
                    messagebox.showerror(self.tr('save.failed_title', 'Save Failed'), str(exc), parent=self.root)
                else:
                    self.log_exception('save document content', exc)
                    self.show_filesystem_error(self.tr('save.failed_title', 'Save Failed'), doc.get('file_path'), exc)
            return False

        doc['text'].edit_modified(False)
        self.update_doc_file_signature(doc)
        if update_recent and not doc.get('is_remote'):
            self.add_recent_file(doc['file_path'])
        self.refresh_tab_title(doc['frame'])
        self.update_status()
        self.save_session()
        return True

    def cancel_doc_autosave(self, doc):
        if not doc:
            return
        autosave_job = doc.get('autosave_job')
        if autosave_job:
            try:
                self.root.after_cancel(autosave_job)
            except tk.TclError:
                pass
            doc['autosave_job'] = None

    def schedule_doc_autosave(self, doc):
        if not doc:
            return
        self.cancel_doc_autosave(doc)
        if not self.autosave_enabled.get():
            return
        if doc.get('large_file_mode'):
            return
        if not doc['text'].edit_modified():
            return
        if not doc.get('file_path'):
            return
        try:
            doc['autosave_job'] = self.root.after(self.autosave_delay_ms, lambda current=doc: self.run_doc_autosave(current))
        except tk.TclError:
            doc['autosave_job'] = None

    def run_doc_autosave(self, doc):
        if not doc:
            return
        doc['autosave_job'] = None
        if not self.autosave_enabled.get():
            return
        if not doc['text'].edit_modified():
            return
        self.save_document_content(doc, autosave=True, show_errors=False, update_recent=False)

    def confirm_exit_app(self):
        for doc in list(self.documents.values()):
            if not self.confirm_close_tab(doc):
                return False
        return True

    def finalize_exit_app(self):
        for doc in list(self.documents.values()):
            self.unregister_doc_from_shared_notes(doc)
            self.cancel_doc_autosave(doc)
            self.cancel_doc_background_index(doc)
        self.stop_single_instance_server()
        self.shutdown_index_process_executor()
        if self.recovery_job:
            try:
                self.root.after_cancel(self.recovery_job)
            except tk.TclError:
                pass
            self.recovery_job = None
        if self.window_layout_job:
            try:
                self.root.after_cancel(self.window_layout_job)
            except tk.TclError:
                pass
            self.window_layout_job = None
        if os.path.exists(self.recovery_path):
            try:
                os.remove(self.recovery_path)
            except OSError as exc:
                self.log_exception("remove recovery file on exit", exc)
        self.persist_editor_identity()
        self.save_session()

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

        find_history = self.sanitize_search_history_entries(session.get('find_history', []))
        find_in_history = self.sanitize_search_history_entries(session.get('find_in_history', []))
        command_history = self.sanitize_command_history_entries(session.get('command_history', []))

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
        try:
            command_panel_height = int(session.get('command_panel_height', self.command_panel_default_height))
        except (TypeError, ValueError):
            command_panel_height = self.command_panel_default_height
        command_panel_height = max(self.command_panel_min_height, min(self.command_panel_max_height, command_panel_height))
        try:
            window_width = int(session.get('window_width', self.default_window_width))
        except (TypeError, ValueError):
            window_width = self.default_window_width
        try:
            window_height = int(session.get('window_height', self.default_window_height))
        except (TypeError, ValueError):
            window_height = self.default_window_height
        window_width = max(240, min(10000, window_width))
        window_height = max(180, min(10000, window_height))
        window_state = 'zoomed' if str(session.get('window_state', 'normal')).strip().lower() == 'zoomed' else 'normal'

        return {
            'open_files': open_files,
            'selected_file': selected_file,
            'recent_files': recent_files,
            'find_history': find_history,
            'find_in_history': find_in_history,
            'command_history': command_history,
            'command_panel_height': command_panel_height,
            'window_width': window_width,
            'window_height': window_height,
            'window_state': window_state,
            'closed_session_files': closed_files,
            'sound_enabled': bool(session.get('sound_enabled', True)),
            'status_bar_enabled': bool(session.get('status_bar_enabled', True)),
            'numbered_lines_enabled': bool(session.get('numbered_lines_enabled', True)),
            'autocomplete_enabled': bool(session.get('autocomplete_enabled', True)),
            'spell_check_enabled': bool(session.get('spell_check_enabled', SpellChecker is not None)),
            'auto_pair_enabled': bool(session.get('auto_pair_enabled', True)),
            'compare_multi_edit_enabled': bool(session.get('compare_multi_edit_enabled', False)),
            'markdown_preview_enabled': bool(session.get('markdown_preview_enabled', False)),
            'sync_page_navigation_enabled': bool(session.get('sync_page_navigation_enabled', False)),
            'edit_with_shell_enabled': bool(session.get('edit_with_shell_enabled', False)),
            'minimap_enabled': bool(session.get('minimap_enabled', True)),
            'breadcrumbs_enabled': bool(session.get('breadcrumbs_enabled', True)),
            'diagnostics_enabled': bool(session.get('diagnostics_enabled', True)),
            'autosave_enabled': bool(session.get('autosave_enabled', True)),
            'current_font_size': current_font_size,
            'syntax_theme': syntax_theme,
            'locale_code': locale_code,
            'compare_file': compare_file,
            'compare_base_file': compare_base_file,
            'hotkey_overrides': self.sanitize_hotkey_overrides(session.get('hotkey_overrides', session.get('hotkeys', {}))),
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

    def reset_startup_trace(self):
        try:
            if os.path.exists(self.startup_trace_path):
                os.remove(self.startup_trace_path)
        except OSError:
            pass

    def trace_startup(self, message):
        try:
            timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
            with open(self.startup_trace_path, 'a', encoding='utf-8') as trace_file:
                trace_file.write(f"[{timestamp}] {message}\n")
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
            messagebox.showerror(
                self.tr('app.crash_title', 'Notepad-X Crash'),
                self.tr(
                    'app.crash_message',
                    'An unexpected error occurred.\nA crash log was written to:\n{crash_log_path}',
                    crash_log_path=self.crash_log_path
                ),
                parent=self.root
            )
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
                if not parent.winfo_viewable():
                    raise tk.TclError
                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                parent_width = max(parent.winfo_width(), parent.winfo_reqwidth())
                parent_height = max(parent.winfo_height(), parent.winfo_reqheight())
                if parent_width <= 1 or parent_height <= 1:
                    raise tk.TclError
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

        screen_width = max(1, window.winfo_screenwidth())
        screen_height = max(1, window.winfo_screenheight())
        x = max(0, min(int(x), max(0, screen_width - width)))
        y = max(0, min(int(y), max(0, screen_height - height)))
        window.geometry(f"{width}x{height}+{x}+{y}")

    def center_window_after_show(self, window, parent=None, attempts=3, delay_ms=60):
        if attempts <= 0:
            return
        try:
            if not window.winfo_exists():
                return
            self.center_window(window, parent)
            window.lift()
        except tk.TclError:
            return
        if attempts <= 1:
            return
        try:
            window.after(
                delay_ms,
                lambda current=window, current_parent=parent, remaining=attempts - 1, current_delay=delay_ms:
                    self.center_window_after_show(current, current_parent, remaining, current_delay)
            )
        except tk.TclError:
            return

    def get_root_window_state(self):
        try:
            state = str(self.root.state() or 'normal').strip().lower()
        except tk.TclError:
            state = 'normal'
        return 'zoomed' if state == 'zoomed' else 'normal'

    def get_window_layout_snapshot(self):
        width = max(240, int(getattr(self, 'window_width', self.default_window_width) or self.default_window_width))
        height = max(180, int(getattr(self, 'window_height', self.default_window_height) or self.default_window_height))
        state = self.get_root_window_state()
        if not self.fullscreen:
            try:
                current_width = int(self.root.winfo_width())
                current_height = int(self.root.winfo_height())
            except (tk.TclError, ValueError):
                current_width = current_height = 0
            if state == 'normal' and current_width > 240 and current_height > 180:
                width = current_width
                height = current_height
        return width, height, state

    def apply_saved_window_layout(self, width=None, height=None, state='normal'):
        try:
            width = int(width if width is not None else self.default_window_width)
        except (TypeError, ValueError):
            width = self.default_window_width
        try:
            height = int(height if height is not None else self.default_window_height)
        except (TypeError, ValueError):
            height = self.default_window_height
        width = max(240, width)
        height = max(180, height)
        state = 'zoomed' if str(state or 'normal').strip().lower() == 'zoomed' else 'normal'
        try:
            screen_width = int(self.root.winfo_screenwidth() or width)
            screen_height = int(self.root.winfo_screenheight() or height)
        except (tk.TclError, ValueError):
            screen_width = width
            screen_height = height
        width = min(width, max(240, screen_width))
        height = min(height, max(180, screen_height))
        self.window_width = width
        self.window_height = height
        self.window_state = state
        try:
            self.root.state('normal')
        except tk.TclError:
            pass
        try:
            self.root.geometry(f"{width}x{height}")
            self.root.update_idletasks()
        except tk.TclError:
            return False
        if state == 'zoomed' and not self.fullscreen:
            try:
                self.root.state('zoomed')
            except tk.TclError:
                try:
                    self.root.wm_state('zoomed')
                except tk.TclError:
                    self.window_state = 'normal'
        return True

    def schedule_window_layout_save(self):
        if self.isolated_session or self._shutdown_requested or not self.main_loop_started:
            return
        if self.window_layout_job:
            try:
                self.root.after_cancel(self.window_layout_job)
            except tk.TclError:
                pass
        try:
            self.window_layout_job = self.root.after(self.window_layout_save_delay_ms, self.save_session)
        except tk.TclError:
            self.window_layout_job = None

    def on_root_configure(self, event=None):
        if getattr(event, 'widget', None) not in (None, self.root):
            return None
        if self._shutdown_requested or self.fullscreen:
            return None
        width, height, state = self.get_window_layout_snapshot()
        self.window_state = state
        if state == 'normal':
            self.window_width = width
            self.window_height = height
        self.schedule_window_layout_save()
        return None

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
        if self._shutdown_requested:
            return
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
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
        if self._shutdown_requested:
            return
        try:
            self.root.after(150, self.process_remote_open_requests)
        except tk.TclError:
            pass

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
            self.schedule_minimap_refresh(doc)
        if self.compare_view:
            self.compare_view['text'].configure(font=font_tuple)
            self.update_line_number_gutter(self.compare_view)
            self.schedule_minimap_refresh(self.compare_view)
        if self.markdown_preview_enabled.get():
            self.schedule_markdown_preview_refresh()

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
            self.schedule_minimap_refresh(doc)
        if self.compare_view:
            self.compare_view['text'].configure(wrap=wrap_mode)
            self.schedule_minimap_refresh(self.compare_view)
        if self.markdown_preview_enabled.get():
            self.schedule_markdown_preview_refresh()

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
                return max(1, round(counters.WorkingSetSize / (1024 * 1024)))
        except Exception:
            pass
        return 0

    def get_process_cpu_usage_percent(self):
        current_sample = (time.process_time(), time.monotonic())
        previous_sample = self.last_cpu_sample
        self.last_cpu_sample = current_sample
        if previous_sample is None:
            return 0.0
        cpu_delta = max(0.0, current_sample[0] - previous_sample[0])
        wall_delta = max(0.0, current_sample[1] - previous_sample[1])
        if wall_delta <= 0.0:
            return self.cpu_used_percent
        normalized_percent = (cpu_delta / wall_delta) * 100.0 / max(1, self.cpu_sample_count)
        return max(0.0, min(100.0, normalized_percent))

    def update_memory_usage(self):
        self.cpu_used_percent = self.get_process_cpu_usage_percent()
        self.memory_used_mb = self.get_memory_usage_mb()
        self.update_status()
        self.root.after(1000, self.update_memory_usage)

    def trim_process_working_set(self):
        gc.collect()
        if not self.is_windows or not self.kernel32 or not self.psapi:
            self.cpu_used_percent = self.get_process_cpu_usage_percent()
            self.memory_used_mb = self.get_memory_usage_mb()
            self.update_status()
            return
        try:
            handle = self.kernel32.GetCurrentProcess()
            empty_working_set = getattr(self.psapi, 'EmptyWorkingSet', None)
            if empty_working_set:
                empty_working_set(handle)
        except Exception:
            pass
        self.cpu_used_percent = self.get_process_cpu_usage_percent()
        self.memory_used_mb = self.get_memory_usage_mb()
        self.update_status()

    def schedule_memory_trim(self):
        for delay_ms in (80, 300, 900):
            try:
                self.root.after(delay_ms, self.trim_process_working_set)
            except tk.TclError:
                break

    def is_blank_startup_tab_state(self):
        if len(self.documents) != 1:
            return False
        doc = next(iter(self.documents.values()), None)
        if not doc or doc.get('file_path') or doc.get('untitled_name') != 'Untitled 1':
            return False
        text_widget = doc.get('text')
        if not text_widget or not text_widget.winfo_exists():
            return False
        try:
            if text_widget.edit_modified():
                return False
            return text_widget.get('1.0', 'end-1c') == ''
        except tk.TclError:
            return False

    def maybe_trim_blank_startup_memory(self):
        if self.is_blank_startup_tab_state():
            self.trim_process_working_set()

    def root_is_ready_for_dialogs(self):
        if self._shutdown_requested:
            return False
        if not self.main_loop_started:
            return False
        try:
            return bool(self.root.winfo_exists() and self.root.winfo_viewable())
        except tk.TclError:
            return False

    def schedule_when_root_ready(self, callback, attempts=20, delay_ms=75):
        if self._shutdown_requested or callback is None:
            return

        def run_when_ready(remaining_attempts):
            if self._shutdown_requested:
                return
            if self.root_is_ready_for_dialogs() or remaining_attempts <= 0:
                self.trace_startup(
                    f"schedule_when_root_ready fired callback={getattr(callback, '__name__', repr(callback))} "
                    f"mainloop={self.main_loop_started} attempts_left={remaining_attempts}"
                )
                callback()
                return
            try:
                self.root.after(
                    delay_ms,
                    lambda next_attempts=remaining_attempts - 1: run_when_ready(next_attempts)
                )
            except tk.TclError:
                pass

        try:
            self.root.after_idle(lambda: run_when_ready(attempts))
        except tk.TclError:
            pass

    def schedule_blank_startup_memory_trim(self):
        for delay_ms in (180, 600, 1400):
            try:
                self.root.after(delay_ms, self.maybe_trim_blank_startup_memory)
            except tk.TclError:
                break

    def schedule_startup_file_opening(self):
        if not self.startup_files:
            return
        pending_files = list(self.startup_files)
        self.startup_files = []
        self.trace_startup(f"schedule_startup_file_opening pending_files={pending_files}")
        self.schedule_when_root_ready(lambda files=pending_files: self.open_startup_files(files))

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

        self.status_tail = tk.Label(
            self.status_left,
            text=self.tr('status.resource_initial', " | CPU: 0.0% Memory: 0MB"),
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

    def get_text_widget_char_count(self, text_widget, start_index='1.0', end_index='end-1c'):
        if not text_widget:
            return 0
        try:
            counts = text_widget.count(start_index, end_index, 'chars')
            if counts:
                return max(0, int(counts[0]))
        except (tk.TclError, TypeError, ValueError, IndexError):
            pass
        try:
            return len(text_widget.get(start_index, end_index))
        except tk.TclError:
            return 0

    def get_text_widget_selected_char_count(self, text_widget):
        if not text_widget:
            return 0
        try:
            sel_start = text_widget.index('sel.first')
            sel_end = text_widget.index('sel.last')
        except tk.TclError:
            return 0
        return self.get_text_widget_char_count(text_widget, sel_start, sel_end)

    def build_status_char_info_for_widget(self, text_widget):
        try:
            total_lines = max(1, int(text_widget.index('end-1c').split('.')[0]))
        except (tk.TclError, ValueError):
            total_lines = 1
        total_chars = self.get_text_widget_char_count(text_widget)
        selected_count = self.get_text_widget_selected_char_count(text_widget)
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
        return total_lines, total_chars, char_info

    def build_editor_status_text(self, doc, text_widget):
        row, col = text_widget.index(tk.INSERT).split('.')
        row = int(row)
        col = int(col) + 1

        if doc and doc.get('virtual_mode'):
            total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
            row = max(1, min(total_lines, doc.get('window_start_line', 1) + row - 1))
            total_chars = int(doc.get('file_size_bytes', 0) or 0)
            char_info = self.tr(
                'status.byte_count',
                '{total_bytes} {bytes_label}',
                total_bytes=f"{total_chars:,}",
                bytes_label=self.tr('status.bytes', 'bytes')
            )
        elif doc and doc.get('large_file_mode'):
            try:
                total_lines = max(1, int(text_widget.index('end-1c').split('.')[0]))
            except (tk.TclError, ValueError):
                total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
            doc['total_file_lines'] = total_lines
            total_chars = max(0, int(doc.get('file_size_bytes', 0) or 0))
            char_info = self.tr(
                'status.byte_count',
                '{total_bytes} {bytes_label}',
                total_bytes=f"{total_chars:,}",
                bytes_label=self.tr('status.bytes', 'bytes')
            )
        else:
            total_lines, total_chars, char_info = self.build_status_char_info_for_widget(text_widget)

        zoom_text = self.get_zoom_text()
        mode_suffix = ""
        if doc and doc.get('virtual_mode'):
            mode_suffix = f" | {self.tr('status.mode.virtual_editable', 'Paged') if self.is_virtual_editable(doc) else self.tr('status.mode.virtual', 'Virtual')}"
        elif doc and doc.get('preview_mode'):
            mode_suffix = f" | {self.tr('status.mode.preview', 'Preview')}"
        elif doc and doc.get('large_file_mode'):
            mode_suffix = f" | {self.tr('status.mode.large', 'Large')}"

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

    def build_compare_status_text(self, doc, text_widget):
        row, col = text_widget.index(tk.INSERT).split('.')
        row = int(row)
        col = int(col) + 1

        if doc and doc.get('virtual_mode'):
            total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
            row = max(1, min(total_lines, doc.get('window_start_line', 1) + row - 1))
            total_chars = int(doc.get('file_size_bytes', 0) or 0)
            char_info = self.tr(
                'status.byte_count',
                '{total_bytes} {bytes_label}',
                total_bytes=f"{total_chars:,}",
                bytes_label=self.tr('status.bytes', 'bytes')
            )
        elif doc and doc.get('large_file_mode'):
            try:
                total_lines = max(1, int(text_widget.index('end-1c').split('.')[0]))
            except (tk.TclError, ValueError):
                total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
            doc['total_file_lines'] = total_lines
            total_chars = max(0, int(doc.get('file_size_bytes', 0) or 0))
            char_info = self.tr(
                'status.byte_count',
                '{total_bytes} {bytes_label}',
                total_bytes=f"{total_chars:,}",
                bytes_label=self.tr('status.bytes', 'bytes')
            )
        else:
            total_lines, total_chars, char_info = self.build_status_char_info_for_widget(text_widget)
        mode_suffix = ""
        if doc and doc.get('virtual_mode'):
            mode_suffix = f" | {self.tr('status.mode.virtual_editable', 'Paged') if self.is_virtual_editable(doc) else self.tr('status.mode.virtual', 'Virtual')}"
        elif doc and doc.get('preview_mode'):
            mode_suffix = f" | {self.tr('status.mode.preview', 'Preview')}"
        elif doc and doc.get('large_file_mode'):
            mode_suffix = f" | {self.tr('status.mode.large', 'Large')}"

        return self.tr(
            'status.compare',
            '{line_label} {row} {of_label} {total_lines}, {col_label} {col} | {char_info}{mode_suffix}',
            line_label=self.tr('status.line', 'Ln'),
            row=row,
            of_label=self.tr('status.of', 'of'),
            total_lines=total_lines,
            col_label=self.tr('status.col', 'Col'),
            col=col,
            char_info=char_info,
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
        status_tail_text = f"{shared_notes_tail}{self.tr('status.resource_usage', ' | CPU: {cpu_percent}% Memory: {memory_mb}MB', cpu_percent=f'{self.cpu_used_percent:.1f}', memory_mb=self.memory_used_mb)}"
        self.status.config(text=status_main_text)
        self.status_tail.config(text=status_tail_text)

        if hasattr(self, 'compare_status') and self.compare_view and self.compare_active:
            try:
                compare_widget = self.compare_view.get('text')
                compare_doc = self.get_doc_for_text_widget(compare_widget) if compare_widget is not None else None
                if compare_widget is None or compare_doc is None:
                    raise tk.TclError
                self.compare_status.config(
                    text=self.build_compare_status_text(compare_doc, compare_widget)
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
        self.update_breadcrumbs()

    def is_side_panel_visible(self):
        return bool(self.compare_active or self.markdown_preview_enabled.get())

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
            gutter.bind('<Button-1>', lambda e, frame=tab_id: self.handle_gutter_click(e, tab_id=frame))
            gutter.bind('<Motion>', lambda e, frame=tab_id: self.update_gutter_cursor(e, tab_id=frame))
        elif doc is not None:
            gutter.bind('<Button-1>', lambda e, target=doc: self.handle_gutter_click(e, target_doc=target))
            gutter.bind('<Motion>', lambda e, target=doc: self.update_gutter_cursor(e, target_doc=target))
        gutter.bind('<Leave>', self.clear_gutter_cursor)
        return gutter

    def get_gutter_font(self):
        gutter_size = max(9, self.current_font_size - 1)
        gutter_font = getattr(self, 'gutter_font', None)
        if gutter_font is not None:
            try:
                current_family = str(gutter_font.actual('family') or '')
                current_size = int(gutter_font.actual('size') or gutter_size)
                if current_family == self.font_family and current_size == gutter_size:
                    return gutter_font
                gutter_font.configure(
                    family=self.font_family,
                    size=gutter_size,
                    weight='normal',
                    slant='roman',
                    underline=0,
                    overstrike=0
                )
                return gutter_font
            except Exception:
                pass
        self.gutter_font = tkfont.Font(family=self.font_family, size=gutter_size)
        return self.gutter_font

    def get_display_line_number(self, doc, local_line):
        if doc.get('virtual_mode'):
            return doc.get('window_start_line', 1) + local_line - 1
        return local_line

    def get_gutter_doc(self, tab_id=None, target_doc=None):
        if target_doc is not None:
            return target_doc
        if tab_id is None:
            return None
        return self.documents.get(str(tab_id))

    def get_gutter_fold_hitbox(self, doc, x, y):
        if not doc:
            return None
        for hitbox in doc.get('gutter_fold_hitboxes') or []:
            if hitbox['x1'] <= x <= hitbox['x2'] and hitbox['y1'] <= y <= hitbox['y2']:
                return hitbox
        return None

    def handle_gutter_click(self, event, tab_id=None, target_doc=None):
        doc = self.get_gutter_doc(tab_id=tab_id, target_doc=target_doc)
        if not doc:
            return "break"
        hitbox = self.get_gutter_fold_hitbox(doc, event.x, event.y)
        if hitbox:
            text_widget = doc.get('text')
            start_line = int(hitbox.get('start', 1))
            if text_widget:
                try:
                    text_widget.mark_set(tk.INSERT, f'{start_line}.0')
                    text_widget.see(f'{start_line}.0')
                    text_widget.focus_set()
                except tk.TclError:
                    pass
            self.toggle_fold_region(doc, hitbox.get('region'))
            return "break"
        return self.copy_line_from_gutter(event, tab_id=tab_id, target_doc=doc)

    def update_gutter_cursor(self, event, tab_id=None, target_doc=None):
        doc = self.get_gutter_doc(tab_id=tab_id, target_doc=target_doc)
        cursor_name = 'hand2' if self.get_gutter_fold_hitbox(doc, event.x, event.y) else 'arrow'
        try:
            event.widget.configure(cursor=cursor_name)
        except tk.TclError:
            pass

    def clear_gutter_cursor(self, event=None):
        if event is None:
            return
        try:
            event.widget.configure(cursor='arrow')
        except tk.TclError:
            pass

    def update_line_number_gutter(self, doc):
        if not doc:
            return
        gutter = doc.get('line_numbers')
        text = doc.get('text')
        if not gutter or not text:
            return
        if not self.numbered_lines_enabled.get():
            return
        try:
            if not gutter.winfo_exists() or not text.winfo_exists() or not gutter.winfo_ismapped():
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

        try:
            gutter_font = self.get_gutter_font()
            desired_gutter_width = max(72, gutter_font.measure('9' * len(str(max_line_number))) + 48)
            current_gutter_width = int(gutter.cget('width'))
            if current_gutter_width != desired_gutter_width:
                gutter.configure(width=desired_gutter_width)

            surface = self.get_syntax_surface_palette()
            gutter.configure(bg=surface['gutter_bg'])
            gutter.delete('all')
            doc['gutter_fold_hitboxes'] = []
            gutter_height = max(gutter.winfo_height(), text.winfo_height(), 1)
            gutter_width = int(gutter.cget('width'))
            current_line = 1
            fold_regions = {
                int(region['start']): dict(region)
                for region in self.get_cached_fold_regions(doc)
            }
            collapsed_regions = self.get_collapsed_fold_regions(doc)
            collapsed_ranges = [
                (int(region.get('start', 1)), int(region.get('end', 1)))
                for region in collapsed_regions.values()
            ]
            try:
                current_line = int(text.index(tk.INSERT).split('.')[0])
            except tk.TclError:
                pass

            gutter.create_rectangle(0, 0, gutter_width - 1, gutter_height, fill=surface['gutter_bg'], outline='')
            gutter.create_rectangle(gutter_width - 1, 0, gutter_width, gutter_height, fill=surface['gutter_divider'], outline='')

            index = text.index('@0,0')
            while True:
                info = text.dlineinfo(index)
                if info is None:
                    break
                local_line = int(index.split('.')[0])
                if any(start_line < local_line <= end_line for start_line, end_line in collapsed_ranges):
                    index = text.index(f"{local_line + 1}.0")
                    continue
                display_line = self.get_display_line_number(doc, local_line)
                y = int(round(info[1]))
                line_height = int(round(info[3]))
                if local_line == current_line:
                    gutter.create_rectangle(0, y, gutter_width - 1, y + line_height, fill=surface['gutter_current_bg'], outline='')
                    line_fg = surface['gutter_current_fg']
                else:
                    line_fg = surface['gutter_fg']
                marker_region = collapsed_regions.get(local_line) or fold_regions.get(local_line)
                if marker_region:
                    marker_size = max(10, min(13, line_height - 4))
                    x1 = 6
                    y1 = int(round(y + max(1, (line_height - marker_size) / 2)))
                    x2 = x1 + marker_size
                    y2 = y1 + marker_size
                    gutter.create_rectangle(
                        x1,
                        y1,
                        x2,
                        y2,
                        fill=surface['text_bg'],
                        outline=surface['gutter_divider'],
                        width=1
                    )
                    mid_x = int(round((x1 + x2) / 2))
                    mid_y = int(round((y1 + y2) / 2))
                    horizontal_inset = max(3, marker_size // 4)
                    vertical_inset = max(3, marker_size // 4)
                    gutter.create_line(
                        x1 + horizontal_inset,
                        mid_y,
                        x2 - horizontal_inset,
                        mid_y,
                        fill=line_fg,
                        width=1
                    )
                    if local_line in collapsed_regions:
                        gutter.create_line(
                            mid_x,
                            y1 + vertical_inset,
                            mid_x,
                            y2 - vertical_inset,
                            fill=line_fg,
                            width=1
                        )
                    doc['gutter_fold_hitboxes'].append(
                        {
                            'x1': x1,
                            'y1': y1,
                            'x2': x2,
                            'y2': y2,
                            'start': int(marker_region.get('start', local_line)),
                            'region': dict(marker_region),
                        }
                    )
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
            self.tr('clipboard.line_copied', 'Copied line {line_number} to clipboard', line_number=display_line),
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
        tk.Label(self.find_frame, text=self.tr('find.panel.find', 'Find:'), bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(8,4), pady=6)
        self.find_entry = tk.Entry(self.find_frame, width=40)
        self.find_entry.pack(side='left', padx=4, pady=6)
        tk.Button(self.find_frame, text=self.tr('menu.edit.find', 'Find'), command=self.find_next)\
            .pack(side='left', padx=4, pady=6)
        self.find_results_label = tk.Label(
            self.find_frame,
            text="",
            bg=self.panel_bg,
            fg='#9aa0a6'
        )
        self.find_results_label.pack(side='left', padx=(4, 8), pady=6)
        tk.Label(self.find_frame, text='|', bg=self.panel_bg, fg='#9aa0a6')\
            .pack(side='left', padx=(4, 4), pady=6)
        tk.Label(self.find_frame, text=self.tr('find.panel.find_in_label', 'Find In:'), bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(0, 4), pady=6)
        self.find_in_query_var = tk.StringVar()
        self.find_in_entry = tk.Entry(self.find_frame, width=42, textvariable=self.find_in_query_var)
        self.find_in_entry.pack(side='left', padx=4, pady=6)
        tk.Button(self.find_frame, text=self.tr('find.panel.browse', 'Browse'), command=self.choose_find_in_directory_and_search)\
            .pack(side='left', padx=4, pady=6)
        self.find_entry.bind('<KeyPress>', self.handle_search_history_keypress)
        self.find_entry.bind('<Return>', self.find_from_input)
        self.find_entry.bind('<KeyRelease>', self.on_find_entry_change)
        self.find_entry.bind('<Escape>', lambda e: self.show_find_panel())   # ← added
        self.find_entry.bind('<FocusIn>', self.on_search_history_entry_focus)
        self.find_entry.bind('<FocusOut>', self.on_search_history_entry_focus_out)
        self.find_entry.bind('<ButtonRelease-1>', self.on_search_history_entry_focus)
        self.find_in_entry.bind('<KeyPress>', self.handle_search_history_keypress)
        self.find_in_entry.bind('<Return>', self.find_in_directory_from_entry)
        self.find_in_entry.bind('<KeyRelease>', self.on_find_in_entry_change)
        self.find_in_entry.bind('<Escape>', lambda e: self.show_find_panel())
        self.find_in_entry.bind('<FocusIn>', self.on_search_history_entry_focus)
        self.find_in_entry.bind('<FocusOut>', self.on_search_history_entry_focus_out)
        self.find_in_entry.bind('<ButtonRelease-1>', self.on_search_history_entry_focus)

        # Replace panel
        self.replace_frame = tk.Frame(self.bottom_frame, bg=self.panel_bg)
        tk.Label(self.replace_frame, text=self.tr('find.panel.find', 'Find:'), bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(8,4), pady=6)
        self.replace_find_entry = tk.Entry(self.replace_frame, width=30)
        self.replace_find_entry.pack(side='left', padx=4, pady=6)
        tk.Label(self.replace_frame, text=self.tr('replace.panel.replace_with', 'Replace with:'), bg=self.panel_bg, fg=self.fg_color)\
            .pack(side='left', padx=(12,4), pady=6)
        self.replace_entry = tk.Entry(self.replace_frame, width=30)
        self.replace_entry.pack(side='left', padx=4, pady=6)
        tk.Button(self.replace_frame, text=self.tr('replace.panel.replace_all', 'Replace All'), command=self.replace_all)\
            .pack(side='left', padx=8, pady=6)
        self.replace_results_label = tk.Label(
            self.replace_frame,
            text="",
            bg=self.panel_bg,
            fg='#9aa0a6'
        )
        self.replace_results_label.pack(side='left', padx=(4, 8), pady=6)

        self.replace_find_entry.bind('<KeyPress>', self.handle_search_history_keypress)
        self.replace_find_entry.bind('<Return>', self.find_from_input)     # ← added
        self.replace_find_entry.bind('<KeyRelease>', self.on_find_entry_change)
        self.replace_find_entry.bind('<Escape>', lambda e: self.show_replace_panel())  # ← added
        self.replace_find_entry.bind('<FocusIn>', self.on_search_history_entry_focus)
        self.replace_find_entry.bind('<FocusOut>', self.on_search_history_entry_focus_out)
        self.replace_find_entry.bind('<ButtonRelease-1>', self.on_search_history_entry_focus)
        self.replace_entry.bind('<Escape>', lambda e: self.show_replace_panel())       # ← added

        # Command panel
        self.command_frame = tk.Frame(self.bottom_frame, bg=self.panel_bg)
        self.command_frame.grid_columnconfigure(0, weight=1)
        self.command_frame.grid_rowconfigure(2, weight=1)
        self.command_resize_grip = tk.Frame(
            self.command_frame,
            bg='#30363d',
            height=6,
            cursor='sb_v_double_arrow'
        )
        self.command_resize_grip.grid(row=0, column=0, sticky='ew')
        self.command_resize_grip.grid_propagate(False)
        self.command_resize_grip.bind('<ButtonPress-1>', self.start_command_panel_resize)
        self.command_resize_grip.bind('<B1-Motion>', self.drag_command_panel_resize)
        self.command_resize_grip.bind('<ButtonRelease-1>', self.finish_command_panel_resize)
        self.command_resize_grip.bind('<Double-Button-1>', self.reset_command_panel_height)

        self.command_controls_frame = tk.Frame(self.command_frame, bg=self.panel_bg)
        self.command_controls_frame.grid(row=1, column=0, sticky='ew')
        self.command_controls_frame.grid_columnconfigure(1, weight=1)
        tk.Label(
            self.command_controls_frame,
            text='Cmd:',
            bg=self.panel_bg,
            fg=self.fg_color
        ).grid(row=0, column=0, sticky='w', padx=(8, 4), pady=6)
        self.command_entry = tk.Entry(self.command_controls_frame)
        self.command_entry.grid(row=0, column=1, sticky='ew', padx=4, pady=6)
        tk.Button(
            self.command_controls_frame,
            text='Run',
            command=self.run_command_panel
        ).grid(row=0, column=2, padx=4, pady=6)
        tk.Button(
            self.command_controls_frame,
            text='Clear',
            command=self.clear_command_output
        ).grid(row=0, column=3, padx=(4, 8), pady=6)
        self.command_body_frame = tk.Frame(self.command_frame, bg=self.panel_bg)
        self.command_body_frame.grid(row=2, column=0, sticky='nsew')
        self.command_body_frame.grid_columnconfigure(0, weight=1)
        self.command_body_frame.grid_rowconfigure(0, weight=1)
        self.command_output = tk.Text(
            self.command_body_frame,
            height=self.command_panel_height,
            wrap='word',
            bg=self.text_bg,
            fg=self.text_fg,
            insertbackground=self.cursor_color,
            font=('Consolas', 10),
            padx=8,
            pady=6,
            borderwidth=0,
            highlightthickness=0,
            relief='flat'
        )
        self.command_output.grid(row=0, column=0, sticky='nsew', padx=(8, 0), pady=(0, 8))
        self.command_output.configure(state='disabled')
        self.command_output_scroll = ttk.Scrollbar(self.command_body_frame, orient='vertical', command=self.command_output.yview)
        self.command_output_scroll.grid(row=0, column=1, sticky='ns', padx=(0, 8), pady=(0, 8))
        self.command_output.configure(yscrollcommand=self.command_output_scroll.set)
        self.command_entry.bind('<Return>', lambda e: self.run_command_panel())
        self.command_entry.bind('<Escape>', self.handle_command_entry_escape)
        self.command_entry.bind('<Up>', self.handle_command_suggestion_keypress)
        self.command_entry.bind('<Down>', self.handle_command_suggestion_keypress)
        self.command_entry.bind('<Tab>', self.handle_command_suggestion_keypress)
        self.command_entry.bind('<KeyRelease>', self.on_command_entry_key_release)
        self.command_entry.bind('<FocusIn>', self.on_command_entry_focus, add='+')
        self.command_entry.bind('<FocusOut>', self.on_command_entry_focus_out, add='+')

        # Hide both initially
        self.find_frame.grid_remove()
        self.replace_frame.grid_remove()
        self.command_frame.grid_remove()
        self.bottom_frame.grid_remove()

    def update_bottom_panel_visibility(self):
        if self.find_panel_visible or self.replace_panel_visible or self.command_panel_visible:
            self.bottom_frame.grid()
        else:
            self.bottom_frame.grid_remove()

    def get_doc_display_path(self, doc):
        if not doc:
            return None
        if doc.get('is_remote') and doc.get('remote_spec'):
            return doc.get('remote_spec')
        if doc.get('display_name'):
            return doc.get('display_name')
        return doc.get('file_path')

    def get_doc_copyable_path(self, doc):
        if not doc:
            return None
        if doc.get('is_remote') and doc.get('remote_spec'):
            return doc.get('remote_spec')
        file_path = doc.get('file_path')
        if isinstance(file_path, str) and file_path.strip():
            return file_path
        display_name = doc.get('display_name')
        if isinstance(display_name, str) and display_name.strip():
            return display_name
        return None

    def get_doc_copyable_name(self, doc):
        if not doc:
            return None
        frame = doc.get('frame')
        if frame is not None and str(frame) in self.documents:
            return self.get_doc_name(frame)
        display_path = self.get_doc_display_path(doc)
        if isinstance(display_path, str) and display_path.strip():
            normalized_path = str(display_path).rstrip('/\\')
            return os.path.basename(normalized_path) or display_path
        untitled_name = doc.get('untitled_name')
        if isinstance(untitled_name, str) and untitled_name.strip():
            return untitled_name
        return None

    def get_breadcrumb_segments(self, doc, text_widget=None):
        if not doc:
            return []
        segments = []
        copy_name = self.get_doc_copyable_name(doc)
        if copy_name:
            segments.append({'kind': 'name', 'text': copy_name, 'copy': copy_name})

        display_path = self.get_doc_display_path(doc)
        copy_path = self.get_doc_copyable_path(doc)
        if display_path and display_path != copy_name:
            segments.append({'kind': 'path', 'text': display_path, 'copy': copy_path or display_path})

        text_widget = text_widget or doc.get('text')
        if text_widget:
            try:
                current_line = int(text_widget.index(tk.INSERT).split('.')[0])
            except (tk.TclError, ValueError):
                current_line = 1
            symbol = self.get_symbol_at_line(doc, current_line)
            if symbol:
                symbol_name = str(symbol.get('name') or '').strip()
                if symbol_name:
                    segments.append({'kind': 'symbol', 'text': symbol_name, 'copy': None})
        return segments

    def get_doc_working_directory(self, doc=None):
        doc = doc or self.get_current_doc()
        if not doc:
            return self.app_dir
        candidate_path = doc.get('remote_shadow_path') or doc.get('file_path')
        if candidate_path:
            try:
                return os.path.dirname(os.path.abspath(candidate_path)) or self.app_dir
            except OSError:
                pass
        return self.app_dir

    def show_command_panel(self, event=None):
        self.hide_search_history_popup()
        self.hide_autocomplete_popup()
        self.hide_command_suggestion_popup()
        if self.find_panel_visible:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.clear_find_highlights()
        if self.replace_panel_visible:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.clear_find_highlights()

        if not self.command_panel_visible:
            self.bottom_frame.grid()
            self.command_frame.grid(sticky='nsew')
            self.command_panel_visible = True
            self.set_command_panel_height(self.command_panel_height, persist=False)
            self.refresh_command_history_list()
            self.command_entry.focus_set()
            try:
                self.root.after_idle(lambda: self.update_command_suggestion_popup(self.command_entry, force_show=True))
            except tk.TclError:
                pass
        else:
            self.hide_command_suggestion_popup()
            self.command_frame.grid_remove()
            self.command_panel_visible = False
            self.focus_last_active_editor()
        self.update_bottom_panel_visibility()
        return "break"

    def on_command_entry_key_release(self, event=None):
        keysym = str(getattr(event, 'keysym', '') or '')
        if keysym in {'Up', 'Down', 'Return', 'Escape', 'Tab', 'ISO_Left_Tab'}:
            return None
        self.command_history_index = None
        self.update_command_suggestion_popup(getattr(event, 'widget', None) or getattr(self, 'command_entry', None))
        return None

    def on_command_history_select(self, event=None):
        listbox = getattr(self, 'command_history_listbox', None)
        if not listbox:
            return None
        selection = listbox.curselection()
        if not selection:
            return None
        selected_index = max(0, int(selection[0]))
        if selected_index >= len(self.command_history):
            return None
        self.command_history_index = selected_index
        value = self.command_history[selected_index]
        self.command_entry.delete(0, tk.END)
        self.command_entry.insert(0, value)
        self.command_entry.icursor(tk.END)
        self.command_entry.focus_set()
        return None

    def select_command_history_value(self, value):
        listbox = getattr(self, 'command_history_listbox', None)
        if not listbox:
            return False
        try:
            if not listbox.winfo_exists():
                return False
        except tk.TclError:
            return False
        listbox.selection_clear(0, tk.END)
        if not self.command_history:
            return False
        normalized = str(value or '').strip().lower()
        if not normalized:
            return False
        selected_index = None
        for index, item in enumerate(self.command_history):
            lowered_item = str(item).lower()
            if lowered_item == normalized:
                selected_index = index
                break
            if selected_index is None and lowered_item.startswith(normalized):
                selected_index = index
        if selected_index is None:
            return False
        listbox.selection_set(selected_index)
        listbox.activate(selected_index)
        listbox.see(selected_index)
        return True

    def refresh_command_history_list(self, selected_value=None):
        listbox = getattr(self, 'command_history_listbox', None)
        if not listbox:
            return
        try:
            if not listbox.winfo_exists():
                return
        except tk.TclError:
            return
        selection = listbox.curselection()
        if selected_value is None and selection:
            selected_index = int(selection[0])
            if 0 <= selected_index < len(self.command_history):
                selected_value = self.command_history[selected_index]
        listbox.delete(0, tk.END)
        for item in self.command_history:
            listbox.insert(tk.END, item)
        self.select_command_history_value(selected_value)

    def record_command_history(self, command_text):
        cleaned_command = str(command_text or '').strip()
        if not cleaned_command:
            return False
        updated_history = [
            item for item in self.command_history
            if str(item).strip().lower() != cleaned_command.lower()
        ]
        updated_history.insert(0, cleaned_command)
        updated_history = updated_history[:self.max_command_history]
        changed = updated_history != self.command_history
        self.command_history = updated_history
        self.command_history_index = None
        self.refresh_command_history_list(selected_value=cleaned_command)
        if changed:
            self.save_session()
        return changed

    def get_command_panel_line_height(self):
        if not hasattr(self, 'command_output') or not self.command_output:
            return 18
        try:
            font_name = self.command_output.cget('font')
            return max(1, tkfont.Font(font=font_name).metrics('linespace'))
        except (tk.TclError, RuntimeError):
            return 18

    def set_command_panel_height(self, height_lines, persist=False):
        try:
            safe_height = int(round(float(height_lines)))
        except (TypeError, ValueError):
            safe_height = self.command_panel_default_height
        safe_height = max(self.command_panel_min_height, min(self.command_panel_max_height, safe_height))
        self.command_panel_height = safe_height
        widgets = [getattr(self, 'command_output', None)]
        for widget in widgets:
            if not widget:
                continue
            try:
                widget.configure(height=safe_height)
            except tk.TclError:
                continue
        if persist:
            self.save_session()
        return safe_height

    def start_command_panel_resize(self, event=None):
        if event is None:
            return "break"
        self.command_panel_resize_active = True
        self.command_panel_resize_origin_y = int(getattr(event, 'y_root', 0))
        self.command_panel_resize_origin_height = int(self.command_panel_height)
        return "break"

    def drag_command_panel_resize(self, event=None):
        if not self.command_panel_resize_active or event is None:
            return "break"
        delta_pixels = self.command_panel_resize_origin_y - int(getattr(event, 'y_root', self.command_panel_resize_origin_y))
        delta_lines = int(round(delta_pixels / max(1, self.get_command_panel_line_height())))
        self.set_command_panel_height(self.command_panel_resize_origin_height + delta_lines, persist=False)
        return "break"

    def finish_command_panel_resize(self, event=None):
        if self.command_panel_resize_active:
            self.command_panel_resize_active = False
            self.save_session()
        return "break"

    def reset_command_panel_height(self, event=None):
        self.command_panel_resize_active = False
        self.set_command_panel_height(self.command_panel_default_height, persist=True)
        return "break"

    def handle_command_history_keypress(self, event=None):
        if not self.command_history:
            return "break"
        keysym = str(getattr(event, 'keysym', '') or '')
        if keysym == 'Up':
            if self.command_history_index is None:
                self.command_history_index = 0
            else:
                self.command_history_index = min(len(self.command_history) - 1, self.command_history_index + 1)
        else:
            if self.command_history_index is None:
                return "break"
            if self.command_history_index <= 0:
                self.command_history_index = None
            else:
                self.command_history_index -= 1
        if self.command_history_index is None:
            value = ''
        else:
            value = self.command_history[self.command_history_index]
        self.command_entry.delete(0, tk.END)
        self.command_entry.insert(0, value)
        self.command_entry.icursor(tk.END)
        self.select_command_history_value(value)
        return "break"

    def append_command_output(self, text):
        if not hasattr(self, 'command_output') or not self.command_output:
            return
        output_widget = self.command_output
        try:
            output_widget.configure(state='normal')
            output_widget.insert(tk.END, str(text))
            current_text = output_widget.get('1.0', 'end-1c')
            if len(current_text) > self.command_output_max_chars:
                trim_chars = len(current_text) - self.command_output_max_chars
                output_widget.delete('1.0', f'1.0+{trim_chars}c')
            output_widget.see(tk.END)
            output_widget.configure(state='disabled')
        except tk.TclError:
            return

    def clear_command_output(self):
        if not hasattr(self, 'command_output') or not self.command_output:
            return "break"
        try:
            self.command_output.configure(state='normal')
            self.command_output.delete('1.0', tk.END)
            self.command_output.configure(state='disabled')
        except tk.TclError:
            pass
        return "break"

    def set_feature_toggle_from_text(self, variable, value_text):
        normalized = str(value_text or '').strip().lower()
        if normalized in {'on', 'true', '1', 'yes'}:
            variable.set(True)
            return True
        if normalized in {'off', 'false', '0', 'no'}:
            variable.set(False)
            return True
        return False

    def strip_named_command_argument(self, value_text):
        value = str(value_text or '').strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            return value[1:-1]
        return value

    def normalize_named_command_value(self, value_text):
        return re.sub(r'[^a-z0-9]+', '', str(value_text or '').strip().lower())

    def format_named_toggle_state(self, label, enabled):
        return f"{label} {'enabled' if enabled else 'disabled'}.\n"

    def get_named_command_catalog(self):
        return [
            (
                'File',
                [
                    ':open [path]',
                    ':open-project [path]',
                    ':remote user@host:/path/file',
                    ':grab-git',
                    ':recent [list|N|clear]',
                    ':new-tab',
                    ':close-tab',
                    ':save',
                    ':save-all',
                    ':save-as',
                    ':save-copy',
                    ':save-run',
                    ':save-encrypted',
                    ':print',
                    ':export-notes',
                    ':exit',
                ],
            ),
            (
                'Edit',
                [
                    ':undo',
                    ':redo',
                    ':cut',
                    ':copy',
                    ':paste',
                    ':select-all',
                    ':find',
                    ':find-next',
                    ':find-prev',
                    ':replace',
                    ':command-panel',
                    ':symbols',
                    ':project-symbols',
                    ':fold',
                    ':fold-all',
                    ':unfold-all',
                    ':lint',
                    ':date',
                    ':time-date',
                    ':font',
                    ':language [list|code]',
                ],
            ),
            (
                'View',
                [
                    ':fullscreen',
                    ':switch-tab',
                    ':currently-editing',
                    ':cycle-notes',
                    ':note-filter [list|all|unread|yellow|green|red|blue]',
                    ':goto-line',
                    ':top',
                    ':bottom',
                    ':theme [list|create|name]',
                    ':syntax-mode [list|auto|plain|python|c|cpp|rust|java|javascript|html|php|xml|sql]',
                    ':compare',
                    ':close-compare',
                    ':preview [on|off]',
                ],
            ),
            (
                'Settings',
                [
                    ':edit-with-notepad-x on|off',
                    ':sound on|off',
                    ':status-bar on|off',
                    ':numbered-lines on|off',
                    ':autocomplete on|off',
                    ':spell-check on|off',
                    ':auto-pair on|off',
                    ':compare-multi-edit on|off',
                    ':minimap on|off',
                    ':breadcrumbs on|off',
                    ':diagnostics on|off',
                    ':autosave on|off',
                    ':word-wrap on|off',
                    ':sync-page-navigation on|off',
                ],
            ),
            (
                'Help',
                [
                    ':help',
                    ':help-contents',
                    ':about',
                ],
            ),
        ]

    def get_named_command_templates(self):
        templates = []
        seen = set()
        for _, commands in self.get_named_command_catalog():
            for command in commands:
                if command in seen:
                    continue
                seen.add(command)
                templates.append(command)
        return templates

    def cancel_command_suggestion_hide_job(self):
        if self.command_suggestion_hide_job:
            try:
                self.root.after_cancel(self.command_suggestion_hide_job)
            except tk.TclError:
                pass
            self.command_suggestion_hide_job = None

    def hide_command_suggestion_popup(self):
        self.cancel_command_suggestion_hide_job()
        popup = getattr(self, 'command_suggestion_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.command_suggestion_popup = None
        self.command_suggestion_listbox = None
        self.command_suggestion_items = []

    def command_suggestion_popup_visible(self):
        popup = getattr(self, 'command_suggestion_popup', None)
        if popup is None:
            return False
        try:
            return popup.winfo_exists()
        except tk.TclError:
            return False

    def maybe_hide_command_suggestion_popup(self):
        self.command_suggestion_hide_job = None
        focus_widget = self.safe_focus_get()
        listbox = getattr(self, 'command_suggestion_listbox', None)
        if focus_widget is not None and focus_widget in (getattr(self, 'command_entry', None), listbox):
            return
        self.hide_command_suggestion_popup()

    def schedule_command_suggestion_popup_hide(self):
        self.cancel_command_suggestion_hide_job()
        try:
            self.command_suggestion_hide_job = self.root.after(120, self.maybe_hide_command_suggestion_popup)
        except tk.TclError:
            self.command_suggestion_hide_job = None

    def get_command_suggestion_matches(self, query_text):
        normalized_query = str(query_text or '').strip().lower()
        suggestions = self.get_named_command_templates()
        if not normalized_query:
            return suggestions

        prefix_matches = []
        substring_matches = []
        for item in suggestions:
            lowered = item.lower()
            if lowered.startswith(normalized_query):
                prefix_matches.append(item)
            elif normalized_query in lowered:
                substring_matches.append(item)
        return prefix_matches + substring_matches

    def select_command_suggestion_index(self, index):
        if not self.command_suggestion_popup_visible() or not self.command_suggestion_listbox:
            return False
        listbox = self.command_suggestion_listbox
        if listbox.size() <= 0:
            return False
        safe_index = max(0, min(listbox.size() - 1, index))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(safe_index)
        listbox.activate(safe_index)
        listbox.see(safe_index)
        return True

    def update_command_suggestion_popup(self, entry_widget=None, force_show=False):
        entry_widget = entry_widget or getattr(self, 'command_entry', None)
        if not entry_widget:
            self.hide_command_suggestion_popup()
            return False
        try:
            if not entry_widget.winfo_exists() or not entry_widget.winfo_ismapped():
                self.hide_command_suggestion_popup()
                return False
            current_text = entry_widget.get()
        except tk.TclError:
            self.hide_command_suggestion_popup()
            return False

        current_text = str(current_text or '')
        if not current_text.startswith(':'):
            self.hide_command_suggestion_popup()
            return False

        query_text = current_text.strip()
        suggestions = self.get_command_suggestion_matches(query_text)
        if not suggestions:
            self.hide_command_suggestion_popup()
            return False

        selected_value = None
        if self.command_suggestion_popup_visible() and self.command_suggestion_listbox:
            selection = self.command_suggestion_listbox.curselection()
            if selection:
                selected_value = self.command_suggestion_listbox.get(selection[0])

        popup = self.command_suggestion_popup
        if not self.command_suggestion_popup_visible():
            popup = self.create_popup_toplevel(self.root)
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
                exportselection=False,
                takefocus=False,
                selectmode='browse',
                width=max(22, min(72, max(len(item) for item in suggestions) + 2)),
                height=min(10, len(suggestions))
            )
            listbox.pack(fill='both', expand=True)
            listbox.bind('<Motion>', self.on_command_suggestion_listbox_motion)
            listbox.bind('<ButtonRelease-1>', self.on_command_suggestion_listbox_click)
            self.command_suggestion_popup = popup
            self.command_suggestion_listbox = listbox
        else:
            listbox = self.command_suggestion_listbox
            listbox.configure(
                width=max(22, min(72, max(len(item) for item in suggestions) + 2)),
                height=min(10, len(suggestions))
            )

        listbox.delete(0, tk.END)
        for suggestion in suggestions:
            listbox.insert(tk.END, suggestion)

        selected_index = 0
        if selected_value in suggestions:
            selected_index = suggestions.index(selected_value)
        self.select_command_suggestion_index(selected_index)

        self.show_popup_toplevel(
            popup,
            entry_widget.winfo_rootx(),
            entry_widget.winfo_rooty() + entry_widget.winfo_height() + 2
        )
        self.command_suggestion_items = suggestions
        return True

    def move_command_suggestion_selection(self, direction):
        if not self.command_suggestion_popup_visible() or not self.command_suggestion_listbox:
            return None
        listbox = self.command_suggestion_listbox
        selection = listbox.curselection()
        current_index = selection[0] if selection else 0
        next_index = max(0, min(listbox.size() - 1, current_index + direction))
        self.select_command_suggestion_index(next_index)
        return "break"

    def accept_command_suggestion_selection(self):
        if not self.command_suggestion_popup_visible() or not self.command_suggestion_listbox:
            return False
        selection = self.command_suggestion_listbox.curselection()
        if selection:
            suggestion = self.command_suggestion_listbox.get(selection[0])
        elif self.command_suggestion_listbox.size() > 0:
            suggestion = self.command_suggestion_listbox.get(0)
        else:
            return False

        entry_widget = getattr(self, 'command_entry', None)
        if not entry_widget:
            self.hide_command_suggestion_popup()
            return False
        try:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, suggestion)
            entry_widget.icursor(tk.END)
            entry_widget.focus_set()
        except tk.TclError:
            self.hide_command_suggestion_popup()
            return False

        self.hide_command_suggestion_popup()
        return True

    def on_command_suggestion_listbox_motion(self, event=None):
        listbox = getattr(self, 'command_suggestion_listbox', None)
        if not listbox:
            return None
        index = listbox.nearest(getattr(event, 'y', 0))
        self.select_command_suggestion_index(index)
        return None

    def on_command_suggestion_listbox_click(self, event=None):
        listbox = getattr(self, 'command_suggestion_listbox', None)
        if not listbox:
            return "break"
        index = listbox.nearest(getattr(event, 'y', 0))
        self.select_command_suggestion_index(index)
        self.accept_command_suggestion_selection()
        return "break"

    def on_command_entry_focus(self, event=None):
        entry_widget = getattr(event, 'widget', None) or getattr(self, 'command_entry', None)
        if not entry_widget:
            return None
        self.cancel_command_suggestion_hide_job()
        try:
            self.root.after_idle(lambda widget=entry_widget: self.update_command_suggestion_popup(widget, force_show=True))
        except tk.TclError:
            return None
        return None

    def on_command_entry_focus_out(self, event=None):
        self.schedule_command_suggestion_popup_hide()
        return None

    def handle_command_suggestion_keypress(self, event=None):
        entry_widget = getattr(event, 'widget', None) or getattr(self, 'command_entry', None)
        if not entry_widget:
            return None

        keysym = str(getattr(event, 'keysym', '') or '')
        popup_visible = self.command_suggestion_popup_visible()

        if keysym in {'Up', 'Down'}:
            if popup_visible:
                return self.move_command_suggestion_selection(-1 if keysym == 'Up' else 1)
            if self.update_command_suggestion_popup(entry_widget, force_show=True):
                if keysym == 'Up' and self.command_suggestion_listbox:
                    self.select_command_suggestion_index(self.command_suggestion_listbox.size() - 1)
                return "break"
            return None

        if keysym in {'Tab', 'ISO_Left_Tab'} and popup_visible:
            if self.accept_command_suggestion_selection():
                return "break"
            return None

        if popup_visible and keysym in {'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}:
            self.hide_command_suggestion_popup()
        return None

    def handle_command_entry_escape(self, event=None):
        if self.command_suggestion_popup_visible():
            self.hide_command_suggestion_popup()
            return "break"
        return self.show_command_panel()

    def run_named_toggle_command(self, command_name, argument_text, variable, callback, label):
        if not argument_text:
            return self.format_named_toggle_state(label, bool(variable.get()))
        if not self.set_feature_toggle_from_text(variable, argument_text):
            return f"Usage: :{command_name} on|off\n"
        if callable(callback):
            callback()
        else:
            self.save_session()
        return self.format_named_toggle_state(label, bool(variable.get()))

    def get_named_syntax_mode_choices(self):
        return [
            ('Auto', 'auto'),
            ('Plain Text', 'plain'),
            ('Python', 'python'),
            ('C', 'c'),
            ('C++', 'cpp'),
            ('Rust', 'rust'),
            ('Java', 'java'),
            ('JavaScript', 'javascript'),
            ('HTML', 'html'),
            ('PHP', 'php'),
            ('XML', 'xml'),
            ('SQL', 'sql'),
        ]

    def get_named_note_filter_choices(self):
        return [
            (self.tr('note.filter.all', 'All'), 'all'),
            (self.tr('note.filter.unread', 'Unread'), 'unread'),
            (self.tr('note.filter.yellow', 'Yellow'), 'yellow'),
            (self.tr('note.filter.green', 'Green'), 'green'),
            (self.tr('note.filter.red', 'Red'), 'red'),
            (self.tr('note.filter.blue', 'Light Blue'), 'blue'),
        ]

    def get_named_command_help_text(self):
        lines = ["Built-in commands:"]
        for heading, commands in self.get_named_command_catalog():
            lines.append(f"{heading}:")
            lines.extend(commands)
        return "\n".join(lines) + "\n"

    def run_named_recent_command(self, argument_text):
        recent_files = [path for path in self.recent_files if isinstance(path, str) and path]
        argument = self.strip_named_command_argument(argument_text)
        normalized = self.normalize_named_command_value(argument)
        if not argument or normalized in {'list', 'ls'}:
            if not recent_files:
                return "Recent files list is empty.\n"
            lines = ["Recent files:"]
            for index, path in enumerate(recent_files, start=1):
                lines.append(f"{index}. {os.path.basename(path)} - {path}")
            return "\n".join(lines) + "\n"
        if normalized in {'clear', 'clearlist'}:
            self.clear_recent_files()
            return "Cleared recent files list.\n"
        target_path = None
        if argument.isdigit():
            selected_index = int(argument) - 1
            if 0 <= selected_index < len(recent_files):
                target_path = recent_files[selected_index]
        else:
            normalized_path = os.path.normcase(os.path.abspath(os.path.expanduser(argument)))
            for path in recent_files:
                if os.path.normcase(os.path.abspath(path)) == normalized_path:
                    target_path = path
                    break
        if not target_path:
            return "Usage: :recent [list|N|clear]\n"
        if self.open_file_path(target_path):
            return f"Opened recent file: {target_path}\n"
        return f"Could not open recent file: {target_path}\n"

    def run_named_language_command(self, argument_text):
        argument = self.strip_named_command_argument(argument_text)
        normalized = self.normalize_named_command_value(argument)
        language_codes = self.get_available_language_codes()
        if not argument or normalized in {'list', 'ls'}:
            current_code = self.locale_code
            lines = [f"Current language: {self.get_language_display_name(current_code)} ({current_code})", "Available languages:"]
            for code in language_codes:
                lines.append(f"- {self.get_language_display_name(code)} ({code})")
            return "\n".join(lines) + "\n"
        for code in language_codes:
            display_name = self.get_language_display_name(code)
            if normalized in {
                self.normalize_named_command_value(code),
                self.normalize_named_command_value(display_name),
            }:
                self.apply_locale(code)
                return f"Language set to {display_name} ({code}).\n"
        return "Usage: :language [list|code]\n"

    def run_named_theme_command(self, argument_text):
        argument = self.strip_named_command_argument(argument_text)
        normalized = self.normalize_named_command_value(argument)
        theme_names = self.get_available_syntax_theme_names()
        if not argument or normalized in {'list', 'ls'}:
            current_theme = self.syntax_theme.get()
            lines = [f"Current syntax theme: {self.get_syntax_theme_label(current_theme)}", "Available syntax themes:"]
            for theme_name in theme_names:
                lines.append(f"- {self.get_syntax_theme_label(theme_name)} ({theme_name})")
            return "\n".join(lines) + "\n"
        if normalized in {'create', 'new', 'newtheme'}:
            self.show_create_theme_dialog()
            return "Opened syntax theme creator.\n"
        for theme_name in theme_names:
            if normalized in {
                self.normalize_named_command_value(theme_name),
                self.normalize_named_command_value(self.get_syntax_theme_label(theme_name)),
            }:
                self.set_syntax_theme(theme_name)
                return f"Syntax theme set to {self.get_syntax_theme_label(theme_name)}.\n"
        return "Usage: :theme [list|create|name]\n"

    def run_named_syntax_mode_command(self, argument_text):
        argument = self.strip_named_command_argument(argument_text)
        normalized = self.normalize_named_command_value(argument)
        choices = self.get_named_syntax_mode_choices()
        current_mode = self.syntax_mode_selection.get() or 'auto'
        if not argument or normalized in {'list', 'ls'}:
            current_label = next((label for label, value in choices if value == current_mode), current_mode)
            lines = [f"Current syntax mode: {current_label} ({current_mode})", "Available syntax modes:"]
            for label, value in choices:
                lines.append(f"- {label} ({value})")
            return "\n".join(lines) + "\n"
        for label, value in choices:
            if normalized in {
                self.normalize_named_command_value(value),
                self.normalize_named_command_value(label),
            }:
                self.set_current_syntax_override(value)
                return f"Syntax mode set to {label}.\n"
        return "Usage: :syntax-mode [list|mode]\n"

    def run_named_note_filter_command(self, argument_text):
        argument = self.strip_named_command_argument(argument_text)
        normalized = self.normalize_named_command_value(argument)
        choices = self.get_named_note_filter_choices()
        current_filter = self.note_filter.get()
        if not argument or normalized in {'list', 'ls'}:
            current_label = next((label for label, value in choices if value == current_filter), current_filter)
            lines = [f"Current note filter: {current_label} ({current_filter})", "Available note filters:"]
            for label, value in choices:
                lines.append(f"- {label} ({value})")
            return "\n".join(lines) + "\n"
        for label, value in choices:
            if normalized in {
                self.normalize_named_command_value(value),
                self.normalize_named_command_value(label),
            }:
                self.note_filter.set(value)
                for doc in self.documents.values():
                    doc['last_note_cycle_tag'] = None
                self.update_status()
                return f"Note filter set to {label}.\n"
        return "Usage: :note-filter [list|all|unread|yellow|green|red|blue]\n"

    def run_named_command(self, command_text):
        normalized = str(command_text or '').strip()
        command_name, _, argument_text = normalized.partition(' ')
        command_name = command_name.lstrip(':').strip().lower()
        argument_text = argument_text.strip()

        if command_name in {'help', '?'}:
            return self.get_named_command_help_text()
        if command_name in {'open', 'file-open'}:
            argument = self.strip_named_command_argument(argument_text)
            if not argument:
                self.open_file()
                return "Opened file picker.\n"
            if self.open_file_path(os.path.expanduser(argument)):
                return f"Opened file: {os.path.abspath(os.path.expanduser(argument))}\n"
            return f"Could not open file: {argument}\n"
        if command_name in {'open-project', 'project-open'}:
            argument = self.strip_named_command_argument(argument_text)
            if not argument:
                self.open_project()
                return "Opened project picker.\n"
            if self.open_project_path(os.path.expanduser(argument)):
                return f"Opened project: {os.path.abspath(os.path.expanduser(argument))}\n"
            return f"Could not open project: {argument}\n"
        if command_name in {'open-remote', 'remote'}:
            argument = self.strip_named_command_argument(argument_text)
            if not argument:
                self.open_remote_file_dialog()
                return "Opened remote file prompt.\n"
            if self.open_remote_file(argument):
                return f"Opened remote file: {argument}\n"
            return f"Remote open failed: {argument}\n"
        if command_name in {'grab-git', 'git'}:
            self.grab_git_project()
            return "Opened Grab Git dialog.\n"
        if command_name in {'recent', 'recent-files'}:
            return self.run_named_recent_command(argument_text)
        if command_name in {'new-tab', 'new'}:
            self.new_tab()
            return "Opened a new tab.\n"
        if command_name in {'close-tab', 'close'}:
            self.close_current_tab()
            return "Closed the current tab.\n"
        if command_name == 'save':
            self.save()
            return "Saved current document.\n"
        if command_name == 'save-all':
            self.save_all()
            return "Saved open documents.\n"
        if command_name == 'save-as':
            self.save_as()
            return "Opened Save As.\n"
        if command_name in {'save-copy', 'save-copy-as'}:
            self.save_copy_as()
            return "Opened Save Copy As.\n"
        if command_name in {'save-run', 'run'}:
            self.save_and_run()
            return "Ran Save and Run.\n"
        if command_name in {'save-encrypted', 'save-as-encrypted'}:
            self.save_encrypted_copy()
            return "Opened Save As Encrypted.\n"
        if command_name == 'print':
            self.print_file()
            return "Opened print for the current document.\n"
        if command_name in {'export-notes', 'notes-export'}:
            self.export_notes_report()
            return "Opened Export Notes.\n"
        if command_name in {'exit', 'quit'}:
            self.root.after_idle(self.exit_app)
            return "Closing Notepad-X.\n"
        if command_name == 'undo':
            self.undo()
            return "Undid the last change.\n"
        if command_name == 'redo':
            self.redo()
            return "Redid the last undone change.\n"
        if command_name == 'cut':
            self.cut()
            return "Cut selection.\n"
        if command_name == 'copy':
            self.copy()
            return "Copied selection.\n"
        if command_name == 'paste':
            self.paste()
            return "Pasted clipboard contents.\n"
        if command_name in {'select-all', 'selectall'}:
            self.select_all()
            return "Selected all text.\n"
        if command_name == 'find':
            self.show_find_panel()
            return "Opened Find panel.\n"
        if command_name in {'find-next', 'next-find', 'next-match'}:
            self.find_next()
            return "Moved to next find match.\n"
        if command_name in {'find-prev', 'find-previous', 'prev-match'}:
            self.find_previous()
            return "Moved to previous find match.\n"
        if command_name == 'replace':
            self.show_replace_panel()
            return "Opened Replace panel.\n"
        if command_name == 'command-panel':
            return "Command panel is already open.\n"
        if command_name == 'symbols':
            self.show_symbol_navigator()
            return "Opened symbol navigator.\n"
        if command_name == 'project-symbols':
            self.show_symbol_navigator(project_scope=True)
            return "Opened project symbol navigator.\n"
        if command_name == 'fold':
            self.toggle_fold_at_cursor()
            return "Toggled fold at cursor.\n"
        if command_name in {'fold-all', 'collapse-all'}:
            self.collapse_all_folds()
            return "Collapsed foldable blocks.\n"
        if command_name in {'unfold-all', 'expand-all'}:
            self.expand_all_folds()
            return "Expanded all folded blocks.\n"
        if command_name == 'lint':
            doc = self.get_current_doc()
            if doc:
                self.run_diagnostics_for_doc(doc, force=True)
                diagnostics = doc.get('diagnostics', [])
                if not diagnostics:
                    return "No diagnostics.\n"
                return "".join(
                    f"{item.get('severity', 'info').upper()}: line {item.get('line', '?')} - {item.get('message', '')}\n"
                    for item in diagnostics[:25]
                )
            return "No active document.\n"
        if command_name in {'date', 'insert-date'}:
            self.insert_date()
            return "Inserted date.\n"
        if command_name in {'time-date', 'datetime', 'insert-time-date'}:
            self.insert_time_date()
            return "Inserted time/date.\n"
        if command_name == 'font':
            self.show_font_dialog()
            return "Opened font dialog.\n"
        if command_name == 'language':
            return self.run_named_language_command(argument_text)
        if command_name in {'fullscreen', 'full-screen'}:
            self.toggle_fullscreen()
            return f"{'Entered' if self.fullscreen else 'Exited'} full screen.\n"
        if command_name in {'switch-tab', 'next-tab'}:
            self.switch_tab_right()
            return "Switched to the next tab.\n"
        if command_name in {'currently-editing', 'editing-panel'}:
            self.toggle_currently_editing_panel()
            return f"{'Opened' if self.currently_editing_panel_visible else 'Closed'} currently editing panel.\n"
        if command_name in {'cycle-notes', 'notes'}:
            self.goto_next_note()
            return "Jumped to the next note.\n"
        if command_name in {'note-filter', 'filter-notes'}:
            return self.run_named_note_filter_command(argument_text)
        if command_name in {'goto-line', 'line'}:
            self.goto_line_dialog()
            return "Opened Go To Line.\n"
        if command_name in {'top', 'document-top'}:
            self.goto_document_start()
            return "Moved to top of document.\n"
        if command_name in {'bottom', 'document-bottom'}:
            self.goto_document_end()
            return "Moved to bottom of document.\n"
        if command_name == 'theme':
            return self.run_named_theme_command(argument_text)
        if command_name in {'syntax-mode', 'mode'}:
            return self.run_named_syntax_mode_command(argument_text)
        if command_name in {'preview', 'markdown-preview'}:
            if argument_text:
                return self.run_named_toggle_command('preview', argument_text, self.markdown_preview_enabled, self.toggle_markdown_preview, 'Markdown preview')
            self.markdown_preview_enabled.set(not self.markdown_preview_enabled.get())
            self.toggle_markdown_preview()
            return self.format_named_toggle_state('Markdown preview', bool(self.markdown_preview_enabled.get()))
        if command_name == 'compare':
            self.show_split_compare()
            return "Opened compare picker.\n"
        if command_name in {'close-compare', 'compare-close'}:
            if self.compare_active:
                self.close_compare_panel()
                return "Closed compare panel.\n"
            return "No compare panel is open.\n"
        toggle_commands = {
            'autosave': (self.autosave_enabled, self.save_session, 'Autosave'),
            'minimap': (self.minimap_enabled, self.toggle_minimap, 'Minimap'),
            'diagnostics': (self.diagnostics_enabled, self.toggle_diagnostics, 'Diagnostics'),
            'edit-with-notepad-x': (self.edit_with_shell_enabled, self.toggle_edit_with_shell, 'Edit with Notepad-X'),
            'edit-with-shell': (self.edit_with_shell_enabled, self.toggle_edit_with_shell, 'Edit with Notepad-X'),
            'sound': (self.sound_enabled, self.toggle_sound, 'Sound'),
            'status-bar': (self.status_bar_enabled, self.toggle_status_bar, 'Status bar'),
            'numbered-lines': (self.numbered_lines_enabled, self.toggle_numbered_lines, 'Numbered lines'),
            'autocomplete': (self.autocomplete_enabled, self.toggle_autocomplete, 'Autocomplete'),
            'spell-check': (self.spell_check_enabled, self.toggle_spell_check, 'Spell check'),
            'auto-pair': (self.auto_pair_enabled, self.save_session, 'Auto pair brackets/quotes'),
            'compare-multi-edit': (self.compare_multi_edit_enabled, self.save_session, 'Compare multi-edit'),
            'breadcrumbs': (self.breadcrumbs_enabled, self.toggle_breadcrumbs, 'Breadcrumbs'),
            'word-wrap': (self.word_wrap_enabled, self.toggle_word_wrap, 'Word wrap'),
            'sync-page-navigation': (self.sync_page_navigation_enabled, self.save_session, 'Sync page navigation'),
        }
        if command_name in toggle_commands:
            variable, callback, label = toggle_commands[command_name]
            return self.run_named_toggle_command(command_name, argument_text, variable, callback, label)
        if command_name in {'help-contents', 'contents'}:
            self.show_help_contents()
            return "Opened help contents.\n"
        if command_name == 'about':
            self.show_about_dialog()
            return "Opened About Notepad-X.\n"
        return None

    def finish_shell_command(self, command_text, result):
        self.command_runner_active = False
        self.command_runner_thread = None
        if not result:
            self.append_command_output("No output.\n[exit 0]\n\n")
            return
        output_parts = []
        stdout_text = str(result.get('stdout') or '')
        stderr_text = str(result.get('stderr') or '')
        if stdout_text:
            output_parts.append(stdout_text.rstrip() + '\n')
        if stderr_text:
            output_parts.append(stderr_text.rstrip() + '\n')
        if not stdout_text and not stderr_text:
            output_parts.append("No output.\n")
        output_parts.append(f"[exit {result.get('returncode', 0)}]\n\n")
        self.append_command_output("".join(output_parts))

    def run_command_panel(self):
        if not hasattr(self, 'command_entry') or not self.command_entry:
            return "break"
        command_text = self.command_entry.get().strip()
        if not command_text:
            return "break"
        normalized_command = command_text.lower()
        self.hide_command_suggestion_popup()

        if normalized_command in {'cls', 'clear'}:
            self.record_command_history(command_text)
            self.clear_command_output()
            self.command_entry.delete(0, tk.END)
            self.command_entry.focus_set()
            return "break"

        named_result = None
        if command_text.startswith(':'):
            named_result = self.run_named_command(command_text)
        if named_result is not None:
            self.record_command_history(command_text)
            self.command_entry.delete(0, tk.END)
            self.append_command_output(named_result if named_result.endswith('\n') else f"{named_result}\n")
            self.command_entry.focus_set()
            return "break"
        if command_text.startswith(':'):
            self.record_command_history(command_text)
            self.command_entry.delete(0, tk.END)
            self.append_command_output("Unknown built-in command. Type :help for available commands.\n")
            self.command_entry.focus_set()
            return "break"

        if self.command_runner_active:
            self.append_command_output("A command is already running.\n")
            return "break"

        cwd = self.get_doc_working_directory()
        self.command_runner_active = True
        self.record_command_history(command_text)
        self.command_entry.delete(0, tk.END)
        self.command_entry.focus_set()

        def worker():
            result = {'returncode': -1, 'stdout': '', 'stderr': ''}
            try:
                completed = subprocess.run(
                    command_text,
                    shell=True,
                    cwd=cwd,
                    capture_output=True,
                    text=True,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
                result['returncode'] = completed.returncode
                result['stdout'] = completed.stdout or ''
                result['stderr'] = completed.stderr or ''
            except Exception as exc:
                result['stderr'] = str(exc)
            self.root.after(0, lambda current=command_text, payload=result: self.finish_shell_command(current, payload))

        self.command_runner_thread = threading.Thread(target=worker, name='NotepadXCommandPanel', daemon=True)
        self.command_runner_thread.start()
        self.append_command_output(f"$ {command_text}\n[running in {cwd}]\n")
        return "break"

    def toggle_minimap(self):
        show_minimap = self.minimap_enabled.get()
        for doc in self.documents.values():
            minimap = doc.get('minimap')
            if not minimap:
                continue
            if show_minimap:
                minimap.grid()
                self.schedule_minimap_refresh(doc)
            else:
                minimap.grid_remove()
        if self.compare_view:
            minimap = self.compare_view.get('minimap')
            if minimap:
                if show_minimap:
                    minimap.grid()
                    self.schedule_minimap_refresh(self.compare_view)
                else:
                    minimap.grid_remove()
        self.save_session()
        return "break"

    def schedule_minimap_refresh(self, doc):
        if not doc:
            return
        existing_job = doc.get('minimap_job')
        if existing_job:
            try:
                self.root.after_cancel(existing_job)
            except tk.TclError:
                pass
        try:
            doc['minimap_job'] = self.root.after(80, lambda current=doc: self.refresh_minimap(current))
        except tk.TclError:
            doc['minimap_job'] = None

    def invalidate_minimap_cache(self, doc, clear_progress=True):
        if not doc:
            return
        doc['minimap_model'] = None
        doc['minimap_model_dirty'] = True
        if clear_progress:
            doc['minimap_progressive_state'] = None

    def create_minimap_model(self, total_lines, sample_step, segment_max_lengths, complete=True):
        total_lines = max(1, int(total_lines or 1))
        sample_step = max(1, int(sample_step or 1))
        segment_count = max(1, (total_lines + sample_step - 1) // sample_step)
        segments = []
        normalized_lengths = []
        for index in range(segment_count):
            start_line = (index * sample_step) + 1
            end_line = min(total_lines, start_line + sample_step - 1)
            segment_length = 0
            if index < len(segment_max_lengths):
                try:
                    segment_length = max(0, int(segment_max_lengths[index] or 0))
                except (TypeError, ValueError):
                    segment_length = 0
            normalized_lengths.append(segment_length)
            segments.append({
                'start_line': start_line,
                'end_line': end_line,
                'length': segment_length,
            })
        return {
            'total_lines': total_lines,
            'sample_step': sample_step,
            'segments': segments,
            'max_length': max(1, max(normalized_lengths or [0])),
            'complete': bool(complete),
        }

    def build_minimap_model_from_content(self, content, total_lines=None):
        content = content if isinstance(content, str) else ''
        line_lengths = [len(line) for line in content.split('\n')]
        total_lines = max(1, int(total_lines or len(line_lengths) or 1))
        if len(line_lengths) < total_lines:
            line_lengths.extend([0] * (total_lines - len(line_lengths)))
        elif len(line_lengths) > total_lines:
            total_lines = len(line_lengths)
        sample_step = max(1, total_lines // self.minimap_max_segments)
        segment_count = max(1, (total_lines + sample_step - 1) // sample_step)
        segment_max_lengths = [0] * segment_count
        for line_number in range(1, total_lines + 1):
            segment_index = min(segment_count - 1, (line_number - 1) // sample_step)
            segment_max_lengths[segment_index] = max(segment_max_lengths[segment_index], line_lengths[line_number - 1])
        return self.create_minimap_model(total_lines, sample_step, segment_max_lengths, complete=True)

    def build_minimap_model_from_line_starts(self, line_starts, file_size):
        normalized_starts = list(line_starts or [0])
        total_lines = max(1, len(normalized_starts))
        sample_step = max(1, total_lines // self.minimap_max_segments)
        segment_count = max(1, (total_lines + sample_step - 1) // sample_step)
        segment_max_lengths = [0] * segment_count
        safe_file_size = max(0, int(file_size or 0))
        for line_index in range(total_lines):
            start_byte = normalized_starts[line_index]
            if line_index + 1 < len(normalized_starts):
                end_byte = normalized_starts[line_index + 1]
            else:
                end_byte = safe_file_size
            line_length = max(0, int(end_byte - start_byte))
            segment_index = min(segment_count - 1, line_index // sample_step)
            segment_max_lengths[segment_index] = max(segment_max_lengths[segment_index], line_length)
        return self.create_minimap_model(total_lines, sample_step, segment_max_lengths, complete=True)

    def start_progressive_minimap_build(self, doc, total_lines_hint=None):
        if not doc:
            return
        total_lines = max(1, int(total_lines_hint or 1))
        sample_step = max(1, total_lines // self.minimap_max_segments)
        segment_count = max(1, (total_lines + sample_step - 1) // sample_step)
        doc['minimap_progressive_state'] = {
            'total_lines': total_lines,
            'sample_step': sample_step,
            'segment_max_lengths': [0] * segment_count,
            'processed_lines': 0,
            'remainder': '',
            'max_length': 1,
        }
        doc['minimap_model'] = self.create_minimap_model(total_lines, sample_step, [0] * segment_count, complete=False)
        doc['minimap_model_dirty'] = False

    def append_progressive_minimap_chunk(self, doc, chunk_text, finalize=False):
        if not doc:
            return
        state = doc.get('minimap_progressive_state')
        if not state:
            return
        buffer = f"{state.get('remainder', '')}{chunk_text or ''}"
        if finalize:
            lines = buffer.split('\n')
            state['remainder'] = ''
        else:
            lines = buffer.split('\n')
            state['remainder'] = lines.pop() if lines else ''
        total_lines = max(1, int(state.get('total_lines', 1) or 1))
        sample_step = max(1, int(state.get('sample_step', 1) or 1))
        segment_max_lengths = state.get('segment_max_lengths') or []
        processed_lines = int(state.get('processed_lines', 0) or 0)
        max_length = max(1, int(state.get('max_length', 1) or 1))
        for line in lines:
            if processed_lines >= total_lines:
                break
            processed_lines += 1
            segment_index = min(len(segment_max_lengths) - 1, (processed_lines - 1) // sample_step)
            line_length = len(line)
            if segment_index >= 0:
                segment_max_lengths[segment_index] = max(segment_max_lengths[segment_index], line_length)
            max_length = max(max_length, line_length)
        if finalize and processed_lines < total_lines:
            processed_lines = total_lines
        state['processed_lines'] = processed_lines
        state['max_length'] = max_length
        doc['minimap_model'] = self.create_minimap_model(total_lines, sample_step, segment_max_lengths, complete=finalize)
        doc['minimap_model_dirty'] = False
        if finalize:
            doc['minimap_progressive_state'] = None

    def draw_minimap_segments(self, minimap, width, height, model):
        if not model:
            return
        total_lines = max(1, int(model.get('total_lines', 1) or 1))
        max_length = max(1, int(model.get('max_length', 1) or 1))
        for segment in model.get('segments', []):
            segment_length = max(0, int(segment.get('length', 0) or 0))
            if segment_length <= 0:
                continue
            start_line = max(1, int(segment.get('start_line', 1) or 1))
            end_line = max(start_line, min(total_lines, int(segment.get('end_line', start_line) or start_line)))
            y0 = int((start_line - 1) / total_lines * height)
            y1 = max(y0 + 1, int(end_line / total_lines * height))
            bar_width = max(2, int((segment_length / max_length) * max(4, width - 6)))
            minimap.create_rectangle(2, y0, min(width - 2, 2 + bar_width), y1, fill='#4f708f', outline='')

    def draw_minimap_overlays(self, doc, minimap, text, width, height, total_lines):
        diagnostics = doc.get('diagnostics', [])
        for diagnostic in diagnostics:
            line_number = max(1, int(diagnostic.get('line', 1)))
            y = int((line_number - 1) / total_lines * height)
            marker_color = '#ff6b6b' if diagnostic.get('severity') == 'error' else '#ffcc66'
            minimap.create_rectangle(width - 4, y, width - 1, min(height, y + 3), fill=marker_color, outline='')

        try:
            top_line = int(text.index('@0,0').split('.')[0])
            bottom_line = int(text.index(f'@0,{max(1, text.winfo_height())}').split('.')[0])
        except (tk.TclError, ValueError):
            top_line = 1
            bottom_line = min(total_lines, max(1, total_lines // 5))
        if doc.get('virtual_mode'):
            window_start_line = max(1, int(doc.get('window_start_line', 1) or 1))
            top_line = max(1, min(total_lines, window_start_line + top_line - 1))
            bottom_line = max(top_line, min(total_lines, window_start_line + bottom_line - 1))
        view_y0 = int((top_line - 1) / total_lines * height)
        view_y1 = max(view_y0 + 6, int(bottom_line / total_lines * height))
        minimap.create_rectangle(0, view_y0, width, view_y1, outline='#9ecbff', width=1)

    def refresh_minimap(self, doc):
        if not doc:
            return
        doc['minimap_job'] = None
        if not self.minimap_enabled.get():
            return
        minimap = doc.get('minimap')
        text = doc.get('text')
        if not minimap or not text:
            return
        try:
            if not minimap.winfo_exists() or not text.winfo_exists() or not minimap.winfo_ismapped():
                return
        except tk.TclError:
            return

        minimap.delete('all')

        width = max(1, minimap.winfo_width())
        height = max(1, minimap.winfo_height())
        if height <= 2:
            height = max(120, minimap.winfo_reqheight())
        model = doc.get('minimap_model')
        if not model or doc.get('minimap_model_dirty', True):
            if doc.get('virtual_mode') and self.is_virtual_index_ready(doc):
                model = self.build_minimap_model_from_line_starts(
                    doc.get('line_starts'),
                    doc.get('file_size_bytes', 0)
                )
            else:
                try:
                    content = text.get('1.0', 'end-1c')
                    total_lines = max(1, int(text.index('end-1c').split('.')[0]))
                except tk.TclError:
                    return
                model = self.build_minimap_model_from_content(content, total_lines=total_lines)
            doc['minimap_model'] = model
            doc['minimap_model_dirty'] = False
        self.draw_minimap_segments(minimap, width, height, model)
        if model.get('complete', True):
            self.draw_minimap_overlays(
                doc,
                minimap,
                text,
                width,
                height,
                max(1, int(model.get('total_lines', 1) or 1))
            )

    def on_minimap_click(self, event, doc):
        if not self.minimap_enabled.get() or not doc:
            return "break"
        text = doc.get('text')
        minimap = doc.get('minimap')
        if not text or not minimap:
            return "break"
        try:
            if doc.get('virtual_mode') and self.is_virtual_index_ready(doc):
                total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
            else:
                total_lines = max(1, int(text.index('end-1c').split('.')[0]))
            height = max(1, minimap.winfo_height())
            target_ratio = min(1.0, max(0.0, float(event.y) / float(height)))
            target_line = max(1, min(total_lines, int(target_ratio * total_lines) + 1))
            if doc.get('virtual_mode') and self.is_virtual_index_ready(doc):
                self.load_virtual_window(doc, target_line)
            else:
                text.mark_set(tk.INSERT, f'{target_line}.0')
                text.see(f'{target_line}.0')
            self.set_last_active_editor_widget(text)
        except (tk.TclError, ValueError, ZeroDivisionError):
            return "break"
        target_doc = self.get_doc_for_text_widget(text)
        if target_doc:
            self.remember_doc_view_state(target_doc)
            self.update_line_number_gutter(target_doc)
        elif doc is self.compare_view:
            self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        self.update_status()
        return "break"

    def build_breadcrumb_text(self, doc, text_widget=None):
        segments = self.get_breadcrumb_segments(doc, text_widget=text_widget)
        return "  >  ".join(segment['text'] for segment in segments if segment.get('text'))

    def update_breadcrumbs(self):
        current_doc = self.get_current_doc()
        label_text = ""
        current_copy_path = self.get_doc_copyable_path(current_doc)
        if self.breadcrumbs_enabled.get():
            label_text = self.build_breadcrumb_text(current_doc, self.get_active_search_widget())
        try:
            if self.breadcrumbs_enabled.get():
                self.breadcrumbs_bar.grid()
            else:
                self.breadcrumbs_bar.grid_remove()
            self.breadcrumbs_label.config(text=label_text, cursor='hand2' if current_copy_path else 'arrow')
        except tk.TclError:
            pass

        compare_text = ""
        compare_copy_path = None
        compare_copy_name = None
        if self.is_side_panel_visible() and self.compare_view:
            compare_text = self.build_breadcrumb_text(self.compare_view, self.compare_view.get('text'))
            compare_copy_path = self.get_doc_copyable_path(self.compare_view)
            compare_copy_name = self.get_doc_copyable_name(self.compare_view)
        try:
            self.compare_breadcrumbs.config(text=compare_text, cursor='hand2' if compare_copy_path else 'arrow')
        except tk.TclError:
            pass
        try:
            self.compare_title.config(cursor='hand2' if compare_copy_name else 'arrow')
        except tk.TclError:
            pass

    def toggle_breadcrumbs(self):
        self.update_breadcrumbs()
        self.save_session()
        return "break"

    def get_breadcrumb_click_value(self, doc, event=None, text_widget=None):
        segments = self.get_breadcrumb_segments(doc, text_widget=text_widget)
        if not segments:
            return None
        if event is None or getattr(event, 'widget', None) is None:
            for segment in segments:
                if segment.get('copy'):
                    return segment['copy']
            return None
        try:
            widget = event.widget
            widget_font = tkfont.Font(font=widget.cget('font'))
            padding_x = int(float(widget.cget('padx') or 0))
            click_x = int(getattr(event, 'x', 0) or 0)
        except (tk.TclError, TypeError, ValueError, AttributeError):
            return None

        separator = "  >  "
        current_x = padding_x
        for segment in segments:
            segment_text = str(segment.get('text') or '')
            segment_width = widget_font.measure(segment_text)
            if current_x <= click_x <= (current_x + segment_width):
                return segment.get('copy')
            current_x += segment_width
            current_x += widget_font.measure(separator)
        return None

    def copy_text_to_clipboard(self, clipboard_text, event=None):
        if not clipboard_text:
            return "break"
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(clipboard_text)
            self.root.update_idletasks()
        except tk.TclError:
            return "break"
        widget = getattr(event, 'widget', None)
        if widget is not None:
            toast_x = widget.winfo_rootx() + getattr(event, 'x', 0)
            toast_y = widget.winfo_rooty() + getattr(event, 'y', 0)
        else:
            toast_x = self.root.winfo_pointerx()
            toast_y = self.root.winfo_pointery()
        self.show_toast(self.tr('clipboard.path_copied', 'Copied to clipboard'), x=toast_x, y=toast_y)
        return "break"

    def copy_breadcrumb_path(self, event=None, target='primary'):
        doc = self.compare_view if target in {'compare', 'compare_title'} else self.get_current_doc()
        if target == 'compare_title':
            clipboard_text = self.get_doc_copyable_name(doc) or self.get_doc_copyable_path(doc)
        else:
            breadcrumb_widget = self.compare_view.get('text') if target == 'compare' and self.compare_view else self.get_active_search_widget()
            clipboard_text = self.get_breadcrumb_click_value(doc, event=event, text_widget=breadcrumb_widget)
        return self.copy_text_to_clipboard(clipboard_text, event)

    def sanitize_search_history_entries(self, entries):
        cleaned_entries = []
        seen = set()
        source_entries = entries if isinstance(entries, list) else []
        for entry in source_entries:
            if not isinstance(entry, str):
                continue
            value = entry.strip()
            if not value:
                continue
            lowered = value.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            cleaned_entries.append(value)
            if len(cleaned_entries) >= self.max_search_history:
                break
        return cleaned_entries

    def sanitize_command_history_entries(self, entries):
        cleaned_entries = []
        seen = set()
        source_entries = entries if isinstance(entries, list) else []
        for entry in source_entries:
            if not isinstance(entry, str):
                continue
            value = entry.strip()
            if not value:
                continue
            lowered = value.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            cleaned_entries.append(value)
            if len(cleaned_entries) >= self.max_command_history:
                break
        return cleaned_entries

    def cancel_search_history_hide_job(self):
        if self.search_history_hide_job:
            try:
                self.root.after_cancel(self.search_history_hide_job)
            except tk.TclError:
                pass
            self.search_history_hide_job = None

    def hide_search_history_popup(self):
        self.cancel_search_history_hide_job()
        popup = getattr(self, 'search_history_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.search_history_popup = None
        self.search_history_listbox = None
        self.search_history_entry = None
        self.search_history_items = []

    def search_history_popup_visible(self):
        popup = getattr(self, 'search_history_popup', None)
        if popup is None:
            return False
        try:
            return popup.winfo_exists()
        except tk.TclError:
            return False

    def maybe_hide_search_history_popup(self):
        self.search_history_hide_job = None
        focus_widget = self.safe_focus_get()
        listbox = getattr(self, 'search_history_listbox', None)
        if focus_widget is not None and focus_widget in (self.search_history_entry, listbox):
            return
        self.hide_search_history_popup()

    def schedule_search_history_popup_hide(self):
        self.cancel_search_history_hide_job()
        try:
            self.search_history_hide_job = self.root.after(120, self.maybe_hide_search_history_popup)
        except tk.TclError:
            self.search_history_hide_job = None

    def get_search_history_for_entry(self, entry_widget):
        if entry_widget in (getattr(self, 'find_entry', None), getattr(self, 'replace_find_entry', None)):
            return list(self.find_history)
        if entry_widget == getattr(self, 'find_in_entry', None):
            return list(self.find_in_history)
        return []

    def get_search_history_matches(self, history_items, query_text):
        normalized_query = str(query_text or '').strip().lower()
        if not normalized_query:
            return list(history_items[:self.max_search_history])

        prefix_matches = []
        substring_matches = []
        for item in history_items:
            lowered = item.lower()
            if lowered.startswith(normalized_query):
                prefix_matches.append(item)
            elif normalized_query in lowered:
                substring_matches.append(item)
        return (prefix_matches + substring_matches)[:self.max_search_history]

    def select_search_history_index(self, index):
        if not self.search_history_popup_visible() or not self.search_history_listbox:
            return False
        listbox = self.search_history_listbox
        if listbox.size() <= 0:
            return False
        safe_index = max(0, min(listbox.size() - 1, index))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(safe_index)
        listbox.activate(safe_index)
        listbox.see(safe_index)
        return True

    def update_search_history_popup(self, entry_widget, force_show=False):
        if not entry_widget:
            self.hide_search_history_popup()
            return False
        try:
            if not entry_widget.winfo_exists() or not entry_widget.winfo_ismapped():
                self.hide_search_history_popup()
                return False
            current_text = entry_widget.get()
        except tk.TclError:
            self.hide_search_history_popup()
            return False

        history_items = self.get_search_history_for_entry(entry_widget)
        if not history_items:
            self.hide_search_history_popup()
            return False

        query_text = current_text.strip()
        suggestions = self.get_search_history_matches(history_items, query_text)
        if not suggestions or (not force_show and not query_text):
            self.hide_search_history_popup()
            return False

        selected_value = None
        if self.search_history_popup_visible() and self.search_history_entry == entry_widget and self.search_history_listbox:
            selection = self.search_history_listbox.curselection()
            if selection:
                selected_value = self.search_history_listbox.get(selection[0])
        elif self.search_history_entry != entry_widget:
            self.hide_search_history_popup()

        popup = self.search_history_popup
        if not self.search_history_popup_visible():
            popup = self.create_popup_toplevel(self.root)
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
                exportselection=False,
                takefocus=False,
                selectmode='browse',
                width=max(18, min(60, max(len(item) for item in suggestions) + 2)),
                height=min(8, len(suggestions))
            )
            listbox.pack(fill='both', expand=True)
            listbox.bind('<Motion>', self.on_search_history_listbox_motion)
            listbox.bind('<ButtonRelease-1>', self.on_search_history_listbox_click)
            self.search_history_popup = popup
            self.search_history_listbox = listbox
        else:
            listbox = self.search_history_listbox
            listbox.configure(
                width=max(18, min(60, max(len(item) for item in suggestions) + 2)),
                height=min(8, len(suggestions))
            )

        listbox.delete(0, tk.END)
        for suggestion in suggestions:
            listbox.insert(tk.END, suggestion)

        selected_index = 0
        if selected_value in suggestions:
            selected_index = suggestions.index(selected_value)
        self.select_search_history_index(selected_index)

        self.show_popup_toplevel(
            popup,
            entry_widget.winfo_rootx(),
            entry_widget.winfo_rooty() + entry_widget.winfo_height() + 2
        )
        self.search_history_entry = entry_widget
        self.search_history_items = suggestions
        return True

    def move_search_history_selection(self, direction):
        if not self.search_history_popup_visible() or not self.search_history_listbox:
            return None
        listbox = self.search_history_listbox
        selection = listbox.curselection()
        current_index = selection[0] if selection else 0
        next_index = max(0, min(listbox.size() - 1, current_index + direction))
        self.select_search_history_index(next_index)
        return "break"

    def accept_search_history_selection(self):
        if not self.search_history_popup_visible() or not self.search_history_listbox or not self.search_history_entry:
            return False
        selection = self.search_history_listbox.curselection()
        if selection:
            suggestion = self.search_history_listbox.get(selection[0])
        elif self.search_history_listbox.size() > 0:
            suggestion = self.search_history_listbox.get(0)
        else:
            return False

        entry_widget = self.search_history_entry
        try:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, suggestion)
            entry_widget.icursor(tk.END)
            entry_widget.focus_set()
        except tk.TclError:
            self.hide_search_history_popup()
            return False

        self.hide_search_history_popup()
        if entry_widget in (getattr(self, 'find_entry', None), getattr(self, 'replace_find_entry', None)):
            self.on_find_entry_change()
        elif entry_widget == getattr(self, 'find_in_entry', None):
            self.on_find_in_entry_change()
        return True

    def on_search_history_listbox_motion(self, event=None):
        listbox = getattr(self, 'search_history_listbox', None)
        if not listbox:
            return None
        index = listbox.nearest(getattr(event, 'y', 0))
        self.select_search_history_index(index)
        return None

    def on_search_history_listbox_click(self, event=None):
        listbox = getattr(self, 'search_history_listbox', None)
        if not listbox:
            return "break"
        index = listbox.nearest(getattr(event, 'y', 0))
        self.select_search_history_index(index)
        self.accept_search_history_selection()
        return "break"

    def on_search_history_entry_focus(self, event=None):
        entry_widget = getattr(event, 'widget', None)
        if not entry_widget:
            return None
        self.cancel_search_history_hide_job()
        try:
            self.root.after_idle(lambda widget=entry_widget: self.update_search_history_popup(widget, force_show=True))
        except tk.TclError:
            return None
        return None

    def on_search_history_entry_focus_out(self, event=None):
        self.schedule_search_history_popup_hide()
        return None

    def handle_search_history_keypress(self, event):
        entry_widget = getattr(event, 'widget', None)
        if not entry_widget:
            return None

        keysym = str(getattr(event, 'keysym', '') or '')
        popup_visible = self.search_history_popup_visible() and self.search_history_entry == entry_widget

        if keysym == 'Escape' and popup_visible:
            self.hide_search_history_popup()
            return "break"

        if keysym in {'Up', 'Down'}:
            if popup_visible:
                return self.move_search_history_selection(-1 if keysym == 'Up' else 1)
            if self.update_search_history_popup(entry_widget, force_show=True):
                if keysym == 'Up' and self.search_history_listbox:
                    self.select_search_history_index(self.search_history_listbox.size() - 1)
                return "break"
            return None

        if keysym == 'Tab' and popup_visible:
            if self.accept_search_history_selection():
                return "break"
            return None

        if popup_visible and keysym in {'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}:
            self.hide_search_history_popup()
        return None

    def add_search_history_entry(self, attribute_name, query):
        cleaned_query = str(query or '').strip()
        if not cleaned_query:
            return False

        existing_history = list(getattr(self, attribute_name, []))
        lowered_query = cleaned_query.lower()
        updated_history = [
            item for item in existing_history
            if str(item).strip().lower() != lowered_query
        ]
        updated_history.insert(0, cleaned_query)
        updated_history = updated_history[:self.max_search_history]
        if updated_history == existing_history:
            return False

        setattr(self, attribute_name, updated_history)
        self.save_session()
        return True

    def add_find_history_entry(self, query):
        return self.add_search_history_entry('find_history', query)

    def add_find_in_history_entry(self, query):
        return self.add_search_history_entry('find_in_history', query)

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
            messagebox.showinfo(
                self.tr('large_file.title', 'Large File Mode'),
                self.tr('large_file.find_unavailable', 'Find is not available in buffered large-file mode yet.'),
                parent=self.root
            )
            return "break"

        self.cancel_find_change_job()
        self.hide_search_history_popup()
        if self.command_panel_visible:
            self.command_frame.grid_remove()
            self.command_panel_visible = False
        if self.replace_panel_visible:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.clear_find_highlights()

        if not self.find_panel_visible:
            self.bottom_frame.grid()
            self.find_frame.grid(sticky='ew')
            self.find_panel_visible = True
            self.update_find_match_summary("")
            self.find_entry.focus_set()
        else:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.find_entry.delete(0, tk.END)
            self.find_in_entry.delete(0, tk.END)
            self.clear_find_highlights()
            self.update_find_match_summary("")
            self.focus_last_active_editor()

        self.update_bottom_panel_visibility()

        return "break"

    def show_replace_panel(self):
        current_doc = self.get_current_doc()
        if current_doc and current_doc.get('virtual_mode'):
            messagebox.showinfo(
                self.tr('large_file.title', 'Large File Mode'),
                self.tr('large_file.replace_unavailable', 'Replace is not available in buffered large-file mode.'),
                parent=self.root
            )
            return "break"

        self.cancel_find_change_job()
        self.hide_search_history_popup()
        if self.command_panel_visible:
            self.command_frame.grid_remove()
            self.command_panel_visible = False
        if self.find_panel_visible:
            self.find_frame.grid_remove()
            self.find_panel_visible = False
            self.clear_find_highlights()

        if not self.replace_panel_visible:
            self.bottom_frame.grid()
            self.replace_frame.grid(sticky='ew')
            self.replace_panel_visible = True
            self.update_find_match_summary("")
            self.replace_find_entry.focus_set()
        else:
            self.replace_frame.grid_remove()
            self.replace_panel_visible = False
            self.replace_find_entry.delete(0, tk.END)
            self.clear_find_highlights()
            self.update_find_match_summary("")
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

    def get_side_panel_text_widget(self):
        if not self.is_side_panel_visible() or not self.compare_view:
            return None
        side_panel_widget = self.compare_view.get('text')
        if not side_panel_widget:
            return None
        try:
            return side_panel_widget if side_panel_widget.winfo_exists() else None
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
        targets = self.get_page_navigation_targets(event)
        if not targets:
            return "break"
        for widget in targets:
            doc = self.get_navigation_doc_for_widget(widget)
            if doc and doc.get('virtual_mode'):
                if self.is_virtual_index_ready(doc):
                    try:
                        self.load_virtual_window(doc, 1)
                    except tk.TclError:
                        continue
                else:
                    doc['pending_virtual_target_line'] = 1
            try:
                widget.mark_set(tk.INSERT, '1.0')
                widget.tag_remove('sel', '1.0', tk.END)
                widget.see('1.0')
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            if doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
                self.schedule_minimap_refresh(doc)
        self.update_status()
        return "break"

    def goto_document_end(self, event=None):
        targets = self.get_page_navigation_targets(event)
        if not targets:
            return "break"
        for widget in targets:
            doc = self.get_navigation_doc_for_widget(widget)
            try:
                if doc and doc.get('virtual_mode'):
                    if self.is_virtual_index_ready(doc):
                        last_line = max(1, int(doc.get('total_file_lines', 1) or 1))
                        self.load_virtual_window(doc, last_line)
                        local_line = max(1, min(
                            max(1, doc.get('window_end_line', last_line) - doc.get('window_start_line', 1) + 1),
                            last_line - doc.get('window_start_line', 1) + 1
                        ))
                        target_index = f'{local_line}.0'
                    else:
                        doc['pending_virtual_target_line'] = 'end'
                        target_index = '1.0'
                else:
                    last_line = widget.index('end-1c').split('.')[0]
                    target_index = f'{last_line}.0'
                widget.mark_set(tk.INSERT, target_index)
                widget.tag_remove('sel', '1.0', tk.END)
                widget.see(target_index)
                self.set_last_active_editor_widget(widget)
            except tk.TclError:
                continue
            if doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
                self.schedule_minimap_refresh(doc)
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

    def get_diagnostic_doc_for_widget(self, widget):
        if not isinstance(widget, tk.Text):
            return None
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and widget == compare_widget:
            return self.compare_view
        for doc in self.documents.values():
            if doc.get('text') == widget:
                return doc
        return None

    def get_diagnostics_for_line(self, doc, line_number):
        if not doc:
            return []
        line_number = max(1, int(line_number))
        diagnostics = [item for item in (doc.get('diagnostics') or []) if int(item.get('line', 0) or 0) == line_number]
        diagnostics.sort(key=lambda item: (0 if item.get('severity') == 'error' else 1, int(item.get('column', 0) or 0)))
        return diagnostics

    def format_diagnostic_tooltip_lines(self, diagnostics):
        lines = []
        for item in diagnostics:
            severity = 'Error' if item.get('severity') == 'error' else 'Warning'
            message = str(item.get('message') or '').strip() or severity
            column_number = item.get('column')
            if column_number not in (None, '', 0):
                lines.append(f"{severity}: {message} (col {column_number})")
            else:
                lines.append(f"{severity}: {message}")
        return lines

    def hide_diagnostic_tooltip(self, event=None):
        job = getattr(self, 'diagnostic_tooltip_job', None)
        if job:
            try:
                self.root.after_cancel(job)
            except tk.TclError:
                pass
        self.diagnostic_tooltip_job = None
        popup = getattr(self, 'diagnostic_tooltip_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        self.diagnostic_tooltip_popup = None
        self.diagnostic_tooltip_doc = None
        self.diagnostic_tooltip_signature = None
        return None

    def show_diagnostic_tooltip(self, doc, widget, x_root, y_root, line_number, diagnostics, signature):
        self.diagnostic_tooltip_job = None
        if signature != self.diagnostic_tooltip_signature:
            return
        if not doc or not diagnostics or widget is None:
            return
        try:
            if not widget.winfo_exists():
                return
        except tk.TclError:
            return

        self.hide_diagnostic_tooltip()
        self.diagnostic_tooltip_signature = signature

        popup = self.create_popup_toplevel(self.root)
        popup.configure(bg='#1f2430')

        accent_color = '#ff6b6b' if any(item.get('severity') == 'error' for item in diagnostics) else '#ffcc66'
        frame = tk.Frame(popup, bg='#1f2430', highlightthickness=1, highlightbackground=accent_color)
        frame.pack(fill='both', expand=True)

        error_count = sum(1 for item in diagnostics if item.get('severity') == 'error')
        warning_count = len(diagnostics) - error_count
        if len(diagnostics) == 1:
            header_text = f"{'Error' if error_count else 'Warning'} on line {line_number}"
        else:
            summary_parts = []
            if error_count:
                summary_parts.append(f"{error_count} error{'s' if error_count != 1 else ''}")
            if warning_count:
                summary_parts.append(f"{warning_count} warning{'s' if warning_count != 1 else ''}")
            header_text = f"Line {line_number}: {', '.join(summary_parts)}"
        tk.Label(
            frame,
            text=header_text,
            bg='#1f2430',
            fg=accent_color,
            font=('Segoe UI', 9, 'bold'),
            anchor=self.ui_anchor_start()
        ).pack(fill='x', padx=10, pady=(8, 2))

        tk.Label(
            frame,
            text="\n".join(self.format_diagnostic_tooltip_lines(diagnostics)),
            bg='#1f2430',
            fg='#f5f5f5',
            font=('Segoe UI', 9),
            justify=self.ui_justify(),
            wraplength=360,
            anchor=self.ui_anchor_start()
        ).pack(fill='both', expand=True, padx=10, pady=(0, 8))

        popup.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        popup_width = popup.winfo_reqwidth()
        popup_height = popup.winfo_reqheight()
        x = min(max(0, x_root + 16), max(0, screen_width - popup_width - 12))
        y = min(max(0, y_root + 20), max(0, screen_height - popup_height - 12))
        self.show_popup_toplevel(popup, x, y)
        self.diagnostic_tooltip_popup = popup
        self.diagnostic_tooltip_doc = doc

    def handle_diagnostic_hover(self, event=None):
        widget = getattr(event, 'widget', None)
        if not isinstance(widget, tk.Text):
            self.hide_diagnostic_tooltip()
            return None
        doc = self.get_diagnostic_doc_for_widget(widget)
        if not doc or not self.diagnostics_enabled.get():
            self.hide_diagnostic_tooltip()
            return None
        try:
            index = widget.index(f"@{event.x},{event.y}")
            line_number = int(str(index).split('.')[0])
        except (tk.TclError, ValueError, AttributeError):
            self.hide_diagnostic_tooltip()
            return None

        diagnostics = self.get_diagnostics_for_line(doc, line_number)
        if not diagnostics:
            self.hide_diagnostic_tooltip()
            return None

        signature = (
            str(widget),
            int(line_number),
            tuple((item.get('severity'), item.get('message'), item.get('column')) for item in diagnostics),
        )
        if signature == self.diagnostic_tooltip_signature and (
            self.diagnostic_tooltip_popup is not None or self.diagnostic_tooltip_job is not None
        ):
            return None

        self.hide_diagnostic_tooltip()
        self.diagnostic_tooltip_signature = signature
        try:
            self.diagnostic_tooltip_job = self.root.after(
                self.diagnostic_tooltip_delay_ms,
                lambda current_doc=doc, current_widget=widget, x=event.x_root, y=event.y_root, line=line_number, items=tuple(diagnostics), current_signature=signature:
                    self.show_diagnostic_tooltip(current_doc, current_widget, x, y, line, list(items), current_signature)
            )
        except tk.TclError:
            self.diagnostic_tooltip_job = None
        return None

    def get_page_navigation_source_widget(self, event=None):
        widget = getattr(event, 'widget', None)
        if isinstance(widget, tk.Text):
            side_panel_widget = self.get_side_panel_text_widget()
            if side_panel_widget is not None and widget == side_panel_widget:
                return widget
            for doc in self.documents.values():
                if doc.get('text') == widget:
                    return widget

        active_widget = self.get_active_search_widget()
        if active_widget is not None:
            return active_widget

        hovered_widget = self.get_hovered_editor_widget()
        if hovered_widget is not None:
            return hovered_widget
        return None

    def get_page_navigation_targets(self, event=None):
        source_widget = self.get_page_navigation_source_widget(event)
        if source_widget is None:
            return []

        if self.sync_page_navigation_enabled.get() and self.is_side_panel_visible():
            targets = []
            if isinstance(self.text, tk.Text):
                targets.append(self.text)
            side_panel_widget = self.get_side_panel_text_widget()
            if side_panel_widget is not None and side_panel_widget not in targets:
                targets.append(side_panel_widget)
            if targets:
                return targets

        return [source_widget]

    def sync_widget_insert_to_visible_line(self, widget):
        try:
            visible_index = widget.index("@0,0")
            visible_line = visible_index.split('.')[0]
            widget.mark_set(tk.INSERT, f"{visible_line}.0")
        except tk.TclError:
            pass

    def get_navigation_doc_for_widget(self, widget):
        side_panel_widget = self.get_side_panel_text_widget()
        if side_panel_widget is not None and widget == side_panel_widget:
            return self.compare_view
        for doc in self.documents.values():
            if doc.get('text') == widget:
                return doc
        return None

    def get_visible_widget_line_bounds(self, widget, doc=None):
        try:
            widget.update_idletasks()
            top_local_line = int(widget.index('@0,0').split('.')[0])
        except (tk.TclError, ValueError, AttributeError):
            return None

        bottom_local_line = top_local_line
        try:
            index = f"{top_local_line}.0"
            while True:
                info = widget.dlineinfo(index)
                if info is None:
                    break
                local_line = int(index.split('.')[0])
                bottom_local_line = max(bottom_local_line, local_line)
                next_index = widget.index(f"{local_line + 1}.0")
                if next_index == index:
                    break
                index = next_index
        except (tk.TclError, ValueError, AttributeError):
            pass

        if doc and doc.get('virtual_mode'):
            window_start = doc.get('window_start_line', 1)
            return (
                window_start + top_local_line - 1,
                window_start + bottom_local_line - 1
            )
        return top_local_line, bottom_local_line

    def scroll_widget_page_by_visible_lines(self, widget, direction):
        doc = self.get_navigation_doc_for_widget(widget)
        bounds = self.get_visible_widget_line_bounds(widget, doc)

        if not bounds:
            try:
                widget.yview_scroll(direction, 'page')
            except tk.TclError:
                return False
            self.sync_widget_insert_to_visible_line(widget)
            return True

        top_line, bottom_line = bounds
        page_lines = max(1, bottom_line - top_line + 1)
        can_rebuffer_virtual_doc = bool(doc and doc.get('virtual_mode') and doc.get('line_starts') is not None)

        if doc and doc.get('virtual_mode'):
            last_line = max(
                1,
                int(doc.get('total_file_lines' if can_rebuffer_virtual_doc else 'window_end_line', bottom_line))
            )
        else:
            try:
                last_line = max(1, int(widget.index('end-1c').split('.')[0]))
            except (tk.TclError, ValueError, AttributeError):
                last_line = bottom_line

        max_top_line = max(1, last_line - page_lines + 1)
        if direction < 0:
            target_top_line = max(1, top_line - page_lines)
        else:
            target_top_line = min(max_top_line, bottom_line + 1)

        try:
            if doc and doc.get('virtual_mode'):
                window_start = doc.get('window_start_line', 1)
                if can_rebuffer_virtual_doc and (
                    target_top_line < window_start or target_top_line > doc.get('window_end_line', window_start)
                ):
                    self.load_virtual_window(doc, target_top_line)
                    window_start = doc.get('window_start_line', 1)
                target_index = f"{max(1, target_top_line - window_start + 1)}.0"
            else:
                target_index = f"{target_top_line}.0"

            widget.mark_set(tk.INSERT, target_index)
            widget.yview(target_index)
        except tk.TclError:
            return False

        self.sync_widget_insert_to_visible_line(widget)
        return True

    def page_up(self, event=None):
        targets = self.get_page_navigation_targets(event)
        if not targets:
            return
        for widget in targets:
            if not self.scroll_widget_page_by_visible_lines(widget, -1):
                continue
            self.set_last_active_editor_widget(widget)
            doc = self.get_navigation_doc_for_widget(widget)
            if doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
                self.schedule_minimap_refresh(doc)
        self.update_status()
        return "break"

    def page_down(self, event=None):
        targets = self.get_page_navigation_targets(event)
        if not targets:
            return
        for widget in targets:
            if not self.scroll_widget_page_by_visible_lines(widget, 1):
                continue
            self.set_last_active_editor_widget(widget)
            doc = self.get_navigation_doc_for_widget(widget)
            if doc:
                self.remember_doc_view_state(doc)
                self.update_line_number_gutter(doc)
                self.schedule_minimap_refresh(doc)
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
        widgets = self.get_find_highlight_widgets()
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

    def refresh_find_highlights(self, query, max_matches=None, allow_short_query=False):
        cleaned_query = str(query or "").strip()
        if not cleaned_query:
            self.clear_find_highlights()
            return False
        if not allow_short_query and len(cleaned_query) < self.live_find_min_chars:
            self.clear_find_highlights()
            return False
        self.clear_find_highlights()
        self.highlight_all_matches(
            cleaned_query,
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

        self.hide_search_history_popup()
        self.add_find_history_entry(query)
        self.update_find_match_summary(query, allow_short_query=True)

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        self.refresh_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_next_in_widget(target_widget, query, wrap=True)
            return "break"

        return self.find_next_across_tabs(query, target_widget)

    def find_previous(self, event=None):
        if self.find_panel_visible:
            query = self.find_entry.get().strip()
        elif self.replace_panel_visible:
            query = self.replace_find_entry.get().strip()
        else:
            return "break"

        if not query:
            return "break"

        self.hide_search_history_popup()
        self.add_find_history_entry(query)
        self.update_find_match_summary(query, allow_short_query=True)

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        self.refresh_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_previous_in_widget(target_widget, query, wrap=True)
            return "break"

        return self.find_previous_across_tabs(query, target_widget)

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

        self.hide_search_history_popup()
        self.add_find_history_entry(query)
        self.update_find_match_summary(query, allow_short_query=True)

        target_widget = self.get_active_search_widget()
        if target_widget is None:
            return "break"

        self.refresh_find_highlights(query, allow_short_query=True)

        compare_widget = self.get_compare_text_widget()
        if target_widget == compare_widget:
            self.search_next_in_widget(target_widget, query, start_index='1.0', wrap=True)
            return "break"

        return self.find_next_across_tabs(query, target_widget, start_from_top=True)

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

        for doc in self.documents.values():
            if doc.get('virtual_mode') or doc.get('preview_mode'):
                continue
            add_widget(doc.get('text'))
        return widgets

    def get_find_highlight_widgets(self):
        widgets = list(self.get_find_target_widgets())
        seen = {str(widget) for widget in widgets}
        compare_widget = self.get_compare_text_widget()
        if compare_widget:
            try:
                if compare_widget.winfo_exists():
                    widget_id = str(compare_widget)
                    if widget_id not in seen:
                        widgets.append(compare_widget)
            except tk.TclError:
                pass
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

    def get_visible_find_query(self):
        try:
            if self.find_panel_visible and self.find_entry.winfo_exists():
                return self.find_entry.get().strip()
            if self.replace_panel_visible and self.replace_find_entry.winfo_exists():
                return self.replace_find_entry.get().strip()
        except tk.TclError:
            return ""
        return ""

    def set_find_match_summary_text(self, text):
        for label_name in ('find_results_label', 'replace_results_label'):
            label = getattr(self, label_name, None)
            if not label:
                continue
            try:
                if label.winfo_exists():
                    label.config(text=text or "")
            except tk.TclError:
                continue

    def count_find_matches(self, query):
        if not query:
            return 0
        total_matches = 0
        for widget in self.get_find_target_widgets():
            total_matches += len(self.find_query_offsets(widget, query, nocase=True))
        return total_matches

    def update_find_match_summary(self, query=None, allow_short_query=False):
        if query is None:
            query = self.get_visible_find_query()
        cleaned_query = str(query or "").strip()
        if not (self.find_panel_visible or self.replace_panel_visible):
            self.set_find_match_summary_text("")
            return 0
        if not cleaned_query or (not allow_short_query and len(cleaned_query) < self.live_find_min_chars):
            self.set_find_match_summary_text("")
            return 0
        match_count = self.count_find_matches(cleaned_query)
        instance_word = self.tr('find.panel.instance_singular', 'instance') if match_count == 1 else self.tr('find.panel.instance_plural', 'instances')
        safe_query = cleaned_query.replace('{', '{{').replace('}', '}}')
        summary_text = self.tr(
            'find.panel.found_summary',
            '| Found: {count} {instance_word} of "{query}"',
            count=match_count,
            instance_word=instance_word,
            query=safe_query
        )
        self.set_find_match_summary_text(summary_text)
        return match_count

    def get_find_in_supported_patterns(self):
        exact_names = set()
        extensions = set()
        all_supported_label = self.tr('filetype.all_supported', 'All Supported')
        all_files_label = self.tr('filetype.all_files', 'All Files')
        for label, pattern in self.get_save_filetypes():
            if label in {all_supported_label, all_files_label}:
                continue
            for token in str(pattern).split():
                normalized = token.strip().lower()
                if normalized.startswith('*.') and len(normalized) > 2:
                    extensions.add(f".{normalized[2:]}")
                elif normalized.startswith('.') and len(normalized) > 1:
                    exact_names.add(normalized)
        return exact_names, extensions

    def get_selected_find_in_directory(self):
        directory = str(getattr(self, 'find_in_selected_directory', '') or '').strip()
        if directory and os.path.isdir(directory):
            return os.path.abspath(directory)
        return ""

    def get_find_in_initial_directory(self):
        selected_directory = self.get_selected_find_in_directory()
        if selected_directory:
            return selected_directory
        current_doc = self.get_current_doc()
        if current_doc:
            file_path = current_doc.get('file_path')
            if file_path and os.path.exists(file_path):
                return os.path.dirname(os.path.abspath(file_path))
        return os.getcwd()

    def prompt_for_find_in_directory(self):
        selected_directory = filedialog.askdirectory(
            parent=self.root,
            title=self.tr('find.in.choose_directory', 'Choose a folder to search'),
            initialdir=self.get_find_in_initial_directory()
        )
        if not selected_directory:
            return ""
        self.find_in_selected_directory = os.path.abspath(selected_directory)
        return self.find_in_selected_directory

    def count_query_matches_in_file(self, file_path, query):
        match_count = 0
        with open(file_path, 'r', encoding='utf-8', errors='replace') as source_file:
            for line in source_file:
                match_count += len(self.find_query_offsets_in_text(line, query, nocase=True))
        return match_count

    def search_query_in_directory(self, directory, query):
        directory = os.path.abspath(directory)
        query_text = str(query or '')
        exact_names, extensions = self.get_find_in_supported_patterns()
        results = []
        scanned_files = 0
        total_matches = 0

        for current_root, dir_names, file_names in os.walk(directory):
            dir_names[:] = [
                name for name in dir_names
                if name.lower() not in {'.git', '__pycache__', '.vs'}
            ]
            for file_name in file_names:
                normalized_name = file_name.lower()
                extension = os.path.splitext(normalized_name)[1]
                if normalized_name not in exact_names and extension not in extensions:
                    continue
                file_path = os.path.join(current_root, file_name)
                scanned_files += 1
                match_count = self.count_query_matches_in_file(file_path, query_text)
                if match_count <= 0:
                    continue
                total_matches += match_count
                results.append({
                    'path': file_path,
                    'relative_path': os.path.relpath(file_path, directory),
                    'match_count': match_count,
                })

        results.sort(key=lambda item: (-item['match_count'], item['relative_path'].lower()))
        return results, scanned_files, total_matches

    def show_find_in_results_dialog(self, directory, query, results, scanned_files, total_matches):
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('find.in.title', 'Find In Files'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color, padx=12, pady=12)
        dialog.minsize(760, 420)

        instance_word = self.tr('find.panel.instance_singular', 'instance') if total_matches == 1 else self.tr('find.panel.instance_plural', 'instances')
        summary = self.tr(
            'find.in.results_summary',
            'Found {match_count} {instance_word} across {file_count} matching file(s).\nDirectory: {directory}\nScanned {scanned_count} supported files.',
            match_count=total_matches,
            instance_word=instance_word,
            file_count=len(results),
            directory=directory,
            scanned_count=scanned_files
        )
        tk.Label(
            dialog,
            text=summary,
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        ).pack(fill='x', pady=(0, 10))

        results_frame = tk.Frame(dialog, bg=self.bg_color)
        results_frame.pack(fill='both', expand=True)
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(
            results_frame,
            columns=('instances', 'file'),
            show='headings',
            selectmode='extended'
        )
        tree.heading('instances', text=self.tr('find.in.column.instances', 'Instances'))
        tree.heading('file', text=self.tr('find.in.column.file', 'File'))
        tree.column('instances', width=90, minwidth=90, anchor='center', stretch=False)
        tree.column('file', width=620, anchor='w', stretch=True)
        tree.grid(row=0, column=0, sticky='nsew')

        scrollbar = ttk.Scrollbar(results_frame, orient='vertical', command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=scrollbar.set)

        row_paths = {}
        for index, item in enumerate(results):
            row_id = f"result_{index}"
            tree.insert('', tk.END, iid=row_id, values=(item['match_count'], item['relative_path']))
            row_paths[row_id] = item['path']

        def close_dialog(event=None):
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()
            return "break"

        def open_selected(event=None):
            selection = tree.selection()
            if not selection:
                messagebox.showinfo(
                    self.tr('find.in.title', 'Find In Files'),
                    self.tr('find.in.select_results', 'Select one or more results to open.'),
                    parent=dialog
                )
                return "break"
            selected_paths = [row_paths[item_id] for item_id in selection if item_id in row_paths]
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()
            for file_path in selected_paths:
                self.open_file_path(file_path)
            return "break"

        button_row = tk.Frame(dialog, bg=self.bg_color)
        button_row.pack(fill='x', pady=(10, 0))
        tk.Button(button_row, text=self.tr('find.in.open_selected', 'Open Selected'), command=open_selected)\
            .pack(side='left')
        tk.Button(button_row, text=self.tr('common.close', 'Close'), command=close_dialog)\
            .pack(side='right')

        if results:
            first_row = 'result_0'
            tree.selection_set(first_row)
            tree.focus(first_row)
        tree.bind('<Double-Button-1>', open_selected)
        dialog.bind('<Return>', open_selected)
        dialog.bind('<Escape>', close_dialog)
        dialog.protocol('WM_DELETE_WINDOW', close_dialog)
        dialog.update_idletasks()
        self.center_window(dialog, self.root)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        tree.focus_set()
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current, self.root))
        dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)

    def start_find_in_directory_search(self, directory, query):
        directory = os.path.abspath(directory)
        safe_query = str(query or '').replace('{', '{{').replace('}', '}}')
        progress_dialog = self.create_toplevel(self.root)
        progress_dialog.title(self.tr('find.in.title', 'Find In Files'))
        progress_dialog.transient(self.root)
        progress_dialog.resizable(False, False)
        progress_dialog.configure(bg=self.bg_color, padx=14, pady=12)

        tk.Label(
            progress_dialog,
            text=self.tr('find.in.searching_prompt', 'Searching:\n{directory}\n\nPlease wait...', directory=directory),
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        ).pack(anchor='w', pady=(0, 10))

        progress_bar = ttk.Progressbar(progress_dialog, orient='horizontal', mode='indeterminate', length=320)
        progress_bar.pack(fill='x')
        progress_bar.start(10)
        progress_dialog.protocol('WM_DELETE_WINDOW', lambda: None)
        progress_dialog.update_idletasks()
        self.center_window(progress_dialog, self.root)
        progress_dialog.lift()
        progress_dialog.attributes('-topmost', True)
        progress_dialog.grab_set()
        progress_dialog.focus_force()
        progress_dialog.after(1, lambda current=progress_dialog: self.center_window_after_show(current, self.root))
        progress_dialog.after(50, lambda: progress_dialog.attributes('-topmost', False) if progress_dialog.winfo_exists() else None)

        result = {
            'done': False,
            'error': None,
            'matches': [],
            'scanned_files': 0,
            'total_matches': 0,
        }

        def worker():
            try:
                matches, scanned_files, total_matches = self.search_query_in_directory(directory, query)
                result['matches'] = matches
                result['scanned_files'] = scanned_files
                result['total_matches'] = total_matches
            except Exception as exc:
                result['error'] = exc
            finally:
                result['done'] = True

        def finish_search():
            if not progress_dialog.winfo_exists():
                return
            if not result['done']:
                progress_dialog.after(120, finish_search)
                return

            progress_bar.stop()
            try:
                progress_dialog.grab_release()
            except tk.TclError:
                pass
            progress_dialog.destroy()

            if result['error'] is not None:
                messagebox.showerror(
                    self.tr('find.in.title', 'Find In Files'),
                    self.tr(
                        'find.in.search_failed',
                        'Notepad-X could not search:\n{directory}\n\n{error_detail}',
                        directory=directory,
                        error_detail=str(result['error']).replace('{', '{{').replace('}', '}}')
                    ),
                    parent=self.root
                )
                return

            matches = list(result.get('matches') or [])
            if not matches:
                messagebox.showinfo(
                    self.tr('find.in.title', 'Find In Files'),
                    self.tr(
                        'find.in.no_matches',
                        'No matching files were found for "{query}" in:\n{directory}',
                        query=safe_query,
                        directory=directory
                    ),
                    parent=self.root
                )
                return

            self.show_find_in_results_dialog(
                directory,
                query,
                matches,
                int(result.get('scanned_files') or 0),
                int(result.get('total_matches') or 0)
            )

        threading.Thread(target=worker, name='NotepadXFindInFiles', daemon=True).start()
        progress_dialog.after(120, finish_search)

    def choose_find_in_directory_and_search(self):
        query = self.find_in_query_var.get().strip()
        if not query:
            messagebox.showinfo(
                self.tr('find.in.title', 'Find In Files'),
                self.tr('find.in.query_required', 'Enter some text in Find In first.'),
                parent=self.root
            )
            return "break"
        selected_directory = self.prompt_for_find_in_directory()
        if not selected_directory:
            return "break"
        self.hide_search_history_popup()
        self.add_find_in_history_entry(query)
        self.start_find_in_directory_search(selected_directory, query)
        return "break"

    def find_in_directory_from_entry(self, event=None):
        query = self.find_in_query_var.get().strip()
        if not query:
            messagebox.showinfo(
                self.tr('find.in.title', 'Find In Files'),
                self.tr('find.in.query_required', 'Enter some text in Find In first.'),
                parent=self.root
            )
            return "break"
        directory = self.get_selected_find_in_directory()
        if not directory:
            directory = self.prompt_for_find_in_directory()
            if not directory:
                return "break"
        self.hide_search_history_popup()
        self.add_find_in_history_entry(query)
        self.start_find_in_directory_search(directory, query)
        return "break"

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

    def build_find_query_pattern(self, query):
        query_text = str(query or '')
        escaped_query = re.escape(query_text)
        if re.fullmatch(r"[\w']+", query_text, flags=re.UNICODE):
            return re.compile(rf"(?<![\w']){escaped_query}(?![\w'])", re.UNICODE)
        return re.compile(escaped_query)

    def find_query_offsets_in_text(self, content, query, start_offset=0, stop_offset=None, max_matches=None, nocase=True):
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

        query_text = str(query or '')
        if not query_text:
            return []

        flags = re.IGNORECASE if nocase else 0
        pattern = self.build_find_query_pattern(query_text)
        if flags:
            pattern = re.compile(pattern.pattern, pattern.flags | flags)

        offsets = []
        search_text = content[search_start:search_stop]
        for match in pattern.finditer(search_text):
            offsets.append(search_start + match.start())
            if max_matches and len(offsets) >= max_matches:
                break
        return offsets

    def find_query_offsets(self, text_widget, query, start_offset=0, stop_offset=None, max_matches=None, nocase=True):
        try:
            if not text_widget or not text_widget.winfo_exists() or not query:
                return []
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return []
        return self.find_query_offsets_in_text(
            content,
            query,
            start_offset=start_offset,
            stop_offset=stop_offset,
            max_matches=max_matches,
            nocase=nocase
        )

    def highlight_all_matches(self, query, max_matches=None, allow_short_query=False):
        widgets = self.get_find_highlight_widgets()
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
        self.refresh_find_highlights(
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

            self.update_find_match_summary(query, allow_short_query=True)
            self.highlight_live_find_matches(query)
        except Exception as exc:
            self.log_exception("live find change", exc)

    def on_find_entry_change(self, event=None):
        try:
            if self.find_panel_visible:
                entry_widget = self.find_entry
                query = entry_widget.get().strip()
            elif self.replace_panel_visible:
                entry_widget = self.replace_find_entry
                query = entry_widget.get().strip()
            else:
                self.hide_search_history_popup()
                return
        except tk.TclError as exc:
            self.log_exception("read live find query", exc)
            return

        self.update_search_history_popup(entry_widget, force_show=True)
        self.cancel_find_change_job()
        self.update_find_match_summary(query, allow_short_query=True)
        if not query:
            self.clear_find_highlights()
            return
        if len(query) < self.live_find_min_chars:
            self.clear_find_highlights()
            return
        try:
            self.find_change_job = self.root.after(30, self.apply_live_find_change)
        except tk.TclError as exc:
            self.log_exception("schedule live find change", exc)

    def on_find_in_entry_change(self, event=None):
        try:
            query = self.find_in_entry.get().strip()
        except tk.TclError as exc:
            self.log_exception("read find in query", exc)
            return

        if not query and not self.find_in_history:
            self.hide_search_history_popup()
            return
        self.update_search_history_popup(self.find_in_entry, force_show=True)

    def replace_all(self):
        query = self.replace_find_entry.get().strip()
        replace_text = self.replace_entry.get()
        if not query:
            return
        self.hide_search_history_popup()
        self.add_find_history_entry(query)
        content = self.text.get('1.0', tk.END)
        pattern = self.build_find_query_pattern(query)
        flags = pattern.flags | re.IGNORECASE
        match_pattern = re.compile(pattern.pattern, flags)
        new_content, replacement_count = match_pattern.subn(replace_text, content)
        self.text.delete('1.0', tk.END)
        self.text.insert('1.0', new_content.rstrip('\n'))
        self.text.edit_modified(True)
        messagebox.showinfo(
            self.tr('replace_all.title', 'Replace All'),
            self.tr('replace_all.completed', 'Replaced {count} occurrence(s).', count=replacement_count),
            parent=self.root
        )
        self.update_status()
        self.update_find_match_summary(query, allow_short_query=True)
        self.highlight_live_find_matches(query)

    def get_hotkey_definitions(self):
        t = self.tr
        return OrderedDict([
            ('open', {'section': t('menu.file', 'File'), 'label': t('menu.file.open', 'Open'), 'default': 'Ctrl+W', 'handler': lambda event=None: self.open_file(event)}),
            ('open_project', {'section': t('menu.file', 'File'), 'label': t('menu.file.open_project', 'Open Project'), 'default': 'Ctrl+Shift+W', 'handler': lambda event=None: self.open_project(event)}),
            ('open_remote', {'section': t('menu.file', 'File'), 'label': 'Open Remote (SSH)', 'default': 'Ctrl+Alt+O', 'handler': lambda event=None: self.open_remote_file_dialog(event)}),
            ('grab_git', {'section': t('menu.file', 'File'), 'label': t('menu.file.grab_git', 'Grab Git'), 'default': 'Ctrl+Shift+G', 'handler': lambda event=None: self.grab_git_project(event)}),
            ('new_tab', {'section': t('menu.file', 'File'), 'label': t('menu.file.new_tab', 'New Tab'), 'default': 'Ctrl+T', 'handler': lambda event=None: self.new_tab(event)}),
            ('close_tab', {'section': t('menu.file', 'File'), 'label': t('menu.file.close_tab', 'Close Tab'), 'default': 'Ctrl+Shift+T', 'handler': lambda event=None: self.close_current_tab(event)}),
            ('save', {'section': t('menu.file', 'File'), 'label': t('menu.file.save', 'Save'), 'default': 'Ctrl+S', 'handler': lambda event=None: self.save(event)}),
            ('save_all', {'section': t('menu.file', 'File'), 'label': t('menu.file.save_all', 'Save All'), 'default': 'Ctrl+Shift+S', 'handler': lambda event=None: self.save_all(event)}),
            ('save_as', {'section': t('menu.file', 'File'), 'label': t('menu.file.save_as', 'Save As'), 'default': 'Ctrl+Shift+Q', 'handler': lambda event=None: self.save_as()}),
            ('save_copy_as', {'section': t('menu.file', 'File'), 'label': t('save.copy_title', 'Save Copy As'), 'default': 'Ctrl+Alt+S', 'handler': lambda event=None: self.save_copy_as()}),
            ('save_and_run', {'section': t('menu.file', 'File'), 'label': t('menu.file.save_and_run', 'Save and Run'), 'default': 'Ctrl+Shift+R', 'handler': lambda event=None: self.save_and_run(event)}),
            ('save_as_encrypted', {'section': t('menu.file', 'File'), 'label': t('menu.file.save_as_encrypted', 'Save As Encrypted'), 'default': 'Ctrl+Shift+E', 'handler': lambda event=None: self.save_encrypted_copy()}),
            ('print', {'section': t('menu.file', 'File'), 'label': t('menu.file.print', 'Print'), 'default': 'Ctrl+P', 'handler': lambda event=None: self.print_file(event)}),
            ('export_notes', {'section': t('menu.file', 'File'), 'label': t('menu.file.export_notes', 'Export Notes'), 'default': 'Ctrl+E', 'handler': lambda event=None: self.export_notes_report()}),
            ('exit', {'section': t('menu.file', 'File'), 'label': t('menu.file.exit', 'Exit'), 'default': 'Ctrl+Shift+X', 'handler': lambda event=None: self.ctrl_shift_x(event)}),
            ('find', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.find', 'Find'), 'default': 'Ctrl+F', 'handler': lambda event=None: self.show_find_panel()}),
            ('find_next', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.find_next', 'Find Next'), 'default': 'F3', 'handler': lambda event=None: self.find_next(event)}),
            ('find_previous', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.find_previous', 'Find Previous'), 'default': 'Shift+F3', 'handler': lambda event=None: self.find_previous(event)}),
            ('replace', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.replace', 'Replace'), 'default': 'Ctrl+R', 'handler': lambda event=None: self.show_replace_panel()}),
            ('command_panel', {'section': t('menu.edit', 'Edit'), 'label': 'Command Panel', 'default': 'Ctrl+Shift+K', 'handler': lambda event=None: self.show_command_panel(event)}),
            ('jump_symbol', {'section': t('menu.edit', 'Edit'), 'label': 'Jump to Symbol', 'default': 'Ctrl+Shift+O', 'handler': lambda event=None: self.show_symbol_navigator(event)}),
            ('project_symbols', {'section': t('menu.edit', 'Edit'), 'label': 'Project Symbols', 'default': 'Ctrl+Alt+P', 'handler': lambda event=None: self.show_symbol_navigator(project_scope=True)}),
            ('toggle_fold', {'section': t('menu.edit', 'Edit'), 'label': 'Toggle Fold', 'default': 'F9', 'handler': lambda event=None: self.toggle_fold_at_cursor(event)}),
            ('collapse_all_folds', {'section': t('menu.edit', 'Edit'), 'label': 'Collapse All Folds', 'default': 'Shift+F9', 'handler': lambda event=None: self.collapse_all_folds(event)}),
            ('expand_all_folds', {'section': t('menu.edit', 'Edit'), 'label': 'Expand All Folds', 'default': 'Ctrl+F9', 'handler': lambda event=None: self.expand_all_folds(event)}),
            ('date', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.date', 'Date'), 'default': 'Ctrl+D', 'handler': lambda event=None: self.insert_date(event)}),
            ('time_date', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.time_date', 'Time/Date'), 'default': 'Ctrl+Shift+D', 'handler': lambda event=None: self.insert_time_date(event)}),
            ('font', {'section': t('menu.edit', 'Edit'), 'label': t('menu.edit.font', 'Font'), 'default': 'Ctrl+Shift+F', 'handler': lambda event=None: self.show_font_dialog(event)}),
            ('fullscreen', {'section': t('menu.view', 'View'), 'label': t('menu.view.full_screen', 'Full Screen'), 'default': 'F11', 'handler': lambda event=None: self.toggle_fullscreen(event)}),
            ('switch_tab', {'section': t('menu.view', 'View'), 'label': t('menu.view.switch_tab', 'Switch Tab'), 'default': 'Ctrl+Tab', 'handler': lambda event=None: self.switch_tab_right(event)}),
            ('currently_editing', {'section': t('menu.view', 'View'), 'label': t('menu.view.currently_editing', 'Currently Editing'), 'default': 'Ctrl+Shift+C', 'handler': lambda event=None: self.toggle_currently_editing_panel(event)}),
            ('cycle_notes', {'section': t('menu.view', 'View'), 'label': t('menu.edit.cycle_notes', 'Cycle Notes'), 'default': 'F4', 'handler': lambda event=None: self.goto_next_note(event)}),
            ('goto_line', {'section': t('menu.view', 'View'), 'label': t('menu.edit.goto_line', 'Go To Line'), 'default': 'Ctrl+G', 'handler': lambda event=None: self.goto_line_dialog(event)}),
            ('top_of_document', {'section': t('menu.view', 'View'), 'label': t('menu.edit.top_of_document', 'Top of Document'), 'default': 'Ctrl+PgUp', 'handler': lambda event=None: self.goto_document_start(event)}),
            ('bottom_of_document', {'section': t('menu.view', 'View'), 'label': t('menu.edit.bottom_of_document', 'Bottom of Document'), 'default': 'Ctrl+PgDn', 'handler': lambda event=None: self.goto_document_end(event)}),
            ('create_theme', {'section': t('menu.view', 'View'), 'label': t('menu.view.create_theme', 'Create Theme'), 'default': 'Ctrl+Alt+T', 'handler': lambda event=None: self.show_create_theme_dialog()}),
            ('compare_tabs', {'section': t('menu.view', 'View'), 'label': t('menu.view.compare_tabs', 'Compare Tabs'), 'default': 'Ctrl+Q', 'handler': lambda event=None: self.show_split_compare()}),
            ('preview_markdown', {'section': t('menu.view', 'View'), 'label': t('menu.view.preview_markdown', 'Preview Markdown'), 'default': 'Ctrl+Shift+P', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.markdown_preview_enabled, self.toggle_markdown_preview)}),
            ('zoom_in', {'section': t('menu.view', 'View'), 'label': 'Zoom In', 'default': 'Ctrl+Plus', 'handler': lambda event=None: self.zoom_in(event)}),
            ('zoom_out', {'section': t('menu.view', 'View'), 'label': 'Zoom Out', 'default': 'Ctrl+Minus', 'handler': lambda event=None: self.zoom_out(event)}),
            ('edit_with_notepadx', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.edit_with_notepadx', 'Edit with Notepad-X'), 'default': 'Ctrl+Alt+X', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.edit_with_shell_enabled, self.toggle_edit_with_shell)}),
            ('sound', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.sound', 'Sound'), 'default': 'Ctrl+Alt+Shift+M', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.sound_enabled, self.toggle_sound)}),
            ('status_bar', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.status_bar', 'Status Bar'), 'default': 'Ctrl+B', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.status_bar_enabled, self.toggle_status_bar)}),
            ('numbered_lines', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.numbered_lines', 'Numbered Lines'), 'default': 'Ctrl+Alt+L', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.numbered_lines_enabled, self.toggle_numbered_lines)}),
            ('autocomplete', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.autocomplete', 'Autocomplete'), 'default': 'Ctrl+Alt+A', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.autocomplete_enabled, self.toggle_autocomplete)}),
            ('spell_check', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.spell_check', 'Spell Check'), 'default': 'F7', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.spell_check_enabled, self.toggle_spell_check)}),
            ('auto_pair', {'section': t('menu.settings', 'Settings'), 'label': 'Auto Pair Brackets/Quotes', 'default': 'Ctrl+Alt+Shift+P', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.auto_pair_enabled, self.save_session)}),
            ('compare_multi_edit', {'section': t('menu.settings', 'Settings'), 'label': 'Compare Multi-Edit', 'default': 'Ctrl+Alt+M', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.compare_multi_edit_enabled, self.save_session)}),
            ('minimap', {'section': t('menu.settings', 'Settings'), 'label': 'Minimap', 'default': 'Ctrl+Alt+I', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.minimap_enabled, self.toggle_minimap)}),
            ('breadcrumbs', {'section': t('menu.settings', 'Settings'), 'label': 'Breadcrumbs', 'default': 'Ctrl+Alt+B', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.breadcrumbs_enabled, self.toggle_breadcrumbs)}),
            ('diagnostics', {'section': t('menu.settings', 'Settings'), 'label': 'Diagnostics', 'default': 'Ctrl+Alt+D', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.diagnostics_enabled, self.toggle_diagnostics)}),
            ('autosave', {'section': t('menu.settings', 'Settings'), 'label': 'Auto Save', 'default': 'Ctrl+Alt+Shift+A', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.autosave_enabled, self.save_session)}),
            ('word_wrap', {'section': t('menu.settings', 'Settings'), 'label': t('menu.view.word_wrap', 'Word Wrap'), 'default': 'Ctrl+Alt+W', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.word_wrap_enabled, self.toggle_word_wrap)}),
            ('sync_page_navigation', {'section': t('menu.settings', 'Settings'), 'label': t('menu.edit.sync_page_navigation', 'Sync PgUp/PgDn in Compare'), 'default': 'Ctrl+Alt+Y', 'handler': lambda event=None: self.toggle_boolean_hotkey(self.sync_page_navigation_enabled, self.save_session)}),
            ('hotkey_settings', {'section': t('menu.settings', 'Settings'), 'label': t('menu.settings.hotkeys', 'Hotkey Settings'), 'default': 'Ctrl+Alt+K', 'handler': lambda event=None: self.show_hotkey_config_dialog()}),
            ('help_contents', {'section': t('menu.help', 'Help'), 'label': t('menu.help.contents', 'Help Contents'), 'default': 'F1', 'handler': lambda event=None: self.show_help_contents()}),
            ('about', {'section': t('menu.help', 'Help'), 'label': t('menu.help.about', 'About Notepad-X'), 'default': 'Shift+F1', 'handler': lambda event=None: self.show_about_dialog()}),
        ])

    def toggle_boolean_hotkey(self, variable, callback):
        variable.set(not bool(variable.get()))
        result = callback()
        return "break" if result is None else result

    def parse_hotkey_shortcut(self, shortcut):
        if shortcut is None:
            return None
        raw = str(shortcut).strip()
        if not raw:
            return None
        raw = raw.replace('++', '+Plus').replace('+-', '+Minus')
        parts = [part.strip() for part in raw.split('+') if part.strip()]
        if not parts:
            return None

        modifier_map = {'ctrl': 'Ctrl', 'control': 'Ctrl', 'alt': 'Alt', 'shift': 'Shift'}
        key_aliases = {
            'pgup': 'PgUp', 'pageup': 'PgUp', 'prior': 'PgUp',
            'pgdn': 'PgDn', 'pagedown': 'PgDn', 'next': 'PgDn',
            'escape': 'Esc', 'esc': 'Esc',
            'enter': 'Enter', 'return': 'Enter',
            'space': 'Space', 'tab': 'Tab',
            'home': 'Home', 'end': 'End',
            'up': 'Up', 'down': 'Down', 'left': 'Left', 'right': 'Right',
            'delete': 'Delete', 'del': 'Delete',
            'backspace': 'Backspace',
            'insert': 'Insert', 'ins': 'Insert',
            'plus': 'Plus', '=': 'Plus', 'equal': 'Plus', 'kp_add': 'Plus',
            'minus': 'Minus', '-': 'Minus', 'subtract': 'Minus', 'kp_subtract': 'Minus',
        }

        modifiers = []
        for token in parts[:-1]:
            normalized = modifier_map.get(token.lower())
            if normalized is None or normalized in modifiers:
                return None
            modifiers.append(normalized)

        key_token = parts[-1]
        key_lower = key_token.lower()
        if re.fullmatch(r'f(?:[1-9]|1[0-2])', key_lower):
            key = key_lower.upper()
        elif len(key_token) == 1 and key_token.isalnum():
            key = key_token.upper()
        else:
            key = key_aliases.get(key_lower)
        if key is None:
            return None

        ordered_modifiers = [name for name in ('Ctrl', 'Alt', 'Shift') if name in modifiers]
        return ordered_modifiers, key

    def normalize_hotkey_shortcut(self, shortcut):
        parsed = self.parse_hotkey_shortcut(shortcut)
        if parsed is None:
            return None
        modifiers, key = parsed
        return '+'.join(list(modifiers) + [key])

    def format_hotkey_display(self, shortcut):
        parsed = self.parse_hotkey_shortcut(shortcut)
        if parsed is None:
            return ''
        modifiers, key = parsed
        display_key = {'Plus': '+', 'Minus': '-', 'PgUp': 'PgUp', 'PgDn': 'PgDn'}.get(key, key)
        return '+'.join(list(modifiers) + [display_key])

    def hotkey_to_tk_sequences(self, shortcut):
        parsed = self.parse_hotkey_shortcut(shortcut)
        if parsed is None:
            return []
        modifiers, key = parsed
        modifier_prefix = ''.join({'Ctrl': 'Control-', 'Alt': 'Alt-', 'Shift': 'Shift-'}[name] for name in modifiers)
        if len(key) == 1 and key.isalpha():
            keysyms = [key.lower(), key.upper()]
        elif len(key) == 1 and key.isdigit():
            keysyms = [key]
        elif key == 'PgUp':
            keysyms = ['Prior']
        elif key == 'PgDn':
            keysyms = ['Next']
        elif key == 'Esc':
            keysyms = ['Escape']
        elif key == 'Enter':
            keysyms = ['Return']
        elif key == 'Space':
            keysyms = ['space']
        elif key == 'Backspace':
            keysyms = ['BackSpace']
        elif key == 'Plus':
            keysyms = ['plus', 'equal', 'KP_Add']
        elif key == 'Minus':
            keysyms = ['minus', 'KP_Subtract']
        else:
            keysyms = [key]

        sequences = []
        seen = set()
        for keysym in keysyms:
            sequence = f"<{modifier_prefix}{keysym}>"
            if sequence not in seen:
                seen.add(sequence)
                sequences.append(sequence)
        return sequences

    def sanitize_hotkey_overrides(self, payload):
        definitions = self.get_hotkey_definitions()
        if not isinstance(payload, dict):
            return {}
        cleaned = {}
        seen_shortcuts = set()
        for action_id in definitions:
            if action_id not in payload:
                continue
            value = payload.get(action_id)
            if value is None or str(value).strip() == '':
                cleaned[action_id] = None
                continue
            normalized = self.normalize_hotkey_shortcut(value)
            if normalized is None or normalized == definitions[action_id].get('default'):
                continue
            if normalized in seen_shortcuts:
                continue
            seen_shortcuts.add(normalized)
            cleaned[action_id] = normalized
        return cleaned

    def get_hotkey_shortcut(self, action_id):
        definition = self.get_hotkey_definitions().get(action_id)
        if definition is None:
            return None
        if action_id in self.hotkey_overrides:
            return self.hotkey_overrides[action_id]
        return definition.get('default')

    def get_hotkey_display(self, action_id, fallback=''):
        shortcut = self.get_hotkey_shortcut(action_id)
        if shortcut is None:
            return ''
        return self.format_hotkey_display(shortcut) or fallback

    def set_hotkey_override(self, action_id, shortcut):
        definitions = self.get_hotkey_definitions()
        if action_id not in definitions:
            return
        if shortcut is None or str(shortcut).strip() == '':
            self.hotkey_overrides[action_id] = None
            return
        normalized = self.normalize_hotkey_shortcut(shortcut)
        if normalized is None:
            return
        if normalized == definitions[action_id].get('default'):
            self.hotkey_overrides.pop(action_id, None)
            return
        self.hotkey_overrides[action_id] = normalized

    def find_hotkey_conflict(self, action_id, shortcut):
        normalized = self.normalize_hotkey_shortcut(shortcut)
        if normalized is None:
            return None
        for other_action_id in self.get_hotkey_definitions():
            if other_action_id == action_id:
                continue
            if self.normalize_hotkey_shortcut(self.get_hotkey_shortcut(other_action_id)) == normalized:
                return other_action_id
        return None

    def invoke_hotkey_action(self, action_id, event=None):
        definition = self.get_hotkey_definitions().get(action_id)
        if definition is None:
            return None
        result = definition['handler'](event)
        return "break" if result is None else result

    def apply_hotkey_bindings(self):
        for sequence in getattr(self, 'hotkey_bound_sequences', set()):
            try:
                self.root.unbind_all(sequence)
            except tk.TclError:
                pass
        self.hotkey_bound_sequences = set()

        seen_sequences = set()
        for action_id in self.get_hotkey_definitions():
            shortcut = self.get_hotkey_shortcut(action_id)
            if shortcut is None:
                continue
            for sequence in self.hotkey_to_tk_sequences(shortcut):
                if sequence in seen_sequences:
                    continue
                seen_sequences.add(sequence)
                try:
                    self.root.bind_all(sequence, lambda event, current=action_id: self.invoke_hotkey_action(current, event))
                    self.hotkey_bound_sequences.add(sequence)
                except tk.TclError:
                    continue

    def refresh_hotkey_configuration(self, rebuild_menu=True, save_session=False):
        self.apply_hotkey_bindings()
        if rebuild_menu and hasattr(self, 'root') and self.root.winfo_exists():
            self.create_menu()
            if self.fullscreen:
                try:
                    self.root.config(menu='')
                except tk.TclError:
                    pass
        if save_session:
            self.save_session()

    def capture_hotkey_from_event(self, event):
        keysym = str(getattr(event, 'keysym', '') or '')
        if not keysym or keysym in {'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'}:
            return None

        state = int(getattr(event, 'state', 0) or 0)
        modifiers = []
        if state & 0x4:
            modifiers.append('Ctrl')
        if state & 0x20000 or state & 0x8:
            modifiers.append('Alt')
        if state & 0x1:
            modifiers.append('Shift')

        key_aliases = {
            'Prior': 'PgUp',
            'Next': 'PgDn',
            'Escape': 'Esc',
            'Return': 'Enter',
            'space': 'Space',
            'Tab': 'Tab',
            'Home': 'Home',
            'End': 'End',
            'Up': 'Up',
            'Down': 'Down',
            'Left': 'Left',
            'Right': 'Right',
            'Delete': 'Delete',
            'BackSpace': 'Backspace',
            'Insert': 'Insert',
            'equal': 'Plus',
            'plus': 'Plus',
            'KP_Add': 'Plus',
            'minus': 'Minus',
            'KP_Subtract': 'Minus',
        }
        if re.fullmatch(r'F(?:[1-9]|1[0-2])', keysym):
            key = keysym
        elif len(keysym) == 1 and keysym.isalnum():
            key = keysym.upper()
        else:
            key = key_aliases.get(keysym)
        if key is None:
            return None
        if not modifiers and not re.fullmatch(r'F(?:[1-9]|1[0-2])', key):
            return None
        if 'Ctrl' not in modifiers and 'Alt' not in modifiers and not re.fullmatch(r'F(?:[1-9]|1[0-2])', key):
            return None
        return '+'.join(modifiers + [key])

    def show_hotkey_config_dialog(self):
        existing = getattr(self, 'hotkey_dialog', None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return "break"
            except tk.TclError:
                pass

        t = self.tr
        dialog = self.create_toplevel(self.root)
        self.hotkey_dialog = dialog
        dialog.title(t('hotkey.dialog.title', 'Hotkey Settings'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color)
        dialog.geometry("860x560")
        dialog.minsize(720, 480)

        def close_dialog(event=None):
            try:
                dialog.destroy()
            except tk.TclError:
                pass
            return "break"

        def clear_dialog_reference(event=None):
            if getattr(event, 'widget', None) is dialog:
                self.hotkey_dialog = None

        dialog.bind('<Destroy>', clear_dialog_reference, add='+')
        dialog.bind('<Escape>', close_dialog)

        outer = tk.Frame(dialog, bg=self.bg_color, padx=14, pady=14)
        outer.pack(fill='both', expand=True)
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=3)
        outer.grid_columnconfigure(1, weight=2)

        tk.Label(
            outer,
            text=t('hotkey.dialog.instructions', 'Select an action, then press the shortcut you want to assign.'),
            bg=self.bg_color,
            fg=self.fg_color,
            anchor='w',
            justify='left'
        ).grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 12))

        list_frame = tk.Frame(outer, bg=self.bg_color)
        list_frame.grid(row=1, column=0, sticky='nsew', padx=(0, 12))
        list_frame.grid_rowconfigure(1, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        tk.Label(
            list_frame,
            text=t('hotkey.dialog.actions', 'Actions'),
            bg=self.bg_color,
            fg=self.fg_color,
            anchor='w'
        ).grid(row=0, column=0, sticky='ew', pady=(0, 6))

        actions_listbox = tk.Listbox(
            list_frame,
            bg=self.text_bg,
            fg=self.text_fg,
            selectbackground=self.select_bg,
            selectforeground='white',
            activestyle='none',
            font=('Consolas', 10),
            exportselection=False
        )
        actions_listbox.grid(row=1, column=0, sticky='nsew')
        actions_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=actions_listbox.yview)
        actions_scrollbar.grid(row=1, column=1, sticky='ns')
        actions_listbox.configure(yscrollcommand=actions_scrollbar.set)

        editor_frame = tk.Frame(outer, bg=self.bg_color)
        editor_frame.grid(row=1, column=1, sticky='nsew')
        editor_frame.grid_columnconfigure(0, weight=1)

        action_name_var = tk.StringVar(value=t('hotkey.dialog.select_action', 'Select an action to edit.'))
        current_var = tk.StringVar(value=t('hotkey.dialog.unassigned', 'Unassigned'))
        capture_var = tk.StringVar(value='')
        status_var = tk.StringVar(value=t('hotkey.dialog.ready', 'Ready.'))
        pending_shortcut = {'value': None}
        definitions = self.get_hotkey_definitions()
        action_ids = list(definitions.keys())

        tk.Label(
            editor_frame,
            textvariable=action_name_var,
            bg=self.bg_color,
            fg='white',
            anchor='w',
            justify='left',
            font=('Segoe UI', 12, 'bold')
        ).grid(row=0, column=0, sticky='ew', pady=(0, 14))

        details_frame = tk.Frame(editor_frame, bg=self.panel_bg, padx=12, pady=12)
        details_frame.grid(row=1, column=0, sticky='ew')
        details_frame.grid_columnconfigure(1, weight=1)

        tk.Label(details_frame, text=t('hotkey.dialog.current', 'Current Shortcut:'), bg=self.panel_bg, fg=self.fg_color, anchor='w').grid(row=0, column=0, sticky='w')
        tk.Label(details_frame, textvariable=current_var, bg=self.panel_bg, fg='white', anchor='w').grid(row=0, column=1, sticky='ew', padx=(10, 0))
        tk.Label(details_frame, text=t('hotkey.dialog.new', 'New Shortcut:'), bg=self.panel_bg, fg=self.fg_color, anchor='w').grid(row=1, column=0, sticky='w', pady=(10, 0))

        capture_entry = tk.Entry(
            details_frame,
            textvariable=capture_var,
            bg=self.text_bg,
            fg='white',
            insertbackground='white',
            relief='flat',
            highlightthickness=1,
            highlightbackground='#3a3a3a',
            highlightcolor=self.cursor_color
        )
        capture_entry.grid(row=1, column=1, sticky='ew', padx=(10, 0), pady=(10, 0))

        button_row = tk.Frame(editor_frame, bg=self.bg_color)
        button_row.grid(row=2, column=0, sticky='ew', pady=(12, 0))
        button_row.grid_columnconfigure(4, weight=1)

        def get_selected_action_id():
            selection = actions_listbox.curselection()
            if not selection:
                return None
            index = int(selection[0])
            return action_ids[index] if 0 <= index < len(action_ids) else None

        def format_action_row(action_id):
            definition = definitions[action_id]
            shortcut_text = self.get_hotkey_display(action_id) or t('hotkey.dialog.unassigned', 'Unassigned')
            return f"{definition['section']} | {definition['label']} | {shortcut_text}"

        def on_action_select(event=None):
            action_id = get_selected_action_id()
            if action_id is None:
                action_name_var.set(t('hotkey.dialog.select_action', 'Select an action to edit.'))
                current_var.set(t('hotkey.dialog.unassigned', 'Unassigned'))
                capture_var.set('')
                pending_shortcut['value'] = None
                return
            definition = definitions[action_id]
            action_name_var.set(f"{definition['section']}  -  {definition['label']}")
            current_display = self.get_hotkey_display(action_id) or t('hotkey.dialog.unassigned', 'Unassigned')
            current_var.set(current_display)
            capture_var.set(current_display if current_display != t('hotkey.dialog.unassigned', 'Unassigned') else '')
            pending_shortcut['value'] = self.get_hotkey_shortcut(action_id)

        def refresh_action_list(selected_action_id=None):
            actions_listbox.delete(0, tk.END)
            for action_id in action_ids:
                actions_listbox.insert(tk.END, format_action_row(action_id))
            if selected_action_id in action_ids:
                index = action_ids.index(selected_action_id)
            elif action_ids:
                index = 0
            else:
                index = None
            if index is not None:
                actions_listbox.selection_clear(0, tk.END)
                actions_listbox.selection_set(index)
                actions_listbox.activate(index)
                actions_listbox.see(index)
            on_action_select()

        def assign_shortcut():
            action_id = get_selected_action_id()
            if action_id is None:
                status_var.set(t('hotkey.dialog.select_action', 'Select an action to edit.'))
                return
            normalized = self.normalize_hotkey_shortcut(pending_shortcut.get('value'))
            if normalized is None:
                status_var.set(t('hotkey.dialog.invalid', 'Use a function key or a shortcut with Ctrl or Alt.'))
                return
            conflict_action_id = self.find_hotkey_conflict(action_id, normalized)
            if conflict_action_id is not None:
                if not messagebox.askyesno(
                    t('hotkey.dialog.conflict_title', 'Shortcut Already In Use'),
                    t(
                        'hotkey.dialog.conflict_message',
                        '{shortcut} is already assigned to {action_label}. Reassign it?',
                        shortcut=self.format_hotkey_display(normalized),
                        action_label=definitions[conflict_action_id]['label']
                    ),
                    parent=dialog
                ):
                    status_var.set(t('hotkey.dialog.ready', 'Ready.'))
                    return
                self.hotkey_overrides[conflict_action_id] = None
            self.set_hotkey_override(action_id, normalized)
            self.refresh_hotkey_configuration(save_session=True)
            refresh_action_list(selected_action_id=action_id)
            status_var.set(
                t(
                    'hotkey.dialog.assigned',
                    'Assigned {shortcut} to {action_label}.',
                    shortcut=self.get_hotkey_display(action_id),
                    action_label=definitions[action_id]['label']
                )
            )

        def clear_shortcut():
            action_id = get_selected_action_id()
            if action_id is None:
                status_var.set(t('hotkey.dialog.select_action', 'Select an action to edit.'))
                return
            self.hotkey_overrides[action_id] = None
            self.refresh_hotkey_configuration(save_session=True)
            refresh_action_list(selected_action_id=action_id)
            status_var.set(t('hotkey.dialog.cleared', 'Cleared {action_label}.', action_label=definitions[action_id]['label']))

        def reset_selected():
            action_id = get_selected_action_id()
            if action_id is None:
                status_var.set(t('hotkey.dialog.select_action', 'Select an action to edit.'))
                return
            self.hotkey_overrides.pop(action_id, None)
            self.refresh_hotkey_configuration(save_session=True)
            refresh_action_list(selected_action_id=action_id)
            status_var.set(t('hotkey.dialog.reset', 'Reset {action_label} to its default shortcut.', action_label=definitions[action_id]['label']))

        def reset_all():
            if not messagebox.askyesno(
                t('hotkey.dialog.reset_all_title', 'Reset All Hotkeys'),
                t('hotkey.dialog.reset_all_message', 'Reset every hotkey to its default value?'),
                parent=dialog
            ):
                return
            self.hotkey_overrides = {}
            self.refresh_hotkey_configuration(save_session=True)
            refresh_action_list(selected_action_id=get_selected_action_id())
            status_var.set(t('hotkey.dialog.ready', 'Ready.'))

        def capture_shortcut(event):
            shortcut = self.capture_hotkey_from_event(event)
            if shortcut is None:
                status_var.set(t('hotkey.dialog.invalid', 'Use a function key or a shortcut with Ctrl or Alt.'))
                return "break"
            pending_shortcut['value'] = shortcut
            capture_var.set(self.format_hotkey_display(shortcut))
            status_var.set(t('hotkey.dialog.captured', 'Captured {shortcut}.', shortcut=self.format_hotkey_display(shortcut)))
            return "break"

        tk.Button(button_row, text=t('hotkey.dialog.assign', 'Assign'), width=12, command=assign_shortcut).grid(row=0, column=0, padx=(0, 8))
        tk.Button(button_row, text=t('hotkey.dialog.clear', 'Clear'), width=12, command=clear_shortcut).grid(row=0, column=1, padx=(0, 8))
        tk.Button(button_row, text=t('hotkey.dialog.reset_action', 'Reset Action'), width=14, command=reset_selected).grid(row=0, column=2, padx=(0, 8))
        tk.Button(button_row, text=t('hotkey.dialog.reset_all', 'Reset All'), width=12, command=reset_all).grid(row=0, column=3, padx=(0, 8))
        tk.Button(button_row, text=t('common.close', 'Close'), width=12, command=close_dialog).grid(row=0, column=5, sticky='e')

        tk.Label(editor_frame, textvariable=status_var, bg=self.bg_color, fg=self.fg_color, anchor='w', justify='left').grid(row=3, column=0, sticky='ew', pady=(14, 0))

        actions_listbox.bind('<<ListboxSelect>>', on_action_select)
        capture_entry.bind('<KeyPress>', capture_shortcut)
        refresh_action_list()
        self.center_window_after_show(dialog, parent=self.root)
        self.root.after_idle(capture_entry.focus_set)
        return "break"

    # ─── Key Bindings ────────────────────────────────────────────
    def bind_keys(self):
        self.root.bind_all('<ButtonRelease-1>', self.maybe_dismiss_transient_ui, add='+')
        self.root.bind_all('<ButtonRelease-3>', self.maybe_dismiss_transient_ui, add='+')
        self.root.bind_all('<Prior>', self.page_up)
        self.root.bind_all('<Next>', self.page_down)
        self.root.bind('<Control-MouseWheel>', self.on_ctrl_mousewheel)
        self.root.bind('<Control-Button-4>', self.on_ctrl_mousewheel)
        self.root.bind('<Control-Button-5>', self.on_ctrl_mousewheel)
        self.apply_hotkey_bindings()

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
        if self.markdown_preview_enabled.get():
            return self.close_markdown_preview()
        if self.compare_active:
            return self.close_compare_panel()
        if self.find_panel_visible or self.replace_panel_visible:
            if self.find_panel_visible:
                self.show_find_panel()
            if self.replace_panel_visible:
                self.show_replace_panel()
            self.focus_last_active_editor()
            return "break"
        if self.command_panel_visible:
            return self.show_command_panel()
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

    def toggle_spell_check(self, event=None):
        if event is not None:
            self.spell_check_enabled.set(not self.spell_check_enabled.get())
        if self.spell_check_enabled.get() and not self.ensure_spellcheck_available(notify=True):
            self.spell_check_enabled.set(False)
            return "break"
        for doc in self.documents.values():
            if self.spell_check_enabled.get():
                self.schedule_spellcheck(doc)
            else:
                self.clear_spellcheck(doc)
        self.save_session()
        return "break"

    def toggle_markdown_preview(self, event=None):
        if event is not None:
            self.markdown_preview_enabled.set(not self.markdown_preview_enabled.get())
        if self.markdown_preview_enabled.get():
            if self.compare_active:
                self.close_compare_panel(persist=False, restore_focus=False)
            self.schedule_markdown_preview_refresh(immediate=True)
        else:
            self.close_markdown_preview(persist=False, restore_focus=False)
        self.save_session()
        return "break"

    def get_edit_with_shell_extensions(self):
        extensions = []
        seen = set()
        all_supported_label = self.tr('filetype.all_supported', 'All Supported')
        all_files_label = self.tr('filetype.all_files', 'All Files')
        for label, pattern in self.get_open_filetypes():
            if label in (all_supported_label, all_files_label):
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
                raise OSError(self.tr('shell_integration.error.unsafe_windows_executable', 'Unsafe executable path for Windows shell integration.'))
            return f'"{executable_path}" "%1"'
        interpreter_path = os.path.abspath(sys.executable)
        script_path = os.path.abspath(__file__)
        if not self.path_looks_safe_for_shell(interpreter_path) or not self.path_looks_safe_for_shell(script_path):
            raise OSError(self.tr('shell_integration.error.unsafe_windows_script', 'Unsafe interpreter or script path for Windows shell integration.'))
        return f'"{interpreter_path}" "{script_path}" "%1"'

    def get_linux_open_command(self):
        if getattr(sys, 'frozen', False):
            executable_path = os.path.abspath(sys.executable)
            if not self.path_looks_safe_for_shell(executable_path):
                raise OSError(self.tr('shell_integration.error.unsafe_linux_executable', 'Unsafe executable path for Linux desktop integration.'))
            return f'"{executable_path}" %F'
        interpreter_path = os.path.abspath(sys.executable)
        script_path = os.path.abspath(__file__)
        if not self.path_looks_safe_for_shell(interpreter_path) or not self.path_looks_safe_for_shell(script_path):
            raise OSError(self.tr('shell_integration.error.unsafe_linux_script', 'Unsafe interpreter or script path for Linux desktop integration.'))
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
                f'Name={self.tr("app.name", "Notepad-X")}',
                f'GenericName={self.tr("shell_integration.generic_name", "Text Editor")}',
                f'Comment={self.tr("shell_integration.desktop_comment", "Edit text files with Notepad-X")}',
                f'Exec={self.get_linux_open_command()}',
                'Terminal=false',
                'Categories=Utility;TextEditor;',
                f'Icon={icon_path}',
                f'MimeType={";".join(self.get_linux_mime_types())};',
                'Actions=EditWithNotepadX;',
                '',
                '[Desktop Action EditWithNotepadX]',
                f'Name={self.tr("menu.view.edit_with_notepadx", "Edit with Notepad-X")}',
                f'Exec={self.get_linux_open_command()}',
                'Terminal=false',
                '',
            ]
            if not self.write_file_atomically(desktop_entry_path, '\n'.join(desktop_entry).rstrip('\n')):
                raise OSError(
                    self.tr(
                        'shell_integration.error.write_linux_desktop_entry',
                        'Could not write Linux desktop entry to {desktop_entry_path}',
                        desktop_entry_path=desktop_entry_path
                    )
                )
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

    def get_registry_string_value(self, root, subkey_path, value_name=''):
        if winreg is None:
            return None
        try:
            with winreg.OpenKey(root, subkey_path, 0, winreg.KEY_READ) as key:
                value, value_type = winreg.QueryValueEx(key, value_name)
        except FileNotFoundError:
            return None
        except OSError:
            return None
        if value_type != winreg.REG_SZ:
            return None
        return str(value)

    def is_windows_edit_with_shell_current(self):
        if not self.is_windows or winreg is None:
            return True
        icon_source = os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else self.icon_path)
        if not self.path_looks_safe_for_shell(icon_source):
            raise OSError(self.tr('shell_integration.error.unsafe_icon_shell', 'Unsafe icon path for Windows shell integration.'))
        open_command = self.get_windows_open_command()
        app_key = rf"Software\Classes\Applications\{self.get_windows_application_registration_name()}"
        if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, app_key, 'ApplicationName') != 'Notepad-X':
            return False
        if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, app_key, 'FriendlyAppName') != 'Notepad-X':
            return False
        if self.get_registry_string_value(
            winreg.HKEY_CURRENT_USER,
            app_key,
            'ApplicationDescription'
        ) != self.tr('shell_integration.app_description', 'Edit supported text and code files with Notepad-X.'):
            return False
        if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, rf"{app_key}\DefaultIcon") != icon_source:
            return False
        if self.get_registry_string_value(
            winreg.HKEY_CURRENT_USER,
            rf"{app_key}\shell\open\command"
        ) != open_command:
            return False
        supported_types_key = rf"{app_key}\SupportedTypes"
        menu_label = self.tr('menu.view.edit_with_notepadx', 'Edit with Notepad-X')
        extensions = self.get_edit_with_shell_extensions()
        for extension in extensions:
            if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, supported_types_key, extension) is None:
                return False
            menu_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\EditWithNotepadX"
            if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, menu_key, 'MUIVerb') != menu_label:
                return False
            if self.get_registry_string_value(winreg.HKEY_CURRENT_USER, menu_key, 'Icon') != icon_source:
                return False
            if self.get_registry_string_value(
                winreg.HKEY_CURRENT_USER,
                rf"{menu_key}\command"
            ) != open_command:
                return False
        return True

    def set_edit_with_shell_app_registration(self, enabled):
        if not self.is_windows or winreg is None:
            return
        app_key = rf"Software\Classes\Applications\{self.get_windows_application_registration_name()}"
        if enabled:
            icon_source = os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else self.icon_path)
            if not self.path_looks_safe_for_shell(icon_source):
                raise OSError(self.tr('shell_integration.error.unsafe_icon_registration', 'Unsafe icon path for Windows application registration.'))
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, app_key) as key:
                winreg.SetValueEx(key, 'ApplicationName', 0, winreg.REG_SZ, 'Notepad-X')
                winreg.SetValueEx(key, 'FriendlyAppName', 0, winreg.REG_SZ, 'Notepad-X')
                winreg.SetValueEx(
                    key,
                    'ApplicationDescription',
                    0,
                    winreg.REG_SZ,
                    self.tr('shell_integration.app_description', 'Edit supported text and code files with Notepad-X.')
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
                raise OSError(self.tr('shell_integration.error.unsafe_icon_shell', 'Unsafe icon path for Windows shell integration.'))
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_key) as key:
                winreg.SetValueEx(key, 'MUIVerb', 0, winreg.REG_SZ, self.tr('menu.view.edit_with_notepadx', 'Edit with Notepad-X'))
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
            did_change = False
            if self.is_windows:
                if winreg is None:
                    return True
                extensions = self.get_edit_with_shell_extensions()
                if enabled:
                    if not self.is_windows_edit_with_shell_current():
                        self.set_edit_with_shell_app_registration(True)
                        for extension in extensions:
                            self.set_edit_with_shell_for_extension(extension, True)
                        did_change = True
                elif self.is_edit_with_shell_registered():
                    self.set_edit_with_shell_app_registration(False)
                    for extension in extensions:
                        self.set_edit_with_shell_for_extension(extension, False)
                    did_change = True
                if did_change:
                    self.notify_windows_shell_change()
            elif self.is_linux:
                self.set_edit_with_shell_linux(enabled)
            return True
        except OSError as exc:
            self.log_exception("sync edit with shell menu", exc)
            if show_errors:
                messagebox.showerror(
                    self.tr('menu.view.edit_with_notepadx', 'Edit with Notepad-X'),
                    self.tr(
                        'shell_integration.update_failed',
                        'Notepad-X could not update the OS shell integration.\n\n{error_detail}',
                        error_detail=exc
                    ),
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
        seen_files = set()
        self.trace_startup(f"open_startup_files raw={list(startup_files or [])}")

        for raw_path in startup_files:
            candidate_path = normalize_startup_path_argument(raw_path)
            if not candidate_path:
                candidate_path = normalize_startup_path_argument(raw_path, base_dir=self.app_dir)
            if not candidate_path:
                continue
            candidate_key = os.path.normcase(candidate_path)
            if candidate_key in seen_files:
                continue
            seen_files.add(candidate_key)
            normalized_files.append(candidate_path)

        self.trace_startup(f"open_startup_files normalized={normalized_files}")

        opened_frames = []
        for file_path in normalized_files:
            if self._shutdown_requested:
                break
            if self.open_file_path(file_path):
                current_doc = self.get_current_doc()
                if current_doc:
                    opened_frames.append(current_doc['frame'])

        if opened_frames and not self._shutdown_requested:
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

    def get_python_module_suggestions(self, doc, prefix):
        suggestions = []
        if not doc or not prefix:
            return suggestions
        doc_path = doc.get('file_path')
        if not doc_path:
            return suggestions
        project_dir = os.path.dirname(os.path.abspath(doc_path))
        try:
            entries = sorted(os.scandir(project_dir), key=lambda entry: entry.name.lower())
        except OSError:
            return suggestions
        for entry in entries:
            if not entry.is_file():
                continue
            stem, extension = os.path.splitext(entry.name)
            if extension.lower() not in {'.py', '.pyw', '.js', '.ts', '.json'}:
                continue
            if not stem or stem == prefix or not stem.lower().startswith(prefix.lower()):
                continue
            suggestions.append(stem)
            if len(suggestions) >= 12:
                break
        return suggestions

    def collect_autocomplete_suggestions(self, doc, prefix):
        syntax_mode = self.get_syntax_mode(doc)
        suggestions = []
        seen = set()

        def add_suggestion(candidate, priority=50):
            candidate_text = str(candidate or '').strip()
            if len(candidate_text) < len(prefix) or candidate_text == prefix:
                return
            if not candidate_text.lower().startswith(prefix.lower()):
                return
            lowered = candidate_text.lower()
            if lowered in seen:
                return
            seen.add(lowered)
            suggestions.append((priority, candidate_text))

        if syntax_mode == 'python':
            for builtin_name in dir(builtins):
                add_suggestion(builtin_name, priority=5)
            for keyword_name in keyword.kwlist:
                add_suggestion(keyword_name, priority=8)

        for keyword_name in self.get_autocomplete_keywords(syntax_mode):
            add_suggestion(keyword_name, priority=10)

        for symbol in self.get_doc_symbols(doc):
            add_suggestion(symbol.get('name'), priority=15)

        for other_doc in self.documents.values():
            if other_doc is doc:
                continue
            for symbol in self.get_doc_symbols(other_doc):
                add_suggestion(symbol.get('name'), priority=20)
            if len(suggestions) >= self.intellisense_max_suggestions:
                break

        if syntax_mode == 'python':
            for module_name in self.get_python_module_suggestions(doc, prefix):
                add_suggestion(module_name, priority=18)

        text = doc.get('text')
        if text:
            try:
                content = text.get('1.0', 'end-1c')
            except tk.TclError:
                content = ''
            for word in re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', content):
                add_suggestion(word, priority=25)
                if len(suggestions) >= self.intellisense_max_suggestions:
                    break

        suggestions.sort(key=lambda item: (item[0], len(item[1]), item[1].lower()))
        return [item[1] for item in suggestions[:self.intellisense_max_suggestions]]

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
        suggestions = self.collect_autocomplete_suggestions(doc, prefix)

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
        previous_prefix = str(getattr(self, 'autocomplete_prefix', '') or '')
        selected_index = 0
        selected_value = None
        if self.autocomplete_popup_visible() and self.autocomplete_listbox:
            selection = self.autocomplete_listbox.curselection()
            if selection:
                selected_index = int(selection[0])
                try:
                    selected_value = self.autocomplete_listbox.get(selected_index)
                except tk.TclError:
                    selected_value = None
        if not self.autocomplete_popup_visible():
            popup = self.create_popup_toplevel(self.root)
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
                height=min(10, len(suggestions))
            )
            listbox.pack(fill='both', expand=True)
            self.autocomplete_popup = popup
            self.autocomplete_listbox = listbox
        else:
            listbox = self.autocomplete_listbox
            listbox.configure(width=max(18, min(42, max(len(word) for word in suggestions) + 2)), height=min(10, len(suggestions)))

        listbox.delete(0, tk.END)
        for suggestion in suggestions:
            listbox.insert(tk.END, suggestion)
        listbox.selection_clear(0, tk.END)
        if prefix == previous_prefix and selected_value in suggestions:
            selected_index = suggestions.index(selected_value)
        elif prefix == previous_prefix and suggestions:
            selected_index = max(0, min(len(suggestions) - 1, selected_index))
        else:
            selected_index = 0
        listbox.selection_set(selected_index)
        listbox.activate(selected_index)
        listbox.see(selected_index)

        self.show_popup_toplevel(popup, text.winfo_rootx() + x, text.winfo_rooty() + y + height + 2)
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
        self.primary_editor_container.grid_rowconfigure(1, weight=1)
        self.primary_editor_container.grid_columnconfigure(0, weight=1)
        self.editor_paned.add(self.primary_editor_container, stretch='always')

        self.breadcrumbs_bar = tk.Frame(self.primary_editor_container, bg='#161b22', height=28)
        self.breadcrumbs_bar.grid(row=0, column=0, sticky='ew')
        self.breadcrumbs_bar.grid_columnconfigure(0, weight=1)
        self.breadcrumbs_label = tk.Label(
            self.breadcrumbs_bar,
            text="",
            bg='#161b22',
            fg='#aab6c6',
            anchor='w',
            padx=10,
            pady=6,
            font=('Segoe UI', 9),
            cursor='arrow'
        )
        self.breadcrumbs_label.grid(row=0, column=0, sticky='ew')
        self.breadcrumbs_label.bind('<Button-1>', lambda e: self.copy_breadcrumb_path(e, target='primary'))

        self.notebook = ttk.Notebook(self.primary_editor_container, style='EditorTabs.TNotebook')
        self.notebook.grid(row=1, column=0, sticky='nsew')
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
        self.notebook.bind('<ButtonPress-1>', self.on_tab_drag_start, add='+')
        self.notebook.bind('<B1-Motion>', self.on_tab_drag_motion, add='+')
        self.notebook.bind('<ButtonRelease-1>', self.on_tab_drag_end, add='+')
        self.notebook.bind('<Button-3>', self.show_tab_context_menu, add='+')

        self.compare_container = tk.Frame(self.editor_paned, bg=self.bg_color)
        self.compare_container.grid_rowconfigure(1, weight=1)
        self.compare_container.grid_columnconfigure(0, weight=1)

        compare_header = tk.Frame(self.compare_container, bg='#161b22')
        compare_header.grid(row=0, column=0, sticky='ew')
        compare_header.grid_columnconfigure(0, weight=1)
        compare_header.grid_rowconfigure(1, weight=0)

        self.compare_title = tk.Label(
            compare_header,
            text="",
            bg='#161b22',
            fg=self.fg_color,
            font=('Segoe UI', 10, 'bold'),
            anchor='w',
            padx=10,
            pady=8,
            cursor='arrow'
        )
        self.compare_title.grid(row=0, column=0, sticky='ew')
        self.compare_title.bind('<Button-1>', lambda e: self.copy_breadcrumb_path(e, target='compare_title'))

        self.compare_breadcrumbs = tk.Label(
            compare_header,
            text="",
            bg='#161b22',
            fg='#8b949e',
            anchor='w',
            padx=10,
            pady=4,
            font=('Segoe UI', 9),
            cursor='arrow'
        )
        self.compare_breadcrumbs.grid(row=1, column=0, sticky='ew')
        self.compare_breadcrumbs.bind('<Button-1>', lambda e: self.copy_breadcrumb_path(e, target='compare'))

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

        self.compare_minimap = tk.Canvas(
            compare_body,
            width=self.minimap_width,
            bg='#11161d',
            highlightthickness=0,
            borderwidth=0,
            relief='flat',
            cursor='hand2'
        )
        self.compare_minimap.grid(row=0, column=2, sticky='ns')
        if not self.minimap_enabled.get():
            self.compare_minimap.grid_remove()

        compare_v_scroll = ttk.Scrollbar(compare_body, orient='vertical', command=self.on_compare_vertical_scroll)
        compare_v_scroll.grid(row=0, column=3, sticky='ns')
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
        self.compare_text.bind('<Control-x>', self.cut)
        self.compare_text.bind('<Control-X>', self.cut)
        self.compare_text.bind('<Control-v>', self.paste)
        self.compare_text.bind('<Control-V>', self.paste)
        self.compare_text.bind('<Control-z>', self.undo)
        self.compare_text.bind('<Control-Z>', self.undo)
        self.compare_text.bind('<Control-Shift-z>', self.redo)
        self.compare_text.bind('<Control-Shift-Z>', self.redo)
        self.compare_text.bind('<Prior>', self.page_up)
        self.compare_text.bind('<Next>', self.page_down)
        self.compare_text.bind('<FocusIn>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<Enter>', self.remember_hovered_editor, add='+')
        self.compare_text.bind('<Motion>', self.remember_hovered_editor, add='+')
        self.compare_text.bind('<Enter>', self.handle_diagnostic_hover, add='+')
        self.compare_text.bind('<Motion>', self.handle_diagnostic_hover, add='+')
        self.compare_text.bind('<Leave>', self.hide_diagnostic_tooltip, add='+')
        self.compare_text.bind('<Button-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<ButtonRelease-1>', self.remember_compare_focus, add='+')
        self.compare_text.bind('<Button-3>', self.show_compare_context_menu)
        self.compare_text.bind('<<Modified>>', self.on_compare_modified)

        self.compare_view = {
            'frame': self.compare_container,
            'text': self.compare_text,
            'line_numbers': self.compare_line_numbers,
            'minimap': self.compare_minimap,
            'notes': {},
            'percolator': None,
            'colorizer': None,
            'theme_effect_job': None,
            'large_file_mode': False,
            'preview_mode': False,
            'virtual_mode': False,
            'line_starts': None,
            'file_size_bytes': 0,
            'syntax_job': None,
            'syntax_mode': None,
            'syntax_override': None,
            'file_path': None,
            'window_start_line': 1,
            'window_end_line': 1,
            'total_file_lines': 1,
            'last_virtual_line': 1,
            'last_virtual_col': 0,
            'pending_virtual_target_line': None,
            'suspend_modified_events': False,
            'fold_ranges': [],
            'fold_ranges_dirty': True,
            'folded_tags': set(),
            'collapsed_fold_regions': {},
            'collapsed_fold_regions_dirty': False,
            'gutter_fold_hitboxes': [],
            'diagnostic_job': None,
            'diagnostics': [],
            'minimap_job': None,
            'minimap_model': None,
            'minimap_model_dirty': True,
            'minimap_progressive_state': None,
            'symbol_cache': None,
            'symbol_cache_signature': None,
            'mirror_guard': False,
            'display_name': None,
            'is_remote': False,
            'remote_spec': None,
            'remote_host': None,
            'remote_path': None,
            'remote_shadow_path': None,
            'virtual_editable': False,
            'virtual_doc_dirty': False,
            'virtual_piece_table': [],
            'virtual_add_buffer_path': None,
            'virtual_add_buffer_size': 0,
            'virtual_revision': 0,
            'virtual_source_path': None,
            'virtual_window_start_byte': 0,
            'virtual_window_end_byte': 0,
            'virtual_hot_chunk_cache': OrderedDict(),
            'virtual_cold_chunk_cache': OrderedDict(),
        }
        self.compare_minimap.bind('<Button-1>', lambda e: self.on_minimap_click(e, self.compare_view))
        self.compare_minimap.bind('<B1-Motion>', lambda e: self.on_minimap_click(e, self.compare_view))
        self.compare_line_numbers.bind('<Button-1>', lambda e: self.handle_gutter_click(e, target_doc=self.compare_view))
        self.compare_line_numbers.bind('<Motion>', lambda e: self.update_gutter_cursor(e, target_doc=self.compare_view))
        self.compare_line_numbers.bind('<Leave>', self.clear_gutter_cursor)
        self.compare_text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        self.compare_text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')
        self.compare_text.tag_config(self.bracket_match_tag, background='#2f81f7', foreground='white')
        self.compare_text.tag_config('diagnostic_error', background='#51202a', foreground='#ffb3ba')
        self.compare_text.tag_config('diagnostic_warning', background='#4b3a14', foreground='#ffd479')
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
        if self.compare_active or self.markdown_preview_enabled.get():
            self.set_compare_sash_position()

    def create_tab(self, file_path=None, content="", select=True):
        tab_frame = tk.Frame(self.notebook, bg=self.bg_color)
        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(1, weight=1)
        tab_frame.grid_columnconfigure(2, weight=0)
        tab_frame.grid_columnconfigure(3, weight=0)

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

        minimap = tk.Canvas(
            tab_frame,
            width=self.minimap_width,
            bg='#11161d',
            highlightthickness=0,
            borderwidth=0,
            relief='flat',
            cursor='hand2'
        )
        minimap.grid(row=0, column=2, sticky='ns')
        minimap.bind('<Button-1>', lambda e, frame=tab_frame: self.on_minimap_click(e, self.documents.get(str(frame))))
        minimap.bind('<B1-Motion>', lambda e, frame=tab_frame: self.on_minimap_click(e, self.documents.get(str(frame))))
        if not self.minimap_enabled.get():
            minimap.grid_remove()

        v_scroll = ttk.Scrollbar(tab_frame, orient='vertical')
        v_scroll.grid(row=0, column=3, sticky='ns')
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
        text.bind('<Control-x>', self.cut)
        text.bind('<Control-X>', self.cut)
        text.bind('<Control-v>', self.paste)
        text.bind('<Control-V>', self.paste)
        text.bind('<Control-z>', self.undo)
        text.bind('<Control-Z>', self.undo)
        text.bind('<Control-Shift-z>', self.redo)
        text.bind('<Control-Shift-Z>', self.redo)
        text.bind('<Prior>', self.page_up)
        text.bind('<Next>', self.page_down)
        text.bind('<FocusIn>', lambda e, frame=tab_frame: self.remember_doc_focus(frame), add='+')
        text.bind('<Enter>', self.remember_hovered_editor, add='+')
        text.bind('<Motion>', self.remember_hovered_editor, add='+')
        text.bind('<Enter>', self.handle_diagnostic_hover, add='+')
        text.bind('<Motion>', self.handle_diagnostic_hover, add='+')
        text.bind('<Leave>', self.hide_diagnostic_tooltip, add='+')
        text.bind('<KeyPress>', lambda e, frame=tab_frame: self.handle_text_keypress(e, frame))
        text.bind('<ButtonRelease-1>', lambda e, frame=tab_frame: self.on_text_click_release(e, frame), add='+')
        text.bind('<Button-3>', lambda e, frame=tab_frame: self.show_text_context_menu(e, frame))

        text.bind('<<Modified>>', lambda e, frame=tab_frame: self.on_text_modified(frame))
        text.tag_config(self.find_matches_tag, background=self.match_bg, foreground='black')
        text.tag_config(self.find_current_tag, background='#ff8c42', foreground='black')
        text.tag_config(self.bracket_match_tag, background='#2f81f7', foreground='white')
        spellcheck_fg, spellcheck_bg = self.get_spellcheck_tag_colors()
        text.tag_config(self.spellcheck_tag, underline=1, foreground=spellcheck_fg, background=spellcheck_bg)
        self.raise_find_tags(text)

        if content:
            text.insert('1.0', content)
        text.edit_modified(False)

        self.documents[str(tab_frame)] = {
            'frame': tab_frame,
            'text': text,
            'line_numbers': line_numbers,
            'minimap': minimap,
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
            'loading_file': False,
            'last_note_cycle_tag': None,
            'theme_effect_job': None,
            'syntax_job': None,
            'spellcheck_job': None,
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
            'background_load_token': None,
            'background_index_future': None,
            'background_index_token': None,
            'background_index_active': False,
            'background_open_new_tab': False,
            'background_bytes_loaded': 0,
            'background_bytes_total': 0,
            'background_lines_loaded': 1,
            'last_virtual_line': 1,
            'last_virtual_col': 0,
            'pending_virtual_target_line': None,
            'pending_insert_content': None,
            'pending_insert_offset': 0,
            'pending_insert_batch_count': 0,
            'fold_ranges': [],
            'fold_ranges_dirty': True,
            'folded_tags': set(),
            'collapsed_fold_regions': {},
            'collapsed_fold_regions_dirty': False,
            'gutter_fold_hitboxes': [],
            'diagnostic_job': None,
            'diagnostics': [],
            'minimap_job': None,
            'minimap_model': None,
            'minimap_model_dirty': True,
            'minimap_progressive_state': None,
            'autosave_job': None,
            'symbol_cache': None,
            'symbol_cache_signature': None,
            'mirror_guard': False,
            'is_remote': False,
            'remote_spec': None,
            'remote_host': None,
            'remote_path': None,
            'remote_shadow_path': None,
            'virtual_editable': False,
            'virtual_doc_dirty': False,
            'virtual_piece_table': [],
            'virtual_add_buffer_path': None,
            'virtual_add_buffer_size': 0,
            'virtual_revision': 0,
            'virtual_source_path': None,
            'virtual_window_start_byte': 0,
            'virtual_window_end_byte': 0,
            'virtual_hot_chunk_cache': OrderedDict(),
            'virtual_cold_chunk_cache': OrderedDict(),
            'display_name': None,
            'load_progress_dialog': None,
            'load_progress_bar': None,
            'load_progress_status_label': None,
            'load_progress_detail_label': None,
        }
        self.apply_syntax_tag_colors(text)
        text.tag_config('diagnostic_error', background='#51202a', foreground='#ffb3ba')
        text.tag_config('diagnostic_warning', background='#4b3a14', foreground='#ffd479')
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
            if doc.get('virtual_mode'):
                row_text, col_text = doc['last_insert_index'].split('.')
                local_row = int(row_text)
                local_col = max(0, int(col_text))
                total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
                doc['last_virtual_line'] = max(
                    1,
                    min(total_lines, int(doc.get('window_start_line', 1) or 1) + local_row - 1)
                )
                doc['last_virtual_col'] = local_col
        except tk.TclError:
            pass
        except (TypeError, ValueError):
            doc['last_virtual_line'] = max(1, int(doc.get('window_start_line', 1) or 1))
            doc['last_virtual_col'] = 0
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
        if doc.get('virtual_mode'):
            if not self.is_virtual_index_ready(doc):
                pending_line = int(doc.get('last_virtual_line', 1) or 1)
                doc['pending_virtual_target_line'] = max(1, pending_line)
                return
            global_line = max(1, min(
                int(doc.get('total_file_lines', 1) or 1),
                int(doc.get('last_virtual_line', doc.get('window_start_line', 1)) or doc.get('window_start_line', 1))
            ))
            global_col = max(0, int(doc.get('last_virtual_col', 0) or 0))
            self.ensure_virtual_line_visible(doc, global_line)
            text = doc.get('text')
            if not text or not text.winfo_exists():
                return
            local_line = max(1, min(
                max(1, int(doc.get('window_end_line', 1) or 1) - int(doc.get('window_start_line', 1) or 1) + 1),
                global_line - int(doc.get('window_start_line', 1) or 1) + 1
            ))
            try:
                line_length = len(text.get(f"{local_line}.0", f"{local_line}.end"))
                text.mark_set(tk.INSERT, f"{local_line}.{min(global_col, line_length)}")
            except tk.TclError:
                text.mark_set(tk.INSERT, '1.0')
        else:
            insert_index = doc.get('last_insert_index') or '1.0'
            try:
                text.mark_set(tk.INSERT, insert_index)
            except tk.TclError:
                text.mark_set(tk.INSERT, '1.0')

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

    def infer_syntax_mode_from_path(self, file_path, content=""):
        file_name = os.path.basename(str(file_path or '')).lower()
        extension = os.path.splitext(file_name)[1].lower()
        extension_modes = {
            '.py': 'python',
            '.pyw': 'python',
            '.rs': 'rust',
            '.c': 'c',
            '.h': 'cpp',
            '.cpp': 'cpp',
            '.cc': 'cpp',
            '.cxx': 'cpp',
            '.hpp': 'cpp',
            '.hh': 'cpp',
            '.hxx': 'cpp',
            '.java': 'java',
            '.js': 'javascript',
            '.jsx': 'javascript',
            '.ts': 'javascript',
            '.tsx': 'javascript',
            '.html': 'html',
            '.htm': 'html',
            '.php': 'php',
            '.xml': 'xml',
            '.json': 'json',
            '.sql': 'sql',
            '.css': 'css',
            '.md': 'markdown',
            '.yml': 'yaml',
            '.yaml': 'yaml',
            '.sh': 'bash',
            '.bat': 'batch',
            '.cmd': 'batch',
        }
        if file_name in {'makefile', 'gnumakefile'}:
            return 'makefile'
        return extension_modes.get(extension) or self.detect_syntax_mode_from_content(content)

    def build_doc_symbol_cache_signature(self, content):
        text = str(content or '')
        if len(text) <= 5000:
            return hashlib.sha1(text.encode('utf-8', errors='replace')).hexdigest()
        preview = text[:2500] + '\n...\n' + text[-2500:]
        return f"{len(text)}:{hashlib.sha1(preview.encode('utf-8', errors='replace')).hexdigest()}"

    def extract_symbols_from_content(self, content, syntax_mode=None, file_path=None):
        text = str(content or '')
        mode = syntax_mode or self.infer_syntax_mode_from_path(file_path, text)
        symbols = []
        seen = set()

        def add_symbol(name, line_number, kind='symbol'):
            clean_name = str(name or '').strip()
            if not clean_name:
                return
            try:
                clean_line = max(1, int(line_number))
            except (TypeError, ValueError):
                clean_line = 1
            key = (clean_name, clean_line, kind)
            if key in seen:
                return
            seen.add(key)
            symbols.append({'name': clean_name, 'line': clean_line, 'kind': kind, 'file_path': file_path})

        if mode == 'python':
            try:
                tree = ast.parse(text or '\n')
            except SyntaxError:
                tree = None
            if tree is not None:
                for node in ast.walk(tree):
                    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        add_symbol(node.name, getattr(node, 'lineno', 1), 'function')
                    elif isinstance(node, ast.ClassDef):
                        add_symbol(node.name, getattr(node, 'lineno', 1), 'class')
                    elif isinstance(node, ast.Import):
                        for alias in node.names:
                            add_symbol(alias.asname or alias.name.split('.')[0], getattr(node, 'lineno', 1), 'import')
                    elif isinstance(node, ast.ImportFrom):
                        for alias in node.names:
                            add_symbol(alias.asname or alias.name, getattr(node, 'lineno', 1), 'import')
                    elif isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name):
                                add_symbol(target.id, getattr(target, 'lineno', 1), 'variable')
            else:
                for match in re.finditer(r'^\s*(def|class)\s+([A-Za-z_]\w*)', text, re.MULTILINE):
                    add_symbol(match.group(2), text.count('\n', 0, match.start()) + 1, match.group(1))
        elif mode == 'markdown':
            for line_number, line in enumerate(text.splitlines(), start=1):
                heading_match = re.match(r'^\s{0,3}(#{1,6})\s+(.+?)\s*$', line)
                if heading_match:
                    add_symbol(heading_match.group(2), line_number, f'h{len(heading_match.group(1))}')
        else:
            patterns = (
                ('class', r'^\s*(?:export\s+)?class\s+([A-Za-z_]\w*)'),
                ('function', r'^\s*(?:export\s+)?(?:async\s+)?function\s+([A-Za-z_]\w*)'),
                ('function', r'^\s*(?:def|fn)\s+([A-Za-z_]\w*)'),
                ('method', r'^\s*([A-Za-z_]\w*)\s*\([^)]*\)\s*\{'),
                ('variable', r'^\s*(?:const|let|var|final|static|pub|private|protected)\s+([A-Za-z_]\w*)'),
                ('type', r'^\s*(?:struct|enum|interface|trait)\s+([A-Za-z_]\w*)'),
                ('import', r'^\s*(?:from|import|use|include)\s+([A-Za-z_][\w./-]*)'),
            )
            lines = text.splitlines()
            for line_number, line in enumerate(lines, start=1):
                for kind, pattern in patterns:
                    match = re.search(pattern, line)
                    if match:
                        add_symbol(match.group(1), line_number, kind)
                        break

        symbols.sort(key=lambda item: (item.get('line', 0), item.get('name', '').lower()))
        return symbols

    def get_doc_symbols(self, doc):
        if not doc:
            return []
        text_widget = doc.get('text')
        if not text_widget:
            return []
        try:
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return []
        signature = self.build_doc_symbol_cache_signature(content)
        if doc.get('symbol_cache_signature') == signature:
            return list(doc.get('symbol_cache') or [])
        symbols = self.extract_symbols_from_content(content, self.get_syntax_mode(doc), self.get_doc_display_path(doc))
        doc['symbol_cache_signature'] = signature
        doc['symbol_cache'] = list(symbols)
        return symbols

    def get_symbol_at_line(self, doc, line_number):
        symbols = self.get_doc_symbols(doc)
        active_symbol = None
        for symbol in symbols:
            if int(symbol.get('line', 1)) <= int(line_number):
                active_symbol = symbol
            else:
                break
        return active_symbol

    def get_symbols_for_scope(self, doc, project_scope=False):
        if not doc:
            return []
        symbols = []
        seen = set()

        def add_scoped_symbol(symbol, source_path=None):
            item = dict(symbol or {})
            item['source_path'] = source_path or doc.get('file_path')
            item['source_display'] = os.path.basename(source_path) if source_path else self.get_doc_name(doc['frame'])
            key = (item.get('source_path'), item.get('line'), item.get('name'))
            if key in seen:
                return
            seen.add(key)
            symbols.append(item)

        for symbol in self.get_doc_symbols(doc):
            add_scoped_symbol(symbol, doc.get('file_path'))

        if project_scope and doc.get('file_path'):
            for candidate_path in self.get_project_source_files(doc['file_path'])[:self.intellisense_max_project_files]:
                if candidate_path == doc.get('file_path'):
                    continue
                try:
                    with open(candidate_path, 'r', encoding='utf-8', errors='replace') as source_file:
                        content = source_file.read()
                except OSError:
                    continue
                mode = self.infer_syntax_mode_from_path(candidate_path, content)
                for symbol in self.extract_symbols_from_content(content, mode, candidate_path):
                    add_scoped_symbol(symbol, candidate_path)

        symbols.sort(key=lambda item: (
            os.path.basename(str(item.get('source_path') or '')).lower(),
            int(item.get('line', 1)),
            str(item.get('name') or '').lower()
        ))
        return symbols

    def goto_symbol_entry(self, symbol):
        if not symbol:
            return
        source_path = symbol.get('source_path')
        current_doc = self.get_current_doc()
        target_doc = current_doc
        if source_path:
            for doc in self.documents.values():
                if doc.get('file_path') and os.path.abspath(doc['file_path']) == os.path.abspath(source_path):
                    target_doc = doc
                    break
            else:
                if not self.open_file_path(source_path):
                    return
                target_doc = self.get_current_doc()
        if not target_doc:
            return
        frame = target_doc.get('frame')
        if frame:
            self.notebook.select(frame)
            self.set_active_document(frame)
        text_widget = target_doc.get('text')
        if not text_widget:
            return
        try:
            line_number = max(1, int(symbol.get('line', 1)))
            text_widget.mark_set(tk.INSERT, f'{line_number}.0')
            text_widget.tag_remove('sel', '1.0', tk.END)
            text_widget.see(f'{line_number}.0')
            text_widget.focus_set()
        except (tk.TclError, ValueError):
            return
        self.set_last_active_editor_widget(text_widget)
        self.update_breadcrumbs()
        self.update_status()

    def show_symbol_navigator(self, event=None, project_scope=False):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        symbols = self.get_symbols_for_scope(doc, project_scope=project_scope)
        if not symbols:
            messagebox.showinfo(
                self.tr('app.name', 'Notepad-X'),
                'No symbols were found for this scope.',
                parent=self.root
            )
            return "break"

        dialog = self.create_toplevel(self.root)
        dialog.title('Project Symbols' if project_scope else 'Symbols')
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color)
        dialog.geometry('720x480')

        container = tk.Frame(dialog, bg=self.bg_color)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        filter_entry = tk.Entry(container)
        filter_entry.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        listbox = tk.Listbox(
            container,
            bg=self.text_bg,
            fg=self.text_fg,
            selectbackground='#264f78',
            selectforeground='white',
            font=('Consolas', 10),
            activestyle='none'
        )
        listbox.grid(row=1, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=listbox.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        listbox.configure(yscrollcommand=scrollbar.set)

        visible_symbols = list(symbols)

        def refresh_list(*_args):
            query = filter_entry.get().strip().lower()
            visible_symbols.clear()
            listbox.delete(0, tk.END)
            for symbol in symbols:
                haystack = f"{symbol.get('name', '')} {symbol.get('kind', '')} {symbol.get('source_display', '')}".lower()
                if query and query not in haystack:
                    continue
                visible_symbols.append(symbol)
                display_name = symbol.get('name', '')
                display_source = symbol.get('source_display', '')
                display_line = symbol.get('line', 1)
                listbox.insert(tk.END, f"{display_name:<36} {display_source:<20} L{display_line}")
            if visible_symbols:
                listbox.selection_set(0)
                listbox.activate(0)

        def open_selected(event=None):
            selection = listbox.curselection()
            if not selection:
                return "break"
            symbol = visible_symbols[selection[0]]
            dialog.destroy()
            self.goto_symbol_entry(symbol)
            return "break"

        filter_entry.bind('<KeyRelease>', refresh_list)
        filter_entry.bind('<Return>', open_selected)
        listbox.bind('<Double-Button-1>', open_selected)
        dialog.bind('<Escape>', lambda e: dialog.destroy())

        refresh_list()
        filter_entry.focus_set()
        self.center_window(dialog, self.root)
        dialog.after(1, lambda current=dialog: self.center_window_after_show(current, self.root))
        return "break"

    def get_indent_fold_regions(self, lines):
        regions = []
        stack = []

        def line_indent(line_text):
            expanded = line_text.expandtabs(4)
            return len(expanded) - len(expanded.lstrip(' '))

        for line_number, raw_line in enumerate(lines, start=1):
            stripped = raw_line.strip()
            if not stripped:
                continue
            indent = line_indent(raw_line)
            while stack and indent <= stack[-1]['indent']:
                start_info = stack.pop()
                end_line = line_number - 1
                if end_line > start_info['line']:
                    regions.append({'start': start_info['line'], 'end': end_line})
            next_indent_candidate = None
            for future_line in lines[line_number:]:
                future_stripped = future_line.strip()
                if not future_stripped:
                    continue
                next_indent_candidate = line_indent(future_line)
                break
            if next_indent_candidate is not None and next_indent_candidate > indent:
                stack.append({'line': line_number, 'indent': indent})
        last_line = len(lines)
        while stack:
            start_info = stack.pop()
            if last_line > start_info['line']:
                regions.append({'start': start_info['line'], 'end': last_line})
        return regions

    def get_brace_fold_regions(self, lines):
        regions = []
        stack = []
        for line_number, raw_line in enumerate(lines, start=1):
            in_string = False
            string_char = ''
            escape = False
            for char in raw_line:
                if escape:
                    escape = False
                    continue
                if in_string:
                    if char == '\\':
                        escape = True
                    elif char == string_char:
                        in_string = False
                    continue
                if char in {'"', "'"}:
                    in_string = True
                    string_char = char
                    continue
                if char == '{':
                    stack.append(line_number)
                elif char == '}' and stack:
                    start_line = stack.pop()
                    if line_number > start_line:
                        regions.append({'start': start_line, 'end': line_number})
        return regions

    def get_markdown_fold_regions(self, lines):
        regions = []
        heading_stack = []
        for line_number, line in enumerate(lines, start=1):
            match = re.match(r'^\s{0,3}(#{1,6})\s+.+$', line)
            if not match:
                continue
            level = len(match.group(1))
            while heading_stack and level <= heading_stack[-1]['level']:
                previous = heading_stack.pop()
                end_line = line_number - 1
                if end_line > previous['line']:
                    regions.append({'start': previous['line'], 'end': end_line})
            heading_stack.append({'line': line_number, 'level': level})
        last_line = len(lines)
        while heading_stack:
            previous = heading_stack.pop()
            if last_line > previous['line']:
                regions.append({'start': previous['line'], 'end': last_line})
        return regions

    def get_fold_regions(self, doc):
        if not doc or doc.get('preview_mode') or doc.get('virtual_mode') or doc.get('large_file_mode'):
            return []
        text_widget = doc.get('text')
        if not text_widget:
            return []
        try:
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return []
        lines = content.splitlines()
        if len(lines) <= 1:
            return []
        mode = self.get_syntax_mode(doc)
        if mode == 'markdown':
            regions = self.get_markdown_fold_regions(lines)
        elif mode in {'c', 'cpp', 'csharp', 'java', 'javascript', 'rust', 'php'}:
            regions = self.get_brace_fold_regions(lines)
            regions.extend(self.get_indent_fold_regions(lines))
        else:
            regions = self.get_indent_fold_regions(lines)
        deduped = {}
        for region in regions:
            start_line = int(region.get('start', 1))
            end_line = int(region.get('end', start_line))
            if end_line <= start_line:
                continue
            deduped[(start_line, end_line)] = {'start': start_line, 'end': end_line}
        return sorted(deduped.values(), key=lambda item: (item['start'], item['end']))

    def invalidate_fold_regions(self, doc):
        if not doc:
            return
        if doc.get('large_file_mode'):
            doc['fold_ranges'] = []
            doc['fold_ranges_dirty'] = False
            return
        doc['fold_ranges_dirty'] = True
        if doc.get('preview_mode') or doc.get('virtual_mode'):
            doc['fold_ranges'] = []

    def invalidate_collapsed_fold_regions(self, doc, clear_cache=False):
        if not doc:
            return
        if clear_cache:
            doc['collapsed_fold_regions'] = {}
        doc['collapsed_fold_regions_dirty'] = True

    def get_cached_fold_regions(self, doc, force=False):
        if not doc:
            return []
        if doc.get('preview_mode') or doc.get('virtual_mode'):
            doc['fold_ranges'] = []
            doc['fold_ranges_dirty'] = False
            return []
        if force or doc.get('fold_ranges_dirty', True):
            doc['fold_ranges'] = list(self.get_fold_regions(doc))
            doc['fold_ranges_dirty'] = False
        return list(doc.get('fold_ranges') or [])

    def make_fold_tag_name(self, start_line, end_line):
        return f"fold_{start_line}_{end_line}"

    def get_collapsed_fold_regions(self, doc, force=False):
        if not doc:
            return {}
        collapsed = doc.setdefault('collapsed_fold_regions', {})
        text_widget = doc.get('text')
        if not text_widget:
            return collapsed
        if not force and not doc.get('collapsed_fold_regions_dirty', False):
            return collapsed
        collapsed = {}
        for tag_name in list(doc.get('folded_tags') or []):
            try:
                ranges = text_widget.tag_ranges(tag_name)
            except tk.TclError:
                continue
            if len(ranges) < 2:
                continue
            try:
                start_line = max(1, int(str(ranges[0]).split('.')[0]) - 1)
                end_index = text_widget.index(f'{ranges[1]} -1c')
                end_line = int(str(end_index).split('.')[0])
            except (tk.TclError, ValueError):
                continue
            collapsed[start_line] = {
                'start': start_line,
                'end': max(start_line + 1, end_line),
                'tag_name': tag_name,
            }
        doc['collapsed_fold_regions'] = collapsed
        doc['collapsed_fold_regions_dirty'] = False
        return collapsed

    def get_top_level_fold_regions(self, regions):
        normalized = sorted(
            (
                {
                    'start': int(region.get('start', 1) or 1),
                    'end': int(region.get('end', int(region.get('start', 1) or 1)) or int(region.get('start', 1) or 1)),
                }
                for region in (regions or [])
            ),
            key=lambda item: (item['start'], -item['end'])
        )
        top_level = []
        current_end = 0
        for region in normalized:
            start_line = region['start']
            end_line = region['end']
            if end_line <= start_line:
                continue
            if start_line <= current_end:
                continue
            top_level.append(region)
            current_end = end_line
        return top_level

    def get_preferred_fold_region_at_line(self, doc, line_number, regions=None, collapsed_regions=None):
        if not doc:
            return None
        try:
            line_number = max(1, int(line_number))
        except (TypeError, ValueError):
            line_number = 1
        if collapsed_regions is None:
            collapsed_regions = self.get_collapsed_fold_regions(doc)
        collapsed_region = collapsed_regions.get(line_number)
        if collapsed_region:
            return collapsed_region
        if regions is None:
            regions = self.get_cached_fold_regions(doc)
        exact_matches = [region for region in regions if int(region.get('start', 0) or 0) == line_number]
        if exact_matches:
            return min(
                exact_matches,
                key=lambda region: (
                    int(region.get('end', line_number) or line_number) - int(region.get('start', line_number) or line_number),
                    -int(region.get('start', line_number) or line_number)
                )
            )
        containing_matches = [
            region for region in regions
            if int(region.get('start', 0) or 0) < line_number <= int(region.get('end', line_number) or line_number)
        ]
        if containing_matches:
            return min(
                containing_matches,
                key=lambda region: (
                    int(region.get('end', line_number) or line_number) - int(region.get('start', line_number) or line_number),
                    -int(region.get('start', line_number) or line_number)
                )
            )
        return None

    def clear_fold_tags(self, doc):
        if not doc:
            return
        text_widget = doc.get('text')
        if not text_widget:
            return
        for tag_name in list(doc.get('folded_tags') or []):
            try:
                text_widget.tag_delete(tag_name)
            except tk.TclError:
                pass
        doc['folded_tags'] = set()
        doc['collapsed_fold_regions'] = {}
        doc['collapsed_fold_regions_dirty'] = False

    def apply_fold_region(self, doc, region):
        if not doc or not region:
            return
        text_widget = doc.get('text')
        if not text_widget:
            return
        start_line = int(region.get('start', 1))
        end_line = int(region.get('end', start_line))
        if end_line <= start_line:
            return
        tag_name = self.make_fold_tag_name(start_line, end_line)
        try:
            text_widget.tag_config(tag_name, elide=True)
            text_widget.tag_add(tag_name, f'{start_line + 1}.0', f'{end_line}.0 lineend +1c')
            doc.setdefault('folded_tags', set()).add(tag_name)
            collapsed = doc.setdefault('collapsed_fold_regions', {})
            collapsed[start_line] = {
                'start': start_line,
                'end': end_line,
                'tag_name': tag_name,
            }
            doc['collapsed_fold_regions_dirty'] = False
        except tk.TclError:
            return

    def toggle_fold_region(self, doc, region):
        if not doc or not region:
            return False
        text_widget = doc.get('text')
        if not text_widget:
            return False
        start_line = int(region.get('start', 1))
        collapsed_region = self.get_collapsed_fold_regions(doc).get(start_line)
        if collapsed_region:
            tag_name = collapsed_region.get('tag_name')
            if tag_name:
                try:
                    text_widget.tag_delete(tag_name)
                except tk.TclError:
                    pass
                doc.setdefault('folded_tags', set()).discard(tag_name)
                doc.setdefault('collapsed_fold_regions', {}).pop(start_line, None)
                doc['collapsed_fold_regions_dirty'] = False
        else:
            self.apply_fold_region(doc, region)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        self.update_status()
        return True

    def expand_all_folds(self, event=None, doc=None):
        target_doc = doc or self.get_navigation_doc_for_widget(self.get_active_search_widget()) or self.get_current_doc()
        if not target_doc:
            return "break"
        self.clear_fold_tags(target_doc)
        self.update_line_number_gutter(target_doc)
        self.schedule_minimap_refresh(target_doc)
        return "break"

    def collapse_all_folds(self, event=None, doc=None):
        target_doc = doc or self.get_navigation_doc_for_widget(self.get_active_search_widget()) or self.get_current_doc()
        if not target_doc:
            return "break"
        self.clear_fold_tags(target_doc)
        regions = self.get_top_level_fold_regions(self.get_cached_fold_regions(target_doc))
        for region in regions:
            self.apply_fold_region(target_doc, region)
        self.update_line_number_gutter(target_doc)
        self.schedule_minimap_refresh(target_doc)
        return "break"

    def toggle_fold_at_cursor(self, event=None):
        widget = self.get_page_navigation_source_widget(event) or self.get_active_search_widget()
        target_doc = self.get_navigation_doc_for_widget(widget) if widget else self.get_current_doc()
        if not target_doc:
            return "break"
        text_widget = target_doc.get('text')
        if not text_widget:
            return "break"
        try:
            current_line = int(text_widget.index(tk.INSERT).split('.')[0])
        except (tk.TclError, ValueError):
            current_line = 1
        regions = self.get_cached_fold_regions(target_doc)
        collapsed_regions = self.get_collapsed_fold_regions(target_doc)
        selected_region = self.get_preferred_fold_region_at_line(
            target_doc,
            current_line,
            regions=regions,
            collapsed_regions=collapsed_regions
        )
        if not selected_region:
            return "break"
        self.toggle_fold_region(target_doc, selected_region)
        return "break"

    def is_log_like_path(self, file_path):
        file_name = os.path.basename(str(file_path or '')).lower()
        return file_name.endswith(('.log', '.trace', '.err', '.out')) or file_name in {'stdout', 'stderr'}

    def extract_log_diagnostics(self, lines):
        diagnostics = []
        in_traceback = False
        for line_number, line in enumerate(lines, start=1):
            stripped = str(line or '').strip()
            if not stripped:
                in_traceback = False
                continue

            if self.log_traceback_header_pattern.match(stripped):
                diagnostics.append({
                    'severity': 'error',
                    'line': line_number,
                    'message': 'Python traceback (most recent call last)'
                })
                in_traceback = True
                if len(diagnostics) >= 25:
                    break
                continue

            exception_match = self.log_exception_line_pattern.match(stripped)
            if in_traceback and self.log_traceback_frame_pattern.match(stripped):
                diagnostics.append({'severity': 'error', 'line': line_number, 'message': stripped})
                if len(diagnostics) >= 25:
                    break
                continue

            if in_traceback and stripped.lower().startswith('during handling of the above exception'):
                diagnostics.append({'severity': 'error', 'line': line_number, 'message': stripped})
                if len(diagnostics) >= 25:
                    break
                continue

            if exception_match:
                severity = 'warning' if exception_match.group(2) == 'Warning' else 'error'
                diagnostics.append({'severity': severity, 'line': line_number, 'message': stripped})
                in_traceback = False
                if len(diagnostics) >= 25:
                    break
                continue

            if self.log_error_level_pattern.match(stripped):
                diagnostics.append({'severity': 'error', 'line': line_number, 'message': stripped})
                in_traceback = False
                if len(diagnostics) >= 25:
                    break
                continue

            if self.log_warning_level_pattern.match(stripped):
                diagnostics.append({'severity': 'warning', 'line': line_number, 'message': stripped})
                in_traceback = False
                if len(diagnostics) >= 25:
                    break

        return diagnostics

    def clear_diagnostics(self, doc):
        if not doc:
            return
        text_widget = doc.get('text')
        if not text_widget:
            return
        if self.diagnostic_tooltip_doc is doc:
            self.hide_diagnostic_tooltip()
        try:
            text_widget.tag_remove('diagnostic_error', '1.0', tk.END)
            text_widget.tag_remove('diagnostic_warning', '1.0', tk.END)
        except tk.TclError:
            return
        doc['diagnostics'] = []
        self.schedule_minimap_refresh(doc)

    def toggle_diagnostics(self):
        if not self.diagnostics_enabled.get():
            for doc in self.documents.values():
                self.clear_diagnostics(doc)
            if self.compare_view:
                self.clear_diagnostics(self.compare_view)
        else:
            for doc in self.documents.values():
                self.schedule_diagnostics(doc)
            if self.compare_view:
                self.schedule_diagnostics(self.compare_view)
        self.save_session()
        return "break"

    def schedule_diagnostics(self, doc):
        if not doc:
            return
        existing_job = doc.get('diagnostic_job')
        if existing_job:
            try:
                self.root.after_cancel(existing_job)
            except tk.TclError:
                pass
        try:
            doc['diagnostic_job'] = self.root.after(self.diagnostics_delay_ms, lambda current=doc: self.run_diagnostics_for_doc(current))
        except tk.TclError:
            doc['diagnostic_job'] = None

    def run_diagnostics_for_doc(self, doc, force=False):
        if not doc:
            return []
        doc['diagnostic_job'] = None
        text_widget = doc.get('text')
        if not text_widget:
            return []
        self.clear_diagnostics(doc)
        if not self.diagnostics_enabled.get() and not force:
            self.schedule_minimap_refresh(doc)
            return []
        if doc.get('preview_mode') or doc.get('virtual_mode') or doc.get('large_file_mode'):
            self.schedule_minimap_refresh(doc)
            return []

        try:
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return []

        diagnostics = []
        syntax_mode = self.get_syntax_mode(doc) or self.infer_syntax_mode_from_path(doc.get('file_path'), content)
        lines = content.splitlines()

        for line_number, line in enumerate(lines, start=1):
            if line.rstrip(' \t') != line:
                diagnostics.append({'severity': 'warning', 'line': line_number, 'message': 'Trailing whitespace'})
            if len(line) > 140:
                diagnostics.append({'severity': 'warning', 'line': line_number, 'message': 'Long line (> 140 chars)'})
            if len(diagnostics) >= 50:
                break

        if self.is_log_like_path(doc.get('file_path')) or self.is_log_like_path(doc.get('remote_path')) or self.is_log_like_path(doc.get('display_name')):
            diagnostics = self.extract_log_diagnostics(lines) + diagnostics

        try:
            if syntax_mode == 'python':
                compile(content or '\n', self.get_doc_display_path(doc) or '<buffer>', 'exec')
            elif syntax_mode == 'json':
                json.loads(content or '{}')
            elif syntax_mode == 'xml':
                ET.fromstring(content or '<root />')
        except SyntaxError as exc:
            diagnostics.insert(0, {
                'severity': 'error',
                'line': int(getattr(exc, 'lineno', 1) or 1),
                'column': int(getattr(exc, 'offset', 1) or 1),
                'message': str(exc.msg or exc)
            })
        except json.JSONDecodeError as exc:
            diagnostics.insert(0, {
                'severity': 'error',
                'line': int(getattr(exc, 'lineno', 1) or 1),
                'column': int(getattr(exc, 'colno', 1) or 1),
                'message': str(exc.msg or exc)
            })
        except ET.ParseError as exc:
            error_message = str(exc)
            line_number = 1
            column_number = 1
            position = getattr(exc, 'position', None)
            if isinstance(position, tuple) and len(position) >= 2:
                line_number = int(position[0] or 1)
                column_number = int(position[1] or 1)
            diagnostics.insert(0, {
                'severity': 'error',
                'line': line_number,
                'column': column_number,
                'message': error_message
            })

        doc['diagnostics'] = diagnostics[:50]
        for diagnostic in doc['diagnostics']:
            tag_name = 'diagnostic_error' if diagnostic.get('severity') == 'error' else 'diagnostic_warning'
            line_number = max(1, int(diagnostic.get('line', 1)))
            try:
                text_widget.tag_add(tag_name, f'{line_number}.0', f'{line_number}.0 lineend')
            except tk.TclError:
                continue
        self.raise_editor_overlay_tags(text_widget)
        self.schedule_minimap_refresh(doc)
        return doc['diagnostics']

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

        self.apply_text_theme_effect(doc)
        self.schedule_spellcheck(doc)

    def build_rainbow_theme_palette(self, color_count=28):
        palette = []
        total_colors = max(2, int(color_count))
        for index in range(total_colors):
            hue = (0.82 * index) / (total_colors - 1)
            red, green, blue = colorsys.hsv_to_rgb(hue, 0.88, 1.0)
            palette.append(f"#{int(red * 255):02x}{int(green * 255):02x}{int(blue * 255):02x}")
        return palette

    def get_current_syntax_theme_definition(self):
        selected = (getattr(self, 'theme_definitions', None) or {}).get(self.syntax_theme.get())
        if selected is None:
            selected = self.sanitize_theme_definition(
                'Default',
                {'name': 'Default', **self.get_builtin_theme_definitions()['Default']}
            )
        return selected

    def get_syntax_theme_text_effect(self):
        selected = self.get_current_syntax_theme_definition()
        return str(selected.get('text_effect') or '').strip().lower()

    def get_syntax_surface_palette(self):
        selected = self.get_current_syntax_theme_definition()
        return dict(selected['surface'])

    def get_syntax_palette(self):
        selected = self.get_current_syntax_theme_definition()
        return dict(selected['syntax'])

    def get_spellcheck_tag_colors(self):
        surface = self.get_syntax_surface_palette()
        bg_value = str(surface.get('text_bg') or '#0d1117').lstrip('#')
        if len(bg_value) != 6:
            bg_value = '0d1117'
        try:
            red = int(bg_value[0:2], 16)
            green = int(bg_value[2:4], 16)
            blue = int(bg_value[4:6], 16)
        except ValueError:
            red, green, blue = (13, 17, 23)
        luminance = (0.2126 * red) + (0.7152 * green) + (0.0722 * blue)
        if luminance < 140:
            return '#ff7b8f', '#35141a'
        return '#9f1239', '#ffd7df'

    def get_rainbow_theme_tag_name(self, palette_index):
        return f"{self.rainbow_theme_tag_prefix}{int(palette_index)}"

    def clear_rainbow_theme_tags(self, text_widget):
        if not text_widget:
            return
        try:
            if not text_widget.winfo_exists():
                return
        except tk.TclError:
            return
        for palette_index in range(len(self.rainbow_theme_palette)):
            try:
                text_widget.tag_remove(self.get_rainbow_theme_tag_name(palette_index), '1.0', tk.END)
            except tk.TclError:
                return

    def configure_rainbow_theme_tags(self, text_widget):
        if not text_widget:
            return
        try:
            if not text_widget.winfo_exists():
                return
        except tk.TclError:
            return
        for palette_index, color_value in enumerate(self.rainbow_theme_palette):
            try:
                text_widget.tag_config(self.get_rainbow_theme_tag_name(palette_index), foreground=color_value)
            except tk.TclError:
                return

    def get_rainbow_theme_chunk_chars(self, content_length):
        if content_length <= 0:
            return 1
        return max(1, (int(content_length) + self.rainbow_theme_target_ranges - 1) // self.rainbow_theme_target_ranges)

    def raise_editor_overlay_tags(self, text_widget):
        if not text_widget:
            return
        try:
            if not text_widget.winfo_exists():
                return
            for tag_name in text_widget.tag_names():
                if str(tag_name).startswith('note_'):
                    text_widget.tag_raise(tag_name)
            text_widget.tag_raise('diagnostic_warning')
            text_widget.tag_raise('diagnostic_error')
            text_widget.tag_raise(self.spellcheck_tag)
            text_widget.tag_raise(self.bracket_match_tag)
            self.raise_find_tags(text_widget)
        except tk.TclError:
            return

    def get_spell_checker(self):
        if self.spell_checker_ready:
            return self.spell_checker
        self.spell_checker_ready = True
        if SpellChecker is None:
            self.spell_checker = None
            return None
        try:
            dictionary_path = self.get_spellcheck_dictionary_path()
            if dictionary_path:
                checker = SpellChecker(language=None, local_dictionary=dictionary_path)
            else:
                checker = SpellChecker(language='en')
            checker.word_frequency.load_words(sorted(self.spellcheck_custom_words))
            self.spell_checker = checker
        except Exception as exc:
            self.log_exception("initialize spell checker", exc)
            self.spell_checker = None
        return self.spell_checker

    def get_spellcheck_dictionary_path(self):
        candidates = []
        for base_dir in (self.app_dir, self.resource_dir):
            if not base_dir:
                continue
            candidates.append(os.path.join(base_dir, 'cfg', 'spellcheck', 'en.json.gz'))
        seen = set()
        for candidate in candidates:
            normalized = os.path.normcase(os.path.abspath(candidate))
            if normalized in seen:
                continue
            seen.add(normalized)
            if os.path.isfile(candidate):
                return candidate
        return None

    def ensure_spellcheck_available(self, notify=False):
        checker = self.get_spell_checker()
        if checker is not None:
            return True
        if notify:
            messagebox.showinfo(
                self.tr('spellcheck.unavailable_title', 'Spell Check Unavailable'),
                self.tr('spellcheck.unavailable_message', 'Spell check needs pyspellchecker and the bundled English dictionary. Rebuild Notepad-X if the menu shows enabled but no words are marked.'),
                parent=self.root
            )
        return False

    def cancel_spellcheck_job(self, doc):
        if not doc:
            return
        job = doc.get('spellcheck_job')
        if not job:
            return
        try:
            self.root.after_cancel(job)
        except tk.TclError:
            pass
        doc['spellcheck_job'] = None

    def clear_spellcheck(self, doc_or_widget):
        if isinstance(doc_or_widget, dict):
            doc = doc_or_widget
            self.cancel_spellcheck_job(doc)
            text_widget = doc.get('text')
        else:
            doc = None
            text_widget = doc_or_widget
        if not isinstance(text_widget, tk.Text):
            return
        try:
            if text_widget.winfo_exists():
                text_widget.tag_remove(self.spellcheck_tag, '1.0', tk.END)
        except tk.TclError:
            pass

    def doc_supports_spellcheck(self, doc):
        if not doc or not self.spell_check_enabled.get():
            return False
        if doc.get('virtual_mode') or doc.get('preview_mode') or doc.get('large_file_mode'):
            return False
        if doc.get('syntax_mode') not in self.spellcheck_supported_modes:
            return False
        text_widget = doc.get('text')
        if not isinstance(text_widget, tk.Text):
            return False
        try:
            if not text_widget.winfo_exists():
                return False
        except tk.TclError:
            return False
        return self.get_spell_checker() is not None

    def is_spellcheck_candidate_word(self, content, match):
        token = match.group(0)
        normalized = token.lower()
        if len(normalized) < 2 or len(normalized) > 32:
            return False
        if normalized in self.spellcheck_custom_words:
            return False
        if token.isupper():
            return False
        if re.search(r'[a-z][A-Z]|[A-Z]{2,}[a-z]', token):
            return False
        start = match.start()
        end = match.end()
        before_char = content[start - 1] if start > 0 else ''
        after_char = content[end] if end < len(content) else ''
        skip_chars = self.spellcheck_skip_neighbor_chars - {'.'}
        if before_char in skip_chars or after_char in skip_chars:
            return False
        if before_char == '.':
            prior_char = content[start - 2] if start > 1 else ''
            if prior_char and (prior_char.isalnum() or prior_char in {'_', '-', '@'}):
                return False
        if after_char == '.':
            next_char = content[end + 1] if end + 1 < len(content) else ''
            if next_char and (next_char.isalnum() or next_char in {'_', '-', '@'}):
                return False
        if before_char == '.' and after_char == '.':
            return False
        return True

    def iter_misspelled_word_ranges(self, content):
        checker = self.get_spell_checker()
        if checker is None or not content or len(content) > self.spellcheck_max_chars:
            return []

        pending = []
        unique_words = set()
        for match in self.spellcheck_token_pattern.finditer(content):
            if len(pending) >= self.spellcheck_max_words:
                break
            if not self.is_spellcheck_candidate_word(content, match):
                continue
            normalized = match.group(0).lower()
            pending.append((match.start(), match.end(), normalized))
            unique_words.add(normalized)

        if not pending:
            return []

        try:
            unknown_words = checker.unknown(unique_words)
        except Exception as exc:
            self.log_exception("spellcheck scan", exc)
            return []

        return [
            (start, end)
            for start, end, normalized in pending
            if normalized in unknown_words
        ]

    def apply_spellcheck(self, doc):
        if not doc:
            return
        doc['spellcheck_job'] = None
        text_widget = doc.get('text')
        self.clear_spellcheck(text_widget)
        if not self.doc_supports_spellcheck(doc):
            return
        try:
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return
        misspelled_ranges = self.iter_misspelled_word_ranges(content)
        if not misspelled_ranges:
            return
        for start_offset, end_offset in misspelled_ranges:
            start_index = self.text_index_from_offset(text_widget, start_offset, content=content)
            end_index = self.text_index_from_offset(text_widget, end_offset, content=content)
            if not start_index or not end_index:
                continue
            try:
                text_widget.tag_add(self.spellcheck_tag, start_index, end_index)
            except tk.TclError:
                continue
        self.raise_editor_overlay_tags(text_widget)

    def schedule_spellcheck(self, doc):
        if not doc:
            return
        if not self.doc_supports_spellcheck(doc):
            self.clear_spellcheck(doc)
            return
        self.cancel_spellcheck_job(doc)
        try:
            doc['spellcheck_job'] = self.root.after(
                self.spellcheck_delay_ms,
                lambda current=doc: self.apply_spellcheck(current)
            )
        except tk.TclError:
            doc['spellcheck_job'] = None
            self.apply_spellcheck(doc)

    def apply_rainbow_theme_to_widget(self, text_widget):
        if not text_widget:
            return
        try:
            if not text_widget.winfo_exists():
                return
            content = text_widget.get('1.0', 'end-1c')
        except tk.TclError:
            return

        self.clear_rainbow_theme_tags(text_widget)
        if not content:
            self.raise_editor_overlay_tags(text_widget)
            return

        self.configure_rainbow_theme_tags(text_widget)
        chunk_chars = self.get_rainbow_theme_chunk_chars(len(content))
        palette_size = len(self.rainbow_theme_palette)
        for offset in range(0, len(content), chunk_chars):
            end_offset = min(len(content), offset + chunk_chars)
            palette_index = (offset // chunk_chars) % palette_size
            text_widget.tag_add(
                self.get_rainbow_theme_tag_name(palette_index),
                f"1.0+{offset}c",
                f"1.0+{end_offset}c"
            )
        self.raise_editor_overlay_tags(text_widget)

    def activate_help_lolcat(self, event=None):
        text_widget = getattr(event, 'widget', None)
        if not isinstance(text_widget, tk.Text):
            return
        text_widget.help_lolcat_active = True
        self.apply_rainbow_theme_to_widget(text_widget)
        return "break"

    def clear_help_lolcat(self, text_widget):
        if not isinstance(text_widget, tk.Text):
            return
        try:
            if not text_widget.winfo_exists():
                return
        except tk.TclError:
            return
        text_widget.help_lolcat_active = False
        self.clear_rainbow_theme_tags(text_widget)

    def close_help_contents_dialog(self, dialog, help_text):
        self.clear_help_lolcat(help_text)
        try:
            if dialog.winfo_exists():
                dialog.destroy()
        except tk.TclError:
            pass

    def cancel_text_theme_effect_job(self, doc):
        if not doc:
            return
        theme_effect_job = doc.get('theme_effect_job')
        if not theme_effect_job:
            return
        try:
            self.root.after_cancel(theme_effect_job)
        except tk.TclError:
            pass
        doc['theme_effect_job'] = None

    def apply_text_theme_effect(self, doc):
        if not doc:
            return
        self.cancel_text_theme_effect_job(doc)
        text_widget = doc.get('text')
        if not text_widget:
            return
        if doc.get('large_file_mode'):
            self.clear_rainbow_theme_tags(text_widget)
            self.raise_editor_overlay_tags(text_widget)
            return
        if self.get_syntax_theme_text_effect() == 'rainbow' and not doc.get('virtual_mode') and not doc.get('preview_mode'):
            self.apply_rainbow_theme_to_widget(text_widget)
            return
        self.clear_rainbow_theme_tags(text_widget)
        self.raise_editor_overlay_tags(text_widget)

    def schedule_text_theme_effect(self, doc):
        if not doc:
            return
        if doc.get('large_file_mode'):
            self.cancel_text_theme_effect_job(doc)
            return
        if self.get_syntax_theme_text_effect() != 'rainbow':
            self.cancel_text_theme_effect_job(doc)
            return
        self.cancel_text_theme_effect_job(doc)
        try:
            doc['theme_effect_job'] = self.root.after(
                self.theme_effect_delay_ms,
                lambda current=doc: self.apply_text_theme_effect(current)
            )
        except tk.TclError:
            doc['theme_effect_job'] = None
            self.apply_text_theme_effect(doc)

    def apply_syntax_tag_colors(self, text_widget):
        surface = self.get_syntax_surface_palette()
        palette = self.get_syntax_palette()
        spellcheck_fg, spellcheck_bg = self.get_spellcheck_tag_colors()
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
        text_widget.tag_config(self.spellcheck_tag, underline=1, foreground=spellcheck_fg, background=spellcheck_bg)

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
        if self.markdown_preview_enabled.get():
            self.schedule_markdown_preview_refresh()
        self.save_session()

    def set_current_syntax_override(self, syntax_mode):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        doc['syntax_override'] = syntax_mode
        self.syntax_mode_selection.set(syntax_mode)
        self.configure_syntax_highlighting(doc['frame'])
        self.invalidate_fold_regions(doc)
        self.update_line_number_gutter(doc)
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

    def is_virtual_index_ready(self, doc):
        if not doc or not doc.get('virtual_mode'):
            return False
        line_starts = doc.get('line_starts')
        return isinstance(line_starts, list) and len(line_starts) > 0

    def get_virtual_line_end_byte(self, doc, line_number):
        line_starts = doc.get('line_starts') or [0]
        total_lines = max(1, int(doc.get('total_file_lines', len(line_starts)) or len(line_starts)))
        safe_line_number = max(1, min(total_lines, int(line_number or 1)))
        if safe_line_number < total_lines and safe_line_number < len(line_starts):
            return line_starts[safe_line_number]
        return max(0, int(doc.get('file_size_bytes', 0) or 0))

    def get_virtual_line_span_bytes(self, doc, start_line, end_line):
        if not self.is_virtual_index_ready(doc):
            return 0
        line_starts = doc.get('line_starts') or [0]
        total_lines = max(1, int(doc.get('total_file_lines', len(line_starts)) or len(line_starts)))
        safe_start_line = max(1, min(total_lines, int(start_line or 1)))
        safe_end_line = max(safe_start_line, min(total_lines, int(end_line or safe_start_line)))
        start_byte = line_starts[safe_start_line - 1]
        end_byte = self.get_virtual_line_end_byte(doc, safe_end_line)
        return max(0, end_byte - start_byte)

    def get_virtual_window_bounds(self, doc, target_line):
        line_starts = doc.get('line_starts') or [0]
        total_lines = max(1, int(doc.get('total_file_lines', len(line_starts)) or len(line_starts)))
        max_lines = max(1, int(self.virtual_file_window_lines or 1))
        max_bytes = max(1, int(self.virtual_file_window_max_bytes or 1))
        file_size_bytes = max(0, int(doc.get('file_size_bytes', 0) or 0))
        if file_size_bytes >= self.huge_file_preview_threshold_bytes:
            max_bytes = min(max_bytes, max(1, int(self.huge_virtual_file_window_max_bytes or 1)))
        safe_target_line = max(1, min(total_lines, int(target_line or 1)))
        start_line = max(1, safe_target_line - (max_lines // 2))
        end_line = min(total_lines, start_line + max_lines - 1)
        start_line = max(1, end_line - max_lines + 1)

        while start_line < end_line and self.get_virtual_line_span_bytes(doc, start_line, end_line) > max_bytes:
            if (safe_target_line - start_line) >= (end_line - safe_target_line):
                start_line += 1
            else:
                end_line -= 1

        expanded = True
        while expanded and (end_line - start_line + 1) < max_lines:
            expanded = False
            if start_line > 1 and self.get_virtual_line_span_bytes(doc, start_line - 1, end_line) <= max_bytes:
                start_line -= 1
                expanded = True
            if end_line < total_lines and self.get_virtual_line_span_bytes(doc, start_line, end_line + 1) <= max_bytes:
                end_line += 1
                expanded = True

        return start_line, end_line

    def read_virtual_line_window(self, doc, start_line, end_line):
        if not self.is_virtual_index_ready(doc):
            return ""
        line_starts = doc.get('line_starts') or [0]
        total_lines = max(1, int(doc.get('total_file_lines', len(line_starts)) or len(line_starts)))
        safe_start_line = max(1, min(total_lines, int(start_line or 1)))
        safe_end_line = max(safe_start_line, min(total_lines, int(end_line or safe_start_line)))

        start_byte = line_starts[safe_start_line - 1]
        if safe_end_line < total_lines and safe_end_line < len(line_starts):
            end_byte = line_starts[safe_end_line]
        else:
            end_byte = int(doc.get('file_size_bytes', 0) or 0)

        if self.is_virtual_editable(doc):
            return self.read_virtual_byte_range_text(doc, start_byte, end_byte)

        with open(doc['file_path'], 'rb') as f:
            f.seek(start_byte)
            data = f.read(end_byte - start_byte)
        return data.decode('utf-8', errors='replace')

    def load_virtual_window(self, doc, target_line=1):
        if not self.is_virtual_index_ready(doc):
            return False
        total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
        target_line = max(1, min(total_lines, int(target_line or 1)))
        if self.is_virtual_editable(doc):
            try:
                current_window_dirty = bool(doc['text'].edit_modified())
            except tk.TclError:
                current_window_dirty = False
            if current_window_dirty:
                predicted_start_line, predicted_end_line = self.get_virtual_window_bounds(doc, target_line)
                if (
                    predicted_start_line != int(doc.get('window_start_line', 0) or 0) or
                    predicted_end_line != int(doc.get('window_end_line', 0) or 0)
                ):
                    if not self.flush_virtual_window_edits(doc, force=True):
                        return False
                    total_lines = max(1, int(doc.get('total_file_lines', 1) or 1))
                    target_line = max(1, min(total_lines, int(target_line or 1)))
        start_line, end_line = self.get_virtual_window_bounds(doc, target_line)
        text = doc['text']
        try:
            current_col = int(text.index(tk.INSERT).split('.')[1])
        except tk.TclError:
            current_col = 0

        if (
            start_line == int(doc.get('window_start_line', 0) or 0) and
            end_line == int(doc.get('window_end_line', 0) or 0)
        ):
            local_line = max(1, target_line - start_line + 1)
            try:
                line_length = len(text.get(f"{local_line}.0", f"{local_line}.end"))
                text.mark_set(tk.INSERT, f"{local_line}.{min(current_col, line_length)}")
                text.see(f"{local_line}.0")
            except tk.TclError:
                return False
            doc['last_virtual_line'] = target_line
            doc['last_virtual_col'] = max(0, min(current_col, line_length))
            self.update_vertical_scrollbar(doc['frame'], None, None, None)
            self.update_line_number_gutter(doc)
            self.schedule_minimap_refresh(doc)
            return True

        content = self.read_virtual_line_window(doc, start_line, end_line)

        doc['suspend_modified_events'] = True
        text.delete('1.0', tk.END)
        text.insert('1.0', content)
        text.edit_modified(False)
        doc['suspend_modified_events'] = False

        doc['window_start_line'] = start_line
        doc['window_end_line'] = end_line
        doc['virtual_window_start_byte'] = int((doc.get('line_starts') or [0])[start_line - 1])
        doc['virtual_window_end_byte'] = self.get_virtual_line_end_byte(doc, end_line)

        local_line = max(1, target_line - start_line + 1)
        line_length = 0
        try:
            line_length = len(text.get(f"{local_line}.0", f"{local_line}.end"))
            text.mark_set(tk.INSERT, f"{local_line}.{min(current_col, line_length)}")
            text.see(f"{local_line}.0")
        except tk.TclError:
            text.mark_set(tk.INSERT, '1.0')

        doc['last_virtual_line'] = target_line
        doc['last_virtual_col'] = max(0, min(current_col, line_length))
        self.update_vertical_scrollbar(doc['frame'], None, None, None)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        return True

    def ensure_virtual_line_visible(self, doc, target_line=None):
        if not doc.get('virtual_mode') or not self.is_virtual_index_ready(doc):
            return False

        if target_line is None:
            try:
                local_line = int(doc['text'].index(tk.INSERT).split('.')[0])
            except tk.TclError:
                local_line = 1
            target_line = doc['window_start_line'] + local_line - 1

        target_line = max(1, min(doc['total_file_lines'], target_line))
        local_line = target_line - doc['window_start_line'] + 1
        visible_lines = max(1, doc['window_end_line'] - doc['window_start_line'] + 1)
        margin_lines = min(self.virtual_file_margin_lines, max(5, visible_lines // 4))

        if local_line <= margin_lines or local_line >= visible_lines - margin_lines:
            return self.load_virtual_window(doc, target_line)
        return True

    def handle_text_activity(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        self.set_last_active_editor_widget(doc['text'])
        self.remember_doc_view_state(doc)
        event_type = str(getattr(event, 'type', ''))
        if event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease', '35', 'KeyRelease'}:
            self.hide_diagnostic_tooltip()
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
        if event_type in {'3', 'KeyRelease'} and not doc.get('virtual_mode') and not doc.get('preview_mode'):
            self.schedule_text_theme_effect(doc)
            self.schedule_spellcheck(doc)
        self.update_bracket_match_highlight(doc)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
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

    def build_auto_indent_insert_text(self, text_widget):
        try:
            insert_index = text_widget.index(tk.INSERT)
            current_line = text_widget.get(f"{insert_index} linestart", f"{insert_index} lineend")
            before_char = text_widget.get(f"{insert_index} -1c")
            after_char = text_widget.get(insert_index)
        except tk.TclError:
            return '\n'
        indent_match = re.match(r'[ \t]*', current_line)
        base_indent = indent_match.group(0) if indent_match else ''
        indent_prefix = self.build_auto_indent_prefix(text_widget, insert_index)
        closing_char = self.bracket_pairs.get(before_char)
        if closing_char and after_char == closing_char:
            return '\n' + indent_prefix + '\n' + base_indent
        return '\n' + indent_prefix

    def insert_auto_indent_newline(self, text_widget):
        newline_text = self.build_auto_indent_insert_text(text_widget)
        if text_widget.tag_ranges('sel'):
            try:
                text_widget.delete('sel.first', 'sel.last')
            except tk.TclError:
                pass
        try:
            insert_index = text_widget.index(tk.INSERT)
        except tk.TclError:
            insert_index = None
        text_widget.insert(tk.INSERT, newline_text)
        if newline_text.count('\n') == 2 and insert_index is not None:
            try:
                line_number = int(insert_index.split('.')[0]) + 1
                prefix = newline_text.split('\n')[1]
                text_widget.mark_set(tk.INSERT, f'{line_number}.{len(prefix)}')
            except (tk.TclError, ValueError, IndexError):
                pass
        return "break"

    def get_compare_mirror_target(self, source_widget):
        if not self.compare_active or not self.compare_multi_edit_enabled.get():
            return None
        compare_widget = self.get_compare_text_widget()
        main_widget = self.text if isinstance(self.text, tk.Text) else None
        if compare_widget is None or main_widget is None:
            return None
        if source_widget == compare_widget and main_widget.winfo_exists():
            return main_widget
        if source_widget == main_widget and compare_widget.winfo_exists():
            return compare_widget
        return None

    def sync_mirror_target_position(self, source_widget, target_widget):
        if not source_widget or not target_widget:
            return
        try:
            insert_index = source_widget.index(tk.INSERT)
            target_widget.mark_set(tk.INSERT, insert_index)
            target_widget.tag_remove('sel', '1.0', tk.END)
            selection = source_widget.tag_ranges('sel')
            if len(selection) >= 2:
                target_widget.tag_add('sel', str(selection[0]), str(selection[1]))
        except tk.TclError:
            return

    def notify_text_widget_changed(self, widget):
        compare_widget = self.get_compare_text_widget()
        if compare_widget is not None and widget == compare_widget:
            self.on_compare_modified()
            return
        doc = self.get_doc_for_text_widget(widget)
        if doc:
            self.on_text_modified(doc['frame'])

    def apply_widget_text_change(self, widget, inserted_text='', delete_selection=False, delete_backward=False, delete_forward=False):
        if not widget:
            return False
        try:
            if delete_selection:
                try:
                    widget.delete('sel.first', 'sel.last')
                except tk.TclError:
                    pass
            if delete_backward:
                try:
                    if widget.compare(tk.INSERT, '>', '1.0'):
                        widget.delete(f'{tk.INSERT} -1c')
                except tk.TclError:
                    pass
            if delete_forward:
                try:
                    widget.delete(tk.INSERT)
                except tk.TclError:
                    pass
            if inserted_text:
                widget.insert(tk.INSERT, inserted_text)
            widget.edit_modified(True)
            widget.see(tk.INSERT)
            return True
        except tk.TclError:
            return False

    def mirror_text_change(self, source_widget, inserted_text='', delete_selection=False, delete_backward=False, delete_forward=False):
        target_widget = self.get_compare_mirror_target(source_widget)
        if target_widget is None:
            return
        self.sync_mirror_target_position(source_widget, target_widget)
        if self.apply_widget_text_change(
            target_widget,
            inserted_text=inserted_text,
            delete_selection=delete_selection,
            delete_backward=delete_backward,
            delete_forward=delete_forward
        ):
            self.notify_text_widget_changed(target_widget)

    def should_auto_pair_quote(self, text_widget, quote_char):
        try:
            before_char = text_widget.get(f'{tk.INSERT} -1c') if text_widget.compare(tk.INSERT, '>', '1.0') else ''
            after_char = text_widget.get(tk.INSERT)
        except tk.TclError:
            return False
        if self.text_has_selection(text_widget):
            return True
        if after_char == quote_char:
            return False
        if before_char and re.match(r'[A-Za-z0-9_]', before_char):
            return False
        return not after_char or after_char.isspace() or after_char in ')]}:;,.'

    def handle_auto_pair_input(self, text_widget, char):
        if not self.auto_pair_enabled.get():
            return None
        selection_ranges = text_widget.tag_ranges('sel')
        mirror_target = self.get_compare_mirror_target(text_widget)

        if char in self.bracket_pairs:
            closing_char = self.bracket_pairs[char]
            if selection_ranges:
                try:
                    selected_text = text_widget.get('sel.first', 'sel.last')
                except tk.TclError:
                    selected_text = ''
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                self.apply_widget_text_change(text_widget, delete_selection=True, inserted_text=char + selected_text + closing_char)
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.apply_widget_text_change(mirror_target, delete_selection=True, inserted_text=char + selected_text + closing_char)
                    try:
                        mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                    except tk.TclError:
                        pass
                    self.notify_text_widget_changed(mirror_target)
            else:
                self.apply_widget_text_change(text_widget, inserted_text=char + closing_char)
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                    self.apply_widget_text_change(mirror_target, inserted_text=char + closing_char)
                    try:
                        mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                    except tk.TclError:
                        pass
                    self.notify_text_widget_changed(mirror_target)
            self.notify_text_widget_changed(text_widget)
            return "break"

        if char in self.reverse_bracket_pairs:
            try:
                next_char = text_widget.get(tk.INSERT)
            except tk.TclError:
                next_char = ''
            if next_char == char:
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} +1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                    try:
                        if mirror_target.get(tk.INSERT) == char:
                            mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} +1c')
                    except tk.TclError:
                        pass
                return "break"

        if char in {'"', "'"} and self.should_auto_pair_quote(text_widget, char):
            if selection_ranges:
                try:
                    selected_text = text_widget.get('sel.first', 'sel.last')
                except tk.TclError:
                    selected_text = ''
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                self.apply_widget_text_change(text_widget, delete_selection=True, inserted_text=char + selected_text + char)
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.apply_widget_text_change(mirror_target, delete_selection=True, inserted_text=char + selected_text + char)
                    try:
                        mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                    except tk.TclError:
                        pass
                    self.notify_text_widget_changed(mirror_target)
            else:
                self.apply_widget_text_change(text_widget, inserted_text=char + char)
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                    self.apply_widget_text_change(mirror_target, inserted_text=char + char)
                    try:
                        mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} -1c')
                    except tk.TclError:
                        pass
                    self.notify_text_widget_changed(mirror_target)
            self.notify_text_widget_changed(text_widget)
            return "break"
        if char in {'"', "'"}:
            try:
                next_char = text_widget.get(tk.INSERT)
            except tk.TclError:
                next_char = ''
            if next_char == char:
                try:
                    text_widget.mark_set(tk.INSERT, f'{tk.INSERT} +1c')
                except tk.TclError:
                    pass
                if mirror_target is not None:
                    self.sync_mirror_target_position(text_widget, mirror_target)
                    try:
                        if mirror_target.get(tk.INSERT) == char:
                            mirror_target.mark_set(tk.INSERT, f'{tk.INSERT} +1c')
                    except tk.TclError:
                        pass
                return "break"
        return None

    def handle_auto_pair_backspace(self, text_widget):
        if not self.auto_pair_enabled.get():
            return None
        if self.text_has_selection(text_widget):
            return None
        try:
            if not text_widget.compare(tk.INSERT, '>', '1.0'):
                return None
            previous_char = text_widget.get(f'{tk.INSERT} -1c')
            next_char = text_widget.get(tk.INSERT)
        except tk.TclError:
            return None
        expected_next = self.bracket_pairs.get(previous_char, previous_char if previous_char in {'"', "'"} else None)
        if expected_next and next_char == expected_next:
            self.apply_widget_text_change(text_widget, delete_backward=True, delete_forward=True)
            self.mirror_text_change(text_widget, delete_backward=True, delete_forward=True)
            self.notify_text_widget_changed(text_widget)
            return "break"
        return None

    def handle_compare_multi_edit_newline(self, text_widget):
        newline_text = self.build_auto_indent_insert_text(text_widget)
        has_selection = bool(text_widget.tag_ranges('sel'))
        mirror_target = self.get_compare_mirror_target(text_widget)
        if mirror_target is not None:
            self.sync_mirror_target_position(text_widget, mirror_target)
        self.apply_widget_text_change(text_widget, delete_selection=has_selection, inserted_text=newline_text)
        if newline_text.count('\n') == 2:
            try:
                line_number = int(text_widget.index(tk.INSERT).split('.')[0]) - 1
                prefix = newline_text.split('\n')[1]
                text_widget.mark_set(tk.INSERT, f'{line_number}.{len(prefix)}')
            except (tk.TclError, ValueError, IndexError):
                pass
        if mirror_target is not None:
            self.apply_widget_text_change(mirror_target, delete_selection=has_selection, inserted_text=newline_text)
            self.notify_text_widget_changed(mirror_target)
        self.notify_text_widget_changed(text_widget)
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

    def handle_autocomplete_keypress(self, event, doc_id):
        if not self.autocomplete_popup_visible() or self.autocomplete_doc_id != doc_id:
            return None
        if event.keysym == 'Up':
            return self.move_autocomplete_selection(-1)
        if event.keysym == 'Down':
            return self.move_autocomplete_selection(1)
        if event.keysym == 'Tab':
            if self.accept_autocomplete_selection():
                return "break"
        if event.keysym in {'Return', 'KP_Enter'}:
            self.hide_autocomplete_popup()
            return None
        if event.keysym == 'Escape':
            self.hide_autocomplete_popup()
            return "break"
        if event.keysym in {'Left', 'Right', 'Home', 'End', 'Prior', 'Next'}:
            self.hide_autocomplete_popup()
        return None

    def perform_mirrored_edit(self, source_widget, inserted_text='', delete_selection=False, delete_backward=False, delete_forward=False):
        mirror_target = self.get_compare_mirror_target(source_widget)
        if mirror_target is not None:
            self.sync_mirror_target_position(source_widget, mirror_target)
        if not self.apply_widget_text_change(
            source_widget,
            inserted_text=inserted_text,
            delete_selection=delete_selection,
            delete_backward=delete_backward,
            delete_forward=delete_forward
        ):
            return False
        if mirror_target is not None:
            self.apply_widget_text_change(
                mirror_target,
                inserted_text=inserted_text,
                delete_selection=delete_selection,
                delete_backward=delete_backward,
                delete_forward=delete_forward
            )
            self.notify_text_widget_changed(mirror_target)
        self.notify_text_widget_changed(source_widget)
        return True

    def handle_editable_text_keypress(self, event, text_widget, doc, doc_id):
        autocomplete_result = self.handle_autocomplete_keypress(event, doc_id)
        if autocomplete_result is not None:
            return autocomplete_result

        state = int(getattr(event, 'state', 0) or 0)
        if (state & 0x4) or (state & 0x20000):
            return None

        keysym = str(getattr(event, 'keysym', '') or '')
        char = str(getattr(event, 'char', '') or '')
        has_selection = self.text_has_selection(text_widget)

        if keysym in {'Return', 'KP_Enter'}:
            return self.handle_compare_multi_edit_newline(text_widget)

        if keysym == 'BackSpace':
            auto_pair_result = self.handle_auto_pair_backspace(text_widget)
            if auto_pair_result is not None:
                return auto_pair_result
            if self.get_compare_mirror_target(text_widget) is not None:
                if self.perform_mirrored_edit(text_widget, delete_selection=has_selection, delete_backward=not has_selection):
                    return "break"
            return None

        if keysym == 'Delete':
            if self.get_compare_mirror_target(text_widget) is not None:
                if self.perform_mirrored_edit(text_widget, delete_selection=has_selection, delete_forward=not has_selection):
                    return "break"
            return None

        if keysym == 'Tab':
            insert_text = self.get_auto_indent_unit('')
            if self.get_compare_mirror_target(text_widget) is not None:
                if self.perform_mirrored_edit(text_widget, inserted_text=insert_text, delete_selection=has_selection):
                    return "break"
            return None

        if len(char) == 1 and ord(char) >= 32:
            auto_pair_result = self.handle_auto_pair_input(text_widget, char)
            if auto_pair_result is not None:
                return auto_pair_result
            if self.get_compare_mirror_target(text_widget) is not None:
                if self.perform_mirrored_edit(text_widget, inserted_text=char, delete_selection=has_selection):
                    return "break"
        return None

    def handle_text_keypress(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return
        self.hide_diagnostic_tooltip()
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
            if not self.is_doc_text_readonly(doc):
                self.hide_autocomplete_popup()
                return self.handle_compare_multi_edit_newline(doc['text'])

        if not self.is_doc_text_readonly(doc):
            return self.handle_editable_text_keypress(event, doc['text'], doc, str(tab_id))

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
        self.hide_diagnostic_tooltip()
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
                return self.handle_compare_multi_edit_newline(compare_doc['text'])
        if source_doc and not source_doc.get('virtual_mode') and not source_doc.get('preview_mode'):
            compare_doc_id = str(compare_doc['frame']) if compare_doc else None
            return self.handle_editable_text_keypress(event, compare_doc['text'], compare_doc, compare_doc_id)

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
        if event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease', '35', 'KeyRelease'}:
            self.hide_diagnostic_tooltip()
        if self.is_shift_selection_navigation(event):
            self.hide_autocomplete_popup()
        elif event_type in {'4', '5', '6', 'Motion', 'ButtonPress', 'ButtonRelease'}:
            self.hide_autocomplete_popup()
        elif event_type not in {'35'}:
            self.update_autocomplete_popup(compare_doc)
        if event_type in {'3', 'KeyRelease'}:
            self.schedule_text_theme_effect(compare_doc)
            self.schedule_spellcheck(compare_doc)
            self.schedule_diagnostics(compare_doc)
        self.update_bracket_match_highlight(compare_doc)
        self.update_line_number_gutter(compare_doc)
        self.schedule_minimap_refresh(compare_doc)
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
            self.invalidate_collapsed_fold_regions(compare_doc)
            self.update_line_number_gutter(compare_doc)
            self.schedule_minimap_refresh(compare_doc)
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
        self.invalidate_collapsed_fold_regions(compare_doc)
        if compare_doc.get('syntax_mode') and compare_doc.get('syntax_mode') != 'python':
            self.schedule_syntax_highlight(compare_doc)
        self.schedule_text_theme_effect(compare_doc)
        compare_doc['symbol_cache_signature'] = None
        compare_doc['symbol_cache'] = None
        self.invalidate_fold_regions(compare_doc)
        self.invalidate_minimap_cache(compare_doc)
        self.schedule_diagnostics(compare_doc)
        self.update_line_number_gutter(compare_doc)
        self.schedule_minimap_refresh(compare_doc)
        self.update_status()
        compare_text.edit_modified(False)

    def on_text_mousewheel(self, event, tab_id):
        doc = self.documents.get(str(tab_id))
        self.hide_diagnostic_tooltip()
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
        self.hide_diagnostic_tooltip()

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
            current_window_lines = max(1, int(doc.get('window_end_line', 1) or 1) - int(doc.get('window_start_line', 1) or 1) + 1)
            step = 1 if unit == 'units' else max(1, current_window_lines // 2)
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
        self.hide_diagnostic_tooltip()
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
        doc['context_menu'] = menu
        self.rebuild_text_context_menu(
            doc,
            action_target,
            note_state='disabled',
            note_action_state='disabled',
            is_readonly_target=False,
            word_index=None
        )

    def rebuild_text_context_menu(self, doc, action_target, note_state, note_action_state, is_readonly_target, word_index=None):
        menu = doc.get('context_menu')
        if menu is None:
            return
        menu.delete(0, tk.END)

        target_widget = doc.get('context_target_widget') or doc.get('text')
        word_info = None
        if not is_readonly_target and isinstance(target_widget, tk.Text):
            word_info = self.get_misspelled_word_info_at_index(target_widget, word_index, doc=doc)

        if word_info:
            suggestions = self.get_spellcheck_suggestions(word_info['word'])
            suggestions = self.format_spellcheck_suggestions(suggestions, word_info)
            if suggestions:
                for suggestion in suggestions:
                    menu.add_command(
                        label=suggestion,
                        command=lambda frame=action_target, start=word_info['start'], end=word_info['end'], replacement=suggestion:
                            self.run_context_menu_action(lambda current_frame=frame, current_start=start, current_end=end, current_replacement=replacement:
                                self.replace_word_in_widget(current_frame, current_start, current_end, current_replacement))
                    )
            else:
                menu.add_command(label=self.tr('context.no_suggestions', 'No suggestions'), state='disabled')
            menu.add_command(
                label=self.tr('context.add_to_dictionary', 'Add to Dictionary'),
                command=lambda word=word_info['normalized']: self.run_context_menu_action(
                    lambda current_word=word: self.add_word_to_spellcheck_dictionary(current_word)
                )
            )
            menu.add_separator()

        menu.add_command(label=self.tr('context.cut', 'Cut'), state='disabled' if is_readonly_target else 'normal', command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.cut))
        menu.add_command(label=self.tr('context.copy', 'Copy'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.copy))
        menu.add_command(label=self.tr('context.paste', 'Paste'), state='disabled' if is_readonly_target else 'normal', command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.paste))
        menu.add_separator()
        menu.add_command(label=self.tr('context.select_all', 'Select All'), command=lambda frame=action_target: self.run_context_menu_widget_action(frame, self.select_all))
        menu.add_command(label=self.tr('context.add_note', 'Add note'), state=note_state, command=lambda frame=action_target: self.run_context_menu_action(lambda: self.add_note_to_selection(frame)))
        menu.add_command(label=self.tr('context.respond', 'Respond'), state=note_action_state, command=lambda frame=action_target: self.run_context_menu_action(lambda: self.respond_to_note(frame)))
        menu.add_command(label=self.tr('context.remove_note', 'Remove note'), state=note_action_state, command=lambda frame=action_target: self.run_context_menu_action(lambda: self.remove_note(frame)))

    def replace_word_in_widget(self, action_target, start, end, replacement):
        target_widget = self.get_context_action_target_widget(action_target)
        if target_widget is None:
            return "break"
        try:
            target_widget.focus_force()
        except tk.TclError:
            pass
        self.set_last_active_editor_widget(target_widget)
        try:
            target_widget.delete(start, end)
            target_widget.insert(start, replacement)
            target_widget.mark_set(tk.INSERT, f"{start}+{len(replacement)}c")
            target_widget.tag_remove('sel', '1.0', tk.END)
        except tk.TclError:
            return "break"
        doc = self.get_doc_for_text_widget(target_widget)
        if doc:
            self.schedule_spellcheck(doc)
            self.update_status()
        return "break"

    def add_word_to_spellcheck_dictionary(self, word):
        checker = self.get_spell_checker()
        normalized = str(word or '').strip().lower()
        if not checker or not normalized:
            return "break"
        self.spellcheck_custom_words.add(normalized)
        try:
            checker.word_frequency.load_words([normalized])
        except Exception as exc:
            self.log_exception("add spellcheck dictionary word", exc)
            return "break"
        for doc in self.documents.values():
            self.schedule_spellcheck(doc)
        return "break"

    def get_spellcheck_suggestions(self, word):
        checker = self.get_spell_checker()
        normalized = str(word or '').strip().lower()
        if not checker or not normalized:
            return []
        try:
            candidates = checker.candidates(normalized) or set()
            correction = checker.correction(normalized)
        except Exception as exc:
            self.log_exception("spellcheck suggestions", exc)
            return []
        ordered = []
        if correction:
            ordered.append(correction)
        for candidate in sorted(candidates):
            if candidate not in ordered:
                ordered.append(candidate)
            if len(ordered) >= 6:
                break
        return ordered[:6]

    def should_capitalize_spellcheck_suggestion(self, text_widget, start_index):
        if not isinstance(text_widget, tk.Text):
            return False
        try:
            leading_text = text_widget.get('1.0', start_index)
        except tk.TclError:
            return False
        if not leading_text:
            return True
        skip_chars = set('\'"`“”‘’()[]{}')
        for char in reversed(leading_text):
            if char.isspace():
                continue
            if char in skip_chars:
                continue
            return char in '.!?'
        return True

    def format_spellcheck_suggestions(self, suggestions, word_info=None):
        if not suggestions:
            return []
        capitalize_first = bool(word_info and word_info.get('capitalize_suggestions'))
        formatted = []
        seen = set()
        for suggestion in suggestions:
            display_value = str(suggestion or '')
            if capitalize_first and display_value:
                display_value = display_value[0].upper() + display_value[1:]
            dedupe_key = display_value.lower()
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            formatted.append(display_value)
        return formatted

    def get_misspelled_word_info_at_index(self, text_widget, index, doc=None):
        if not isinstance(text_widget, tk.Text):
            return None
        doc = doc or self.get_doc_for_text_widget(text_widget)
        if not self.doc_supports_spellcheck(doc):
            return None
        try:
            resolved_index = text_widget.index(index or tk.INSERT)
            line_start = text_widget.index(f"{resolved_index} linestart")
            line_text = text_widget.get(line_start, f"{resolved_index} lineend")
            line_column = int(resolved_index.split('.', 1)[1])
        except (tk.TclError, ValueError, IndexError):
            return None

        checker = self.get_spell_checker()
        if checker is None:
            return None

        for match in self.spellcheck_token_pattern.finditer(line_text):
            if not (match.start() <= line_column < match.end()):
                continue
            if not self.is_spellcheck_candidate_word(line_text, match):
                return None
            normalized = match.group(0).lower()
            if normalized in self.spellcheck_custom_words:
                return None
            try:
                if normalized not in checker.unknown([normalized]):
                    return None
            except Exception as exc:
                self.log_exception("spellcheck word lookup", exc)
                return None
            return {
                'word': match.group(0),
                'normalized': normalized,
                'start': f"{line_start}+{match.start()}c",
                'end': f"{line_start}+{match.end()}c",
                'capitalize_suggestions': self.should_capitalize_spellcheck_suggestion(
                    text_widget,
                    f"{line_start}+{match.start()}c"
                )
            }
        return None

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
        self.tab_context_menu_tab_id = None
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
        if self.markdown_preview_enabled.get():
            doc = self.compare_view
        else:
            doc = self.documents.get(self.compare_source_tab) if self.compare_source_tab else None
        return self.show_context_menu_for_doc(event, doc, select_tab=False)

    def show_context_menu_for_doc(self, event, doc, select_tab):
        if not doc:
            return "break"
        self.hide_diagnostic_tooltip()

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

        is_readonly_target = self.is_doc_text_readonly(doc)
        if is_readonly_target or doc.get('virtual_mode'):
            note_state = 'disabled'
        note_action_state = 'normal' if doc.get('context_note_tag') else 'disabled'
        self.rebuild_text_context_menu(
            doc,
            doc.get('frame') if select_tab else '__compare__',
            note_state=note_state,
            note_action_state=note_action_state,
            is_readonly_target=is_readonly_target,
            word_index=index
        )
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

        popup = self.create_popup_toplevel(self.root)
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
        try:
            popup.deiconify()
            popup.lift()
        except tk.TclError:
            pass
        popup.bind('<Escape>', lambda e, current=doc: self.hide_note_popup(current))
        doc['note_popup'] = popup

    def prompt_text_input(self, title, prompt, initialvalue="", parent=None):
        parent = parent or self.root
        self.hide_autocomplete_popup()
        self.autocomplete_suspended += 1
        dialog = self.create_toplevel(parent)
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

        tk.Button(button_row, text=self.tr('common.ok', 'OK'), width=10, command=submit).pack(side='right' if self.is_rtl_locale() else 'left')
        tk.Button(button_row, text=self.tr('common.cancel', 'Cancel'), width=10, command=cancel).pack(side='left' if self.is_rtl_locale() else 'right')

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

    def prompt_note_input(self, title, prompt, initialvalue="", parent=None):
        return self.prompt_text_input(title, prompt, initialvalue=initialvalue, parent=parent)

    def prompt_note_color(self, title=None, initialvalue='yellow', parent=None):
        parent = parent or self.root
        self.hide_autocomplete_popup()
        self.autocomplete_suspended += 1
        dialog = self.create_toplevel(parent)
        dialog.title(title or self.tr('note.color.title', 'Note Color'))
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg='#f0f0f0', padx=14, pady=12)

        result = {'value': None}
        color_var = tk.StringVar(value=self.normalize_note_color(initialvalue))

        tk.Label(
            dialog,
            text=self.tr('note.color.prompt', 'Color:'),
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

        tk.Button(button_row, text=self.tr('common.ok', 'OK'), width=10, command=submit).pack(side='right' if self.is_rtl_locale() else 'left')
        tk.Button(button_row, text=self.tr('common.cancel', 'Cancel'), width=10, command=cancel).pack(side='left' if self.is_rtl_locale() else 'right')

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
            messagebox.showinfo(
                self.tr('menu.file.export_notes', 'Export Notes'),
                self.tr('export.notes.none', 'There are no notes to export in this tab.'),
                parent=self.root
            )
            return "break"
        initial_name = f"{self.get_doc_name(doc['frame'])}-notes.json"
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr('menu.file.export_notes', 'Export Notes'),
            defaultextension=".json",
            initialfile=initial_name,
            filetypes=[
                (self.tr('filetype.json', 'JSON'), "*.json"),
                (self.tr('filetype.markdown', 'Markdown'), "*.md"),
                (self.tr('filetype.all_files', 'All Files'), "*.*")
            ]
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
                markdown_parts = [
                    f"# {self.tr('export.notes.markdown.title', 'Notes Export for {doc_name}', doc_name=self.get_doc_name(doc['frame']))}\n"
                ]
                range_label = self.tr('export.notes.markdown.range', 'Range')
                author_label = self.tr('note.popup.author', 'Author')
                color_label = self.tr('export.notes.markdown.color', 'Color')
                created_label = self.tr('note.popup.created', 'Created')
                selected_text_heading = self.tr('export.notes.markdown.selected_text_heading', 'Selected Text')
                code_note_heading = self.tr('export.notes.markdown.code_note_heading', 'Code Note')
                responses_heading = self.tr('note.popup.responses', 'Responses')
                unknown_label = self.tr('common.unknown', 'Unknown')
                for row in note_rows:
                    markdown_parts.append(
                        f"\n## {self.tr('export.notes.markdown.note_heading', 'Note {note_id}', note_id=row['id'])}\n"
                    )
                    markdown_parts.append(
                        f"\n- {range_label}: `{row['selection_start']}` to `{row['selection_end']}`\n"
                    )
                    if row['author']:
                        markdown_parts.append(f"- {author_label}: {row['author']}\n")
                    if row['color']:
                        markdown_parts.append(f"- {color_label}: {row['color']}\n")
                    if row['created_at']:
                        markdown_parts.append(f"- {created_label}: {row['created_at']}\n")
                    markdown_parts.append(f"\n### {selected_text_heading}\n\n```\n")
                    markdown_parts.append(row['selected_text'])
                    markdown_parts.append(f"\n```\n\n### {code_note_heading}\n\n")
                    markdown_parts.append(row['note'] or '')
                    markdown_parts.append("\n")
                    responses = row.get('responses') or []
                    if responses:
                        markdown_parts.append(f"\n### {responses_heading}\n")
                        for response in responses:
                            author = response.get('author_label') or response.get('author_id') or unknown_label
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
            messagebox.showinfo(
                self.tr('menu.file.export_notes', 'Export Notes'),
                self.tr('export.notes.saved', 'Notes exported to:\n{output_path}', output_path=output_path),
                parent=self.root
            )
        except PermissionError as exc:
            self.show_filesystem_error(self.tr('menu.file.export_notes', 'Export Notes'), output_path, exc)
        except OSError as exc:
            self.log_exception("export notes report", exc)
            self.show_filesystem_error(self.tr('menu.file.export_notes', 'Export Notes'), output_path, exc)
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
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)
        except OSError as exc:
            self.log_exception("write doc notes payload", exc)
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)
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
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)
        except OSError as exc:
            self.log_exception("sync single note to sidecar", exc)
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)

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
        if not doc or not doc.get('file_path') or doc.get('virtual_mode') or doc.get('preview_mode'):
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
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)
        except OSError as exc:
            self.log_exception("persist doc notes", exc)
            self.show_filesystem_error(self.tr('code_notes.title', 'Code Notes'), sidecar_path, exc)
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
        keep_loading_state = False
        self.begin_doc_load(doc)
        try:
            self.cancel_doc_background_index(doc)
            self.close_doc_load_progress(doc)
            self.reset_virtual_backing_store(doc, remove_files=True)
            self.invalidate_minimap_cache(doc)
            self.clear_fold_tags(doc)
            doc['diagnostics'] = []
            doc['encrypted_file'] = False
            doc['encryption_header'] = None
            doc['encryption_key'] = None
            doc['file_size_bytes'] = 0
            doc['background_loading'] = False
            doc['background_load_kind'] = None
            doc['background_load_file_path'] = None
            doc['background_load_token'] = None
            doc['background_index_future'] = None
            doc['background_index_token'] = None
            doc['background_index_active'] = False
            doc['background_bytes_loaded'] = 0
            doc['background_bytes_total'] = 0
            doc['background_lines_loaded'] = 1
            doc.pop('pending_insert_content', None)
            doc['pending_insert_offset'] = 0
            doc['pending_insert_batch_count'] = 0
            doc['line_starts'] = None
            doc['total_file_lines'] = 1
            doc['window_start_line'] = 1
            doc['window_end_line'] = 1
            doc['last_virtual_line'] = 1
            doc['last_virtual_col'] = 0
            doc['pending_virtual_target_line'] = None

            text = doc['text']
            text.configure(state='normal')
            text.delete('1.0', tk.END)

            file_size = os.path.getsize(file_path)
            doc['file_size_bytes'] = file_size

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
                keep_loading_state = True
                return True
            else:
                if not doc['encrypted_file'] and file_size >= self.large_file_threshold_bytes:
                    self.start_background_text_load(doc, file_path)
                    self.refresh_tab_title(doc['frame'])
                    if str(doc['frame']) == self.notebook.select():
                        self.update_status()
                    keep_loading_state = True
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

            text.edit_modified(False)
            text.mark_set(tk.INSERT, '1.0')
            text.tag_remove('sel', '1.0', tk.END)
            text.see('1.0')
            doc['last_insert_index'] = '1.0'
            doc['last_yview'] = 0.0
            doc['last_xview'] = 0.0
            doc['symbol_cache_signature'] = None
            doc['symbol_cache'] = None
            self.update_doc_file_signature(doc)
            self.configure_syntax_highlighting(doc['frame'])
            self.restore_doc_notes(doc)
            self.register_doc_for_shared_notes(doc)
            self.invalidate_fold_regions(doc)
            self.schedule_diagnostics(doc)
            self.update_line_number_gutter(doc)
            self.schedule_minimap_refresh(doc)
            self.refresh_tab_title(doc['frame'])
            if self.compare_active and self.compare_source_tab == str(doc['frame']):
                self.refresh_compare_panel()
            if self.markdown_preview_enabled.get() and str(doc['frame']) == self.notebook.select():
                self.schedule_markdown_preview_refresh()
            if str(doc['frame']) == self.notebook.select():
                self.update_status()
            return True
        except EncryptedFileOpenCancelled:
            self.trace_startup(f"load_content_into_doc cancelled={file_path}")
            return False
        except RuntimeError as exc:
            self.log_exception("load content into doc", exc)
            messagebox.showerror(self.tr('file.open_failed_title', 'Open Failed'), str(exc), parent=self.root)
            return False
        except (OSError, UnicodeDecodeError, ValueError) as exc:
            self.log_exception("load content into doc", exc)
            messagebox.showerror(
                self.tr('file.open_failed_title', 'Open Failed'),
                self.tr(
                    'file.open_failed_message',
                    'Notepad-X could not open:\n{file_path}\n\n{error_detail}',
                    file_path=file_path,
                    error_detail=exc
                ),
                parent=self.root
            )
            return False
        finally:
            if not keep_loading_state:
                self.end_doc_load(doc)

    def get_session_state(self):
        current_doc = self.get_current_doc()
        selected_tab_id = self.notebook.select()
        selected_tab_doc = self.documents.get(str(selected_tab_id)) if selected_tab_id else None
        selected_file = current_doc['file_path'] if current_doc and current_doc['file_path'] and not current_doc.get('is_remote') else None
        window_width, window_height, window_state = self.get_window_layout_snapshot()
        compare_file = None
        compare_base_file = None
        if self.compare_active and self.compare_source_tab:
            compare_doc = self.documents.get(self.compare_source_tab)
            if compare_doc and not compare_doc.get('is_remote') and compare_doc.get('file_path') and os.path.exists(compare_doc['file_path']):
                compare_file = compare_doc['file_path']
            if selected_tab_doc and not selected_tab_doc.get('is_remote') and selected_tab_doc.get('file_path') and os.path.exists(selected_tab_doc['file_path']):
                compare_base_file = selected_tab_doc['file_path']
        open_files = []

        for tab_id in self.notebook.tabs():
            doc = self.documents.get(str(tab_id))
            if not doc:
                continue
            if doc.get('is_remote'):
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
            'find_history': self.sanitize_search_history_entries(self.find_history),
            'find_in_history': self.sanitize_search_history_entries(self.find_in_history),
            'command_history': self.sanitize_command_history_entries(self.command_history),
            'command_panel_height': int(self.command_panel_height),
            'window_width': int(window_width),
            'window_height': int(window_height),
            'window_state': window_state,
            'closed_session_files': [
                path for path in self.closed_session_files
                if isinstance(path, str)
            ][:self.max_session_files],
            'sound_enabled': bool(self.sound_enabled.get()),
            'status_bar_enabled': bool(self.status_bar_enabled.get()),
            'numbered_lines_enabled': bool(self.numbered_lines_enabled.get()),
            'autocomplete_enabled': bool(self.autocomplete_enabled.get()),
            'spell_check_enabled': bool(self.spell_check_enabled.get()),
            'auto_pair_enabled': bool(self.auto_pair_enabled.get()),
            'compare_multi_edit_enabled': bool(self.compare_multi_edit_enabled.get()),
            'markdown_preview_enabled': bool(self.markdown_preview_enabled.get()),
            'sync_page_navigation_enabled': bool(self.sync_page_navigation_enabled.get()),
            'edit_with_shell_enabled': bool(self.edit_with_shell_enabled.get()),
            'minimap_enabled': bool(self.minimap_enabled.get()),
            'breadcrumbs_enabled': bool(self.breadcrumbs_enabled.get()),
            'diagnostics_enabled': bool(self.diagnostics_enabled.get()),
            'autosave_enabled': bool(self.autosave_enabled.get()),
            'current_font_size': int(self.current_font_size),
            'syntax_theme': self.syntax_theme.get(),
            'locale_code': self.locale_code,
            'compare_file': compare_file,
            'compare_base_file': compare_base_file,
            'hotkey_overrides': {
                action_id: value for action_id, value in self.hotkey_overrides.items()
                if action_id in self.get_hotkey_definitions()
            },
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
            if doc.get('large_file_mode'):
                continue
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
        self.invalidate_minimap_cache(doc)
        doc['diagnostics'] = []
        doc['suspend_modified_events'] = True
        try:
            text.delete('1.0', tk.END)
            text.insert('1.0', content)
            text.edit_modified(bool(modified))
        finally:
            doc['suspend_modified_events'] = False
        self.refresh_tab_title(doc['frame'])
        self.update_doc_file_signature(doc)
        self.invalidate_fold_regions(doc)
        self.schedule_diagnostics(doc)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)

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
        if not messagebox.askyesno(
            self.tr('recover.tabs_title', 'Recover Tabs'),
            self.tr('recover.tabs_message', 'Notepad-X found unsaved tabs from a previous crash. Restore them?'),
            parent=self.root
        ):
            return
        current_doc = self.get_current_doc()
        if current_doc and not current_doc.get('file_path') and not current_doc['text'].edit_modified() and not current_doc['text'].get('1.0', 'end-1c').strip():
            self.notebook.forget(current_doc['frame'])
            self.documents.pop(str(current_doc['frame']), None)
        selected_recovery_key = recovery.get('selected_recovery_key')
        selected_tab = None
        for recovered_tab in recovery_tabs:
            if self._shutdown_requested:
                break
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
                        if self._shutdown_requested:
                            break
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
        if self._shutdown_requested:
            return
        if selected_tab is None and self.documents:
            selected_tab = next(iter(self.documents.values()))['frame']
        if selected_tab is not None:
            self.notebook.select(selected_tab)
            self.set_active_document(selected_tab)

    def schedule_startup_recovery_restore(self):
        if self.startup_recovery_restore_scheduled or self.isolated_session:
            return
        self.startup_recovery_restore_scheduled = True
        self.schedule_when_root_ready(self.run_startup_recovery_restore)

    def run_startup_recovery_restore(self):
        self.startup_recovery_restore_scheduled = False
        if self._shutdown_requested:
            return
        self.restore_recovery_state()

    def restore_session_tabs(self, open_files, selected_file=None, compare_base_file=None, compare_file=None):
        if self._shutdown_requested:
            return

        open_files = list(open_files or [])
        self.trace_startup(f"restore_session_tabs open_files={open_files}")
        if not open_files:
            if self.markdown_preview_enabled.get():
                self.schedule_markdown_preview_refresh(immediate=True)
            self.restore_recovery_state()
            return

        fallback_doc = self.get_current_doc()
        if fallback_doc:
            if fallback_doc.get('file_path') or fallback_doc['text'].edit_modified():
                fallback_doc = None

        restored_tabs = {}
        for file_path in open_files:
            if self._shutdown_requested:
                break
            if fallback_doc and not restored_tabs and str(fallback_doc['frame']) in self.documents:
                doc = fallback_doc
                doc['file_path'] = file_path
                doc['background_open_new_tab'] = False
                if not self.load_content_into_doc(doc, file_path):
                    self.cleanup_failed_file_open(doc)
                    continue
                doc['background_open_new_tab'] = False
                self.notebook.tab(doc['frame'], text=self.get_doc_title(doc['frame']))
                restored_tabs[file_path] = doc['frame']
                continue
            tab_id = self.create_tab(file_path=file_path, select=False)
            doc = self.documents[str(tab_id)]
            if not self.load_content_into_doc(doc, file_path):
                if self._shutdown_requested:
                    break
                try:
                    self.notebook.forget(doc['frame'])
                except tk.TclError:
                    pass
                self.documents.pop(str(tab_id), None)
                continue
            restored_tabs[file_path] = doc['frame']

        primary_file = selected_file
        if isinstance(compare_base_file, str) and compare_base_file in restored_tabs:
            primary_file = compare_base_file

        selected_tab = restored_tabs.get(primary_file)
        if selected_tab is None and restored_tabs:
            selected_tab = next(iter(restored_tabs.values()))
        if selected_tab is None and fallback_doc and str(fallback_doc['frame']) in self.documents:
            selected_tab = fallback_doc['frame']
            self.refresh_tab_title(selected_tab)

        if selected_tab is not None:
            self.notebook.select(selected_tab)
            self.set_active_document(selected_tab)
        if self._shutdown_requested:
            return
        if self.markdown_preview_enabled.get():
            self.schedule_markdown_preview_refresh(immediate=True)

        compare_tab = restored_tabs.get(compare_file) if isinstance(compare_file, str) else None
        if compare_tab is not None and compare_tab != selected_tab:
            compare_doc = self.documents.get(str(compare_tab))
            if compare_doc:
                self.start_inline_compare(compare_doc)
        self.restore_recovery_state()

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
            self.schedule_startup_recovery_restore()
            return

        raw_session = self.read_json_file(self.session_path, "restore session", None)
        has_saved_window_layout = isinstance(raw_session, dict) and any(
            key in raw_session for key in ('window_width', 'window_height', 'window_state')
        )
        session = self.sanitize_session_payload(raw_session)
        if session is None:
            self.schedule_startup_recovery_restore()
            return

        open_files = list(session.get('open_files', []))
        self.closed_session_files = set(session.get('closed_session_files', []))
        open_files = [path for path in open_files if path not in self.closed_session_files]
        self.recent_files = list(session.get('recent_files', []))[:self.max_recent_files]
        self.find_history = self.sanitize_search_history_entries(session.get('find_history', []))
        self.find_in_history = self.sanitize_search_history_entries(session.get('find_in_history', []))
        self.command_history = self.sanitize_command_history_entries(session.get('command_history', []))
        self.command_panel_height = max(
            self.command_panel_min_height,
            min(self.command_panel_max_height, int(session.get('command_panel_height', self.command_panel_default_height)))
        )
        self.hotkey_overrides = dict(session.get('hotkey_overrides', {}))
        self.window_layout_restored = False
        if has_saved_window_layout:
            self.window_layout_restored = self.apply_saved_window_layout(
                session.get('window_width', self.default_window_width),
                session.get('window_height', self.default_window_height),
                session.get('window_state', 'normal')
            )
        self.sound_enabled.set(bool(session.get('sound_enabled', True)))
        self.status_bar_enabled.set(bool(session.get('status_bar_enabled', True)))
        self.numbered_lines_enabled.set(bool(session.get('numbered_lines_enabled', True)))
        self.autocomplete_enabled.set(bool(session.get('autocomplete_enabled', True)))
        self.spell_check_enabled.set(bool(session.get('spell_check_enabled', SpellChecker is not None)) and self.ensure_spellcheck_available(notify=False))
        self.auto_pair_enabled.set(bool(session.get('auto_pair_enabled', True)))
        self.compare_multi_edit_enabled.set(bool(session.get('compare_multi_edit_enabled', False)))
        self.markdown_preview_enabled.set(bool(session.get('markdown_preview_enabled', False)))
        self.sync_page_navigation_enabled.set(bool(session.get('sync_page_navigation_enabled', False)))
        self.minimap_enabled.set(bool(session.get('minimap_enabled', True)))
        self.breadcrumbs_enabled.set(bool(session.get('breadcrumbs_enabled', True)))
        self.diagnostics_enabled.set(bool(session.get('diagnostics_enabled', True)))
        self.autosave_enabled.set(bool(session.get('autosave_enabled', True)))
        saved_edit_with_shell = bool(session.get('edit_with_shell_enabled', False))
        shell_registered = self.is_edit_with_shell_registered()
        self.edit_with_shell_enabled.set(saved_edit_with_shell or shell_registered)
        self.apply_locale(session.get('locale_code', self.locale_code), persist=False)
        if self.edit_with_shell_enabled.get():
            self.sync_edit_with_shell_menu(show_errors=False)
        saved_font_size = session.get('current_font_size', self.base_font_size)
        try:
            self.current_font_size = max(self.min_font_size, min(self.max_font_size, int(saved_font_size)))
        except (TypeError, ValueError):
            self.current_font_size = self.base_font_size
        saved_theme = str(session.get('syntax_theme', 'Default'))
        if saved_theme not in self.get_available_syntax_theme_names():
            saved_theme = 'Default'
        self.syntax_theme.set(saved_theme)
        self.create_menu()
        for doc in self.documents.values():
            self.apply_syntax_tag_colors(doc['text'])
            self.apply_text_theme_effect(doc)
            self.update_line_number_gutter(doc)
        if getattr(self, 'compare_view', None):
            self.apply_syntax_tag_colors(self.compare_text)
            self.apply_text_theme_effect(self.compare_view)
            self.update_line_number_gutter(self.compare_view)
        if self.status_bar_enabled.get():
            self.status_frame.grid()
        else:
            self.status_frame.grid_remove()
        self.toggle_numbered_lines()
        self.toggle_minimap()
        self.toggle_breadcrumbs()
        self.set_command_panel_height(self.command_panel_height, persist=False)
        self.refresh_command_history_list()
        self.refresh_recent_files_menu()
        if not open_files:
            if self.markdown_preview_enabled.get():
                self.schedule_markdown_preview_refresh()
            self.schedule_startup_recovery_restore()
            return

        selected_file = session.get('selected_file')
        compare_base_file = session.get('compare_base_file')
        compare_file = session.get('compare_file')
        self.trace_startup(
            f"restore_session open_files={open_files} selected={selected_file} "
            f"compare_base={compare_base_file} compare={compare_file}"
        )
        # Wait until the main window is viewable before reopening files so encrypted-file prompts can render reliably.
        self.schedule_when_root_ready(
            lambda files=tuple(open_files), selected=selected_file, base=compare_base_file, compare=compare_file:
            self.restore_session_tabs(files, selected, base, compare)
        )

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
            self.update_breadcrumbs()
            return
        self.text = doc['text']
        self.current_file = doc['file_path']
        self.syntax_mode_selection.set(doc.get('syntax_override') or 'auto')
        self.restore_doc_view_state(doc)
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
        self.schedule_diagnostics(doc)
        self.update_window_title()
        if self.markdown_preview_enabled.get():
            self.markdown_preview_source_tab = str(tab_id)
            self.schedule_markdown_preview_refresh(immediate=True)
        self.update_status()

    def on_tab_changed(self, event=None):
        self.hide_autocomplete_popup()
        self.hide_search_history_popup()
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
        if self.find_panel_visible or self.replace_panel_visible:
            query = self.get_visible_find_query()
            self.update_find_match_summary(query, allow_short_query=True)
            if query and len(query) >= self.live_find_min_chars:
                self.highlight_live_find_matches(query)
            else:
                self.clear_find_highlights()

    def get_doc_name(self, tab_id):
        doc = self.documents[str(tab_id)]
        if doc.get('display_name'):
            return os.path.basename(str(doc.get('display_name')))
        if doc['file_path']:
            return os.path.basename(doc['file_path'])
        return doc['untitled_name']

    def get_doc_title(self, tab_id):
        doc = self.documents[str(tab_id)]
        title = self.get_doc_name(tab_id)
        if self.doc_has_unsaved_changes(doc):
            title += " *"
        return title

    def refresh_tab_title(self, tab_id):
        self.notebook.tab(tab_id, text=self.get_doc_title(tab_id))
        if str(tab_id) == self.notebook.select():
            self.update_window_title()
        if self.compare_active and self.compare_source_tab == str(tab_id):
            self.refresh_compare_header()
        if self.markdown_preview_enabled.get() and self.markdown_preview_source_tab == str(tab_id):
            self.schedule_markdown_preview_refresh()
        self.save_session()

    def update_window_title(self):
        doc = self.get_current_doc()
        if not doc:
            self.root.title(self.app_name)
            return
        title = self.get_doc_name(doc['frame'])
        if self.doc_has_unsaved_changes(doc):
            title += " *"
        self.root.title(f"{self.app_name} - {title}")

    def on_text_modified(self, tab_id):
        if str(tab_id) not in self.documents:
            return
        doc = self.documents[str(tab_id)]
        if doc.get('loading_file') or doc.get('suspend_modified_events'):
            doc['text'].edit_modified(False)
            return
        self.remember_doc_view_state(doc)
        doc['symbol_cache_signature'] = None
        doc['symbol_cache'] = None
        if self.is_virtual_editable(doc):
            self.refresh_tab_title(tab_id)
            self.update_line_number_gutter(doc)
            if str(tab_id) == self.notebook.select():
                self.update_status()
            return
        is_large_file = bool(doc.get('large_file_mode'))
        if is_large_file:
            try:
                doc['total_file_lines'] = max(1, int(doc['text'].index('end-1c').split('.')[0]))
            except (tk.TclError, TypeError, ValueError):
                pass
        else:
            self.invalidate_fold_regions(doc)
            self.invalidate_collapsed_fold_regions(doc)
            self.invalidate_minimap_cache(doc)
        if not doc.get('file_path'):
            self.configure_syntax_highlighting(tab_id)
        if doc.get('syntax_mode') and doc.get('syntax_mode') != 'python':
            self.schedule_syntax_highlight(doc)
        self.schedule_text_theme_effect(doc)
        self.schedule_spellcheck(doc)
        self.schedule_diagnostics(doc)
        self.schedule_doc_autosave(doc)
        if self.markdown_preview_enabled.get() and self.markdown_preview_source_tab == str(tab_id):
            self.schedule_markdown_preview_refresh()
        self.update_line_number_gutter(doc)
        self.schedule_minimap_refresh(doc)
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

    def get_tab_id_at_position(self, x, y):
        try:
            tab_index = self.notebook.index(f"@{int(x)},{int(y)}")
        except (tk.TclError, ValueError, TypeError):
            return None
        tab_ids = list(self.notebook.tabs())
        if 0 <= tab_index < len(tab_ids):
            return str(tab_ids[tab_index])
        return None

    def select_tab_by_id(self, tab_id, focus_editor=False):
        if not tab_id or str(tab_id) not in self.documents:
            return None
        try:
            self.notebook.select(tab_id)
        except tk.TclError:
            return None
        self.set_active_document(tab_id)
        if focus_editor:
            doc = self.documents.get(str(tab_id))
            text_widget = doc.get('text') if doc else None
            if text_widget and text_widget.winfo_exists():
                try:
                    text_widget.focus_set()
                except tk.TclError:
                    pass
        return self.documents.get(str(tab_id))

    def get_tab_close_candidate_ids(self, target_tab_id=None, mode='single'):
        tab_ids = [str(tab_id) for tab_id in self.notebook.tabs() if str(tab_id) in self.documents]
        if not tab_ids:
            return []
        if target_tab_id is None or target_tab_id not in tab_ids:
            target_tab_id = str(self.notebook.select()) if self.notebook.select() in tab_ids else tab_ids[0]
        target_index = tab_ids.index(target_tab_id)
        if mode == 'others':
            return [tab_id for tab_id in tab_ids if tab_id != target_tab_id]
        if mode == 'left':
            return tab_ids[:target_index]
        if mode == 'right':
            return tab_ids[target_index + 1:]
        if mode == 'all':
            return list(tab_ids)
        return [target_tab_id]

    def close_tab_by_id(self, tab_id, recreate_if_empty=True, confirm=True):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return False
        if (
            len(self.documents) == 1 and
            not doc.get('file_path') and
            doc.get('untitled_name') == 'Untitled 1'
        ):
            return False
        if confirm and not self.confirm_close_tab(doc):
            return False

        closed_file_path = doc.get('file_path')
        if closed_file_path and not doc.get('is_remote'):
            self.closed_session_files.add(closed_file_path)

        if self.compare_active and self.compare_source_tab == str(doc['frame']):
            self.close_compare_panel()

        self.unregister_doc_from_shared_notes(doc)
        tab_key = str(doc['frame'])
        try:
            self.notebook.forget(doc['frame'])
        except tk.TclError:
            return False
        self.documents.pop(tab_key, None)
        self.dispose_doc_resources(doc)

        if not self.documents:
            if recreate_if_empty:
                new_tab = self.create_tab()
                self.set_active_document(new_tab)
                self.current_file = None
            else:
                self.save_session()
                self.root.quit()
                return True
        else:
            try:
                selected_tab_id = self.notebook.select()
            except tk.TclError:
                selected_tab_id = None
            if selected_tab_id:
                self.set_active_document(selected_tab_id)
        if not any(existing_doc.get('file_path') for existing_doc in self.documents.values()):
            self.current_file = None
        self.cpu_used_percent = self.get_process_cpu_usage_percent()
        self.memory_used_mb = self.get_memory_usage_mb()
        self.update_status()
        self.schedule_memory_trim()
        self.save_session()
        return True

    def close_tab_group(self, target_tab_id=None, mode='single', recreate_if_empty=True):
        tab_ids = self.get_tab_close_candidate_ids(target_tab_id=target_tab_id, mode=mode)
        if not tab_ids:
            return "break"
        target_tab_id = str(target_tab_id) if target_tab_id is not None else None
        for tab_id in list(tab_ids):
            if str(tab_id) not in self.documents:
                continue
            if not self.close_tab_by_id(tab_id, recreate_if_empty=recreate_if_empty, confirm=True):
                return "break"
        if target_tab_id and target_tab_id in self.documents:
            self.select_tab_by_id(target_tab_id)
        return "break"

    def copy_tab_file_name(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        file_name = self.get_doc_name(doc['frame'])
        return self.copy_text_to_clipboard(file_name)

    def copy_tab_file_path(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        return self.copy_text_to_clipboard(self.get_doc_display_path(doc))

    def reveal_tab_in_file_manager(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        file_path = doc.get('file_path')
        if not file_path or not os.path.exists(file_path):
            messagebox.showinfo(
                self.tr('tab.menu.reveal_title', 'Reveal in File Explorer'),
                self.tr('tab.menu.reveal_unavailable', 'This tab does not have a local file path to reveal yet.'),
                parent=self.root
            )
            return "break"
        try:
            if self.is_windows:
                subprocess.Popen(
                    ['explorer.exe', '/select,', file_path],
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
            elif self.is_linux:
                target_dir = os.path.dirname(file_path) or file_path
                subprocess.Popen(['xdg-open', target_dir], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                webbrowser.open(os.path.dirname(file_path) or file_path)
        except OSError as exc:
            self.log_exception('reveal tab in file manager', exc)
            messagebox.showerror(self.tr('tab.menu.reveal_title', 'Reveal in File Explorer'), str(exc), parent=self.root)
        return "break"

    def build_new_window_launch_args(self, file_path):
        if getattr(sys, 'frozen', False):
            return [os.path.abspath(sys.executable), '--isolated', file_path]
        return [sys.executable, os.path.abspath(__file__), '--isolated', file_path]

    def move_tab_to_new_window(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        self.select_tab_by_id(tab_id)
        doc = self.documents.get(str(tab_id))
        if not doc:
            return "break"
        if doc.get('is_remote'):
            messagebox.showinfo(
                self.tr('tab.menu.new_window_title', 'Open in New Window'),
                self.tr('tab.menu.new_window_remote', 'Remote tabs cannot be moved to a separate window yet.'),
                parent=self.root
            )
            return "break"
        if doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
            local_path = doc.get('file_path')
            if not local_path or not os.path.exists(local_path):
                messagebox.showinfo(
                    self.tr('tab.menu.new_window_title', 'Open in New Window'),
                    self.tr('tab.menu.new_window_unavailable', 'Only saved local tabs can be moved to a separate window.'),
                    parent=self.root
                )
                return "break"
        if not doc.get('file_path'):
            if not self.save_as():
                return "break"
            doc = self.documents.get(str(tab_id))
            if not doc or not doc.get('file_path'):
                return "break"
        elif self.doc_has_unsaved_changes(doc) and not self.is_doc_text_readonly(doc):
            if not self.save_document_content(doc, autosave=False, show_errors=True, update_recent=True):
                return "break"

        file_path = doc.get('file_path')
        if not file_path or not os.path.exists(file_path):
            messagebox.showinfo(
                self.tr('tab.menu.new_window_title', 'Open in New Window'),
                self.tr('tab.menu.new_window_unavailable', 'Only saved local tabs can be moved to a separate window.'),
                parent=self.root
            )
            return "break"
        command_args = self.build_new_window_launch_args(file_path)
        try:
            subprocess.Popen(
                command_args,
                cwd=self.app_dir,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0) if not getattr(sys, 'frozen', False) else 0
            )
        except OSError as exc:
            self.log_exception('move tab to new window', exc)
            messagebox.showerror(self.tr('tab.menu.new_window_title', 'Open in New Window'), str(exc), parent=self.root)
            return "break"
        self.close_tab_by_id(tab_id, recreate_if_empty=True, confirm=False)
        return "break"

    def create_tab_context_menu(self):
        menu = tk.Menu(self.root, tearoff=0)
        self.tab_context_menu = menu
        return menu

    def rebuild_tab_context_menu(self, tab_id):
        doc = self.documents.get(str(tab_id))
        if not doc:
            return None
        menu = self.tab_context_menu or self.create_tab_context_menu()
        menu.delete(0, tk.END)
        save_disabled = bool(doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)))
        has_local_path = bool(doc.get('file_path') and os.path.exists(doc.get('file_path')))

        menu.add_command(label=self.tr('tab.menu.save', 'Save'), state='disabled' if save_disabled else 'normal', command=lambda current=tab_id: self.run_tab_menu_action(current, self.save))
        menu.add_command(label=self.tr('tab.menu.save_as', 'Save As...'), state='disabled' if save_disabled else 'normal', command=lambda current=tab_id: self.run_tab_menu_action(current, self.save_as))
        menu.add_command(label=self.tr('tab.menu.save_copy_as', 'Save Copy As...'), command=lambda current=tab_id: self.run_tab_menu_action(current, self.save_copy_as))
        menu.add_separator()
        menu.add_command(label=self.tr('tab.menu.copy_name', 'Copy File Name'), command=lambda current=tab_id: self.run_tab_menu_action(current, lambda: self.copy_tab_file_name(current)))
        menu.add_command(label=self.tr('tab.menu.copy_path', 'Copy Full Path'), command=lambda current=tab_id: self.run_tab_menu_action(current, lambda: self.copy_tab_file_path(current)))
        menu.add_command(label=self.tr('tab.menu.reveal', 'Reveal in File Explorer'), state='normal' if has_local_path else 'disabled', command=lambda current=tab_id: self.run_tab_menu_action(current, lambda: self.reveal_tab_in_file_manager(current)))
        menu.add_separator()
        menu.add_command(label=self.tr('tab.menu.new_window', 'Open in New Notepad-X Window'), command=lambda current=tab_id: self.run_tab_menu_action(current, lambda: self.move_tab_to_new_window(current)))
        menu.add_separator()
        menu.add_command(label=self.tr('tab.menu.close', 'Close Tab'), command=lambda current=tab_id: self.run_tab_menu_action(current, lambda: self.close_tab_group(current, mode='single')))
        return menu

    def run_tab_menu_action(self, tab_id, callback):
        self.dismiss_context_menu()
        if callback is None:
            return "break"
        current = self.select_tab_by_id(tab_id)
        if current is None:
            return "break"
        try:
            self.root.after_idle(callback)
        except tk.TclError:
            return callback()
        return "break"

    def show_tab_context_menu(self, event):
        tab_id = self.get_tab_id_at_position(getattr(event, 'x', 0), getattr(event, 'y', 0))
        if not tab_id:
            self.dismiss_context_menu()
            return None
        self.select_tab_by_id(tab_id)
        menu = self.rebuild_tab_context_menu(tab_id)
        if menu is None:
            return None
        self.dismiss_context_menu()
        self.active_context_menu = menu
        self.tab_context_menu_tab_id = str(tab_id)
        self.context_menu_posted_at = time.monotonic()
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                menu.grab_release()
            except tk.TclError:
                pass
        return "break"

    def close_current_tab(self, event=None, recreate_if_empty=True):
        doc = self.get_current_doc()
        if not doc:
            return "break"
        self.close_tab_by_id(doc['frame'], recreate_if_empty=recreate_if_empty, confirm=True)
        return "break"

    def dispose_doc_resources(self, doc):
        if not doc:
            return
        self.cancel_doc_background_index(doc)
        self.reset_virtual_backing_store(doc, remove_files=True)
        self.close_doc_load_progress(doc)
        self.cancel_text_theme_effect_job(doc)
        self.cancel_doc_autosave(doc)
        syntax_job = doc.get('syntax_job')
        if syntax_job:
            try:
                self.root.after_cancel(syntax_job)
            except tk.TclError:
                pass
            doc['syntax_job'] = None
        diagnostic_job = doc.get('diagnostic_job')
        if diagnostic_job:
            try:
                self.root.after_cancel(diagnostic_job)
            except tk.TclError:
                pass
            doc['diagnostic_job'] = None
        minimap_job = doc.get('minimap_job')
        if minimap_job:
            try:
                self.root.after_cancel(minimap_job)
            except tk.TclError:
                pass
            doc['minimap_job'] = None
        spellcheck_job = doc.get('spellcheck_job')
        if spellcheck_job:
            try:
                self.root.after_cancel(spellcheck_job)
            except tk.TclError:
                pass
            doc['spellcheck_job'] = None

        text_widget = doc.get('text')
        if text_widget:
            try:
                if text_widget.winfo_exists():
                    text_widget.tag_remove(self.spellcheck_tag, '1.0', tk.END)
                    text_widget.configure(undo=False, autoseparators=False)
                    text_widget.edit_reset()
                    text_widget.delete('1.0', tk.END)
            except tk.TclError:
                pass

        doc['notes'] = {}
        doc['context_note_tag'] = None
        doc['percolator'] = None
        doc['colorizer'] = None
        doc['line_starts'] = None
        doc['pending_insert_content'] = None
        doc['background_loading'] = False
        doc['background_load_kind'] = None
        doc['background_load_file_path'] = None
        doc['background_load_token'] = None
        doc['background_index_future'] = None
        doc['background_index_token'] = None
        doc['background_index_active'] = False
        doc['background_bytes_loaded'] = 0
        doc['background_bytes_total'] = 0
        doc['background_lines_loaded'] = 1

        frame = doc.get('frame')
        if frame:
            try:
                if frame.winfo_exists():
                    frame.destroy()
            except tk.TclError:
                pass

        gc.collect()

    # ─── Menu ────────────────────────────────────────────────────
    def create_menu(self):
        t = self.tr
        hk = self.get_hotkey_display
        self.menu = tk.Menu(self.root, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a', activeforeground='white')
        self.root.config(menu=self.menu)

        file_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.file', 'File'), menu=file_menu)
        file_menu.add_command(label=t('menu.file.open', 'Open'), command=self.open_file, accelerator=hk('open'))
        file_menu.add_command(label=t('menu.file.open_project', 'Open Project'), command=self.open_project, accelerator=hk('open_project'))
        file_menu.add_command(label='Open Remote (SSH)', command=self.open_remote_file_dialog, accelerator=hk('open_remote'))
        file_menu.add_command(label=t('menu.file.grab_git', 'Grab Git'), command=self.grab_git_project, accelerator=hk('grab_git'))
        self.recent_menu = tk.Menu(file_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                                   activebackground='#3a3a3a')
        file_menu.add_cascade(label=t('menu.file.recent', 'Recent'), menu=self.recent_menu)
        self.refresh_recent_files_menu()
        file_menu.add_command(label=t('menu.file.new_tab', 'New Tab'), command=self.new_tab, accelerator=hk('new_tab'))
        file_menu.add_command(label=t('menu.file.close_tab', 'Close Tab'), command=self.close_current_tab, accelerator=hk('close_tab'))
        file_menu.add_command(label=t('menu.file.save', 'Save'), command=self.save, accelerator=hk('save'))
        file_menu.add_command(label=t('menu.file.save_all', 'Save All'), command=self.save_all, accelerator=hk('save_all'))
        file_menu.add_command(label=t('menu.file.save_as', 'Save As'), command=self.save_as, accelerator=hk('save_as'))
        file_menu.add_command(label=t('save.copy_title', 'Save Copy As'), command=self.save_copy_as, accelerator=hk('save_copy_as'))
        file_menu.add_command(label=t('menu.file.save_and_run', 'Save and Run'), command=self.save_and_run, accelerator=hk('save_and_run'))
        file_menu.add_command(label=t('menu.file.save_as_encrypted', 'Save As Encrypted'), command=self.save_encrypted_copy, accelerator=hk('save_as_encrypted'))
        file_menu.add_command(label=t('menu.file.print', 'Print'), command=self.print_file, accelerator=hk('print'))
        file_menu.add_command(label=t('menu.file.export_notes', 'Export Notes'), command=self.export_notes_report, accelerator=hk('export_notes'))
        file_menu.add_separator()
        file_menu.add_command(label=t('menu.file.exit', 'Exit'), command=self.exit_app, accelerator=hk('exit'))

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
        edit_menu.add_command(label=t('menu.edit.find', 'Find'), command=self.show_find_panel, accelerator=hk('find'))
        edit_menu.add_command(label=t('menu.edit.find_next', 'Find Next'), command=self.find_next, accelerator=hk('find_next'))
        edit_menu.add_command(label=t('menu.edit.find_previous', 'Find Previous'), command=self.find_previous, accelerator=hk('find_previous'))
        edit_menu.add_command(label=t('menu.edit.replace', 'Replace'), command=self.show_replace_panel, accelerator=hk('replace'))
        edit_menu.add_command(label='Command Panel', command=self.show_command_panel, accelerator=hk('command_panel'))
        edit_menu.add_command(label='Jump to Symbol', command=self.show_symbol_navigator, accelerator=hk('jump_symbol'))
        edit_menu.add_command(label='Project Symbols', command=lambda: self.show_symbol_navigator(project_scope=True), accelerator=hk('project_symbols'))
        edit_menu.add_separator()
        edit_menu.add_command(label='Toggle Fold', command=self.toggle_fold_at_cursor, accelerator=hk('toggle_fold'))
        edit_menu.add_command(label='Collapse All Folds', command=self.collapse_all_folds, accelerator=hk('collapse_all_folds'))
        edit_menu.add_command(label='Expand All Folds', command=self.expand_all_folds, accelerator=hk('expand_all_folds'))
        edit_menu.add_separator()
        edit_menu.add_command(label=t('menu.edit.date', 'Date'), command=self.insert_date, accelerator=hk('date'))
        edit_menu.add_command(label=t('menu.edit.time_date', 'Time/Date'), command=self.insert_time_date, accelerator=hk('time_date'))
        edit_menu.add_command(label=t('menu.edit.font', 'Font'), command=self.show_font_dialog, accelerator=hk('font'))
        self.language_menu = tk.Menu(edit_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a', postcommand=self.refresh_language_menu)
        edit_menu.add_cascade(label=t('menu.edit.language', 'Language'), menu=self.language_menu)
        self.refresh_language_menu()
        view_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.view', 'View'), menu=view_menu)
        view_menu.add_command(label=t('menu.view.full_screen', 'Full Screen'), command=self.toggle_fullscreen, accelerator=hk('fullscreen'))
        view_menu.add_command(label=t('menu.view.switch_tab', 'Switch Tab'), command=self.switch_tab_right, accelerator=hk('switch_tab'))
        view_menu.add_command(label=t('menu.view.currently_editing', 'Currently Editing'), command=self.toggle_currently_editing_panel, accelerator=hk('currently_editing'))
        view_menu.add_command(label=t('menu.edit.cycle_notes', 'Cycle Notes'), command=self.goto_next_note, accelerator=hk('cycle_notes'))
        note_filter_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label=t('menu.edit.filter_notes', 'Filter Notes'), menu=note_filter_menu)
        note_filter_menu.add_radiobutton(label=t('note.filter.all', 'All'), variable=self.note_filter, value='all')
        note_filter_menu.add_radiobutton(label=t('note.filter.unread', 'Unread'), variable=self.note_filter, value='unread')
        note_filter_menu.add_radiobutton(label=t('note.filter.yellow', 'Yellow'), variable=self.note_filter, value='yellow')
        note_filter_menu.add_radiobutton(label=t('note.filter.green', 'Green'), variable=self.note_filter, value='green')
        note_filter_menu.add_radiobutton(label=t('note.filter.red', 'Red'), variable=self.note_filter, value='red')
        note_filter_menu.add_radiobutton(label=t('note.filter.blue', 'Light Blue'), variable=self.note_filter, value='blue')
        view_menu.add_command(label=t('menu.edit.goto_line', 'Go To Line'), command=self.goto_line_dialog, accelerator=hk('goto_line'))
        view_menu.add_command(label=t('menu.edit.top_of_document', 'Top of Document'), command=self.goto_document_start, accelerator=hk('top_of_document'))
        view_menu.add_command(label=t('menu.edit.bottom_of_document', 'Bottom of Document'), command=self.goto_document_end, accelerator=hk('bottom_of_document'))
        syntax_theme_menu = tk.Menu(view_menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color, activebackground='#3a3a3a')
        view_menu.add_cascade(label=t('menu.view.syntax_theme', 'Syntax Theme'), menu=syntax_theme_menu)
        syntax_theme_menu.add_command(label=t('menu.view.create_theme', 'Create Theme'), command=self.show_create_theme_dialog, accelerator=hk('create_theme'))
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
        view_menu.add_command(label=t('menu.view.compare_tabs', 'Compare Tabs'), command=self.show_split_compare, accelerator=hk('compare_tabs'))
        view_menu.add_command(label=t('menu.view.close_compare_tabs', 'Close Compare Tabs'), command=self.close_compare_panel, accelerator=hk('exit'))

        settings_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                                activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.settings', 'Settings'), menu=settings_menu)
        settings_menu.add_checkbutton(label=t('menu.view.edit_with_notepadx', 'Edit with Notepad-X'), variable=self.edit_with_shell_enabled, command=self.toggle_edit_with_shell, accelerator=hk('edit_with_notepadx'))
        settings_menu.add_checkbutton(label=t('menu.view.sound', 'Sound'), variable=self.sound_enabled, command=self.toggle_sound, accelerator=hk('sound'))
        settings_menu.add_separator()
        settings_menu.add_checkbutton(label=t('menu.view.status_bar', 'Status Bar'), variable=self.status_bar_enabled, command=self.toggle_status_bar, accelerator=hk('status_bar'))
        settings_menu.add_checkbutton(label=t('menu.view.numbered_lines', 'Numbered Lines'), variable=self.numbered_lines_enabled, command=self.toggle_numbered_lines, accelerator=hk('numbered_lines'))
        settings_menu.add_checkbutton(label=t('menu.view.autocomplete', 'Autocomplete'), variable=self.autocomplete_enabled, command=self.toggle_autocomplete, accelerator=hk('autocomplete'))
        settings_menu.add_checkbutton(label=t('menu.view.spell_check', 'Spell Check'), variable=self.spell_check_enabled, command=self.toggle_spell_check, accelerator=hk('spell_check'))
        settings_menu.add_checkbutton(label='Auto Pair Brackets/Quotes', variable=self.auto_pair_enabled, command=self.save_session, accelerator=hk('auto_pair'))
        settings_menu.add_checkbutton(label='Compare Multi-Edit', variable=self.compare_multi_edit_enabled, command=self.save_session, accelerator=hk('compare_multi_edit'))
        settings_menu.add_checkbutton(label='Minimap', variable=self.minimap_enabled, command=self.toggle_minimap, accelerator=hk('minimap'))
        settings_menu.add_checkbutton(label='Breadcrumbs', variable=self.breadcrumbs_enabled, command=self.toggle_breadcrumbs, accelerator=hk('breadcrumbs'))
        settings_menu.add_checkbutton(label='Diagnostics', variable=self.diagnostics_enabled, command=self.toggle_diagnostics, accelerator=hk('diagnostics'))
        settings_menu.add_checkbutton(label='Auto Save', variable=self.autosave_enabled, command=self.save_session, accelerator=hk('autosave'))
        settings_menu.add_checkbutton(label=t('menu.view.word_wrap', 'Word Wrap'), variable=self.word_wrap_enabled, command=self.toggle_word_wrap, accelerator=hk('word_wrap'))
        settings_menu.add_checkbutton(label=t('menu.view.preview_markdown', 'Preview Markdown'), variable=self.markdown_preview_enabled, command=self.toggle_markdown_preview, accelerator=hk('preview_markdown'))
        settings_menu.add_checkbutton(label=t('menu.edit.sync_page_navigation', 'Sync PgUp/PgDn in Compare'), variable=self.sync_page_navigation_enabled, command=self.save_session, accelerator=hk('sync_page_navigation'))
        settings_menu.add_separator()
        settings_menu.add_command(label=t('menu.settings.hotkeys', 'Hotkey Settings'), command=self.show_hotkey_config_dialog, accelerator=hk('hotkey_settings'))

        help_menu = tk.Menu(self.menu, tearoff=0, bg='#2d2d2d', fg=self.fg_color,
                            activebackground='#3a3a3a')
        self.menu.add_cascade(label=t('menu.help', 'Help'), menu=help_menu)
        help_menu.add_command(label=t('menu.help.contents', 'Help Contents'), command=self.show_help_contents, accelerator=hk('help_contents'))
        help_menu.add_command(label=t('menu.help.about', 'About Notepad-X'), command=self.show_about_dialog, accelerator=hk('about'))

    def show_help_contents(self):
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('app.help_title', 'Notepad-X Help'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color)
        dialog.geometry("900x650")

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
        help_text.bind('<Control-Button-1>', self.activate_help_lolcat)

        close_button = tk.Button(
            dialog,
            text=self.tr('common.close', 'Close'),
            command=lambda current=dialog, widget=help_text: self.close_help_contents_dialog(current, widget),
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

        dialog.bind('<Escape>', lambda e, current=dialog, widget=help_text: self.close_help_contents_dialog(current, widget))
        dialog.protocol('WM_DELETE_WINDOW', lambda current=dialog, widget=help_text: self.close_help_contents_dialog(current, widget))
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
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('app.compare_title', 'Compare With Tab'))
        dialog.transient(self.root)
        dialog.configure(bg=self.bg_color, padx=12, pady=12)
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

    def schedule_markdown_preview_refresh(self, immediate=False):
        if not self.markdown_preview_enabled.get():
            return
        if self.markdown_preview_refresh_job:
            try:
                self.root.after_cancel(self.markdown_preview_refresh_job)
            except tk.TclError:
                pass
        if immediate:
            self.markdown_preview_refresh_job = self.root.after_idle(self.refresh_markdown_preview)
        else:
            self.markdown_preview_refresh_job = self.root.after(self.markdown_preview_delay_ms, self.refresh_markdown_preview)

    def configure_markdown_preview_tags(self, text_widget):
        family = self.font_family
        size = int(self.current_font_size)
        code_family = 'Consolas' if self.is_windows else 'DejaVu Sans Mono'

        text_widget.tag_config('md_heading_1', font=(family, size + 9, 'bold'), spacing1=12, spacing3=6)
        text_widget.tag_config('md_heading_2', font=(family, size + 6, 'bold'), spacing1=10, spacing3=5)
        text_widget.tag_config('md_heading_3', font=(family, size + 4, 'bold'), spacing1=8, spacing3=4)
        text_widget.tag_config('md_heading_4', font=(family, size + 2, 'bold'), spacing1=6, spacing3=3)
        text_widget.tag_config('md_heading_5', font=(family, size + 1, 'bold'), spacing1=5, spacing3=3)
        text_widget.tag_config('md_heading_6', font=(family, size, 'bold'), spacing1=4, spacing3=2)
        text_widget.tag_config('md_bold', font=(family, size, 'bold'))
        text_widget.tag_config('md_italic', font=(family, size, 'italic'))
        text_widget.tag_config('md_code_inline', font=(code_family, size), background='#1f2630', foreground='#ffb86c')
        text_widget.tag_config('md_code_block', font=(code_family, size), background='#161b22', foreground='#d2e4ff', lmargin1=16, lmargin2=16, spacing1=5, spacing3=5)
        text_widget.tag_config('md_quote', foreground='#8b949e', lmargin1=18, lmargin2=18, spacing1=3, spacing3=3)
        text_widget.tag_config('md_list', lmargin1=14, lmargin2=24)
        text_widget.tag_config('md_rule', foreground='#7d8590')

    def insert_markdown_segment(self, text_widget, content, tags=()):
        if content == "":
            return
        start_index = text_widget.index('end-1c')
        text_widget.insert('end-1c', content)
        if tags:
            end_index = text_widget.index('end-1c')
            for tag in tags:
                text_widget.tag_add(tag, start_index, end_index)

    def render_markdown_inline(self, text_widget, line_text, base_tags=()):
        source = self.markdown_link_pattern.sub(r'\1 (\2)', line_text)
        position = 0
        while position < len(source):
            best_match = None
            best_tag = None
            for tag_name, pattern in self.markdown_inline_patterns:
                match = pattern.search(source, position)
                if match is None:
                    continue
                if best_match is None or match.start() < best_match.start():
                    best_match = match
                    best_tag = tag_name
            if best_match is None:
                self.insert_markdown_segment(text_widget, source[position:], base_tags)
                break
            if best_match.start() > position:
                self.insert_markdown_segment(text_widget, source[position:best_match.start()], base_tags)
            groups = [group for group in best_match.groups() if group]
            token_text = groups[-1] if groups else best_match.group(0)
            self.insert_markdown_segment(text_widget, token_text, tuple(base_tags) + (best_tag,))
            position = best_match.end()

    def render_markdown_preview(self, text_widget, markdown_text):
        lines = str(markdown_text or '').splitlines()
        if not lines:
            self.insert_markdown_segment(text_widget, self.tr('markdown.preview.empty', 'Nothing to preview.'), ('md_quote',))
            return

        in_code_block = False
        for raw_line in lines:
            line = raw_line.rstrip('\r')
            stripped = line.strip()
            if stripped.startswith('```'):
                in_code_block = not in_code_block
                continue
            if in_code_block:
                self.insert_markdown_segment(text_widget, line + '\n', ('md_code_block',))
                continue
            if not stripped:
                text_widget.insert(tk.END, '\n')
                continue
            compact_rule = stripped.replace(' ', '')
            if len(compact_rule) >= 3 and set(compact_rule) <= {'-', '*', '_'}:
                self.insert_markdown_segment(text_widget, '-' * 32 + '\n', ('md_rule',))
                continue
            heading_match = self.markdown_heading_pattern.match(line)
            if heading_match:
                level = len(heading_match.group(1))
                self.render_markdown_inline(text_widget, heading_match.group(2).strip(), (f'md_heading_{level}',))
                text_widget.insert(tk.END, '\n')
                continue
            quote_match = self.markdown_quote_pattern.match(line)
            if quote_match:
                self.render_markdown_inline(text_widget, quote_match.group(1), ('md_quote',))
                text_widget.insert(tk.END, '\n')
                continue
            list_match = self.markdown_list_pattern.match(line)
            if list_match:
                indent = ' ' * (len(list_match.group(1)) // 2)
                self.insert_markdown_segment(text_widget, f"{indent}- ", ('md_list',))
                self.render_markdown_inline(text_widget, list_match.group(3), ('md_list',))
                text_widget.insert(tk.END, '\n')
                continue
            ordered_match = self.markdown_ordered_list_pattern.match(line)
            if ordered_match:
                indent = ' ' * (len(ordered_match.group(1)) // 2)
                self.insert_markdown_segment(text_widget, f"{indent}{ordered_match.group(2)}. ", ('md_list',))
                self.render_markdown_inline(text_widget, ordered_match.group(3), ('md_list',))
                text_widget.insert(tk.END, '\n')
                continue
            self.render_markdown_inline(text_widget, line)
            text_widget.insert(tk.END, '\n')

    def refresh_markdown_preview(self):
        self.markdown_preview_refresh_job = None
        if not self.markdown_preview_enabled.get():
            return
        doc = self.get_current_doc()
        if not doc:
            self.close_markdown_preview(persist=False, restore_focus=False)
            return

        self.markdown_preview_source_tab = str(doc['frame'])
        compare_container_visible = False
        try:
            compare_container_visible = bool(self.compare_container.winfo_ismapped())
        except tk.TclError:
            compare_container_visible = False
        if not compare_container_visible:
            self.rebuild_editor_panes()
        self.compare_title.config(text=self.tr('markdown.preview.header', 'Markdown Preview: {title}', title=self.get_doc_title(doc['frame'])))

        compare_doc = self.compare_view
        compare_doc['file_path'] = doc.get('file_path')
        compare_doc['display_name'] = f"Preview: {self.get_doc_name(doc['frame'])}"
        compare_doc['syntax_override'] = None
        compare_doc['large_file_mode'] = False
        compare_doc['preview_mode'] = True
        compare_doc['virtual_mode'] = False
        compare_doc['is_remote'] = False
        compare_doc['remote_spec'] = None
        compare_doc['remote_host'] = None
        compare_doc['remote_path'] = None
        compare_doc['remote_shadow_path'] = None
        compare_doc['window_start_line'] = 1
        compare_doc['window_end_line'] = 1
        compare_doc['total_file_lines'] = 1
        self.invalidate_minimap_cache(compare_doc)
        compare_doc['diagnostics'] = []

        compare_text = compare_doc['text']
        try:
            previous_yview = compare_text.yview()[0]
        except (tk.TclError, IndexError):
            previous_yview = 0.0

        try:
            markdown_text = doc['text'].get('1.0', 'end-1c')
        except tk.TclError:
            markdown_text = ''

        compare_doc['suspend_modified_events'] = True
        try:
            compare_text.configure(state='normal')
            compare_text.delete('1.0', tk.END)
            self.configure_markdown_preview_tags(compare_text)
            self.render_markdown_preview(compare_text, markdown_text)
            compare_text.edit_modified(False)
            compare_text.mark_set(tk.INSERT, '1.0')
            compare_text.configure(state='disabled')
        finally:
            compare_doc['suspend_modified_events'] = False

        try:
            compare_text.yview_moveto(float(previous_yview))
        except (tk.TclError, TypeError, ValueError):
            pass
        compare_doc['total_file_lines'] = max(1, int(compare_text.index('end-1c').split('.')[0]))
        compare_doc['window_end_line'] = compare_doc['total_file_lines']
        self.update_line_number_gutter(compare_doc)
        self.schedule_minimap_refresh(compare_doc)
        self.schedule_diagnostics(compare_doc)
        if not compare_container_visible:
            self.schedule_compare_layout_refresh()
        self.update_status()

    def close_markdown_preview(self, event=None, persist=True, restore_focus=True):
        if self.markdown_preview_refresh_job:
            try:
                self.root.after_cancel(self.markdown_preview_refresh_job)
            except tk.TclError:
                pass
            self.markdown_preview_refresh_job = None
        self.markdown_preview_source_tab = None
        self.markdown_preview_enabled.set(False)
        if self.compare_view:
            try:
                self.compare_view['text'].configure(state='normal')
                self.compare_view['text'].delete('1.0', tk.END)
            except tk.TclError:
                pass
            self.compare_view['preview_mode'] = False
            self.compare_view['virtual_mode'] = False
            self.compare_view['large_file_mode'] = False
            self.compare_view['file_path'] = None
            self.compare_view['display_name'] = None
            self.compare_view['syntax_override'] = None
            self.compare_view['syntax_mode'] = None
            self.compare_view['is_remote'] = False
            self.compare_view['remote_spec'] = None
            self.compare_view['remote_host'] = None
            self.compare_view['remote_path'] = None
            self.compare_view['remote_shadow_path'] = None
            self.compare_view['window_start_line'] = 1
            self.compare_view['window_end_line'] = 1
            self.compare_view['total_file_lines'] = 1
        self.compare_title.config(text="")
        try:
            self.editor_paned.forget(self.compare_container)
        except tk.TclError:
            pass
        self.rebuild_editor_panes()
        self.root.after_idle(self.set_currently_editing_sash_position)
        if hasattr(self, 'compare_status'):
            self.compare_status.place_forget()
        if persist:
            self.save_session()
        self.update_status()
        if restore_focus and self.text and self.text.winfo_exists():
            self.text.focus_set()
        return "break"

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
        compare_doc['display_name'] = doc.get('display_name')
        compare_doc['syntax_override'] = doc.get('syntax_override')
        compare_doc['large_file_mode'] = bool(doc.get('large_file_mode'))
        compare_doc['preview_mode'] = bool(doc.get('preview_mode'))
        compare_doc['virtual_mode'] = bool(doc.get('virtual_mode'))
        compare_doc['is_remote'] = bool(doc.get('is_remote'))
        compare_doc['remote_spec'] = doc.get('remote_spec')
        compare_doc['remote_host'] = doc.get('remote_host')
        compare_doc['remote_path'] = doc.get('remote_path')
        compare_doc['remote_shadow_path'] = doc.get('remote_shadow_path')
        compare_doc['window_start_line'] = doc.get('window_start_line', 1)
        compare_doc['window_end_line'] = doc.get('window_end_line', 1)
        compare_doc['total_file_lines'] = doc.get('total_file_lines', 1)
        self.invalidate_minimap_cache(compare_doc)
        compare_doc['diagnostics'] = []

        self.refresh_compare_header()

        compare_text = compare_doc['text']
        try:
            compare_doc['last_insert_index'] = compare_text.index(tk.INSERT)
            compare_doc['last_yview'] = compare_text.yview()[0]
            compare_doc['last_xview'] = compare_text.xview()[0]
        except (tk.TclError, IndexError):
            pass

        compare_doc['suspend_modified_events'] = True
        compare_text.configure(state='normal')
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
        compare_doc['symbol_cache_signature'] = None
        compare_doc['symbol_cache'] = None
        self.invalidate_fold_regions(compare_doc)
        self.schedule_diagnostics(compare_doc)
        self.sync_compare_note_tags(doc)
        self.update_line_number_gutter(compare_doc)
        self.schedule_minimap_refresh(compare_doc)
        self.update_status()

    def set_compare_sash_position(self):
        if not self.is_side_panel_visible():
            return
        try:
            self.editor_paned.update_idletasks()
            width = self.editor_paned.winfo_width()
            sash_width = int(self.editor_paned.cget('sashwidth') or 0)
            available_width = max(0, width - sash_width)
            if available_width > 0:
                sash_x = available_width // 2
                if available_width >= 480:
                    sash_x = max(240, min(sash_x, available_width - 240))
                self.editor_paned.sash_place(0, sash_x, 0)
                self.position_compare_status()
        except tk.TclError:
            pass

    def get_currently_editing_sidebar_width(self):
        font = getattr(self, 'currently_editing_measure_font', None)
        if font is None:
            self.currently_editing_measure_font = tkfont.Font(family='Consolas', size=10)
            font = self.currently_editing_measure_font
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
            if self.is_side_panel_visible():
                self.editor_paned.add(self.compare_container, stretch='always')
        except tk.TclError:
            pass

    def set_currently_editing_sash_position(self):
        return

    def schedule_compare_layout_refresh(self):
        if not self.is_side_panel_visible():
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
        if self.markdown_preview_enabled.get():
            self.close_markdown_preview(persist=False, restore_focus=False)
        self.compare_source_tab = str(source_doc['frame'])
        self.compare_active = True
        self.rebuild_editor_panes()
        self.refresh_compare_panel()
        self.root.after_idle(self.set_compare_sash_position)
        self.root.after_idle(self.set_currently_editing_sash_position)
        self.update_status()
        self.save_session()
        return "break"

    def close_compare_panel(self, event=None, persist=True, restore_focus=True):
        if self.compare_refresh_job:
            try:
                self.root.after_cancel(self.compare_refresh_job)
            except tk.TclError:
                pass
            self.compare_refresh_job = None

        if self.compare_view:
            self.cancel_text_theme_effect_job(self.compare_view)
            if self.compare_view.get('colorizer') is not None and self.compare_view.get('percolator') is not None:
                try:
                    self.compare_view['percolator'].removefilter(self.compare_view['colorizer'])
                except Exception:
                    pass
            self.compare_view['percolator'] = None
            self.compare_view['colorizer'] = None
            self.compare_view['theme_effect_job'] = None
            self.compare_view['syntax_job'] = None
            self.compare_view['syntax_mode'] = None
            self.compare_view['file_path'] = None
            self.compare_view['display_name'] = None
            self.compare_view['syntax_override'] = None
            self.compare_view['large_file_mode'] = False
            self.compare_view['preview_mode'] = False
            self.compare_view['virtual_mode'] = False
            self.compare_view['is_remote'] = False
            self.compare_view['remote_spec'] = None
            self.compare_view['remote_host'] = None
            self.compare_view['remote_path'] = None
            self.compare_view['remote_shadow_path'] = None
            self.compare_view['text'].configure(state='normal')
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
        if persist:
            self.save_session()
        if restore_focus and self.text and self.text.winfo_exists():
            self.text.focus_set()
        return "break"

    def show_about_dialog(self):
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('app.about_title', 'About Notepad-X'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=24, pady=20)
        dialog.pong_after_id = None

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
                text=self.tr('about.icon_placeholder', '[Icon]'),
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
            text=self.tr('about.pong.title', 'Pong-X'),
            bg=self.bg_color,
            fg=self.fg_color,
            font=('Segoe UI', 14, 'bold')
        )
        header.pack(pady=(0, 8))

        score_label = tk.Label(
            dialog,
            text=self.tr(
                'about.pong.score',
                '{left_label} {left_score}   {right_label} {right_score}',
                left_label=self.tr('about.pong.user', 'User'),
                left_score=0,
                right_label=self.tr('about.pong.computer', 'Computer'),
                right_score=0
            ),
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
            text=self.tr(
                'about.pong.info',
                'Player 1 keys: W & S Player 2 keys: Up & Down, Press Up/Down once to start PVP. Press R to restart.'
            ),
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
            text=self.tr(
                'about.pong.score',
                '{left_label} {left_score}   {right_label} {right_score}',
                left_label=self.tr('about.pong.user', 'User') if state['mode'] == 'ai' else self.tr('about.pong.player1', 'Player 1'),
                left_score=state['left_score'],
                right_label=self.tr('about.pong.computer', 'Computer') if state['mode'] == 'ai' else self.tr('about.pong.player2', 'Player 2'),
                right_score=state['right_score']
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
    def normalize_grab_git_repo_identifier(self, value):
        repo_text = str(value or '').strip()
        if not repo_text:
            return None
        if any(char in repo_text for char in '\r\n\0'):
            return None
        repo_text = repo_text.replace('\\', '/').rstrip('/')
        github_match = re.match(r'^(?:https?://)?(?:www\.)?github\.com/([^/]+)/([^/]+?)(?:\.git)?$', repo_text, re.IGNORECASE)
        if github_match:
            owner = github_match.group(1).strip()
            repo_name = github_match.group(2).strip()
        else:
            parts = [segment.strip() for segment in repo_text.split('/') if segment.strip()]
            if len(parts) != 2:
                return None
            owner, repo_name = parts
        if repo_name.lower().endswith('.git'):
            repo_name = repo_name[:-4]
        if not owner or not repo_name:
            return None
        if not re.fullmatch(r'[A-Za-z0-9](?:[A-Za-z0-9._-]{0,98}[A-Za-z0-9])?', owner):
            return None
        if not re.fullmatch(r'[A-Za-z0-9._-]{1,100}', repo_name):
            return None
        return f"{owner}/{repo_name}"

    def build_grab_git_clone_error_message(self, repo_identifier, error_detail):
        detail_text = str(error_detail or '').strip()
        lowered = detail_text.lower()
        if 'repository not found' in lowered or 'not found' in lowered:
            return self.tr(
                'grab_git.not_found',
                'GitHub could not find "{repo_identifier}".\n\nCheck the username/project and try again.'
            ).format(repo_identifier=repo_identifier)
        if 'could not read username' in lowered or 'authentication failed' in lowered:
            return self.tr(
                'grab_git.private_repo',
                'That GitHub project could not be downloaded.\n\nIt may be private or require authentication.'
            )
        return self.tr(
            'grab_git.clone_failed',
            'Notepad-X could not download that GitHub project.\n\n{error_detail}'
        ).format(error_detail=detail_text or self.tr('grab_git.unknown_failure', 'Unknown git clone failure.'))

    def prompt_grab_git_repository(self, parent=None):
        parent = parent or self.root
        dialog = self.create_toplevel(parent)
        dialog.title(self.tr('grab_git.title', 'Grab Git'))
        dialog.transient(parent)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=14, pady=12)

        result = {'value': None}
        repo_var = tk.StringVar()

        tk.Label(
            dialog,
            text=self.tr(
                'grab_git.prompt',
                'Enter the GitHub project as:\nusername/project'
            ),
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        ).pack(anchor='w', pady=(0, 8))

        repo_entry = tk.Entry(dialog, textvariable=repo_var, width=36)
        repo_entry.pack(fill='x', pady=(0, 8))

        button_row = tk.Frame(dialog, bg=self.bg_color)
        button_row.pack(fill='x')

        def submit(event=None):
            normalized = self.normalize_grab_git_repo_identifier(repo_var.get())
            if not normalized:
                messagebox.showwarning(
                    self.tr('grab_git.title', 'Grab Git'),
                    self.tr(
                        'grab_git.invalid',
                        'Enter the project as username/project.\n\nExample:\nopenai/openai-python'
                    ),
                    parent=dialog
                )
                return "break"
            result['value'] = normalized
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = None
            dialog.destroy()
            return "break"

        tk.Button(button_row, text=self.tr('common.ok', 'OK'), width=10, command=submit).pack(side='left')
        tk.Button(button_row, text=self.tr('common.close', 'Close'), width=10, command=cancel).pack(side='right')

        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol('WM_DELETE_WINDOW', cancel)
        dialog.update_idletasks()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        repo_entry.focus_force()
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(1, lambda current=dialog: self.center_window_after_show(current, parent))
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.hide_autocomplete_popup()

    def get_grab_git_project_files(self, root_dir):
        project_extensions = {
            '.txt', '.md', '.rst', '.log', '.ini', '.cfg', '.conf', '.toml', '.yaml', '.yml',
            '.json', '.xml', '.csv', '.tsv', '.html', '.htm', '.css', '.js', '.ts', '.tsx',
            '.jsx', '.py', '.pyw', '.rs', '.c', '.h', '.cpp', '.cc', '.cxx', '.hpp', '.hh',
            '.hxx', '.java', '.php', '.sql', '.sh', '.bat', '.cmd', '.ps1', '.psm1', '.psd1',
            '.ps1xml', '.pssc', '.psrc', '.psc1', '.cdxml', '.asm', '.s',
            '.tex', '.vb', '.vbs', '.pas', '.pl', '.pm', '.diff', '.patch', '.nsi', '.nsh',
            '.iss', '.rc', '.as', '.mx', '.asp', '.aspx', '.au3', '.ml', '.mli', '.sml',
            '.thy', '.for', '.f', '.f90', '.f95', '.f2k', '.lsp', '.lisp', '.mak', '.m',
            '.nfo', '.st', '.xsd', '.xsml', '.xsl', '.kml', '.gitignore', '.gitattributes',
            '.editorconfig', '.env', '.sample'
        }
        preferred_names = {
            'readme', 'readme.md', 'license', 'license.txt', 'copying', 'dockerfile',
            'makefile', 'cmakelists.txt'
        }
        root_dir = os.path.abspath(root_dir)
        candidate_files = []

        for current_root, dir_names, file_names in os.walk(root_dir):
            dir_names[:] = [name for name in dir_names if name.lower() not in {'.git', '.vs', '__pycache__'}]
            for file_name in sorted(file_names, key=str.lower):
                lower_name = file_name.lower()
                if self.is_notepadx_support_file(lower_name):
                    continue
                full_path = os.path.join(current_root, file_name)
                extension = os.path.splitext(lower_name)[1]
                if lower_name in preferred_names or extension in project_extensions:
                    candidate_files.append(full_path)

        candidate_files.sort(key=lambda path: os.path.relpath(path, root_dir).lower())
        return candidate_files

    def prompt_grab_git_files_to_open(self, root_dir, file_paths, parent=None):
        parent = parent or self.root
        dialog = self.create_toplevel(parent)
        dialog.title(self.tr('grab_git.open_title', 'Open Downloaded Project Files'))
        dialog.transient(parent)
        dialog.configure(bg=self.bg_color, padx=12, pady=12)
        dialog.minsize(760, 420)
        dialog.grid_rowconfigure(1, weight=1)
        dialog.grid_columnconfigure(0, weight=1)

        tk.Label(
            dialog,
            text=self.tr(
                'grab_git.open_prompt',
                'Downloaded project root:\n{root_dir}\n\nSelect one or more files to open.'
            ).format(root_dir=root_dir),
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        ).grid(row=0, column=0, sticky='ew', pady=(0, 10))

        list_frame = tk.Frame(dialog, bg=self.bg_color)
        list_frame.grid(row=1, column=0, sticky='nsew')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        listbox = tk.Listbox(
            list_frame,
            selectmode='extended',
            activestyle='dotbox',
            exportselection=False,
            bg=self.text_bg,
            fg=self.text_fg,
            selectbackground=self.select_bg,
            selectforeground='white',
            font=('Consolas', 10)
        )
        listbox.grid(row=0, column=0, sticky='nsew')

        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        listbox.configure(yscrollcommand=scrollbar.set)

        relative_paths = [os.path.relpath(path, root_dir) for path in file_paths]
        for relative_path in relative_paths:
            listbox.insert(tk.END, relative_path)
        if relative_paths:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.see(0)

        result = {'value': None}

        def submit(event=None):
            selection = listbox.curselection()
            if not selection:
                messagebox.showinfo(
                    self.tr('grab_git.open_title', 'Open Downloaded Project Files'),
                    self.tr('grab_git.select_files', 'Select one or more files to open.'),
                    parent=dialog
                )
                return "break"
            result['value'] = [file_paths[index] for index in selection]
            dialog.destroy()
            return "break"

        def cancel(event=None):
            result['value'] = []
            dialog.destroy()
            return "break"

        button_row = tk.Frame(dialog, bg=self.bg_color)
        button_row.grid(row=2, column=0, sticky='ew', pady=(10, 0))
        tk.Button(button_row, text=self.tr('grab_git.open_selected', 'Open Selected'), width=14, command=submit).pack(side='left')
        tk.Button(button_row, text=self.tr('common.close', 'Close'), width=10, command=cancel).pack(side='right')

        listbox.bind('<Double-Button-1>', submit)
        dialog.bind('<Return>', submit)
        dialog.bind('<Escape>', cancel)
        dialog.protocol('WM_DELETE_WINDOW', cancel)
        dialog.update_idletasks()
        self.center_window(dialog, parent)
        dialog.lift()
        dialog.attributes('-topmost', True)
        dialog.grab_set()
        listbox.focus_set()
        try:
            dialog.wait_visibility()
            self.center_window(dialog, parent)
            dialog.after(1, lambda current=dialog: self.center_window_after_show(current, parent))
            dialog.after(50, lambda: dialog.attributes('-topmost', False) if dialog.winfo_exists() else None)
            parent.wait_window(dialog)
            return result['value']
        finally:
            self.hide_autocomplete_popup()

    def grab_git_project(self, event=None):
        git_executable = shutil.which('git')
        if not git_executable:
            messagebox.showerror(
                self.tr('grab_git.title', 'Grab Git'),
                self.tr('grab_git.git_missing', 'Git could not be found on this system.'),
                parent=self.root
            )
            return "break"

        repo_identifier = self.prompt_grab_git_repository(parent=self.root)
        if not repo_identifier:
            return "break"

        destination_dir = filedialog.askdirectory(
            parent=self.root,
            title=self.tr('grab_git.choose_folder', 'Choose where to save the GitHub project')
        )
        if not destination_dir:
            return "break"

        repo_name = repo_identifier.split('/', 1)[1]
        clone_target = os.path.abspath(os.path.join(destination_dir, repo_name))
        if os.path.exists(clone_target):
            messagebox.showwarning(
                self.tr('grab_git.title', 'Grab Git'),
                self.tr('grab_git.exists', 'That destination folder already exists.\nChoose another location or remove the existing folder first.'),
                parent=self.root
            )
            return "break"

        progress_dialog = self.create_toplevel(self.root)
        progress_dialog.title(self.tr('grab_git.downloading', 'Downloading Project'))
        progress_dialog.transient(self.root)
        progress_dialog.resizable(False, False)
        progress_dialog.configure(bg=self.bg_color, padx=14, pady=12)

        tk.Label(
            progress_dialog,
            text=self.tr(
                'grab_git.downloading_prompt',
                'Downloading:\n{repo_identifier}\n\nPlease wait...'
            ).format(repo_identifier=repo_identifier),
            bg=self.bg_color,
            fg=self.fg_color,
            justify='left',
            anchor='w'
        ).pack(anchor='w', pady=(0, 10))

        progress_bar = ttk.Progressbar(progress_dialog, orient='horizontal', mode='indeterminate', length=320)
        progress_bar.pack(fill='x')
        progress_bar.start(10)
        progress_dialog.protocol('WM_DELETE_WINDOW', lambda: None)
        progress_dialog.update_idletasks()
        self.center_window(progress_dialog, self.root)
        progress_dialog.lift()
        progress_dialog.attributes('-topmost', True)
        progress_dialog.grab_set()
        progress_dialog.focus_force()
        progress_dialog.after(1, lambda current=progress_dialog: self.center_window_after_show(current, self.root))
        progress_dialog.after(50, lambda: progress_dialog.attributes('-topmost', False) if progress_dialog.winfo_exists() else None)

        result = {'done': False, 'returncode': 1, 'stdout': '', 'stderr': '', 'error': None}

        def worker():
            clone_url = f"https://github.com/{repo_identifier}.git"
            try:
                completed = subprocess.run(
                    [git_executable, 'clone', clone_url, clone_target],
                    capture_output=True,
                    text=True,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
                result['returncode'] = completed.returncode
                result['stdout'] = completed.stdout or ''
                result['stderr'] = completed.stderr or ''
            except Exception as exc:
                result['error'] = exc
            finally:
                result['done'] = True

        def finish_clone():
            if not progress_dialog.winfo_exists():
                return
            if not result['done']:
                progress_dialog.after(120, finish_clone)
                return

            progress_bar.stop()
            try:
                progress_dialog.grab_release()
            except tk.TclError:
                pass
            progress_dialog.destroy()

            if result.get('error') is not None or result.get('returncode') != 0:
                error_detail = str(
                    result.get('error')
                    or result.get('stderr')
                    or result.get('stdout')
                    or self.tr('grab_git.unknown_failure', 'Unknown git clone failure.')
                )
                messagebox.showerror(
                    self.tr('grab_git.title', 'Grab Git'),
                    self.build_grab_git_clone_error_message(repo_identifier, error_detail),
                    parent=self.root
                )
                return

            candidate_files = self.get_grab_git_project_files(clone_target)
            if not candidate_files:
                messagebox.showinfo(
                    self.tr('grab_git.title', 'Grab Git'),
                    self.tr('grab_git.no_files', 'The project downloaded successfully, but no matching files were found to open.'),
                    parent=self.root
                )
                return

            selected_files = self.prompt_grab_git_files_to_open(clone_target, candidate_files, parent=self.root)
            if not selected_files:
                return
            for selected_file in selected_files:
                self.open_file_path(selected_file)

        threading.Thread(target=worker, name='NotepadXGrabGitClone', daemon=True).start()
        finish_clone()
        return "break"

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
            messagebox.showwarning(
                self.tr('file.missing_title', 'File Missing'),
                self.tr('file.missing_message', 'That file could not be found.'),
                parent=self.root
            )
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
                self.cleanup_failed_file_open(current_doc)
                return False
            current_doc['background_open_new_tab'] = False
            self.notebook.tab(current_doc['frame'], text=self.get_doc_title(current_doc['frame']))
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
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr('menu.file.open', 'Open'),
            filetypes=self.get_open_filetypes()
        )
        if file_path:
            self.open_file_path(file_path)
        return "break"

    def open_project_path(self, file_path):
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning(
                self.tr('file.missing_title', 'File Missing'),
                self.tr('file.missing_message', 'That file could not be found.'),
                parent=self.root
            )
            return False
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
        return True

    def open_project(self, event=None):
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr('menu.file.open_project', 'Open Project'),
            filetypes=self.get_save_filetypes()
        )
        if not file_path:
            return "break"
        self.open_project_path(file_path)
        return "break"

    def open_recent_file(self, file_path):
        self.open_file_path(file_path)

    def write_file_atomically(self, file_path, content):
        directory = os.path.dirname(file_path) or '.'
        target_mode = None
        temp_path = None
        if not self.is_windows:
            try:
                target_mode = stat.S_IMODE(os.stat(file_path).st_mode)
            except OSError:
                target_mode = 0o664
        try:
            fd, temp_path = tempfile.mkstemp(prefix='notepadx-save-', suffix='.tmp', dir=directory)
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
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError as exc:
                    self.log_exception("cleanup temp save file", exc)

    def get_save_filetypes(self, include_encrypted=False):
        t = self.tr
        filetypes = [
            (t('filetype.all_supported', 'All Supported'), "*.txt *.md *.log .gitignore *.py *.pyw *.c *.cpp *.cxx *.cc *.h *.hpp *.hxx *.hh *.cs *.rs *.java *.js *.html *.htm *.php *.xml *.sql *.css *.json *.ini *.bat *.cmd *.ps1 *.psm1 *.psd1 *.ps1xml *.pssc *.psrc *.psc1 *.cdxml *.sh *.asm *.s *.tex *.vb *.vbs *.pas *.pl *.pm *.diff *.patch *.nsi *.nsh *.iss *.rc *.as *.mx *.asp *.aspx *.au3 *.ml *.mli *.sml *.thy *.for *.f *.f90 *.f95 *.f2k *.lsp *.lisp *.mak *.m *.nfo *.st *.xsd *.xsml *.xsl *.kml"),
            (t('filetype.text_document', 'Text Document'), "*.txt"),
            (t('filetype.markdown', 'Markdown'), "*.md"),
            (t('filetype.log', 'Log File'), "*.log"),
            (t('filetype.git_ignore', 'Git Ignore'), ".gitignore"),
            (t('filetype.python', 'Python'), "*.py *.pyw"),
            (t('filetype.c_headers', 'C / Headers'), "*.c *.h"),
            (t('filetype.cpp_headers', 'C++ / Headers'), "*.cpp *.cxx *.cc *.hpp *.hxx *.hh"),
            (t('filetype.csharp', 'C#'), "*.cs"),
            (t('filetype.rust', 'Rust'), "*.rs"),
            (t('filetype.java', 'Java'), "*.java"),
            (t('filetype.javascript', 'JavaScript'), "*.js"),
            (t('filetype.html', 'HTML'), "*.html *.htm"),
            (t('filetype.php', 'PHP'), "*.php *.php3 *.phtml"),
            (t('filetype.xml', 'XML'), "*.xml *.xsd *.xsml *.xsl *.kml"),
            (t('filetype.sql', 'SQL'), "*.sql"),
            (t('filetype.css', 'CSS'), "*.css"),
            (t('filetype.json', 'JSON'), "*.json"),
            (t('filetype.ini_config', 'INI / Config'), "*.ini *.inf *.reg *.url"),
            (t('filetype.batch', 'Batch'), "*.bat *.cmd"),
            (t('filetype.powershell', 'PowerShell'), "*.ps1 *.psm1 *.psd1 *.ps1xml *.pssc *.psrc *.psc1 *.cdxml"),
            (t('filetype.shell', 'Shell'), "*.sh *.bsh"),
            (t('filetype.assembly', 'Assembly'), "*.asm *.s"),
            (t('filetype.pascal', 'Pascal'), "*.pas *.inc"),
            (t('filetype.perl', 'Perl'), "*.pl *.pm *.plx"),
            (t('filetype.diff_patch', 'Diff / Patch'), "*.diff *.patch"),
            (t('filetype.vb_vbscript', 'VB / VBScript'), "*.vb *.vbs"),
            (t('filetype.actionscript', 'ActionScript'), "*.as *.mx"),
            (t('filetype.asp_aspx', 'ASP / ASPX'), "*.asp *.aspx"),
            (t('filetype.autoit', 'AutoIt'), "*.au3"),
            (t('filetype.caml', 'Caml'), "*.ml *.mli *.sml *.thy"),
            (t('filetype.fortran', 'Fortran'), "*.f *.for *.f90 *.f95 *.f2k"),
            (t('filetype.inno_setup', 'Inno Setup'), "*.iss"),
            (t('filetype.lisp', 'Lisp'), "*.lsp *.lisp"),
            (t('filetype.makefile', 'Makefile'), "*.mak"),
            (t('filetype.matlab', 'Matlab'), "*.m"),
            (t('filetype.nfo', 'NFO'), "*.nfo"),
            (t('filetype.nsis', 'NSIS'), "*.nsi *.nsh"),
            (t('filetype.resource', 'Resource'), "*.rc"),
            (t('filetype.smalltalk', 'Smalltalk'), "*.st"),
            (t('filetype.tex', 'TeX'), "*.tex"),
            (t('filetype.all_files', 'All Files'), "*.*"),
        ]
        if include_encrypted:
            return [(t('filetype.encrypted', 'Notepad-X Encrypted'), "*.npxe"), *filetypes]
        return filetypes

    def get_open_filetypes(self):
        t = self.tr
        all_supported_label = t('filetype.all_supported', 'All Supported')
        all_files_label = t('filetype.all_files', 'All Files')
        encrypted_entry = (t('filetype.encrypted', 'Notepad-X Encrypted'), "*.npxe")
        all_supported_pattern = ""
        all_files_entry = (all_files_label, "*.*")
        other_filetypes = []

        for label, pattern in self.get_save_filetypes():
            if label == all_supported_label:
                all_supported_pattern = str(pattern)
            elif label == all_files_label:
                all_files_entry = (label, pattern)
            else:
                other_filetypes.append((label, pattern))

        supported_patterns = list(dict.fromkeys(
            token for token in f"{all_supported_pattern} {encrypted_entry[1]}".split() if token
        ))
        all_supported_entry = (all_supported_label, " ".join(supported_patterns))
        return [all_supported_entry, encrypted_entry, *other_filetypes, all_files_entry]

    def get_primary_save_extension(self, pattern):
        for token in str(pattern or '').split():
            token = token.strip()
            if not token or token == '*.*':
                continue
            if token.startswith('*.') and len(token) > 2:
                return token[1:]
            if token.startswith('.'):
                return token
        return None

    def get_selected_save_extension(self, selected_filetype, filetypes, fallback_extension='.txt'):
        selected_value = str(selected_filetype or '').strip().lower()
        fallback_value = fallback_extension if str(fallback_extension).startswith('.') else f".{str(fallback_extension).lstrip('.')}"

        for label, pattern in filetypes:
            extension = self.get_primary_save_extension(pattern)
            if not extension:
                continue
            label_value = str(label or '').strip().lower()
            pattern_value = str(pattern or '').strip().lower()
            if selected_value and (selected_value == label_value or selected_value == pattern_value or selected_value in pattern_value.split()):
                return extension
        return fallback_value

    def normalize_save_file_path(self, file_path, selected_filetype='', filetypes=None, fallback_extension='.txt'):
        if not file_path:
            return file_path

        base_name = os.path.basename(file_path)
        if not base_name:
            return file_path
        if base_name.startswith('.') and base_name.count('.') == 1:
            return file_path

        _, extension = os.path.splitext(base_name)
        if extension:
            return file_path

        effective_filetypes = filetypes if filetypes is not None else self.get_save_filetypes()
        selected_extension = self.get_selected_save_extension(
            selected_filetype,
            effective_filetypes,
            fallback_extension=fallback_extension
        )
        if not selected_extension.startswith('.'):
            selected_extension = f".{selected_extension.lstrip('.')}"
        return f"{file_path}{selected_extension}"

    def prompt_plain_text_save_path(self, title, initialfile=None):
        filetypes = self.get_save_filetypes()
        dialog_options = {
            'parent': self.root,
            'title': title,
            'defaultextension': '.txt',
            'filetypes': filetypes,
        }
        if initialfile:
            dialog_options['initialfile'] = initialfile

        selected_filetype = tk.StringVar(master=self.root, value='')
        try:
            file_path = filedialog.asksaveasfilename(typevariable=selected_filetype, **dialog_options)
        except tk.TclError:
            file_path = filedialog.asksaveasfilename(**dialog_options)
            selected_value = ''
        else:
            selected_value = selected_filetype.get()

        return self.normalize_save_file_path(
            file_path,
            selected_filetype=selected_value,
            filetypes=filetypes,
            fallback_extension='.txt'
        )

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
        display_name_keys = {
            'python': ('filetype.python', 'Python'),
            'javascript': ('filetype.javascript', 'JavaScript'),
            'php': ('filetype.php', 'PHP'),
            'batch': ('filetype.batch', 'Batch'),
            'powershell': ('filetype.powershell', 'PowerShell'),
            'shell': ('filetype.shell', 'Shell'),
            'html': ('filetype.html', 'HTML'),
        }
        locale_entry = display_name_keys.get(language)
        if locale_entry:
            return self.tr(locale_entry[0], locale_entry[1])
        return str(language or '').title() or self.tr('common.unknown', 'Unknown')

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
        if doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
            messagebox.showinfo(
                self.tr('run.title', 'Save and Run'),
                self.tr('run.large_file_unavailable', 'Save and Run is not available for buffered large-file tabs.'),
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
                self.tr('run.unsafe_path', 'That file path could not be sent to a runtime safely.'),
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
        return self.save_document_content(doc, autosave=False, show_errors=True, update_recent=True)

    def save_all(self, event=None):
        original_tab = self.notebook.select()
        any_saved = False

        for tab_id in list(self.notebook.tabs()):
            doc = self.documents.get(str(tab_id))
            if not doc or doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
                continue

            self.notebook.select(tab_id)
            self.set_active_document(tab_id)

            if doc['file_path'] or self.doc_has_unsaved_changes(doc):
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
        if doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
            messagebox.showinfo(
                self.tr('large_file.title', 'Large File Mode'),
                self.tr('large_file.save_as_disabled', 'Save As is disabled for buffered large-file tabs.'),
                parent=self.root
            )
            return False
        file_path = self.prompt_plain_text_save_path(self.tr('save.as_title', 'Save As'))
        if file_path:
            old_file_path = doc.get('file_path')
            if old_file_path and old_file_path != file_path:
                self.unregister_doc_from_shared_notes(doc)
            doc['file_path'] = os.path.abspath(file_path)
            self.clear_remote_metadata(doc)
            doc['display_name'] = None
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
        output_path = self.prompt_plain_text_save_path(
            self.tr('save.copy_title', 'Save Copy As'),
            initialfile=suggested_name
        )
        if not output_path:
            return "break"
        try:
            if doc.get('preview_mode') or (doc.get('virtual_mode') and not self.is_virtual_editable(doc)):
                with open(doc['file_path'], 'rb') as src, open(output_path, 'wb') as dst:
                    while True:
                        chunk = src.read(self.file_load_chunk_size)
                        if not chunk:
                            break
                        dst.write(chunk)
            elif self.is_virtual_editable(doc):
                if not self.flush_virtual_window_edits(doc, force=True):
                    return "break"
                self.write_virtual_document_atomically(doc, output_path)
            else:
                self.write_file_atomically(output_path, doc['text'].get('1.0', tk.END).rstrip('\n'))
            messagebox.showinfo(
                self.tr('save.copy_title', 'Save Copy As'),
                self.tr('save.copy_saved', 'Copy saved to:\n{output_path}', output_path=output_path),
                parent=self.root
            )
        except PermissionError as exc:
            self.show_filesystem_error(self.tr('save.copy_title', 'Save Copy As'), output_path, exc)
        except RuntimeError as exc:
            messagebox.showerror(self.tr('save.copy_title', 'Save Copy As'), str(exc), parent=self.root)
        except ValueError as exc:
            messagebox.showerror(self.tr('save.copy_title', 'Save Copy As'), str(exc), parent=self.root)
        except OSError as exc:
            self.log_exception("save copy as", exc)
            self.show_filesystem_error(self.tr('save.copy_title', 'Save Copy As'), output_path, exc)
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
            title=self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'),
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
                    self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'),
                    self.tr(
                        'large_file.encryption_disabled',
                        'Encryption is not available for buffered large-file or preview tabs.'
                    ),
                    parent=self.root
                )
                return "break"
            self.write_encrypted_text_file(
                output_path,
                doc['text'].get('1.0', tk.END).rstrip('\n'),
                passphrase=encryption_options.get('passphrase'),
                original_name=suggested_name
            )
            messagebox.showinfo(
                self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'),
                self.tr(
                    'save.encrypted_copy_saved',
                    'Encrypted copy saved to:\n{output_path}',
                    output_path=output_path
                ),
                parent=self.root
            )
        except PermissionError as exc:
            self.show_filesystem_error(self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'), output_path, exc)
        except RuntimeError as exc:
            messagebox.showerror(self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'), str(exc), parent=self.root)
        except ValueError as exc:
            messagebox.showerror(self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'), str(exc), parent=self.root)
        except OSError as exc:
            self.log_exception("save encrypted copy", exc)
            self.show_filesystem_error(self.tr('save.encrypted_copy_title', 'Save Encrypted Copy'), output_path, exc)
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
            messagebox.showerror(
                self.tr('print.failed_title', 'Print Failed'),
                self.tr('print.unsafe_path', 'That file path could not be sent to the print command safely.'),
                parent=self.root
            )
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
                        raise OSError(
                            self.tr(
                                'print.windows_failed_code',
                                'Windows print action failed with code {code}',
                                code=result
                            )
                        )
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
            print_error = OSError(self.tr('print.unavailable', 'Print is not available on this platform.'))
        self.log_exception("print file", print_error)
        messagebox.showerror(self.tr('print.failed_title', 'Print Failed'), str(print_error), parent=self.root)
        return "break"

    def exit_app(self, event=None):
        if not self.confirm_exit_app():
            return "break"
        self.finalize_exit_app()
        self.request_app_shutdown()
        return "break"

    def confirm_close_tab(self, doc):
        if not self.doc_has_unsaved_changes(doc):
            return True

        self.notebook.select(doc['frame'])
        self.set_active_document(doc['frame'])
        answer = messagebox.askyesnocancel(
            self.tr('save.title', 'Save'),
            self.tr(
                'save.close_prompt',
                'Save changes to {file_name} before closing?',
                file_name=self.get_doc_name(doc['frame'])
            ),
            parent=self.root
        )
        if answer is True:
            return self.save()
        return answer is False

    def current_doc_is_large_readonly(self):
        doc = self.get_current_doc()
        return bool(doc and self.is_doc_text_readonly(doc))

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
        if doc and self.is_doc_text_readonly(doc):
            return "break"

        try:
            selection = target.get('sel.first', 'sel.last')
        except tk.TclError:
            return "break"

        self.root.clipboard_clear()
        self.root.clipboard_append(selection)
        self.root.update_idletasks()

        mirror_target = self.get_compare_mirror_target(target)
        if mirror_target is not None:
            self.sync_mirror_target_position(target, mirror_target)
        try:
            target.delete('sel.first', 'sel.last')
            target.edit_modified(True)
        except tk.TclError:
            return "break"
        if mirror_target is not None:
            try:
                mirror_target.delete('sel.first', 'sel.last')
                mirror_target.edit_modified(True)
            except tk.TclError:
                pass

        self.set_last_active_editor_widget(target)
        if mirror_target is not None:
            self.notify_text_widget_changed(mirror_target)
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
        if doc and self.is_doc_text_readonly(doc):
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

        mirror_target = self.get_compare_mirror_target(target)
        if mirror_target is not None:
            self.sync_mirror_target_position(target, mirror_target)
        target.insert(tk.INSERT, clipboard_text)
        try:
            target.edit_modified(True)
        except tk.TclError:
            pass
        if mirror_target is not None:
            try:
                mirror_target.delete('sel.first', 'sel.last')
            except tk.TclError:
                pass
            try:
                mirror_target.insert(tk.INSERT, clipboard_text)
                mirror_target.edit_modified(True)
            except tk.TclError:
                pass
        target.see(tk.INSERT)
        self.set_last_active_editor_widget(target)
        if mirror_target is not None:
            self.notify_text_widget_changed(mirror_target)
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
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('font.title', 'Font'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color, padx=16, pady=16)

        tk.Label(dialog, text=self.tr('font.family', 'Font:'), bg=self.bg_color, fg=self.fg_color).grid(row=0, column=0, sticky='w', pady=(0, 8))
        tk.Label(dialog, text=self.tr('font.size', 'Size:'), bg=self.bg_color, fg=self.fg_color).grid(row=0, column=1, sticky='w', padx=(12, 0), pady=(0, 8))

        families = sorted(set(tkfont.families()))
        family_var = tk.StringVar(value=self.font_family)
        size_var = tk.StringVar(value=str(self.current_font_size))

        family_combo = ttk.Combobox(dialog, textvariable=family_var, values=families, state='readonly', width=28)
        family_combo.grid(row=1, column=0, sticky='ew')
        size_combo = ttk.Combobox(dialog, textvariable=size_var, values=[str(size) for size in range(6, 33)], state='readonly', width=8)
        size_combo.grid(row=1, column=1, sticky='ew', padx=(12, 0))

        preview = tk.Label(
            dialog,
            text=self.tr('font.preview', 'AaBbYyZz 0123456789'),
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
                messagebox.showwarning(
                    self.tr('font.invalid_size_title', 'Invalid Size'),
                    self.tr('font.invalid_size_message', 'Choose a valid font size.'),
                    parent=dialog
                )
                return

            self.font_family = family_var.get() or self.font_family
            self.current_font_size = max(self.min_font_size, min(self.max_font_size, new_size))
            self.update_font()
            self.update_status()
            dialog.destroy()

        tk.Button(
            button_row,
            text=self.tr('common.ok', 'OK'),
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
            text=self.tr('common.cancel', 'Cancel'),
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
        dialog = self.create_toplevel(self.root)
        dialog.title(self.tr('goto_line.title', 'Go To Line'))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color)

        tk.Label(dialog, text=self.tr('goto_line.prompt', 'Line Number:'), bg=self.bg_color, fg=self.fg_color)\
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
                    messagebox.showwarning(
                        self.tr('goto_line.invalid_title', 'Invalid'),
                        self.tr('goto_line.invalid_range', 'Line number out of range.'),
                        parent=dialog
                    )
            except ValueError:
                messagebox.showwarning(
                    self.tr('goto_line.invalid_title', 'Invalid'),
                    self.tr('goto_line.invalid_number', 'Enter a valid number.'),
                    parent=dialog
                )

        tk.Button(dialog, text=self.tr('common.go', 'Go'), command=goto)\
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
    multiprocessing.freeze_support()
    raw_args = get_process_launch_arguments()
    isolated_mode = '--isolated' in {arg.lower() for arg in raw_args}
    startup_files = [arg for arg in raw_args if arg.lower() != '--isolated']
    app_dir = get_notepadx_app_dir()
    if not isolated_mode and send_files_to_running_notepadx(app_dir, startup_files):
        sys.exit(0)
    NotepadX(isolated_session=isolated_mode, startup_files=startup_files)

