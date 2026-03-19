# Notepad-X

Notepad-X is a tabbed Tkinter text editor built for plain text, source code, code review notes, and safer large-file handling.

It keeps the simple feel of a classic desktop editor, but adds project opening, syntax highlighting, shared code notes, note approvals and denials, live find, persistent sessions, and a few playful extras.

## Features

- Tabbed editing with persistent file-backed sessions
- Recent files and `Open Project`
- Live Find and Find/Replace
- Syntax highlighting for many common code and config formats
- Large-file protection with buffered virtual mode
- Code notes on selected text
- Shared note sidecars with unread tracking
- Allow / Deny review flow with reviewer name and reply
- Status bar with line info, memory usage, note sync state, and live clock
- Word Wrap, Full Screen, zoom controls, font picker, printing
- Built-in Help viewer and About dialog

## Code Notes

Notepad-X lets you select a section of code, right-click, and attach a note to it.

Notes support:

- yellow highlight for normal notes
- green highlight for allowed changes
- red highlight for denied changes
- unread tracking between editors
- `F3` to jump unread notes
- `F4` to cycle all notes
- shared sidecar files for collaboration

## Large File Handling

Very large files are protected with a virtual buffered mode so Notepad-X does not try to load the entire file into the Tk text widget at once.

In large-file virtual mode:

- navigation stays usable
- line tracking still works
- the whole file is not fully loaded into memory
- editing and saving are disabled on purpose

## Included Assets

- `gfx/Notepad-X.ico`
- `audio/note.mp3`
- `audio/delete_note.mp3`
- `Notepad-X-help.txt`

## Project Files

Main app:

- `Notepad-X.py`

Help file:

- `Notepad-X-help.txt`

Runtime support files are intentionally excluded from Git:

- `Notepad-X.session.json`
- `Notepad-X.editor.json`
- `*.notepadx.notes.json`

## Requirements

- Windows
- Python 3.11 recommended
- Tkinter available in the Python install

## Run

```powershell
python Notepad-X.py
```

## Build EXE

If you want to package it with PyInstaller:

```powershell
python -m PyInstaller --noconfirm --clean --onedir --windowed --name "Notepad-X" --icon "gfx\Notepad-X.ico" --add-data "gfx;gfx" --add-data "audio;audio" --add-data "Notepad-X-help.txt;." "Notepad-X.py"
```

## Main Shortcuts

- `Ctrl+W` Open
- `Ctrl+Shift+W` Open Project
- `Ctrl+T` New Tab
- `Ctrl+Shift+T` Close Tab
- `Ctrl+S` Save
- `Ctrl+Shift+S` Save all
- `Ctrl+P` Print
- `Ctrl+Shift+X` Exit or close Find/Replace
- `Ctrl+F` Find
- `Ctrl+R` Replace
- `F3` Find next or next unread note
- `F4` Cycle notes
- `Ctrl+G` Go To Line
- `Ctrl+D` Date
- `Ctrl+Shift+D` Time/Date
- `Ctrl+Shift+F` Font
- `Ctrl+A` Select All
- `Ctrl+B` Show / hide status bar
- `Ctrl+Tab` Switch Tab
- `Ctrl+Left` Switch Tab Left
- `Ctrl+Right` Switch Tab Right
- `Ctrl++` Zoom in
- `Ctrl+-` Zoom out
- `F11` Full Screen

## Notes

Notepad-X is currently implemented as a single Python file for the main application logic. That keeps distribution simple, though the project can be modularized further later if needed.

