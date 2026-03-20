# Notepad-X

Notepad-X is a tabbed Windows text editor built with Tkinter for plain text, source code, shared code notes, and safer handling of large files.

It keeps a simple desktop-editor feel, but adds project opening, persistent sessions, syntax highlighting, live search, collaborative note sidecars, inline compare mode, recovery, and a built-in help viewer.

## Features

- Tabbed editing with persistent file-backed sessions
- Recent files and `Open Project`
- Drag-reorder tabs
- GitHub-style line number gutter with click-to-copy line support
- Local autocomplete popup with syntax keywords and current-document word matching
- Live Find and Find/Replace
- Optional `Search across all tabs`
- Syntax highlighting for many source and config formats
- Syntax theme presets and manual syntax override per tab
- Large-file protection with buffered virtual mode
- `Save Copy As` for huge read-only files
- Code notes on selected text
- Shared note sidecars with unread tracking
- Allow / Deny review flow with reviewer name and reply
- Export notes to JSON or Markdown
- Inline compare mode inside the main editor
- `Find Next` and `F3` follow the active pane during compare mode
- Autosave recovery for unsaved untitled tabs after a crash
- Crash logging for important failures and unhandled exceptions
- Status bar with line info, memory usage, note sync state, editor ID, and live clock
- Word Wrap, Sound toggle, Full Screen, zoom controls, font picker, printing
- `View > Numbered Lines` toggle with saved preference
- `View > Autocomplete` toggle with saved preference
- Built-in Help viewer and About dialog

## Compare Mode

Notepad-X can compare two open tabs side by side inside the main window.

- `View > Compare Tabs` or `Ctrl+Q` opens compare mode
- the normal editor stays usable on the left
- the compared file appears on the right
- syntax highlighting is applied on the compare side too
- `Find Next` and `F3` follow whichever compare pane you last clicked
- `Ctrl+Shift+X` closes compare mode

## Line Numbers

Notepad-X includes a GitHub-style line number gutter on the left side of the editor.

- line numbers are enabled by default
- `View > Numbered Lines` hides or shows the gutter
- the setting is remembered across launches
- clicking a line number copies that whole line to the clipboard
- a small in-window notification appears beside the clicked gutter line

## Autocomplete

Notepad-X includes a lightweight local autocomplete system for normal editable tabs.

- enabled by default
- `View > Autocomplete` hides or shows it
- the setting is remembered across launches
- suggestions combine syntax keywords with matching words from the current tab
- the popup appears under the caret while typing
- `Up` / `Down` move through suggestions
- `Tab` or `Enter` accepts the selected suggestion
- `Esc` closes the popup

## Code Notes

You can select text, right-click, and attach a note to the selection.

Notes support:

- yellow highlight for normal notes
- green highlight for allowed changes
- red highlight for denied changes
- note author and timestamp display
- unread tracking between editors
- `F3` to jump unread notes
- `F4` to cycle notes
- shared sidecar files for collaboration
- export to JSON or Markdown

## Large File Handling

Very large files are protected with a buffered virtual mode so Notepad-X does not try to load the entire file into the Tk text widget at once.

In large-file virtual mode:

- navigation stays usable
- line tracking still works
- only a moving window of the file is loaded
- editing is disabled
- direct saving is disabled
- `Save Copy As` is available for copying the source file elsewhere

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
- `Ctrl+Shift+Q` Save Copy As
- `Ctrl+P` Print
- `Ctrl+E` Export Notes
- `Ctrl+Shift+X` Close compare / close Find or Replace / exit
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
- `Ctrl+Q` Compare Tabs
- `Ctrl++` Zoom in
- `Ctrl+-` Zoom out
- `F11` Full Screen

## Support Files

Notepad-X keeps some support files next to the app, including:

- `Notepad-X.session.json`
- `Notepad-X.editor.json`
- `Notepad-X.recovery.json`
- `Notepad-X.crash.log`
- `*.notepadx.notes.json`

On Windows, Notepad-X attempts to mark support files as hidden.

## Notes

Notepad-X is currently implemented as a single Python file for the main application logic. That keeps distribution simple, though it can be modularized further later if needed.
