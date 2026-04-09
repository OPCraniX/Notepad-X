# Notepad-X

Notepad-X is a Windows tabbed text editor built for everyday writing, code editing, side-by-side comparison, shared notes, and safer handling of very large files.

It keeps the familiar single-window desktop editor workflow, then layers in compare mode, live Markdown preview, syntax-aware editing, recovery tools, remote file access, and a growing set of coding helpers without turning into a full IDE.

## Highlights

- Tabbed editing with session restore, caret/scroll restore, and crash recovery
- Side-by-side compare mode and live Markdown preview in the shared right pane
- Compare multi-edit for mirrored editing while comparing normal text tabs
- Smarter autocomplete with document symbols, keywords, builtins, and nearby project file suggestions
- Symbol navigation for the current file and project-wide symbol lookup
- Code folding with gutter `+` / `-` controls plus `F9`, `Shift+F9`, and `Ctrl+F9`
- Auto-pairing for brackets and quotes with matching-pair highlight
- Minimap, breadcrumbs, numbered lines, and synced `PgUp` / `PgDn` navigation in compare/preview
- Real-time diagnostics for warnings and syntax problems, including hover tooltips
- Docked command panel for quick commands and editor actions
- SSH-style remote open/save for `user@host:/absolute/path` files using `scp`
- Autosave, atomic saves, backup snapshots, and large-file virtual mode
- Tab right-click actions for save, copy path/name, reveal, and move-to-new-window
- A dedicated `Settings` menu for editor toggles and checkmark-based options
- Shared notes with replies, unread tracking, filters, and export
- `Save and Run`, encrypted `.npxe` saves, theme support, locale switching, and Windows shell integration

## What Notepad-X Does Well

### Editing and Coding

- Auto-indent keeps line indentation flowing naturally
- Auto-pair inserts matching brackets and quotes while typing
- Matching bracket highlight shows paired `()`, `[]`, and `{}`
- Autocomplete suggests:
  document words
  symbols from the current file
  symbols from other open files
  syntax keywords
  Python keywords and builtins
  nearby module/file names from the working folder
- Symbol navigation supports:
  current-file symbol lookup with `Ctrl+Shift+O`
  project symbol lookup with `Ctrl+Alt+P`
- Folding supports:
  gutter fold boxes
  `F9` to toggle the nearest foldable section containing the caret
  `Shift+F9` to collapse all folds
  `Ctrl+F9` to expand all folds
- Diagnostics can flag:
  trailing whitespace
  long lines over 140 characters
  Python syntax errors
  JSON parse errors
  XML parse errors
- Log and crash files can also surface traceback/error lines as hoverable diagnostics
- Hovering a highlighted diagnostic line shows the issue in a small info bubble

### Navigation and Search

- `Ctrl+F` opens live Find
- `Ctrl+R` opens Replace
- `Find In Files` searches across folders and opens results in a dedicated window
- `Ctrl+G` jumps to a line
- Minimap provides a compressed file overview and click-to-jump navigation
- Breadcrumbs show the current file/path context and nearest symbol
- Clicking the filename in breadcrumbs copies just the filename; clicking the path copies the full location
- `Ctrl+PgUp` and `Ctrl+PgDn` jump to the top and bottom of the current document
- Optional synced `PgUp` / `PgDn` keeps the left editor and right compare/preview pane moving together

### Compare, Preview, and Collaboration

- `Ctrl+Q` opens compare mode
- Compare mode keeps the source tab on the left and the selected comparison tab on the right
- Compare multi-edit can mirror typing, delete, paste, tab, and newline edits across both panes
- `Ctrl+Shift+P` opens live Markdown preview in the same right-side workspace
- Right-clicking a tab opens file-specific actions like save, copy name/path, reveal, and move to a new Notepad-X window
- `Ctrl+Shift+C` toggles the `Currently Editing` sidebar for shared file visibility
- Shared notes support:
  colored note markers
  threaded replies
  unread tracking
  note cycling
  JSON and Markdown export

### Files, Safety, and Recovery

- Normal saves use atomic replace behavior
- Autosave can write changes back automatically
- Backup snapshots are created on save
- Session restore remembers open files, view state, window size/state, and major editor toggles
- Crash recovery can restore untitled work and modified tabs
- External-file conflict detection warns before overwriting newer disk content
- Large files can load in the background
- Extremely large files can fall back to virtual mode for responsive navigation
- `Save As Encrypted` creates passphrase-protected `.npxe` files
- `Open Remote (SSH)` fetches and saves remote files through `scp`

## Command Panel

Press `Ctrl+Shift+K` to open the bottom command panel.

It supports quick built-in commands like:

- `:help`
- `:symbols`
- `:project-symbols`
- `:fold`
- `:fold-all`
- `:unfold-all`
- `:lint`
- `:preview`
- `:compare`
- `:remote user@host:/absolute/path`
- `:autosave on|off`
- `:minimap on|off`
- `:diagnostics on|off`

It can also run shell commands from the current document directory. `cls` or `clear` wipes the output pane, submitted commands are stored in a visible history list, and the panel height can be resized by dragging its top grip. It is a docked command runner, not a full interactive terminal emulator.

## Remote Files

- Use `File > Open Remote (SSH)` or `Ctrl+Alt+O`
- Remote paths use this format:
  `user@host:/absolute/path/to/file`
- Notepad-X pulls the file locally through `scp`, edits it normally, and pushes changes back on save
- This is a file open/save workflow, not a full remote workspace explorer

## Large Files

- Large files can load asynchronously to keep the UI responsive
- Very large files can open in virtual mode
- Virtual mode keeps navigation available while limiting direct editing features
- Binary-like content opens in preview-style mode instead of normal editing mode

## Settings Menu

- `Settings` collects the checkmark/toggle options that used to be split across `Edit` and `View`
- It includes status bar, numbered lines, autocomplete, spell check, auto-pair, compare multi-edit, minimap, breadcrumbs, diagnostics, autosave, word wrap, Markdown preview, sync page navigation, sound, and `Edit with Notepad-X`

## Main Shortcuts

- `Ctrl+W` Open
- `Ctrl+Shift+W` Open Project
- `Ctrl+Alt+O` Open Remote (SSH)
- `Ctrl+Shift+G` Grab Git
- `Ctrl+T` New Tab
- `Ctrl+Shift+T` Close Tab
- `Ctrl+S` Save
- `Ctrl+Shift+S` Save All
- `Ctrl+Shift+Q` Save As
- `Ctrl+Shift+E` Save As Encrypted
- `Ctrl+Shift+R` Save and Run
- `Ctrl+E` Export Notes
- `Ctrl+P` Print
- `Ctrl+F` Find
- `Ctrl+R` Replace
- `Ctrl+Shift+K` Command Panel
- `Ctrl+Shift+O` Jump to Symbol
- `Ctrl+Alt+P` Project Symbols
- `Ctrl+G` Go To Line
- `F3` Find Next or next unread note
- `Shift+F3` Find Previous
- `F4` Cycle Notes
- `F7` Toggle Spell Check
- `F9` Toggle Fold at Caret Section
- `Shift+F9` Collapse All Folds
- `Ctrl+F9` Expand All Folds
- `Ctrl+Shift+P` Preview Markdown
- `Ctrl+Shift+C` Currently Editing
- `Ctrl+Q` Compare Tabs
- `Ctrl+Shift+X` Close compare/preview/panels or exit
- `Ctrl+B` Toggle Status Bar
- `Ctrl+Tab` Switch Tab
- `Ctrl+PgUp` Top of Document
- `Ctrl+PgDn` Bottom of Document
- `Ctrl++` Zoom In
- `Ctrl+-` Zoom Out
- `Ctrl+Mouse Wheel` Zoom
- `F11` Full Screen

## Benchmark Snapshot

![EXE benchmark comparison](gfx/exe_benchmark.png)

The latest packaged EXE benchmark for Notepad-X `1.0.6.0` was run on April 8, 2026 over 10 launches on Windows 11 Pro 25H2. In this snapshot, Notepad-X opens faster than Microsoft Notepad on average, uses dramatically less memory once running, and stays lean after startup, while paying for that richer editor startup path with a noticeably larger burst of CPU and disk activity during initialization. In practice, the results show the current build getting the window on screen quickly and then settling down after loading its extra features such as recovery, compare tools, diagnostics, minimap, and theming support.
