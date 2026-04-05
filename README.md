# Notepad-X

Notepad-X is a tabbed desktop text editor for Windows focused on fast everyday editing, side-by-side comparison, shared notes, and safer handling of large files.

It keeps a familiar desktop editor workflow while adding session restore, live search, folder-wide search, Markdown preview, syntax-aware themes, encrypted saves, note sharing, and crash recovery.

## Benchmark Snapshot

![EXE benchmark comparison](gfx/exe_benchmark.png)

The benchmark compares Microsoft Notepad and Notepad-X across launch time, memory use, CPU activity, and disk activity. In this snapshot, Notepad-X launches slightly faster on average and uses much less memory, while showing heavier startup CPU and disk work as the app initializes its extra editor features. The result is a lighter runtime footprint with a busier startup profile.

## Highlights

- Tabbed editing with persistent sessions and caret/scroll restore
- Live Find, Find/Replace, automatic cross-tab search, and folder-wide Find In Files
- Side-by-side compare mode inside the main window
- Shared notes with unread tracking, threaded replies, filters, and export
- Live Markdown preview in the right pane
- Syntax themes, syntax mode overrides, autocomplete, and spell check
- `Save and Run` for supported script and HTML files
- `Save As Encrypted` for passphrase-protected `.npxe` files
- Background large-file loading plus buffered virtual mode for very large files
- Crash recovery, conflict detection, and atomic save behavior
- Language switching with friendly locale names and Arabic RTL support
- Built-in Help, About dialog, and Windows shell integration toggle

## How It Works

### Tabs and Editing

- Every document opens in its own tab
- Tabs remember caret position and scroll position when switching away and back
- Duplicate opens focus the existing tab instead of creating a second copy
- GitHub-style line numbers can be shown or hidden from `View > Numbered Lines`
- Auto-indent, bracket matching, autocomplete, and spell check support normal editing workflows

### Search and Navigation

- `Ctrl+F` opens the Find panel
- The `Find` button and `F3` move forward through matches
- `Shift+F3` moves backward through matches
- The Replace panel uses the same live highlighting as Find
- Find and `Find In:` remember recent search strings and suggest them in a popup while you type
- `Up` and `Down` move through saved search suggestions, `Tab` accepts one, and `Enter` searches the current text
- Search continues across open tabs by default
- `Find In:` accepts a folder-search query, and `Browse` chooses the folder to search
- Pressing `Enter` in `Find In:` searches the most recently selected folder or prompts for one if none has been chosen yet

### Compare and Preview

- `View > Compare Tabs` opens two tabs side by side in the main window
- The right pane can also show live Markdown preview
- Compare mode and Markdown preview share the same right-side workspace
- `View > Currently Editing` adds a far-right sidebar for active-editor visibility on shared files

### Notes and Collaboration

- Notes attach to selected text through the context menu
- Notes support color tags, threaded replies, unread tracking, and per-editor state
- `F3` jumps unread notes when Find and Replace are closed
- `F4` cycles note markers using the current note filter
- Notes export to JSON or Markdown

### File Handling and Safety

- Normal saves use an atomic temp-file replace pattern
- Save conflict checks warn before overwriting a file that changed on disk
- Recovery restores into tabs instead of overwriting existing files
- Large files stay responsive by loading in the background
- Very large files can open in buffered virtual mode with limited editing features

## Feature Guide

### Grab Git

`File > Grab Git` downloads a public GitHub repository from a `username/project` entry, lets a folder be chosen for the download, and then opens one or more selected files from the downloaded project.

### Compare Mode

- `View > Compare Tabs` or `Ctrl+Q` opens compare mode
- The active tab stays on the left and the compared file appears on the right
- Both panes remain editable for normal tabs
- Compare mode has its own bottom status readout for the right pane
- Search commands follow whichever compare pane was clicked most recently
- `Ctrl+Shift+X` closes compare mode

<p align="center">
  <img src="gfx/compairing_files.png" alt="Compare mode screenshot in Notepad-X" width="900">
</p>

### Search, Replace, and Find In Files

- Live search highlights matches while typing without moving the caret
- Pressing `Enter` in the Find or Replace query box jumps to the first match from the top
- The Find panel includes a `Find In:` query field and `Browse`
- Recent search strings are stored for both `Find` and `Find In:` and appear in an autocomplete-style suggestion list
- `Up` and `Down` step through saved suggestions, `Tab` accepts one, and `Esc` closes the list
- Match highlighting follows the visible query across open tabs, including after switching tabs
- `Browse` chooses the folder for Find In Files without replacing the query text in the box
- Find In Files searches supported files under the selected folder and shows match counts in an `Open Selected` results window
- In compare mode, search follows the active pane

### Shared Notes

- Notes attach to selected text from the context menu
- Available note colors are Yellow, Green, Red, and Light Blue
- Notes show author, machine name, LAN IP, and local timestamp information
- Notes support threaded replies and unread tracking
- `View > Filter Notes` controls which notes are included when cycling with `F4`

<p align="center">
  <img src="gfx/codenotes.png" alt="Shared notes screenshot in Notepad-X" width="760">
</p>

### Markdown Preview

- `View > Preview Markdown` or `Ctrl+Shift+P` opens a live rendered preview in the right pane
- The preview updates while the source tab changes
- Headings, lists, quotes, rules, fenced code blocks, emphasis, inline code, and links are rendered
- Opening preview closes compare mode first

### Themes, Syntax, and Language

- Built-in themes are included out of the box
- `View > Syntax Theme > Create Theme` builds new custom themes
- `View > Syntax Mode` overrides automatic syntax detection for the active tab
- `Edit > Language` switches the visible UI language
- Locale entries use friendly native names when available

<p align="center">
  <img src="gfx/create_syntax_theme.png" alt="Create theme dialog in Notepad-X" width="420">
</p>

<p align="center">
  <img src="gfx/languages.png" alt="Language menu screenshot in Notepad-X" width="520">
</p>

### Save and Run

- `File > Save and Run` or `Ctrl+Shift+R` saves first, then launches the current file
- Supported launch targets include Python, JavaScript, PHP, batch, PowerShell, shell, and HTML
- HTML opens in the default browser
- Buffered large-file tabs and preview tabs cannot use Save and Run

### Encrypted Files

- `Save As Encrypted` creates `.npxe` encrypted copies
- Encrypted save suggests `file.ext.npxe` automatically
- Opening an `.npxe` file prompts for the passphrase
- Normal `Save` and `Save As` remain plain-text workflows

### Large Files

- Large files load in the background to keep the UI responsive
- Extremely large files can fall back to buffered virtual mode
- Virtual mode keeps navigation available while disabling direct editing and normal save operations
- `Save Copy As` remains available for copying the source file elsewhere
- Binary-like files open in a safer preview-style mode instead of being treated as normal editable text

### File Safety

- Atomic saves reduce the risk of partial writes
- Conflict detection warns before overwriting newer on-disk content
- Recovery can restore untitled work and modified file-backed tabs after a crash
- Session and note support files are also written atomically

## Main Shortcuts

- `Ctrl+W` Open
- `Ctrl+Shift+W` Open Project
- `Ctrl+Shift+G` Grab Git
- `Ctrl+T` New Tab
- `Ctrl+Shift+T` Close Tab
- `Ctrl+S` Save
- `Ctrl+Shift+S` Save all
- `Ctrl+Shift+Q` Save As
- `Ctrl+Shift+R` Save and Run
- `Ctrl+Shift+E` Save As Encrypted
- `Ctrl+E` Export Notes
- `Ctrl+P` Print
- `Ctrl+Shift+X` Close Markdown Preview / compare / Find or Replace / exit
- `Ctrl+F` Find
- `Ctrl+R` Replace
- `F3` Find next or next unread note
- `Shift+F3` Find previous
- `F4` Cycle notes
- `F7` Spell Check
- `Ctrl+G` Go To Line
- `Ctrl+Shift+F` Font
- `Ctrl+Shift+P` Preview Markdown
- `Ctrl+Shift+C` Currently Editing
- `Ctrl+Q` Compare Tabs
- `Ctrl+B` Show or hide status bar
- `Ctrl+Tab` Switch tab
- `Ctrl+PgUp` Top of document
- `Ctrl+PgDn` Bottom of document
- `Ctrl++` Zoom in
- `Ctrl+-` Zoom out
- `F11` Full Screen
