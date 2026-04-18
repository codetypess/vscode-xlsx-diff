# XLSX Diff

A VS Code extension for visually comparing and editing `.xlsx` workbooks — directly inside the editor, with first-class Git / SCM integration.

![XLSX Diff screenshot](./media/preview.png)

---

## Features

- **Side-by-side spreadsheet diff** — view left and right workbooks in a clean table layout with colour-coded differences (modified / added / removed).
- **Git & SCM integration** — click any `.xlsx` file in the Source Control panel to open the diff view automatically.
- **Sheet tabs** — navigate between multiple worksheets; tabs show diff markers when a sheet contains changes.
- **Row filters** — toggle between *All rows*, *Diff rows only*, or *Same rows only*.
- **Diff navigation** — jump to the previous / next changed cell with a single click or keyboard shortcut.
- **Pagination** — large workbooks are split into pages for performance.
- **Row height sync** — left and right rows are height-matched so multi-line cells align correctly.
- **Formula display** — cells containing formulas show an `fx` badge.
- **Cell editing** — double-click any cell in a local (non-read-only) workbook to edit it inline. Press **Enter** or **Tab** to save; press **Escape** to cancel.
- **Swap** — swap left and right sides with one click.
- **Auto-reload** — the diff view refreshes automatically when a watched file changes on disk.
- **Bilingual UI** — Chinese (Simplified) and English, following the VS Code display language setting.

---

## Usage

### Open the single-file editor

Opening a local `.xlsx` file now uses the XLSX editor automatically. You can still run **XLSX Diff: Open XLSX Editor** from the Explorer, editor context menu, or Command Palette if you want to open it explicitly.

In the editor view you can:

- browse sheets with tabs
- move between cells with the arrow keys, `Enter`, and `Tab`
- page through large used ranges with `PageUp` and `PageDown`
- press `Cmd/Ctrl+F` to search workbook values and formulas, then step through matches
- press `Cmd/Ctrl+G` to jump to a cell such as `A1` or `Sheet1!B2`
- double-click a cell to edit it inline
- paste tab-separated spreadsheet data into the current selection rectangle
- press **Backspace** or **Delete** to clear the selected cell
- press **Cmd/Ctrl+Z** and **Shift+Cmd/Ctrl+Z** or **Ctrl+Y** for session undo / redo
- press **⌘S** / **Ctrl+S** or click **Save** to write staged edits back to the workbook
- use the toolbar **Undo**, **Redo**, **Search**, and **Go** controls for mouse-driven editing
- keep a workbook open in read-only mode when the file system is not writable
- leave formula cells read-only; formula cells are visible but cannot be edited or overwritten from the editor

### Compare two files

1. Right-click an `.xlsx` file in the **Explorer** and choose **XLSX Diff: Compare Active XLSX With…** — then pick the second file.
2. Or open the Command Palette (`⌘⇧P` / `Ctrl+Shift+P`) and run **XLSX Diff: Compare Two XLSX Files**.

### Compare from Source Control

Open the **Source Control** panel, then click any `.xlsx` file listed under *Changes*. The XLSX diff view opens automatically instead of VS Code's built-in binary diff.

### Keyboard shortcuts in the diff view

| Action | Key |
|---|---|
| Next diff | Click **↓ Next Diff** button |
| Prev diff | Click **↑ Prev Diff** button |
| Next page | Click **→ Next Page** button |
| Prev page | Click **← Prev Page** button |

### Editing a cell

1. Make sure the target file is a local `.xlsx` file (not a Git history version or read-only file).
2. **Double-click** a non-formula cell you want to edit.
3. Type the new value.
4. Press **Enter** or **Tab** to stage the change in the current session.
5. Press **⌘S** / **Ctrl+S** or click **Save** to write all staged edits to disk.
6. Press **Escape** to cancel without saving.

> Formula cells stay read-only. Session undo / redo only affects staged edits that have not been saved yet.

---

## Settings

| Setting | Values | Default | Description |
|---|---|---|---|
| `xlsx-diff.displayLanguage` | `auto`, `en`, `zh-cn` | `auto` | Controls the language used in diff panel prompts and UI labels. `auto` follows the active VS Code display language. |

---

## Git difftool integration

You can configure Git to use this extension as your difftool for `.xlsx` files.

**`~/.gitconfig`:**

```ini
[diff]
    tool = xlsx-vscode
[difftool "xlsx-vscode"]
    cmd = "node /absolute/path/to/vscode-xlsx-diff/scripts/xlsx-difftool.mjs" "$LOCAL" "$REMOTE" "$MERGED"
    prompt = false
```

**`.gitattributes`** in your repository:

```gitattributes
*.xlsx diff=xlsx
```

Running `git difftool` on an `.xlsx` file will then open the XLSX Diff panel in VS Code.

If you are using a different publisher ID, set the environment variable before running the script:

```bash
export VSCODE_XLSX_DIFF_EXTENSION_ID=your-publisher.xlsx-diff
```

---

## Local development

```bash
# Install dependencies
npm install

# Start watch build (extension + webview)
npm run watch

# Press F5 in VS Code to launch the Extension Development Host
# Then right-click an .xlsx file in Explorer or SCM view

# Type-check all targets
npm run check-types

# Run tests
npm test
```

---

## License

MIT. See [LICENSE](./LICENSE).

## Current behavior notes

- Comparison is value/formula oriented
- renamed sheets are matched by content signature
- merged range changes are surfaced as sheet-level warnings
- row pagination defaults to `200` rows per page
- the current loader scans each sheet's used range eagerly

## Known gaps

- no style-level diff visualization yet
- no chart / pivot / macro diff yet
- no merge tool flow yet
- large sheets are paginated in the UI, but workbook parsing is still eager
