# XLSX Diff

A VS Code extension for visually comparing and editing `.xlsx` workbooks — directly inside the editor, with first-class Git / SVN / SCM integration.

![XLSX Diff screenshot](./media/preview.png)

---

## Features

- **Side-by-side spreadsheet diff** — view left and right workbooks in a clean table layout with colour-coded differences (modified / added / removed).
- **Git, SVN & SCM integration** — click any `.xlsx` file in the Source Control panel to open the diff view automatically.
- **Sheet tabs** — navigate between multiple worksheets; tabs show diff markers when a sheet contains changes.
- **Row filters** — toggle between _All rows_, _Diff rows only_, or _Same rows only_.
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

Open the **Source Control** panel, then click any `.xlsx` file listed under _Changes_. The XLSX diff view opens automatically instead of VS Code's built-in binary diff.

Git support works with VS Code's built-in Git extension. SVN support works with the `johnstoncode.svn-scm` extension and its `svn:` read-only resources.

### Keyboard shortcuts in the diff view

| Action    | Key                          |
| --------- | ---------------------------- |
| Next diff | Click **↓ Next Diff** button |
| Prev diff | Click **↑ Prev Diff** button |

### Editing a cell

1. Make sure the target file is a local `.xlsx` file (not a Git/SVN history version or read-only file).
2. **Double-click** a non-formula cell you want to edit.
3. Type the new value.
4. Press **Enter** or **Tab** to stage the change in the current session.
5. Press **⌘S** / **Ctrl+S** or click **Save** to write all staged edits to disk.
6. Press **Escape** to cancel without saving.

> Formula cells stay read-only. Cell edit undo / redo remains available after save during the current editor session. Structural edit history still resets on save.

---

## Settings

| Setting                     | Values                | Default | Description                                                                                                         |
| --------------------------- | --------------------- | ------- | ------------------------------------------------------------------------------------------------------------------- |
| `xlsx-diff.displayLanguage` | `auto`, `en`, `zh-CN` | `auto`  | Controls the language used in diff panel prompts and UI labels. `auto` follows the active VS Code display language. |
| `xlsx-diff.compareFormula` | `boolean` | `false` | Controls whether formula changes count as cell diffs when the displayed value is unchanged. |

---

## External tool integration

The extension already exposes a `vscode://` compare entrypoint. This repository ships CLI helpers that forward external tool arguments into that entrypoint and open the XLSX Diff panel in VS Code.

If you want stable commands on your `PATH`, run this once from the repository root:

```bash
npm link
```

That gives you:

- `xlsx-difftool` for tools that pass two workbook paths
- `xlsx-svn-diffwrap` for Subversion-style external diff arguments

If you are using a different publisher ID, set it before invoking either helper:

```bash
export VSCODE_XLSX_DIFF_EXTENSION_ID=your-publisher.xlsx-diff
```

### Direct CLI usage

```bash
xlsx-difftool left.xlsx right.xlsx
```

If you do not want to link the commands globally, invoke the script directly:

```bash
node /absolute/path/to/vscode-xlsx-diff/scripts/xlsx-difftool.mjs left.xlsx right.xlsx
```

### Git difftool integration

You can configure Git to use this extension as the difftool for `.xlsx` files.

**`~/.gitconfig`:**

```ini
[diff]
    tool = xlsx-vscode
[difftool "xlsx-vscode"]
    cmd = "xlsx-difftool" "$LOCAL" "$REMOTE" "$MERGED"
    prompt = false
```

If you prefer not to install the helper on `PATH`, point Git at the script directly:

```ini
[difftool "xlsx-vscode"]
    cmd = "node /absolute/path/to/vscode-xlsx-diff/scripts/xlsx-difftool.mjs" "$LOCAL" "$REMOTE" "$MERGED"
    prompt = false
```

**`.gitattributes`** in your repository:

```gitattributes
*.xlsx diff=xlsx
```

Running `git difftool` on an `.xlsx` file will then open the XLSX Diff panel in VS Code.

### SVN external diff integration

Subversion passes GNU `diff`-style arguments to external diff tools. The `xlsx-svn-diffwrap` helper consumes those arguments, forwards the workbook pair to VS Code, and exits with the standard `diff` status code `1`.

Direct command-line usage:

```bash
svn diff --diff-cmd xlsx-svn-diffwrap
```

If you prefer repository-wide or user-wide configuration, point Subversion's `diff-cmd` helper at the installed `xlsx-svn-diffwrap` command or the script path directly.

### Custom external tools

Any tool that can launch a command with two workbook paths can reuse the same helper:

```bash
xlsx-difftool "$LEFT_XLSX" "$RIGHT_XLSX"
```

The helpers are non-blocking: they ask VS Code to open the XLSX Diff panel and then return immediately.

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
