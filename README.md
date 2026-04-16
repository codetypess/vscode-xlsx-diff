# XLSX Diff

VS Code extension for comparing `.xlsx` workbooks with a spreadsheet-oriented diff view.

Current implementation includes:

- `fastxlsx`-based workbook loading
- side-by-side Webview diff UI
- sheet tabs
- row pagination
- filters for `All / Diffs / Same`
- diff navigation
- URI handler for external launches
- Git difftool bridge script

## Commands

- `XLSX Diff: Compare Two XLSX Files`
- `XLSX Diff: Compare Active XLSX With...`

## Local development

1. Install dependencies:

   ```bash
   npm install
   ```

   This project uses `tsx` to run the ESM build/debug script.

2. Start the extension in an Extension Development Host:

   - Press `F5` in VS Code
   - Open the Command Palette
   - Run one of the `XLSX Diff` commands

3. Run checks:

   ```bash
   npm run compile
   npm test
   ```

The test setup is configured to reuse a locally installed VS Code when possible so it does not need to download a fresh build every run.

## Git difftool setup

The repository includes `scripts/xlsx-difftool.mjs`, which converts Git's `$LOCAL` and `$REMOTE` files into a `vscode://` URI handled by the extension.

Example Git config:

```ini
[diff]
    tool = xlsx-vscode
[difftool "xlsx-vscode"]
    cmd = "node /absolute/path/to/vscode-xlsx-diff/scripts/xlsx-difftool.mjs" "$LOCAL" "$REMOTE" "$MERGED"
    prompt = false
```

Example `.gitattributes`:

```gitattributes
*.xlsx diff=xlsx
```

If you publish the extension under a different identifier, set:

```bash
export VSCODE_XLSX_DIFF_EXTENSION_ID=your-publisher.xlsx-diff
```

before running the difftool script.

## How the Git bridge works

1. Git calls `xlsx-difftool` with `$LOCAL` and `$REMOTE`
2. The script builds a `vscode://<publisher>.xlsx-diff/compare?...` URI
3. VS Code routes the URI to the extension's `UriHandler`
4. The extension opens the XLSX diff panel with those two files

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
