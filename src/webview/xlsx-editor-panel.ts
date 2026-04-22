import * as path from "node:path";
import * as vscode from "vscode";
import { DEFAULT_PAGE_SIZE } from "../constants";
import { getCellAddress, getColumnNumber } from "../core/model/cells";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import {
    type CellEdit,
    type SheetEdit,
    type SheetViewEdit,
    type WorkbookEditState,
} from "../core/fastxlsx/write-cell-value";
import type { EditorPanelState, EditorRenderModel, WorkbookSnapshot } from "../core/model/types";
import { getHtmlLanguageTag, isChineseDisplayLanguage } from "../display-language";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    type EditorSheetEntry,
    normalizeEditorPanelState,
    setActiveEditorSheet,
    setEditorViewportStartRow,
    setSelectedEditorCell,
} from "./editor-render-model";
import { hasLockedView } from "./view-lock";
import { XlsxEditorDocument } from "./xlsx-editor-document";
import { getWorkbookResourceName } from "../workbook/resource-uri";
import { rememberRecentWorkbookResourceUri } from "../scm/recent-workbook-resource-context";

interface SearchOptions {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

type WebviewMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "addSheet" }
    | { type: "deleteSheet"; sheetKey: string }
    | { type: "renameSheet"; sheetKey: string }
    | { type: "setViewportStartRow"; rowNumber: number }
    | {
          type: "search";
          query: string;
          direction: "next" | "prev";
          options: SearchOptions;
      }
    | { type: "gotoCell"; reference: string }
    | { type: "selectCell"; rowNumber: number; columnNumber: number }
    | {
          type: "setPendingEdits";
          edits: Array<{
              sheetKey: string;
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
      }
    | { type: "requestSave" }
    | { type: "pendingEditStateChanged"; hasPendingEdits: boolean }
    | { type: "undoSheetEdit" }
    | { type: "redoSheetEdit" }
    | { type: "toggleViewLock"; rowCount: number; columnCount: number }
    | { type: "reload" };

export interface XlsxEditorPanelController {
    onPendingStateChanged(state: WorkbookEditState): Promise<void> | void;
    onRequestSave(): Promise<void> | void;
    onRequestRevert(): Promise<void> | void;
}

interface WebviewStrings {
    loading: string;
    reload: string;
    size: string;
    modified: string;
    sheet: string;
    rows: string;
    noRows: string;
    visibleRows: string;
    readOnly: string;
    save: string;
    lockView: string;
    unlockView: string;
    addSheet: string;
    deleteSheet: string;
    renameSheet: string;
    renameSheetPrompt: string;
    renameSheetTitle: string;
    sheetNameEmpty: string;
    sheetNameDuplicate: string;
    sheetNameTooLong: string;
    sheetNameInvalidChars: string;
    undo: string;
    redo: string;
    searchPlaceholder: string;
    findPrev: string;
    findNext: string;
    gotoPlaceholder: string;
    goto: string;
    totalSheets: string;
    totalRows: string;
    nonEmptyCells: string;
    selectedCell: string;
    noCellSelected: string;
    mergedRanges: string;
    pendingChanges: string;
    noRowsAvailable: string;
    readOnlyBadge: string;
    localChangesBlockedReload: string;
    confirmReloadDiscard: string;
    discardChangesAndReload: string;
    keepEditing: string;
    displayLanguageRefreshBlocked: string;
    noSearchMatches: string;
    invalidCellReference: string;
    invalidSearchPattern: string;
    searchRegex: string;
    searchMatchCase: string;
    searchWholeWord: string;
}

function getNonce(): string {
    return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function toErrorMessage(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
}

function escapeWatcherGlobSegment(value: string): string {
    return value.replace(/[{}\[\]*?]/g, "[$&]");
}

function escapeRegex(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function createSearchPattern(query: string, options: SearchOptions): RegExp {
    const source = options.isRegexp ? query.trim() : escapeRegex(query.trim());
    const wrappedSource = options.wholeWord ? `\\b(?:${source})\\b` : source;
    return new RegExp(wrappedSource, options.matchCase ? "" : "i");
}

function getWebviewStrings(): WebviewStrings {
    if (isChineseDisplayLanguage()) {
        return {
            loading: "正在加载 XLSX 编辑器...",
            reload: "刷新",
            undo: "撤销",
            redo: "重做",
            searchPlaceholder: "搜索值或公式",
            findPrev: "上一个",
            findNext: "下一个",
            gotoPlaceholder: "A1 或 Sheet1!B2",
            goto: "定位",
            size: "大小",
            modified: "修改时间",
            sheet: "工作表",
            rows: "行",
            noRows: "无行",
            visibleRows: "可见行",
            readOnly: "只读",
            save: "保存",
            lockView: "锁定视图",
            unlockView: "取消锁定视图",
            addSheet: "添加工作表",
            deleteSheet: "删除工作表",
            renameSheet: "重命名工作表",
            renameSheetPrompt: "输入新的工作表名称",
            renameSheetTitle: "重命名工作表",
            sheetNameEmpty: "工作表名称不能为空。",
            sheetNameDuplicate: "工作表名称已存在。",
            sheetNameTooLong: "工作表名称不能超过 31 个字符。",
            sheetNameInvalidChars: "工作表名称不能包含 \\ / ? * [ ] : 等字符。",
            totalSheets: "总工作表",
            totalRows: "总行数",
            nonEmptyCells: "非空单元格",
            selectedCell: "当前单元格",
            noCellSelected: "未选择",
            mergedRanges: "合并区域",
            pendingChanges: "待保存修改",
            noRowsAvailable: "当前视图没有可显示的行。",
            readOnlyBadge: "只读模式",
            localChangesBlockedReload:
                "工作簿文件已在磁盘上变化。请先保存或放弃当前未保存修改，再刷新。",
            confirmReloadDiscard: "刷新会丢弃当前未保存修改，是否继续？",
            discardChangesAndReload: "放弃修改并刷新",
            keepEditing: "继续编辑",
            displayLanguageRefreshBlocked: "当前有未保存修改，语言变更将在保存或手动刷新后生效。",
            noSearchMatches: "没有找到匹配的单元格。",
            invalidCellReference:
                "无法定位该单元格，请使用 A1 或 Sheet1!B2 格式，并确保目标在当前工作簿范围内。",
            invalidSearchPattern: "搜索表达式无效。",
            searchRegex: "使用正则表达式",
            searchMatchCase: "区分大小写",
            searchWholeWord: "匹配整个单词",
        };
    }

    return {
        loading: "Loading XLSX editor...",
        reload: "Reload",
        undo: "Undo",
        redo: "Redo",
        searchPlaceholder: "Search values or formulas",
        findPrev: "Prev Match",
        findNext: "Next Match",
        gotoPlaceholder: "A1 or Sheet1!B2",
        goto: "Go",
        size: "Size",
        modified: "Modified",
        sheet: "Sheet",
        rows: "Rows",
        noRows: "No rows",
        visibleRows: "Visible rows",
        readOnly: "Read-only",
        save: "Save",
        lockView: "Lock View",
        unlockView: "Unlock View",
        addSheet: "Add Sheet",
        deleteSheet: "Delete Sheet",
        renameSheet: "Rename Sheet",
        renameSheetPrompt: "Enter a new sheet name",
        renameSheetTitle: "Rename Sheet",
        sheetNameEmpty: "Sheet name cannot be empty.",
        sheetNameDuplicate: "A sheet with this name already exists.",
        sheetNameTooLong: "Sheet names must be 31 characters or fewer.",
        sheetNameInvalidChars: "Sheet names cannot contain \\ / ? * [ ] or :.",
        totalSheets: "Sheets",
        totalRows: "Rows",
        nonEmptyCells: "Non-empty cells",
        selectedCell: "Selected cell",
        noCellSelected: "None",
        mergedRanges: "Merged ranges",
        pendingChanges: "Pending changes",
        noRowsAvailable: "No rows available in this view.",
        readOnlyBadge: "Read-only",
        localChangesBlockedReload:
            "The workbook changed on disk. Save or discard your pending edits before reloading.",
        confirmReloadDiscard: "Reloading will discard your pending edits. Continue?",
        discardChangesAndReload: "Discard Changes and Reload",
        keepEditing: "Keep Editing",
        displayLanguageRefreshBlocked:
            "Pending edits are open. Display language changes will apply after you save or reload the editor.",
        noSearchMatches: "No matching cells were found.",
        invalidCellReference:
            "Unable to locate that cell. Use A1 or Sheet1!B2 and stay within the workbook range.",
        invalidSearchPattern: "The search pattern is invalid.",
        searchRegex: "Use Regular Expression",
        searchMatchCase: "Match Case",
        searchWholeWord: "Match Whole Word",
    };
}

interface WorkingSheetEntry extends EditorSheetEntry {
    sheet: WorkbookSnapshot["sheets"][number];
}

interface StructuralSnapshot {
    state: EditorPanelState;
    sheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
}

interface StructuralHistoryEntry {
    before: StructuralSnapshot;
    after: StructuralSnapshot;
    resetPendingHistory: boolean;
}

function cloneCellEdit(edit: CellEdit): CellEdit {
    return { ...edit };
}

function cloneSheetEdit(edit: SheetEdit): SheetEdit {
    return { ...edit };
}

function cloneViewEdit(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
    };
}

function cloneSheetSnapshot(
    sheet: WorkbookSnapshot["sheets"][number]
): WorkbookSnapshot["sheets"][number] {
    return {
        ...sheet,
        mergedRanges: [...sheet.mergedRanges],
        freezePane: sheet.freezePane ? { ...sheet.freezePane } : null,
        cells: { ...sheet.cells },
    };
}

function createFreezePaneSnapshot(
    columnCount: number,
    rowCount: number
): WorkbookSnapshot["sheets"][number]["freezePane"] {
    if (columnCount <= 0 && rowCount <= 0) {
        return null;
    }

    return {
        columnCount,
        rowCount,
        topLeftCell: getCellAddress(rowCount + 1, columnCount + 1),
        activePane:
            rowCount > 0 && columnCount > 0
                ? "bottomRight"
                : rowCount > 0
                  ? "bottomLeft"
                  : "topRight",
    };
}

function areFreezePaneCountsEqual(
    left: WorkbookSnapshot["sheets"][number]["freezePane"],
    right: WorkbookSnapshot["sheets"][number]["freezePane"]
): boolean {
    return (
        (left?.columnCount ?? 0) === (right?.columnCount ?? 0) &&
        (left?.rowCount ?? 0) === (right?.rowCount ?? 0)
    );
}

function cloneSheetEntry(entry: WorkingSheetEntry): WorkingSheetEntry {
    return {
        key: entry.key,
        index: entry.index,
        sheet: cloneSheetSnapshot(entry.sheet),
    };
}

function cloneEditorState(state: EditorPanelState): EditorPanelState {
    return {
        ...state,
        selectedCell: state.selectedCell ? { ...state.selectedCell } : null,
    };
}

function reindexWorkingSheetEntries(sheetEntries: WorkingSheetEntry[]): WorkingSheetEntry[] {
    return sheetEntries.map((entry, index) => ({
        ...entry,
        index,
    }));
}

function createWorkingSheetEntries(workbook: WorkbookSnapshot): WorkingSheetEntry[] {
    return workbook.sheets.map((sheet, index) => ({
        key: `sheet:${index}`,
        index,
        sheet: cloneSheetSnapshot(sheet),
    }));
}

export class XlsxEditorPanel {
    private static readonly panels = new Map<number, XlsxEditorPanel>();
    private static nextPanelId = 1;

    private readonly panel: vscode.WebviewPanel;
    private readonly extensionUri: vscode.Uri;
    private readonly document: XlsxEditorDocument;
    private readonly controller: XlsxEditorPanelController;
    private readonly disposables: vscode.Disposable[] = [];
    private readonly fileWatchers: vscode.Disposable[] = [];
    private readonly panelId: number;

    private workbookUri: vscode.Uri;
    private workbook: WorkbookSnapshot | null = null;
    private state: EditorPanelState = {
        activeSheetKey: null,
        viewportStartRow: 1,
        selectedCell: null,
    };
    private isWebviewReady = false;
    private hasPendingRender = false;
    private isReloading = false;
    private hasQueuedReload = false;
    private autoRefreshTimer: ReturnType<typeof setTimeout> | undefined;
    private suppressAutoRefreshUntil = 0;
    private hasWarnedPendingExternalChange = false;
    private pendingCellEdits: CellEdit[] = [];
    private pendingSheetEdits: SheetEdit[] = [];
    private pendingViewEdits: SheetViewEdit[] = [];
    private workingSheetEntries: WorkingSheetEntry[] = [];
    private sheetUndoStack: StructuralHistoryEntry[] = [];
    private sheetRedoStack: StructuralHistoryEntry[] = [];
    private nextNewSheetId = 1;

    private constructor(
        panel: vscode.WebviewPanel,
        extensionUri: vscode.Uri,
        document: XlsxEditorDocument,
        controller: XlsxEditorPanelController,
        workbookUri: vscode.Uri,
        panelId: number
    ) {
        this.panel = panel;
        this.extensionUri = extensionUri;
        this.document = document;
        this.controller = controller;
        this.workbookUri = workbookUri;
        this.panelId = panelId;
        this.panel.webview.options = {
            enableScripts: true,
            localResourceRoots: [extensionUri],
        };

        this.panel.webview.html = this.getHtml();
        this.panel.onDidDispose(
            () => {
                XlsxEditorPanel.panels.delete(this.panelId);
                this.dispose();
            },
            null,
            this.disposables
        );
        this.panel.webview.onDidReceiveMessage(
            (message: WebviewMessage) => {
                void this.handleMessage(message);
            },
            null,
            this.disposables
        );
        this.refreshFileWatchers();
    }

    public static async resolveCustomEditor(
        extensionUri: vscode.Uri,
        document: XlsxEditorDocument,
        panel: vscode.WebviewPanel,
        controller: XlsxEditorPanelController
    ): Promise<void> {
        rememberRecentWorkbookResourceUri(document.uri, "customEditorPanel");
        const panelId = XlsxEditorPanel.nextPanelId;
        XlsxEditorPanel.nextPanelId += 1;
        panel.title = getWorkbookResourceName(document.uri);
        const instance = new XlsxEditorPanel(
            panel,
            extensionUri,
            document,
            controller,
            document.uri,
            panelId
        );
        XlsxEditorPanel.panels.set(panelId, instance);
        await instance.enqueueReload();
    }

    public static async refreshDocument(
        document: XlsxEditorDocument,
        options: { silent?: boolean; clearPendingEdits?: boolean } = {}
    ): Promise<void> {
        await Promise.all(
            [...XlsxEditorPanel.panels.values()]
                .filter((panel) => panel.document === document)
                .map((panel) => panel.enqueueReload(options))
        );
    }

    public static async refreshAll(): Promise<void> {
        await Promise.all(
            [...XlsxEditorPanel.panels.values()].map((panel) =>
                panel.refreshForDisplayLanguageChange()
            )
        );
    }

    private dispose(): void {
        if (this.autoRefreshTimer) {
            clearTimeout(this.autoRefreshTimer);
            this.autoRefreshTimer = undefined;
        }

        this.disposeFileWatchers();

        for (const disposable of this.disposables) {
            disposable.dispose();
        }
    }

    private disposeFileWatchers(): void {
        for (const disposable of this.fileWatchers) {
            disposable.dispose();
        }

        this.fileWatchers.length = 0;
    }

    private refreshFileWatchers(): void {
        this.disposeFileWatchers();

        if (this.workbookUri.scheme !== "file") {
            return;
        }

        const watcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(
                vscode.Uri.file(path.dirname(this.workbookUri.fsPath)),
                escapeWatcherGlobSegment(path.basename(this.workbookUri.fsPath))
            )
        );
        const scheduleRefresh = () => {
            this.scheduleAutoRefresh();
        };

        this.fileWatchers.push(watcher);
        this.fileWatchers.push(watcher.onDidChange(scheduleRefresh));
        this.fileWatchers.push(watcher.onDidCreate(scheduleRefresh));
        this.fileWatchers.push(watcher.onDidDelete(scheduleRefresh));
    }

    private scheduleAutoRefresh(): void {
        if (Date.now() < this.suppressAutoRefreshUntil) {
            if (this.autoRefreshTimer) {
                clearTimeout(this.autoRefreshTimer);
                this.autoRefreshTimer = undefined;
            }

            return;
        }

        if (this.document.hasPendingEdits()) {
            if (!this.hasWarnedPendingExternalChange) {
                this.hasWarnedPendingExternalChange = true;
                void vscode.window.showWarningMessage(
                    getWebviewStrings().localChangesBlockedReload
                );
            }

            return;
        }

        if (this.autoRefreshTimer) {
            clearTimeout(this.autoRefreshTimer);
        }

        this.autoRefreshTimer = setTimeout(() => {
            this.autoRefreshTimer = undefined;
            void this.enqueueReload().catch((error) => {
                void this.handleError(error);
            });
        }, 250);
    }

    private async enqueueReload({
        silent = false,
        clearPendingEdits = false,
    }: { silent?: boolean; clearPendingEdits?: boolean } = {}): Promise<void> {
        if (this.isReloading) {
            this.hasQueuedReload = true;
            return;
        }

        this.isReloading = true;
        let reloadError: unknown;

        try {
            await this.reloadModel({ silent, clearPendingEdits });
        } catch (error) {
            reloadError = error;
        } finally {
            this.isReloading = false;

            if (this.hasQueuedReload) {
                this.hasQueuedReload = false;
                await this.enqueueReload();
            }
        }

        if (reloadError) {
            throw reloadError;
        }
    }

    private async handleError(error: unknown): Promise<void> {
        const errorMessage = toErrorMessage(error);
        console.error(error);
        await vscode.window.showErrorMessage(errorMessage);
        if (this.isWebviewReady) {
            await this.panel.webview.postMessage({
                type: "error",
                message: errorMessage,
            });
        }
    }

    private async refreshForDisplayLanguageChange(): Promise<void> {
        if (this.document.hasPendingEdits()) {
            void vscode.window.showWarningMessage(
                getWebviewStrings().displayLanguageRefreshBlocked
            );
            return;
        }

        this.isWebviewReady = false;
        this.hasPendingRender = Boolean(this.workbook);
        this.panel.webview.html = this.getHtml();
        await this.enqueueReload();
    }

    private getHtml(): string {
        const webview = this.panel.webview;
        const nonce = getNonce();
        const webviewStrings = getWebviewStrings();
        const strings = JSON.stringify(webviewStrings).replace(/</g, "\\u003c");
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "editor-panel.js")
        );
        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "panel.css")
        );
        const codiconStyleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "codicons", "codicon.css")
        );

        return `<!DOCTYPE html>
<html lang="${getHtmlLanguageTag()}">
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource};" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${codiconStyleUri}" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Editor</title>
</head>
<body>
	<div id="app" class="loading-shell">
		<div class="loading-shell__message">${webviewStrings.loading}</div>
	</div>
	<script nonce="${nonce}">window.__XLSX_EDITOR_STRINGS__ = ${strings};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
    }

    private async confirmDiscardPendingEdits(): Promise<boolean> {
        if (!this.document.hasPendingEdits()) {
            return true;
        }

        const strings = getWebviewStrings();
        const choice = await vscode.window.showWarningMessage(
            strings.confirmReloadDiscard,
            { modal: true },
            strings.discardChangesAndReload,
            strings.keepEditing
        );

        return choice === strings.discardChangesAndReload;
    }

    private async handleMessage(message: WebviewMessage): Promise<void> {
        try {
            switch (message.type) {
                case "ready":
                    this.isWebviewReady = true;
                    if (this.hasPendingRender) {
                        await this.render();
                    }
                    return;
                case "setSheet":
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }
                    this.state = setActiveEditorSheet(
                        this.getWorkingWorkbook()!,
                        this.state,
                        message.sheetKey,
                        this.getSheetEntries()
                    );
                    await this.render();
                    return;
                case "addSheet":
                    await this.addPendingSheet();
                    return;
                case "deleteSheet":
                    await this.deletePendingSheet(message.sheetKey);
                    return;
                case "renameSheet":
                    await this.renamePendingSheet(message.sheetKey);
                    return;
                case "setViewportStartRow":
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }
                    this.state = setEditorViewportStartRow(
                        this.getWorkingWorkbook()!,
                        this.state,
                        message.rowNumber,
                        this.getSheetEntries()
                    );
                    await this.render(undefined, { silent: true });
                    return;
                case "search": {
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }

                    let match: { sheetKey: string; rowNumber: number; columnNumber: number } | null;
                    try {
                        match = this.findSearchMatch(
                            message.query,
                            message.direction,
                            message.options
                        );
                    } catch (error) {
                        await vscode.window.showInformationMessage(
                            `${getWebviewStrings().invalidSearchPattern} ${toErrorMessage(error)}`
                        );
                        return;
                    }

                    if (!match) {
                        await vscode.window.showInformationMessage(
                            getWebviewStrings().noSearchMatches
                        );
                        return;
                    }

                    this.revealCell(match.sheetKey, match.rowNumber, match.columnNumber);
                    await this.render();
                    return;
                }
                case "gotoCell":
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }

                    if (!this.gotoCellReference(message.reference)) {
                        await vscode.window.showInformationMessage(
                            getWebviewStrings().invalidCellReference
                        );
                        return;
                    }

                    await this.render();
                    return;
                case "selectCell":
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }
                    this.state = setSelectedEditorCell(
                        this.getWorkingWorkbook()!,
                        this.state,
                        message.rowNumber,
                        message.columnNumber,
                        this.getSheetEntries()
                    );
                    return;
                case "pendingEditStateChanged":
                    if (!message.hasPendingEdits) {
                        this.hasWarnedPendingExternalChange = false;
                    }
                    return;
                case "setPendingEdits": {
                    if (!this.workbook) {
                        return;
                    }

                    const cellEdits: CellEdit[] = message.edits.flatMap((edit) => {
                        const sheet = this.getSheetEntries().find(
                            (candidate) => candidate.key === edit.sheetKey
                        )?.sheet;

                        return sheet
                            ? [
                                  {
                                      sheetName: sheet.name,
                                      rowNumber: edit.rowNumber,
                                      columnNumber: edit.columnNumber,
                                      value: edit.value,
                                  },
                              ]
                            : [];
                    });

                    this.pendingCellEdits = cellEdits;
                    await this.controller.onPendingStateChanged({
                        cellEdits: this.pendingCellEdits,
                        sheetEdits: this.pendingSheetEdits,
                    });
                    if (cellEdits.length === 0 && !this.document.hasPendingEdits()) {
                        this.hasWarnedPendingExternalChange = false;
                    }
                    return;
                }
                case "requestSave":
                    this.suppressAutoRefreshUntil = Date.now() + 2000;
                    await this.controller.onRequestSave();
                    return;
                case "undoSheetEdit":
                    await this.undoStructuralEdit();
                    return;
                case "redoSheetEdit":
                    await this.redoStructuralEdit();
                    return;
                case "toggleViewLock":
                    await this.toggleViewLock(message.columnCount, message.rowCount);
                    return;
                case "reload":
                    if (this.document.hasPendingEdits()) {
                        await this.controller.onRequestRevert();
                        return;
                    }

                    await this.enqueueReload({ clearPendingEdits: true });
                    return;
            }
        } catch (error) {
            await this.handleError(error);
        }
    }

    private revealCell(sheetKey: string, rowNumber: number, columnNumber: number): void {
        const workbook = this.getWorkingWorkbook();
        if (!workbook) {
            return;
        }

        this.state = setSelectedEditorCell(
            workbook,
            setActiveEditorSheet(workbook, this.state, sheetKey, this.getSheetEntries()),
            rowNumber,
            columnNumber,
            this.getSheetEntries()
        );
    }

    private getWorkingWorkbook(): WorkbookSnapshot | null {
        if (!this.workbook) {
            return null;
        }

        return {
            ...this.workbook,
            sheets: this.workingSheetEntries.map((entry) => entry.sheet),
        };
    }

    private getSheetEntries(): Array<{
        key: string;
        index: number;
        sheet: WorkbookSnapshot["sheets"][number];
    }> {
        return this.workingSheetEntries;
    }

    private createPendingState(): WorkbookEditState {
        return {
            cellEdits: this.pendingCellEdits.map(cloneCellEdit),
            sheetEdits: this.pendingSheetEdits.map(cloneSheetEdit),
            viewEdits: this.pendingViewEdits.map(cloneViewEdit),
        };
    }

    private captureStructuralSnapshot(): StructuralSnapshot {
        return {
            state: cloneEditorState(this.state),
            sheetEntries: this.workingSheetEntries.map(cloneSheetEntry),
            pendingCellEdits: this.pendingCellEdits.map(cloneCellEdit),
            pendingSheetEdits: this.pendingSheetEdits.map(cloneSheetEdit),
            pendingViewEdits: this.pendingViewEdits.map(cloneViewEdit),
        };
    }

    private restoreStructuralSnapshot(snapshot: StructuralSnapshot): void {
        this.state = cloneEditorState(snapshot.state);
        this.workingSheetEntries = reindexWorkingSheetEntries(
            snapshot.sheetEntries.map(cloneSheetEntry)
        );
        this.pendingCellEdits = snapshot.pendingCellEdits.map(cloneCellEdit);
        this.pendingSheetEdits = snapshot.pendingSheetEdits.map(cloneSheetEdit);
        this.pendingViewEdits = snapshot.pendingViewEdits.map(cloneViewEdit);
    }

    private async commitStructuralMutation(
        mutate: () => void,
        { resetPendingHistory = true }: { resetPendingHistory?: boolean } = {}
    ): Promise<void> {
        const before = this.captureStructuralSnapshot();
        mutate();
        this.workingSheetEntries = reindexWorkingSheetEntries(this.workingSheetEntries);
        const after = this.captureStructuralSnapshot();
        this.sheetUndoStack.push({ before, after, resetPendingHistory });
        this.sheetRedoStack.length = 0;
        await this.syncPendingState();
        await this.render(undefined, {
            useModelSelection: true,
            replacePendingEdits: this.pendingCellEdits,
            resetPendingHistory,
        });
    }

    private getOriginalSheetSnapshot(
        sheetKey: string
    ): WorkbookSnapshot["sheets"][number] | undefined {
        return this.workbook?.sheets.find((_sheet, index) => `sheet:${index}` === sheetKey);
    }

    private updatePendingViewLock(
        sheetKey: string,
        sheetName: string,
        freezePane: WorkbookSnapshot["sheets"][number]["freezePane"]
    ): void {
        const originalFreezePane = this.getOriginalSheetSnapshot(sheetKey)?.freezePane ?? null;
        const nextFreezePane = freezePane ?? null;

        if (areFreezePaneCountsEqual(originalFreezePane, nextFreezePane)) {
            this.pendingViewEdits = this.pendingViewEdits.filter(
                (edit) => edit.sheetKey !== sheetKey
            );
            return;
        }

        const nextEdit: SheetViewEdit = {
            sheetKey,
            sheetName,
            freezePane: nextFreezePane
                ? {
                      columnCount: nextFreezePane.columnCount,
                      rowCount: nextFreezePane.rowCount,
                  }
                : null,
        };

        const existingIndex = this.pendingViewEdits.findIndex((edit) => edit.sheetKey === sheetKey);
        if (existingIndex >= 0) {
            this.pendingViewEdits = this.pendingViewEdits.map((edit, index) =>
                index === existingIndex ? nextEdit : edit
            );
            return;
        }

        this.pendingViewEdits = [...this.pendingViewEdits, nextEdit];
    }

    private findSheetEntry(sheetKey: string): WorkingSheetEntry | undefined {
        return this.workingSheetEntries.find((entry) => entry.key === sheetKey);
    }

    private getActiveSheetEntry(): WorkingSheetEntry | undefined {
        return this.workingSheetEntries.find((entry) => entry.key === this.state.activeSheetKey);
    }

    private getNextWorkingSheetKey(): string {
        const key = `sheet:new:${this.nextNewSheetId}`;
        this.nextNewSheetId += 1;
        return key;
    }

    private validateSheetName(value: string, currentSheetKey?: string): string | undefined {
        const trimmed = value.trim();
        const strings = getWebviewStrings();

        if (!trimmed) {
            return strings.sheetNameEmpty;
        }

        if (trimmed.length > 31) {
            return strings.sheetNameTooLong;
        }

        if (/[\\/:?*\[\]]/.test(trimmed)) {
            return strings.sheetNameInvalidChars;
        }

        if (
            this.workingSheetEntries.some(
                (entry) =>
                    entry.key !== currentSheetKey &&
                    entry.sheet.name.toLocaleLowerCase() === trimmed.toLocaleLowerCase()
            )
        ) {
            return strings.sheetNameDuplicate;
        }

        return undefined;
    }

    private updatePendingRename(sheetKey: string, previousName: string, nextName: string): void {
        const addSheetEdit = this.pendingSheetEdits.find(
            (edit) => edit.type === "addSheet" && edit.sheetKey === sheetKey
        );
        if (addSheetEdit?.type === "addSheet") {
            addSheetEdit.sheetName = nextName;
            this.pendingSheetEdits = this.pendingSheetEdits.filter(
                (edit) => !(edit.type === "renameSheet" && edit.sheetKey === sheetKey)
            );
            return;
        }

        const renameEdit = this.pendingSheetEdits.find(
            (edit) => edit.type === "renameSheet" && edit.sheetKey === sheetKey
        );
        if (renameEdit?.type === "renameSheet") {
            renameEdit.nextSheetName = nextName;
            return;
        }

        this.pendingSheetEdits = [
            ...this.pendingSheetEdits,
            {
                type: "renameSheet",
                sheetKey,
                sheetName: previousName,
                nextSheetName: nextName,
            },
        ];
    }

    private updatePendingViewRename(sheetKey: string, nextName: string): void {
        this.pendingViewEdits = this.pendingViewEdits.map((edit) =>
            edit.sheetKey === sheetKey ? { ...edit, sheetName: nextName } : edit
        );
    }

    private async toggleViewLock(columnCount: number, rowCount: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (!workbook || workbook.isReadonly || !activeEntry) {
            return;
        }

        const isLocked = hasLockedView(activeEntry.sheet.freezePane);

        const nextColumnCount = Math.max(
            0,
            Math.min(columnCount, Math.max(activeEntry.sheet.columnCount - 1, 0))
        );
        const nextRowCount = Math.max(
            0,
            Math.min(rowCount, Math.max(activeEntry.sheet.rowCount - 1, 0))
        );

        if (!isLocked && nextColumnCount === 0 && nextRowCount === 0) {
            return;
        }

        await this.commitStructuralMutation(
            () => {
                const entry = this.getActiveSheetEntry();
                if (!entry) {
                    return;
                }

                const nextFreezePane = isLocked
                    ? null
                    : createFreezePaneSnapshot(nextColumnCount, nextRowCount);
                entry.sheet = {
                    ...entry.sheet,
                    freezePane: nextFreezePane,
                };
                this.updatePendingViewLock(entry.key, entry.sheet.name, nextFreezePane);
            },
            { resetPendingHistory: false }
        );
    }

    private async undoStructuralEdit(): Promise<void> {
        const entry = this.sheetUndoStack.pop();
        if (!entry) {
            return;
        }

        this.restoreStructuralSnapshot(entry.before);
        this.sheetRedoStack.push(entry);
        await this.syncPendingState();
        await this.render(undefined, {
            useModelSelection: true,
            replacePendingEdits: this.pendingCellEdits,
            resetPendingHistory: entry.resetPendingHistory,
        });
    }

    private async redoStructuralEdit(): Promise<void> {
        const entry = this.sheetRedoStack.pop();
        if (!entry) {
            return;
        }

        this.restoreStructuralSnapshot(entry.after);
        this.sheetUndoStack.push(entry);
        await this.syncPendingState();
        await this.render(undefined, {
            useModelSelection: true,
            replacePendingEdits: this.pendingCellEdits,
            resetPendingHistory: entry.resetPendingHistory,
        });
    }

    private async syncPendingState(): Promise<void> {
        await this.controller.onPendingStateChanged(this.createPendingState());

        if (
            this.pendingCellEdits.length === 0 &&
            this.pendingSheetEdits.length === 0 &&
            !this.document.hasPendingEdits()
        ) {
            this.hasWarnedPendingExternalChange = false;
        }
    }

    private getNewSheetName(): string {
        const baseName = isChineseDisplayLanguage() ? "工作表" : "Sheet";
        const existingNames = new Set(this.workingSheetEntries.map((entry) => entry.sheet.name));

        let suffix = 1;
        while (existingNames.has(`${baseName}${suffix}`)) {
            suffix += 1;
        }

        return `${baseName}${suffix}`;
    }

    private getInsertSheetIndex(): number {
        if (this.workingSheetEntries.length === 0) {
            return 0;
        }

        const activeSheetIndex = this.workingSheetEntries.findIndex(
            (sheet) => sheet.key === this.state.activeSheetKey
        );

        return activeSheetIndex >= 0 ? activeSheetIndex + 1 : this.workingSheetEntries.length;
    }

    private async addPendingSheet(): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        if (!workbook || workbook.isReadonly) {
            return;
        }

        const sheetName = this.getNewSheetName();
        const targetIndex = this.getInsertSheetIndex();
        await this.commitStructuralMutation(() => {
            const sheetKey = this.getNextWorkingSheetKey();
            const newEntry: WorkingSheetEntry = {
                key: sheetKey,
                index: targetIndex,
                sheet: {
                    name: sheetName,
                    rowCount: DEFAULT_PAGE_SIZE,
                    columnCount: 26,
                    mergedRanges: [],
                    cells: {},
                    signature: `pending:${sheetKey}:${sheetName}`,
                },
            };
            this.workingSheetEntries = [
                ...this.workingSheetEntries.slice(0, targetIndex),
                newEntry,
                ...this.workingSheetEntries.slice(targetIndex),
            ];
            this.pendingSheetEdits = [
                ...this.pendingSheetEdits,
                {
                    type: "addSheet",
                    sheetKey,
                    sheetName,
                    targetIndex,
                },
            ];
            this.state = setActiveEditorSheet(
                this.getWorkingWorkbook()!,
                this.state,
                sheetKey,
                this.getSheetEntries()
            );
        });
    }

    private async deletePendingSheet(sheetKey: string): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        if (!workbook || workbook.isReadonly || this.workingSheetEntries.length <= 1) {
            return;
        }

        const targetIndex = this.workingSheetEntries.findIndex((sheet) => sheet.key === sheetKey);
        if (targetIndex < 0) {
            return;
        }

        const deletedEntry = this.workingSheetEntries[targetIndex];
        if (!deletedEntry) {
            return;
        }

        await this.commitStructuralMutation(() => {
            this.workingSheetEntries = this.workingSheetEntries.filter(
                (entry) => entry.key !== sheetKey
            );
            this.pendingCellEdits = this.pendingCellEdits.filter(
                (edit) => edit.sheetName !== deletedEntry.sheet.name
            );
            this.pendingViewEdits = this.pendingViewEdits.filter(
                (edit) => edit.sheetKey !== sheetKey
            );

            const pendingAddIndex = this.pendingSheetEdits.findIndex(
                (edit) => edit.type === "addSheet" && edit.sheetKey === sheetKey
            );
            if (pendingAddIndex >= 0) {
                this.pendingSheetEdits = this.pendingSheetEdits.filter(
                    (edit) => edit.sheetKey !== sheetKey
                );
            } else {
                this.pendingSheetEdits = [
                    ...this.pendingSheetEdits,
                    {
                        type: "deleteSheet",
                        sheetKey,
                        sheetName: deletedEntry.sheet.name,
                        targetIndex,
                    },
                ];
            }

            const fallbackSheet =
                this.workingSheetEntries[Math.max(0, targetIndex - 1)] ??
                this.workingSheetEntries[0];
            this.state = fallbackSheet
                ? setActiveEditorSheet(
                      this.getWorkingWorkbook()!,
                      this.state,
                      fallbackSheet.key,
                      this.getSheetEntries()
                  )
                : createInitialEditorPanelState(this.getWorkingWorkbook()!, this.getSheetEntries());
        });
    }

    private async renamePendingSheet(sheetKey: string): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const targetEntry = this.findSheetEntry(sheetKey);
        if (!workbook || workbook.isReadonly || !targetEntry) {
            return;
        }

        const strings = getWebviewStrings();
        const nextName = await vscode.window.showInputBox({
            prompt: strings.renameSheetPrompt,
            title: strings.renameSheetTitle,
            value: targetEntry.sheet.name,
            validateInput: (value) => this.validateSheetName(value, sheetKey),
        });

        const trimmedName = nextName?.trim();
        if (!trimmedName || trimmedName === targetEntry.sheet.name) {
            return;
        }

        await this.commitStructuralMutation(() => {
            const entry = this.findSheetEntry(sheetKey);
            if (!entry) {
                return;
            }

            const previousName = entry.sheet.name;
            entry.sheet = {
                ...entry.sheet,
                name: trimmedName,
                signature: `pending:${sheetKey}:${trimmedName}`,
            };
            this.pendingCellEdits = this.pendingCellEdits.map((edit) =>
                edit.sheetName === previousName ? { ...edit, sheetName: trimmedName } : edit
            );
            this.updatePendingRename(sheetKey, previousName, trimmedName);
            this.updatePendingViewRename(sheetKey, trimmedName);
            this.state = setActiveEditorSheet(
                this.getWorkingWorkbook()!,
                {
                    ...this.state,
                    activeSheetKey: sheetKey,
                },
                sheetKey,
                this.getSheetEntries()
            );
        });
    }

    private findSearchMatch(
        query: string,
        direction: "next" | "prev",
        options: SearchOptions
    ): { sheetKey: string; rowNumber: number; columnNumber: number } | null {
        const normalizedQuery = query.trim();
        const workbook = this.getWorkingWorkbook();
        if (!workbook || !normalizedQuery) {
            return null;
        }

        const pattern = createSearchPattern(normalizedQuery, options);
        const sheetEntries = this.getSheetEntries();
        const activeSheetEntry =
            sheetEntries.find((entry) => entry.key === this.state.activeSheetKey) ??
            sheetEntries[0];
        if (!activeSheetEntry) {
            return null;
        }

        const matches = Object.values(activeSheetEntry.sheet.cells)
            .filter((cell) => {
                const value = cell.displayValue;
                const formula = cell.formula ?? "";
                return pattern.test(value) || pattern.test(formula);
            })
            .map((cell) => ({
                sheetKey: activeSheetEntry.key,
                rowNumber: cell.rowNumber,
                columnNumber: cell.columnNumber,
            }));

        if (matches.length === 0) {
            return null;
        }

        matches.sort((left, right) => {
            if (left.rowNumber !== right.rowNumber) {
                return left.rowNumber - right.rowNumber;
            }

            return left.columnNumber - right.columnNumber;
        });

        const selectionOnCurrentPage =
            this.state.selectedCell &&
            this.state.selectedCell.rowNumber >= 1 &&
            this.state.selectedCell.rowNumber <= activeSheetEntry.sheet.rowCount;
        const anchor = {
            rowNumber: selectionOnCurrentPage ? this.state.selectedCell!.rowNumber : 1,
            columnNumber: selectionOnCurrentPage ? this.state.selectedCell!.columnNumber : 0,
        };
        const compare = (
            candidate: { rowNumber: number; columnNumber: number },
            current: { rowNumber: number; columnNumber: number }
        ): number => {
            if (candidate.rowNumber !== current.rowNumber) {
                return candidate.rowNumber - current.rowNumber;
            }

            return candidate.columnNumber - current.columnNumber;
        };

        if (direction === "prev") {
            for (let index = matches.length - 1; index >= 0; index -= 1) {
                if (compare(matches[index], anchor) < 0) {
                    return matches[index];
                }
            }

            return matches[matches.length - 1];
        }

        return matches.find((match) => compare(match, anchor) > 0) ?? matches[0];
    }

    private gotoCellReference(reference: string): boolean {
        if (!this.workbook) {
            return false;
        }

        const trimmedReference = reference.trim();
        if (!trimmedReference) {
            return false;
        }

        const separatorIndex = trimmedReference.lastIndexOf("!");
        const sheetName =
            separatorIndex > 0 ? trimmedReference.slice(0, separatorIndex).trim() : null;
        const address =
            separatorIndex > 0
                ? trimmedReference.slice(separatorIndex + 1).trim()
                : trimmedReference;
        const addressMatch = /^([A-Za-z]+)(\d+)$/.exec(address);
        if (!addressMatch) {
            return false;
        }

        const columnNumber = getColumnNumber(addressMatch[1]);
        const rowNumber = Number(addressMatch[2]);
        if (!columnNumber || rowNumber < 1) {
            return false;
        }

        const sheetEntries = this.getSheetEntries();
        const targetSheet = sheetName
            ? sheetEntries.find(
                  (entry) =>
                      entry.sheet.name === sheetName ||
                      entry.sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()
              )
            : (sheetEntries.find((entry) => entry.key === this.state.activeSheetKey) ??
              sheetEntries[0]);

        if (
            !targetSheet ||
            rowNumber > targetSheet.sheet.rowCount ||
            columnNumber > targetSheet.sheet.columnCount
        ) {
            return false;
        }

        this.revealCell(targetSheet.key, rowNumber, columnNumber);
        return true;
    }

    private async reloadModel({
        silent = false,
        clearPendingEdits = false,
    }: { silent?: boolean; clearPendingEdits?: boolean } = {}): Promise<void> {
        const webviewStrings = getWebviewStrings();

        if (!silent) {
            this.panel.title = webviewStrings.loading;

            if (this.isWebviewReady) {
                await this.panel.webview.postMessage({
                    type: "loading",
                    message: webviewStrings.loading,
                });
            }
        }

        this.workbook = await loadWorkbookSnapshot(this.document.getReadUri());
        this.workbook.filePath = this.workbookUri.fsPath;
        this.workbook.fileName = getWorkbookResourceName(this.workbookUri);
        this.workingSheetEntries = createWorkingSheetEntries(this.workbook);
        this.nextNewSheetId = 1;
        if (clearPendingEdits) {
            this.pendingCellEdits = [];
            this.pendingSheetEdits = [];
            this.pendingViewEdits = [];
            this.sheetUndoStack.length = 0;
            this.sheetRedoStack.length = 0;
        }
        this.state = this.workingSheetEntries.length
            ? normalizeEditorPanelState(
                  this.getWorkingWorkbook()!,
                  this.state.activeSheetKey
                      ? this.state
                      : createInitialEditorPanelState(
                            this.getWorkingWorkbook()!,
                            this.getSheetEntries()
                        ),
                  this.getSheetEntries()
              )
            : createInitialEditorPanelState(this.getWorkingWorkbook()!, this.getSheetEntries());

        const renderModel = createEditorRenderModel(this.getWorkingWorkbook()!, this.state, {
            hasPendingEdits: clearPendingEdits ? false : this.document.hasPendingEdits(),
            sheetEntries: this.getSheetEntries(),
            canUndoStructuralEdits: this.sheetUndoStack.length > 0,
            canRedoStructuralEdits: this.sheetRedoStack.length > 0,
        });

        if (clearPendingEdits) {
            this.hasWarnedPendingExternalChange = false;
        }
        this.panel.title = renderModel.title;
        await this.render(renderModel, {
            silent,
            clearPendingEdits,
            useModelSelection: true,
        });
    }

    private async render(
        renderModel?: EditorRenderModel,
        {
            silent = false,
            clearPendingEdits = false,
            useModelSelection = true,
            replacePendingEdits,
            resetPendingHistory = false,
        }: {
            silent?: boolean;
            clearPendingEdits?: boolean;
            useModelSelection?: boolean;
            replacePendingEdits?: CellEdit[];
            resetPendingHistory?: boolean;
        } = {}
    ): Promise<void> {
        if (!this.workbook) {
            return;
        }

        const workbook = this.getWorkingWorkbook();
        if (!workbook) {
            return;
        }

        const payload =
            renderModel ??
            createEditorRenderModel(workbook, this.state, {
                hasPendingEdits: this.document.hasPendingEdits(),
                sheetEntries: this.getSheetEntries(),
                canUndoStructuralEdits: this.sheetUndoStack.length > 0,
                canRedoStructuralEdits: this.sheetRedoStack.length > 0,
            });
        this.panel.title = payload.title;

        if (!this.isWebviewReady) {
            this.hasPendingRender = true;
            return;
        }

        this.hasPendingRender = false;
        await this.panel.webview.postMessage({
            type: "render",
            payload,
            silent,
            clearPendingEdits,
            useModelSelection,
            replacePendingEdits:
                replacePendingEdits?.map((edit) => ({
                    sheetKey:
                        this.getSheetEntries().find((entry) => entry.sheet.name === edit.sheetName)
                            ?.key ?? edit.sheetName,
                    rowNumber: edit.rowNumber,
                    columnNumber: edit.columnNumber,
                    value: edit.value,
                })) ?? undefined,
            resetPendingHistory,
        });
    }
}
