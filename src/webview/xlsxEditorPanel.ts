import * as path from "node:path";
import * as vscode from "vscode";
import { DEFAULT_PAGE_SIZE } from "../constants";
import { getColumnNumber } from "../core/model/cells";
import { loadWorkbookSnapshot } from "../core/fastxlsx/loadWorkbookSnapshot";
import {
    type CellEdit,
    type SheetEdit,
    type WorkbookEditState,
} from "../core/fastxlsx/writeCellValue";
import type { EditorPanelState, EditorRenderModel, WorkbookSnapshot } from "../core/model/types";
import { getHtmlLanguageTag, isChineseDisplayLanguage } from "../displayLanguage";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    getEditorSheetKey,
    moveEditorPageCursor,
    normalizeEditorPanelState,
    setActiveEditorSheet,
    setEditorCurrentPage,
    setSelectedEditorCell,
} from "./editorRenderModel";
import { XlsxEditorDocument } from "./xlsxEditorDocument";
import { getWorkbookResourceName } from "../workbook/resourceUri";

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
    | { type: "setPage"; page: number }
    | { type: "prevPage" }
    | { type: "nextPage" }
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
    | { type: "reload" };

export interface XlsxEditorPanelController {
    onPendingStateChanged(state: WorkbookEditState): Promise<void> | void;
    onRequestSave(): Promise<void> | void;
    onRequestRevert(): Promise<void> | void;
}

interface WebviewStrings {
    loading: string;
    reload: string;
    prevPage: string;
    nextPage: string;
    size: string;
    modified: string;
    sheet: string;
    rows: string;
    noRows: string;
    page: string;
    visibleRows: string;
    readOnly: string;
    save: string;
    addSheet: string;
    deleteSheet: string;
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
            prevPage: "上一页",
            nextPage: "下一页",
            size: "大小",
            modified: "修改时间",
            sheet: "工作表",
            rows: "行",
            noRows: "无行",
            page: "页码",
            visibleRows: "可见行",
            readOnly: "只读",
            save: "保存",
            addSheet: "添加工作表",
            deleteSheet: "删除工作表",
            totalSheets: "总工作表",
            totalRows: "总行数",
            nonEmptyCells: "非空单元格",
            selectedCell: "当前单元格",
            noCellSelected: "未选择",
            mergedRanges: "合并区域",
            pendingChanges: "待保存修改",
            noRowsAvailable: "当前页没有可显示的行。",
            readOnlyBadge: "只读模式",
            localChangesBlockedReload:
                "工作簿文件已在磁盘上变化。请先保存或放弃当前未保存修改，再刷新。",
            confirmReloadDiscard: "刷新会丢弃当前未保存修改，是否继续？",
            discardChangesAndReload: "放弃修改并刷新",
            keepEditing: "继续编辑",
            displayLanguageRefreshBlocked: "当前有未保存修改，语言变更将在保存或手动刷新后生效。",
            noSearchMatches: "当前页没有找到匹配的单元格。",
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
        prevPage: "Prev Page",
        nextPage: "Next Page",
        size: "Size",
        modified: "Modified",
        sheet: "Sheet",
        rows: "Rows",
        noRows: "No rows",
        page: "Page",
        visibleRows: "Visible rows",
        readOnly: "Read-only",
        save: "Save",
        addSheet: "Add Sheet",
        deleteSheet: "Delete Sheet",
        totalSheets: "Sheets",
        totalRows: "Rows",
        nonEmptyCells: "Non-empty cells",
        selectedCell: "Selected cell",
        noCellSelected: "None",
        mergedRanges: "Merged ranges",
        pendingChanges: "Pending changes",
        noRowsAvailable: "No rows available on this page.",
        readOnlyBadge: "Read-only",
        localChangesBlockedReload:
            "The workbook changed on disk. Save or discard your pending edits before reloading.",
        confirmReloadDiscard: "Reloading will discard your pending edits. Continue?",
        discardChangesAndReload: "Discard Changes and Reload",
        keepEditing: "Keep Editing",
        displayLanguageRefreshBlocked:
            "Pending edits are open. Display language changes will apply after you save or reload the editor.",
        noSearchMatches: "No matching cells were found on this page.",
        invalidCellReference:
            "Unable to locate that cell. Use A1 or Sheet1!B2 and stay within the workbook range.",
        invalidSearchPattern: "The search pattern is invalid.",
        searchRegex: "Use Regular Expression",
        searchMatchCase: "Match Case",
        searchWholeWord: "Match Whole Word",
    };
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
        currentPage: 1,
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
            vscode.Uri.joinPath(this.extensionUri, "media", "editorPanel.js")
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
                    if (!this.workbook) {
                        return;
                    }
                    this.state = setActiveEditorSheet(this.workbook, this.state, message.sheetKey);
                    await this.render();
                    return;
                case "addSheet":
                    await this.addPendingSheet();
                    return;
                case "deleteSheet":
                    await this.deletePendingSheet(message.sheetKey);
                    return;
                case "setPage":
                    if (!this.workbook) {
                        return;
                    }
                    this.state = setEditorCurrentPage(this.workbook, this.state, message.page);
                    await this.render();
                    return;
                case "prevPage":
                    if (!this.workbook) {
                        return;
                    }
                    this.state = moveEditorPageCursor(this.workbook, this.state, -1);
                    await this.render();
                    return;
                case "nextPage":
                    if (!this.workbook) {
                        return;
                    }
                    this.state = moveEditorPageCursor(this.workbook, this.state, 1);
                    await this.render();
                    return;
                case "search": {
                    if (!this.workbook) {
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
                    if (!this.workbook) {
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
                    if (!this.workbook) {
                        return;
                    }
                    this.state = setSelectedEditorCell(
                        this.workbook,
                        this.state,
                        message.rowNumber,
                        message.columnNumber
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
                        const sheet = this.workbook!.sheets.find(
                            (candidate) => getEditorSheetKey(candidate) === edit.sheetKey
                        );

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
        if (!this.workbook) {
            return;
        }

        this.state = setSelectedEditorCell(
            this.workbook,
            setActiveEditorSheet(this.workbook, this.state, sheetKey),
            rowNumber,
            columnNumber
        );
    }

    private getSheetEntries(): Array<{
        key: string;
        index: number;
        sheet: WorkbookSnapshot["sheets"][number];
    }> {
        if (!this.workbook) {
            return [];
        }

        return this.workbook.sheets.map((sheet, index) => ({
            key: getEditorSheetKey(sheet),
            index,
            sheet,
        }));
    }

    private createPendingState(): WorkbookEditState {
        return {
            cellEdits: this.pendingCellEdits,
            sheetEdits: this.pendingSheetEdits,
        };
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
        const existingNames = new Set(this.workbook?.sheets.map((sheet) => sheet.name) ?? []);

        let suffix = 1;
        while (existingNames.has(`${baseName}${suffix}`)) {
            suffix += 1;
        }

        return `${baseName}${suffix}`;
    }

    private getInsertSheetIndex(): number {
        if (!this.workbook || this.workbook.sheets.length === 0) {
            return 0;
        }

        const activeSheetIndex = this.workbook.sheets.findIndex(
            (sheet) => getEditorSheetKey(sheet) === this.state.activeSheetKey
        );

        return activeSheetIndex >= 0 ? activeSheetIndex + 1 : this.workbook.sheets.length;
    }

    private async addPendingSheet(): Promise<void> {
        if (!this.workbook || this.workbook.isReadonly) {
            return;
        }

        const sheetName = this.getNewSheetName();
        const targetIndex = this.getInsertSheetIndex();
        this.workbook = {
            ...this.workbook,
            sheets: [
                ...this.workbook.sheets.slice(0, targetIndex),
                {
                    name: sheetName,
                    rowCount: DEFAULT_PAGE_SIZE,
                    columnCount: 26,
                    mergedRanges: [],
                    cells: {},
                    signature: `pending:${sheetName}`,
                },
                ...this.workbook.sheets.slice(targetIndex),
            ],
        };

        this.pendingSheetEdits = [
            ...this.pendingSheetEdits,
            {
                type: "addSheet",
                sheetName,
                targetIndex,
            },
        ];
        this.state = setActiveEditorSheet(this.workbook, this.state, sheetName);
        await this.syncPendingState();
        await this.render(undefined, {
            useModelSelection: true,
            replacePendingEdits: this.pendingCellEdits,
            resetPendingHistory: true,
        });
    }

    private async deletePendingSheet(sheetKey: string): Promise<void> {
        if (!this.workbook || this.workbook.isReadonly || this.workbook.sheets.length <= 1) {
            return;
        }

        const targetIndex = this.workbook.sheets.findIndex(
            (sheet) => getEditorSheetKey(sheet) === sheetKey
        );
        if (targetIndex < 0) {
            return;
        }

        const [deletedSheet] = this.workbook.sheets.slice(targetIndex, targetIndex + 1);
        if (!deletedSheet) {
            return;
        }

        this.workbook = {
            ...this.workbook,
            sheets: this.workbook.sheets.filter((sheet) => sheet.name !== deletedSheet.name),
        };
        this.pendingCellEdits = this.pendingCellEdits.filter(
            (edit) => edit.sheetName !== deletedSheet.name
        );

        const pendingAddIndex = this.pendingSheetEdits.findIndex(
            (edit) => edit.type === "addSheet" && edit.sheetName === deletedSheet.name
        );
        if (pendingAddIndex >= 0) {
            this.pendingSheetEdits = this.pendingSheetEdits.filter(
                (_edit, index) => index !== pendingAddIndex
            );
        } else {
            this.pendingSheetEdits = [
                ...this.pendingSheetEdits,
                {
                    type: "deleteSheet",
                    sheetName: deletedSheet.name,
                },
            ];
        }

        const fallbackSheet =
            this.workbook.sheets[Math.max(0, targetIndex - 1)] ?? this.workbook.sheets[0];
        this.state = fallbackSheet
            ? setActiveEditorSheet(this.workbook, this.state, fallbackSheet.name)
            : createInitialEditorPanelState(this.workbook);
        await this.syncPendingState();
        await this.render(undefined, {
            useModelSelection: true,
            replacePendingEdits: this.pendingCellEdits,
            resetPendingHistory: true,
        });
    }

    private findSearchMatch(
        query: string,
        direction: "next" | "prev",
        options: SearchOptions
    ): { sheetKey: string; rowNumber: number; columnNumber: number } | null {
        const normalizedQuery = query.trim();
        if (!this.workbook || !normalizedQuery) {
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

        const pageStartRow = (this.state.currentPage - 1) * DEFAULT_PAGE_SIZE + 1;
        const pageEndRow = Math.min(
            activeSheetEntry.sheet.rowCount,
            this.state.currentPage * DEFAULT_PAGE_SIZE
        );
        const matches = Object.values(activeSheetEntry.sheet.cells)
            .filter((cell) => {
                if (cell.rowNumber < pageStartRow || cell.rowNumber > pageEndRow) {
                    return false;
                }

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
            this.state.selectedCell.rowNumber >= pageStartRow &&
            this.state.selectedCell.rowNumber <= pageEndRow;
        const anchor = {
            rowNumber: selectionOnCurrentPage ? this.state.selectedCell!.rowNumber : pageStartRow,
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
        if (clearPendingEdits) {
            this.pendingCellEdits = [];
            this.pendingSheetEdits = [];
        }
        this.state = this.workbook.sheets.length
            ? normalizeEditorPanelState(
                  this.workbook,
                  this.state.activeSheetKey
                      ? this.state
                      : createInitialEditorPanelState(this.workbook)
              )
            : createInitialEditorPanelState(this.workbook);

        const renderModel = createEditorRenderModel(this.workbook, this.state, {
            hasPendingEdits: clearPendingEdits ? false : this.document.hasPendingEdits(),
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

        const payload =
            renderModel ??
            createEditorRenderModel(this.workbook, this.state, {
                hasPendingEdits: this.document.hasPendingEdits(),
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
                    sheetKey: edit.sheetName,
                    rowNumber: edit.rowNumber,
                    columnNumber: edit.columnNumber,
                    value: edit.value,
                })) ?? undefined,
            resetPendingHistory,
        });
    }
}
