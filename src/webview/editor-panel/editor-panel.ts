import * as path from "node:path";
import * as vscode from "vscode";
import { DEFAULT_PAGE_SIZE } from "../../constants";
import { loadWorkbookSnapshot } from "../../core/fastxlsx/load-workbook-snapshot";
import {
    type CellEdit,
    type SheetEdit,
    type SheetViewEdit,
} from "../../core/fastxlsx/write-cell-value";
import type { EditorPanelState, EditorRenderModel, WorkbookSnapshot } from "../../core/model/types";
import { getHtmlLanguageTag } from "../../display-language";
import { toErrorMessage } from "../../error-message";
import { getRuntimeMessages } from "../../i18n";
import { rememberRecentWorkbookResourceUri } from "../../scm/recent-workbook-resource-context";
import { getWorkbookResourceName } from "../../workbook/resource-uri";
import { createWebviewNonce, escapeWatcherGlobSegment } from "../webview-utils";
import {
    getInsertEditorSheetIndex,
    getNewEditorSheetName,
    resolveEditorSearchResult,
    resolveEditorCellReference,
    validateEditorSheetName,
} from "./editor-panel-logic";
import {
    applyGridSheetEditToSheet,
    areColumnWidthsEquivalent,
    areFreezePaneCountsEqual,
    captureStructuralSnapshot,
    cloneColumnWidths,
    createCommittedWorkbookState,
    createFreezePaneSnapshot,
    createPendingWorkbookEditState,
    isGridSheetEdit,
    restorePendingWorkbookState,
    setSheetColumnWidthSnapshot,
    createWorkingSheetEntries,
    createWorkingWorkbook,
    mapPendingCellEditsToWebview,
    reindexWorkingSheetEntries,
    restoreStructuralSnapshot,
    shiftPendingCellEditsForGridSheetEdit,
} from "./editor-panel-state";
import type {
    EditorPanelStrings,
    EditorSearchResultMessage,
    EditorWebviewMessage,
    StructuralHistoryEntry,
    WorkingSheetEntry,
    XlsxEditorPanelController,
} from "./editor-panel-types";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    normalizeEditorPanelState,
    setActiveEditorSheet,
    setSelectedEditorCell,
} from "./editor-render-model";
import { hasLockedView } from "../view-lock";
import { XlsxEditorDocument } from "./xlsx-editor-document";

function getWebviewStrings(): EditorPanelStrings {
    return getRuntimeMessages().editorPanel;
}

function normalizeWorkbookColumnWidth(columnWidth: number): number {
    return Math.round(columnWidth * 256) / 256;
}

export class XlsxEditorPanel {
    private static readonly panels = new Map<number, XlsxEditorPanel>();
    private static readonly confirmedSaveBypassDocuments = new WeakSet<XlsxEditorDocument>();
    private static readonly pendingSaveConfirmations = new WeakMap<
        XlsxEditorDocument,
        Promise<boolean>
    >();
    private static nextPanelId = 1;
    private static readonly LOCAL_SAVE_AUTO_REFRESH_DELAY_MS = 3000;
    private static readonly LOCAL_SAVE_IGNORED_REFRESH_EVENTS = 4;

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
        selectedCell: null,
    };
    private isWebviewReady = false;
    private hasPendingRender = false;
    private isReloading = false;
    private hasQueuedReload = false;
    private autoRefreshTimer: ReturnType<typeof setTimeout> | undefined;
    private suppressAutoRefreshUntil = 0;
    private ignoredAutoRefreshTriggerCount = 0;
    private ignoredAutoRefreshUntil = 0;
    private isSavingDocument = false;
    private hasWarnedPendingExternalChange = false;
    private hasPendingExternalWorkbookChange = false;
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
            (message: EditorWebviewMessage) => {
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

    public static async beginDocumentSave(document: XlsxEditorDocument): Promise<void> {
        await Promise.all(
            [...XlsxEditorPanel.panels.values()]
                .filter((panel) => panel.document === document)
                .map((panel) => panel.startDocumentSave())
        );
    }

    public static async confirmDocumentSave(document: XlsxEditorDocument): Promise<boolean> {
        if (XlsxEditorPanel.confirmedSaveBypassDocuments.has(document)) {
            return true;
        }

        const pendingConfirmation = XlsxEditorPanel.pendingSaveConfirmations.get(document);
        if (pendingConfirmation) {
            return pendingConfirmation;
        }

        const matchingPanels = [...XlsxEditorPanel.panels.values()].filter(
            (panel) => panel.document === document
        );
        if (matchingPanels.length === 0) {
            return true;
        }

        const panelNeedingConfirmation =
            matchingPanels.find((panel) => panel.hasPendingExternalWorkbookChange) ??
            matchingPanels[0]!;
        const confirmationPromise = (async () => {
            try {
                return await panelNeedingConfirmation.confirmSaveIfNeeded();
            } finally {
                XlsxEditorPanel.pendingSaveConfirmations.delete(document);
            }
        })();
        XlsxEditorPanel.pendingSaveConfirmations.set(document, confirmationPromise);
        return confirmationPromise;
    }

    private static allowNextConfirmedSave(document: XlsxEditorDocument): void {
        XlsxEditorPanel.confirmedSaveBypassDocuments.add(document);
    }

    private static clearConfirmedSaveBypass(document: XlsxEditorDocument): void {
        XlsxEditorPanel.confirmedSaveBypassDocuments.delete(document);
    }

    public static async commitDocumentSave(document: XlsxEditorDocument): Promise<void> {
        XlsxEditorPanel.clearConfirmedSaveBypass(document);
        await Promise.all(
            [...XlsxEditorPanel.panels.values()]
                .filter((panel) => panel.document === document)
                .map((panel) => panel.handleDocumentSave())
        );
    }

    public static async failDocumentSave(document: XlsxEditorDocument): Promise<void> {
        XlsxEditorPanel.clearConfirmedSaveBypass(document);
        await Promise.all(
            [...XlsxEditorPanel.panels.values()]
                .filter((panel) => panel.document === document)
                .map((panel) => panel.cancelDocumentSave())
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
        this.clearAutoRefreshTimer();

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
        const scheduleRefresh = (eventType: "change" | "create" | "delete") => {
            this.scheduleAutoRefresh(eventType);
        };

        this.fileWatchers.push(watcher);
        this.fileWatchers.push(watcher.onDidChange(() => scheduleRefresh("change")));
        this.fileWatchers.push(watcher.onDidCreate(() => scheduleRefresh("create")));
        this.fileWatchers.push(watcher.onDidDelete(() => scheduleRefresh("delete")));
    }

    private scheduleAutoRefresh(trigger: "change" | "create" | "delete"): void {
        if (this.isSavingDocument) {
            this.clearAutoRefreshTimer();
            return;
        }

        if (this.shouldIgnoreAutoRefreshTrigger()) {
            this.clearAutoRefreshTimer();
            return;
        }

        if (Date.now() < this.suppressAutoRefreshUntil) {
            this.clearAutoRefreshTimer();

            return;
        }

        if (this.document.hasPendingEdits()) {
            this.hasPendingExternalWorkbookChange = true;
            if (!this.hasWarnedPendingExternalChange) {
                this.hasWarnedPendingExternalChange = true;
                void vscode.window.showWarningMessage(
                    getWebviewStrings().localChangesBlockedReload
                );
            }

            return;
        }

        this.clearAutoRefreshTimer();

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

    private async postSearchResult(message: EditorSearchResultMessage): Promise<void> {
        if (!this.isWebviewReady) {
            return;
        }

        await this.panel.webview.postMessage(message);
    }

    private getHtml(): string {
        const webview = this.panel.webview;
        const nonce = createWebviewNonce();
        const webviewStrings = getWebviewStrings();
        const strings = JSON.stringify(webviewStrings).replace(/</g, "\\u003c");
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "editor-panel.js")
        );
        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "editor-panel.css")
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

    private async handleMessage(message: EditorWebviewMessage): Promise<void> {
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
                case "insertRow":
                    await this.insertPendingRow(message.rowNumber);
                    return;
                case "deleteRow":
                    await this.deletePendingRow(message.rowNumber);
                    return;
                case "insertColumn":
                    await this.insertPendingColumn(message.columnNumber);
                    return;
                case "deleteColumn":
                    await this.deletePendingColumn(message.columnNumber);
                    return;
                case "promptColumnWidth":
                    await this.promptPendingColumnWidth(message.columnNumber);
                    return;
                case "setColumnWidth":
                    await this.setPendingColumnWidth(message.columnNumber, message.width);
                    return;
                case "search": {
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }

                    const result = resolveEditorSearchResult(
                        this.getSheetEntries(),
                        this.state,
                        message,
                        {
                            pendingEdits: this.pendingCellEdits,
                        }
                    );
                    if (result.status === "invalid-pattern") {
                        await this.postSearchResult({
                            type: "searchResult",
                            status: "invalid-pattern",
                            scope: message.scope,
                            message: getWebviewStrings().invalidSearchPattern,
                        });
                        return;
                    }

                    if (result.status === "no-match" || !result.match) {
                        await this.postSearchResult({
                            type: "searchResult",
                            status: "no-match",
                            scope: message.scope,
                            message: getWebviewStrings().noSearchMatches,
                        });
                        return;
                    }

                    this.state = setSelectedEditorCell(
                        this.getWorkingWorkbook()!,
                        this.state,
                        result.match.rowNumber,
                        result.match.columnNumber,
                        this.getSheetEntries()
                    );
                    await this.postSearchResult({
                        type: "searchResult",
                        status: "matched",
                        scope: message.scope,
                        match: result.match,
                        matchCount: result.matchCount,
                        matchIndex: result.matchIndex,
                    });
                    return;
                }
                case "gotoCell":
                    if (!this.getWorkingWorkbook()) {
                        return;
                    }

                    const targetCell = resolveEditorCellReference(
                        this.getSheetEntries(),
                        this.state.activeSheetKey,
                        message.reference
                    );
                    if (!targetCell) {
                        await vscode.window.showInformationMessage(
                            getWebviewStrings().invalidCellReference
                        );
                        return;
                    }

                    this.revealCell(
                        targetCell.sheetKey,
                        targetCell.rowNumber,
                        targetCell.columnNumber
                    );
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
                        if (this.hasPendingExternalWorkbookChange) {
                            this.hasPendingExternalWorkbookChange = false;
                            await this.enqueueReload({ clearPendingEdits: true });
                        }
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
                    await this.controller.onPendingStateChanged(this.createPendingState());
                    if (cellEdits.length === 0 && !this.document.hasPendingEdits()) {
                        this.hasWarnedPendingExternalChange = false;
                        if (this.hasPendingExternalWorkbookChange) {
                            this.hasPendingExternalWorkbookChange = false;
                            await this.enqueueReload({ clearPendingEdits: true });
                        }
                    }
                    return;
                }
                case "requestSave":
                    await this.requestDocumentSave();
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

        return createWorkingWorkbook(
            this.workbook,
            this.workingSheetEntries,
            this.pendingCellEdits
        );
    }

    private async commitSavedState(): Promise<void> {
        if (!this.workbook) {
            await this.enqueueReload({ silent: true, clearPendingEdits: true });
            return;
        }

        const activeSheetName = this.getActiveSheetEntry()?.sheet.name ?? null;
        const selectedCell = this.state.selectedCell ? { ...this.state.selectedCell } : null;
        const committedState = createCommittedWorkbookState(
            this.workbook,
            this.workingSheetEntries,
            this.pendingCellEdits
        );

        this.workbook = committedState.workbook;
        this.workingSheetEntries = committedState.sheetEntries;
        this.nextNewSheetId = 1;
        this.pendingCellEdits = [];
        this.pendingSheetEdits = [];
        this.pendingViewEdits = [];
        this.sheetUndoStack.length = 0;
        this.sheetRedoStack.length = 0;
        this.hasWarnedPendingExternalChange = false;
        this.hasPendingExternalWorkbookChange = false;

        const nextActiveEntry =
            (activeSheetName
                ? this.workingSheetEntries.find((entry) => entry.sheet.name === activeSheetName)
                : null) ??
            this.workingSheetEntries[0] ??
            null;

        this.state = normalizeEditorPanelState(
            this.getWorkingWorkbook()!,
            {
                activeSheetKey: nextActiveEntry?.key ?? null,
                selectedCell,
            },
            this.getSheetEntries()
        );

        await this.render(undefined, {
            silent: true,
            clearPendingEdits: true,
            preservePendingHistory: true,
            useModelSelection: true,
        });
    }

    private async confirmSaveIfNeeded(): Promise<boolean> {
        if (!this.hasPendingExternalWorkbookChange || !this.document.hasPendingEdits()) {
            return true;
        }

        const strings = getWebviewStrings();
        const saveAnywayAction = strings.externalChangesSaveAnyway;
        const reloadAction = strings.externalChangesReload;
        const choice = await vscode.window.showWarningMessage(
            strings.externalChangesSavePrompt,
            { modal: true },
            saveAnywayAction,
            reloadAction
        );

        if (choice === saveAnywayAction) {
            return true;
        }

        if (choice === reloadAction) {
            await this.controller.onRequestRevert();
        }

        return false;
    }

    private async requestDocumentSave(): Promise<void> {
        const shouldSave = await XlsxEditorPanel.confirmDocumentSave(this.document);
        if (!shouldSave) {
            if (this.document.hasPendingEdits()) {
                await this.render(undefined, {
                    silent: true,
                    useModelSelection: true,
                    replacePendingEdits: this.pendingCellEdits,
                });
            }
            return;
        }

        XlsxEditorPanel.allowNextConfirmedSave(this.document);
        await this.controller.onRequestSave();
    }

    private async handleDocumentSave(): Promise<void> {
        this.isSavingDocument = false;
        this.noteLocalSaveCompletion();
        if (this.hasPendingExternalWorkbookChange) {
            this.hasPendingExternalWorkbookChange = false;
            await this.enqueueReload({ silent: true, clearPendingEdits: true });
            return;
        }

        await this.commitSavedState();
    }

    private startDocumentSave(): void {
        this.isSavingDocument = true;
        this.clearAutoRefreshTimer();
        this.suppressAutoRefreshUntil = Number.POSITIVE_INFINITY;
    }

    private cancelDocumentSave(): void {
        this.isSavingDocument = false;
        this.clearAutoRefreshTimer();
        this.suppressAutoRefreshUntil = Date.now() + 1500;
    }

    private getSheetEntries(): WorkingSheetEntry[] {
        return this.workingSheetEntries;
    }

    private createPendingState() {
        return createPendingWorkbookEditState(
            this.pendingCellEdits,
            this.pendingSheetEdits,
            this.pendingViewEdits
        );
    }

    private clearAutoRefreshTimer(): void {
        if (!this.autoRefreshTimer) {
            return;
        }

        clearTimeout(this.autoRefreshTimer);
        this.autoRefreshTimer = undefined;
    }

    private noteLocalSaveCompletion(): void {
        this.suppressAutoRefreshUntil =
            Date.now() + XlsxEditorPanel.LOCAL_SAVE_AUTO_REFRESH_DELAY_MS;
        this.ignoredAutoRefreshTriggerCount = Math.max(
            this.ignoredAutoRefreshTriggerCount,
            XlsxEditorPanel.LOCAL_SAVE_IGNORED_REFRESH_EVENTS
        );
        this.ignoredAutoRefreshUntil = Math.max(
            this.ignoredAutoRefreshUntil,
            Date.now() + XlsxEditorPanel.LOCAL_SAVE_AUTO_REFRESH_DELAY_MS
        );
        this.clearAutoRefreshTimer();
    }

    private shouldIgnoreAutoRefreshTrigger(): boolean {
        if (this.ignoredAutoRefreshTriggerCount <= 0) {
            return false;
        }

        if (Date.now() > this.ignoredAutoRefreshUntil) {
            this.ignoredAutoRefreshTriggerCount = 0;
            this.ignoredAutoRefreshUntil = 0;
            return false;
        }

        this.ignoredAutoRefreshTriggerCount -= 1;
        if (this.ignoredAutoRefreshTriggerCount === 0) {
            this.ignoredAutoRefreshUntil = 0;
        }

        return true;
    }

    private captureStructuralSnapshot() {
        return captureStructuralSnapshot(
            this.state,
            this.workingSheetEntries,
            this.pendingCellEdits,
            this.pendingSheetEdits,
            this.pendingViewEdits
        );
    }

    private restoreStructuralSnapshot(snapshot: StructuralHistoryEntry["before"]): void {
        const restored = restoreStructuralSnapshot(snapshot);
        this.state = restored.state;
        this.workingSheetEntries = restored.sheetEntries;
        this.pendingCellEdits = restored.pendingCellEdits;
        this.pendingSheetEdits = restored.pendingSheetEdits;
        this.pendingViewEdits = restored.pendingViewEdits;
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

    private getStructuralBaselineSheetSnapshot(
        sheetKey: string
    ): WorkbookSnapshot["sheets"][number] | undefined {
        if (!this.workbook) {
            return undefined;
        }

        return restorePendingWorkbookState(this.workbook, {
            cellEdits: [],
            sheetEdits: this.pendingSheetEdits,
            viewEdits: [],
        }).sheetEntries.find((entry) => entry.key === sheetKey)?.sheet;
    }

    private syncPendingSheetViewEdit(sheetKey: string): void {
        const entry = this.findSheetEntry(sheetKey);
        if (!entry) {
            this.pendingViewEdits = this.pendingViewEdits.filter(
                (edit) => edit.sheetKey !== sheetKey
            );
            return;
        }

        const baselineSheet = this.getStructuralBaselineSheetSnapshot(sheetKey);
        const hasFreezePaneChange = !areFreezePaneCountsEqual(
            baselineSheet?.freezePane ?? null,
            entry.sheet.freezePane ?? null
        );
        const hasColumnWidthChange = !areColumnWidthsEquivalent(
            baselineSheet?.columnWidths,
            entry.sheet.columnWidths
        );
        if (!hasFreezePaneChange && !hasColumnWidthChange) {
            this.pendingViewEdits = this.pendingViewEdits.filter(
                (edit) => edit.sheetKey !== sheetKey
            );
            return;
        }

        const nextEdit: SheetViewEdit = {
            sheetKey,
            sheetName: entry.sheet.name,
            freezePane: entry.sheet.freezePane
                ? {
                      columnCount: entry.sheet.freezePane.columnCount,
                      rowCount: entry.sheet.freezePane.rowCount,
                  }
                : null,
            ...(hasColumnWidthChange
                ? {
                      columnWidths: cloneColumnWidths(entry.sheet.columnWidths),
                  }
                : {}),
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
        return validateEditorSheetName(
            value,
            this.workingSheetEntries,
            getWebviewStrings(),
            currentSheetKey
        );
    }

    private validateColumnWidth(value: string): string | undefined {
        const trimmedValue = value.trim();
        if (trimmedValue.length === 0) {
            return undefined;
        }

        const nextWidth = Number(trimmedValue);
        if (!Number.isFinite(nextWidth) || nextWidth <= 0) {
            return getWebviewStrings().invalidColumnWidth;
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

    private updatePendingGridEditRename(sheetKey: string, nextName: string): void {
        this.pendingSheetEdits = this.pendingSheetEdits.map((edit) =>
            isGridSheetEdit(edit) && edit.sheetKey === sheetKey
                ? {
                      ...edit,
                      sheetName: nextName,
                  }
                : edit
        );
    }

    private setActiveSheetSelection(rowNumber: number, columnNumber: number): void {
        const workbook = this.getWorkingWorkbook();
        const activeSheetKey = this.state.activeSheetKey;
        if (!workbook || !activeSheetKey) {
            return;
        }

        this.state = setSelectedEditorCell(
            workbook,
            {
                ...this.state,
                activeSheetKey,
            },
            rowNumber,
            columnNumber,
            this.getSheetEntries()
        );
    }

    private async promptPendingColumnWidth(columnNumber: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            !Number.isInteger(columnNumber) ||
            columnNumber < 1 ||
            columnNumber > activeEntry.sheet.columnCount
        ) {
            return;
        }

        const strings = getWebviewStrings();
        const currentWidth = activeEntry.sheet.columnWidths?.[columnNumber - 1] ?? null;
        const nextValue = await vscode.window.showInputBox({
            prompt: strings.setColumnWidthPrompt,
            title: strings.setColumnWidthTitle,
            value: currentWidth === null ? "" : String(currentWidth),
            validateInput: (value) => this.validateColumnWidth(value),
        });
        if (nextValue === undefined) {
            return;
        }

        const trimmedValue = nextValue.trim();
        await this.setPendingColumnWidth(
            columnNumber,
            trimmedValue.length === 0 ? null : normalizeWorkbookColumnWidth(Number(trimmedValue))
        );
    }

    private async setPendingColumnWidth(
        columnNumber: number,
        nextWidth: number | null
    ): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            !Number.isInteger(columnNumber) ||
            columnNumber < 1 ||
            columnNumber > activeEntry.sheet.columnCount
        ) {
            return;
        }

        const normalizedWidth =
            nextWidth === null ? null : normalizeWorkbookColumnWidth(nextWidth);
        if (normalizedWidth !== null && (!Number.isFinite(normalizedWidth) || normalizedWidth <= 0)) {
            return;
        }

        const currentWidth = activeEntry.sheet.columnWidths?.[columnNumber - 1] ?? null;
        if (currentWidth === normalizedWidth) {
            return;
        }

        await this.commitStructuralMutation(
            () => {
                const entry = this.getActiveSheetEntry();
                if (!entry) {
                    return;
                }

                entry.sheet = {
                    ...entry.sheet,
                    columnWidths: setSheetColumnWidthSnapshot(
                        entry.sheet.columnWidths,
                        columnNumber,
                        normalizedWidth
                    ),
                };
                this.syncPendingSheetViewEdit(entry.key);
            },
            {
                resetPendingHistory: false,
            }
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

        const nextFreezePane = isLocked
            ? null
            : createFreezePaneSnapshot(nextColumnCount, nextRowCount);
        await this.commitStructuralMutation(
            () => {
                const entry = this.getActiveSheetEntry();
                if (!entry) {
                    return;
                }

                entry.sheet = {
                    ...entry.sheet,
                    freezePane: nextFreezePane,
                };
                this.syncPendingSheetViewEdit(entry.key);
            },
            {
                resetPendingHistory: false,
            }
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
            if (this.hasPendingExternalWorkbookChange) {
                this.hasPendingExternalWorkbookChange = false;
                await this.enqueueReload({ clearPendingEdits: true });
            }
        }
    }

    private getNewSheetName(): string {
        return getNewEditorSheetName(
            this.workingSheetEntries,
            getRuntimeMessages().workbook.newSheetBaseName
        );
    }

    private getInsertSheetIndex(): number {
        return getInsertEditorSheetIndex(this.workingSheetEntries, this.state.activeSheetKey);
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
                    visibility: "visible",
                    mergedRanges: [],
                    columnWidths: [],
                    cells: {},
                    freezePane: null,
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
                this.pendingSheetEdits = this.pendingSheetEdits.filter(
                    (edit) => !(isGridSheetEdit(edit) && edit.sheetKey === sheetKey)
                );
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
            this.updatePendingGridEditRename(sheetKey, trimmedName);
            this.syncPendingSheetViewEdit(sheetKey);
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

    private async insertPendingRow(rowNumber: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            !Number.isInteger(rowNumber) ||
            rowNumber < 1 ||
            rowNumber > activeEntry.sheet.rowCount + 1
        ) {
            return;
        }

        await this.commitStructuralMutation(() => {
            const entry = this.getActiveSheetEntry();
            if (!entry) {
                return;
            }

            const edit: SheetEdit = {
                type: "insertRow",
                sheetKey: entry.key,
                sheetName: entry.sheet.name,
                rowNumber,
                count: 1,
            };
            entry.sheet = applyGridSheetEditToSheet(entry.sheet, edit);
            this.pendingCellEdits = shiftPendingCellEditsForGridSheetEdit(
                this.pendingCellEdits,
                edit
            );
            this.pendingSheetEdits = [...this.pendingSheetEdits, edit];
            this.setActiveSheetSelection(
                rowNumber,
                Math.min(
                    Math.max(this.state.selectedCell?.columnNumber ?? 1, 1),
                    Math.max(entry.sheet.columnCount, 1)
                )
            );
        });
    }

    private async deletePendingRow(rowNumber: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            activeEntry.sheet.rowCount <= 1 ||
            !Number.isInteger(rowNumber) ||
            rowNumber < 1 ||
            rowNumber > activeEntry.sheet.rowCount
        ) {
            return;
        }

        await this.commitStructuralMutation(() => {
            const entry = this.getActiveSheetEntry();
            if (!entry) {
                return;
            }

            const edit: SheetEdit = {
                type: "deleteRow",
                sheetKey: entry.key,
                sheetName: entry.sheet.name,
                rowNumber,
                count: 1,
            };
            entry.sheet = applyGridSheetEditToSheet(entry.sheet, edit);
            this.pendingCellEdits = shiftPendingCellEditsForGridSheetEdit(
                this.pendingCellEdits,
                edit
            );
            this.pendingSheetEdits = [...this.pendingSheetEdits, edit];
            this.setActiveSheetSelection(
                Math.min(rowNumber, Math.max(entry.sheet.rowCount, 1)),
                Math.min(
                    Math.max(this.state.selectedCell?.columnNumber ?? 1, 1),
                    Math.max(entry.sheet.columnCount, 1)
                )
            );
        });
    }

    private async insertPendingColumn(columnNumber: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            !Number.isInteger(columnNumber) ||
            columnNumber < 1 ||
            columnNumber > activeEntry.sheet.columnCount + 1
        ) {
            return;
        }

        await this.commitStructuralMutation(() => {
            const entry = this.getActiveSheetEntry();
            if (!entry) {
                return;
            }

            const edit: SheetEdit = {
                type: "insertColumn",
                sheetKey: entry.key,
                sheetName: entry.sheet.name,
                columnNumber,
                count: 1,
            };
            entry.sheet = applyGridSheetEditToSheet(entry.sheet, edit);
            this.pendingCellEdits = shiftPendingCellEditsForGridSheetEdit(
                this.pendingCellEdits,
                edit
            );
            this.pendingSheetEdits = [...this.pendingSheetEdits, edit];
            this.syncPendingSheetViewEdit(entry.key);
            this.setActiveSheetSelection(
                Math.min(
                    Math.max(this.state.selectedCell?.rowNumber ?? 1, 1),
                    Math.max(entry.sheet.rowCount, 1)
                ),
                columnNumber
            );
        });
    }

    private async deletePendingColumn(columnNumber: number): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (
            !workbook ||
            workbook.isReadonly ||
            !activeEntry ||
            activeEntry.sheet.columnCount <= 1 ||
            !Number.isInteger(columnNumber) ||
            columnNumber < 1 ||
            columnNumber > activeEntry.sheet.columnCount
        ) {
            return;
        }

        await this.commitStructuralMutation(() => {
            const entry = this.getActiveSheetEntry();
            if (!entry) {
                return;
            }

            const edit: SheetEdit = {
                type: "deleteColumn",
                sheetKey: entry.key,
                sheetName: entry.sheet.name,
                columnNumber,
                count: 1,
            };
            entry.sheet = applyGridSheetEditToSheet(entry.sheet, edit);
            this.pendingCellEdits = shiftPendingCellEditsForGridSheetEdit(
                this.pendingCellEdits,
                edit
            );
            this.pendingSheetEdits = [...this.pendingSheetEdits, edit];
            this.syncPendingSheetViewEdit(entry.key);
            this.setActiveSheetSelection(
                Math.min(
                    Math.max(this.state.selectedCell?.rowNumber ?? 1, 1),
                    Math.max(entry.sheet.rowCount, 1)
                ),
                Math.min(columnNumber, Math.max(entry.sheet.columnCount, 1))
            );
        });
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
        const documentPendingState = this.document.getPendingState();
        const shouldRestoreDocumentPendingState =
            !clearPendingEdits &&
            this.pendingCellEdits.length === 0 &&
            this.pendingSheetEdits.length === 0 &&
            this.pendingViewEdits.length === 0 &&
            documentPendingState.cellEdits.length +
                documentPendingState.sheetEdits.length +
                (documentPendingState.viewEdits?.length ?? 0) >
                0;
        if (shouldRestoreDocumentPendingState) {
            const restored = restorePendingWorkbookState(this.workbook, documentPendingState);
            this.workingSheetEntries = restored.sheetEntries;
            this.pendingCellEdits = restored.pendingCellEdits;
            this.pendingSheetEdits = restored.pendingSheetEdits;
            this.pendingViewEdits = restored.pendingViewEdits;
            this.nextNewSheetId = restored.nextNewSheetId;
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
            this.hasPendingExternalWorkbookChange = false;
        }
        this.panel.title = renderModel.title;
        await this.render(renderModel, {
            silent,
            clearPendingEdits,
            useModelSelection: true,
            replacePendingEdits:
                clearPendingEdits || this.pendingCellEdits.length === 0
                    ? undefined
                    : this.pendingCellEdits,
        });
    }

    private async render(
        renderModel?: EditorRenderModel,
        {
            silent = false,
            clearPendingEdits = false,
            preservePendingHistory = false,
            reuseActiveSheetData = false,
            useModelSelection = true,
            replacePendingEdits,
            resetPendingHistory = false,
        }: {
            silent?: boolean;
            clearPendingEdits?: boolean;
            preservePendingHistory?: boolean;
            reuseActiveSheetData?: boolean;
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
            preservePendingHistory,
            reuseActiveSheetData,
            useModelSelection,
            replacePendingEdits:
                replacePendingEdits !== undefined
                    ? mapPendingCellEditsToWebview(replacePendingEdits, this.getSheetEntries())
                    : undefined,
            resetPendingHistory,
        });
    }
}
