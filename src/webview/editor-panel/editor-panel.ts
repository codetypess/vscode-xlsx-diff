import * as path from "node:path";
import * as vscode from "vscode";
import { DEFAULT_PAGE_SIZE } from "../../constants";
import {
    type CellEdit,
    type SheetEdit,
    type SheetViewEdit,
} from "../../core/fastxlsx/write-cell-value";
import {
    areCellAlignmentsEqual,
    type EditorAlignmentPatch,
} from "../../core/model/alignment";
import { createCellKey } from "../../core/model/cells";
import type {
    EditorPanelState,
    EditorRenderModel,
    SheetSnapshot,
    WorkbookSnapshot,
} from "../../core/model/types";
import { toErrorMessage } from "../../error-message";
import { getRuntimeMessages } from "../../i18n";
import { rememberRecentWorkbookResourceUri } from "../../scm/recent-workbook-resource-context";
import {
    getInsertEditorSheetIndex,
    getNewEditorSheetName,
    resolveEditorSearchResult,
    resolveEditorCellReference,
    validateEditorSheetName,
} from "./editor-panel-logic";
import {
    createSessionStatusMessage,
    isEditorWebviewOutgoingMessage,
    type EditorWebviewOutgoingMessage,
} from "../shared/session-protocol";
import type {
    EditorAlignmentTargetKind,
    EditorPanelStrings,
    EditorSearchResultMessage,
    StructuralHistoryEntry,
    WorkingSheetEntry,
    XlsxEditorPanelController,
} from "./editor-panel-types";
import type { SelectionRange } from "./editor-selection-range";
import { hasLockedView } from "../shared/view-lock";
import { getWorkbookResourceName } from "../../workbook/resource-uri";
import { escapeWatcherGlobSegment } from "../webview-utils";
import { logPerf, summarizePendingStateForPerf } from "./editor-perf-log";
import {
    applyAlignmentPatchToSheetSnapshot,
    applyGridSheetEditToSheet,
    areAutoFiltersEquivalent,
    areCellAlignmentsEquivalent,
    areColumnWidthsEquivalent,
    areColumnAlignmentsEquivalent,
    areFreezePaneCountsEqual,
    areRowHeightsEquivalent,
    areRowAlignmentsEquivalent,
    captureStructuralSnapshot,
    cloneAutoFilterSnapshot,
    createFreezePaneSnapshot,
    createPendingWorkbookEditState,
    isGridSheetEdit,
    restorePendingWorkbookState,
    setSheetColumnWidthSnapshot,
    setSheetRowHeightSnapshot,
    createWorkingWorkbook,
    reindexWorkingSheetEntries,
    restoreStructuralSnapshot,
    shiftPendingCellEditsForGridSheetEdit,
} from "./editor-panel-state";
import {
    createActiveSheetAlignmentRenderPatch,
    createEditorPanelHtml,
    createEditorRenderPayload,
    createEditorSessionMessage,
    type ActiveSheetAlignmentRenderPatch,
} from "./editor-panel-protocol-adapter";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    normalizeEditorPanelState,
    setActiveEditorSheet,
    setSelectedEditorCell,
} from "./editor-render-model";
import {
    commitEditorWorkbookSession,
    loadEditorWorkbookSession,
} from "./editor-workbook-session";
import { XlsxEditorDocument } from "./xlsx-editor-document";

function getWebviewStrings(): EditorPanelStrings {
    return getRuntimeMessages().editorPanel;
}

function normalizeWorkbookColumnWidth(columnWidth: number): number {
    return Math.round(columnWidth * 256) / 256;
}

function normalizeWorkbookRowHeight(rowHeight: number): number {
    return Math.round(rowHeight * 100) / 100;
}

function countRecordEntries(record: Readonly<Record<string, unknown>> | undefined): number {
    return Object.keys(record ?? {}).length;
}

function summarizeSheetForPerf(
    sheet:
        | {
              name?: string;
              rowCount?: number;
              columnCount?: number;
              cells?: Readonly<Record<string, unknown>>;
              cellAlignments?: Readonly<Record<string, unknown>>;
              rowAlignments?: Readonly<Record<string, unknown>>;
              columnAlignments?: Readonly<Record<string, unknown>>;
              rowHeights?: Readonly<Record<string, unknown>>;
              columnWidths?: readonly unknown[];
          }
        | null
        | undefined
): Record<string, unknown> {
    return {
        sheetName: sheet?.name ?? null,
        rowCount: sheet?.rowCount ?? null,
        columnCount: sheet?.columnCount ?? null,
        cellCount: countRecordEntries(sheet?.cells),
        cellAlignmentCount: countRecordEntries(sheet?.cellAlignments),
        rowAlignmentCount: countRecordEntries(sheet?.rowAlignments),
        columnAlignmentCount: countRecordEntries(sheet?.columnAlignments),
        rowHeightCount: countRecordEntries(sheet?.rowHeights),
        columnWidthCount: sheet?.columnWidths?.length ?? 0,
    };
}

function logEditorPanelPerf(event: string, details: Record<string, unknown> = {}): void {
    logPerf("host", event, details);
}

interface NumericSheetDimensionPromptOptions {
    currentValue: number | null;
    prompt: string;
    title: string;
    invalidMessage: string;
    normalize(value: number): number;
}

interface NumericSheetDimensionMutationOptions {
    currentValue: number | null;
    nextValue: number | null;
    normalize(value: number): number;
    updateSheet(sheet: SheetSnapshot, nextValue: number | null): SheetSnapshot;
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
    private static isDebugMode = false;

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
    private hasSentSessionInit = false;
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
    private lastRenderedModel: EditorRenderModel | null = null;
    private activePerfTraceId: string | null = null;

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
            (message: unknown) => {
                if (!isEditorWebviewOutgoingMessage(message)) {
                    return;
                }

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

    public static setDebugMode(isDebugMode: boolean): void {
        XlsxEditorPanel.isDebugMode = isDebugMode;
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
        const scheduleRefresh = () => {
            this.scheduleAutoRefresh();
        };

        this.fileWatchers.push(watcher);
        this.fileWatchers.push(watcher.onDidChange(scheduleRefresh));
        this.fileWatchers.push(watcher.onDidCreate(scheduleRefresh));
        this.fileWatchers.push(watcher.onDidDelete(scheduleRefresh));
    }

    private scheduleAutoRefresh(): void {
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
            await this.panel.webview.postMessage(
                createSessionStatusMessage("editor", "error", errorMessage)
            );
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
        this.hasSentSessionInit = false;
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
        return createEditorPanelHtml({
            webview: this.panel.webview,
            extensionUri: this.extensionUri,
            strings: getWebviewStrings(),
            isDebugMode: XlsxEditorPanel.isDebugMode,
        });
    }

    private async handleMessage(message: EditorWebviewOutgoingMessage): Promise<void> {
        try {
            switch (message.type) {
                case "ready":
                    this.isWebviewReady = true;
                    this.hasSentSessionInit = false;
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
                case "promptRowHeight":
                    await this.promptPendingRowHeight(message.rowNumber);
                    return;
                case "setRowHeight":
                    await this.setPendingRowHeight(message.rowNumber, message.height);
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
                case "setAlignment":
                    logEditorPanelPerf("handleMessage:setAlignment", {
                        perfTraceId: message.perfTraceId ?? null,
                        target: message.target,
                        selection: message.selection,
                        alignment: message.alignment,
                    });
                    await this.setPendingAlignment(
                        message.target,
                        message.selection,
                        message.alignment,
                        message.perfTraceId
                    );
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
                case "setFilterState":
                    await this.setPendingFilterState(message.sheetKey, message.filterState);
                    return;
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
        const startedAt = performance.now();
        logEditorPanelPerf("commitSavedState:start", {
            panelId: this.panelId,
            hasWorkbook: Boolean(this.workbook),
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
        if (!this.workbook) {
            await this.enqueueReload({ silent: true, clearPendingEdits: true });
            logEditorPanelPerf("commitSavedState:reloaded", {
                panelId: this.panelId,
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                activeSheetKey: this.state.activeSheetKey,
            });
            return;
        }

        const committedSession = commitEditorWorkbookSession({
            workbook: this.workbook,
            workingSheetEntries: this.workingSheetEntries,
            pendingCellEdits: this.pendingCellEdits,
            currentPanelState: this.state,
        });

        this.workbook = committedSession.workbook;
        this.workingSheetEntries = committedSession.workingSheetEntries;
        this.nextNewSheetId = committedSession.nextNewSheetId;
        this.pendingCellEdits = committedSession.pendingCellEdits;
        this.pendingSheetEdits = committedSession.pendingSheetEdits;
        this.pendingViewEdits = committedSession.pendingViewEdits;
        this.sheetUndoStack.length = 0;
        this.sheetRedoStack.length = 0;
        this.hasWarnedPendingExternalChange = false;
        this.hasPendingExternalWorkbookChange = false;
        this.state = committedSession.panelState;

        await this.render(undefined, {
            silent: true,
            clearPendingEdits: true,
            preservePendingHistory: true,
            useModelSelection: true,
        });
        logEditorPanelPerf("commitSavedState:done", {
            panelId: this.panelId,
            durationMs: Number((performance.now() - startedAt).toFixed(2)),
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
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
        const startedAt = performance.now();
        logEditorPanelPerf("requestDocumentSave:start", {
            panelId: this.panelId,
            activeSheetKey: this.state.activeSheetKey,
            hasPendingExternalWorkbookChange: this.hasPendingExternalWorkbookChange,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
        const shouldSave = await XlsxEditorPanel.confirmDocumentSave(this.document);
        if (!shouldSave) {
            logEditorPanelPerf("requestDocumentSave:cancelled", {
                panelId: this.panelId,
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                activeSheetKey: this.state.activeSheetKey,
                hasPendingExternalWorkbookChange: this.hasPendingExternalWorkbookChange,
            });
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
        logEditorPanelPerf("requestDocumentSave:dispatched", {
            panelId: this.panelId,
            durationMs: Number((performance.now() - startedAt).toFixed(2)),
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
    }

    private async handleDocumentSave(): Promise<void> {
        const startedAt = performance.now();
        logEditorPanelPerf("handleDocumentSave:start", {
            panelId: this.panelId,
            activeSheetKey: this.state.activeSheetKey,
            hasPendingExternalWorkbookChange: this.hasPendingExternalWorkbookChange,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
        this.isSavingDocument = false;
        this.noteLocalSaveCompletion();
        if (this.hasPendingExternalWorkbookChange) {
            this.hasPendingExternalWorkbookChange = false;
            await this.enqueueReload({ silent: true, clearPendingEdits: true });
            logEditorPanelPerf("handleDocumentSave:reloaded", {
                panelId: this.panelId,
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                activeSheetKey: this.state.activeSheetKey,
            });
            return;
        }

        await this.commitSavedState();
        logEditorPanelPerf("handleDocumentSave:done", {
            panelId: this.panelId,
            durationMs: Number((performance.now() - startedAt).toFixed(2)),
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
    }

    private startDocumentSave(): void {
        this.isSavingDocument = true;
        this.clearAutoRefreshTimer();
        this.suppressAutoRefreshUntil = Number.POSITIVE_INFINITY;
        logEditorPanelPerf("startDocumentSave", {
            panelId: this.panelId,
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
    }

    private cancelDocumentSave(): void {
        this.isSavingDocument = false;
        this.clearAutoRefreshTimer();
        this.suppressAutoRefreshUntil = Date.now() + 1500;
        logEditorPanelPerf("cancelDocumentSave", {
            panelId: this.panelId,
            activeSheetKey: this.state.activeSheetKey,
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
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
        {
            resetPendingHistory = true,
            activeSheetAlignmentRenderPatch,
        }: {
            resetPendingHistory?: boolean;
            activeSheetAlignmentRenderPatch?: ActiveSheetAlignmentRenderPatch;
        } = {}
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
            ...(activeSheetAlignmentRenderPatch
                ? {
                      activeSheetAlignmentRenderPatch,
                  }
                : {}),
        });
    }

    private getStructuralBaselineSheetSnapshot(
        sheetKey: string
    ): WorkbookSnapshot["sheets"][number] | undefined {
        if (!this.workbook) {
            return undefined;
        }

        if (this.pendingSheetEdits.length === 0) {
            const match = /^sheet:(\d+)$/.exec(sheetKey);
            const sheetIndex = match ? Number(match[1]) : Number.NaN;
            return Number.isInteger(sheetIndex) ? this.workbook.sheets[sheetIndex] : undefined;
        }

        return restorePendingWorkbookState(this.workbook, {
            cellEdits: [],
            sheetEdits: this.pendingSheetEdits,
            viewEdits: [],
        }).sheetEntries.find((entry) => entry.key === sheetKey)?.sheet;
    }

    private createSheetFreezePaneEdit(
        sheet: WorkbookSnapshot["sheets"][number]
    ): SheetViewEdit["freezePane"] {
        return sheet.freezePane
            ? {
                  columnCount: sheet.freezePane.columnCount,
                  rowCount: sheet.freezePane.rowCount,
              }
            : null;
    }

    private findPendingViewEdit(sheetKey: string): SheetViewEdit | undefined {
        return this.pendingViewEdits.find((edit) => edit.sheetKey === sheetKey);
    }

    private ensurePendingViewEdit(entry: WorkingSheetEntry): SheetViewEdit {
        const existingEdit = this.findPendingViewEdit(entry.key);
        if (existingEdit) {
            existingEdit.sheetName = entry.sheet.name;
            existingEdit.freezePane = this.createSheetFreezePaneEdit(entry.sheet);
            return existingEdit;
        }

        const nextEdit: SheetViewEdit = {
            sheetKey: entry.key,
            sheetName: entry.sheet.name,
            freezePane: this.createSheetFreezePaneEdit(entry.sheet),
        };
        this.pendingViewEdits = [...this.pendingViewEdits, nextEdit];
        return nextEdit;
    }

    private cleanupPendingViewEdit(
        entry: WorkingSheetEntry,
        baselineSheet: WorkbookSnapshot["sheets"][number] | undefined
    ): void {
        const edit = this.findPendingViewEdit(entry.key);
        if (!edit) {
            return;
        }

        const hasFreezePaneChange = !areFreezePaneCountsEqual(
            baselineSheet?.freezePane ?? null,
            entry.sheet.freezePane ?? null
        );
        if (
            hasFreezePaneChange ||
            edit.autoFilter !== undefined ||
            edit.columnWidths !== undefined ||
            edit.rowHeights !== undefined ||
            edit.cellAlignments !== undefined ||
            edit.rowAlignments !== undefined ||
            edit.columnAlignments !== undefined
        ) {
            edit.sheetName = entry.sheet.name;
            edit.freezePane = this.createSheetFreezePaneEdit(entry.sheet);
            return;
        }

        this.pendingViewEdits = this.pendingViewEdits.filter(
            (candidate) => candidate.sheetKey !== entry.key
        );
    }

    private updatePendingAlignmentEditKeys(
        edit: SheetViewEdit,
        property: "cellAlignments" | "rowAlignments" | "columnAlignments",
        dirtyKeyProperty:
            | "dirtyCellAlignmentKeys"
            | "dirtyRowAlignmentKeys"
            | "dirtyColumnAlignmentKeys",
        alignments:
            | SheetSnapshot["cellAlignments"]
            | SheetSnapshot["rowAlignments"]
            | SheetSnapshot["columnAlignments"],
        dirtyKeys: Set<string>
    ): void {
        if (dirtyKeys.size === 0) {
            delete edit[property];
            delete edit[dirtyKeyProperty];
            return;
        }

        edit[property] = alignments;
        edit[dirtyKeyProperty] = [...dirtyKeys].sort((left, right) => left.localeCompare(right));
    }

    private syncPendingAlignmentViewEdit(
        entry: WorkingSheetEntry,
        target: EditorAlignmentTargetKind,
        selection: SelectionRange,
        perfTraceId?: string
    ): void {
        const baselineSheet = this.getStructuralBaselineSheetSnapshot(entry.key);
        const edit = this.ensurePendingViewEdit(entry);
        const logDirtyState = () => {
            logEditorPanelPerf("syncPendingAlignmentViewEdit", {
                perfTraceId: perfTraceId ?? null,
                target,
                selection,
                dirtyCellAlignmentKeyCount: edit.dirtyCellAlignmentKeys?.length ?? 0,
                dirtyRowAlignmentKeyCount: edit.dirtyRowAlignmentKeys?.length ?? 0,
                dirtyColumnAlignmentKeyCount: edit.dirtyColumnAlignmentKeys?.length ?? 0,
                ...summarizeSheetForPerf(entry.sheet),
            });
        };

        const dirtyCellAlignmentKeys = new Set(edit.dirtyCellAlignmentKeys ?? []);
        const dirtyRowAlignmentKeys = new Set(edit.dirtyRowAlignmentKeys ?? []);
        const dirtyColumnAlignmentKeys = new Set(edit.dirtyColumnAlignmentKeys ?? []);

        const updateDirtyKey = (
            dirtyKeys: Set<string>,
            key: string,
            baselineAlignment: Parameters<typeof areCellAlignmentsEqual>[0],
            currentAlignment: Parameters<typeof areCellAlignmentsEqual>[1]
        ) => {
            if (areCellAlignmentsEqual(baselineAlignment, currentAlignment)) {
                dirtyKeys.delete(key);
                return;
            }

            dirtyKeys.add(key);
        };

        if (target === "cell" || target === "range") {
            for (
                let rowNumber = selection.startRow;
                rowNumber <= selection.endRow;
                rowNumber += 1
            ) {
                for (
                    let columnNumber = selection.startColumn;
                    columnNumber <= selection.endColumn;
                    columnNumber += 1
                ) {
                    const cellKey = createCellKey(rowNumber, columnNumber);
                    updateDirtyKey(
                        dirtyCellAlignmentKeys,
                        cellKey,
                        baselineSheet?.cellAlignments?.[cellKey],
                        entry.sheet.cellAlignments?.[cellKey]
                    );
                }
            }

            this.updatePendingAlignmentEditKeys(
                edit,
                "cellAlignments",
                "dirtyCellAlignmentKeys",
                entry.sheet.cellAlignments,
                dirtyCellAlignmentKeys
            );
            this.cleanupPendingViewEdit(entry, baselineSheet);
            logDirtyState();
            return;
        }

        if (target === "row") {
            for (
                let rowNumber = selection.startRow;
                rowNumber <= selection.endRow;
                rowNumber += 1
            ) {
                const rowKey = String(rowNumber);
                updateDirtyKey(
                    dirtyRowAlignmentKeys,
                    rowKey,
                    baselineSheet?.rowAlignments?.[rowKey],
                    entry.sheet.rowAlignments?.[rowKey]
                );
            }

            this.updatePendingAlignmentEditKeys(
                edit,
                "rowAlignments",
                "dirtyRowAlignmentKeys",
                entry.sheet.rowAlignments,
                dirtyRowAlignmentKeys
            );
            this.cleanupPendingViewEdit(entry, baselineSheet);
            logDirtyState();
            return;
        }

        for (
            let columnNumber = selection.startColumn;
            columnNumber <= selection.endColumn;
            columnNumber += 1
        ) {
            const columnKey = String(columnNumber);
            updateDirtyKey(
                dirtyColumnAlignmentKeys,
                columnKey,
                baselineSheet?.columnAlignments?.[columnKey],
                entry.sheet.columnAlignments?.[columnKey]
            );
        }

        this.updatePendingAlignmentEditKeys(
            edit,
            "columnAlignments",
            "dirtyColumnAlignmentKeys",
            entry.sheet.columnAlignments,
            dirtyColumnAlignmentKeys
        );
        this.cleanupPendingViewEdit(entry, baselineSheet);
        logDirtyState();
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
        const hasAutoFilterChange = !areAutoFiltersEquivalent(
            baselineSheet?.autoFilter ?? null,
            entry.sheet.autoFilter ?? null
        );
        const hasRowHeightChange = !areRowHeightsEquivalent(
            baselineSheet?.rowHeights,
            entry.sheet.rowHeights
        );
        const hasColumnWidthChange = !areColumnWidthsEquivalent(
            baselineSheet?.columnWidths,
            entry.sheet.columnWidths
        );
        const hasCellAlignmentChange = !areCellAlignmentsEquivalent(
            baselineSheet?.cellAlignments,
            entry.sheet.cellAlignments
        );
        const hasRowAlignmentChange = !areRowAlignmentsEquivalent(
            baselineSheet?.rowAlignments,
            entry.sheet.rowAlignments
        );
        const hasColumnAlignmentChange = !areColumnAlignmentsEquivalent(
            baselineSheet?.columnAlignments,
            entry.sheet.columnAlignments
        );
        if (
            !hasFreezePaneChange &&
            !hasAutoFilterChange &&
            !hasColumnWidthChange &&
            !hasRowHeightChange &&
            !hasCellAlignmentChange &&
            !hasRowAlignmentChange &&
            !hasColumnAlignmentChange
        ) {
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
            ...(hasAutoFilterChange
                ? {
                      autoFilter: cloneAutoFilterSnapshot(entry.sheet.autoFilter),
                  }
                : {}),
            ...(hasColumnWidthChange
                ? {
                      columnWidths: entry.sheet.columnWidths,
                  }
                : {}),
            ...(hasRowHeightChange
                ? {
                      rowHeights: entry.sheet.rowHeights,
                  }
                : {}),
            ...(hasCellAlignmentChange
                ? {
                      cellAlignments: entry.sheet.cellAlignments,
                  }
                : {}),
            ...(hasRowAlignmentChange
                ? {
                      rowAlignments: entry.sheet.rowAlignments,
                  }
                : {}),
            ...(hasColumnAlignmentChange
                ? {
                      columnAlignments: entry.sheet.columnAlignments,
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

    private getEditableActiveSheetEntry(): WorkingSheetEntry | null {
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (!workbook || workbook.isReadonly || !activeEntry) {
            return null;
        }

        return activeEntry;
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
        return this.validateOptionalPositiveNumber(value, getWebviewStrings().invalidColumnWidth);
    }

    private validateRowHeight(value: string): string | undefined {
        return this.validateOptionalPositiveNumber(value, getWebviewStrings().invalidRowHeight);
    }

    private validateOptionalPositiveNumber(
        value: string,
        invalidMessage: string
    ): string | undefined {
        const trimmedValue = value.trim();
        if (trimmedValue.length === 0) {
            return undefined;
        }

        const nextValue = Number(trimmedValue);
        if (!Number.isFinite(nextValue) || nextValue <= 0) {
            return invalidMessage;
        }

        return undefined;
    }

    private isValidSheetDimensionTarget(targetNumber: number, totalCount: number): boolean {
        return Number.isInteger(targetNumber) && targetNumber >= 1 && targetNumber <= totalCount;
    }

    private async promptPendingDimensionValue({
        currentValue,
        prompt,
        title,
        invalidMessage,
        normalize,
    }: NumericSheetDimensionPromptOptions): Promise<number | null | undefined> {
        const nextValue = await vscode.window.showInputBox({
            prompt,
            title,
            value: currentValue === null ? "" : String(currentValue),
            validateInput: (value) => this.validateOptionalPositiveNumber(value, invalidMessage),
        });
        if (nextValue === undefined) {
            return undefined;
        }

        const trimmedValue = nextValue.trim();
        return trimmedValue.length === 0 ? null : normalize(Number(trimmedValue));
    }

    private async commitPendingActiveSheetDimensionValue({
        currentValue,
        nextValue,
        normalize,
        updateSheet,
    }: NumericSheetDimensionMutationOptions): Promise<void> {
        const normalizedValue = nextValue === null ? null : normalize(nextValue);
        if (
            normalizedValue !== null &&
            (!Number.isFinite(normalizedValue) || normalizedValue <= 0)
        ) {
            return;
        }

        if (currentValue === normalizedValue) {
            return;
        }

        await this.commitStructuralMutation(
            () => {
                const entry = this.getActiveSheetEntry();
                if (!entry) {
                    return;
                }

                entry.sheet = updateSheet(entry.sheet, normalizedValue);
                this.syncPendingSheetViewEdit(entry.key);
            },
            {
                resetPendingHistory: false,
            }
        );
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
        const activeEntry = this.getEditableActiveSheetEntry();
        if (
            !activeEntry ||
            !this.isValidSheetDimensionTarget(columnNumber, activeEntry.sheet.columnCount)
        ) {
            return;
        }

        const strings = getWebviewStrings();
        const currentWidth = activeEntry.sheet.columnWidths?.[columnNumber - 1] ?? null;
        const nextWidth = await this.promptPendingDimensionValue({
            currentValue: currentWidth,
            prompt: strings.setColumnWidthPrompt,
            title: strings.setColumnWidthTitle,
            invalidMessage: strings.invalidColumnWidth,
            normalize: normalizeWorkbookColumnWidth,
        });
        if (nextWidth === undefined) {
            return;
        }

        await this.setPendingColumnWidth(columnNumber, nextWidth);
    }

    private async promptPendingRowHeight(rowNumber: number): Promise<void> {
        const activeEntry = this.getEditableActiveSheetEntry();
        if (
            !activeEntry ||
            !this.isValidSheetDimensionTarget(rowNumber, activeEntry.sheet.rowCount)
        ) {
            return;
        }

        const strings = getWebviewStrings();
        const currentHeight = activeEntry.sheet.rowHeights?.[String(rowNumber)] ?? null;
        const nextHeight = await this.promptPendingDimensionValue({
            currentValue: currentHeight,
            prompt: strings.setRowHeightPrompt,
            title: strings.setRowHeightTitle,
            invalidMessage: strings.invalidRowHeight,
            normalize: normalizeWorkbookRowHeight,
        });
        if (nextHeight === undefined) {
            return;
        }

        await this.setPendingRowHeight(rowNumber, nextHeight);
    }

    private async setPendingColumnWidth(
        columnNumber: number,
        nextWidth: number | null
    ): Promise<void> {
        const activeEntry = this.getEditableActiveSheetEntry();
        if (
            !activeEntry ||
            !this.isValidSheetDimensionTarget(columnNumber, activeEntry.sheet.columnCount)
        ) {
            return;
        }

        await this.commitPendingActiveSheetDimensionValue({
            currentValue: activeEntry.sheet.columnWidths?.[columnNumber - 1] ?? null,
            nextValue: nextWidth,
            normalize: normalizeWorkbookColumnWidth,
            updateSheet: (sheet, normalizedWidth) => ({
                ...sheet,
                columnWidths: setSheetColumnWidthSnapshot(
                    sheet.columnWidths,
                    columnNumber,
                    normalizedWidth
                ),
            }),
        });
    }

    private async setPendingRowHeight(rowNumber: number, nextHeight: number | null): Promise<void> {
        const activeEntry = this.getEditableActiveSheetEntry();
        if (
            !activeEntry ||
            !this.isValidSheetDimensionTarget(rowNumber, activeEntry.sheet.rowCount)
        ) {
            return;
        }

        await this.commitPendingActiveSheetDimensionValue({
            currentValue: activeEntry.sheet.rowHeights?.[String(rowNumber)] ?? null,
            nextValue: nextHeight,
            normalize: normalizeWorkbookRowHeight,
            updateSheet: (sheet, normalizedHeight) => ({
                ...sheet,
                rowHeights: setSheetRowHeightSnapshot(
                    sheet.rowHeights,
                    rowNumber,
                    normalizedHeight
                ),
            }),
        });
    }

    private async setPendingAlignment(
        target: EditorAlignmentTargetKind,
        selection: SelectionRange,
        alignment: EditorAlignmentPatch,
        perfTraceId?: string
    ): Promise<void> {
        const startedAt = performance.now();
        const workbook = this.getWorkingWorkbook();
        const activeEntry = this.getActiveSheetEntry();
        if (!workbook || workbook.isReadonly || !activeEntry) {
            return;
        }

        if (!alignment.horizontal && !alignment.vertical) {
            return;
        }

        const nextAlignment: EditorAlignmentPatch = {
            ...(alignment.horizontal ? { horizontal: alignment.horizontal } : {}),
            ...(alignment.vertical ? { vertical: alignment.vertical } : {}),
        };

        const normalizedSelection: SelectionRange = {
            startRow: Math.min(selection.startRow, selection.endRow),
            endRow: Math.max(selection.startRow, selection.endRow),
            startColumn: Math.min(selection.startColumn, selection.endColumn),
            endColumn: Math.max(selection.startColumn, selection.endColumn),
        };
        const isSelectionInBounds =
            Number.isInteger(normalizedSelection.startRow) &&
            Number.isInteger(normalizedSelection.endRow) &&
            Number.isInteger(normalizedSelection.startColumn) &&
            Number.isInteger(normalizedSelection.endColumn) &&
            normalizedSelection.startRow >= 1 &&
            normalizedSelection.startColumn >= 1 &&
            normalizedSelection.endRow <= activeEntry.sheet.rowCount &&
            normalizedSelection.endColumn <= activeEntry.sheet.columnCount;
        if (!isSelectionInBounds) {
            return;
        }

        const selectsAllColumns =
            normalizedSelection.startColumn === 1 &&
            normalizedSelection.endColumn === activeEntry.sheet.columnCount;
        const selectsAllRows =
            normalizedSelection.startRow === 1 &&
            normalizedSelection.endRow === activeEntry.sheet.rowCount;
        if (
            (target === "cell" &&
                (normalizedSelection.startRow !== normalizedSelection.endRow ||
                    normalizedSelection.startColumn !== normalizedSelection.endColumn)) ||
            (target === "row" && !selectsAllColumns) ||
            (target === "column" && !selectsAllRows)
        ) {
            return;
        }

        logEditorPanelPerf("setPendingAlignment:start", {
            perfTraceId: perfTraceId ?? null,
            target,
            selection: normalizedSelection,
            alignment: nextAlignment,
            ...summarizeSheetForPerf(activeEntry.sheet),
            ...summarizePendingStateForPerf(this.createPendingState()),
        });
        const previewSheet = applyAlignmentPatchToSheetSnapshot(
            activeEntry.sheet,
            target,
            normalizedSelection,
            nextAlignment
        );
        logEditorPanelPerf("setPendingAlignment:preview", {
            perfTraceId: perfTraceId ?? null,
            durationMs: Number((performance.now() - startedAt).toFixed(2)),
            sheetChanged: previewSheet !== activeEntry.sheet,
            ...summarizeSheetForPerf(previewSheet),
        });
        if (previewSheet === activeEntry.sheet) {
            return;
        }

        this.activePerfTraceId = perfTraceId ?? null;
        try {
            await this.commitStructuralMutation(
                () => {
                    const entry = this.getActiveSheetEntry();
                    if (!entry) {
                        return;
                    }

                    entry.sheet = previewSheet;
                    this.syncPendingAlignmentViewEdit(
                        entry,
                        target,
                        normalizedSelection,
                        perfTraceId
                    );
                },
                {
                    resetPendingHistory: false,
                    activeSheetAlignmentRenderPatch: createActiveSheetAlignmentRenderPatch(
                        activeEntry.key,
                        target,
                        normalizedSelection
                    ),
                }
            );
        } finally {
            logEditorPanelPerf("setPendingAlignment:done", {
                perfTraceId: perfTraceId ?? null,
                totalDurationMs: Number((performance.now() - startedAt).toFixed(2)),
                ...summarizePendingStateForPerf(this.createPendingState()),
            });
            this.activePerfTraceId = null;
        }
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

    private async setPendingFilterState(
        sheetKey: string,
        filterState: WorkbookSnapshot["sheets"][number]["autoFilter"]
    ): Promise<void> {
        const workbook = this.getWorkingWorkbook();
        const entry = this.findSheetEntry(sheetKey);
        if (!workbook || workbook.isReadonly || !entry) {
            return;
        }

        if (areAutoFiltersEquivalent(entry.sheet.autoFilter ?? null, filterState ?? null)) {
            return;
        }

        await this.commitStructuralMutation(
            () => {
                const activeEntry = this.findSheetEntry(sheetKey);
                if (!activeEntry) {
                    return;
                }

                activeEntry.sheet = {
                    ...activeEntry.sheet,
                    autoFilter: cloneAutoFilterSnapshot(filterState),
                };
                this.syncPendingSheetViewEdit(activeEntry.key);
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
        const pendingState = this.createPendingState();
        const startedAt = performance.now();
        await this.controller.onPendingStateChanged(pendingState);
        if (this.activePerfTraceId) {
            logEditorPanelPerf("syncPendingState", {
                perfTraceId: this.activePerfTraceId,
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                ...summarizePendingStateForPerf(pendingState),
            });
        }

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
                    rowHeights: {},
                    cellAlignments: {},
                    rowAlignments: {},
                    columnAlignments: {},
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
            this.syncPendingSheetViewEdit(entry.key);
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
            this.syncPendingSheetViewEdit(entry.key);
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
                await this.panel.webview.postMessage(
                    createSessionStatusMessage("editor", "loading", webviewStrings.loading)
                );
            }
        }

        const reloadedSession = await loadEditorWorkbookSession({
            readUri: this.document.getReadUri(),
            workbookUri: this.workbookUri,
            clearPendingEdits,
            currentPanelState: this.state,
            pendingCellEdits: this.pendingCellEdits,
            pendingSheetEdits: this.pendingSheetEdits,
            pendingViewEdits: this.pendingViewEdits,
            documentPendingState: this.document.getPendingState(),
            documentHasPendingEdits: clearPendingEdits ? false : this.document.hasPendingEdits(),
            canUndoStructuralEdits: this.sheetUndoStack.length > 0,
            canRedoStructuralEdits: this.sheetRedoStack.length > 0,
        });
        this.workbook = reloadedSession.workbook;
        this.workingSheetEntries = reloadedSession.workingSheetEntries;
        this.pendingCellEdits = reloadedSession.pendingCellEdits;
        this.pendingSheetEdits = reloadedSession.pendingSheetEdits;
        this.pendingViewEdits = reloadedSession.pendingViewEdits;
        this.nextNewSheetId = reloadedSession.nextNewSheetId;
        this.state = reloadedSession.panelState;
        if (clearPendingEdits) {
            this.sheetUndoStack.length = 0;
            this.sheetRedoStack.length = 0;
        }

        if (clearPendingEdits) {
            this.hasWarnedPendingExternalChange = false;
            this.hasPendingExternalWorkbookChange = false;
        }
        this.panel.title = reloadedSession.renderModel.title;
        await this.render(reloadedSession.renderModel, {
            silent,
            clearPendingEdits,
            useModelSelection: true,
            replacePendingEdits:
                clearPendingEdits || reloadedSession.pendingCellEdits.length === 0
                    ? undefined
                    : reloadedSession.pendingCellEdits,
        });
    }

    private async render(
        renderModel?: EditorRenderModel,
        {
            silent = false,
            clearPendingEdits = false,
            preservePendingHistory = false,
            reuseActiveSheetData = true,
            useModelSelection = true,
            replacePendingEdits,
            resetPendingHistory = false,
            activeSheetAlignmentRenderPatch,
        }: {
            silent?: boolean;
            clearPendingEdits?: boolean;
            preservePendingHistory?: boolean;
            reuseActiveSheetData?: boolean;
            useModelSelection?: boolean;
            replacePendingEdits?: CellEdit[];
            resetPendingHistory?: boolean;
            activeSheetAlignmentRenderPatch?: ActiveSheetAlignmentRenderPatch;
        } = {}
    ): Promise<void> {
        if (!this.workbook) {
            return;
        }

        const renderStartedAt = performance.now();
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

        const renderPayload = createEditorRenderPayload({
            renderModel: payload,
            previousRenderModel: this.lastRenderedModel,
            reuseActiveSheetData,
            alignmentRenderPatch: activeSheetAlignmentRenderPatch,
        });
        if (this.activePerfTraceId) {
            logEditorPanelPerf("render:payload", {
                perfTraceId: this.activePerfTraceId,
                durationMs: Number((performance.now() - renderStartedAt).toFixed(2)),
                reuseActiveSheetData,
                payloadHasCells: renderPayload.activeSheet.cells !== undefined,
                payloadHasColumns: renderPayload.activeSheet.columns !== undefined,
                payloadHasColumnWidths: renderPayload.activeSheet.columnWidths !== undefined,
                payloadHasRowHeights: renderPayload.activeSheet.rowHeights !== undefined,
                payloadHasCellAlignments: renderPayload.activeSheet.cellAlignments !== undefined,
                payloadHasRowAlignments: renderPayload.activeSheet.rowAlignments !== undefined,
                payloadHasColumnAlignments:
                    renderPayload.activeSheet.columnAlignments !== undefined,
                payloadCellAlignmentDirtyKeyCount:
                    renderPayload.activeSheet.cellAlignmentDirtyKeys?.length ?? 0,
                payloadRowAlignmentDirtyKeyCount:
                    renderPayload.activeSheet.rowAlignmentDirtyKeys?.length ?? 0,
                payloadColumnAlignmentDirtyKeyCount:
                    renderPayload.activeSheet.columnAlignmentDirtyKeys?.length ?? 0,
                ...summarizeSheetForPerf(payload.activeSheet),
            });
        }
        this.hasPendingRender = false;
        const sessionMessage = createEditorSessionMessage({
            hasSentSessionInit: this.hasSentSessionInit,
            renderPayload,
            silent,
            clearPendingEdits,
            preservePendingHistory,
            reuseActiveSheetData,
            useModelSelection,
            replacePendingEdits,
            sheetEntries: this.getSheetEntries(),
            resetPendingHistory,
            perfTraceId: this.activePerfTraceId ?? null,
        });
        await this.panel.webview.postMessage(sessionMessage);
        this.hasSentSessionInit = true;
        if (this.activePerfTraceId) {
            logEditorPanelPerf("render:postMessage", {
                perfTraceId: this.activePerfTraceId,
                durationMs: Number((performance.now() - renderStartedAt).toFixed(2)),
                ...summarizeSheetForPerf(payload.activeSheet),
            });
        }
        this.lastRenderedModel = payload;
    }
}
