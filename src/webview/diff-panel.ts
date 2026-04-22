import * as path from "node:path";
import * as vscode from "vscode";
import { WEBVIEW_TYPE_DIFF_PANEL } from "../constants";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import { writeCellValues, type CellEdit } from "../core/fastxlsx/write-cell-value";
import type { WorkbookDiffModel } from "../core/model/types";
import { getHtmlLanguageTag, isChineseDisplayLanguage } from "../display-language";
import { getWorkbookResourceName } from "../workbook/resource-uri";
import { createDiffPanelRenderModel } from "./diff-panel-model";
import type { DiffPanelRenderModel } from "./diff-panel-types";

type WebviewMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | {
          type: "saveEdits";
          edits: Array<{
              sheetKey: string;
              side: "left" | "right";
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
      }
    | { type: "swap" }
    | { type: "reload" };

interface WebviewStrings {
    loading: string;
    reload: string;
    swap: string;
    all: string;
    diffs: string;
    same: string;
    allRows: string;
    diffRows: string;
    sameRows: string;
    prevDiff: string;
    nextDiff: string;
    sheets: string;
    diffCells: string;
    diffRowsShort: string;
    rows: string;
    columns: string;
    filter: string;
    visibleRows: string;
    currentDiff: string;
    selected: string;
    save: string;
    none: string;
    modified: string;
    size: string;
    readOnly: string;
    mergedRangesChanged: string;
    noSheet: string;
    noRows: string;
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

function getWebviewStrings(): WebviewStrings {
    if (isChineseDisplayLanguage()) {
        return {
            loading: "正在加载 XLSX 对比...",
            reload: "刷新",
            swap: "交换",
            all: "全部",
            diffs: "差异",
            same: "相同",
            allRows: "全部行",
            diffRows: "差异行",
            sameRows: "相同行",
            prevDiff: "上一处差异",
            nextDiff: "下一处差异",
            sheets: "工作表",
            diffCells: "差异单元格",
            diffRowsShort: "差异行",
            rows: "行",
            columns: "列",
            filter: "筛选",
            visibleRows: "可见行",
            currentDiff: "当前差异",
            selected: "选中",
            save: "保存",
            none: "-",
            modified: "修改时间",
            size: "大小",
            readOnly: "只读",
            mergedRangesChanged: "合并区域已变化",
            noSheet: "没有可显示的工作表。",
            noRows: "当前工作表没有可显示的行。",
        };
    }

    return {
        loading: "Loading XLSX diff...",
        reload: "Reload",
        swap: "Swap",
        all: "All",
        diffs: "Diffs",
        same: "Same",
        allRows: "All Rows",
        diffRows: "Diff Rows",
        sameRows: "Same Rows",
        prevDiff: "Prev Diff",
        nextDiff: "Next Diff",
        sheets: "Sheets",
        diffCells: "Diff Cells",
        diffRowsShort: "Diff Rows",
        rows: "Rows",
        columns: "Columns",
        filter: "Filter",
        visibleRows: "Visible Rows",
        currentDiff: "Current Diff",
        selected: "Selected",
        save: "Save",
        none: "-",
        modified: "Modified",
        size: "Size",
        readOnly: "Read-only",
        mergedRangesChanged: "Merged ranges changed",
        noSheet: "No sheet is available.",
        noRows: "No rows are available in this sheet.",
    };
}

export class XlsxDiffPanel {
    private static readonly panels = new Map<string, XlsxDiffPanel>();

    private readonly panel: vscode.WebviewPanel;
    private readonly extensionUri: vscode.Uri;
    private readonly panelKey: string;
    private readonly disposables: vscode.Disposable[] = [];
    private readonly fileWatchers: vscode.Disposable[] = [];

    private leftFileUri: vscode.Uri;
    private rightFileUri: vscode.Uri;
    private diffModel: WorkbookDiffModel | null = null;
    private activeSheetKey: string | null = null;
    private isWebviewReady = false;
    private hasPendingRender = false;
    private isReloading = false;
    private hasQueuedReload = false;
    private autoRefreshTimer: ReturnType<typeof setTimeout> | undefined;
    private suppressAutoRefreshUntil = 0;

    private constructor(
        panel: vscode.WebviewPanel,
        extensionUri: vscode.Uri,
        leftFileUri: vscode.Uri,
        rightFileUri: vscode.Uri,
        panelKey: string
    ) {
        this.panel = panel;
        this.extensionUri = extensionUri;
        this.leftFileUri = leftFileUri;
        this.rightFileUri = rightFileUri;
        this.panelKey = panelKey;

        this.panel.webview.html = this.getHtml();
        this.panel.onDidDispose(
            () => {
                XlsxDiffPanel.panels.delete(this.panelKey);
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

    public static async create(
        extensionUri: vscode.Uri,
        leftFileUri: vscode.Uri,
        rightFileUri: vscode.Uri,
        viewColumn: vscode.ViewColumn = vscode.ViewColumn.Active
    ): Promise<void> {
        const panelKey = XlsxDiffPanel.getPanelKey(leftFileUri, rightFileUri);
        const existingPanel = XlsxDiffPanel.panels.get(panelKey);
        if (existingPanel) {
            existingPanel.setFileUris(leftFileUri, rightFileUri);
            existingPanel.panel.reveal(viewColumn, true);
            await existingPanel.enqueueReload();
            return;
        }

        const panel = vscode.window.createWebviewPanel(
            WEBVIEW_TYPE_DIFF_PANEL,
            `${getWorkbookResourceName(leftFileUri)} ↔ ${getWorkbookResourceName(rightFileUri)}`,
            viewColumn,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [extensionUri],
            }
        );

        panel.iconPath = {
            light: vscode.Uri.joinPath(extensionUri, "media", "icon.png"),
            dark: vscode.Uri.joinPath(extensionUri, "media", "icon.png"),
        };

        const instance = new XlsxDiffPanel(
            panel,
            extensionUri,
            leftFileUri,
            rightFileUri,
            panelKey
        );
        XlsxDiffPanel.panels.set(panelKey, instance);
        await instance.enqueueReload();
    }

    public static async refreshAll(): Promise<void> {
        await Promise.all(
            [...XlsxDiffPanel.panels.values()].map((panel) =>
                panel.refreshForDisplayLanguageChange()
            )
        );
    }

    private static getPanelKey(leftFileUri: vscode.Uri, rightFileUri: vscode.Uri): string {
        return [leftFileUri.toString(), rightFileUri.toString()].sort().join("::");
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

    private setFileUris(leftFileUri: vscode.Uri, rightFileUri: vscode.Uri): void {
        this.leftFileUri = leftFileUri;
        this.rightFileUri = rightFileUri;
        this.refreshFileWatchers();
    }

    private refreshFileWatchers(): void {
        this.disposeFileWatchers();

        const watchTargets = new Map<string, vscode.Uri>();
        for (const uri of [this.leftFileUri, this.rightFileUri]) {
            if (uri.scheme !== "file") {
                continue;
            }

            watchTargets.set(uri.toString(), uri);
        }

        for (const uri of watchTargets.values()) {
            const watcher = vscode.workspace.createFileSystemWatcher(
                new vscode.RelativePattern(
                    vscode.Uri.file(path.dirname(uri.fsPath)),
                    escapeWatcherGlobSegment(path.basename(uri.fsPath))
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
    }

    private scheduleAutoRefresh(): void {
        if (Date.now() < this.suppressAutoRefreshUntil) {
            if (this.autoRefreshTimer) {
                clearTimeout(this.autoRefreshTimer);
            }

            const delay = Math.max(0, this.suppressAutoRefreshUntil - Date.now()) + 50;
            this.autoRefreshTimer = setTimeout(() => {
                this.autoRefreshTimer = undefined;
                void this.enqueueReload({ silent: true, clearPendingEdits: true }).catch((error) => {
                    void this.handleError(error);
                });
            }, delay);
            return;
        }

        if (this.autoRefreshTimer) {
            clearTimeout(this.autoRefreshTimer);
        }

        this.autoRefreshTimer = setTimeout(() => {
            this.autoRefreshTimer = undefined;
            void this.enqueueReload({ silent: true, clearPendingEdits: true }).catch((error) => {
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
                await this.enqueueReload({ silent: true });
            }
        }

        if (reloadError) {
            throw reloadError;
        }
    }

    private async refreshForDisplayLanguageChange(): Promise<void> {
        this.isWebviewReady = false;
        this.hasPendingRender = Boolean(this.diffModel);
        this.panel.webview.html = this.getHtml();
        await this.enqueueReload({ silent: true });
    }

    private getHtml(): string {
        const webview = this.panel.webview;
        const nonce = getNonce();
        const webviewStrings = getWebviewStrings();
        const strings = JSON.stringify(webviewStrings).replace(/</g, "\\u003c");
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "panel.js")
        );
        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "diff-panel.css")
        );
        const codiconStyleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, "media", "codicons", "codicon.css")
        );

        return `<!DOCTYPE html>
<html lang="${getHtmlLanguageTag()}">
<head>
	<meta charset="UTF-8" />
	<meta
		http-equiv="Content-Security-Policy"
		content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource};"
	/>
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${codiconStyleUri}" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Diff</title>
</head>
<body>
	<div id="app" class="v2-loading">${webviewStrings.loading}</div>
	<script nonce="${nonce}">window.__XLSX_DIFF_STRINGS__ = ${strings};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
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
                    if (!this.diffModel) {
                        return;
                    }

                    if (!this.diffModel.sheets.some((sheet) => sheet.key === message.sheetKey)) {
                        return;
                    }

                    this.activeSheetKey = message.sheetKey;
                    await this.render();
                    return;
                case "saveEdits": {
                    if (!this.diffModel || message.edits.length === 0) {
                        return;
                    }

                    const leftCellEdits: CellEdit[] = message.edits
                        .filter((edit) => edit.side === "left")
                        .flatMap((edit) => {
                            const sheet = this.diffModel!.sheets.find((item) => item.key === edit.sheetKey);
                            return sheet?.leftSheet
                                ? [
                                      {
                                          sheetName: sheet.leftSheet.name,
                                          rowNumber: edit.rowNumber,
                                          columnNumber: edit.columnNumber,
                                          value: edit.value,
                                      },
                                  ]
                                : [];
                        });

                    const rightCellEdits: CellEdit[] = message.edits
                        .filter((edit) => edit.side === "right")
                        .flatMap((edit) => {
                            const sheet = this.diffModel!.sheets.find((item) => item.key === edit.sheetKey);
                            return sheet?.rightSheet
                                ? [
                                      {
                                          sheetName: sheet.rightSheet.name,
                                          rowNumber: edit.rowNumber,
                                          columnNumber: edit.columnNumber,
                                          value: edit.value,
                                      },
                                  ]
                                : [];
                        });

                    this.suppressAutoRefreshUntil = Date.now() + 2000;

                    await Promise.all([
                        leftCellEdits.length > 0
                            ? writeCellValues(this.leftFileUri, leftCellEdits)
                            : Promise.resolve(),
                        rightCellEdits.length > 0
                            ? writeCellValues(this.rightFileUri, rightCellEdits)
                            : Promise.resolve(),
                    ]);

                    await this.enqueueReload({ silent: true, clearPendingEdits: true });
                    return;
                }
                case "swap": {
                    const previousLeftFileUri = this.leftFileUri;
                    this.setFileUris(this.rightFileUri, previousLeftFileUri);
                    await this.enqueueReload({ clearPendingEdits: true });
                    return;
                }
                case "reload":
                    await this.enqueueReload({ clearPendingEdits: true });
                    return;
            }
        } catch (error) {
            await this.handleError(error);
        }
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

        const [leftWorkbook, rightWorkbook] = await Promise.all([
            loadWorkbookSnapshot(this.leftFileUri),
            loadWorkbookSnapshot(this.rightFileUri),
        ]);

        this.diffModel = buildWorkbookDiff(leftWorkbook, rightWorkbook);
        this.activeSheetKey =
            this.diffModel.sheets.find((sheet) => sheet.key === this.activeSheetKey)?.key ??
            this.diffModel.sheets[0]?.key ??
            null;

        await this.render(undefined, { clearPendingEdits });
    }

    private async render(
        renderModel?: DiffPanelRenderModel,
        { clearPendingEdits = false }: { clearPendingEdits?: boolean } = {}
    ): Promise<void> {
        if (!this.diffModel) {
            return;
        }

        const payload = renderModel ?? createDiffPanelRenderModel(this.diffModel, this.activeSheetKey);
        this.panel.title = payload.title;

        if (!this.isWebviewReady) {
            this.hasPendingRender = true;
            return;
        }

        this.hasPendingRender = false;
        await this.panel.webview.postMessage({
            type: "render",
            payload,
            clearPendingEdits,
        });
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
}
