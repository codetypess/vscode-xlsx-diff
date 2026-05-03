import * as vscode from "vscode";
import { WEBVIEW_TYPE_EDITOR_PANEL } from "../../constants";
import { type WorkbookEditState } from "../../core/fastxlsx/write-cell-value";
import { rememberRecentWorkbookResourceUri } from "../../scm/recent-workbook-resource-context";
import { withWorkbookSaveProgress } from "../../workbook/save-progress";
import { readEditorBackupState, writeEditorBackupState } from "./editor-backup-state";
import {
    formatPerfLog,
    logPerf,
    summarizePendingStateForPerf,
    toPerfErrorMessage,
} from "./editor-perf-log";
import { XlsxEditorDocument } from "./xlsx-editor-document";

function logProviderPerf(event: string, details: Record<string, unknown> = {}): void {
    logPerf("provider", event, details);
}

async function getXlsxEditorPanelModule(): Promise<typeof import("./editor-panel")> {
    return import("./editor-panel");
}

export class XlsxCustomEditorProvider
    implements vscode.CustomEditorProvider<XlsxEditorDocument>, vscode.Disposable
{
    private readonly onDidChangeCustomDocumentEmitter = new vscode.EventEmitter<
        vscode.CustomDocumentContentChangeEvent<XlsxEditorDocument>
    >();
    private readonly inFlightSaves = new WeakMap<XlsxEditorDocument, Promise<void>>();

    public static register(context: vscode.ExtensionContext): vscode.Disposable {
        const provider = new XlsxCustomEditorProvider(context.extensionUri);
        context.subscriptions.push(provider);

        return vscode.window.registerCustomEditorProvider(WEBVIEW_TYPE_EDITOR_PANEL, provider, {
            webviewOptions: {
                retainContextWhenHidden: true,
            },
        });
    }

    private constructor(private readonly extensionUri: vscode.Uri) {}

    public readonly onDidChangeCustomDocument = this.onDidChangeCustomDocumentEmitter.event;

    public dispose(): void {
        this.onDidChangeCustomDocumentEmitter.dispose();
    }

    public async openCustomDocument(
        uri: vscode.Uri,
        openContext: vscode.CustomDocumentOpenContext
    ): Promise<XlsxEditorDocument> {
        rememberRecentWorkbookResourceUri(uri, "customEditorDocument");
        const backupUri = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : undefined;
        if (!backupUri) {
            return new XlsxEditorDocument(uri);
        }

        const backupState = await readEditorBackupState(backupUri).catch(() => null);
        if (backupState) {
            return new XlsxEditorDocument(uri, {
                backupState,
            });
        }

        return new XlsxEditorDocument(uri, {
            backupUri,
        });
    }

    public async resolveCustomEditor(
        document: XlsxEditorDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        const { XlsxEditorPanel } = await getXlsxEditorPanelModule();
        await XlsxEditorPanel.resolveCustomEditor(this.extensionUri, document, webviewPanel, {
            onPendingStateChanged: async (state: WorkbookEditState) => {
                const startedAt = performance.now();
                if (!document.replacePendingState(state)) {
                    console.info(
                        formatPerfLog("provider", "onPendingStateChanged:no-op", {
                            durationMs: Number((performance.now() - startedAt).toFixed(2)),
                            ...summarizePendingStateForPerf(state),
                        })
                    );
                    return;
                }

                console.info(
                    formatPerfLog("provider", "onPendingStateChanged:updated", {
                        durationMs: Number((performance.now() - startedAt).toFixed(2)),
                        ...summarizePendingStateForPerf(state),
                    })
                );
                this.onDidChangeCustomDocumentEmitter.fire({ document });
            },
            onRequestSave: async () => {
                await vscode.commands.executeCommand("workbench.action.files.save");
            },
            onRequestRevert: async () => {
                await vscode.commands.executeCommand("workbench.action.files.revert");
            },
        });

        if (document.consumeInitialDirtyState()) {
            this.onDidChangeCustomDocumentEmitter.fire({ document });
        }
    }

    public async saveCustomDocument(
        document: XlsxEditorDocument,
        _cancellation: vscode.CancellationToken
    ): Promise<void> {
        await this.saveDocument(document);
    }

    public async saveCustomDocumentAs(
        document: XlsxEditorDocument,
        destination: vscode.Uri,
        _cancellation: vscode.CancellationToken
    ): Promise<void> {
        const startedAt = performance.now();
        const pendingState = document.getPendingState();
        logProviderPerf("saveDocumentAs:start", {
            sourcePath: document.getReadUri().fsPath,
            destinationPath: destination.fsPath,
            ...summarizePendingStateForPerf(pendingState),
        });

        try {
            await document.saveTo(destination);
            logProviderPerf("saveDocumentAs:done", {
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                sourcePath: document.getReadUri().fsPath,
                destinationPath: destination.fsPath,
                ...summarizePendingStateForPerf(pendingState),
            });
        } catch (error) {
            logProviderPerf("saveDocumentAs:error", {
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                sourcePath: document.getReadUri().fsPath,
                destinationPath: destination.fsPath,
                errorMessage: toPerfErrorMessage(error),
                ...summarizePendingStateForPerf(pendingState),
            });
            throw error;
        }
    }

    public async revertCustomDocument(
        document: XlsxEditorDocument,
        _cancellation: vscode.CancellationToken
    ): Promise<void> {
        await this.revertDocument(document);
    }

    public async backupCustomDocument(
        document: XlsxEditorDocument,
        context: vscode.CustomDocumentBackupContext,
        _cancellation: vscode.CancellationToken
    ): Promise<vscode.CustomDocumentBackup> {
        await writeEditorBackupState(context.destination, document.getPendingState());

        return {
            id: context.destination.toString(),
            delete: () => {
                void vscode.workspace.fs.delete(context.destination).then(
                    () => undefined,
                    () => undefined
                );
            },
        };
    }

    private async revertDocument(document: XlsxEditorDocument): Promise<void> {
        document.markReverted();
        const { XlsxEditorPanel } = await getXlsxEditorPanelModule();
        await XlsxEditorPanel.refreshDocument(document, {
            clearPendingEdits: true,
        });
    }

    private async saveDocument(document: XlsxEditorDocument): Promise<void> {
        const inFlightSave = this.inFlightSaves.get(document);
        if (inFlightSave) {
            logProviderPerf("saveDocument:reuseInFlight", {
                documentPath: document.uri.fsPath,
            });
            return inFlightSave;
        }

        const savePromise = this.runSaveDocument(document).finally(() => {
            this.inFlightSaves.delete(document);
        });
        this.inFlightSaves.set(document, savePromise);
        return savePromise;
    }

    private async runSaveDocument(document: XlsxEditorDocument): Promise<void> {
        const confirmStartedAt = performance.now();
        logProviderPerf("saveDocument:confirm:start", {
            documentPath: document.uri.fsPath,
        });
        const { XlsxEditorPanel } = await getXlsxEditorPanelModule();
        const confirmed = await XlsxEditorPanel.confirmDocumentSave(document);
        logProviderPerf("saveDocument:confirm:done", {
            durationMs: Number((performance.now() - confirmStartedAt).toFixed(2)),
            documentPath: document.uri.fsPath,
            confirmed,
        });
        if (!confirmed) {
            return;
        }

        const pendingState = document.getPendingState();
        const startedAt = performance.now();
        logProviderPerf("saveDocument:start", {
            documentPath: document.uri.fsPath,
            ...summarizePendingStateForPerf(pendingState),
        });
        await XlsxEditorPanel.beginDocumentSave(document);

        try {
            await withWorkbookSaveProgress(
                () => document.saveTo(document.uri),
                { workbookUri: document.uri }
            );
            document.markSaved();
            // Saving already clears VS Code's dirty state for the custom editor.
            // Emitting another content-change event here marks the tab dirty again.
            await XlsxEditorPanel.commitDocumentSave(document);
            logProviderPerf("saveDocument:done", {
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                documentPath: document.uri.fsPath,
                ...summarizePendingStateForPerf(pendingState),
            });
        } catch (error) {
            logProviderPerf("saveDocument:error", {
                durationMs: Number((performance.now() - startedAt).toFixed(2)),
                documentPath: document.uri.fsPath,
                errorMessage: toPerfErrorMessage(error),
                ...summarizePendingStateForPerf(pendingState),
            });
            await XlsxEditorPanel.failDocumentSave(document);
            throw error;
        }
    }
}
