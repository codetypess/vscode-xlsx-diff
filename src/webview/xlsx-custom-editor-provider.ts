import * as vscode from "vscode";
import { WEBVIEW_TYPE_EDITOR_PANEL } from "../constants";
import { type WorkbookEditState } from "../core/fastxlsx/write-cell-value";
import { rememberRecentWorkbookResourceUri } from "../scm/recent-workbook-resource-context";
import { withWorkbookSaveProgress } from "../workbook/save-progress";
import { readEditorBackupState, writeEditorBackupState } from "./editor-backup-state";
import { XlsxEditorPanel } from "./editor-panel";
import { XlsxEditorDocument } from "./xlsx-editor-document";

export class XlsxCustomEditorProvider
    implements vscode.CustomEditorProvider<XlsxEditorDocument>, vscode.Disposable
{
    private readonly onDidChangeCustomDocumentEmitter = new vscode.EventEmitter<
        vscode.CustomDocumentContentChangeEvent<XlsxEditorDocument>
    >();

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
        await XlsxEditorPanel.resolveCustomEditor(this.extensionUri, document, webviewPanel, {
            onPendingStateChanged: async (state: WorkbookEditState) => {
                if (!document.replacePendingState(state)) {
                    return;
                }

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
        await document.saveTo(destination);
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
        await XlsxEditorPanel.refreshDocument(document, {
            clearPendingEdits: true,
        });
    }

    private async saveDocument(document: XlsxEditorDocument): Promise<void> {
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
        } catch (error) {
            await XlsxEditorPanel.failDocumentSave(document);
            throw error;
        }
    }
}
