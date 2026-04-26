/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import { XlsxEditorPanel } from "../webview/editor-panel";
import { XlsxCustomEditorProvider } from "../webview/xlsx-custom-editor-provider";
import { XlsxEditorDocument } from "../webview/xlsx-editor-document";

suite("Xlsx custom editor provider", () => {
    test("save clears pending edits without re-emitting a dirty change event", async () => {
        const ProviderConstructor = XlsxCustomEditorProvider as unknown as {
            new (extensionUri: vscode.Uri): XlsxCustomEditorProvider;
        };
        const provider = new ProviderConstructor(vscode.Uri.file("/tmp"));
        const document = new XlsxEditorDocument(vscode.Uri.file("/tmp/editor.xlsx"));
        let emittedChangeEvents = 0;
        let savedTo = "";

        document.replacePendingState({
            cellEdits: [
                {
                    sheetName: "Sheet1",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "updated",
                },
            ],
            sheetEdits: [],
            viewEdits: [],
        });

        const originalSaveTo = document.saveTo.bind(document);
        const originalWithProgress = vscode.window.withProgress;
        const originalConfirmDocumentSave = XlsxEditorPanel.confirmDocumentSave;
        const originalBeginDocumentSave = XlsxEditorPanel.beginDocumentSave;
        const originalCommitDocumentSave = XlsxEditorPanel.commitDocumentSave;
        const originalFailDocumentSave = XlsxEditorPanel.failDocumentSave;
        const subscription = provider.onDidChangeCustomDocument(() => {
            emittedChangeEvents += 1;
        });

        document.saveTo = async (destination: vscode.Uri): Promise<void> => {
            savedTo = destination.toString();
        };
        (
            vscode.window as {
                withProgress: typeof vscode.window.withProgress;
            }
        ).withProgress = async (_options, task) =>
            task({ report: () => undefined }, {} as vscode.CancellationToken);
        XlsxEditorPanel.confirmDocumentSave = async (): Promise<boolean> => true;
        XlsxEditorPanel.beginDocumentSave = async (): Promise<void> => {};
        XlsxEditorPanel.commitDocumentSave = async (): Promise<void> => {};
        XlsxEditorPanel.failDocumentSave = async (): Promise<void> => {};

        try {
            await provider.saveCustomDocument(document, {} as vscode.CancellationToken);

            assert.strictEqual(savedTo, document.uri.toString());
            assert.strictEqual(document.hasPendingEdits(), false);
            assert.strictEqual(emittedChangeEvents, 0);
        } finally {
            document.saveTo = originalSaveTo;
            (
                vscode.window as {
                    withProgress: typeof vscode.window.withProgress;
                }
            ).withProgress = originalWithProgress;
            XlsxEditorPanel.confirmDocumentSave = originalConfirmDocumentSave;
            XlsxEditorPanel.beginDocumentSave = originalBeginDocumentSave;
            XlsxEditorPanel.commitDocumentSave = originalCommitDocumentSave;
            XlsxEditorPanel.failDocumentSave = originalFailDocumentSave;
            subscription.dispose();
            provider.dispose();
        }
    });

    test("save is skipped when the editor panel declines confirmation", async () => {
        const ProviderConstructor = XlsxCustomEditorProvider as unknown as {
            new (extensionUri: vscode.Uri): XlsxCustomEditorProvider;
        };
        const provider = new ProviderConstructor(vscode.Uri.file("/tmp"));
        const document = new XlsxEditorDocument(vscode.Uri.file("/tmp/editor.xlsx"));
        const originalSaveTo = document.saveTo.bind(document);
        const originalConfirmDocumentSave = XlsxEditorPanel.confirmDocumentSave;
        let attemptedSave = false;

        document.saveTo = async (): Promise<void> => {
            attemptedSave = true;
        };
        XlsxEditorPanel.confirmDocumentSave = async (): Promise<boolean> => false;

        try {
            await provider.saveCustomDocument(document, {} as vscode.CancellationToken);

            assert.strictEqual(attemptedSave, false);
        } finally {
            document.saveTo = originalSaveTo;
            XlsxEditorPanel.confirmDocumentSave = originalConfirmDocumentSave;
            provider.dispose();
        }
    });

    test("toolbar save routes through VS Code's save command", async () => {
        const ProviderConstructor = XlsxCustomEditorProvider as unknown as {
            new (extensionUri: vscode.Uri): XlsxCustomEditorProvider;
        };
        const provider = new ProviderConstructor(vscode.Uri.file("/tmp"));
        const document = new XlsxEditorDocument(vscode.Uri.file("/tmp/editor.xlsx"));
        const originalResolveCustomEditor = XlsxEditorPanel.resolveCustomEditor;
        const originalExecuteCommand = vscode.commands.executeCommand;
        let capturedController:
            | {
                  onRequestSave(): void | Promise<void>;
              }
            | undefined;
        const executedCommands: string[] = [];

        XlsxEditorPanel.resolveCustomEditor = async (
            _extensionUri,
            _document,
            _panel,
            controller
        ): Promise<void> => {
            capturedController = controller;
        };
        (
            vscode.commands as {
                executeCommand: typeof vscode.commands.executeCommand;
            }
        ).executeCommand = <T = unknown>(command: string): Thenable<T> => {
            executedCommands.push(command);
            return Promise.resolve(undefined as T);
        };

        try {
            await provider.resolveCustomEditor(
                document,
                {} as vscode.WebviewPanel,
                {} as vscode.CancellationToken
            );

            assert.ok(capturedController);
            await capturedController.onRequestSave();
            assert.deepStrictEqual(executedCommands, ["workbench.action.files.save"]);
        } finally {
            XlsxEditorPanel.resolveCustomEditor = originalResolveCustomEditor;
            (
                vscode.commands as {
                    executeCommand: typeof vscode.commands.executeCommand;
                }
            ).executeCommand = originalExecuteCommand;
            provider.dispose();
        }
    });

    test("concurrent save requests share one in-flight save", async () => {
        const ProviderConstructor = XlsxCustomEditorProvider as unknown as {
            new (extensionUri: vscode.Uri): XlsxCustomEditorProvider;
        };
        const provider = new ProviderConstructor(vscode.Uri.file("/tmp"));
        const document = new XlsxEditorDocument(vscode.Uri.file("/tmp/editor.xlsx"));
        const originalSaveTo = document.saveTo.bind(document);
        const originalWithProgress = vscode.window.withProgress;
        const originalConfirmDocumentSave = XlsxEditorPanel.confirmDocumentSave;
        const originalBeginDocumentSave = XlsxEditorPanel.beginDocumentSave;
        const originalCommitDocumentSave = XlsxEditorPanel.commitDocumentSave;
        const originalFailDocumentSave = XlsxEditorPanel.failDocumentSave;
        let resolveSave: (() => void) | undefined;
        let saveCalls = 0;

        document.replacePendingState({
            cellEdits: [
                {
                    sheetName: "Sheet1",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "updated",
                },
            ],
            sheetEdits: [],
            viewEdits: [],
        });

        document.saveTo = async (): Promise<void> => {
            saveCalls += 1;
            await new Promise<void>((resolve) => {
                resolveSave = resolve;
            });
        };
        (
            vscode.window as {
                withProgress: typeof vscode.window.withProgress;
            }
        ).withProgress = async (_options, task) =>
            task({ report: () => undefined }, {} as vscode.CancellationToken);
        XlsxEditorPanel.confirmDocumentSave = async (): Promise<boolean> => true;
        XlsxEditorPanel.beginDocumentSave = async (): Promise<void> => {};
        XlsxEditorPanel.commitDocumentSave = async (): Promise<void> => {};
        XlsxEditorPanel.failDocumentSave = async (): Promise<void> => {};

        try {
            const firstSave = provider.saveCustomDocument(document, {} as vscode.CancellationToken);
            const secondSave = provider.saveCustomDocument(document, {} as vscode.CancellationToken);

            await Promise.resolve();
            assert.strictEqual(saveCalls, 1);

            resolveSave?.();
            await Promise.all([firstSave, secondSave]);
            assert.strictEqual(saveCalls, 1);
        } finally {
            document.saveTo = originalSaveTo;
            (
                vscode.window as {
                    withProgress: typeof vscode.window.withProgress;
                }
            ).withProgress = originalWithProgress;
            XlsxEditorPanel.confirmDocumentSave = originalConfirmDocumentSave;
            XlsxEditorPanel.beginDocumentSave = originalBeginDocumentSave;
            XlsxEditorPanel.commitDocumentSave = originalCommitDocumentSave;
            XlsxEditorPanel.failDocumentSave = originalFailDocumentSave;
            provider.dispose();
        }
    });
});
