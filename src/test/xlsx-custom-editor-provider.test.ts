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
        const originalCommitDocumentSave = XlsxEditorPanel.commitDocumentSave;
        const subscription = provider.onDidChangeCustomDocument(() => {
            emittedChangeEvents += 1;
        });

        document.saveTo = async (destination: vscode.Uri): Promise<void> => {
            savedTo = destination.toString();
        };
        XlsxEditorPanel.commitDocumentSave = async (): Promise<void> => {};

        try {
            await provider.saveCustomDocument(document, {} as vscode.CancellationToken);

            assert.strictEqual(savedTo, document.uri.toString());
            assert.strictEqual(document.hasPendingEdits(), false);
            assert.strictEqual(emittedChangeEvents, 0);
        } finally {
            document.saveTo = originalSaveTo;
            XlsxEditorPanel.commitDocumentSave = originalCommitDocumentSave;
            subscription.dispose();
            provider.dispose();
        }
    });
});
