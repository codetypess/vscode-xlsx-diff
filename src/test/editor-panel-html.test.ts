/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import { XlsxEditorPanel } from "../webview/editor-panel";

suite("Editor panel html", () => {
    test("injects the debug flag into the webview bootstrap script", () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const originalDebugMode = (XlsxEditorPanel as any).isDebugMode;

        panel.extensionUri = vscode.Uri.file("/tmp");
        panel.panel = {
            webview: {
                asWebviewUri: (uri: vscode.Uri) => uri,
            },
        };

        try {
            XlsxEditorPanel.setDebugMode(true);

            const html = panel.getHtml();

            assert.match(html, /window\.__XLSX_EDITOR_DEBUG__ = true/);
        } finally {
            XlsxEditorPanel.setDebugMode(originalDebugMode);
        }
    });
});
