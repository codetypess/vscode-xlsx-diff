/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import { resolveWorkbookUriForOpenEditor } from "../commands/open-editor";

suite("Open editor command", () => {
    test("prefers the command resource when present", () => {
        const commandUri = vscode.Uri.file("/tmp/command.xlsx");
        const activeUri = vscode.Uri.file("/tmp/active.xlsx");

        const resolved = resolveWorkbookUriForOpenEditor(commandUri, activeUri);

        assert.strictEqual(resolved?.fsPath, commandUri.fsPath);
    });

    test("falls back to the active workbook when command resource is absent", () => {
        const activeUri = vscode.Uri.file("/tmp/active.xlsx");

        const resolved = resolveWorkbookUriForOpenEditor(undefined, activeUri);

        assert.strictEqual(resolved?.fsPath, activeUri.fsPath);
    });

    test("accepts explorer resource wrappers", () => {
        const commandUri = vscode.Uri.file("/tmp/wrapped.xlsx");

        const resolved = resolveWorkbookUriForOpenEditor({ resourceUri: commandUri });

        assert.strictEqual(resolved?.fsPath, commandUri.fsPath);
    });
});
