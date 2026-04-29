/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import { XlsxEditorPanel } from "../webview/editor-panel";

suite("Editor panel row heights", () => {
    test("prompts using workbook row heights directly", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                rowCount: 5,
                rowHeights: {
                    "2": 16,
                },
            },
        };
        const originalShowInputBox = vscode.window.showInputBox;
        let capturedOptions: vscode.InputBoxOptions | undefined;
        let capturedUpdate: { rowNumber: number; height: number | null } | null = null;

        panel.getWorkingWorkbook = () => ({ isReadonly: false });
        panel.getActiveSheetEntry = () => activeEntry;
        panel.setPendingRowHeight = async (rowNumber: number, height: number | null) => {
            capturedUpdate = { rowNumber, height };
        };
        (
            vscode.window as {
                showInputBox: typeof vscode.window.showInputBox;
            }
        ).showInputBox = async (options) => {
            capturedOptions = options;
            return "32";
        };

        try {
            await panel.promptPendingRowHeight(2);

            assert.strictEqual(capturedOptions?.value, "16");
            assert.deepStrictEqual(capturedUpdate, {
                rowNumber: 2,
                height: 32,
            });
        } finally {
            (
                vscode.window as {
                    showInputBox: typeof vscode.window.showInputBox;
                }
            ).showInputBox = originalShowInputBox;
        }
    });

    test("stores direct row height updates in sheet snapshots", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                rowCount: 5,
                rowHeights: {
                    "2": 16,
                },
            },
        };
        let committedOptions: { resetPendingHistory?: boolean } | null = null;
        let syncedSheetKey: string | null = null;

        panel.getWorkingWorkbook = () => ({ isReadonly: false });
        panel.getActiveSheetEntry = () => activeEntry;
        panel.commitStructuralMutation = async (
            mutate: () => void,
            options: { resetPendingHistory?: boolean }
        ) => {
            committedOptions = options;
            mutate();
        };
        panel.syncPendingSheetViewEdit = (sheetKey: string) => {
            syncedSheetKey = sheetKey;
        };

        await panel.setPendingRowHeight(3, 18.13);

        assert.deepStrictEqual(activeEntry.sheet.rowHeights, {
            "2": 16,
            "3": 18.13,
        });
        assert.deepStrictEqual(committedOptions, { resetPendingHistory: false });
        assert.strictEqual(syncedSheetKey, "sheet:0");
    });
});
