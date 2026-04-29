/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { XlsxEditorPanel } from "../webview/editor-panel";

suite("Editor panel column widths", () => {
    test("normalizes direct column width updates before storing them", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                columnCount: 3,
                columnWidths: [8.7109375, null, 12],
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

        await panel.setPendingColumnWidth(2, 15.1239);

        assert.deepStrictEqual(activeEntry.sheet.columnWidths, [8.7109375, 15.125, 12]);
        assert.deepStrictEqual(committedOptions, { resetPendingHistory: false });
        assert.strictEqual(syncedSheetKey, "sheet:0");
    });

    test("skips direct column width updates when the normalized width is unchanged", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                columnCount: 3,
                columnWidths: [8.7109375, 15.125, 12],
            },
        };
        let commitCount = 0;

        panel.getWorkingWorkbook = () => ({ isReadonly: false });
        panel.getActiveSheetEntry = () => activeEntry;
        panel.commitStructuralMutation = async () => {
            commitCount += 1;
        };
        panel.syncPendingSheetViewEdit = () => undefined;

        await panel.setPendingColumnWidth(2, 15.1249);

        assert.strictEqual(commitCount, 0);
        assert.deepStrictEqual(activeEntry.sheet.columnWidths, [8.7109375, 15.125, 12]);
    });

    test("drops trailing default-width placeholders after resetting the last custom column", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                columnCount: 3,
                columnWidths: [8.7109375, 15.125],
            },
        };
        let committedOptions: { resetPendingHistory?: boolean } | null = null;

        panel.getWorkingWorkbook = () => ({ isReadonly: false });
        panel.getActiveSheetEntry = () => activeEntry;
        panel.commitStructuralMutation = async (
            mutate: () => void,
            options: { resetPendingHistory?: boolean }
        ) => {
            committedOptions = options;
            mutate();
        };
        panel.syncPendingSheetViewEdit = () => undefined;

        await panel.setPendingColumnWidth(2, null);

        assert.deepStrictEqual(activeEntry.sheet.columnWidths, [8.7109375]);
        assert.deepStrictEqual(committedOptions, { resetPendingHistory: false });
    });
});
