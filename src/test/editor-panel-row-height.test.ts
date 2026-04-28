/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { XlsxEditorPanel } from "../webview/editor-panel";

suite("Editor panel row heights", () => {
    test("normalizes direct row height updates before storing them", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                rowCount: 3,
                rowHeights: { "1": 15, "3": 20 },
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

        await panel.setPendingRowHeight(2, 18.126);

        assert.deepStrictEqual(activeEntry.sheet.rowHeights, { "1": 15, "2": 18.13, "3": 20 });
        assert.deepStrictEqual(committedOptions, { resetPendingHistory: false });
        assert.strictEqual(syncedSheetKey, "sheet:0");
    });

    test("skips direct row height updates when the normalized height is unchanged", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: {
                name: "Sheet1",
                rowCount: 3,
                rowHeights: { "2": 18.13 },
            },
        };
        let commitCount = 0;

        panel.getWorkingWorkbook = () => ({ isReadonly: false });
        panel.getActiveSheetEntry = () => activeEntry;
        panel.commitStructuralMutation = async () => {
            commitCount += 1;
        };
        panel.syncPendingSheetViewEdit = () => undefined;

        await panel.setPendingRowHeight(2, 18.129);

        assert.strictEqual(commitCount, 0);
        assert.deepStrictEqual(activeEntry.sheet.rowHeights, { "2": 18.13 });
    });
});
