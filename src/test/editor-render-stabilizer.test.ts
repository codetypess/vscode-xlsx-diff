/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { EditorRenderModel, SheetFreezePaneSnapshot } from "../core/model/types";
import { stabilizeIncomingRenderModel } from "../webview/editor-panel/editor-render-stabilizer";

function createRenderModel(
    value: string,
    {
        sheetKey = "sheet:0",
        rowCount = 5,
        columnCount = 4,
        columns = ["A", "B", "C", "D"],
        freezePane = null,
    }: {
        sheetKey?: string;
        rowCount?: number;
        columnCount?: number;
        columns?: string[];
        freezePane?: SheetFreezePaneSnapshot | null;
    } = {}
): EditorRenderModel {
    return {
        title: "Workbook",
        hasPendingEdits: false,
        canEdit: true,
        canUndoStructuralEdits: false,
        canRedoStructuralEdits: false,
        sheets: [{ key: sheetKey, label: "Sheet1", isActive: true }],
        selection: null,
        activeSheet: {
            key: sheetKey,
            rowCount,
            columnCount,
            columns,
            freezePane,
            autoFilter: null,
            cells: {
                "1:1": {
                    key: "1:1",
                    rowNumber: 1,
                    columnNumber: 1,
                    address: "A1",
                    displayValue: value,
                    formula: null,
                    styleId: null,
                },
            },
        },
    };
}

suite("Editor render stabilizer", () => {
    test("prefers the fresh payload when cell reuse is disabled", () => {
        const previousModel = createRenderModel("old");
        const nextModel = createRenderModel("new");

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: false,
        });

        assert.strictEqual(stabilized, nextModel);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "new");
    });

    test("can reuse active sheet cells when explicitly allowed", () => {
        const previousModel = createRenderModel("old");
        const nextModel = createRenderModel("new");

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        assert.notStrictEqual(stabilized, nextModel);
        assert.strictEqual(stabilized.activeSheet.cells, previousModel.activeSheet.cells);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "old");
    });

    test("does not reuse active sheet data when the active sheet shape changes", () => {
        const previousModel = createRenderModel("old");
        const nextModel = createRenderModel("new", { rowCount: 6 });

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        assert.strictEqual(stabilized, nextModel);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "new");
    });
});
