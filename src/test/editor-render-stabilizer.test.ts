/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type {
    EditorRenderModel,
    EditorRenderPayload,
    SheetFreezePaneSnapshot,
} from "../core/model/types";
import { stabilizeIncomingRenderModel } from "../webview/editor-panel/editor-render-stabilizer";

function createRenderModel(
    value: string,
    {
        sheetKey = "sheet:0",
        rowCount = 5,
        columnCount = 4,
        columns = ["A", "B", "C", "D"],
        freezePane = null,
        rowHeights = {},
    }: {
        sheetKey?: string;
        rowCount?: number;
        columnCount?: number;
        columns?: string[];
        freezePane?: SheetFreezePaneSnapshot | null;
        rowHeights?: Record<string, number | null>;
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
            columnWidths: [],
            rowHeights,
            cellAlignments: {},
            rowAlignments: {},
            columnAlignments: {},
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

function omitActiveSheetFields(
    model: EditorRenderModel,
    fields: Array<keyof EditorRenderModel["activeSheet"]>
): EditorRenderPayload {
    const activeSheet = { ...model.activeSheet } as Partial<EditorRenderModel["activeSheet"]>;
    for (const field of fields) {
        delete activeSheet[field];
    }

    return {
        ...model,
        activeSheet: activeSheet as EditorRenderPayload["activeSheet"],
    };
}

suite("Editor render stabilizer", () => {
    test("prefers the fresh payload when cell reuse is disabled", () => {
        const previousModel = createRenderModel("old");
        const nextModel = createRenderModel("new");

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: false,
        });

        assert.notStrictEqual(stabilized, previousModel);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "new");
    });

    test("can reuse active sheet cells when explicitly allowed", () => {
        const previousModel = createRenderModel("old");
        const nextModel = omitActiveSheetFields(createRenderModel("new"), ["cells"]);

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        assert.notStrictEqual(stabilized, nextModel);
        assert.strictEqual(stabilized.activeSheet.cells, previousModel.activeSheet.cells);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "old");
    });

    test("can reuse omitted metadata fields when explicitly allowed", () => {
        const previousModel = createRenderModel("old", {
            rowHeights: {
                "2": 42,
            },
        });
        const nextModel = omitActiveSheetFields(createRenderModel("old"), [
            "columns",
            "columnWidths",
            "rowHeights",
            "cellAlignments",
            "rowAlignments",
            "columnAlignments",
        ]);

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        assert.strictEqual(stabilized.activeSheet.columns, previousModel.activeSheet.columns);
        assert.strictEqual(
            stabilized.activeSheet.rowHeights,
            previousModel.activeSheet.rowHeights
        );
        assert.strictEqual(
            stabilized.activeSheet.cellAlignments,
            previousModel.activeSheet.cellAlignments
        );
    });

    test("does not reuse active sheet data when the active sheet shape changes", () => {
        const previousModel = createRenderModel("old");
        const nextModel = createRenderModel("new", { rowCount: 6 });

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        assert.notStrictEqual(stabilized, previousModel);
        assert.strictEqual(stabilized.activeSheet.cells["1:1"]?.displayValue, "new");
    });

    test("applies cell alignment patches by dirty keys when active sheet data is reused", () => {
        const previousModel = createRenderModel("old");
        previousModel.activeSheet.cellAlignments = {
            "1:1": { horizontal: "center" },
            "2:2": { vertical: "bottom" },
        };
        const nextModel = omitActiveSheetFields(createRenderModel("old"), [
            "columns",
            "columnWidths",
            "rowHeights",
            "cells",
            "rowAlignments",
            "columnAlignments",
        ]);
        nextModel.activeSheet.cellAlignments = {};
        nextModel.activeSheet.cellAlignmentDirtyKeys = ["1:1"];

        const stabilized = stabilizeIncomingRenderModel(previousModel, nextModel, {
            canReuseActiveSheetData: true,
        });

        const cellAlignments = stabilized.activeSheet.cellAlignments;
        assert.ok(cellAlignments);
        assert.strictEqual(cellAlignments["1:1"], undefined);
        assert.deepStrictEqual(cellAlignments["2:2"], {
            vertical: "bottom",
        });
    });
});
