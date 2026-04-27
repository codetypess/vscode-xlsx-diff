/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    setActiveEditorSheet,
    setSelectedEditorCell,
} from "../webview/editor-panel/editor-render-model";

function createCell(rowNumber: number, columnNumber: number, displayValue: string): CellSnapshot {
    return {
        key: `${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        address: `R${rowNumber}C${columnNumber}`,
        displayValue,
        formula: null,
        styleId: null,
    };
}

function createSheet(
    name: string,
    cells: CellSnapshot[],
    rowCount = 1,
    columnCount = 1,
    mergedRanges: string[] = [],
    freezePane: SheetSnapshot["freezePane"] = null
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        visibility: "visible",
        mergedRanges,
        freezePane,
        cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
        signature: `${name}-signature`,
    };
}

function createWorkbook(
    overrides: Partial<WorkbookSnapshot>,
    sheets: SheetSnapshot[]
): WorkbookSnapshot {
    return {
        filePath: "/tmp/editor.xlsx",
        fileName: "editor.xlsx",
        fileSize: 128,
        modifiedTime: new Date("2026-04-18T06:51:00.000Z").toISOString(),
        sheets,
        ...overrides,
    };
}

suite("Editor render model", () => {
    test("keeps title, readonly state, and edit flags aligned", () => {
        const workbook = createWorkbook(
            {
                detailLabel: "Commit",
                detailValue: "d4ce7e0",
                titleDetail: "d4ce7e0",
                modifiedTimeLabel: "Apr 18, 2026, 6:51 AM",
                isReadonly: true,
            },
            [createSheet("Sheet1", [createCell(1, 1, "value")])]
        );

        const renderModel = createEditorRenderModel(
            workbook,
            createInitialEditorPanelState(workbook),
            { hasPendingEdits: true }
        );

        assert.strictEqual(renderModel.title, "editor.xlsx (d4ce7e0)");
        assert.strictEqual(renderModel.hasPendingEdits, true);
        assert.strictEqual(renderModel.canEdit, false);
    });

    test("keeps the selected sparse cell in the render model", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(205, 1, "tail")], 205, 1),
        ]);

        const state = setSelectedEditorCell(
            workbook,
            createInitialEditorPanelState(workbook),
            205,
            1
        );
        const renderModel = createEditorRenderModel(workbook, state);

        assert.strictEqual(renderModel.selection?.address, "R205C1");
        assert.strictEqual(renderModel.activeSheet.rowCount, 205);
        assert.strictEqual(renderModel.activeSheet.cells["205:1"]?.displayValue, "tail");
    });

    test("surfaces full active sheet columns and sparse cells", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(2, 2, "value"), createCell(4, 3, "tail")], 400, 3),
        ]);
        const renderModel = createEditorRenderModel(
            workbook,
            createInitialEditorPanelState(workbook)
        );

        assert.deepStrictEqual(renderModel.activeSheet.columns, ["A", "B", "C"]);
        assert.strictEqual(renderModel.activeSheet.cells["2:2"]?.displayValue, "value");
        assert.strictEqual(renderModel.activeSheet.cells["4:3"]?.displayValue, "tail");
    });

    test("switches sheets and preserves sparse cells as blanks", () => {
        const workbook = createWorkbook({}, [
            createSheet("First", [createCell(1, 1, "left")], 1, 2),
            createSheet("Second", [createCell(2, 2, "value")], 2, 2, ["A1:B2"]),
        ]);
        const initialState = createInitialEditorPanelState(workbook);
        const targetSheetKey = createEditorRenderModel(workbook, initialState).sheets[1].key;
        const state = setActiveEditorSheet(workbook, initialState, targetSheetKey);
        const renderModel = createEditorRenderModel(workbook, state);

        assert.strictEqual(renderModel.activeSheet.key, targetSheetKey);
        assert.strictEqual(renderModel.sheets[1]?.label, "Second");
        assert.strictEqual(renderModel.sheets[1]?.isActive, true);
        assert.deepStrictEqual(renderModel.activeSheet.columns, ["A", "B"]);
        assert.strictEqual(renderModel.activeSheet.cells["2:2"]?.displayValue, "value");
    });

    test("surfaces freeze pane state for the active sheet", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(3, 3, "value")], 10, 10, [], {
                columnCount: 1,
                rowCount: 2,
                topLeftCell: "B3",
                activePane: "bottomRight",
            }),
        ]);

        const renderModel = createEditorRenderModel(
            workbook,
            createInitialEditorPanelState(workbook)
        );

        assert.deepStrictEqual(renderModel.activeSheet.freezePane, {
            columnCount: 1,
            rowCount: 2,
            topLeftCell: "B3",
            activePane: "bottomRight",
        });
        assert.strictEqual(renderModel.activeSheet.rowCount, 10);
        assert.strictEqual(renderModel.activeSheet.columnCount, 10);
    });

    test("keeps the full active sheet cell map available for large locked sheets", () => {
        const workbook = createWorkbook({}, [
            createSheet(
                "Sheet1",
                [createCell(240, 2, "near-tail"), createCell(360, 2, "far-tail")],
                400,
                4,
                [],
                {
                    columnCount: 1,
                    rowCount: 2,
                    topLeftCell: "B3",
                    activePane: "bottomRight",
                }
            ),
        ]);

        const renderModel = createEditorRenderModel(
            workbook,
            createInitialEditorPanelState(workbook)
        );

        assert.strictEqual(renderModel.activeSheet.cells["240:2"]?.displayValue, "near-tail");
        assert.strictEqual(renderModel.activeSheet.cells["360:2"]?.displayValue, "far-tail");
        assert.deepStrictEqual(renderModel.activeSheet.freezePane, {
            columnCount: 1,
            rowCount: 2,
            topLeftCell: "B3",
            activePane: "bottomRight",
        });
    });

    test("loads a sparse tail cell after selection moves the viewport window", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(240, 2, "tail")], 400, 4, [], {
                columnCount: 1,
                rowCount: 2,
                topLeftCell: "B3",
                activePane: "bottomRight",
            }),
        ]);

        const selectedState = setSelectedEditorCell(
            workbook,
            createInitialEditorPanelState(workbook),
            240,
            2
        );
        const renderModel = createEditorRenderModel(workbook, selectedState);

        assert.strictEqual(renderModel.selection?.address, "R240C2");
        assert.strictEqual(renderModel.activeSheet.cells["240:2"]?.displayValue, "tail");
    });

    test("keeps frozen-row selections in the render model", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(240, 2, "tail")], 400, 4, [], {
                columnCount: 1,
                rowCount: 2,
                topLeftCell: "B3",
                activePane: "bottomRight",
            }),
        ]);

        const selectedState = setSelectedEditorCell(
            workbook,
            createInitialEditorPanelState(workbook),
            1,
            1
        );
        const renderModel = createEditorRenderModel(workbook, selectedState);

        assert.strictEqual(renderModel.selection?.address, "A1");
        assert.deepStrictEqual(renderModel.activeSheet.freezePane, {
            columnCount: 1,
            rowCount: 2,
            topLeftCell: "B3",
            activePane: "bottomRight",
        });
    });

    test("preserves selections beyond the current used range", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(1, 1, "head")], 5, 3),
        ]);

        const selectedState = setSelectedEditorCell(
            workbook,
            createInitialEditorPanelState(workbook),
            12,
            8
        );
        const renderModel = createEditorRenderModel(workbook, selectedState);

        assert.deepStrictEqual(renderModel.selection, {
            rowNumber: 12,
            columnNumber: 8,
            key: "12:8",
            address: "H12",
            value: "",
            formula: null,
            isPresent: false,
        });
    });
});
