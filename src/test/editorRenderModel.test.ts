/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    moveEditorPageCursor,
    setActiveEditorSheet,
    setSelectedEditorCell,
    setEditorViewportStartRow,
} from "../webview/editorRenderModel";

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
    freezePane: SheetSnapshot["freezePane"] = null,
    dimensions: {
        rowHeights?: Record<number, number>;
        columnWidths?: Record<number, number>;
    } = {}
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        mergedRanges,
        freezePane,
        rowHeights: { ...(dimensions.rowHeights ?? {}) },
        columnWidths: { ...(dimensions.columnWidths ?? {}) },
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
    test("keeps file detail, readonly state, and save flags aligned", () => {
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
        assert.deepStrictEqual(renderModel.file, {
            fileName: "editor.xlsx",
            filePath: "/tmp/editor.xlsx",
            fileSizeLabel: "128 B",
            detailLabel: "Commit",
            detailValue: "d4ce7e0",
            modifiedTimeLabel: "Apr 18, 2026, 6:51 AM",
            isReadonly: true,
        });
        assert.strictEqual(renderModel.hasPendingEdits, true);
        assert.strictEqual(renderModel.canEdit, false);
        assert.strictEqual(renderModel.canSave, false);
    });

    test("moves selection to the page that contains the clicked cell", () => {
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

        assert.strictEqual(renderModel.page.currentPage, 2);
        assert.strictEqual(renderModel.selection?.address, "R205C1");
        assert.strictEqual(renderModel.page.rows[0].rowNumber, 201);
        assert.strictEqual(renderModel.page.rows[4].cells[0].isSelected, true);
    });

    test("renders a sliding row window when viewport start row changes", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(240, 1, "tail")], 400, 1),
        ]);

        const state = setEditorViewportStartRow(
            workbook,
            createInitialEditorPanelState(workbook),
            121
        );
        const renderModel = createEditorRenderModel(workbook, state);

        assert.strictEqual(renderModel.page.startRow, 121);
        assert.strictEqual(renderModel.page.rows[0]?.rowNumber, 121);
        assert.strictEqual(renderModel.page.endRow, 320);
        assert.strictEqual(renderModel.page.rows[renderModel.page.rows.length - 1]?.rowNumber, 320);
    });

    test("moves page navigation across sheet boundaries", () => {
        const workbook = createWorkbook({}, [
            createSheet("First", [createCell(205, 1, "tail")], 205, 1),
            createSheet("Second", [createCell(1, 1, "head")], 1, 1),
        ]);

        const secondPageState = setSelectedEditorCell(
            workbook,
            createInitialEditorPanelState(workbook),
            205,
            1
        );
        const crossSheetState = moveEditorPageCursor(workbook, secondPageState, 1);
        const renderModel = createEditorRenderModel(workbook, crossSheetState);

        assert.strictEqual(renderModel.activeSheet.label, "Second");
        assert.strictEqual(renderModel.page.currentPage, 1);
        assert.strictEqual(renderModel.selection?.address, "R1C1");
        assert.strictEqual(renderModel.canPrevPage, true);
        assert.strictEqual(renderModel.canNextPage, false);
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

        assert.strictEqual(renderModel.activeSheet.label, "Second");
        assert.strictEqual(renderModel.activeSheet.hasMergedRanges, true);
        assert.strictEqual(renderModel.page.rows[0].cells[0].value, "");
        assert.strictEqual(renderModel.page.rows[1].cells[1].value, "value");
        assert.strictEqual(renderModel.summary.totalNonEmptyCells, 2);
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
        assert.deepStrictEqual(
            renderModel.page.frozenRows.map((row) => row.rowNumber),
            [1, 2]
        );
    });

    test("surfaces explicit row heights and column widths for the active sheet", () => {
        const workbook = createWorkbook({}, [
            createSheet(
                "Sheet1",
                [createCell(3, 3, "value")],
                10,
                10,
                [],
                null,
                {
                    rowHeights: { 2: 24 },
                    columnWidths: { 3: 18.5 },
                }
            ),
        ]);

        const renderModel = createEditorRenderModel(
            workbook,
            createInitialEditorPanelState(workbook)
        );

        assert.deepStrictEqual(renderModel.activeSheet.rowHeights, { 2: 24 });
        assert.deepStrictEqual(renderModel.activeSheet.columnWidths, { 3: 18.5 });
    });

    test("keeps frozen rows separate when virtualizing a locked sheet", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(240, 2, "tail")], 400, 4, [], {
                columnCount: 1,
                rowCount: 2,
                topLeftCell: "B3",
                activePane: "bottomRight",
            }),
        ]);

        const state = setEditorViewportStartRow(
            workbook,
            createInitialEditorPanelState(workbook),
            121
        );
        const renderModel = createEditorRenderModel(workbook, state);

        assert.deepStrictEqual(
            renderModel.page.frozenRows.map((row) => row.rowNumber),
            [1, 2]
        );
        assert.strictEqual(renderModel.page.startRow, 121);
        assert.strictEqual(renderModel.page.rows[0]?.rowNumber, 121);
        assert.strictEqual(renderModel.page.endRow, 320);
    });

    test("keeps the scrollable window below frozen rows on shorter locked sheets", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(150, 2, "tail")], 150, 4, [], {
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

        assert.deepStrictEqual(
            renderModel.page.frozenRows.map((row) => row.rowNumber),
            [1, 2]
        );
        assert.strictEqual(renderModel.page.startRow, 3);
        assert.strictEqual(renderModel.page.rows[0]?.rowNumber, 3);
        assert.strictEqual(renderModel.page.endRow, 150);
    });

    test("keeps viewport position when selecting a frozen row", () => {
        const workbook = createWorkbook({}, [
            createSheet("Sheet1", [createCell(240, 2, "tail")], 400, 4, [], {
                columnCount: 1,
                rowCount: 2,
                topLeftCell: "B3",
                activePane: "bottomRight",
            }),
        ]);

        const scrolledState = setEditorViewportStartRow(
            workbook,
            createInitialEditorPanelState(workbook),
            121
        );
        const selectedState = setSelectedEditorCell(workbook, scrolledState, 1, 1);
        const renderModel = createEditorRenderModel(workbook, selectedState);

        assert.strictEqual(renderModel.selection?.address, "R1C1");
        assert.strictEqual(renderModel.page.startRow, 121);
        assert.strictEqual(renderModel.page.rows[0]?.rowNumber, 121);
    });
});
