/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey, getCellAddress } from "../core/model/cells";
import type { CellSnapshot, EditorPanelState, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    findEditorSearchMatch,
    getInsertEditorSheetIndex,
    getNewEditorSheetName,
    resolveEditorCellReference,
    validateEditorSheetName,
} from "../webview/editor-panel-logic";
import {
    captureStructuralSnapshot,
    createPendingWorkbookEditState,
    createWorkingSheetEntries,
    mapPendingCellEditsToWebview,
    restoreStructuralSnapshot,
} from "../webview/editor-panel-state";

function createCell(
    rowNumber: number,
    columnNumber: number,
    value: string,
    formula: string | null = null
): CellSnapshot {
    return {
        key: createCellKey(rowNumber, columnNumber),
        rowNumber,
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        displayValue: value,
        formula,
        styleId: null,
    };
}

function createSheet(
    name: string,
    cells: CellSnapshot[],
    { rowCount = 10, columnCount = 4 }: { rowCount?: number; columnCount?: number } = {}
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        mergedRanges: [],
        freezePane: null,
        cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
        signature: `${name}:${cells.map((cell) => cell.address).join("|")}`,
    };
}

function createWorkbook(fileName: string, sheets: SheetSnapshot[]): WorkbookSnapshot {
    return {
        filePath: `/tmp/${fileName}`,
        fileName,
        fileSize: 0,
        modifiedTime: new Date("2026-01-01T00:00:00.000Z").toISOString(),
        sheets,
        isReadonly: false,
    };
}

suite("Editor panel logic", () => {
    test("search wraps around the active sheet and respects the current selection", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "beta"),
                createCell(4, 1, "alpha"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 1, columnNumber: 1 },
        };

        const nextMatch = findEditorSearchMatch(sheetEntries, state, "alpha", "next", {
            isRegexp: false,
            matchCase: false,
            wholeWord: false,
        });
        const previousMatch = findEditorSearchMatch(sheetEntries, state, "alpha", "prev", {
            isRegexp: false,
            matchCase: false,
            wholeWord: false,
        });

        assert.deepStrictEqual(nextMatch, {
            sheetKey: "sheet:0",
            rowNumber: 4,
            columnNumber: 1,
        });
        assert.deepStrictEqual(previousMatch, {
            sheetKey: "sheet:0",
            rowNumber: 4,
            columnNumber: 1,
        });
    });

    test("goto cell resolves sheet names case-insensitively and rejects out-of-range cells", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "alpha")], { rowCount: 3, columnCount: 3 }),
            createSheet("Data", [createCell(2, 2, "beta")], { rowCount: 5, columnCount: 5 }),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);

        assert.deepStrictEqual(
            resolveEditorCellReference(sheetEntries, "sheet:0", "data!B2"),
            {
                sheetKey: "sheet:1",
                rowNumber: 2,
                columnNumber: 2,
            }
        );
        assert.strictEqual(
            resolveEditorCellReference(sheetEntries, "sheet:0", "Sheet1!D4"),
            null
        );
    });

    test("sheet helpers validate names and derive insertion defaults", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", []),
            createSheet("Sheet2", []),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const strings = {
            sheetNameDuplicate: "duplicate",
            sheetNameEmpty: "empty",
            sheetNameInvalidChars: "invalid",
            sheetNameTooLong: "too-long",
        };

        assert.strictEqual(
            validateEditorSheetName(" sheet1 ", sheetEntries, strings, undefined),
            "duplicate"
        );
        assert.strictEqual(validateEditorSheetName("bad/name", sheetEntries, strings), "invalid");
        assert.strictEqual(
            getNewEditorSheetName(sheetEntries, { isChinese: false }),
            "Sheet3"
        );
        assert.strictEqual(getInsertEditorSheetIndex(sheetEntries, "sheet:0"), 1);
    });
});

suite("Editor panel state helpers", () => {
    test("structural snapshots restore cloned state and preserve pending edit mappings", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "alpha")], { rowCount: 5, columnCount: 3 }),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 1, columnNumber: 1 },
        };
        const pendingCellEdits = [
            {
                sheetName: "Sheet1",
                rowNumber: 1,
                columnNumber: 1,
                value: "changed",
            },
        ];
        const pendingSheetEdits = [
            {
                type: "renameSheet" as const,
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                nextSheetName: "Renamed",
            },
        ];
        const pendingViewEdits = [
            {
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                freezePane: { rowCount: 1, columnCount: 1 },
            },
        ];

        const snapshot = captureStructuralSnapshot(
            state,
            sheetEntries,
            pendingCellEdits,
            pendingSheetEdits,
            pendingViewEdits
        );

        state.selectedCell!.rowNumber = 9;
        sheetEntries[0]!.sheet.name = "Mutated";
        pendingCellEdits[0]!.value = "mutated";
        pendingViewEdits[0]!.freezePane!.rowCount = 9;

        const restored = restoreStructuralSnapshot(snapshot);
        const pendingState = createPendingWorkbookEditState(
            restored.pendingCellEdits,
            restored.pendingSheetEdits,
            restored.pendingViewEdits
        );

        assert.deepStrictEqual(restored.state.selectedCell, { rowNumber: 1, columnNumber: 1 });
        assert.strictEqual(restored.sheetEntries[0]!.sheet.name, "Sheet1");
        assert.strictEqual(restored.pendingCellEdits[0]!.value, "changed");
        assert.deepStrictEqual(restored.pendingViewEdits[0]!.freezePane, {
            rowCount: 1,
            columnCount: 1,
        });
        assert.deepStrictEqual(
            mapPendingCellEditsToWebview(restored.pendingCellEdits, restored.sheetEntries),
            [
                {
                    sheetKey: "sheet:0",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "changed",
                },
            ]
        );
        assert.deepStrictEqual(pendingState.viewEdits, [
            {
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                freezePane: { rowCount: 1, columnCount: 1 },
            },
        ]);
    });
});
