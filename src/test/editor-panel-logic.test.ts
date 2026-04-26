/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey, getCellAddress } from "../core/model/cells";
import type { CellSnapshot, EditorPanelState, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    findEditorSearchMatch,
    getInsertEditorSheetIndex,
    getNewEditorSheetName,
    resolveEditorReplaceResultInSheet,
    resolveEditorSearchResult,
    resolveEditorCellReference,
    validateEditorSheetName,
} from "../webview/editor-panel/editor-panel-logic";
import {
    applyGridSheetEditToSheet,
    captureStructuralSnapshot,
    createCommittedWorkbookState,
    createPendingWorkbookEditState,
    createWorkingWorkbook,
    restorePendingWorkbookState,
    createWorkingSheetEntries,
    mapPendingCellEditsToWebview,
    restoreStructuralSnapshot,
    shiftPendingCellEditsForGridSheetEdit,
} from "../webview/editor-panel/editor-panel-state";

function createCell(
    rowNumber: number,
    columnNumber: number,
    value: string,
    formula: string | null = null,
    styleId: number | null = null
): CellSnapshot {
    return {
        key: createCellKey(rowNumber, columnNumber),
        rowNumber,
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        displayValue: value,
        formula,
        styleId,
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

    test("search can be limited to the selected range", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "alpha"),
                createCell(3, 1, "alpha"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 2, columnNumber: 1 },
        };

        const nextMatch = findEditorSearchMatch(
            sheetEntries,
            state,
            "alpha",
            "next",
            {
                isRegexp: false,
                matchCase: false,
                wholeWord: false,
            },
            {
                scope: "selection",
                selectionRange: {
                    startRow: 2,
                    endRow: 3,
                    startColumn: 1,
                    endColumn: 1,
                },
            }
        );

        assert.deepStrictEqual(nextMatch, {
            sheetKey: "sheet:0",
            rowNumber: 3,
            columnNumber: 1,
        });
    });

    test("search matches pending edits before the saved snapshot", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "beta"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 1, columnNumber: 1 },
        };

        const nextMatch = findEditorSearchMatch(
            sheetEntries,
            state,
            "alpha",
            "next",
            {
                isRegexp: false,
                matchCase: false,
                wholeWord: false,
            },
            {
                pendingEdits: [
                    {
                        sheetName: "Sheet1",
                        rowNumber: 2,
                        columnNumber: 1,
                        value: "alpha",
                    },
                ],
            }
        );

        assert.deepStrictEqual(nextMatch, {
            sheetKey: "sheet:0",
            rowNumber: 2,
            columnNumber: 1,
        });
    });

    test("invalid regexp search returns an invalid-pattern result", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "alpha")]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 1, columnNumber: 1 },
        };

        const result = resolveEditorSearchResult(
            sheetEntries,
            state,
            {
                query: "[alpha",
                direction: "next",
                options: {
                    isRegexp: true,
                    matchCase: false,
                    wholeWord: false,
                },
                scope: "sheet",
            }
        );

        assert.strictEqual(result.status, "invalid-pattern");
    });

    test("search result includes the current match index and total count", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "alpha"),
                createCell(4, 1, "alpha"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const state: EditorPanelState = {
            activeSheetKey: "sheet:0",
            selectedCell: { rowNumber: 2, columnNumber: 1 },
        };

        const result = resolveEditorSearchResult(
            sheetEntries,
            state,
            {
                query: "alpha",
                direction: "next",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                scope: "sheet",
            }
        );

        assert.deepStrictEqual(result, {
            status: "matched",
            match: {
                sheetKey: "sheet:0",
                rowNumber: 4,
                columnNumber: 1,
            },
            matchCount: 3,
            matchIndex: 3,
        });
    });

    test("replace updates the current editable match and advances to the next one", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "alpha"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const activeSheetEntry = sheetEntries[0]!;

        const result = resolveEditorReplaceResultInSheet(
            {
                key: activeSheetEntry.key,
                rowCount: activeSheetEntry.sheet.rowCount,
                columnCount: activeSheetEntry.sheet.columnCount,
                cells: activeSheetEntry.sheet.cells,
            },
            { rowNumber: 1, columnNumber: 1 },
            {
                query: "alpha",
                replacement: "gamma",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                mode: "single",
            }
        );

        assert.deepStrictEqual(result, {
            status: "replaced",
            changes: [
                {
                    rowNumber: 1,
                    columnNumber: 1,
                    beforeValue: "alpha",
                    afterValue: "gamma",
                },
            ],
            replacedCellCount: 1,
            match: {
                sheetKey: "sheet:0",
                rowNumber: 1,
                columnNumber: 1,
            },
            nextMatch: {
                sheetKey: "sheet:0",
                rowNumber: 2,
                columnNumber: 1,
            },
        });
    });

    test("replace all respects selection scope and pending edits", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [
                createCell(1, 1, "alpha"),
                createCell(2, 1, "beta"),
                createCell(3, 1, "alpha"),
            ]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const activeSheetEntry = sheetEntries[0]!;

        const result = resolveEditorReplaceResultInSheet(
            {
                key: activeSheetEntry.key,
                rowCount: activeSheetEntry.sheet.rowCount,
                columnCount: activeSheetEntry.sheet.columnCount,
                cells: activeSheetEntry.sheet.cells,
            },
            { rowNumber: 2, columnNumber: 1 },
            {
                query: "alpha",
                replacement: "gamma",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                scope: "selection",
                selectionRange: {
                    startRow: 2,
                    endRow: 3,
                    startColumn: 1,
                    endColumn: 1,
                },
                pendingEdits: [
                    {
                        rowNumber: 2,
                        columnNumber: 1,
                        value: "alpha",
                    },
                ],
                mode: "all",
            }
        );

        assert.deepStrictEqual(result, {
            status: "replaced",
            changes: [
                {
                    rowNumber: 2,
                    columnNumber: 1,
                    beforeValue: "alpha",
                    afterValue: "gamma",
                },
                {
                    rowNumber: 3,
                    columnNumber: 1,
                    beforeValue: "alpha",
                    afterValue: "gamma",
                },
            ],
            replacedCellCount: 2,
            match: {
                sheetKey: "sheet:0",
                rowNumber: 2,
                columnNumber: 1,
            },
        });
    });

    test("replace skips formula cells because they are not editable", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "alpha", '=LOWER("ALPHA")')]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const activeSheetEntry = sheetEntries[0]!;

        const result = resolveEditorReplaceResultInSheet(
            {
                key: activeSheetEntry.key,
                rowCount: activeSheetEntry.sheet.rowCount,
                columnCount: activeSheetEntry.sheet.columnCount,
                cells: activeSheetEntry.sheet.cells,
            },
            { rowNumber: 1, columnNumber: 1 },
            {
                query: "alpha",
                replacement: "gamma",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                mode: "all",
            }
        );

        assert.deepStrictEqual(result, {
            status: "no-match",
        });
    });

    test("replace reports no-change when the replacement keeps the same values", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "alpha")]),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const activeSheetEntry = sheetEntries[0]!;

        const result = resolveEditorReplaceResultInSheet(
            {
                key: activeSheetEntry.key,
                rowCount: activeSheetEntry.sheet.rowCount,
                columnCount: activeSheetEntry.sheet.columnCount,
                cells: activeSheetEntry.sheet.cells,
            },
            { rowNumber: 1, columnNumber: 1 },
            {
                query: "alpha",
                replacement: "alpha",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                mode: "all",
            }
        );

        assert.deepStrictEqual(result, {
            status: "no-change",
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
        assert.strictEqual(getNewEditorSheetName(sheetEntries, "Sheet"), "Sheet3");
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

    test("commits working sheets and pending cell edits into a saved workbook snapshot", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(2, 2, "before", null, 7)], {
                rowCount: 5,
                columnCount: 4,
            }),
            createSheet("Sheet2", [], { rowCount: 4, columnCount: 3 }),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        sheetEntries[0]!.sheet = {
            ...sheetEntries[0]!.sheet,
            name: "Renamed",
            freezePane: {
                columnCount: 1,
                rowCount: 1,
                topLeftCell: "B2",
                activePane: "bottomRight",
            },
        };

        const committedState = createCommittedWorkbookState(workbook, sheetEntries, [
            {
                sheetName: "Renamed",
                rowNumber: 2,
                columnNumber: 2,
                value: "after",
            },
            {
                sheetName: "Sheet2",
                rowNumber: 4,
                columnNumber: 3,
                value: "tail",
            },
        ]);
        const committedWorkbook = committedState.workbook;

        assert.deepStrictEqual(
            workbook.sheets.map((sheet) => sheet.name),
            ["Sheet1", "Sheet2"]
        );
        assert.deepStrictEqual(
            committedWorkbook.sheets.map((sheet) => sheet.name),
            ["Renamed", "Sheet2"]
        );
        assert.deepStrictEqual(committedWorkbook.sheets[0]!.freezePane, {
            columnCount: 1,
            rowCount: 1,
            topLeftCell: "B2",
            activePane: "bottomRight",
        });
        assert.strictEqual(committedWorkbook.sheets[0]!.cells["2:2"]!.displayValue, "after");
        assert.strictEqual(committedWorkbook.sheets[0]!.cells["2:2"]!.styleId, 7);
        assert.strictEqual(committedWorkbook.sheets[1]!.cells["4:3"]!.displayValue, "tail");
        assert.strictEqual(committedWorkbook.sheets[1]!.cells["4:3"]!.address, "C4");
        assert.strictEqual(committedWorkbook.sheets[1]!.rowCount, 4);
        assert.strictEqual(committedWorkbook.sheets[1]!.columnCount, 3);
        assert.strictEqual(committedState.sheetEntries[0]!.sheet.cells["2:2"]!.displayValue, "after");
        assert.strictEqual(sheetEntries[0]!.sheet.cells["2:2"]!.displayValue, "before");
    });

    test("expands sheet bounds for pending and committed edits beyond the used range", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "head")], { rowCount: 5, columnCount: 3 }),
        ]);
        const sheetEntries = createWorkingSheetEntries(workbook);
        const pendingCellEdits = [
            {
                sheetName: "Sheet1",
                rowNumber: 12,
                columnNumber: 8,
                value: "tail",
            },
        ];

        const workingWorkbook = createWorkingWorkbook(workbook, sheetEntries, pendingCellEdits);
        const committedState = createCommittedWorkbookState(workbook, sheetEntries, pendingCellEdits);

        assert.strictEqual(workingWorkbook.sheets[0]!.rowCount, 12);
        assert.strictEqual(workingWorkbook.sheets[0]!.columnCount, 8);
        assert.strictEqual(committedState.workbook.sheets[0]!.rowCount, 12);
        assert.strictEqual(committedState.workbook.sheets[0]!.columnCount, 8);
        assert.strictEqual(committedState.workbook.sheets[0]!.cells["12:8"]!.displayValue, "tail");
        assert.strictEqual(committedState.workbook.sheets[0]!.cells["12:8"]!.address, "H12");
    });

    test("restores working state from pending workbook edits", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "base")], { rowCount: 5, columnCount: 3 }),
        ]);

        const restored = restorePendingWorkbookState(workbook, {
            cellEdits: [
                {
                    sheetName: "Renamed",
                    rowNumber: 2,
                    columnNumber: 2,
                    value: "changed",
                },
            ],
            sheetEdits: [
                {
                    type: "renameSheet",
                    sheetKey: "sheet:0",
                    sheetName: "Sheet1",
                    nextSheetName: "Renamed",
                },
                {
                    type: "addSheet",
                    sheetKey: "sheet:new:3",
                    sheetName: "Added",
                    targetIndex: 1,
                },
            ],
            viewEdits: [
                {
                    sheetKey: "sheet:0",
                    sheetName: "Renamed",
                    freezePane: {
                        rowCount: 1,
                        columnCount: 1,
                    },
                },
            ],
        });

        assert.deepStrictEqual(
            restored.sheetEntries.map((entry) => entry.sheet.name),
            ["Renamed", "Added"]
        );
        assert.deepStrictEqual(restored.sheetEntries[0]!.sheet.freezePane, {
            rowCount: 1,
            columnCount: 1,
            topLeftCell: "B2",
            activePane: "bottomRight",
        });
        assert.strictEqual(restored.pendingCellEdits[0]!.sheetName, "Renamed");
        assert.strictEqual(restored.pendingSheetEdits.length, 2);
        assert.strictEqual(restored.pendingViewEdits.length, 1);
        assert.strictEqual(restored.nextNewSheetId, 4);
    });

    test("restores pending row and column edits into working sheet snapshots", () => {
        const workbook = createWorkbook("editor.xlsx", [
            createSheet(
                "Sheet1",
                [
                    createCell(1, 1, "A1"),
                    createCell(2, 2, "B2"),
                    createCell(3, 3, "C3"),
                ],
                { rowCount: 3, columnCount: 3 }
            ),
        ]);

        const restored = restorePendingWorkbookState(workbook, {
            cellEdits: [],
            sheetEdits: [
                {
                    type: "insertRow",
                    sheetKey: "sheet:0",
                    sheetName: "Sheet1",
                    rowNumber: 2,
                    count: 1,
                },
                {
                    type: "deleteColumn",
                    sheetKey: "sheet:0",
                    sheetName: "Sheet1",
                    columnNumber: 1,
                    count: 1,
                },
            ],
            viewEdits: [],
        });

        assert.strictEqual(restored.sheetEntries[0]!.sheet.rowCount, 4);
        assert.strictEqual(restored.sheetEntries[0]!.sheet.columnCount, 2);
        assert.strictEqual(restored.sheetEntries[0]!.sheet.cells["1:1"], undefined);
        assert.strictEqual(restored.sheetEntries[0]!.sheet.cells["3:1"]?.displayValue, "B2");
        assert.strictEqual(restored.sheetEntries[0]!.sheet.cells["4:2"]?.displayValue, "C3");
    });

    test("shifts pending cell edits when rows and columns move", () => {
        const movedAfterRowInsert = shiftPendingCellEditsForGridSheetEdit(
            [
                {
                    sheetName: "Sheet1",
                    rowNumber: 2,
                    columnNumber: 2,
                    value: "pending",
                },
            ],
            {
                type: "insertRow",
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                rowNumber: 2,
                count: 1,
            }
        );
        const movedAfterColumnDelete = shiftPendingCellEditsForGridSheetEdit(
            movedAfterRowInsert,
            {
                type: "deleteColumn",
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                columnNumber: 1,
                count: 1,
            }
        );

        assert.deepStrictEqual(movedAfterColumnDelete, [
            {
                sheetName: "Sheet1",
                rowNumber: 3,
                columnNumber: 1,
                value: "pending",
            },
        ]);
    });

    test("applies grid edits directly to sheet snapshots", () => {
        const sheet = createSheet(
            "Sheet1",
            [createCell(2, 2, "B2"), createCell(3, 3, "C3")],
            { rowCount: 3, columnCount: 3 }
        );

        const updatedSheet = applyGridSheetEditToSheet(sheet, {
            type: "insertColumn",
            sheetKey: "sheet:0",
            sheetName: "Sheet1",
            columnNumber: 2,
            count: 1,
        });

        assert.strictEqual(updatedSheet.columnCount, 4);
        assert.strictEqual(updatedSheet.cells["2:3"]?.displayValue, "B2");
        assert.strictEqual(updatedSheet.cells["3:4"]?.displayValue, "C3");
    });
});
