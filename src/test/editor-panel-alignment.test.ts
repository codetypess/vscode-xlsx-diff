/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { SheetSnapshot } from "../core/model/types";
import { XlsxEditorPanel } from "../webview/editor-panel";
import {
    applyAlignmentPatchToSheetSnapshot,
    applyGridSheetEditToSheet,
} from "../webview/editor-panel/editor-panel-state";

function createSheet(
    overrides: Partial<SheetSnapshot> = {}
): SheetSnapshot {
    return {
        name: "Sheet1",
        rowCount: 5,
        columnCount: 4,
        visibility: "visible",
        mergedRanges: [],
        columnWidths: [],
        rowHeights: {},
        cellAlignments: {},
        rowAlignments: {},
        columnAlignments: {},
        freezePane: null,
        cells: {},
        signature: "sheet-signature",
        ...overrides,
    };
}

suite("Editor panel alignment", () => {
    test("applies alignment patches to single cells and rectangular ranges", () => {
        const sheet = createSheet({
            cellAlignments: {
                "2:2": {
                    vertical: "bottom",
                },
            },
        });

        const alignedCellSheet = applyAlignmentPatchToSheetSnapshot(
            sheet,
            "cell",
            {
                startRow: 2,
                endRow: 2,
                startColumn: 2,
                endColumn: 2,
            },
            {
                horizontal: "right",
            }
        );
        const alignedRangeSheet = applyAlignmentPatchToSheetSnapshot(
            alignedCellSheet,
            "range",
            {
                startRow: 1,
                endRow: 2,
                startColumn: 1,
                endColumn: 2,
            },
            {
                vertical: "center",
            }
        );

        assert.deepStrictEqual(alignedCellSheet.cellAlignments?.["2:2"], {
            horizontal: "right",
            vertical: "bottom",
        });
        assert.deepStrictEqual(alignedRangeSheet.cellAlignments?.["1:1"], {
            vertical: "center",
        });
        assert.deepStrictEqual(alignedRangeSheet.cellAlignments?.["1:2"], {
            vertical: "center",
        });
        assert.deepStrictEqual(alignedRangeSheet.cellAlignments?.["2:1"], {
            vertical: "center",
        });
        assert.deepStrictEqual(alignedRangeSheet.cellAlignments?.["2:2"], {
            horizontal: "right",
            vertical: "center",
        });
    });

    test("applies alignment patches to full rows and columns", () => {
        const sheet = createSheet({
            rowAlignments: {
                "2": {
                    vertical: "bottom",
                },
            },
            columnAlignments: {
                "3": {
                    horizontal: "right",
                },
            },
        });

        const alignedRowsSheet = applyAlignmentPatchToSheetSnapshot(
            sheet,
            "row",
            {
                startRow: 2,
                endRow: 3,
                startColumn: 1,
                endColumn: 4,
            },
            {
                horizontal: "center",
            }
        );
        const alignedColumnsSheet = applyAlignmentPatchToSheetSnapshot(
            alignedRowsSheet,
            "column",
            {
                startRow: 1,
                endRow: 5,
                startColumn: 3,
                endColumn: 4,
            },
            {
                vertical: "top",
            }
        );

        assert.deepStrictEqual(alignedRowsSheet.rowAlignments?.["2"], {
            horizontal: "center",
            vertical: "bottom",
        });
        assert.deepStrictEqual(alignedRowsSheet.rowAlignments?.["3"], {
            horizontal: "center",
        });
        assert.deepStrictEqual(alignedColumnsSheet.columnAlignments?.["3"], {
            horizontal: "right",
            vertical: "top",
        });
        assert.deepStrictEqual(alignedColumnsSheet.columnAlignments?.["4"], {
            vertical: "top",
        });
    });

    test("moves cell, row, and column alignments with row and column edits", () => {
        const sheet = createSheet({
            cellAlignments: {
                "2:2": {
                    horizontal: "right",
                },
            },
            rowAlignments: {
                "2": {
                    vertical: "center",
                },
            },
            columnAlignments: {
                "2": {
                    horizontal: "center",
                },
            },
        });

        const insertedRowSheet = applyGridSheetEditToSheet(sheet, {
            type: "insertRow",
            sheetKey: "sheet:0",
            sheetName: "Sheet1",
            rowNumber: 2,
            count: 1,
        });
        const deletedColumnSheet = applyGridSheetEditToSheet(insertedRowSheet, {
            type: "deleteColumn",
            sheetKey: "sheet:0",
            sheetName: "Sheet1",
            columnNumber: 1,
            count: 1,
        });

        assert.deepStrictEqual(deletedColumnSheet.cellAlignments, {
            "3:1": {
                horizontal: "right",
            },
        });
        assert.deepStrictEqual(deletedColumnSheet.rowAlignments, {
            "3": {
                vertical: "center",
            },
        });
        assert.deepStrictEqual(deletedColumnSheet.columnAlignments, {
            "1": {
                horizontal: "center",
            },
        });
    });

    test("routes toolbar alignment requests through structural mutation sync", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const activeEntry = {
            key: "sheet:0",
            sheet: createSheet(),
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

        await panel.setPendingAlignment(
            "row",
            {
                startRow: 2,
                endRow: 3,
                startColumn: 1,
                endColumn: 4,
            },
            {
                horizontal: "center",
            }
        );

        assert.deepStrictEqual(activeEntry.sheet.rowAlignments, {
            "2": {
                horizontal: "center",
            },
            "3": {
                horizontal: "center",
            },
        });
        assert.deepStrictEqual(committedOptions, { resetPendingHistory: false });
        assert.strictEqual(syncedSheetKey, "sheet:0");
    });
});
