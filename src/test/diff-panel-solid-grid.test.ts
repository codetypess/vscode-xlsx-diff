/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type {
    DiffPanelColumnView,
    DiffPanelRowView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "../webview/diff-panel/diff-panel-types";
import {
    beginDiffCellEdit,
    clampDiffHorizontalScroll,
    createSaveDiffEditsMessage,
    finalizeDiffCellEdit,
    filterDiffRows,
    getDiffPreviewState,
    getPendingDiffEditKey,
    getRenderedDiffCellValue,
    getSelectedDiffCellState,
    getDiffTrackWidth,
    getWrappedDiffIndex,
} from "../webview-solid/diff-panel/grid-helpers";

function createCell(overrides: Partial<DiffPanelSparseCellView>): DiffPanelSparseCellView {
    return {
        key: "cell:1:1",
        columnNumber: 1,
        address: "A1",
        status: "modified",
        diffIndex: 0,
        leftPresent: true,
        rightPresent: true,
        leftValue: "left",
        rightValue: "right",
        leftFormula: null,
        rightFormula: null,
        ...overrides,
    };
}

function createRow(overrides: Partial<DiffPanelRowView>): DiffPanelRowView {
    return {
        rowNumber: 1,
        leftRowNumber: 1,
        rightRowNumber: 1,
        hasDiff: true,
        diffTone: "modified",
        cells: [createCell({})],
        ...overrides,
    };
}

function createColumn(columnNumber: number): DiffPanelColumnView {
    return {
        columnNumber,
        leftColumnNumber: columnNumber,
        rightColumnNumber: columnNumber,
        columnWidth: null,
        leftLabel: String.fromCharCode(64 + columnNumber),
        rightLabel: String.fromCharCode(64 + columnNumber),
    };
}

function createSheet(overrides: Partial<DiffPanelSheetView>): DiffPanelSheetView {
    return {
        key: "sheet:1",
        kind: "matched",
        label: "Sheet1",
        leftName: "Sheet1",
        rightName: "Sheet1",
        rowCount: 2,
        columnCount: 2,
        columns: [createColumn(1), createColumn(2)],
        rows: [
            createRow({
                rowNumber: 1,
                cells: [
                    createCell({
                        key: "cell:1:1",
                        columnNumber: 1,
                        address: "A1",
                        diffIndex: 0,
                        leftValue: "alpha",
                        rightValue: "alpha changed",
                    }),
                ],
            }),
            createRow({
                rowNumber: 2,
                hasDiff: false,
                diffTone: "equal",
                cells: [
                    createCell({
                        key: "cell:2:1",
                        columnNumber: 1,
                        address: "A2",
                        status: "equal",
                        diffIndex: null,
                        leftValue: "beta",
                        rightValue: "beta",
                    }),
                ],
            }),
        ],
        diffRows: [1],
        diffCells: [
            {
                key: "diff:A1",
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                diffIndex: 0,
            },
        ],
        diffRowCount: 1,
        diffCellCount: 1,
        mergedRangesChanged: false,
        freezePaneChanged: false,
        visibilityChanged: false,
        sheetOrderChanged: false,
        ...overrides,
    };
}

suite("Solid diff grid helpers", () => {
    test("filters rows by the active row mode", () => {
        const rows = [
            createRow({ rowNumber: 1, hasDiff: true }),
            createRow({ rowNumber: 2, hasDiff: false }),
        ];

        assert.deepStrictEqual(
            filterDiffRows(rows, "diffs").map((row) => row.rowNumber),
            [1]
        );
        assert.deepStrictEqual(
            filterDiffRows(rows, "same").map((row) => row.rowNumber),
            [2]
        );
        assert.strictEqual(filterDiffRows(rows, "all"), rows);
    });

    test("wraps diff navigation at the sheet boundaries", () => {
        assert.strictEqual(getWrappedDiffIndex(0, 3, -1), 2);
        assert.strictEqual(getWrappedDiffIndex(2, 3, 1), 0);
        assert.strictEqual(getWrappedDiffIndex(1, 3, 1), 2);
        assert.strictEqual(getWrappedDiffIndex(0, 0, 1), 0);
    });

    test("builds preview state from the active diff index", () => {
        const sheet = createSheet({
            diffCells: [
                {
                    key: "diff:A1",
                    rowNumber: 1,
                    columnNumber: 1,
                    address: "A1",
                    diffIndex: 0,
                },
                {
                    key: "diff:B1",
                    rowNumber: 1,
                    columnNumber: 2,
                    address: "B1",
                    diffIndex: 1,
                },
            ],
            rows: [
                createRow({
                    rowNumber: 1,
                    cells: [
                        createCell({
                            key: "cell:1:1",
                            columnNumber: 1,
                            address: "A1",
                            diffIndex: 0,
                            leftValue: "alpha",
                            rightValue: "alpha changed",
                        }),
                        createCell({
                            key: "cell:1:2",
                            columnNumber: 2,
                            address: "B1",
                            diffIndex: 1,
                            leftValue: "beta",
                            rightValue: "beta changed",
                        }),
                    ],
                }),
            ],
        });

        const preview = getDiffPreviewState(sheet, 1);

        assert.deepStrictEqual(preview, {
            address: "B1",
            rowNumber: 1,
            columnNumber: 2,
            leftValue: "beta",
            rightValue: "beta changed",
            leftPresent: true,
            rightPresent: true,
            index: 1,
            total: 2,
        });
        assert.strictEqual(getDiffPreviewState(sheet, 99)?.address, "B1");
        assert.strictEqual(getDiffPreviewState(null, 0), null);
    });

    test("selecting a diff cell keeps selection and preview navigation aligned", () => {
        const row = createRow({ rowNumber: 3 });
        const diffCell = createCell({ columnNumber: 2, diffIndex: 4 });
        const equalCell = createCell({ columnNumber: 3, diffIndex: null, status: "equal" });

        assert.deepStrictEqual(getSelectedDiffCellState(row, diffCell), {
            selectedCell: {
                rowNumber: 3,
                columnNumber: 2,
            },
            activeDiffIndex: 4,
        });
        assert.deepStrictEqual(getSelectedDiffCellState(row, equalCell), {
            selectedCell: {
                rowNumber: 3,
                columnNumber: 3,
            },
            activeDiffIndex: null,
        });
    });

    test("starts inline editing from the staged value when one already exists", () => {
        const row = createRow({ rowNumber: 4 });
        const cell = createCell({ columnNumber: 2, leftValue: "before", rightValue: "after" });
        const pendingEdits = {
            [getPendingDiffEditKey(4, 2, "left")]: {
                sheetKey: "sheet:1",
                side: "left" as const,
                rowNumber: 4,
                columnNumber: 2,
                value: "staged",
            },
        };

        const nextEditingState = beginDiffCellEdit({
            activeSheetKey: "sheet:1",
            side: "left",
            sideEditable: true,
            pendingEdits,
            row,
            cell,
        });

        assert.deepStrictEqual(nextEditingState, {
            editingCell: {
                sheetKey: "sheet:1",
                side: "left",
                rowNumber: 4,
                columnNumber: 2,
                value: "staged",
            },
            selectedCell: {
                rowNumber: 4,
                columnNumber: 2,
            },
        });
        assert.strictEqual(
            beginDiffCellEdit({
                activeSheetKey: "sheet:1",
                side: "left",
                sideEditable: false,
                pendingEdits: {},
                row,
                cell,
            }),
            null
        );
        assert.strictEqual(
            beginDiffCellEdit({
                activeSheetKey: "sheet:1",
                side: "right",
                sideEditable: true,
                pendingEdits: {},
                row,
                cell: createCell({ rightPresent: false }),
            }),
            null
        );
    });

    test("commits and cancels inline edits without corrupting staged state", () => {
        const existingPendingEdits = {
            [getPendingDiffEditKey(2, 1, "right")]: {
                sheetKey: "sheet:1",
                side: "right" as const,
                rowNumber: 2,
                columnNumber: 1,
                value: "kept",
            },
        };

        const cancelResult = finalizeDiffCellEdit(
            existingPendingEdits,
            {
                sheetKey: "sheet:1",
                side: "left",
                rowNumber: 1,
                columnNumber: 1,
                value: "discard",
            },
            "cancel"
        );

        assert.strictEqual(cancelResult.editingCell, null);
        assert.strictEqual(cancelResult.pendingEdits, existingPendingEdits);

        const commitResult = finalizeDiffCellEdit(
            existingPendingEdits,
            {
                sheetKey: "sheet:1",
                side: "left",
                rowNumber: 1,
                columnNumber: 1,
                value: "committed",
            },
            "commit"
        );

        assert.strictEqual(commitResult.editingCell, null);
        assert.deepStrictEqual(commitResult.pendingEdits, {
            ...existingPendingEdits,
            [getPendingDiffEditKey(1, 1, "left")]: {
                sheetKey: "sheet:1",
                side: "left",
                rowNumber: 1,
                columnNumber: 1,
                value: "committed",
            },
        });
    });

    test("renders staged values and creates save messages from pending edits", () => {
        const row = createRow({ rowNumber: 6 });
        const cell = createCell({ columnNumber: 3, leftValue: "left", rightValue: "right" });
        const pendingEdits = {
            [getPendingDiffEditKey(6, 3, "right")]: {
                sheetKey: "sheet:1",
                side: "right" as const,
                rowNumber: 6,
                columnNumber: 3,
                value: "updated",
            },
        };

        assert.strictEqual(getRenderedDiffCellValue(pendingEdits, row, cell, "right"), "updated");
        assert.strictEqual(getRenderedDiffCellValue({}, row, cell, "left"), "left");
        assert.deepStrictEqual(createSaveDiffEditsMessage(pendingEdits), {
            type: "saveEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    side: "right",
                    rowNumber: 6,
                    columnNumber: 3,
                    value: "updated",
                },
            ],
        });
        assert.strictEqual(createSaveDiffEditsMessage({}), null);
    });

    test("clamps horizontal scrolling to the visible track", () => {
        assert.strictEqual(
            getDiffTrackWidth([createColumn(1), createColumn(2), createColumn(3)]),
            360
        );
        assert.strictEqual(clampDiffHorizontalScroll(-20, 360, 180), 0);
        assert.strictEqual(clampDiffHorizontalScroll(50, 360, 180), 50);
        assert.strictEqual(clampDiffHorizontalScroll(400, 360, 180), 180);
        assert.strictEqual(clampDiffHorizontalScroll(40, 120, 240), 0);
    });
});
