/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    buildFillChanges,
    createAutoFillDownPreviewRange,
    createFillPreviewRange,
} from "../webview/editor-panel/editor-fill-drag";

function createCellValueKey(rowNumber: number, columnNumber: number): string {
    return `${rowNumber}:${columnNumber}`;
}

function createValueGetter(values: Record<string, string>): (rowNumber: number, columnNumber: number) => string {
    return (rowNumber: number, columnNumber: number): string =>
        values[createCellValueKey(rowNumber, columnNumber)] ?? "";
}

suite("Editor fill drag helpers", () => {
    test("creates a downward preview range", () => {
        assert.deepStrictEqual(
            createFillPreviewRange(
                {
                    startRow: 2,
                    endRow: 3,
                    startColumn: 4,
                    endColumn: 5,
                },
                {
                    rowNumber: 6,
                    columnNumber: 5,
                },
                {
                    minRow: 1,
                    maxRow: 20,
                    minColumn: 1,
                    maxColumn: 20,
                }
            ),
            {
                startRow: 2,
                endRow: 6,
                startColumn: 4,
                endColumn: 5,
            }
        );
    });

    test("creates an upward preview range", () => {
        assert.deepStrictEqual(
            createFillPreviewRange(
                {
                    startRow: 5,
                    endRow: 7,
                    startColumn: 2,
                    endColumn: 2,
                },
                {
                    rowNumber: 3,
                    columnNumber: 2,
                },
                {
                    minRow: 1,
                    maxRow: 20,
                    minColumn: 1,
                    maxColumn: 20,
                }
            ),
            {
                startRow: 3,
                endRow: 7,
                startColumn: 2,
                endColumn: 2,
            }
        );
    });

    test("creates a leftward preview range", () => {
        assert.deepStrictEqual(
            createFillPreviewRange(
                {
                    startRow: 4,
                    endRow: 4,
                    startColumn: 3,
                    endColumn: 5,
                },
                {
                    rowNumber: 4,
                    columnNumber: 1,
                },
                {
                    minRow: 1,
                    maxRow: 20,
                    minColumn: 1,
                    maxColumn: 20,
                }
            ),
            {
                startRow: 4,
                endRow: 4,
                startColumn: 1,
                endColumn: 5,
            }
        );
    });

    test("creates a diagonal preview range", () => {
        assert.deepStrictEqual(
            createFillPreviewRange(
                {
                    startRow: 3,
                    endRow: 4,
                    startColumn: 3,
                    endColumn: 4,
                },
                {
                    rowNumber: 6,
                    columnNumber: 7,
                },
                {
                    minRow: 1,
                    maxRow: 20,
                    minColumn: 1,
                    maxColumn: 20,
                }
            ),
            {
                startRow: 3,
                endRow: 6,
                startColumn: 3,
                endColumn: 7,
            }
        );
    });

    test("clamps preview ranges to workbook bounds", () => {
        assert.deepStrictEqual(
            createFillPreviewRange(
                {
                    startRow: 2,
                    endRow: 2,
                    startColumn: 2,
                    endColumn: 2,
                },
                {
                    rowNumber: -4,
                    columnNumber: 99,
                },
                {
                    minRow: 1,
                    maxRow: 6,
                    minColumn: 1,
                    maxColumn: 4,
                }
            ),
            {
                startRow: 1,
                endRow: 2,
                startColumn: 2,
                endColumn: 4,
            }
        );
    });

    test("treats targets inside the source range as a no-op preview", () => {
        assert.strictEqual(
            createFillPreviewRange(
                {
                    startRow: 2,
                    endRow: 4,
                    startColumn: 2,
                    endColumn: 4,
                },
                {
                    rowNumber: 3,
                    columnNumber: 3,
                },
                {
                    minRow: 1,
                    maxRow: 20,
                    minColumn: 1,
                    maxColumn: 20,
                }
            ),
            null
        );
    });

    test("creates an auto-fill-down preview range from adjacent data", () => {
        assert.deepStrictEqual(
            createAutoFillDownPreviewRange({
                sourceRange: {
                    startRow: 2,
                    endRow: 3,
                    startColumn: 2,
                    endColumn: 2,
                },
                bounds: {
                    minRow: 1,
                    maxRow: 12,
                    minColumn: 1,
                    maxColumn: 6,
                },
                getCellValue: createValueGetter({
                    "4:1": "x",
                    "5:1": "y",
                    "6:1": "z",
                }),
            }),
            {
                startRow: 2,
                endRow: 6,
                startColumn: 2,
                endColumn: 2,
            }
        );
    });

    test("uses the longer adjacent data column for auto-fill-down", () => {
        assert.deepStrictEqual(
            createAutoFillDownPreviewRange({
                sourceRange: {
                    startRow: 2,
                    endRow: 3,
                    startColumn: 3,
                    endColumn: 4,
                },
                bounds: {
                    minRow: 1,
                    maxRow: 12,
                    minColumn: 1,
                    maxColumn: 8,
                },
                getCellValue: createValueGetter({
                    "4:2": "left-1",
                    "5:2": "left-2",
                    "6:2": "left-3",
                    "4:5": "right-1",
                    "5:5": "right-2",
                    "6:5": "right-3",
                    "7:5": "right-4",
                    "8:5": "right-5",
                }),
            }),
            {
                startRow: 2,
                endRow: 8,
                startColumn: 3,
                endColumn: 4,
            }
        );
    });

    test("falls back to the display bounds when auto-fill-down has no adjacent data below", () => {
        assert.deepStrictEqual(
            createAutoFillDownPreviewRange({
                sourceRange: {
                    startRow: 4,
                    endRow: 5,
                    startColumn: 1,
                    endColumn: 1,
                },
                bounds: {
                    minRow: 1,
                    maxRow: 10,
                    minColumn: 1,
                    maxColumn: 4,
                },
                getCellValue: createValueGetter({
                    "6:2": "",
                    "7:2": "later-value",
                }),
            }),
            {
                startRow: 4,
                endRow: 10,
                startColumn: 1,
                endColumn: 1,
            }
        );
    });

    test("returns null when auto-fill-down is already at the bottom bound", () => {
        assert.strictEqual(
            createAutoFillDownPreviewRange({
                sourceRange: {
                    startRow: 9,
                    endRow: 10,
                    startColumn: 1,
                    endColumn: 1,
                },
                bounds: {
                    minRow: 1,
                    maxRow: 10,
                    minColumn: 1,
                    maxColumn: 4,
                },
                getCellValue: createValueGetter({}),
            }),
            null
        );
    });

    test("auto-fill-down reuses the fill algorithm for arithmetic series", () => {
        const previewRange = createAutoFillDownPreviewRange({
            sourceRange: {
                startRow: 2,
                endRow: 3,
                startColumn: 2,
                endColumn: 2,
            },
            bounds: {
                minRow: 1,
                maxRow: 12,
                minColumn: 1,
                maxColumn: 6,
            },
            getCellValue: createValueGetter({
                "2:2": "4",
                "3:2": "5",
                "4:1": "row-4",
                "5:1": "row-5",
                "6:1": "row-6",
            }),
        });

        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 2,
                endRow: 3,
                startColumn: 2,
                endColumn: 2,
            },
            previewRange,
            getCellValue: createValueGetter({
                "2:2": "4",
                "3:2": "5",
                "4:2": "",
                "5:2": "",
                "6:2": "",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, columnNumber, afterValue }) => ({
                rowNumber,
                columnNumber,
                afterValue,
            })),
            [
                { rowNumber: 4, columnNumber: 2, afterValue: "6" },
                { rowNumber: 5, columnNumber: 2, afterValue: "7" },
                { rowNumber: 6, columnNumber: 2, afterValue: "8" },
            ]
        );
    });

    test("repeats a single-cell seed across the fill area", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 2,
                endColumn: 2,
            },
            previewRange: {
                startRow: 2,
                endRow: 4,
                startColumn: 2,
                endColumn: 2,
            },
            getCellValue: createValueGetter({
                "2:2": "x",
                "3:2": "",
                "4:2": "",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, columnNumber, afterValue }) => ({
                rowNumber,
                columnNumber,
                afterValue,
            })),
            [
                {
                    rowNumber: 3,
                    columnNumber: 2,
                    afterValue: "x",
                },
                {
                    rowNumber: 4,
                    columnNumber: 2,
                    afterValue: "x",
                },
            ]
        );
    });

    test("tiles a rectangular seed during diagonal expansion", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 1,
                endRow: 2,
                startColumn: 1,
                endColumn: 2,
            },
            previewRange: {
                startRow: 1,
                endRow: 4,
                startColumn: 1,
                endColumn: 4,
            },
            getCellValue: createValueGetter({
                "1:1": "A",
                "1:2": "B",
                "2:1": "C",
                "2:2": "D",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, columnNumber, afterValue }) => ({
                rowNumber,
                columnNumber,
                afterValue,
            })),
            [
                { rowNumber: 1, columnNumber: 3, afterValue: "A" },
                { rowNumber: 1, columnNumber: 4, afterValue: "B" },
                { rowNumber: 2, columnNumber: 3, afterValue: "C" },
                { rowNumber: 2, columnNumber: 4, afterValue: "D" },
                { rowNumber: 3, columnNumber: 1, afterValue: "A" },
                { rowNumber: 3, columnNumber: 2, afterValue: "B" },
                { rowNumber: 3, columnNumber: 3, afterValue: "A" },
                { rowNumber: 3, columnNumber: 4, afterValue: "B" },
                { rowNumber: 4, columnNumber: 1, afterValue: "C" },
                { rowNumber: 4, columnNumber: 2, afterValue: "D" },
                { rowNumber: 4, columnNumber: 3, afterValue: "C" },
                { rowNumber: 4, columnNumber: 4, afterValue: "D" },
            ]
        );
    });

    test("extends arithmetic numeric row seeds forward", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 2,
                endColumn: 3,
            },
            previewRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 2,
                endColumn: 5,
            },
            getCellValue: createValueGetter({
                "2:2": "1",
                "2:3": "3",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ columnNumber, afterValue }) => ({
                columnNumber,
                afterValue,
            })),
            [
                { columnNumber: 4, afterValue: "5" },
                { columnNumber: 5, afterValue: "7" },
            ]
        );
    });

    test("extends arithmetic numeric column seeds backward", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 3,
                endRow: 4,
                startColumn: 2,
                endColumn: 2,
            },
            previewRange: {
                startRow: 1,
                endRow: 4,
                startColumn: 2,
                endColumn: 2,
            },
            getCellValue: createValueGetter({
                "3:2": "10",
                "4:2": "13",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, afterValue }) => ({
                rowNumber,
                afterValue,
            })),
            [
                { rowNumber: 1, afterValue: "4" },
                { rowNumber: 2, afterValue: "7" },
            ]
        );
    });

    test("extends arithmetic numeric column seeds forward", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 10,
                endRow: 11,
                startColumn: 1,
                endColumn: 1,
            },
            previewRange: {
                startRow: 10,
                endRow: 14,
                startColumn: 1,
                endColumn: 1,
            },
            getCellValue: createValueGetter({
                "10:1": "4",
                "11:1": "5",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, afterValue }) => ({
                rowNumber,
                afterValue,
            })),
            [
                { rowNumber: 12, afterValue: "6" },
                { rowNumber: 13, afterValue: "7" },
                { rowNumber: 14, afterValue: "8" },
            ]
        );
    });

    test("keeps extending a vertical series when pointer drift also expands columns", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 10,
                endRow: 11,
                startColumn: 1,
                endColumn: 1,
            },
            previewRange: {
                startRow: 10,
                endRow: 13,
                startColumn: 1,
                endColumn: 2,
            },
            getCellValue: createValueGetter({
                "10:1": "4",
                "11:1": "5",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, columnNumber, afterValue }) => ({
                rowNumber,
                columnNumber,
                afterValue,
            })),
            [
                { rowNumber: 10, columnNumber: 2, afterValue: "4" },
                { rowNumber: 11, columnNumber: 2, afterValue: "5" },
                { rowNumber: 12, columnNumber: 1, afterValue: "6" },
                { rowNumber: 12, columnNumber: 2, afterValue: "6" },
                { rowNumber: 13, columnNumber: 1, afterValue: "7" },
                { rowNumber: 13, columnNumber: 2, afterValue: "7" },
            ]
        );
    });

    test("falls back to repeating for single numeric seeds", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 1,
                endRow: 1,
                startColumn: 1,
                endColumn: 1,
            },
            previewRange: {
                startRow: 1,
                endRow: 1,
                startColumn: 1,
                endColumn: 3,
            },
            getCellValue: createValueGetter({
                "1:1": "9",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ columnNumber, afterValue }) => ({
                columnNumber,
                afterValue,
            })),
            [
                { columnNumber: 2, afterValue: "9" },
                { columnNumber: 3, afterValue: "9" },
            ]
        );
    });

    test("falls back to repeating for mixed numeric seeds", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 1,
                endColumn: 3,
            },
            previewRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 1,
                endColumn: 5,
            },
            getCellValue: createValueGetter({
                "2:1": "1",
                "2:2": "x",
                "2:3": "3",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ columnNumber, afterValue }) => ({
                columnNumber,
                afterValue,
            })),
            [
                { columnNumber: 4, afterValue: "1" },
                { columnNumber: 5, afterValue: "x" },
            ]
        );
    });

    test("falls back to repeating for non-arithmetic numeric seeds", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 1,
                endRow: 3,
                startColumn: 2,
                endColumn: 2,
            },
            previewRange: {
                startRow: 1,
                endRow: 5,
                startColumn: 2,
                endColumn: 2,
            },
            getCellValue: createValueGetter({
                "1:2": "1",
                "2:2": "2",
                "3:2": "4",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(
            changes.map(({ rowNumber, afterValue }) => ({
                rowNumber,
                afterValue,
            })),
            [
                { rowNumber: 4, afterValue: "1" },
                { rowNumber: 5, afterValue: "2" },
            ]
        );
    });

    test("excludes source cells and preserves effective before values", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 2,
                endColumn: 3,
            },
            previewRange: {
                startRow: 2,
                endRow: 2,
                startColumn: 1,
                endColumn: 3,
            },
            getCellValue: createValueGetter({
                "2:1": "pending-old",
                "2:2": "A",
                "2:3": "B",
            }),
            getModelValue: createValueGetter({
                "2:1": "model-old",
            }),
            canEditCell: () => true,
        });

        assert.deepStrictEqual(changes, [
            {
                sheetKey: "sheet:1",
                rowNumber: 2,
                columnNumber: 1,
                modelValue: "model-old",
                beforeValue: "pending-old",
                afterValue: "B",
            },
        ]);
    });

    test("skips non-editable target cells", () => {
        const changes = buildFillChanges({
            sheetKey: "sheet:1",
            sourceRange: {
                startRow: 1,
                endRow: 1,
                startColumn: 1,
                endColumn: 2,
            },
            previewRange: {
                startRow: 1,
                endRow: 1,
                startColumn: 1,
                endColumn: 4,
            },
            getCellValue: createValueGetter({
                "1:1": "A",
                "1:2": "B",
            }),
            getModelValue: createValueGetter({}),
            canEditCell: (_rowNumber, columnNumber) => columnNumber !== 3,
        });

        assert.deepStrictEqual(
            changes.map(({ columnNumber, afterValue }) => ({
                columnNumber,
                afterValue,
            })),
            [{ columnNumber: 4, afterValue: "B" }]
        );
    });
});
