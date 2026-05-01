/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    createEditorColumnSelectionState,
    createEditorExtendedCellSelectionState,
    createEditorRowSelectionState,
    createEditorSingleCellSelectionState,
    resolveEditorSelectionRange,
} from "../webview-solid/editor-panel/selection-range-state-helpers";

suite("Solid editor selection range state", () => {
    test("resolves a single-cell selection range when no override exists", () => {
        assert.deepStrictEqual(
            resolveEditorSelectionRange(
                {
                    rowNumber: 3,
                    columnNumber: 4,
                },
                null
            ),
            {
                startRow: 3,
                endRow: 3,
                startColumn: 4,
                endColumn: 4,
            }
        );
    });

    test("prefers an explicit selection range override", () => {
        assert.deepStrictEqual(
            resolveEditorSelectionRange(
                {
                    rowNumber: 3,
                    columnNumber: 4,
                },
                {
                    startRow: 2,
                    endRow: 5,
                    startColumn: 1,
                    endColumn: 4,
                }
            ),
            {
                startRow: 2,
                endRow: 5,
                startColumn: 1,
                endColumn: 4,
            }
        );
    });

    test("creates a collapsed single-cell selection state", () => {
        assert.deepStrictEqual(
            createEditorSingleCellSelectionState({
                rowNumber: 2,
                columnNumber: 2,
            }),
            {
                focusCell: {
                    rowNumber: 2,
                    columnNumber: 2,
                },
                anchorCell: {
                    rowNumber: 2,
                    columnNumber: 2,
                },
                selectionRange: null,
            }
        );
    });

    test("extends a cell selection from the anchor to the focus cell", () => {
        assert.deepStrictEqual(
            createEditorExtendedCellSelectionState({
                anchorCell: {
                    rowNumber: 2,
                    columnNumber: 2,
                },
                focusCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
            }),
            {
                focusCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
                anchorCell: {
                    rowNumber: 2,
                    columnNumber: 2,
                },
                selectionRange: {
                    startRow: 2,
                    endRow: 4,
                    startColumn: 2,
                    endColumn: 5,
                },
            }
        );
    });

    test("creates full-row selections and row spans", () => {
        assert.deepStrictEqual(
            createEditorRowSelectionState({
                anchorCell: null,
                focusCell: {
                    rowNumber: 6,
                    columnNumber: 3,
                },
                columnCount: 8,
                extend: false,
            }),
            {
                focusCell: {
                    rowNumber: 6,
                    columnNumber: 3,
                },
                anchorCell: {
                    rowNumber: 6,
                    columnNumber: 3,
                },
                selectionRange: {
                    startRow: 6,
                    endRow: 6,
                    startColumn: 1,
                    endColumn: 8,
                },
            }
        );

        assert.deepStrictEqual(
            createEditorRowSelectionState({
                anchorCell: {
                    rowNumber: 3,
                    columnNumber: 2,
                },
                focusCell: {
                    rowNumber: 6,
                    columnNumber: 2,
                },
                columnCount: 8,
                extend: true,
            }).selectionRange,
            {
                startRow: 3,
                endRow: 6,
                startColumn: 1,
                endColumn: 8,
            }
        );
    });

    test("creates full-column selections and column spans", () => {
        assert.deepStrictEqual(
            createEditorColumnSelectionState({
                anchorCell: null,
                focusCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
                rowCount: 12,
                extend: false,
            }),
            {
                focusCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
                anchorCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
                selectionRange: {
                    startRow: 1,
                    endRow: 12,
                    startColumn: 5,
                    endColumn: 5,
                },
            }
        );

        assert.deepStrictEqual(
            createEditorColumnSelectionState({
                anchorCell: {
                    rowNumber: 4,
                    columnNumber: 2,
                },
                focusCell: {
                    rowNumber: 4,
                    columnNumber: 5,
                },
                rowCount: 12,
                extend: true,
            }).selectionRange,
            {
                startRow: 1,
                endRow: 12,
                startColumn: 2,
                endColumn: 5,
            }
        );
    });
});
