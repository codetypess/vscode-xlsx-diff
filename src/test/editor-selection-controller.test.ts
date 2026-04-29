/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    clearPendingSelectionAfterRender,
    clearSelectedCell,
    createSelectionControllerState,
    getExpandedSelectionRange,
    getSelectionRange,
    isActiveSelectionCell,
    isSimpleSelectionState,
    setPendingSelectionAfterRender,
    setSelectedCell,
    startCellSelectionDrag,
    stopSelectionDrag,
    syncSelectionAnchorToSelectedCell,
} from "../webview/editor-panel/editor-selection-controller";

suite("Editor selection controller", () => {
    test("creates an empty selection state by default", () => {
        assert.deepStrictEqual(createSelectionControllerState(), {
            selectedCell: null,
            selectionAnchorCell: null,
            selectionRangeOverride: null,
            pendingSelectionAfterRender: null,
            suppressAutoSelection: false,
            selectionDragState: null,
        });
    });

    test("selects a simple cell and keeps it as the anchor", () => {
        const state = setSelectedCell(createSelectionControllerState(), {
            rowNumber: 4,
            columnNumber: 3,
        });

        assert.deepStrictEqual(state.selectedCell, { rowNumber: 4, columnNumber: 3 });
        assert.deepStrictEqual(state.selectionAnchorCell, { rowNumber: 4, columnNumber: 3 });
        assert.strictEqual(state.selectionRangeOverride, null);
        assert.strictEqual(isSimpleSelectionState(state), true);
        assert.strictEqual(isActiveSelectionCell(state, 4, 3), true);
    });

    test("derives expanded ranges from anchor and focus cells", () => {
        const state = setSelectedCell(
            createSelectionControllerState(),
            {
                rowNumber: 7,
                columnNumber: 5,
            },
            {
                anchorCell: {
                    rowNumber: 2,
                    columnNumber: 3,
                },
            }
        );

        assert.deepStrictEqual(getSelectionRange(state), {
            startRow: 2,
            endRow: 7,
            startColumn: 3,
            endColumn: 5,
        });
        assert.deepStrictEqual(getExpandedSelectionRange(state), {
            startRow: 2,
            endRow: 7,
            startColumn: 3,
            endColumn: 5,
        });
        assert.strictEqual(isSimpleSelectionState(state), false);
    });

    test("clears selection and suppresses auto selection", () => {
        const state = clearSelectedCell(
            setSelectedCell(createSelectionControllerState(), {
                rowNumber: 9,
                columnNumber: 2,
            })
        );

        assert.deepStrictEqual(state.selectedCell, null);
        assert.deepStrictEqual(state.selectionAnchorCell, null);
        assert.deepStrictEqual(state.selectionRangeOverride, null);
        assert.strictEqual(state.suppressAutoSelection, true);
    });

    test("tracks pending selection until it is cleared", () => {
        const pendingState = setPendingSelectionAfterRender(createSelectionControllerState(), {
            rowNumber: 12,
            columnNumber: 8,
            reveal: true,
        });
        assert.deepStrictEqual(pendingState.pendingSelectionAfterRender, {
            rowNumber: 12,
            columnNumber: 8,
            reveal: true,
        });

        const clearedState = clearPendingSelectionAfterRender(pendingState);
        assert.strictEqual(clearedState.pendingSelectionAfterRender, null);
    });

    test("stops drag only when the pointer matches", () => {
        const draggingState = startCellSelectionDrag(
            setSelectedCell(createSelectionControllerState(), {
                rowNumber: 3,
                columnNumber: 3,
            }),
            42,
            { rowNumber: 3, columnNumber: 3 }
        );

        assert.deepStrictEqual(
            stopSelectionDrag(draggingState, 7).selectionDragState,
            draggingState.selectionDragState
        );
        assert.strictEqual(stopSelectionDrag(draggingState, 42).selectionDragState, null);
    });

    test("can re-sync the anchor to the current focused cell", () => {
        const state = syncSelectionAnchorToSelectedCell(
            setSelectedCell(
                createSelectionControllerState(),
                { rowNumber: 6, columnNumber: 4 },
                {
                    anchorCell: { rowNumber: 2, columnNumber: 1 },
                }
            )
        );

        assert.deepStrictEqual(state.selectionAnchorCell, { rowNumber: 6, columnNumber: 4 });
    });
});
