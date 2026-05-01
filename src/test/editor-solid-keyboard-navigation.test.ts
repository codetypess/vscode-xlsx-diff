/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { EditorActiveSheetView, EditorSelectionView } from "../core/model/types";
import {
    getEditorKeyboardNavigationDelta,
    getEditorKeyboardPageDirection,
    getNextEditorKeyboardNavigationTarget,
    getNextEditorViewportPageNavigationTarget,
    isEditorClearCellKey,
} from "../webview-solid/editor-panel/keyboard-navigation-helpers";

function createActiveSheet(): EditorActiveSheetView {
    return {
        key: "sheet:1",
        rowCount: 10,
        columnCount: 8,
        columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
        cells: {},
        freezePane: null,
        autoFilter: null,
    };
}

function createSelection(rowNumber: number, columnNumber: number): EditorSelectionView {
    return {
        key: `${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        address: `${String.fromCharCode(64 + columnNumber)}${rowNumber}`,
        value: "",
        formula: null,
        isPresent: false,
    };
}

suite("Solid editor keyboard navigation", () => {
    test("maps plain navigation keys to deltas", () => {
        assert.deepStrictEqual(getEditorKeyboardNavigationDelta("ArrowUp"), {
            rowDelta: -1,
            columnDelta: 0,
        });
        assert.deepStrictEqual(getEditorKeyboardNavigationDelta("ArrowRight"), {
            rowDelta: 0,
            columnDelta: 1,
        });
        assert.deepStrictEqual(getEditorKeyboardNavigationDelta("Tab"), {
            rowDelta: 0,
            columnDelta: 1,
        });
        assert.deepStrictEqual(getEditorKeyboardNavigationDelta("Enter"), {
            rowDelta: 1,
            columnDelta: 0,
        });
        assert.strictEqual(getEditorKeyboardNavigationDelta("Escape"), null);
    });

    test("recognizes clear-cell keys", () => {
        assert.strictEqual(isEditorClearCellKey("Backspace"), true);
        assert.strictEqual(isEditorClearCellKey("Delete"), true);
        assert.strictEqual(isEditorClearCellKey("Enter"), false);
    });

    test("maps page keys to directions", () => {
        assert.strictEqual(getEditorKeyboardPageDirection("PageUp"), -1);
        assert.strictEqual(getEditorKeyboardPageDirection("PageDown"), 1);
        assert.strictEqual(getEditorKeyboardPageDirection("ArrowDown"), null);
    });

    test("clamps navigation targets to sheet bounds", () => {
        assert.deepStrictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(1, 1),
                delta: { rowDelta: -1, columnDelta: -1 },
            }),
            { rowNumber: 1, columnNumber: 1 }
        );
        assert.deepStrictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(10, 8),
                delta: { rowDelta: 1, columnDelta: 1 },
            }),
            { rowNumber: 10, columnNumber: 8 }
        );
    });

    test("falls back to A1 when there is no current selection", () => {
        assert.deepStrictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: null,
                delta: { rowDelta: 0, columnDelta: 0 },
            }),
            { rowNumber: 1, columnNumber: 1 }
        );
    });

    test("moves through visible filtered rows in display order", () => {
        assert.deepStrictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(4, 2),
                delta: { rowDelta: 1, columnDelta: 0 },
                visibleRowNumbers: [1, 4, 7],
            }),
            { rowNumber: 7, columnNumber: 2 }
        );
        assert.deepStrictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(4, 2),
                delta: { rowDelta: -1, columnDelta: 0 },
                visibleRowNumbers: [1, 7, 4],
            }),
            { rowNumber: 7, columnNumber: 2 }
        );
    });

    test("returns null for empty sheets", () => {
        assert.strictEqual(
            getNextEditorKeyboardNavigationTarget({
                activeSheet: {
                    ...createActiveSheet(),
                    rowCount: 0,
                },
                selection: createSelection(1, 1),
                delta: { rowDelta: 1, columnDelta: 0 },
            }),
            null
        );
    });

    test("moves by the visible page size while preserving the current column", () => {
        assert.deepStrictEqual(
            getNextEditorViewportPageNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(4, 3),
                direction: 1,
                visibleRowCount: 5,
            }),
            { rowNumber: 8, columnNumber: 3 }
        );
        assert.deepStrictEqual(
            getNextEditorViewportPageNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(2, 3),
                direction: -1,
                visibleRowCount: 5,
            }),
            { rowNumber: 1, columnNumber: 3 }
        );
    });

    test("moves by page through filtered rows", () => {
        assert.deepStrictEqual(
            getNextEditorViewportPageNavigationTarget({
                activeSheet: createActiveSheet(),
                selection: createSelection(1, 3),
                direction: 1,
                visibleRowCount: 3,
                visibleRowNumbers: [1, 4, 9, 2],
            }),
            { rowNumber: 9, columnNumber: 3 }
        );
    });
});
