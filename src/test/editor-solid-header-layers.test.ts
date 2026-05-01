/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    createInitialEditorGridViewportState,
    deriveEditorGridMetrics,
} from "../webview-solid/editor-panel/grid-foundation";
import {
    createEditorColumnHeaderSelection,
    createEditorRowHeaderSelection,
    deriveEditorGridHeaderLayers,
} from "../webview-solid/editor-panel/header-layer-helpers";

suite("Solid editor header layers", () => {
    test("derives scrollable and frozen row and column headers", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 200,
                columnCount: 20,
                freezePane: {
                    rowCount: 2,
                    columnCount: 2,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
                scrollTop: 120,
                scrollLeft: 240,
            })
        );

        const layers = deriveEditorGridHeaderLayers({
            metrics,
            columnLabels: ["A", "B", "C", "D"],
            selectedRowNumber: 5,
            selectedColumnNumber: 4,
            selectionRange: null,
        });

        assert.strictEqual(layers.headerHeight, 28);
        assert.strictEqual(layers.rowHeaderWidth, 56);
        assert.deepStrictEqual(
            layers.frozenColumnHeaders.map((item) => item.columnNumber),
            [1, 2]
        );
        assert.deepStrictEqual(
            layers.frozenRowHeaders.map((item) => item.rowNumber),
            [1, 2]
        );
        assert.ok(layers.scrollableColumnHeaders.length > 0);
        assert.ok(layers.scrollableRowHeaders.length > 0);
        assert.strictEqual(
            layers.scrollableColumnHeaders.find((item) => item.columnNumber === 4)?.isActive,
            true
        );
        assert.strictEqual(
            layers.scrollableRowHeaders.find((item) => item.rowNumber === 5)?.isActive,
            true
        );
        assert.strictEqual(layers.frozenColumnHeaders[0]?.label, "A");
        assert.ok((layers.scrollableRowHeaders[0]?.top ?? 0) > 28);
    });

    test("falls back to generated column labels for padded columns", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 5,
                columnCount: 1,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 960,
            })
        );

        const layers = deriveEditorGridHeaderLayers({
            metrics,
            columnLabels: ["A"],
            selectedRowNumber: null,
            selectedColumnNumber: null,
            selectionRange: null,
        });

        assert.strictEqual(layers.scrollableColumnHeaders[1]?.label, "B");
    });

    test("creates row and column header selection anchors", () => {
        assert.deepStrictEqual(createEditorRowHeaderSelection(8, 3), {
            rowNumber: 8,
            columnNumber: 3,
        });
        assert.deepStrictEqual(createEditorRowHeaderSelection(8, null), {
            rowNumber: 8,
            columnNumber: 1,
        });
        assert.deepStrictEqual(createEditorColumnHeaderSelection(5, 4), {
            rowNumber: 4,
            columnNumber: 5,
        });
        assert.deepStrictEqual(createEditorColumnHeaderSelection(5, null), {
            rowNumber: 1,
            columnNumber: 5,
        });
    });

    test("marks row and column headers active across an expanded selection range", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 20,
                columnCount: 10,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 420,
                viewportWidth: 640,
            })
        );

        const layers = deriveEditorGridHeaderLayers({
            metrics,
            columnLabels: ["A", "B", "C", "D"],
            selectedRowNumber: 4,
            selectedColumnNumber: 3,
            selectionRange: {
                startRow: 4,
                endRow: 6,
                startColumn: 3,
                endColumn: 5,
            },
        });

        assert.strictEqual(
            layers.scrollableRowHeaders.find((item) => item.rowNumber === 5)?.isActive,
            true
        );
        assert.strictEqual(
            layers.scrollableColumnHeaders.find((item) => item.columnNumber === 4)?.isActive,
            true
        );
    });
});
