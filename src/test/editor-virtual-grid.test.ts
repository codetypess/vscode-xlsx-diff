/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    EDITOR_VIRTUAL_COLUMN_WIDTH,
    EDITOR_VIRTUAL_HEADER_HEIGHT,
    EDITOR_VIRTUAL_ROW_HEIGHT,
    clampEditorScrollPosition,
    createEditorColumnWindow,
    createEditorRowWindow,
    getEditorContentSize,
    getEditorRowHeaderWidth,
    getEditorScrollPositionForCell,
    getFrozenEditorCounts,
    getVisibleFrozenEditorCounts,
} from "../webview/editor-panel/editor-virtual-grid";

suite("Editor virtual grid helpers", () => {
    test("creates a scrollable row window below frozen rows", () => {
        const window = createEditorRowWindow({
            totalRows: 1000,
            frozenRowCount: 2,
            scrollTop: 28 * 100,
            viewportHeight: 600,
        });

        assert.strictEqual(window.startRowNumber, 63);
        assert.strictEqual(window.rowNumbers[0], 63);
        assert.ok(window.endRowNumber > window.startRowNumber);
        assert.strictEqual(window.topSpacerHeight, 60 * EDITOR_VIRTUAL_ROW_HEIGHT);
    });

    test("creates a scrollable column window after frozen columns", () => {
        const window = createEditorColumnWindow({
            totalColumns: 200,
            frozenColumnCount: 1,
            scrollLeft: 120 * 20,
            viewportWidth: 900,
            rowHeaderWidth: 56,
        });

        assert.strictEqual(window.startColumnNumber, 14);
        assert.strictEqual(window.columnNumbers[0], 14);
        assert.ok(window.trailingSpacerWidth > 0);
    });

    test("computes content size and target scroll positions for a cell", () => {
        const rowHeaderWidth = getEditorRowHeaderWidth(1000);
        const contentSize = getEditorContentSize({
            rowCount: 1000,
            columnCount: 50,
            rowHeaderWidth,
        });
        const target = getEditorScrollPositionForCell({
            rowNumber: 50,
            columnNumber: 10,
            frozenRowCount: 2,
            frozenColumnCount: 1,
            viewportHeight: 600,
            viewportWidth: 900,
            rowHeaderWidth,
        });

        assert.strictEqual(
            contentSize.width,
            rowHeaderWidth + 50 * EDITOR_VIRTUAL_COLUMN_WIDTH
        );
        assert.strictEqual(
            contentSize.height,
            EDITOR_VIRTUAL_HEADER_HEIGHT + 1000 * EDITOR_VIRTUAL_ROW_HEIGHT
        );
        assert.ok((target.top ?? 0) > 0);
        assert.ok((target.left ?? 0) > 0);
    });

    test("clamps freeze counts and scroll positions safely", () => {
        assert.deepStrictEqual(
            getFrozenEditorCounts({
                rowCount: 10,
                columnCount: 5,
                freezePane: { rowCount: 100, columnCount: 100 },
            }),
            { rowCount: 9, columnCount: 4 }
        );
        assert.strictEqual(clampEditorScrollPosition(-10, 200), 0);
        assert.strictEqual(clampEditorScrollPosition(250, 200), 200);
    });

    test("clips rendered frozen panes to the current viewport", () => {
        assert.deepStrictEqual(
            getVisibleFrozenEditorCounts({
                frozenRowCount: 200,
                frozenColumnCount: 20,
                viewportHeight: 600,
                viewportWidth: 900,
                rowHeaderWidth: 56,
            }),
            {
                rowCount: 20,
                columnCount: 7,
            }
        );
    });
});
