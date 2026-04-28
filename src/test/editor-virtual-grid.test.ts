/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    convertPixelsToWorkbookColumnWidth,
    convertWorkbookColumnWidthToPixels,
    stabilizeColumnPixelWidth,
} from "../webview/column-layout";
import {
    EDITOR_VIRTUAL_COLUMN_WIDTH,
    EDITOR_VIRTUAL_HEADER_HEIGHT,
    EDITOR_VIRTUAL_ROW_HEIGHT,
    clampEditorScrollPosition,
    createEditorColumnWindow,
    createEditorPixelColumnLayout,
    createEditorRowWindow,
    getEditorDisplayColumnLayout,
    getEditorDisplayGridDimensions,
    getEditorFrozenColumnsWidth,
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
        const sheetColumnLayout = createEditorPixelColumnLayout({ columnCount: 200 });
        const columnLayout = getEditorDisplayColumnLayout(sheetColumnLayout, 200);
        const window = createEditorColumnWindow({
            columnLayout,
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
        const sheetColumnLayout = createEditorPixelColumnLayout({ columnCount: 50 });
        const columnLayout = getEditorDisplayColumnLayout(sheetColumnLayout, 50);
        const contentSize = getEditorContentSize({
            rowCount: 1000,
            columnLayout,
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
            columnLayout,
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
        const columnLayout = createEditorPixelColumnLayout({ columnCount: 20 });
        assert.deepStrictEqual(
            getVisibleFrozenEditorCounts({
                frozenRowCount: 200,
                frozenColumnCount: 20,
                viewportHeight: 600,
                viewportWidth: 900,
                rowHeaderWidth: 56,
                columnLayout,
            }),
            {
                rowCount: 20,
                columnCount: 8,
            }
        );
    });

    test("pads displayed rows and columns to fill the viewport", () => {
        const columnLayout = createEditorPixelColumnLayout({ columnCount: 3 });
        assert.deepStrictEqual(
            getEditorDisplayGridDimensions({
                rowCount: 5,
                columnCount: 3,
                viewportHeight: 600,
                viewportWidth: 960,
                columnLayout,
            }),
            {
                rowCount: 29,
                columnCount: 11,
                rowHeaderWidth: 56,
            }
        );
    });

    test("adds extra editable rows and columns even when the sheet already fills the viewport", () => {
        const columnLayout = createEditorPixelColumnLayout({ columnCount: 12 });
        assert.deepStrictEqual(
            getEditorDisplayGridDimensions({
                rowCount: 40,
                columnCount: 12,
                viewportHeight: 600,
                viewportWidth: 960,
                columnLayout,
            }),
            {
                rowCount: 48,
                columnCount: 15,
                rowHeaderWidth: 56,
            }
        );
    });

    test("honors workbook-backed variable column widths", () => {
        const sheetColumnLayout = createEditorPixelColumnLayout({
            columnCount: 4,
            columnWidths: [8.7109375, 20, null, 12],
            maximumDigitWidth: 7,
        });
        const columnLayout = getEditorDisplayColumnLayout(sheetColumnLayout, 4);

        assert.strictEqual(getEditorFrozenColumnsWidth(columnLayout, 2), 201);
        assert.strictEqual(
            getEditorDisplayGridDimensions({
                rowCount: 10,
                columnCount: 4,
                viewportHeight: 600,
                viewportWidth: 960,
                columnLayout: sheetColumnLayout,
            }).columnCount,
            12
        );
    });

    test("normalizes drag-resized pixel widths to stable workbook-backed sizes", () => {
        const maximumDigitWidth = 7;

        for (const pixelWidth of [40, 61, 84, 120, 140, 240]) {
            const stabilizedPixelWidth = stabilizeColumnPixelWidth(pixelWidth, maximumDigitWidth);
            const workbookWidth = convertPixelsToWorkbookColumnWidth(
                stabilizedPixelWidth,
                maximumDigitWidth
            );
            assert.strictEqual(
                convertWorkbookColumnWidthToPixels(workbookWidth, maximumDigitWidth),
                stabilizedPixelWidth
            );
        }
    });
});
