/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    createColumnSelectionRange,
    createRowSelectionRange,
    createSelectionRange,
    hasExpandedSelectionRange,
} from "../webview/editor-panel/editor-selection-range";

suite("Editor selection range helpers", () => {
    test("builds a rectangular range from two cells", () => {
        assert.deepStrictEqual(
            createSelectionRange(
                { rowNumber: 8, columnNumber: 5 },
                { rowNumber: 3, columnNumber: 2 }
            ),
            {
                startRow: 3,
                endRow: 8,
                startColumn: 2,
                endColumn: 5,
            }
        );
    });

    test("detects when a selection range expands beyond one cell", () => {
        assert.strictEqual(
            hasExpandedSelectionRange({
                startRow: 5,
                endRow: 5,
                startColumn: 2,
                endColumn: 7,
            }),
            true
        );
        assert.strictEqual(
            hasExpandedSelectionRange({
                startRow: 5,
                endRow: 5,
                startColumn: 2,
                endColumn: 2,
            }),
            false
        );
    });

    test("creates full-row ranges from row headers", () => {
        assert.deepStrictEqual(createRowSelectionRange(6, 12), {
            startRow: 6,
            endRow: 6,
            startColumn: 1,
            endColumn: 12,
        });
        assert.strictEqual(createRowSelectionRange(0, 12), null);
    });

    test("creates full-column ranges from column headers", () => {
        assert.deepStrictEqual(createColumnSelectionRange(4, 128), {
            startRow: 1,
            endRow: 128,
            startColumn: 4,
            endColumn: 4,
        });
        assert.strictEqual(createColumnSelectionRange(4, 0), null);
    });
});
