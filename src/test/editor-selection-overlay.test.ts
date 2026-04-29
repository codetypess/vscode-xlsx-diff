/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { clipSelectionRangeToVisibleGrid } from "../webview/editor-panel/editor-selection-overlay";

suite("Editor selection overlay helpers", () => {
    test("clips a selection to the visible rows and columns", () => {
        assert.deepStrictEqual(
            clipSelectionRangeToVisibleGrid(
                {
                    startRow: 2,
                    endRow: 8,
                    startColumn: 3,
                    endColumn: 9,
                },
                [1, 2, 4, 7, 10],
                [1, 3, 5, 6, 9]
            ),
            {
                startRow: 2,
                endRow: 7,
                startColumn: 3,
                endColumn: 9,
            }
        );
    });

    test("returns null when the selection is fully outside visible rows", () => {
        assert.strictEqual(
            clipSelectionRangeToVisibleGrid(
                {
                    startRow: 20,
                    endRow: 30,
                    startColumn: 2,
                    endColumn: 4,
                },
                [1, 2, 4, 7, 10],
                [1, 3, 5, 6, 9]
            ),
            null
        );
    });

    test("returns null when the selection is fully outside visible columns", () => {
        assert.strictEqual(
            clipSelectionRangeToVisibleGrid(
                {
                    startRow: 2,
                    endRow: 8,
                    startColumn: 20,
                    endColumn: 30,
                },
                [1, 2, 4, 7, 10],
                [1, 3, 5, 6, 9]
            ),
            null
        );
    });
});
