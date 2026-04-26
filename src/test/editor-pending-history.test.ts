/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { rebasePendingHistory } from "../webview/editor-panel/editor-pending-history";

suite("Editor pending history", () => {
    test("rebases saved cells in undo and redo history to the current saved value", () => {
        const rebased = rebasePendingHistory(
            [
                {
                    changes: [
                        {
                            sheetKey: "sheet:0",
                            rowNumber: 1,
                            columnNumber: 1,
                            modelValue: "before",
                            beforeValue: "before",
                            afterValue: "mid",
                        },
                        {
                            sheetKey: "sheet:1",
                            rowNumber: 2,
                            columnNumber: 3,
                            modelValue: "base",
                            beforeValue: "base",
                            afterValue: "saved",
                        },
                    ],
                },
            ],
            [
                {
                    changes: [
                        {
                            sheetKey: "sheet:0",
                            rowNumber: 1,
                            columnNumber: 1,
                            modelValue: "before",
                            beforeValue: "mid",
                            afterValue: "after",
                        },
                    ],
                },
            ],
            [
                {
                    sheetKey: "sheet:0",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "mid",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "saved",
                },
            ]
        );

        assert.strictEqual(rebased.undoStack[0]!.changes[0]!.modelValue, "mid");
        assert.strictEqual(rebased.redoStack[0]!.changes[0]!.modelValue, "mid");
        assert.strictEqual(rebased.undoStack[0]!.changes[1]!.modelValue, "saved");
    });

    test("keeps the original model value when a cell is back on the saved baseline", () => {
        const rebased = rebasePendingHistory(
            [
                {
                    changes: [
                        {
                            sheetKey: "sheet:0",
                            rowNumber: 4,
                            columnNumber: 2,
                            modelValue: "original",
                            beforeValue: "original",
                            afterValue: "changed",
                        },
                    ],
                },
            ],
            [],
            []
        );

        assert.strictEqual(rebased.undoStack[0]!.changes[0]!.modelValue, "original");
    });
});
