/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    getToolbarCellEditTargetKey,
    shouldResetToolbarCellValueDraft,
} from "../webview/editor-toolbar-input";

suite("Editor toolbar input helpers", () => {
    test("keeps the draft while editing the same cell", () => {
        const target = {
            sheetKey: "sheet:1",
            rowNumber: 5,
            columnNumber: 4,
        };

        assert.strictEqual(shouldResetToolbarCellValueDraft(target, target, true), false);
        assert.strictEqual(getToolbarCellEditTargetKey(target), "sheet:1:5:4");
    });

    test("resets the draft after the active cell changes", () => {
        assert.strictEqual(
            shouldResetToolbarCellValueDraft(
                {
                    sheetKey: "sheet:1",
                    rowNumber: 5,
                    columnNumber: 4,
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 10,
                    columnNumber: 6,
                },
                true
            ),
            true
        );
    });

    test("resets the draft when the cell is no longer editable", () => {
        assert.strictEqual(
            shouldResetToolbarCellValueDraft(
                {
                    sheetKey: "sheet:1",
                    rowNumber: 5,
                    columnNumber: 4,
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 5,
                    columnNumber: 4,
                },
                false
            ),
            true
        );
    });
});
