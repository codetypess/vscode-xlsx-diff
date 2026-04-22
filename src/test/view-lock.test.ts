/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getFreezePaneCountsForCell, hasLockedView } from "../webview/view-lock";

suite("View lock helpers", () => {
    test("treats only positive freeze counts as a locked view", () => {
        assert.strictEqual(hasLockedView(null), false);
        assert.strictEqual(
            hasLockedView({
                columnCount: 0,
                rowCount: 0,
                topLeftCell: "A1",
                activePane: null,
            }),
            false
        );
        assert.strictEqual(
            hasLockedView({
                columnCount: 1,
                rowCount: 0,
                topLeftCell: "B1",
                activePane: "topRight",
            }),
            true
        );
    });

    test("freezes rows above and columns left of the selected cell", () => {
        assert.deepStrictEqual(getFreezePaneCountsForCell({ rowNumber: 1, columnNumber: 1 }), {
            rowCount: 0,
            columnCount: 0,
        });
        assert.deepStrictEqual(getFreezePaneCountsForCell({ rowNumber: 2, columnNumber: 2 }), {
            rowCount: 1,
            columnCount: 1,
        });
        assert.deepStrictEqual(getFreezePaneCountsForCell({ rowNumber: 1, columnNumber: 4 }), {
            rowCount: 0,
            columnCount: 3,
        });
    });
});
