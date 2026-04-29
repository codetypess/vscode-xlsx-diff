/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    getCellContentAlignmentStyle,
    getToolbarHorizontalAlignment,
} from "../webview/editor-panel/editor-cell-alignment";

suite("Editor cell alignment helpers", () => {
    test("maps horizontal alignment to a fixed-height content style", () => {
        assert.deepStrictEqual(
            getCellContentAlignmentStyle({
                horizontal: "center",
            }),
            {
                justifyContent: "center",
                alignItems: "flex-start",
                textAlign: "center",
                height: "100%",
                maxHeight: "100%",
            }
        );
    });

    test("ignores vertical alignment and keeps horizontal right alignment", () => {
        assert.deepStrictEqual(
            getCellContentAlignmentStyle({
                horizontal: "right",
                vertical: "bottom",
            }),
            {
                justifyContent: "flex-end",
                alignItems: "flex-start",
                textAlign: "right",
                height: "100%",
                maxHeight: "100%",
            }
        );
    });

    test("normalizes toolbar alignment state", () => {
        assert.strictEqual(
            getToolbarHorizontalAlignment({ horizontal: "centerContinuous" }),
            "center"
        );
        assert.strictEqual(getToolbarHorizontalAlignment(null), undefined);
    });
});
