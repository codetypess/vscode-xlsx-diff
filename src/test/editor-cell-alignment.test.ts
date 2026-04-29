/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    getCellContentAlignmentStyle,
    getToolbarHorizontalAlignment,
    getToolbarVerticalAlignment,
} from "../webview/editor-panel/editor-cell-alignment";

suite("Editor cell alignment helpers", () => {
    test("maps centered alignments to a full-height content style", () => {
        assert.deepStrictEqual(
            getCellContentAlignmentStyle({
                horizontal: "center",
                vertical: "center",
            }),
            {
                justifyContent: "center",
                alignItems: "center",
                textAlign: "center",
                height: "100%",
                maxHeight: "100%",
            }
        );
    });

    test("maps bottom-right alignment correctly", () => {
        assert.deepStrictEqual(
            getCellContentAlignmentStyle({
                horizontal: "right",
                vertical: "bottom",
            }),
            {
                justifyContent: "flex-end",
                alignItems: "flex-end",
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
        assert.strictEqual(getToolbarVerticalAlignment({ vertical: "bottom" }), "bottom");
        assert.strictEqual(getToolbarHorizontalAlignment(null), undefined);
        assert.strictEqual(getToolbarVerticalAlignment(null), undefined);
    });
});
