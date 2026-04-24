/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getSelectionPreviewInlineDiff } from "../webview/selection-preview-diff";

suite("Selection preview inline diff", () => {
    test("keeps identical text unhighlighted", () => {
        assert.deepStrictEqual(getSelectionPreviewInlineDiff("same", "same"), {
            before: "same",
            changed: "",
            after: "",
            hasDifference: false,
        });
    });

    test("highlights the changed middle segment", () => {
        assert.deepStrictEqual(getSelectionPreviewInlineDiff("abc123xyz", "abc456xyz"), {
            before: "abc",
            changed: "123",
            after: "xyz",
            hasDifference: true,
        });
    });

    test("handles inserted suffix text on the other side", () => {
        assert.deepStrictEqual(getSelectionPreviewInlineDiff("abc", "abcd"), {
            before: "abc",
            changed: "",
            after: "",
            hasDifference: true,
        });
    });

    test("highlights a fully different value", () => {
        assert.deepStrictEqual(getSelectionPreviewInlineDiff("left", "down"), {
            before: "",
            changed: "left",
            after: "",
            hasDifference: true,
        });
    });
});
