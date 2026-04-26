/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getDiffRowHeaderWidth } from "../webview/diff-panel/diff-grid-layout";

suite("Diff grid layout helpers", () => {
    test("keeps the default width for up to three-digit row numbers", () => {
        assert.strictEqual(getDiffRowHeaderWidth(999), 56);
    });

    test("expands the width when four-digit row numbers need room for a diff marker", () => {
        assert.strictEqual(getDiffRowHeaderWidth(1018), 65);
    });
});
