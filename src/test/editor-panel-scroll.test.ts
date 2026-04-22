/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getApproximateViewportStartRowForScrollPosition } from "../webview/editor-panel-scroll";

suite("Editor panel scroll helpers", () => {
    test("keeps the viewport at the base row when scrolling is unnecessary", () => {
        assert.strictEqual(
            getApproximateViewportStartRowForScrollPosition({
                baseRow: 1,
                maxStartRow: 1,
                totalScrollableRows: 40,
                visibleRowCount: 40,
                scrollTop: 0,
                maxScrollTop: 0,
            }),
            1
        );
    });

    test("maps a dragged scrollbar position to a later viewport window", () => {
        assert.strictEqual(
            getApproximateViewportStartRowForScrollPosition({
                baseRow: 1,
                maxStartRow: 801,
                totalScrollableRows: 1000,
                visibleRowCount: 200,
                scrollTop: 4000,
                maxScrollTop: 8000,
            }),
            461
        );
    });

    test("respects frozen-row offsets and clamps near the end of the sheet", () => {
        assert.strictEqual(
            getApproximateViewportStartRowForScrollPosition({
                baseRow: 6,
                maxStartRow: 806,
                totalScrollableRows: 995,
                visibleRowCount: 200,
                scrollTop: 9000,
                maxScrollTop: 9000,
            }),
            806
        );
    });
});
