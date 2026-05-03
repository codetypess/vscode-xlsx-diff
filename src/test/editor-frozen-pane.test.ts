/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    createInitialEditorGridViewportState,
    deriveEditorGridMetrics,
} from "../webview/editor-panel/grid-foundation";
import { deriveEditorFrozenPaneOverlayLayout } from "../webview/editor-panel/frozen-pane-helpers";

suite("Editor frozen pane helpers", () => {
    test("hides frozen overlays when the sheet has no frozen panes", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 40,
                columnCount: 12,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 960,
            })
        );

        assert.deepStrictEqual(deriveEditorFrozenPaneOverlayLayout(metrics), {
            headerHeight: 28,
            rowHeaderWidth: 56,
            frozenRowsHeight: 0,
            frozenColumnsWidth: 0,
            showTopOverlay: false,
            showLeftOverlay: false,
            showCornerOverlay: false,
        });
    });

    test("derives top, left, and corner overlay geometry from frozen metrics", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 200,
                columnCount: 20,
                freezePane: {
                    rowCount: 3,
                    columnCount: 2,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
            })
        );

        const layout = deriveEditorFrozenPaneOverlayLayout(metrics);
        assert.strictEqual(layout.headerHeight, 28);
        assert.strictEqual(layout.rowHeaderWidth, 56);
        assert.ok(layout.frozenRowsHeight > 0);
        assert.ok(layout.frozenColumnsWidth > 0);
        assert.strictEqual(layout.showTopOverlay, true);
        assert.strictEqual(layout.showLeftOverlay, true);
        assert.strictEqual(layout.showCornerOverlay, true);
    });
});
