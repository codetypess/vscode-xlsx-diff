/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey } from "../core/model/cells";
import {
    createInitialEditorGridViewportState,
    deriveEditorGridMetrics,
} from "../webview/editor-panel/grid-foundation";
import { deriveEditorSelectionOverlayLayers } from "../webview/editor-panel/selection-overlay-helpers";

suite("Editor selection overlays", () => {
    test("derives active row, active column, range, and primary rects for a visible body selection", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 100,
                columnCount: 20,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 420,
                viewportWidth: 640,
                scrollTop: 84,
                scrollLeft: 120,
            })
        );

        const overlays = deriveEditorSelectionOverlayLayers({
            metrics,
            selection: {
                key: createCellKey(5, 4),
                rowNumber: 5,
                columnNumber: 4,
                address: "D5",
                value: "body",
                formula: null,
                isPresent: true,
            },
            selectionRangeOverride: null,
        });

        assert.ok(overlays.body.activeRowRect);
        assert.ok(overlays.body.activeColumnRect);
        assert.ok(overlays.body.rangeRect);
        assert.ok(overlays.body.primaryRect);
        assert.strictEqual(overlays.top.primaryRect, null);
        assert.strictEqual(overlays.left.primaryRect, null);
        assert.strictEqual(overlays.corner.primaryRect, null);
    });

    test("assigns the primary rect to the frozen corner layer for frozen selections", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 40,
                columnCount: 12,
                freezePane: {
                    rowCount: 2,
                    columnCount: 2,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 420,
                viewportWidth: 640,
            })
        );

        const overlays = deriveEditorSelectionOverlayLayers({
            metrics,
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "corner",
                formula: null,
                isPresent: true,
            },
            selectionRangeOverride: null,
        });

        assert.ok(overlays.corner.primaryRect);
        assert.ok(overlays.corner.rangeRect);
        assert.ok(overlays.top.activeRowRect);
        assert.ok(overlays.left.activeColumnRect);
        assert.strictEqual(overlays.body.primaryRect, null);
    });

    test("returns empty overlay layers when the selected cell is outside the visible windows", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 40,
                columnCount: 12,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 180,
                viewportWidth: 260,
            })
        );

        const overlays = deriveEditorSelectionOverlayLayers({
            metrics,
            selection: {
                key: createCellKey(40, 12),
                rowNumber: 40,
                columnNumber: 12,
                address: "L40",
                value: "far",
                formula: null,
                isPresent: true,
            },
            selectionRangeOverride: null,
        });

        assert.strictEqual(overlays.body.primaryRect, null);
        assert.strictEqual(overlays.body.activeRowRect, null);
        assert.strictEqual(overlays.body.activeColumnRect, null);
        assert.strictEqual(overlays.top.primaryRect, null);
        assert.strictEqual(overlays.left.primaryRect, null);
        assert.strictEqual(overlays.corner.primaryRect, null);
    });

    test("derives an expanded range rect when a selection range override is present", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 100,
                columnCount: 20,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 420,
                viewportWidth: 640,
                scrollTop: 84,
                scrollLeft: 120,
            })
        );

        const overlays = deriveEditorSelectionOverlayLayers({
            metrics,
            selection: {
                key: createCellKey(6, 5),
                rowNumber: 6,
                columnNumber: 5,
                address: "E6",
                value: "focus",
                formula: null,
                isPresent: true,
            },
            selectionRangeOverride: {
                startRow: 4,
                endRow: 6,
                startColumn: 3,
                endColumn: 5,
            },
        });

        assert.ok(overlays.body.rangeRect);
        assert.ok((overlays.body.rangeRect?.width ?? 0) > (overlays.body.primaryRect?.width ?? 0));
        assert.strictEqual(overlays.body.rangeRect?.showTopBorder, true);
        assert.strictEqual(overlays.body.rangeRect?.showRightBorder, true);
    });
});
