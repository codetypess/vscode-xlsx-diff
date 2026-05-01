/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    DEFAULT_SOLID_EDITOR_VIEWPORT_HEIGHT,
    DEFAULT_SOLID_EDITOR_VIEWPORT_WIDTH,
    applyEditorGridViewportPatch,
    areEditorGridMetricsEqual,
    createInitialEditorGridViewportState,
    deriveEditorGridMetrics,
    getEditorDisplayRowNumber,
    getEditorGridActualRowTop,
    reuseEquivalentEditorGridMetrics,
} from "../webview-solid/editor-panel/grid-foundation";

suite("Solid editor grid foundation", () => {
    test("creates and patches viewport state with clamped values", () => {
        const initialState = createInitialEditorGridViewportState({
            scrollTop: -10,
            viewportWidth: 0,
        });

        assert.deepStrictEqual(initialState, {
            scrollTop: 0,
            scrollLeft: 0,
            viewportHeight: DEFAULT_SOLID_EDITOR_VIEWPORT_HEIGHT,
            viewportWidth: 0,
        });

        assert.deepStrictEqual(
            applyEditorGridViewportPatch(initialState, {
                scrollLeft: 42,
                scrollTop: -5,
                viewportHeight: 720,
            }),
            {
                scrollTop: 0,
                scrollLeft: 42,
                viewportHeight: 720,
                viewportWidth: 0,
            }
        );
    });

    test("derives padded display metrics for a small sheet", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 5,
                columnCount: 3,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: DEFAULT_SOLID_EDITOR_VIEWPORT_WIDTH,
            })
        );

        assert.strictEqual(metrics.displayRowCount, 29);
        assert.strictEqual(metrics.displayColumnCount, 11);
        assert.strictEqual(metrics.rowHeaderWidth, 56);
        assert.strictEqual(metrics.window.frozenRowNumbers.length, 0);
        assert.strictEqual(metrics.window.frozenColumnNumbers.length, 0);
        assert.ok(metrics.window.rowNumbers.length > 0);
        assert.ok(metrics.window.columnNumbers.length > 0);
    });

    test("clips visible frozen panes to the current viewport", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 200,
                columnCount: 20,
                freezePane: {
                    rowCount: 200,
                    columnCount: 20,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
            })
        );

        assert.strictEqual(metrics.frozenRowCount, 199);
        assert.strictEqual(metrics.frozenColumnCount, 19);
        assert.strictEqual(metrics.visibleFrozenRowCount, 21);
        assert.strictEqual(metrics.visibleFrozenColumnCount, 8);
        assert.deepStrictEqual(metrics.window.frozenRowNumbers.slice(0, 3), [1, 2, 3]);
        assert.deepStrictEqual(metrics.window.frozenColumnNumbers.slice(0, 3), [1, 2, 3]);
    });

    test("maps filtered visible rows to display rows", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 10,
                columnCount: 3,
                visibleRows: [1, 4, 7],
                hiddenRows: [2, 3, 5, 6],
                rowHeights: {
                    "4": 48,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 180,
                viewportWidth: 420,
            })
        );

        assert.deepStrictEqual(metrics.window.rowNumbers.slice(0, 3), [1, 4, 7]);
        assert.strictEqual(getEditorDisplayRowNumber(metrics, 4), 2);
        assert.strictEqual(getEditorGridActualRowTop(metrics, 4), 28);
        assert.strictEqual(metrics.rowLayout.overriddenPixelHeights["2"], 84);
        assert.deepStrictEqual(metrics.rowState.hiddenActualRowNumbers, [2, 3, 5, 6]);
    });

    test("reuses equivalent metrics when the rendered window stays unchanged", () => {
        const baseInput = {
            rowCount: 1000,
            columnCount: 200,
        };
        const first = deriveEditorGridMetrics(
            baseInput,
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
                scrollTop: 2800,
                scrollLeft: 2400,
            })
        );
        const second = deriveEditorGridMetrics(
            baseInput,
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
                scrollTop: 2810,
                scrollLeft: 2400,
            })
        );

        assert.strictEqual(areEditorGridMetricsEqual(first, second), true);
        assert.strictEqual(reuseEquivalentEditorGridMetrics(first, second), first);
    });

    test("replaces metrics when the rendered window changes", () => {
        const baseInput = {
            rowCount: 1000,
            columnCount: 200,
        };
        const first = deriveEditorGridMetrics(
            baseInput,
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
                scrollTop: 2800,
                scrollLeft: 2400,
            })
        );
        const second = deriveEditorGridMetrics(
            baseInput,
            createInitialEditorGridViewportState({
                viewportHeight: 600,
                viewportWidth: 900,
                scrollTop: 5600,
                scrollLeft: 4800,
            })
        );

        assert.strictEqual(areEditorGridMetricsEqual(first, second), false);
        assert.strictEqual(reuseEquivalentEditorGridMetrics(first, second), second);
    });
});
