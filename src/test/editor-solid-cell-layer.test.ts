/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey } from "../core/model/cells";
import type { EditorActiveSheetView } from "../core/model/types";
import type { EditorPendingEdit } from "../webview/editor-panel/editor-panel-types";
import {
    deriveEditorGridMetrics,
    createInitialEditorGridViewportState,
} from "../webview-solid/editor-panel/grid-foundation";
import {
    deriveEditorGridCellLayers,
    reuseEquivalentEditorGridCellLayers,
} from "../webview-solid/editor-panel/cell-layer-helpers";
import type { EditorSheetFilterState } from "../webview/editor-panel/editor-panel-filter";

function createActiveSheet(): EditorActiveSheetView {
    return {
        key: "sheet:1",
        rowCount: 20,
        columnCount: 10,
        columns: ["A", "B", "C", "D"],
        columnWidths: undefined,
        rowHeights: undefined,
        cellAlignments: undefined,
        rowAlignments: undefined,
        columnAlignments: undefined,
        cells: {
            [createCellKey(1, 1)]: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                displayValue: "Frozen",
                formula: null,
                styleId: null,
            },
            [createCellKey(2, 2)]: {
                key: createCellKey(2, 2),
                rowNumber: 2,
                columnNumber: 2,
                address: "B2",
                displayValue: "Base",
                formula: "=1+1",
                styleId: null,
            },
        },
        freezePane: {
            rowCount: 1,
            columnCount: 1,
            topLeftCell: "B2",
            activePane: "bottomRight",
        },
        autoFilter: null,
    };
}

suite("Solid editor cell layers", () => {
    test("derives body, top, left, and corner cells from visible windows", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 20,
                columnCount: 10,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet: createActiveSheet(),
            pendingEdits: [],
            selection: null,
        });

        assert.strictEqual(layers.corner.length, 1);
        assert.ok(layers.top.length > 0);
        assert.ok(layers.left.length > 0);
        assert.ok(layers.body.length > 0);
        assert.strictEqual(layers.corner[0]?.displayValue, "Frozen");
    });

    test("prefers pending edit values and marks the selected cell", () => {
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: 20,
                columnCount: 10,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                },
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );
        const pendingEdits: EditorPendingEdit[] = [
            {
                sheetKey: "sheet:1",
                rowNumber: 2,
                columnNumber: 2,
                value: "Draft",
            },
        ];

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet: createActiveSheet(),
            pendingEdits,
            selection: {
                key: createCellKey(2, 2),
                rowNumber: 2,
                columnNumber: 2,
                address: "B2",
                value: "Draft",
                formula: "=1+1",
                isPresent: true,
            },
        });

        const selectedCell = layers.body.find(
            (item) => item.rowNumber === 2 && item.columnNumber === 2
        );
        assert.ok(selectedCell);
        assert.strictEqual(selectedCell?.displayValue, "Draft");
        assert.strictEqual(selectedCell?.isPending, true);
        assert.strictEqual(selectedCell?.isSelected, true);
        assert.strictEqual(selectedCell?.formula, null);
    });

    test("fills padded visible cells with empty display values", () => {
        const activeSheet = createActiveSheet();
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 900,
                viewportWidth: 1280,
            })
        );

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: null,
        });

        const paddedCell = layers.body.find(
            (item) =>
                item.rowNumber > activeSheet.rowCount || item.columnNumber > activeSheet.columnCount
        );
        assert.ok(paddedCell);
        assert.strictEqual(paddedCell?.displayValue, "");
    });

    test("derives overflow spill metrics for long values with blank trailing cells", () => {
        const activeSheet: EditorActiveSheetView = {
            ...createActiveSheet(),
            columnCount: 4,
            cells: {
                [createCellKey(1, 1)]: {
                    key: createCellKey(1, 1),
                    rowNumber: 1,
                    columnNumber: 1,
                    address: "A1",
                    displayValue: "Long text",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(1, 4)]: {
                    key: createCellKey(1, 4),
                    rowNumber: 1,
                    columnNumber: 4,
                    address: "D1",
                    displayValue: "stop",
                    formula: null,
                    styleId: null,
                },
            },
        };

        const metrics = deriveEditorGridMetrics(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
                columnWidths: activeSheet.columnWidths,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: null,
        });
        const overflowingCell = layers.body.find(
            (item) => item.rowNumber === 1 && item.columnNumber === 1
        );

        assert.ok(overflowingCell);
        assert.strictEqual(overflowingCell?.spillsIntoNextCells, true);
        assert.strictEqual(overflowingCell?.displayMaxWidthPx, 346);
        assert.strictEqual(overflowingCell?.contentMaxHeightPx, 22);
        assert.strictEqual(overflowingCell?.visibleLineCount, 1);
    });

    test("derives overflow spill metrics across the frozen column boundary", () => {
        const activeSheet: EditorActiveSheetView = {
            ...createActiveSheet(),
            columnCount: 4,
            freezePane: {
                rowCount: 1,
                columnCount: 1,
                topLeftCell: "B2",
                activePane: "bottomRight",
            },
            cells: {
                [createCellKey(1, 1)]: {
                    key: createCellKey(1, 1),
                    rowNumber: 1,
                    columnNumber: 1,
                    address: "A1",
                    displayValue: "Long text",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(1, 4)]: {
                    key: createCellKey(1, 4),
                    rowNumber: 1,
                    columnNumber: 4,
                    address: "D1",
                    displayValue: "stop",
                    formula: null,
                    styleId: null,
                },
            },
        };

        const metrics = deriveEditorGridMetrics(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
                columnWidths: activeSheet.columnWidths,
                freezePane: activeSheet.freezePane,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: null,
        });
        const overflowingCell = layers.corner.find(
            (item) => item.rowNumber === 1 && item.columnNumber === 1
        );

        assert.ok(overflowingCell);
        assert.strictEqual(overflowingCell?.spillsIntoNextCells, true);
        assert.strictEqual(overflowingCell?.displayMaxWidthPx, 346);
    });

    test("marks active filter header cells", () => {
        const activeSheet = createActiveSheet();
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );
        const filterState: EditorSheetFilterState = {
            range: {
                startRow: 1,
                endRow: 10,
                startColumn: 1,
                endColumn: 2,
            },
            sort: {
                columnNumber: 2,
                direction: "asc",
            },
            includedValuesByColumn: {
                "1": ["Frozen"],
            },
        };

        const layers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: null,
            filterState,
        });
        const firstHeader = layers.body.find(
            (item) => item.rowNumber === 1 && item.columnNumber === 1
        );
        const secondHeader = layers.body.find(
            (item) => item.rowNumber === 1 && item.columnNumber === 2
        );

        assert.ok(firstHeader);
        assert.strictEqual(firstHeader?.isFilterHeader, true);
        assert.strictEqual(firstHeader?.isColumnFilterActive, true);
        assert.ok(secondHeader);
        assert.strictEqual(secondHeader?.isFilterHeader, true);
        assert.strictEqual(secondHeader?.isColumnFilterActive, true);
    });

    test("reuses stable cell objects for unaffected visible cells", () => {
        const activeSheet = createActiveSheet();
        const metrics = deriveEditorGridMetrics(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
                freezePane: activeSheet.freezePane,
            },
            createInitialEditorGridViewportState({
                viewportHeight: 360,
                viewportWidth: 540,
            })
        );

        const baseLayers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: null,
        });
        const nextLayers = deriveEditorGridCellLayers({
            metrics,
            activeSheet,
            pendingEdits: [],
            selection: {
                key: createCellKey(2, 2),
                rowNumber: 2,
                columnNumber: 2,
                address: "B2",
                value: "Base",
                formula: "=1+1",
                isPresent: true,
            },
        });
        const reusedLayers = reuseEquivalentEditorGridCellLayers(baseLayers, nextLayers);

        const stableCell = baseLayers.body.find(
            (item) => item.rowNumber === 3 && item.columnNumber === 3
        );
        const reusedStableCell = reusedLayers.body.find(
            (item) => item.rowNumber === 3 && item.columnNumber === 3
        );
        const changedCell = baseLayers.body.find(
            (item) => item.rowNumber === 2 && item.columnNumber === 2
        );
        const reusedChangedCell = reusedLayers.body.find(
            (item) => item.rowNumber === 2 && item.columnNumber === 2
        );

        assert.ok(stableCell);
        assert.strictEqual(reusedStableCell, stableCell);
        assert.ok(changedCell);
        assert.notStrictEqual(reusedChangedCell, changedCell);
        assert.strictEqual(reusedChangedCell?.isSelected, true);
    });
});
