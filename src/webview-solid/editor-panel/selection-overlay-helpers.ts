import type { EditorSelectionView } from "../../core/model/types";
import {
    hasExpandedSelectionRange,
    type SelectionRange,
} from "../../webview/editor-panel/editor-selection-range";
import { clipSelectionRangeToVisibleGrid } from "../../webview/editor-panel/editor-selection-overlay";
import {
    getEditorColumnLeft,
    getEditorColumnWidth,
} from "../../webview/editor-panel/editor-virtual-grid";
import {
    getEditorGridActualRowHeight,
    getEditorGridActualRowTop,
    type EditorGridMetrics,
} from "./grid-foundation";
import { EDITOR_GRID_HEADER_HEIGHT } from "./header-layer-helpers";
import { resolveEditorSelectionRange } from "./selection-range-state-helpers";

export type EditorSelectionOverlayLayerKind = "body" | "top" | "left" | "corner";

export interface EditorSelectionOverlayRect {
    top: number;
    left: number;
    width: number;
    height: number;
}

export interface EditorSelectionOverlayRangeRect extends EditorSelectionOverlayRect {
    showTopBorder: boolean;
    showRightBorder: boolean;
    showBottomBorder: boolean;
    showLeftBorder: boolean;
}

export interface EditorSelectionOverlayLayer {
    activeRowRect: EditorSelectionOverlayRect | null;
    activeColumnRect: EditorSelectionOverlayRect | null;
    rangeRect: EditorSelectionOverlayRangeRect | null;
    primaryRect: EditorSelectionOverlayRect | null;
}

export interface EditorSelectionOverlayLayers {
    body: EditorSelectionOverlayLayer;
    top: EditorSelectionOverlayLayer;
    left: EditorSelectionOverlayLayer;
    corner: EditorSelectionOverlayLayer;
}

function createSingleCellRange(selection: EditorSelectionView | null): SelectionRange | null {
    if (!selection) {
        return null;
    }

    return {
        startRow: selection.rowNumber,
        endRow: selection.rowNumber,
        startColumn: selection.columnNumber,
        endColumn: selection.columnNumber,
    };
}

function getEditorGridTop(metrics: EditorGridMetrics, rowNumber: number): number {
    return EDITOR_GRID_HEADER_HEIGHT + (getEditorGridActualRowTop(metrics, rowNumber) ?? 0);
}

function getEditorGridLeft(metrics: EditorGridMetrics, columnNumber: number): number {
    return metrics.rowHeaderWidth + getEditorColumnLeft(metrics.columnLayout, columnNumber);
}

function getGridLayerRowNumbers(
    metrics: EditorGridMetrics,
    layer: EditorSelectionOverlayLayerKind
): readonly number[] {
    if (layer === "top" || layer === "corner") {
        return metrics.window.frozenRowNumbers;
    }

    return metrics.window.rowNumbers;
}

function getGridLayerColumnNumbers(
    metrics: EditorGridMetrics,
    layer: EditorSelectionOverlayLayerKind
): readonly number[] {
    if (layer === "left" || layer === "corner") {
        return metrics.window.frozenColumnNumbers;
    }

    return metrics.window.columnNumbers;
}

function getGridLayerForCell(
    metrics: EditorGridMetrics,
    selection: EditorSelectionView | null
): EditorSelectionOverlayLayerKind | null {
    if (!selection) {
        return null;
    }

    const isFrozenRow = metrics.window.frozenRowNumbers.includes(selection.rowNumber);
    const isFrozenColumn = metrics.window.frozenColumnNumbers.includes(selection.columnNumber);
    const isVisibleRow = isFrozenRow || metrics.window.rowNumbers.includes(selection.rowNumber);
    const isVisibleColumn =
        isFrozenColumn || metrics.window.columnNumbers.includes(selection.columnNumber);

    if (!isVisibleRow || !isVisibleColumn) {
        return null;
    }

    if (isFrozenRow && isFrozenColumn) {
        return "corner";
    }

    if (isFrozenRow) {
        return "top";
    }

    if (isFrozenColumn) {
        return "left";
    }

    return "body";
}

function createSelectionOverlayRect(
    metrics: EditorGridMetrics,
    selectionRange: SelectionRange | null,
    layer: EditorSelectionOverlayLayerKind
): EditorSelectionOverlayRect | null {
    const visibleRange = clipSelectionRangeToVisibleGrid(
        selectionRange,
        getGridLayerRowNumbers(metrics, layer),
        getGridLayerColumnNumbers(metrics, layer)
    );
    if (!visibleRange) {
        return null;
    }

    const top = getEditorGridTop(metrics, visibleRange.startRow);
    const bottomTop = getEditorGridTop(metrics, visibleRange.endRow);
    const left = getEditorGridLeft(metrics, visibleRange.startColumn);
    const right =
        getEditorGridLeft(metrics, visibleRange.endColumn) +
        getEditorColumnWidth(metrics.columnLayout, visibleRange.endColumn);
    const bottom = bottomTop + (getEditorGridActualRowHeight(metrics, visibleRange.endRow) ?? 0);

    return {
        top,
        left,
        width: right - left,
        height: bottom - top,
    };
}

function createSelectionOverlayRangeRect(
    metrics: EditorGridMetrics,
    selectionRange: SelectionRange | null,
    layer: EditorSelectionOverlayLayerKind
): EditorSelectionOverlayRangeRect | null {
    const visibleRange = clipSelectionRangeToVisibleGrid(
        selectionRange,
        getGridLayerRowNumbers(metrics, layer),
        getGridLayerColumnNumbers(metrics, layer)
    );
    if (!selectionRange || !visibleRange) {
        return null;
    }

    const rect = createSelectionOverlayRect(metrics, selectionRange, layer);
    if (!rect) {
        return null;
    }

    return {
        ...rect,
        showTopBorder: visibleRange.startRow === selectionRange.startRow,
        showRightBorder: visibleRange.endColumn === selectionRange.endColumn,
        showBottomBorder: visibleRange.endRow === selectionRange.endRow,
        showLeftBorder: visibleRange.startColumn === selectionRange.startColumn,
    };
}

function createSelectionOverlayCellRect(
    metrics: EditorGridMetrics,
    selection: EditorSelectionView | null,
    layer: EditorSelectionOverlayLayerKind
): EditorSelectionOverlayRect | null {
    if (!selection || getGridLayerForCell(metrics, selection) !== layer) {
        return null;
    }

    return createSelectionOverlayRect(metrics, createSingleCellRange(selection), layer);
}

function getLastNumber(values: readonly number[]): number | null {
    return values[values.length - 1] ?? null;
}

function createSelectionOverlayRowRect(
    metrics: EditorGridMetrics,
    rowNumber: number | null,
    layer: EditorSelectionOverlayLayerKind
): EditorSelectionOverlayRect | null {
    const columnNumbers = getGridLayerColumnNumbers(metrics, layer);
    const startColumn = columnNumbers[0] ?? null;
    const endColumn = getLastNumber(columnNumbers);
    if (rowNumber === null || startColumn === null || endColumn === null) {
        return null;
    }

    return createSelectionOverlayRect(
        metrics,
        {
            startRow: rowNumber,
            endRow: rowNumber,
            startColumn,
            endColumn,
        },
        layer
    );
}

function createSelectionOverlayColumnRect(
    metrics: EditorGridMetrics,
    columnNumber: number | null,
    layer: EditorSelectionOverlayLayerKind
): EditorSelectionOverlayRect | null {
    const rowNumbers = getGridLayerRowNumbers(metrics, layer);
    const startRow = rowNumbers[0] ?? null;
    const endRow = getLastNumber(rowNumbers);
    if (columnNumber === null || startRow === null || endRow === null) {
        return null;
    }

    return createSelectionOverlayRect(
        metrics,
        {
            startRow,
            endRow,
            startColumn: columnNumber,
            endColumn: columnNumber,
        },
        layer
    );
}

function createEmptySelectionOverlayLayer(): EditorSelectionOverlayLayer {
    return {
        activeRowRect: null,
        activeColumnRect: null,
        rangeRect: null,
        primaryRect: null,
    };
}

export function deriveEditorSelectionOverlayLayers({
    metrics,
    selection,
    selectionRangeOverride,
    forcePrimaryRect = false,
}: {
    metrics: EditorGridMetrics;
    selection: EditorSelectionView | null;
    selectionRangeOverride: SelectionRange | null;
    forcePrimaryRect?: boolean;
}): EditorSelectionOverlayLayers {
    const selectionRange = resolveEditorSelectionRange(selection, selectionRangeOverride);
    const showPrimaryRect = forcePrimaryRect || !hasExpandedSelectionRange(selectionRange);
    const activeRowNumber = selection?.rowNumber ?? null;
    const activeColumnNumber = selection?.columnNumber ?? null;

    const createLayer = (layer: EditorSelectionOverlayLayerKind): EditorSelectionOverlayLayer => ({
        activeRowRect: createSelectionOverlayRowRect(metrics, activeRowNumber, layer),
        activeColumnRect: createSelectionOverlayColumnRect(metrics, activeColumnNumber, layer),
        rangeRect: createSelectionOverlayRangeRect(metrics, selectionRange, layer),
        primaryRect: showPrimaryRect
            ? createSelectionOverlayCellRect(metrics, selection, layer)
            : null,
    });

    if (!selection) {
        return {
            body: createEmptySelectionOverlayLayer(),
            top: createEmptySelectionOverlayLayer(),
            left: createEmptySelectionOverlayLayer(),
            corner: createEmptySelectionOverlayLayer(),
        };
    }

    return {
        body: createLayer("body"),
        top: createLayer("top"),
        left: createLayer("left"),
        corner: createLayer("corner"),
    };
}
