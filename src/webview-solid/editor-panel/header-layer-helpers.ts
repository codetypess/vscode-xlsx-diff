import { getColumnLabel } from "../../core/model/cells";
import type { SelectionRange } from "../../webview/editor-panel/editor-selection-range";
import {
    getEditorColumnLeft,
    getEditorColumnWidth,
} from "../../webview/editor-panel/editor-virtual-grid";
import {
    getEditorGridActualRowHeight,
    getEditorGridActualRowTop,
    type EditorGridMetrics,
} from "./grid-foundation";

export const EDITOR_GRID_HEADER_HEIGHT = 28;

export interface EditorColumnHeaderItem {
    key: string;
    columnNumber: number;
    label: string;
    left: number;
    width: number;
    isFrozen: boolean;
    isActive: boolean;
}

export interface EditorRowHeaderItem {
    key: string;
    rowNumber: number;
    top: number;
    height: number;
    isFrozen: boolean;
    isActive: boolean;
}

export interface EditorGridHeaderLayers {
    headerHeight: number;
    rowHeaderWidth: number;
    frozenColumnHeaders: EditorColumnHeaderItem[];
    scrollableColumnHeaders: EditorColumnHeaderItem[];
    frozenRowHeaders: EditorRowHeaderItem[];
    scrollableRowHeaders: EditorRowHeaderItem[];
}

function isColumnWithinSelectionRange(
    selectionRange: SelectionRange | null,
    columnNumber: number
): boolean {
    return Boolean(
        selectionRange &&
        columnNumber >= selectionRange.startColumn &&
        columnNumber <= selectionRange.endColumn
    );
}

function isRowWithinSelectionRange(
    selectionRange: SelectionRange | null,
    rowNumber: number
): boolean {
    return Boolean(
        selectionRange && rowNumber >= selectionRange.startRow && rowNumber <= selectionRange.endRow
    );
}

export function deriveEditorGridHeaderLayers({
    metrics,
    columnLabels,
    selectedRowNumber,
    selectedColumnNumber,
    selectionRange,
}: {
    metrics: EditorGridMetrics;
    columnLabels: readonly string[];
    selectedRowNumber: number | null;
    selectedColumnNumber: number | null;
    selectionRange: SelectionRange | null;
}): EditorGridHeaderLayers {
    const createColumnHeader = (
        columnNumber: number,
        isFrozen: boolean
    ): EditorColumnHeaderItem => ({
        key: `${isFrozen ? "frozen" : "scroll"}:column:${columnNumber}`,
        columnNumber,
        label: columnLabels[columnNumber - 1] ?? getColumnLabel(columnNumber),
        left: metrics.rowHeaderWidth + getEditorColumnLeft(metrics.columnLayout, columnNumber),
        width: getEditorColumnWidth(metrics.columnLayout, columnNumber),
        isFrozen,
        isActive:
            isColumnWithinSelectionRange(selectionRange, columnNumber) ||
            selectedColumnNumber === columnNumber,
    });
    const createRowHeader = (rowNumber: number, isFrozen: boolean): EditorRowHeaderItem => ({
        key: `${isFrozen ? "frozen" : "scroll"}:row:${rowNumber}`,
        rowNumber,
        top: EDITOR_GRID_HEADER_HEIGHT + (getEditorGridActualRowTop(metrics, rowNumber) ?? 0),
        height: getEditorGridActualRowHeight(metrics, rowNumber) ?? 0,
        isFrozen,
        isActive:
            isRowWithinSelectionRange(selectionRange, rowNumber) || selectedRowNumber === rowNumber,
    });

    return {
        headerHeight: EDITOR_GRID_HEADER_HEIGHT,
        rowHeaderWidth: metrics.rowHeaderWidth,
        frozenColumnHeaders: metrics.window.frozenColumnNumbers.map((columnNumber) =>
            createColumnHeader(columnNumber, true)
        ),
        scrollableColumnHeaders: metrics.window.columnNumbers.map((columnNumber) =>
            createColumnHeader(columnNumber, false)
        ),
        frozenRowHeaders: metrics.window.frozenRowNumbers.map((rowNumber) =>
            createRowHeader(rowNumber, true)
        ),
        scrollableRowHeaders: metrics.window.rowNumbers.map((rowNumber) =>
            createRowHeader(rowNumber, false)
        ),
    };
}

export function createEditorRowHeaderSelection(
    rowNumber: number,
    selectedColumnNumber: number | null | undefined
): { rowNumber: number; columnNumber: number } {
    return {
        rowNumber: Math.max(1, Math.trunc(rowNumber)),
        columnNumber: Math.max(1, Math.trunc(selectedColumnNumber ?? 1)),
    };
}

export function createEditorColumnHeaderSelection(
    columnNumber: number,
    selectedRowNumber: number | null | undefined
): { rowNumber: number; columnNumber: number } {
    return {
        rowNumber: Math.max(1, Math.trunc(selectedRowNumber ?? 1)),
        columnNumber: Math.max(1, Math.trunc(columnNumber)),
    };
}
