import { createCellKey } from "../../core/model/cells";
import type { CellAlignmentSnapshot } from "../../core/model/alignment";
import type { EditorActiveSheetView, EditorSelectionView } from "../../core/model/types";
import { getCellOverflowMetrics } from "./editor-cell-overflow";
import type { EditorPendingEdit } from "./editor-panel-types";
import {
    getEditorColumnLeft,
    getEditorColumnWidth,
} from "./editor-virtual-grid";
import {
    getEditorGridActualRowHeight,
    getEditorGridActualRowTop,
    type EditorGridMetrics,
} from "./grid-foundation";
import { EDITOR_GRID_HEADER_HEIGHT } from "./header-layer-helpers";
import { getEffectiveEditorCellAlignment } from "./alignment-helpers";
import {
    isEditorFilterHeaderCell,
    type EditorSheetFilterState,
} from "./editor-panel-filter";

export type EditorGridCellLayerKind = "body" | "top" | "left" | "corner";

const GRID_CELL_VERTICAL_PADDING_PX = 6;
const GRID_CELL_FONT_SIZE_PX = 13;
const GRID_CELL_LINE_HEIGHT_MULTIPLIER = 1.25;
const GRID_CELL_LINE_HEIGHT_PX = GRID_CELL_FONT_SIZE_PX * GRID_CELL_LINE_HEIGHT_MULTIPLIER;

export interface EditorGridCellItem {
    key: string;
    rowNumber: number;
    columnNumber: number;
    top: number;
    left: number;
    width: number;
    height: number;
    displayValue: string;
    formula: string | null;
    alignment: CellAlignmentSnapshot | null;
    contentMaxHeightPx: number;
    visibleLineCount: number;
    spillsIntoNextCells: boolean;
    displayMaxWidthPx: number | null;
    isPending: boolean;
    isSelected: boolean;
    isFilterHeader: boolean;
    isColumnFilterActive: boolean;
    layer: EditorGridCellLayerKind;
}

export interface EditorGridCellLayers {
    body: EditorGridCellItem[];
    top: EditorGridCellItem[];
    left: EditorGridCellItem[];
    corner: EditorGridCellItem[];
}

function areEditorGridCellItemsEqual(left: EditorGridCellItem, right: EditorGridCellItem): boolean {
    return (
        left.key === right.key &&
        left.rowNumber === right.rowNumber &&
        left.columnNumber === right.columnNumber &&
        left.top === right.top &&
        left.left === right.left &&
        left.width === right.width &&
        left.height === right.height &&
        left.displayValue === right.displayValue &&
        left.formula === right.formula &&
        left.alignment === right.alignment &&
        left.contentMaxHeightPx === right.contentMaxHeightPx &&
        left.visibleLineCount === right.visibleLineCount &&
        left.spillsIntoNextCells === right.spillsIntoNextCells &&
        left.displayMaxWidthPx === right.displayMaxWidthPx &&
        left.isPending === right.isPending &&
        left.isSelected === right.isSelected &&
        left.isFilterHeader === right.isFilterHeader &&
        left.isColumnFilterActive === right.isColumnFilterActive &&
        left.layer === right.layer
    );
}

function reuseEquivalentEditorGridCellLayer(
    previous: EditorGridCellItem[] | null,
    next: readonly EditorGridCellItem[]
): EditorGridCellItem[] {
    if (!previous || previous.length === 0) {
        return [...next];
    }

    if (next.length === 0) {
        return [];
    }

    const previousByKey = new Map(previous.map((item) => [item.key, item]));
    const reused = next.map((item) => {
        const previousItem = previousByKey.get(item.key);
        return previousItem && areEditorGridCellItemsEqual(previousItem, item)
            ? previousItem
            : item;
    });

    if (
        previous.length === reused.length &&
        reused.every((item, index) => item === previous[index])
    ) {
        return previous;
    }

    return reused;
}

function getPendingCellValue(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    pendingEdits: readonly EditorPendingEdit[]
): string | undefined {
    return pendingEdits.find(
        (edit) =>
            edit.sheetKey === sheetKey &&
            edit.rowNumber === rowNumber &&
            edit.columnNumber === columnNumber
    )?.value;
}

function getEditorGridCellTextLayoutMetrics(cellHeight: number): {
    contentMaxHeightPx: number;
    visibleLineCount: number;
} {
    const contentMaxHeightPx = Math.max(18, cellHeight - GRID_CELL_VERTICAL_PADDING_PX);
    return {
        contentMaxHeightPx,
        visibleLineCount: Math.max(1, Math.floor(contentMaxHeightPx / GRID_CELL_LINE_HEIGHT_PX)),
    };
}

function getEditorFrozenCellOverflowColumnNumbers(
    frozenColumnNumbers: readonly number[],
    scrollableColumnNumbers: readonly number[]
): number[] {
    if (frozenColumnNumbers.length === 0) {
        return [...scrollableColumnNumbers];
    }

    const seenColumnNumbers = new Set<number>();
    const columnNumbers: number[] = [];

    for (const columnNumber of [...frozenColumnNumbers, ...scrollableColumnNumbers]) {
        if (seenColumnNumbers.has(columnNumber)) {
            continue;
        }

        seenColumnNumbers.add(columnNumber);
        columnNumbers.push(columnNumber);
    }

    return columnNumbers;
}

function createEditorGridCellItem({
    layer,
    rowNumber,
    columnNumber,
    visibleColumnNumbers,
    visibleColumnIndex,
    metrics,
    activeSheet,
    pendingEdits,
    selection,
    filterState,
}: {
    layer: EditorGridCellLayerKind;
    rowNumber: number;
    columnNumber: number;
    visibleColumnNumbers: readonly number[];
    visibleColumnIndex: number;
    metrics: EditorGridMetrics;
    activeSheet: EditorActiveSheetView;
    pendingEdits: readonly EditorPendingEdit[];
    selection: EditorSelectionView | null;
    filterState: EditorSheetFilterState | null;
}): EditorGridCellItem {
    const cell = activeSheet.cells[createCellKey(rowNumber, columnNumber)];
    const pendingValue = getPendingCellValue(
        activeSheet.key,
        rowNumber,
        columnNumber,
        pendingEdits
    );
    const displayValue = pendingValue ?? cell?.displayValue ?? "";
    const formula = pendingValue !== undefined ? null : (cell?.formula ?? null);
    const alignment = getEffectiveEditorCellAlignment(activeSheet, rowNumber, columnNumber);
    const width = getEditorColumnWidth(metrics.columnLayout, columnNumber);
    const height =
        getEditorGridActualRowHeight(metrics, rowNumber) ?? metrics.rowLayout.fallbackPixelHeight;
    const isFilterHeader = isEditorFilterHeaderCell(filterState, rowNumber, columnNumber);
    const isColumnFilterActive =
        Boolean(filterState?.includedValuesByColumn[String(columnNumber)]) ||
        filterState?.sort?.columnNumber === columnNumber;
    const overflowMetrics = getCellOverflowMetrics({
        value: displayValue,
        alignment,
        baseColumnWidth: width,
        visibleColumnNumbers,
        visibleColumnIndex,
        getColumnWidth: (nextColumnNumber) =>
            getEditorColumnWidth(metrics.columnLayout, nextColumnNumber),
        getTrailingCellState: (nextColumnNumber) => {
            const nextCell = activeSheet.cells[createCellKey(rowNumber, nextColumnNumber)];
            const nextPendingValue = getPendingCellValue(
                activeSheet.key,
                rowNumber,
                nextColumnNumber,
                pendingEdits
            );
            return {
                value: nextPendingValue ?? nextCell?.displayValue ?? "",
                formula: nextPendingValue !== undefined ? null : (nextCell?.formula ?? null),
                blocksOverflow: isEditorFilterHeaderCell(filterState, rowNumber, nextColumnNumber),
            };
        },
    });
    const { contentMaxHeightPx, visibleLineCount } = getEditorGridCellTextLayoutMetrics(height);

    return {
        key: `${layer}:${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        top: EDITOR_GRID_HEADER_HEIGHT + (getEditorGridActualRowTop(metrics, rowNumber) ?? 0),
        left: metrics.rowHeaderWidth + getEditorColumnLeft(metrics.columnLayout, columnNumber),
        width,
        height,
        displayValue,
        formula,
        alignment,
        contentMaxHeightPx,
        visibleLineCount,
        spillsIntoNextCells: overflowMetrics.spillsIntoNextCells,
        displayMaxWidthPx: overflowMetrics.spillsIntoNextCells
            ? overflowMetrics.contentWidthPx
            : null,
        isPending: pendingValue !== undefined,
        isSelected: selection?.rowNumber === rowNumber && selection?.columnNumber === columnNumber,
        isFilterHeader,
        isColumnFilterActive,
        layer,
    };
}

export function deriveEditorGridCellLayers({
    metrics,
    activeSheet,
    pendingEdits,
    selection,
    filterState,
}: {
    metrics: EditorGridMetrics;
    activeSheet: EditorActiveSheetView;
    pendingEdits: readonly EditorPendingEdit[];
    selection: EditorSelectionView | null;
    filterState?: EditorSheetFilterState | null;
}): EditorGridCellLayers {
    const body: EditorGridCellItem[] = [];
    const top: EditorGridCellItem[] = [];
    const left: EditorGridCellItem[] = [];
    const corner: EditorGridCellItem[] = [];
    const frozenCellOverflowColumnNumbers = getEditorFrozenCellOverflowColumnNumbers(
        metrics.window.frozenColumnNumbers,
        metrics.window.columnNumbers
    );

    for (const rowNumber of metrics.window.rowNumbers) {
        for (
            let visibleColumnIndex = 0;
            visibleColumnIndex < metrics.window.columnNumbers.length;
            visibleColumnIndex += 1
        ) {
            const columnNumber = metrics.window.columnNumbers[visibleColumnIndex];
            if (columnNumber === undefined) {
                continue;
            }

            body.push(
                createEditorGridCellItem({
                    layer: "body",
                    rowNumber,
                    columnNumber,
                    visibleColumnNumbers: metrics.window.columnNumbers,
                    visibleColumnIndex,
                    metrics,
                    activeSheet,
                    pendingEdits,
                    selection,
                    filterState: filterState ?? null,
                })
            );
        }

        for (
            let visibleColumnIndex = 0;
            visibleColumnIndex < metrics.window.frozenColumnNumbers.length;
            visibleColumnIndex += 1
        ) {
            const columnNumber = metrics.window.frozenColumnNumbers[visibleColumnIndex];
            if (columnNumber === undefined) {
                continue;
            }

            left.push(
                createEditorGridCellItem({
                    layer: "left",
                    rowNumber,
                    columnNumber,
                    visibleColumnNumbers: frozenCellOverflowColumnNumbers,
                    visibleColumnIndex,
                    metrics,
                    activeSheet,
                    pendingEdits,
                    selection,
                    filterState: filterState ?? null,
                })
            );
        }
    }

    for (const rowNumber of metrics.window.frozenRowNumbers) {
        for (
            let visibleColumnIndex = 0;
            visibleColumnIndex < metrics.window.columnNumbers.length;
            visibleColumnIndex += 1
        ) {
            const columnNumber = metrics.window.columnNumbers[visibleColumnIndex];
            if (columnNumber === undefined) {
                continue;
            }

            top.push(
                createEditorGridCellItem({
                    layer: "top",
                    rowNumber,
                    columnNumber,
                    visibleColumnNumbers: metrics.window.columnNumbers,
                    visibleColumnIndex,
                    metrics,
                    activeSheet,
                    pendingEdits,
                    selection,
                    filterState: filterState ?? null,
                })
            );
        }

        for (
            let visibleColumnIndex = 0;
            visibleColumnIndex < metrics.window.frozenColumnNumbers.length;
            visibleColumnIndex += 1
        ) {
            const columnNumber = metrics.window.frozenColumnNumbers[visibleColumnIndex];
            if (columnNumber === undefined) {
                continue;
            }

            corner.push(
                createEditorGridCellItem({
                    layer: "corner",
                    rowNumber,
                    columnNumber,
                    visibleColumnNumbers: frozenCellOverflowColumnNumbers,
                    visibleColumnIndex,
                    metrics,
                    activeSheet,
                    pendingEdits,
                    selection,
                    filterState: filterState ?? null,
                })
            );
        }
    }

    return {
        body,
        top,
        left,
        corner,
    };
}

export function reuseEquivalentEditorGridCellLayers(
    previous: EditorGridCellLayers | null,
    next: EditorGridCellLayers
): EditorGridCellLayers {
    if (!previous) {
        return next;
    }

    const body = reuseEquivalentEditorGridCellLayer(previous.body, next.body);
    const top = reuseEquivalentEditorGridCellLayer(previous.top, next.top);
    const left = reuseEquivalentEditorGridCellLayer(previous.left, next.left);
    const corner = reuseEquivalentEditorGridCellLayer(previous.corner, next.corner);

    if (
        body === previous.body &&
        top === previous.top &&
        left === previous.left &&
        corner === previous.corner
    ) {
        return previous;
    }

    return {
        body,
        top,
        left,
        corner,
    };
}
