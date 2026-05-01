import {
    EDITOR_EXTRA_PADDING_ROWS,
    clampEditorScrollPosition,
    createEditorColumnWindow,
    createEditorPixelColumnLayout,
    createEditorPixelRowLayout,
    createEditorRowWindow,
    getEditorContentSize,
    getEditorDisplayColumnLayout,
    getEditorDisplayGridDimensions,
    getEditorDisplayRowLayout,
    getEditorFrozenColumnsWidth,
    getEditorFrozenRowsHeight,
    getEditorRowHeight,
    getEditorRowTop,
    getFrozenEditorCounts,
    getVisibleFrozenEditorCounts,
} from "../../webview/editor-panel/editor-virtual-grid";
import type { SheetFreezePaneSnapshot, SheetRowHeightsSnapshot } from "../../core/model/types";
import type { PixelColumnLayout } from "../../webview/column-layout";
import type { PixelRowLayout } from "../../webview/row-layout";

export const DEFAULT_SOLID_EDITOR_VIEWPORT_HEIGHT = 480;
export const DEFAULT_SOLID_EDITOR_VIEWPORT_WIDTH = 960;

export interface EditorGridViewportState {
    scrollTop: number;
    scrollLeft: number;
    viewportHeight: number;
    viewportWidth: number;
}

export interface EditorGridViewportPatch {
    scrollTop?: number;
    scrollLeft?: number;
    viewportHeight?: number;
    viewportWidth?: number;
}

export interface EditorGridMetricsInput {
    rowCount: number;
    columnCount: number;
    rowHeaderLabelCount?: number;
    visibleRows?: readonly number[];
    hiddenRows?: readonly number[];
    rowHeights?: SheetRowHeightsSnapshot;
    columnWidths?: readonly (number | null)[];
    freezePane?: Pick<SheetFreezePaneSnapshot, "rowCount" | "columnCount"> | null;
    maximumDigitWidth?: number;
}

export interface EditorGridWindowState {
    rowNumbers: number[];
    columnNumbers: number[];
    frozenRowNumbers: number[];
    frozenColumnNumbers: number[];
    topSpacerHeight: number;
    bottomSpacerHeight: number;
    leadingSpacerWidth: number;
    trailingSpacerWidth: number;
}

export interface EditorGridDisplayRowState {
    actualRowNumbers: number[];
    actualToDisplayRowNumbers: Record<string, number>;
    hiddenActualRowNumbers: number[];
}

export interface EditorGridMetrics {
    viewport: EditorGridViewportState;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
    rowState: EditorGridDisplayRowState;
    rowHeaderWidth: number;
    contentWidth: number;
    contentHeight: number;
    stickyTopHeight: number;
    stickyLeftWidth: number;
    frozenRowCount: number;
    frozenColumnCount: number;
    visibleFrozenRowCount: number;
    visibleFrozenColumnCount: number;
    displayRowCount: number;
    displayColumnCount: number;
    window: EditorGridWindowState;
}

function createSequentialNumbers(count: number): number[] {
    return Array.from({ length: Math.max(0, count) }, (_, index) => index + 1);
}

function areNumberArraysEqual(left: readonly number[], right: readonly number[]): boolean {
    if (left.length !== right.length) {
        return false;
    }

    for (let index = 0; index < left.length; index += 1) {
        if (left[index] !== right[index]) {
            return false;
        }
    }

    return true;
}

function createEditorGridDisplayRowState({
    sourceRowCount,
    visibleRows,
    hiddenRows,
    totalDisplayRowCount,
}: {
    sourceRowCount: number;
    visibleRows: readonly number[];
    hiddenRows: readonly number[];
    totalDisplayRowCount: number;
}): EditorGridDisplayRowState {
    const actualRowNumbers = [...visibleRows];
    let nextSyntheticRowNumber = sourceRowCount + 1;
    while (actualRowNumbers.length < totalDisplayRowCount) {
        actualRowNumbers.push(nextSyntheticRowNumber);
        nextSyntheticRowNumber += 1;
    }

    return {
        actualRowNumbers,
        actualToDisplayRowNumbers: Object.fromEntries(
            actualRowNumbers.map((actualRowNumber, index) => [String(actualRowNumber), index + 1])
        ),
        hiddenActualRowNumbers: [...hiddenRows],
    };
}

function mapActualRowHeightsToDisplayRowHeights(
    actualToDisplayRowNumbers: Readonly<Record<string, number>>,
    rowHeights: SheetRowHeightsSnapshot | undefined
): Record<string, number | null> | undefined {
    const displayRowHeights = Object.fromEntries(
        Object.entries(rowHeights ?? {}).flatMap(([actualRowNumber, rowHeight]) => {
            const displayRowNumber = actualToDisplayRowNumbers[actualRowNumber];
            if (!Number.isInteger(displayRowNumber)) {
                return [];
            }

            return [[String(displayRowNumber), rowHeight] as const];
        })
    );

    return Object.keys(displayRowHeights).length > 0 ? displayRowHeights : undefined;
}

function getActualRowNumberAtDisplayRow(
    rowState: EditorGridDisplayRowState,
    displayRowNumber: number
): number | null {
    return rowState.actualRowNumbers[displayRowNumber - 1] ?? null;
}

function mapDisplayRowNumbersToActualRowNumbers(
    rowState: EditorGridDisplayRowState,
    displayRowNumbers: readonly number[]
): number[] {
    return displayRowNumbers
        .map((displayRowNumber) => getActualRowNumberAtDisplayRow(rowState, displayRowNumber))
        .filter((rowNumber): rowNumber is number => rowNumber !== null);
}

function areColumnLayoutsEqual(left: PixelColumnLayout, right: PixelColumnLayout): boolean {
    return (
        left.actualColumnCount === right.actualColumnCount &&
        left.totalColumnCount === right.totalColumnCount &&
        left.actualColumnsWidth === right.actualColumnsWidth &&
        left.totalWidth === right.totalWidth &&
        left.fallbackPixelWidth === right.fallbackPixelWidth &&
        areNumberArraysEqual(left.pixelWidths, right.pixelWidths)
    );
}

function areDisplayRowStatesEqual(
    left: EditorGridDisplayRowState,
    right: EditorGridDisplayRowState
): boolean {
    return (
        areNumberArraysEqual(left.actualRowNumbers, right.actualRowNumbers) &&
        areNumberArraysEqual(left.hiddenActualRowNumbers, right.hiddenActualRowNumbers)
    );
}

function areRowLayoutsEqual(left: PixelRowLayout, right: PixelRowLayout): boolean {
    return (
        left.totalRowCount === right.totalRowCount &&
        left.totalHeight === right.totalHeight &&
        left.fallbackPixelHeight === right.fallbackPixelHeight &&
        areNumberArraysEqual(left.overriddenRowNumbers, right.overriddenRowNumbers) &&
        left.overriddenRowNumbers.every(
            (rowNumber) =>
                left.overriddenPixelHeights[String(rowNumber)] ===
                right.overriddenPixelHeights[String(rowNumber)]
        )
    );
}

export function createInitialEditorGridViewportState(
    overrides: Partial<EditorGridViewportState> = {}
): EditorGridViewportState {
    return {
        scrollTop: Math.max(0, overrides.scrollTop ?? 0),
        scrollLeft: Math.max(0, overrides.scrollLeft ?? 0),
        viewportHeight: Math.max(
            0,
            overrides.viewportHeight ?? DEFAULT_SOLID_EDITOR_VIEWPORT_HEIGHT
        ),
        viewportWidth: Math.max(0, overrides.viewportWidth ?? DEFAULT_SOLID_EDITOR_VIEWPORT_WIDTH),
    };
}

export function applyEditorGridViewportPatch(
    state: EditorGridViewportState,
    patch: EditorGridViewportPatch
): EditorGridViewportState {
    return {
        scrollTop: Math.max(0, patch.scrollTop ?? state.scrollTop),
        scrollLeft: Math.max(0, patch.scrollLeft ?? state.scrollLeft),
        viewportHeight: Math.max(0, patch.viewportHeight ?? state.viewportHeight),
        viewportWidth: Math.max(0, patch.viewportWidth ?? state.viewportWidth),
    };
}

export function deriveEditorGridMetrics(
    input: EditorGridMetricsInput,
    viewport: EditorGridViewportState
): EditorGridMetrics {
    const normalizedViewport = createInitialEditorGridViewportState(viewport);
    const visibleRows = input.visibleRows ?? createSequentialNumbers(input.rowCount);
    const hiddenRows = input.hiddenRows ?? [];
    const baseActualToDisplayRowNumbers = Object.fromEntries(
        visibleRows.map((actualRowNumber, index) => [String(actualRowNumber), index + 1])
    );
    const baseVisibleRowLayout = createEditorPixelRowLayout({
        rowCount: visibleRows.length,
        rowHeights: mapActualRowHeightsToDisplayRowHeights(
            baseActualToDisplayRowNumbers,
            input.rowHeights
        ),
    });
    const sheetColumnLayout = createEditorPixelColumnLayout({
        columnCount: input.columnCount,
        columnWidths: input.columnWidths,
        maximumDigitWidth: input.maximumDigitWidth,
    });
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: visibleRows.length,
        columnCount: input.columnCount,
        rowHeaderLabelCount:
            input.rowHeaderLabelCount ?? Math.max(input.rowCount + EDITOR_EXTRA_PADDING_ROWS, 1),
        viewportHeight: normalizedViewport.viewportHeight,
        viewportWidth: normalizedViewport.viewportWidth,
        rowLayout: baseVisibleRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const rowState = createEditorGridDisplayRowState({
        sourceRowCount: input.rowCount,
        visibleRows,
        hiddenRows,
        totalDisplayRowCount: displayGrid.rowCount,
    });
    const rowLayout = getEditorDisplayRowLayout(
        createEditorPixelRowLayout({
            rowCount: rowState.actualRowNumbers.length,
            rowHeights: mapActualRowHeightsToDisplayRowHeights(
                rowState.actualToDisplayRowNumbers,
                input.rowHeights
            ),
        }),
        displayGrid.rowCount
    );
    const columnLayout = getEditorDisplayColumnLayout(sheetColumnLayout, displayGrid.columnCount);
    const frozenCounts = getFrozenEditorCounts({
        rowCount: Math.min(rowState.actualRowNumbers.length, input.rowCount),
        columnCount: input.columnCount,
        freezePane: input.freezePane,
    });
    const visibleFrozenCounts = getVisibleFrozenEditorCounts({
        frozenRowCount: frozenCounts.rowCount,
        frozenColumnCount: frozenCounts.columnCount,
        viewportHeight: normalizedViewport.viewportHeight,
        viewportWidth: normalizedViewport.viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
        rowLayout,
        columnLayout: sheetColumnLayout,
    });
    const contentSize = getEditorContentSize({
        rowCount: displayGrid.rowCount,
        rowLayout,
        columnLayout,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const clampedViewport = {
        ...normalizedViewport,
        scrollTop: clampEditorScrollPosition(
            normalizedViewport.scrollTop,
            Math.max(0, contentSize.height - normalizedViewport.viewportHeight)
        ),
        scrollLeft: clampEditorScrollPosition(
            normalizedViewport.scrollLeft,
            Math.max(0, contentSize.width - normalizedViewport.viewportWidth)
        ),
    };
    const rowWindow = createEditorRowWindow({
        rowLayout,
        totalRows: displayGrid.rowCount,
        frozenRowCount: frozenCounts.rowCount,
        scrollTop: clampedViewport.scrollTop,
        viewportHeight: clampedViewport.viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        columnLayout,
        frozenColumnCount: frozenCounts.columnCount,
        scrollLeft: clampedViewport.scrollLeft,
        viewportWidth: clampedViewport.viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });

    return {
        viewport: clampedViewport,
        rowLayout,
        columnLayout,
        rowState,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
        contentWidth: contentSize.width,
        contentHeight: contentSize.height,
        stickyTopHeight: 28 + getEditorFrozenRowsHeight(rowLayout, frozenCounts.rowCount),
        stickyLeftWidth:
            displayGrid.rowHeaderWidth +
            getEditorFrozenColumnsWidth(columnLayout, frozenCounts.columnCount),
        frozenRowCount: frozenCounts.rowCount,
        frozenColumnCount: frozenCounts.columnCount,
        visibleFrozenRowCount: visibleFrozenCounts.rowCount,
        visibleFrozenColumnCount: visibleFrozenCounts.columnCount,
        displayRowCount: displayGrid.rowCount,
        displayColumnCount: displayGrid.columnCount,
        window: {
            rowNumbers: mapDisplayRowNumbersToActualRowNumbers(rowState, rowWindow.rowNumbers),
            columnNumbers: columnWindow.columnNumbers,
            frozenRowNumbers: mapDisplayRowNumbersToActualRowNumbers(
                rowState,
                createSequentialNumbers(visibleFrozenCounts.rowCount)
            ),
            frozenColumnNumbers: createSequentialNumbers(visibleFrozenCounts.columnCount),
            topSpacerHeight: rowWindow.topSpacerHeight,
            bottomSpacerHeight: rowWindow.bottomSpacerHeight,
            leadingSpacerWidth: columnWindow.leadingSpacerWidth,
            trailingSpacerWidth: columnWindow.trailingSpacerWidth,
        },
    };
}

export function areEditorGridMetricsEqual(
    left: EditorGridMetrics,
    right: EditorGridMetrics
): boolean {
    return (
        left.rowHeaderWidth === right.rowHeaderWidth &&
        left.contentWidth === right.contentWidth &&
        left.contentHeight === right.contentHeight &&
        left.stickyTopHeight === right.stickyTopHeight &&
        left.stickyLeftWidth === right.stickyLeftWidth &&
        left.frozenRowCount === right.frozenRowCount &&
        left.frozenColumnCount === right.frozenColumnCount &&
        left.visibleFrozenRowCount === right.visibleFrozenRowCount &&
        left.visibleFrozenColumnCount === right.visibleFrozenColumnCount &&
        left.displayRowCount === right.displayRowCount &&
        left.displayColumnCount === right.displayColumnCount &&
        left.viewport.viewportHeight === right.viewport.viewportHeight &&
        left.viewport.viewportWidth === right.viewport.viewportWidth &&
        areDisplayRowStatesEqual(left.rowState, right.rowState) &&
        areRowLayoutsEqual(left.rowLayout, right.rowLayout) &&
        areColumnLayoutsEqual(left.columnLayout, right.columnLayout) &&
        areNumberArraysEqual(left.window.rowNumbers, right.window.rowNumbers) &&
        areNumberArraysEqual(left.window.columnNumbers, right.window.columnNumbers) &&
        areNumberArraysEqual(left.window.frozenRowNumbers, right.window.frozenRowNumbers) &&
        areNumberArraysEqual(left.window.frozenColumnNumbers, right.window.frozenColumnNumbers) &&
        left.window.topSpacerHeight === right.window.topSpacerHeight &&
        left.window.bottomSpacerHeight === right.window.bottomSpacerHeight &&
        left.window.leadingSpacerWidth === right.window.leadingSpacerWidth &&
        left.window.trailingSpacerWidth === right.window.trailingSpacerWidth
    );
}

export function reuseEquivalentEditorGridMetrics(
    previous: EditorGridMetrics | null,
    next: EditorGridMetrics
): EditorGridMetrics {
    return previous && areEditorGridMetricsEqual(previous, next) ? previous : next;
}

export function getEditorDisplayRowNumber(
    metrics: EditorGridMetrics,
    actualRowNumber: number
): number | null {
    return metrics.rowState.actualToDisplayRowNumbers[String(actualRowNumber)] ?? null;
}

export function getEditorActualRowNumberAtDisplayRow(
    metrics: EditorGridMetrics,
    displayRowNumber: number
): number | null {
    return getActualRowNumberAtDisplayRow(metrics.rowState, displayRowNumber);
}

export function getEditorGridActualRowTop(
    metrics: EditorGridMetrics,
    actualRowNumber: number
): number | null {
    const displayRowNumber = getEditorDisplayRowNumber(metrics, actualRowNumber);
    return displayRowNumber === null ? null : getEditorRowTop(metrics.rowLayout, displayRowNumber);
}

export function getEditorGridActualRowHeight(
    metrics: EditorGridMetrics,
    actualRowNumber: number
): number | null {
    const displayRowNumber = getEditorDisplayRowNumber(metrics, actualRowNumber);
    return displayRowNumber === null
        ? null
        : getEditorRowHeight(metrics.rowLayout, displayRowNumber);
}
