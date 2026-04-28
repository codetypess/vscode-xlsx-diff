import { DEFAULT_EDITOR_WINDOW_OVERSCAN } from "../../constants";
import {
    DEFAULT_COLUMN_PIXEL_WIDTH,
    DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
    type PixelColumnLayout,
    createPixelColumnLayout,
    extendPixelColumnLayout,
    findColumnIndexForOffset,
    getPixelColumnLeft,
    getPixelColumnOffset,
    getPixelColumnWidth,
} from "../column-layout";
import {
    DEFAULT_ROW_PIXEL_HEIGHT,
    type PixelRowLayout,
    createPixelRowLayout,
    extendPixelRowLayout,
    getPixelRowHeight,
    getPixelRowOffset,
    getPixelRowTop,
    getPixelRowWindow,
} from "../row-layout";

export const EDITOR_VIRTUAL_ROW_HEIGHT = DEFAULT_ROW_PIXEL_HEIGHT;
export const EDITOR_VIRTUAL_COLUMN_WIDTH = DEFAULT_COLUMN_PIXEL_WIDTH;
export const EDITOR_VIRTUAL_HEADER_HEIGHT = 28;
export const EDITOR_VIRTUAL_COLUMN_OVERSCAN = 8;
export const EDITOR_EXTRA_PADDING_ROWS = 8;
export const EDITOR_EXTRA_PADDING_COLUMNS = 3;

function clamp(value: number, min: number, max: number): number {
    return Math.min(Math.max(value, min), max);
}

export function getEditorRowHeaderWidth(totalRows: number): number {
    const digits = String(Math.max(totalRows, 1)).length;
    return Math.max(56, digits * 9 + 24);
}

export function createEditorPixelColumnLayout({
    columnCount,
    columnWidths,
    maximumDigitWidth = DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
}: {
    columnCount: number;
    columnWidths?: readonly (number | null)[];
    maximumDigitWidth?: number;
}): PixelColumnLayout {
    return createPixelColumnLayout({
        columnWidths: (columnWidths ?? []).slice(0, columnCount),
        totalColumnCount: columnCount,
        maximumDigitWidth,
        fallbackPixelWidth: EDITOR_VIRTUAL_COLUMN_WIDTH,
    });
}

export function createEditorPixelRowLayout({
    rowCount,
    rowHeights,
}: {
    rowCount: number;
    rowHeights?: Readonly<Record<string, number | null>>;
}): PixelRowLayout {
    return createPixelRowLayout({
        rowHeights,
        totalRowCount: rowCount,
        fallbackPixelHeight: EDITOR_VIRTUAL_ROW_HEIGHT,
    });
}

export function getEditorDisplayColumnLayout(
    layout: PixelColumnLayout,
    totalColumnCount: number
): PixelColumnLayout {
    return extendPixelColumnLayout(layout, totalColumnCount);
}

export function getEditorDisplayRowLayout(
    layout: PixelRowLayout,
    totalRowCount: number
): PixelRowLayout {
    return extendPixelRowLayout(layout, totalRowCount);
}

export function getEditorColumnLeft(layout: PixelColumnLayout, columnNumber: number): number {
    return getPixelColumnLeft(layout, columnNumber);
}

export function getEditorColumnWidth(layout: PixelColumnLayout, columnNumber: number): number {
    return getPixelColumnWidth(layout, columnNumber);
}

export function getEditorFrozenColumnsWidth(
    layout: PixelColumnLayout,
    frozenColumnCount: number
): number {
    return getPixelColumnOffset(layout, frozenColumnCount);
}

export function getEditorRowTop(layout: PixelRowLayout, rowNumber: number): number {
    return getPixelRowTop(layout, rowNumber);
}

export function getEditorRowHeight(layout: PixelRowLayout, rowNumber: number): number {
    return getPixelRowHeight(layout, rowNumber);
}

export function getEditorFrozenRowsHeight(
    layout: PixelRowLayout,
    frozenRowCount: number
): number {
    return getPixelRowOffset(layout, frozenRowCount);
}

export function getMinimumVisibleEditorRowCount(
    viewportHeight: number,
    rowLayout: PixelRowLayout
): number {
    const availableHeight = Math.max(0, viewportHeight - EDITOR_VIRTUAL_HEADER_HEIGHT);
    if (availableHeight <= 0) {
        return 1;
    }

    let visibleHeight = 0;
    let visibleRowCount = 0;
    while (visibleHeight < availableHeight) {
        visibleRowCount += 1;
        visibleHeight += getEditorRowHeight(rowLayout, visibleRowCount);
    }

    return Math.max(1, visibleRowCount);
}

export function getMinimumVisibleEditorColumnCount({
    viewportWidth,
    rowHeaderWidth,
    columnLayout,
}: {
    viewportWidth: number;
    rowHeaderWidth: number;
    columnLayout: PixelColumnLayout;
}): number {
    const availableWidth = Math.max(0, viewportWidth - rowHeaderWidth);
    if (availableWidth <= 0) {
        return 1;
    }

    let visibleWidth = 0;
    let visibleColumnCount = 0;
    while (visibleWidth < availableWidth) {
        visibleColumnCount += 1;
        visibleWidth += getEditorColumnWidth(columnLayout, visibleColumnCount);
    }

    return Math.max(1, visibleColumnCount);
}

export function getEditorDisplayGridDimensions({
    rowCount,
    columnCount,
    rowHeaderLabelCount = rowCount,
    viewportHeight,
    viewportWidth,
    rowLayout,
    columnLayout,
}: {
    rowCount: number;
    columnCount: number;
    rowHeaderLabelCount?: number;
    viewportHeight: number;
    viewportWidth: number;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
}): { rowCount: number; columnCount: number; rowHeaderWidth: number } {
    const displayRowCount =
        Math.max(rowCount, getMinimumVisibleEditorRowCount(viewportHeight, rowLayout)) +
        EDITOR_EXTRA_PADDING_ROWS;
    const rowHeaderWidth = getEditorRowHeaderWidth(
        Math.max(displayRowCount, rowHeaderLabelCount)
    );

    return {
        rowCount: displayRowCount,
        columnCount:
            Math.max(
                columnCount,
                getMinimumVisibleEditorColumnCount({
                    viewportWidth,
                    rowHeaderWidth,
                    columnLayout,
                })
            ) + EDITOR_EXTRA_PADDING_COLUMNS,
        rowHeaderWidth,
    };
}

export function getFrozenEditorCounts({
    rowCount,
    columnCount,
    freezePane,
}: {
    rowCount: number;
    columnCount: number;
    freezePane:
        | {
              rowCount: number;
              columnCount: number;
          }
        | null
        | undefined;
}): { rowCount: number; columnCount: number } {
    return {
        rowCount: Math.max(0, Math.min(freezePane?.rowCount ?? 0, Math.max(rowCount - 1, 0))),
        columnCount: Math.max(
            0,
            Math.min(freezePane?.columnCount ?? 0, Math.max(columnCount - 1, 0))
        ),
    };
}

export function getVisibleFrozenEditorCounts({
    frozenRowCount,
    frozenColumnCount,
    viewportHeight,
    viewportWidth,
    rowHeaderWidth,
    rowLayout,
    columnLayout,
}: {
    frozenRowCount: number;
    frozenColumnCount: number;
    viewportHeight: number;
    viewportWidth: number;
    rowHeaderWidth: number;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
}): { rowCount: number; columnCount: number } {
    const availableFrozenHeight = Math.max(0, viewportHeight - EDITOR_VIRTUAL_HEADER_HEIGHT);
    let visibleFrozenRowCount = 0;
    let consumedFrozenHeight = 0;
    while (visibleFrozenRowCount < frozenRowCount && consumedFrozenHeight < availableFrozenHeight) {
        visibleFrozenRowCount += 1;
        consumedFrozenHeight += getEditorRowHeight(rowLayout, visibleFrozenRowCount);
    }

    const availableFrozenWidth = Math.max(0, viewportWidth - rowHeaderWidth);
    let visibleFrozenColumnCount = 0;
    let consumedFrozenWidth = 0;
    while (
        visibleFrozenColumnCount < frozenColumnCount &&
        consumedFrozenWidth < availableFrozenWidth
    ) {
        visibleFrozenColumnCount += 1;
        consumedFrozenWidth += getEditorColumnWidth(columnLayout, visibleFrozenColumnCount);
    }

    return {
        rowCount: Math.min(frozenRowCount, visibleFrozenRowCount),
        columnCount: Math.min(frozenColumnCount, visibleFrozenColumnCount),
    };
}

export function createEditorRowWindow({
    rowLayout,
    totalRows,
    frozenRowCount,
    scrollTop,
    viewportHeight,
    overscan = DEFAULT_EDITOR_WINDOW_OVERSCAN,
}: {
    rowLayout: PixelRowLayout;
    totalRows: number;
    frozenRowCount: number;
    scrollTop: number;
    viewportHeight: number;
    overscan?: number;
}): {
    rowNumbers: number[];
    startRowNumber: number;
    endRowNumber: number;
    topSpacerHeight: number;
    bottomSpacerHeight: number;
} {
    const scrollableRowCount = Math.max(0, totalRows - frozenRowCount);
    if (scrollableRowCount <= 0) {
        return {
            rowNumbers: [],
            startRowNumber: frozenRowCount + 1,
            endRowNumber: frozenRowCount,
            topSpacerHeight: 0,
            bottomSpacerHeight: 0,
        };
    }

    const frozenRowsHeight = getEditorFrozenRowsHeight(rowLayout, frozenRowCount);
    const scrollableViewportHeight = Math.max(
        0,
        viewportHeight - EDITOR_VIRTUAL_HEADER_HEIGHT - frozenRowsHeight
    );
    const rowWindow = getPixelRowWindow(
        rowLayout,
        frozenRowsHeight + scrollTop,
        scrollableViewportHeight,
        overscan
    );
    const startIndex = Math.max(frozenRowCount, rowWindow.startIndex);
    const endIndex = Math.max(startIndex, rowWindow.endIndex);
    const rowNumbers = Array.from(
        { length: Math.max(0, endIndex - startIndex) },
        (_, index) => startIndex + index + 1
    );
    const totalScrollableHeight = Math.max(0, rowLayout.totalHeight - frozenRowsHeight);

    return {
        rowNumbers,
        startRowNumber: rowNumbers[0] ?? frozenRowCount + 1,
        endRowNumber: rowNumbers[rowNumbers.length - 1] ?? frozenRowCount,
        topSpacerHeight: Math.max(0, getPixelRowOffset(rowLayout, startIndex) - frozenRowsHeight),
        bottomSpacerHeight: Math.max(
            0,
            totalScrollableHeight -
                Math.max(0, getPixelRowOffset(rowLayout, endIndex) - frozenRowsHeight)
        ),
    };
}

export function createEditorColumnWindow({
    columnLayout,
    frozenColumnCount,
    scrollLeft,
    viewportWidth,
    rowHeaderWidth,
    overscan = EDITOR_VIRTUAL_COLUMN_OVERSCAN,
}: {
    columnLayout: PixelColumnLayout;
    frozenColumnCount: number;
    scrollLeft: number;
    viewportWidth: number;
    rowHeaderWidth: number;
    overscan?: number;
}): {
    columnNumbers: number[];
    startColumnNumber: number;
    endColumnNumber: number;
    leadingSpacerWidth: number;
    trailingSpacerWidth: number;
} {
    const scrollableColumnCount = Math.max(0, columnLayout.totalColumnCount - frozenColumnCount);
    if (scrollableColumnCount <= 0) {
        return {
            columnNumbers: [],
            startColumnNumber: frozenColumnCount + 1,
            endColumnNumber: frozenColumnCount,
            leadingSpacerWidth: 0,
            trailingSpacerWidth: 0,
        };
    }

    const frozenColumnsWidth = getEditorFrozenColumnsWidth(columnLayout, frozenColumnCount);
    const stickyWidth = rowHeaderWidth + frozenColumnsWidth;
    const scrollableViewportWidth = Math.max(
        EDITOR_VIRTUAL_COLUMN_WIDTH,
        viewportWidth - stickyWidth
    );
    const startVisibleIndex = Math.max(
        frozenColumnCount,
        findColumnIndexForOffset(columnLayout, frozenColumnsWidth + scrollLeft)
    );
    const endVisibleIndex = Math.max(
        frozenColumnCount,
        findColumnIndexForOffset(
            columnLayout,
            frozenColumnsWidth + Math.max(scrollLeft, scrollLeft + scrollableViewportWidth - 1)
        )
    );
    const startIndex = Math.max(frozenColumnCount, startVisibleIndex - overscan);
    const endIndex = Math.min(columnLayout.totalColumnCount, endVisibleIndex + overscan + 1);
    const columnNumbers = Array.from(
        { length: Math.max(0, endIndex - startIndex) },
        (_, index) => startIndex + index + 1
    );

    return {
        columnNumbers,
        startColumnNumber: columnNumbers[0] ?? frozenColumnCount + 1,
        endColumnNumber: columnNumbers[columnNumbers.length - 1] ?? frozenColumnCount,
        leadingSpacerWidth: Math.max(0, getPixelColumnOffset(columnLayout, startIndex) - frozenColumnsWidth),
        trailingSpacerWidth: Math.max(
            0,
            columnLayout.totalWidth -
                getPixelColumnOffset(columnLayout, endIndex) -
                frozenColumnsWidth
        ),
    };
}

export function getEditorContentSize({
    rowCount,
    rowLayout,
    columnLayout,
    rowHeaderWidth,
}: {
    rowCount: number;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
    rowHeaderWidth: number;
}): { width: number; height: number } {
    return {
        width: rowHeaderWidth + columnLayout.totalWidth,
        height:
            EDITOR_VIRTUAL_HEADER_HEIGHT +
            getEditorDisplayRowLayout(rowLayout, rowCount).totalHeight,
    };
}

export function getEditorScrollPositionForCell({
    rowNumber,
    columnNumber,
    frozenRowCount,
    frozenColumnCount,
    viewportHeight,
    viewportWidth,
    rowHeaderWidth,
    rowLayout,
    columnLayout,
}: {
    rowNumber: number;
    columnNumber: number;
    frozenRowCount: number;
    frozenColumnCount: number;
    viewportHeight: number;
    viewportWidth: number;
    rowHeaderWidth: number;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
}): { top: number | null; left: number | null } {
    const stickyTop =
        EDITOR_VIRTUAL_HEADER_HEIGHT + getEditorFrozenRowsHeight(rowLayout, frozenRowCount);
    const frozenColumnsWidth = getEditorFrozenColumnsWidth(columnLayout, frozenColumnCount);
    const stickyLeft = rowHeaderWidth + frozenColumnsWidth;

    return {
        top:
            rowNumber <= frozenRowCount
                ? null
                : Math.max(
                      0,
                      getEditorRowTop(rowLayout, rowNumber) -
                          getEditorFrozenRowsHeight(rowLayout, frozenRowCount) -
                          Math.max(
                              0,
                              viewportHeight -
                                  stickyTop -
                                  getEditorRowHeight(rowLayout, rowNumber)
                          )
                  ),
        left:
            columnNumber <= frozenColumnCount
                ? null
                : Math.max(
                      0,
                      getEditorColumnLeft(columnLayout, columnNumber) -
                          frozenColumnsWidth -
                          Math.max(
                              0,
                              viewportWidth -
                                  stickyLeft -
                                  getEditorColumnWidth(columnLayout, columnNumber)
                          )
                  ),
    };
}

export function getEditorVirtualCellKey(rowNumber: number, columnNumber: number): string {
    return `${rowNumber}:${columnNumber}`;
}

export function clampEditorScrollPosition(value: number, maxValue: number): number {
    return clamp(value, 0, Math.max(maxValue, 0));
}
