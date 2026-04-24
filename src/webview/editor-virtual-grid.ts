import { DEFAULT_EDITOR_WINDOW_OVERSCAN } from "../constants";

export const EDITOR_VIRTUAL_ROW_HEIGHT = 28;
export const EDITOR_VIRTUAL_COLUMN_WIDTH = 120;
export const EDITOR_VIRTUAL_HEADER_HEIGHT = 28;
export const EDITOR_VIRTUAL_COLUMN_OVERSCAN = 8;

function clamp(value: number, min: number, max: number): number {
    return Math.min(Math.max(value, min), max);
}

export function getEditorRowHeaderWidth(totalRows: number): number {
    const digits = String(Math.max(totalRows, 1)).length;
    return Math.max(56, digits * 9 + 24);
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
}: {
    frozenRowCount: number;
    frozenColumnCount: number;
    viewportHeight: number;
    viewportWidth: number;
    rowHeaderWidth: number;
}): { rowCount: number; columnCount: number } {
    const visibleFrozenRowCount = Math.max(
        0,
        Math.floor(
            Math.max(0, viewportHeight - EDITOR_VIRTUAL_HEADER_HEIGHT) / EDITOR_VIRTUAL_ROW_HEIGHT
        )
    );
    const visibleFrozenColumnCount = Math.max(
        0,
        Math.floor(
            Math.max(0, viewportWidth - rowHeaderWidth) / EDITOR_VIRTUAL_COLUMN_WIDTH
        )
    );

    return {
        rowCount: Math.min(frozenRowCount, visibleFrozenRowCount),
        columnCount: Math.min(frozenColumnCount, visibleFrozenColumnCount),
    };
}

export function createEditorRowWindow({
    totalRows,
    frozenRowCount,
    scrollTop,
    viewportHeight,
    overscan = DEFAULT_EDITOR_WINDOW_OVERSCAN,
}: {
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

    const scrollableViewportHeight = Math.max(
        0,
        viewportHeight - EDITOR_VIRTUAL_HEADER_HEIGHT - frozenRowCount * EDITOR_VIRTUAL_ROW_HEIGHT
    );
    const startIndex = Math.max(
        0,
        Math.floor(scrollTop / EDITOR_VIRTUAL_ROW_HEIGHT) - overscan
    );
    const visibleRowCount =
        Math.ceil(
            Math.max(scrollableViewportHeight, EDITOR_VIRTUAL_ROW_HEIGHT) /
                EDITOR_VIRTUAL_ROW_HEIGHT
        ) +
        overscan * 2;
    const endIndex = Math.min(scrollableRowCount, startIndex + visibleRowCount);
    const rowNumbers = Array.from(
        { length: Math.max(0, endIndex - startIndex) },
        (_, index) => frozenRowCount + startIndex + index + 1
    );

    return {
        rowNumbers,
        startRowNumber: rowNumbers[0] ?? frozenRowCount + 1,
        endRowNumber: rowNumbers[rowNumbers.length - 1] ?? frozenRowCount,
        topSpacerHeight: startIndex * EDITOR_VIRTUAL_ROW_HEIGHT,
        bottomSpacerHeight: Math.max(0, (scrollableRowCount - endIndex) * EDITOR_VIRTUAL_ROW_HEIGHT),
    };
}

export function createEditorColumnWindow({
    totalColumns,
    frozenColumnCount,
    scrollLeft,
    viewportWidth,
    rowHeaderWidth,
    overscan = EDITOR_VIRTUAL_COLUMN_OVERSCAN,
}: {
    totalColumns: number;
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
    const scrollableColumnCount = Math.max(0, totalColumns - frozenColumnCount);
    if (scrollableColumnCount <= 0) {
        return {
            columnNumbers: [],
            startColumnNumber: frozenColumnCount + 1,
            endColumnNumber: frozenColumnCount,
            leadingSpacerWidth: 0,
            trailingSpacerWidth: 0,
        };
    }

    const stickyWidth =
        rowHeaderWidth + frozenColumnCount * EDITOR_VIRTUAL_COLUMN_WIDTH;
    const scrollableViewportWidth = Math.max(
        EDITOR_VIRTUAL_COLUMN_WIDTH,
        viewportWidth - stickyWidth
    );
    const startIndex = Math.max(
        0,
        Math.floor(scrollLeft / EDITOR_VIRTUAL_COLUMN_WIDTH) - overscan
    );
    const visibleColumnCount =
        Math.ceil(scrollableViewportWidth / EDITOR_VIRTUAL_COLUMN_WIDTH) + overscan * 2;
    const endIndex = Math.min(scrollableColumnCount, startIndex + visibleColumnCount);
    const columnNumbers = Array.from(
        { length: Math.max(0, endIndex - startIndex) },
        (_, index) => frozenColumnCount + startIndex + index + 1
    );

    return {
        columnNumbers,
        startColumnNumber: columnNumbers[0] ?? frozenColumnCount + 1,
        endColumnNumber: columnNumbers[columnNumbers.length - 1] ?? frozenColumnCount,
        leadingSpacerWidth: startIndex * EDITOR_VIRTUAL_COLUMN_WIDTH,
        trailingSpacerWidth: Math.max(
            0,
            (scrollableColumnCount - endIndex) * EDITOR_VIRTUAL_COLUMN_WIDTH
        ),
    };
}

export function getEditorContentSize({
    rowCount,
    columnCount,
    rowHeaderWidth,
}: {
    rowCount: number;
    columnCount: number;
    rowHeaderWidth: number;
}): { width: number; height: number } {
    return {
        width: rowHeaderWidth + columnCount * EDITOR_VIRTUAL_COLUMN_WIDTH,
        height: EDITOR_VIRTUAL_HEADER_HEIGHT + rowCount * EDITOR_VIRTUAL_ROW_HEIGHT,
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
}: {
    rowNumber: number;
    columnNumber: number;
    frozenRowCount: number;
    frozenColumnCount: number;
    viewportHeight: number;
    viewportWidth: number;
    rowHeaderWidth: number;
}): { top: number | null; left: number | null } {
    const stickyTop =
        EDITOR_VIRTUAL_HEADER_HEIGHT + frozenRowCount * EDITOR_VIRTUAL_ROW_HEIGHT;
    const stickyLeft = rowHeaderWidth + frozenColumnCount * EDITOR_VIRTUAL_COLUMN_WIDTH;

    return {
        top:
            rowNumber <= frozenRowCount
                ? null
                : Math.max(
                      0,
                      (rowNumber - frozenRowCount - 1) * EDITOR_VIRTUAL_ROW_HEIGHT -
                          Math.max(
                              0,
                              viewportHeight - stickyTop - EDITOR_VIRTUAL_ROW_HEIGHT
                          )
                  ),
        left:
            columnNumber <= frozenColumnCount
                ? null
                : Math.max(
                      0,
                      (columnNumber - frozenColumnCount - 1) * EDITOR_VIRTUAL_COLUMN_WIDTH -
                          Math.max(
                              0,
                              viewportWidth - stickyLeft - EDITOR_VIRTUAL_COLUMN_WIDTH
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
