import type { SelectionPositionLike, SelectionRange } from "./editor-selection-range";
import type { SheetAutoFilterSnapshot } from "../../core/model/types";

export type EditorFilterSortDirection = "asc" | "desc";

export interface EditorFilterCellSource {
    rowCount: number;
    columnCount: number;
    getCellValue(rowNumber: number, columnNumber: number): string;
}

export interface EditorSheetFilterState {
    range: SelectionRange;
    sort: {
        columnNumber: number;
        direction: EditorFilterSortDirection;
    } | null;
    includedValuesByColumn: Record<string, readonly string[]>;
}

export interface EditorFilterValueOption {
    value: string;
    count: number;
}

export interface EditorVisibleRowResult {
    visibleRows: number[];
    hiddenRows: number[];
}

const FILTER_VALUE_COLLATOR = new Intl.Collator(undefined, {
    numeric: true,
    sensitivity: "base",
});

function hasFilterCellContent(value: string): boolean {
    return value.length > 0;
}

function rowHasContent(
    source: EditorFilterCellSource,
    rowNumber: number,
    startColumn = 1,
    endColumn = source.columnCount
): boolean {
    for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
        if (hasFilterCellContent(source.getCellValue(rowNumber, columnNumber))) {
            return true;
        }
    }

    return false;
}

function columnHasContent(
    source: EditorFilterCellSource,
    columnNumber: number,
    startRow = 1,
    endRow = source.rowCount
): boolean {
    for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
        if (hasFilterCellContent(source.getCellValue(rowNumber, columnNumber))) {
            return true;
        }
    }

    return false;
}

function getContiguousRowRange(
    source: EditorFilterCellSource,
    rowNumber: number,
    startColumn = 1,
    endColumn = source.columnCount
): { startRow: number; endRow: number } | null {
    if (
        rowNumber < 1 ||
        rowNumber > source.rowCount ||
        !rowHasContent(source, rowNumber, startColumn, endColumn)
    ) {
        return null;
    }

    let startRow = rowNumber;
    let endRow = rowNumber;
    while (startRow > 1 && rowHasContent(source, startRow - 1, startColumn, endColumn)) {
        startRow -= 1;
    }
    while (endRow < source.rowCount && rowHasContent(source, endRow + 1, startColumn, endColumn)) {
        endRow += 1;
    }

    return { startRow, endRow };
}

function getContiguousColumnRange(
    source: EditorFilterCellSource,
    columnNumber: number,
    startRow = 1,
    endRow = source.rowCount
): { startColumn: number; endColumn: number } | null {
    if (
        columnNumber < 1 ||
        columnNumber > source.columnCount ||
        !columnHasContent(source, columnNumber, startRow, endRow)
    ) {
        return null;
    }

    let startColumn = columnNumber;
    let endColumn = columnNumber;
    while (startColumn > 1 && columnHasContent(source, startColumn - 1, startRow, endRow)) {
        startColumn -= 1;
    }
    while (
        endColumn < source.columnCount &&
        columnHasContent(source, endColumn + 1, startRow, endRow)
    ) {
        endColumn += 1;
    }

    return { startColumn, endColumn };
}

function getRangeColumnNumbers(range: SelectionRange): number[] {
    return Array.from(
        { length: Math.max(0, range.endColumn - range.startColumn + 1) },
        (_, index) => range.startColumn + index
    );
}

function getRangeDataRows(range: SelectionRange): number[] {
    return Array.from(
        { length: Math.max(0, range.endRow - range.startRow) },
        (_, index) => range.startRow + index + 1
    );
}

function compareFilterValues(left: string, right: string): number {
    if (!left && !right) {
        return 0;
    }

    if (!left) {
        return 1;
    }

    if (!right) {
        return -1;
    }

    return FILTER_VALUE_COLLATOR.compare(left, right);
}

function normalizeIncludedValues(values: readonly string[]): readonly string[] {
    const normalized = [...new Set(values)].sort(compareFilterValues);
    return normalized;
}

function passesColumnFilters(
    source: EditorFilterCellSource,
    filterState: EditorSheetFilterState,
    rowNumber: number
): boolean {
    return getRangeColumnNumbers(filterState.range).every((columnNumber) => {
        const includedValues = filterState.includedValuesByColumn[String(columnNumber)];
        if (!includedValues) {
            return true;
        }

        return includedValues.includes(source.getCellValue(rowNumber, columnNumber));
    });
}

function compareRowsBySortColumn(
    source: EditorFilterCellSource,
    rowNumber: number,
    otherRowNumber: number,
    sort: NonNullable<EditorSheetFilterState["sort"]>
): number {
    const leftValue = source.getCellValue(rowNumber, sort.columnNumber);
    const rightValue = source.getCellValue(otherRowNumber, sort.columnNumber);
    const compared = compareFilterValues(leftValue, rightValue);

    if (compared !== 0) {
        return sort.direction === "asc" ? compared : -compared;
    }

    return rowNumber - otherRowNumber;
}

export function normalizeEditorFilterRange(
    source: Pick<EditorFilterCellSource, "rowCount" | "columnCount">,
    range: SelectionRange | null
): SelectionRange | null {
    if (!range) {
        return null;
    }

    const startRow = Math.max(1, Math.min(range.startRow, source.rowCount));
    const endRow = Math.max(1, Math.min(range.endRow, source.rowCount));
    const startColumn = Math.max(1, Math.min(range.startColumn, source.columnCount));
    const endColumn = Math.max(1, Math.min(range.endColumn, source.columnCount));

    if (startRow >= endRow || startColumn > endColumn) {
        return null;
    }

    return {
        startRow,
        endRow,
        startColumn,
        endColumn,
    };
}

export function canCreateEditorFilterRange(
    source: Pick<EditorFilterCellSource, "rowCount" | "columnCount">,
    range: SelectionRange | null
): boolean {
    return Boolean(normalizeEditorFilterRange(source, range));
}

export function resolveEditorFilterRangeFromSelection(
    source: Pick<EditorFilterCellSource, "rowCount" | "columnCount">,
    range: SelectionRange | null
): SelectionRange | null {
    if (!range) {
        return null;
    }

    const normalizedRange = normalizeEditorFilterRange(source, range);
    if (normalizedRange) {
        return normalizedRange;
    }

    if (range.startRow === range.endRow) {
        return normalizeEditorFilterRange(source, {
            startRow: range.startRow,
            endRow: source.rowCount,
            startColumn: range.startColumn,
            endColumn: range.endColumn,
        });
    }

    return null;
}

export function resolveEditorFilterRangeFromActiveCell(
    source: EditorFilterCellSource,
    cell: SelectionPositionLike | null
): SelectionRange | null {
    if (!cell) {
        return null;
    }

    const initialRowRange = getContiguousRowRange(source, cell.rowNumber);
    if (!initialRowRange) {
        return null;
    }

    const initialColumnRange = getContiguousColumnRange(
        source,
        cell.columnNumber,
        initialRowRange.startRow,
        initialRowRange.endRow
    );
    if (!initialColumnRange) {
        return null;
    }

    const refinedRowRange = getContiguousRowRange(
        source,
        cell.rowNumber,
        initialColumnRange.startColumn,
        initialColumnRange.endColumn
    );
    if (!refinedRowRange) {
        return null;
    }

    const refinedColumnRange = getContiguousColumnRange(
        source,
        cell.columnNumber,
        refinedRowRange.startRow,
        refinedRowRange.endRow
    );
    if (!refinedColumnRange) {
        return null;
    }

    return normalizeEditorFilterRange(source, {
        startRow: refinedRowRange.startRow,
        endRow: refinedRowRange.endRow,
        startColumn: refinedColumnRange.startColumn,
        endColumn: refinedColumnRange.endColumn,
    });
}

export function createEditorSheetFilterState(
    source: Pick<EditorFilterCellSource, "rowCount" | "columnCount">,
    range: SelectionRange | null
): EditorSheetFilterState | null {
    const normalizedRange = resolveEditorFilterRangeFromSelection(source, range);
    if (!normalizedRange) {
        return null;
    }

    return {
        range: normalizedRange,
        sort: null,
        includedValuesByColumn: {},
    };
}

export function createEditorSheetFilterStateFromSnapshot(
    snapshot: SheetAutoFilterSnapshot | null
): EditorSheetFilterState | null {
    if (!snapshot) {
        return null;
    }

    return {
        range: {
            ...snapshot.range,
        },
        sort: snapshot.sort
            ? {
                  ...snapshot.sort,
              }
            : null,
        includedValuesByColumn: {},
    };
}

export function createEditorSheetFilterSnapshot(
    filterState: EditorSheetFilterState | null
): SheetAutoFilterSnapshot | null {
    if (!filterState) {
        return null;
    }

    return {
        range: {
            ...filterState.range,
        },
        sort: filterState.sort
            ? {
                  ...filterState.sort,
              }
            : null,
    };
}

export function toggleEditorSheetFilterState(
    source: Pick<EditorFilterCellSource, "rowCount" | "columnCount">,
    filterState: EditorSheetFilterState | null,
    range: SelectionRange | null
): EditorSheetFilterState | null {
    if (filterState) {
        return null;
    }

    return createEditorSheetFilterState(source, range);
}

export function isEditorFilterHeaderCell(
    filterState: EditorSheetFilterState | null,
    rowNumber: number,
    columnNumber: number
): boolean {
    if (!filterState) {
        return false;
    }

    return (
        rowNumber === filterState.range.startRow &&
        columnNumber >= filterState.range.startColumn &&
        columnNumber <= filterState.range.endColumn
    );
}

export function getEditorFilterColumnValues(
    source: EditorFilterCellSource,
    filterState: EditorSheetFilterState,
    columnNumber: number
): EditorFilterValueOption[] {
    if (
        columnNumber < filterState.range.startColumn ||
        columnNumber > filterState.range.endColumn
    ) {
        return [];
    }

    const counts = new Map<string, number>();
    for (const rowNumber of getRangeDataRows(filterState.range)) {
        const value = source.getCellValue(rowNumber, columnNumber);
        counts.set(value, (counts.get(value) ?? 0) + 1);
    }

    return [...counts.entries()]
        .sort(([leftValue], [rightValue]) => compareFilterValues(leftValue, rightValue))
        .map(([value, count]) => ({ value, count }));
}

export function getEditorVisibleRows(
    source: EditorFilterCellSource,
    filterState: EditorSheetFilterState | null
): EditorVisibleRowResult {
    if (!filterState) {
        return {
            visibleRows: Array.from({ length: source.rowCount }, (_, index) => index + 1),
            hiddenRows: [],
        };
    }

    const visibleRows = Array.from(
        { length: Math.max(0, filterState.range.startRow - 1) },
        (_, index) => index + 1
    );
    const dataRows = getRangeDataRows(filterState.range);
    const filteredRows = dataRows.filter((rowNumber) =>
        passesColumnFilters(source, filterState, rowNumber)
    );

    if (filterState.sort) {
        filteredRows.sort((leftRowNumber, rightRowNumber) =>
            compareRowsBySortColumn(source, leftRowNumber, rightRowNumber, filterState.sort!)
        );
    }

    visibleRows.push(filterState.range.startRow, ...filteredRows);
    for (
        let rowNumber = filterState.range.endRow + 1;
        rowNumber <= source.rowCount;
        rowNumber += 1
    ) {
        visibleRows.push(rowNumber);
    }

    return {
        visibleRows,
        hiddenRows: dataRows.filter((rowNumber) => !filteredRows.includes(rowNumber)),
    };
}

export function updateEditorFilterIncludedValues(
    filterState: EditorSheetFilterState,
    columnNumber: number,
    includedValues: readonly string[] | null
): EditorSheetFilterState {
    const nextIncludedValuesByColumn = { ...filterState.includedValuesByColumn };
    if (!includedValues) {
        delete nextIncludedValuesByColumn[String(columnNumber)];
    } else {
        nextIncludedValuesByColumn[String(columnNumber)] = normalizeIncludedValues(includedValues);
    }

    return {
        ...filterState,
        includedValuesByColumn: nextIncludedValuesByColumn,
    };
}

export function updateEditorFilterSort(
    filterState: EditorSheetFilterState,
    columnNumber: number,
    direction: EditorFilterSortDirection | null
): EditorSheetFilterState {
    if (!direction) {
        if (filterState.sort?.columnNumber !== columnNumber) {
            return filterState;
        }

        return {
            ...filterState,
            sort: null,
        };
    }

    return {
        ...filterState,
        sort: {
            columnNumber,
            direction,
        },
    };
}

export function clearEditorFilterColumn(
    filterState: EditorSheetFilterState,
    columnNumber: number
): EditorSheetFilterState {
    const nextFilterState = updateEditorFilterIncludedValues(filterState, columnNumber, null);
    return nextFilterState.sort?.columnNumber === columnNumber
        ? {
              ...nextFilterState,
              sort: null,
          }
        : nextFilterState;
}
