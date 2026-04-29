import { type SelectionRange } from "./editor-selection-range";

function findFirstVisibleValueInRange(
    values: readonly number[],
    start: number,
    end: number
): number | null {
    return values.find((value) => value >= start && value <= end) ?? null;
}

function findLastVisibleValueInRange(
    values: readonly number[],
    start: number,
    end: number
): number | null {
    for (let index = values.length - 1; index >= 0; index -= 1) {
        const value = values[index];
        if (value === undefined) {
            continue;
        }

        if (value >= start && value <= end) {
            return value;
        }
    }

    return null;
}

export function clipSelectionRangeToVisibleGrid(
    selectionRange: SelectionRange | null,
    visibleRowNumbers: readonly number[],
    visibleColumnNumbers: readonly number[]
): SelectionRange | null {
    if (!selectionRange || visibleRowNumbers.length === 0 || visibleColumnNumbers.length === 0) {
        return null;
    }

    const startRow = findFirstVisibleValueInRange(
        visibleRowNumbers,
        selectionRange.startRow,
        selectionRange.endRow
    );
    const endRow = findLastVisibleValueInRange(
        visibleRowNumbers,
        selectionRange.startRow,
        selectionRange.endRow
    );
    const startColumn = findFirstVisibleValueInRange(
        visibleColumnNumbers,
        selectionRange.startColumn,
        selectionRange.endColumn
    );
    const endColumn = findLastVisibleValueInRange(
        visibleColumnNumbers,
        selectionRange.startColumn,
        selectionRange.endColumn
    );

    if (
        startRow === null ||
        endRow === null ||
        startColumn === null ||
        endColumn === null
    ) {
        return null;
    }

    return {
        startRow,
        endRow,
        startColumn,
        endColumn,
    };
}
