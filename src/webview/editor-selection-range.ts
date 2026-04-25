export interface SelectionPositionLike {
    rowNumber: number;
    columnNumber: number;
}

export interface SelectionRange {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

export function createSelectionRange(
    anchorCell: SelectionPositionLike | null,
    focusCell: SelectionPositionLike | null
): SelectionRange | null {
    if (!anchorCell || !focusCell) {
        return null;
    }

    return {
        startRow: Math.min(anchorCell.rowNumber, focusCell.rowNumber),
        endRow: Math.max(anchorCell.rowNumber, focusCell.rowNumber),
        startColumn: Math.min(anchorCell.columnNumber, focusCell.columnNumber),
        endColumn: Math.max(anchorCell.columnNumber, focusCell.columnNumber),
    };
}

export function hasExpandedSelectionRange(range: SelectionRange | null): boolean {
    return Boolean(
        range && (range.startRow !== range.endRow || range.startColumn !== range.endColumn)
    );
}

export function createRowSelectionRange(
    rowNumber: number,
    columnCount: number
): SelectionRange | null {
    if (!Number.isInteger(rowNumber) || !Number.isInteger(columnCount)) {
        return null;
    }

    if (rowNumber < 1 || columnCount < 1) {
        return null;
    }

    return {
        startRow: rowNumber,
        endRow: rowNumber,
        startColumn: 1,
        endColumn: columnCount,
    };
}

export function createColumnSelectionRange(
    columnNumber: number,
    rowCount: number
): SelectionRange | null {
    if (!Number.isInteger(columnNumber) || !Number.isInteger(rowCount)) {
        return null;
    }

    if (columnNumber < 1 || rowCount < 1) {
        return null;
    }

    return {
        startRow: 1,
        endRow: rowCount,
        startColumn: columnNumber,
        endColumn: columnNumber,
    };
}
