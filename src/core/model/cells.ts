export function createCellKey(rowNumber: number, columnNumber: number): string {
    return `${rowNumber}:${columnNumber}`;
}

export function getColumnLabel(columnNumber: number): string {
    let current = columnNumber;
    let label = "";

    while (current > 0) {
        const remainder = (current - 1) % 26;
        label = String.fromCharCode(65 + remainder) + label;
        current = Math.floor((current - 1) / 26);
    }

    return label;
}

export function getColumnNumber(columnLabel: string): number | null {
    const normalized = columnLabel.trim().toUpperCase();
    if (!/^[A-Z]+$/.test(normalized)) {
        return null;
    }

    let value = 0;
    for (const character of normalized) {
        value = value * 26 + (character.charCodeAt(0) - 64);
    }

    return value;
}

export function getCellAddress(rowNumber: number, columnNumber: number): string {
    return `${getColumnLabel(columnNumber)}${rowNumber}`;
}

export function parseCellAddress(
    address: string
): { rowNumber: number; columnNumber: number } | null {
    const normalizedAddress = address.trim().toUpperCase().replaceAll("$", "");
    const match = /^([A-Z]+)(\d+)$/.exec(normalizedAddress);
    if (!match) {
        return null;
    }

    const columnNumber = getColumnNumber(match[1]);
    const rowNumber = Number(match[2]);
    if (columnNumber === null || !Number.isInteger(rowNumber) || rowNumber < 1) {
        return null;
    }

    return {
        rowNumber,
        columnNumber,
    };
}

export function parseRangeAddress(range: string): {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
} | null {
    const [startAddress, endAddress = startAddress] = range.split(":", 2);
    const start = parseCellAddress(startAddress);
    const end = parseCellAddress(endAddress);
    if (!start || !end) {
        return null;
    }

    return {
        startRow: Math.min(start.rowNumber, end.rowNumber),
        endRow: Math.max(start.rowNumber, end.rowNumber),
        startColumn: Math.min(start.columnNumber, end.columnNumber),
        endColumn: Math.max(start.columnNumber, end.columnNumber),
    };
}

export function getRangeAddress(range: {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}): string {
    return `${getCellAddress(range.startRow, range.startColumn)}:${getCellAddress(
        range.endRow,
        range.endColumn
    )}`;
}

export function normalizeCellTextLineEndings(value: string): string {
    return value.replace(/\r\n?/g, "\n");
}

export function hasComparableCellContent(
    displayValue: string | null | undefined,
    formula: string | null | undefined
): boolean {
    return (
        (formula !== null && formula !== undefined) ||
        (displayValue !== null && displayValue !== undefined && displayValue !== "")
    );
}
