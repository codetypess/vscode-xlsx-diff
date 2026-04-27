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

export function normalizeCellTextLineEndings(value: string): string {
    return value.replace(/\r\n?/g, "\n");
}
