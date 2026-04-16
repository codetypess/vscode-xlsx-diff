export function createCellKey(rowNumber: number, columnNumber: number): string {
	return `${rowNumber}:${columnNumber}`;
}

export function getColumnLabel(columnNumber: number): string {
	let current = columnNumber;
	let label = '';

	while (current > 0) {
		const remainder = (current - 1) % 26;
		label = String.fromCharCode(65 + remainder) + label;
		current = Math.floor((current - 1) / 26);
	}

	return label;
}

export function getCellAddress(rowNumber: number, columnNumber: number): string {
	return `${getColumnLabel(columnNumber)}${rowNumber}`;
}
