import { createHash } from 'node:crypto';
import { stat } from 'node:fs/promises';
import * as path from 'node:path';
import { createCellKey, getCellAddress } from '../model/cells';
import {
	type CellSnapshot,
	type SheetSnapshot,
	type WorkbookSnapshot,
} from '../model/types';

interface SheetReader {
	name: string;
	rowCount: number;
	columnCount: number;
	getDisplayValue(rowNumber: number, columnNumber: number): string | null;
	getFormula(rowNumber: number, columnNumber: number): string | null;
	getStyleId(rowNumber: number, columnNumber: number): number | null;
	getMergedRanges(): string[];
}

interface WorkbookReader {
	getSheet(sheetName: string): SheetReader;
	getSheetNames(): string[];
}

function createSheetSignature(sheet: SheetSnapshot): string {
	const hash = createHash('sha1');
	hash.update(`${sheet.name}\n`);
	hash.update(`${sheet.rowCount}:${sheet.columnCount}\n`);

	for (const mergedRange of sheet.mergedRanges) {
		hash.update(`merge:${mergedRange}\n`);
	}

	for (const cell of Object.values(sheet.cells).sort((left, right) =>
		left.key.localeCompare(right.key),
	)) {
		hash.update(
			`${cell.address}\u0000${cell.displayValue}\u0000${cell.formula ?? ''}\n`,
		);
	}

	return hash.digest('hex');
}

function loadSheetSnapshot(workbook: WorkbookReader, sheetName: string): SheetSnapshot {
	const sheet = workbook.getSheet(sheetName);
	const cells: Record<string, CellSnapshot> = {};

	for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
		for (
			let columnNumber = 1;
			columnNumber <= sheet.columnCount;
			columnNumber += 1
		) {
			const displayValue = sheet.getDisplayValue(rowNumber, columnNumber);
			const formula = sheet.getFormula(rowNumber, columnNumber);

			if (displayValue === null && formula === null) {
				continue;
			}

			const key = createCellKey(rowNumber, columnNumber);
			cells[key] = {
				key,
				rowNumber,
				columnNumber,
				address: getCellAddress(rowNumber, columnNumber),
				displayValue: displayValue ?? '',
				formula,
				styleId: sheet.getStyleId(rowNumber, columnNumber),
			};
		}
	}

	const snapshot: SheetSnapshot = {
		name: sheet.name,
		rowCount: sheet.rowCount,
		columnCount: sheet.columnCount,
		mergedRanges: [...sheet.getMergedRanges()].sort((left, right) =>
			left.localeCompare(right),
		),
		cells,
		signature: '',
	};

	snapshot.signature = createSheetSignature(snapshot);
	return snapshot;
}

export async function loadWorkbookSnapshot(
	filePath: string,
): Promise<WorkbookSnapshot> {
	const { Workbook } = await import('fastxlsx');
	const [workbook, fileStats] = await Promise.all([
		Workbook.open(filePath),
		stat(filePath),
	]);
	const sheets = workbook
		.getSheetNames()
		.map((sheetName) => loadSheetSnapshot(workbook, sheetName));

	return {
		filePath,
		fileName: path.basename(filePath),
		fileSize: fileStats.size,
		modifiedTime: fileStats.mtime.toISOString(),
		sheets,
	};
}
