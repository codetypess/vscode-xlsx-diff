/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from 'assert';
import { buildWorkbookDiff } from '../core/diff/buildWorkbookDiff';
import type {
	CellSnapshot,
	SheetSnapshot,
	WorkbookSnapshot,
} from '../core/model/types';
import { createInitialPanelState, createRenderModel } from '../webview/renderModel';

function createCell(
	rowNumber: number,
	columnNumber: number,
	displayValue: string,
): CellSnapshot {
	return {
		key: `${rowNumber}:${columnNumber}`,
		rowNumber,
		columnNumber,
		address: `R${rowNumber}C${columnNumber}`,
		displayValue,
		formula: null,
		styleId: null,
	};
}

function createSheet(name: string, cells: CellSnapshot[]): SheetSnapshot {
	return {
		name,
		rowCount: 1,
		columnCount: 1,
		mergedRanges: [],
		cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
		signature: `${name}-signature`,
	};
}

function createWorkbook(
	overrides: Partial<WorkbookSnapshot>,
	cellValue: string,
): WorkbookSnapshot {
	return {
		filePath: '/tmp/item.xlsx',
		fileName: 'item.xlsx',
		fileSize: 128,
		modifiedTime: new Date('2026-04-18T06:51:00.000Z').toISOString(),
		sheets: [createSheet('item', [createCell(1, 1, cellValue)])],
		...overrides,
	};
}

suite('Render model', () => {
	test('keeps commit detail, title hints, and read-only state aligned', () => {
		const left = createWorkbook(
			{
				detailLabel: 'Commit',
				detailValue: 'd4ce7e0',
				titleDetail: 'd4ce7e0',
				modifiedTimeLabel: 'Apr 18, 2026, 6:51 AM',
				isReadonly: true,
			},
			'left',
		);
		const right = createWorkbook({}, 'right');

		const diff = buildWorkbookDiff(left, right);
		const renderModel = createRenderModel(diff, createInitialPanelState(diff));

		assert.strictEqual(renderModel.title, 'item.xlsx @ d4ce7e0 ↔ item.xlsx');
		assert.deepStrictEqual(renderModel.leftFile, {
			fileName: 'item.xlsx',
			filePath: '/tmp/item.xlsx',
			fileSizeLabel: '128 B',
			detailLabel: 'Commit',
			detailValue: 'd4ce7e0',
			modifiedTimeLabel: 'Apr 18, 2026, 6:51 AM',
			isReadonly: true,
		});
		assert.strictEqual(renderModel.rightFile.isReadonly, false);
	});
});