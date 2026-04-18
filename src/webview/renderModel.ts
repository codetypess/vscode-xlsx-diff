import { DEFAULT_PAGE_SIZE } from '../constants';
import { createPageSlice } from '../core/paging/createPageSlice';
import { isChineseDisplayLanguage } from '../displayLanguage';
import {
	type CellDiffStatus,
	type PanelState,
	type RenderModel,
	type RowFilterMode,
	type SheetDiffModel,
	type SheetTabView,
	type WorkbookDiffModel,
	type WorkbookSnapshot,
} from '../core/model/types';

function getUntitledSheetLabel(): string {
	return isChineseDisplayLanguage() ? '未命名工作表' : 'Untitled Sheet';
}

function getWorkbookTitle(workbook: WorkbookSnapshot): string {
	return workbook.titleDetail
		? `${workbook.fileName} @ ${workbook.titleDetail}`
		: workbook.fileName;
}

function formatFileSize(bytes: number): string {
	if (bytes < 1024) {
		return `${bytes} B`;
	}

	const units = ['KB', 'MB', 'GB'];
	let value = bytes / 1024;
	let index = 0;

	while (value >= 1024 && index < units.length - 1) {
		value /= 1024;
		index += 1;
	}

	return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[index]}`;
}

function formatModifiedTime(value: string): string {
	return new Intl.DateTimeFormat(undefined, {
		dateStyle: 'medium',
		timeStyle: 'short',
	}).format(new Date(value));
}

function getSheetLabel(sheet: SheetDiffModel): string {
	if (sheet.kind === 'renamed') {
		return `${sheet.leftSheetName} -> ${sheet.rightSheetName}`;
	}

	return sheet.rightSheetName ?? sheet.leftSheetName ?? getUntitledSheetLabel();
}

function getSheetHasDiff(sheet: SheetDiffModel): boolean {
	return (
		sheet.kind !== 'matched' ||
		sheet.diffCellCount > 0 ||
		sheet.mergedRangesChanged
	);
}

function getSheetDiffTone(sheet: SheetDiffModel): CellDiffStatus {
	if (sheet.kind === 'added') {
		return 'added';
	}

	if (sheet.kind === 'removed') {
		return 'removed';
	}

	return getSheetHasDiff(sheet) ? 'modified' : 'equal';
}

function getFilteredRowCount(
	sheet: SheetDiffModel,
	filter: RowFilterMode,
): number {
	switch (filter) {
		case 'diffs':
			return sheet.diffRows.length;
		case 'same':
			return sheet.kind === 'added' || sheet.kind === 'removed'
				? 0
				: Math.max(0, sheet.rowCount - sheet.diffRows.length);
		case 'all':
		default:
			return sheet.rowCount;
	}
}

function clampPage(
	sheet: SheetDiffModel,
	filter: RowFilterMode,
	page: number,
): number {
	const totalRows = getFilteredRowCount(sheet, filter);
	const totalPages = Math.max(
		1,
		Math.ceil(Math.max(totalRows, 1) / DEFAULT_PAGE_SIZE),
	);
	return Math.min(Math.max(page, 1), totalPages);
}

export function createInitialPanelState(diff: WorkbookDiffModel): PanelState {
	const firstSheet = diff.sheets[0];

	return {
		activeSheetKey: firstSheet?.key ?? null,
		filter: 'all',
		currentPage: 1,
		highlightedDiffRow: firstSheet?.diffRows[0] ?? null,
	};
}

export function normalizePanelState(
	diff: WorkbookDiffModel,
	state: PanelState,
): PanelState {
	const activeSheet =
		diff.sheets.find((sheet) => sheet.key === state.activeSheetKey) ??
		diff.sheets[0] ??
		null;

	if (!activeSheet) {
		return {
			activeSheetKey: null,
			filter: 'all',
			currentPage: 1,
			highlightedDiffRow: null,
		};
	}

	const highlightedDiffRow =
		state.filter === 'same' || activeSheet.diffRows.length === 0
			? null
			: activeSheet.diffRows.includes(state.highlightedDiffRow ?? -1)
				? state.highlightedDiffRow
				: activeSheet.diffRows[0] ?? null;

	return {
		activeSheetKey: activeSheet.key,
		filter: state.filter,
		currentPage: clampPage(activeSheet, state.filter, state.currentPage),
		highlightedDiffRow,
	};
}

export function setActiveSheet(
	diff: WorkbookDiffModel,
	state: PanelState,
	sheetKey: string,
): PanelState {
	const activeSheet = diff.sheets.find((sheet) => sheet.key === sheetKey);
	if (!activeSheet) {
		return state;
	}

	return normalizePanelState(diff, {
		activeSheetKey: activeSheet.key,
		filter: state.filter,
		currentPage: 1,
		highlightedDiffRow: activeSheet.diffRows[0] ?? null,
	});
}

export function setFilterMode(
	diff: WorkbookDiffModel,
	state: PanelState,
	filter: RowFilterMode,
): PanelState {
	return normalizePanelState(diff, {
		...state,
		filter,
		currentPage: 1,
		highlightedDiffRow: filter === 'same' ? null : state.highlightedDiffRow,
	});
}

export function setCurrentPage(
	diff: WorkbookDiffModel,
	state: PanelState,
	currentPage: number,
): PanelState {
	return normalizePanelState(diff, {
		...state,
		currentPage,
	});
}

export function moveDiffCursor(
	diff: WorkbookDiffModel,
	state: PanelState,
	direction: -1 | 1,
): PanelState {
	const normalizedState = normalizePanelState(diff, state);
	const activeSheet = diff.sheets.find(
		(sheet) => sheet.key === normalizedState.activeSheetKey,
	);

	if (!activeSheet || activeSheet.diffRows.length === 0) {
		return normalizedState;
	}

	const filter = normalizedState.filter === 'same' ? 'diffs' : normalizedState.filter;
	const currentIndex = normalizedState.highlightedDiffRow
		? activeSheet.diffRows.indexOf(normalizedState.highlightedDiffRow)
		: direction > 0
			? -1
			: activeSheet.diffRows.length;
	const nextIndex = Math.min(
		Math.max(currentIndex + direction, 0),
		activeSheet.diffRows.length - 1,
	);
	const nextHighlightedRow = activeSheet.diffRows[nextIndex];
	const nextPage =
		filter === 'diffs'
			? Math.floor(nextIndex / DEFAULT_PAGE_SIZE) + 1
			: Math.floor((nextHighlightedRow - 1) / DEFAULT_PAGE_SIZE) + 1;

	return normalizePanelState(diff, {
		activeSheetKey: activeSheet.key,
		filter,
		currentPage: nextPage,
		highlightedDiffRow: nextHighlightedRow,
	});
}

export function createRenderModel(
	diff: WorkbookDiffModel,
	state: PanelState,
): RenderModel {
	const normalizedState = normalizePanelState(diff, state);
	const activeSheet =
		diff.sheets.find((sheet) => sheet.key === normalizedState.activeSheetKey) ??
		diff.sheets[0];
	const page = createPageSlice(
		activeSheet,
		normalizedState.filter,
		normalizedState.currentPage,
		normalizedState.highlightedDiffRow,
	);
	const currentDiffIndex =
		normalizedState.highlightedDiffRow === null
			? -1
			: activeSheet.diffRows.indexOf(normalizedState.highlightedDiffRow);

	const sheets: SheetTabView[] = diff.sheets.map((sheet) => ({
		key: sheet.key,
		label: getSheetLabel(sheet),
		kind: sheet.kind,
		diffRowCount: sheet.diffRows.length,
		diffCellCount: sheet.diffCellCount,
		mergedRangesChanged: sheet.mergedRangesChanged,
		hasDiff: getSheetHasDiff(sheet),
		diffTone: getSheetDiffTone(sheet),
		isActive: sheet.key === activeSheet.key,
	}));

	return {
		title: `${getWorkbookTitle(diff.left)} ↔ ${getWorkbookTitle(diff.right)}`,
		leftFile: {
			fileName: diff.left.fileName,
			filePath: diff.left.filePath,
			fileSizeLabel: formatFileSize(diff.left.fileSize),
			detailLabel: diff.left.detailLabel,
			detailValue: diff.left.detailValue,
			modifiedTimeLabel:
				diff.left.modifiedTimeLabel ?? formatModifiedTime(diff.left.modifiedTime),
			isReadonly: diff.left.isReadonly ?? false,
		},
		rightFile: {
			fileName: diff.right.fileName,
			filePath: diff.right.filePath,
			fileSizeLabel: formatFileSize(diff.right.fileSize),
			detailLabel: diff.right.detailLabel,
			detailValue: diff.right.detailValue,
			modifiedTimeLabel:
				diff.right.modifiedTimeLabel ?? formatModifiedTime(diff.right.modifiedTime),
			isReadonly: diff.right.isReadonly ?? false,
		},
		summary: {
			totalSheets: diff.sheets.length,
			diffSheets: diff.totalDiffSheets,
			diffRows: diff.totalDiffRows,
			diffCells: diff.totalDiffCells,
		},
		activeSheet: {
			key: activeSheet.key,
			label: getSheetLabel(activeSheet),
			kind: activeSheet.kind,
			leftName: activeSheet.leftSheetName,
			rightName: activeSheet.rightSheetName,
			hasDiff: getSheetHasDiff(activeSheet),
			mergedRangesChanged: activeSheet.mergedRangesChanged,
		},
		filter: normalizedState.filter,
		page,
		sheets,
		canPrevPage: page.currentPage > 1,
		canNextPage: page.currentPage < page.totalPages,
		canPrevDiff: currentDiffIndex > 0,
		canNextDiff:
			currentDiffIndex >= 0 && currentDiffIndex < activeSheet.diffRows.length - 1,
	};
}
