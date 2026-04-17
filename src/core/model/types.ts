export type RowFilterMode = 'all' | 'diffs' | 'same';
export type SheetComparisonKind = 'matched' | 'renamed' | 'added' | 'removed';
export type CellDiffStatus = 'equal' | 'modified' | 'added' | 'removed';

export interface CellSnapshot {
	key: string;
	rowNumber: number;
	columnNumber: number;
	address: string;
	displayValue: string;
	formula: string | null;
	styleId: number | null;
}

export interface SheetSnapshot {
	name: string;
	rowCount: number;
	columnCount: number;
	mergedRanges: string[];
	cells: Record<string, CellSnapshot>;
	signature: string;
}

export interface WorkbookSnapshot {
	filePath: string;
	fileName: string;
	fileSize: number;
	modifiedTime: string;
	modifiedTimeLabel?: string;
	sheets: SheetSnapshot[];
}

export interface SheetDiffModel {
	key: string;
	kind: SheetComparisonKind;
	leftSheet: SheetSnapshot | null;
	rightSheet: SheetSnapshot | null;
	leftSheetName: string | null;
	rightSheetName: string | null;
	rowCount: number;
	columnCount: number;
	diffRows: number[];
	diffCellCount: number;
	mergedRangesChanged: boolean;
}

export interface WorkbookDiffModel {
	left: WorkbookSnapshot;
	right: WorkbookSnapshot;
	sheets: SheetDiffModel[];
	totalDiffSheets: number;
	totalDiffRows: number;
	totalDiffCells: number;
}

export interface GridCellView {
	key: string;
	address: string;
	status: CellDiffStatus;
	leftPresent: boolean;
	rightPresent: boolean;
	leftValue: string;
	rightValue: string;
	leftFormula: string | null;
	rightFormula: string | null;
}

export interface GridRowView {
	rowNumber: number;
	hasDiff: boolean;
	isHighlighted: boolean;
	cells: GridCellView[];
}

export interface PageSlice {
	filter: RowFilterMode;
	currentPage: number;
	totalPages: number;
	totalRows: number;
	visibleRowCount: number;
	rangeLabel: string;
	columns: string[];
	rows: GridRowView[];
	diffRowCount: number;
	diffCellCount: number;
	sameRowCount: number;
	highlightedDiffRow: number | null;
	mergedRangesChanged: boolean;
}

export interface PanelState {
	activeSheetKey: string | null;
	filter: RowFilterMode;
	currentPage: number;
	highlightedDiffRow: number | null;
}

export interface SheetTabView {
	key: string;
	label: string;
	kind: SheetComparisonKind;
	diffRowCount: number;
	diffCellCount: number;
	mergedRangesChanged: boolean;
	hasDiff: boolean;
	isActive: boolean;
}

export interface WorkbookFileView {
	fileName: string;
	filePath: string;
	fileSizeLabel: string;
	modifiedTimeLabel: string;
}

export interface RenderModel {
	title: string;
	leftFile: WorkbookFileView;
	rightFile: WorkbookFileView;
	summary: {
		totalSheets: number;
		diffSheets: number;
		diffRows: number;
		diffCells: number;
	};
	activeSheet: {
		key: string;
		label: string;
		kind: SheetComparisonKind;
		leftName: string | null;
		rightName: string | null;
		hasDiff: boolean;
		mergedRangesChanged: boolean;
	};
	filter: RowFilterMode;
	page: PageSlice;
	sheets: SheetTabView[];
	canPrevPage: boolean;
	canNextPage: boolean;
	canPrevDiff: boolean;
	canNextDiff: boolean;
}
