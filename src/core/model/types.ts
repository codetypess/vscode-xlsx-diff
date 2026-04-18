export type RowFilterMode = 'all' | 'diffs' | 'same';
export type SheetComparisonKind = 'matched' | 'renamed' | 'added' | 'removed';
export type CellDiffStatus = 'equal' | 'modified' | 'added' | 'removed';

export interface DiffCellLocation {
	key: string;
	rowNumber: number;
	columnNumber: number;
	address: string;
}

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
	detailLabel?: string;
	detailValue?: string;
	titleDetail?: string;
	isReadonly?: boolean;
	sheets: SheetSnapshot[];
}

export interface EditorSelectedCell {
	rowNumber: number;
	columnNumber: number;
}

export interface EditorSelectionView extends EditorSelectedCell {
	key: string;
	address: string;
	value: string;
	formula: string | null;
	isPresent: boolean;
}

export interface EditorGridCellView {
	key: string;
	address: string;
	value: string;
	formula: string | null;
	isPresent: boolean;
	isSelected: boolean;
}

export interface EditorGridRowView {
	rowNumber: number;
	isSelected: boolean;
	cells: EditorGridCellView[];
}

export interface EditorPageSlice {
	currentPage: number;
	totalPages: number;
	totalRows: number;
	visibleRowCount: number;
	rangeLabel: string;
	columns: string[];
	rows: EditorGridRowView[];
}

export interface EditorPanelState {
	activeSheetKey: string | null;
	currentPage: number;
	selectedCell: EditorSelectedCell | null;
}

export interface EditorSheetTabView {
	key: string;
	label: string;
	rowCount: number;
	columnCount: number;
	hasData: boolean;
	isActive: boolean;
}

export interface EditorActiveSheetView {
	key: string;
	label: string;
	rowCount: number;
	columnCount: number;
	hasData: boolean;
	mergedRangeCount: number;
	hasMergedRanges: boolean;
}

export interface EditorWorkbookSummary {
	totalSheets: number;
	totalRows: number;
	totalNonEmptyCells: number;
}

export interface EditorRenderModel {
	title: string;
	file: WorkbookFileView;
	summary: EditorWorkbookSummary;
	activeSheet: EditorActiveSheetView;
	selection: EditorSelectionView | null;
	hasPendingEdits: boolean;
	canSave: boolean;
	canEdit: boolean;
	page: EditorPageSlice;
	sheets: EditorSheetTabView[];
	canPrevPage: boolean;
	canNextPage: boolean;
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
	diffCells: DiffCellLocation[];
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
	diffTone: CellDiffStatus;
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
	highlightedDiffCell: DiffCellLocation | null;
	mergedRangesChanged: boolean;
}

export interface PanelState {
	activeSheetKey: string | null;
	filter: RowFilterMode;
	currentPage: number;
	highlightedDiffCellKey: string | null;
}

export interface SheetTabView {
	key: string;
	label: string;
	kind: SheetComparisonKind;
	diffRowCount: number;
	diffCellCount: number;
	mergedRangesChanged: boolean;
	hasDiff: boolean;
	diffTone: CellDiffStatus;
	isActive: boolean;
}

export interface WorkbookFileView {
	fileName: string;
	filePath: string;
	fileSizeLabel: string;
	detailLabel?: string;
	detailValue?: string;
	modifiedTimeLabel: string;
	isReadonly: boolean;
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
