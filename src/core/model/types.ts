export type SheetComparisonKind = "matched" | "renamed" | "added" | "removed";
export type CellDiffStatus = "equal" | "modified" | "added" | "removed";

export interface DiffCellLocation {
    key: string;
    rowNumber: number;
    columnNumber: number;
    address: string;
    diffIndex: number;
}

export interface DiffRowAlignment {
    rowNumber: number;
    leftRowNumber: number | null;
    rightRowNumber: number | null;
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

export interface SheetFreezePaneSnapshot {
    columnCount: number;
    rowCount: number;
    topLeftCell: string;
    activePane: "bottomLeft" | "topRight" | "bottomRight" | null;
}

export interface SheetSnapshot {
    name: string;
    rowCount: number;
    columnCount: number;
    mergedRanges: string[];
    freezePane?: SheetFreezePaneSnapshot | null;
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
    totalRows: number;
    visibleRowCount: number;
    rangeLabel: string;
    startRow: number;
    endRow: number;
    columns: string[];
    frozenRows: EditorGridRowView[];
    rows: EditorGridRowView[];
}

export interface EditorPanelState {
    activeSheetKey: string | null;
    viewportStartRow: number;
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
    freezePane: SheetFreezePaneSnapshot | null;
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
    canUndoStructuralEdits: boolean;
    canRedoStructuralEdits: boolean;
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
    alignedRows: DiffRowAlignment[];
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

export interface WorkbookFileView {
    fileName: string;
    filePath: string;
    fileSizeLabel: string;
    detailLabel?: string;
    detailValue?: string;
    modifiedTimeLabel: string;
    isReadonly: boolean;
}
