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

export interface DiffColumnAlignment {
    columnNumber: number;
    leftColumnNumber: number | null;
    rightColumnNumber: number | null;
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
    detailFacts?: WorkbookDetailFact[];
    titleDetail?: string;
    isReadonly?: boolean;
    sheets: SheetSnapshot[];
}

export interface WorkbookDetailFact {
    label: string;
    value: string;
    titleValue?: string;
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

export interface EditorPanelState {
    activeSheetKey: string | null;
    selectedCell: EditorSelectedCell | null;
}

export interface EditorSheetTabView {
    key: string;
    label: string;
    isActive: boolean;
}

export interface EditorActiveSheetView {
    key: string;
    rowCount: number;
    columnCount: number;
    columns: string[];
    cells: Record<string, CellSnapshot>;
    freezePane: SheetFreezePaneSnapshot | null;
}

export interface EditorRenderModel {
    title: string;
    activeSheet: EditorActiveSheetView;
    selection: EditorSelectionView | null;
    hasPendingEdits: boolean;
    canEdit: boolean;
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
    alignedColumns: DiffColumnAlignment[];
    diffRows: number[];
    diffCells: DiffCellLocation[];
    diffCellCount: number;
    mergedRangesChanged: boolean;
    freezePaneChanged: boolean;
}

export interface WorkbookDiffModel {
    left: WorkbookSnapshot;
    right: WorkbookSnapshot;
    sheets: SheetDiffModel[];
    totalDiffSheets: number;
    totalDiffRows: number;
    totalDiffCells: number;
}
