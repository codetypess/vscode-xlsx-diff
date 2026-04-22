import type { CellDiffStatus, SheetComparisonKind } from "../core/model/types";

export interface DiffPanelFileView {
    title: string;
    path: string;
    sizeLabel: string;
    detailLabel?: string;
    detailValue?: string;
    modifiedLabel: string;
    isReadonly: boolean;
}

export interface DiffPanelSummaryView {
    totalSheets: number;
    diffSheets: number;
    diffRows: number;
    diffCells: number;
}

export interface DiffPanelSheetTabView {
    key: string;
    label: string;
    kind: SheetComparisonKind;
    rowCount: number;
    columnCount: number;
    diffRowCount: number;
    diffCellCount: number;
    mergedRangesChanged: boolean;
    hasDiff: boolean;
    diffTone: CellDiffStatus;
    isActive: boolean;
}

export interface DiffPanelDiffCellView {
    key: string;
    rowNumber: number;
    columnNumber: number;
    address: string;
    diffIndex: number;
}

export interface DiffPanelSparseCellView {
    key: string;
    columnNumber: number;
    address: string;
    status: CellDiffStatus;
    diffIndex: number | null;
    leftPresent: boolean;
    rightPresent: boolean;
    leftValue: string;
    rightValue: string;
    leftFormula: string | null;
    rightFormula: string | null;
}

export interface DiffPanelRowView {
    rowNumber: number;
    leftRowNumber: number | null;
    rightRowNumber: number | null;
    hasDiff: boolean;
    diffTone: CellDiffStatus;
    cells: DiffPanelSparseCellView[];
}

export interface DiffPanelSheetView {
    key: string;
    label: string;
    kind: SheetComparisonKind;
    leftName: string | null;
    rightName: string | null;
    rowCount: number;
    columnCount: number;
    columns: string[];
    rows: DiffPanelRowView[];
    diffRows: number[];
    diffCells: DiffPanelDiffCellView[];
    diffRowCount: number;
    diffCellCount: number;
    mergedRangesChanged: boolean;
}

export interface DiffPanelRenderModel {
    title: string;
    leftFile: DiffPanelFileView;
    rightFile: DiffPanelFileView;
    summary: DiffPanelSummaryView;
    sheets: DiffPanelSheetTabView[];
    activeSheet: DiffPanelSheetView | null;
}
