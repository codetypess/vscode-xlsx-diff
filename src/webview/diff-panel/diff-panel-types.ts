import type { CellDiffStatus } from "../../core/model/types";

export interface DiffPanelFileView {
    title: string;
    path: string;
    sizeLabel: string;
    detailFacts: DiffPanelFileFactView[];
    modifiedLabel: string;
    isReadonly: boolean;
}

export interface DiffPanelFileFactView {
    label: string;
    value: string;
    title?: string;
}

export interface DiffPanelSheetTabView {
    key: string;
    label: string;
    diffRowCount: number;
    diffCellCount: number;
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
}

export interface DiffPanelRenderModel {
    title: string;
    leftFile: DiffPanelFileView;
    rightFile: DiffPanelFileView;
    sheets: DiffPanelSheetTabView[];
    activeSheet: DiffPanelSheetView | null;
}
