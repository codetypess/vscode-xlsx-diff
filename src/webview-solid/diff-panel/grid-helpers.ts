import type {
    DiffPanelColumnView,
    DiffPanelRowView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "../../webview/diff-panel/diff-panel-types";
import type { DiffWebviewOutgoingMessage } from "../shared/session-protocol";

export type RowFilterMode = "all" | "diffs" | "same";

export interface SelectedDiffCell {
    rowNumber: number;
    columnNumber: number;
    sourceRowNumber: number | null;
    sourceColumnNumber: number | null;
}

export interface DraftEditState {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    sourceRowNumber: number | null;
    sourceColumnNumber: number | null;
    side: "left" | "right";
    value: string;
    modelValue: string;
}

export interface PendingDiffEdit {
    sheetKey: string;
    side: "left" | "right";
    rowNumber: number;
    columnNumber: number;
    sourceRowNumber: number | null;
    sourceColumnNumber: number | null;
    value: string;
}

export interface DiffPreviewState {
    address: string;
    rowNumber: number;
    columnNumber: number;
    leftValue: string;
    rightValue: string;
    leftPresent: boolean;
    rightPresent: boolean;
    index: number;
    total: number;
}

export const DIFF_GRID_COLUMN_WIDTH_PX = 120;

export function getPendingDiffEditKey(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    side: "left" | "right"
): string {
    return `${sheetKey}:${side}:${rowNumber}:${columnNumber}`;
}

export function getSourceRowNumber(
    row: DiffPanelRowView | null,
    side: "left" | "right"
): number | null {
    return side === "left" ? (row?.leftRowNumber ?? null) : (row?.rightRowNumber ?? null);
}

export function createSelectedDiffCell(
    sheet: DiffPanelSheetView,
    row: DiffPanelRowView | null,
    rowNumber: number,
    columnNumber: number,
    side: "left" | "right"
): SelectedDiffCell {
    const column = sheet.columns[columnNumber - 1];

    return {
        rowNumber,
        columnNumber,
        sourceRowNumber: getSourceRowNumber(row, side),
        sourceColumnNumber:
            side === "left" ? (column?.leftColumnNumber ?? null) : (column?.rightColumnNumber ?? null),
    };
}

export function filterDiffRows(rows: DiffPanelRowView[], mode: RowFilterMode): DiffPanelRowView[] {
    switch (mode) {
        case "all":
            return rows;
        case "same":
            return rows.filter((row) => !row.hasDiff);
        case "diffs":
        default:
            return rows.filter((row) => row.hasDiff);
    }
}

export function getWrappedDiffIndex(
    currentIndex: number,
    diffCount: number,
    offset: number
): number {
    if (diffCount <= 0) {
        return 0;
    }

    const nextIndex = currentIndex + offset;
    if (nextIndex < 0) {
        return diffCount - 1;
    }

    if (nextIndex >= diffCount) {
        return 0;
    }

    return nextIndex;
}

export function getDiffTrackWidth(
    columns: DiffPanelColumnView[],
    columnWidthPx = DIFF_GRID_COLUMN_WIDTH_PX
): number {
    return columns.length * columnWidthPx;
}

export function clampDiffHorizontalScroll(
    scrollLeft: number,
    trackWidth: number,
    viewportWidth: number
): number {
    const maxScrollLeft = Math.max(trackWidth - viewportWidth, 0);
    return Math.max(0, Math.min(scrollLeft, maxScrollLeft));
}

export function getDiffPreviewState(
    activeSheet: DiffPanelSheetView | null,
    activeDiffIndex: number
): DiffPreviewState | null {
    if (!activeSheet || activeSheet.diffCells.length === 0) {
        return null;
    }

    const safeIndex = Math.max(0, Math.min(activeDiffIndex, activeSheet.diffCells.length - 1));
    const diffCell = activeSheet.diffCells[safeIndex];
    const row = activeSheet.rows.find((candidate) => candidate.rowNumber === diffCell.rowNumber);
    const cell = row?.cells.find((candidate) => candidate.columnNumber === diffCell.columnNumber);

    return {
        address: diffCell.address,
        rowNumber: diffCell.rowNumber,
        columnNumber: diffCell.columnNumber,
        leftValue: cell?.leftValue ?? "",
        rightValue: cell?.rightValue ?? "",
        leftPresent: cell?.leftPresent ?? false,
        rightPresent: cell?.rightPresent ?? false,
        index: safeIndex,
        total: activeSheet.diffCells.length,
    };
}

export function getSelectedDiffCellState(
    sheet: DiffPanelSheetView,
    row: DiffPanelRowView,
    cell: DiffPanelSparseCellView,
    side: "left" | "right"
): {
    selectedCell: SelectedDiffCell;
    activeDiffIndex: number | null;
} {
    return {
        selectedCell: createSelectedDiffCell(sheet, row, row.rowNumber, cell.columnNumber, side),
        activeDiffIndex: cell.diffIndex,
    };
}

export function beginDiffCellEdit({
    activeSheetKey,
    side,
    sideEditable,
    pendingEdits,
    selection,
    cell,
}: {
    activeSheetKey: string | null;
    side: "left" | "right";
    sideEditable: boolean;
    pendingEdits: Record<string, PendingDiffEdit>;
    selection: SelectedDiffCell;
    cell: DiffPanelSparseCellView | null;
}): {
    editingCell: DraftEditState;
    selectedCell: SelectedDiffCell;
} | null {
    if (!activeSheetKey || !sideEditable) {
        return null;
    }

    if (selection.sourceRowNumber === null || selection.sourceColumnNumber === null) {
        return null;
    }

    const pendingEdit = pendingEdits[
        getPendingDiffEditKey(
            activeSheetKey,
            selection.rowNumber,
            selection.columnNumber,
            side
        )
    ];
    const modelValue = side === "left" ? (cell?.leftValue ?? "") : (cell?.rightValue ?? "");

    return {
        editingCell: {
            sheetKey: activeSheetKey,
            rowNumber: selection.rowNumber,
            columnNumber: selection.columnNumber,
            sourceRowNumber: selection.sourceRowNumber,
            sourceColumnNumber: selection.sourceColumnNumber,
            side,
            value: pendingEdit?.value ?? modelValue,
            modelValue,
        },
        selectedCell: selection,
    };
}

export function applyPendingDiffEdit(
    pendingEdits: Record<string, PendingDiffEdit>,
    draft: DraftEditState
): Record<string, PendingDiffEdit> {
    const key = getPendingDiffEditKey(
        draft.sheetKey,
        draft.rowNumber,
        draft.columnNumber,
        draft.side
    );

    if (draft.value === draft.modelValue) {
        if (!Object.hasOwn(pendingEdits, key)) {
            return pendingEdits;
        }

        const nextPendingEdits = { ...pendingEdits };
        delete nextPendingEdits[key];
        return nextPendingEdits;
    }

    return {
        ...pendingEdits,
        [key]: {
            sheetKey: draft.sheetKey,
            side: draft.side,
            rowNumber: draft.rowNumber,
            columnNumber: draft.columnNumber,
            sourceRowNumber: draft.sourceRowNumber,
            sourceColumnNumber: draft.sourceColumnNumber,
            value: draft.value,
        },
    };
}

export function finalizeDiffCellEdit(
    pendingEdits: Record<string, PendingDiffEdit>,
    draft: DraftEditState | null,
    disposition: "commit" | "cancel"
): {
    pendingEdits: Record<string, PendingDiffEdit>;
    editingCell: DraftEditState | null;
} {
    if (!draft) {
        return {
            pendingEdits,
            editingCell: null,
        };
    }

    if (disposition === "cancel") {
        return {
            pendingEdits,
            editingCell: null,
        };
    }

    return {
        pendingEdits: applyPendingDiffEdit(pendingEdits, draft),
        editingCell: null,
    };
}

export function getRenderedDiffCellValue(
    pendingEdits: Record<string, PendingDiffEdit>,
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    cell: DiffPanelSparseCellView | null,
    side: "left" | "right"
): string {
    const pendingEdit = pendingEdits[
        getPendingDiffEditKey(sheetKey, rowNumber, columnNumber, side)
    ];
    if (pendingEdit) {
        return pendingEdit.value;
    }

    return side === "left" ? (cell?.leftValue ?? "") : (cell?.rightValue ?? "");
}

export function createSaveDiffEditsMessage(
    pendingEdits: Record<string, PendingDiffEdit>
): DiffWebviewOutgoingMessage | null {
    const edits = Object.values(pendingEdits)
        .filter(
            (edit): edit is PendingDiffEdit & {
                sourceRowNumber: number;
                sourceColumnNumber: number;
            } => edit.sourceRowNumber !== null && edit.sourceColumnNumber !== null
        )
        .map((edit) => ({
            sheetKey: edit.sheetKey,
            side: edit.side,
            rowNumber: edit.sourceRowNumber,
            columnNumber: edit.sourceColumnNumber,
            value: edit.value,
        }));
    if (edits.length === 0) {
        return null;
    }

    return {
        type: "saveEdits",
        edits,
    };
}
