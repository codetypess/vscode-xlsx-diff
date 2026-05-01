import type {
    DiffPanelColumnView,
    DiffPanelRowView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "../../webview/diff-panel/diff-panel-types";
import type {
    DiffWebviewOutgoingMessage,
    DiffWebviewPendingEdit,
} from "../shared/session-protocol";

export type RowFilterMode = "all" | "diffs" | "same";

export interface SelectedDiffCell {
    rowNumber: number;
    columnNumber: number;
}

export interface DraftEditState {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    side: "left" | "right";
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
    rowNumber: number,
    columnNumber: number,
    side: "left" | "right"
): string {
    return `${side}:${rowNumber}:${columnNumber}`;
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
    row: DiffPanelRowView,
    cell: DiffPanelSparseCellView
): {
    selectedCell: SelectedDiffCell;
    activeDiffIndex: number | null;
} {
    return {
        selectedCell: {
            rowNumber: row.rowNumber,
            columnNumber: cell.columnNumber,
        },
        activeDiffIndex: cell.diffIndex,
    };
}

export function beginDiffCellEdit({
    activeSheetKey,
    side,
    sideEditable,
    pendingEdits,
    row,
    cell,
}: {
    activeSheetKey: string | null;
    side: "left" | "right";
    sideEditable: boolean;
    pendingEdits: Record<string, DiffWebviewPendingEdit>;
    row: DiffPanelRowView;
    cell: DiffPanelSparseCellView;
}): {
    editingCell: DraftEditState;
    selectedCell: SelectedDiffCell;
} | null {
    if (!activeSheetKey || !sideEditable) {
        return null;
    }

    if ((side === "left" && !cell.leftPresent) || (side === "right" && !cell.rightPresent)) {
        return null;
    }

    const pendingEdit = pendingEdits[getPendingDiffEditKey(row.rowNumber, cell.columnNumber, side)];

    return {
        editingCell: {
            sheetKey: activeSheetKey,
            rowNumber: row.rowNumber,
            columnNumber: cell.columnNumber,
            side,
            value: pendingEdit?.value ?? (side === "left" ? cell.leftValue : cell.rightValue),
        },
        selectedCell: {
            rowNumber: row.rowNumber,
            columnNumber: cell.columnNumber,
        },
    };
}

export function finalizeDiffCellEdit(
    pendingEdits: Record<string, DiffWebviewPendingEdit>,
    draft: DraftEditState | null,
    disposition: "commit" | "cancel"
): {
    pendingEdits: Record<string, DiffWebviewPendingEdit>;
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

    const edit: DiffWebviewPendingEdit = {
        sheetKey: draft.sheetKey,
        side: draft.side,
        rowNumber: draft.rowNumber,
        columnNumber: draft.columnNumber,
        value: draft.value,
    };

    return {
        pendingEdits: {
            ...pendingEdits,
            [getPendingDiffEditKey(edit.rowNumber, edit.columnNumber, edit.side)]: edit,
        },
        editingCell: null,
    };
}

export function getRenderedDiffCellValue(
    pendingEdits: Record<string, DiffWebviewPendingEdit>,
    row: DiffPanelRowView,
    cell: DiffPanelSparseCellView,
    side: "left" | "right"
): string {
    const pendingEdit = pendingEdits[getPendingDiffEditKey(row.rowNumber, cell.columnNumber, side)];
    if (pendingEdit) {
        return pendingEdit.value;
    }

    return side === "left" ? cell.leftValue : cell.rightValue;
}

export function createSaveDiffEditsMessage(
    pendingEdits: Record<string, DiffWebviewPendingEdit>
): DiffWebviewOutgoingMessage | null {
    const edits = Object.values(pendingEdits);
    if (edits.length === 0) {
        return null;
    }

    return {
        type: "saveEdits",
        edits,
    };
}
