import { createCellKey } from "../../core/model/cells";
import type { EditorActiveSheetView, EditorSelectionView } from "../../core/model/types";
import type { EditorPendingEdit } from "../../webview/editor-panel/editor-panel-types";

export interface EditorCellEditingState {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    modelValue: string;
    draftValue: string;
}

function isCellPositionInBounds(
    rowNumber: number,
    columnNumber: number
): boolean {
    return (
        Number.isInteger(rowNumber) &&
        Number.isInteger(columnNumber) &&
        rowNumber >= 1 &&
        columnNumber >= 1
    );
}

export function getEditorPendingEditValue(
    pendingEdits: readonly EditorPendingEdit[],
    sheetKey: string,
    rowNumber: number,
    columnNumber: number
): string | null {
    return (
        pendingEdits.find(
            (edit) =>
                edit.sheetKey === sheetKey &&
                edit.rowNumber === rowNumber &&
                edit.columnNumber === columnNumber
        )?.value ?? null
    );
}

export function createEditorCellEditingState({
    activeSheet,
    rowNumber,
    columnNumber,
    canEdit,
    pendingEdits,
}: {
    activeSheet: EditorActiveSheetView;
    rowNumber: number;
    columnNumber: number;
    canEdit: boolean;
    pendingEdits: readonly EditorPendingEdit[];
}): EditorCellEditingState | null {
    if (!canEdit || !isCellPositionInBounds(rowNumber, columnNumber)) {
        return null;
    }

    const key = createCellKey(rowNumber, columnNumber);
    const cell = activeSheet.cells[key];
    if (cell?.formula) {
        return null;
    }

    const modelValue = cell?.displayValue ?? "";
    const draftValue =
        getEditorPendingEditValue(pendingEdits, activeSheet.key, rowNumber, columnNumber) ??
        modelValue;

    return {
        sheetKey: activeSheet.key,
        rowNumber,
        columnNumber,
        modelValue,
        draftValue,
    };
}

export function isEditorCellEditingActive({
    editingCell,
    sheetKey,
    rowNumber,
    columnNumber,
}: {
    editingCell: EditorCellEditingState | null;
    sheetKey: string | null;
    rowNumber: number;
    columnNumber: number;
}): boolean {
    return Boolean(
        editingCell &&
        sheetKey === editingCell.sheetKey &&
        editingCell.rowNumber === rowNumber &&
        editingCell.columnNumber === columnNumber
    );
}

export function applyCommittedEditorCellEdit({
    pendingEdits,
    editingCell,
    nextValue,
}: {
    pendingEdits: readonly EditorPendingEdit[];
    editingCell: EditorCellEditingState;
    nextValue: string;
}): EditorPendingEdit[] {
    const remainingEdits = pendingEdits.filter(
        (edit) =>
            !(
                edit.sheetKey === editingCell.sheetKey &&
                edit.rowNumber === editingCell.rowNumber &&
                edit.columnNumber === editingCell.columnNumber
            )
    );
    if (nextValue === editingCell.modelValue) {
        return remainingEdits;
    }

    return [
        ...remainingEdits,
        {
            sheetKey: editingCell.sheetKey,
            rowNumber: editingCell.rowNumber,
            columnNumber: editingCell.columnNumber,
            value: nextValue,
        },
    ];
}

export function getEditorSelectionDisplayValue({
    activeSheetKey,
    selection,
    editingCell,
    pendingEdits,
}: {
    activeSheetKey: string | null;
    selection: EditorSelectionView | null;
    editingCell: EditorCellEditingState | null;
    pendingEdits: readonly EditorPendingEdit[];
}): string {
    if (!selection) {
        return "";
    }

    if (
        editingCell &&
        activeSheetKey === editingCell.sheetKey &&
        editingCell.rowNumber === selection.rowNumber &&
        editingCell.columnNumber === selection.columnNumber
    ) {
        return editingCell.draftValue;
    }

    if (activeSheetKey) {
        const pendingValue = getEditorPendingEditValue(
            pendingEdits,
            activeSheetKey,
            selection.rowNumber,
            selection.columnNumber
        );
        if (pendingValue !== null) {
            return pendingValue;
        }
    }

    return selection.value;
}
