import type { EditorSelectedCell, EditorSelectionView } from "../../core/model/types";
import {
    createColumnSelectionRange,
    createColumnSelectionSpanRange,
    createRowSelectionRange,
    createRowSelectionSpanRange,
    createSelectionRange,
    hasExpandedSelectionRange,
    type SelectionRange,
} from "./editor-selection-range";

export interface EditorSelectionRangeState {
    focusCell: EditorSelectedCell | null;
    anchorCell: EditorSelectedCell | null;
    selectionRange: SelectionRange | null;
}

function normalizeExpandedRange(range: SelectionRange | null): SelectionRange | null {
    return hasExpandedSelectionRange(range) ? range : null;
}

export function resolveEditorSelectionRange(
    selection: Pick<EditorSelectionView, "rowNumber" | "columnNumber"> | null,
    selectionRangeOverride: SelectionRange | null
): SelectionRange | null {
    return selectionRangeOverride ?? createSelectionRange(selection, selection);
}

export function createEditorSingleCellSelectionState(
    focusCell: EditorSelectedCell | null
): EditorSelectionRangeState {
    return {
        focusCell,
        anchorCell: focusCell,
        selectionRange: null,
    };
}

export function createEditorExtendedCellSelectionState({
    anchorCell,
    focusCell,
}: {
    anchorCell: EditorSelectedCell | null;
    focusCell: EditorSelectedCell | null;
}): EditorSelectionRangeState {
    const effectiveAnchor = anchorCell ?? focusCell;
    return {
        focusCell,
        anchorCell: effectiveAnchor,
        selectionRange: normalizeExpandedRange(createSelectionRange(effectiveAnchor, focusCell)),
    };
}

export function createEditorAnchoredRangeSelectionState({
    anchorCell,
    previewCell,
}: {
    anchorCell: EditorSelectedCell | null;
    previewCell: EditorSelectedCell | null;
}): EditorSelectionRangeState {
    const effectiveAnchor = anchorCell ?? previewCell;
    return {
        focusCell: effectiveAnchor,
        anchorCell: effectiveAnchor,
        selectionRange: normalizeExpandedRange(createSelectionRange(effectiveAnchor, previewCell)),
    };
}

export function createEditorRowSelectionState({
    anchorCell,
    focusCell,
    columnCount,
    extend,
}: {
    anchorCell: EditorSelectedCell | null;
    focusCell: EditorSelectedCell;
    columnCount: number;
    extend: boolean;
}): EditorSelectionRangeState {
    const effectiveAnchor = extend ? (anchorCell ?? focusCell) : focusCell;
    return {
        focusCell,
        anchorCell: effectiveAnchor,
        selectionRange: extend
            ? createRowSelectionSpanRange(
                  effectiveAnchor.rowNumber,
                  focusCell.rowNumber,
                  columnCount
              )
            : createRowSelectionRange(focusCell.rowNumber, columnCount),
    };
}

export function createEditorColumnSelectionState({
    anchorCell,
    focusCell,
    rowCount,
    extend,
}: {
    anchorCell: EditorSelectedCell | null;
    focusCell: EditorSelectedCell;
    rowCount: number;
    extend: boolean;
}): EditorSelectionRangeState {
    const effectiveAnchor = extend ? (anchorCell ?? focusCell) : focusCell;
    return {
        focusCell,
        anchorCell: effectiveAnchor,
        selectionRange: extend
            ? createColumnSelectionSpanRange(
                  effectiveAnchor.columnNumber,
                  focusCell.columnNumber,
                  rowCount
              )
            : createColumnSelectionRange(focusCell.columnNumber, rowCount),
    };
}
