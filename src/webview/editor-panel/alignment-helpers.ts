import { createCellKey } from "../../core/model/cells";
import {
    mergeCellAlignments,
    type CellAlignmentSnapshot,
    type EditorAlignmentPatch,
} from "../../core/model/alignment";
import type { EditorActiveSheetView, EditorSelectedCell } from "../../core/model/types";
import {
    getCellContentAlignmentStyle as getLegacyCellContentAlignmentStyle,
    getToolbarHorizontalAlignment,
    getToolbarVerticalAlignment,
} from "./editor-cell-alignment";
import type { EditorAlignmentTargetKind } from "./editor-panel-types";
import type { SelectionRange } from "./editor-selection-range";

export function getCellContentAlignmentStyle(
    alignment: CellAlignmentSnapshot | null
): Record<string, string> {
    const style = getLegacyCellContentAlignmentStyle(alignment);
    return {
        "justify-content": style.justifyContent,
        "align-items": style.alignItems,
        "text-align": style.textAlign,
        height: style.height,
        "max-height": style.maxHeight,
    };
}

function hasExpandedSelectionRange(range: SelectionRange | null): range is SelectionRange {
    return Boolean(
        range &&
            (range.startRow !== range.endRow || range.startColumn !== range.endColumn)
    );
}

export function getEffectiveEditorCellAlignment(
    activeSheet: EditorActiveSheetView,
    rowNumber: number,
    columnNumber: number
): CellAlignmentSnapshot | null {
    const cellKey = createCellKey(rowNumber, columnNumber);
    return mergeCellAlignments(
        activeSheet.columnAlignments?.[String(columnNumber)] ?? null,
        activeSheet.rowAlignments?.[String(rowNumber)] ?? null,
        activeSheet.cellAlignments?.[cellKey] ?? null
    );
}

export function getActiveEditorToolbarAlignment({
    activeSheet,
    selection,
}: {
    activeSheet: EditorActiveSheetView | null;
    selection: EditorSelectedCell | null;
}): EditorAlignmentPatch {
    if (!activeSheet || !selection) {
        return {};
    }

    const alignment = getEffectiveEditorCellAlignment(
        activeSheet,
        selection.rowNumber,
        selection.columnNumber
    );
    const horizontal = getToolbarHorizontalAlignment(alignment);
    const vertical = getToolbarVerticalAlignment(alignment);

    return {
        ...(horizontal ? { horizontal } : {}),
        ...(vertical ? { vertical } : {}),
    };
}

export function getActiveEditorAlignmentSelectionTarget({
    activeSheet,
    selection,
    selectionRange,
}: {
    activeSheet: EditorActiveSheetView | null;
    selection: EditorSelectedCell | null;
    selectionRange: SelectionRange | null;
}): {
    target: EditorAlignmentTargetKind;
    selection: SelectionRange;
} | null {
    if (!activeSheet || !selection) {
        return null;
    }

    if (!hasExpandedSelectionRange(selectionRange)) {
        return {
            target: "cell",
            selection: {
                startRow: selection.rowNumber,
                endRow: selection.rowNumber,
                startColumn: selection.columnNumber,
                endColumn: selection.columnNumber,
            },
        };
    }

    const selectsAllColumns =
        selectionRange.startColumn === 1 && selectionRange.endColumn === activeSheet.columnCount;
    const selectsAllRows =
        selectionRange.startRow === 1 && selectionRange.endRow === activeSheet.rowCount;

    if (selectsAllColumns !== selectsAllRows) {
        return {
            target: selectsAllColumns ? "row" : "column",
            selection: selectionRange,
        };
    }

    return {
        target: "range",
        selection: selectionRange,
    };
}
