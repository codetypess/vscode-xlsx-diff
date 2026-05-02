import { createCellKey, getCellAddress } from "../../core/model/cells";
import type { EditorActiveSheetView, EditorSelectionView } from "../../core/model/types";

export function createOptimisticEditorSelection({
    activeSheet,
    rowNumber,
    columnNumber,
}: {
    activeSheet: EditorActiveSheetView;
    rowNumber: number;
    columnNumber: number;
}): EditorSelectionView | null {
    if (
        !Number.isInteger(rowNumber) ||
        !Number.isInteger(columnNumber) ||
        rowNumber < 1 ||
        columnNumber < 1
    ) {
        return null;
    }

    const key = createCellKey(rowNumber, columnNumber);
    const cell = activeSheet.cells[key];

    return {
        key,
        rowNumber,
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        value: cell?.displayValue ?? "",
        formula: cell?.formula ?? null,
        isPresent: Boolean(cell),
    };
}
