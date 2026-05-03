import type { SheetFreezePaneSnapshot } from "../../core/model/types";

export function hasLockedView(
    freezePane: SheetFreezePaneSnapshot | null | undefined
): freezePane is SheetFreezePaneSnapshot {
    return Boolean(freezePane && (freezePane.columnCount > 0 || freezePane.rowCount > 0));
}

export function getFreezePaneCountsForCell(
    cell: { rowNumber: number; columnNumber: number } | null | undefined
): { rowCount: number; columnCount: number } | null {
    if (!cell) {
        return null;
    }

    return {
        rowCount: Math.max(Math.trunc(cell.rowNumber) - 1, 0),
        columnCount: Math.max(Math.trunc(cell.columnNumber) - 1, 0),
    };
}
