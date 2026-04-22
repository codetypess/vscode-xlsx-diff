import { getCellAddress } from "../core/model/cells";
import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../core/fastxlsx/write-cell-value";
import type { EditorPanelState, WorkbookSnapshot } from "../core/model/types";
import type {
    EditorPendingEdit,
    RestoredStructuralState,
    StructuralSnapshot,
    WorkingSheetEntry,
} from "./editor-panel-types";

export function cloneCellEdit(edit: CellEdit): CellEdit {
    return { ...edit };
}

export function cloneSheetEdit(edit: SheetEdit): SheetEdit {
    return { ...edit };
}

export function cloneViewEdit(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
    };
}

export function cloneSheetSnapshot(
    sheet: WorkbookSnapshot["sheets"][number]
): WorkbookSnapshot["sheets"][number] {
    return {
        ...sheet,
        mergedRanges: [...sheet.mergedRanges],
        freezePane: sheet.freezePane ? { ...sheet.freezePane } : null,
        cells: { ...sheet.cells },
    };
}

export function createFreezePaneSnapshot(
    columnCount: number,
    rowCount: number
): WorkbookSnapshot["sheets"][number]["freezePane"] {
    if (columnCount <= 0 && rowCount <= 0) {
        return null;
    }

    return {
        columnCount,
        rowCount,
        topLeftCell: getCellAddress(rowCount + 1, columnCount + 1),
        activePane:
            rowCount > 0 && columnCount > 0
                ? "bottomRight"
                : rowCount > 0
                  ? "bottomLeft"
                  : "topRight",
    };
}

export function areFreezePaneCountsEqual(
    left: WorkbookSnapshot["sheets"][number]["freezePane"],
    right: WorkbookSnapshot["sheets"][number]["freezePane"]
): boolean {
    return (
        (left?.columnCount ?? 0) === (right?.columnCount ?? 0) &&
        (left?.rowCount ?? 0) === (right?.rowCount ?? 0)
    );
}

export function cloneSheetEntry(entry: WorkingSheetEntry): WorkingSheetEntry {
    return {
        key: entry.key,
        index: entry.index,
        sheet: cloneSheetSnapshot(entry.sheet),
    };
}

export function cloneEditorState(state: EditorPanelState): EditorPanelState {
    return {
        ...state,
        selectedCell: state.selectedCell ? { ...state.selectedCell } : null,
    };
}

export function reindexWorkingSheetEntries(
    sheetEntries: WorkingSheetEntry[]
): WorkingSheetEntry[] {
    return sheetEntries.map((entry, index) => ({
        ...entry,
        index,
    }));
}

export function createWorkingSheetEntries(workbook: WorkbookSnapshot): WorkingSheetEntry[] {
    return workbook.sheets.map((sheet, index) => ({
        key: `sheet:${index}`,
        index,
        sheet: cloneSheetSnapshot(sheet),
    }));
}

export function createWorkingWorkbook(
    workbook: WorkbookSnapshot,
    sheetEntries: WorkingSheetEntry[]
): WorkbookSnapshot {
    return {
        ...workbook,
        sheets: sheetEntries.map((entry) => entry.sheet),
    };
}

export function createPendingWorkbookEditState(
    pendingCellEdits: CellEdit[],
    pendingSheetEdits: SheetEdit[],
    pendingViewEdits: SheetViewEdit[]
): WorkbookEditState {
    return {
        cellEdits: pendingCellEdits.map(cloneCellEdit),
        sheetEdits: pendingSheetEdits.map(cloneSheetEdit),
        viewEdits: pendingViewEdits.map(cloneViewEdit),
    };
}

export function captureStructuralSnapshot(
    state: EditorPanelState,
    sheetEntries: WorkingSheetEntry[],
    pendingCellEdits: CellEdit[],
    pendingSheetEdits: SheetEdit[],
    pendingViewEdits: SheetViewEdit[]
): StructuralSnapshot {
    return {
        state: cloneEditorState(state),
        sheetEntries: sheetEntries.map(cloneSheetEntry),
        pendingCellEdits: pendingCellEdits.map(cloneCellEdit),
        pendingSheetEdits: pendingSheetEdits.map(cloneSheetEdit),
        pendingViewEdits: pendingViewEdits.map(cloneViewEdit),
    };
}

export function restoreStructuralSnapshot(
    snapshot: StructuralSnapshot
): RestoredStructuralState {
    return {
        state: cloneEditorState(snapshot.state),
        sheetEntries: reindexWorkingSheetEntries(snapshot.sheetEntries.map(cloneSheetEntry)),
        pendingCellEdits: snapshot.pendingCellEdits.map(cloneCellEdit),
        pendingSheetEdits: snapshot.pendingSheetEdits.map(cloneSheetEdit),
        pendingViewEdits: snapshot.pendingViewEdits.map(cloneViewEdit),
    };
}

export function mapPendingCellEditsToWebview(
    pendingCellEdits: CellEdit[],
    sheetEntries: WorkingSheetEntry[]
): EditorPendingEdit[] {
    return pendingCellEdits.map((edit) => ({
        sheetKey:
            sheetEntries.find((entry) => entry.sheet.name === edit.sheetName)?.key ?? edit.sheetName,
        rowNumber: edit.rowNumber,
        columnNumber: edit.columnNumber,
        value: edit.value,
    }));
}
