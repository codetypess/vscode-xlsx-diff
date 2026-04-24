import { createCellKey, getCellAddress } from "../core/model/cells";
import { DEFAULT_PAGE_SIZE } from "../constants";
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

export function createCommittedWorkbookState(
    workbook: WorkbookSnapshot,
    sheetEntries: WorkingSheetEntry[],
    pendingCellEdits: CellEdit[]
): {
    workbook: WorkbookSnapshot;
    sheetEntries: WorkingSheetEntry[];
} {
    const sheetEntryIndexes = new Map(
        sheetEntries.map((entry, index) => [entry.sheet.name, index] as const)
    );
    const clonedSheetIndexes = new Set<number>();
    let committedEntries = sheetEntries;

    for (const edit of pendingCellEdits) {
        const entryIndex = sheetEntryIndexes.get(edit.sheetName);
        if (entryIndex === undefined) {
            continue;
        }

        if (committedEntries === sheetEntries) {
            committedEntries = [...sheetEntries];
        }

        if (!clonedSheetIndexes.has(entryIndex)) {
            const sourceEntry = committedEntries[entryIndex]!;
            committedEntries[entryIndex] = {
                ...sourceEntry,
                sheet: {
                    ...sourceEntry.sheet,
                    cells: { ...sourceEntry.sheet.cells },
                },
            };
            clonedSheetIndexes.add(entryIndex);
        }

        const entry = committedEntries[entryIndex]!;
        const key = createCellKey(edit.rowNumber, edit.columnNumber);
        const currentCell = entry.sheet.cells[key];
        entry.sheet = {
            ...entry.sheet,
            cells: {
                ...entry.sheet.cells,
                [key]: {
                    key,
                    rowNumber: edit.rowNumber,
                    columnNumber: edit.columnNumber,
                    address: currentCell?.address ?? getCellAddress(edit.rowNumber, edit.columnNumber),
                    displayValue: edit.value,
                    formula: null,
                    styleId: currentCell?.styleId ?? null,
                },
            },
        };
    }

    const committedWorkbook: WorkbookSnapshot = {
        ...workbook,
        sheets: committedEntries.map((entry) => ({ ...entry.sheet })),
    };

    return {
        workbook: committedWorkbook,
        sheetEntries: committedEntries,
    };
}

export function restorePendingWorkbookState(
    workbook: WorkbookSnapshot,
    pendingState: Readonly<WorkbookEditState>
): {
    sheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
    nextNewSheetId: number;
} {
    let sheetEntries = createWorkingSheetEntries(workbook);
    const pendingCellEdits = pendingState.cellEdits.map(cloneCellEdit);
    const pendingSheetEdits = pendingState.sheetEdits.map(cloneSheetEdit);
    const pendingViewEdits = (pendingState.viewEdits ?? []).map(cloneViewEdit);

    for (const edit of pendingSheetEdits) {
        if (edit.type === "addSheet") {
            const newEntry: WorkingSheetEntry = {
                key: edit.sheetKey,
                index: edit.targetIndex,
                sheet: {
                    name: edit.sheetName,
                    rowCount: DEFAULT_PAGE_SIZE,
                    columnCount: 26,
                    mergedRanges: [],
                    cells: {},
                    freezePane: null,
                    signature: `pending:${edit.sheetKey}:${edit.sheetName}`,
                },
            };
            sheetEntries = reindexWorkingSheetEntries([
                ...sheetEntries.slice(0, edit.targetIndex),
                newEntry,
                ...sheetEntries.slice(edit.targetIndex),
            ]);
            continue;
        }

        if (edit.type === "deleteSheet") {
            sheetEntries = reindexWorkingSheetEntries(
                sheetEntries.filter(
                    (entry) =>
                        entry.key !== edit.sheetKey && entry.sheet.name !== edit.sheetName
                )
            );
            continue;
        }

        sheetEntries = sheetEntries.map((entry) =>
            entry.key === edit.sheetKey
                ? {
                      ...entry,
                      sheet: {
                          ...entry.sheet,
                          name: edit.nextSheetName,
                          signature: `pending:${edit.sheetKey}:${edit.nextSheetName}`,
                      },
                  }
                : entry
        );
    }

    for (const edit of pendingViewEdits) {
        sheetEntries = sheetEntries.map((entry) =>
            entry.key === edit.sheetKey
                ? {
                      ...entry,
                      sheet: {
                          ...entry.sheet,
                          freezePane: edit.freezePane
                              ? createFreezePaneSnapshot(
                                    edit.freezePane.columnCount,
                                    edit.freezePane.rowCount
                                )
                              : null,
                      },
                  }
                : entry
        );
    }

    const nextNewSheetId =
        sheetEntries.reduce((maxId, entry) => {
            const match = /^sheet:new:(\d+)$/.exec(entry.key);
            return match ? Math.max(maxId, Number(match[1])) : maxId;
        }, 0) + 1;

    return {
        sheetEntries,
        pendingCellEdits,
        pendingSheetEdits,
        pendingViewEdits,
        nextNewSheetId,
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
