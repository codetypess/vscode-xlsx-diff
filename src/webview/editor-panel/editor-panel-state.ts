import { createCellKey, getCellAddress } from "../../core/model/cells";
import { DEFAULT_PAGE_SIZE } from "../../constants";
import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../../core/fastxlsx/write-cell-value";
import type { EditorPanelState, WorkbookSnapshot } from "../../core/model/types";
import type {
    EditorPendingEdit,
    RestoredStructuralState,
    StructuralSnapshot,
    WorkingSheetEntry,
} from "./editor-panel-types";

export type GridSheetEdit = Extract<
    SheetEdit,
    { type: "insertRow" | "deleteRow" | "insertColumn" | "deleteColumn" }
>;

export function cloneCellEdit(edit: CellEdit): CellEdit {
    return { ...edit };
}

export function cloneSheetEdit(edit: SheetEdit): SheetEdit {
    return { ...edit };
}

export function cloneColumnWidths(
    columnWidths: readonly (number | null | undefined)[] | undefined
): Array<number | null> {
    return (columnWidths ?? []).map((columnWidth) => columnWidth ?? null);
}

function normalizeColumnWidthsForComparison(
    columnWidths: readonly (number | null | undefined)[] | undefined
): Array<number | null> {
    const normalizedWidths = cloneColumnWidths(columnWidths);
    while (
        normalizedWidths.length > 0 &&
        normalizedWidths[normalizedWidths.length - 1] === null
    ) {
        normalizedWidths.pop();
    }

    return normalizedWidths;
}

export function areColumnWidthsEquivalent(
    left: readonly (number | null | undefined)[] | undefined,
    right: readonly (number | null | undefined)[] | undefined
): boolean {
    const normalizedLeft = normalizeColumnWidthsForComparison(left);
    const normalizedRight = normalizeColumnWidthsForComparison(right);
    if (normalizedLeft.length !== normalizedRight.length) {
        return false;
    }

    return normalizedLeft.every((columnWidth, index) => columnWidth === normalizedRight[index]);
}

export function setSheetColumnWidthSnapshot(
    columnWidths: readonly (number | null | undefined)[] | undefined,
    columnNumber: number,
    nextWidth: number | null
): Array<number | null> {
    const nextColumnWidths = cloneColumnWidths(columnWidths);
    while (nextColumnWidths.length < columnNumber) {
        nextColumnWidths.push(null);
    }

    nextColumnWidths[columnNumber - 1] = nextWidth;
    return nextColumnWidths;
}

export function cloneViewEdit(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
        ...(edit.columnWidths
            ? {
                  columnWidths: cloneColumnWidths(edit.columnWidths),
              }
            : {}),
    };
}

export function cloneSheetSnapshot(
    sheet: WorkbookSnapshot["sheets"][number]
): WorkbookSnapshot["sheets"][number] {
    return {
        ...sheet,
        mergedRanges: [...sheet.mergedRanges],
        columnWidths: cloneColumnWidths(sheet.columnWidths),
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

export function reindexWorkingSheetEntries(sheetEntries: WorkingSheetEntry[]): WorkingSheetEntry[] {
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

function expandSheetBoundsForCellEdit(
    sheet: WorkbookSnapshot["sheets"][number],
    edit: CellEdit
): WorkbookSnapshot["sheets"][number] {
    const nextRowCount = Math.max(sheet.rowCount, edit.rowNumber);
    const nextColumnCount = Math.max(sheet.columnCount, edit.columnNumber);
    if (nextRowCount === sheet.rowCount && nextColumnCount === sheet.columnCount) {
        return sheet;
    }

    return {
        ...sheet,
        rowCount: nextRowCount,
        columnCount: nextColumnCount,
        signature: createPendingSheetSignature(
            sheet.name,
            nextRowCount,
            nextColumnCount,
            Object.keys(sheet.cells).length
        ),
    };
}

export function createWorkingWorkbook(
    workbook: WorkbookSnapshot,
    sheetEntries: WorkingSheetEntry[],
    pendingCellEdits: CellEdit[] = []
): WorkbookSnapshot {
    const sheetsByName = new Map(
        sheetEntries.map((entry) => [entry.sheet.name, entry.sheet] as const)
    );

    for (const edit of pendingCellEdits) {
        const sheet = sheetsByName.get(edit.sheetName);
        if (!sheet) {
            continue;
        }

        sheetsByName.set(edit.sheetName, expandSheetBoundsForCellEdit(sheet, edit));
    }

    return {
        ...workbook,
        sheets: sheetEntries.map((entry) => sheetsByName.get(entry.sheet.name) ?? entry.sheet),
    };
}

export function isGridSheetEdit(edit: SheetEdit): edit is GridSheetEdit {
    return (
        edit.type === "insertRow" ||
        edit.type === "deleteRow" ||
        edit.type === "insertColumn" ||
        edit.type === "deleteColumn"
    );
}

function createPendingSheetSignature(
    sheetName: string,
    rowCount: number,
    columnCount: number,
    cellCount: number
): string {
    return `pending:${sheetName}:${rowCount}:${columnCount}:${cellCount}`;
}

function transformCellPosition(
    rowNumber: number,
    columnNumber: number,
    edit: GridSheetEdit
): { rowNumber: number; columnNumber: number } | null {
    if (edit.type === "insertRow") {
        return {
            rowNumber: rowNumber >= edit.rowNumber ? rowNumber + edit.count : rowNumber,
            columnNumber,
        };
    }

    if (edit.type === "deleteRow") {
        const lastDeletedRow = edit.rowNumber + edit.count - 1;
        if (rowNumber >= edit.rowNumber && rowNumber <= lastDeletedRow) {
            return null;
        }

        return {
            rowNumber: rowNumber > lastDeletedRow ? rowNumber - edit.count : rowNumber,
            columnNumber,
        };
    }

    if (edit.type === "insertColumn") {
        return {
            rowNumber,
            columnNumber:
                columnNumber >= edit.columnNumber ? columnNumber + edit.count : columnNumber,
        };
    }

    const lastDeletedColumn = edit.columnNumber + edit.count - 1;
    if (columnNumber >= edit.columnNumber && columnNumber <= lastDeletedColumn) {
        return null;
    }

    return {
        rowNumber,
        columnNumber: columnNumber > lastDeletedColumn ? columnNumber - edit.count : columnNumber,
    };
}

function transformColumnWidths(
    columnWidths: WorkbookSnapshot["sheets"][number]["columnWidths"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["columnWidths"] {
    if (edit.type !== "insertColumn" && edit.type !== "deleteColumn") {
        return cloneColumnWidths(columnWidths);
    }

    const nextWidths = cloneColumnWidths(columnWidths);
    if (edit.type === "insertColumn") {
        nextWidths.splice(
            edit.columnNumber - 1,
            0,
            ...Array.from({ length: edit.count }, () => null)
        );
        return nextWidths;
    }

    nextWidths.splice(edit.columnNumber - 1, edit.count);
    return nextWidths;
}

export function shiftPendingCellEditsForGridSheetEdit(
    pendingCellEdits: CellEdit[],
    edit: GridSheetEdit
): CellEdit[] {
    return pendingCellEdits.flatMap((pendingEdit) => {
        if (pendingEdit.sheetName !== edit.sheetName) {
            return [pendingEdit];
        }

        const nextPosition = transformCellPosition(
            pendingEdit.rowNumber,
            pendingEdit.columnNumber,
            edit
        );
        if (!nextPosition) {
            return [];
        }

        return [
            {
                ...pendingEdit,
                rowNumber: nextPosition.rowNumber,
                columnNumber: nextPosition.columnNumber,
            },
        ];
    });
}

export function applyGridSheetEditToSheet(
    sheet: WorkbookSnapshot["sheets"][number],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number] {
    if (sheet.name !== edit.sheetName) {
        return sheet;
    }

    const nextCount = Math.max(0, edit.count);
    if (nextCount === 0) {
        return sheet;
    }

    const nextCells = Object.fromEntries(
        Object.values(sheet.cells)
            .flatMap((cell) => {
                const nextPosition = transformCellPosition(cell.rowNumber, cell.columnNumber, edit);
                if (!nextPosition) {
                    return [];
                }

                const nextCell = {
                    ...cell,
                    key: createCellKey(nextPosition.rowNumber, nextPosition.columnNumber),
                    rowNumber: nextPosition.rowNumber,
                    columnNumber: nextPosition.columnNumber,
                    address: getCellAddress(nextPosition.rowNumber, nextPosition.columnNumber),
                };
                return [[nextCell.key, nextCell] as const];
            })
            .sort((left, right) => left[0].localeCompare(right[0]))
    );

    const rowCount =
        edit.type === "insertRow"
            ? sheet.rowCount + nextCount
            : edit.type === "deleteRow"
              ? Math.max(0, sheet.rowCount - nextCount)
              : sheet.rowCount;
    const columnCount =
        edit.type === "insertColumn"
            ? sheet.columnCount + nextCount
            : edit.type === "deleteColumn"
              ? Math.max(0, sheet.columnCount - nextCount)
              : sheet.columnCount;

    return {
        ...sheet,
        rowCount,
        columnCount,
        columnWidths: transformColumnWidths(sheet.columnWidths, edit),
        cells: nextCells,
        signature: createPendingSheetSignature(
            sheet.name,
            rowCount,
            columnCount,
            Object.keys(nextCells).length
        ),
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
        const nextRowCount = Math.max(entry.sheet.rowCount, edit.rowNumber);
        const nextColumnCount = Math.max(entry.sheet.columnCount, edit.columnNumber);
        entry.sheet = {
            ...entry.sheet,
            rowCount: nextRowCount,
            columnCount: nextColumnCount,
            cells: {
                ...entry.sheet.cells,
                [key]: {
                    key,
                    rowNumber: edit.rowNumber,
                    columnNumber: edit.columnNumber,
                    address:
                        currentCell?.address ?? getCellAddress(edit.rowNumber, edit.columnNumber),
                    displayValue: edit.value,
                    formula: null,
                    styleId: currentCell?.styleId ?? null,
                },
            },
            signature: createPendingSheetSignature(
                entry.sheet.name,
                nextRowCount,
                nextColumnCount,
                Object.keys(entry.sheet.cells).length + (currentCell ? 0 : 1)
            ),
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
                    visibility: "visible",
                    mergedRanges: [],
                    columnWidths: [],
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
                    (entry) => entry.key !== edit.sheetKey && entry.sheet.name !== edit.sheetName
                )
            );
            continue;
        }

        if (edit.type === "renameSheet") {
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
    }

    for (const edit of pendingSheetEdits) {
        if (!isGridSheetEdit(edit)) {
            continue;
        }

        sheetEntries = sheetEntries.map((entry) =>
            entry.key === edit.sheetKey
                ? {
                      ...entry,
                      sheet: applyGridSheetEditToSheet(entry.sheet, edit),
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
                          columnWidths: edit.columnWidths
                              ? cloneColumnWidths(edit.columnWidths)
                              : cloneColumnWidths(entry.sheet.columnWidths),
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

export function restoreStructuralSnapshot(snapshot: StructuralSnapshot): RestoredStructuralState {
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
            sheetEntries.find((entry) => entry.sheet.name === edit.sheetName)?.key ??
            edit.sheetName,
        rowNumber: edit.rowNumber,
        columnNumber: edit.columnNumber,
        value: edit.value,
    }));
}
