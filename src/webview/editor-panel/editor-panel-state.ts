import { createCellKey, getCellAddress } from "../../core/model/cells";
import {
    applyCellAlignmentPatch,
    areCellAlignmentsEqual,
    areCellAlignmentMapsEquivalent,
    cloneCellAlignmentMap,
    type CellAlignmentSnapshot,
    type EditorAlignmentPatch,
} from "../../core/model/alignment";
import { DEFAULT_PAGE_SIZE } from "../../constants";
import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../../core/fastxlsx/write-cell-value";
import type { EditorPanelState, WorkbookSnapshot } from "../../core/model/types";
import type {
    EditorAlignmentTargetKind,
    EditorPendingEdit,
    RestoredStructuralState,
    StructuralSnapshot,
    WorkingSheetEntry,
} from "./editor-panel-types";
import type { SelectionRange } from "./editor-selection-range";

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
    const nextColumnWidths = (columnWidths ?? []).map((columnWidth) => columnWidth ?? null);
    while (
        nextColumnWidths.length > 0 &&
        nextColumnWidths[nextColumnWidths.length - 1] === null
    ) {
        nextColumnWidths.pop();
    }

    return nextColumnWidths;
}

export function cloneRowHeights(
    rowHeights: Readonly<Record<string, number | null>> | undefined
): Record<string, number | null> {
    return { ...(rowHeights ?? {}) };
}

export function cloneCellAlignments(
    cellAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): Record<string, CellAlignmentSnapshot> {
    return cloneCellAlignmentMap(cellAlignments);
}

export function cloneRowAlignments(
    rowAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): Record<string, CellAlignmentSnapshot> {
    return cloneCellAlignmentMap(rowAlignments);
}

export function cloneColumnAlignments(
    columnAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): Record<string, CellAlignmentSnapshot> {
    return cloneCellAlignmentMap(columnAlignments);
}

function copyAlignmentEntries(
    alignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): Record<string, CellAlignmentSnapshot> {
    return { ...(alignments ?? {}) };
}

function normalizeColumnWidthsForComparison(
    columnWidths: readonly (number | null | undefined)[] | undefined
): Array<number | null> {
    const normalizedWidths = cloneColumnWidths(columnWidths);
    while (normalizedWidths.length > 0 && normalizedWidths[normalizedWidths.length - 1] === null) {
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

function normalizeRowHeightsForComparison(
    rowHeights: Readonly<Record<string, number | null>> | undefined
): Record<string, number> {
    return Object.fromEntries(
        Object.entries(rowHeights ?? {})
            .filter(([, rowHeight]) => rowHeight !== null)
            .sort(
                ([leftRowNumber], [rightRowNumber]) =>
                    Number(leftRowNumber) - Number(rightRowNumber)
            )
    ) as Record<string, number>;
}

export function areRowHeightsEquivalent(
    left: Readonly<Record<string, number | null>> | undefined,
    right: Readonly<Record<string, number | null>> | undefined
): boolean {
    const normalizedLeft = normalizeRowHeightsForComparison(left);
    const normalizedRight = normalizeRowHeightsForComparison(right);
    const leftKeys = Object.keys(normalizedLeft);
    const rightKeys = Object.keys(normalizedRight);
    if (leftKeys.length !== rightKeys.length) {
        return false;
    }

    return leftKeys.every((rowNumber) => normalizedLeft[rowNumber] === normalizedRight[rowNumber]);
}

export function areCellAlignmentsEquivalent(
    left: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    right: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): boolean {
    return areCellAlignmentMapsEquivalent(left, right);
}

export function areRowAlignmentsEquivalent(
    left: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    right: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): boolean {
    return areCellAlignmentMapsEquivalent(left, right);
}

export function areColumnAlignmentsEquivalent(
    left: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    right: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): boolean {
    return areCellAlignmentMapsEquivalent(left, right);
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
    return cloneColumnWidths(nextColumnWidths);
}

export function setSheetRowHeightSnapshot(
    rowHeights: Readonly<Record<string, number | null>> | undefined,
    rowNumber: number,
    nextHeight: number | null
): Record<string, number | null> {
    const nextRowHeights = cloneRowHeights(rowHeights);
    if (nextHeight === null) {
        delete nextRowHeights[String(rowNumber)];
        return nextRowHeights;
    }

    nextRowHeights[String(rowNumber)] = nextHeight;
    return nextRowHeights;
}

function setAlignmentSnapshotEntry(
    alignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    key: string,
    patch: EditorAlignmentPatch
): Record<string, CellAlignmentSnapshot> {
    const nextAlignments = cloneCellAlignmentMap(alignments);
    const nextAlignment = applyCellAlignmentPatch(nextAlignments[key], patch);
    if (!nextAlignment) {
        delete nextAlignments[key];
        return nextAlignments;
    }

    nextAlignments[key] = nextAlignment;
    return nextAlignments;
}

export function setSheetCellAlignmentSnapshot(
    cellAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    rowNumber: number,
    columnNumber: number,
    patch: EditorAlignmentPatch
): Record<string, CellAlignmentSnapshot> {
    return setAlignmentSnapshotEntry(cellAlignments, createCellKey(rowNumber, columnNumber), patch);
}

export function setSheetRowAlignmentSnapshot(
    rowAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    rowNumber: number,
    patch: EditorAlignmentPatch
): Record<string, CellAlignmentSnapshot> {
    return setAlignmentSnapshotEntry(rowAlignments, String(rowNumber), patch);
}

export function setSheetColumnAlignmentSnapshot(
    columnAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    columnNumber: number,
    patch: EditorAlignmentPatch
): Record<string, CellAlignmentSnapshot> {
    return setAlignmentSnapshotEntry(columnAlignments, String(columnNumber), patch);
}

export function applyAlignmentPatchToSheetSnapshot(
    sheet: WorkbookSnapshot["sheets"][number],
    target: EditorAlignmentTargetKind,
    selection: SelectionRange,
    patch: EditorAlignmentPatch
): WorkbookSnapshot["sheets"][number] {
    if (target === "cell" || target === "range") {
        let nextCellAlignments: Record<string, CellAlignmentSnapshot> | null = null;

        for (let rowNumber = selection.startRow; rowNumber <= selection.endRow; rowNumber += 1) {
            for (
                let columnNumber = selection.startColumn;
                columnNumber <= selection.endColumn;
                columnNumber += 1
            ) {
                const cellKey = createCellKey(rowNumber, columnNumber);
                const currentAlignment = sheet.cellAlignments?.[cellKey];
                const nextAlignment = applyCellAlignmentPatch(currentAlignment, patch);
                if (areCellAlignmentsEqual(currentAlignment, nextAlignment)) {
                    continue;
                }

                nextCellAlignments ??= copyAlignmentEntries(sheet.cellAlignments);
                if (nextAlignment) {
                    nextCellAlignments[cellKey] = nextAlignment;
                } else {
                    delete nextCellAlignments[cellKey];
                }
            }
        }

        return nextCellAlignments
            ? {
                  ...sheet,
                  cellAlignments: nextCellAlignments,
              }
            : sheet;
    }

    if (target === "row") {
        let nextRowAlignments: Record<string, CellAlignmentSnapshot> | null = null;

        for (let rowNumber = selection.startRow; rowNumber <= selection.endRow; rowNumber += 1) {
            const rowKey = String(rowNumber);
            const currentAlignment = sheet.rowAlignments?.[rowKey];
            const nextAlignment = applyCellAlignmentPatch(currentAlignment, patch);
            if (areCellAlignmentsEqual(currentAlignment, nextAlignment)) {
                continue;
            }

            nextRowAlignments ??= copyAlignmentEntries(sheet.rowAlignments);
            if (nextAlignment) {
                nextRowAlignments[rowKey] = nextAlignment;
            } else {
                delete nextRowAlignments[rowKey];
            }
        }

        return nextRowAlignments
            ? {
                  ...sheet,
                  rowAlignments: nextRowAlignments,
              }
            : sheet;
    }

    let nextColumnAlignments: Record<string, CellAlignmentSnapshot> | null = null;

    for (
        let columnNumber = selection.startColumn;
        columnNumber <= selection.endColumn;
        columnNumber += 1
    ) {
        const columnKey = String(columnNumber);
        const currentAlignment = sheet.columnAlignments?.[columnKey];
        const nextAlignment = applyCellAlignmentPatch(currentAlignment, patch);
        if (areCellAlignmentsEqual(currentAlignment, nextAlignment)) {
            continue;
        }

        nextColumnAlignments ??= copyAlignmentEntries(sheet.columnAlignments);
        if (nextAlignment) {
            nextColumnAlignments[columnKey] = nextAlignment;
        } else {
            delete nextColumnAlignments[columnKey];
        }
    }

    return nextColumnAlignments
        ? {
              ...sheet,
              columnAlignments: nextColumnAlignments,
          }
        : sheet;
}

export function cloneViewEdit(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
        ...(edit.dirtyCellAlignmentKeys
            ? {
                  dirtyCellAlignmentKeys: [...edit.dirtyCellAlignmentKeys],
              }
            : {}),
        ...(edit.dirtyRowAlignmentKeys
            ? {
                  dirtyRowAlignmentKeys: [...edit.dirtyRowAlignmentKeys],
              }
            : {}),
        ...(edit.dirtyColumnAlignmentKeys
            ? {
                  dirtyColumnAlignmentKeys: [...edit.dirtyColumnAlignmentKeys],
              }
            : {}),
        ...(edit.autoFilter !== undefined
            ? {
                  autoFilter: cloneAutoFilterSnapshot(edit.autoFilter),
              }
            : {}),
        ...(edit.columnWidths
            ? {
                  columnWidths: cloneColumnWidths(edit.columnWidths),
              }
            : {}),
        ...(edit.rowHeights
            ? {
                  rowHeights: cloneRowHeights(edit.rowHeights),
              }
            : {}),
        ...(edit.cellAlignments
            ? {
                  cellAlignments: cloneCellAlignments(edit.cellAlignments),
              }
            : {}),
        ...(edit.rowAlignments
            ? {
                  rowAlignments: cloneRowAlignments(edit.rowAlignments),
              }
            : {}),
        ...(edit.columnAlignments
            ? {
                  columnAlignments: cloneColumnAlignments(edit.columnAlignments),
              }
            : {}),
    };
}

function cloneViewEditForStructuralSnapshot(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
        ...(edit.dirtyCellAlignmentKeys
            ? {
                  dirtyCellAlignmentKeys: [...edit.dirtyCellAlignmentKeys],
              }
            : {}),
        ...(edit.dirtyRowAlignmentKeys
            ? {
                  dirtyRowAlignmentKeys: [...edit.dirtyRowAlignmentKeys],
              }
            : {}),
        ...(edit.dirtyColumnAlignmentKeys
            ? {
                  dirtyColumnAlignmentKeys: [...edit.dirtyColumnAlignmentKeys],
              }
            : {}),
        ...(edit.autoFilter !== undefined
            ? {
                  autoFilter: cloneAutoFilterSnapshot(edit.autoFilter),
              }
            : {}),
        ...(edit.columnWidths
            ? {
                  columnWidths: edit.columnWidths,
              }
            : {}),
        ...(edit.rowHeights
            ? {
                  rowHeights: edit.rowHeights,
              }
            : {}),
        ...(edit.cellAlignments
            ? {
                  cellAlignments: edit.cellAlignments,
              }
            : {}),
        ...(edit.rowAlignments
            ? {
                  rowAlignments: edit.rowAlignments,
              }
            : {}),
        ...(edit.columnAlignments
            ? {
                  columnAlignments: edit.columnAlignments,
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
        rowHeights: cloneRowHeights(sheet.rowHeights),
        cellAlignments: cloneCellAlignments(sheet.cellAlignments),
        rowAlignments: cloneRowAlignments(sheet.rowAlignments),
        columnAlignments: cloneColumnAlignments(sheet.columnAlignments),
        freezePane: sheet.freezePane ? { ...sheet.freezePane } : null,
        autoFilter: cloneAutoFilterSnapshot(sheet.autoFilter),
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

export function cloneAutoFilterSnapshot(
    autoFilter: WorkbookSnapshot["sheets"][number]["autoFilter"] | undefined
): WorkbookSnapshot["sheets"][number]["autoFilter"] {
    if (autoFilter === undefined) {
        return undefined;
    }

    if (autoFilter === null) {
        return null;
    }

    return {
        range: {
            ...autoFilter.range,
        },
        sort: autoFilter.sort
            ? {
                  ...autoFilter.sort,
              }
            : null,
    };
}

export function areAutoFiltersEquivalent(
    left: WorkbookSnapshot["sheets"][number]["autoFilter"] | undefined,
    right: WorkbookSnapshot["sheets"][number]["autoFilter"] | undefined
): boolean {
    if (left === right) {
        return true;
    }

    if (!left || !right) {
        return left === right;
    }

    return (
        left.range.startRow === right.range.startRow &&
        left.range.endRow === right.range.endRow &&
        left.range.startColumn === right.range.startColumn &&
        left.range.endColumn === right.range.endColumn &&
        (left.sort?.columnNumber ?? null) === (right.sort?.columnNumber ?? null) &&
        (left.sort?.direction ?? null) === (right.sort?.direction ?? null)
    );
}

export function cloneSheetEntry(entry: WorkingSheetEntry): WorkingSheetEntry {
    return {
        key: entry.key,
        index: entry.index,
        sheet: cloneSheetSnapshot(entry.sheet),
    };
}

function cloneSheetSnapshotForStructuralSnapshot(
    sheet: WorkbookSnapshot["sheets"][number]
): WorkbookSnapshot["sheets"][number] {
    return {
        ...sheet,
        mergedRanges: [...sheet.mergedRanges],
        columnWidths: sheet.columnWidths,
        rowHeights: sheet.rowHeights,
        cellAlignments: sheet.cellAlignments,
        rowAlignments: sheet.rowAlignments,
        columnAlignments: sheet.columnAlignments,
        freezePane: sheet.freezePane ? { ...sheet.freezePane } : null,
        autoFilter: cloneAutoFilterSnapshot(sheet.autoFilter),
        cells: sheet.cells,
    };
}

function cloneSheetEntryForStructuralSnapshot(entry: WorkingSheetEntry): WorkingSheetEntry {
    return {
        key: entry.key,
        index: entry.index,
        sheet: cloneSheetSnapshotForStructuralSnapshot(entry.sheet),
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
        return cloneColumnWidths(nextWidths);
    }

    nextWidths.splice(edit.columnNumber - 1, edit.count);
    return cloneColumnWidths(nextWidths);
}

function transformRowHeights(
    rowHeights: WorkbookSnapshot["sheets"][number]["rowHeights"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["rowHeights"] {
    if (edit.type !== "insertRow" && edit.type !== "deleteRow") {
        return cloneRowHeights(rowHeights);
    }

    const nextRowHeights = Object.fromEntries(
        Object.entries(rowHeights ?? {}).flatMap(([rowNumberText, rowHeight]) => {
            const rowNumber = Number(rowNumberText);
            if (!Number.isInteger(rowNumber) || rowNumber < 1) {
                return [];
            }

            if (edit.type === "insertRow") {
                return [
                    [
                        String(rowNumber >= edit.rowNumber ? rowNumber + edit.count : rowNumber),
                        rowHeight,
                    ] as const,
                ];
            }

            const lastDeletedRow = edit.rowNumber + edit.count - 1;
            if (rowNumber >= edit.rowNumber && rowNumber <= lastDeletedRow) {
                return [];
            }

            return [
                [
                    String(rowNumber > lastDeletedRow ? rowNumber - edit.count : rowNumber),
                    rowHeight,
                ] as const,
            ];
        })
    );

    return nextRowHeights;
}

function transformCellAlignments(
    cellAlignments: WorkbookSnapshot["sheets"][number]["cellAlignments"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["cellAlignments"] {
    return Object.fromEntries(
        Object.entries(cellAlignments ?? {})
            .flatMap(([cellKey, alignment]) => {
                const [rowNumberText, columnNumberText] = cellKey.split(":");
                const rowNumber = Number(rowNumberText);
                const columnNumber = Number(columnNumberText);
                if (!Number.isInteger(rowNumber) || !Number.isInteger(columnNumber)) {
                    return [];
                }

                const nextPosition = transformCellPosition(rowNumber, columnNumber, edit);
                if (!nextPosition) {
                    return [];
                }

                return [
                    [
                        createCellKey(nextPosition.rowNumber, nextPosition.columnNumber),
                        alignment,
                    ] as const,
                ];
            })
            .sort(([leftCellKey], [rightCellKey]) => leftCellKey.localeCompare(rightCellKey))
    );
}

function transformIndexedAlignments(
    alignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    edit: GridSheetEdit,
    axis: "row" | "column"
): Record<string, CellAlignmentSnapshot> {
    const isRowAxis = axis === "row";
    if (
        (isRowAxis && edit.type !== "insertRow" && edit.type !== "deleteRow") ||
        (!isRowAxis && edit.type !== "insertColumn" && edit.type !== "deleteColumn")
    ) {
        return cloneCellAlignmentMap(alignments);
    }

    return Object.fromEntries(
        Object.entries(alignments ?? {})
            .flatMap(([indexText, alignment]) => {
                const index = Number(indexText);
                if (!Number.isInteger(index) || index < 1) {
                    return [];
                }

                if (isRowAxis) {
                    const nextPosition = transformCellPosition(index, 1, edit);
                    return nextPosition
                        ? [[String(nextPosition.rowNumber), alignment] as const]
                        : [];
                }

                const nextPosition = transformCellPosition(1, index, edit);
                return nextPosition
                    ? [[String(nextPosition.columnNumber), alignment] as const]
                    : [];
            })
            .sort(
                ([leftIndexText], [rightIndexText]) =>
                    Number(leftIndexText) - Number(rightIndexText)
            )
    );
}

function transformRowAlignments(
    rowAlignments: WorkbookSnapshot["sheets"][number]["rowAlignments"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["rowAlignments"] {
    return transformIndexedAlignments(rowAlignments, edit, "row");
}

function transformColumnAlignments(
    columnAlignments: WorkbookSnapshot["sheets"][number]["columnAlignments"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["columnAlignments"] {
    return transformIndexedAlignments(columnAlignments, edit, "column");
}

function transformRangeAxis(
    start: number,
    end: number,
    editStart: number,
    count: number,
    kind: "insert" | "delete"
): { start: number; end: number } | null {
    if (kind === "insert") {
        if (editStart <= start) {
            return {
                start: start + count,
                end: end + count,
            };
        }

        if (editStart <= end) {
            return {
                start,
                end: end + count,
            };
        }

        return { start, end };
    }

    const lastDeletedIndex = editStart + count - 1;
    if (lastDeletedIndex < start) {
        return {
            start: start - count,
            end: end - count,
        };
    }

    if (editStart > end) {
        return { start, end };
    }

    const deletedBeforeStart = Math.max(0, Math.min(lastDeletedIndex, start - 1) - editStart + 1);
    const deletedWithinRange = Math.max(
        0,
        Math.min(lastDeletedIndex, end) - Math.max(editStart, start) + 1
    );
    const nextStart = start - deletedBeforeStart;
    const nextEnd = end - deletedBeforeStart - deletedWithinRange;

    return nextStart <= nextEnd ? { start: nextStart, end: nextEnd } : null;
}

function transformAutoFilter(
    autoFilter: WorkbookSnapshot["sheets"][number]["autoFilter"],
    edit: GridSheetEdit
): WorkbookSnapshot["sheets"][number]["autoFilter"] {
    if (!autoFilter) {
        return null;
    }

    const nextRowRange =
        edit.type === "insertRow"
            ? transformRangeAxis(
                  autoFilter.range.startRow,
                  autoFilter.range.endRow,
                  edit.rowNumber,
                  edit.count,
                  "insert"
              )
            : edit.type === "deleteRow"
              ? transformRangeAxis(
                    autoFilter.range.startRow,
                    autoFilter.range.endRow,
                    edit.rowNumber,
                    edit.count,
                    "delete"
                )
              : {
                    start: autoFilter.range.startRow,
                    end: autoFilter.range.endRow,
                };
    if (!nextRowRange) {
        return null;
    }

    const nextColumnRange =
        edit.type === "insertColumn"
            ? transformRangeAxis(
                  autoFilter.range.startColumn,
                  autoFilter.range.endColumn,
                  edit.columnNumber,
                  edit.count,
                  "insert"
              )
            : edit.type === "deleteColumn"
              ? transformRangeAxis(
                    autoFilter.range.startColumn,
                    autoFilter.range.endColumn,
                    edit.columnNumber,
                    edit.count,
                    "delete"
                )
              : {
                    start: autoFilter.range.startColumn,
                    end: autoFilter.range.endColumn,
                };
    if (!nextColumnRange) {
        return null;
    }

    const nextSort = autoFilter.sort
        ? (() => {
              if (edit.type !== "insertColumn" && edit.type !== "deleteColumn") {
                  return { ...autoFilter.sort };
              }

              const transformed = transformCellPosition(1, autoFilter.sort.columnNumber, edit);
              return transformed
                  ? {
                        ...autoFilter.sort,
                        columnNumber: transformed.columnNumber,
                    }
                  : null;
          })()
        : null;

    return {
        range: {
            startRow: nextRowRange.start,
            endRow: nextRowRange.end,
            startColumn: nextColumnRange.start,
            endColumn: nextColumnRange.end,
        },
        sort: nextSort,
    };
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
        rowHeights: transformRowHeights(sheet.rowHeights, edit),
        cellAlignments: transformCellAlignments(sheet.cellAlignments, edit),
        rowAlignments: transformRowAlignments(sheet.rowAlignments, edit),
        columnAlignments: transformColumnAlignments(sheet.columnAlignments, edit),
        autoFilter: transformAutoFilter(sheet.autoFilter ?? null, edit),
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
                    rowHeights: {},
                    cellAlignments: {},
                    rowAlignments: {},
                    columnAlignments: {},
                    cells: {},
                    freezePane: null,
                    autoFilter: null,
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
                          rowHeights: edit.rowHeights
                              ? cloneRowHeights(edit.rowHeights)
                              : cloneRowHeights(entry.sheet.rowHeights),
                          cellAlignments: edit.cellAlignments
                              ? cloneCellAlignments(edit.cellAlignments)
                              : cloneCellAlignments(entry.sheet.cellAlignments),
                          rowAlignments: edit.rowAlignments
                              ? cloneRowAlignments(edit.rowAlignments)
                              : cloneRowAlignments(entry.sheet.rowAlignments),
                          columnAlignments: edit.columnAlignments
                              ? cloneColumnAlignments(edit.columnAlignments)
                              : cloneColumnAlignments(entry.sheet.columnAlignments),
                          freezePane: edit.freezePane
                              ? createFreezePaneSnapshot(
                                    edit.freezePane.columnCount,
                                    edit.freezePane.rowCount
                                )
                              : null,
                          autoFilter:
                              edit.autoFilter !== undefined
                                  ? cloneAutoFilterSnapshot(edit.autoFilter)
                                  : cloneAutoFilterSnapshot(entry.sheet.autoFilter),
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
        // View edits can carry large per-sheet alignment maps. Keep references here and
        // clone when persisting/document boundaries actually need ownership.
        viewEdits: pendingViewEdits,
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
        sheetEntries: sheetEntries.map(cloneSheetEntryForStructuralSnapshot),
        pendingCellEdits: pendingCellEdits.map(cloneCellEdit),
        pendingSheetEdits: pendingSheetEdits.map(cloneSheetEdit),
        pendingViewEdits: pendingViewEdits.map(cloneViewEditForStructuralSnapshot),
    };
}

export function restoreStructuralSnapshot(snapshot: StructuralSnapshot): RestoredStructuralState {
    return {
        state: cloneEditorState(snapshot.state),
        sheetEntries: reindexWorkingSheetEntries(snapshot.sheetEntries.map(cloneSheetEntry)),
        pendingCellEdits: snapshot.pendingCellEdits.map(cloneCellEdit),
        pendingSheetEdits: snapshot.pendingSheetEdits.map(cloneSheetEdit),
        pendingViewEdits: snapshot.pendingViewEdits.map(cloneViewEditForStructuralSnapshot),
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
