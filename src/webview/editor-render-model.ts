import { createCellKey, getColumnLabel } from "../core/model/cells";
import {
    type EditorPanelState,
    type EditorRenderModel,
    type EditorSelectedCell,
    type EditorSelectionView,
    type EditorSheetTabView,
    type SheetSnapshot,
    type WorkbookSnapshot,
} from "../core/model/types";
import { getRuntimeMessages } from "../i18n";

export interface EditorSheetEntry {
    key: string;
    sheet: SheetSnapshot;
    index: number;
}

function getUntitledSheetLabel(): string {
    return getRuntimeMessages().workbook.untitledSheet;
}

function getWorkbookTitle(workbook: WorkbookSnapshot): string {
    return workbook.titleDetail
        ? `${workbook.fileName} (${workbook.titleDetail})`
        : workbook.fileName;
}

export function getEditorSheetKey(sheet: SheetSnapshot): string {
    return sheet.name;
}

function getSheetLabel(sheet: SheetSnapshot): string {
    return sheet.name || getUntitledSheetLabel();
}

function getSheetEntries(workbook: WorkbookSnapshot): Array<{
    key: string;
    sheet: SheetSnapshot;
    index: number;
}> {
    return workbook.sheets.map((sheet, index) => ({
        key: getEditorSheetKey(sheet),
        sheet,
        index,
    }));
}

function resolveSheetEntries(
    workbook: WorkbookSnapshot,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorSheetEntry[] {
    return sheetEntriesOverride ?? getSheetEntries(workbook);
}

function getDefaultSelectedCell(sheet: SheetSnapshot): EditorSelectedCell | null {
    if (sheet.rowCount === 0 || sheet.columnCount === 0) {
        return null;
    }

    return {
        rowNumber: 1,
        columnNumber: 1,
    };
}

function clampSelectedCell(
    sheet: SheetSnapshot,
    selectedCell: EditorSelectedCell | null
): EditorSelectedCell | null {
    if (!selectedCell || sheet.rowCount === 0 || sheet.columnCount === 0) {
        return null;
    }

    return {
        rowNumber: Math.min(Math.max(selectedCell.rowNumber, 1), sheet.rowCount),
        columnNumber: Math.min(Math.max(selectedCell.columnNumber, 1), sheet.columnCount),
    };
}

function createSelectionView(
    sheet: SheetSnapshot,
    selectedCell: EditorSelectedCell | null
): EditorSelectionView | null {
    if (!selectedCell) {
        return null;
    }

    const key = createCellKey(selectedCell.rowNumber, selectedCell.columnNumber);
    const cell = sheet.cells[key];

    return {
        ...selectedCell,
        key,
        address:
            cell?.address ??
            `${getColumnLabel(selectedCell.columnNumber)}${selectedCell.rowNumber}`,
        value: cell?.displayValue ?? "",
        formula: cell?.formula ?? null,
        isPresent: Boolean(cell),
    };
}

export function createInitialEditorPanelState(
    workbook: WorkbookSnapshot,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const firstSheet = resolveSheetEntries(workbook, sheetEntriesOverride)[0];

    return {
        activeSheetKey: firstSheet?.key ?? null,
        selectedCell: firstSheet ? getDefaultSelectedCell(firstSheet.sheet) : null,
    };
}

export function normalizeEditorPanelState(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const sheetEntries = resolveSheetEntries(workbook, sheetEntriesOverride);
    const activeSheetEntry =
        sheetEntries.find((entry) => entry.key === state.activeSheetKey) ?? sheetEntries[0] ?? null;

    if (!activeSheetEntry) {
        return {
            activeSheetKey: null,
            selectedCell: null,
        };
    }

    const selectedCell =
        clampSelectedCell(activeSheetEntry.sheet, state.selectedCell) ??
        getDefaultSelectedCell(activeSheetEntry.sheet);

    return {
        activeSheetKey: activeSheetEntry.key,
        selectedCell,
    };
}

export function setActiveEditorSheet(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    sheetKey: string,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const sheetEntry = resolveSheetEntries(workbook, sheetEntriesOverride).find(
        (entry) => entry.key === sheetKey
    );
    if (!sheetEntry) {
        return state;
    }

    return normalizeEditorPanelState(
        workbook,
        {
            activeSheetKey: sheetEntry.key,
            selectedCell: getDefaultSelectedCell(sheetEntry.sheet),
        },
        sheetEntriesOverride
    );
}

export function setSelectedEditorCell(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    rowNumber: number,
    columnNumber: number,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const normalizedState = normalizeEditorPanelState(workbook, state, sheetEntriesOverride);
    const sheetEntry = resolveSheetEntries(workbook, sheetEntriesOverride).find(
        (entry) => entry.key === normalizedState.activeSheetKey
    );

    if (!sheetEntry) {
        return normalizedState;
    }

    const selectedCell = clampSelectedCell(sheetEntry.sheet, {
        rowNumber,
        columnNumber,
    });

    if (!selectedCell) {
        return normalizedState;
    }

    return normalizeEditorPanelState(
        workbook,
        {
            ...normalizedState,
            selectedCell,
        },
        sheetEntriesOverride
    );
}

export function createEditorRenderModel(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    options: {
        hasPendingEdits?: boolean;
        sheetEntries?: EditorSheetEntry[];
        canUndoStructuralEdits?: boolean;
        canRedoStructuralEdits?: boolean;
    } = {}
): EditorRenderModel {
    const normalizedState = normalizeEditorPanelState(workbook, state, options.sheetEntries);
    const sheetEntries = resolveSheetEntries(workbook, options.sheetEntries);
    const hasPendingEdits = options.hasPendingEdits ?? false;

    if (sheetEntries.length === 0) {
        return {
            title: getWorkbookTitle(workbook),
            activeSheet: {
                key: "",
                rowCount: 0,
                columnCount: 0,
                columns: [],
                cells: {},
                freezePane: null,
            },
            selection: null,
            hasPendingEdits,
            canEdit: !(workbook.isReadonly ?? false),
            sheets: [],
            canUndoStructuralEdits: options.canUndoStructuralEdits ?? false,
            canRedoStructuralEdits: options.canRedoStructuralEdits ?? false,
        };
    }

    const activeSheetEntry =
        sheetEntries.find((entry) => entry.key === normalizedState.activeSheetKey) ??
        sheetEntries[0];
    const activeSheet = activeSheetEntry.sheet;
    const columns = Array.from({ length: activeSheet.columnCount }, (_, index) =>
        getColumnLabel(index + 1)
    );
    const selection = createSelectionView(activeSheet, normalizedState.selectedCell);
    const sheets: EditorSheetTabView[] = sheetEntries.map((entry) => ({
        key: entry.key,
        label: getSheetLabel(entry.sheet),
        isActive: entry.key === activeSheetEntry.key,
    }));

    return {
        title: getWorkbookTitle(workbook),
        activeSheet: {
            key: activeSheetEntry.key,
            rowCount: activeSheet.rowCount,
            columnCount: activeSheet.columnCount,
            columns,
            cells: activeSheet.cells,
            freezePane: activeSheet.freezePane ?? null,
        },
        selection,
        hasPendingEdits,
        canEdit: !(workbook.isReadonly ?? false),
        sheets,
        canUndoStructuralEdits: options.canUndoStructuralEdits ?? false,
        canRedoStructuralEdits: options.canRedoStructuralEdits ?? false,
    };
}
