import { DEFAULT_EDITOR_WINDOW_SIZE, DEFAULT_PAGE_SIZE } from "../constants";
import { createCellKey, getColumnLabel } from "../core/model/cells";
import { isChineseDisplayLanguage } from "../displayLanguage";
import {
    type EditorGridCellView,
    type EditorGridRowView,
    type EditorPanelState,
    type EditorRenderModel,
    type EditorSelectedCell,
    type EditorSelectionView,
    type EditorSheetTabView,
    type SheetSnapshot,
    type WorkbookFileView,
    type WorkbookSnapshot,
} from "../core/model/types";

export interface EditorSheetEntry {
    key: string;
    sheet: SheetSnapshot;
    index: number;
}

function getUntitledSheetLabel(): string {
    return isChineseDisplayLanguage() ? "未命名工作表" : "Untitled Sheet";
}

function getWorkbookTitle(workbook: WorkbookSnapshot): string {
    return workbook.titleDetail
        ? `${workbook.fileName} (${workbook.titleDetail})`
        : workbook.fileName;
}

function formatFileSize(bytes: number): string {
    if (bytes < 1024) {
        return `${bytes} B`;
    }

    const units = ["KB", "MB", "GB"];
    let value = bytes / 1024;
    let index = 0;

    while (value >= 1024 && index < units.length - 1) {
        value /= 1024;
        index += 1;
    }

    return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[index]}`;
}

function formatModifiedTime(value: string): string {
    return new Intl.DateTimeFormat(undefined, {
        dateStyle: "medium",
        timeStyle: "short",
    }).format(new Date(value));
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

function getDefaultSelectedCell(sheet: SheetSnapshot, currentPage = 1): EditorSelectedCell | null {
    if (sheet.rowCount === 0 || sheet.columnCount === 0) {
        return null;
    }

    const firstVisibleRow = Math.min(
        sheet.rowCount,
        Math.max(1, (currentPage - 1) * DEFAULT_PAGE_SIZE + 1)
    );

    return {
        rowNumber: firstVisibleRow,
        columnNumber: 1,
    };
}

function getPageStartRow(currentPage: number): number {
    return Math.max(1, (currentPage - 1) * DEFAULT_PAGE_SIZE + 1);
}

function clampPage(sheet: SheetSnapshot, currentPage: number): number {
    const totalPages = Math.max(1, Math.ceil(Math.max(sheet.rowCount, 1) / DEFAULT_PAGE_SIZE));
    return Math.min(Math.max(currentPage, 1), totalPages);
}

function getMinViewportStartRow(sheet: SheetSnapshot): number {
    if (sheet.rowCount <= 0) {
        return 1;
    }

    return Math.min((sheet.freezePane?.rowCount ?? 0) + 1, sheet.rowCount);
}

function clampViewportStartRow(sheet: SheetSnapshot, viewportStartRow: number): number {
    if (sheet.rowCount <= 0) {
        return 1;
    }

    const minStartRow = getMinViewportStartRow(sheet);
    const maxStartRow = Math.max(sheet.rowCount - DEFAULT_EDITOR_WINDOW_SIZE + 1, minStartRow);
    return Math.min(Math.max(Math.trunc(viewportStartRow), minStartRow), maxStartRow);
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

function getCellPage(rowNumber: number): number {
    return Math.floor((Math.max(rowNumber, 1) - 1) / DEFAULT_PAGE_SIZE) + 1;
}

function getSelectionForPage(
    sheet: SheetSnapshot,
    currentPage: number,
    selectedCell: EditorSelectedCell | null
): EditorSelectedCell | null {
    const normalizedSelection = clampSelectedCell(sheet, selectedCell);
    const pageStartRow = (currentPage - 1) * DEFAULT_PAGE_SIZE + 1;
    const pageEndRow = Math.min(sheet.rowCount, currentPage * DEFAULT_PAGE_SIZE);

    if (
        normalizedSelection &&
        normalizedSelection.rowNumber >= pageStartRow &&
        normalizedSelection.rowNumber <= pageEndRow
    ) {
        return normalizedSelection;
    }

    if (sheet.rowCount === 0 || sheet.columnCount === 0) {
        return null;
    }

    return {
        rowNumber: pageStartRow,
        columnNumber: normalizedSelection?.columnNumber ?? 1,
    };
}

function createWorkbookFileView(workbook: WorkbookSnapshot): WorkbookFileView {
    return {
        fileName: workbook.fileName,
        filePath: workbook.filePath,
        fileSizeLabel: formatFileSize(workbook.fileSize),
        detailLabel: workbook.detailLabel,
        detailValue: workbook.detailValue,
        modifiedTimeLabel: workbook.modifiedTimeLabel ?? formatModifiedTime(workbook.modifiedTime),
        isReadonly: workbook.isReadonly ?? false,
    };
}

function createRowsForRange(
    sheet: SheetSnapshot,
    startRow: number,
    endRow: number,
    selectedCell: EditorSelectedCell | null,
    columns: string[]
): EditorGridRowView[] {
    if (sheet.rowCount === 0 || startRow > endRow) {
        return [];
    }

    const rows: EditorGridRowView[] = [];

    for (
        let rowNumber = Math.max(startRow, 1);
        rowNumber <= Math.min(endRow, sheet.rowCount);
        rowNumber += 1
    ) {
        const cells: EditorGridCellView[] = columns.map((columnLabel, columnIndex) => {
            const columnNumber = columnIndex + 1;
            const cellKey = createCellKey(rowNumber, columnNumber);
            const cell = sheet.cells[cellKey];
            const isSelected =
                selectedCell?.rowNumber === rowNumber && selectedCell.columnNumber === columnNumber;

            return {
                key: cellKey,
                address: cell?.address ?? `${columnLabel}${rowNumber}`,
                value: cell?.displayValue ?? "",
                formula: cell?.formula ?? null,
                isPresent: Boolean(cell),
                isSelected,
            };
        });

        rows.push({
            rowNumber,
            isSelected: selectedCell?.rowNumber === rowNumber,
            cells,
        });
    }

    return rows;
}

function createPageRows(
    sheet: SheetSnapshot,
    viewportStartRow: number,
    selectedCell: EditorSelectedCell | null,
    columns: string[]
): EditorGridRowView[] {
    const startRow = clampViewportStartRow(sheet, viewportStartRow);
    return createRowsForRange(
        sheet,
        startRow,
        Math.min(sheet.rowCount, startRow + DEFAULT_EDITOR_WINDOW_SIZE - 1),
        selectedCell,
        columns
    );
}

function createFrozenRows(
    sheet: SheetSnapshot,
    selectedCell: EditorSelectedCell | null,
    columns: string[]
): EditorGridRowView[] {
    const frozenRowCount = Math.max(
        0,
        Math.min(sheet.freezePane?.rowCount ?? 0, Math.max(sheet.rowCount - 1, 0))
    );

    return createRowsForRange(sheet, 1, frozenRowCount, selectedCell, columns);
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

function getWorkbookSummary(workbook: WorkbookSnapshot): {
    totalSheets: number;
    totalRows: number;
    totalNonEmptyCells: number;
} {
    return workbook.sheets.reduce(
        (summary, sheet) => ({
            totalSheets: summary.totalSheets + 1,
            totalRows: summary.totalRows + sheet.rowCount,
            totalNonEmptyCells: summary.totalNonEmptyCells + Object.keys(sheet.cells).length,
        }),
        { totalSheets: 0, totalRows: 0, totalNonEmptyCells: 0 }
    );
}

export function createInitialEditorPanelState(
    workbook: WorkbookSnapshot,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const firstSheet = resolveSheetEntries(workbook, sheetEntriesOverride)[0];

    return {
        activeSheetKey: firstSheet?.key ?? null,
        currentPage: 1,
        viewportStartRow: 1,
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
            currentPage: 1,
            viewportStartRow: 1,
            selectedCell: null,
        };
    }

    const viewportStartRow = clampViewportStartRow(
        activeSheetEntry.sheet,
        state.viewportStartRow || getPageStartRow(state.currentPage)
    );
    const currentPage = clampPage(activeSheetEntry.sheet, state.currentPage);
    const selectedCell =
        clampSelectedCell(activeSheetEntry.sheet, state.selectedCell) ??
        getDefaultSelectedCell(activeSheetEntry.sheet, currentPage);

    return {
        activeSheetKey: activeSheetEntry.key,
        currentPage,
        viewportStartRow,
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
            currentPage: 1,
            viewportStartRow: 1,
            selectedCell: getDefaultSelectedCell(sheetEntry.sheet),
        },
        sheetEntriesOverride
    );
}

export function setEditorCurrentPage(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    currentPage: number,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const normalizedState = normalizeEditorPanelState(workbook, state, sheetEntriesOverride);
    const sheetEntry = resolveSheetEntries(workbook, sheetEntriesOverride).find(
        (entry) => entry.key === normalizedState.activeSheetKey
    );

    if (!sheetEntry) {
        return normalizedState;
    }

    const nextPage = clampPage(sheetEntry.sheet, currentPage);

    return normalizeEditorPanelState(
        workbook,
        {
            ...normalizedState,
            currentPage: nextPage,
            viewportStartRow: getPageStartRow(nextPage),
            selectedCell: getSelectionForPage(
                sheetEntry.sheet,
                nextPage,
                normalizedState.selectedCell
            ),
        },
        sheetEntriesOverride
    );
}

export function moveEditorPageCursor(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    direction: -1 | 1,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const normalizedState = normalizeEditorPanelState(workbook, state, sheetEntriesOverride);
    const sheetEntries = resolveSheetEntries(workbook, sheetEntriesOverride);
    const activeSheetIndex = sheetEntries.findIndex(
        (entry) => entry.key === normalizedState.activeSheetKey
    );

    if (activeSheetIndex < 0) {
        return normalizedState;
    }

    const activeSheet = sheetEntries[activeSheetIndex].sheet;
    const totalPages = Math.max(
        1,
        Math.ceil(Math.max(activeSheet.rowCount, 1) / DEFAULT_PAGE_SIZE)
    );

    if (direction < 0 && normalizedState.currentPage > 1) {
        return setEditorCurrentPage(
            workbook,
            normalizedState,
            normalizedState.currentPage - 1,
            sheetEntriesOverride
        );
    }

    if (direction > 0 && normalizedState.currentPage < totalPages) {
        return setEditorCurrentPage(
            workbook,
            normalizedState,
            normalizedState.currentPage + 1,
            sheetEntriesOverride
        );
    }

    const adjacentSheet =
        direction < 0 ? sheetEntries[activeSheetIndex - 1] : sheetEntries[activeSheetIndex + 1];

    if (!adjacentSheet) {
        return normalizedState;
    }

    const targetPage = direction < 0 ? clampPage(adjacentSheet.sheet, Number.MAX_SAFE_INTEGER) : 1;

    return normalizeEditorPanelState(
        workbook,
        {
            activeSheetKey: adjacentSheet.key,
            currentPage: targetPage,
            viewportStartRow: getPageStartRow(targetPage),
            selectedCell: getDefaultSelectedCell(adjacentSheet.sheet, targetPage),
        },
        sheetEntriesOverride
    );
}

export function setEditorViewportStartRow(
    workbook: WorkbookSnapshot,
    state: EditorPanelState,
    viewportStartRow: number,
    sheetEntriesOverride?: EditorSheetEntry[]
): EditorPanelState {
    const normalizedState = normalizeEditorPanelState(workbook, state, sheetEntriesOverride);
    const sheetEntry = resolveSheetEntries(workbook, sheetEntriesOverride).find(
        (entry) => entry.key === normalizedState.activeSheetKey
    );

    if (!sheetEntry) {
        return normalizedState;
    }

    const nextViewportStartRow = clampViewportStartRow(sheetEntry.sheet, viewportStartRow);
    return normalizeEditorPanelState(
        workbook,
        {
            ...normalizedState,
            currentPage: getCellPage(nextViewportStartRow),
            viewportStartRow: nextViewportStartRow,
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

    const frozenRowCount = Math.max(
        0,
        Math.min(sheetEntry.sheet.freezePane?.rowCount ?? 0, sheetEntry.sheet.rowCount)
    );
    let viewportStartRow = normalizedState.viewportStartRow;
    if (selectedCell.rowNumber > frozenRowCount) {
        if (
            selectedCell.rowNumber < normalizedState.viewportStartRow ||
            selectedCell.rowNumber >=
                normalizedState.viewportStartRow + DEFAULT_EDITOR_WINDOW_SIZE
        ) {
            viewportStartRow = clampViewportStartRow(
                sheetEntry.sheet,
                getPageStartRow(getCellPage(selectedCell.rowNumber))
            );
        }
    }

    return normalizeEditorPanelState(
        workbook,
        {
            ...normalizedState,
            currentPage: getCellPage(selectedCell.rowNumber),
            viewportStartRow,
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
    const file = createWorkbookFileView(workbook);
    const hasPendingEdits = options.hasPendingEdits ?? false;

    if (sheetEntries.length === 0) {
        return {
            title: getWorkbookTitle(workbook),
            file,
            summary: getWorkbookSummary(workbook),
            activeSheet: {
                key: "",
                label: getUntitledSheetLabel(),
                rowCount: 0,
                columnCount: 0,
                hasData: false,
                mergedRangeCount: 0,
                hasMergedRanges: false,
                freezePane: null,
                rowHeights: {},
                columnWidths: {},
            },
            selection: null,
            hasPendingEdits,
            canSave: hasPendingEdits && !file.isReadonly,
            canEdit: !file.isReadonly,
            page: {
                currentPage: 1,
                totalPages: 1,
                totalRows: 0,
                visibleRowCount: 0,
                rangeLabel: "No rows",
                startRow: 0,
                endRow: 0,
                columns: [],
                frozenRows: [],
                rows: [],
            },
            sheets: [],
            canPrevPage: false,
            canNextPage: false,
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
    const pageRows = createPageRows(
        activeSheet,
        normalizedState.viewportStartRow,
        normalizedState.selectedCell,
        columns
    );
    const frozenRows = createFrozenRows(activeSheet, normalizedState.selectedCell, columns);
    const totalPages = Math.max(
        1,
        Math.ceil(Math.max(activeSheet.rowCount, 1) / DEFAULT_PAGE_SIZE)
    );
    const sheetIndex = sheetEntries.findIndex((entry) => entry.key === activeSheetEntry.key);
    const selection = createSelectionView(activeSheet, normalizedState.selectedCell);
    const sheets: EditorSheetTabView[] = sheetEntries.map((entry) => ({
        key: entry.key,
        label: getSheetLabel(entry.sheet),
        rowCount: entry.sheet.rowCount,
        columnCount: entry.sheet.columnCount,
        hasData: Object.keys(entry.sheet.cells).length > 0,
        isActive: entry.key === activeSheetEntry.key,
    }));

    return {
        title: getWorkbookTitle(workbook),
        file,
        summary: getWorkbookSummary(workbook),
        activeSheet: {
            key: activeSheetEntry.key,
            label: getSheetLabel(activeSheet),
            rowCount: activeSheet.rowCount,
            columnCount: activeSheet.columnCount,
            hasData: Object.keys(activeSheet.cells).length > 0,
            mergedRangeCount: activeSheet.mergedRanges.length,
            hasMergedRanges: activeSheet.mergedRanges.length > 0,
            freezePane: activeSheet.freezePane ?? null,
            rowHeights: { ...activeSheet.rowHeights },
            columnWidths: { ...activeSheet.columnWidths },
        },
        selection,
        hasPendingEdits,
        canSave: hasPendingEdits && !file.isReadonly,
        canEdit: !file.isReadonly,
        page: {
            currentPage: normalizedState.currentPage,
            totalPages,
            totalRows: activeSheet.rowCount,
            visibleRowCount: pageRows.length,
            rangeLabel:
                pageRows.length === 0
                    ? "No rows"
                    : `${pageRows[0].rowNumber}-${pageRows[pageRows.length - 1].rowNumber}`,
            startRow: pageRows[0]?.rowNumber ?? 0,
            endRow: pageRows[pageRows.length - 1]?.rowNumber ?? 0,
            columns,
            frozenRows,
            rows: pageRows,
        },
        sheets,
        canPrevPage: normalizedState.currentPage > 1 || sheetIndex > 0,
        canNextPage:
            normalizedState.currentPage < totalPages || sheetIndex < sheetEntries.length - 1,
        canUndoStructuralEdits: options.canUndoStructuralEdits ?? false,
        canRedoStructuralEdits: options.canRedoStructuralEdits ?? false,
    };
}
