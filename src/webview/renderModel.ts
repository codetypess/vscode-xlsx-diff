import { DEFAULT_PAGE_SIZE } from "../constants";
import { createPageSlice } from "../core/paging/createPageSlice";
import { isChineseDisplayLanguage } from "../displayLanguage";
import {
    type CellDiffStatus,
    type PanelState,
    type RenderModel,
    type RowFilterMode,
    type SheetDiffModel,
    type SheetTabView,
    type WorkbookDiffModel,
    type WorkbookSnapshot,
} from "../core/model/types";

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

function getSheetLabel(sheet: SheetDiffModel): string {
    if (sheet.kind === "renamed") {
        return `${sheet.leftSheetName} -> ${sheet.rightSheetName}`;
    }

    return sheet.rightSheetName ?? sheet.leftSheetName ?? getUntitledSheetLabel();
}

function getSheetHasDiff(sheet: SheetDiffModel): boolean {
    return sheet.kind !== "matched" || sheet.diffCellCount > 0 || sheet.mergedRangesChanged;
}

function getSheetDiffTone(sheet: SheetDiffModel): CellDiffStatus {
    if (sheet.kind === "added") {
        return "added";
    }

    if (sheet.kind === "removed") {
        return "removed";
    }

    return getSheetHasDiff(sheet) ? "modified" : "equal";
}

function getFilteredRowCount(sheet: SheetDiffModel, filter: RowFilterMode): number {
    switch (filter) {
        case "diffs":
            return sheet.diffRows.length;
        case "same":
            return sheet.kind === "added" || sheet.kind === "removed"
                ? 0
                : Math.max(0, sheet.rowCount - sheet.diffRows.length);
        case "all":
        default:
            return sheet.rowCount;
    }
}

function getTotalPages(sheet: SheetDiffModel, filter: RowFilterMode): number {
    return Math.max(
        1,
        Math.ceil(Math.max(getFilteredRowCount(sheet, filter), 1) / DEFAULT_PAGE_SIZE)
    );
}

function getFirstDiffCellKey(sheet: SheetDiffModel): string | null {
    return sheet.diffCells[0]?.key ?? null;
}

function getDiffCellPage(sheet: SheetDiffModel, filter: RowFilterMode, rowNumber: number): number {
    if (filter === "diffs") {
        const diffRowIndex = sheet.diffRows.indexOf(rowNumber);
        return diffRowIndex >= 0 ? Math.floor(diffRowIndex / DEFAULT_PAGE_SIZE) + 1 : 1;
    }

    return Math.floor((rowNumber - 1) / DEFAULT_PAGE_SIZE) + 1;
}

function clampPage(sheet: SheetDiffModel, filter: RowFilterMode, page: number): number {
    const totalPages = getTotalPages(sheet, filter);
    return Math.min(Math.max(page, 1), totalPages);
}

export function createInitialPanelState(diff: WorkbookDiffModel): PanelState {
    const firstSheet = diff.sheets[0];

    return {
        activeSheetKey: firstSheet?.key ?? null,
        filter: "all",
        currentPage: 1,
        highlightedDiffCellKey: getFirstDiffCellKey(firstSheet),
    };
}

export function normalizePanelState(diff: WorkbookDiffModel, state: PanelState): PanelState {
    const activeSheet =
        diff.sheets.find((sheet) => sheet.key === state.activeSheetKey) ?? diff.sheets[0] ?? null;

    if (!activeSheet) {
        return {
            activeSheetKey: null,
            filter: "all",
            currentPage: 1,
            highlightedDiffCellKey: null,
        };
    }

    const highlightedDiffCellKey =
        state.filter === "same" || activeSheet.diffCells.length === 0
            ? null
            : activeSheet.diffCells.some((cell) => cell.key === state.highlightedDiffCellKey)
              ? state.highlightedDiffCellKey
              : getFirstDiffCellKey(activeSheet);

    return {
        activeSheetKey: activeSheet.key,
        filter: state.filter,
        currentPage: clampPage(activeSheet, state.filter, state.currentPage),
        highlightedDiffCellKey,
    };
}

export function setActiveSheet(
    diff: WorkbookDiffModel,
    state: PanelState,
    sheetKey: string
): PanelState {
    const activeSheet = diff.sheets.find((sheet) => sheet.key === sheetKey);
    if (!activeSheet) {
        return state;
    }

    return normalizePanelState(diff, {
        activeSheetKey: activeSheet.key,
        filter: state.filter,
        currentPage: 1,
        highlightedDiffCellKey: getFirstDiffCellKey(activeSheet),
    });
}

export function setFilterMode(
    diff: WorkbookDiffModel,
    state: PanelState,
    filter: RowFilterMode
): PanelState {
    return normalizePanelState(diff, {
        ...state,
        filter,
        currentPage: 1,
        highlightedDiffCellKey: filter === "same" ? null : state.highlightedDiffCellKey,
    });
}

export function setCurrentPage(
    diff: WorkbookDiffModel,
    state: PanelState,
    currentPage: number
): PanelState {
    return normalizePanelState(diff, {
        ...state,
        currentPage,
    });
}

export function movePageCursor(
    diff: WorkbookDiffModel,
    state: PanelState,
    direction: -1 | 1
): PanelState {
    const normalizedState = normalizePanelState(diff, state);
    const activeSheetIndex = diff.sheets.findIndex(
        (sheet) => sheet.key === normalizedState.activeSheetKey
    );

    if (activeSheetIndex < 0) {
        return normalizedState;
    }

    const activeSheet = diff.sheets[activeSheetIndex];
    const activeSheetTotalPages = getTotalPages(activeSheet, normalizedState.filter);

    if (direction < 0 && normalizedState.currentPage > 1) {
        return normalizePanelState(diff, {
            ...normalizedState,
            currentPage: normalizedState.currentPage - 1,
        });
    }

    if (direction > 0 && normalizedState.currentPage < activeSheetTotalPages) {
        return normalizePanelState(diff, {
            ...normalizedState,
            currentPage: normalizedState.currentPage + 1,
        });
    }

    const adjacentSheet =
        direction < 0 ? diff.sheets[activeSheetIndex - 1] : diff.sheets[activeSheetIndex + 1];

    if (!adjacentSheet) {
        return normalizedState;
    }

    return normalizePanelState(diff, {
        activeSheetKey: adjacentSheet.key,
        filter: normalizedState.filter,
        currentPage: direction < 0 ? getTotalPages(adjacentSheet, normalizedState.filter) : 1,
        highlightedDiffCellKey:
            normalizedState.filter === "same" ? null : getFirstDiffCellKey(adjacentSheet),
    });
}

export function setHighlightedDiffCell(
    diff: WorkbookDiffModel,
    state: PanelState,
    rowNumber: number,
    columnNumber?: number
): PanelState {
    const normalizedState = normalizePanelState(diff, state);
    const activeSheet = diff.sheets.find((sheet) => sheet.key === normalizedState.activeSheetKey);

    if (!activeSheet) {
        return normalizedState;
    }

    const targetCell = activeSheet.diffCells.find(
        (cell) =>
            cell.rowNumber === rowNumber &&
            (columnNumber === undefined || cell.columnNumber === columnNumber)
    );

    if (!targetCell) {
        return normalizedState;
    }

    const filter = normalizedState.filter === "same" ? "diffs" : normalizedState.filter;
    const currentPage = getDiffCellPage(activeSheet, filter, targetCell.rowNumber);

    return normalizePanelState(diff, {
        ...normalizedState,
        filter,
        currentPage,
        highlightedDiffCellKey: targetCell.key,
    });
}

export function setHighlightedDiffRow(
    diff: WorkbookDiffModel,
    state: PanelState,
    rowNumber: number
): PanelState {
    return setHighlightedDiffCell(diff, state, rowNumber);
}

export function moveDiffCursor(
    diff: WorkbookDiffModel,
    state: PanelState,
    direction: -1 | 1
): PanelState {
    const normalizedState = normalizePanelState(diff, state);
    const activeSheet = diff.sheets.find((sheet) => sheet.key === normalizedState.activeSheetKey);

    if (!activeSheet || activeSheet.diffCells.length === 0) {
        return normalizedState;
    }

    const filter = normalizedState.filter === "same" ? "diffs" : normalizedState.filter;
    const currentIndex = normalizedState.highlightedDiffCellKey
        ? activeSheet.diffCells.findIndex(
              (cell) => cell.key === normalizedState.highlightedDiffCellKey
          )
        : direction > 0
          ? -1
          : activeSheet.diffCells.length;
    const nextIndex = Math.min(
        Math.max(currentIndex + direction, 0),
        activeSheet.diffCells.length - 1
    );
    const nextHighlightedCell = activeSheet.diffCells[nextIndex];
    const nextPage = getDiffCellPage(activeSheet, filter, nextHighlightedCell.rowNumber);

    return normalizePanelState(diff, {
        activeSheetKey: activeSheet.key,
        filter,
        currentPage: nextPage,
        highlightedDiffCellKey: nextHighlightedCell.key,
    });
}

export function createRenderModel(diff: WorkbookDiffModel, state: PanelState): RenderModel {
    const normalizedState = normalizePanelState(diff, state);
    const activeSheet =
        diff.sheets.find((sheet) => sheet.key === normalizedState.activeSheetKey) ?? diff.sheets[0];
    const page = createPageSlice(
        activeSheet,
        normalizedState.filter,
        normalizedState.currentPage,
        normalizedState.highlightedDiffCellKey
    );
    const currentDiffIndex =
        normalizedState.highlightedDiffCellKey === null
            ? -1
            : activeSheet.diffCells.findIndex(
                  (cell) => cell.key === normalizedState.highlightedDiffCellKey
              );
    const activeSheetIndex = diff.sheets.findIndex((sheet) => sheet.key === activeSheet.key);

    const sheets: SheetTabView[] = diff.sheets.map((sheet) => ({
        key: sheet.key,
        label: getSheetLabel(sheet),
        kind: sheet.kind,
        diffRowCount: sheet.diffRows.length,
        diffCellCount: sheet.diffCellCount,
        mergedRangesChanged: sheet.mergedRangesChanged,
        hasDiff: getSheetHasDiff(sheet),
        diffTone: getSheetDiffTone(sheet),
        isActive: sheet.key === activeSheet.key,
    }));

    return {
        title: `${getWorkbookTitle(diff.left)} ↔ ${getWorkbookTitle(diff.right)}`,
        leftFile: {
            fileName: diff.left.fileName,
            filePath: diff.left.filePath,
            fileSizeLabel: formatFileSize(diff.left.fileSize),
            detailLabel: diff.left.detailLabel,
            detailValue: diff.left.detailValue,
            modifiedTimeLabel:
                diff.left.modifiedTimeLabel ?? formatModifiedTime(diff.left.modifiedTime),
            isReadonly: diff.left.isReadonly ?? false,
        },
        rightFile: {
            fileName: diff.right.fileName,
            filePath: diff.right.filePath,
            fileSizeLabel: formatFileSize(diff.right.fileSize),
            detailLabel: diff.right.detailLabel,
            detailValue: diff.right.detailValue,
            modifiedTimeLabel:
                diff.right.modifiedTimeLabel ?? formatModifiedTime(diff.right.modifiedTime),
            isReadonly: diff.right.isReadonly ?? false,
        },
        summary: {
            totalSheets: diff.sheets.length,
            diffSheets: diff.totalDiffSheets,
            diffRows: diff.totalDiffRows,
            diffCells: diff.totalDiffCells,
        },
        activeSheet: {
            key: activeSheet.key,
            label: getSheetLabel(activeSheet),
            kind: activeSheet.kind,
            leftName: activeSheet.leftSheetName,
            rightName: activeSheet.rightSheetName,
            hasDiff: getSheetHasDiff(activeSheet),
            mergedRangesChanged: activeSheet.mergedRangesChanged,
        },
        filter: normalizedState.filter,
        page,
        sheets,
        canPrevPage: page.currentPage > 1 || activeSheetIndex > 0,
        canNextPage:
            page.currentPage < page.totalPages || activeSheetIndex < diff.sheets.length - 1,
        canPrevDiff: currentDiffIndex > 0,
        canNextDiff: currentDiffIndex >= 0 && currentDiffIndex < activeSheet.diffCells.length - 1,
    };
}
