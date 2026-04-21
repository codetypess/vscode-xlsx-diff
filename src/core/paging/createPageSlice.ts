import { DEFAULT_PAGE_SIZE } from "../../constants";
import { createCellKey, getColumnLabel } from "../model/cells";
import {
    type CellDiffStatus,
    type DiffCellLocation,
    type GridCellView,
    type PageSlice,
    type RowFilterMode,
    type SheetDiffModel,
} from "../model/types";

interface CachedGridRowView {
    rowNumber: number;
    hasDiff: boolean;
    diffTone: CellDiffStatus;
    cells: GridCellView[];
}

interface CachedPageBase {
    rowNumbers: number[];
    rangeLabel: string;
    rows: CachedGridRowView[];
    columnDiffTones: Array<CellDiffStatus | null>;
}

interface SheetDiffDerivedCache {
    columns: string[];
    diffRowsSet: Set<number>;
    diffCellByKey: Map<string, DiffCellLocation>;
    diffCellIndexByKey: Map<string, number>;
    sameRows: number[];
    pageBases: Map<string, CachedPageBase>;
}

const sheetDiffDerivedCache = new WeakMap<SheetDiffModel, SheetDiffDerivedCache>();

function buildSameRows(sheet: SheetDiffModel, diffRowsSet: Set<number>): number[] {
    if (sheet.kind === "added" || sheet.kind === "removed") {
        return [];
    }

    const sameRows: number[] = [];
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        if (!diffRowsSet.has(rowNumber)) {
            sameRows.push(rowNumber);
        }
    }

    return sameRows;
}

function getSheetDiffDerivedCache(sheet: SheetDiffModel): SheetDiffDerivedCache {
    const cached = sheetDiffDerivedCache.get(sheet);
    if (cached) {
        return cached;
    }

    const diffRowsSet = new Set(sheet.diffRows);
    const nextCache: SheetDiffDerivedCache = {
        columns: Array.from({ length: sheet.columnCount }, (_, index) => getColumnLabel(index + 1)),
        diffRowsSet,
        diffCellByKey: new Map(sheet.diffCells.map((cell) => [cell.key, cell] as const)),
        diffCellIndexByKey: new Map(sheet.diffCells.map((cell) => [cell.key, cell.diffIndex] as const)),
        sameRows: buildSameRows(sheet, diffRowsSet),
        pageBases: new Map<string, CachedPageBase>(),
    };

    sheetDiffDerivedCache.set(sheet, nextCache);
    return nextCache;
}

function getSameRowCount(sheet: SheetDiffModel): number {
    return getSheetDiffDerivedCache(sheet).sameRows.length;
}

function getDiffTonePriority(status: CellDiffStatus): number {
    switch (status) {
        case "modified":
            return 3;
        case "removed":
            return 2;
        case "added":
            return 1;
        case "equal":
        default:
            return 0;
    }
}

function mergeRowDiffTone(current: CellDiffStatus, next: CellDiffStatus): CellDiffStatus {
    return getDiffTonePriority(next) > getDiffTonePriority(current) ? next : current;
}

function mergeColumnDiffTone(
    current: CellDiffStatus | null,
    next: CellDiffStatus
): CellDiffStatus | null {
    if (next === "equal") {
        return current;
    }

    if (!current) {
        return next;
    }

    return getDiffTonePriority(next) > getDiffTonePriority(current) ? next : current;
}

function getFilteredRowCount(sheet: SheetDiffModel, filter: RowFilterMode): number {
    switch (filter) {
        case "diffs":
            return sheet.diffRows.length;
        case "same":
            return getSameRowCount(sheet);
        case "all":
        default:
            return sheet.rowCount;
    }
}

function clampPage(totalRows: number, currentPage: number): number {
    const totalPages = Math.max(1, Math.ceil(Math.max(totalRows, 1) / DEFAULT_PAGE_SIZE));
    return Math.min(Math.max(currentPage, 1), totalPages);
}

function getRowNumbersForPage(
    cache: SheetDiffDerivedCache,
    sheet: SheetDiffModel,
    filter: RowFilterMode,
    page: number
): number[] {
    const offset = (page - 1) * DEFAULT_PAGE_SIZE;

    switch (filter) {
        case "diffs":
            return sheet.diffRows.slice(offset, offset + DEFAULT_PAGE_SIZE);
        case "same":
            return cache.sameRows.slice(offset, offset + DEFAULT_PAGE_SIZE);
        case "all":
        default: {
            const rowNumbers: number[] = [];
            const lastRow = Math.min(sheet.rowCount, offset + DEFAULT_PAGE_SIZE);

            for (let rowNumber = offset + 1; rowNumber <= lastRow; rowNumber += 1) {
                rowNumbers.push(rowNumber);
            }

            return rowNumbers;
        }
    }
}

function resolveCellStatus(
    sheet: SheetDiffModel,
    leftCellPresent: boolean,
    rightCellPresent: boolean,
    leftValue: string,
    rightValue: string,
    leftFormula: string | null,
    rightFormula: string | null
): CellDiffStatus {
    if (sheet.kind === "added") {
        return rightCellPresent ? "added" : "equal";
    }

    if (sheet.kind === "removed") {
        return leftCellPresent ? "removed" : "equal";
    }

    if (leftCellPresent && !rightCellPresent) {
        return "removed";
    }

    if (!leftCellPresent && rightCellPresent) {
        return "added";
    }

    if (leftValue !== rightValue || leftFormula !== rightFormula) {
        return "modified";
    }

    return "equal";
}

function getHighlightedDiffCell(
    cache: SheetDiffDerivedCache,
    highlightedDiffCellKey: string | null
): DiffCellLocation | null {
    if (!highlightedDiffCellKey) {
        return null;
    }

    return cache.diffCellByKey.get(highlightedDiffCellKey) ?? null;
}

function getCachedPageBase(
    cache: SheetDiffDerivedCache,
    sheet: SheetDiffModel,
    filter: RowFilterMode,
    page: number
): CachedPageBase {
    const cacheKey = `${filter}:${page}`;
    const cached = cache.pageBases.get(cacheKey);
    if (cached) {
        return cached;
    }

    const rowNumbers = getRowNumbersForPage(cache, sheet, filter, page);
    const columnDiffTones = Array.from({ length: sheet.columnCount }, () => null as CellDiffStatus | null);
    const rows = rowNumbers.map((rowNumber) => {
        let diffTone: CellDiffStatus = "equal";
        const cells = Array.from({ length: sheet.columnCount }, (_, columnIndex) => {
            const columnNumber = columnIndex + 1;
            const cellKey = createCellKey(rowNumber, columnNumber);
            const leftCell = sheet.leftSheet?.cells[cellKey];
            const rightCell = sheet.rightSheet?.cells[cellKey];
            const leftValue = leftCell?.displayValue ?? "";
            const rightValue = rightCell?.displayValue ?? "";
            const leftFormula = leftCell?.formula ?? null;
            const rightFormula = rightCell?.formula ?? null;
            const status = resolveCellStatus(
                sheet,
                Boolean(leftCell),
                Boolean(rightCell),
                leftValue,
                rightValue,
                leftFormula,
                rightFormula
            );

            diffTone = mergeRowDiffTone(diffTone, status);
            columnDiffTones[columnIndex] = mergeColumnDiffTone(columnDiffTones[columnIndex], status);

            return {
                key: cellKey,
                address:
                    leftCell?.address ??
                    rightCell?.address ??
                    `${cache.columns[columnIndex]}${rowNumber}`,
                status,
                diffIndex: cache.diffCellIndexByKey.get(cellKey) ?? null,
                leftPresent: Boolean(leftCell),
                rightPresent: Boolean(rightCell),
                leftValue,
                rightValue,
                leftFormula,
                rightFormula,
            };
        });

        return {
            rowNumber,
            hasDiff: cache.diffRowsSet.has(rowNumber),
            diffTone,
            cells,
        };
    });

    const nextBase: CachedPageBase = {
        rowNumbers,
        rangeLabel:
            rowNumbers.length === 0
                ? "No rows"
                : `${rowNumbers[0]}-${rowNumbers[rowNumbers.length - 1]}`,
        rows,
        columnDiffTones,
    };

    cache.pageBases.set(cacheKey, nextBase);
    return nextBase;
}

export function createPageSlice(
    sheet: SheetDiffModel,
    filter: RowFilterMode,
    currentPage: number,
    highlightedDiffCellKey: string | null
): PageSlice {
    const cache = getSheetDiffDerivedCache(sheet);
    const totalRows = getFilteredRowCount(sheet, filter);
    const normalizedPage = clampPage(totalRows, currentPage);
    const pageBase = getCachedPageBase(cache, sheet, filter, normalizedPage);
    const highlightedDiffCell = getHighlightedDiffCell(cache, highlightedDiffCellKey);
    const rows = pageBase.rows.map((row) => ({
        ...row,
        isHighlighted:
            highlightedDiffCell !== null && row.rowNumber === highlightedDiffCell.rowNumber,
    }));

    return {
        filter,
        currentPage: normalizedPage,
        totalPages: Math.max(1, Math.ceil(Math.max(totalRows, 1) / DEFAULT_PAGE_SIZE)),
        totalRows,
        visibleRowCount: pageBase.rowNumbers.length,
        rangeLabel: pageBase.rangeLabel,
        columns: cache.columns,
        columnDiffTones: pageBase.columnDiffTones,
        rows,
        diffRowCount: sheet.diffRows.length,
        diffCellCount: sheet.diffCellCount,
        sameRowCount: getSameRowCount(sheet),
        highlightedDiffRow: highlightedDiffCell?.rowNumber ?? null,
        highlightedDiffCell,
        mergedRangesChanged: sheet.mergedRangesChanged,
    };
}
