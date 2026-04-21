import { DEFAULT_PAGE_SIZE } from "../../constants";
import { createCellKey, getColumnLabel } from "../model/cells";
import {
    type CellDiffStatus,
    type DiffCellLocation,
    type PageSlice,
    type RowFilterMode,
    type SheetDiffModel,
} from "../model/types";

function getSameRowCount(sheet: SheetDiffModel): number {
    if (sheet.kind === "added" || sheet.kind === "removed") {
        return 0;
    }

    return Math.max(0, sheet.rowCount - sheet.diffRows.length);
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
    sheet: SheetDiffModel,
    filter: RowFilterMode,
    page: number
): number[] {
    const offset = (page - 1) * DEFAULT_PAGE_SIZE;

    switch (filter) {
        case "diffs":
            return sheet.diffRows.slice(offset, offset + DEFAULT_PAGE_SIZE);
        case "same": {
            const diffRows = new Set(sheet.diffRows);
            const rowNumbers: number[] = [];
            let skipped = 0;

            for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
                if (diffRows.has(rowNumber)) {
                    continue;
                }

                if (skipped < offset) {
                    skipped += 1;
                    continue;
                }

                rowNumbers.push(rowNumber);
                if (rowNumbers.length === DEFAULT_PAGE_SIZE) {
                    break;
                }
            }

            return rowNumbers;
        }
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

function resolveRowDiffTone(cells: { status: CellDiffStatus }[]): CellDiffStatus {
    if (cells.some((cell) => cell.status === "modified")) {
        return "modified";
    }

    if (cells.some((cell) => cell.status === "removed")) {
        return "removed";
    }

    if (cells.some((cell) => cell.status === "added")) {
        return "added";
    }

    return "equal";
}

function getHighlightedDiffCell(
    sheet: SheetDiffModel,
    highlightedDiffCellKey: string | null
): DiffCellLocation | null {
    if (!highlightedDiffCellKey) {
        return null;
    }

    return sheet.diffCells.find((cell) => cell.key === highlightedDiffCellKey) ?? null;
}

export function createPageSlice(
    sheet: SheetDiffModel,
    filter: RowFilterMode,
    currentPage: number,
    highlightedDiffCellKey: string | null
): PageSlice {
    const totalRows = getFilteredRowCount(sheet, filter);
    const normalizedPage = clampPage(totalRows, currentPage);
    const rowNumbers = getRowNumbersForPage(sheet, filter, normalizedPage);
    const diffRows = new Set(sheet.diffRows);
    const diffCellIndexByKey = new Map(
        sheet.diffCells.map((cell) => [cell.key, cell.diffIndex] as const)
    );
    const highlightedDiffCell = getHighlightedDiffCell(sheet, highlightedDiffCellKey);
    const columns = Array.from({ length: sheet.columnCount }, (_, index) =>
        getColumnLabel(index + 1)
    );
    const rows = rowNumbers.map((rowNumber) => {
        const cells = Array.from({ length: sheet.columnCount }, (_, columnIndex) => {
            const columnNumber = columnIndex + 1;
            const cellKey = createCellKey(rowNumber, columnNumber);
            const leftCell = sheet.leftSheet?.cells[cellKey];
            const rightCell = sheet.rightSheet?.cells[cellKey];
            const leftValue = leftCell?.displayValue ?? "";
            const rightValue = rightCell?.displayValue ?? "";
            const leftFormula = leftCell?.formula ?? null;
            const rightFormula = rightCell?.formula ?? null;
            const diffIndex = diffCellIndexByKey.get(cellKey) ?? null;

            return {
                key: cellKey,
                address:
                    leftCell?.address ??
                    rightCell?.address ??
                    `${columns[columnIndex]}${rowNumber}`,
                status: resolveCellStatus(
                    sheet,
                    Boolean(leftCell),
                    Boolean(rightCell),
                    leftValue,
                    rightValue,
                    leftFormula,
                    rightFormula
                ),
                diffIndex,
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
            hasDiff: diffRows.has(rowNumber),
            isHighlighted:
                highlightedDiffCell !== null && rowNumber === highlightedDiffCell.rowNumber,
            diffTone: resolveRowDiffTone(cells),
            cells,
        };
    });

    return {
        filter,
        currentPage: normalizedPage,
        totalPages: Math.max(1, Math.ceil(Math.max(totalRows, 1) / DEFAULT_PAGE_SIZE)),
        totalRows,
        visibleRowCount: rowNumbers.length,
        rangeLabel:
            rowNumbers.length === 0
                ? "No rows"
                : `${rowNumbers[0]}-${rowNumbers[rowNumbers.length - 1]}`,
        columns,
        rows,
        diffRowCount: sheet.diffRows.length,
        diffCellCount: sheet.diffCellCount,
        sameRowCount: getSameRowCount(sheet),
        highlightedDiffRow: highlightedDiffCell?.rowNumber ?? null,
        highlightedDiffCell,
        mergedRangesChanged: sheet.mergedRangesChanged,
    };
}
