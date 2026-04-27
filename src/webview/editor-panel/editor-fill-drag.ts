import type { PendingHistoryChange } from "./editor-pending-history";
import {
    isCellWithinSelectionRange,
    type SelectionPositionLike,
    type SelectionRange,
} from "./editor-selection-range";

export interface FillBounds {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
}

export interface BuildFillChangesOptions {
    sheetKey: string;
    sourceRange: SelectionRange;
    previewRange: SelectionRange | null;
    getCellValue(rowNumber: number, columnNumber: number): string;
    getModelValue(rowNumber: number, columnNumber: number): string;
    canEditCell(rowNumber: number, columnNumber: number): boolean;
}

export interface AutoFillDownPreviewRangeOptions {
    sourceRange: SelectionRange;
    bounds: FillBounds;
    getCellValue(rowNumber: number, columnNumber: number): string;
}

interface NumericProgression {
    firstValue: number;
    step: number;
}

function clamp(value: number, min: number, max: number): number {
    return Math.max(min, Math.min(value, max));
}

function isRangeExpandedBeyondSource(
    sourceRange: SelectionRange,
    previewRange: SelectionRange
): boolean {
    return (
        previewRange.startRow < sourceRange.startRow ||
        previewRange.endRow > sourceRange.endRow ||
        previewRange.startColumn < sourceRange.startColumn ||
        previewRange.endColumn > sourceRange.endColumn
    );
}

function getRangeHeight(range: SelectionRange): number {
    return range.endRow - range.startRow + 1;
}

function getRangeWidth(range: SelectionRange): number {
    return range.endColumn - range.startColumn + 1;
}

function hasCellContent(value: string): boolean {
    return value !== "";
}

function modulo(value: number, divisor: number): number {
    return ((value % divisor) + divisor) % divisor;
}

function parsePureNumber(value: string): number | null {
    const normalizedValue = value.trim();
    if (!normalizedValue) {
        return null;
    }

    if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalizedValue)) {
        return null;
    }

    const numericValue = Number(normalizedValue);
    return Number.isFinite(numericValue) ? numericValue : null;
}

function areNumbersEffectivelyEqual(left: number, right: number): boolean {
    return Math.abs(left - right) <= 1e-9 * Math.max(1, Math.abs(left), Math.abs(right));
}

function getNumericProgression(values: string[]): NumericProgression | null {
    if (values.length < 2) {
        return null;
    }

    const numericValues = values.map((value) => parsePureNumber(value));
    if (numericValues.some((value) => value === null)) {
        return null;
    }

    const resolvedValues = numericValues as number[];
    const step = resolvedValues[1]! - resolvedValues[0]!;
    for (let index = 2; index < resolvedValues.length; index += 1) {
        const previousValue = resolvedValues[index - 1]!;
        const currentValue = resolvedValues[index]!;
        if (!areNumbersEffectivelyEqual(currentValue - previousValue, step)) {
            return null;
        }
    }

    return {
        firstValue: resolvedValues[0]!,
        step,
    };
}

function formatNumericFillValue(value: number): string {
    if (Object.is(value, -0)) {
        return "0";
    }

    const normalizedValue = Number(value.toPrecision(15));
    return Object.is(normalizedValue, -0) ? "0" : String(normalizedValue);
}

function getSeedValues(
    sourceRange: SelectionRange,
    getCellValue: (rowNumber: number, columnNumber: number) => string
): string[][] {
    const values: string[][] = [];

    for (let rowNumber = sourceRange.startRow; rowNumber <= sourceRange.endRow; rowNumber += 1) {
        const rowValues: string[] = [];
        for (
            let columnNumber = sourceRange.startColumn;
            columnNumber <= sourceRange.endColumn;
            columnNumber += 1
        ) {
            rowValues.push(getCellValue(rowNumber, columnNumber));
        }
        values.push(rowValues);
    }

    return values;
}

function createTiledValueGetter(sourceRange: SelectionRange, seedValues: string[][]) {
    const sourceHeight = getRangeHeight(sourceRange);
    const sourceWidth = getRangeWidth(sourceRange);

    return (rowNumber: number, columnNumber: number): string =>
        seedValues[modulo(rowNumber - sourceRange.startRow, sourceHeight)]?.[
            modulo(columnNumber - sourceRange.startColumn, sourceWidth)
        ] ?? "";
}

function createSeriesValueGetter(
    sourceRange: SelectionRange,
    previewRange: SelectionRange,
    seedValues: string[][]
): ((rowNumber: number, columnNumber: number) => string) | null {
    const sourceHeight = getRangeHeight(sourceRange);
    const sourceWidth = getRangeWidth(sourceRange);
    const expandsRows =
        previewRange.startRow < sourceRange.startRow || previewRange.endRow > sourceRange.endRow;
    const expandsColumns =
        previewRange.startColumn < sourceRange.startColumn ||
        previewRange.endColumn > sourceRange.endColumn;

    if (sourceHeight === 1 && sourceWidth >= 2 && expandsColumns) {
        const progression = getNumericProgression(seedValues[0] ?? []);
        if (!progression) {
            return null;
        }

        return (_rowNumber: number, columnNumber: number): string =>
            formatNumericFillValue(
                progression.firstValue + (columnNumber - sourceRange.startColumn) * progression.step
            );
    }

    if (sourceWidth === 1 && sourceHeight >= 2 && expandsRows) {
        const progression = getNumericProgression(seedValues.map((rowValues) => rowValues[0] ?? ""));
        if (!progression) {
            return null;
        }

        return (rowNumber: number, _columnNumber: number): string =>
            formatNumericFillValue(
                progression.firstValue + (rowNumber - sourceRange.startRow) * progression.step
            );
    }

    return null;
}

export function createFillPreviewRange(
    sourceRange: SelectionRange,
    targetCell: SelectionPositionLike,
    bounds: FillBounds
): SelectionRange | null {
    const targetRow = clamp(targetCell.rowNumber, bounds.minRow, bounds.maxRow);
    const targetColumn = clamp(targetCell.columnNumber, bounds.minColumn, bounds.maxColumn);
    const previewRange = {
        startRow: Math.min(sourceRange.startRow, targetRow),
        endRow: Math.max(sourceRange.endRow, targetRow),
        startColumn: Math.min(sourceRange.startColumn, targetColumn),
        endColumn: Math.max(sourceRange.endColumn, targetColumn),
    };

    return isRangeExpandedBeyondSource(sourceRange, previewRange) ? previewRange : null;
}

function getContiguousFilledEndRow(
    startRow: number,
    columnNumber: number,
    maxRow: number,
    getCellValue: (rowNumber: number, columnNumber: number) => string
): number {
    let endRow = startRow - 1;

    for (let rowNumber = startRow; rowNumber <= maxRow; rowNumber += 1) {
        if (!hasCellContent(getCellValue(rowNumber, columnNumber))) {
            break;
        }

        endRow = rowNumber;
    }

    return endRow;
}

export function createAutoFillDownPreviewRange({
    sourceRange,
    bounds,
    getCellValue,
}: AutoFillDownPreviewRangeOptions): SelectionRange | null {
    const startRow = sourceRange.endRow + 1;
    if (startRow > bounds.maxRow) {
        return null;
    }

    const leftColumnNumber =
        sourceRange.startColumn > bounds.minColumn ? sourceRange.startColumn - 1 : null;
    const rightColumnNumber =
        sourceRange.endColumn < bounds.maxColumn ? sourceRange.endColumn + 1 : null;
    const leftEndRow =
        leftColumnNumber === null
            ? null
            : getContiguousFilledEndRow(startRow, leftColumnNumber, bounds.maxRow, getCellValue);
    const rightEndRow =
        rightColumnNumber === null
            ? null
            : getContiguousFilledEndRow(startRow, rightColumnNumber, bounds.maxRow, getCellValue);
    let previewEndRow = sourceRange.endRow;

    if (leftEndRow !== null) {
        previewEndRow = Math.max(previewEndRow, leftEndRow);
    }

    if (rightEndRow !== null) {
        previewEndRow = Math.max(previewEndRow, rightEndRow);
    }

    if (previewEndRow > sourceRange.endRow) {
        return {
            startRow: sourceRange.startRow,
            endRow: previewEndRow,
            startColumn: sourceRange.startColumn,
            endColumn: sourceRange.endColumn,
        };
    }

    if (bounds.maxRow > sourceRange.endRow) {
        return {
            startRow: sourceRange.startRow,
            endRow: bounds.maxRow,
            startColumn: sourceRange.startColumn,
            endColumn: sourceRange.endColumn,
        };
    }

    return null;
}

export function isCellWithinFillPreviewArea(
    sourceRange: SelectionRange,
    previewRange: SelectionRange | null,
    rowNumber: number,
    columnNumber: number
): boolean {
    return Boolean(
        previewRange &&
            isCellWithinSelectionRange(previewRange, rowNumber, columnNumber) &&
            !isCellWithinSelectionRange(sourceRange, rowNumber, columnNumber)
    );
}

export function buildFillChanges({
    sheetKey,
    sourceRange,
    previewRange,
    getCellValue,
    getModelValue,
    canEditCell,
}: BuildFillChangesOptions): PendingHistoryChange[] {
    if (!previewRange || !isRangeExpandedBeyondSource(sourceRange, previewRange)) {
        return [];
    }

    const seedValues = getSeedValues(sourceRange, getCellValue);
    const getSeriesValue = createSeriesValueGetter(sourceRange, previewRange, seedValues);
    const getTiledValue = createTiledValueGetter(sourceRange, seedValues);
    const getAfterValue = getSeriesValue ?? getTiledValue;
    const changes: PendingHistoryChange[] = [];

    for (let rowNumber = previewRange.startRow; rowNumber <= previewRange.endRow; rowNumber += 1) {
        for (
            let columnNumber = previewRange.startColumn;
            columnNumber <= previewRange.endColumn;
            columnNumber += 1
        ) {
            if (!isCellWithinFillPreviewArea(sourceRange, previewRange, rowNumber, columnNumber)) {
                continue;
            }

            if (!canEditCell(rowNumber, columnNumber)) {
                continue;
            }

            const beforeValue = getCellValue(rowNumber, columnNumber);
            const afterValue = getAfterValue(rowNumber, columnNumber);
            if (beforeValue === afterValue) {
                continue;
            }

            changes.push({
                sheetKey,
                rowNumber,
                columnNumber,
                modelValue: getModelValue(rowNumber, columnNumber),
                beforeValue,
                afterValue,
            });
        }
    }

    return changes;
}
