import type { CellAlignmentSnapshot } from "../../core/model/alignment";

const GRID_CELL_HORIZONTAL_PADDING_PX = 14;

export interface EditorCellOverflowNeighborState {
    value: string;
    formula: string | null;
    blocksOverflow?: boolean;
}

export interface EditorCellOverflowMetrics {
    contentWidthPx: number;
    spillsIntoNextCells: boolean;
}

export function canCellContentSpillRight(
    alignment: CellAlignmentSnapshot | null | undefined
): boolean {
    if (!alignment) {
        return true;
    }

    if (alignment.wrapText || alignment.shrinkToFit) {
        return false;
    }

    return (
        alignment.horizontal === undefined ||
        alignment.horizontal === "general" ||
        alignment.horizontal === "left"
    );
}

export function hasBlockingOverflowContent(
    cell: EditorCellOverflowNeighborState | null | undefined
): boolean {
    if (!cell) {
        return false;
    }

    return cell.blocksOverflow === true || cell.value.length > 0 || Boolean(cell.formula);
}

export function getCellOverflowMetrics({
    value,
    alignment,
    baseColumnWidth,
    visibleColumnNumbers,
    visibleColumnIndex,
    getColumnWidth,
    getTrailingCellState,
}: {
    value: string;
    alignment: CellAlignmentSnapshot | null | undefined;
    baseColumnWidth: number;
    visibleColumnNumbers: readonly number[];
    visibleColumnIndex: number;
    getColumnWidth(columnNumber: number): number;
    getTrailingCellState(columnNumber: number): EditorCellOverflowNeighborState | null;
}): EditorCellOverflowMetrics {
    const contentWidthPx = Math.max(baseColumnWidth - GRID_CELL_HORIZONTAL_PADDING_PX, 0);
    if (!value || !canCellContentSpillRight(alignment)) {
        return {
            contentWidthPx,
            spillsIntoNextCells: false,
        };
    }

    if (
        !Number.isInteger(visibleColumnIndex) ||
        visibleColumnIndex < 0 ||
        visibleColumnIndex >= visibleColumnNumbers.length
    ) {
        return {
            contentWidthPx,
            spillsIntoNextCells: false,
        };
    }

    let nextContentWidthPx = contentWidthPx;
    let spillsIntoNextCells = false;
    for (let index = visibleColumnIndex + 1; index < visibleColumnNumbers.length; index += 1) {
        const nextColumnNumber = visibleColumnNumbers[index];
        const previousColumnNumber = visibleColumnNumbers[index - 1];
        if (
            nextColumnNumber === undefined ||
            previousColumnNumber === undefined ||
            nextColumnNumber !== previousColumnNumber + 1
        ) {
            break;
        }

        if (hasBlockingOverflowContent(getTrailingCellState(nextColumnNumber))) {
            break;
        }

        nextContentWidthPx += getColumnWidth(nextColumnNumber);
        spillsIntoNextCells = true;
    }

    return {
        contentWidthPx: nextContentWidthPx,
        spillsIntoNextCells,
    };
}
