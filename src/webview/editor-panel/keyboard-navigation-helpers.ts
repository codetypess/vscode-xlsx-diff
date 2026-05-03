import type { EditorActiveSheetView, EditorSelectionView } from "../../core/model/types";

export interface EditorKeyboardNavigationDelta {
    rowDelta: number;
    columnDelta: number;
}

export interface EditorKeyboardNavigationTarget {
    rowNumber: number;
    columnNumber: number;
}

export type EditorKeyboardPageDirection = -1 | 1;

function clamp(value: number, minimum: number, maximum: number): number {
    return Math.max(minimum, Math.min(maximum, value));
}

export function getEditorKeyboardNavigationDelta(
    key: string
): EditorKeyboardNavigationDelta | null {
    switch (key) {
        case "ArrowUp":
            return { rowDelta: -1, columnDelta: 0 };
        case "ArrowDown":
            return { rowDelta: 1, columnDelta: 0 };
        case "ArrowLeft":
            return { rowDelta: 0, columnDelta: -1 };
        case "ArrowRight":
            return { rowDelta: 0, columnDelta: 1 };
        case "Tab":
            return { rowDelta: 0, columnDelta: 1 };
        case "Enter":
            return { rowDelta: 1, columnDelta: 0 };
        default:
            return null;
    }
}

export function isEditorClearCellKey(key: string): boolean {
    return key === "Backspace" || key === "Delete";
}

export function getEditorKeyboardPageDirection(key: string): EditorKeyboardPageDirection | null {
    switch (key) {
        case "PageUp":
            return -1;
        case "PageDown":
            return 1;
        default:
            return null;
    }
}

export function getNextEditorKeyboardNavigationTarget({
    activeSheet,
    selection,
    delta,
    visibleRowNumbers,
}: {
    activeSheet: EditorActiveSheetView;
    selection: EditorSelectionView | null;
    delta: EditorKeyboardNavigationDelta;
    visibleRowNumbers?: readonly number[];
}): EditorKeyboardNavigationTarget | null {
    if (activeSheet.rowCount < 1 || activeSheet.columnCount < 1) {
        return null;
    }

    const rowNumber =
        delta.rowDelta !== 0 && visibleRowNumbers && visibleRowNumbers.length > 0
            ? getNextVisibleRowNumber({
                  visibleRowNumbers,
                  currentRowNumber: selection?.rowNumber ?? visibleRowNumbers[0] ?? 1,
                  rowDelta: delta.rowDelta,
              })
            : clamp((selection?.rowNumber ?? 1) + delta.rowDelta, 1, activeSheet.rowCount);
    const columnNumber = clamp(
        (selection?.columnNumber ?? 1) + delta.columnDelta,
        1,
        activeSheet.columnCount
    );

    return {
        rowNumber,
        columnNumber,
    };
}

export function getNextEditorViewportPageNavigationTarget({
    activeSheet,
    selection,
    direction,
    visibleRowCount,
    visibleRowNumbers,
}: {
    activeSheet: EditorActiveSheetView;
    selection: EditorSelectionView | null;
    direction: EditorKeyboardPageDirection;
    visibleRowCount: number;
    visibleRowNumbers?: readonly number[];
}): EditorKeyboardNavigationTarget | null {
    if (activeSheet.rowCount < 1 || activeSheet.columnCount < 1) {
        return null;
    }

    const rowStep = Math.max(1, visibleRowCount - 1);
    return {
        rowNumber:
            visibleRowNumbers && visibleRowNumbers.length > 0
                ? getNextVisibleRowNumber({
                      visibleRowNumbers,
                      currentRowNumber: selection?.rowNumber ?? visibleRowNumbers[0] ?? 1,
                      rowDelta: direction * rowStep,
                  })
                : clamp((selection?.rowNumber ?? 1) + direction * rowStep, 1, activeSheet.rowCount),
        columnNumber: clamp(selection?.columnNumber ?? 1, 1, activeSheet.columnCount),
    };
}

function getNextVisibleRowNumber({
    visibleRowNumbers,
    currentRowNumber,
    rowDelta,
}: {
    visibleRowNumbers: readonly number[];
    currentRowNumber: number;
    rowDelta: number;
}): number {
    if (visibleRowNumbers.length === 0) {
        return currentRowNumber;
    }

    const currentIndex = visibleRowNumbers.indexOf(currentRowNumber);
    const fallbackIndex = visibleRowNumbers.findIndex((rowNumber) => rowNumber >= currentRowNumber);
    const startIndex =
        currentIndex >= 0
            ? currentIndex
            : fallbackIndex >= 0
              ? fallbackIndex
              : visibleRowNumbers.length - 1;
    const nextIndex = clamp(startIndex + rowDelta, 0, visibleRowNumbers.length - 1);
    return visibleRowNumbers[nextIndex] ?? currentRowNumber;
}
