import { getEditorScrollPositionForCell } from "./editor-virtual-grid";
import type { EditorActiveSheetView } from "../../core/model/types";
import type {
    EditorGridMetrics,
    EditorGridMetricsInput,
    EditorGridViewportPatch,
} from "./grid-foundation";
import { getEditorDisplayRowNumber } from "./grid-foundation";

export interface EditorGridViewportElementLike {
    scrollTop: number;
    scrollLeft: number;
    clientHeight: number;
    clientWidth: number;
}

export interface EditorGridScrollableElementLike extends EditorGridViewportElementLike {
    scrollTo(position: { top?: number; left?: number }): void;
}

export function createEditorGridMetricsInputFromSheet(
    sheet: EditorActiveSheetView
): EditorGridMetricsInput {
    return {
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        rowHeaderLabelCount: sheet.rowCount,
        rowHeights: sheet.rowHeights,
        columnWidths: sheet.columnWidths,
        freezePane: sheet.freezePane,
    };
}

export function createEditorGridViewportPatchFromElement(
    element: EditorGridViewportElementLike
): EditorGridViewportPatch {
    return {
        scrollTop: element.scrollTop,
        scrollLeft: element.scrollLeft,
        viewportHeight: element.clientHeight,
        viewportWidth: element.clientWidth,
    };
}

export function revealEditorGridCellInViewport(
    element: EditorGridScrollableElementLike,
    metrics: EditorGridMetrics,
    rowNumber: number,
    columnNumber: number
): boolean {
    const displayRowNumber = getEditorDisplayRowNumber(metrics, rowNumber);
    if (displayRowNumber === null) {
        return false;
    }

    const nextScrollPosition = getEditorScrollPositionForCell({
        rowNumber: displayRowNumber,
        columnNumber,
        frozenRowCount: metrics.frozenRowCount,
        frozenColumnCount: metrics.frozenColumnCount,
        viewportHeight: metrics.viewport.viewportHeight,
        viewportWidth: metrics.viewport.viewportWidth,
        rowHeaderWidth: metrics.rowHeaderWidth,
        rowLayout: metrics.rowLayout,
        columnLayout: metrics.columnLayout,
    });
    const nextTop = nextScrollPosition.top ?? element.scrollTop;
    const nextLeft = nextScrollPosition.left ?? element.scrollLeft;
    if (nextTop === element.scrollTop && nextLeft === element.scrollLeft) {
        return false;
    }

    element.scrollTo({
        top: nextTop,
        left: nextLeft,
    });
    return true;
}
