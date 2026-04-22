import { DEFAULT_EDITOR_WINDOW_OVERSCAN } from "../constants";

function clamp(value: number, min: number, max: number): number {
    return Math.min(Math.max(value, min), max);
}

export function getApproximateViewportStartRowForScrollPosition({
    baseRow,
    maxStartRow,
    totalScrollableRows,
    visibleRowCount,
    scrollTop,
    maxScrollTop,
}: {
    baseRow: number;
    maxStartRow: number;
    totalScrollableRows: number;
    visibleRowCount: number;
    scrollTop: number;
    maxScrollTop: number;
}): number {
    if (
        totalScrollableRows <= Math.max(visibleRowCount, 0) ||
        maxStartRow <= baseRow ||
        maxScrollTop <= 0
    ) {
        return baseRow;
    }

    const progress = clamp(scrollTop / maxScrollTop, 0, 1);
    const firstVisibleRow =
        baseRow + Math.round(progress * Math.max(totalScrollableRows - 1, 0));

    return clamp(
        firstVisibleRow - DEFAULT_EDITOR_WINDOW_OVERSCAN,
        baseRow,
        maxStartRow
    );
}
