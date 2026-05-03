import type { EditorGridMetrics } from "./grid-foundation";

const EDITOR_HEADER_HEIGHT = 28;

export interface EditorFrozenPaneOverlayLayout {
    headerHeight: number;
    rowHeaderWidth: number;
    frozenRowsHeight: number;
    frozenColumnsWidth: number;
    showTopOverlay: boolean;
    showLeftOverlay: boolean;
    showCornerOverlay: boolean;
}

export function deriveEditorFrozenPaneOverlayLayout(
    metrics: EditorGridMetrics
): EditorFrozenPaneOverlayLayout {
    const frozenRowsHeight = Math.max(0, metrics.stickyTopHeight - EDITOR_HEADER_HEIGHT);
    const frozenColumnsWidth = Math.max(0, metrics.stickyLeftWidth - metrics.rowHeaderWidth);
    const showTopOverlay = frozenRowsHeight > 0;
    const showLeftOverlay = frozenColumnsWidth > 0;

    return {
        headerHeight: EDITOR_HEADER_HEIGHT,
        rowHeaderWidth: metrics.rowHeaderWidth,
        frozenRowsHeight,
        frozenColumnsWidth,
        showTopOverlay,
        showLeftOverlay,
        showCornerOverlay: showTopOverlay && showLeftOverlay,
    };
}
