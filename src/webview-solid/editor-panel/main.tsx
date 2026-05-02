import { For, Show, createEffect, createMemo, createSignal, onCleanup, onMount } from "solid-js";
import { render } from "solid-js/web";
import { createCellKey, getCellAddress } from "../../core/model/cells";
import type { EditorAlignmentPatch } from "../../core/model/alignment";
import type { EditorActiveSheetView, EditorSelectedCell } from "../../core/model/types";
import {
    EDITOR_EXTRA_PADDING_ROWS,
    getEditorColumnLeft,
    getEditorColumnWidth,
    getMinimumVisibleEditorRowCount,
} from "../../webview/editor-panel/editor-virtual-grid";
import {
    DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
    MAX_COLUMN_PIXEL_WIDTH,
    MIN_COLUMN_PIXEL_WIDTH,
    convertPixelsToWorkbookColumnWidth,
    stabilizeColumnPixelWidth,
} from "../../webview/column-layout";
import {
    buildFillChanges,
    createFillPreviewRange,
    isCellWithinFillPreviewArea,
    type FillBounds,
} from "../../webview/editor-panel/editor-fill-drag";
import {
    rebasePendingHistory,
    type PendingHistoryChange,
    type PendingHistoryEntry,
} from "../../webview/editor-panel/editor-pending-history";
import {
    MAX_ROW_PIXEL_HEIGHT,
    MIN_ROW_PIXEL_HEIGHT,
    convertPixelsToWorkbookRowHeight,
    stabilizeRowPixelHeight,
} from "../../webview/row-layout";
import {
    isCellWithinSelectionRange,
    type SelectionRange,
} from "../../webview/editor-panel/editor-selection-range";
import {
    clearEditorFilterColumn,
    createEditorSheetFilterSnapshot,
    createEditorSheetFilterStateFromSnapshot,
    getEditorFilterColumnValues,
    getEditorVisibleRows,
    isEditorFilterHeaderCell,
    normalizeEditorFilterRange,
    resolveEditorFilterRangeFromActiveCell,
    resolveEditorFilterRangeFromSelection,
    toggleEditorSheetFilterState,
    updateEditorFilterIncludedValues,
    updateEditorFilterSort,
    type EditorFilterCellSource,
    type EditorFilterSortDirection,
    type EditorFilterValueOption,
    type EditorSheetFilterState,
} from "../../webview/editor-panel/editor-panel-filter";
import { resolveEditorReplaceResultInSheet } from "../../webview/editor-panel/editor-panel-logic";
import type {
    EditorPendingEdit,
    EditorPanelStrings,
    EditorSearchDirection,
    EditorSearchResultMessage,
    EditorSearchScope,
    SearchOptions,
} from "../../webview/editor-panel/editor-panel-types";
import {
    createWebviewReadyMessage,
    isEditorSessionIncomingMessage,
    type EditorSessionIncomingMessage,
    type EditorWebviewOutgoingMessage,
} from "../shared/session-protocol";
import {
    createEditorAddSheetMessage,
    createEditorDeleteSheetMessage,
    createEditorGotoMessage,
    createEditorRenameSheetMessage,
    createEditorSearchMessage,
    createEditorSetSheetMessage,
    getEditorSheetContextMenuState,
    getEditorShellCapabilities,
} from "./shell-helpers";
import {
    applyEditorGridViewportPatch,
    createInitialEditorGridViewportState,
    deriveEditorGridMetrics,
    getEditorDisplayRowNumber,
    getEditorGridActualRowHeight,
    getEditorGridActualRowTop,
    reuseEquivalentEditorGridMetrics,
    type EditorGridMetrics,
} from "./grid-foundation";
import {
    createEditorColumnHeaderSelection,
    createEditorRowHeaderSelection,
    deriveEditorGridHeaderLayers,
    EDITOR_GRID_HEADER_HEIGHT,
} from "./header-layer-helpers";
import {
    deriveEditorGridCellLayers,
    type EditorGridCellItem,
    type EditorGridCellLayers,
    type EditorGridCellLayerKind,
    reuseEquivalentEditorGridCellLayers,
} from "./cell-layer-helpers";
import {
    getActiveEditorAlignmentSelectionTarget,
    getActiveEditorToolbarAlignment,
    getCellContentAlignmentStyle,
} from "./alignment-helpers";
import {
    createEditorCellEditingState,
    getEditorPendingEditValue,
    getEditorSelectionDisplayValue,
    isEditorCellEditingActive,
    type EditorCellEditingState,
} from "./editing-surface-helpers";
import { createInitialEditorSessionState, reduceEditorSessionMessage } from "./session";
import {
    deriveEditorSelectionOverlayLayers,
    type EditorSelectionOverlayLayer,
    type EditorSelectionOverlayRangeRect,
    type EditorSelectionOverlayRect,
} from "./selection-overlay-helpers";
import { createOptimisticEditorSelection } from "./selection-model-helpers";
import {
    createEditorAnchoredRangeSelectionState,
    createEditorColumnSelectionState,
    createEditorExtendedCellSelectionState,
    createEditorRowSelectionState,
    createEditorSingleCellSelectionState,
    resolveEditorSelectionRange,
    type EditorSelectionRangeState,
} from "./selection-range-state-helpers";
import {
    createEditorGridMetricsInputFromSheet,
    createEditorGridViewportPatchFromElement,
    revealEditorGridCellInViewport,
} from "./viewport-services";
import {
    getEditorKeyboardNavigationDelta,
    getEditorKeyboardPageDirection,
    getNextEditorKeyboardNavigationTarget,
    getNextEditorViewportPageNavigationTarget,
    isEditorClearCellKey,
} from "./keyboard-navigation-helpers";

interface VsCodeApi {
    postMessage(message: EditorWebviewOutgoingMessage): void;
}

function getVsCodeApi(): VsCodeApi {
    const candidate = (globalThis as Record<string, unknown>).acquireVsCodeApi;
    if (typeof candidate === "function") {
        return (candidate as () => VsCodeApi)();
    }

    return {
        postMessage: () => undefined,
    };
}

type EditorSearchMode = "find" | "replace";
type SearchFeedbackTone = "success" | "warn" | "error";
type SearchFeedbackStatus = EditorSearchResultMessage["status"] | "replaced" | "no-change";

interface SearchFeedbackState {
    status: SearchFeedbackStatus;
    tone: SearchFeedbackTone;
    message: string;
}

interface SheetContextMenuState {
    sheetKey: string;
    x: number;
    y: number;
}

type GridContextMenuState =
    | {
          kind: "cell";
          rowNumber: number;
          columnNumber: number;
          x: number;
          y: number;
      }
    | {
          kind: "row";
          rowNumber: number;
          x: number;
          y: number;
      }
    | {
          kind: "column";
          columnNumber: number;
          x: number;
          y: number;
      };

interface SearchPanelPosition {
    left: number;
    top: number;
}

interface SearchPanelDragState {
    pointerId: number;
    offsetX: number;
    offsetY: number;
}

interface FilterMenuState {
    sheetKey: string;
    columnNumber: number;
    left: number;
    top: number;
}

interface EditorGridFillDragState {
    pointerId: number;
    sourceRange: SelectionRange;
    previewRange: SelectionRange | null;
}

interface EditorColumnResizeState {
    pointerId: number;
    sheetKey: string;
    columnNumber: number;
    startClientX: number;
    startPixelWidth: number;
    previewPixelWidth: number;
    isDragging: boolean;
}

interface EditorRowResizeState {
    pointerId: number;
    sheetKey: string;
    rowNumber: number;
    startClientY: number;
    startPixelHeight: number;
    previewPixelHeight: number;
    isDragging: boolean;
}

type EditorGridSelectionDragKind = "cell" | "row" | "column";

interface EditorGridSelectionDragState {
    kind: EditorGridSelectionDragKind;
    pointerId: number;
    anchorCell: EditorSelectedCell;
    lastTargetKey: string;
}

type EditorGridSelectionTarget =
    | {
          kind: "cell";
          rowNumber: number;
          columnNumber: number;
      }
    | {
          kind: "row";
          rowNumber: number;
      }
    | {
          kind: "column";
          columnNumber: number;
      };

interface EditorGridFillHandleLayout {
    layer: EditorGridCellLayerKind;
    rowNumber: number;
    columnNumber: number;
    left: number;
    top: number;
}

const FILL_HANDLE_SIZE = 8;
const FILL_HANDLE_HALF_SIZE = FILL_HANDLE_SIZE / 2;

function isTextInputTarget(target: EventTarget | null): boolean {
    if (!(target instanceof HTMLElement)) {
        return false;
    }

    return Boolean(
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target.isContentEditable ||
        target.closest('input, textarea, [contenteditable="true"], [contenteditable=""]')
    );
}

function isEditableTextInputTarget(target: EventTarget | null): boolean {
    if (!(target instanceof Element)) {
        return false;
    }

    if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
        return !target.readOnly && !target.disabled;
    }

    if (target instanceof HTMLElement && target.isContentEditable) {
        return true;
    }

    const editableTarget = target.closest(
        'input, textarea, [contenteditable="true"], [contenteditable=""]'
    );
    if (
        editableTarget instanceof HTMLInputElement ||
        editableTarget instanceof HTMLTextAreaElement
    ) {
        return !editableTarget.readOnly && !editableTarget.disabled;
    }

    return editableTarget instanceof HTMLElement ? editableTarget.isContentEditable : false;
}

function isSearchPanelInteractiveTarget(target: EventTarget | null): boolean {
    if (!(target instanceof Element)) {
        return false;
    }

    return Boolean(
        target.closest(
            'button, input, textarea, select, a, [role="button"], [role="tab"], .search-strip__input-wrap'
        )
    );
}

function parseEditorGridCoordinate(value: string | undefined): number | null {
    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed >= 1 ? parsed : null;
}

function getEditorGridSelectionTargetKey(target: EditorGridSelectionTarget): string {
    switch (target.kind) {
        case "cell":
            return `cell:${target.rowNumber}:${target.columnNumber}`;
        case "row":
            return `row:${target.rowNumber}`;
        case "column":
            return `column:${target.columnNumber}`;
    }
}

function resolveEditorGridSelectionTarget(
    target: EventTarget | null
): EditorGridSelectionTarget | null {
    if (!(target instanceof Element)) {
        return null;
    }

    const cellElement = target.closest('[data-role="grid-cell"]');
    if (cellElement instanceof HTMLElement) {
        const rowNumber = parseEditorGridCoordinate(cellElement.dataset.rowNumber);
        const columnNumber = parseEditorGridCoordinate(cellElement.dataset.columnNumber);
        if (rowNumber !== null && columnNumber !== null) {
            return {
                kind: "cell",
                rowNumber,
                columnNumber,
            };
        }
    }

    const rowHeaderElement = target.closest('[data-role="grid-row-header"]');
    if (rowHeaderElement instanceof HTMLElement) {
        const rowNumber = parseEditorGridCoordinate(rowHeaderElement.dataset.rowNumber);
        if (rowNumber !== null) {
            return {
                kind: "row",
                rowNumber,
            };
        }
    }

    const columnHeaderElement = target.closest('[data-role="grid-column-header"]');
    if (columnHeaderElement instanceof HTMLElement) {
        const columnNumber = parseEditorGridCoordinate(columnHeaderElement.dataset.columnNumber);
        if (columnNumber !== null) {
            return {
                kind: "column",
                columnNumber,
            };
        }
    }

    return null;
}

function areSelectionRangesEqual(
    left: SelectionRange | null,
    right: SelectionRange | null
): boolean {
    if (left === right) {
        return true;
    }

    if (!left || !right) {
        return left === right;
    }

    return (
        left.startRow === right.startRow &&
        left.endRow === right.endRow &&
        left.startColumn === right.startColumn &&
        left.endColumn === right.endColumn
    );
}

function hasExpandedSelectionRange(range: SelectionRange | null): range is SelectionRange {
    return Boolean(
        range && (range.startRow !== range.endRow || range.startColumn !== range.endColumn)
    );
}

function focusSearchInput(): void {
    const input = document.querySelector<HTMLInputElement>('[data-role="search-input"]');
    if (!input) {
        return;
    }

    input.focus();
    input.select();
}

function focusReplaceInput(): void {
    const input = document.querySelector<HTMLInputElement>('[data-role="replace-input"]');
    if (!input) {
        return;
    }

    input.focus();
    input.select();
}

function focusGotoInput(): void {
    const input = document.querySelector<HTMLInputElement>('[data-role="goto-input"]');
    if (!input) {
        return;
    }

    input.focus();
    input.select();
}

function normalizeWorkbookColumnWidth(columnWidth: number): number {
    return Math.round(columnWidth * 256) / 256;
}

function normalizeWorkbookRowHeight(rowHeight: number): number {
    return Math.round(rowHeight * 100) / 100;
}

function getEditorPendingEditKey(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number
): string {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
}

function getEditorGridFillBounds(metrics: EditorGridMetrics | null): FillBounds | null {
    if (!metrics || metrics.displayRowCount < 1 || metrics.displayColumnCount < 1) {
        return null;
    }

    return {
        minRow: 1,
        maxRow:
            metrics.rowState.actualRowNumbers[metrics.rowState.actualRowNumbers.length - 1] ?? 1,
        minColumn: 1,
        maxColumn: metrics.displayColumnCount,
    };
}

function getEditorGridLayerForCell(
    metrics: EditorGridMetrics,
    cell: Pick<EditorSelectedCell, "rowNumber" | "columnNumber"> | null
): EditorGridCellLayerKind | null {
    if (!cell) {
        return null;
    }

    const isFrozenRow = metrics.window.frozenRowNumbers.includes(cell.rowNumber);
    const isFrozenColumn = metrics.window.frozenColumnNumbers.includes(cell.columnNumber);
    const isVisibleRow = isFrozenRow || metrics.window.rowNumbers.includes(cell.rowNumber);
    const isVisibleColumn =
        isFrozenColumn || metrics.window.columnNumbers.includes(cell.columnNumber);

    if (!isVisibleRow || !isVisibleColumn) {
        return null;
    }

    if (isFrozenRow && isFrozenColumn) {
        return "corner";
    }

    if (isFrozenRow) {
        return "top";
    }

    if (isFrozenColumn) {
        return "left";
    }

    return "body";
}

function clampSearchPanelPosition(
    position: SearchPanelPosition,
    {
        appElement,
        panelElement,
    }: {
        appElement?: { clientWidth: number; clientHeight: number } | null;
        panelElement?: { offsetWidth: number; offsetHeight: number } | null;
    } = {}
): SearchPanelPosition {
    if (!appElement || !panelElement) {
        return position;
    }

    const margin = 12;
    const maxLeft = Math.max(margin, appElement.clientWidth - panelElement.offsetWidth - margin);
    const maxTop = Math.max(margin, appElement.clientHeight - panelElement.offsetHeight - margin);

    return {
        left: Math.max(margin, Math.min(position.left, maxLeft)),
        top: Math.max(margin, Math.min(position.top, maxTop)),
    };
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === "object";
}

function isEditorSearchResultMessage(value: unknown): value is EditorSearchResultMessage {
    return (
        isRecord(value) &&
        value.type === "searchResult" &&
        (value.status === "matched" ||
            value.status === "no-match" ||
            value.status === "invalid-pattern") &&
        (value.scope === "sheet" || value.scope === "selection")
    );
}

function formatTemplate(template: string, values: Record<string, string>): string {
    return template.replace(/\{(\w+)\}/g, (_match, key: string) => values[key] ?? "");
}

function getSelectionOverlayStyle(
    rect: EditorSelectionOverlayRect,
    borders?: Pick<
        EditorSelectionOverlayRangeRect,
        "showTopBorder" | "showRightBorder" | "showBottomBorder" | "showLeftBorder"
    >
): Record<string, string> {
    return {
        width: `${rect.width}px`,
        height: `${rect.height}px`,
        transform: `translate(${rect.left}px, ${rect.top}px)`,
        ...(borders
            ? {
                  "--selection-overlay-border-top": borders.showTopBorder ? "2px" : "0px",
                  "--selection-overlay-border-right": borders.showRightBorder ? "2px" : "0px",
                  "--selection-overlay-border-bottom": borders.showBottomBorder ? "2px" : "0px",
                  "--selection-overlay-border-left": borders.showLeftBorder ? "2px" : "0px",
              }
            : {}),
    };
}

function SelectionOverlayLayerView(props: {
    overlay: EditorSelectionOverlayLayer;
    offsetLeft: number;
    offsetTop: number;
    isSearchFocused?: boolean;
}) {
    const renderRect = (
        rect: EditorSelectionOverlayRect,
        className: string,
        borders?: Pick<
            EditorSelectionOverlayRangeRect,
            "showTopBorder" | "showRightBorder" | "showBottomBorder" | "showLeftBorder"
        >
    ) => (
        <div
            aria-hidden
            class={className}
            style={getSelectionOverlayStyle(
                {
                    ...rect,
                    left: rect.left - props.offsetLeft,
                    top: rect.top - props.offsetTop,
                },
                borders
            )}
        />
    );

    return (
        <>
            <Show when={props.overlay.activeRowRect}>
                {(rect) =>
                    renderRect(
                        rect(),
                        "editor-grid__selection-overlay editor-grid__selection-overlay--active-row"
                    )
                }
            </Show>
            <Show when={props.overlay.activeColumnRect}>
                {(rect) =>
                    renderRect(
                        rect(),
                        "editor-grid__selection-overlay editor-grid__selection-overlay--active-column"
                    )
                }
            </Show>
            <Show when={props.overlay.rangeRect}>
                {(rect) =>
                    renderRect(
                        rect(),
                        "editor-grid__selection-overlay editor-grid__selection-overlay--range",
                        rect()
                    )
                }
            </Show>
            <Show when={props.overlay.primaryRect}>
                {(rect) =>
                    renderRect(
                        rect(),
                        `editor-grid__selection-overlay editor-grid__selection-overlay--primary${
                            props.isSearchFocused
                                ? " editor-grid__selection-overlay--search-focus"
                                : ""
                        }`
                    )
                }
            </Show>
        </>
    );
}

function CellEditorInput(props: {
    edit: EditorCellEditingState;
    onUpdateDraft: (value: string) => void;
    onCommit: () => void;
    onCancel: () => void;
}) {
    let inputElement: HTMLInputElement | undefined;

    onMount(() => {
        inputElement?.focus();
        inputElement?.select();
    });

    return (
        <input
            ref={(element) => {
                inputElement = element;
            }}
            class="grid__cell-input"
            data-role="grid-cell-input"
            type="text"
            value={props.edit.draftValue}
            onBlur={() => {
                setTimeout(() => props.onCommit(), 0);
            }}
            onInput={(event) => props.onUpdateDraft(event.currentTarget.value)}
            onClick={(event) => event.stopPropagation()}
            onDblClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === "Tab") {
                    event.preventDefault();
                    props.onCommit();
                } else if (event.key === "Escape") {
                    event.preventDefault();
                    props.onCancel();
                }
            }}
        />
    );
}

function EditorGridCornerHeaderView(props: { rowHeaderWidth: number; headerHeight: number }) {
    return (
        <div
            class="editor-grid__item editor-grid__item--corner"
            style={{
                width: `${props.rowHeaderWidth}px`,
                height: `${props.headerHeight}px`,
                transform: "translate(0px, 0px)",
                display: "flex",
                "align-items": "center",
                "justify-content": "center",
                color: "var(--vscode-descriptionForeground)",
                "font-size": "11px",
                "font-weight": "600",
                "text-transform": "uppercase",
                "letter-spacing": "0.08em",
            }}
        >
            <span aria-hidden>#</span>
        </div>
    );
}

function EditorGridColumnHeaderView(props: {
    left: number;
    width: number;
    height: number;
    columnNumber: number;
    label: string;
    isActive: boolean;
    canResize: boolean;
    isResizing: boolean;
    onPointerSelectStart: (pointerId: number, options?: { extend?: boolean }) => void;
    onPointerResizeStart: (
        pointerId: number,
        columnNumber: number,
        startPixelWidth: number,
        clientX: number
    ) => void;
    onOpenContextMenu: (columnNumber: number, clientX: number, clientY: number) => void;
}) {
    return (
        <div
            class="editor-grid__item editor-grid__item--header"
            data-role="grid-column-header"
            data-column-number={props.columnNumber}
            style={{
                width: `${props.width}px`,
                height: `${props.height}px`,
                transform: `translate(${props.left}px, 0px)`,
                background: props.isActive
                    ? "color-mix(in srgb, var(--vscode-editorInfo-foreground, #3794ff) 10%, var(--vscode-editor-background) 90%)"
                    : undefined,
            }}
            onPointerDown={(event) => {
                if (event.button !== 0) {
                    return;
                }

                event.preventDefault();
                props.onPointerSelectStart(event.pointerId, {
                    extend: event.shiftKey,
                });
            }}
            onContextMenu={(event) => {
                event.preventDefault();
                props.onOpenContextMenu(props.columnNumber, event.clientX, event.clientY);
            }}
        >
            <div class="grid__column">
                <span class="grid__column-label">{props.label}</span>
                <Show when={props.canResize}>
                    <span
                        aria-hidden
                        class="grid__column-resize-handle"
                        classList={{ "is-active": props.isResizing }}
                        data-role="grid-column-resize-handle"
                        data-column-number={props.columnNumber}
                        onPointerDown={(event) => {
                            if (event.button !== 0) {
                                return;
                            }

                            event.preventDefault();
                            event.stopPropagation();
                            props.onPointerResizeStart(
                                event.pointerId,
                                props.columnNumber,
                                props.width,
                                event.clientX
                            );
                        }}
                    />
                </Show>
            </div>
        </div>
    );
}

function EditorGridRowHeaderView(props: {
    top: number;
    height: number;
    rowNumber: number;
    rowHeaderWidth: number;
    isActive: boolean;
    canResize: boolean;
    isResizing: boolean;
    onPointerSelectStart: (pointerId: number, options?: { extend?: boolean }) => void;
    onPointerResizeStart: (
        pointerId: number,
        rowNumber: number,
        startPixelHeight: number,
        clientY: number
    ) => void;
    onOpenContextMenu: (rowNumber: number, clientX: number, clientY: number) => void;
}) {
    return (
        <div
            class="editor-grid__item editor-grid__item--row-header"
            data-role="grid-row-header"
            data-row-number={props.rowNumber}
            style={{
                width: `${props.rowHeaderWidth}px`,
                height: `${props.height}px`,
                transform: `translate(0px, ${props.top}px)`,
                background: props.isActive
                    ? "color-mix(in srgb, var(--vscode-editorInfo-foreground, #3794ff) 10%, var(--vscode-editor-background) 90%)"
                    : undefined,
            }}
            onPointerDown={(event) => {
                if (event.button !== 0) {
                    return;
                }

                event.preventDefault();
                props.onPointerSelectStart(event.pointerId, {
                    extend: event.shiftKey,
                });
            }}
            onContextMenu={(event) => {
                event.preventDefault();
                props.onOpenContextMenu(props.rowNumber, event.clientX, event.clientY);
            }}
        >
            <div class="grid__row-number">
                <span class="grid__row-label">{props.rowNumber}</span>
                <Show when={props.canResize}>
                    <span
                        aria-hidden
                        class="grid__row-resize-handle"
                        classList={{ "is-active": props.isResizing }}
                        data-role="grid-row-resize-handle"
                        data-row-number={props.rowNumber}
                        onPointerDown={(event) => {
                            if (event.button !== 0) {
                                return;
                            }

                            event.preventDefault();
                            event.stopPropagation();
                            props.onPointerResizeStart(
                                event.pointerId,
                                props.rowNumber,
                                props.height,
                                event.clientY
                            );
                        }}
                    />
                </Show>
            </div>
        </div>
    );
}

function EditorGridFillHandleView(props: {
    layout: EditorGridFillHandleLayout;
    isActive: boolean;
    onPointerDown: (pointerId: number) => void;
}) {
    return (
        <span
            aria-hidden
            class="editor-grid__fill-handle"
            classList={{ "is-active": props.isActive }}
            data-role="grid-fill-handle"
            data-row-number={props.layout.rowNumber}
            data-column-number={props.layout.columnNumber}
            style={{
                width: `${FILL_HANDLE_SIZE}px`,
                height: `${FILL_HANDLE_SIZE}px`,
                transform: `translate(${props.layout.left}px, ${props.layout.top}px)`,
            }}
            onPointerDown={(event) => {
                if (event.button !== 0) {
                    return;
                }

                event.preventDefault();
                event.stopPropagation();
                props.onPointerDown(event.pointerId);
            }}
            onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
            }}
        />
    );
}

function GridCellValueView(props: { value: string }) {
    if (!props.value) {
        return null;
    }

    return <span class="grid__cell-value">{props.value}</span>;
}

function GridCellFormulaBadgeView(props: { formula: string | null }) {
    if (!props.formula) {
        return null;
    }

    return (
        <span class="cell__formula" title={props.formula}>
            fx
        </span>
    );
}

function EditorGridCellView(props: {
    item: EditorGridCellItem;
    sheetKey: string | null;
    offsetLeft: number;
    offsetTop: number;
    editingCell: EditorCellEditingState | null;
    fillSourceRange: SelectionRange | null;
    fillPreviewRange: SelectionRange | null;
    onPointerSelectStart: (
        rowNumber: number,
        columnNumber: number,
        pointerId: number,
        options?: { extend?: boolean }
    ) => void;
    onStartEdit: (rowNumber: number, columnNumber: number) => void;
    onUpdateDraft: (value: string) => void;
    onCommitEdit: () => void;
    onCancelEdit: () => void;
    isFilterMenuOpen: boolean;
    onOpenFilterMenu: (rowNumber: number, columnNumber: number, element: HTMLButtonElement) => void;
    onOpenContextMenu: (
        rowNumber: number,
        columnNumber: number,
        clientX: number,
        clientY: number
    ) => void;
}) {
    const isEditing = () =>
        isEditorCellEditingActive({
            editingCell: props.editingCell,
            sheetKey: props.sheetKey,
            rowNumber: props.item.rowNumber,
            columnNumber: props.item.columnNumber,
        });
    const isFillPreview = () =>
        Boolean(
            props.fillSourceRange &&
            isCellWithinFillPreviewArea(
                props.fillSourceRange,
                props.fillPreviewRange,
                props.item.rowNumber,
                props.item.columnNumber
            )
        );
    const shouldSpillIntoNextCells = () => props.item.spillsIntoNextCells && !isEditing();
    const style = () => ({
        width: `${props.item.width}px`,
        height: `${props.item.height}px`,
        "min-width": `${props.item.width}px`,
        "max-width": `${props.item.width}px`,
        transform: `translate(${props.item.left - props.offsetLeft}px, ${props.item.top - props.offsetTop}px)`,
        "--grid-column-max-width": `${props.item.width}px`,
        "--grid-cell-content-max-height": `${props.item.contentMaxHeightPx}px`,
        "--grid-cell-line-clamp": String(props.item.visibleLineCount),
        "--grid-cell-display-max-width":
            shouldSpillIntoNextCells() && props.item.displayMaxWidthPx !== null
                ? `${props.item.displayMaxWidthPx}px`
                : undefined,
        background: props.item.isSelected
            ? "color-mix(in srgb, var(--vscode-editorInfo-foreground, #3794ff) 12%, var(--vscode-editor-background) 88%)"
            : undefined,
    });

    return (
        <div
            class="editor-grid__item grid__cell"
            classList={{
                "grid__cell--filter-header": props.item.isFilterHeader,
                "grid__cell--pending": props.item.isPending,
                "grid__cell--fill-preview": isFillPreview(),
                "grid__cell--editing": isEditing(),
                "grid__cell--overflow-spill": shouldSpillIntoNextCells(),
            }}
            data-role="grid-cell"
            data-cell-address={getCellAddress(props.item.rowNumber, props.item.columnNumber)}
            data-row-number={props.item.rowNumber}
            data-column-number={props.item.columnNumber}
            title={props.item.formula ?? props.item.displayValue}
            style={style()}
            onPointerDown={(event) => {
                if (event.button !== 0 || isTextInputTarget(event.target)) {
                    return;
                }

                event.preventDefault();
                props.onPointerSelectStart(
                    props.item.rowNumber,
                    props.item.columnNumber,
                    event.pointerId,
                    {
                        extend: event.shiftKey,
                    }
                );
            }}
            onDblClick={(event) => {
                event.preventDefault();
                props.onStartEdit(props.item.rowNumber, props.item.columnNumber);
            }}
            onContextMenu={(event) => {
                if (isTextInputTarget(event.target)) {
                    return;
                }

                event.preventDefault();
                props.onOpenContextMenu(
                    props.item.rowNumber,
                    props.item.columnNumber,
                    event.clientX,
                    event.clientY
                );
            }}
        >
            <Show
                when={isEditing() ? props.editingCell : null}
                fallback={
                    <span
                        class="grid__cell-content"
                        style={getCellContentAlignmentStyle(props.item.alignment)}
                    >
                        <GridCellValueView value={props.item.displayValue} />
                    </span>
                }
            >
                {(edit) => (
                    <CellEditorInput
                        edit={edit()}
                        onUpdateDraft={props.onUpdateDraft}
                        onCommit={props.onCommitEdit}
                        onCancel={props.onCancelEdit}
                    />
                )}
            </Show>
            <GridCellFormulaBadgeView formula={props.item.formula} />
            <Show when={props.item.isFilterHeader}>
                <button
                    class="grid__cell-filterButton"
                    classList={{
                        "is-active": props.item.isColumnFilterActive,
                        "is-open": props.isFilterMenuOpen,
                    }}
                    data-role="filter-trigger"
                    type="button"
                    onPointerDown={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                    }}
                    onClick={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        props.onOpenFilterMenu(
                            props.item.rowNumber,
                            props.item.columnNumber,
                            event.currentTarget
                        );
                    }}
                >
                    <span class="codicon codicon-chevron-down" aria-hidden />
                </button>
            </Show>
        </div>
    );
}

function getFilterOptionLabel(value: string, blankValueLabel: string): string {
    return value ? value : blankValueLabel;
}

function EditorFilterMenuView(props: {
    strings: Pick<
        EditorPanelStrings,
        | "sortAscending"
        | "sortDescending"
        | "filterSearchPlaceholder"
        | "filterSelectAll"
        | "filterClearColumn"
        | "filterBlankValue"
        | "filterNoValues"
    >;
    menu: FilterMenuState | null;
    filterState: EditorSheetFilterState | null;
    options: readonly EditorFilterValueOption[];
    onSort: (columnNumber: number, direction: EditorFilterSortDirection) => void;
    onSetIncludedValues: (columnNumber: number, includedValues: readonly string[] | null) => void;
    onClearColumn: (columnNumber: number) => void;
}) {
    const [query, setQuery] = createSignal("");
    createEffect(() => {
        props.menu?.sheetKey;
        props.menu?.columnNumber;
        setQuery("");
    });

    const isOpen = createMemo(
        () =>
            Boolean(props.menu && props.filterState) &&
            props.menu!.columnNumber >= props.filterState!.range.startColumn &&
            props.menu!.columnNumber <= props.filterState!.range.endColumn
    );
    const allValues = createMemo(() => props.options.map((option) => option.value));
    const includedValues = createMemo(() => {
        const menu = props.menu;
        const filterState = props.filterState;
        if (!menu || !filterState) {
            return [];
        }

        return filterState.includedValuesByColumn[String(menu.columnNumber)] ?? allValues();
    });
    const includedValuesSet = createMemo(() => new Set(includedValues()));
    const filteredOptions = createMemo(() => {
        const normalizedQuery = query().trim().toLowerCase();
        if (!normalizedQuery) {
            return props.options;
        }

        return props.options.filter((option) =>
            getFilterOptionLabel(option.value, props.strings.filterBlankValue)
                .toLowerCase()
                .includes(normalizedQuery)
        );
    });
    const isAllSelected = createMemo(
        () => props.options.length > 0 && includedValues().length === props.options.length
    );
    const hasColumnCriteria = createMemo(() => {
        const menu = props.menu;
        const filterState = props.filterState;
        if (!menu || !filterState) {
            return false;
        }

        return (
            Boolean(filterState.includedValuesByColumn[String(menu.columnNumber)]) ||
            filterState.sort?.columnNumber === menu.columnNumber
        );
    });

    return (
        <Show when={isOpen() && props.menu && props.filterState}>
            <div
                class="filter-menu"
                data-role="filter-menu"
                style={{
                    left: `${props.menu!.left}px`,
                    top: `${props.menu!.top}px`,
                }}
            >
                <div class="filter-menu__sorts">
                    <button
                        class="filter-menu__sortButton"
                        classList={{
                            "is-active":
                                props.filterState!.sort?.columnNumber ===
                                    props.menu!.columnNumber &&
                                props.filterState!.sort.direction === "asc",
                        }}
                        type="button"
                        onClick={() => props.onSort(props.menu!.columnNumber, "asc")}
                    >
                        <span class="codicon codicon-arrow-up" aria-hidden />
                        <span>{props.strings.sortAscending}</span>
                    </button>
                    <button
                        class="filter-menu__sortButton"
                        classList={{
                            "is-active":
                                props.filterState!.sort?.columnNumber ===
                                    props.menu!.columnNumber &&
                                props.filterState!.sort.direction === "desc",
                        }}
                        type="button"
                        onClick={() => props.onSort(props.menu!.columnNumber, "desc")}
                    >
                        <span class="codicon codicon-arrow-down" aria-hidden />
                        <span>{props.strings.sortDescending}</span>
                    </button>
                </div>
                <div class="filter-menu__search">
                    <span class="codicon codicon-search filter-menu__searchIcon" aria-hidden />
                    <input
                        class="filter-menu__searchInput"
                        placeholder={props.strings.filterSearchPlaceholder}
                        type="text"
                        value={query()}
                        onInput={(event) => setQuery(event.currentTarget.value)}
                    />
                </div>
                <div class="filter-menu__values">
                    <Show
                        when={props.options.length > 0}
                        fallback={
                            <div class="filter-menu__empty">{props.strings.filterNoValues}</div>
                        }
                    >
                        <label class="filter-menu__option">
                            <input
                                checked={isAllSelected()}
                                type="checkbox"
                                onChange={(event) =>
                                    props.onSetIncludedValues(
                                        props.menu!.columnNumber,
                                        event.currentTarget.checked ? null : []
                                    )
                                }
                            />
                            <span>{props.strings.filterSelectAll}</span>
                        </label>
                        <div class="filter-menu__divider" />
                        <div class="filter-menu__optionList">
                            <For each={filteredOptions()}>
                                {(option) => {
                                    const optionLabel = () =>
                                        getFilterOptionLabel(
                                            option.value,
                                            props.strings.filterBlankValue
                                        );
                                    const isChecked = () => includedValuesSet().has(option.value);
                                    return (
                                        <label class="filter-menu__option">
                                            <input
                                                checked={isChecked()}
                                                type="checkbox"
                                                onChange={(event) => {
                                                    const nextIncludedValues = new Set(
                                                        includedValues()
                                                    );
                                                    if (event.currentTarget.checked) {
                                                        nextIncludedValues.add(option.value);
                                                    } else {
                                                        nextIncludedValues.delete(option.value);
                                                    }

                                                    const normalizedIncludedValues =
                                                        allValues().filter((value) =>
                                                            nextIncludedValues.has(value)
                                                        );
                                                    props.onSetIncludedValues(
                                                        props.menu!.columnNumber,
                                                        normalizedIncludedValues.length ===
                                                            allValues().length
                                                            ? null
                                                            : normalizedIncludedValues
                                                    );
                                                }}
                                            />
                                            <span class="filter-menu__optionLabel">
                                                {optionLabel()}
                                            </span>
                                            <span class="filter-menu__optionCount">
                                                {option.count}
                                            </span>
                                        </label>
                                    );
                                }}
                            </For>
                        </div>
                    </Show>
                </div>
                <div class="filter-menu__footer">
                    <button
                        class="filter-menu__clearButton"
                        disabled={!hasColumnCriteria()}
                        type="button"
                        onClick={() => props.onClearColumn(props.menu!.columnNumber)}
                    >
                        {props.strings.filterClearColumn}
                    </button>
                </div>
            </div>
        </Show>
    );
}

function getEditorStrings(): Pick<
    EditorPanelStrings,
    | "search"
    | "searchFind"
    | "searchReplace"
    | "searchReplaceComingSoon"
    | "filterSelection"
    | "clearFilterRange"
    | "searchScopeLabel"
    | "searchScopeSelection"
    | "searchScopeSelectionDisabled"
    | "searchScopeWholeSheet"
    | "searchClose"
    | "loading"
    | "reload"
    | "save"
    | "readOnly"
    | "moreSheets"
    | "addSheet"
    | "deleteSheet"
    | "renameSheet"
    | "insertRowAbove"
    | "insertRowBelow"
    | "deleteRow"
    | "setRowHeight"
    | "insertColumnLeft"
    | "insertColumnRight"
    | "deleteColumn"
    | "setColumnWidth"
    | "undo"
    | "redo"
    | "searchPlaceholder"
    | "replacePlaceholder"
    | "findPrev"
    | "findNext"
    | "replaceAll"
    | "replaceCount"
    | "replaceNoEditableMatches"
    | "replaceNoChanges"
    | "sortAscending"
    | "sortDescending"
    | "filterSearchPlaceholder"
    | "filterSelectAll"
    | "filterClearColumn"
    | "filterBlankValue"
    | "filterNoValues"
    | "gotoPlaceholder"
    | "goto"
    | "selectedCell"
    | "noCellSelected"
    | "noRowsAvailable"
    | "noSearchMatches"
    | "invalidSearchPattern"
    | "searchMatchFound"
    | "searchMatchFoundInSelection"
    | "searchMatchSummary"
    | "searchRegex"
    | "searchMatchCase"
    | "searchWholeWord"
    | "alignLeft"
    | "alignCenter"
    | "alignRight"
    | "alignTop"
    | "alignMiddle"
    | "alignBottom"
> {
    const strings = (globalThis as Record<string, unknown>).__XLSX_EDITOR_STRINGS__ as
        | Partial<EditorPanelStrings>
        | undefined;
    return {
        search: strings?.search ?? "Search",
        searchFind: strings?.searchFind ?? "Find",
        searchReplace: strings?.searchReplace ?? "Replace",
        searchReplaceComingSoon: strings?.searchReplaceComingSoon ?? "Coming soon",
        filterSelection: strings?.filterSelection ?? "Filter Selection",
        clearFilterRange: strings?.clearFilterRange ?? "Clear Filter",
        searchScopeLabel: strings?.searchScopeLabel ?? "Search scope:",
        searchScopeSelection: strings?.searchScopeSelection ?? "Selected range",
        searchScopeSelectionDisabled:
            strings?.searchScopeSelectionDisabled ?? "Select multiple cells to enable",
        searchScopeWholeSheet: strings?.searchScopeWholeSheet ?? "Whole sheet",
        searchClose: strings?.searchClose ?? "Close",
        loading: strings?.loading ?? "Loading editor...",
        reload: strings?.reload ?? "Reload",
        save: strings?.save ?? "Save",
        readOnly: strings?.readOnly ?? "Read-only",
        moreSheets: strings?.moreSheets ?? "More",
        addSheet: strings?.addSheet ?? "Add Sheet",
        deleteSheet: strings?.deleteSheet ?? "Delete Sheet",
        renameSheet: strings?.renameSheet ?? "Rename Sheet",
        insertRowAbove: strings?.insertRowAbove ?? "Insert Row Above",
        insertRowBelow: strings?.insertRowBelow ?? "Insert Row Below",
        deleteRow: strings?.deleteRow ?? "Delete Row",
        setRowHeight: strings?.setRowHeight ?? "Set Row Height",
        insertColumnLeft: strings?.insertColumnLeft ?? "Insert Column Left",
        insertColumnRight: strings?.insertColumnRight ?? "Insert Column Right",
        deleteColumn: strings?.deleteColumn ?? "Delete Column",
        setColumnWidth: strings?.setColumnWidth ?? "Set Column Width",
        undo: strings?.undo ?? "Undo",
        redo: strings?.redo ?? "Redo",
        searchPlaceholder: strings?.searchPlaceholder ?? "Search values or formulas",
        replacePlaceholder: strings?.replacePlaceholder ?? "Replace in matching cells",
        findPrev: strings?.findPrev ?? "Prev Match",
        findNext: strings?.findNext ?? "Next Match",
        replaceAll: strings?.replaceAll ?? "Replace All",
        replaceCount: strings?.replaceCount ?? "Replaced {count} matching cells.",
        replaceNoEditableMatches:
            strings?.replaceNoEditableMatches ?? "No editable matching cells were found.",
        replaceNoChanges: strings?.replaceNoChanges ?? "No values changed.",
        sortAscending: strings?.sortAscending ?? "Sort A to Z",
        sortDescending: strings?.sortDescending ?? "Sort Z to A",
        filterSearchPlaceholder: strings?.filterSearchPlaceholder ?? "Search values",
        filterSelectAll: strings?.filterSelectAll ?? "Select All",
        filterClearColumn: strings?.filterClearColumn ?? "Clear Column",
        filterBlankValue: strings?.filterBlankValue ?? "(Blanks)",
        filterNoValues: strings?.filterNoValues ?? "No values",
        gotoPlaceholder: strings?.gotoPlaceholder ?? "A1 or Sheet1!B2",
        goto: strings?.goto ?? "Go",
        selectedCell: strings?.selectedCell ?? "Selected Cell",
        noCellSelected: strings?.noCellSelected ?? "No cell selected",
        noRowsAvailable: strings?.noRowsAvailable ?? "No rows available",
        noSearchMatches: strings?.noSearchMatches ?? "No matching cells were found.",
        invalidSearchPattern: strings?.invalidSearchPattern ?? "The search pattern is invalid.",
        searchMatchFound: strings?.searchMatchFound ?? "Found at {address}.",
        searchMatchFoundInSelection:
            strings?.searchMatchFoundInSelection ?? "Found at {address} in selected range {range}.",
        searchMatchSummary: strings?.searchMatchSummary ?? "Match {index} of {count}.",
        searchRegex: strings?.searchRegex ?? "Regex",
        searchMatchCase: strings?.searchMatchCase ?? "Match Case",
        searchWholeWord: strings?.searchWholeWord ?? "Whole Word",
        alignLeft: strings?.alignLeft ?? "Align Left",
        alignCenter: strings?.alignCenter ?? "Align Center",
        alignRight: strings?.alignRight ?? "Align Right",
        alignTop: strings?.alignTop ?? "Align Top",
        alignMiddle: strings?.alignMiddle ?? "Align Middle",
        alignBottom: strings?.alignBottom ?? "Align Bottom",
    };
}

export function EditorBootstrapApp() {
    const vscode = getVsCodeApi();
    const [session, setSession] = createSignal(createInitialEditorSessionState());
    const [viewportState, setViewportState] = createSignal(createInitialEditorGridViewportState());
    const [pendingUndoHistory, setPendingUndoHistory] = createSignal<PendingHistoryEntry[]>([]);
    const [pendingRedoHistory, setPendingRedoHistory] = createSignal<PendingHistoryEntry[]>([]);
    const [viewportElement, setViewportElement] = createSignal<HTMLDivElement | null>(null);
    const [selectionAnchorCell, setSelectionAnchorCell] = createSignal<EditorSelectedCell | null>(
        null
    );
    const [selectionRangeOverride, setSelectionRangeOverride] = createSignal<SelectionRange | null>(
        null
    );
    const [sheetFilterStates, setSheetFilterStates] = createSignal<
        Record<string, EditorSheetFilterState | null>
    >({});
    const [filterMenu, setFilterMenu] = createSignal<FilterMenuState | null>(null);
    const [searchOpen, setSearchOpen] = createSignal(false);
    const [searchPanelPosition, setSearchPanelPosition] = createSignal<SearchPanelPosition | null>(
        null
    );
    const [sheetContextMenu, setSheetContextMenu] = createSignal<SheetContextMenuState | null>(
        null
    );
    const [gridContextMenu, setGridContextMenu] = createSignal<GridContextMenuState | null>(null);
    const [searchMode, setSearchMode] = createSignal<EditorSearchMode>("find");
    const [searchQuery, setSearchQuery] = createSignal("");
    const [replaceQuery, setReplaceQuery] = createSignal("");
    const [gotoReference, setGotoReference] = createSignal("");
    const [isEditingGotoReference, setIsEditingGotoReference] = createSignal(false);
    const [editingCell, setEditingCell] = createSignal<EditorCellEditingState | null>(null);
    const [columnResizeState, setColumnResizeState] = createSignal<EditorColumnResizeState | null>(
        null
    );
    const [rowResizeState, setRowResizeState] = createSignal<EditorRowResizeState | null>(null);
    const [fillDragState, setFillDragState] = createSignal<EditorGridFillDragState | null>(null);
    const [searchOptions, setSearchOptions] = createSignal<SearchOptions>({
        isRegexp: false,
        matchCase: false,
        wholeWord: false,
    });
    const [searchFeedback, setSearchFeedback] = createSignal<SearchFeedbackState | null>(null);
    const strings = getEditorStrings();
    let appElement: HTMLDivElement | undefined;
    let searchPanelElement: HTMLDivElement | undefined;
    let searchPanelDragState: SearchPanelDragState | null = null;
    let selectionDragState: EditorGridSelectionDragState | null = null;
    let previousPendingEditsSnapshot: EditorPendingEdit[] = [];
    let lastEditingDraftSyncKey: string | null = null;

    const createSearchFeedback = (message: EditorSearchResultMessage): SearchFeedbackState => {
        switch (message.status) {
            case "invalid-pattern":
                return {
                    status: "invalid-pattern",
                    tone: "error",
                    message: message.message ?? strings.invalidSearchPattern,
                };
            case "no-match":
                return {
                    status: "no-match",
                    tone: "warn",
                    message: message.message ?? strings.noSearchMatches,
                };
            case "matched": {
                const address = message.match
                    ? getCellAddress(message.match.rowNumber, message.match.columnNumber)
                    : "";
                const summary =
                    typeof message.matchIndex === "number" && typeof message.matchCount === "number"
                        ? formatTemplate(strings.searchMatchSummary, {
                              index: String(message.matchIndex),
                              count: String(message.matchCount),
                          })
                        : "";
                const locationMessage = formatTemplate(
                    message.scope === "selection"
                        ? strings.searchMatchFoundInSelection
                        : strings.searchMatchFound,
                    {
                        address,
                        range: searchScopeSummary(),
                    }
                );
                return {
                    status: "matched",
                    tone: "success",
                    message: [locationMessage, summary].filter(Boolean).join(" "),
                };
            }
        }
    };

    onMount(() => {
        const syncClampedSearchPanelPosition = (panelElementOverride?: HTMLDivElement | null) => {
            const currentPosition = searchPanelPosition();
            const resolvedPanelElement = panelElementOverride ?? searchPanelElement ?? null;
            if (!currentPosition || !appElement || !resolvedPanelElement) {
                return;
            }

            const nextPosition = clampSearchPanelPosition(currentPosition, {
                appElement,
                panelElement: resolvedPanelElement,
            });
            if (
                nextPosition.left !== currentPosition.left ||
                nextPosition.top !== currentPosition.top
            ) {
                setSearchPanelPosition(nextPosition);
            }
        };

        const beginSearchPanelDrag = (pointerId: number, clientX: number, clientY: number) => {
            if (!appElement || !searchPanelElement) {
                return;
            }

            const appRect = appElement.getBoundingClientRect();
            const panelRect = searchPanelElement.getBoundingClientRect();
            const initialPosition = clampSearchPanelPosition(
                {
                    left: panelRect.left - appRect.left,
                    top: panelRect.top - appRect.top,
                },
                {
                    appElement,
                    panelElement: searchPanelElement,
                }
            );

            setSearchPanelPosition(initialPosition);
            searchPanelDragState = {
                pointerId,
                offsetX: clientX - panelRect.left,
                offsetY: clientY - panelRect.top,
            };
            globalThis.getSelection?.()?.removeAllRanges();
        };

        const updateSearchPanelDrag = (pointerId: number, clientX: number, clientY: number) => {
            if (
                !searchPanelDragState ||
                searchPanelDragState.pointerId !== pointerId ||
                !appElement ||
                !searchPanelElement
            ) {
                return;
            }

            const appRect = appElement.getBoundingClientRect();
            setSearchPanelPosition(
                clampSearchPanelPosition(
                    {
                        left: clientX - appRect.left - searchPanelDragState.offsetX,
                        top: clientY - appRect.top - searchPanelDragState.offsetY,
                    },
                    {
                        appElement,
                        panelElement: searchPanelElement,
                    }
                )
            );
        };

        const stopSearchPanelDrag = (pointerId?: number) => {
            if (!searchPanelDragState) {
                return;
            }

            if (pointerId !== undefined && searchPanelDragState.pointerId !== pointerId) {
                return;
            }

            searchPanelDragState = null;
        };

        const stopSelectionDrag = (pointerId?: number) => {
            if (!selectionDragState) {
                return;
            }

            if (pointerId !== undefined && selectionDragState.pointerId !== pointerId) {
                return;
            }

            const focusCell = selection();
            selectionDragState = null;
            if (!focusCell) {
                return;
            }

            postMessage({
                type: "selectCell",
                rowNumber: focusCell.rowNumber,
                columnNumber: focusCell.columnNumber,
            });
        };

        const stopFillDrag = (
            pointerId?: number,
            options: {
                commit?: boolean;
            } = {}
        ) => {
            const currentFillDragState = fillDragState();
            if (!currentFillDragState) {
                return;
            }

            if (pointerId !== undefined && currentFillDragState.pointerId !== pointerId) {
                return;
            }

            setFillDragState(null);

            if (!options.commit || !currentFillDragState.previewRange) {
                return;
            }

            applyFillDragChanges(
                currentFillDragState.sourceRange,
                currentFillDragState.previewRange
            );
        };

        const handleMessage = (
            event: MessageEvent<EditorSessionIncomingMessage | EditorSearchResultMessage>
        ) => {
            const payload = event.data;
            if (isEditorSessionIncomingMessage(payload)) {
                setSession((current) => reduceEditorSessionMessage(current, payload));
                return;
            }

            if (isEditorSearchResultMessage(payload)) {
                setSearchFeedback(createSearchFeedback(payload));
                if (payload.status === "matched" && payload.match) {
                    setGotoReference(
                        getCellAddress(payload.match.rowNumber, payload.match.columnNumber)
                    );
                    revealSearchMatch(payload.match, payload.scope);
                }
            }
        };
        const handlePointerDown = (event: PointerEvent) => {
            const target = event.target;
            if (!(target instanceof Element)) {
                setFilterMenu(null);
                setSheetContextMenu(null);
                setGridContextMenu(null);
                return;
            }

            if (
                filterMenu() &&
                !target.closest('[data-role="filter-menu"]') &&
                !target.closest('[data-role="filter-trigger"]')
            ) {
                setFilterMenu(null);
            }

            if (
                target.closest(".context-menu") ||
                target.closest('[data-role="sheet-tab"]') ||
                target.closest('[data-role="sheet-menu-toggle"]')
            ) {
                return;
            }

            setSheetContextMenu(null);
            setGridContextMenu(null);
        };
        const handlePointerMove = (event: PointerEvent) => {
            const activeColumnResize = columnResizeState();
            if (activeColumnResize?.isDragging && activeColumnResize.pointerId === event.pointerId) {
                if ((event.buttons & 1) === 0) {
                    stopColumnResize(event.pointerId, { commit: true });
                    return;
                }

                event.preventDefault();
                updateColumnResize(event.pointerId, event.clientX);
                return;
            }

            const activeRowResize = rowResizeState();
            if (activeRowResize?.isDragging && activeRowResize.pointerId === event.pointerId) {
                if ((event.buttons & 1) === 0) {
                    stopRowResize(event.pointerId, { commit: true });
                    return;
                }

                event.preventDefault();
                updateRowResize(event.pointerId, event.clientY);
                return;
            }

            updateSearchPanelDrag(event.pointerId, event.clientX, event.clientY);
            const currentFillDragState = fillDragState();
            if (currentFillDragState && currentFillDragState.pointerId === event.pointerId) {
                if ((event.buttons & 1) === 0) {
                    stopFillDrag(event.pointerId, { commit: true });
                    return;
                }

                const target = resolveEditorGridSelectionTarget(event.target);
                if (!target || target.kind !== "cell") {
                    return;
                }

                updateFillDrag(target);
                return;
            }

            if (!selectionDragState || selectionDragState.pointerId !== event.pointerId) {
                return;
            }

            if ((event.buttons & 1) === 0) {
                stopSelectionDrag(event.pointerId);
                return;
            }

            const target = resolveEditorGridSelectionTarget(event.target);
            if (!target) {
                return;
            }

            updateSelectionDrag(target);
        };
        const handlePointerUp = (event: PointerEvent) => {
            stopColumnResize(event.pointerId, { commit: true });
            stopRowResize(event.pointerId, { commit: true });
            stopSearchPanelDrag(event.pointerId);
            stopFillDrag(event.pointerId, { commit: true });
            stopSelectionDrag(event.pointerId);
        };
        const handlePointerCancel = (event: PointerEvent) => {
            stopColumnResize(event.pointerId);
            stopRowResize(event.pointerId);
            stopSearchPanelDrag(event.pointerId);
            stopFillDrag(event.pointerId);
            stopSelectionDrag(event.pointerId);
        };
        const handleResize = () => {
            syncClampedSearchPanelPosition();
        };
        const handleKeyDown = (event: KeyboardEvent) => {
            const normalizedKey = event.key.toLowerCase();
            const hasPrimaryModifier = event.metaKey || event.ctrlKey;
            const editableTextInputTarget = isEditableTextInputTarget(event.target);

            if (event.key === "Escape") {
                if (filterMenu()) {
                    event.preventDefault();
                    setFilterMenu(null);
                    return;
                }

                if (sheetContextMenu() || gridContextMenu()) {
                    event.preventDefault();
                    setSheetContextMenu(null);
                    setGridContextMenu(null);
                    return;
                }

                if (searchOpen() && !editableTextInputTarget) {
                    event.preventDefault();
                    closeSearchPanel();
                    return;
                }

                if (editingCell()) {
                    return;
                }

                setSheetContextMenu(null);
                setGridContextMenu(null);
                return;
            }

            if (!event.altKey && hasPrimaryModifier && !editableTextInputTarget) {
                if (normalizedKey === "f") {
                    event.preventDefault();
                    openSearchPanel("find");
                    return;
                }

                if (normalizedKey === "h") {
                    event.preventDefault();
                    openSearchPanel("replace");
                    return;
                }

                if (normalizedKey === "g" && !editingCell()) {
                    event.preventDefault();
                    focusGotoInput();
                    return;
                }

                if (normalizedKey === "z" && !event.shiftKey) {
                    if (canUndoEdits()) {
                        event.preventDefault();
                        undoPendingEdits();
                    }
                    return;
                }

                const isRedoShortcut =
                    (normalizedKey === "z" && event.shiftKey) ||
                    (normalizedKey === "y" && event.ctrlKey && !event.metaKey && !event.shiftKey);
                if (isRedoShortcut) {
                    if (canRedoEdits()) {
                        event.preventDefault();
                        redoPendingEdits();
                    }
                    return;
                }
            }

            if (editableTextInputTarget || editingCell()) {
                return;
            }

            if (!event.altKey && !event.ctrlKey && !event.metaKey && event.shiftKey) {
                switch (event.key) {
                    case "ArrowUp":
                    case "ArrowDown":
                    case "ArrowLeft":
                    case "ArrowRight":
                        event.preventDefault();
                        navigateSelectionByKey(event.key, { extend: true });
                        return;
                }
            }

            if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
                return;
            }

            if (isEditorClearCellKey(event.key)) {
                event.preventDefault();
                clearSelectedCellValue();
                return;
            }

            const pageDirection = getEditorKeyboardPageDirection(event.key);
            if (pageDirection) {
                event.preventDefault();
                navigateSelectionByViewportPage(pageDirection);
                return;
            }

            if (getEditorKeyboardNavigationDelta(event.key)) {
                event.preventDefault();
                navigateSelectionByKey(event.key);
            }
        };

        window.addEventListener("message", handleMessage);
        window.addEventListener("pointerdown", handlePointerDown);
        window.addEventListener("pointermove", handlePointerMove);
        window.addEventListener("pointerup", handlePointerUp);
        window.addEventListener("pointercancel", handlePointerCancel);
        window.addEventListener("keydown", handleKeyDown);
        window.addEventListener("resize", handleResize);
        vscode.postMessage(createWebviewReadyMessage());

        onCleanup(() => {
            window.removeEventListener("message", handleMessage);
            window.removeEventListener("pointerdown", handlePointerDown);
            window.removeEventListener("pointermove", handlePointerMove);
            window.removeEventListener("pointerup", handlePointerUp);
            window.removeEventListener("pointercancel", handlePointerCancel);
            window.removeEventListener("keydown", handleKeyDown);
            window.removeEventListener("resize", handleResize);
        });
    });

    const workbook = () => session().document.workbook;
    const selection = () => session().ui.selection;
    const activeSheet = () => workbook().activeSheet;
    const activeSheetView = createMemo<EditorActiveSheetView | null>(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return null;
        }

        return {
            ...sheet,
            columns: sheet.columns ?? [],
            cells: sheet.cells ?? {},
        };
    });
    const selectionRange = createMemo(() =>
        resolveEditorSelectionRange(selection(), selectionRangeOverride())
    );
    const expandedSelectionRange = createMemo<SelectionRange | null>(() => {
        const currentSelectionRange = selectionRange();
        return hasExpandedSelectionRange(currentSelectionRange) ? currentSelectionRange : null;
    });
    const fillSourceRange = createMemo<SelectionRange | null>(
        () => fillDragState()?.sourceRange ?? selectionRange()
    );
    const fillPreviewRange = createMemo<SelectionRange | null>(
        () => fillDragState()?.previewRange ?? null
    );
    const selectionSignature = createMemo(() => {
        const currentSelection = selection();
        return currentSelection
            ? `${currentSelection.rowNumber}:${currentSelection.columnNumber}`
            : null;
    });
    let previousGridMetrics: EditorGridMetrics | null = null;
    let previousCellLayers: EditorGridCellLayers | null = null;
    let preserveLocalSelectionRangeOnNextSelectionSync = false;
    let previousActiveSheetKeyForUiReset: string | null | undefined;

    const getCellSignature = (cell: EditorSelectedCell | null): string | null =>
        cell ? `${cell.rowNumber}:${cell.columnNumber}` : null;

    const applyLocalSelectionRangeState = (nextState: EditorSelectionRangeState) => {
        preserveLocalSelectionRangeOnNextSelectionSync =
            getCellSignature(nextState.focusCell) !== selectionSignature();
        setSelectionAnchorCell(nextState.anchorCell);
        setSelectionRangeOverride(nextState.selectionRange);
    };

    const syncViewportFromElement = (element: {
        scrollTop: number;
        scrollLeft: number;
        clientHeight: number;
        clientWidth: number;
    }) => {
        setViewportState((current) =>
            applyEditorGridViewportPatch(current, createEditorGridViewportPatchFromElement(element))
        );
    };

    createEffect(() => {
        const activeEdit = editingCell();
        const sheet = activeSheet();
        const currentSelection = selection();
        if (!activeEdit) {
            return;
        }

        if (
            !sheet ||
            sheet.key !== activeEdit.sheetKey ||
            !currentSelection ||
            currentSelection.rowNumber !== activeEdit.rowNumber ||
            currentSelection.columnNumber !== activeEdit.columnNumber
        ) {
            setEditingCell(null);
        }
    });

    createEffect(() => {
        const activeSheetKey = workbook().activeSheet?.key ?? null;
        if (activeSheetKey === previousActiveSheetKeyForUiReset) {
            return;
        }

        previousActiveSheetKeyForUiReset = activeSheetKey;
        setSearchFeedback(null);
        setFilterMenu(null);
        setSheetContextMenu(null);
        setGridContextMenu(null);
        setFillDragState(null);
    });

    createEffect(() => {
        const activeResize = columnResizeState();
        if (!activeResize) {
            return;
        }

        const sheet = activeSheetView();
        if (!sheet || activeResize.sheetKey !== sheet.key) {
            setColumnResizeState(null);
            return;
        }

        if (activeResize.isDragging) {
            return;
        }

        const currentWidth = normalizeWorkbookColumnWidth(
            sheet.columnWidths?.[activeResize.columnNumber - 1] ?? 0
        );
        if (currentWidth === getPreviewWorkbookColumnWidth(activeResize.previewPixelWidth)) {
            setColumnResizeState(null);
        }
    });

    createEffect(() => {
        const activeResize = rowResizeState();
        if (!activeResize) {
            return;
        }

        const sheet = activeSheetView();
        if (!sheet || activeResize.sheetKey !== sheet.key) {
            setRowResizeState(null);
            return;
        }

        if (activeResize.isDragging) {
            return;
        }

        const currentHeight = normalizeWorkbookRowHeight(
            sheet.rowHeights?.[String(activeResize.rowNumber)] ?? 16
        );
        if (currentHeight === getPreviewWorkbookRowHeight(activeResize.previewPixelHeight)) {
            setRowResizeState(null);
        }
    });

    createEffect(() => {
        const sheetKey = activeSheet()?.key ?? null;
        const element = viewportElement();
        if (!sheetKey || !element) {
            return;
        }

        syncViewportFromElement(element);
        let resizeObserver: ResizeObserver | null = null;
        if (typeof ResizeObserver !== "undefined") {
            resizeObserver = new ResizeObserver(() => {
                syncViewportFromElement(element);
            });
            resizeObserver.observe(element);
        }

        onCleanup(() => {
            resizeObserver?.disconnect();
        });
    });

    createEffect(() => {
        if (!searchOpen()) {
            searchPanelDragState = null;
            searchPanelElement = undefined;
        }
    });

    createEffect(() => {
        const currentPosition = searchPanelPosition();
        if (!searchOpen() || !currentPosition || !appElement || !searchPanelElement) {
            return;
        }

        const nextPosition = clampSearchPanelPosition(currentPosition, {
            appElement,
            panelElement: searchPanelElement,
        });
        if (
            nextPosition.left !== currentPosition.left ||
            nextPosition.top !== currentPosition.top
        ) {
            setSearchPanelPosition(nextPosition);
        }
    });

    createEffect(() => {
        const signature = selectionSignature();
        if (!signature) {
            preserveLocalSelectionRangeOnNextSelectionSync = false;
            setSelectionAnchorCell(null);
            setSelectionRangeOverride(null);
            return;
        }

        if (preserveLocalSelectionRangeOnNextSelectionSync) {
            preserveLocalSelectionRangeOnNextSelectionSync = false;
            return;
        }

        const [rowNumberText, columnNumberText] = signature.split(":");
        const rowNumber = Number(rowNumberText);
        const columnNumber = Number(columnNumberText);

        setSelectionAnchorCell({
            rowNumber,
            columnNumber,
        });
        setSelectionRangeOverride(null);
    });

    const selectionAddressLabel = createMemo(() => {
        const currentSelection = selection();
        if (!currentSelection) {
            return "";
        }

        const currentSelectionRange = expandedSelectionRange();
        if (currentSelectionRange) {
            return `${getCellAddress(currentSelectionRange.startRow, currentSelectionRange.startColumn)}:${getCellAddress(currentSelectionRange.endRow, currentSelectionRange.endColumn)}`;
        }

        return currentSelection.address;
    });

    createEffect(() => {
        const currentAddress = selectionAddressLabel();
        if (!isEditingGotoReference()) {
            setGotoReference(currentAddress);
        }
    });

    createEffect(() => {
        const editingDrafts = session().ui.editingDrafts;
        const pendingEditsSyncKey = editingDrafts.pendingEdits
            .map((edit) => `${edit.sheetKey}:${edit.rowNumber}:${edit.columnNumber}:${edit.value}`)
            .join("|");
        const nextSyncKey = `${editingDrafts.clearRequested ? 1 : 0}:${editingDrafts.preservePendingHistory ? 1 : 0}:${editingDrafts.resetPendingHistory ? 1 : 0}:${pendingEditsSyncKey}`;
        if (nextSyncKey === lastEditingDraftSyncKey) {
            return;
        }

        if (editingDrafts.clearRequested) {
            setSheetFilterStates({});
            setFilterMenu(null);
            if (editingDrafts.preservePendingHistory) {
                const rebasedHistory = rebasePendingHistory(
                    pendingUndoHistory(),
                    pendingRedoHistory(),
                    previousPendingEditsSnapshot
                );
                setPendingUndoHistory(rebasedHistory.undoStack);
                setPendingRedoHistory(rebasedHistory.redoStack);
            } else {
                setPendingUndoHistory([]);
                setPendingRedoHistory([]);
            }
        } else if (editingDrafts.resetPendingHistory) {
            setPendingUndoHistory([]);
            setPendingRedoHistory([]);
        }

        previousPendingEditsSnapshot = editingDrafts.pendingEdits.map((edit) => ({ ...edit }));
        lastEditingDraftSyncKey = nextSyncKey;
    });

    const getModelCellValue = (rowNumber: number, columnNumber: number): string => {
        const sheet = activeSheetView();
        if (!sheet) {
            return "";
        }

        return sheet.cells[createCellKey(rowNumber, columnNumber)]?.displayValue ?? "";
    };

    const getEffectiveCellValue = (rowNumber: number, columnNumber: number): string => {
        const sheet = activeSheetView();
        if (!sheet) {
            return "";
        }

        return (
            getEditorPendingEditValue(
                session().ui.editingDrafts.pendingEdits,
                sheet.key,
                rowNumber,
                columnNumber
            ) ?? getModelCellValue(rowNumber, columnNumber)
        );
    };

    const getActiveSheetPendingEdits = () => {
        const sheet = activeSheetView();
        if (!sheet) {
            return [];
        }

        return session().ui.editingDrafts.pendingEdits
            .filter((edit) => edit.sheetKey === sheet.key)
            .map((edit) => ({
                rowNumber: edit.rowNumber,
                columnNumber: edit.columnNumber,
                value: edit.value,
            }));
    };

    const getVisibleSearchCells = () => {
        const sheet = activeSheetView();
        if (!sheet) {
            return {};
        }

        const filterState = activeSheetFilterState();
        if (!filterState) {
            return sheet.cells;
        }

        const visibleRows = new Set(
            visibleRowResult().visibleRows.filter((rowNumber) => rowNumber <= sheet.rowCount)
        );

        return Object.fromEntries(
            Object.entries(sheet.cells).filter(([, cell]) => visibleRows.has(cell.rowNumber))
        );
    };

    const createFilterCellSource = (): EditorFilterCellSource | null => {
        const sheet = activeSheetView();
        if (!sheet) {
            return null;
        }

        return {
            rowCount: sheet.rowCount,
            columnCount: sheet.columnCount,
            getCellValue: getEffectiveCellValue,
        };
    };

    const getStoredSheetFilterState = (
        sheetKey: string
    ): EditorSheetFilterState | null | undefined => {
        const states = sheetFilterStates();
        return Object.hasOwn(states, sheetKey) ? (states[sheetKey] ?? null) : undefined;
    };

    const activeSheetFilterState = createMemo<EditorSheetFilterState | null>(() => {
        const sheet = activeSheetView();
        const source = createFilterCellSource();
        if (!sheet || !source) {
            return null;
        }

        const storedState =
            getStoredSheetFilterState(sheet.key) ??
            createEditorSheetFilterStateFromSnapshot(sheet.autoFilter ?? null);
        if (!storedState) {
            return null;
        }

        const normalizedRange = normalizeEditorFilterRange(source, storedState.range);
        return normalizedRange
            ? {
                  ...storedState,
                  range: normalizedRange,
              }
            : null;
    });

    const visibleRowResult = createMemo(() => {
        const sheet = activeSheetView();
        const source = createFilterCellSource();
        if (!sheet || !source) {
            return {
                visibleRows: [] as number[],
                hiddenRows: [] as number[],
            };
        }

        return getEditorVisibleRows(source, activeSheetFilterState());
    });

    const selectedCellValue = createMemo(() =>
        getEditorSelectionDisplayValue({
            activeSheetKey: activeSheet()?.key ?? null,
            selection: selection(),
            editingCell: editingCell(),
            pendingEdits: session().ui.editingDrafts.pendingEdits,
        })
    );
    const hasPendingEdits = createMemo(
        () => workbook().hasPendingEdits || session().ui.editingDrafts.pendingEdits.length > 0
    );
    const gridMetrics = createMemo(() => {
        const sheet = activeSheetView();
        if (!sheet) {
            previousGridMetrics = null;
            return null;
        }

        const rows = visibleRowResult();
        const activeColumnResize = columnResizeState();
        const activeRowResize = rowResizeState();
        const previewColumnWidths =
            activeColumnResize?.sheetKey === sheet.key
                ? (() => {
                      const nextColumnWidths = [...(sheet.columnWidths ?? [])];
                      nextColumnWidths[activeColumnResize.columnNumber - 1] =
                          getPreviewWorkbookColumnWidth(activeColumnResize.previewPixelWidth);
                      return nextColumnWidths;
                  })()
                : sheet.columnWidths;
        const previewRowHeights =
            activeRowResize?.sheetKey === sheet.key
                ? {
                      ...(sheet.rowHeights ?? {}),
                      [String(activeRowResize.rowNumber)]: getPreviewWorkbookRowHeight(
                          activeRowResize.previewPixelHeight
                      ),
                  }
                : sheet.rowHeights;
        const nextMetrics = deriveEditorGridMetrics(
            {
                ...createEditorGridMetricsInputFromSheet(sheet),
                rowHeaderLabelCount: sheet.rowCount + EDITOR_EXTRA_PADDING_ROWS,
                visibleRows: rows.visibleRows,
                hiddenRows: rows.hiddenRows,
                columnWidths: previewColumnWidths,
                rowHeights: previewRowHeights,
                maximumDigitWidth: DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
            },
            viewportState()
        );
        previousGridMetrics = reuseEquivalentEditorGridMetrics(previousGridMetrics, nextMetrics);
        return previousGridMetrics;
    });
    const headerLayers = createMemo(() => {
        const metrics = gridMetrics();
        const sheet = activeSheetView();
        if (!metrics || !sheet) {
            return null;
        }

        return deriveEditorGridHeaderLayers({
            metrics,
            columnLabels: sheet.columns,
            selectedRowNumber: selection()?.rowNumber ?? null,
            selectedColumnNumber: selection()?.columnNumber ?? null,
            selectionRange: selectionRange(),
        });
    });
    const cellLayers = createMemo(() => {
        const metrics = gridMetrics();
        const sheet = activeSheetView();
        if (!metrics || !sheet) {
            previousCellLayers = null;
            return null;
        }

        const nextLayers = deriveEditorGridCellLayers({
            metrics,
            activeSheet: sheet,
            pendingEdits: session().ui.editingDrafts.pendingEdits,
            selection: selection(),
            filterState: activeSheetFilterState(),
        });
        previousCellLayers = reuseEquivalentEditorGridCellLayers(previousCellLayers, nextLayers);
        return previousCellLayers;
    });
    const isSearchFocusedSelection = createMemo(() => {
        const feedback = searchFeedback();
        return (
            feedback?.status === "matched" ||
            feedback?.status === "replaced" ||
            feedback?.status === "no-change"
        );
    });
    const selectionOverlayLayers = createMemo(() => {
        const metrics = gridMetrics();
        if (!metrics) {
            return null;
        }

        return deriveEditorSelectionOverlayLayers({
            metrics,
            selection: selection(),
            selectionRangeOverride: selectionRangeOverride(),
            forcePrimaryRect: isSearchFocusedSelection(),
        });
    });
    const hasEditableCellInFillSourceRange = createMemo(() => {
        const range = fillSourceRange();
        if (!range || !workbook().canEdit || editingCell()) {
            return false;
        }

        for (let rowNumber = range.startRow; rowNumber <= range.endRow; rowNumber += 1) {
            for (
                let columnNumber = range.startColumn;
                columnNumber <= range.endColumn;
                columnNumber += 1
            ) {
                if (canEditGridCell(rowNumber, columnNumber)) {
                    return true;
                }
            }
        }

        return false;
    });
    const fillHandleLayout = createMemo<EditorGridFillHandleLayout | null>(() => {
        const metrics = gridMetrics();
        const range = fillSourceRange();
        if (!metrics || !range || !hasEditableCellInFillSourceRange()) {
            return null;
        }

        const handleCell = {
            rowNumber: range.endRow,
            columnNumber: range.endColumn,
        };
        const layer = getEditorGridLayerForCell(metrics, handleCell);
        if (!layer) {
            return null;
        }

        return {
            layer,
            rowNumber: handleCell.rowNumber,
            columnNumber: handleCell.columnNumber,
            left:
                metrics.rowHeaderWidth +
                getEditorColumnLeft(metrics.columnLayout, handleCell.columnNumber) +
                getEditorColumnWidth(metrics.columnLayout, handleCell.columnNumber) -
                FILL_HANDLE_HALF_SIZE,
            top:
                EDITOR_GRID_HEADER_HEIGHT +
                (getEditorGridActualRowTop(metrics, handleCell.rowNumber) ?? 0) +
                (getEditorGridActualRowHeight(metrics, handleCell.rowNumber) ?? 0) -
                FILL_HANDLE_HALF_SIZE,
        };
    });
    const activeToolbarAlignment = createMemo(() =>
        getActiveEditorToolbarAlignment({
            activeSheet: activeSheetView(),
            selection: selection(),
        })
    );
    const canApplyAlignment = createMemo(() => Boolean(workbook().canEdit && selection()));

    const shellCapabilities = createMemo(() => getEditorShellCapabilities(workbook()));
    const canUndoEdits = createMemo(
        () => pendingUndoHistory().length > 0 || shellCapabilities().canUndo
    );
    const canRedoEdits = createMemo(
        () => pendingRedoHistory().length > 0 || shellCapabilities().canRedo
    );
    const sheetContextMenuState = createMemo(() =>
        getEditorSheetContextMenuState({
            canEdit: workbook().canEdit,
            sheetCount: workbook().sheets.length,
        })
    );
    const gridContextMenuStyle = createMemo(() => {
        const menu = gridContextMenu();
        if (!menu) {
            return null;
        }

        const maxHeight = menu.kind === "cell" ? 344 : 168;

        return {
            left: `${Math.max(8, Math.min(menu.x, window.innerWidth - 212))}px`,
            top: `${Math.max(8, Math.min(menu.y, window.innerHeight - maxHeight))}px`,
        };
    });
    const activeSheetTabKey = createMemo(
        () => workbook().sheets.find((item) => item.isActive)?.key ?? null
    );
    const candidateFilterRange = createMemo<SelectionRange | null>(() => {
        const sheet = activeSheetView();
        const source = createFilterCellSource();
        if (!sheet || !source) {
            return null;
        }

        const expandedSelection = selectionRange();
        if (hasExpandedSelectionRange(expandedSelection)) {
            return resolveEditorFilterRangeFromSelection(
                {
                    rowCount: sheet.rowCount,
                    columnCount: sheet.columnCount,
                },
                expandedSelection
            );
        }

        return resolveEditorFilterRangeFromActiveCell(source, selection());
    });
    const hasActiveFilterState = createMemo(() => Boolean(activeSheetFilterState()));
    const canToggleFilter = createMemo(
        () => Boolean(candidateFilterRange()) || hasActiveFilterState()
    );
    const filterActionLabel = createMemo(() => {
        if (candidateFilterRange()) {
            return strings.filterSelection;
        }

        return hasActiveFilterState() ? strings.clearFilterRange : strings.filterSelection;
    });
    const filterMenuOptions = createMemo<readonly EditorFilterValueOption[]>(() => {
        const source = createFilterCellSource();
        const filterState = activeSheetFilterState();
        const menu = filterMenu();
        if (!source || !filterState || !menu) {
            return [];
        }

        return getEditorFilterColumnValues(source, filterState, menu.columnNumber);
    });
    const sheetContextTargetKey = createMemo(() => sheetContextMenu()?.sheetKey ?? null);
    const effectiveSearchScope = createMemo<EditorSearchScope>(() =>
        expandedSelectionRange() ? "selection" : "sheet"
    );
    const activeSearchSelectionRange = createMemo<SelectionRange | null>(() =>
        effectiveSearchScope() === "selection" ? expandedSelectionRange() : null
    );
    const searchScopeSummary = createMemo(() => {
        const range = activeSearchSelectionRange();
        if (!range) {
            return strings.searchScopeWholeSheet;
        }

        return `${getCellAddress(range.startRow, range.startColumn)}:${getCellAddress(range.endRow, range.endColumn)}`;
    });
    const canRunSearch = createMemo(() => searchQuery().trim().length > 0);
    const canRunReplace = createMemo(
        () => Boolean(activeSheetView()) && workbook().canEdit && canRunSearch()
    );
    const searchFeedbackClass = createMemo(() => {
        const feedback = searchFeedback();
        if (!feedback) {
            return "search-strip__feedback";
        }

        return `search-strip__feedback search-strip__feedback--${feedback.tone}`;
    });

    const postMessage = (message: EditorWebviewOutgoingMessage) => {
        vscode.postMessage(message);
    };

    const syncPendingEdits = (nextPendingEdits: readonly EditorPendingEdit[]) => {
        const pendingEdits = [...nextPendingEdits];
        setSession((current) => ({
            ...current,
            ui: {
                ...current.ui,
                editingDrafts: {
                    ...current.ui.editingDrafts,
                    pendingEdits,
                },
            },
        }));
        postMessage({
            type: "setPendingEdits",
            edits: pendingEdits,
        });
    };

    const buildPendingEditsByKey = () =>
        new Map(
            session().ui.editingDrafts.pendingEdits.map((edit) => [
                getEditorPendingEditKey(edit.sheetKey, edit.rowNumber, edit.columnNumber),
                { ...edit },
            ])
        );

    const applyPendingHistoryEntry = (entry: PendingHistoryEntry, direction: "undo" | "redo") => {
        const pendingEditsByKey = buildPendingEditsByKey();
        for (const change of entry.changes) {
            const nextValue = direction === "undo" ? change.beforeValue : change.afterValue;
            const key = getEditorPendingEditKey(
                change.sheetKey,
                change.rowNumber,
                change.columnNumber
            );
            if (nextValue === change.modelValue) {
                pendingEditsByKey.delete(key);
                continue;
            }

            pendingEditsByKey.set(key, {
                sheetKey: change.sheetKey,
                rowNumber: change.rowNumber,
                columnNumber: change.columnNumber,
                value: nextValue,
            });
        }

        syncPendingEdits(Array.from(pendingEditsByKey.values()));
    };

    const applyPendingEditChanges = (
        changes: readonly PendingHistoryChange[],
        options: { recordHistory?: boolean } = {}
    ) => {
        const effectiveChanges = changes.filter(
            (change) => change.beforeValue !== change.afterValue
        );
        if (effectiveChanges.length === 0) {
            return false;
        }

        if (options.recordHistory ?? true) {
            setPendingUndoHistory((current) => [
                ...current,
                {
                    changes: effectiveChanges.map((change) => ({ ...change })),
                },
            ]);
            setPendingRedoHistory([]);
        }

        const pendingEditsByKey = buildPendingEditsByKey();
        for (const change of effectiveChanges) {
            const key = getEditorPendingEditKey(
                change.sheetKey,
                change.rowNumber,
                change.columnNumber
            );
            if (change.afterValue === change.modelValue) {
                pendingEditsByKey.delete(key);
                continue;
            }

            pendingEditsByKey.set(key, {
                sheetKey: change.sheetKey,
                rowNumber: change.rowNumber,
                columnNumber: change.columnNumber,
                value: change.afterValue,
            });
        }

        syncPendingEdits(Array.from(pendingEditsByKey.values()));
        return true;
    };

    const undoPendingEdits = () => {
        const currentUndoHistory = pendingUndoHistory();
        const entry = currentUndoHistory.at(-1);
        if (!entry) {
            if (shellCapabilities().canUndo) {
                postMessage({ type: "undoSheetEdit" });
            }
            return;
        }

        setPendingUndoHistory(currentUndoHistory.slice(0, -1));
        applyPendingHistoryEntry(entry, "undo");
        setPendingRedoHistory((current) => [...current, entry]);
    };

    const redoPendingEdits = () => {
        const currentRedoHistory = pendingRedoHistory();
        const entry = currentRedoHistory.at(-1);
        if (!entry) {
            if (shellCapabilities().canRedo) {
                postMessage({ type: "redoSheetEdit" });
            }
            return;
        }

        setPendingRedoHistory(currentRedoHistory.slice(0, -1));
        applyPendingHistoryEntry(entry, "redo");
        setPendingUndoHistory((current) => [...current, entry]);
    };

    const canEditGridCell = (rowNumber: number, columnNumber: number): boolean => {
        const sheet = activeSheetView();
        const metrics = gridMetrics();
        if (!sheet || !metrics || !workbook().canEdit) {
            return false;
        }

        if (
            getEditorDisplayRowNumber(metrics, rowNumber) === null ||
            columnNumber < 1 ||
            columnNumber > metrics.displayColumnCount
        ) {
            return false;
        }

        return !sheet.cells[createCellKey(rowNumber, columnNumber)]?.formula;
    };

    const applyFillDragChanges = (sourceRange: SelectionRange, previewRange: SelectionRange) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return false;
        }

        const changes = buildFillChanges({
            sheetKey: sheet.key,
            sourceRange,
            previewRange,
            getCellValue: getEffectiveCellValue,
            getModelValue: getModelCellValue,
            canEditCell: canEditGridCell,
        });
        return applyPendingEditChanges(changes);
    };

    const commitEditingCell = () => {
        const activeEdit = editingCell();
        if (!activeEdit) {
            return;
        }

        setEditingCell(null);
        applyPendingEditChanges([
            {
                sheetKey: activeEdit.sheetKey,
                rowNumber: activeEdit.rowNumber,
                columnNumber: activeEdit.columnNumber,
                modelValue: activeEdit.modelValue,
                beforeValue:
                    getEditorPendingEditValue(
                        session().ui.editingDrafts.pendingEdits,
                        activeEdit.sheetKey,
                        activeEdit.rowNumber,
                        activeEdit.columnNumber
                    ) ?? activeEdit.modelValue,
                afterValue: activeEdit.draftValue,
            },
        ]);
    };

    const cancelEditingCell = () => {
        setEditingCell(null);
    };

    const updateEditingDraft = (value: string) => {
        setEditingCell((current) =>
            current
                ? {
                      ...current,
                      draftValue: value,
                  }
                : current
        );
    };

    const navigateSelectionByKey = (key: string, options: { extend?: boolean } = {}) => {
        const sheet = activeSheetView();
        const metrics = gridMetrics();
        if (!sheet || !metrics) {
            return;
        }

        const delta = getEditorKeyboardNavigationDelta(key);
        if (!delta) {
            return;
        }

        const target = getNextEditorKeyboardNavigationTarget({
            activeSheet: sheet,
            selection: selection(),
            delta,
            visibleRowNumbers: metrics.rowState.actualRowNumbers.filter(
                (rowNumber) => rowNumber <= sheet.rowCount
            ),
        });
        if (!target) {
            return;
        }

        selectCell(target.rowNumber, target.columnNumber, {
            reveal: true,
            extend: options.extend,
        });
    };

    const navigateSelectionByViewportPage = (direction: -1 | 1) => {
        const sheet = activeSheetView();
        const metrics = gridMetrics();
        if (!sheet || !metrics) {
            return;
        }

        const target = getNextEditorViewportPageNavigationTarget({
            activeSheet: sheet,
            selection: selection(),
            direction,
            visibleRowCount: getMinimumVisibleEditorRowCount(
                metrics.viewport.viewportHeight,
                metrics.rowLayout
            ),
            visibleRowNumbers: metrics.rowState.actualRowNumbers.filter(
                (rowNumber) => rowNumber <= sheet.rowCount
            ),
        });
        if (!target) {
            return;
        }

        selectCell(target.rowNumber, target.columnNumber, { reveal: true });
    };

    const clearSelectedCellValue = () => {
        const sheet = activeSheetView();
        const currentSelection = selection();
        if (!sheet || !currentSelection) {
            return;
        }

        const editableCell = createEditorCellEditingState({
            activeSheet: sheet,
            rowNumber: currentSelection.rowNumber,
            columnNumber: currentSelection.columnNumber,
            canEdit: workbook().canEdit,
            pendingEdits: session().ui.editingDrafts.pendingEdits,
        });
        if (!editableCell) {
            return;
        }

        applyPendingEditChanges([
            {
                sheetKey: editableCell.sheetKey,
                rowNumber: editableCell.rowNumber,
                columnNumber: editableCell.columnNumber,
                modelValue: editableCell.modelValue,
                beforeValue:
                    getEditorPendingEditValue(
                        session().ui.editingDrafts.pendingEdits,
                        editableCell.sheetKey,
                        editableCell.rowNumber,
                        editableCell.columnNumber
                    ) ?? editableCell.modelValue,
                afterValue: "",
            },
        ]);
    };

    const revealSearchMatch = (
        match: { rowNumber: number; columnNumber: number },
        scope: EditorSearchResultMessage["scope"],
        options: { syncHost?: boolean } = {}
    ) => {
        const focusCell = {
            rowNumber: match.rowNumber,
            columnNumber: match.columnNumber,
        };
        const preservedSelectionRange = scope === "selection" ? selectionRange() : null;
        const nextState: EditorSelectionRangeState = preservedSelectionRange
            ? {
                  focusCell,
                  anchorCell: selectionAnchorCell() ?? selection() ?? focusCell,
                  selectionRange: preservedSelectionRange,
              }
            : createEditorSingleCellSelectionState(focusCell);

        ((options.syncHost ?? false) ? applySelectionState : applySelectionStateLocally)(nextState, {
            reveal: true,
        });
    };

    const applyOptimisticSelection = (rowNumber: number, columnNumber: number) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return;
        }

        const nextSelection = createOptimisticEditorSelection({
            activeSheet: sheet,
            rowNumber,
            columnNumber,
        });
        if (!nextSelection) {
            return;
        }

        setSession((current) => {
            const currentSelection = current.ui.selection;
            if (
                currentSelection &&
                currentSelection.rowNumber === nextSelection.rowNumber &&
                currentSelection.columnNumber === nextSelection.columnNumber
            ) {
                return current;
            }

            return {
                ...current,
                ui: {
                    ...current.ui,
                    selection: nextSelection,
                },
            };
        });
    };

    const applySelectionStateLocally = (
        nextState: EditorSelectionRangeState,
        options: { reveal?: boolean } = {}
    ) => {
        const focusCell = nextState.focusCell;
        if (!focusCell) {
            return;
        }

        const activeEdit = editingCell();
        const sheetKey = activeSheet()?.key ?? null;
        if (
            activeEdit &&
            (activeEdit.sheetKey !== sheetKey ||
                activeEdit.rowNumber !== focusCell.rowNumber ||
                activeEdit.columnNumber !== focusCell.columnNumber)
        ) {
            commitEditingCell();
        }

        applyLocalSelectionRangeState(nextState);
        applyOptimisticSelection(focusCell.rowNumber, focusCell.columnNumber);
        if (options.reveal) {
            const metrics = gridMetrics();
            const element = viewportElement();
            if (metrics && element) {
                revealEditorGridCellInViewport(
                    element,
                    metrics,
                    focusCell.rowNumber,
                    focusCell.columnNumber
                );
            }
        }
    };

    const applySelectionState = (
        nextState: EditorSelectionRangeState,
        options: { reveal?: boolean } = {}
    ) => {
        const focusCell = nextState.focusCell;
        if (!focusCell) {
            return;
        }

        applySelectionStateLocally(nextState, options);
        postMessage({
            type: "selectCell",
            rowNumber: focusCell.rowNumber,
            columnNumber: focusCell.columnNumber,
        });
    };

    const beginFillDrag = (pointerId: number, sourceRange: SelectionRange) => {
        selectionDragState = null;
        setFillDragState({
            pointerId,
            sourceRange,
            previewRange: null,
        });
        globalThis.getSelection?.()?.removeAllRanges();
    };

    const updateFillDrag = (target: Extract<EditorGridSelectionTarget, { kind: "cell" }>) => {
        const currentFillDragState = fillDragState();
        const bounds = getEditorGridFillBounds(gridMetrics());
        if (!currentFillDragState || !bounds) {
            return;
        }

        const nextPreviewRange = createFillPreviewRange(
            currentFillDragState.sourceRange,
            {
                rowNumber: target.rowNumber,
                columnNumber: target.columnNumber,
            },
            bounds
        );
        if (areSelectionRangesEqual(currentFillDragState.previewRange, nextPreviewRange)) {
            return;
        }

        setFillDragState({
            ...currentFillDragState,
            previewRange: nextPreviewRange,
        });
    };

    const startEditingCell = (rowNumber: number, columnNumber: number) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return;
        }

        const activeEdit = editingCell();
        if (
            activeEdit &&
            (activeEdit.sheetKey !== sheet.key ||
                activeEdit.rowNumber !== rowNumber ||
                activeEdit.columnNumber !== columnNumber)
        ) {
            commitEditingCell();
        }

        const nextEdit = createEditorCellEditingState({
            activeSheet: sheet,
            rowNumber,
            columnNumber,
            canEdit: workbook().canEdit,
            pendingEdits: session().ui.editingDrafts.pendingEdits,
        });
        if (!nextEdit) {
            return;
        }

        applyOptimisticSelection(rowNumber, columnNumber);
        setEditingCell(nextEdit);
    };

    const selectCell = (
        rowNumber: number,
        columnNumber: number,
        options: { reveal?: boolean; extend?: boolean; syncHost?: boolean } = {}
    ) => {
        const nextState = options.extend
            ? createEditorExtendedCellSelectionState({
                  anchorCell: selectionAnchorCell(),
                  focusCell: { rowNumber, columnNumber },
              })
            : createEditorSingleCellSelectionState({ rowNumber, columnNumber });

        ((options.syncHost ?? true) ? applySelectionState : applySelectionStateLocally)(nextState, {
            reveal: options.reveal,
        });
    };

    const beginCellSelectionDrag = (
        rowNumber: number,
        columnNumber: number,
        pointerId: number,
        options: { extend?: boolean } = {}
    ) => {
        const anchorCell =
            options.extend && selectionAnchorCell()
                ? selectionAnchorCell()!
                : { rowNumber, columnNumber };
        selectionDragState = {
            kind: "cell",
            pointerId,
            anchorCell,
            lastTargetKey: getEditorGridSelectionTargetKey({
                kind: "cell",
                rowNumber,
                columnNumber,
            }),
        };
        selectCell(rowNumber, columnNumber, {
            extend: options.extend,
            syncHost: false,
        });
    };

    const updateSelectionDrag = (target: EditorGridSelectionTarget) => {
        const currentDragState = selectionDragState;
        const sheet = activeSheetView();
        if (!currentDragState || !sheet || currentDragState.kind !== target.kind) {
            return;
        }

        const nextTargetKey = getEditorGridSelectionTargetKey(target);
        if (nextTargetKey === currentDragState.lastTargetKey) {
            return;
        }

        selectionDragState = {
            ...currentDragState,
            lastTargetKey: nextTargetKey,
        };

        switch (target.kind) {
            case "cell":
                applySelectionStateLocally(
                    createEditorAnchoredRangeSelectionState({
                        anchorCell: currentDragState.anchorCell,
                        previewCell: {
                            rowNumber: target.rowNumber,
                            columnNumber: target.columnNumber,
                        },
                    })
                );
                return;
            case "row":
                applySelectionStateLocally(
                    createEditorRowSelectionState({
                        anchorCell: currentDragState.anchorCell,
                        focusCell: createEditorRowHeaderSelection(
                            target.rowNumber,
                            currentDragState.anchorCell.columnNumber
                        ),
                        columnCount: sheet.columnCount,
                        extend: true,
                    })
                );
                return;
            case "column":
                applySelectionStateLocally(
                    createEditorColumnSelectionState({
                        anchorCell: currentDragState.anchorCell,
                        focusCell: createEditorColumnHeaderSelection(
                            target.columnNumber,
                            currentDragState.anchorCell.rowNumber
                        ),
                        rowCount: sheet.rowCount,
                        extend: true,
                    })
                );
        }
    };

    const beginRowSelectionDrag = (
        rowNumber: number,
        pointerId: number,
        options: { extend?: boolean } = {}
    ) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return;
        }

        const anchorCell =
            options.extend && selectionAnchorCell()
                ? selectionAnchorCell()!
                : createEditorRowHeaderSelection(
                      rowNumber,
                      selection()?.columnNumber ?? selectionAnchorCell()?.columnNumber ?? null
                  );
        selectionDragState = {
            kind: "row",
            pointerId,
            anchorCell,
            lastTargetKey: getEditorGridSelectionTargetKey({
                kind: "row",
                rowNumber,
            }),
        };
        applySelectionStateLocally(
            createEditorRowSelectionState({
                anchorCell,
                focusCell: createEditorRowHeaderSelection(rowNumber, anchorCell.columnNumber),
                columnCount: sheet.columnCount,
                extend: options.extend ?? false,
            })
        );
    };

    const beginColumnSelectionDrag = (
        columnNumber: number,
        pointerId: number,
        options: { extend?: boolean } = {}
    ) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return;
        }

        const anchorCell =
            options.extend && selectionAnchorCell()
                ? selectionAnchorCell()!
                : createEditorColumnHeaderSelection(
                      columnNumber,
                      selection()?.rowNumber ?? selectionAnchorCell()?.rowNumber ?? null
                  );
        selectionDragState = {
            kind: "column",
            pointerId,
            anchorCell,
            lastTargetKey: getEditorGridSelectionTargetKey({
                kind: "column",
                columnNumber,
            }),
        };
        applySelectionStateLocally(
            createEditorColumnSelectionState({
                anchorCell,
                focusCell: createEditorColumnHeaderSelection(columnNumber, anchorCell.rowNumber),
                rowCount: sheet.rowCount,
                extend: options.extend ?? false,
            })
        );
    };

    const toggleSearchOption = (key: keyof SearchOptions) => {
        setSearchOptions((current) => ({
            ...current,
            [key]: !current[key],
        }));
        setSearchFeedback(null);
    };

    const submitSearch = (direction: EditorSearchDirection) => {
        const message = createEditorSearchMessage(searchQuery(), direction, searchOptions(), {
            scope: effectiveSearchScope(),
            selectionRange: activeSearchSelectionRange(),
        });
        if (message) {
            postMessage(message);
        }
    };

    const openSearchPanel = (mode: EditorSearchMode = "find") => {
        if (!searchOpen()) {
            setSearchOpen(true);
            setSearchFeedback(null);
        }

        if (searchMode() !== mode) {
            setSearchMode(mode);
        }

        queueMicrotask(() => {
            if (mode === "replace" && searchQuery().trim().length > 0) {
                focusReplaceInput();
                return;
            }

            focusSearchInput();
        });
    };

    const closeSearchPanel = () => {
        if (!searchOpen()) {
            return;
        }

        searchPanelDragState = null;
        setSearchOpen(false);
        setSearchFeedback(null);
    };

    const getPreviewWorkbookColumnWidth = (pixelWidth: number) =>
        normalizeWorkbookColumnWidth(
            convertPixelsToWorkbookColumnWidth(pixelWidth, DEFAULT_MAXIMUM_DIGIT_WIDTH_PX)
        );

    const getPreviewWorkbookRowHeight = (pixelHeight: number) =>
        normalizeWorkbookRowHeight(convertPixelsToWorkbookRowHeight(pixelHeight));

    const beginColumnResize = (
        pointerId: number,
        columnNumber: number,
        startPixelWidth: number,
        clientX: number
    ) => {
        const sheet = activeSheetView();
        if (!sheet || !workbook().canEdit) {
            return;
        }

        commitEditingCell();
        setFilterMenu(null);
        setSheetContextMenu(null);
        setGridContextMenu(null);
        setFillDragState(null);
        selectionDragState = null;
        globalThis.getSelection?.()?.removeAllRanges();
        setColumnResizeState({
            pointerId,
            sheetKey: sheet.key,
            columnNumber,
            startClientX: clientX,
            startPixelWidth,
            previewPixelWidth: startPixelWidth,
            isDragging: true,
        });
    };

    const updateColumnResize = (pointerId: number, clientX: number) => {
        const activeResize = columnResizeState();
        if (!activeResize?.isDragging || activeResize.pointerId !== pointerId) {
            return;
        }

        const nextPixelWidth = Math.max(
            MIN_COLUMN_PIXEL_WIDTH,
            Math.min(
                MAX_COLUMN_PIXEL_WIDTH,
                stabilizeColumnPixelWidth(
                    Math.round(activeResize.startPixelWidth + clientX - activeResize.startClientX),
                    DEFAULT_MAXIMUM_DIGIT_WIDTH_PX
                )
            )
        );
        if (nextPixelWidth === activeResize.previewPixelWidth) {
            return;
        }

        setColumnResizeState({
            ...activeResize,
            previewPixelWidth: nextPixelWidth,
        });
    };

    const stopColumnResize = (pointerId?: number, options: { commit?: boolean } = {}) => {
        const activeResize = columnResizeState();
        if (!activeResize) {
            return;
        }

        if (pointerId !== undefined && activeResize.pointerId !== pointerId) {
            return;
        }

        const didChange = activeResize.previewPixelWidth !== activeResize.startPixelWidth;
        if (!(options.commit ?? false) || !didChange) {
            setColumnResizeState(null);
            return;
        }

        setColumnResizeState({
            ...activeResize,
            startPixelWidth: activeResize.previewPixelWidth,
            isDragging: false,
        });
        postMessage({
            type: "setColumnWidth",
            columnNumber: activeResize.columnNumber,
            width: getPreviewWorkbookColumnWidth(activeResize.previewPixelWidth),
        });
    };

    const beginRowResize = (
        pointerId: number,
        rowNumber: number,
        startPixelHeight: number,
        clientY: number
    ) => {
        const sheet = activeSheetView();
        if (!sheet || !workbook().canEdit) {
            return;
        }

        commitEditingCell();
        setFilterMenu(null);
        setSheetContextMenu(null);
        setGridContextMenu(null);
        setFillDragState(null);
        selectionDragState = null;
        globalThis.getSelection?.()?.removeAllRanges();
        setRowResizeState({
            pointerId,
            sheetKey: sheet.key,
            rowNumber,
            startClientY: clientY,
            startPixelHeight,
            previewPixelHeight: startPixelHeight,
            isDragging: true,
        });
    };

    const updateRowResize = (pointerId: number, clientY: number) => {
        const activeResize = rowResizeState();
        if (!activeResize?.isDragging || activeResize.pointerId !== pointerId) {
            return;
        }

        const nextPixelHeight = Math.max(
            MIN_ROW_PIXEL_HEIGHT,
            Math.min(
                MAX_ROW_PIXEL_HEIGHT,
                stabilizeRowPixelHeight(
                    Math.round(activeResize.startPixelHeight + clientY - activeResize.startClientY)
                )
            )
        );
        if (nextPixelHeight === activeResize.previewPixelHeight) {
            return;
        }

        setRowResizeState({
            ...activeResize,
            previewPixelHeight: nextPixelHeight,
        });
    };

    const stopRowResize = (pointerId?: number, options: { commit?: boolean } = {}) => {
        const activeResize = rowResizeState();
        if (!activeResize) {
            return;
        }

        if (pointerId !== undefined && activeResize.pointerId !== pointerId) {
            return;
        }

        const didChange = activeResize.previewPixelHeight !== activeResize.startPixelHeight;
        if (!(options.commit ?? false) || !didChange) {
            setRowResizeState(null);
            return;
        }

        setRowResizeState({
            ...activeResize,
            startPixelHeight: activeResize.previewPixelHeight,
            isDragging: false,
        });
        postMessage({
            type: "setRowHeight",
            rowNumber: activeResize.rowNumber,
            height: getPreviewWorkbookRowHeight(activeResize.previewPixelHeight),
        });
    };

    const submitReplace = (mode: "single" | "all") => {
        const normalizedQuery = searchQuery().trim();
        const sheet = activeSheetView();
        if (!normalizedQuery) {
            focusSearchInput();
            return;
        }

        if (!sheet || !workbook().canEdit) {
            return;
        }

        const result = resolveEditorReplaceResultInSheet(
            {
                key: sheet.key,
                rowCount: sheet.rowCount,
                columnCount: sheet.columnCount,
                cells: getVisibleSearchCells(),
            },
            selection(),
            {
                query: normalizedQuery,
                replacement: replaceQuery(),
                options: searchOptions(),
                scope: effectiveSearchScope(),
                selectionRange: activeSearchSelectionRange() ?? undefined,
                pendingEdits: getActiveSheetPendingEdits(),
                mode,
            }
        );

        if (result.status === "invalid-pattern") {
            setSearchFeedback({
                status: "invalid-pattern",
                tone: "error",
                message: strings.invalidSearchPattern,
            });
            return;
        }

        if (result.status === "no-match") {
            setSearchFeedback({
                status: "no-match",
                tone: "warn",
                message: strings.replaceNoEditableMatches,
            });
            return;
        }

        if (result.status === "no-change") {
            setSearchFeedback({
                status: "no-change",
                tone: "warn",
                message: strings.replaceNoChanges,
            });
            if (result.match) {
                revealSearchMatch(result.match, effectiveSearchScope(), { syncHost: true });
            }
            return;
        }

        const changes = (result.changes ?? []).map((change) => ({
            sheetKey: sheet.key,
            rowNumber: change.rowNumber,
            columnNumber: change.columnNumber,
            modelValue: getModelCellValue(change.rowNumber, change.columnNumber),
            beforeValue: change.beforeValue,
            afterValue: change.afterValue,
        }));
        applyPendingEditChanges(changes);
        setSearchFeedback({
            status: "replaced",
            tone: "success",
            message: formatTemplate(strings.replaceCount, {
                count: String(result.replacedCellCount ?? changes.length),
            }),
        });
        if (result.match) {
            revealSearchMatch(result.nextMatch ?? result.match, effectiveSearchScope(), {
                syncHost: true,
            });
        }
    };

    const submitGoto = () => {
        const message = createEditorGotoMessage(gotoReference());
        if (message) {
            postMessage(message);
        }
    };

    const resetGotoReference = () => {
        setIsEditingGotoReference(false);
        setGotoReference(selectionAddressLabel());
    };

    const setActiveSheetLocalFilterState = (
        sheetKey: string,
        filterState: EditorSheetFilterState | null
    ) => {
        setSheetFilterStates((current) => ({
            ...current,
            [sheetKey]: filterState,
        }));
        if (!filterState && filterMenu()?.sheetKey === sheetKey) {
            setFilterMenu(null);
        }
    };

    const syncFilterStateToHost = (
        sheetKey: string,
        filterState: EditorSheetFilterState | null
    ) => {
        postMessage({
            type: "setFilterState",
            sheetKey,
            filterState: createEditorSheetFilterSnapshot(filterState),
        });
    };

    const getNearestVisibleRowNumber = (
        visibleRows: readonly number[],
        rowNumber: number
    ): number | null => {
        if (visibleRows.length === 0) {
            return null;
        }

        if (visibleRows.includes(rowNumber)) {
            return rowNumber;
        }

        return (
            visibleRows.find((visibleRowNumber) => visibleRowNumber >= rowNumber) ??
            visibleRows[visibleRows.length - 1] ??
            null
        );
    };

    const ensureSelectedCellVisibleForFilter = (
        filterState: EditorSheetFilterState | null
    ): boolean => {
        const source = createFilterCellSource();
        const currentSelection = selection();
        const sheet = activeSheetView();
        if (!source || !currentSelection || !sheet || currentSelection.rowNumber > sheet.rowCount) {
            return false;
        }

        const nextVisibleRowNumber = getNearestVisibleRowNumber(
            getEditorVisibleRows(source, filterState).visibleRows,
            currentSelection.rowNumber
        );
        if (nextVisibleRowNumber === null || nextVisibleRowNumber === currentSelection.rowNumber) {
            return false;
        }

        selectCell(nextVisibleRowNumber, currentSelection.columnNumber, {
            reveal: true,
        });
        return true;
    };

    const submitFilterToggle = () => {
        const sheet = activeSheetView();
        const source = createFilterCellSource();
        if (!sheet || !source) {
            return;
        }

        const currentFilterState = activeSheetFilterState();
        const nextFilterState = toggleEditorSheetFilterState(
            source,
            currentFilterState,
            candidateFilterRange()
        );
        if (!nextFilterState && !currentFilterState) {
            return;
        }

        setActiveSheetLocalFilterState(sheet.key, nextFilterState);
        syncFilterStateToHost(sheet.key, nextFilterState);
        setFilterMenu(null);
        ensureSelectedCellVisibleForFilter(nextFilterState);
    };

    const openFilterMenu = (
        rowNumber: number,
        columnNumber: number,
        element: HTMLButtonElement
    ) => {
        const sheet = activeSheetView();
        const filterState = activeSheetFilterState();
        if (
            !sheet ||
            !filterState ||
            !isEditorFilterHeaderCell(filterState, rowNumber, columnNumber)
        ) {
            return;
        }

        const currentMenu = filterMenu();
        if (currentMenu?.sheetKey === sheet.key && currentMenu.columnNumber === columnNumber) {
            setFilterMenu(null);
            return;
        }

        const appRect = appElement?.getBoundingClientRect();
        const triggerRect = element.getBoundingClientRect();
        const rawLeft = appRect ? triggerRect.right - appRect.left - 260 : triggerRect.left;
        const rawTop = appRect ? triggerRect.bottom - appRect.top + 6 : triggerRect.bottom + 6;
        const appWidth = appElement?.clientWidth ?? window.innerWidth;
        const appHeight = appElement?.clientHeight ?? window.innerHeight;

        setFilterMenu({
            sheetKey: sheet.key,
            columnNumber,
            left: Math.max(8, Math.min(rawLeft, appWidth - 288)),
            top: Math.max(8, Math.min(rawTop, appHeight - 360)),
        });
        closeSearchPanel();
        setSheetContextMenu(null);
    };

    const applyActiveSheetFilterSort = (
        columnNumber: number,
        direction: EditorFilterSortDirection
    ) => {
        const sheet = activeSheetView();
        const filterState = activeSheetFilterState();
        if (!sheet || !filterState) {
            return;
        }

        const nextFilterState = updateEditorFilterSort(filterState, columnNumber, direction);
        setActiveSheetLocalFilterState(sheet.key, nextFilterState);
        syncFilterStateToHost(sheet.key, nextFilterState);
        setFilterMenu(null);
        ensureSelectedCellVisibleForFilter(nextFilterState);
    };

    const setActiveSheetFilterIncludedValues = (
        columnNumber: number,
        includedValues: readonly string[] | null
    ) => {
        const sheet = activeSheetView();
        const filterState = activeSheetFilterState();
        if (!sheet || !filterState) {
            return;
        }

        const nextFilterState = updateEditorFilterIncludedValues(
            filterState,
            columnNumber,
            includedValues
        );
        setActiveSheetLocalFilterState(sheet.key, nextFilterState);
        syncFilterStateToHost(sheet.key, nextFilterState);
        ensureSelectedCellVisibleForFilter(nextFilterState);
    };

    const clearActiveSheetFilterColumn = (columnNumber: number) => {
        const sheet = activeSheetView();
        const filterState = activeSheetFilterState();
        if (!sheet || !filterState) {
            return;
        }

        const nextFilterState = clearEditorFilterColumn(filterState, columnNumber);
        setActiveSheetLocalFilterState(sheet.key, nextFilterState);
        syncFilterStateToHost(sheet.key, nextFilterState);
        setFilterMenu(null);
        ensureSelectedCellVisibleForFilter(nextFilterState);
    };

    const openCellContextMenu = (
        rowNumber: number,
        columnNumber: number,
        x: number,
        y: number
    ) => {
        const sheet = activeSheetView();
        if (!sheet) {
            return;
        }

        if (!isCellWithinSelectionRange(selectionRange(), rowNumber, columnNumber)) {
            selectCell(rowNumber, columnNumber, {
                syncHost: true,
            });
        }

        setGridContextMenu({
            kind: "cell",
            rowNumber,
            columnNumber,
            x,
            y,
        });
        setSheetContextMenu(null);
        setFilterMenu(null);
        closeSearchPanel();
    };

    const openRowContextMenu = (rowNumber: number, x: number, y: number) => {
        const sheet = activeSheetView();
        if (!sheet || !workbook().canEdit) {
            return;
        }

        const focusCell = createEditorRowHeaderSelection(
            rowNumber,
            selection()?.columnNumber ?? selectionAnchorCell()?.columnNumber ?? null
        );
        applySelectionState(
            createEditorRowSelectionState({
                anchorCell: focusCell,
                focusCell,
                columnCount: sheet.columnCount,
                extend: false,
            })
        );
        setGridContextMenu({
            kind: "row",
            rowNumber,
            x,
            y,
        });
        setSheetContextMenu(null);
        setFilterMenu(null);
        closeSearchPanel();
    };

    const openColumnContextMenu = (columnNumber: number, x: number, y: number) => {
        const sheet = activeSheetView();
        if (!sheet || !workbook().canEdit) {
            return;
        }

        const focusCell = createEditorColumnHeaderSelection(
            columnNumber,
            selection()?.rowNumber ?? selectionAnchorCell()?.rowNumber ?? null
        );
        applySelectionState(
            createEditorColumnSelectionState({
                anchorCell: focusCell,
                focusCell,
                rowCount: sheet.rowCount,
                extend: false,
            })
        );
        setGridContextMenu({
            kind: "column",
            columnNumber,
            x,
            y,
        });
        setSheetContextMenu(null);
        setFilterMenu(null);
        closeSearchPanel();
    };

    const requestInsertRow = (rowNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "insertRow", rowNumber });
    };

    const requestDeleteRow = (rowNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "deleteRow", rowNumber });
    };

    const requestPromptRowHeight = (rowNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "promptRowHeight", rowNumber });
    };

    const requestInsertColumn = (columnNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "insertColumn", columnNumber });
    };

    const requestDeleteColumn = (columnNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "deleteColumn", columnNumber });
    };

    const requestPromptColumnWidth = (columnNumber: number) => {
        setGridContextMenu(null);
        postMessage({ type: "promptColumnWidth", columnNumber });
    };

    const openSheetContextMenu = (sheetKey: string, x: number, y: number) => {
        const safeX = Math.max(8, Math.min(x, window.innerWidth - 196));
        const safeY = Math.max(8, Math.min(y, window.innerHeight - 152));
        setSheetContextMenu({
            sheetKey,
            x: safeX,
            y: safeY,
        });
        setFilterMenu(null);
        setGridContextMenu(null);
        closeSearchPanel();
    };

    const applyToolbarAlignment = (alignment: EditorAlignmentPatch) => {
        const target = getActiveEditorAlignmentSelectionTarget({
            activeSheet: activeSheetView(),
            selection: selection(),
            selectionRange: selectionRange(),
        });
        if (!target) {
            return;
        }

        commitEditingCell();
        postMessage({
            type: "setAlignment",
            target: target.target,
            selection: target.selection,
            alignment,
        });
    };

    const panelMessage = () => session().ui.panel.statusMessage;

    return (
        <div
            class="app app--editor"
            ref={(element) => {
                appElement = element;
            }}
        >
            <div class="toolbar toolbar--editor">
                <div class="toolbar__group toolbar__group--grow">
                    <label class="toolbar__field toolbar__field--address">
                        <span class="toolbar__field-label">#</span>
                        <input
                            class="toolbar__input"
                            data-role="goto-input"
                            type="text"
                            value={gotoReference()}
                            placeholder={strings.gotoPlaceholder}
                            onFocus={(event) => {
                                setIsEditingGotoReference(true);
                                event.currentTarget.select();
                            }}
                            onBlur={() => resetGotoReference()}
                            onInput={(event) => setGotoReference(event.currentTarget.value)}
                            onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                    event.preventDefault();
                                    setIsEditingGotoReference(false);
                                    submitGoto();
                                    return;
                                }

                                if (event.key === "Escape") {
                                    event.preventDefault();
                                    resetGotoReference();
                                }
                            }}
                        />
                    </label>
                    <label class="toolbar__field toolbar__field--cell-value">
                        <span class="toolbar__field-label">T</span>
                        <input
                            class="toolbar__input"
                            data-role="selected-cell-value"
                            readonly
                            value={selectedCellValue()}
                            placeholder={strings.noCellSelected}
                        />
                    </label>
                </div>
                <div class="toolbar__group">
                    <Show when={shellCapabilities().isReadOnly}>
                        <span class="badge badge--warn" data-role="read-only-badge">
                            {strings.readOnly}
                        </span>
                    </Show>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        classList={{
                            "is-active": hasActiveFilterState(),
                        }}
                        data-role="filter-toggle"
                        type="button"
                        title={filterActionLabel()}
                        disabled={!canToggleFilter()}
                        onClick={() => {
                            submitFilterToggle();
                            closeSearchPanel();
                        }}
                    >
                        <span class="toolbar__button-icon codicon codicon-filter" />
                    </button>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        classList={{ "is-active": searchOpen() }}
                        data-role="search-toggle"
                        type="button"
                        title={strings.search}
                        onClick={() => {
                            if (searchOpen()) {
                                closeSearchPanel();
                                return;
                            }

                            openSearchPanel(searchMode());
                        }}
                    >
                        <span class="toolbar__button-icon codicon codicon-search" />
                    </button>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        data-role="undo-button"
                        type="button"
                        title={strings.undo}
                        disabled={!canUndoEdits()}
                        onClick={undoPendingEdits}
                    >
                        <span class="toolbar__button-icon toolbar__button-icon--flip codicon codicon-redo" />
                    </button>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        data-role="redo-button"
                        type="button"
                        title={strings.redo}
                        disabled={!canRedoEdits()}
                        onClick={redoPendingEdits}
                    >
                        <span class="toolbar__button-icon codicon codicon-redo" />
                    </button>
                    <div class="toolbar__segmented" role="group" aria-label="Horizontal alignment">
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().horizontal === "left",
                            }}
                            data-role="align-left"
                            type="button"
                            title={strings.alignLeft}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ horizontal: "left" })}
                        >
                            L
                        </button>
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().horizontal === "center",
                            }}
                            data-role="align-center"
                            type="button"
                            title={strings.alignCenter}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ horizontal: "center" })}
                        >
                            C
                        </button>
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().horizontal === "right",
                            }}
                            data-role="align-right"
                            type="button"
                            title={strings.alignRight}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ horizontal: "right" })}
                        >
                            R
                        </button>
                    </div>
                    <div class="toolbar__segmented" role="group" aria-label="Vertical alignment">
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().vertical === "top",
                            }}
                            data-role="align-top"
                            type="button"
                            title={strings.alignTop}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ vertical: "top" })}
                        >
                            T
                        </button>
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().vertical === "center",
                            }}
                            data-role="align-middle"
                            type="button"
                            title={strings.alignMiddle}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ vertical: "center" })}
                        >
                            M
                        </button>
                        <button
                            class="toolbar__toggle"
                            classList={{
                                "is-active": activeToolbarAlignment().vertical === "bottom",
                            }}
                            data-role="align-bottom"
                            type="button"
                            title={strings.alignBottom}
                            disabled={!canApplyAlignment()}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyToolbarAlignment({ vertical: "bottom" })}
                        >
                            B
                        </button>
                    </div>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        data-role="reload-button"
                        type="button"
                        title={strings.reload}
                        onClick={() => postMessage({ type: "reload" })}
                    >
                        <span class="toolbar__button-icon codicon codicon-refresh" />
                    </button>
                    <button
                        class="toolbar__button toolbar__button--icon"
                        classList={{ "is-dirty": hasPendingEdits() }}
                        data-role="save-button"
                        type="button"
                        title={strings.save}
                        disabled={shellCapabilities().isReadOnly || !hasPendingEdits()}
                        onClick={() => {
                            commitEditingCell();
                            postMessage({ type: "requestSave" });
                        }}
                    >
                        <span class="toolbar__button-icon codicon codicon-save" />
                    </button>
                </div>
            </div>

            <EditorFilterMenuView
                strings={strings}
                menu={filterMenu()}
                filterState={activeSheetFilterState()}
                options={filterMenuOptions()}
                onSort={applyActiveSheetFilterSort}
                onSetIncludedValues={setActiveSheetFilterIncludedValues}
                onClearColumn={clearActiveSheetFilterColumn}
            />

            <Show when={searchOpen()}>
                <div
                    class="search-strip-shell"
                    data-role="search-strip"
                    ref={(element) => {
                        searchPanelElement = element;
                        const currentPosition = searchPanelPosition();
                        if (!currentPosition || !appElement) {
                            return;
                        }

                        const nextPosition = clampSearchPanelPosition(currentPosition, {
                            appElement,
                            panelElement: element,
                        });
                        if (
                            nextPosition.left !== currentPosition.left ||
                            nextPosition.top !== currentPosition.top
                        ) {
                            setSearchPanelPosition(nextPosition);
                        }
                    }}
                    style={
                        searchPanelPosition()
                            ? {
                                  left: `${searchPanelPosition()!.left}px`,
                                  top: `${searchPanelPosition()!.top}px`,
                                  right: "auto",
                              }
                            : undefined
                    }
                    onPointerDown={(event) => {
                        if (event.button !== 0 || isSearchPanelInteractiveTarget(event.target)) {
                            return;
                        }

                        event.preventDefault();
                        if (!appElement || !searchPanelElement) {
                            return;
                        }

                        const appRect = appElement.getBoundingClientRect();
                        const panelRect = searchPanelElement.getBoundingClientRect();
                        const initialPosition = clampSearchPanelPosition(
                            {
                                left: panelRect.left - appRect.left,
                                top: panelRect.top - appRect.top,
                            },
                            {
                                appElement,
                                panelElement: searchPanelElement,
                            }
                        );

                        setSearchPanelPosition(initialPosition);
                        searchPanelDragState = {
                            pointerId: event.pointerId,
                            offsetX: event.clientX - panelRect.left,
                            offsetY: event.clientY - panelRect.top,
                        };
                        globalThis.getSelection?.()?.removeAllRanges();
                    }}
                >
                    <div class="search-strip">
                        <div class="search-strip__header">
                            <div class="search-strip__tabs">
                                <button
                                    class="search-strip__tab"
                                    classList={{ "is-active": searchMode() === "find" }}
                                    type="button"
                                    onClick={() => {
                                        setSearchMode("find");
                                        setSearchFeedback(null);
                                    }}
                                >
                                    {strings.searchFind}
                                </button>
                                <button
                                    class="search-strip__tab"
                                    classList={{ "is-active": searchMode() === "replace" }}
                                    type="button"
                                    onClick={() => {
                                        setSearchMode("replace");
                                        setSearchFeedback(null);
                                    }}
                                >
                                    <span>{strings.searchReplace}</span>
                                </button>
                            </div>
                            <div class="search-strip__header-tools">
                                <button
                                    class="toolbar__button search-strip__close-button"
                                    data-role="search-close"
                                    type="button"
                                    title={strings.searchClose}
                                    onClick={closeSearchPanel}
                                >
                                    <span class="codicon codicon-close" />
                                </button>
                            </div>
                        </div>

                        <form
                            class="search-strip__row search-strip__row--primary"
                            data-role="search-form"
                            onSubmit={(event) => {
                                event.preventDefault();
                                submitSearch("next");
                            }}
                        >
                            <div
                                class="search-strip__input-wrap"
                                classList={{ "is-invalid": searchFeedback()?.tone === "error" }}
                            >
                                <span class="search-strip__input-icon codicon codicon-search" />
                                <input
                                    class="search-strip__input"
                                    data-role="search-input"
                                    value={searchQuery()}
                                    placeholder={strings.searchPlaceholder}
                                    onInput={(event) => {
                                        setSearchQuery(event.currentTarget.value);
                                        setSearchFeedback(null);
                                    }}
                                />
                                <div class="search-strip__input-tools">
                                    <button
                                        class="search-strip__icon-toggle"
                                        classList={{ "is-active": searchOptions().isRegexp }}
                                        type="button"
                                        title={strings.searchRegex}
                                        onClick={() => toggleSearchOption("isRegexp")}
                                    >
                                        <span class="codicon codicon-symbol-namespace" />
                                    </button>
                                    <button
                                        class="search-strip__icon-toggle"
                                        classList={{ "is-active": searchOptions().matchCase }}
                                        type="button"
                                        title={strings.searchMatchCase}
                                        onClick={() => toggleSearchOption("matchCase")}
                                    >
                                        <span class="codicon codicon-case-sensitive" />
                                    </button>
                                    <button
                                        class="search-strip__icon-toggle"
                                        classList={{ "is-active": searchOptions().wholeWord }}
                                        type="button"
                                        title={strings.searchWholeWord}
                                        onClick={() => toggleSearchOption("wholeWord")}
                                    >
                                        <span class="codicon codicon-whole-word" />
                                    </button>
                                </div>
                            </div>
                            <div class="search-strip__actions">
                                <button
                                    class="toolbar__button"
                                    data-role="search-prev"
                                    type="button"
                                    title={strings.findPrev}
                                    disabled={!canRunSearch()}
                                    onClick={() => submitSearch("prev")}
                                >
                                    <span class="codicon codicon-arrow-up" />
                                </button>
                                <button
                                    class="toolbar__button"
                                    data-role="search-next"
                                    type="submit"
                                    title={strings.findNext}
                                    disabled={!canRunSearch()}
                                >
                                    <span class="codicon codicon-arrow-down" />
                                </button>
                            </div>
                        </form>

                        <Show when={searchMode() === "replace"}>
                            <div class="search-strip__row search-strip__row--replace">
                                <div class="search-strip__input-wrap">
                                    <span class="search-strip__input-icon codicon codicon-replace" />
                                    <input
                                        class="search-strip__input"
                                        data-role="replace-input"
                                        value={replaceQuery()}
                                        placeholder={strings.replacePlaceholder}
                                        onInput={(event) => {
                                            setReplaceQuery(event.currentTarget.value);
                                            setSearchFeedback(null);
                                        }}
                                        onKeyDown={(event) => {
                                            if (event.key === "Enter") {
                                                event.preventDefault();
                                                submitReplace(
                                                    event.ctrlKey || event.metaKey ? "all" : "single"
                                                );
                                                return;
                                            }

                                            if (event.key === "Escape") {
                                                event.preventDefault();
                                                event.stopPropagation();
                                                closeSearchPanel();
                                            }
                                        }}
                                    />
                                </div>
                                <div class="search-strip__replace-actions">
                                    <button
                                        class="toolbar__button"
                                        data-role="replace-button"
                                        type="button"
                                        title={strings.searchReplace}
                                        disabled={!canRunReplace()}
                                        onClick={() => submitReplace("single")}
                                    >
                                        <span class="codicon codicon-replace" />
                                    </button>
                                    <button
                                        class="toolbar__button"
                                        data-role="replace-all-button"
                                        type="button"
                                        title={strings.replaceAll}
                                        disabled={!canRunReplace()}
                                        onClick={() => submitReplace("all")}
                                    >
                                        <span class="codicon codicon-replace-all" />
                                    </button>
                                </div>
                            </div>
                        </Show>

                        <div class="search-strip__row search-strip__row--meta">
                            <div class="search-strip__scope-summary">
                                <span class="search-strip__scope-summary-label">
                                    {strings.searchScopeLabel}
                                </span>
                                <span
                                    class="search-strip__scope-summary-value"
                                    classList={{ "is-selection": Boolean(activeSearchSelectionRange()) }}
                                >
                                    {searchScopeSummary()}
                                </span>
                            </div>
                            <Show when={searchFeedback()}>
                                {(feedback) => (
                                    <div class={searchFeedbackClass()}>{feedback().message}</div>
                                )}
                            </Show>
                        </div>
                    </div>
                </div>
            </Show>

            <div class="panes panes--single">
                <div class="pane pane--editor">
                    <Show
                        when={activeSheet()}
                        fallback={
                            session().mode === "error" && panelMessage() ? (
                                <div class="empty-shell">
                                    <div class="empty-shell__message">{panelMessage()}</div>
                                </div>
                            ) : session().initialized ? (
                                <div class="empty-shell">
                                    <div class="empty-shell__message">
                                        {strings.noRowsAvailable}
                                    </div>
                                </div>
                            ) : null
                        }
                    >
                        <Show when={gridMetrics()}>
                            {(metrics) => (
                                <div
                                    class="pane__table editor-grid-shell"
                                    data-role="editor-grid-shell"
                                >
                                    <div
                                        class="editor-grid__viewport"
                                        data-role="editor-grid-viewport"
                                        ref={(element) => {
                                            setViewportElement(element);
                                            syncViewportFromElement(element);
                                        }}
                                        onScroll={(event) =>
                                            syncViewportFromElement(event.currentTarget)
                                        }
                                    >
                                        <div
                                            class="editor-grid__canvas"
                                            style={{
                                                width: `${metrics().contentWidth}px`,
                                                height: `${metrics().contentHeight}px`,
                                            }}
                                        >
                                            <div class="editor-grid__layer editor-grid__layer--body">
                                                <Show when={cellLayers()}>
                                                    {(layers) => (
                                                        <For each={layers().body}>
                                                            {(item) => (
                                                                <EditorGridCellView
                                                                    item={item}
                                                                    sheetKey={
                                                                        activeSheet()?.key ?? null
                                                                    }
                                                                    offsetLeft={0}
                                                                    offsetTop={0}
                                                                    editingCell={editingCell()}
                                                                    fillSourceRange={fillSourceRange()}
                                                                    fillPreviewRange={fillPreviewRange()}
                                                                    onPointerSelectStart={
                                                                        beginCellSelectionDrag
                                                                    }
                                                                    onStartEdit={startEditingCell}
                                                                    onUpdateDraft={
                                                                        updateEditingDraft
                                                                    }
                                                                    onCommitEdit={commitEditingCell}
                                                                    onCancelEdit={cancelEditingCell}
                                                                    isFilterMenuOpen={
                                                                        filterMenu()?.sheetKey ===
                                                                            activeSheet()?.key &&
                                                                        filterMenu()
                                                                            ?.columnNumber ===
                                                                            item.columnNumber
                                                                    }
                                                                    onOpenFilterMenu={
                                                                        openFilterMenu
                                                                    }
                                                                    onOpenContextMenu={
                                                                        openCellContextMenu
                                                                    }
                                                                />
                                                            )}
                                                        </For>
                                                    )}
                                                </Show>
                                                <Show when={selectionOverlayLayers()?.body} keyed>
                                                    {(overlay) => (
                                                        <SelectionOverlayLayerView
                                                            overlay={overlay}
                                                            offsetLeft={0}
                                                            offsetTop={0}
                                                            isSearchFocused={
                                                                isSearchFocusedSelection()
                                                            }
                                                        />
                                                    )}
                                                </Show>
                                                <Show when={fillHandleLayout()?.layer === "body"}>
                                                    <EditorGridFillHandleView
                                                        layout={fillHandleLayout()!}
                                                        isActive={Boolean(fillPreviewRange())}
                                                        onPointerDown={(pointerId) => {
                                                            const range = fillSourceRange();
                                                            if (!range) {
                                                                return;
                                                            }

                                                            beginFillDrag(pointerId, range);
                                                        }}
                                                    />
                                                </Show>
                                            </div>
                                            <div
                                                class="editor-grid__overlay editor-grid__overlay--top"
                                                data-role="editor-grid-top-overlay"
                                                style={{
                                                    width: `${metrics().contentWidth}px`,
                                                    height: `${metrics().stickyTopHeight}px`,
                                                }}
                                            >
                                                <div
                                                    class="editor-grid__track editor-grid__track--x"
                                                    style={{
                                                        width: `${metrics().contentWidth}px`,
                                                        height: `${metrics().stickyTopHeight}px`,
                                                    }}
                                                >
                                                    <Show when={headerLayers()}>
                                                        {(headers) => (
                                                            <For
                                                                each={
                                                                    headers()
                                                                        .scrollableColumnHeaders
                                                                }
                                                            >
                                                                {(item) => (
                                                                    <EditorGridColumnHeaderView
                                                                        columnNumber={
                                                                            item.columnNumber
                                                                        }
                                                                        left={item.left}
                                                                        width={item.width}
                                                                        height={
                                                                            headers().headerHeight
                                                                        }
                                                                        label={item.label}
                                                                        isActive={item.isActive}
                                                                        canResize={
                                                                            workbook().canEdit
                                                                        }
                                                                        isResizing={
                                                                            columnResizeState()
                                                                                ?.sheetKey ===
                                                                                activeSheet()
                                                                                    ?.key &&
                                                                            columnResizeState()
                                                                                ?.columnNumber ===
                                                                                item.columnNumber
                                                                        }
                                                                        onPointerSelectStart={(
                                                                            pointerId,
                                                                            options
                                                                        ) =>
                                                                            beginColumnSelectionDrag(
                                                                                item.columnNumber,
                                                                                pointerId,
                                                                                options
                                                                            )
                                                                        }
                                                                        onPointerResizeStart={(
                                                                            pointerId,
                                                                            columnNumber,
                                                                            startPixelWidth,
                                                                            clientX
                                                                        ) =>
                                                                            beginColumnResize(
                                                                                pointerId,
                                                                                columnNumber,
                                                                                startPixelWidth,
                                                                                clientX
                                                                            )
                                                                        }
                                                                        onOpenContextMenu={
                                                                            openColumnContextMenu
                                                                        }
                                                                    />
                                                                )}
                                                            </For>
                                                        )}
                                                    </Show>
                                                    <Show when={cellLayers()}>
                                                        {(layers) => (
                                                            <For each={layers().top}>
                                                                {(item) => (
                                                                    <EditorGridCellView
                                                                        item={item}
                                                                        sheetKey={
                                                                            activeSheet()?.key ??
                                                                            null
                                                                        }
                                                                        offsetLeft={0}
                                                                        offsetTop={0}
                                                                        editingCell={editingCell()}
                                                                        fillSourceRange={fillSourceRange()}
                                                                        fillPreviewRange={fillPreviewRange()}
                                                                        onPointerSelectStart={
                                                                            beginCellSelectionDrag
                                                                        }
                                                                        onStartEdit={
                                                                            startEditingCell
                                                                        }
                                                                        onUpdateDraft={
                                                                            updateEditingDraft
                                                                        }
                                                                        onCommitEdit={
                                                                            commitEditingCell
                                                                        }
                                                                        onCancelEdit={
                                                                            cancelEditingCell
                                                                        }
                                                                        isFilterMenuOpen={
                                                                            filterMenu()
                                                                                ?.sheetKey ===
                                                                                activeSheet()
                                                                                    ?.key &&
                                                                            filterMenu()
                                                                                ?.columnNumber ===
                                                                                item.columnNumber
                                                                        }
                                                                        onOpenFilterMenu={
                                                                            openFilterMenu
                                                                        }
                                                                        onOpenContextMenu={
                                                                            openCellContextMenu
                                                                        }
                                                                    />
                                                                )}
                                                            </For>
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={selectionOverlayLayers()?.top}
                                                        keyed
                                                    >
                                                        {(overlay) => (
                                                            <SelectionOverlayLayerView
                                                                overlay={overlay}
                                                                offsetLeft={0}
                                                                offsetTop={0}
                                                                isSearchFocused={
                                                                    isSearchFocusedSelection()
                                                                }
                                                            />
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={fillHandleLayout()?.layer === "top"}
                                                    >
                                                        <EditorGridFillHandleView
                                                            layout={fillHandleLayout()!}
                                                            isActive={Boolean(fillPreviewRange())}
                                                            onPointerDown={(pointerId) => {
                                                                const range = fillSourceRange();
                                                                if (!range) {
                                                                    return;
                                                                }

                                                                beginFillDrag(pointerId, range);
                                                            }}
                                                        />
                                                    </Show>
                                                </div>
                                            </div>
                                            <div
                                                class="editor-grid__overlay editor-grid__overlay--left"
                                                data-role="editor-grid-left-overlay"
                                                style={{
                                                    width: `${metrics().stickyLeftWidth}px`,
                                                    height: `${metrics().contentHeight}px`,
                                                }}
                                            >
                                                <div
                                                    class="editor-grid__track editor-grid__track--y"
                                                    style={{
                                                        width: `${metrics().stickyLeftWidth}px`,
                                                        height: `${metrics().contentHeight}px`,
                                                    }}
                                                >
                                                    <Show when={headerLayers()}>
                                                        {(headers) => (
                                                            <For
                                                                each={
                                                                    headers().scrollableRowHeaders
                                                                }
                                                            >
                                                                {(item) => (
                                                                    <EditorGridRowHeaderView
                                                                        top={item.top}
                                                                        height={item.height}
                                                                        rowNumber={item.rowNumber}
                                                                        rowHeaderWidth={
                                                                            headers().rowHeaderWidth
                                                                        }
                                                                        isActive={item.isActive}
                                                                        canResize={
                                                                            workbook().canEdit
                                                                        }
                                                                        isResizing={
                                                                            rowResizeState()
                                                                                ?.sheetKey ===
                                                                                activeSheet()
                                                                                    ?.key &&
                                                                            rowResizeState()
                                                                                ?.rowNumber ===
                                                                                item.rowNumber
                                                                        }
                                                                        onPointerSelectStart={(
                                                                            pointerId,
                                                                            options
                                                                        ) =>
                                                                            beginRowSelectionDrag(
                                                                                item.rowNumber,
                                                                                pointerId,
                                                                                options
                                                                            )
                                                                        }
                                                                        onPointerResizeStart={(
                                                                            pointerId,
                                                                            rowNumber,
                                                                            startPixelHeight,
                                                                            clientY
                                                                        ) =>
                                                                            beginRowResize(
                                                                                pointerId,
                                                                                rowNumber,
                                                                                startPixelHeight,
                                                                                clientY
                                                                            )
                                                                        }
                                                                        onOpenContextMenu={
                                                                            openRowContextMenu
                                                                        }
                                                                    />
                                                                )}
                                                            </For>
                                                        )}
                                                    </Show>
                                                    <Show when={cellLayers()}>
                                                        {(layers) => (
                                                            <For each={layers().left}>
                                                                {(item) => (
                                                                    <EditorGridCellView
                                                                        item={item}
                                                                        sheetKey={
                                                                            activeSheet()?.key ??
                                                                            null
                                                                        }
                                                                        offsetLeft={0}
                                                                        offsetTop={0}
                                                                        editingCell={editingCell()}
                                                                        fillSourceRange={fillSourceRange()}
                                                                        fillPreviewRange={fillPreviewRange()}
                                                                        onPointerSelectStart={
                                                                            beginCellSelectionDrag
                                                                        }
                                                                        onStartEdit={
                                                                            startEditingCell
                                                                        }
                                                                        onUpdateDraft={
                                                                            updateEditingDraft
                                                                        }
                                                                        onCommitEdit={
                                                                            commitEditingCell
                                                                        }
                                                                        onCancelEdit={
                                                                            cancelEditingCell
                                                                        }
                                                                        isFilterMenuOpen={
                                                                            filterMenu()
                                                                                ?.sheetKey ===
                                                                                activeSheet()
                                                                                    ?.key &&
                                                                            filterMenu()
                                                                                ?.columnNumber ===
                                                                                item.columnNumber
                                                                        }
                                                                        onOpenFilterMenu={
                                                                            openFilterMenu
                                                                        }
                                                                        onOpenContextMenu={
                                                                            openCellContextMenu
                                                                        }
                                                                    />
                                                                )}
                                                            </For>
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={selectionOverlayLayers()?.left}
                                                        keyed
                                                    >
                                                        {(overlay) => (
                                                            <SelectionOverlayLayerView
                                                                overlay={overlay}
                                                                offsetLeft={0}
                                                                offsetTop={0}
                                                                isSearchFocused={
                                                                    isSearchFocusedSelection()
                                                                }
                                                            />
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={fillHandleLayout()?.layer === "left"}
                                                    >
                                                        <EditorGridFillHandleView
                                                            layout={fillHandleLayout()!}
                                                            isActive={Boolean(fillPreviewRange())}
                                                            onPointerDown={(pointerId) => {
                                                                const range = fillSourceRange();
                                                                if (!range) {
                                                                    return;
                                                                }

                                                                beginFillDrag(pointerId, range);
                                                            }}
                                                        />
                                                    </Show>
                                                </div>
                                            </div>
                                            <div
                                                class="editor-grid__overlay editor-grid__overlay--corner"
                                                data-role="editor-grid-corner-overlay"
                                                style={{
                                                    width: `${metrics().stickyLeftWidth}px`,
                                                    height: `${metrics().stickyTopHeight}px`,
                                                }}
                                            >
                                                <div
                                                    class="editor-grid__track"
                                                    style={{
                                                        width: `${metrics().stickyLeftWidth}px`,
                                                        height: `${metrics().stickyTopHeight}px`,
                                                    }}
                                                >
                                                    <Show when={headerLayers()}>
                                                        {(headers) => (
                                                            <>
                                                                <EditorGridCornerHeaderView
                                                                    rowHeaderWidth={
                                                                        headers().rowHeaderWidth
                                                                    }
                                                                    headerHeight={
                                                                        headers().headerHeight
                                                                    }
                                                                />
                                                                <For
                                                                    each={
                                                                        headers()
                                                                            .frozenColumnHeaders
                                                                    }
                                                                >
                                                                    {(item) => (
                                                                        <EditorGridColumnHeaderView
                                                                            columnNumber={
                                                                                item.columnNumber
                                                                            }
                                                                            left={item.left}
                                                                            width={item.width}
                                                                            height={
                                                                                headers()
                                                                                    .headerHeight
                                                                            }
                                                                            label={item.label}
                                                                            isActive={item.isActive}
                                                                            canResize={
                                                                                workbook().canEdit
                                                                            }
                                                                            isResizing={
                                                                                columnResizeState()
                                                                                    ?.sheetKey ===
                                                                                    activeSheet()
                                                                                        ?.key &&
                                                                                columnResizeState()
                                                                                    ?.columnNumber ===
                                                                                    item.columnNumber
                                                                            }
                                                                            onPointerSelectStart={(
                                                                                pointerId,
                                                                                options
                                                                            ) =>
                                                                                beginColumnSelectionDrag(
                                                                                    item.columnNumber,
                                                                                    pointerId,
                                                                                    options
                                                                                )
                                                                            }
                                                                            onPointerResizeStart={(
                                                                                pointerId,
                                                                                columnNumber,
                                                                                startPixelWidth,
                                                                                clientX
                                                                            ) =>
                                                                                beginColumnResize(
                                                                                    pointerId,
                                                                                    columnNumber,
                                                                                    startPixelWidth,
                                                                                    clientX
                                                                                )
                                                                            }
                                                                            onOpenContextMenu={
                                                                                openColumnContextMenu
                                                                            }
                                                                        />
                                                                    )}
                                                                </For>
                                                                <For
                                                                    each={
                                                                        headers().frozenRowHeaders
                                                                    }
                                                                >
                                                                    {(item) => (
                                                                        <EditorGridRowHeaderView
                                                                            top={item.top}
                                                                            height={item.height}
                                                                            rowNumber={
                                                                                item.rowNumber
                                                                            }
                                                                            rowHeaderWidth={
                                                                                headers()
                                                                                    .rowHeaderWidth
                                                                            }
                                                                            isActive={item.isActive}
                                                                            canResize={
                                                                                workbook().canEdit
                                                                            }
                                                                            isResizing={
                                                                                rowResizeState()
                                                                                    ?.sheetKey ===
                                                                                    activeSheet()
                                                                                        ?.key &&
                                                                                rowResizeState()
                                                                                    ?.rowNumber ===
                                                                                    item.rowNumber
                                                                            }
                                                                            onPointerSelectStart={(
                                                                                pointerId,
                                                                                options
                                                                            ) =>
                                                                                beginRowSelectionDrag(
                                                                                    item.rowNumber,
                                                                                    pointerId,
                                                                                    options
                                                                                )
                                                                            }
                                                                            onPointerResizeStart={(
                                                                                pointerId,
                                                                                rowNumber,
                                                                                startPixelHeight,
                                                                                clientY
                                                                            ) =>
                                                                                beginRowResize(
                                                                                    pointerId,
                                                                                    rowNumber,
                                                                                    startPixelHeight,
                                                                                    clientY
                                                                                )
                                                                            }
                                                                            onOpenContextMenu={
                                                                                openRowContextMenu
                                                                            }
                                                                        />
                                                                    )}
                                                                </For>
                                                            </>
                                                        )}
                                                    </Show>
                                                    <Show when={cellLayers()}>
                                                        {(layers) => (
                                                            <For each={layers().corner}>
                                                                {(item) => (
                                                                    <EditorGridCellView
                                                                        item={item}
                                                                        sheetKey={
                                                                            activeSheet()?.key ??
                                                                            null
                                                                        }
                                                                        offsetLeft={0}
                                                                        offsetTop={0}
                                                                        editingCell={editingCell()}
                                                                        fillSourceRange={fillSourceRange()}
                                                                        fillPreviewRange={fillPreviewRange()}
                                                                        onPointerSelectStart={
                                                                            beginCellSelectionDrag
                                                                        }
                                                                        onStartEdit={
                                                                            startEditingCell
                                                                        }
                                                                        onUpdateDraft={
                                                                            updateEditingDraft
                                                                        }
                                                                        onCommitEdit={
                                                                            commitEditingCell
                                                                        }
                                                                        onCancelEdit={
                                                                            cancelEditingCell
                                                                        }
                                                                        isFilterMenuOpen={
                                                                            filterMenu()
                                                                                ?.sheetKey ===
                                                                                activeSheet()
                                                                                    ?.key &&
                                                                            filterMenu()
                                                                                ?.columnNumber ===
                                                                                item.columnNumber
                                                                        }
                                                                        onOpenFilterMenu={
                                                                            openFilterMenu
                                                                        }
                                                                        onOpenContextMenu={
                                                                            openCellContextMenu
                                                                        }
                                                                    />
                                                                )}
                                                            </For>
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={selectionOverlayLayers()?.corner}
                                                        keyed
                                                    >
                                                        {(overlay) => (
                                                            <SelectionOverlayLayerView
                                                                overlay={overlay}
                                                                offsetLeft={0}
                                                                offsetTop={0}
                                                                isSearchFocused={
                                                                    isSearchFocusedSelection()
                                                                }
                                                            />
                                                        )}
                                                    </Show>
                                                    <Show
                                                        when={
                                                            fillHandleLayout()?.layer === "corner"
                                                        }
                                                    >
                                                        <EditorGridFillHandleView
                                                            layout={fillHandleLayout()!}
                                                            isActive={Boolean(fillPreviewRange())}
                                                            onPointerDown={(pointerId) => {
                                                                const range = fillSourceRange();
                                                                if (!range) {
                                                                    return;
                                                                }

                                                                beginFillDrag(pointerId, range);
                                                            }}
                                                        />
                                                    </Show>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            )}
                        </Show>
                    </Show>
                </div>
            </div>

            <div class="footer">
                <div class="tabs">
                    <div class="tabs__viewport">
                        <div class="tabs__content">
                            <div class="tabs__list">
                                <For each={workbook().sheets}>
                                    {(sheet) => (
                                        <button
                                            class="tab"
                                            classList={{ "is-active": sheet.isActive }}
                                            data-role="sheet-tab"
                                            data-sheet-key={sheet.key}
                                            type="button"
                                            onContextMenu={(event) => {
                                                event.preventDefault();
                                                openSheetContextMenu(
                                                    sheet.key,
                                                    event.clientX,
                                                    event.clientY
                                                );
                                            }}
                                            onClick={() =>
                                                postMessage(createEditorSetSheetMessage(sheet.key))
                                            }
                                        >
                                            <span class="tab__label">{sheet.label}</span>
                                        </button>
                                    )}
                                </For>
                                <button
                                    class="tab"
                                    data-role="sheet-menu-toggle"
                                    type="button"
                                    disabled={!activeSheetTabKey()}
                                    onClick={(event) => {
                                        const targetSheetKey = activeSheetTabKey();
                                        if (!targetSheetKey) {
                                            return;
                                        }

                                        const rect = event.currentTarget.getBoundingClientRect();
                                        openSheetContextMenu(
                                            targetSheetKey,
                                            rect.left,
                                            rect.top - 8
                                        );
                                    }}
                                >
                                    <span class="tab__label">{strings.moreSheets}</span>
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <Show when={sheetContextMenu()}>
                {(menu) => (
                    <div
                        class="context-menu"
                        data-role="sheet-context-menu"
                        style={{
                            left: `${menu().x}px`,
                            top: `${menu().y}px`,
                        }}
                    >
                        <button
                            class="context-menu__item"
                            data-role="sheet-context-add"
                            type="button"
                            disabled={!sheetContextMenuState().canAddSheet}
                            onClick={() => {
                                postMessage(createEditorAddSheetMessage());
                                setSheetContextMenu(null);
                                return;
                            }}
                        >
                            <span class="context-menu__icon codicon codicon-add" />
                            <span>{strings.addSheet}</span>
                        </button>
                        <button
                            class="context-menu__item"
                            data-role="sheet-context-rename"
                            type="button"
                            disabled={
                                !sheetContextTargetKey() || !sheetContextMenuState().canRenameSheet
                            }
                            onClick={() => {
                                const targetSheetKey = sheetContextTargetKey();
                                if (!targetSheetKey) {
                                    return;
                                }

                                postMessage(createEditorRenameSheetMessage(targetSheetKey));
                                setSheetContextMenu(null);
                            }}
                        >
                            <span class="context-menu__icon codicon codicon-edit" />
                            <span>{strings.renameSheet}</span>
                        </button>
                        <button
                            class="context-menu__item context-menu__item--danger"
                            data-role="sheet-context-delete"
                            type="button"
                            disabled={
                                !sheetContextTargetKey() || !sheetContextMenuState().canDeleteSheet
                            }
                            onClick={() => {
                                const targetSheetKey = sheetContextTargetKey();
                                if (!targetSheetKey) {
                                    return;
                                }

                                postMessage(createEditorDeleteSheetMessage(targetSheetKey));
                                setSheetContextMenu(null);
                            }}
                        >
                            <span class="context-menu__icon codicon codicon-trash" />
                            <span>{strings.deleteSheet}</span>
                        </button>
                    </div>
                )}
            </Show>
            <Show when={gridContextMenu()}>
                {(menu) => (
                    <div
                        class="context-menu"
                        classList={{ "context-menu--cell": menu().kind === "cell" }}
                        data-role={
                            menu().kind === "cell"
                                ? "cell-context-menu"
                                : menu().kind === "row"
                                  ? "row-context-menu"
                                  : "column-context-menu"
                        }
                        style={gridContextMenuStyle() ?? {}}
                    >
                        <Show when={menu().kind === "cell"}>
                            <button
                                class="context-menu__item"
                                data-role="cell-context-find"
                                type="button"
                                onClick={() => {
                                    setGridContextMenu(null);
                                    openSearchPanel("find");
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-search" />
                                <span class="context-menu__label">{strings.searchFind}</span>
                                <span
                                    class="context-menu__shortcut"
                                    data-role="cell-context-find-shortcut"
                                >
                                    Ctrl/Cmd+F
                                </span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="cell-context-replace"
                                type="button"
                                onClick={() => {
                                    setGridContextMenu(null);
                                    openSearchPanel("replace");
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-replace" />
                                <span class="context-menu__label">{strings.searchReplace}</span>
                                <span
                                    class="context-menu__shortcut"
                                    data-role="cell-context-replace-shortcut"
                                >
                                    Ctrl/Cmd+H
                                </span>
                            </button>
                            <div class="context-menu__separator" aria-hidden="true" />
                            <button
                                class="context-menu__item"
                                data-role="cell-context-filter"
                                type="button"
                                disabled={!canToggleFilter()}
                                onClick={() => {
                                    setGridContextMenu(null);
                                    submitFilterToggle();
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-filter" />
                                <span class="context-menu__label">{filterActionLabel()}</span>
                            </button>
                            <div class="context-menu__separator" aria-hidden="true" />
                            <button
                                class="context-menu__item"
                                data-role="cell-context-insert-row-above"
                                type="button"
                                disabled={!workbook().canEdit}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestInsertRow(current.rowNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span class="context-menu__label">{strings.insertRowAbove}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="cell-context-insert-row-below"
                                type="button"
                                disabled={!workbook().canEdit}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestInsertRow(current.rowNumber + 1);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span class="context-menu__label">{strings.insertRowBelow}</span>
                            </button>
                            <button
                                class="context-menu__item context-menu__item--danger"
                                data-role="cell-context-delete-row"
                                type="button"
                                disabled={
                                    !workbook().canEdit || (activeSheetView()?.rowCount ?? 0) <= 1
                                }
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestDeleteRow(current.rowNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-trash" />
                                <span class="context-menu__label">{strings.deleteRow}</span>
                            </button>
                            <div class="context-menu__separator" aria-hidden="true" />
                            <button
                                class="context-menu__item"
                                data-role="cell-context-insert-column-left"
                                type="button"
                                disabled={!workbook().canEdit}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestInsertColumn(current.columnNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span class="context-menu__label">{strings.insertColumnLeft}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="cell-context-insert-column-right"
                                type="button"
                                disabled={!workbook().canEdit}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestInsertColumn(current.columnNumber + 1);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span class="context-menu__label">
                                    {strings.insertColumnRight}
                                </span>
                            </button>
                            <button
                                class="context-menu__item context-menu__item--danger"
                                data-role="cell-context-delete-column"
                                type="button"
                                disabled={
                                    !workbook().canEdit ||
                                    (activeSheetView()?.columnCount ?? 0) <= 1
                                }
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "cell") {
                                        requestDeleteColumn(current.columnNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-trash" />
                                <span class="context-menu__label">{strings.deleteColumn}</span>
                            </button>
                        </Show>
                        <Show when={menu().kind === "row"}>
                            <button
                                class="context-menu__item"
                                data-role="row-context-height"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "row") {
                                        requestPromptRowHeight(current.rowNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-symbol-number" />
                                <span>{strings.setRowHeight}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="row-context-insert-above"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "row") {
                                        requestInsertRow(current.rowNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span>{strings.insertRowAbove}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="row-context-insert-below"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "row") {
                                        requestInsertRow(current.rowNumber + 1);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span>{strings.insertRowBelow}</span>
                            </button>
                            <button
                                class="context-menu__item context-menu__item--danger"
                                data-role="row-context-delete"
                                type="button"
                                disabled={(activeSheetView()?.rowCount ?? 0) <= 1}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "row") {
                                        requestDeleteRow(current.rowNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-trash" />
                                <span>{strings.deleteRow}</span>
                            </button>
                        </Show>
                        <Show when={menu().kind === "column"}>
                            <button
                                class="context-menu__item"
                                data-role="column-context-width"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "column") {
                                        requestPromptColumnWidth(current.columnNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-symbol-number" />
                                <span>{strings.setColumnWidth}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="column-context-insert-left"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "column") {
                                        requestInsertColumn(current.columnNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span>{strings.insertColumnLeft}</span>
                            </button>
                            <button
                                class="context-menu__item"
                                data-role="column-context-insert-right"
                                type="button"
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "column") {
                                        requestInsertColumn(current.columnNumber + 1);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-add" />
                                <span>{strings.insertColumnRight}</span>
                            </button>
                            <button
                                class="context-menu__item context-menu__item--danger"
                                data-role="column-context-delete"
                                type="button"
                                disabled={(activeSheetView()?.columnCount ?? 0) <= 1}
                                onClick={() => {
                                    const current = menu();
                                    if (current.kind === "column") {
                                        requestDeleteColumn(current.columnNumber);
                                    }
                                }}
                            >
                                <span class="context-menu__icon codicon codicon-trash" />
                                <span>{strings.deleteColumn}</span>
                            </button>
                        </Show>
                    </div>
                )}
            </Show>
        </div>
    );
}

export function mountEditorBootstrapApp(target: HTMLElement) {
    return render(() => <EditorBootstrapApp />, target);
}

if (typeof document !== "undefined") {
    const rootElement = document.getElementById("app");
    if (rootElement) {
        mountEditorBootstrapApp(rootElement);
    }
}
