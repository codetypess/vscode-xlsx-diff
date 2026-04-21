import * as React from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import {
    DEFAULT_EDITOR_WINDOW_OVERSCAN,
    DEFAULT_EDITOR_WINDOW_SIZE,
} from "../constants";
import type {
    EditorGridCellView,
    EditorGridRowView,
    EditorRenderModel,
    EditorSheetTabView,
} from "../core/model/types";
import { getColumnLabel } from "../core/model/cells";
import { getFreezePaneCountsForCell, hasLockedView } from "./viewLock";

interface VsCodeApi {
    postMessage(message: OutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

type OutgoingMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "addSheet" }
    | { type: "deleteSheet"; sheetKey: string }
    | { type: "renameSheet"; sheetKey: string }
    | { type: "setViewportStartRow"; rowNumber: number }
    | {
          type: "search";
          query: string;
          direction: "next" | "prev";
          options: SearchOptions;
      }
    | { type: "gotoCell"; reference: string }
    | { type: "selectCell"; rowNumber: number; columnNumber: number }
    | {
          type: "setPendingEdits";
          edits: Array<{
              sheetKey: string;
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
      }
    | { type: "requestSave" }
    | { type: "pendingEditStateChanged"; hasPendingEdits: boolean }
    | { type: "undoSheetEdit" }
    | { type: "redoSheetEdit" }
    | { type: "toggleViewLock"; rowCount: number; columnCount: number }
    | { type: "reload" };

type IncomingMessage =
    | { type: "loading"; message: string }
    | { type: "error"; message: string }
    | {
          type: "render";
          payload: EditorRenderModel;
          silent?: boolean;
          clearPendingEdits?: boolean;
          useModelSelection?: boolean;
          replacePendingEdits?: Array<{
              sheetKey: string;
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
          resetPendingHistory?: boolean;
      };

interface SearchOptions {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

interface CellPosition {
    rowNumber: number;
    columnNumber: number;
}

interface CellRange {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

interface EditingCell extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingEdit extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingEditChange extends CellPosition {
    sheetKey: string;
    modelValue: string;
    beforeValue: string;
    afterValue: string;
}

interface HistoryEntry {
    changes: PendingEditChange[];
}

interface PendingSelection extends CellPosition {
    reveal: boolean;
}

interface PendingSummary {
    sheetKeys: Set<string>;
    rows: Set<number>;
    columns: Set<number>;
}

interface ScrollState {
    top: number;
    left: number;
}

interface TabContextMenuState {
    sheetKey: string;
    x: number;
    y: number;
}

interface SelectionDragState {
    anchorCell: CellPosition;
    pointerId: number;
}

type FrozenPaneName = "top-left" | "top-right" | "bottom-left" | "bottom-right";

type ViewState =
    | { kind: "loading"; message: string }
    | { kind: "error"; message: string }
    | {
          kind: "app";
          model: EditorRenderModel;
          revealSelection: boolean;
          revision: number;
          scrollState: ScrollState | null;
      };

const DEFAULT_STRINGS = {
    loading: "Loading XLSX editor...",
    reload: "Reload",
    undo: "Undo",
    redo: "Redo",
    searchPlaceholder: "Search values or formulas",
    findPrev: "Prev Match",
    findNext: "Next Match",
    gotoPlaceholder: "A1 or Sheet1!B2",
    goto: "Go",
    searchRegex: "Use Regular Expression",
    searchMatchCase: "Match Case",
    searchWholeWord: "Match Whole Word",
    size: "Size",
    modified: "Modified",
    sheet: "Sheet",
    rows: "Rows",
    noRows: "No rows",
    visibleRows: "Visible rows",
    readOnly: "Read-only",
    save: "Save",
    lockView: "Lock View",
    unlockView: "Unlock View",
    addSheet: "Add Sheet",
    deleteSheet: "Delete Sheet",
    renameSheet: "Rename Sheet",
    selectedCell: "Selected cell",
    noCellSelected: "None",
    totalSheets: "Sheets",
    totalRows: "Rows",
    nonEmptyCells: "Non-empty cells",
    mergedRanges: "Merged ranges",
    noRowsAvailable: "No rows available in this view.",
    readOnlyBadge: "Read-only",
};

type Strings = typeof DEFAULT_STRINGS;

const STRINGS: Strings =
    ((globalThis as Record<string, unknown>).__XLSX_EDITOR_STRINGS__ as Strings | undefined) ??
    DEFAULT_STRINGS;

const vscode = acquireVsCodeApi();
const pendingEdits = new Map<string, PendingEdit>();
const undoStack: HistoryEntry[] = [];
const redoStack: HistoryEntry[] = [];

let model: EditorRenderModel | null = null;
let selectedCell: CellPosition | null = null;
let selectionAnchorCell: CellPosition | null = null;
let editingCell: EditingCell | null = null;
let isSaving = false;
let lastPendingNotification: boolean | null = null;
let pendingSelectionAfterRender: PendingSelection | null = null;
let suppressAutoSelection = false;
let searchQuery = "";
let gotoReference = "";
let lastPendingEditsSyncKey: string | null = null;
let viewRevision = 0;
let setViewState: React.Dispatch<React.SetStateAction<ViewState>> | null = null;
let tabContextMenu: TabContextMenuState | null = null;
let selectionDragState: SelectionDragState | null = null;
let suppressNextCellClick = false;
let frozenPaneLayoutFrame = 0;
let deferredGridLayoutTimer = 0;
let isSyncingFrozenPaneScroll = false;
let lastRequestedViewportStartRow: number | null = null;
let queuedViewportStartRow: number | null = null;
let lastObservedPaneScrollTop: number | null = null;
let lastMeasuredSheetKey: string | null = null;
let pendingViewportScrollFrame = 0;
let pendingViewportScrollPane: HTMLElement | null = null;
let lastScrollActivityAt = 0;
const rowHeightByNumber = new Map<number, number>();
const columnWidthByNumber = new Map<number, number>();
let searchOptions: SearchOptions = {
    isRegexp: false,
    matchCase: false,
    wholeWord: false,
};

const ESTIMATED_EDITOR_ROW_HEIGHT = 28;
const EDITOR_COLUMN_MIN_WIDTH = 108;
const EDITOR_COLUMN_MAX_WIDTH = 240;
const ACTIVE_SCROLL_LAYOUT_DELAY_MS = 140;
const WHEEL_DELTA_LINE_MODE = 1;
const WHEEL_DELTA_PAGE_MODE = 2;
const WHEEL_LINE_SCROLL_PIXELS = 40;

function getViewportShiftRowCount(): number {
    return Math.max(1, DEFAULT_EDITOR_WINDOW_SIZE - DEFAULT_EDITOR_WINDOW_OVERSCAN * 2);
}

function getNow(): number {
    return globalThis.performance?.now() ?? Date.now();
}

function recordScrollActivity(): void {
    lastScrollActivityAt = getNow();
}

function isActivelyScrolling(): boolean {
    return getNow() - lastScrollActivityAt < ACTIVE_SCROLL_LAYOUT_DELAY_MS;
}

function getPendingEditKey(sheetKey: string, rowNumber: number, columnNumber: number): string {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
}

function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}

function serializePendingEdits(edits: PendingEdit[]): string {
    return JSON.stringify(
        [...edits]
            .sort((left, right) => {
                if (left.sheetKey !== right.sheetKey) {
                    return left.sheetKey.localeCompare(right.sheetKey);
                }

                if (left.rowNumber !== right.rowNumber) {
                    return left.rowNumber - right.rowNumber;
                }

                if (left.columnNumber !== right.columnNumber) {
                    return left.columnNumber - right.columnNumber;
                }

                return left.value.localeCompare(right.value);
            })
            .map((edit) => ({
                sheetKey: edit.sheetKey,
                rowNumber: edit.rowNumber,
                columnNumber: edit.columnNumber,
                value: edit.value,
            }))
    );
}

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

function clearBrowserTextSelection(): void {
    globalThis.getSelection?.()?.removeAllRanges();
}

function getCellTooltip(address: string, value: string, formula: string | null): string {
    const lines = [address];

    if (value) {
        lines.push(value);
    }

    if (formula) {
        lines.push(`fx ${formula}`);
    }

    return lines.join("\n");
}

function getRenderableRows(currentModel: EditorRenderModel | null): EditorGridRowView[] {
    if (!currentModel) {
        return [];
    }

    const rows: EditorGridRowView[] = [];
    const seenRows = new Set<number>();
    for (const row of [...currentModel.page.frozenRows, ...currentModel.page.rows]) {
        if (seenRows.has(row.rowNumber)) {
            continue;
        }

        seenRows.add(row.rowNumber);
        rows.push(row);
    }

    return rows;
}

function getApproximateRowHeight(
    currentModel: EditorRenderModel | null,
    rowNumber: number
): number {
    return rowHeightByNumber.get(rowNumber) ?? ESTIMATED_EDITOR_ROW_HEIGHT;
}

function getApproximateRowSpanHeight(
    currentModel: EditorRenderModel | null,
    startRow: number,
    endRow: number
): number {
    if (endRow < startRow) {
        return 0;
    }

    let height = 0;
    for (let rowNumber = Math.max(startRow, 1); rowNumber <= endRow; rowNumber += 1) {
        height += getApproximateRowHeight(currentModel, rowNumber);
    }

    return height;
}

function getVisibleScrollableRowRange(
    currentModel: EditorRenderModel,
    pane: HTMLElement
): { firstVisibleRow: number; lastVisibleRow: number } {
    const rows = currentModel.page.rows;
    const baseRow = getScrollableViewportBaseRow(currentModel);
    if (rows.length === 0) {
        return {
            firstVisibleRow: baseRow,
            lastVisibleRow: baseRow,
        };
    }

    let rowIndex = 0;
    let offsetWithinWindow = Math.max(0, pane.scrollTop - getTopSpacerHeight(currentModel));
    while (rowIndex < rows.length) {
        const rowHeight = getApproximateRowHeight(currentModel, rows[rowIndex].rowNumber);
        if (offsetWithinWindow < rowHeight) {
            break;
        }

        offsetWithinWindow -= rowHeight;
        rowIndex += 1;
    }

    const clampedRowIndex = Math.min(rowIndex, rows.length - 1);
    const firstVisibleRow = rows[clampedRowIndex]?.rowNumber ?? rows[0].rowNumber;
    let lastVisibleRow = firstVisibleRow;
    let remainingHeight = pane.clientHeight + offsetWithinWindow;

    for (let visibleRowIndex = clampedRowIndex; visibleRowIndex < rows.length; visibleRowIndex += 1) {
        const row = rows[visibleRowIndex];
        if (!row) {
            break;
        }

        lastVisibleRow = row.rowNumber;
        remainingHeight -= getApproximateRowHeight(currentModel, row.rowNumber);
        if (remainingHeight <= 0) {
            break;
        }
    }

    return { firstVisibleRow, lastVisibleRow };
}

function getTopSpacerHeight(currentModel: EditorRenderModel): number {
    const baseRow = getScrollableViewportBaseRow(currentModel);
    if (currentModel.page.startRow <= baseRow) {
        return 0;
    }

    return getApproximateRowSpanHeight(currentModel, baseRow, currentModel.page.startRow - 1);
}

function getBottomSpacerHeight(currentModel: EditorRenderModel): number {
    if (currentModel.page.endRow <= 0 || currentModel.page.endRow >= currentModel.page.totalRows) {
        return 0;
    }

    return getApproximateRowSpanHeight(
        currentModel,
        currentModel.page.endRow + 1,
        currentModel.page.totalRows
    );
}

function getScrollableViewportBaseRow(currentModel: EditorRenderModel | null): number {
    if (!currentModel) {
        return 1;
    }

    return Math.max(
        1,
        hasLockedView(currentModel.activeSheet.freezePane)
            ? Math.min(
                  (currentModel.activeSheet.freezePane?.rowCount ?? 0) + 1,
                  currentModel.page.totalRows
              )
            : 1
    );
}

function getScrollableViewportRowCount(currentModel: EditorRenderModel | null): number {
    if (!currentModel) {
        return 0;
    }

    return Math.max(
        0,
        currentModel.page.totalRows - getScrollableViewportBaseRow(currentModel) + 1
    );
}

function getMaxViewportStartRow(currentModel: EditorRenderModel | null): number {
    if (!currentModel) {
        return 1;
    }

    const baseRow = getScrollableViewportBaseRow(currentModel);
    return Math.max(currentModel.page.totalRows - DEFAULT_EDITOR_WINDOW_SIZE + 1, baseRow);
}

function hasPendingViewportRequest(currentModel: EditorRenderModel | null): boolean {
    return Boolean(
        currentModel &&
            lastRequestedViewportStartRow !== null &&
            lastRequestedViewportStartRow !== currentModel.page.startRow
    );
}

function requestViewportStartRow(rowNumber: number): void {
    if (!model) {
        return;
    }

    const minRowNumber = getScrollableViewportBaseRow(model);
    const clampedRowNumber = Math.max(
        minRowNumber,
        Math.min(rowNumber, getMaxViewportStartRow(model))
    );
    if (clampedRowNumber === model.page.startRow || clampedRowNumber === lastRequestedViewportStartRow) {
        return;
    }

    if (hasPendingViewportRequest(model)) {
        queuedViewportStartRow = clampedRowNumber;
        return;
    }

    queuedViewportStartRow = null;
    lastRequestedViewportStartRow = clampedRowNumber;
    vscode.postMessage({ type: "setViewportStartRow", rowNumber: clampedRowNumber });
}

function maybeRequestViewportForScroll(pane: HTMLElement): void {
    if (!model) {
        return;
    }

    if (getScrollableViewportRowCount(model) <= model.page.visibleRowCount) {
        return;
    }

    const baseRow = getScrollableViewportBaseRow(model);
    const currentStartRow = model.page.startRow;
    const currentEndRow = model.page.endRow;
    const { firstVisibleRow, lastVisibleRow } = getVisibleScrollableRowRange(model, pane);
    const previousScrollTop = lastObservedPaneScrollTop;
    lastObservedPaneScrollTop = pane.scrollTop;
    if (previousScrollTop === null || pane.scrollTop === previousScrollTop) {
        return;
    }

    const isScrollingDown = pane.scrollTop > previousScrollTop;
    const shiftRowCount = getViewportShiftRowCount();

    if (isScrollingDown) {
        const remainingRowsBelowViewport = currentEndRow - lastVisibleRow;
        if (remainingRowsBelowViewport > DEFAULT_EDITOR_WINDOW_OVERSCAN) {
            return;
        }

        requestViewportStartRow(currentStartRow + shiftRowCount);
        return;
    }

    const rowsAboveViewport = firstVisibleRow - currentStartRow;
    if (rowsAboveViewport > DEFAULT_EDITOR_WINDOW_OVERSCAN) {
        return;
    }

    requestViewportStartRow(currentStartRow - shiftRowCount);
}

function scheduleViewportRequestForScroll(pane: HTMLElement): void {
    recordScrollActivity();
    pendingViewportScrollPane = pane;

    if (pendingViewportScrollFrame) {
        return;
    }

    pendingViewportScrollFrame = requestAnimationFrame(() => {
        pendingViewportScrollFrame = 0;
        const nextPane = pendingViewportScrollPane;
        pendingViewportScrollPane = null;
        if (!nextPane) {
            return;
        }

        maybeRequestViewportForScroll(nextPane);
    });
}

function hasFrozenRowGap(currentModel: EditorRenderModel | null): boolean {
    if (!currentModel || currentModel.page.frozenRows.length === 0 || currentModel.page.rows.length === 0) {
        return false;
    }

    const lastFrozenRow = currentModel.page.frozenRows[currentModel.page.frozenRows.length - 1];
    const firstPageRow = currentModel.page.rows[0];
    return Boolean(lastFrozenRow && firstPageRow && lastFrozenRow.rowNumber + 1 < firstPageRow.rowNumber);
}

function getFirstRenderableRow(currentModel: EditorRenderModel | null): EditorGridRowView | null {
    return getRenderableRows(currentModel)[0] ?? null;
}

function getVisibleFreezeCounts(
    currentModel: EditorRenderModel | null
): { rowCount: number; columnCount: number } {
    if (!currentModel || !hasLockedView(currentModel.activeSheet.freezePane)) {
        return { rowCount: 0, columnCount: 0 };
    }

    return {
        rowCount: Math.max(
            0,
            Math.min(
                currentModel.activeSheet.freezePane.rowCount,
                Math.max(currentModel.activeSheet.rowCount - 1, 0)
            )
        ),
        columnCount: Math.max(
            0,
            Math.min(
                currentModel.activeSheet.freezePane.columnCount,
                Math.max(currentModel.activeSheet.columnCount - 1, 0)
            )
        ),
    };
}

function getFrozenRowsForCurrentView(currentModel: EditorRenderModel): EditorGridRowView[] {
    const { rowCount } = getVisibleFreezeCounts(currentModel);
    return currentModel.page.frozenRows.filter((row) => row.rowNumber <= rowCount);
}

function getScrollableRowsForCurrentView(currentModel: EditorRenderModel): EditorGridRowView[] {
    const { rowCount } = getVisibleFreezeCounts(currentModel);
    return currentModel.page.rows.filter((row) => row.rowNumber > rowCount);
}

function getCellView(rowNumber: number, columnNumber: number): EditorGridCellView | null {
    const row = getRenderableRows(model).find((item) => item.rowNumber === rowNumber);
    return row?.cells[columnNumber - 1] ?? null;
}

function isGridCellEditable(cell: EditorGridCellView | null): boolean {
    return Boolean(model?.canEdit && !cell?.formula);
}

function canEditCellAt(rowNumber: number, columnNumber: number): boolean {
    return isGridCellEditable(getCellView(rowNumber, columnNumber));
}

function getCellAddressLabel(rowNumber: number, columnNumber: number): string {
    return `${getColumnLabel(columnNumber)}${rowNumber}`;
}

function getSelectionRange(): CellRange | null {
    if (!selectedCell) {
        return null;
    }

    const anchor = selectionAnchorCell ?? selectedCell;
    return {
        startRow: Math.min(anchor.rowNumber, selectedCell.rowNumber),
        endRow: Math.max(anchor.rowNumber, selectedCell.rowNumber),
        startColumn: Math.min(anchor.columnNumber, selectedCell.columnNumber),
        endColumn: Math.max(anchor.columnNumber, selectedCell.columnNumber),
    };
}

function hasExpandedSelection(range: CellRange | null = getSelectionRange()): boolean {
    return Boolean(
        range && (range.startRow !== range.endRow || range.startColumn !== range.endColumn)
    );
}

function isRowInSelection(rowNumber: number): boolean {
    const range = getSelectionRange();
    return Boolean(range && rowNumber >= range.startRow && rowNumber <= range.endRow);
}

function isColumnInSelection(columnNumber: number): boolean {
    const range = getSelectionRange();
    return Boolean(range && columnNumber >= range.startColumn && columnNumber <= range.endColumn);
}

function isCellInSelection(rowNumber: number, columnNumber: number): boolean {
    const range = getSelectionRange();
    return Boolean(
        range &&
        rowNumber >= range.startRow &&
        rowNumber <= range.endRow &&
        columnNumber >= range.startColumn &&
        columnNumber <= range.endColumn
    );
}

function getSelectionOutlineClasses(
    rowNumber: number,
    columnNumber: number,
    range: CellRange | null
): Array<string | false> {
    if (!range) {
        return [];
    }

    return [
        rowNumber === range.startRow && "grid__cell--selection-top",
        rowNumber === range.endRow && "grid__cell--selection-bottom",
        columnNumber === range.startColumn && "grid__cell--selection-left",
        columnNumber === range.endColumn && "grid__cell--selection-right",
    ];
}

function getSelectionStartCell(): CellPosition | null {
    const range = getSelectionRange();
    if (!range) {
        return null;
    }

    return {
        rowNumber: range.startRow,
        columnNumber: range.startColumn,
    };
}

function isActiveSelectionCell(rowNumber: number, columnNumber: number): boolean {
    return selectedCell?.rowNumber === rowNumber && selectedCell.columnNumber === columnNumber;
}

function getEffectiveCellValue(rowNumber: number, columnNumber: number): string {
    if (!model) {
        return "";
    }

    return (
        pendingEdits.get(getPendingEditKey(model.activeSheet.key, rowNumber, columnNumber))
            ?.value ?? getCellModelValue(rowNumber, columnNumber)
    );
}

function serializeSelectionToClipboard(): string | null {
    const range = getSelectionRange();
    if (!range) {
        return null;
    }

    const rows: string[] = [];
    for (let rowNumber = range.startRow; rowNumber <= range.endRow; rowNumber += 1) {
        const cells: string[] = [];
        for (
            let columnNumber = range.startColumn;
            columnNumber <= range.endColumn;
            columnNumber += 1
        ) {
            cells.push(getEffectiveCellValue(rowNumber, columnNumber));
        }

        rows.push(cells.join("\t"));
    }

    return rows.join("\n");
}

function getGridWidth(grid: string[][]): number {
    return grid.reduce((maxWidth, row) => Math.max(maxWidth, row.length), 0);
}

function getPasteGridForSelection(grid: string[][], selectionRange: CellRange): string[][] {
    if (!hasExpandedSelection(selectionRange)) {
        return grid;
    }

    const selectionHeight = selectionRange.endRow - selectionRange.startRow + 1;
    const selectionWidth = selectionRange.endColumn - selectionRange.startColumn + 1;
    const gridWidth = getGridWidth(grid);

    if (grid.length === 1 && gridWidth === 1) {
        const value = grid[0]?.[0] ?? "";
        return Array.from({ length: selectionHeight }, () =>
            Array.from({ length: selectionWidth }, () => value)
        );
    }

    if (grid.length === selectionHeight && gridWidth === selectionWidth) {
        return Array.from({ length: selectionHeight }, (_, rowIndex) =>
            Array.from(
                { length: selectionWidth },
                (_, columnIndex) => grid[rowIndex]?.[columnIndex] ?? ""
            )
        );
    }

    return grid;
}

function getSelectedCellAddress(): string {
    const range = getSelectionRange();
    if (!range || !selectedCell) {
        return STRINGS.noCellSelected;
    }

    if (hasExpandedSelection(range)) {
        return `${getCellAddressLabel(range.startRow, range.startColumn)}:${getCellAddressLabel(range.endRow, range.endColumn)}`;
    }

    const cell = getCellView(selectedCell.rowNumber, selectedCell.columnNumber);
    return cell?.address ?? getCellAddressLabel(selectedCell.rowNumber, selectedCell.columnNumber);
}

function getViewLockTarget(): { rowCount: number; columnCount: number } | null {
    if (!model) {
        return null;
    }

    const anchorCell =
        selectedCell ??
        (model.selection
            ? {
                  rowNumber: model.selection.rowNumber,
                  columnNumber: model.selection.columnNumber,
              }
            : null) ??
        getSelectionStartCell() ??
        (getFirstRenderableRow(model)
            ? {
                  rowNumber: getFirstRenderableRow(model)!.rowNumber,
                  columnNumber: 1,
              }
            : null);
    if (!anchorCell) {
        return null;
    }

    return getFreezePaneCountsForCell(anchorCell);
}

function isViewLocked(): boolean {
    return hasLockedView(model?.activeSheet.freezePane);
}

function canToggleViewLock(): boolean {
    if (!model?.canEdit) {
        return false;
    }

    if (isViewLocked()) {
        return true;
    }

    const target = getViewLockTarget();
    return Boolean(target && (target.rowCount > 0 || target.columnCount > 0));
}

function toggleViewLock(): void {
    if (!model?.canEdit) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    const target = getViewLockTarget();
    if (!isViewLocked() && (!target || (target.rowCount === 0 && target.columnCount === 0))) {
        renderApp({ commitEditing: false });
        return;
    }

    vscode.postMessage({
        type: "toggleViewLock",
        rowCount: target?.rowCount ?? 0,
        columnCount: target?.columnCount ?? 0,
    });
}

function closeTabContextMenu({ refresh = true }: { refresh?: boolean } = {}): void {
    if (!tabContextMenu) {
        return;
    }

    tabContextMenu = null;
    if (refresh) {
        renderApp({ commitEditing: false });
    }
}

function openTabContextMenu(sheetKey: string, x: number, y: number): void {
    tabContextMenu = { sheetKey, x, y };
    renderApp({ commitEditing: false });
}

function requestAddSheet(): void {
    closeTabContextMenu({ refresh: false });
    vscode.postMessage({ type: "addSheet" });
}

function requestDeleteSheet(sheetKey: string): void {
    closeTabContextMenu({ refresh: false });
    vscode.postMessage({ type: "deleteSheet", sheetKey });
}

function requestRenameSheet(sheetKey: string): void {
    closeTabContextMenu({ refresh: false });
    vscode.postMessage({ type: "renameSheet", sheetKey });
}

function getCellPositionFromElement(element: Element | null): CellPosition | null {
    const cell = element?.closest<HTMLElement>('[data-role="grid-cell"]');
    if (!cell) {
        return null;
    }

    const rowNumber = Number(cell.dataset.rowNumber);
    const columnNumber = Number(cell.dataset.columnNumber);
    if (!Number.isInteger(rowNumber) || !Number.isInteger(columnNumber)) {
        return null;
    }

    return { rowNumber, columnNumber };
}

function startSelectionDrag(pointerId: number, anchorCell: CellPosition): void {
    selectionDragState = {
        anchorCell,
        pointerId,
    };
    suppressNextCellClick = false;
}

function stopSelectionDrag(pointerId?: number): void {
    if (!selectionDragState) {
        return;
    }

    if (pointerId !== undefined && selectionDragState.pointerId !== pointerId) {
        return;
    }

    selectionDragState = null;
    if (selectedCell) {
        syncSelectedCellToHost();
    }
}

function updateSelectionDrag(targetCell: CellPosition): void {
    if (!selectionDragState) {
        return;
    }

    if (isActiveSelectionCell(targetCell.rowNumber, targetCell.columnNumber)) {
        return;
    }

    suppressNextCellClick = true;
    setSelectedCellLocal(targetCell, {
        reveal: false,
        syncHost: false,
        anchorCell: selectionDragState.anchorCell,
    });
}

function clampScrollPosition(value: number, maxValue: number): number {
    return Math.max(0, Math.min(value, Math.max(maxValue, 0)));
}

function getPaneScrollState(): ScrollState | null {
    const pane =
        document.querySelector<HTMLElement>('[data-role="grid-scroll-main"]') ??
        document.querySelector<HTMLElement>(".pane__table");
    if (!pane) {
        return null;
    }

    return {
        top: pane.scrollTop,
        left: pane.scrollLeft,
    };
}

function restorePaneScrollState(scrollState: ScrollState | null): void {
    if (!scrollState) {
        return;
    }

    const pane =
        document.querySelector<HTMLElement>('[data-role="grid-scroll-main"]') ??
        document.querySelector<HTMLElement>(".pane__table");
    if (!pane) {
        return;
    }

    pane.scrollTop = clampScrollPosition(scrollState.top, pane.scrollHeight - pane.clientHeight);
    pane.scrollLeft = clampScrollPosition(scrollState.left, pane.scrollWidth - pane.clientWidth);
    lastObservedPaneScrollTop = pane.scrollTop;

    if (pane.dataset.pane === "bottom-right") {
        syncFrozenPaneScroll("bottom-right", pane);
    }
}

function clampGridColumnWidth(width: number): number {
    return Math.max(EDITOR_COLUMN_MIN_WIDTH, Math.min(width, EDITOR_COLUMN_MAX_WIDTH));
}

function applyCachedGridColumnWidths(): void {
    for (const column of document.querySelectorAll<HTMLElement>(
        '[data-role="grid-column-col"][data-column-number]'
    )) {
        const columnNumber = Number(column.dataset.columnNumber);
        if (!Number.isInteger(columnNumber)) {
            continue;
        }

        const width = columnWidthByNumber.get(columnNumber);
        column.style.width = width ? `${width}px` : "";
    }
}

function hasMeasuredVisibleGridColumnWidths(): boolean {
    const columnElements = document.querySelectorAll<HTMLElement>(
        '[data-role="grid-column-col"][data-column-number]'
    );
    if (columnElements.length === 0) {
        return true;
    }

    for (const column of columnElements) {
        const columnNumber = Number(column.dataset.columnNumber);
        if (!Number.isInteger(columnNumber)) {
            continue;
        }

        if (!columnWidthByNumber.has(columnNumber)) {
            return false;
        }
    }

    return true;
}

function measureVisibleGridColumnWidths(): void {
    const columnElements = Array.from(
        document.querySelectorAll<HTMLElement>('[data-role="grid-column-col"][data-column-number]')
    );

    for (const column of columnElements) {
        column.style.width = "";
    }

    for (const element of document.querySelectorAll<HTMLElement>(
        'th.grid__column[data-column-number], td.grid__cell[data-column-number]'
    )) {
        const columnNumber = Number(element.dataset.columnNumber);
        if (!Number.isInteger(columnNumber)) {
            continue;
        }

        const width = clampGridColumnWidth(Math.ceil(element.getBoundingClientRect().width));
        if (width <= 0) {
            continue;
        }

        columnWidthByNumber.set(
            columnNumber,
            Math.max(columnWidthByNumber.get(columnNumber) ?? 0, width)
        );
    }

    applyCachedGridColumnWidths();
}

function setGridColumnWidths({
    measureColumns = true,
}: {
    measureColumns?: boolean;
} = {}): void {
    const rowHeaderColumns = Array.from(
        document.querySelectorAll<HTMLElement>('[data-role="grid-row-header-col"]')
    );
    const rowHeaderCells = Array.from(
        document.querySelectorAll<HTMLElement>('th.grid__row-number, td.grid__row-number')
    );
    for (const column of rowHeaderColumns) {
        column.style.width = "";
    }

    const rowHeaderWidth = Math.ceil(
        Math.max(56, ...rowHeaderCells.map((cell) => cell.getBoundingClientRect().width), 0)
    );
    for (const column of rowHeaderColumns) {
        column.style.width = `${rowHeaderWidth}px`;
    }

    if (measureColumns || !hasMeasuredVisibleGridColumnWidths()) {
        measureVisibleGridColumnWidths();
        return;
    }

    applyCachedGridColumnWidths();
}

function cancelScheduledGridLayoutSync(): void {
    if (frozenPaneLayoutFrame) {
        cancelAnimationFrame(frozenPaneLayoutFrame);
        frozenPaneLayoutFrame = 0;
    }

    if (deferredGridLayoutTimer) {
        clearTimeout(deferredGridLayoutTimer);
        deferredGridLayoutTimer = 0;
    }

    if (pendingViewportScrollFrame) {
        cancelAnimationFrame(pendingViewportScrollFrame);
        pendingViewportScrollFrame = 0;
    }

    pendingViewportScrollPane = null;
}

function performGridLayoutSync({
    revealSelection = false,
    viewLocked = false,
    forceColumnMeasurement = false,
}: {
    revealSelection?: boolean;
    viewLocked?: boolean;
    forceColumnMeasurement?: boolean;
} = {}): void {
    const shouldMeasureColumns =
        forceColumnMeasurement || !isActivelyScrolling() || !hasMeasuredVisibleGridColumnWidths();

    setGridColumnWidths({ measureColumns: shouldMeasureColumns });
    if (viewLocked) {
        syncFrozenPaneRowHeights();
    }
    cacheRenderedRowHeights();
    if (revealSelection) {
        revealSelectedCell();
    }

    if (shouldMeasureColumns) {
        return;
    }

    if (deferredGridLayoutTimer) {
        clearTimeout(deferredGridLayoutTimer);
    }

    deferredGridLayoutTimer = window.setTimeout(() => {
        deferredGridLayoutTimer = 0;
        scheduleGridLayoutSync({ viewLocked, forceColumnMeasurement: true });
    }, ACTIVE_SCROLL_LAYOUT_DELAY_MS);
}

function scheduleGridLayoutSync({
    revealSelection = false,
    viewLocked = false,
    forceColumnMeasurement = false,
}: {
    revealSelection?: boolean;
    viewLocked?: boolean;
    forceColumnMeasurement?: boolean;
} = {}): void {
    if (frozenPaneLayoutFrame) {
        cancelAnimationFrame(frozenPaneLayoutFrame);
    }

    frozenPaneLayoutFrame = requestAnimationFrame(() => {
        frozenPaneLayoutFrame = 0;
        performGridLayoutSync({
            revealSelection,
            viewLocked,
            forceColumnMeasurement,
        });
    });
}

function resetGridMeasurementState(): void {
    cancelScheduledGridLayoutSync();
    lastObservedPaneScrollTop = null;
    lastMeasuredSheetKey = null;
    rowHeightByNumber.clear();
    columnWidthByNumber.clear();
}

function scheduleGridMetricRefresh({
    revealSelection = false,
    viewLocked = false,
    forceColumnMeasurement = false,
}: {
    revealSelection?: boolean;
    viewLocked?: boolean;
    forceColumnMeasurement?: boolean;
} = {}): void {
    clearFrozenPaneLayout();
    scheduleGridLayoutSync({
        revealSelection,
        viewLocked,
        forceColumnMeasurement,
    });
}

function syncGridColumnWidths(): void {
    setGridColumnWidths({ measureColumns: true });
}

function scheduleFrozenPaneLayoutSync({
    revealSelection = false,
    forceColumnMeasurement = false,
}: {
    revealSelection?: boolean;
    forceColumnMeasurement?: boolean;
} = {}): void {
    scheduleGridMetricRefresh({
        revealSelection,
        viewLocked: hasLockedView(model?.activeSheet.freezePane),
        forceColumnMeasurement,
    });
}

function cacheRenderedRowHeights(): void {
    const nextHeights = new Map<number, number>();

    for (const row of document.querySelectorAll<HTMLTableRowElement>('[data-role="grid-row"]')) {
        const rowNumber = Number(row.dataset.rowNumber);
        if (!Number.isInteger(rowNumber)) {
            continue;
        }

        const rowHeight = Math.ceil(row.getBoundingClientRect().height);
        if (rowHeight <= 0) {
            continue;
        }

        nextHeights.set(rowNumber, Math.max(nextHeights.get(rowNumber) ?? 0, rowHeight));
    }

    for (const [rowNumber, rowHeight] of nextHeights) {
        rowHeightByNumber.set(rowNumber, rowHeight);
    }
}

function clearFrozenPaneLayout(): void {
    for (const element of document.querySelectorAll<HTMLElement>(
        ".grid__column--frozen, .grid__row-number--frozen, .grid__cell--frozen-row, .grid__cell--frozen-column, .grid__cell--frozen-intersection"
    )) {
        element.classList.remove(
            "grid__column--frozen",
            "grid__row-number--frozen",
            "grid__cell--frozen-row",
            "grid__cell--frozen-column",
            "grid__cell--frozen-intersection"
        );
        element.style.removeProperty("--grid-freeze-top");
        element.style.removeProperty("--grid-freeze-left");
        element.style.removeProperty("--grid-freeze-background");
    }
}

function getFrozenPaneBackgroundColor(element: HTMLElement, fallbackColor: string): string {
    const backgroundColor = globalThis.getComputedStyle(element).backgroundColor;
    if (backgroundColor === "rgba(0, 0, 0, 0)" || backgroundColor === "transparent") {
        return fallbackColor;
    }

    return backgroundColor;
}

function applyFrozenPaneLayout(freezePane: EditorRenderModel["activeSheet"]["freezePane"]): void {
    clearFrozenPaneLayout();

    if (!freezePane || (freezePane.columnCount <= 0 && freezePane.rowCount <= 0)) {
        return;
    }

    const pane = document.querySelector<HTMLElement>(".pane__table");
    const table = pane?.querySelector<HTMLTableElement>(".grid");
    if (!pane || !table) {
        return;
    }

    const fallbackColor = globalThis.getComputedStyle(pane).backgroundColor;
    const headerRow = table.querySelector<HTMLTableRowElement>("thead tr");
    const cornerHeader = table.querySelector<HTMLElement>("thead th.grid__row-number");
    const headerHeight = headerRow?.getBoundingClientRect().height ?? 0;
    const rowHeaderWidth = cornerHeader?.getBoundingClientRect().width ?? 0;

    const frozenColumnOffsets = new Map<number, number>();
    let currentLeft = rowHeaderWidth;
    for (const header of table.querySelectorAll<HTMLElement>(
        "thead th.grid__column[data-column-number]"
    )) {
        const columnNumber = Number(header.dataset.columnNumber);
        if (!Number.isInteger(columnNumber) || columnNumber > freezePane.columnCount) {
            continue;
        }

        frozenColumnOffsets.set(columnNumber, currentLeft);
        header.classList.add("grid__column--frozen");
        header.style.setProperty("--grid-freeze-left", `${currentLeft}px`);
        header.style.setProperty(
            "--grid-freeze-background",
            getFrozenPaneBackgroundColor(header, fallbackColor)
        );
        currentLeft += header.getBoundingClientRect().width;
    }

    const frozenRowOffsets = new Map<number, number>();
    let currentTop = headerHeight;
    for (const row of table.querySelectorAll<HTMLTableRowElement>(
        'tbody tr[data-role="grid-row"]'
    )) {
        const rowNumber = Number(row.dataset.rowNumber);
        if (!Number.isInteger(rowNumber) || rowNumber > freezePane.rowCount) {
            continue;
        }

        frozenRowOffsets.set(rowNumber, currentTop);
        const rowHeader = row.querySelector<HTMLElement>('th[data-role="grid-row-header"]');
        if (rowHeader) {
            rowHeader.classList.add("grid__row-number--frozen");
            rowHeader.style.setProperty("--grid-freeze-top", `${currentTop}px`);
            rowHeader.style.setProperty(
                "--grid-freeze-background",
                getFrozenPaneBackgroundColor(rowHeader, fallbackColor)
            );
        }

        currentTop += row.getBoundingClientRect().height;
    }

    for (const cell of table.querySelectorAll<HTMLElement>('td[data-role="grid-cell"]')) {
        const rowNumber = Number(cell.dataset.rowNumber);
        const columnNumber = Number(cell.dataset.columnNumber);
        const topOffset = frozenRowOffsets.get(rowNumber);
        const leftOffset = frozenColumnOffsets.get(columnNumber);

        if (topOffset === undefined && leftOffset === undefined) {
            continue;
        }

        if (topOffset !== undefined) {
            cell.classList.add("grid__cell--frozen-row");
            cell.style.setProperty("--grid-freeze-top", `${topOffset}px`);
        }

        if (leftOffset !== undefined) {
            cell.classList.add("grid__cell--frozen-column");
            cell.style.setProperty("--grid-freeze-left", `${leftOffset}px`);
        }

        if (topOffset !== undefined && leftOffset !== undefined) {
            cell.classList.add("grid__cell--frozen-intersection");
        }

        cell.style.setProperty(
            "--grid-freeze-background",
            getFrozenPaneBackgroundColor(cell, fallbackColor)
        );
    }
}

function getFrozenPaneElement(name: FrozenPaneName): HTMLElement | null {
    return document.querySelector<HTMLElement>(`[data-pane="${name}"]`);
}

function syncFrozenPaneScroll(
    sourcePaneName: FrozenPaneName,
    sourcePane: HTMLElement,
    { force = false }: { force?: boolean } = {}
): void {
    if (isSyncingFrozenPaneScroll && !force) {
        return;
    }

    const topRightPane = getFrozenPaneElement("top-right");
    const bottomLeftPane = getFrozenPaneElement("bottom-left");
    const bottomRightPane = getFrozenPaneElement("bottom-right");
    if (!bottomRightPane) {
        return;
    }

    isSyncingFrozenPaneScroll = true;

    if (sourcePaneName === "top-right") {
        bottomRightPane.scrollLeft = clampScrollPosition(
            sourcePane.scrollLeft,
            bottomRightPane.scrollWidth - bottomRightPane.clientWidth
        );
    } else if (sourcePaneName === "bottom-left") {
        bottomRightPane.scrollTop = clampScrollPosition(
            sourcePane.scrollTop,
            bottomRightPane.scrollHeight - bottomRightPane.clientHeight
        );
    } else if (sourcePaneName === "bottom-right") {
        if (topRightPane) {
            topRightPane.scrollLeft = clampScrollPosition(
                sourcePane.scrollLeft,
                topRightPane.scrollWidth - topRightPane.clientWidth
            );
        }

        if (bottomLeftPane) {
            bottomLeftPane.scrollTop = clampScrollPosition(
                sourcePane.scrollTop,
                bottomLeftPane.scrollHeight - bottomLeftPane.clientHeight
            );
        }
    }

    requestAnimationFrame(() => {
        isSyncingFrozenPaneScroll = false;
    });
}

function normalizeWheelDelta(delta: number, deltaMode: number, viewportSize: number): number {
    if (deltaMode === WHEEL_DELTA_LINE_MODE) {
        return delta * WHEEL_LINE_SCROLL_PIXELS;
    }

    if (deltaMode === WHEEL_DELTA_PAGE_MODE) {
        return delta * viewportSize;
    }

    return delta;
}

function forwardFrozenPaneWheel(event: React.WheelEvent<HTMLElement>): void {
    const bottomRightPane = getFrozenPaneElement("bottom-right");
    if (!bottomRightPane) {
        return;
    }

    let deltaX = normalizeWheelDelta(
        event.deltaX,
        event.deltaMode,
        bottomRightPane.clientWidth
    );
    let deltaY = normalizeWheelDelta(
        event.deltaY,
        event.deltaMode,
        bottomRightPane.clientHeight
    );

    if (event.shiftKey && deltaX === 0 && deltaY !== 0) {
        deltaX = deltaY;
        deltaY = 0;
    }

    if (deltaX === 0 && deltaY === 0) {
        return;
    }

    event.preventDefault();

    const nextLeft = clampScrollPosition(
        bottomRightPane.scrollLeft + deltaX,
        bottomRightPane.scrollWidth - bottomRightPane.clientWidth
    );
    const nextTop = clampScrollPosition(
        bottomRightPane.scrollTop + deltaY,
        bottomRightPane.scrollHeight - bottomRightPane.clientHeight
    );

    bottomRightPane.scrollLeft = nextLeft;
    bottomRightPane.scrollTop = nextTop;
    syncFrozenPaneScroll("bottom-right", bottomRightPane, { force: true });
    scheduleViewportRequestForScroll(bottomRightPane);
}

function syncFrozenPaneRowHeights(): void {
    const rowFragments = Array.from(
        document.querySelectorAll<HTMLTableRowElement>('[data-role="grid-row"]')
    );
    if (rowFragments.length === 0) {
        return;
    }

    const headerRows = Array.from(
        document.querySelectorAll<HTMLTableRowElement>('[data-role="grid-header-row"]')
    );
    for (const row of [...headerRows, ...rowFragments]) {
        row.style.height = "";
    }

    if (headerRows.length > 1) {
        const headerHeight = Math.ceil(
            Math.max(...headerRows.map((row) => row.getBoundingClientRect().height))
        );
        if (headerHeight > 0) {
            for (const row of headerRows) {
                row.style.height = `${headerHeight}px`;
            }
        }
    }

    const rowGroups = new Map<number, HTMLTableRowElement[]>();
    for (const row of rowFragments) {
        const rowNumber = Number(row.dataset.rowNumber);
        if (!Number.isInteger(rowNumber)) {
            continue;
        }

        const entries = rowGroups.get(rowNumber) ?? [];
        entries.push(row);
        rowGroups.set(rowNumber, entries);
    }

    for (const [rowNumber, rows] of rowGroups) {
        const rowHeight = Math.ceil(Math.max(...rows.map((row) => row.getBoundingClientRect().height)));
        if (rowHeight <= 0) {
            continue;
        }

        for (const row of rows) {
            row.style.height = `${rowHeight}px`;
        }
    }
}

function getSelectedCellPaneName(
    currentModel: EditorRenderModel | null,
    cell: CellPosition
): FrozenPaneName | null {
    if (!currentModel || !hasLockedView(currentModel.activeSheet.freezePane)) {
        return null;
    }

    const { rowCount, columnCount } = getVisibleFreezeCounts(currentModel);
    const isFrozenRow = cell.rowNumber <= rowCount;
    const isFrozenColumn = cell.columnNumber <= columnCount;

    if (isFrozenRow && isFrozenColumn) {
        return "top-left";
    }

    if (isFrozenRow) {
        return "top-right";
    }

    if (isFrozenColumn) {
        return "bottom-left";
    }

    return "bottom-right";
}

function getStickyPaneInsets(pane: HTMLElement): { top: number; left: number } {
    const paneRect = pane.getBoundingClientRect();
    const headerRow = pane.querySelector("thead tr");
    const firstColumn = pane.querySelector("thead th:first-child");
    let top = headerRow?.getBoundingClientRect().height ?? 0;
    let left = firstColumn?.getBoundingClientRect().width ?? 0;

    for (const rowHeader of pane.querySelectorAll<HTMLElement>("th.grid__row-number--frozen")) {
        top = Math.max(top, rowHeader.getBoundingClientRect().bottom - paneRect.top);
    }

    for (const columnHeader of pane.querySelectorAll<HTMLElement>(
        "thead th.grid__column--frozen"
    )) {
        left = Math.max(left, columnHeader.getBoundingClientRect().right - paneRect.left);
    }

    return {
        top,
        left,
    };
}

function revealSelectedCell(): void {
    if (!selectedCell) {
        return;
    }

    const paneName = getSelectedCellPaneName(model, selectedCell);
    const pane =
        (paneName ? getFrozenPaneElement(paneName) : null) ??
        document.querySelector<HTMLElement>(".pane__table");
    const target = pane?.querySelector<HTMLElement>(
        `[data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"][data-column-number="${selectedCell.columnNumber}"]`
    );
    if (!pane || !target) {
        return;
    }

    const paneRect = pane.getBoundingClientRect();
    const elementRect = target.getBoundingClientRect();
    const stickyInsets = paneName ? { top: 0, left: 0 } : getStickyPaneInsets(pane);
    let top = pane.scrollTop;
    let left = pane.scrollLeft;

    if (elementRect.top < paneRect.top + stickyInsets.top) {
        top -= paneRect.top + stickyInsets.top - elementRect.top;
    } else if (elementRect.bottom > paneRect.bottom) {
        top += elementRect.bottom - paneRect.bottom;
    }

    if (elementRect.left < paneRect.left + stickyInsets.left) {
        left -= paneRect.left + stickyInsets.left - elementRect.left;
    } else if (elementRect.right > paneRect.right) {
        left += elementRect.right - paneRect.right;
    }

    pane.scrollTop = clampScrollPosition(top, pane.scrollHeight - pane.clientHeight);
    pane.scrollLeft = clampScrollPosition(left, pane.scrollWidth - pane.clientWidth);

    if (paneName) {
        syncFrozenPaneScroll(paneName, pane);
    }
}

function isCellVisible(cell: CellPosition | null): boolean {
    if (!cell || !model) {
        return false;
    }

    return getRenderableRows(model).some(
        (row) => row.rowNumber === cell.rowNumber && cell.columnNumber <= row.cells.length
    );
}

function syncSelectedCellToHost(): void {
    if (!model || !selectedCell) {
        return;
    }

    vscode.postMessage({
        type: "selectCell",
        rowNumber: selectedCell.rowNumber,
        columnNumber: selectedCell.columnNumber,
    });
}

function isSimpleSelection(
    cell: CellPosition | null,
    anchorCell: CellPosition | null = selectionAnchorCell
): boolean {
    if (!cell) {
        return anchorCell === null;
    }

    const anchor = anchorCell ?? cell;
    return anchor.rowNumber === cell.rowNumber && anchor.columnNumber === cell.columnNumber;
}

function getRenderedCellElements(
    cell: Pick<CellPosition, "rowNumber" | "columnNumber"> | null
): HTMLElement[] {
    if (!cell) {
        return [];
    }

    return Array.from(
        document.querySelectorAll<HTMLElement>(
            `[data-role="grid-cell"][data-row-number="${cell.rowNumber}"][data-column-number="${cell.columnNumber}"]`
        )
    );
}

function setRenderedPrimarySelectionState(
    cell: Pick<CellPosition, "rowNumber" | "columnNumber"> | null,
    isSelected: boolean
): void {
    for (const element of getRenderedCellElements(cell)) {
        element.classList.toggle("grid__cell--selected-range", isSelected);
        element.classList.toggle("grid__cell--selected", isSelected);
        element.classList.toggle("grid__cell--selection-top", isSelected);
        element.classList.toggle("grid__cell--selection-right", isSelected);
        element.classList.toggle("grid__cell--selection-bottom", isSelected);
        element.classList.toggle("grid__cell--selection-left", isSelected);
        element.setAttribute("aria-selected", String(isSelected));
    }
}

function setRenderedRowSelectionState(rowNumber: number, isActive: boolean): void {
    for (const header of document.querySelectorAll<HTMLElement>(
        `[data-role="grid-row-header"][data-row-number="${rowNumber}"]`
    )) {
        header.classList.toggle("grid__row-number--active", isActive);
    }

    for (const cell of document.querySelectorAll<HTMLElement>(
        `[data-role="grid-cell"][data-row-number="${rowNumber}"]`
    )) {
        cell.classList.toggle("grid__cell--active-row", isActive);
    }
}

function setRenderedColumnSelectionState(columnNumber: number, isActive: boolean): void {
    for (const header of document.querySelectorAll<HTMLElement>(
        `th.grid__column[data-column-number="${columnNumber}"]`
    )) {
        header.classList.toggle("grid__column--active", isActive);
    }

    for (const cell of document.querySelectorAll<HTMLElement>(
        `[data-role="grid-cell"][data-column-number="${columnNumber}"]`
    )) {
        cell.classList.toggle("grid__cell--active-column", isActive);
    }
}

function updateSelectedCellAddressBadge(): void {
    const badge = document.querySelector<HTMLElement>('[data-role="selected-cell-address"]');
    if (!badge) {
        return;
    }

    const address = getSelectedCellAddress();
    badge.textContent = address;
    badge.title = `${STRINGS.selectedCell}: ${address}`;
}

function syncLocalSimpleSelectionDom(
    previousCell: CellPosition | null,
    nextCell: CellPosition | null
): void {
    if (
        previousCell &&
        (!nextCell ||
            previousCell.rowNumber !== nextCell.rowNumber ||
            previousCell.columnNumber !== nextCell.columnNumber)
    ) {
        setRenderedPrimarySelectionState(previousCell, false);
    }

    if (previousCell && (!nextCell || previousCell.rowNumber !== nextCell.rowNumber)) {
        setRenderedRowSelectionState(previousCell.rowNumber, false);
    }

    if (previousCell && (!nextCell || previousCell.columnNumber !== nextCell.columnNumber)) {
        setRenderedColumnSelectionState(previousCell.columnNumber, false);
    }

    if (
        nextCell &&
        (!previousCell ||
            previousCell.rowNumber !== nextCell.rowNumber ||
            previousCell.columnNumber !== nextCell.columnNumber)
    ) {
        setRenderedPrimarySelectionState(nextCell, true);
    }

    if (nextCell && (!previousCell || previousCell.rowNumber !== nextCell.rowNumber)) {
        setRenderedRowSelectionState(nextCell.rowNumber, true);
    }

    if (nextCell && (!previousCell || previousCell.columnNumber !== nextCell.columnNumber)) {
        setRenderedColumnSelectionState(nextCell.columnNumber, true);
    }

    updateSelectedCellAddressBadge();
}

function canUseLockedViewLocalSelectionUpdate(
    nextCell: CellPosition | null,
    {
        reveal,
        anchorCell,
    }: {
        reveal: boolean;
        anchorCell: CellPosition | null;
    }
): boolean {
    return Boolean(
        nextCell &&
            model &&
            hasLockedView(model.activeSheet.freezePane) &&
            !editingCell &&
            !reveal &&
            isCellVisible(nextCell) &&
            !hasExpandedSelection() &&
            isSimpleSelection(nextCell, anchorCell)
    );
}

function setSelectedCellLocal(
    nextCell: CellPosition | null,
    {
        reveal = false,
        syncHost = true,
        anchorCell,
    }: {
        reveal?: boolean;
        syncHost?: boolean;
        anchorCell?: CellPosition | null;
    } = {}
): void {
    const previousCell = selectedCell;
    const nextAnchorCell = nextCell ? (anchorCell ?? nextCell) : null;
    const canUseLocalUpdate = canUseLockedViewLocalSelectionUpdate(nextCell, {
        reveal,
        anchorCell: nextAnchorCell,
    });

    selectedCell = nextCell;
    selectionAnchorCell = nextAnchorCell;
    suppressAutoSelection = nextCell === null;
    clearBrowserTextSelection();

    if (syncHost) {
        syncSelectedCellToHost();
    }

    if (canUseLocalUpdate) {
        syncLocalSimpleSelectionDom(previousCell, nextCell);
        return;
    }

    renderApp({ commitEditing: false, revealSelection: reveal });
}

function prepareSelectionForRender({
    revealSelection,
    useModelSelection,
}: {
    revealSelection: boolean;
    useModelSelection: boolean;
}): boolean {
    let shouldReveal = revealSelection;

    if (
        useModelSelection &&
        model?.selection &&
        isCellVisible({
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        })
    ) {
        selectedCell = {
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        };
        selectionAnchorCell = selectedCell;
        suppressAutoSelection = false;
        pendingSelectionAfterRender = null;
        syncSelectedCellToHost();
        return shouldReveal;
    }

    if (pendingSelectionAfterRender && isCellVisible(pendingSelectionAfterRender)) {
        selectedCell = {
            rowNumber: pendingSelectionAfterRender.rowNumber,
            columnNumber: pendingSelectionAfterRender.columnNumber,
        };
        selectionAnchorCell = selectedCell;
        suppressAutoSelection = false;
        shouldReveal = pendingSelectionAfterRender.reveal;
        pendingSelectionAfterRender = null;
        syncSelectedCellToHost();
        return shouldReveal;
    }

    pendingSelectionAfterRender = null;

    if (isCellVisible(selectedCell)) {
        if (!isCellVisible(selectionAnchorCell)) {
            selectionAnchorCell = selectedCell;
        }
        syncSelectedCellToHost();
        return shouldReveal;
    }

    if (suppressAutoSelection) {
        return shouldReveal;
    }

    selectedCell = model?.selection
        ? {
              rowNumber: model.selection.rowNumber,
              columnNumber: model.selection.columnNumber,
          }
        : getFirstRenderableRow(model)
          ? { rowNumber: getFirstRenderableRow(model)!.rowNumber, columnNumber: 1 }
          : null;

    if (selectedCell) {
        selectionAnchorCell = selectedCell;
        suppressAutoSelection = false;
        syncSelectedCellToHost();
    }

    return shouldReveal;
}

function getCellModelValue(rowNumber: number, columnNumber: number): string {
    return getCellView(rowNumber, columnNumber)?.value ?? "";
}

function getCellFormula(rowNumber: number, columnNumber: number): string | null {
    return getCellView(rowNumber, columnNumber)?.formula ?? null;
}

function notifyPendingEditState(): void {
    const hasPendingEdits = pendingEdits.size > 0;
    if (lastPendingNotification === hasPendingEdits) {
        return;
    }

    lastPendingNotification = hasPendingEdits;
    vscode.postMessage({ type: "pendingEditStateChanged", hasPendingEdits });
}

function syncPendingEditsToHost(): void {
    if (!model) {
        return;
    }

    const edits = Array.from(pendingEdits.values());
    const serializedEdits = serializePendingEdits(edits);
    if (serializedEdits === lastPendingEditsSyncKey) {
        return;
    }

    lastPendingEditsSyncKey = serializedEdits;
    vscode.postMessage({ type: "setPendingEdits", edits });
}

function setPendingCellValue(change: PendingEditChange, value: string): void {
    const key = getPendingEditKey(change.sheetKey, change.rowNumber, change.columnNumber);
    if (value === change.modelValue) {
        pendingEdits.delete(key);
        return;
    }

    pendingEdits.set(key, {
        sheetKey: change.sheetKey,
        rowNumber: change.rowNumber,
        columnNumber: change.columnNumber,
        value,
    });
}

function applyPendingSideEffects(): void {
    notifyPendingEditState();
    syncPendingEditsToHost();
}

function applyEditChanges(
    changes: PendingEditChange[],
    {
        recordHistory = true,
        refresh = true,
        revealSelection = false,
    }: {
        recordHistory?: boolean;
        refresh?: boolean;
        revealSelection?: boolean;
    } = {}
): void {
    const effectiveChanges = changes.filter((change) => change.beforeValue !== change.afterValue);
    if (effectiveChanges.length === 0) {
        notifyPendingEditState();
        if (refresh) {
            renderApp({ commitEditing: false, revealSelection });
        }
        return;
    }

    if (recordHistory) {
        undoStack.push({ changes: effectiveChanges });
        redoStack.length = 0;
    }

    for (const change of effectiveChanges) {
        setPendingCellValue(change, change.afterValue);
    }

    applyPendingSideEffects();

    if (refresh) {
        renderApp({ commitEditing: false, revealSelection });
    }
}

function applyHistoryEntry(entry: HistoryEntry, direction: "undo" | "redo"): void {
    for (const change of entry.changes) {
        setPendingCellValue(change, direction === "undo" ? change.beforeValue : change.afterValue);
    }

    applyPendingSideEffects();
}

function undoPendingEdits(): void {
    const entry = undoStack.pop();
    if (!entry) {
        if (model?.canUndoStructuralEdits) {
            vscode.postMessage({ type: "undoSheetEdit" });
            return;
        }

        renderApp({ commitEditing: false });
        return;
    }

    applyHistoryEntry(entry, "undo");
    redoStack.push(entry);
    renderApp({ commitEditing: false, revealSelection: true });
}

function redoPendingEdits(): void {
    const entry = redoStack.pop();
    if (!entry) {
        if (model?.canRedoStructuralEdits) {
            vscode.postMessage({ type: "redoSheetEdit" });
            return;
        }

        renderApp({ commitEditing: false });
        return;
    }

    applyHistoryEntry(entry, "redo");
    undoStack.push(entry);
    renderApp({ commitEditing: false, revealSelection: true });
}

function commitEdit(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    value: string,
    {
        refresh = true,
        revealSelection = false,
    }: { refresh?: boolean; revealSelection?: boolean } = {}
): void {
    const modelValue = getCellModelValue(rowNumber, columnNumber);
    const beforeValue =
        pendingEdits.get(getPendingEditKey(sheetKey, rowNumber, columnNumber))?.value ?? modelValue;

    applyEditChanges(
        [
            {
                sheetKey,
                rowNumber,
                columnNumber,
                modelValue,
                beforeValue,
                afterValue: value,
            },
        ],
        { refresh, revealSelection }
    );
}

function finishEdit({
    mode,
    clearSelection = false,
    refresh = true,
}: {
    mode: "commit" | "cancel";
    clearSelection?: boolean;
    refresh?: boolean;
}): void {
    const session = editingCell;
    if (!session) {
        return;
    }

    editingCell = null;

    if (mode === "commit") {
        commitEdit(session.sheetKey, session.rowNumber, session.columnNumber, session.value, {
            refresh: false,
        });
    }

    if (clearSelection) {
        selectedCell = null;
        selectionAnchorCell = null;
        suppressAutoSelection = true;
    }

    if (refresh) {
        renderApp({ commitEditing: false });
    }
}

function clearSelectedCellValue(): void {
    const range = getSelectionRange();
    if (!model || !range) {
        return;
    }

    const changes: PendingEditChange[] = [];
    for (let rowNumber = range.startRow; rowNumber <= range.endRow; rowNumber += 1) {
        for (
            let columnNumber = range.startColumn;
            columnNumber <= range.endColumn;
            columnNumber += 1
        ) {
            if (!canEditCellAt(rowNumber, columnNumber)) {
                continue;
            }

            const modelValue = getCellModelValue(rowNumber, columnNumber);
            changes.push({
                sheetKey: model.activeSheet.key,
                rowNumber,
                columnNumber,
                modelValue,
                beforeValue:
                    pendingEdits.get(
                        getPendingEditKey(model.activeSheet.key, rowNumber, columnNumber)
                    )?.value ?? modelValue,
                afterValue: "",
            });
        }
    }

    applyEditChanges(changes, { revealSelection: true });
}

function isClearSelectedCellKey(event: KeyboardEvent): boolean {
    if (event.altKey || event.ctrlKey || event.metaKey) {
        return false;
    }

    return (
        event.key === "Backspace" ||
        event.key === "Delete" ||
        event.code === "Backspace" ||
        event.code === "Delete"
    );
}

function triggerSave(): void {
    if (!model || isSaving) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", clearSelection: true, refresh: false });
    }

    if (!model.hasPendingEdits && pendingEdits.size === 0) {
        renderApp({ commitEditing: false });
        return;
    }

    isSaving = true;
    renderApp({ commitEditing: false });
    vscode.postMessage({ type: "requestSave" });
}

function normalizePastedRows(text: string): string[][] {
    const lines = text.replaceAll("\r", "").split("\n");
    if (lines.length > 0 && lines[lines.length - 1] === "") {
        lines.pop();
    }

    return lines.map((line) => line.split("\t"));
}

function applyPastedGrid(grid: string[][]): void {
    const selectionRange = getSelectionRange();
    const anchorCell = getSelectionStartCell();
    if (!model || !selectionRange || !anchorCell || grid.length === 0 || !model.canEdit) {
        return;
    }

    const pasteGrid = getPasteGridForSelection(grid, selectionRange);
    const maxRow = model.activeSheet.rowCount;
    const maxColumn = model.activeSheet.columnCount;
    const changes: PendingEditChange[] = [];

    for (let rowOffset = 0; rowOffset < pasteGrid.length; rowOffset += 1) {
        const targetRow = anchorCell.rowNumber + rowOffset;
        if (targetRow > maxRow) {
            break;
        }

        const values = pasteGrid[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
            const targetColumn = anchorCell.columnNumber + columnOffset;
            if (targetColumn > maxColumn) {
                break;
            }

            if (!canEditCellAt(targetRow, targetColumn)) {
                continue;
            }

            const modelValue = getCellModelValue(targetRow, targetColumn);
            changes.push({
                sheetKey: model.activeSheet.key,
                rowNumber: targetRow,
                columnNumber: targetColumn,
                modelValue,
                beforeValue:
                    pendingEdits.get(
                        getPendingEditKey(model.activeSheet.key, targetRow, targetColumn)
                    )?.value ?? modelValue,
                afterValue: values[columnOffset] ?? "",
            });
        }
    }

    applyEditChanges(changes, { revealSelection: true });
}

function startEditCell(rowNumber: number, columnNumber: number, currentValue: string): void {
    if (!model) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    clearBrowserTextSelection();
    selectedCell = { rowNumber, columnNumber };
    selectionAnchorCell = selectedCell;
    suppressAutoSelection = false;
    editingCell = {
        sheetKey: model.activeSheet.key,
        rowNumber,
        columnNumber,
        value: currentValue,
    };
    syncSelectedCellToHost();
    renderApp({ commitEditing: false, revealSelection: true });
}

function getSelectionBounds(): {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
} | null {
    if (!model) {
        return null;
    }

    const currentSelectedCell = selectedCell;
    const boundsRows =
        hasFrozenRowGap(model) && currentSelectedCell
            ? getFrozenRowsForCurrentView(model).some(
                  (row) => row.rowNumber === currentSelectedCell.rowNumber
              )
                ? getFrozenRowsForCurrentView(model)
                : model.page.rows
            : getRenderableRows(model);
    if (boundsRows.length === 0) {
        return null;
    }

    return {
        minRow: boundsRows[0]!.rowNumber,
        maxRow: boundsRows[boundsRows.length - 1]!.rowNumber,
        minColumn: 1,
        maxColumn: model.page.columns.length,
    };
}

function ensureSelection(): CellPosition | null {
    if (selectedCell) {
        selectionAnchorCell ??= selectedCell;
        suppressAutoSelection = false;
        return selectedCell;
    }

    if (model?.selection) {
        selectedCell = {
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        };
        selectionAnchorCell = selectedCell;
        suppressAutoSelection = false;
        return selectedCell;
    }

    const firstRow = getFirstRenderableRow(model);
    if (firstRow) {
        selectedCell = {
            rowNumber: firstRow.rowNumber,
            columnNumber: 1,
        };
        selectionAnchorCell = selectedCell;
        suppressAutoSelection = false;
        return selectedCell;
    }

    return null;
}

function moveSelection(
    rowDelta: number,
    columnDelta: number,
    { extend = false }: { extend?: boolean } = {}
): void {
    const selection = ensureSelection();
    const bounds = getSelectionBounds();
    if (!selection || !bounds) {
        return;
    }

    const nextRow = Math.max(
        bounds.minRow,
        Math.min(bounds.maxRow, selection.rowNumber + rowDelta)
    );
    const nextColumn = Math.max(
        bounds.minColumn,
        Math.min(bounds.maxColumn, selection.columnNumber + columnDelta)
    );

    setSelectedCellLocal(
        { rowNumber: nextRow, columnNumber: nextColumn },
        {
            reveal: true,
            anchorCell: extend ? (selectionAnchorCell ?? selection) : undefined,
        }
    );
}

function moveSelectionByViewportWindow(direction: -1 | 1): void {
    const selection = ensureSelection();
    if (!model || !selection) {
        return;
    }

    const shiftRowCount = Math.max(1, model.page.rows.length - 1);
    const nextRowNumber = Math.max(
        1,
        Math.min(model.activeSheet.rowCount, selection.rowNumber + direction * shiftRowCount)
    );
    const nextViewportStartRow = model.page.startRow + direction * shiftRowCount;

    if (
        nextRowNumber === selection.rowNumber &&
        nextViewportStartRow === model.page.startRow
    ) {
        return;
    }

    pendingSelectionAfterRender = {
        rowNumber: nextRowNumber,
        columnNumber: selection.columnNumber,
        reveal: true,
    };

    vscode.postMessage({ type: "setViewportStartRow", rowNumber: nextViewportStartRow });
}

function submitSearch(direction: "next" | "prev"): void {
    const query = searchQuery.trim();
    if (!query) {
        return;
    }

    vscode.postMessage({ type: "search", query, direction, options: searchOptions });
}

function submitGoto(): void {
    const reference = gotoReference.trim();
    if (!reference) {
        return;
    }

    vscode.postMessage({ type: "gotoCell", reference });
}

function focusToolbarInput(role: "search" | "goto"): void {
    const selector = role === "search" ? '[data-role="search-input"]' : '[data-role="goto-input"]';
    const input = document.querySelector<HTMLInputElement>(selector);
    if (!input) {
        return;
    }

    input.focus();
    input.select();
}

function getPendingSummary(activeSheetKey: string): PendingSummary {
    const summary: PendingSummary = {
        sheetKeys: new Set<string>(),
        rows: new Set<number>(),
        columns: new Set<number>(),
    };

    for (const pendingEdit of pendingEdits.values()) {
        summary.sheetKeys.add(pendingEdit.sheetKey);

        if (pendingEdit.sheetKey !== activeSheetKey) {
            continue;
        }

        summary.rows.add(pendingEdit.rowNumber);
        summary.columns.add(pendingEdit.columnNumber);
    }

    return summary;
}

function updateView(view: ViewState): void {
    const setView = setViewState;
    if (!setView) {
        return;
    }

    flushSync(() => {
        setView(view);
    });
}

function renderLoading(message: string): void {
    queuedViewportStartRow = null;
    resetGridMeasurementState();
    updateView({ kind: "loading", message });
}

function renderError(message: string): void {
    queuedViewportStartRow = null;
    resetGridMeasurementState();
    updateView({ kind: "error", message });
}

function renderApp({
    commitEditing = true,
    revealSelection = false,
    useModelSelection = false,
}: {
    commitEditing?: boolean;
    revealSelection?: boolean;
    useModelSelection?: boolean;
} = {}): void {
    if (!model) {
        renderLoading(STRINGS.loading);
        return;
    }

    if (commitEditing) {
        finishEdit({ mode: "commit", refresh: false });
    }

    const scrollState = getPaneScrollState();
    const shouldRevealSelection = prepareSelectionForRender({
        revealSelection,
        useModelSelection,
    });
    viewRevision += 1;

    updateView({
        kind: "app",
        model,
        revealSelection: shouldRevealSelection,
        revision: viewRevision,
        scrollState,
    });
}

function PendingMarker({ extraClass }: { extraClass?: string }): React.ReactElement {
    return (
        <span
            className={classNames(["diff-marker", "diff-marker--pending", extraClass])}
            aria-hidden
        />
    );
}

function CellValue({
    value,
    formula,
}: {
    value: string;
    formula: string | null;
}): React.ReactElement | null {
    if (!value && !formula) {
        return null;
    }

    return (
        <>
            {value ? <span className="grid__cell-value">{value}</span> : null}
            {formula ? (
                <span className="cell__formula" title={formula}>
                    fx
                </span>
            ) : null}
        </>
    );
}

function ToolbarButton({
    actionLabel,
    icon,
    disabled = false,
    isActive = false,
    iconOnly = false,
    iconMirrored = false,
    onClick,
}: {
    actionLabel: string;
    icon: string;
    disabled?: boolean;
    isActive?: boolean;
    iconOnly?: boolean;
    iconMirrored?: boolean;
    onClick(): void;
}): React.ReactElement {
    return (
        <button
            aria-label={actionLabel}
            className={classNames([
                "toolbar__button",
                isActive && "is-active",
                iconOnly && "toolbar__button--icon",
            ])}
            disabled={disabled}
            title={actionLabel}
            type="button"
            onClick={onClick}
        >
            <span
                className={classNames([
                    "codicon",
                    icon,
                    "toolbar__button-icon",
                    iconMirrored && "toolbar__button-icon--flip",
                ])}
                aria-hidden
            />
            {iconOnly ? null : <span>{actionLabel}</span>}
        </button>
    );
}

function SearchToggle({
    isActive,
    label,
    icon,
    onClick,
}: {
    isActive: boolean;
    label: string;
    icon: string;
    onClick(): void;
}): React.ReactElement {
    return (
        <button
            aria-label={label}
            className={classNames(["codicon", icon, "toolbar__toggle", isActive && "is-active"])}
            title={label}
            type="button"
            onClick={onClick}
        />
    );
}

function EditorToolbar({ currentModel }: { currentModel: EditorRenderModel }): React.ReactElement {
    const hasPendingEdits = pendingEdits.size > 0 || currentModel.hasPendingEdits;
    const canUndo = undoStack.length > 0 || currentModel.canUndoStructuralEdits;
    const canRedo = redoStack.length > 0 || currentModel.canRedoStructuralEdits;
    const viewLocked = hasLockedView(currentModel.activeSheet.freezePane);
    const viewLockActionLabel = viewLocked ? STRINGS.unlockView : STRINGS.lockView;

    return (
        <header className="toolbar toolbar--editor">
            <div className="toolbar__group toolbar__group--grow">
                <label className="toolbar__field">
                    <span className="codicon codicon-search toolbar__field-icon" aria-hidden />
                    <input
                        className="toolbar__input"
                        data-role="search-input"
                        defaultValue={searchQuery}
                        placeholder={STRINGS.searchPlaceholder}
                        type="text"
                        onChange={(event) => {
                            searchQuery = event.currentTarget.value;
                        }}
                        onKeyDown={(event) => {
                            if (event.key !== "Enter") {
                                return;
                            }

                            event.preventDefault();
                            submitSearch(event.shiftKey ? "prev" : "next");
                        }}
                    />
                    <span className="toolbar__field-actions">
                        <SearchToggle
                            isActive={searchOptions.isRegexp}
                            label={STRINGS.searchRegex}
                            icon="codicon-regex"
                            onClick={() => {
                                searchOptions = {
                                    ...searchOptions,
                                    isRegexp: !searchOptions.isRegexp,
                                };
                                renderApp({ commitEditing: false });
                            }}
                        />
                        <SearchToggle
                            isActive={searchOptions.matchCase}
                            label={STRINGS.searchMatchCase}
                            icon="codicon-case-sensitive"
                            onClick={() => {
                                searchOptions = {
                                    ...searchOptions,
                                    matchCase: !searchOptions.matchCase,
                                };
                                renderApp({ commitEditing: false });
                            }}
                        />
                        <SearchToggle
                            isActive={searchOptions.wholeWord}
                            label={STRINGS.searchWholeWord}
                            icon="codicon-whole-word"
                            onClick={() => {
                                searchOptions = {
                                    ...searchOptions,
                                    wholeWord: !searchOptions.wholeWord,
                                };
                                renderApp({ commitEditing: false });
                            }}
                        />
                    </span>
                </label>
                <ToolbarButton
                    actionLabel={STRINGS.findPrev}
                    icon="codicon-arrow-up"
                    iconOnly
                    onClick={() => submitSearch("prev")}
                />
                <ToolbarButton
                    actionLabel={STRINGS.findNext}
                    icon="codicon-arrow-down"
                    iconOnly
                    onClick={() => submitSearch("next")}
                />
                <label className="toolbar__field toolbar__field--goto">
                    <span className="codicon codicon-target toolbar__field-icon" aria-hidden />
                    <input
                        className="toolbar__input"
                        data-role="goto-input"
                        defaultValue={gotoReference}
                        placeholder={STRINGS.gotoPlaceholder}
                        type="text"
                        onChange={(event) => {
                            gotoReference = event.currentTarget.value;
                        }}
                        onKeyDown={(event) => {
                            if (event.key !== "Enter") {
                                return;
                            }

                            event.preventDefault();
                            submitGoto();
                        }}
                    />
                </label>
                <ToolbarButton
                    actionLabel={STRINGS.goto}
                    icon="codicon-target"
                    onClick={submitGoto}
                />
            </div>
            <div className="toolbar__group">
                <ToolbarButton
                    actionLabel={STRINGS.undo}
                    disabled={!currentModel.canEdit || !canUndo || isSaving}
                    icon="codicon-redo"
                    iconMirrored
                    iconOnly={true}
                    onClick={undoPendingEdits}
                />
                <ToolbarButton
                    actionLabel={STRINGS.redo}
                    disabled={!currentModel.canEdit || !canRedo || isSaving}
                    icon="codicon-redo"
                    iconOnly={true}
                    onClick={redoPendingEdits}
                />
                <ToolbarButton
                    actionLabel={STRINGS.reload}
                    icon="codicon-refresh"
                    iconOnly={true}
                    onClick={() => vscode.postMessage({ type: "reload" })}
                />
                <ToolbarButton
                    actionLabel={viewLockActionLabel}
                    disabled={!canToggleViewLock() || isSaving}
                    icon={viewLocked ? "codicon-lock" : "codicon-unlock"}
                    iconOnly={true}
                    isActive={viewLocked}
                    onClick={toggleViewLock}
                />
                <ToolbarButton
                    actionLabel={STRINGS.save}
                    disabled={!currentModel.canEdit || !hasPendingEdits || isSaving}
                    icon="codicon-save"
                    iconOnly={true}
                    isActive={hasPendingEdits}
                    onClick={triggerSave}
                />
            </div>
        </header>
    );
}

function CellEditor({ edit }: { edit: EditingCell }): React.ReactElement {
    const inputRef = React.useRef<HTMLInputElement | null>(null);

    React.useLayoutEffect(() => {
        inputRef.current?.focus();
        inputRef.current?.select();
    }, []);

    return (
        <input
            ref={inputRef}
            className="grid__cell-input"
            defaultValue={edit.value}
            type="text"
            onBlur={() => {
                setTimeout(() => finishEdit({ mode: "commit", clearSelection: true }), 0);
            }}
            onChange={(event) => {
                if (editingCell === edit) {
                    editingCell.value = event.currentTarget.value;
                }
            }}
            onClick={(event) => event.stopPropagation()}
            onDoubleClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === "Tab") {
                    event.preventDefault();
                    finishEdit({ mode: "commit", clearSelection: true });
                } else if (event.key === "Escape") {
                    event.preventDefault();
                    finishEdit({ mode: "cancel" });
                }
            }}
        />
    );
}

function GridCell({
    cell,
    columnNumber,
    row,
    selectionRange,
    hasExpandedRange,
}: {
    cell: EditorGridCellView;
    columnNumber: number;
    row: EditorGridRowView;
    selectionRange: CellRange | null;
    hasExpandedRange: boolean;
}): React.ReactElement {
    const pendingKey = getPendingEditKey(model!.activeSheet.key, row.rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const value = pendingEdit ? pendingEdit.value : cell.value;
    const formula = pendingEdit ? null : cell.formula;
    const editable = isGridCellEditable(cell);
    const isPrimarySelection =
        !hasExpandedRange &&
        selectedCell?.rowNumber === row.rowNumber &&
        selectedCell.columnNumber === columnNumber;
    const isSelected = Boolean(
        selectionRange &&
        row.rowNumber >= selectionRange.startRow &&
        row.rowNumber <= selectionRange.endRow &&
        columnNumber >= selectionRange.startColumn &&
        columnNumber <= selectionRange.endColumn
    );
    const isActiveRow = Boolean(
        selectionRange &&
            row.rowNumber >= selectionRange.startRow &&
            row.rowNumber <= selectionRange.endRow
    );
    const isActiveColumn = Boolean(
        selectionRange &&
            columnNumber >= selectionRange.startColumn &&
            columnNumber <= selectionRange.endColumn
    );
    const isEditing =
        editingCell?.rowNumber === row.rowNumber && editingCell.columnNumber === columnNumber;

    return (
        <td
            aria-selected={isSelected}
            className={classNames([
                "grid__cell",
                isSelected && "grid__cell--selected-range",
                isPrimarySelection && "grid__cell--selected",
                isActiveRow && "grid__cell--active-row",
                isActiveColumn && "grid__cell--active-column",
                !editable && "grid__cell--locked",
                pendingEdit && "grid__cell--pending",
                isEditing && "grid__cell--editing",
                ...getSelectionOutlineClasses(row.rowNumber, columnNumber, selectionRange),
            ])}
            data-column-number={columnNumber}
            data-editable={editable}
            data-role="grid-cell"
            data-row-number={row.rowNumber}
            title={getCellTooltip(cell.address, value, formula)}
            onPointerDown={(event) => {
                if (event.button !== 0) {
                    return;
                }

                closeTabContextMenu({ refresh: false });
                startSelectionDrag(event.pointerId, { rowNumber: row.rowNumber, columnNumber });
                setSelectedCellLocal(
                    { rowNumber: row.rowNumber, columnNumber },
                    {
                        syncHost: false,
                        anchorCell: { rowNumber: row.rowNumber, columnNumber },
                    }
                );
            }}
            onClick={(event) => {
                if (suppressNextCellClick) {
                    suppressNextCellClick = false;
                    event.preventDefault();
                    return;
                }

                if (editingCell) {
                    finishEdit({ mode: "commit", refresh: false });
                }

                setSelectedCellLocal(
                    { rowNumber: row.rowNumber, columnNumber },
                    {
                        syncHost: true,
                        anchorCell:
                            event.shiftKey && selectedCell
                                ? (selectionAnchorCell ?? selectedCell)
                                : undefined,
                    }
                );
            }}
            onDoubleClick={(event) => {
                if (!model?.canEdit || !editable) {
                    return;
                }

                event.preventDefault();
                startEditCell(row.rowNumber, columnNumber, value);
            }}
        >
            <div className="grid__cell-content">
                {isEditing && editingCell ? (
                    <CellEditor edit={editingCell} />
                ) : (
                    <CellValue formula={formula} value={value} />
                )}
            </div>
        </td>
    );
}

function GridColumnHeaderCell({
    currentModel,
    pendingSummary,
    columnNumber,
    selectionRange,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    columnNumber: number;
    selectionRange: CellRange | null;
}): React.ReactElement {
    const hasPending = pendingSummary.columns.has(columnNumber);
    const isActiveColumn = Boolean(
        selectionRange &&
            columnNumber >= selectionRange.startColumn &&
            columnNumber <= selectionRange.endColumn
    );

    return (
        <th
            key={columnNumber}
            className={classNames([
                "grid__column",
                hasPending && "grid__column--diff",
                hasPending && "grid__column--pending",
                isActiveColumn && "grid__column--active",
            ])}
            data-column-number={columnNumber}
        >
            <span className="grid__column-label">
                {hasPending ? <PendingMarker /> : null}
                <span>{currentModel.page.columns[columnNumber - 1]}</span>
            </span>
        </th>
    );
}

function GridRowHeaderCell({
    pendingSummary,
    row,
    selectionRange,
}: {
    pendingSummary: PendingSummary;
    row: EditorGridRowView;
    selectionRange: CellRange | null;
}): React.ReactElement {
    const hasPending = pendingSummary.rows.has(row.rowNumber);
    const isActiveRow = Boolean(
        selectionRange &&
            row.rowNumber >= selectionRange.startRow &&
            row.rowNumber <= selectionRange.endRow
    );

    return (
        <th
            className={classNames([
                "grid__row-number",
                hasPending && "grid__row-number--pending",
                isActiveRow && "grid__row-number--active",
            ])}
            data-role="grid-row-header"
            data-row-number={row.rowNumber}
        >
            <span className="grid__row-label">
                {hasPending ? <PendingMarker /> : null}
                <span>{row.rowNumber}</span>
            </span>
        </th>
    );
}

function GridSectionTable({
    currentModel,
    pendingSummary,
    rows,
    columnNumbers,
    includeHeader,
    includeRowHeaders,
    split = false,
    topSpacerHeight = 0,
    bottomSpacerHeight = 0,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    rows: EditorGridRowView[];
    columnNumbers: number[];
    includeHeader: boolean;
    includeRowHeaders: boolean;
    split?: boolean;
    topSpacerHeight?: number;
    bottomSpacerHeight?: number;
}): React.ReactElement | null {
    const hasHeader = includeHeader && (includeRowHeaders || columnNumbers.length > 0);
    const hasBody = rows.length > 0 && (includeRowHeaders || columnNumbers.length > 0);
    const selectionRange = getSelectionRange();
    const hasExpandedRange = hasExpandedSelection(selectionRange);
    if (!hasHeader && !hasBody) {
        return null;
    }

    return (
        <table className={classNames(["grid", split && "grid--split"])}>
            <colgroup>
                {includeRowHeaders ? <col data-role="grid-row-header-col" /> : null}
                {columnNumbers.map((columnNumber) => (
                    <col
                        key={columnNumber}
                        data-column-number={columnNumber}
                        data-role="grid-column-col"
                    />
                ))}
            </colgroup>
            {includeHeader ? (
                <thead>
                    <tr data-role="grid-header-row">
                        {includeRowHeaders ? <th className="grid__row-number">#</th> : null}
                        {columnNumbers.map((columnNumber) => (
                            <GridColumnHeaderCell
                                key={columnNumber}
                                currentModel={currentModel}
                                pendingSummary={pendingSummary}
                                columnNumber={columnNumber}
                                selectionRange={selectionRange}
                            />
                        ))}
                    </tr>
                </thead>
            ) : null}
            <tbody>
                {topSpacerHeight > 0 ? (
                    <tr aria-hidden="true" className="grid__spacer-row">
                        <td
                            className="grid__spacer-cell"
                            colSpan={columnNumbers.length + (includeRowHeaders ? 1 : 0)}
                            style={{ height: `${topSpacerHeight}px` }}
                        />
                    </tr>
                ) : null}
                {rows.map((row) => {
                    return (
                        <tr
                            key={`${row.rowNumber}:${includeRowHeaders ? "row" : "cell"}:${columnNumbers[0] ?? 0}`}
                            data-role="grid-row"
                            data-row-number={row.rowNumber}
                        >
                            {includeRowHeaders ? (
                                <GridRowHeaderCell
                                    pendingSummary={pendingSummary}
                                    row={row}
                                    selectionRange={selectionRange}
                                />
                            ) : null}
                            {columnNumbers.map((columnNumber) => (
                                <GridCell
                                    key={`${row.rowNumber}:${columnNumber}`}
                                    cell={row.cells[columnNumber - 1]!}
                                    columnNumber={columnNumber}
                                    row={row}
                                    selectionRange={selectionRange}
                                    hasExpandedRange={hasExpandedRange}
                                />
                            ))}
                        </tr>
                    );
                })}
                {bottomSpacerHeight > 0 ? (
                    <tr aria-hidden="true" className="grid__spacer-row">
                        <td
                            className="grid__spacer-cell"
                            colSpan={columnNumbers.length + (includeRowHeaders ? 1 : 0)}
                            style={{ height: `${bottomSpacerHeight}px` }}
                        />
                    </tr>
                ) : null}
            </tbody>
        </table>
    );
}

function EditorTable({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement | null {
    if (currentModel.page.rows.length === 0) {
        return <div className="empty-table">{STRINGS.noRowsAvailable}</div>;
    }

    return (
        <GridSectionTable
            currentModel={currentModel}
            pendingSummary={pendingSummary}
            rows={currentModel.page.rows}
            columnNumbers={currentModel.page.columns.map((_, index) => index + 1)}
            includeHeader={true}
            includeRowHeaders={true}
            topSpacerHeight={getTopSpacerHeight(currentModel)}
            bottomSpacerHeight={getBottomSpacerHeight(currentModel)}
        />
    );
}

function FrozenEditorTable({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    const { columnCount } = getVisibleFreezeCounts(currentModel);
    const frozenRows = getFrozenRowsForCurrentView(currentModel);
    const scrollableRows = getScrollableRowsForCurrentView(currentModel);
    const topSpacerHeight = getTopSpacerHeight(currentModel);
    const bottomSpacerHeight = getBottomSpacerHeight(currentModel);
    const frozenColumnNumbers = Array.from({ length: columnCount }, (_, index) => index + 1);
    const scrollColumnNumbers = Array.from(
        { length: Math.max(currentModel.page.columns.length - columnCount, 0) },
        (_, index) => columnCount + index + 1
    );

    return (
        <div className="pane__table pane__table--split">
            <div className="freeze-grid">
                <div
                    className="freeze-grid__pane freeze-grid__pane--top-left"
                    data-pane="top-left"
                    onWheel={forwardFrozenPaneWheel}
                >
                    <GridSectionTable
                        currentModel={currentModel}
                        pendingSummary={pendingSummary}
                        rows={frozenRows}
                        columnNumbers={frozenColumnNumbers}
                        includeHeader={true}
                        includeRowHeaders={true}
                        split={true}
                    />
                </div>
                <div
                    className="freeze-grid__pane freeze-grid__pane--top-right"
                    data-pane="top-right"
                    onWheel={forwardFrozenPaneWheel}
                    onScroll={(event) => syncFrozenPaneScroll("top-right", event.currentTarget)}
                >
                    <GridSectionTable
                        currentModel={currentModel}
                        pendingSummary={pendingSummary}
                        rows={frozenRows}
                        columnNumbers={scrollColumnNumbers}
                        includeHeader={true}
                        includeRowHeaders={false}
                        split={true}
                    />
                </div>
                <div
                    className="freeze-grid__pane freeze-grid__pane--bottom-left"
                    data-pane="bottom-left"
                    onWheel={forwardFrozenPaneWheel}
                    onScroll={(event) => {
                        syncFrozenPaneScroll("bottom-left", event.currentTarget);
                        scheduleViewportRequestForScroll(event.currentTarget);
                    }}
                >
                    <GridSectionTable
                        currentModel={currentModel}
                        pendingSummary={pendingSummary}
                        rows={scrollableRows}
                        columnNumbers={frozenColumnNumbers}
                        includeHeader={false}
                        includeRowHeaders={true}
                        split={true}
                        topSpacerHeight={topSpacerHeight}
                        bottomSpacerHeight={bottomSpacerHeight}
                    />
                </div>
                <div
                    className="freeze-grid__pane freeze-grid__pane--bottom-right"
                    data-pane="bottom-right"
                    data-role="grid-scroll-main"
                    onScroll={(event) => {
                        syncFrozenPaneScroll("bottom-right", event.currentTarget);
                        scheduleViewportRequestForScroll(event.currentTarget);
                    }}
                >
                    <GridSectionTable
                        currentModel={currentModel}
                        pendingSummary={pendingSummary}
                        rows={scrollableRows}
                        columnNumbers={scrollColumnNumbers}
                        includeHeader={false}
                        includeRowHeaders={false}
                        split={true}
                        topSpacerHeight={topSpacerHeight}
                        bottomSpacerHeight={bottomSpacerHeight}
                    />
                </div>
            </div>
        </div>
    );
}

function EditorPane({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    const viewLocked = hasLockedView(currentModel.activeSheet.freezePane);
    const hasVisibleRows =
        currentModel.page.rows.length > 0 || currentModel.page.frozenRows.length > 0;

    return (
        <section className="pane pane--single pane--editor">
            {!hasVisibleRows ? (
                <div className="pane__table">
                    <div className="empty-table">{STRINGS.noRowsAvailable}</div>
                </div>
            ) : viewLocked ? (
                <FrozenEditorTable currentModel={currentModel} pendingSummary={pendingSummary} />
            ) : (
                <div
                    className="pane__table"
                    onScroll={(event) => scheduleViewportRequestForScroll(event.currentTarget)}
                >
                    <EditorTable currentModel={currentModel} pendingSummary={pendingSummary} />
                </div>
            )}
        </section>
    );
}

function Tabs({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    return (
        <div
            className="tabs"
            onContextMenu={(event) => {
                const target = event.target;
                if (target instanceof HTMLElement && target.closest('[data-role="sheet-tab"]')) {
                    return;
                }

                event.preventDefault();
                openTabContextMenu(currentModel.activeSheet.key, event.clientX, event.clientY);
            }}
        >
            {currentModel.sheets.map((sheet: EditorSheetTabView) => {
                const hasPending = pendingSummary.sheetKeys.has(sheet.key);

                return (
                    <button
                        key={sheet.key}
                        className={classNames(["tab", sheet.isActive && "is-active"])}
                        data-role="sheet-tab"
                        title={sheet.label}
                        type="button"
                        onClick={() =>
                            vscode.postMessage({ type: "setSheet", sheetKey: sheet.key })
                        }
                        onContextMenu={(event) => {
                            event.preventDefault();
                            openTabContextMenu(sheet.key, event.clientX, event.clientY);
                        }}
                    >
                        {hasPending ? <PendingMarker extraClass="tab__marker" /> : null}
                        <span className="tab__label">{sheet.label}</span>
                    </button>
                );
            })}
        </div>
    );
}

function TabContextMenu({
    currentModel,
}: {
    currentModel: EditorRenderModel;
}): React.ReactElement | null {
    if (!tabContextMenu || !currentModel.canEdit) {
        return null;
    }

    const { sheetKey } = tabContextMenu;
    const menuStyle: React.CSSProperties = {
        left: Math.max(8, Math.min(tabContextMenu.x, window.innerWidth - 188)),
        top: Math.max(8, Math.min(tabContextMenu.y, window.innerHeight - 132)),
    };
    const disableDelete = currentModel.sheets.length <= 1;

    return (
        <div className="context-menu" data-role="tab-context-menu" style={menuStyle}>
            <button className="context-menu__item" type="button" onClick={requestAddSheet}>
                <span className="codicon codicon-add context-menu__icon" aria-hidden />
                <span>{STRINGS.addSheet}</span>
            </button>
            <button
                className="context-menu__item"
                type="button"
                onClick={() => requestRenameSheet(sheetKey)}
            >
                <span className="codicon codicon-edit context-menu__icon" aria-hidden />
                <span>{STRINGS.renameSheet}</span>
            </button>
            <button
                className="context-menu__item context-menu__item--danger"
                disabled={disableDelete}
                type="button"
                onClick={() => requestDeleteSheet(sheetKey)}
            >
                <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                <span>{STRINGS.deleteSheet}</span>
            </button>
        </div>
    );
}

function Status({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    return (
        <footer className="footer">
            <Tabs currentModel={currentModel} pendingSummary={pendingSummary} />
        </footer>
    );
}

function EditorApp({ view }: { view: Extract<ViewState, { kind: "app" }> }): React.ReactElement {
    const pendingSummary = getPendingSummary(view.model.activeSheet.key);
    const viewLocked = hasLockedView(view.model.activeSheet.freezePane);

    React.useLayoutEffect(() => {
        restorePaneScrollState(view.scrollState);
        if (viewLocked) {
            clearFrozenPaneLayout();
            scheduleFrozenPaneLayoutSync({
                revealSelection: view.revealSelection,
                forceColumnMeasurement: true,
            });
            return;
        }

        scheduleGridMetricRefresh({
            revealSelection: view.revealSelection,
            viewLocked,
        });
    }, [
        viewLocked,
        view.model.activeSheet.freezePane,
        view.revision,
        view.revealSelection,
        view.scrollState,
    ]);

    return (
        <div className="app app--editor">
            <EditorToolbar currentModel={view.model} />
            <section className="panes panes--single">
                <EditorPane currentModel={view.model} pendingSummary={pendingSummary} />
            </section>
            <Status currentModel={view.model} pendingSummary={pendingSummary} />
            <TabContextMenu currentModel={view.model} />
        </div>
    );
}

function Shell({
    kind,
    message,
}: {
    kind: "loading" | "error";
    message: string;
}): React.ReactElement {
    const className = kind === "loading" ? "loading-shell" : "empty-shell";
    const messageClassName = kind === "loading" ? "loading-shell__message" : "empty-shell__message";

    return (
        <div className={className}>
            <div className={messageClassName}>{message}</div>
        </div>
    );
}

function Root(): React.ReactElement {
    const [view, setView] = React.useState<ViewState>({
        kind: "loading",
        message: STRINGS.loading,
    });

    setViewState = setView;

    if (view.kind === "app") {
        return <EditorApp view={view} />;
    }

    return <Shell kind={view.kind} message={view.message} />;
}

window.addEventListener("resize", () => {
    if (hasLockedView(model?.activeSheet.freezePane)) {
        clearFrozenPaneLayout();
        scheduleFrozenPaneLayoutSync({ forceColumnMeasurement: true });
        return;
    }

    scheduleGridMetricRefresh({
        viewLocked: false,
        forceColumnMeasurement: true,
    });
});

window.addEventListener("message", (event: MessageEvent<IncomingMessage>) => {
    const message = event.data;

    if (message.type === "loading") {
        renderLoading(message.message);
        return;
    }

    if (message.type === "error") {
        isSaving = false;
        renderError(message.message);
        return;
    }

    if (message.type === "render") {
        if (message.payload.activeSheet.key !== lastMeasuredSheetKey) {
            rowHeightByNumber.clear();
            columnWidthByNumber.clear();
            lastMeasuredSheetKey = message.payload.activeSheet.key;
            lastObservedPaneScrollTop = null;
        }

        model = message.payload;
        lastRequestedViewportStartRow = message.payload.page.startRow;
        isSaving = false;

        if (message.clearPendingEdits) {
            pendingEdits.clear();
            undoStack.length = 0;
            redoStack.length = 0;
            editingCell = null;
            selectionAnchorCell = selectedCell;
            lastPendingNotification = null;
            lastPendingEditsSyncKey = serializePendingEdits([]);
            notifyPendingEditState();
        } else if (message.replacePendingEdits) {
            pendingEdits.clear();
            for (const edit of message.replacePendingEdits) {
                pendingEdits.set(
                    getPendingEditKey(edit.sheetKey, edit.rowNumber, edit.columnNumber),
                    edit
                );
            }

            if (message.resetPendingHistory) {
                undoStack.length = 0;
                redoStack.length = 0;
                editingCell = null;
            }

            lastPendingNotification = null;
            lastPendingEditsSyncKey = serializePendingEdits(Array.from(pendingEdits.values()));
            notifyPendingEditState();
        }

        if (queuedViewportStartRow !== null && queuedViewportStartRow !== message.payload.page.startRow) {
            const nextViewportStartRow = queuedViewportStartRow;
            queuedViewportStartRow = null;
            requestViewportStartRow(nextViewportStartRow);
        }

        renderApp({
            revealSelection: !message.silent,
            useModelSelection: message.useModelSelection,
        });
    }
});

document.addEventListener("keydown", (event: KeyboardEvent) => {
    const isTextInputContext = isTextInputTarget(event.target);

    if (event.key === "Escape" && tabContextMenu) {
        event.preventDefault();
        closeTabContextMenu();
        return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        triggerSave();
        return;
    }

    if (isTextInputContext) {
        return;
    }

    if (!editingCell && (event.ctrlKey || event.metaKey) && !event.altKey) {
        if (event.key.toLowerCase() === "f") {
            event.preventDefault();
            focusToolbarInput("search");
            return;
        }

        if (event.key.toLowerCase() === "g") {
            event.preventDefault();
            focusToolbarInput("goto");
            return;
        }

        if (event.key.toLowerCase() === "z") {
            event.preventDefault();
            if (event.shiftKey) {
                redoPendingEdits();
            } else {
                undoPendingEdits();
            }
            return;
        }

        if (event.key.toLowerCase() === "y") {
            event.preventDefault();
            redoPendingEdits();
            return;
        }
    }

    if (editingCell) {
        return;
    }

    if (!event.altKey && !event.ctrlKey && !event.metaKey && event.shiftKey) {
        switch (event.key) {
            case "ArrowUp":
                event.preventDefault();
                moveSelection(-1, 0, { extend: true });
                return;
            case "ArrowDown":
                event.preventDefault();
                moveSelection(1, 0, { extend: true });
                return;
            case "ArrowLeft":
                event.preventDefault();
                moveSelection(0, -1, { extend: true });
                return;
            case "ArrowRight":
                event.preventDefault();
                moveSelection(0, 1, { extend: true });
                return;
        }
    }

    if (isClearSelectedCellKey(event) && selectedCell) {
        event.preventDefault();
        clearSelectedCellValue();
        return;
    }

    if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
        return;
    }

    switch (event.key) {
        case "ArrowUp":
            event.preventDefault();
            moveSelection(-1, 0);
            return;
        case "ArrowDown":
            event.preventDefault();
            moveSelection(1, 0);
            return;
        case "ArrowLeft":
            event.preventDefault();
            moveSelection(0, -1);
            return;
        case "ArrowRight":
            event.preventDefault();
            moveSelection(0, 1);
            return;
        case "Tab":
            event.preventDefault();
            moveSelection(0, 1);
            return;
        case "Enter":
            event.preventDefault();
            moveSelection(1, 0);
            return;
        case "PageUp":
            event.preventDefault();
            moveSelectionByViewportWindow(-1);
            return;
        case "PageDown":
            event.preventDefault();
            moveSelectionByViewportWindow(1);
            return;
    }
});

document.addEventListener("copy", (event: ClipboardEvent) => {
    if (isTextInputTarget(event.target) || editingCell || !model || !selectedCell) {
        return;
    }

    const text = serializeSelectionToClipboard();
    if (!text) {
        return;
    }

    event.preventDefault();
    event.clipboardData?.setData("text/plain", text);
});

document.addEventListener("pointerdown", (event: PointerEvent) => {
    if (!tabContextMenu) {
        return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
        closeTabContextMenu();
        return;
    }

    if (target.closest('[data-role="tab-context-menu"]')) {
        return;
    }

    closeTabContextMenu();
});

document.addEventListener("pointermove", (event: PointerEvent) => {
    if (!selectionDragState || selectionDragState.pointerId !== event.pointerId) {
        return;
    }

    if ((event.buttons & 1) === 0) {
        stopSelectionDrag(event.pointerId);
        return;
    }

    const targetCell = getCellPositionFromElement(
        document.elementFromPoint(event.clientX, event.clientY)
    );
    if (!targetCell) {
        return;
    }

    updateSelectionDrag(targetCell);
});

document.addEventListener("pointerup", (event: PointerEvent) => {
    stopSelectionDrag(event.pointerId);
});

document.addEventListener("pointercancel", (event: PointerEvent) => {
    stopSelectionDrag(event.pointerId);
});

document.addEventListener("paste", (event: ClipboardEvent) => {
    if (
        isTextInputTarget(event.target) ||
        editingCell ||
        !model ||
        !selectedCell ||
        !model.canEdit
    ) {
        return;
    }

    const text = event.clipboardData?.getData("text/plain");
    if (!text) {
        return;
    }

    event.preventDefault();
    applyPastedGrid(normalizePastedRows(text));
});

const rootElement = document.getElementById("app");
if (!rootElement) {
    throw new Error("Missing #app root element");
}

rootElement.removeAttribute("class");
const root = createRoot(rootElement);
flushSync(() => {
    root.render(<Root />);
});
renderLoading(STRINGS.loading);
vscode.postMessage({ type: "ready" });
