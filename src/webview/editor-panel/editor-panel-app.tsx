import * as React from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import type {
    CellSnapshot,
    EditorRenderModel,
    EditorSheetTabView,
} from "../../core/model/types";
import { createCellKey, getColumnLabel } from "../../core/model/cells";
import { formatI18nMessage, RUNTIME_MESSAGES } from "../../i18n/catalog";
import {
    EDITOR_VIRTUAL_COLUMN_WIDTH,
    EDITOR_VIRTUAL_HEADER_HEIGHT,
    EDITOR_VIRTUAL_ROW_HEIGHT,
    clampEditorScrollPosition,
    createEditorColumnWindow,
    createEditorRowWindow,
    getEditorDisplayGridDimensions,
    getEditorContentSize,
    getFrozenEditorCounts,
    getVisibleFrozenEditorCounts,
} from "./editor-virtual-grid";
import {
    isSelectionFocusCell,
    shouldResetInvisibleSelectionAnchor,
    shouldSyncLocalSelectionDomFromModelSelection,
    shouldUseLocalSimpleSelectionUpdate,
} from "./editor-selection-render";
import {
    createColumnSelectionRange,
    createRowSelectionRange,
    createSelectionRange,
    hasExpandedSelectionRange,
    type SelectionRange as CellRange,
} from "./editor-selection-range";
import {
    getEditorToolbarSyncSnapshot,
    notifyEditorToolbarSync,
    subscribeEditorToolbarSync,
} from "./editor-toolbar-sync";
import {
    getToolbarCellEditTargetKey,
    shouldResetToolbarCellValueDraft,
    type ToolbarCellEditTarget,
} from "./editor-toolbar-input";
import { getMaxVisibleSheetTabsForWidth, partitionSheetTabs } from "../editor-sheet-tabs";
import type {
    EditorSearchResultMessage,
    EditorSearchScope,
    SearchOptions,
} from "./editor-panel-types";
import { resolveEditorReplaceResultInSheet } from "./editor-panel-logic";
import {
    rebasePendingHistory,
    type PendingHistoryChange as PendingEditChange,
    type PendingHistoryEntry as HistoryEntry,
} from "./editor-pending-history";
import { stabilizeIncomingRenderModel } from "./editor-render-stabilizer";
import { getFreezePaneCountsForCell, hasLockedView } from "../view-lock";

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
    | { type: "insertRow"; rowNumber: number }
    | { type: "deleteRow"; rowNumber: number }
    | { type: "insertColumn"; columnNumber: number }
    | { type: "deleteColumn"; columnNumber: number }
    | {
          type: "search";
          query: string;
          direction: "next" | "prev";
          options: SearchOptions;
          scope: EditorSearchScope;
          selectionRange?: CellRange;
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
    | EditorSearchResultMessage
    | {
          type: "render";
          payload: EditorRenderModel;
          silent?: boolean;
          clearPendingEdits?: boolean;
          preservePendingHistory?: boolean;
          reuseActiveSheetData?: boolean;
          useModelSelection?: boolean;
          replacePendingEdits?: Array<{
              sheetKey: string;
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
          resetPendingHistory?: boolean;
      };

interface CellPosition {
    rowNumber: number;
    columnNumber: number;
}

interface EditorGridCellView {
    key: string;
    address: string;
    value: string;
    formula: string | null;
    isPresent: boolean;
    isSelected: boolean;
}

interface EditingCell extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingEdit extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingSelection extends CellPosition {
    reveal: boolean;
}

interface PendingSummary {
    sheetKeys: Set<string>;
    rows: Set<number>;
    columns: Set<number>;
}

type SearchPanelMode = "find" | "replace";
type SearchPanelFeedbackStatus = EditorSearchResultMessage["status"] | "replaced" | "no-change";

interface SearchPanelFeedback {
    status: SearchPanelFeedbackStatus;
    message?: string;
}

interface ScrollState {
    top: number;
    left: number;
}

type ContextMenuState =
    | {
          kind: "tab";
          sheetKey: string;
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

interface SelectionDragState {
    anchorCell: CellPosition;
    pointerId: number;
}

interface SearchPanelPosition {
    left: number;
    top: number;
}

interface SearchPanelDragState {
    pointerId: number;
    offsetX: number;
    offsetY: number;
}

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

const DEFAULT_STRINGS = RUNTIME_MESSAGES.en.editorPanel;

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
let selectionRangeOverride: CellRange | null = null;
let editingCell: EditingCell | null = null;
let isSaving = false;
let lastPendingNotification: boolean | null = null;
let pendingSelectionAfterRender: PendingSelection | null = null;
let suppressAutoSelection = false;
let lastPendingEditsSyncKey: string | null = null;
let viewRevision = 0;
let setViewState: React.Dispatch<React.SetStateAction<ViewState>> | null = null;
let contextMenu: ContextMenuState | null = null;
let selectionDragState: SelectionDragState | null = null;
let searchPanelDragState: SearchPanelDragState | null = null;
let suppressNextCellClick = false;

const WHEEL_DELTA_LINE_MODE = 1;
const WHEEL_DELTA_PAGE_MODE = 2;
const WHEEL_LINE_SCROLL_PIXELS = 40;
const DEFAULT_EDITOR_VIEWPORT_HEIGHT = 480;
const DEFAULT_EDITOR_VIEWPORT_WIDTH = 960;
const SHEET_TAB_ITEM_GAP = 1;
const SHEET_TAB_ESTIMATED_WIDTH = 120;
const SHEET_TAB_VISIBLE_MAX_WIDTH = 144;
const SHEET_TAB_OVERFLOW_TRIGGER_WIDTH = 32;
const DEFAULT_SEARCH_OPTIONS: SearchOptions = {
    isRegexp: false,
    matchCase: false,
    wholeWord: false,
};

let isSearchPanelOpen = false;
let searchMode: SearchPanelMode = "find";
let searchQuery = "";
let replaceValue = "";
let searchOptions: SearchOptions = { ...DEFAULT_SEARCH_OPTIONS };
let searchScope: EditorSearchScope = "sheet";
let searchFeedback: SearchPanelFeedback | null = null;
let searchPanelPosition: SearchPanelPosition | null = null;
let searchSelectionRange: CellRange | null = null;

interface VirtualViewportState {
    scrollTop: number;
    scrollLeft: number;
    viewportHeight: number;
    viewportWidth: number;
    rowHeaderWidth: number;
    frozenRowCount: number;
    frozenColumnCount: number;
    frozenRowNumbers: number[];
    frozenColumnNumbers: number[];
    rowNumbers: number[];
    columnNumbers: number[];
}

function getViewportElement(): HTMLElement | null {
    return document.querySelector<HTMLElement>('[data-role="grid-scroll-main"]');
}

function getEditorAppElement(): HTMLElement | null {
    return document.querySelector<HTMLElement>('[data-role="editor-app"]');
}

function getSearchPanelShellElement(): HTMLElement | null {
    return document.querySelector<HTMLElement>('[data-role="search-panel-shell"]');
}

function clampSearchPanelPosition(
    position: SearchPanelPosition,
    {
        appElement = getEditorAppElement(),
        panelElement = getSearchPanelShellElement(),
    }: {
        appElement?: HTMLElement | null;
        panelElement?: HTMLElement | null;
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

function syncSearchPanelShellPosition(): void {
    if (!searchPanelPosition) {
        return;
    }

    const shell = getSearchPanelShellElement();
    if (!shell) {
        return;
    }

    const nextPosition = clampSearchPanelPosition(searchPanelPosition, {
        panelElement: shell,
    });
    searchPanelPosition = nextPosition;
    shell.style.left = `${nextPosition.left}px`;
    shell.style.top = `${nextPosition.top}px`;
    shell.style.right = "auto";
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

function getVirtualViewportState(currentModel: EditorRenderModel | null): VirtualViewportState | null {
    if (!currentModel) {
        return null;
    }

    const pane = getViewportElement();
    const viewportHeight = pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT;
    const viewportWidth = pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH;
    const scrollTop = pane?.scrollTop ?? 0;
    const scrollLeft = pane?.scrollLeft ?? 0;
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        viewportHeight,
        viewportWidth,
    });
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        freezePane: currentModel.activeSheet.freezePane,
    });
    const { rowCount: visibleFrozenRowCount, columnCount: visibleFrozenColumnCount } =
        getVisibleFrozenEditorCounts({
            frozenRowCount,
            frozenColumnCount,
            viewportHeight,
            viewportWidth,
            rowHeaderWidth: displayGrid.rowHeaderWidth,
        });
    const rowWindow = createEditorRowWindow({
        totalRows: displayGrid.rowCount,
        frozenRowCount,
        scrollTop,
        viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        totalColumns: displayGrid.columnCount,
        frozenColumnCount,
        scrollLeft,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });

    return {
        scrollTop,
        scrollLeft,
        viewportHeight,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
        frozenRowCount,
        frozenColumnCount,
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        rowNumbers: rowWindow.rowNumbers,
        columnNumbers: columnWindow.columnNumbers,
    };
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

function getCellSnapshot(
    rowNumber: number,
    columnNumber: number,
    currentModel: EditorRenderModel | null = model
): CellSnapshot | null {
    if (!currentModel) {
        return null;
    }

    return currentModel.activeSheet.cells[createCellKey(rowNumber, columnNumber)] ?? null;
}

function getCellView(
    rowNumber: number,
    columnNumber: number,
    currentModel: EditorRenderModel | null = model
): EditorGridCellView | null {
    if (!currentModel) {
        return null;
    }

    const cell = getCellSnapshot(rowNumber, columnNumber, currentModel);
    return {
        key: createCellKey(rowNumber, columnNumber),
        address: cell?.address ?? `${getColumnLabel(columnNumber)}${rowNumber}`,
        value: cell?.displayValue ?? "",
        formula: cell?.formula ?? null,
        isPresent: Boolean(cell),
        isSelected:
            selectedCell?.rowNumber === rowNumber && selectedCell.columnNumber === columnNumber,
    };
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
    return (
        selectionRangeOverride ??
        createSelectionRange(selectionAnchorCell ?? selectedCell, selectedCell)
    );
}

function hasExpandedSelection(range: CellRange | null = getSelectionRange()): boolean {
    return hasExpandedSelectionRange(range);
}

function getExpandedSelectionRange(): CellRange | null {
    const range = getSelectionRange();
    return hasExpandedSelection(range) ? range : null;
}

function getSelectionRangeAddress(range: CellRange): string {
    return `${getCellAddressLabel(range.startRow, range.startColumn)}:${getCellAddressLabel(range.endRow, range.endColumn)}`;
}

function normalizeSearchPanelState(): void {
    const currentSelectionRange = getExpandedSelectionRange();

    if (searchScope !== "selection") {
        searchSelectionRange = currentSelectionRange;
        return;
    }

    if (!currentSelectionRange) {
        searchScope = "sheet";
        searchSelectionRange = null;
        searchFeedback = null;
        return;
    }

    searchSelectionRange = currentSelectionRange;
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

function getActiveHighlightCell(): CellPosition | null {
    return selectionAnchorCell ?? selectedCell;
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
        return getSelectionRangeAddress(range);
    }

    const cell = getCellView(selectedCell.rowNumber, selectedCell.columnNumber);
    return cell?.address ?? getCellAddressLabel(selectedCell.rowNumber, selectedCell.columnNumber);
}

function getActiveCellAddress(): string {
    if (!selectedCell) {
        return STRINGS.noCellSelected;
    }

    const cell = getCellView(selectedCell.rowNumber, selectedCell.columnNumber);
    return cell?.address ?? getCellAddressLabel(selectedCell.rowNumber, selectedCell.columnNumber);
}

function getSelectedCellToolbarValue(): string {
    if (!model || !selectedCell || hasExpandedSelection()) {
        return "";
    }

    if (
        editingCell &&
        editingCell.sheetKey === model.activeSheet.key &&
        editingCell.rowNumber === selectedCell.rowNumber &&
        editingCell.columnNumber === selectedCell.columnNumber
    ) {
        return editingCell.value;
    }

    return getEffectiveCellValue(selectedCell.rowNumber, selectedCell.columnNumber);
}

function canEditSelectedCellValue(): boolean {
    if (!model || !selectedCell || hasExpandedSelection()) {
        return false;
    }

    return canEditCellAt(selectedCell.rowNumber, selectedCell.columnNumber);
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
        (model.activeSheet.rowCount > 0 && model.activeSheet.columnCount > 0
            ? {
                  rowNumber: 1,
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

function closeContextMenu({ refresh = true }: { refresh?: boolean } = {}): void {
    if (!contextMenu) {
        return;
    }

    contextMenu = null;
    if (refresh) {
        renderApp({ commitEditing: false });
    }
}

function openTabContextMenu(sheetKey: string, x: number, y: number): void {
    contextMenu = { kind: "tab", sheetKey, x, y };
    renderApp({ commitEditing: false });
}

function openRowContextMenu(rowNumber: number, x: number, y: number): void {
    contextMenu = { kind: "row", rowNumber, x, y };
    renderApp({ commitEditing: false });
}

function openColumnContextMenu(columnNumber: number, x: number, y: number): void {
    contextMenu = { kind: "column", columnNumber, x, y };
    renderApp({ commitEditing: false });
}

function requestAddSheet(): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "addSheet" });
}

function requestDeleteSheet(sheetKey: string): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "deleteSheet", sheetKey });
}

function requestRenameSheet(sheetKey: string): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "renameSheet", sheetKey });
}

function requestInsertRow(rowNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "insertRow", rowNumber });
}

function requestDeleteRow(rowNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "deleteRow", rowNumber });
}

function requestInsertColumn(columnNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "insertColumn", columnNumber });
}

function requestDeleteColumn(columnNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "deleteColumn", columnNumber });
}

function beginSearchPanelDrag(pointerId: number, clientX: number, clientY: number): void {
    const appElement = getEditorAppElement();
    const panelElement = getSearchPanelShellElement();
    if (!appElement || !panelElement) {
        return;
    }

    const appRect = appElement.getBoundingClientRect();
    const panelRect = panelElement.getBoundingClientRect();
    const initialPosition = {
        left: panelRect.left - appRect.left,
        top: panelRect.top - appRect.top,
    };

    searchPanelPosition = clampSearchPanelPosition(initialPosition, {
        appElement,
        panelElement,
    });
    searchPanelDragState = {
        pointerId,
        offsetX: clientX - panelRect.left,
        offsetY: clientY - panelRect.top,
    };
    clearBrowserTextSelection();
    syncSearchPanelShellPosition();
}

function updateSearchPanelDrag(clientX: number, clientY: number): void {
    if (!searchPanelDragState) {
        return;
    }

    const appElement = getEditorAppElement();
    const panelElement = getSearchPanelShellElement();
    if (!appElement || !panelElement) {
        return;
    }

    const appRect = appElement.getBoundingClientRect();
    searchPanelPosition = clampSearchPanelPosition(
        {
            left: clientX - appRect.left - searchPanelDragState.offsetX,
            top: clientY - appRect.top - searchPanelDragState.offsetY,
        },
        {
            appElement,
            panelElement,
        }
    );
    syncSearchPanelShellPosition();
}

function stopSearchPanelDrag(pointerId?: number): void {
    if (!searchPanelDragState) {
        return;
    }

    if (pointerId !== undefined && searchPanelDragState.pointerId !== pointerId) {
        return;
    }

    searchPanelDragState = null;
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

function normalizeWheelDelta(delta: number, deltaMode: number, viewportSize: number): number {
    if (deltaMode === WHEEL_DELTA_LINE_MODE) {
        return delta * WHEEL_LINE_SCROLL_PIXELS;
    }

    if (deltaMode === WHEEL_DELTA_PAGE_MODE) {
        return delta * viewportSize;
    }

    return delta;
}

function forwardVirtualGridWheel(event: React.WheelEvent<HTMLElement>): void {
    const viewport = getViewportElement();
    if (!viewport) {
        return;
    }

    let deltaX = normalizeWheelDelta(event.deltaX, event.deltaMode, viewport.clientWidth);
    let deltaY = normalizeWheelDelta(event.deltaY, event.deltaMode, viewport.clientHeight);

    if (event.shiftKey && deltaX === 0 && deltaY !== 0) {
        deltaX = deltaY;
        deltaY = 0;
    }

    if (deltaX === 0 && deltaY === 0) {
        return;
    }

    event.preventDefault();
    viewport.scrollLeft = clampScrollPosition(
        viewport.scrollLeft + deltaX,
        viewport.scrollWidth - viewport.clientWidth
    );
    viewport.scrollTop = clampScrollPosition(
        viewport.scrollTop + deltaY,
        viewport.scrollHeight - viewport.clientHeight
    );
}

function revealSelectedCell(): void {
    if (!selectedCell || !model) {
        return;
    }

    const pane = getViewportElement();
    if (!pane) {
        return;
    }

    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: model.activeSheet.rowCount,
        columnCount: model.activeSheet.columnCount,
        viewportHeight: pane.clientHeight,
        viewportWidth: pane.clientWidth,
    });
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: model.activeSheet.rowCount,
        columnCount: model.activeSheet.columnCount,
        freezePane: model.activeSheet.freezePane,
    });
    const contentSize = getEditorContentSize({
        rowCount: displayGrid.rowCount,
        columnCount: displayGrid.columnCount,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const stickyTop =
        EDITOR_VIRTUAL_HEADER_HEIGHT + frozenRowCount * EDITOR_VIRTUAL_ROW_HEIGHT;
    const stickyLeft =
        displayGrid.rowHeaderWidth + frozenColumnCount * EDITOR_VIRTUAL_COLUMN_WIDTH;
    let nextTop = pane.scrollTop;
    let nextLeft = pane.scrollLeft;

    if (selectedCell.rowNumber > frozenRowCount) {
        const cellTop =
            EDITOR_VIRTUAL_HEADER_HEIGHT +
            (selectedCell.rowNumber - 1) * EDITOR_VIRTUAL_ROW_HEIGHT;
        const cellBottom = cellTop + EDITOR_VIRTUAL_ROW_HEIGHT;
        const visibleTop = pane.scrollTop + stickyTop;
        const visibleBottom = pane.scrollTop + pane.clientHeight;

        if (cellTop < visibleTop) {
            nextTop = cellTop - stickyTop;
        } else if (cellBottom > visibleBottom) {
            nextTop = cellBottom - pane.clientHeight;
        }
    }

    if (selectedCell.columnNumber > frozenColumnCount) {
        const cellLeft =
            displayGrid.rowHeaderWidth +
            (selectedCell.columnNumber - 1) * EDITOR_VIRTUAL_COLUMN_WIDTH;
        const cellRight = cellLeft + EDITOR_VIRTUAL_COLUMN_WIDTH;
        const visibleLeft = pane.scrollLeft + stickyLeft;
        const visibleRight = pane.scrollLeft + pane.clientWidth;

        if (cellLeft < visibleLeft) {
            nextLeft = cellLeft - stickyLeft;
        } else if (cellRight > visibleRight) {
            nextLeft = cellRight - pane.clientWidth;
        }
    }

    pane.scrollTop = clampEditorScrollPosition(nextTop, contentSize.height - pane.clientHeight);
    pane.scrollLeft = clampEditorScrollPosition(nextLeft, contentSize.width - pane.clientWidth);
}

function isCellVisible(cell: CellPosition | null): boolean {
    if (!cell || !model) {
        return false;
    }

    const viewportState = getVirtualViewportState(model);
    if (!viewportState) {
        return false;
    }

    const rowVisible =
        viewportState.frozenRowNumbers.includes(cell.rowNumber) ||
        viewportState.rowNumbers.includes(cell.rowNumber);
    const columnVisible =
        viewportState.frozenColumnNumbers.includes(cell.columnNumber) ||
        viewportState.columnNumbers.includes(cell.columnNumber);

    return rowVisible && columnVisible;
}

function isRowVisibleInViewport(
    rowNumber: number | null | undefined,
    viewportState: VirtualViewportState | null
): boolean {
    return Boolean(
        rowNumber &&
            viewportState &&
            (viewportState.frozenRowNumbers.includes(rowNumber) ||
                viewportState.rowNumbers.includes(rowNumber))
    );
}

function isColumnVisibleInViewport(
    columnNumber: number | null | undefined,
    viewportState: VirtualViewportState | null
): boolean {
    return Boolean(
        columnNumber &&
            viewportState &&
            (viewportState.frozenColumnNumbers.includes(columnNumber) ||
                viewportState.columnNumbers.includes(columnNumber))
    );
}

function getPreferredSelectionRowNumber(): number {
    if (!model) {
        return 1;
    }

    const viewportState = getVirtualViewportState(model);
    const candidate = selectedCell?.rowNumber;
    if (
        candidate &&
        candidate >= 1 &&
        candidate <= model.activeSheet.rowCount &&
        (!viewportState || isRowVisibleInViewport(candidate, viewportState))
    ) {
        return candidate;
    }

    return (
        viewportState?.frozenRowNumbers[0] ??
        viewportState?.rowNumbers[0] ??
        Math.min(model.activeSheet.rowCount, Math.max(candidate ?? 1, 1))
    );
}

function getPreferredSelectionColumnNumber(): number {
    if (!model) {
        return 1;
    }

    const viewportState = getVirtualViewportState(model);
    const candidate = selectedCell?.columnNumber;
    if (
        candidate &&
        candidate >= 1 &&
        candidate <= model.activeSheet.columnCount &&
        (!viewportState || isColumnVisibleInViewport(candidate, viewportState))
    ) {
        return candidate;
    }

    return (
        viewportState?.frozenColumnNumbers[0] ??
        viewportState?.columnNumbers[0] ??
        Math.min(model.activeSheet.columnCount, Math.max(candidate ?? 1, 1))
    );
}

function selectEntireRow(rowNumber: number): void {
    if (!model) {
        return;
    }

    const selectionRange = createRowSelectionRange(rowNumber, model.activeSheet.columnCount);
    if (!selectionRange) {
        return;
    }

    const forceRender = Boolean(editingCell);
    if (forceRender) {
        finishEdit({ mode: "commit", refresh: false });
    }

    setSelectedCellLocal(
        {
            rowNumber,
            columnNumber: getPreferredSelectionColumnNumber(),
        },
        {
            syncHost: true,
            selectionRange,
            forceRender,
        }
    );
}

function selectEntireColumn(columnNumber: number): void {
    if (!model) {
        return;
    }

    const selectionRange = createColumnSelectionRange(
        columnNumber,
        model.activeSheet.rowCount
    );
    if (!selectionRange) {
        return;
    }

    const forceRender = Boolean(editingCell);
    if (forceRender) {
        finishEdit({ mode: "commit", refresh: false });
    }

    setSelectedCellLocal(
        {
            rowNumber: getPreferredSelectionRowNumber(),
            columnNumber,
        },
        {
            syncHost: true,
            selectionRange,
            forceRender,
        }
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
    anchorCell: CellPosition | null = selectionAnchorCell,
    selectionRange: CellRange | null = selectionRangeOverride
): boolean {
    if (!cell) {
        return anchorCell === null && selectionRange === null;
    }

    return !hasExpandedSelectionRange(selectionRange ?? createSelectionRange(anchorCell ?? cell, cell));
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
        `[data-role="grid-column-header"][data-column-number="${columnNumber}"]`
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

function canUseLocalSimpleSelectionUpdate(
    nextCell: CellPosition | null,
    {
        anchorCell,
        selectionRange,
        forceRender = false,
    }: {
        anchorCell: CellPosition | null;
        selectionRange: CellRange | null;
        forceRender?: boolean;
    }
): boolean {
    return shouldUseLocalSimpleSelectionUpdate({
        hasNextCell: Boolean(nextCell),
        hasModel: Boolean(model),
        hasEditingCell: Boolean(editingCell),
        isNextCellVisible: isCellVisible(nextCell),
        hasExpandedSelection: hasExpandedSelection(),
        isSimpleSelection: isSimpleSelection(nextCell, anchorCell, selectionRange),
        forceRender,
    });
}

function setSelectedCellLocal(
    nextCell: CellPosition | null,
    {
        reveal = false,
        syncHost = true,
        anchorCell,
        selectionRange,
        clearSearchFeedback = true,
        forceRender = false,
    }: {
        reveal?: boolean;
        syncHost?: boolean;
        anchorCell?: CellPosition | null;
        selectionRange?: CellRange | null;
        clearSearchFeedback?: boolean;
        forceRender?: boolean;
    } = {}
): void {
    const previousCell = selectedCell;
    const nextAnchorCell = nextCell ? (anchorCell ?? nextCell) : null;
    const nextSelectionRange =
        nextCell && nextAnchorCell
            ? (selectionRange ?? createSelectionRange(nextAnchorCell, nextCell))
            : null;
    const shouldClearSearchFeedback = clearSearchFeedback && Boolean(searchFeedback);
    const canUseLocalUpdate = canUseLocalSimpleSelectionUpdate(nextCell, {
        anchorCell: nextAnchorCell,
        selectionRange: nextSelectionRange,
        forceRender: forceRender || shouldClearSearchFeedback,
    });

    selectedCell = nextCell;
    selectionAnchorCell = nextAnchorCell;
    selectionRangeOverride = selectionRange ?? null;
    suppressAutoSelection = nextCell === null;
    clearBrowserTextSelection();
    if (shouldClearSearchFeedback) {
        searchFeedback = null;
    }

    if (syncHost) {
        syncSelectedCellToHost();
    }

    if (canUseLocalUpdate) {
        syncLocalSimpleSelectionDom(previousCell, nextCell);
        notifyEditorToolbarSync();
        if (reveal) {
            revealSelectedCell();
        }
        return;
    }

    renderApp({
        commitEditing: false,
        revealSelection: reveal,
    });
}

function prepareSelectionForRender({
    revealSelection,
    useModelSelection,
}: {
    revealSelection: boolean;
    useModelSelection: boolean;
}): boolean {
    let shouldReveal = revealSelection;
    const previousCell = selectedCell;
    const previousAnchorCell = selectionAnchorCell;

    if (useModelSelection && model?.selection) {
        selectedCell = {
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        };
        selectionAnchorCell = selectedCell;
        selectionRangeOverride = null;
        suppressAutoSelection = false;
        pendingSelectionAfterRender = null;
        if (
            shouldSyncLocalSelectionDomFromModelSelection(
                previousCell,
                previousAnchorCell,
                selectedCell
            )
        ) {
            syncLocalSimpleSelectionDom(previousCell, selectedCell);
        }
        syncSelectedCellToHost();
        return shouldReveal;
    }

    if (pendingSelectionAfterRender) {
        selectedCell = {
            rowNumber: pendingSelectionAfterRender.rowNumber,
            columnNumber: pendingSelectionAfterRender.columnNumber,
        };
        selectionAnchorCell = selectedCell;
        selectionRangeOverride = null;
        suppressAutoSelection = false;
        shouldReveal = pendingSelectionAfterRender.reveal;
        pendingSelectionAfterRender = null;
        syncSelectedCellToHost();
        return shouldReveal;
    }

    pendingSelectionAfterRender = null;

    if (isCellVisible(selectedCell)) {
        if (
            shouldResetInvisibleSelectionAnchor({
                hasSelectionRangeOverride: Boolean(selectionRangeOverride),
                hasExpandedSelection: hasExpandedSelection(),
                isAnchorVisible: isCellVisible(selectionAnchorCell),
            })
        ) {
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
        : model && model.activeSheet.rowCount > 0 && model.activeSheet.columnCount > 0
          ? { rowNumber: 1, columnNumber: 1 }
          : null;
    selectionRangeOverride = null;

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
        selectionRangeOverride = null;
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
    selectionRangeOverride = null;
    suppressAutoSelection = false;
    editingCell = {
        sheetKey: model.activeSheet.key,
        rowNumber,
        columnNumber,
        value: currentValue,
    };
    syncSelectedCellToHost();
    renderApp({
        commitEditing: false,
        revealSelection: true,
    });
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

    if (model.activeSheet.rowCount === 0 || model.activeSheet.columnCount === 0) {
        return null;
    }

    return {
        minRow: 1,
        maxRow: model.activeSheet.rowCount,
        minColumn: 1,
        maxColumn: model.activeSheet.columnCount,
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
        selectionRangeOverride = null;
        suppressAutoSelection = false;
        return selectedCell;
    }

    if (model && model.activeSheet.rowCount > 0 && model.activeSheet.columnCount > 0) {
        selectedCell = {
            rowNumber: 1,
            columnNumber: 1,
        };
        selectionAnchorCell = selectedCell;
        selectionRangeOverride = null;
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

    const viewportState = getVirtualViewportState(model);
    const shiftRowCount = Math.max(
        1,
        Math.floor(
            Math.max(
                (viewportState?.viewportHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT) -
                    EDITOR_VIRTUAL_HEADER_HEIGHT,
                EDITOR_VIRTUAL_ROW_HEIGHT
            ) / EDITOR_VIRTUAL_ROW_HEIGHT
        ) - 1
    );
    const nextRowNumber = Math.max(
        1,
        Math.min(model.activeSheet.rowCount, selection.rowNumber + direction * shiftRowCount)
    );
    if (nextRowNumber === selection.rowNumber) {
        return;
    }

    setSelectedCellLocal(
        {
            rowNumber: nextRowNumber,
            columnNumber: selection.columnNumber,
        },
        { reveal: true }
    );
}

function submitGoto(reference: string): void {
    const normalizedReference = reference.trim();
    if (!normalizedReference) {
        return;
    }

    vscode.postMessage({ type: "gotoCell", reference: normalizedReference });
}

function commitSelectedCellValue(value: string): void {
    if (!model || !selectedCell || !canEditSelectedCellValue()) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    commitEdit(model.activeSheet.key, selectedCell.rowNumber, selectedCell.columnNumber, value, {
        refresh: true,
        revealSelection: true,
    });
}

function commitToolbarCellValue(target: ToolbarCellEditTarget, value: string): void {
    if (!model || target.sheetKey !== model.activeSheet.key || !model.canEdit) {
        return;
    }

    if (!canEditCellAt(target.rowNumber, target.columnNumber)) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    commitEdit(target.sheetKey, target.rowNumber, target.columnNumber, value, {
        refresh: true,
        revealSelection: true,
    });
}

function focusToolbarInput(role: "position" | "value"): void {
    const selector =
        role === "position"
            ? '[data-role="position-input"]'
            : '[data-role="cell-value-input"]';
    const input = document.querySelector<HTMLInputElement>(selector);
    if (!input) {
        return;
    }

    input.focus();
    input.select();
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

function openSearchPanel(mode: SearchPanelMode = "find"): void {
    let shouldRender = false;

    if (!isSearchPanelOpen) {
        const currentSelectionRange = getExpandedSelectionRange();
        isSearchPanelOpen = true;
        searchScope = currentSelectionRange ? "selection" : "sheet";
        searchSelectionRange = currentSelectionRange;
        searchFeedback = null;
        shouldRender = true;
    }

    if (searchMode !== mode) {
        searchMode = mode;
        shouldRender = true;
    }

    if (shouldRender) {
        renderApp({ commitEditing: false });
    }

    requestAnimationFrame(() => {
        if (mode === "replace" && searchQuery.trim().length > 0) {
            focusReplaceInput();
            return;
        }

        focusSearchInput();
    });
}

function closeSearchPanel(): void {
    if (!isSearchPanelOpen) {
        return;
    }

    stopSearchPanelDrag();
    isSearchPanelOpen = false;
    searchScope = "sheet";
    searchSelectionRange = null;
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function updateSearchQuery(value: string): void {
    if (searchQuery === value && !searchFeedback) {
        return;
    }

    searchQuery = value;
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function updateReplaceValue(value: string): void {
    if (replaceValue === value && !searchFeedback) {
        return;
    }

    replaceValue = value;
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function setSearchMode(mode: SearchPanelMode): void {
    if (searchMode === mode) {
        return;
    }

    searchMode = mode;
    searchFeedback = null;
    renderApp({ commitEditing: false });
    requestAnimationFrame(() => {
        if (mode === "replace" && searchQuery.trim().length > 0) {
            focusReplaceInput();
            return;
        }

        focusSearchInput();
    });
}

function updateSearchScope(scope: EditorSearchScope): void {
    if (scope === "selection") {
        const currentSelectionRange = getExpandedSelectionRange();
        if (!currentSelectionRange) {
            return;
        }

        searchSelectionRange = currentSelectionRange;
    } else {
        searchSelectionRange = getExpandedSelectionRange();
    }

    if (searchScope === scope && !searchFeedback) {
        return;
    }

    searchScope = scope;
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function toggleSearchOption(option: keyof SearchOptions): void {
    searchOptions = {
        ...searchOptions,
        [option]: !searchOptions[option],
    };
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function getEffectiveSearchScope(): EditorSearchScope {
    return searchScope === "selection" && Boolean(searchSelectionRange ?? getExpandedSelectionRange())
        ? "selection"
        : "sheet";
}

function getActiveSearchSelectionRange(scope: EditorSearchScope): CellRange | null {
    return scope === "selection" ? (searchSelectionRange ?? getExpandedSelectionRange()) : null;
}

function getActiveSheetPendingEdits(): Array<{
    rowNumber: number;
    columnNumber: number;
    value: string;
}> {
    if (!model) {
        return [];
    }

    return Array.from(pendingEdits.values())
        .filter((edit) => edit.sheetKey === model!.activeSheet.key)
        .map((edit) => ({
            rowNumber: edit.rowNumber,
            columnNumber: edit.columnNumber,
            value: edit.value,
        }));
}

function revealSearchPanelMatch(
    match: { rowNumber: number; columnNumber: number },
    scope: EditorSearchScope,
    { syncHost = true }: { syncHost?: boolean } = {}
): void {
    const preservedSelectionRange = getActiveSearchSelectionRange(scope);
    const fallbackAnchor = {
        rowNumber: match.rowNumber,
        columnNumber: match.columnNumber,
    };

    setSelectedCellLocal(
        {
            rowNumber: match.rowNumber,
            columnNumber: match.columnNumber,
        },
        {
            reveal: true,
            syncHost,
            anchorCell:
                scope === "selection"
                    ? (selectionAnchorCell ?? selectedCell ?? fallbackAnchor)
                    : undefined,
            selectionRange: preservedSelectionRange ?? undefined,
            clearSearchFeedback: false,
            forceRender: true,
        }
    );
}

function submitSearch(direction: "next" | "prev"): void {
    const normalizedQuery = searchQuery.trim();
    if (!normalizedQuery) {
        focusSearchInput();
        return;
    }

    const effectiveScope = getEffectiveSearchScope();
    const selectionRange = getActiveSearchSelectionRange(effectiveScope);

    searchFeedback = null;
    renderApp({ commitEditing: false });
    vscode.postMessage({
        type: "search",
        query: normalizedQuery,
        direction,
        options: searchOptions,
        scope: effectiveScope,
        selectionRange: selectionRange ?? undefined,
    });
}

function getSearchMatchFeedbackMessage(
    match: { rowNumber: number; columnNumber: number },
    scope: EditorSearchScope,
    selectionRange: CellRange | null,
    matchCount?: number,
    matchIndex?: number
): string {
    const address = getCellAddressLabel(match.rowNumber, match.columnNumber);
    const locationMessage =
        scope === "selection" && selectionRange
            ? formatI18nMessage(STRINGS.searchMatchFoundInSelection, {
                  address,
                  range: getSelectionRangeAddress(selectionRange),
              })
            : formatI18nMessage(STRINGS.searchMatchFound, { address });
    if (!matchCount || !matchIndex) {
        return locationMessage;
    }

    const summaryMessage = formatI18nMessage(STRINGS.searchMatchSummary, {
        count: matchCount,
        index: matchIndex,
    });

    return `${locationMessage} ${summaryMessage}`;
}

function submitReplace(mode: "single" | "all"): void {
    const normalizedQuery = searchQuery.trim();
    if (!normalizedQuery) {
        focusSearchInput();
        return;
    }

    if (!model || !model.canEdit) {
        return;
    }

    const effectiveScope = getEffectiveSearchScope();
    const selectionRange = getActiveSearchSelectionRange(effectiveScope);
    const result = resolveEditorReplaceResultInSheet(
        {
            key: model.activeSheet.key,
            rowCount: model.activeSheet.rowCount,
            columnCount: model.activeSheet.columnCount,
            cells: model.activeSheet.cells,
        },
        selectedCell,
        {
            query: normalizedQuery,
            replacement: replaceValue,
            options: searchOptions,
            scope: effectiveScope,
            selectionRange: selectionRange ?? undefined,
            pendingEdits: getActiveSheetPendingEdits(),
            mode,
        }
    );

    if (result.status === "invalid-pattern") {
        searchFeedback = {
            status: "invalid-pattern",
            message: STRINGS.invalidSearchPattern,
        };
        renderApp({ commitEditing: false, sync: true });
        return;
    }

    if (result.status === "no-match") {
        searchFeedback = {
            status: "no-match",
            message: STRINGS.replaceNoEditableMatches,
        };
        renderApp({ commitEditing: false, sync: true });
        return;
    }

    if (result.status === "no-change") {
        searchFeedback = {
            status: "no-change",
            message: STRINGS.replaceNoChanges,
        };

        if (result.match) {
            revealSearchPanelMatch(result.match, effectiveScope);
            return;
        }

        renderApp({ commitEditing: false, sync: true });
        return;
    }

    const activeSheetKey = model.activeSheet.key;
    const changes = (result.changes ?? []).map((change) => ({
        sheetKey: activeSheetKey,
        rowNumber: change.rowNumber,
        columnNumber: change.columnNumber,
        modelValue: getCellModelValue(change.rowNumber, change.columnNumber),
        beforeValue: change.beforeValue,
        afterValue: change.afterValue,
    }));

    applyEditChanges(changes, { refresh: false });
    searchFeedback = {
        status: "replaced",
        message: formatI18nMessage(STRINGS.replaceCount, {
            count: result.replacedCellCount ?? changes.length,
        }),
    };

    if (result.match) {
        revealSearchPanelMatch(result.nextMatch ?? result.match, effectiveScope);
        return;
    }

    renderApp({ commitEditing: false, sync: true });
}

function handleSearchResult(message: EditorSearchResultMessage): void {
    if (message.status !== "matched" || !message.match) {
        searchFeedback = {
            status: message.status,
            message: message.message,
        };
        renderApp({ commitEditing: false, sync: true });
        return;
    }

    searchFeedback = {
        status: "matched",
        message: getSearchMatchFeedbackMessage(
            message.match,
            message.scope,
            getActiveSearchSelectionRange(message.scope),
            message.matchCount,
            message.matchIndex
        ),
    };
    revealSearchPanelMatch(message.match, message.scope, { syncHost: false });
}

function getPositionInputValue(): string {
    return selectedCell ? getSelectedCellAddress() : "";
}

function getCellValueInputValue(): string {
    return selectedCell ? getSelectedCellToolbarValue() : "";
}

function getCellValueInputPlaceholder(): string {
    if (!selectedCell) {
        return STRINGS.noCellSelected;
    }

    return hasExpandedSelection() ? STRINGS.multipleCellsSelected : "";
}

function submitGotoSelection(reference: string): void {
    if (!reference) {
        return;
    }

    submitGoto(reference);
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

function updateView(view: ViewState, { sync = false }: { sync?: boolean } = {}): void {
    const setView = setViewState;
    if (!setView) {
        return;
    }

    if (sync || view.kind !== "app") {
        flushSync(() => {
            setView(view);
        });
        return;
    }

    setView(view);
}

function renderLoading(message: string): void {
    updateView({ kind: "loading", message }, { sync: true });
}

function renderError(message: string): void {
    updateView({ kind: "error", message }, { sync: true });
}

function renderApp({
    commitEditing = true,
    revealSelection = false,
    useModelSelection = false,
    sync = false,
}: {
    commitEditing?: boolean;
    revealSelection?: boolean;
    useModelSelection?: boolean;
    sync?: boolean;
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
    normalizeSearchPanelState();
    viewRevision += 1;
    updateView({
        kind: "app",
        model,
        revealSelection: shouldRevealSelection,
        revision: viewRevision,
        scrollState,
    }, { sync });
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
    isLoading = false,
    iconOnly = false,
    iconMirrored = false,
    onClick,
}: {
    actionLabel: string;
    icon: string;
    disabled?: boolean;
    isActive?: boolean;
    isLoading?: boolean;
    iconOnly?: boolean;
    iconMirrored?: boolean;
    onClick(): void;
}): React.ReactElement {
    const displayedIcon = isLoading ? "codicon-loading" : icon;

    return (
        <button
            aria-label={actionLabel}
            className={classNames([
                "toolbar__button",
                isActive && "is-active",
                isLoading && "is-loading",
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
                    displayedIcon,
                    "toolbar__button-icon",
                    iconMirrored && "toolbar__button-icon--flip",
                    isLoading && "toolbar__button-icon--spin",
                ])}
                aria-hidden
            />
            {iconOnly ? null : <span>{actionLabel}</span>}
        </button>
    );
}

function SearchPanel({
    currentModel,
}: {
    currentModel: EditorRenderModel;
}): React.ReactElement | null {
    const shellRef = React.useRef<HTMLElement | null>(null);

    React.useLayoutEffect(() => {
        if (!isSearchPanelOpen) {
            return;
        }

        const shell = shellRef.current;
        if (!shell || !searchPanelPosition) {
            return;
        }

        const nextPosition = clampSearchPanelPosition(searchPanelPosition, {
            panelElement: shell,
        });
        if (
            nextPosition.left !== searchPanelPosition.left ||
            nextPosition.top !== searchPanelPosition.top
        ) {
            searchPanelPosition = nextPosition;
            shell.style.left = `${nextPosition.left}px`;
            shell.style.top = `${nextPosition.top}px`;
            shell.style.right = "auto";
        }
    });

    if (!isSearchPanelOpen) {
        return null;
    }

    const currentSelectionRange = getExpandedSelectionRange();
    const selectionRange =
        searchScope === "selection"
            ? (searchSelectionRange ?? currentSelectionRange)
            : currentSelectionRange;
    const isSelectionScopeAvailable = Boolean(currentSelectionRange);
    const effectiveScope = getEffectiveSearchScope();
    const isReplaceMode = searchMode === "replace";
    const isQueryEmpty = searchQuery.trim().length === 0;
    const hasSearchableGrid =
        currentModel.activeSheet.rowCount > 0 && currentModel.activeSheet.columnCount > 0;
    const canReplace = !isQueryEmpty && hasSearchableGrid && currentModel.canEdit;
    const feedbackToneClass =
        searchFeedback?.status === "matched" || searchFeedback?.status === "replaced"
            ? "search-strip__feedback--success"
            : searchFeedback?.status === "invalid-pattern"
            ? "search-strip__feedback--error"
            : searchFeedback?.status === "no-match" || searchFeedback?.status === "no-change"
              ? "search-strip__feedback--warn"
              : undefined;
    const panelStyle: React.CSSProperties | undefined = searchPanelPosition
        ? {
              left: `${searchPanelPosition.left}px`,
              top: `${searchPanelPosition.top}px`,
              right: "auto",
          }
        : undefined;

    return (
        <section
            ref={shellRef}
            className="search-strip-shell"
            data-role="search-panel-shell"
            style={panelStyle}
            onPointerDown={(event) => {
                if (event.button !== 0 || isSearchPanelInteractiveTarget(event.target)) {
                    return;
                }

                event.preventDefault();
                closeContextMenu({ refresh: false });
                beginSearchPanelDrag(event.pointerId, event.clientX, event.clientY);
            }}
        >
            <div className="search-strip" data-role="search-panel" role="search">
                <div className="search-strip__header">
                    <div className="search-strip__tabs" role="tablist" aria-label={STRINGS.search}>
                        <button
                            aria-selected={!isReplaceMode}
                            className={classNames([
                                "search-strip__tab",
                                !isReplaceMode && "is-active",
                            ])}
                            role="tab"
                            tabIndex={!isReplaceMode ? 0 : -1}
                            type="button"
                            onClick={() => setSearchMode("find")}
                        >
                            <span className="codicon codicon-search search-strip__tab-icon" aria-hidden />
                            <span>{STRINGS.searchFind}</span>
                        </button>
                        <button
                            aria-selected={isReplaceMode}
                            className={classNames([
                                "search-strip__tab",
                                isReplaceMode && "is-active",
                            ])}
                            role="tab"
                            tabIndex={isReplaceMode ? 0 : -1}
                            type="button"
                            onClick={() => setSearchMode("replace")}
                        >
                            <span className="codicon codicon-replace search-strip__tab-icon" aria-hidden />
                            <span>{STRINGS.searchReplace}</span>
                        </button>
                    </div>
                </div>
                <div className="search-strip__row search-strip__row--primary">
                    <div
                        className={classNames([
                            "search-strip__input-wrap",
                            searchFeedback?.status === "invalid-pattern" && "is-invalid",
                        ])}
                    >
                        <span
                            className="codicon codicon-search search-strip__input-icon"
                            aria-hidden
                        />
                        <input
                            aria-label={STRINGS.search}
                            className="search-strip__input"
                            data-role="search-input"
                            placeholder={STRINGS.searchPlaceholder}
                            type="text"
                            value={searchQuery}
                            onChange={(event) => {
                                updateSearchQuery(event.currentTarget.value);
                            }}
                            onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                    event.preventDefault();
                                    submitSearch(event.shiftKey ? "prev" : "next");
                                    return;
                                }

                                if (event.key === "Escape") {
                                    event.preventDefault();
                                    event.stopPropagation();
                                    closeSearchPanel();
                                }
                            }}
                        />
                        <div className="search-strip__input-tools" role="group" aria-label={STRINGS.search}>
                            <button
                                aria-label={STRINGS.searchRegex}
                                aria-pressed={searchOptions.isRegexp}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    searchOptions.isRegexp && "is-active",
                                ])}
                                title={STRINGS.searchRegex}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => toggleSearchOption("isRegexp")}
                            >
                                <span className="codicon codicon-regex" aria-hidden />
                            </button>
                            <button
                                aria-label={STRINGS.searchMatchCase}
                                aria-pressed={searchOptions.matchCase}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    searchOptions.matchCase && "is-active",
                                ])}
                                title={STRINGS.searchMatchCase}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => toggleSearchOption("matchCase")}
                            >
                                <span className="codicon codicon-case-sensitive" aria-hidden />
                            </button>
                            <button
                                aria-label={STRINGS.searchWholeWord}
                                aria-pressed={searchOptions.wholeWord}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    searchOptions.wholeWord && "is-active",
                                ])}
                                title={STRINGS.searchWholeWord}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => toggleSearchOption("wholeWord")}
                            >
                                <span className="codicon codicon-whole-word" aria-hidden />
                            </button>
                        </div>
                    </div>
                    <div className="search-strip__actions">
                        <ToolbarButton
                            actionLabel={STRINGS.findPrev}
                            disabled={isQueryEmpty || !hasSearchableGrid}
                            icon="codicon-arrow-up"
                            iconOnly={true}
                            onClick={() => submitSearch("prev")}
                        />
                        <ToolbarButton
                            actionLabel={STRINGS.findNext}
                            disabled={isQueryEmpty || !hasSearchableGrid}
                            icon="codicon-arrow-down"
                            iconOnly={true}
                            onClick={() => submitSearch("next")}
                        />
                        <ToolbarButton
                            actionLabel={STRINGS.searchClose}
                            icon="codicon-close"
                            iconOnly={true}
                            onClick={closeSearchPanel}
                        />
                    </div>
                </div>
                {isReplaceMode ? (
                    <div className="search-strip__row search-strip__row--replace">
                        <div className="search-strip__input-wrap">
                            <span
                                className="codicon codicon-replace search-strip__input-icon"
                                aria-hidden
                            />
                            <input
                                aria-label={STRINGS.searchReplace}
                                className="search-strip__input"
                                data-role="replace-input"
                                placeholder={STRINGS.replacePlaceholder}
                                type="text"
                                value={replaceValue}
                                onChange={(event) => {
                                    updateReplaceValue(event.currentTarget.value);
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
                        <div className="search-strip__replace-actions">
                            <ToolbarButton
                                actionLabel={STRINGS.searchReplace}
                                disabled={!canReplace}
                                icon="codicon-replace"
                                iconOnly={true}
                                onClick={() => submitReplace("single")}
                            />
                            <ToolbarButton
                                actionLabel={STRINGS.replaceAll}
                                disabled={!canReplace}
                                icon="codicon-replace-all"
                                iconOnly={true}
                                onClick={() => submitReplace("all")}
                            />
                        </div>
                    </div>
                ) : null}
                <div className="search-strip__row search-strip__row--meta">
                    <div className="search-strip__scope-group">
                        <button
                            aria-pressed={effectiveScope === "sheet"}
                            className={classNames([
                                "search-strip__scope-button",
                                effectiveScope === "sheet" && "is-active",
                            ])}
                            type="button"
                            onClick={() => updateSearchScope("sheet")}
                        >
                            {STRINGS.searchScopeSheet}
                        </button>
                        <button
                            aria-pressed={effectiveScope === "selection"}
                            className={classNames([
                                "search-strip__scope-button",
                                effectiveScope === "selection" && "is-active",
                            ])}
                            disabled={!isSelectionScopeAvailable}
                            title={
                                isSelectionScopeAvailable
                                    ? STRINGS.searchScopeSelection
                                    : STRINGS.searchScopeSelectionDisabled
                            }
                            type="button"
                            onClick={() => updateSearchScope("selection")}
                        >
                            {STRINGS.searchScopeSelection}
                        </button>
                    </div>
                    {selectionRange ? (
                        <span className="search-strip__range">
                            {getSelectionRangeAddress(selectionRange)}
                        </span>
                    ) : (
                        <span className="search-strip__hint">
                            {STRINGS.searchScopeSelectionDisabled}
                        </span>
                    )}
                </div>
                {searchFeedback?.message ? (
                    <div className={classNames(["search-strip__feedback", feedbackToneClass])}>
                        {searchFeedback.message}
                    </div>
                ) : null}
            </div>
        </section>
    );
}

function EditorToolbar({ currentModel }: { currentModel: EditorRenderModel }): React.ReactElement {
    React.useSyncExternalStore(
        subscribeEditorToolbarSync,
        getEditorToolbarSyncSnapshot,
        getEditorToolbarSyncSnapshot
    );

    const hasPendingEdits = pendingEdits.size > 0 || currentModel.hasPendingEdits;
    const canUndo = undoStack.length > 0 || currentModel.canUndoStructuralEdits;
    const canRedo = redoStack.length > 0 || currentModel.canRedoStructuralEdits;
    const viewLocked = hasLockedView(currentModel.activeSheet.freezePane);
    const viewLockActionLabel = viewLocked ? STRINGS.unlockView : STRINGS.lockView;
    const activeCellAddress = getPositionInputValue();
    const selectedCellValue = getCellValueInputValue();
    const cellValuePlaceholder = getCellValueInputPlaceholder();
    const selectedCellEditable = currentModel.canEdit && canEditSelectedCellValue();
    const activeCellEditTarget: ToolbarCellEditTarget | null = selectedCell
        ? {
              sheetKey: currentModel.activeSheet.key,
              rowNumber: selectedCell.rowNumber,
              columnNumber: selectedCell.columnNumber,
          }
        : null;
    const activeCellEditTargetKey = getToolbarCellEditTargetKey(activeCellEditTarget);
    const [positionInputValue, setPositionInputValue] = React.useState(activeCellAddress);
    const [cellValueInputValue, setCellValueInputValue] = React.useState(selectedCellValue);
    const [isEditingPosition, setIsEditingPosition] = React.useState(false);
    const [isEditingCellValue, setIsEditingCellValue] = React.useState(false);
    const [cellValueEditTarget, setCellValueEditTarget] =
        React.useState<ToolbarCellEditTarget | null>(null);
    const cellValueEditTargetKey = getToolbarCellEditTargetKey(cellValueEditTarget);
    const showCellValueActions =
        isEditingCellValue &&
        !shouldResetToolbarCellValueDraft(
            cellValueEditTarget,
            activeCellEditTarget,
            selectedCellEditable
        );

    React.useEffect(() => {
        if (!isEditingPosition) {
            setPositionInputValue(activeCellAddress);
        }
    }, [activeCellAddress, isEditingPosition]);

    React.useEffect(() => {
        if (
            isEditingCellValue &&
            shouldResetToolbarCellValueDraft(
                cellValueEditTarget,
                activeCellEditTarget,
                selectedCellEditable
            )
        ) {
            setIsEditingCellValue(false);
            setCellValueEditTarget(null);
            setCellValueInputValue(selectedCellValue);
            return;
        }

        if (!isEditingCellValue) {
            setCellValueInputValue(selectedCellValue);
        }
    }, [
        activeCellEditTargetKey,
        cellValueEditTarget,
        cellValueEditTargetKey,
        isEditingCellValue,
        selectedCellEditable,
        selectedCellValue,
    ]);

    const resetPositionInput = () => {
        setIsEditingPosition(false);
        setPositionInputValue(activeCellAddress);
    };

    const commitPositionInput = () => {
        const nextReference = positionInputValue.trim();
        setIsEditingPosition(false);
        if (!nextReference) {
            setPositionInputValue(activeCellAddress);
            return;
        }

        submitGotoSelection(nextReference);
    };

    const resetCellValueInput = () => {
        setIsEditingCellValue(false);
        setCellValueEditTarget(null);
        setCellValueInputValue(selectedCellValue);
    };

    const commitCellValueInput = () => {
        const target = cellValueEditTarget;
        setIsEditingCellValue(false);
        setCellValueEditTarget(null);
        if (
            !target ||
            shouldResetToolbarCellValueDraft(target, activeCellEditTarget, selectedCellEditable)
        ) {
            setCellValueInputValue(selectedCellValue);
            return;
        }

        commitToolbarCellValue(target, cellValueInputValue);
    };

    return (
        <header className="toolbar toolbar--editor">
            <div className="toolbar__group toolbar__group--grow">
                <label className="toolbar__field toolbar__field--address">
                    <span className="toolbar__field-label">#</span>
                    <input
                        className="toolbar__input"
                        data-role="position-input"
                        value={positionInputValue}
                        placeholder={STRINGS.gotoPlaceholder}
                        type="text"
                        onFocus={(event) => {
                            setIsEditingPosition(true);
                            event.currentTarget.select();
                        }}
                        onChange={(event) => {
                            setPositionInputValue(event.currentTarget.value);
                        }}
                        onBlur={() => {
                            resetPositionInput();
                        }}
                        onKeyDown={(event) => {
                            if (event.key === "Enter") {
                                event.preventDefault();
                                commitPositionInput();
                                return;
                            }

                            if (event.key === "Escape") {
                                event.preventDefault();
                                resetPositionInput();
                            }
                        }}
                    />
                </label>
                <label className="toolbar__field toolbar__field--cell-value">
                    <span className="toolbar__field-label">T</span>
                    <input
                        className="toolbar__input"
                        data-role="cell-value-input"
                        value={cellValueInputValue}
                        placeholder={cellValuePlaceholder}
                        readOnly={!selectedCellEditable}
                        type="text"
                        onFocus={(event) => {
                            if (editingCell) {
                                finishEdit({ mode: "commit", refresh: false });
                            }

                            if (!selectedCellEditable || !activeCellEditTarget) {
                                return;
                            }

                            setIsEditingCellValue(true);
                            setCellValueEditTarget(activeCellEditTarget);
                            event.currentTarget.select();
                        }}
                        onChange={(event) => {
                            setCellValueInputValue(event.currentTarget.value);
                        }}
                        onKeyDown={(event) => {
                            if (event.key === "Enter") {
                                event.preventDefault();
                                commitCellValueInput();
                                return;
                            }

                            if (event.key === "Escape") {
                                event.preventDefault();
                                resetCellValueInput();
                            }
                        }}
                    />
                    {showCellValueActions ? (
                        <span className="toolbar__field-actions" aria-label="Cell value actions">
                            <button
                                type="button"
                                className="toolbar__toggle"
                                aria-label={STRINGS.cancelInput}
                                title={STRINGS.cancelInput}
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={resetCellValueInput}
                            >
                                <span className="codicon codicon-close toolbar__toggle-icon" aria-hidden />
                            </button>
                            <button
                                type="button"
                                className="toolbar__toggle is-active"
                                aria-label={STRINGS.confirmInput}
                                title={STRINGS.confirmInput}
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={commitCellValueInput}
                            >
                                <span className="codicon codicon-check toolbar__toggle-icon" aria-hidden />
                            </button>
                        </span>
                    ) : null}
                </label>
            </div>
            <div className="toolbar__group">
                <ToolbarButton
                    actionLabel={STRINGS.search}
                    icon="codicon-search"
                    iconOnly={true}
                    isActive={isSearchPanelOpen}
                    onClick={() => openSearchPanel("find")}
                />
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
                    disabled={!currentModel.canEdit || isSaving}
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
                    isLoading={isSaving}
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
                    notifyEditorToolbarSync();
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

function areSelectionRangesEqual(left: CellRange | null, right: CellRange | null): boolean {
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

function areCellPositionsEqual(
    left: Pick<CellPosition, "rowNumber" | "columnNumber"> | null,
    right: Pick<CellPosition, "rowNumber" | "columnNumber"> | null
): boolean {
    if (left === right) {
        return true;
    }

    if (!left || !right) {
        return left === right;
    }

    return left.rowNumber === right.rowNumber && left.columnNumber === right.columnNumber;
}

function areEditingCellsEqual(left: EditingCell | null, right: EditingCell | null): boolean {
    if (left === right) {
        return true;
    }

    if (!left || !right) {
        return left === right;
    }

    return (
        left.sheetKey === right.sheetKey &&
        left.rowNumber === right.rowNumber &&
        left.columnNumber === right.columnNumber &&
        left.value === right.value
    );
}

function arePendingEditsEqual(
    left: PendingEdit | undefined,
    right: PendingEdit | undefined
): boolean {
    if (left === right) {
        return true;
    }

    if (!left || !right) {
        return left === right;
    }

    return (
        left.sheetKey === right.sheetKey &&
        left.rowNumber === right.rowNumber &&
        left.columnNumber === right.columnNumber &&
        left.value === right.value
    );
}

function isActiveHighlightRow(activeRowNumber: number | null, rowNumber: number): boolean {
    return activeRowNumber === rowNumber;
}

function isActiveHighlightColumn(
    activeColumnNumber: number | null,
    columnNumber: number
): boolean {
    return activeColumnNumber === columnNumber;
}

function isCellWithinSelectionRange(
    selectionRange: CellRange | null,
    rowNumber: number,
    columnNumber: number
): boolean {
    return Boolean(
        selectionRange &&
            rowNumber >= selectionRange.startRow &&
            rowNumber <= selectionRange.endRow &&
            columnNumber >= selectionRange.startColumn &&
            columnNumber <= selectionRange.endColumn
    );
}

interface EditorVirtualGridMetrics extends VirtualViewportState {
    contentWidth: number;
    contentHeight: number;
    frozenRowNumbers: number[];
    frozenColumnNumbers: number[];
    stickyTopHeight: number;
    stickyLeftWidth: number;
}

function createSequentialNumbers(count: number): number[] {
    return Array.from({ length: count }, (_, index) => index + 1);
}

function areNumberArraysEqual(left: number[], right: number[]): boolean {
    if (left.length !== right.length) {
        return false;
    }

    for (let index = 0; index < left.length; index += 1) {
        if (left[index] !== right[index]) {
            return false;
        }
    }

    return true;
}

function hasEditorVirtualGridLayoutChanged(
    previous: EditorVirtualGridMetrics,
    next: EditorVirtualGridMetrics
): boolean {
    return (
        previous.viewportHeight !== next.viewportHeight ||
        previous.viewportWidth !== next.viewportWidth ||
        previous.rowHeaderWidth !== next.rowHeaderWidth ||
        previous.frozenRowCount !== next.frozenRowCount ||
        previous.frozenColumnCount !== next.frozenColumnCount ||
        previous.contentWidth !== next.contentWidth ||
        previous.contentHeight !== next.contentHeight ||
        previous.stickyTopHeight !== next.stickyTopHeight ||
        previous.stickyLeftWidth !== next.stickyLeftWidth ||
        !areNumberArraysEqual(previous.rowNumbers, next.rowNumbers) ||
        !areNumberArraysEqual(previous.columnNumbers, next.columnNumbers) ||
        !areNumberArraysEqual(previous.frozenRowNumbers, next.frozenRowNumbers) ||
        !areNumberArraysEqual(previous.frozenColumnNumbers, next.frozenColumnNumbers)
    );
}

function getEditorGridItemStyle({
    top,
    left,
    width,
    height,
}: {
    top: number;
    left: number;
    width: number;
    height: number;
}): React.CSSProperties {
    return {
        width: `${width}px`,
        height: `${height}px`,
        transform: `translate(${left}px, ${top}px)`,
    };
}

function getEditorGridTop(rowNumber: number): number {
    return EDITOR_VIRTUAL_HEADER_HEIGHT + (rowNumber - 1) * EDITOR_VIRTUAL_ROW_HEIGHT;
}

function getEditorGridLeft(rowHeaderWidth: number, columnNumber: number): number {
    return rowHeaderWidth + (columnNumber - 1) * EDITOR_VIRTUAL_COLUMN_WIDTH;
}

function createEditorVirtualGridMetrics(
    currentModel: EditorRenderModel,
    element: HTMLElement | null,
    fallbackScrollState?: ScrollState | null
): EditorVirtualGridMetrics {
    const viewportHeight = element?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT;
    const viewportWidth = element?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH;
    const scrollTop = element?.scrollTop ?? fallbackScrollState?.top ?? 0;
    const scrollLeft = element?.scrollLeft ?? fallbackScrollState?.left ?? 0;
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        viewportHeight,
        viewportWidth,
    });
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        freezePane: currentModel.activeSheet.freezePane,
    });
    const {
        rowCount: visibleFrozenRowCount,
        columnCount: visibleFrozenColumnCount,
    } = getVisibleFrozenEditorCounts({
        frozenRowCount,
        frozenColumnCount,
        viewportHeight,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const rowWindow = createEditorRowWindow({
        totalRows: displayGrid.rowCount,
        frozenRowCount,
        scrollTop,
        viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        totalColumns: displayGrid.columnCount,
        frozenColumnCount,
        scrollLeft,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const contentSize = getEditorContentSize({
        rowCount: displayGrid.rowCount,
        columnCount: displayGrid.columnCount,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });

    return {
        scrollTop,
        scrollLeft,
        viewportHeight,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
        frozenRowCount,
        frozenColumnCount,
        rowNumbers: rowWindow.rowNumbers,
        columnNumbers: columnWindow.columnNumbers,
        contentWidth: contentSize.width,
        contentHeight: contentSize.height,
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        stickyTopHeight:
            EDITOR_VIRTUAL_HEADER_HEIGHT + frozenRowCount * EDITOR_VIRTUAL_ROW_HEIGHT,
        stickyLeftWidth:
            displayGrid.rowHeaderWidth + frozenColumnCount * EDITOR_VIRTUAL_COLUMN_WIDTH,
    };
}

function useEditorVirtualGrid(
    currentModel: EditorRenderModel,
    initialScrollState: ScrollState | null,
    revision: number
): {
    viewportRef: React.RefObject<HTMLDivElement | null>;
    metrics: EditorVirtualGridMetrics;
    handleScroll(event: React.UIEvent<HTMLDivElement>): void;
} {
    const viewportRef = React.useRef<HTMLDivElement | null>(null);
    const scrollFrameRef = React.useRef(0);
    const latestScrollElementRef = React.useRef<HTMLDivElement | null>(null);
    const [metrics, setMetrics] = React.useState<EditorVirtualGridMetrics>(() =>
        createEditorVirtualGridMetrics(currentModel, null, initialScrollState)
    );
    const syncMetrics = React.useEffectEvent(
        (
            element: HTMLDivElement | null,
            fallbackScrollState?: ScrollState | null,
            { force = false }: { force?: boolean } = {}
        ) => {
            const nextMetrics = createEditorVirtualGridMetrics(
                currentModel,
                element,
                fallbackScrollState
            );
            setMetrics((previous) =>
                force || hasEditorVirtualGridLayoutChanged(previous, nextMetrics)
                    ? nextMetrics
                    : previous
            );
        }
    );

    React.useLayoutEffect(() => {
        const element = viewportRef.current;
        if (!element) {
            syncMetrics(null, initialScrollState, { force: true });
            return;
        }

        const displayGrid = getEditorDisplayGridDimensions({
            rowCount: currentModel.activeSheet.rowCount,
            columnCount: currentModel.activeSheet.columnCount,
            viewportHeight: element.clientHeight,
            viewportWidth: element.clientWidth,
        });
        const contentSize = getEditorContentSize({
            rowCount: displayGrid.rowCount,
            columnCount: displayGrid.columnCount,
            rowHeaderWidth: displayGrid.rowHeaderWidth,
        });
        const nextTop = clampEditorScrollPosition(
            initialScrollState?.top ?? element.scrollTop,
            contentSize.height - element.clientHeight
        );
        const nextLeft = clampEditorScrollPosition(
            initialScrollState?.left ?? element.scrollLeft,
            contentSize.width - element.clientWidth
        );

        if (element.scrollTop !== nextTop) {
            element.scrollTop = nextTop;
        }

        if (element.scrollLeft !== nextLeft) {
            element.scrollLeft = nextLeft;
        }

        syncMetrics(element, undefined, { force: true });

        let resizeObserver: ResizeObserver | null = null;
        const handleResize = () => {
            syncMetrics(element);
        };

        if (typeof ResizeObserver !== "undefined") {
            resizeObserver = new ResizeObserver(() => {
                syncMetrics(element);
            });
            resizeObserver.observe(element);
        } else {
            window.addEventListener("resize", handleResize);
        }

        return () => {
            if (scrollFrameRef.current) {
                cancelAnimationFrame(scrollFrameRef.current);
                scrollFrameRef.current = 0;
            }

            resizeObserver?.disconnect();
            window.removeEventListener("resize", handleResize);
        };
    }, [
        currentModel.activeSheet.key,
        currentModel.activeSheet.rowCount,
        currentModel.activeSheet.columnCount,
        currentModel.activeSheet.freezePane?.rowCount ?? 0,
        currentModel.activeSheet.freezePane?.columnCount ?? 0,
        initialScrollState?.top ?? 0,
        initialScrollState?.left ?? 0,
        revision,
    ]);

    const handleScroll = (event: React.UIEvent<HTMLDivElement>) => {
        const element = event.currentTarget;
        latestScrollElementRef.current = element;
        if (scrollFrameRef.current) {
            return;
        }

        scrollFrameRef.current = requestAnimationFrame(() => {
            scrollFrameRef.current = 0;
            const element = latestScrollElementRef.current;
            if (!element) {
                return;
            }

            syncMetrics(element);
        });
    };

    return {
        viewportRef,
        metrics,
        handleScroll,
    };
}

const EditorCornerHeader = React.memo(function EditorCornerHeader({
    rowHeaderWidth,
}: {
    rowHeaderWidth: number;
}): React.ReactElement {
    return (
        <div
            aria-hidden
            className="editor-grid__item editor-grid__item--corner editor-grid__item--header grid__row-number"
            style={getEditorGridItemStyle({
                top: 0,
                left: 0,
                width: rowHeaderWidth,
                height: EDITOR_VIRTUAL_HEADER_HEIGHT,
            })}
        >
            <span className="grid__row-label">
                <span>#</span>
            </span>
        </div>
    );
});

const EditorColumnHeaderCell = React.memo(function EditorColumnHeaderCell({
    label,
    columnNumber,
    hasPending,
    activeColumnNumber,
    top,
    left,
}: {
    label: string;
    columnNumber: number;
    hasPending: boolean;
    activeColumnNumber: number | null;
    top: number;
    left: number;
}): React.ReactElement {
    const isActiveColumn = isActiveHighlightColumn(activeColumnNumber, columnNumber);

    return (
        <div
            className={classNames([
                "editor-grid__item",
                "editor-grid__item--header",
                "grid__column",
                hasPending && "grid__column--diff",
                hasPending && "grid__column--pending",
                isActiveColumn && "grid__column--active",
            ])}
            data-column-number={columnNumber}
            data-role="grid-column-header"
            style={getEditorGridItemStyle({
                top,
                left,
                width: EDITOR_VIRTUAL_COLUMN_WIDTH,
                height: EDITOR_VIRTUAL_HEADER_HEIGHT,
            })}
            onClick={() => {
                closeContextMenu({ refresh: false });
                selectEntireColumn(columnNumber);
            }}
            onContextMenu={(event) => {
                event.preventDefault();
                selectEntireColumn(columnNumber);
                openColumnContextMenu(columnNumber, event.clientX, event.clientY);
            }}
        >
            <span className="grid__column-label">
                {hasPending ? <PendingMarker /> : null}
                <span>{label}</span>
            </span>
        </div>
    );
},
(previous, next) =>
    previous.label === next.label &&
    previous.columnNumber === next.columnNumber &&
    previous.hasPending === next.hasPending &&
    previous.activeColumnNumber === next.activeColumnNumber &&
    previous.top === next.top &&
    previous.left === next.left
);

const EditorRowHeaderCell = React.memo(function EditorRowHeaderCell({
    rowNumber,
    hasPending,
    activeRowNumber,
    top,
    rowHeaderWidth,
}: {
    rowNumber: number;
    hasPending: boolean;
    activeRowNumber: number | null;
    top: number;
    rowHeaderWidth: number;
}): React.ReactElement {
    const isActiveRow = isActiveHighlightRow(activeRowNumber, rowNumber);

    return (
        <div
            className={classNames([
                "editor-grid__item",
                "editor-grid__item--row-header",
                "grid__row-number",
                hasPending && "grid__row-number--pending",
                isActiveRow && "grid__row-number--active",
            ])}
            data-role="grid-row-header"
            data-row-number={rowNumber}
            style={getEditorGridItemStyle({
                top,
                left: 0,
                width: rowHeaderWidth,
                height: EDITOR_VIRTUAL_ROW_HEIGHT,
            })}
            onClick={() => {
                closeContextMenu({ refresh: false });
                selectEntireRow(rowNumber);
            }}
            onContextMenu={(event) => {
                event.preventDefault();
                selectEntireRow(rowNumber);
                openRowContextMenu(rowNumber, event.clientX, event.clientY);
            }}
        >
            <span className="grid__row-label">
                {hasPending ? <PendingMarker /> : null}
                <span>{rowNumber}</span>
            </span>
        </div>
    );
},
(previous, next) =>
    previous.rowNumber === next.rowNumber &&
    previous.hasPending === next.hasPending &&
    previous.activeRowNumber === next.activeRowNumber &&
    previous.top === next.top &&
    previous.rowHeaderWidth === next.rowHeaderWidth
);

const EditorVirtualCell = React.memo(function EditorVirtualCell({
    currentModel,
    rowNumber,
    columnNumber,
    top,
    left,
    selectionRange,
    activeRowNumber,
    activeColumnNumber,
    currentSelection,
    activeEditingCell,
    pendingEdit,
}: {
    currentModel: EditorRenderModel;
    rowNumber: number;
    columnNumber: number;
    top: number;
    left: number;
    selectionRange: CellRange | null;
    activeRowNumber: number | null;
    activeColumnNumber: number | null;
    currentSelection: CellPosition | null;
    activeEditingCell: EditingCell | null;
    pendingEdit: PendingEdit | undefined;
}): React.ReactElement {
    const cell =
        getCellView(rowNumber, columnNumber, currentModel) ?? {
            key: createCellKey(rowNumber, columnNumber),
            address: getCellAddressLabel(rowNumber, columnNumber),
            value: "",
            formula: null,
            isPresent: false,
            isSelected: false,
        };
    const value = pendingEdit?.value ?? cell.value;
    const formula = pendingEdit ? null : cell.formula;
    const editable = Boolean(currentModel.canEdit && !cell.formula);
    const isPrimarySelection = isSelectionFocusCell(currentSelection, rowNumber, columnNumber);
    const isSelected = isCellWithinSelectionRange(selectionRange, rowNumber, columnNumber);
    const isSearchFocusedSelection =
        searchFeedback?.status === "matched" ||
        searchFeedback?.status === "replaced" ||
        searchFeedback?.status === "no-change";
    const showPrimarySelectionFrame =
        isPrimarySelection &&
        (!hasExpandedSelectionRange(selectionRange) || isSearchFocusedSelection);
    const isActiveRow = isActiveHighlightRow(activeRowNumber, rowNumber);
    const isActiveColumn = isActiveHighlightColumn(activeColumnNumber, columnNumber);
    const isEditing =
        activeEditingCell?.rowNumber === rowNumber &&
        activeEditingCell.columnNumber === columnNumber;

    return (
        <div
            aria-selected={isSelected}
            className={classNames([
                "editor-grid__item",
                "grid__cell",
                isSelected && "grid__cell--selected-range",
                showPrimarySelectionFrame && "grid__cell--selected",
                isActiveRow && "grid__cell--active-row",
                isActiveColumn && "grid__cell--active-column",
                !editable && "grid__cell--locked",
                pendingEdit && "grid__cell--pending",
                isEditing && "grid__cell--editing",
                ...getSelectionOutlineClasses(rowNumber, columnNumber, selectionRange),
            ])}
            data-column-number={columnNumber}
            data-editable={editable}
            data-role="grid-cell"
            data-row-number={rowNumber}
            style={getEditorGridItemStyle({
                top,
                left,
                width: EDITOR_VIRTUAL_COLUMN_WIDTH,
                height: EDITOR_VIRTUAL_ROW_HEIGHT,
            })}
            title={getCellTooltip(cell.address, value, formula)}
            onPointerDown={(event) => {
                if (event.button !== 0) {
                    return;
                }

                closeContextMenu({ refresh: false });
                startSelectionDrag(event.pointerId, { rowNumber, columnNumber });
                setSelectedCellLocal(
                    { rowNumber, columnNumber },
                    {
                        syncHost: false,
                        anchorCell: { rowNumber, columnNumber },
                    }
                );
            }}
            onClick={(event) => {
                if (suppressNextCellClick) {
                    suppressNextCellClick = false;
                    event.preventDefault();
                    return;
                }

                const forceRender = Boolean(editingCell);
                if (forceRender) {
                    finishEdit({ mode: "commit", refresh: false });
                }

                setSelectedCellLocal(
                    { rowNumber, columnNumber },
                    {
                        syncHost: true,
                        anchorCell:
                            event.shiftKey && selectedCell
                                ? (selectionAnchorCell ?? selectedCell)
                                : undefined,
                        forceRender,
                    }
                );
            }}
            onDoubleClick={(event) => {
                if (!currentModel.canEdit || !editable) {
                    return;
                }

                event.preventDefault();
                startEditCell(rowNumber, columnNumber, value);
            }}
        >
            <div className="grid__cell-content">
                {isEditing && activeEditingCell?.sheetKey === currentModel.activeSheet.key ? (
                    <CellEditor edit={activeEditingCell} />
                ) : (
                    <CellValue formula={formula} value={value} />
                )}
            </div>
        </div>
    );
},
(previous, next) =>
    previous.currentModel.activeSheet.key === next.currentModel.activeSheet.key &&
    previous.currentModel.activeSheet.cells === next.currentModel.activeSheet.cells &&
    previous.currentModel.canEdit === next.currentModel.canEdit &&
    previous.rowNumber === next.rowNumber &&
    previous.columnNumber === next.columnNumber &&
    previous.top === next.top &&
    previous.left === next.left &&
    previous.activeRowNumber === next.activeRowNumber &&
    previous.activeColumnNumber === next.activeColumnNumber &&
    areSelectionRangesEqual(previous.selectionRange, next.selectionRange) &&
    areCellPositionsEqual(previous.currentSelection, next.currentSelection) &&
    areEditingCellsEqual(previous.activeEditingCell, next.activeEditingCell) &&
    arePendingEditsEqual(previous.pendingEdit, next.pendingEdit)
);

function EditorVirtualGrid({
    currentModel,
    pendingSummary,
    view,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    view: Extract<ViewState, { kind: "app" }>;
}): React.ReactElement {
    const { viewportRef, metrics, handleScroll } = useEditorVirtualGrid(
        currentModel,
        view.scrollState,
        view.revision
    );
    const selectionRange = getSelectionRange();
    const currentSelection = selectedCell;
    const activeHighlightCell = getActiveHighlightCell();
    const activeRowNumber = activeHighlightCell?.rowNumber ?? null;
    const activeColumnNumber = activeHighlightCell?.columnNumber ?? null;
    const activeEditingCell = editingCell;
    const bodyItems: React.ReactElement[] = [];
    const topItems: React.ReactElement[] = [];
    const leftItems: React.ReactElement[] = [];
    const cornerItems: React.ReactElement[] = [
        <EditorCornerHeader
            key="corner"
            rowHeaderWidth={metrics.rowHeaderWidth}
        />,
    ];
    const createGridCellItem = (
        keyPrefix: "body" | "left" | "top" | "corner",
        rowNumber: number,
        columnNumber: number
    ): React.ReactElement => {
        const top = getEditorGridTop(rowNumber);
        const left = getEditorGridLeft(metrics.rowHeaderWidth, columnNumber);
        const key = `${keyPrefix}:${rowNumber}:${columnNumber}`;

        const pendingEdit = pendingEdits.get(
            getPendingEditKey(currentModel.activeSheet.key, rowNumber, columnNumber)
        );
        return (
            <EditorVirtualCell
                key={key}
                currentModel={currentModel}
                rowNumber={rowNumber}
                columnNumber={columnNumber}
                top={top}
                left={left}
                selectionRange={selectionRange}
                activeRowNumber={activeRowNumber}
                activeColumnNumber={activeColumnNumber}
                currentSelection={currentSelection}
                activeEditingCell={activeEditingCell}
                pendingEdit={pendingEdit}
            />
        );
    };

    for (const columnNumber of metrics.columnNumbers) {
        topItems.push(
            <EditorColumnHeaderCell
                key={`top:header:${columnNumber}`}
                label={
                    currentModel.activeSheet.columns[columnNumber - 1] ??
                    getColumnLabel(columnNumber)
                }
                columnNumber={columnNumber}
                hasPending={pendingSummary.columns.has(columnNumber)}
                activeColumnNumber={activeColumnNumber}
                top={0}
                left={getEditorGridLeft(metrics.rowHeaderWidth, columnNumber)}
            />
        );
    }

    for (const columnNumber of metrics.frozenColumnNumbers) {
        cornerItems.push(
            <EditorColumnHeaderCell
                key={`corner:header:${columnNumber}`}
                label={
                    currentModel.activeSheet.columns[columnNumber - 1] ??
                    getColumnLabel(columnNumber)
                }
                columnNumber={columnNumber}
                hasPending={pendingSummary.columns.has(columnNumber)}
                activeColumnNumber={activeColumnNumber}
                top={0}
                left={getEditorGridLeft(metrics.rowHeaderWidth, columnNumber)}
            />
        );
    }

    for (const rowNumber of metrics.rowNumbers) {
        leftItems.push(
            <EditorRowHeaderCell
                key={`left:row:${rowNumber}`}
                rowNumber={rowNumber}
                hasPending={pendingSummary.rows.has(rowNumber)}
                activeRowNumber={activeRowNumber}
                top={getEditorGridTop(rowNumber)}
                rowHeaderWidth={metrics.rowHeaderWidth}
            />
        );

        for (const columnNumber of metrics.columnNumbers) {
            bodyItems.push(createGridCellItem("body", rowNumber, columnNumber));
        }

        for (const columnNumber of metrics.frozenColumnNumbers) {
            leftItems.push(createGridCellItem("left", rowNumber, columnNumber));
        }
    }

    for (const rowNumber of metrics.frozenRowNumbers) {
        cornerItems.push(
            <EditorRowHeaderCell
                key={`corner:row:${rowNumber}`}
                rowNumber={rowNumber}
                hasPending={pendingSummary.rows.has(rowNumber)}
                activeRowNumber={activeRowNumber}
                top={getEditorGridTop(rowNumber)}
                rowHeaderWidth={metrics.rowHeaderWidth}
            />
        );

        for (const columnNumber of metrics.columnNumbers) {
            topItems.push(createGridCellItem("top", rowNumber, columnNumber));
        }

        for (const columnNumber of metrics.frozenColumnNumbers) {
            cornerItems.push(createGridCellItem("corner", rowNumber, columnNumber));
        }
    }

    React.useLayoutEffect(() => {
        if (!view.revealSelection) {
            return;
        }

        revealSelectedCell();
    }, [view.revision, view.revealSelection]);

    return (
        <div className="pane__table editor-grid-shell">
            <div
                ref={viewportRef}
                className="editor-grid__viewport"
                data-role="grid-scroll-main"
                onScroll={handleScroll}
            >
                <div
                    className="editor-grid__canvas"
                    style={{
                        width: `${metrics.contentWidth}px`,
                        height: `${metrics.contentHeight}px`,
                    }}
                >
                    <div className="editor-grid__layer editor-grid__layer--body">{bodyItems}</div>
                    <div
                        className="editor-grid__overlay editor-grid__overlay--top"
                        style={{
                            width: `${metrics.contentWidth}px`,
                            height: `${metrics.stickyTopHeight}px`,
                        }}
                        onWheel={forwardVirtualGridWheel}
                    >
                        <div
                            className="editor-grid__track editor-grid__track--x"
                            style={{
                                width: `${metrics.contentWidth}px`,
                                height: `${metrics.stickyTopHeight}px`,
                            }}
                        >
                            {topItems}
                        </div>
                    </div>
                    <div
                        className="editor-grid__overlay editor-grid__overlay--left"
                        style={{
                            width: `${metrics.stickyLeftWidth}px`,
                            height: `${metrics.contentHeight}px`,
                        }}
                        onWheel={forwardVirtualGridWheel}
                    >
                        <div
                            className="editor-grid__track editor-grid__track--y"
                            style={{
                                width: `${metrics.stickyLeftWidth}px`,
                                height: `${metrics.contentHeight}px`,
                            }}
                        >
                            {leftItems}
                        </div>
                    </div>
                    <div
                        className="editor-grid__overlay editor-grid__overlay--corner"
                        style={{
                            width: `${metrics.stickyLeftWidth}px`,
                            height: `${metrics.stickyTopHeight}px`,
                        }}
                        onWheel={forwardVirtualGridWheel}
                    >
                        <div
                            className="editor-grid__track"
                            style={{
                                width: `${metrics.stickyLeftWidth}px`,
                                height: `${metrics.stickyTopHeight}px`,
                            }}
                        >
                            {cornerItems}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}

function EditorPane({
    currentModel,
    pendingSummary,
    view,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    view: Extract<ViewState, { kind: "app" }>;
}): React.ReactElement {
    const hasVisibleCells =
        currentModel.activeSheet.rowCount > 0 && currentModel.activeSheet.columnCount > 0;
    const gridInstanceKey = [
        currentModel.activeSheet.key,
        currentModel.activeSheet.rowCount,
        currentModel.activeSheet.columnCount,
        currentModel.activeSheet.freezePane?.rowCount ?? 0,
        currentModel.activeSheet.freezePane?.columnCount ?? 0,
    ].join(":");

    return (
        <section className="pane pane--single pane--editor">
            {!hasVisibleCells ? (
                <div className="pane__table">
                    <div className="empty-table">{STRINGS.noRowsAvailable}</div>
                </div>
            ) : (
                <EditorVirtualGrid
                    key={gridInstanceKey}
                    currentModel={currentModel}
                    pendingSummary={pendingSummary}
                    view={view}
                />
            )}
        </section>
    );
}

function areSheetTabWidthsEqual(
    left: Record<string, number>,
    right: Record<string, number>
): boolean {
    const leftKeys = Object.keys(left);
    const rightKeys = Object.keys(right);
    if (leftKeys.length !== rightKeys.length) {
        return false;
    }

    return leftKeys.every((key) => left[key] === right[key]);
}

function getMaxVisibleSheetTabs(
    tabs: readonly EditorSheetTabView[],
    containerWidth: number,
    measuredTabWidths: Record<string, number>
): number {
    return getMaxVisibleSheetTabsForWidth(tabs, {
        containerWidth,
        getTabWidth: (tab) => measuredTabWidths[tab.key] ?? SHEET_TAB_ESTIMATED_WIDTH,
        itemGap: SHEET_TAB_ITEM_GAP,
        overflowTriggerWidth: SHEET_TAB_OVERFLOW_TRIGGER_WIDTH,
    });
}

function useObservedElementWidth<TElement extends HTMLElement>(
    ref: React.RefObject<TElement | null>
): number {
    const [width, setWidth] = React.useState(0);

    React.useLayoutEffect(() => {
        const element = ref.current;
        if (!element) {
            return;
        }

        let frameId = 0;
        const updateWidth = (): void => {
            cancelAnimationFrame(frameId);
            frameId = requestAnimationFrame(() => {
                setWidth(Math.round(element.getBoundingClientRect().width));
            });
        };

        updateWidth();

        const observer = new ResizeObserver(() => {
            updateWidth();
        });
        observer.observe(element);
        window.addEventListener("resize", updateWidth);
        window.visualViewport?.addEventListener("resize", updateWidth);

        return () => {
            cancelAnimationFrame(frameId);
            observer.disconnect();
            window.removeEventListener("resize", updateWidth);
            window.visualViewport?.removeEventListener("resize", updateWidth);
        };
    }, [ref]);

    return width;
}

function Tabs({
    currentModel,
    pendingSummary,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    const viewportRef = React.useRef<HTMLDivElement | null>(null);
    const overflowRef = React.useRef<HTMLDivElement | null>(null);
    const measureRef = React.useRef<HTMLDivElement | null>(null);
    const [isOverflowOpen, setIsOverflowOpen] = React.useState(false);
    const [measuredTabWidths, setMeasuredTabWidths] = React.useState<Record<string, number>>({});
    const viewportWidth = useObservedElementWidth(viewportRef);
    const pendingSheetKeySignature = Array.from(pendingSummary.sheetKeys).sort().join("\0");
    const maxVisibleTabs = getMaxVisibleSheetTabs(currentModel.sheets, viewportWidth, measuredTabWidths);
    const tabLayout = partitionSheetTabs(currentModel.sheets, maxVisibleTabs);

    React.useLayoutEffect(() => {
        const measureRoot = measureRef.current;
        if (!measureRoot) {
            return;
        }

        const nextMeasuredTabWidths: Record<string, number> = {};
        for (const element of measureRoot.querySelectorAll<HTMLElement>('[data-role="sheet-tab-measure"]')) {
            const sheetKey = element.dataset.sheetKey;
            if (!sheetKey) {
                continue;
            }

            nextMeasuredTabWidths[sheetKey] = Math.min(
                SHEET_TAB_VISIBLE_MAX_WIDTH,
                Math.ceil(element.getBoundingClientRect().width)
            );
        }

        setMeasuredTabWidths((currentWidths) =>
            areSheetTabWidthsEqual(currentWidths, nextMeasuredTabWidths)
                ? currentWidths
                : nextMeasuredTabWidths
        );
    }, [
        currentModel.sheets,
        currentModel.activeSheet.key,
        pendingSheetKeySignature,
    ]);

    React.useEffect(() => {
        setIsOverflowOpen(false);
    }, [currentModel.activeSheet.key, currentModel.sheets.length, tabLayout.hasOverflow]);

    React.useEffect(() => {
        if (!isOverflowOpen) {
            return;
        }

        const handlePointerDown = (event: PointerEvent): void => {
            const target = event.target;
            if (!(target instanceof HTMLElement)) {
                setIsOverflowOpen(false);
                return;
            }

            if (overflowRef.current?.contains(target)) {
                return;
            }

            setIsOverflowOpen(false);
        };

        const handleKeyDown = (event: KeyboardEvent): void => {
            if (event.key === "Escape") {
                setIsOverflowOpen(false);
            }
        };

        document.addEventListener("pointerdown", handlePointerDown);
        document.addEventListener("keydown", handleKeyDown);

        return () => {
            document.removeEventListener("pointerdown", handlePointerDown);
            document.removeEventListener("keydown", handleKeyDown);
        };
    }, [isOverflowOpen]);

    const setSheet = (sheetKey: string): void => {
        closeContextMenu({ refresh: false });
        setIsOverflowOpen(false);
        vscode.postMessage({ type: "setSheet", sheetKey });
    };

    return (
        <div
            className="tabs"
            onContextMenu={(event) => {
                const target = event.target;
                if (
                    target instanceof HTMLElement &&
                    target.closest('[data-role="sheet-tab"], [data-role="sheet-tab-overflow"]')
                ) {
                    return;
                }

                event.preventDefault();
                openTabContextMenu(currentModel.activeSheet.key, event.clientX, event.clientY);
            }}
        >
            <div ref={viewportRef} className="tabs__viewport">
                <div className="tabs__content">
                    <div className="tabs__list">
                        {tabLayout.visibleTabs.map((sheet: EditorSheetTabView) => {
                            const hasPending = pendingSummary.sheetKeys.has(sheet.key);

                            return (
                                <button
                                    key={sheet.key}
                                    className={classNames(["tab", sheet.isActive && "is-active"])}
                                    data-role="sheet-tab"
                                    title={sheet.label}
                                    type="button"
                                    onClick={() => setSheet(sheet.key)}
                                    onContextMenu={(event) => {
                                        event.preventDefault();
                                        setIsOverflowOpen(false);
                                        openTabContextMenu(sheet.key, event.clientX, event.clientY);
                                    }}
                                >
                                    {hasPending ? <PendingMarker extraClass="tab__marker" /> : null}
                                    <span className="tab__label">{sheet.label}</span>
                                </button>
                            );
                        })}
                    </div>
                    {tabLayout.hasOverflow ? (
                        <div
                            ref={overflowRef}
                            className="tabs__overflow"
                            data-role="sheet-tab-overflow"
                            onContextMenu={(event) => {
                                event.preventDefault();
                            }}
                        >
                            <button
                                aria-label={STRINGS.moreSheets}
                                aria-expanded={isOverflowOpen}
                                aria-haspopup="menu"
                                className={classNames([
                                    "tab",
                                    "tab--overflowTrigger",
                                    isOverflowOpen && "is-active",
                                ])}
                                title={STRINGS.moreSheets}
                                type="button"
                                onClick={() => {
                                    closeContextMenu({ refresh: false });
                                    setIsOverflowOpen((open) => !open);
                                }}
                            >
                                <span className="codicon codicon-more tab__icon" aria-hidden />
                                <span className="tabs__overflowCount" aria-hidden>
                                    {tabLayout.overflowTabs.length}
                                </span>
                            </button>
                            {isOverflowOpen ? (
                                <div
                                    className="tabs__overflowMenu"
                                    data-role="sheet-tab-overflow"
                                    role="menu"
                                >
                                    {tabLayout.overflowTabs.map((sheet: EditorSheetTabView) => {
                                        const hasPending = pendingSummary.sheetKeys.has(sheet.key);

                                        return (
                                            <button
                                                key={sheet.key}
                                                className="context-menu__item tabs__overflowItem"
                                                role="menuitem"
                                                title={sheet.label}
                                                type="button"
                                                onClick={() => setSheet(sheet.key)}
                                                onContextMenu={(event) => {
                                                    event.preventDefault();
                                                    setIsOverflowOpen(false);
                                                    openTabContextMenu(
                                                        sheet.key,
                                                        event.clientX,
                                                        event.clientY
                                                    );
                                                }}
                                            >
                                                {hasPending ? (
                                                    <PendingMarker extraClass="tab__marker" />
                                                ) : null}
                                                <span className="tabs__overflowLabel">
                                                    {sheet.label}
                                                </span>
                                            </button>
                                        );
                                    })}
                                </div>
                            ) : null}
                        </div>
                    ) : null}
                </div>
            </div>
            <div ref={measureRef} aria-hidden className="tabs__measure">
                {currentModel.sheets.map((sheet: EditorSheetTabView) => {
                    const hasPending = pendingSummary.sheetKeys.has(sheet.key);

                    return (
                        <button
                            key={sheet.key}
                            className={classNames([
                                "tab",
                                "tabs__measureTab",
                                sheet.isActive && "is-active",
                            ])}
                            data-role="sheet-tab-measure"
                            data-sheet-key={sheet.key}
                            tabIndex={-1}
                            type="button"
                        >
                            {hasPending ? <PendingMarker extraClass="tab__marker" /> : null}
                            <span className="tab__label">{sheet.label}</span>
                        </button>
                    );
                })}
            </div>
        </div>
    );
}

function TabContextMenu({
    currentModel,
}: {
    currentModel: EditorRenderModel;
}): React.ReactElement | null {
    if (!contextMenu || !currentModel.canEdit) {
        return null;
    }

    const menu = contextMenu;
    const menuStyle: React.CSSProperties = {
        left: Math.max(8, Math.min(menu.x, window.innerWidth - 188)),
        top: Math.max(8, Math.min(menu.y, window.innerHeight - 132)),
    };

    if (menu.kind === "tab") {
        const disableDelete = currentModel.sheets.length <= 1;
        return (
            <div className="context-menu" data-role="context-menu" style={menuStyle}>
                <button className="context-menu__item" type="button" onClick={requestAddSheet}>
                    <span className="codicon codicon-add context-menu__icon" aria-hidden />
                    <span>{STRINGS.addSheet}</span>
                </button>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => requestRenameSheet(menu.sheetKey)}
                >
                    <span className="codicon codicon-edit context-menu__icon" aria-hidden />
                    <span>{STRINGS.renameSheet}</span>
                </button>
                <button
                    className="context-menu__item context-menu__item--danger"
                    disabled={disableDelete}
                    type="button"
                    onClick={() => requestDeleteSheet(menu.sheetKey)}
                >
                    <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                    <span>{STRINGS.deleteSheet}</span>
                </button>
            </div>
        );
    }

    if (menu.kind === "row") {
        return (
            <div className="context-menu" data-role="context-menu" style={menuStyle}>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => requestInsertRow(menu.rowNumber)}
                >
                    <span className="codicon codicon-add context-menu__icon" aria-hidden />
                    <span>{STRINGS.insertRowAbove}</span>
                </button>
                <button
                    className="context-menu__item context-menu__item--danger"
                    disabled={currentModel.activeSheet.rowCount <= 1}
                    type="button"
                    onClick={() => requestDeleteRow(menu.rowNumber)}
                >
                    <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                    <span>{STRINGS.deleteRow}</span>
                </button>
            </div>
        );
    }

    return (
        <div className="context-menu" data-role="context-menu" style={menuStyle}>
            <button
                className="context-menu__item"
                type="button"
                onClick={() => requestInsertColumn(menu.columnNumber)}
            >
                <span className="codicon codicon-add context-menu__icon" aria-hidden />
                <span>{STRINGS.insertColumnLeft}</span>
            </button>
            <button
                className="context-menu__item context-menu__item--danger"
                disabled={currentModel.activeSheet.columnCount <= 1}
                type="button"
                onClick={() => requestDeleteColumn(menu.columnNumber)}
            >
                <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                <span>{STRINGS.deleteColumn}</span>
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

    return (
        <div className="app app--editor" data-role="editor-app">
            <EditorToolbar currentModel={view.model} />
            <SearchPanel currentModel={view.model} />
            <section className="panes panes--single">
                <EditorPane
                    currentModel={view.model}
                    pendingSummary={pendingSummary}
                    view={view}
                />
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
    if (!model) {
        return;
    }

    renderApp({ commitEditing: false });
});

window.addEventListener("message", (event: MessageEvent<IncomingMessage>) => {
    const message = event.data;

    if (message.type === "loading") {
        stopSearchPanelDrag();
        isSearchPanelOpen = false;
        searchMode = "find";
        searchQuery = "";
        replaceValue = "";
        searchOptions = { ...DEFAULT_SEARCH_OPTIONS };
        searchScope = "sheet";
        searchSelectionRange = null;
        searchFeedback = null;
        renderLoading(message.message);
        return;
    }

    if (message.type === "error") {
        stopSearchPanelDrag();
        isSaving = false;
        renderError(message.message);
        return;
    }

    if (message.type === "searchResult") {
        handleSearchResult(message);
        return;
    }

    if (message.type === "render") {
        model = stabilizeIncomingRenderModel(model, message.payload, {
            canReuseActiveSheetData: Boolean(message.reuseActiveSheetData),
        });
        isSaving = false;

        if (message.clearPendingEdits) {
            const pendingEditsBeforeClear = Array.from(pendingEdits.values());
            pendingEdits.clear();
            if (message.preservePendingHistory) {
                const rebasedHistory = rebasePendingHistory(
                    undoStack,
                    redoStack,
                    pendingEditsBeforeClear
                );
                undoStack.length = 0;
                undoStack.push(...rebasedHistory.undoStack);
                redoStack.length = 0;
                redoStack.push(...rebasedHistory.redoStack);
            } else {
                undoStack.length = 0;
                redoStack.length = 0;
            }
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

        renderApp({
            revealSelection: !message.silent,
            useModelSelection: message.useModelSelection,
            sync: true,
        });
    }
});

document.addEventListener("keydown", (event: KeyboardEvent) => {
    const isTextInputContext = isTextInputTarget(event.target);

    if (event.key === "Escape" && contextMenu) {
        event.preventDefault();
        closeContextMenu();
        return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        triggerSave();
        return;
    }

    if (
        !editingCell &&
        (event.ctrlKey || event.metaKey) &&
        !event.altKey &&
        event.key.toLowerCase() === "f"
    ) {
        event.preventDefault();
        openSearchPanel("find");
        return;
    }

    if (
        !editingCell &&
        (event.ctrlKey || event.metaKey) &&
        !event.altKey &&
        event.key.toLowerCase() === "h"
    ) {
        event.preventDefault();
        openSearchPanel("replace");
        return;
    }

    if (event.key === "Escape" && isSearchPanelOpen && !isTextInputContext) {
        event.preventDefault();
        closeSearchPanel();
        return;
    }

    if (isTextInputContext) {
        return;
    }

    if (!editingCell && (event.ctrlKey || event.metaKey) && !event.altKey) {
        if (event.key.toLowerCase() === "g") {
            event.preventDefault();
            focusToolbarInput("position");
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
    if (!contextMenu) {
        return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
        closeContextMenu();
        return;
    }

    if (target.closest('[data-role="context-menu"]')) {
        return;
    }

    closeContextMenu();
});

document.addEventListener("pointermove", (event: PointerEvent) => {
    if (searchPanelDragState && searchPanelDragState.pointerId === event.pointerId) {
        event.preventDefault();
        updateSearchPanelDrag(event.clientX, event.clientY);
        return;
    }

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
    stopSearchPanelDrag(event.pointerId);
    stopSelectionDrag(event.pointerId);
});

document.addEventListener("pointercancel", (event: PointerEvent) => {
    stopSearchPanelDrag(event.pointerId);
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
