import * as React from "react";
import { AiOutlineSortAscending, AiOutlineSortDescending } from "react-icons/ai";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import type { CellSnapshot, EditorRenderModel } from "../../core/model/types";
import {
    mergeCellAlignments,
    type CellAlignmentSnapshot,
    type EditorAlignmentPatch,
} from "../../core/model/alignment";
import { createCellKey, getColumnLabel } from "../../core/model/cells";
import { formatI18nMessage, RUNTIME_MESSAGES } from "../../i18n/catalog";
import {
    EDITOR_EXTRA_PADDING_ROWS,
    EDITOR_VIRTUAL_COLUMN_WIDTH,
    EDITOR_VIRTUAL_HEADER_HEIGHT,
    clampEditorScrollPosition,
    createEditorColumnWindow,
    createEditorPixelColumnLayout,
    createEditorPixelRowLayout,
    createEditorRowWindow,
    getEditorColumnLeft,
    getEditorColumnWidth,
    getEditorDisplayGridDimensions,
    getEditorDisplayColumnLayout,
    getEditorContentSize,
    getEditorFrozenColumnsWidth,
    getEditorFrozenRowsHeight,
    getEditorRowHeight,
    getEditorRowTop,
    getFrozenEditorCounts,
    getVisibleFrozenEditorCounts,
} from "./editor-virtual-grid";
import {
    DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
    MAX_COLUMN_PIXEL_WIDTH,
    MIN_COLUMN_PIXEL_WIDTH,
    convertPixelsToWorkbookColumnWidth,
    stabilizeColumnPixelWidth,
    getFontShorthand,
    measureMaximumDigitWidth,
    type PixelColumnLayout,
} from "../column-layout";
import { type PixelRowLayout } from "../row-layout";
import {
    clearPendingSelectionAfterRender as clearControllerPendingSelectionAfterRender,
    clearSelectedCell as clearControllerSelectedCell,
    getActiveHighlightCell as getControllerActiveHighlightCell,
    getExpandedSelectionRange as getControllerExpandedSelectionRange,
    getSelectionExtendAnchorCell as getControllerSelectionExtendAnchorCell,
    getSelectionRange as getControllerSelectionRange,
    isActiveSelectionCell as isControllerActiveSelectionCell,
    setSelectedCell as setControllerSelectedCell,
    setSelectionAnchorCell as setControllerSelectionAnchorCell,
    setSuppressAutoSelection as setControllerSuppressAutoSelection,
    startCellSelectionDrag as startControllerCellSelectionDrag,
    startColumnSelectionDrag as startControllerColumnSelectionDrag,
    startRowSelectionDrag as startControllerRowSelectionDrag,
    stopSelectionDrag as stopControllerSelectionDrag,
    syncSelectionAnchorToSelectedCell as syncControllerSelectionAnchorToSelectedCell,
    type PendingSelection,
    type SelectionControllerState,
    type SelectionDragState,
} from "./editor-selection-controller";
import { clipSelectionRangeToVisibleGrid } from "./editor-selection-overlay";
import {
    createColumnSelectionSpanRange,
    createRowSelectionSpanRange,
    createSelectionRange,
    hasExpandedSelectionRange,
    type SelectionRange as CellRange,
} from "./editor-selection-range";
import {
    buildFillChanges,
    createAutoFillDownPreviewRange,
    createFillPreviewRange,
    isCellWithinFillPreviewArea,
} from "./editor-fill-drag";
import {
    CellFormulaBadge,
    CellValue,
    EditorToolbar,
    PendingMarker,
    SearchPanel,
    Shell,
    TabContextMenu,
    Tabs,
} from "./editor-panel-chrome";
import {
    classNames,
    type ContextMenuState,
    type PendingSummary,
    type SearchPanelFeedback,
    type SearchPanelMode,
    type SearchPanelPosition,
} from "./editor-panel-ui-shared";
import {
    getCellContentAlignmentStyle,
    getToolbarHorizontalAlignment,
    getToolbarVerticalAlignment,
} from "./editor-cell-alignment";
import { getCellOverflowMetrics } from "./editor-cell-overflow";
import { notifyEditorToolbarSync } from "./editor-toolbar-sync";
import { type ToolbarCellEditTarget } from "./editor-toolbar-input";
import type {
    EditorPanelStrings,
    EditorSearchResultMessage,
    EditorSearchScope,
    EditorWebviewMessage,
    EditorAlignmentTargetKind,
    SearchOptions,
} from "./editor-panel-types";
import { resolveEditorReplaceResultInSheet } from "./editor-panel-logic";
import {
    canCreateEditorFilterRange,
    clearEditorFilterColumn,
    createEditorSheetFilterSnapshot,
    createEditorSheetFilterState,
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
    type EditorSheetFilterState,
} from "./editor-panel-filter";
import {
    rebasePendingHistory,
    type PendingHistoryChange as PendingEditChange,
    type PendingHistoryEntry as HistoryEntry,
} from "./editor-pending-history";
import { stabilizeIncomingRenderModel } from "./editor-render-stabilizer";
import { getFreezePaneCountsForCell, hasLockedView } from "../view-lock";
import { ImSortAlphaAsc, ImSortAlphaDesc } from "react-icons/im";

interface VsCodeApi {
    postMessage(message: EditorWebviewMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

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

interface AlignmentSelectionTarget {
    target: EditorAlignmentTargetKind;
    selection: CellRange;
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

interface ScrollState {
    top: number;
    left: number;
}

interface FillDragState {
    pointerId: number;
    sourceRange: CellRange;
    previewRange: CellRange | null;
}

interface FillHandleClickState {
    rowNumber: number;
    columnNumber: number;
    timeStamp: number;
}

type GridLayerKind = "body" | "top" | "left" | "corner";

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

interface EditorDisplayRowState {
    actualRowNumbers: number[];
    actualToDisplayRowNumbers: Record<string, number>;
    hiddenActualRowNumbers: number[];
}

interface ColumnResizeState {
    pointerId: number;
    sheetKey: string;
    columnNumber: number;
    startClientX: number;
    startPixelWidth: number;
    previewPixelWidth: number;
    isDragging: boolean;
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

const DEFAULT_STRINGS: EditorPanelStrings = RUNTIME_MESSAGES.en.editorPanel;
type Strings = EditorPanelStrings;

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
let fillDragState: FillDragState | null = null;
let lastFillHandleClickState: FillHandleClickState | null = null;
let searchPanelDragState: SearchPanelDragState | null = null;
let columnResizeState: ColumnResizeState | null = null;
let suppressNextCellClick = false;
let suppressNextRowHeaderClick = false;
let suppressNextColumnHeaderClick = false;

const WHEEL_DELTA_LINE_MODE = 1;
const WHEEL_DELTA_PAGE_MODE = 2;
const WHEEL_LINE_SCROLL_PIXELS = 40;
const DEFAULT_EDITOR_VIEWPORT_HEIGHT = 480;
const DEFAULT_EDITOR_VIEWPORT_WIDTH = 960;
const GRID_CELL_VERTICAL_PADDING_PX = 6;
const GRID_CELL_FONT_SIZE_PX = 13;
const GRID_CELL_LINE_HEIGHT_MULTIPLIER = 1.25;
const GRID_CELL_LINE_HEIGHT_PX = GRID_CELL_FONT_SIZE_PX * GRID_CELL_LINE_HEIGHT_MULTIPLIER;
const FILL_HANDLE_DOUBLE_CLICK_MS = 400;
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
let editorMaximumDigitWidth = DEFAULT_MAXIMUM_DIGIT_WIDTH_PX;
const sheetFilterStates = new Map<string, EditorSheetFilterState | null>();
let filterMenuState: FilterMenuState | null = null;
const IS_DEBUG_MODE = Boolean(
    (globalThis as Record<string, unknown>).__XLSX_EDITOR_DEBUG__ === true
);
let latestVirtualGridCache: {
    model: EditorRenderModel;
    metrics: EditorVirtualGridMetrics;
} | null = null;

interface DebugRenderStats {
    renderedRowCount: number;
    renderedColumnCount: number;
}

type EditorSelectionState = SelectionControllerState<CellPosition>;

let debugRenderStats: DebugRenderStats | null = null;
const debugRenderStatsListeners = new Set<() => void>();

function getSelectionControllerState(): EditorSelectionState {
    return {
        selectedCell,
        selectionAnchorCell,
        selectionRangeOverride,
        pendingSelectionAfterRender,
        suppressAutoSelection,
        selectionDragState,
    };
}

function applySelectionControllerState(state: EditorSelectionState): EditorSelectionState {
    selectedCell = state.selectedCell;
    selectionAnchorCell = state.selectionAnchorCell;
    selectionRangeOverride = state.selectionRangeOverride;
    pendingSelectionAfterRender = state.pendingSelectionAfterRender;
    suppressAutoSelection = state.suppressAutoSelection;
    selectionDragState = state.selectionDragState;
    return state;
}

function updateSelectionControllerState(
    updater: (state: EditorSelectionState) => EditorSelectionState
): EditorSelectionState {
    return applySelectionControllerState(updater(getSelectionControllerState()));
}

function getDebugRenderStatsSnapshot(): DebugRenderStats | null {
    return debugRenderStats;
}

function subscribeDebugRenderStats(listener: () => void): () => void {
    debugRenderStatsListeners.add(listener);
    return () => {
        debugRenderStatsListeners.delete(listener);
    };
}

function areDebugRenderStatsEqual(
    left: DebugRenderStats | null,
    right: DebugRenderStats | null
): boolean {
    return (
        left?.renderedRowCount === right?.renderedRowCount &&
        left?.renderedColumnCount === right?.renderedColumnCount
    );
}

function setDebugRenderStats(nextStats: DebugRenderStats | null): void {
    if (areDebugRenderStatsEqual(debugRenderStats, nextStats)) {
        return;
    }

    debugRenderStats = nextStats;
    for (const listener of debugRenderStatsListeners) {
        listener();
    }
}

function normalizeWorkbookColumnWidth(columnWidth: number): number {
    return Math.round(columnWidth * 256) / 256;
}

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

function useMeasuredMaximumDigitWidth(target: Element | null): number {
    const [maximumDigitWidth, setMaximumDigitWidth] = React.useState(
        DEFAULT_MAXIMUM_DIGIT_WIDTH_PX
    );

    React.useLayoutEffect(() => {
        const updateMaximumDigitWidth = (): void => {
            setMaximumDigitWidth(measureMaximumDigitWidth(getFontShorthand(target)));
        };

        updateMaximumDigitWidth();
        window.addEventListener("resize", updateMaximumDigitWidth);
        window.visualViewport?.addEventListener("resize", updateMaximumDigitWidth);

        return () => {
            window.removeEventListener("resize", updateMaximumDigitWidth);
            window.visualViewport?.removeEventListener("resize", updateMaximumDigitWidth);
        };
    }, [target]);

    return maximumDigitWidth;
}

function getEffectiveSheetColumnWidths(
    currentModel: EditorRenderModel,
    maximumDigitWidth: number
): readonly (number | null)[] | undefined {
    const activeResize = columnResizeState;
    if (!activeResize || activeResize.sheetKey !== currentModel.activeSheet.key) {
        return currentModel.activeSheet.columnWidths;
    }

    const nextColumnWidths = [...(currentModel.activeSheet.columnWidths ?? [])];
    nextColumnWidths[activeResize.columnNumber - 1] = normalizeWorkbookColumnWidth(
        convertPixelsToWorkbookColumnWidth(activeResize.previewPixelWidth, maximumDigitWidth)
    );
    return nextColumnWidths;
}

function getSheetColumnLayout(currentModel: EditorRenderModel): PixelColumnLayout {
    return createEditorPixelColumnLayout({
        columnCount: currentModel.activeSheet.columnCount,
        columnWidths: getEffectiveSheetColumnWidths(currentModel, editorMaximumDigitWidth),
        maximumDigitWidth: editorMaximumDigitWidth,
    });
}

function getEffectiveCellValueForModel(
    currentModel: EditorRenderModel,
    rowNumber: number,
    columnNumber: number
): string {
    return (
        pendingEdits.get(getPendingEditKey(currentModel.activeSheet.key, rowNumber, columnNumber))
            ?.value ??
        getCellView(rowNumber, columnNumber, currentModel)?.value ??
        ""
    );
}

function createEditorFilterCellSource(currentModel: EditorRenderModel): EditorFilterCellSource {
    return {
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        getCellValue: (rowNumber, columnNumber) =>
            getEffectiveCellValueForModel(currentModel, rowNumber, columnNumber),
    };
}

function getStoredSheetFilterState(sheetKey: string): EditorSheetFilterState | null | undefined {
    return sheetFilterStates.has(sheetKey) ? (sheetFilterStates.get(sheetKey) ?? null) : undefined;
}

function closeFilterMenu({ refresh = true }: { refresh?: boolean } = {}): void {
    if (!filterMenuState) {
        return;
    }

    filterMenuState = null;
    if (refresh) {
        renderApp({ commitEditing: false });
    }
}

function getActiveSheetFilterState(
    currentModel: EditorRenderModel | null = model
): EditorSheetFilterState | null {
    if (!currentModel) {
        return null;
    }

    const sheetKey = currentModel.activeSheet.key;
    const filterState =
        getStoredSheetFilterState(sheetKey) ??
        createEditorSheetFilterStateFromSnapshot(currentModel.activeSheet.autoFilter);
    if (!filterState) {
        return null;
    }

    const normalizedRange = normalizeEditorFilterRange(
        createEditorFilterCellSource(currentModel),
        filterState.range
    );
    if (!normalizedRange) {
        if (sheetFilterStates.has(sheetKey)) {
            sheetFilterStates.set(sheetKey, null);
        }
        if (filterMenuState?.sheetKey === sheetKey) {
            filterMenuState = null;
        }
        return null;
    }

    if (!areSelectionRangesEqual(normalizedRange, filterState.range)) {
        const nextFilterState: EditorSheetFilterState = {
            ...filterState,
            range: normalizedRange,
        };
        sheetFilterStates.set(sheetKey, nextFilterState);
        return nextFilterState;
    }

    return filterState;
}

function setSheetFilterState(sheetKey: string, filterState: EditorSheetFilterState | null): void {
    sheetFilterStates.set(sheetKey, filterState);
    if (!filterState && filterMenuState?.sheetKey === sheetKey) {
        filterMenuState = null;
    }
}

function syncActiveSheetFilterStateToHost(
    currentModel: EditorRenderModel,
    filterState: EditorSheetFilterState | null
): void {
    vscode.postMessage({
        type: "setFilterState",
        sheetKey: currentModel.activeSheet.key,
        filterState: createEditorSheetFilterSnapshot(filterState),
    });
}

function getVisibleActualRowsForModel(currentModel: EditorRenderModel): {
    visibleRows: number[];
    hiddenRows: number[];
} {
    return getEditorVisibleRows(
        createEditorFilterCellSource(currentModel),
        getActiveSheetFilterState(currentModel)
    );
}

function createEditorDisplayRowState(
    currentModel: EditorRenderModel,
    visibleRows: readonly number[],
    hiddenRows: readonly number[],
    totalDisplayRowCount: number
): EditorDisplayRowState {
    const actualRowNumbers = [...visibleRows];
    let nextSyntheticRowNumber = currentModel.activeSheet.rowCount + 1;
    while (actualRowNumbers.length < totalDisplayRowCount) {
        actualRowNumbers.push(nextSyntheticRowNumber);
        nextSyntheticRowNumber += 1;
    }

    return {
        actualRowNumbers,
        actualToDisplayRowNumbers: Object.fromEntries(
            actualRowNumbers.map((actualRowNumber, index) => [String(actualRowNumber), index + 1])
        ),
        hiddenActualRowNumbers: [...hiddenRows],
    };
}

function getDisplayRowNumber(
    rowState: EditorDisplayRowState,
    actualRowNumber: number
): number | null {
    return rowState.actualToDisplayRowNumbers[String(actualRowNumber)] ?? null;
}

function getActualRowNumberAtDisplayRow(
    rowState: EditorDisplayRowState,
    displayRowNumber: number
): number | null {
    return rowState.actualRowNumbers[displayRowNumber - 1] ?? null;
}

function getNearestVisibleRowNumber(
    currentModel: EditorRenderModel,
    rowNumber: number
): number | null {
    const visibleRows = getVisibleActualRowsForModel(currentModel).visibleRows;
    if (visibleRows.length === 0) {
        return null;
    }

    if (visibleRows.includes(rowNumber)) {
        return rowNumber;
    }

    const nextVisibleRowNumber = visibleRows.find(
        (visibleRowNumber) => visibleRowNumber >= rowNumber
    );
    if (nextVisibleRowNumber) {
        return nextVisibleRowNumber;
    }

    return visibleRows[visibleRows.length - 1] ?? null;
}

function ensureSelectedCellVisibleForFilter(
    currentModel: EditorRenderModel | null = model
): boolean {
    if (!currentModel || !selectedCell) {
        return false;
    }

    if (selectedCell.rowNumber > currentModel.activeSheet.rowCount) {
        return false;
    }

    const nextVisibleRowNumber = getNearestVisibleRowNumber(currentModel, selectedCell.rowNumber);
    if (nextVisibleRowNumber === null || nextVisibleRowNumber === selectedCell.rowNumber) {
        return false;
    }

    updateSelectionControllerState((state) =>
        setControllerSelectedCell(state, {
            rowNumber: nextVisibleRowNumber,
            columnNumber: selectedCell!.columnNumber,
        })
    );
    syncSelectedCellToHost();
    return true;
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

function syncSearchPanelShellPosition(
    shell: HTMLElement | null = getSearchPanelShellElement()
): void {
    if (!searchPanelPosition) {
        return;
    }

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

function getVirtualViewportState(
    currentModel: EditorRenderModel | null
): VirtualViewportState | null {
    if (!currentModel) {
        return null;
    }

    const cachedMetrics = getCachedVirtualGridMetrics(currentModel);
    if (cachedMetrics) {
        return createVirtualViewportStateFromMetrics(cachedMetrics, getViewportElement());
    }

    const pane = getViewportElement();
    const viewportHeight = pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT;
    const viewportWidth = pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH;
    const scrollTop = pane?.scrollTop ?? 0;
    const scrollLeft = pane?.scrollLeft ?? 0;
    const sheetColumnLayout = getSheetColumnLayout(currentModel);
    const { visibleRowResult, rowLayout: baseVisibleRowLayout } =
        createBaseVisibleRowLayout(currentModel);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: visibleRowResult.visibleRows.length,
        columnCount: currentModel.activeSheet.columnCount,
        rowHeaderLabelCount: getFilterAwareRowHeaderLabelCount(
            currentModel,
            visibleRowResult.visibleRows.length
        ),
        viewportHeight,
        viewportWidth,
        rowLayout: baseVisibleRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const displayRowState = createEditorDisplayRowState(
        currentModel,
        visibleRowResult.visibleRows,
        visibleRowResult.hiddenRows,
        displayGrid.rowCount
    );
    const displayRowLayout = createEditorPixelRowLayout({
        rowCount: displayRowState.actualRowNumbers.length,
    });
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: displayRowState.actualRowNumbers.length,
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
            rowLayout: displayRowLayout,
            columnLayout: sheetColumnLayout,
        });
    const rowWindow = createEditorRowWindow({
        rowLayout: displayRowLayout,
        totalRows: displayGrid.rowCount,
        frozenRowCount,
        scrollTop,
        viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        columnLayout: displayColumnLayout,
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
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount)
            .map((displayRowNumber) =>
                getActualRowNumberAtDisplayRow(displayRowState, displayRowNumber)
            )
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        rowNumbers: rowWindow.rowNumbers
            .map((displayRowNumber) =>
                getActualRowNumberAtDisplayRow(displayRowState, displayRowNumber)
            )
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        columnNumbers: columnWindow.columnNumbers,
    };
}

function getPendingEditKey(sheetKey: string, rowNumber: number, columnNumber: number): string {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
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

function clearSuppressedSelectionClicks(): void {
    suppressNextCellClick = false;
    suppressNextRowHeaderClick = false;
    suppressNextColumnHeaderClick = false;
}

function scheduleSuppressedSelectionClickReset(): void {
    if (!suppressNextCellClick && !suppressNextRowHeaderClick && !suppressNextColumnHeaderClick) {
        return;
    }

    globalThis.setTimeout(() => {
        clearSuppressedSelectionClicks();
    }, 0);
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

function getEffectiveCellAlignment(
    rowNumber: number,
    columnNumber: number,
    currentModel: EditorRenderModel | null = model
): CellAlignmentSnapshot | null {
    if (!currentModel) {
        return null;
    }

    const cellKey = createCellKey(rowNumber, columnNumber);
    return mergeCellAlignments(
        currentModel.activeSheet.columnAlignments?.[String(columnNumber)] ?? null,
        currentModel.activeSheet.rowAlignments?.[String(rowNumber)] ?? null,
        currentModel.activeSheet.cellAlignments?.[cellKey] ?? null
    );
}

function getActiveToolbarAlignment(
    currentModel: EditorRenderModel | null = model
): EditorAlignmentPatch {
    if (!currentModel || !selectedCell) {
        return {};
    }

    const alignment = getEffectiveCellAlignment(
        selectedCell.rowNumber,
        selectedCell.columnNumber,
        currentModel
    );
    const horizontal = getToolbarHorizontalAlignment(alignment);
    const vertical = getToolbarVerticalAlignment(alignment);
    return {
        ...(horizontal ? { horizontal } : {}),
        ...(vertical ? { vertical } : {}),
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
    return getControllerSelectionRange(getSelectionControllerState());
}

function hasExpandedSelection(range: CellRange | null = getSelectionRange()): boolean {
    return hasExpandedSelectionRange(range);
}

function getExpandedSelectionRange(): CellRange | null {
    return getControllerExpandedSelectionRange(getSelectionControllerState());
}

function getActiveAlignmentSelectionTarget(
    currentModel: EditorRenderModel | null = model
): AlignmentSelectionTarget | null {
    if (!currentModel || !selectedCell) {
        return null;
    }

    const expandedSelection = getExpandedSelectionRange();
    if (!expandedSelection) {
        return {
            target: "cell",
            selection: {
                startRow: selectedCell.rowNumber,
                endRow: selectedCell.rowNumber,
                startColumn: selectedCell.columnNumber,
                endColumn: selectedCell.columnNumber,
            },
        };
    }

    const selectsAllColumns =
        expandedSelection.startColumn === 1 &&
        expandedSelection.endColumn === currentModel.activeSheet.columnCount;
    const selectsAllRows =
        expandedSelection.startRow === 1 &&
        expandedSelection.endRow === currentModel.activeSheet.rowCount;

    if (selectsAllColumns !== selectsAllRows) {
        return {
            target: selectsAllColumns ? "row" : "column",
            selection: expandedSelection,
        };
    }

    return {
        target: "range",
        selection: expandedSelection,
    };
}

function getSelectionRangeAddress(range: CellRange): string {
    return `${getCellAddressLabel(range.startRow, range.startColumn)}:${getCellAddressLabel(range.endRow, range.endColumn)}`;
}

function normalizeSearchPanelState(): void {
    const currentSelectionRange = getExpandedSelectionRange();

    if (!isSearchPanelOpen) {
        searchSelectionRange = currentSelectionRange;
        return;
    }

    if (currentSelectionRange) {
        const shouldAutoSwitchToSelection =
            searchScope !== "selection" &&
            !areSelectionRangesEqual(searchSelectionRange, currentSelectionRange);
        searchSelectionRange = currentSelectionRange;
        if (shouldAutoSwitchToSelection) {
            searchScope = "selection";
            searchFeedback = null;
        }
        return;
    }

    if (!currentSelectionRange) {
        if (searchScope === "selection") {
            searchScope = "sheet";
            searchFeedback = null;
        }
        searchSelectionRange = null;
        return;
    }
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
    return getControllerActiveHighlightCell(getSelectionControllerState());
}

function getSelectionExtendAnchorCell(): CellPosition | null {
    return getControllerSelectionExtendAnchorCell(getSelectionControllerState());
}

function getSelectionExtendAnchorRowNumber(): number | null {
    return getSelectionExtendAnchorCell()?.rowNumber ?? null;
}

function getSelectionExtendAnchorColumnNumber(): number | null {
    return getSelectionExtendAnchorCell()?.columnNumber ?? null;
}

function isActiveSelectionCell(rowNumber: number, columnNumber: number): boolean {
    return isControllerActiveSelectionCell(
        getSelectionControllerState(),
        rowNumber,
        columnNumber
    );
}

function getEffectiveCellValue(rowNumber: number, columnNumber: number): string {
    if (!model) {
        return "";
    }

    return getEffectiveCellValueForModel(model, rowNumber, columnNumber);
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

function getFilterToolbarActionLabel(currentModel: EditorRenderModel): string {
    if (
        canCreateEditorFilterRange(
            createEditorFilterCellSource(currentModel),
            getCandidateFilterRange(currentModel)
        )
    ) {
        return STRINGS.filterSelection;
    }

    return getActiveSheetFilterState(currentModel)
        ? STRINGS.clearFilterRange
        : STRINGS.filterSelection;
}

function getCandidateFilterRange(currentModel: EditorRenderModel): CellRange | null {
    const source = createEditorFilterCellSource(currentModel);
    const expandedSelection = getExpandedSelectionRange();

    return expandedSelection
        ? resolveEditorFilterRangeFromSelection(source, expandedSelection)
        : resolveEditorFilterRangeFromActiveCell(source, selectedCell);
}

function canToggleFilterForCurrentSelection(currentModel: EditorRenderModel): boolean {
    return (
        canCreateEditorFilterRange(
            createEditorFilterCellSource(currentModel),
            getCandidateFilterRange(currentModel)
        ) || Boolean(getActiveSheetFilterState(currentModel))
    );
}

function toggleActiveSheetFilter(): void {
    if (!model) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    const currentFilterState = getActiveSheetFilterState(model);
    const nextFilterState = toggleEditorSheetFilterState(
        createEditorFilterCellSource(model),
        currentFilterState,
        getCandidateFilterRange(model)
    );

    if (!nextFilterState && !currentFilterState) {
        return;
    }

    setSheetFilterState(model.activeSheet.key, nextFilterState);
    syncActiveSheetFilterStateToHost(model, nextFilterState);
    closeFilterMenu({ refresh: false });
    ensureSelectedCellVisibleForFilter(model);
    renderApp({ commitEditing: false, sync: true });
}

function openFilterMenu(rowNumber: number, columnNumber: number, element: HTMLElement): void {
    const currentModel = model;
    const filterState = currentModel ? getActiveSheetFilterState(currentModel) : null;
    if (
        !currentModel ||
        !filterState ||
        !isEditorFilterHeaderCell(filterState, rowNumber, columnNumber)
    ) {
        return;
    }

    if (
        filterMenuState?.sheetKey === currentModel.activeSheet.key &&
        filterMenuState.columnNumber === columnNumber
    ) {
        closeFilterMenu({ refresh: false });
        renderApp({ commitEditing: false, sync: true });
        return;
    }

    const appRect = getEditorAppElement()?.getBoundingClientRect();
    const triggerRect = element.getBoundingClientRect();
    const rawLeft = appRect ? triggerRect.right - appRect.left - 260 : triggerRect.left;
    const rawTop = appRect ? triggerRect.bottom - appRect.top + 6 : triggerRect.bottom + 6;
    const appWidth = getEditorAppElement()?.clientWidth ?? window.innerWidth;
    const appHeight = getEditorAppElement()?.clientHeight ?? window.innerHeight;

    filterMenuState = {
        sheetKey: currentModel.activeSheet.key,
        columnNumber,
        left: Math.max(8, Math.min(rawLeft, appWidth - 288)),
        top: Math.max(8, Math.min(rawTop, appHeight - 360)),
    };
    renderApp({ commitEditing: false, sync: true });
}

function applyActiveSheetFilterSort(
    columnNumber: number,
    direction: EditorFilterSortDirection
): void {
    if (!model) {
        return;
    }

    const filterState = getActiveSheetFilterState(model);
    if (!filterState) {
        return;
    }

    const nextFilterState = updateEditorFilterSort(filterState, columnNumber, direction);
    setSheetFilterState(model.activeSheet.key, nextFilterState);
    syncActiveSheetFilterStateToHost(model, nextFilterState);
    closeFilterMenu({ refresh: false });
    ensureSelectedCellVisibleForFilter(model);
    renderApp({ commitEditing: false, sync: true });
}

function setActiveSheetFilterIncludedValues(
    columnNumber: number,
    includedValues: readonly string[] | null
): void {
    if (!model) {
        return;
    }

    const filterState = getActiveSheetFilterState(model);
    if (!filterState) {
        return;
    }

    const nextFilterState = updateEditorFilterIncludedValues(
        filterState,
        columnNumber,
        includedValues
    );
    setSheetFilterState(model.activeSheet.key, nextFilterState);
    syncActiveSheetFilterStateToHost(model, nextFilterState);
    ensureSelectedCellVisibleForFilter(model);
    renderApp({ commitEditing: false, sync: true });
}

function clearActiveSheetFilterColumn(columnNumber: number): void {
    if (!model) {
        return;
    }

    const filterState = getActiveSheetFilterState(model);
    if (!filterState) {
        return;
    }

    const nextFilterState = clearEditorFilterColumn(filterState, columnNumber);
    setSheetFilterState(model.activeSheet.key, nextFilterState);
    syncActiveSheetFilterStateToHost(model, nextFilterState);
    closeFilterMenu({ refresh: false });
    ensureSelectedCellVisibleForFilter(model);
    renderApp({ commitEditing: false, sync: true });
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
    // Prompt-driven actions need an immediate re-render because no grid mutation follows.
    closeContextMenu();
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

function requestPromptRowHeight(rowNumber: number): void {
    // Prompt-driven actions need an immediate re-render because no grid mutation follows.
    closeContextMenu();
    vscode.postMessage({ type: "promptRowHeight", rowNumber });
}

function requestInsertColumn(columnNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "insertColumn", columnNumber });
}

function requestDeleteColumn(columnNumber: number): void {
    closeContextMenu({ refresh: false });
    vscode.postMessage({ type: "deleteColumn", columnNumber });
}

function requestPromptColumnWidth(columnNumber: number): void {
    // Prompt-driven actions need an immediate re-render because no grid mutation follows.
    closeContextMenu();
    vscode.postMessage({ type: "promptColumnWidth", columnNumber });
}

function beginColumnResize(
    pointerId: number,
    sheetKey: string,
    columnNumber: number,
    startPixelWidth: number,
    clientX: number
): void {
    if (!model?.canEdit) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    closeContextMenu({ refresh: false });
    clearBrowserTextSelection();
    suppressNextColumnHeaderClick = true;
    columnResizeState = {
        pointerId,
        sheetKey,
        columnNumber,
        startClientX: clientX,
        startPixelWidth,
        previewPixelWidth: startPixelWidth,
        isDragging: true,
    };
    renderApp({ commitEditing: false });
}

function updateColumnResize(clientX: number): void {
    if (!columnResizeState?.isDragging) {
        return;
    }

    const nextPixelWidth = Math.max(
        MIN_COLUMN_PIXEL_WIDTH,
        Math.min(
            MAX_COLUMN_PIXEL_WIDTH,
            stabilizeColumnPixelWidth(
                Math.round(
                    columnResizeState.startPixelWidth + clientX - columnResizeState.startClientX
                ),
                editorMaximumDigitWidth
            )
        )
    );
    if (nextPixelWidth === columnResizeState.previewPixelWidth) {
        return;
    }

    columnResizeState = {
        ...columnResizeState,
        previewPixelWidth: nextPixelWidth,
    };
    renderApp({ commitEditing: false });
}

function stopColumnResize(pointerId?: number, { commit = false }: { commit?: boolean } = {}): void {
    if (!columnResizeState) {
        return;
    }

    if (pointerId !== undefined && columnResizeState.pointerId !== pointerId) {
        return;
    }

    const activeResize = columnResizeState;
    const didChange = activeResize.previewPixelWidth !== activeResize.startPixelWidth;
    if (!commit || !didChange) {
        columnResizeState = null;
        renderApp({ commitEditing: false });
        return;
    }

    columnResizeState = {
        ...activeResize,
        startPixelWidth: activeResize.previewPixelWidth,
        isDragging: false,
    };
    renderApp({ commitEditing: false });
    vscode.postMessage({
        type: "setColumnWidth",
        columnNumber: activeResize.columnNumber,
        width: normalizeWorkbookColumnWidth(
            convertPixelsToWorkbookColumnWidth(
                activeResize.previewPixelWidth,
                editorMaximumDigitWidth
            )
        ),
    });
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

function getRowNumberFromElement(element: Element | null): number | null {
    const target =
        element?.closest<HTMLElement>('[data-role="grid-row-header"]') ??
        element?.closest<HTMLElement>('[data-role="grid-cell"]');
    if (!target) {
        return null;
    }

    const rowNumber = Number(target.dataset.rowNumber);
    return Number.isInteger(rowNumber) ? rowNumber : null;
}

function getColumnNumberFromElement(element: Element | null): number | null {
    const target =
        element?.closest<HTMLElement>('[data-role="grid-column-header"]') ??
        element?.closest<HTMLElement>('[data-role="grid-cell"]');
    if (!target) {
        return null;
    }

    const columnNumber = Number(target.dataset.columnNumber);
    return Number.isInteger(columnNumber) ? columnNumber : null;
}

function clearFillDragState(): void {
    fillDragState = null;
}

function clearFillHandleClickState(): void {
    lastFillHandleClickState = null;
}

function rememberFillHandleClick(sourceRange: CellRange, timeStamp: number): void {
    lastFillHandleClickState = {
        rowNumber: sourceRange.endRow,
        columnNumber: sourceRange.endColumn,
        timeStamp,
    };
}

function isRepeatedFillHandleClick(sourceRange: CellRange, timeStamp: number): boolean {
    return Boolean(
        lastFillHandleClickState &&
        lastFillHandleClickState.rowNumber === sourceRange.endRow &&
        lastFillHandleClickState.columnNumber === sourceRange.endColumn &&
        timeStamp >= lastFillHandleClickState.timeStamp &&
        timeStamp - lastFillHandleClickState.timeStamp <= FILL_HANDLE_DOUBLE_CLICK_MS
    );
}

function hasEditableCellInRange(range: CellRange): boolean {
    for (let rowNumber = range.startRow; rowNumber <= range.endRow; rowNumber += 1) {
        for (
            let columnNumber = range.startColumn;
            columnNumber <= range.endColumn;
            columnNumber += 1
        ) {
            if (canEditCellAt(rowNumber, columnNumber)) {
                return true;
            }
        }
    }

    return false;
}

function canShowFillHandle(range: CellRange | null = getSelectionRange()): boolean {
    return Boolean(model?.canEdit && !editingCell && range && hasEditableCellInRange(range));
}

function getFillHandleCell(range: CellRange | null): CellPosition | null {
    if (!range || !canShowFillHandle(range)) {
        return null;
    }

    return {
        rowNumber: range.endRow,
        columnNumber: range.endColumn,
    };
}

function startFillDrag(pointerId: number, sourceRange: CellRange): void {
    fillDragState = {
        pointerId,
        sourceRange,
        previewRange: null,
    };
    clearBrowserTextSelection();
    clearSuppressedSelectionClicks();
}

function updateFillDrag(targetCell: CellPosition): void {
    if (!fillDragState) {
        return;
    }

    const bounds = getSelectionBounds();
    if (!bounds) {
        return;
    }

    const nextPreviewRange = createFillPreviewRange(fillDragState.sourceRange, targetCell, bounds);
    if (areSelectionRangesEqual(fillDragState.previewRange, nextPreviewRange)) {
        return;
    }

    fillDragState = {
        ...fillDragState,
        previewRange: nextPreviewRange,
    };
    renderApp({ commitEditing: false });
}

function applyFillChanges(sourceRange: CellRange, previewRange: CellRange | null): boolean {
    if (!model || !previewRange) {
        return false;
    }

    const changes = buildFillChanges({
        sheetKey: model.activeSheet.key,
        sourceRange,
        previewRange,
        getCellValue: getEffectiveCellValue,
        getModelValue: getCellModelValue,
        canEditCell: canEditCellAt,
    });
    if (changes.length === 0) {
        return false;
    }

    applyEditChanges(changes, { revealSelection: true });
    return true;
}

function stopFillDrag(
    pointerId?: number,
    {
        commit = false,
        timeStamp,
    }: {
        commit?: boolean;
        timeStamp?: number;
    } = {}
): void {
    if (!fillDragState) {
        return;
    }

    if (pointerId !== undefined && fillDragState.pointerId !== pointerId) {
        return;
    }

    const currentFillDrag = fillDragState;
    fillDragState = null;

    if (!currentFillDrag.previewRange) {
        if (commit && typeof timeStamp === "number") {
            if (isRepeatedFillHandleClick(currentFillDrag.sourceRange, timeStamp)) {
                clearFillHandleClickState();
                handleFillHandleAutoFill(currentFillDrag.sourceRange);
                return;
            }

            rememberFillHandleClick(currentFillDrag.sourceRange, timeStamp);
            return;
        }

        clearFillHandleClickState();
        return;
    }

    clearFillHandleClickState();

    if (commit && applyFillChanges(currentFillDrag.sourceRange, currentFillDrag.previewRange)) {
        return;
    }

    if (currentFillDrag.previewRange) {
        renderApp({ commitEditing: false });
    }
}

function commitAutoFillDown(sourceRange: CellRange): void {
    const bounds = getSelectionBounds();
    if (!bounds) {
        return;
    }

    const previewRange = createAutoFillDownPreviewRange({
        sourceRange,
        bounds,
        getCellValue: getEffectiveCellValue,
    });
    if (!previewRange) {
        return;
    }

    applyFillChanges(sourceRange, previewRange);
}

function handleFillHandleAutoFill(sourceRange: CellRange): void {
    clearFillDragState();
    clearFillHandleClickState();
    closeContextMenu({ refresh: false });
    commitAutoFillDown(sourceRange);
}

function startSelectionDrag(pointerId: number, anchorCell: CellPosition): void {
    updateSelectionControllerState((state) =>
        startControllerCellSelectionDrag(state, pointerId, anchorCell)
    );
    clearSuppressedSelectionClicks();
}

function stopSelectionDrag(pointerId?: number): void {
    const currentState = getSelectionControllerState();
    if (!currentState.selectionDragState) {
        return;
    }

    const nextState = stopControllerSelectionDrag(currentState, pointerId);
    if (nextState === currentState) {
        return;
    }

    applySelectionControllerState(nextState);
    if (selectedCell) {
        syncSelectedCellToHost();
    }
}

function updateCellSelectionDrag(targetCell: CellPosition): void {
    if (!selectionDragState || selectionDragState.kind !== "cell") {
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

interface VirtualGridWheelLikeEvent {
    deltaX: number;
    deltaY: number;
    deltaMode: number;
    shiftKey: boolean;
    preventDefault(): void;
}

function forwardVirtualGridWheel(event: VirtualGridWheelLikeEvent): void {
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

    const cachedMetrics = getCachedVirtualGridMetrics(model);
    if (cachedMetrics) {
        let nextTop = pane.scrollTop;
        let nextLeft = pane.scrollLeft;

        const selectedDisplayRowNumber = getDisplayRowNumber(
            cachedMetrics.rowState,
            selectedCell.rowNumber
        );
        if (
            selectedDisplayRowNumber !== null &&
            selectedDisplayRowNumber > cachedMetrics.frozenRowCount
        ) {
            const cellTop =
                EDITOR_VIRTUAL_HEADER_HEIGHT +
                getEditorRowTop(cachedMetrics.rowLayout, selectedDisplayRowNumber);
            const cellBottom =
                cellTop + getEditorRowHeight(cachedMetrics.rowLayout, selectedDisplayRowNumber);
            const visibleTop = pane.scrollTop + cachedMetrics.stickyTopHeight;
            const visibleBottom = pane.scrollTop + pane.clientHeight;

            if (cellTop < visibleTop) {
                nextTop = cellTop - cachedMetrics.stickyTopHeight;
            } else if (cellBottom > visibleBottom) {
                nextTop = cellBottom - pane.clientHeight;
            }
        }

        if (selectedCell.columnNumber > cachedMetrics.frozenColumnCount) {
            const cellLeft =
                cachedMetrics.rowHeaderWidth +
                getEditorColumnLeft(cachedMetrics.columnLayout, selectedCell.columnNumber);
            const cellRight =
                cellLeft +
                getEditorColumnWidth(cachedMetrics.columnLayout, selectedCell.columnNumber);
            const visibleLeft = pane.scrollLeft + cachedMetrics.stickyLeftWidth;
            const visibleRight = pane.scrollLeft + pane.clientWidth;

            if (cellLeft < visibleLeft) {
                nextLeft = cellLeft - cachedMetrics.stickyLeftWidth;
            } else if (cellRight > visibleRight) {
                nextLeft = cellRight - pane.clientWidth;
            }
        }

        pane.scrollTop = clampEditorScrollPosition(
            nextTop,
            cachedMetrics.contentHeight - pane.clientHeight
        );
        pane.scrollLeft = clampEditorScrollPosition(
            nextLeft,
            cachedMetrics.contentWidth - pane.clientWidth
        );
        return;
    }

    const sheetColumnLayout = getSheetColumnLayout(model);
    const { visibleRowResult, rowLayout: baseVisibleRowLayout } = createBaseVisibleRowLayout(model);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: visibleRowResult.visibleRows.length,
        columnCount: model.activeSheet.columnCount,
        rowHeaderLabelCount: getFilterAwareRowHeaderLabelCount(
            model,
            visibleRowResult.visibleRows.length
        ),
        viewportHeight: pane.clientHeight,
        viewportWidth: pane.clientWidth,
        rowLayout: baseVisibleRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const displayRowState = createEditorDisplayRowState(
        model,
        visibleRowResult.visibleRows,
        visibleRowResult.hiddenRows,
        displayGrid.rowCount
    );
    const displayRowLayout = createEditorPixelRowLayout({
        rowCount: displayRowState.actualRowNumbers.length,
    });
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: displayRowState.actualRowNumbers.length,
        columnCount: model.activeSheet.columnCount,
        freezePane: model.activeSheet.freezePane,
    });
    const contentSize = getEditorContentSize({
        rowCount: displayGrid.rowCount,
        rowLayout: displayRowLayout,
        columnLayout: displayColumnLayout,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const stickyTop =
        EDITOR_VIRTUAL_HEADER_HEIGHT + getEditorFrozenRowsHeight(displayRowLayout, frozenRowCount);
    const stickyLeft =
        displayGrid.rowHeaderWidth +
        getEditorFrozenColumnsWidth(displayColumnLayout, frozenColumnCount);
    let nextTop = pane.scrollTop;
    let nextLeft = pane.scrollLeft;

    const selectedDisplayRowNumber = getDisplayRowNumber(displayRowState, selectedCell.rowNumber);
    if (selectedDisplayRowNumber !== null && selectedDisplayRowNumber > frozenRowCount) {
        const cellTop =
            EDITOR_VIRTUAL_HEADER_HEIGHT +
            getEditorRowTop(displayRowLayout, selectedDisplayRowNumber);
        const cellBottom = cellTop + getEditorRowHeight(displayRowLayout, selectedDisplayRowNumber);
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
            getEditorColumnLeft(displayColumnLayout, selectedCell.columnNumber);
        const cellRight =
            cellLeft + getEditorColumnWidth(displayColumnLayout, selectedCell.columnNumber);
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

    if (getRenderedCellElements(cell).length > 0) {
        return true;
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

function isCellSelectableInCurrentModel(
    cell: CellPosition | null,
    currentModel: EditorRenderModel | null = model
): boolean {
    if (!cell || !currentModel) {
        return false;
    }

    if (
        cell.rowNumber < 1 ||
        cell.rowNumber > currentModel.activeSheet.rowCount ||
        cell.columnNumber < 1 ||
        cell.columnNumber > currentModel.activeSheet.columnCount
    ) {
        return false;
    }

    return getNearestVisibleRowNumber(currentModel, cell.rowNumber) === cell.rowNumber;
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

function setRowSelectionLocal(
    rowNumber: number,
    {
        anchorRowNumber = rowNumber,
        reveal = false,
        syncHost = true,
        forceRender = false,
    }: {
        anchorRowNumber?: number;
        reveal?: boolean;
        syncHost?: boolean;
        forceRender?: boolean;
    } = {}
): void {
    if (!model) {
        return;
    }

    const selectionBounds = getSelectionBounds();
    const selectionRange = createRowSelectionSpanRange(
        anchorRowNumber,
        rowNumber,
        selectionBounds?.maxColumn ?? model.activeSheet.columnCount
    );
    if (!selectionRange) {
        return;
    }

    const columnNumber = getPreferredSelectionColumnNumber();
    setSelectedCellLocal(
        {
            rowNumber,
            columnNumber,
        },
        {
            reveal,
            syncHost,
            anchorCell: {
                rowNumber: anchorRowNumber,
                columnNumber,
            },
            selectionRange,
            forceRender,
        }
    );
}

function setColumnSelectionLocal(
    columnNumber: number,
    {
        anchorColumnNumber = columnNumber,
        reveal = false,
        syncHost = true,
        forceRender = false,
    }: {
        anchorColumnNumber?: number;
        reveal?: boolean;
        syncHost?: boolean;
        forceRender?: boolean;
    } = {}
): void {
    if (!model) {
        return;
    }

    const selectionBounds = getSelectionBounds();
    const selectionRange = createColumnSelectionSpanRange(
        anchorColumnNumber,
        columnNumber,
        selectionBounds?.maxRow ?? model.activeSheet.rowCount
    );
    if (!selectionRange) {
        return;
    }

    const rowNumber = getPreferredSelectionRowNumber();
    setSelectedCellLocal(
        {
            rowNumber,
            columnNumber,
        },
        {
            reveal,
            syncHost,
            anchorCell: {
                rowNumber,
                columnNumber: anchorColumnNumber,
            },
            selectionRange,
            forceRender,
        }
    );
}

function startRowSelectionDrag(pointerId: number, anchorRowNumber: number): void {
    updateSelectionControllerState((state) =>
        startControllerRowSelectionDrag(state, pointerId, anchorRowNumber)
    );
    clearSuppressedSelectionClicks();
    setRowSelectionLocal(anchorRowNumber, {
        anchorRowNumber,
        syncHost: false,
    });
}

function updateRowSelectionDrag(targetRowNumber: number): void {
    if (!selectionDragState || selectionDragState.kind !== "row") {
        return;
    }

    if (selectedCell?.rowNumber === targetRowNumber) {
        return;
    }

    suppressNextRowHeaderClick = true;
    setRowSelectionLocal(targetRowNumber, {
        anchorRowNumber: selectionDragState.anchorRowNumber,
        syncHost: false,
    });
}

function startColumnSelectionDrag(pointerId: number, anchorColumnNumber: number): void {
    updateSelectionControllerState((state) =>
        startControllerColumnSelectionDrag(state, pointerId, anchorColumnNumber)
    );
    clearSuppressedSelectionClicks();
    setColumnSelectionLocal(anchorColumnNumber, {
        anchorColumnNumber,
        syncHost: false,
    });
}

function updateColumnSelectionDrag(targetColumnNumber: number): void {
    if (!selectionDragState || selectionDragState.kind !== "column") {
        return;
    }

    if (selectedCell?.columnNumber === targetColumnNumber) {
        return;
    }

    suppressNextColumnHeaderClick = true;
    setColumnSelectionLocal(targetColumnNumber, {
        anchorColumnNumber: selectionDragState.anchorColumnNumber,
        syncHost: false,
    });
}

function selectEntireRow(rowNumber: number): void {
    if (!model) {
        return;
    }

    const forceRender = Boolean(editingCell);
    if (forceRender) {
        finishEdit({ mode: "commit", refresh: false });
    }

    setRowSelectionLocal(rowNumber, { syncHost: true, forceRender });
}

function selectEntireColumn(columnNumber: number): void {
    if (!model) {
        return;
    }

    const forceRender = Boolean(editingCell);
    if (forceRender) {
        finishEdit({ mode: "commit", refresh: false });
    }

    setColumnSelectionLocal(columnNumber, { syncHost: true, forceRender });
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
    void forceRender;
    const nextSelectionState = setControllerSelectedCell(
        getSelectionControllerState(),
        nextCell,
        {
            anchorCell,
            selectionRangeOverride: selectionRange ?? null,
        }
    );
    const shouldClearSearchFeedback = clearSearchFeedback && Boolean(searchFeedback);

    clearFillDragState();
    applySelectionControllerState(nextSelectionState);
    clearBrowserTextSelection();
    if (shouldClearSearchFeedback) {
        searchFeedback = null;
    }

    if (syncHost) {
        syncSelectedCellToHost();
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

    if (useModelSelection && model?.selection) {
        const nextRowNumber =
            getNearestVisibleRowNumber(model, model.selection.rowNumber) ??
            model.selection.rowNumber;
        clearFillDragState();
        applySelectionControllerState(
            clearControllerPendingSelectionAfterRender(
                setControllerSelectedCell(getSelectionControllerState(), {
                    rowNumber: nextRowNumber,
                    columnNumber: model.selection.columnNumber,
                })
            )
        );
        syncSelectedCellToHost();
        return shouldReveal;
    }

    if (pendingSelectionAfterRender) {
        const nextRowNumber =
            getNearestVisibleRowNumber(model!, pendingSelectionAfterRender.rowNumber) ??
            pendingSelectionAfterRender.rowNumber;
        clearFillDragState();
        shouldReveal = pendingSelectionAfterRender.reveal;
        applySelectionControllerState(
            clearControllerPendingSelectionAfterRender(
                setControllerSelectedCell(getSelectionControllerState(), {
                    rowNumber: nextRowNumber,
                    columnNumber: pendingSelectionAfterRender.columnNumber,
                })
            )
        );
        syncSelectedCellToHost();
        return shouldReveal;
    }

    applySelectionControllerState(
        clearControllerPendingSelectionAfterRender(getSelectionControllerState())
    );

    ensureSelectedCellVisibleForFilter(model);

    if (isCellSelectableInCurrentModel(selectedCell)) {
        if (
            !selectionRangeOverride &&
            !hasExpandedSelection() &&
            !isCellSelectableInCurrentModel(selectionAnchorCell)
        ) {
            updateSelectionControllerState((state) =>
                syncControllerSelectionAnchorToSelectedCell(state)
            );
        }
        syncSelectedCellToHost();
        return shouldReveal;
    }

    if (suppressAutoSelection) {
        return shouldReveal;
    }

    const nextSelectedCell = model?.selection
        ? {
              rowNumber:
                  getNearestVisibleRowNumber(model, model.selection.rowNumber) ??
                  model.selection.rowNumber,
              columnNumber: model.selection.columnNumber,
          }
        : model && model.activeSheet.rowCount > 0 && model.activeSheet.columnCount > 0
          ? { rowNumber: getNearestVisibleRowNumber(model, 1) ?? 1, columnNumber: 1 }
          : null;
    clearFillDragState();
    updateSelectionControllerState((state) =>
        nextSelectedCell
            ? setControllerSelectedCell(state, nextSelectedCell)
            : {
                  ...state,
                  selectedCell: null,
                  selectionRangeOverride: null,
              }
    );

    if (selectedCell) {
        syncSelectedCellToHost();
    }

    return shouldReveal;
}

function getCellModelValue(rowNumber: number, columnNumber: number): string {
    return getCellView(rowNumber, columnNumber)?.value ?? "";
}

function getGridCellTextLayoutMetrics(cellHeight: number): {
    contentMaxHeightPx: number;
    visibleLineCount: number;
} {
    const contentMaxHeightPx = Math.max(18, cellHeight - GRID_CELL_VERTICAL_PADDING_PX);
    return {
        contentMaxHeightPx,
        visibleLineCount: Math.max(
            1,
            Math.floor(contentMaxHeightPx / GRID_CELL_LINE_HEIGHT_PX)
        ),
    };
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

    clearFillDragState();
    editingCell = null;

    if (mode === "commit") {
        commitEdit(session.sheetKey, session.rowNumber, session.columnNumber, session.value, {
            refresh: false,
        });
    }

    if (clearSelection) {
        applySelectionControllerState(clearControllerSelectedCell(getSelectionControllerState()));
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
    const bounds = getSelectionBounds();
    if (!bounds) {
        return;
    }

    const maxRow = bounds.maxRow;
    const maxColumn = bounds.maxColumn;
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
    clearFillDragState();
    applySelectionControllerState(
        setControllerSelectedCell(getSelectionControllerState(), {
            rowNumber,
            columnNumber,
        })
    );
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

    const cachedMetrics = getCachedVirtualGridMetrics(model);
    if (cachedMetrics) {
        return {
            minRow: 1,
            maxRow:
                cachedMetrics.rowState.actualRowNumbers[
                    cachedMetrics.rowState.actualRowNumbers.length - 1
                ] ?? 1,
            minColumn: 1,
            maxColumn: cachedMetrics.columnLayout.totalColumnCount,
        };
    }

    const pane = getViewportElement();
    const { displayGrid, rowState } = createDisplayGridRowState(
        model,
        pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT,
        pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH
    );

    return {
        minRow: 1,
        maxRow: rowState.actualRowNumbers[rowState.actualRowNumbers.length - 1] ?? 1,
        minColumn: 1,
        maxColumn: displayGrid.columnCount,
    };
}

function ensureSelection(): CellPosition | null {
    if (selectedCell) {
        let nextState = getSelectionControllerState();
        if (model && selectedCell.rowNumber <= model.activeSheet.rowCount) {
            const nextVisibleRowNumber = getNearestVisibleRowNumber(model, selectedCell.rowNumber);
            if (nextVisibleRowNumber !== null && nextVisibleRowNumber !== selectedCell.rowNumber) {
                nextState = {
                    ...nextState,
                    selectedCell: {
                        rowNumber: nextVisibleRowNumber,
                        columnNumber: nextState.selectedCell!.columnNumber,
                    },
                };
            }
        }
        if (!nextState.selectionAnchorCell && nextState.selectedCell) {
            nextState = setControllerSelectionAnchorCell(nextState, nextState.selectedCell);
        }
        nextState = setControllerSuppressAutoSelection(nextState, false);
        applySelectionControllerState(nextState);
        return selectedCell;
    }

    if (model?.selection) {
        const rowNumber =
            getNearestVisibleRowNumber(model, model.selection.rowNumber) ??
            model.selection.rowNumber;
        applySelectionControllerState(
            setControllerSelectedCell(getSelectionControllerState(), {
                rowNumber,
                columnNumber: model.selection.columnNumber,
            })
        );
        return selectedCell;
    }

    if (model && model.activeSheet.rowCount > 0 && model.activeSheet.columnCount > 0) {
        const rowNumber = getNearestVisibleRowNumber(model, 1) ?? 1;
        applySelectionControllerState(
            setControllerSelectedCell(getSelectionControllerState(), {
                rowNumber,
                columnNumber: 1,
            })
        );
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
    if (!selection || !bounds || !model) {
        return;
    }

    const pane = getViewportElement();
    const rowState =
        getCachedVirtualGridMetrics(model)?.rowState ??
        createDisplayGridRowState(
            model,
            pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT,
            pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH
        ).rowState;
    const currentDisplayRowNumber = getDisplayRowNumber(rowState, selection.rowNumber) ?? 1;
    const nextDisplayRowNumber = Math.max(
        1,
        Math.min(rowState.actualRowNumbers.length, currentDisplayRowNumber + rowDelta)
    );
    const nextRow =
        rowDelta === 0
            ? selection.rowNumber
            : (getActualRowNumberAtDisplayRow(rowState, nextDisplayRowNumber) ??
              selection.rowNumber);
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
    const pane = getViewportElement();
    const rowState =
        getCachedVirtualGridMetrics(model)?.rowState ??
        createDisplayGridRowState(
            model,
            pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT,
            pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH
        ).rowState;
    const shiftRowCount = Math.max(1, (viewportState?.rowNumbers.length ?? 1) - 1);
    const currentDisplayRowNumber = getDisplayRowNumber(rowState, selection.rowNumber) ?? 1;
    const nextDisplayRowNumber = Math.max(
        1,
        Math.min(
            rowState.actualRowNumbers.length,
            currentDisplayRowNumber + direction * shiftRowCount
        )
    );
    const nextRowNumber =
        getActualRowNumberAtDisplayRow(rowState, nextDisplayRowNumber) ?? selection.rowNumber;
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
        role === "position" ? '[data-role="position-input"]' : '[data-role="cell-value-input"]';
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

function toggleSearchOption(option: keyof SearchOptions): void {
    searchOptions = {
        ...searchOptions,
        [option]: !searchOptions[option],
    };
    searchFeedback = null;
    renderApp({ commitEditing: false });
}

function getEffectiveSearchScope(): EditorSearchScope {
    return searchScope === "selection" &&
        Boolean(searchSelectionRange ?? getExpandedSelectionRange())
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

function getVisibleSearchCellsForModel(
    currentModel: EditorRenderModel
): Record<string, CellSnapshot> {
    const filterState = getActiveSheetFilterState(currentModel);
    if (!filterState) {
        return currentModel.activeSheet.cells;
    }

    const visibleRows = new Set(
        getVisibleActualRowsForModel(currentModel).visibleRows.filter(
            (rowNumber) => rowNumber <= currentModel.activeSheet.rowCount
        )
    );

    return Object.fromEntries(
        Object.entries(currentModel.activeSheet.cells).filter(([, cell]) =>
            visibleRows.has(cell.rowNumber)
        )
    );
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
    if (!normalizedQuery || !model) {
        focusSearchInput();
        return;
    }

    const effectiveScope = getEffectiveSearchScope();
    const selectionRange = getActiveSearchSelectionRange(effectiveScope);

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
            cells: getVisibleSearchCellsForModel(model),
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
        refreshCurrentAppView({ sync: true });
        return;
    }

    if (result.status === "no-match") {
        searchFeedback = {
            status: "no-match",
            message: STRINGS.replaceNoEditableMatches,
        };
        refreshCurrentAppView({ sync: true });
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

        refreshCurrentAppView({ sync: true });
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
        refreshCurrentAppView({ sync: true });
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

function refreshCurrentAppView({ sync = false }: { sync?: boolean } = {}): void {
    if (!model) {
        return;
    }

    updateView(
        {
            kind: "app",
            model,
            revealSelection: false,
            revision: viewRevision,
            scrollState: getPaneScrollState(),
        },
        { sync }
    );
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
    updateView(
        {
            kind: "app",
            model,
            revealSelection: shouldRevealSelection,
            revision: viewRevision,
            scrollState,
        },
        { sync }
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

function getFilterOptionLabel(value: string): string {
    return value ? value : STRINGS.filterBlankValue;
}

function EditorFilterMenu({
    currentModel,
    filterState,
    menuState,
}: {
    currentModel: EditorRenderModel;
    filterState: EditorSheetFilterState | null;
    menuState: FilterMenuState | null;
}): React.ReactElement | null {
    const [query, setQuery] = React.useState("");

    React.useEffect(() => {
        setQuery("");
    }, [currentModel.activeSheet.key, menuState?.columnNumber]);

    if (
        !filterState ||
        !menuState ||
        menuState.sheetKey !== currentModel.activeSheet.key ||
        menuState.columnNumber < filterState.range.startColumn ||
        menuState.columnNumber > filterState.range.endColumn
    ) {
        return null;
    }

    const filterOptions = getEditorFilterColumnValues(
        createEditorFilterCellSource(currentModel),
        filterState,
        menuState.columnNumber
    );
    const allValues = filterOptions.map((option) => option.value);
    const includedValues =
        filterState.includedValuesByColumn[String(menuState.columnNumber)] ?? allValues;
    const includedValuesSet = new Set(includedValues);
    const normalizedQuery = query.trim().toLowerCase();
    const filteredOptions = normalizedQuery
        ? filterOptions.filter((option) =>
              getFilterOptionLabel(option.value).toLowerCase().includes(normalizedQuery)
          )
        : filterOptions;
    const isAllSelected =
        filterOptions.length > 0 && includedValues.length === filterOptions.length;
    const hasColumnCriteria =
        Boolean(filterState.includedValuesByColumn[String(menuState.columnNumber)]) ||
        filterState.sort?.columnNumber === menuState.columnNumber;

    return (
        <div
            className="filter-menu"
            data-role="filter-menu"
            style={{
                left: `${menuState.left}px`,
                top: `${menuState.top}px`,
            }}
        >
            <div className="filter-menu__sorts">
                <button
                    className={classNames([
                        "filter-menu__sortButton",
                        filterState.sort?.columnNumber === menuState.columnNumber &&
                            filterState.sort.direction === "asc" &&
                            "is-active",
                    ])}
                    type="button"
                    onClick={() => applyActiveSheetFilterSort(menuState.columnNumber, "asc")}
                >
                    <ImSortAlphaAsc aria-hidden />
                    <span>{STRINGS.sortAscending}</span>
                </button>
                <button
                    className={classNames([
                        "filter-menu__sortButton",
                        filterState.sort?.columnNumber === menuState.columnNumber &&
                            filterState.sort.direction === "desc" &&
                            "is-active",
                    ])}
                    type="button"
                    onClick={() => applyActiveSheetFilterSort(menuState.columnNumber, "desc")}
                >
                    <ImSortAlphaDesc aria-hidden />
                    <span>{STRINGS.sortDescending}</span>
                </button>
            </div>
            <div className="filter-menu__search">
                <span className="codicon codicon-search filter-menu__searchIcon" aria-hidden />
                <input
                    className="filter-menu__searchInput"
                    placeholder={STRINGS.filterSearchPlaceholder}
                    type="text"
                    value={query}
                    onChange={(event) => setQuery(event.currentTarget.value)}
                />
            </div>
            <div className="filter-menu__values">
                {filterOptions.length > 0 ? (
                    <>
                        <label className="filter-menu__option">
                            <input
                                checked={isAllSelected}
                                type="checkbox"
                                onChange={(event) =>
                                    setActiveSheetFilterIncludedValues(
                                        menuState.columnNumber,
                                        event.currentTarget.checked ? null : []
                                    )
                                }
                            />
                            <span>{STRINGS.filterSelectAll}</span>
                        </label>
                        <div className="filter-menu__divider" />
                        <div className="filter-menu__optionList">
                            {filteredOptions.map((option) => {
                                const optionLabel = getFilterOptionLabel(option.value);
                                const isChecked = includedValuesSet.has(option.value);
                                return (
                                    <label
                                        key={`${menuState.columnNumber}:${option.value}`}
                                        className="filter-menu__option"
                                    >
                                        <input
                                            checked={isChecked}
                                            type="checkbox"
                                            onChange={(event) => {
                                                const nextIncludedValues = new Set(includedValues);
                                                if (event.currentTarget.checked) {
                                                    nextIncludedValues.add(option.value);
                                                } else {
                                                    nextIncludedValues.delete(option.value);
                                                }

                                                const normalizedIncludedValues = allValues.filter(
                                                    (value) => nextIncludedValues.has(value)
                                                );
                                                setActiveSheetFilterIncludedValues(
                                                    menuState.columnNumber,
                                                    normalizedIncludedValues.length ===
                                                        allValues.length
                                                        ? null
                                                        : normalizedIncludedValues
                                                );
                                            }}
                                        />
                                        <span className="filter-menu__optionLabel">
                                            {optionLabel}
                                        </span>
                                        <span className="filter-menu__optionCount">
                                            {option.count}
                                        </span>
                                    </label>
                                );
                            })}
                        </div>
                    </>
                ) : (
                    <div className="filter-menu__empty">{STRINGS.filterNoValues}</div>
                )}
            </div>
            <div className="filter-menu__footer">
                <button
                    className="filter-menu__clearButton"
                    disabled={!hasColumnCriteria}
                    type="button"
                    onClick={() => clearActiveSheetFilterColumn(menuState.columnNumber)}
                >
                    {STRINGS.filterClearColumn}
                </button>
            </div>
        </div>
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

function isActiveHighlightColumn(activeColumnNumber: number | null, columnNumber: number): boolean {
    return activeColumnNumber === columnNumber;
}

function isRowHeaderWithinSelectionRange(
    selectionRange: CellRange | null,
    rowNumber: number,
    columnCount: number
): boolean {
    return Boolean(
        selectionRange &&
        selectionRange.startColumn === 1 &&
        selectionRange.endColumn === columnCount &&
        rowNumber >= selectionRange.startRow &&
        rowNumber <= selectionRange.endRow
    );
}

function isColumnHeaderWithinSelectionRange(
    selectionRange: CellRange | null,
    columnNumber: number,
    rowCount: number
): boolean {
    return Boolean(
        selectionRange &&
        selectionRange.startRow === 1 &&
        selectionRange.endRow === rowCount &&
        columnNumber >= selectionRange.startColumn &&
        columnNumber <= selectionRange.endColumn
    );
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
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
    rowState: EditorDisplayRowState;
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

function getCachedVirtualGridMetrics(
    currentModel: EditorRenderModel | null
): EditorVirtualGridMetrics | null {
    if (!currentModel || !latestVirtualGridCache || latestVirtualGridCache.model !== currentModel) {
        return null;
    }

    return latestVirtualGridCache.metrics;
}

function createVirtualViewportStateFromMetrics(
    metrics: EditorVirtualGridMetrics,
    element: HTMLElement | null
): VirtualViewportState {
    const viewportHeight = element?.clientHeight ?? metrics.viewportHeight;
    const viewportWidth = element?.clientWidth ?? metrics.viewportWidth;
    const scrollTop = element?.scrollTop ?? metrics.scrollTop;
    const scrollLeft = element?.scrollLeft ?? metrics.scrollLeft;
    const { rowCount: visibleFrozenRowCount, columnCount: visibleFrozenColumnCount } =
        getVisibleFrozenEditorCounts({
            frozenRowCount: metrics.frozenRowCount,
            frozenColumnCount: metrics.frozenColumnCount,
            viewportHeight,
            viewportWidth,
            rowHeaderWidth: metrics.rowHeaderWidth,
            rowLayout: metrics.rowLayout,
            columnLayout: metrics.columnLayout,
        });
    const rowWindow = createEditorRowWindow({
        rowLayout: metrics.rowLayout,
        totalRows: metrics.rowLayout.totalRowCount,
        frozenRowCount: metrics.frozenRowCount,
        scrollTop,
        viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        columnLayout: metrics.columnLayout,
        frozenColumnCount: metrics.frozenColumnCount,
        scrollLeft,
        viewportWidth,
        rowHeaderWidth: metrics.rowHeaderWidth,
    });

    return {
        scrollTop,
        scrollLeft,
        viewportHeight,
        viewportWidth,
        rowHeaderWidth: metrics.rowHeaderWidth,
        frozenRowCount: metrics.frozenRowCount,
        frozenColumnCount: metrics.frozenColumnCount,
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount)
            .map((displayRowNumber) =>
                getActualRowNumberAtDisplayRow(metrics.rowState, displayRowNumber)
            )
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        rowNumbers: rowWindow.rowNumbers
            .map((displayRowNumber) =>
                getActualRowNumberAtDisplayRow(metrics.rowState, displayRowNumber)
            )
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        columnNumbers: columnWindow.columnNumbers,
    };
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

function areColumnLayoutsEqual(left: PixelColumnLayout, right: PixelColumnLayout): boolean {
    return (
        left.totalColumnCount === right.totalColumnCount &&
        left.totalWidth === right.totalWidth &&
        left.fallbackPixelWidth === right.fallbackPixelWidth &&
        areNumberArraysEqual(left.pixelWidths, right.pixelWidths)
    );
}

function areRowLayoutsEqual(left: PixelRowLayout, right: PixelRowLayout): boolean {
    return (
        left.totalRowCount === right.totalRowCount &&
        left.totalHeight === right.totalHeight &&
        left.fallbackPixelHeight === right.fallbackPixelHeight &&
        areNumberArraysEqual(left.overriddenRowNumbers, right.overriddenRowNumbers) &&
        left.overriddenRowNumbers.every(
            (rowNumber) =>
                left.overriddenPixelHeights[String(rowNumber)] ===
                right.overriddenPixelHeights[String(rowNumber)]
        )
    );
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
        !areRowLayoutsEqual(previous.rowLayout, next.rowLayout) ||
        !areColumnLayoutsEqual(previous.columnLayout, next.columnLayout) ||
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

function getEditorGridTop(rowLayout: PixelRowLayout, rowNumber: number): number {
    return EDITOR_VIRTUAL_HEADER_HEIGHT + getEditorRowTop(rowLayout, rowNumber);
}

function getEditorGridLeft(
    rowHeaderWidth: number,
    columnLayout: PixelColumnLayout,
    columnNumber: number
): number {
    return rowHeaderWidth + getEditorColumnLeft(columnLayout, columnNumber);
}

function getEditorGridTopForActualRow(
    metrics: Pick<EditorVirtualGridMetrics, "rowLayout" | "rowState">,
    actualRowNumber: number
): number | null {
    const displayRowNumber = getDisplayRowNumber(metrics.rowState, actualRowNumber);
    return displayRowNumber === null ? null : getEditorGridTop(metrics.rowLayout, displayRowNumber);
}

function getEditorGridRowHeightForActualRow(
    metrics: Pick<EditorVirtualGridMetrics, "rowLayout" | "rowState">,
    actualRowNumber: number
): number | null {
    const displayRowNumber = getDisplayRowNumber(metrics.rowState, actualRowNumber);
    return displayRowNumber === null
        ? null
        : getEditorRowHeight(metrics.rowLayout, displayRowNumber);
}

function getGridLayerForCell(
    metrics: EditorVirtualGridMetrics,
    cell: CellPosition | null
): GridLayerKind | null {
    if (!cell) {
        return null;
    }

    const isFrozenRow = metrics.frozenRowNumbers.includes(cell.rowNumber);
    const isFrozenColumn = metrics.frozenColumnNumbers.includes(cell.columnNumber);
    const isVisibleRow = isFrozenRow || metrics.rowNumbers.includes(cell.rowNumber);
    const isVisibleColumn = isFrozenColumn || metrics.columnNumbers.includes(cell.columnNumber);

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

interface SelectionOverlayRect {
    top: number;
    left: number;
    width: number;
    height: number;
}

interface SelectionOverlayRangeRect extends SelectionOverlayRect {
    showTopBorder: boolean;
    showRightBorder: boolean;
    showBottomBorder: boolean;
    showLeftBorder: boolean;
}

function getGridLayerRowNumbers(
    metrics: EditorVirtualGridMetrics,
    layer: GridLayerKind
): readonly number[] {
    if (layer === "top" || layer === "corner") {
        return metrics.frozenRowNumbers;
    }

    return metrics.rowNumbers;
}

function getGridLayerColumnNumbers(
    metrics: EditorVirtualGridMetrics,
    layer: GridLayerKind
): readonly number[] {
    if (layer === "left" || layer === "corner") {
        return metrics.frozenColumnNumbers;
    }

    return metrics.columnNumbers;
}

function createSelectionOverlayRect(
    metrics: EditorVirtualGridMetrics,
    selectionRange: CellRange | null,
    layer: GridLayerKind
): SelectionOverlayRect | null {
    const visibleRange = clipSelectionRangeToVisibleGrid(
        selectionRange,
        getGridLayerRowNumbers(metrics, layer),
        getGridLayerColumnNumbers(metrics, layer)
    );
    if (!visibleRange) {
        return null;
    }

    const top = getEditorGridTopForActualRow(metrics, visibleRange.startRow);
    const bottomTop = getEditorGridTopForActualRow(metrics, visibleRange.endRow);
    if (top === null || bottomTop === null) {
        return null;
    }

    const left = getEditorGridLeft(
        metrics.rowHeaderWidth,
        metrics.columnLayout,
        visibleRange.startColumn
    );
    const right =
        getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, visibleRange.endColumn) +
        getEditorColumnWidth(metrics.columnLayout, visibleRange.endColumn);
    const bottom =
        bottomTop + (getEditorGridRowHeightForActualRow(metrics, visibleRange.endRow) ?? 0);

    return {
        top,
        left,
        width: right - left,
        height: bottom - top,
    };
}

function createSelectionOverlayRangeRect(
    metrics: EditorVirtualGridMetrics,
    selectionRange: CellRange | null,
    layer: GridLayerKind
): SelectionOverlayRangeRect | null {
    const visibleRange = clipSelectionRangeToVisibleGrid(
        selectionRange,
        getGridLayerRowNumbers(metrics, layer),
        getGridLayerColumnNumbers(metrics, layer)
    );
    if (!selectionRange || !visibleRange) {
        return null;
    }

    const rect = createSelectionOverlayRect(metrics, selectionRange, layer);
    if (!rect) {
        return null;
    }

    return {
        ...rect,
        showTopBorder: visibleRange.startRow === selectionRange.startRow,
        showRightBorder: visibleRange.endColumn === selectionRange.endColumn,
        showBottomBorder: visibleRange.endRow === selectionRange.endRow,
        showLeftBorder: visibleRange.startColumn === selectionRange.startColumn,
    };
}

function createSelectionOverlayCellRect(
    metrics: EditorVirtualGridMetrics,
    cell: CellPosition | null,
    layer: GridLayerKind
): SelectionOverlayRect | null {
    if (!cell || getGridLayerForCell(metrics, cell) !== layer) {
        return null;
    }

    return createSelectionOverlayRect(
        metrics,
        {
            startRow: cell.rowNumber,
            endRow: cell.rowNumber,
            startColumn: cell.columnNumber,
            endColumn: cell.columnNumber,
        },
        layer
    );
}

function getLastNumber(values: readonly number[]): number | null {
    return values[values.length - 1] ?? null;
}

function createSelectionOverlayRowRect(
    metrics: EditorVirtualGridMetrics,
    rowNumber: number | null,
    layer: GridLayerKind
): SelectionOverlayRect | null {
    const columnNumbers = getGridLayerColumnNumbers(metrics, layer);
    const startColumn = columnNumbers[0] ?? null;
    const endColumn = getLastNumber(columnNumbers);
    if (rowNumber === null || startColumn === null || endColumn === null) {
        return null;
    }

    return createSelectionOverlayRect(
        metrics,
        {
            startRow: rowNumber,
            endRow: rowNumber,
            startColumn,
            endColumn,
        },
        layer
    );
}

function createSelectionOverlayColumnRect(
    metrics: EditorVirtualGridMetrics,
    columnNumber: number | null,
    layer: GridLayerKind
): SelectionOverlayRect | null {
    const rowNumbers = getGridLayerRowNumbers(metrics, layer);
    const startRow = rowNumbers[0] ?? null;
    const endRow = getLastNumber(rowNumbers);
    if (columnNumber === null || startRow === null || endRow === null) {
        return null;
    }

    return createSelectionOverlayRect(
        metrics,
        {
            startRow,
            endRow,
            startColumn: columnNumber,
            endColumn: columnNumber,
        },
        layer
    );
}

function getSelectionOverlayStyle(
    rect: SelectionOverlayRect,
    borders?: Pick<
        SelectionOverlayRangeRect,
        "showTopBorder" | "showRightBorder" | "showBottomBorder" | "showLeftBorder"
    >
): React.CSSProperties {
    return {
        ...getEditorGridItemStyle(rect),
        ...(borders
            ? ({
                  "--selection-overlay-border-top": borders.showTopBorder ? "2px" : "0px",
                  "--selection-overlay-border-right": borders.showRightBorder ? "2px" : "0px",
                  "--selection-overlay-border-bottom": borders.showBottomBorder ? "2px" : "0px",
                  "--selection-overlay-border-left": borders.showLeftBorder ? "2px" : "0px",
              } as React.CSSProperties)
            : {}),
    };
}

function EditorSelectionOverlay({
    metrics,
    layer,
    selectionRange,
    activeRowNumber,
    activeColumnNumber,
    currentSelection,
    showPrimarySelectionFrame,
}: {
    metrics: EditorVirtualGridMetrics;
    layer: GridLayerKind;
    selectionRange: CellRange | null;
    activeRowNumber: number | null;
    activeColumnNumber: number | null;
    currentSelection: CellPosition | null;
    showPrimarySelectionFrame: boolean;
}): React.ReactElement | null {
    const activeRowRect = createSelectionOverlayRowRect(metrics, activeRowNumber, layer);
    const activeColumnRect = createSelectionOverlayColumnRect(metrics, activeColumnNumber, layer);
    const rangeRect = createSelectionOverlayRangeRect(metrics, selectionRange, layer);
    const primaryCellRect = showPrimarySelectionFrame
        ? createSelectionOverlayCellRect(metrics, currentSelection, layer)
        : null;

    if (!activeRowRect && !activeColumnRect && !rangeRect && !primaryCellRect) {
        return null;
    }

    return (
        <>
            {activeRowRect ? (
                <div
                    aria-hidden
                    className="editor-grid__selection-overlay editor-grid__selection-overlay--active-row"
                    style={getSelectionOverlayStyle(activeRowRect)}
                />
            ) : null}
            {activeColumnRect ? (
                <div
                    aria-hidden
                    className="editor-grid__selection-overlay editor-grid__selection-overlay--active-column"
                    style={getSelectionOverlayStyle(activeColumnRect)}
                />
            ) : null}
            {rangeRect ? (
                <div
                    aria-hidden
                    className="editor-grid__selection-overlay editor-grid__selection-overlay--range"
                    style={getSelectionOverlayStyle(rangeRect, rangeRect)}
                />
            ) : null}
            {primaryCellRect ? (
                <div
                    aria-hidden
                    className="editor-grid__selection-overlay editor-grid__selection-overlay--primary"
                    style={getSelectionOverlayStyle(primaryCellRect)}
                />
            ) : null}
        </>
    );
}

function getFilterAwareRowHeaderLabelCount(
    currentModel: EditorRenderModel,
    visibleActualRowCount: number
): number {
    return Math.max(
        currentModel.activeSheet.rowCount + EDITOR_EXTRA_PADDING_ROWS,
        visibleActualRowCount
    );
}

function mapActualRowHeightsToDisplayRowHeights(
    actualToDisplayRowNumbers: Readonly<Record<string, number>>,
    rowHeights: Readonly<Record<string, number | null>> | undefined
): Record<string, number | null> | undefined {
    const displayRowHeights = Object.fromEntries(
        Object.entries(rowHeights ?? {}).flatMap(([actualRowNumber, rowHeight]) => {
            const displayRowNumber = actualToDisplayRowNumbers[actualRowNumber];
            if (!Number.isInteger(displayRowNumber)) {
                return [];
            }

            return [[String(displayRowNumber), rowHeight] as const];
        })
    );

    return Object.keys(displayRowHeights).length > 0 ? displayRowHeights : undefined;
}

function createBaseVisibleRowLayout(currentModel: EditorRenderModel): {
    visibleRowResult: {
        visibleRows: number[];
        hiddenRows: number[];
    };
    rowLayout: PixelRowLayout;
} {
    const visibleRowResult = getVisibleActualRowsForModel(currentModel);
    const actualToDisplayRowNumbers = Object.fromEntries(
        visibleRowResult.visibleRows.map((actualRowNumber, index) => [
            String(actualRowNumber),
            index + 1,
        ])
    );
    return {
        visibleRowResult,
        rowLayout: createEditorPixelRowLayout({
            rowCount: visibleRowResult.visibleRows.length,
            rowHeights: mapActualRowHeightsToDisplayRowHeights(
                actualToDisplayRowNumbers,
                currentModel.activeSheet.rowHeights
            ),
        }),
    };
}

function createDisplayGridRowState(
    currentModel: EditorRenderModel,
    viewportHeight: number,
    viewportWidth: number
): {
    displayGrid: {
        rowCount: number;
        columnCount: number;
        rowHeaderWidth: number;
    };
    rowState: EditorDisplayRowState;
} {
    const sheetColumnLayout = getSheetColumnLayout(currentModel);
    const { visibleRowResult, rowLayout: baseVisibleRowLayout } =
        createBaseVisibleRowLayout(currentModel);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: visibleRowResult.visibleRows.length,
        columnCount: currentModel.activeSheet.columnCount,
        rowHeaderLabelCount: getFilterAwareRowHeaderLabelCount(
            currentModel,
            visibleRowResult.visibleRows.length
        ),
        viewportHeight,
        viewportWidth,
        rowLayout: baseVisibleRowLayout,
        columnLayout: sheetColumnLayout,
    });

    return {
        displayGrid,
        rowState: createEditorDisplayRowState(
            currentModel,
            visibleRowResult.visibleRows,
            visibleRowResult.hiddenRows,
            displayGrid.rowCount
        ),
    };
}

function createEditorVirtualGridMetrics(
    currentModel: EditorRenderModel,
    sheetColumnLayout: PixelColumnLayout,
    visibleRowResult: {
        visibleRows: number[];
        hiddenRows: number[];
    },
    baseVisibleRowLayout: PixelRowLayout,
    element: HTMLElement | null,
    fallbackScrollState?: ScrollState | null
): EditorVirtualGridMetrics {
    const viewportHeight = element?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT;
    const viewportWidth = element?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH;
    const scrollTop = element?.scrollTop ?? fallbackScrollState?.top ?? 0;
    const scrollLeft = element?.scrollLeft ?? fallbackScrollState?.left ?? 0;
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: visibleRowResult.visibleRows.length,
        columnCount: currentModel.activeSheet.columnCount,
        rowHeaderLabelCount: getFilterAwareRowHeaderLabelCount(
            currentModel,
            visibleRowResult.visibleRows.length
        ),
        viewportHeight,
        viewportWidth,
        rowLayout: baseVisibleRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const rowState = createEditorDisplayRowState(
        currentModel,
        visibleRowResult.visibleRows,
        visibleRowResult.hiddenRows,
        displayGrid.rowCount
    );
    const displayRowLayout = createEditorPixelRowLayout({
        rowCount: rowState.actualRowNumbers.length,
        rowHeights: mapActualRowHeightsToDisplayRowHeights(
            rowState.actualToDisplayRowNumbers,
            currentModel.activeSheet.rowHeights
        ),
    });
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: rowState.actualRowNumbers.length,
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
            rowLayout: displayRowLayout,
            columnLayout: sheetColumnLayout,
        });
    const rowWindow = createEditorRowWindow({
        rowLayout: displayRowLayout,
        totalRows: displayGrid.rowCount,
        frozenRowCount,
        scrollTop,
        viewportHeight,
    });
    const columnWindow = createEditorColumnWindow({
        columnLayout: displayColumnLayout,
        frozenColumnCount,
        scrollLeft,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });
    const contentSize = getEditorContentSize({
        rowCount: displayGrid.rowCount,
        rowLayout: displayRowLayout,
        columnLayout: displayColumnLayout,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
    });

    return {
        rowLayout: displayRowLayout,
        columnLayout: displayColumnLayout,
        rowState,
        scrollTop,
        scrollLeft,
        viewportHeight,
        viewportWidth,
        rowHeaderWidth: displayGrid.rowHeaderWidth,
        frozenRowCount,
        frozenColumnCount,
        rowNumbers: rowWindow.rowNumbers
            .map((displayRowNumber) => getActualRowNumberAtDisplayRow(rowState, displayRowNumber))
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        columnNumbers: columnWindow.columnNumbers,
        contentWidth: contentSize.width,
        contentHeight: contentSize.height,
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount)
            .map((displayRowNumber) => getActualRowNumberAtDisplayRow(rowState, displayRowNumber))
            .filter((rowNumber): rowNumber is number => rowNumber !== null),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        stickyTopHeight:
            EDITOR_VIRTUAL_HEADER_HEIGHT +
            getEditorFrozenRowsHeight(displayRowLayout, frozenRowCount),
        stickyLeftWidth:
            displayGrid.rowHeaderWidth +
            getEditorFrozenColumnsWidth(displayColumnLayout, frozenColumnCount),
    };
}

function useEditorVirtualGrid(
    currentModel: EditorRenderModel,
    initialScrollState: ScrollState | null,
    revision: number,
    maximumDigitWidth: number
): {
    viewportRef: React.RefObject<HTMLDivElement | null>;
    metrics: EditorVirtualGridMetrics;
    handleScroll(event: React.UIEvent<HTMLDivElement>): void;
} {
    const viewportRef = React.useRef<HTMLDivElement | null>(null);
    const scrollFrameRef = React.useRef(0);
    const latestScrollElementRef = React.useRef<HTMLDivElement | null>(null);
    const activeResizePreviewPixelWidth =
        columnResizeState?.sheetKey === currentModel.activeSheet.key
            ? columnResizeState.previewPixelWidth
            : null;
    const activeResizeColumnNumber =
        columnResizeState?.sheetKey === currentModel.activeSheet.key
            ? columnResizeState.columnNumber
            : null;
    const { visibleRowResult, rowLayout: baseVisibleRowLayout } = React.useMemo(
        () => createBaseVisibleRowLayout(currentModel),
        [
            currentModel.activeSheet.key,
            currentModel.activeSheet.rowCount,
            currentModel.activeSheet.columnCount,
            currentModel.activeSheet.rowHeights,
            revision,
        ]
    );
    const sheetColumnLayout = React.useMemo(
        () =>
            createEditorPixelColumnLayout({
                columnCount: currentModel.activeSheet.columnCount,
                columnWidths: getEffectiveSheetColumnWidths(currentModel, maximumDigitWidth),
                maximumDigitWidth,
            }),
        [
            currentModel.activeSheet.columnCount,
            currentModel.activeSheet.columnWidths,
            currentModel.activeSheet.key,
            maximumDigitWidth,
            activeResizeColumnNumber,
            activeResizePreviewPixelWidth,
        ]
    );
    const [metrics, setMetrics] = React.useState<EditorVirtualGridMetrics>(() =>
        createEditorVirtualGridMetrics(
            currentModel,
            sheetColumnLayout,
            visibleRowResult,
            baseVisibleRowLayout,
            null,
            initialScrollState
        )
    );
    const syncMetrics = React.useEffectEvent(
        (
            element: HTMLDivElement | null,
            fallbackScrollState?: ScrollState | null,
            { force = false }: { force?: boolean } = {}
        ) => {
            const nextMetrics = createEditorVirtualGridMetrics(
                currentModel,
                sheetColumnLayout,
                visibleRowResult,
                baseVisibleRowLayout,
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
        latestVirtualGridCache = {
            model: currentModel,
            metrics,
        };

        return () => {
            if (
                latestVirtualGridCache?.model === currentModel &&
                latestVirtualGridCache.metrics === metrics
            ) {
                latestVirtualGridCache = null;
            }
        };
    }, [currentModel, metrics]);

    React.useLayoutEffect(() => {
        const element = viewportRef.current;
        if (!element) {
            syncMetrics(null, initialScrollState, { force: true });
            return;
        }

        const displayGrid = getEditorDisplayGridDimensions({
            rowCount: visibleRowResult.visibleRows.length,
            columnCount: currentModel.activeSheet.columnCount,
            rowHeaderLabelCount: getFilterAwareRowHeaderLabelCount(
                currentModel,
                visibleRowResult.visibleRows.length
            ),
            viewportHeight: element.clientHeight,
            viewportWidth: element.clientWidth,
            rowLayout: baseVisibleRowLayout,
            columnLayout: sheetColumnLayout,
        });
        const displayRowState = createEditorDisplayRowState(
            currentModel,
            visibleRowResult.visibleRows,
            visibleRowResult.hiddenRows,
            displayGrid.rowCount
        );
        const displayRowLayout = createEditorPixelRowLayout({
            rowCount: displayRowState.actualRowNumbers.length,
            rowHeights: mapActualRowHeightsToDisplayRowHeights(
                displayRowState.actualToDisplayRowNumbers,
                currentModel.activeSheet.rowHeights
            ),
        });
        const displayColumnLayout = getEditorDisplayColumnLayout(
            sheetColumnLayout,
            displayGrid.columnCount
        );
        const contentSize = getEditorContentSize({
            rowCount: displayGrid.rowCount,
            rowLayout: displayRowLayout,
            columnLayout: displayColumnLayout,
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
        currentModel.activeSheet.columnWidths,
        currentModel.activeSheet.rowHeights,
        currentModel.activeSheet.freezePane?.rowCount ?? 0,
        currentModel.activeSheet.freezePane?.columnCount ?? 0,
        activeResizeColumnNumber,
        activeResizePreviewPixelWidth,
        visibleRowResult.visibleRows.join(","),
        initialScrollState?.top ?? 0,
        initialScrollState?.left ?? 0,
        maximumDigitWidth,
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

const EditorColumnHeaderCell = React.memo(
    function EditorColumnHeaderCell({
        sheetKey,
        label,
        columnNumber,
        width,
        canResize,
        isResizing,
        hasPending,
        isActive,
        top,
        left,
    }: {
        sheetKey: string;
        label: string;
        columnNumber: number;
        width: number;
        canResize: boolean;
        isResizing: boolean;
        hasPending: boolean;
        isActive: boolean;
        top: number;
        left: number;
    }): React.ReactElement {
        return (
            <div
                className={classNames([
                    "editor-grid__item",
                    "editor-grid__item--header",
                    "grid__column",
                    hasPending && "grid__column--diff",
                    hasPending && "grid__column--pending",
                    isActive && "grid__column--active",
                    isResizing && "grid__column--resizing",
                ])}
                data-column-number={columnNumber}
                data-role="grid-column-header"
                style={
                    {
                        ...getEditorGridItemStyle({
                            top,
                            left,
                            width,
                            height: EDITOR_VIRTUAL_HEADER_HEIGHT,
                        }),
                        minWidth: `${width}px`,
                        maxWidth: `${width}px`,
                        "--grid-column-max-width": `${width}px`,
                    } as React.CSSProperties
                }
                onPointerDown={(event) => {
                    if (event.button !== 0) {
                        return;
                    }

                    closeContextMenu({ refresh: false });
                    const anchorColumnNumber =
                        event.shiftKey && getSelectionExtendAnchorColumnNumber() !== null
                            ? getSelectionExtendAnchorColumnNumber()!
                            : columnNumber;
                    startColumnSelectionDrag(event.pointerId, anchorColumnNumber);
                    setColumnSelectionLocal(columnNumber, {
                        anchorColumnNumber,
                        syncHost: false,
                    });
                }}
                onClick={(event) => {
                    if (suppressNextColumnHeaderClick) {
                        suppressNextColumnHeaderClick = false;
                        event.preventDefault();
                        return;
                    }

                    closeContextMenu({ refresh: false });
                    if (event.shiftKey && getSelectionExtendAnchorColumnNumber() !== null) {
                        setColumnSelectionLocal(columnNumber, {
                            anchorColumnNumber: getSelectionExtendAnchorColumnNumber()!,
                            syncHost: true,
                        });
                        return;
                    }

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
                {canResize ? (
                    <span
                        aria-hidden
                        className={classNames([
                            "grid__column-resize-handle",
                            isResizing && "is-active",
                        ])}
                        data-role="grid-column-resize-handle"
                        onPointerDown={(event) => {
                            if (event.button !== 0) {
                                return;
                            }

                            event.preventDefault();
                            event.stopPropagation();
                            beginColumnResize(
                                event.pointerId,
                                sheetKey,
                                columnNumber,
                                width,
                                event.clientX
                            );
                        }}
                    />
                ) : null}
            </div>
        );
    },
    (previous, next) =>
        previous.sheetKey === next.sheetKey &&
        previous.label === next.label &&
        previous.columnNumber === next.columnNumber &&
        previous.width === next.width &&
        previous.canResize === next.canResize &&
        previous.isResizing === next.isResizing &&
        previous.hasPending === next.hasPending &&
        previous.isActive === next.isActive &&
        previous.top === next.top &&
        previous.left === next.left
);

const EditorRowHeaderCell = React.memo(
    function EditorRowHeaderCell({
        rowNumber,
        height,
        hasPending,
        isActive,
        top,
        rowHeaderWidth,
    }: {
        rowNumber: number;
        height: number;
        hasPending: boolean;
        isActive: boolean;
        top: number;
        rowHeaderWidth: number;
    }): React.ReactElement {
        return (
            <div
                className={classNames([
                    "editor-grid__item",
                    "editor-grid__item--row-header",
                    "grid__row-number",
                    hasPending && "grid__row-number--pending",
                    isActive && "grid__row-number--active",
                ])}
                data-role="grid-row-header"
                data-row-number={rowNumber}
                style={getEditorGridItemStyle({
                    top,
                    left: 0,
                    width: rowHeaderWidth,
                    height,
                })}
                onPointerDown={(event) => {
                    if (event.button !== 0) {
                        return;
                    }

                    closeContextMenu({ refresh: false });
                    const anchorRowNumber =
                        event.shiftKey && getSelectionExtendAnchorRowNumber() !== null
                            ? getSelectionExtendAnchorRowNumber()!
                            : rowNumber;
                    startRowSelectionDrag(event.pointerId, anchorRowNumber);
                    setRowSelectionLocal(rowNumber, {
                        anchorRowNumber,
                        syncHost: false,
                    });
                }}
                onClick={(event) => {
                    if (suppressNextRowHeaderClick) {
                        suppressNextRowHeaderClick = false;
                        event.preventDefault();
                        return;
                    }

                    closeContextMenu({ refresh: false });
                    if (event.shiftKey && getSelectionExtendAnchorRowNumber() !== null) {
                        setRowSelectionLocal(rowNumber, {
                            anchorRowNumber: getSelectionExtendAnchorRowNumber()!,
                            syncHost: true,
                        });
                        return;
                    }

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
        previous.height === next.height &&
        previous.hasPending === next.hasPending &&
        previous.isActive === next.isActive &&
        previous.top === next.top &&
        previous.rowHeaderWidth === next.rowHeaderWidth
);

const FILL_HANDLE_SIZE = 8;
const FILL_HANDLE_HALF_SIZE = FILL_HANDLE_SIZE / 2;

function EditorFillHandle({
    rowNumber,
    columnNumber,
    rowHeaderWidth,
    rowLayout,
    rowState,
    columnLayout,
    isActive,
    onPointerDown,
    onDoubleClick,
}: {
    rowNumber: number;
    columnNumber: number;
    rowHeaderWidth: number;
    rowLayout: PixelRowLayout;
    rowState: EditorDisplayRowState;
    columnLayout: PixelColumnLayout;
    isActive: boolean;
    onPointerDown(event: React.PointerEvent<HTMLSpanElement>): void;
    onDoubleClick(event: React.MouseEvent<HTMLSpanElement>): void;
}): React.ReactElement {
    const displayRowNumber = getDisplayRowNumber(rowState, rowNumber) ?? 1;
    return (
        <span
            aria-hidden
            className={classNames(["editor-grid__fill-handle", isActive && "is-active"])}
            style={getEditorGridItemStyle({
                top:
                    getEditorGridTop(rowLayout, displayRowNumber) +
                    getEditorRowHeight(rowLayout, displayRowNumber) -
                    FILL_HANDLE_HALF_SIZE,
                left:
                    getEditorGridLeft(rowHeaderWidth, columnLayout, columnNumber) +
                    getEditorColumnWidth(columnLayout, columnNumber) -
                    FILL_HANDLE_HALF_SIZE,
                width: FILL_HANDLE_SIZE,
                height: FILL_HANDLE_SIZE,
            })}
            onPointerDown={onPointerDown}
            onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
            }}
            onDoubleClick={onDoubleClick}
        />
    );
}

const EditorVirtualCell = React.memo(
    function EditorVirtualCell({
        currentModel,
        rowNumber,
        columnNumber,
        width,
        height,
        top,
        left,
        isSelectionFocus,
        activeEditingCell,
        pendingEdit,
        fillSourceRange,
        fillPreviewRange,
        activeFilterState,
        isFilterMenuOpen,
        contentWidthPx,
        spillsIntoNextCells,
    }: {
        currentModel: EditorRenderModel;
        rowNumber: number;
        columnNumber: number;
        width: number;
        height: number;
        top: number;
        left: number;
        isSelectionFocus: boolean;
        activeEditingCell: EditingCell | null;
        pendingEdit: PendingEdit | undefined;
        fillSourceRange: CellRange | null;
        fillPreviewRange: CellRange | null;
        activeFilterState: EditorSheetFilterState | null;
        isFilterMenuOpen: boolean;
        contentWidthPx: number;
        spillsIntoNextCells: boolean;
    }): React.ReactElement {
        const cell = getCellView(rowNumber, columnNumber, currentModel) ?? {
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
        const isEditing =
            activeEditingCell?.rowNumber === rowNumber &&
            activeEditingCell.columnNumber === columnNumber;
        const isFillPreview = Boolean(
            fillSourceRange &&
            isCellWithinFillPreviewArea(fillSourceRange, fillPreviewRange, rowNumber, columnNumber)
        );
        const isFilterHeader = isEditorFilterHeaderCell(activeFilterState, rowNumber, columnNumber);
        const isColumnFilterActive =
            Boolean(activeFilterState?.includedValuesByColumn[String(columnNumber)]) ||
            activeFilterState?.sort?.columnNumber === columnNumber;
        const contentAlignmentStyle = getCellContentAlignmentStyle(
            getEffectiveCellAlignment(rowNumber, columnNumber, currentModel)
        );
        const shouldSpillIntoNextCells = spillsIntoNextCells && !isEditing;
        const { contentMaxHeightPx, visibleLineCount } = getGridCellTextLayoutMetrics(height);

        return (
            <div
                aria-selected={isSelectionFocus}
                className={classNames([
                    "editor-grid__item",
                    "grid__cell",
                    !editable && "grid__cell--locked",
                    isFilterHeader && "grid__cell--filter-header",
                    pendingEdit && "grid__cell--pending",
                    isFillPreview && "grid__cell--fill-preview",
                    isEditing && "grid__cell--editing",
                    shouldSpillIntoNextCells && "grid__cell--overflow-spill",
                ])}
                data-column-number={columnNumber}
                data-editable={editable}
                data-role="grid-cell"
                data-row-number={rowNumber}
                style={
                    {
                        ...getEditorGridItemStyle({
                            top,
                            left,
                            width,
                            height,
                        }),
                        minWidth: `${width}px`,
                        maxWidth: `${width}px`,
                        "--grid-column-max-width": `${width}px`,
                        "--grid-cell-content-max-height": `${contentMaxHeightPx}px`,
                        "--grid-cell-line-clamp": String(visibleLineCount),
                        ...(shouldSpillIntoNextCells
                            ? {
                                  "--grid-cell-display-max-width": `${contentWidthPx}px`,
                              }
                            : {}),
                    } as React.CSSProperties
                }
                title={getCellTooltip(cell.address, value, formula)}
                onPointerDown={(event) => {
                    if (event.button !== 0) {
                        return;
                    }

                    closeContextMenu({ refresh: false });
                    const anchorCell =
                        event.shiftKey && getSelectionExtendAnchorCell()
                            ? getSelectionExtendAnchorCell()!
                            : { rowNumber, columnNumber };
                    startSelectionDrag(event.pointerId, anchorCell);
                    setSelectedCellLocal(
                        { rowNumber, columnNumber },
                        {
                            syncHost: false,
                            anchorCell,
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
                                event.shiftKey && getSelectionExtendAnchorCell()
                                    ? getSelectionExtendAnchorCell()
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
                <div className="grid__cell-content" style={contentAlignmentStyle}>
                    {isEditing && activeEditingCell?.sheetKey === currentModel.activeSheet.key ? (
                        <CellEditor edit={activeEditingCell} />
                    ) : (
                        <CellValue value={value} />
                    )}
                </div>
                <CellFormulaBadge formula={formula} />
                {isFilterHeader ? (
                    <button
                        className={classNames([
                            "grid__cell-filterButton",
                            isColumnFilterActive && "is-active",
                            isFilterMenuOpen && "is-open",
                        ])}
                        data-role="filter-trigger"
                        type="button"
                        onPointerDown={(event) => {
                            event.preventDefault();
                            event.stopPropagation();
                        }}
                        onClick={(event) => {
                            event.preventDefault();
                            event.stopPropagation();
                            openFilterMenu(rowNumber, columnNumber, event.currentTarget);
                        }}
                    >
                        <span className="codicon codicon-chevron-down" aria-hidden />
                    </button>
                ) : null}
            </div>
        );
    },
    (previous, next) =>
        previous.currentModel.activeSheet.key === next.currentModel.activeSheet.key &&
        previous.currentModel.activeSheet.cells === next.currentModel.activeSheet.cells &&
        previous.currentModel.activeSheet.cellAlignments ===
            next.currentModel.activeSheet.cellAlignments &&
        previous.currentModel.activeSheet.rowAlignments ===
            next.currentModel.activeSheet.rowAlignments &&
        previous.currentModel.activeSheet.columnAlignments ===
            next.currentModel.activeSheet.columnAlignments &&
        previous.currentModel.canEdit === next.currentModel.canEdit &&
        previous.rowNumber === next.rowNumber &&
        previous.columnNumber === next.columnNumber &&
        previous.width === next.width &&
        previous.height === next.height &&
        previous.top === next.top &&
        previous.left === next.left &&
        previous.isSelectionFocus === next.isSelectionFocus &&
        areEditingCellsEqual(previous.activeEditingCell, next.activeEditingCell) &&
        arePendingEditsEqual(previous.pendingEdit, next.pendingEdit) &&
        areSelectionRangesEqual(previous.fillSourceRange, next.fillSourceRange) &&
        areSelectionRangesEqual(previous.fillPreviewRange, next.fillPreviewRange) &&
        previous.activeFilterState === next.activeFilterState &&
        previous.isFilterMenuOpen === next.isFilterMenuOpen &&
        previous.contentWidthPx === next.contentWidthPx &&
        previous.spillsIntoNextCells === next.spillsIntoNextCells
);

function EditorVirtualGrid({
    currentModel,
    pendingSummary,
    view,
    maximumDigitWidth,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    view: Extract<ViewState, { kind: "app" }>;
    maximumDigitWidth: number;
}): React.ReactElement {
    const { viewportRef, metrics, handleScroll } = useEditorVirtualGrid(
        currentModel,
        view.scrollState,
        view.revision,
        maximumDigitWidth
    );
    const selectionRange = getSelectionRange();
    const fillSourceRange = fillDragState?.sourceRange ?? selectionRange;
    const fillPreviewRange = fillDragState?.previewRange ?? null;
    const fillHandleCell = getFillHandleCell(fillSourceRange);
    const fillHandleLayer = getGridLayerForCell(metrics, fillHandleCell);
    const currentSelection = selectedCell;
    const activeHighlightCell = getActiveHighlightCell();
    const activeRowNumber = activeHighlightCell?.rowNumber ?? null;
    const activeColumnNumber = activeHighlightCell?.columnNumber ?? null;
    const activeEditingCell = editingCell;
    const activeFilterState = getActiveSheetFilterState(currentModel);
    const topOverlayRef = React.useRef<HTMLDivElement | null>(null);
    const leftOverlayRef = React.useRef<HTMLDivElement | null>(null);
    const cornerOverlayRef = React.useRef<HTMLDivElement | null>(null);
    const handleOverlayWheel = React.useEffectEvent((event: WheelEvent) => {
        forwardVirtualGridWheel(event);
    });
    const isSearchFocusedSelection =
        searchFeedback?.status === "matched" ||
        searchFeedback?.status === "replaced" ||
        searchFeedback?.status === "no-change";
    const showPrimarySelectionFrame =
        Boolean(currentSelection) &&
        (!hasExpandedSelectionRange(selectionRange) || isSearchFocusedSelection);
    const bodyItems: React.ReactElement[] = [];
    const topItems: React.ReactElement[] = [];
    const leftItems: React.ReactElement[] = [];
    const cornerItems: React.ReactElement[] = [
        <EditorCornerHeader key="corner" rowHeaderWidth={metrics.rowHeaderWidth} />,
    ];
    const createGridCellItem = (
        keyPrefix: "body" | "left" | "top" | "corner",
        rowNumber: number,
        columnNumber: number,
        visibleColumnNumbers: readonly number[],
        visibleColumnIndex: number
    ): React.ReactElement => {
        const top = getEditorGridTopForActualRow(metrics, rowNumber) ?? 0;
        const left = getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, columnNumber);
        const width = getEditorColumnWidth(metrics.columnLayout, columnNumber);
        const height =
            getEditorGridRowHeightForActualRow(metrics, rowNumber) ??
            metrics.rowLayout.fallbackPixelHeight;
        const key = `${keyPrefix}:${rowNumber}:${columnNumber}`;

        const pendingEdit = pendingEdits.get(
            getPendingEditKey(currentModel.activeSheet.key, rowNumber, columnNumber)
        );
        const displayedValue =
            pendingEdit?.value ??
            getCellSnapshot(rowNumber, columnNumber, currentModel)?.displayValue ??
            "";
        const overflowMetrics = getCellOverflowMetrics({
            value: displayedValue,
            alignment: getEffectiveCellAlignment(rowNumber, columnNumber, currentModel),
            baseColumnWidth: width,
            visibleColumnNumbers,
            visibleColumnIndex,
            getColumnWidth: (nextColumnNumber) =>
                getEditorColumnWidth(metrics.columnLayout, nextColumnNumber),
            getTrailingCellState: (nextColumnNumber) => {
                const nextPendingEdit = pendingEdits.get(
                    getPendingEditKey(currentModel.activeSheet.key, rowNumber, nextColumnNumber)
                );
                const nextCell = getCellSnapshot(rowNumber, nextColumnNumber, currentModel);
                return {
                    value: nextPendingEdit?.value ?? nextCell?.displayValue ?? "",
                    formula: nextPendingEdit ? null : (nextCell?.formula ?? null),
                    blocksOverflow: isEditorFilterHeaderCell(
                        activeFilterState,
                        rowNumber,
                        nextColumnNumber
                    ),
                };
            },
        });
        return (
            <EditorVirtualCell
                key={key}
                currentModel={currentModel}
                rowNumber={rowNumber}
                columnNumber={columnNumber}
                width={width}
                height={height}
                top={top}
                left={left}
                isSelectionFocus={
                    Boolean(
                        currentSelection &&
                            currentSelection.rowNumber === rowNumber &&
                            currentSelection.columnNumber === columnNumber
                    )
                }
                activeEditingCell={activeEditingCell}
                pendingEdit={pendingEdit}
                fillSourceRange={fillSourceRange}
                fillPreviewRange={fillPreviewRange}
                activeFilterState={activeFilterState}
                isFilterMenuOpen={
                    filterMenuState?.sheetKey === currentModel.activeSheet.key &&
                    filterMenuState.columnNumber === columnNumber
                }
                contentWidthPx={overflowMetrics.contentWidthPx}
                spillsIntoNextCells={overflowMetrics.spillsIntoNextCells}
            />
        );
    };

    for (const columnNumber of metrics.columnNumbers) {
        topItems.push(
            <EditorColumnHeaderCell
                key={`top:header:${columnNumber}`}
                sheetKey={currentModel.activeSheet.key}
                label={
                    currentModel.activeSheet.columns[columnNumber - 1] ??
                    getColumnLabel(columnNumber)
                }
                columnNumber={columnNumber}
                width={getEditorColumnWidth(metrics.columnLayout, columnNumber)}
                canResize={currentModel.canEdit}
                isResizing={
                    columnResizeState?.sheetKey === currentModel.activeSheet.key &&
                    columnResizeState.columnNumber === columnNumber
                }
                hasPending={pendingSummary.columns.has(columnNumber)}
                isActive={
                    isActiveHighlightColumn(activeColumnNumber, columnNumber) ||
                    isColumnHeaderWithinSelectionRange(
                        selectionRange,
                        columnNumber,
                        metrics.rowState.actualRowNumbers[
                            metrics.rowState.actualRowNumbers.length - 1
                        ] ?? 1
                    )
                }
                top={0}
                left={getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, columnNumber)}
            />
        );
    }

    for (const columnNumber of metrics.frozenColumnNumbers) {
        cornerItems.push(
            <EditorColumnHeaderCell
                key={`corner:header:${columnNumber}`}
                sheetKey={currentModel.activeSheet.key}
                label={
                    currentModel.activeSheet.columns[columnNumber - 1] ??
                    getColumnLabel(columnNumber)
                }
                columnNumber={columnNumber}
                width={getEditorColumnWidth(metrics.columnLayout, columnNumber)}
                canResize={currentModel.canEdit}
                isResizing={
                    columnResizeState?.sheetKey === currentModel.activeSheet.key &&
                    columnResizeState.columnNumber === columnNumber
                }
                hasPending={pendingSummary.columns.has(columnNumber)}
                isActive={
                    isActiveHighlightColumn(activeColumnNumber, columnNumber) ||
                    isColumnHeaderWithinSelectionRange(
                        selectionRange,
                        columnNumber,
                        metrics.rowState.actualRowNumbers[
                            metrics.rowState.actualRowNumbers.length - 1
                        ] ?? 1
                    )
                }
                top={0}
                left={getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, columnNumber)}
            />
        );
    }

    for (const rowNumber of metrics.rowNumbers) {
        leftItems.push(
            <EditorRowHeaderCell
                key={`left:row:${rowNumber}`}
                rowNumber={rowNumber}
                height={getEditorGridRowHeightForActualRow(metrics, rowNumber) ?? 0}
                hasPending={pendingSummary.rows.has(rowNumber)}
                isActive={
                    isActiveHighlightRow(activeRowNumber, rowNumber) ||
                    isRowHeaderWithinSelectionRange(
                        selectionRange,
                        rowNumber,
                        metrics.columnLayout.totalColumnCount
                    )
                }
                top={getEditorGridTopForActualRow(metrics, rowNumber) ?? 0}
                rowHeaderWidth={metrics.rowHeaderWidth}
            />
        );

        for (const [visibleColumnIndex, columnNumber] of metrics.columnNumbers.entries()) {
            bodyItems.push(
                createGridCellItem(
                    "body",
                    rowNumber,
                    columnNumber,
                    metrics.columnNumbers,
                    visibleColumnIndex
                )
            );
        }

        for (const [visibleColumnIndex, columnNumber] of metrics.frozenColumnNumbers.entries()) {
            leftItems.push(
                createGridCellItem(
                    "left",
                    rowNumber,
                    columnNumber,
                    metrics.frozenColumnNumbers,
                    visibleColumnIndex
                )
            );
        }
    }

    for (const rowNumber of metrics.frozenRowNumbers) {
        cornerItems.push(
            <EditorRowHeaderCell
                key={`corner:row:${rowNumber}`}
                rowNumber={rowNumber}
                height={getEditorGridRowHeightForActualRow(metrics, rowNumber) ?? 0}
                hasPending={pendingSummary.rows.has(rowNumber)}
                isActive={
                    isActiveHighlightRow(activeRowNumber, rowNumber) ||
                    isRowHeaderWithinSelectionRange(
                        selectionRange,
                        rowNumber,
                        metrics.columnLayout.totalColumnCount
                    )
                }
                top={getEditorGridTopForActualRow(metrics, rowNumber) ?? 0}
                rowHeaderWidth={metrics.rowHeaderWidth}
            />
        );

        for (const [visibleColumnIndex, columnNumber] of metrics.columnNumbers.entries()) {
            topItems.push(
                createGridCellItem(
                    "top",
                    rowNumber,
                    columnNumber,
                    metrics.columnNumbers,
                    visibleColumnIndex
                )
            );
        }

        for (const [visibleColumnIndex, columnNumber] of metrics.frozenColumnNumbers.entries()) {
            cornerItems.push(
                createGridCellItem(
                    "corner",
                    rowNumber,
                    columnNumber,
                    metrics.frozenColumnNumbers,
                    visibleColumnIndex
                )
            );
        }
    }

    const fillHandle =
        fillHandleCell && fillHandleLayer ? (
            <EditorFillHandle
                key={`fill-handle:${fillHandleCell.rowNumber}:${fillHandleCell.columnNumber}`}
                rowNumber={fillHandleCell.rowNumber}
                columnNumber={fillHandleCell.columnNumber}
                rowHeaderWidth={metrics.rowHeaderWidth}
                rowLayout={metrics.rowLayout}
                rowState={metrics.rowState}
                columnLayout={metrics.columnLayout}
                isActive={Boolean(fillPreviewRange)}
                onPointerDown={(event) => {
                    if (event.button !== 0 || !fillSourceRange) {
                        return;
                    }

                    event.preventDefault();
                    event.stopPropagation();
                    closeContextMenu({ refresh: false });
                    startFillDrag(event.pointerId, fillSourceRange);
                }}
                onDoubleClick={(event) => {
                    if (!fillSourceRange) {
                        return;
                    }

                    event.preventDefault();
                    event.stopPropagation();
                    handleFillHandleAutoFill(fillSourceRange);
                }}
            />
        ) : null;

    if (fillHandle && fillHandleLayer === "body") {
        bodyItems.push(fillHandle);
    } else if (fillHandle && fillHandleLayer === "top") {
        topItems.push(fillHandle);
    } else if (fillHandle && fillHandleLayer === "left") {
        leftItems.push(fillHandle);
    } else if (fillHandle && fillHandleLayer === "corner") {
        cornerItems.push(fillHandle);
    }

    React.useLayoutEffect(() => {
        if (!view.revealSelection) {
            return;
        }

        revealSelectedCell();
    }, [view.revision, view.revealSelection]);

    React.useEffect(() => {
        if (!IS_DEBUG_MODE) {
            return;
        }

        setDebugRenderStats({
            renderedRowCount: metrics.frozenRowNumbers.length + metrics.rowNumbers.length,
            renderedColumnCount: metrics.frozenColumnNumbers.length + metrics.columnNumbers.length,
        });
    }, [
        metrics.frozenColumnNumbers.length,
        metrics.frozenRowNumbers.length,
        metrics.columnNumbers.length,
        metrics.rowNumbers.length,
    ]);

    React.useEffect(() => {
        if (!IS_DEBUG_MODE) {
            return;
        }

        return () => {
            setDebugRenderStats(null);
        };
    }, []);

    React.useEffect(() => {
        const overlays = [
            topOverlayRef.current,
            leftOverlayRef.current,
            cornerOverlayRef.current,
        ].filter((overlay): overlay is HTMLDivElement => Boolean(overlay));
        if (overlays.length === 0) {
            return;
        }

        const listener = (event: WheelEvent) => {
            handleOverlayWheel(event);
        };
        const options: AddEventListenerOptions = { passive: false };

        for (const overlay of overlays) {
            overlay.addEventListener("wheel", listener, options);
        }

        return () => {
            for (const overlay of overlays) {
                overlay.removeEventListener("wheel", listener);
            }
        };
    }, [handleOverlayWheel]);

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
                    <div className="editor-grid__layer editor-grid__layer--body">
                        {bodyItems}
                        <EditorSelectionOverlay
                            metrics={metrics}
                            layer="body"
                            selectionRange={selectionRange}
                            activeRowNumber={activeRowNumber}
                            activeColumnNumber={activeColumnNumber}
                            currentSelection={currentSelection}
                            showPrimarySelectionFrame={showPrimarySelectionFrame}
                        />
                    </div>
                    <div
                        ref={topOverlayRef}
                        className="editor-grid__overlay editor-grid__overlay--top"
                        style={{
                            width: `${metrics.contentWidth}px`,
                            height: `${metrics.stickyTopHeight}px`,
                        }}
                    >
                        <div
                            className="editor-grid__track editor-grid__track--x"
                            style={{
                                width: `${metrics.contentWidth}px`,
                                height: `${metrics.stickyTopHeight}px`,
                            }}
                        >
                            {topItems}
                            <EditorSelectionOverlay
                                metrics={metrics}
                                layer="top"
                                selectionRange={selectionRange}
                                activeRowNumber={activeRowNumber}
                                activeColumnNumber={activeColumnNumber}
                                currentSelection={currentSelection}
                                showPrimarySelectionFrame={showPrimarySelectionFrame}
                            />
                        </div>
                    </div>
                    <div
                        ref={leftOverlayRef}
                        className="editor-grid__overlay editor-grid__overlay--left"
                        style={{
                            width: `${metrics.stickyLeftWidth}px`,
                            height: `${metrics.contentHeight}px`,
                        }}
                    >
                        <div
                            className="editor-grid__track editor-grid__track--y"
                            style={{
                                width: `${metrics.stickyLeftWidth}px`,
                                height: `${metrics.contentHeight}px`,
                            }}
                        >
                            {leftItems}
                            <EditorSelectionOverlay
                                metrics={metrics}
                                layer="left"
                                selectionRange={selectionRange}
                                activeRowNumber={activeRowNumber}
                                activeColumnNumber={activeColumnNumber}
                                currentSelection={currentSelection}
                                showPrimarySelectionFrame={showPrimarySelectionFrame}
                            />
                        </div>
                    </div>
                    <div
                        ref={cornerOverlayRef}
                        className="editor-grid__overlay editor-grid__overlay--corner"
                        style={{
                            width: `${metrics.stickyLeftWidth}px`,
                            height: `${metrics.stickyTopHeight}px`,
                        }}
                    >
                        <div
                            className="editor-grid__track"
                            style={{
                                width: `${metrics.stickyLeftWidth}px`,
                                height: `${metrics.stickyTopHeight}px`,
                            }}
                        >
                            {cornerItems}
                            <EditorSelectionOverlay
                                metrics={metrics}
                                layer="corner"
                                selectionRange={selectionRange}
                                activeRowNumber={activeRowNumber}
                                activeColumnNumber={activeColumnNumber}
                                currentSelection={currentSelection}
                                showPrimarySelectionFrame={showPrimarySelectionFrame}
                            />
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
    maximumDigitWidth,
}: {
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    view: Extract<ViewState, { kind: "app" }>;
    maximumDigitWidth: number;
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
                    maximumDigitWidth={maximumDigitWidth}
                />
            )}
        </section>
    );
}

function closeContextMenuSilently(): void {
    closeContextMenu({ refresh: false });
}

function requestSetSheet(sheetKey: string): void {
    vscode.postMessage({ type: "setSheet", sheetKey });
}

function requestReload(): void {
    vscode.postMessage({ type: "reload" });
}

function applyToolbarAlignment(alignment: EditorAlignmentPatch): void {
    const target = getActiveAlignmentSelectionTarget();
    if (!target) {
        return;
    }

    vscode.postMessage({
        type: "setAlignment",
        target: target.target,
        selection: target.selection,
        alignment,
    });
}

function finishEditingFromToolbar(): void {
    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }
}

function getActiveToolbarCellEditTarget(
    currentModel: EditorRenderModel
): ToolbarCellEditTarget | null {
    if (!selectedCell) {
        return null;
    }

    return {
        sheetKey: currentModel.activeSheet.key,
        rowNumber: selectedCell.rowNumber,
        columnNumber: selectedCell.columnNumber,
    };
}

function EditorApp({ view }: { view: Extract<ViewState, { kind: "app" }> }): React.ReactElement {
    const currentModel = view.model;
    const maximumDigitWidth = useMeasuredMaximumDigitWidth(getEditorAppElement());
    const currentDebugRenderStats = React.useSyncExternalStore(
        subscribeDebugRenderStats,
        getDebugRenderStatsSnapshot,
        getDebugRenderStatsSnapshot
    );
    editorMaximumDigitWidth = maximumDigitWidth;
    const pendingSummary = getPendingSummary(currentModel.activeSheet.key);
    const activeFilterState = getActiveSheetFilterState(currentModel);
    const activeToolbarAlignment = getActiveToolbarAlignment(currentModel);
    const currentSelectionRange = getExpandedSelectionRange();
    const effectiveScope = getEffectiveSearchScope();
    const selectionRange =
        effectiveScope === "selection" ? (searchSelectionRange ?? currentSelectionRange) : null;
    const hasSearchableGrid =
        currentModel.activeSheet.rowCount > 0 && currentModel.activeSheet.columnCount > 0;
    const canReplace = searchQuery.trim().length > 0 && hasSearchableGrid && currentModel.canEdit;
    const hasPendingEdits = pendingEdits.size > 0 || currentModel.hasPendingEdits;
    const canUndo = undoStack.length > 0 || currentModel.canUndoStructuralEdits;
    const canRedo = redoStack.length > 0 || currentModel.canRedoStructuralEdits;
    const viewLocked = hasLockedView(currentModel.activeSheet.freezePane);
    const filterActionLabel = getFilterToolbarActionLabel(currentModel);
    const canToggleFilter = canToggleFilterForCurrentSelection(currentModel);

    return (
        <div
            className={classNames([
                "app",
                "app--editor",
                columnResizeState?.isDragging && "app--column-resizing",
            ])}
            data-role="editor-app"
        >
            <EditorToolbar
                strings={STRINGS}
                currentModel={currentModel}
                isSearchPanelOpen={isSearchPanelOpen}
                filterActionLabel={filterActionLabel}
                isFilterButtonEnabled={canToggleFilter}
                hasActiveFilter={Boolean(activeFilterState)}
                isSaving={isSaving}
                hasPendingEdits={hasPendingEdits}
                canUndo={canUndo}
                canRedo={canRedo}
                viewLocked={viewLocked}
                debugRenderStats={IS_DEBUG_MODE ? currentDebugRenderStats : null}
                getPositionInputValue={getPositionInputValue}
                getCellValueInputValue={getCellValueInputValue}
                getCellValueInputPlaceholder={getCellValueInputPlaceholder}
                activeAlignment={activeToolbarAlignment}
                canEditSelectedCellValue={canEditSelectedCellValue}
                getActiveCellEditTarget={() => getActiveToolbarCellEditTarget(currentModel)}
                onOpenSearch={openSearchPanel}
                onToggleFilter={toggleActiveSheetFilter}
                onUndo={undoPendingEdits}
                onRedo={redoPendingEdits}
                onReload={requestReload}
                onToggleViewLock={toggleViewLock}
                onSave={triggerSave}
                onSubmitGoto={submitGotoSelection}
                onCommitCellValue={commitToolbarCellValue}
                onApplyAlignment={applyToolbarAlignment}
                onFinishGridEdit={finishEditingFromToolbar}
            />
            <SearchPanel
                strings={STRINGS}
                isOpen={isSearchPanelOpen}
                mode={searchMode}
                query={searchQuery}
                replaceValue={replaceValue}
                options={searchOptions}
                feedback={searchFeedback}
                scopeSummary={
                    selectionRange
                        ? getSelectionRangeAddress(selectionRange)
                        : STRINGS.searchScopeSheet
                }
                hasSelectionScope={Boolean(selectionRange)}
                position={searchPanelPosition}
                hasSearchableGrid={hasSearchableGrid}
                canReplace={canReplace}
                isInteractiveTarget={isSearchPanelInteractiveTarget}
                onSyncPosition={syncSearchPanelShellPosition}
                onBeginDrag={(pointerId, clientX, clientY) => {
                    closeContextMenuSilently();
                    beginSearchPanelDrag(pointerId, clientX, clientY);
                }}
                onClose={closeSearchPanel}
                onModeChange={setSearchMode}
                onQueryChange={updateSearchQuery}
                onReplaceValueChange={updateReplaceValue}
                onToggleOption={toggleSearchOption}
                onSubmitSearch={submitSearch}
                onSubmitReplace={submitReplace}
            />
            <EditorFilterMenu
                currentModel={currentModel}
                filterState={activeFilterState}
                menuState={filterMenuState}
            />
            <section className="panes panes--single">
                <EditorPane
                    currentModel={currentModel}
                    pendingSummary={pendingSummary}
                    view={view}
                    maximumDigitWidth={maximumDigitWidth}
                />
            </section>
            <footer className="footer">
                <Tabs
                    strings={STRINGS}
                    currentModel={currentModel}
                    pendingSummary={pendingSummary}
                    onSetSheet={requestSetSheet}
                    onOpenTabContextMenu={openTabContextMenu}
                    onCloseContextMenu={closeContextMenuSilently}
                />
            </footer>
            <TabContextMenu
                strings={STRINGS}
                currentModel={currentModel}
                contextMenu={contextMenu}
                onRequestAddSheet={requestAddSheet}
                onRequestDeleteSheet={requestDeleteSheet}
                onRequestRenameSheet={requestRenameSheet}
                onRequestInsertRow={requestInsertRow}
                onRequestDeleteRow={requestDeleteRow}
                onRequestPromptRowHeight={requestPromptRowHeight}
                onRequestInsertColumn={requestInsertColumn}
                onRequestDeleteColumn={requestDeleteColumn}
                onRequestPromptColumnWidth={requestPromptColumnWidth}
            />
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
        clearFillDragState();
        columnResizeState = null;
        stopSearchPanelDrag();
        sheetFilterStates.clear();
        filterMenuState = null;
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
        clearFillDragState();
        columnResizeState = null;
        stopSearchPanelDrag();
        sheetFilterStates.clear();
        filterMenuState = null;
        isSaving = false;
        renderError(message.message);
        return;
    }

    if (message.type === "searchResult") {
        handleSearchResult(message);
        return;
    }

    if (message.type === "render") {
        clearFillDragState();
        columnResizeState = null;
        model = stabilizeIncomingRenderModel(model, message.payload, {
            canReuseActiveSheetData: Boolean(message.reuseActiveSheetData),
        });
        isSaving = false;
        if (filterMenuState && filterMenuState.sheetKey !== model.activeSheet.key) {
            filterMenuState = null;
        }

        if (message.clearPendingEdits) {
            sheetFilterStates.clear();
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
            updateSelectionControllerState((state) =>
                syncControllerSelectionAnchorToSelectedCell(state)
            );
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

    if (event.key === "Escape" && filterMenuState) {
        event.preventDefault();
        closeFilterMenu();
        return;
    }

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
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
        closeFilterMenu();
        closeContextMenu();
        return;
    }

    if (
        filterMenuState &&
        !target.closest('[data-role="filter-menu"]') &&
        !target.closest('[data-role="filter-trigger"]')
    ) {
        closeFilterMenu();
    }

    if (contextMenu && !target.closest('[data-role="context-menu"]')) {
        closeContextMenu();
    }
});

document.addEventListener("pointermove", (event: PointerEvent) => {
    if (columnResizeState?.isDragging && columnResizeState.pointerId === event.pointerId) {
        if ((event.buttons & 1) === 0) {
            stopColumnResize(event.pointerId, { commit: true });
            return;
        }

        event.preventDefault();
        updateColumnResize(event.clientX);
        return;
    }

    if (searchPanelDragState && searchPanelDragState.pointerId === event.pointerId) {
        event.preventDefault();
        updateSearchPanelDrag(event.clientX, event.clientY);
        return;
    }

    if (fillDragState && fillDragState.pointerId === event.pointerId) {
        if ((event.buttons & 1) === 0) {
            stopFillDrag(event.pointerId, { commit: true, timeStamp: event.timeStamp });
            return;
        }

        const targetElement = document.elementFromPoint(event.clientX, event.clientY);
        if (!targetElement) {
            return;
        }

        const targetCell = getCellPositionFromElement(targetElement);
        if (!targetCell) {
            return;
        }

        event.preventDefault();
        updateFillDrag(targetCell);
        return;
    }

    if (!selectionDragState || selectionDragState.pointerId !== event.pointerId) {
        return;
    }

    if ((event.buttons & 1) === 0) {
        stopSelectionDrag(event.pointerId);
        return;
    }

    const targetElement = document.elementFromPoint(event.clientX, event.clientY);
    if (!targetElement) {
        return;
    }

    if (selectionDragState.kind === "cell") {
        const targetCell = getCellPositionFromElement(targetElement);
        if (!targetCell) {
            return;
        }

        updateCellSelectionDrag(targetCell);
        return;
    }

    if (selectionDragState.kind === "row") {
        const targetRowNumber = getRowNumberFromElement(targetElement);
        if (targetRowNumber === null) {
            return;
        }

        updateRowSelectionDrag(targetRowNumber);
        return;
    }

    const targetColumnNumber = getColumnNumberFromElement(targetElement);
    if (targetColumnNumber === null) {
        return;
    }

    updateColumnSelectionDrag(targetColumnNumber);
});

document.addEventListener("pointerup", (event: PointerEvent) => {
    stopColumnResize(event.pointerId, { commit: true });
    stopSearchPanelDrag(event.pointerId);
    stopFillDrag(event.pointerId, { commit: true, timeStamp: event.timeStamp });
    stopSelectionDrag(event.pointerId);
    scheduleSuppressedSelectionClickReset();
});

document.addEventListener("pointercancel", (event: PointerEvent) => {
    stopColumnResize(event.pointerId);
    stopSearchPanelDrag(event.pointerId);
    stopFillDrag(event.pointerId);
    stopSelectionDrag(event.pointerId);
    scheduleSuppressedSelectionClickReset();
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
