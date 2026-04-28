import * as React from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import type { CellSnapshot, EditorRenderModel } from "../../core/model/types";
import { createCellKey, getColumnLabel } from "../../core/model/cells";
import { formatI18nMessage, RUNTIME_MESSAGES } from "../../i18n/catalog";
import {
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
    getEditorDisplayRowLayout,
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
import {
    MAX_ROW_PIXEL_HEIGHT,
    MIN_ROW_PIXEL_HEIGHT,
    convertPixelsToWorkbookRowHeight,
    stabilizeRowPixelHeight,
    type PixelRowLayout,
} from "../row-layout";
import {
    isSelectionFocusCell,
    shouldResetInvisibleSelectionAnchor,
    shouldSyncLocalSelectionDomFromModelSelection,
    shouldUseLocalSimpleSelectionUpdate,
} from "./editor-selection-render";
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
import { notifyEditorToolbarSync } from "./editor-toolbar-sync";
import { type ToolbarCellEditTarget } from "./editor-toolbar-input";
import type {
    EditorPanelStrings,
    EditorSearchResultMessage,
    EditorSearchScope,
    EditorWebviewMessage,
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

interface ScrollState {
    top: number;
    left: number;
}

type SelectionDragState =
    | {
          kind: "cell";
          anchorCell: CellPosition;
          pointerId: number;
      }
    | {
          kind: "row";
          anchorRowNumber: number;
          pointerId: number;
      }
    | {
          kind: "column";
          anchorColumnNumber: number;
          pointerId: number;
      };

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

interface ColumnResizeState {
    pointerId: number;
    sheetKey: string;
    columnNumber: number;
    startClientX: number;
    startPixelWidth: number;
    previewPixelWidth: number;
    isDragging: boolean;
}

interface RowResizeState {
    pointerId: number;
    sheetKey: string;
    rowNumber: number;
    startClientY: number;
    startPixelHeight: number;
    previewPixelHeight: number;
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
let rowResizeState: RowResizeState | null = null;
let suppressNextCellClick = false;
let suppressNextRowHeaderClick = false;
let suppressNextColumnHeaderClick = false;

const WHEEL_DELTA_LINE_MODE = 1;
const WHEEL_DELTA_PAGE_MODE = 2;
const WHEEL_LINE_SCROLL_PIXELS = 40;
const DEFAULT_EDITOR_VIEWPORT_HEIGHT = 480;
const DEFAULT_EDITOR_VIEWPORT_WIDTH = 960;
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

function normalizeWorkbookColumnWidth(columnWidth: number): number {
    return Math.round(columnWidth * 256) / 256;
}

function normalizeWorkbookRowHeight(rowHeight: number): number {
    return Math.round(rowHeight * 100) / 100;
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
    const [maximumDigitWidth, setMaximumDigitWidth] = React.useState(DEFAULT_MAXIMUM_DIGIT_WIDTH_PX);

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

function getEffectiveSheetRowHeights(
    currentModel: EditorRenderModel
): Readonly<Record<string, number | null>> | undefined {
    const activeResize = rowResizeState;
    if (!activeResize || activeResize.sheetKey !== currentModel.activeSheet.key) {
        return currentModel.activeSheet.rowHeights;
    }

    return {
        ...(currentModel.activeSheet.rowHeights ?? {}),
        [String(activeResize.rowNumber)]: normalizeWorkbookRowHeight(
            convertPixelsToWorkbookRowHeight(activeResize.previewPixelHeight)
        ),
    };
}

function getSheetColumnLayout(currentModel: EditorRenderModel): PixelColumnLayout {
    return createEditorPixelColumnLayout({
        columnCount: currentModel.activeSheet.columnCount,
        columnWidths: getEffectiveSheetColumnWidths(currentModel, editorMaximumDigitWidth),
        maximumDigitWidth: editorMaximumDigitWidth,
    });
}

function getSheetRowLayout(currentModel: EditorRenderModel): PixelRowLayout {
    return createEditorPixelRowLayout({
        rowCount: currentModel.activeSheet.rowCount,
        rowHeights: getEffectiveSheetRowHeights(currentModel),
    });
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

    const pane = getViewportElement();
    const viewportHeight = pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT;
    const viewportWidth = pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH;
    const scrollTop = pane?.scrollTop ?? 0;
    const scrollLeft = pane?.scrollLeft ?? 0;
    const sheetRowLayout = getSheetRowLayout(currentModel);
    const sheetColumnLayout = getSheetColumnLayout(currentModel);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: currentModel.activeSheet.rowCount,
        columnCount: currentModel.activeSheet.columnCount,
        viewportHeight,
        viewportWidth,
        rowLayout: sheetRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const displayRowLayout = getEditorDisplayRowLayout(sheetRowLayout, displayGrid.rowCount);
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
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
        frozenRowNumbers: createSequentialNumbers(visibleFrozenRowCount),
        frozenColumnNumbers: createSequentialNumbers(visibleFrozenColumnCount),
        rowNumbers: rowWindow.rowNumbers,
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

function getSelectionExtendAnchorCell(): CellPosition | null {
    return selectionAnchorCell ?? selectedCell;
}

function getSelectionExtendAnchorRowNumber(): number | null {
    return getSelectionExtendAnchorCell()?.rowNumber ?? null;
}

function getSelectionExtendAnchorColumnNumber(): number | null {
    return getSelectionExtendAnchorCell()?.columnNumber ?? null;
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

function requestPromptRowHeight(rowNumber: number): void {
    closeContextMenu({ refresh: false });
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
    closeContextMenu({ refresh: false });
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
                Math.round(columnResizeState.startPixelWidth + clientX - columnResizeState.startClientX),
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

function stopColumnResize(
    pointerId?: number,
    { commit = false }: { commit?: boolean } = {}
): void {
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

function beginRowResize(
    pointerId: number,
    sheetKey: string,
    rowNumber: number,
    startPixelHeight: number,
    clientY: number
): void {
    if (!model?.canEdit) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    closeContextMenu({ refresh: false });
    clearBrowserTextSelection();
    suppressNextRowHeaderClick = true;
    rowResizeState = {
        pointerId,
        sheetKey,
        rowNumber,
        startClientY: clientY,
        startPixelHeight,
        previewPixelHeight: startPixelHeight,
        isDragging: true,
    };
    renderApp({ commitEditing: false });
}

function updateRowResize(clientY: number): void {
    if (!rowResizeState?.isDragging) {
        return;
    }

    const nextPixelHeight = Math.max(
        MIN_ROW_PIXEL_HEIGHT,
        Math.min(
            MAX_ROW_PIXEL_HEIGHT,
            stabilizeRowPixelHeight(
                Math.round(rowResizeState.startPixelHeight + clientY - rowResizeState.startClientY)
            )
        )
    );
    if (nextPixelHeight === rowResizeState.previewPixelHeight) {
        return;
    }

    rowResizeState = {
        ...rowResizeState,
        previewPixelHeight: nextPixelHeight,
    };
    renderApp({ commitEditing: false });
}

function stopRowResize(
    pointerId?: number,
    { commit = false }: { commit?: boolean } = {}
): void {
    if (!rowResizeState) {
        return;
    }

    if (pointerId !== undefined && rowResizeState.pointerId !== pointerId) {
        return;
    }

    const activeResize = rowResizeState;
    const didChange = activeResize.previewPixelHeight !== activeResize.startPixelHeight;
    if (!commit || !didChange) {
        rowResizeState = null;
        renderApp({ commitEditing: false });
        return;
    }

    rowResizeState = {
        ...activeResize,
        startPixelHeight: activeResize.previewPixelHeight,
        isDragging: false,
    };
    renderApp({ commitEditing: false });
    vscode.postMessage({
        type: "setRowHeight",
        rowNumber: activeResize.rowNumber,
        height: normalizeWorkbookRowHeight(
            convertPixelsToWorkbookRowHeight(activeResize.previewPixelHeight)
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
    selectionDragState = {
        kind: "cell",
        anchorCell,
        pointerId,
    };
    clearSuppressedSelectionClicks();
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

    const sheetRowLayout = getSheetRowLayout(model);
    const sheetColumnLayout = getSheetColumnLayout(model);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: model.activeSheet.rowCount,
        columnCount: model.activeSheet.columnCount,
        viewportHeight: pane.clientHeight,
        viewportWidth: pane.clientWidth,
        rowLayout: sheetRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const displayRowLayout = getEditorDisplayRowLayout(sheetRowLayout, displayGrid.rowCount);
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
    const { rowCount: frozenRowCount, columnCount: frozenColumnCount } = getFrozenEditorCounts({
        rowCount: model.activeSheet.rowCount,
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

    if (selectedCell.rowNumber > frozenRowCount) {
        const cellTop =
            EDITOR_VIRTUAL_HEADER_HEIGHT + getEditorRowTop(displayRowLayout, selectedCell.rowNumber);
        const cellBottom = cellTop + getEditorRowHeight(displayRowLayout, selectedCell.rowNumber);
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
        const cellRight = cellLeft + getEditorColumnWidth(displayColumnLayout, selectedCell.columnNumber);
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
    selectionDragState = {
        kind: "row",
        anchorRowNumber,
        pointerId,
    };
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
    selectionDragState = {
        kind: "column",
        anchorColumnNumber,
        pointerId,
    };
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

function isSimpleSelection(
    cell: CellPosition | null,
    anchorCell: CellPosition | null = selectionAnchorCell,
    selectionRange: CellRange | null = selectionRangeOverride
): boolean {
    if (!cell) {
        return anchorCell === null && selectionRange === null;
    }

    return !hasExpandedSelectionRange(
        selectionRange ?? createSelectionRange(anchorCell ?? cell, cell)
    );
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
    const shouldForceRenderForFillHandle = Boolean(getFillHandleCell(nextSelectionRange));
    const canUseLocalUpdate = canUseLocalSimpleSelectionUpdate(nextCell, {
        anchorCell: nextAnchorCell,
        selectionRange: nextSelectionRange,
        forceRender: forceRender || shouldClearSearchFeedback || shouldForceRenderForFillHandle,
    });

    clearFillDragState();
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
        clearFillDragState();
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
        clearFillDragState();
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
    clearFillDragState();
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

    clearFillDragState();
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

    const pane = getViewportElement();
    const sheetRowLayout = getSheetRowLayout(model);
    const sheetColumnLayout = getSheetColumnLayout(model);
    const displayGrid = getEditorDisplayGridDimensions({
        rowCount: model.activeSheet.rowCount,
        columnCount: model.activeSheet.columnCount,
        viewportHeight: pane?.clientHeight ?? DEFAULT_EDITOR_VIEWPORT_HEIGHT,
        viewportWidth: pane?.clientWidth ?? DEFAULT_EDITOR_VIEWPORT_WIDTH,
        rowLayout: sheetRowLayout,
        columnLayout: sheetColumnLayout,
    });

    return {
        minRow: 1,
        maxRow: displayGrid.rowCount,
        minColumn: 1,
        maxColumn: displayGrid.columnCount,
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
        (viewportState?.rowNumbers.length ?? 1) - 1
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

function createEditorVirtualGridMetrics(
    currentModel: EditorRenderModel,
    sheetRowLayout: PixelRowLayout,
    sheetColumnLayout: PixelColumnLayout,
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
        rowLayout: sheetRowLayout,
        columnLayout: sheetColumnLayout,
    });
    const displayRowLayout = getEditorDisplayRowLayout(sheetRowLayout, displayGrid.rowCount);
    const displayColumnLayout = getEditorDisplayColumnLayout(
        sheetColumnLayout,
        displayGrid.columnCount
    );
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
    const activeResizePreviewPixelHeight =
        rowResizeState?.sheetKey === currentModel.activeSheet.key
            ? rowResizeState.previewPixelHeight
            : null;
    const activeResizeRowNumber =
        rowResizeState?.sheetKey === currentModel.activeSheet.key
            ? rowResizeState.rowNumber
            : null;
    const activeResizePreviewPixelWidth =
        columnResizeState?.sheetKey === currentModel.activeSheet.key
            ? columnResizeState.previewPixelWidth
            : null;
    const activeResizeColumnNumber =
        columnResizeState?.sheetKey === currentModel.activeSheet.key
            ? columnResizeState.columnNumber
            : null;
    const sheetRowLayout = React.useMemo(
        () =>
            createEditorPixelRowLayout({
                rowCount: currentModel.activeSheet.rowCount,
                rowHeights: getEffectiveSheetRowHeights(currentModel),
            }),
        [
            currentModel.activeSheet.key,
            currentModel.activeSheet.rowCount,
            currentModel.activeSheet.rowHeights,
            activeResizeRowNumber,
            activeResizePreviewPixelHeight,
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
            sheetRowLayout,
            sheetColumnLayout,
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
                sheetRowLayout,
                sheetColumnLayout,
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
            rowLayout: sheetRowLayout,
            columnLayout: sheetColumnLayout,
        });
        const displayRowLayout = getEditorDisplayRowLayout(sheetRowLayout, displayGrid.rowCount);
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
        currentModel.activeSheet.rowHeights,
        currentModel.activeSheet.columnWidths,
        currentModel.activeSheet.freezePane?.rowCount ?? 0,
        currentModel.activeSheet.freezePane?.columnCount ?? 0,
        activeResizeRowNumber,
        activeResizePreviewPixelHeight,
        activeResizeColumnNumber,
        activeResizePreviewPixelWidth,
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
        activeColumnNumber,
        selectionRange,
        rowCount,
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
        activeColumnNumber: number | null;
        selectionRange: CellRange | null;
        rowCount: number;
        top: number;
        left: number;
    }): React.ReactElement {
        const isActiveColumn =
            isActiveHighlightColumn(activeColumnNumber, columnNumber) ||
            isColumnHeaderWithinSelectionRange(selectionRange, columnNumber, rowCount);

        return (
            <div
                className={classNames([
                    "editor-grid__item",
                    "editor-grid__item--header",
                    "grid__column",
                    hasPending && "grid__column--diff",
                    hasPending && "grid__column--pending",
                    isActiveColumn && "grid__column--active",
                    isResizing && "grid__column--resizing",
                ])}
                data-column-number={columnNumber}
                data-role="grid-column-header"
                style={{
                    ...getEditorGridItemStyle({
                        top,
                        left,
                        width,
                        height: EDITOR_VIRTUAL_HEADER_HEIGHT,
                    }),
                    minWidth: `${width}px`,
                    maxWidth: `${width}px`,
                    "--grid-column-max-width": `${width}px`,
                } as React.CSSProperties}
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
        previous.activeColumnNumber === next.activeColumnNumber &&
        previous.rowCount === next.rowCount &&
        areSelectionRangesEqual(previous.selectionRange, next.selectionRange) &&
        previous.top === next.top &&
        previous.left === next.left
);

const EditorRowHeaderCell = React.memo(
    function EditorRowHeaderCell({
        sheetKey,
        rowNumber,
        height,
        canResize,
        isResizing,
        hasPending,
        activeRowNumber,
        selectionRange,
        columnCount,
        top,
        rowHeaderWidth,
    }: {
        sheetKey: string;
        rowNumber: number;
        height: number;
        canResize: boolean;
        isResizing: boolean;
        hasPending: boolean;
        activeRowNumber: number | null;
        selectionRange: CellRange | null;
        columnCount: number;
        top: number;
        rowHeaderWidth: number;
    }): React.ReactElement {
        const isActiveRow =
            isActiveHighlightRow(activeRowNumber, rowNumber) ||
            isRowHeaderWithinSelectionRange(selectionRange, rowNumber, columnCount);

        return (
            <div
                className={classNames([
                    "editor-grid__item",
                    "editor-grid__item--row-header",
                    "grid__row-number",
                    isResizing && "grid__row-number--resizing",
                    hasPending && "grid__row-number--pending",
                    isActiveRow && "grid__row-number--active",
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
                {canResize ? (
                    <span
                        aria-hidden
                        className={classNames(["grid__row-resize-handle", isResizing && "is-active"])}
                        data-role="grid-row-resize-handle"
                        onPointerDown={(event) => {
                            if (event.button !== 0) {
                                return;
                            }

                            event.preventDefault();
                            event.stopPropagation();
                            beginRowResize(event.pointerId, sheetKey, rowNumber, height, event.clientY);
                        }}
                    />
                ) : null}
            </div>
        );
    },
    (previous, next) =>
        previous.sheetKey === next.sheetKey &&
        previous.rowNumber === next.rowNumber &&
        previous.height === next.height &&
        previous.canResize === next.canResize &&
        previous.isResizing === next.isResizing &&
        previous.hasPending === next.hasPending &&
        previous.activeRowNumber === next.activeRowNumber &&
        previous.columnCount === next.columnCount &&
        areSelectionRangesEqual(previous.selectionRange, next.selectionRange) &&
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
    columnLayout,
    isActive,
    onPointerDown,
    onDoubleClick,
}: {
    rowNumber: number;
    columnNumber: number;
    rowHeaderWidth: number;
    rowLayout: PixelRowLayout;
    columnLayout: PixelColumnLayout;
    isActive: boolean;
    onPointerDown(event: React.PointerEvent<HTMLSpanElement>): void;
    onDoubleClick(event: React.MouseEvent<HTMLSpanElement>): void;
}): React.ReactElement {
    return (
        <span
            aria-hidden
            className={classNames(["editor-grid__fill-handle", isActive && "is-active"])}
            style={getEditorGridItemStyle({
                top:
                    getEditorGridTop(rowLayout, rowNumber) +
                    getEditorRowHeight(rowLayout, rowNumber) -
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
        selectionRange,
        activeRowNumber,
        activeColumnNumber,
        currentSelection,
        activeEditingCell,
        pendingEdit,
        fillSourceRange,
        fillPreviewRange,
    }: {
        currentModel: EditorRenderModel;
        rowNumber: number;
        columnNumber: number;
        width: number;
        height: number;
        top: number;
        left: number;
        selectionRange: CellRange | null;
        activeRowNumber: number | null;
        activeColumnNumber: number | null;
        currentSelection: CellPosition | null;
        activeEditingCell: EditingCell | null;
        pendingEdit: PendingEdit | undefined;
        fillSourceRange: CellRange | null;
        fillPreviewRange: CellRange | null;
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
        const isFillPreview = Boolean(
            fillSourceRange &&
            isCellWithinFillPreviewArea(fillSourceRange, fillPreviewRange, rowNumber, columnNumber)
        );

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
                    isFillPreview && "grid__cell--fill-preview",
                    isEditing && "grid__cell--editing",
                    ...getSelectionOutlineClasses(rowNumber, columnNumber, selectionRange),
                ])}
                data-column-number={columnNumber}
                data-editable={editable}
                data-role="grid-cell"
                data-row-number={rowNumber}
                style={{
                    ...getEditorGridItemStyle({
                        top,
                        left,
                        width,
                        height,
                    }),
                    minWidth: `${width}px`,
                    maxWidth: `${width}px`,
                    "--grid-column-max-width": `${width}px`,
                } as React.CSSProperties}
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
        previous.width === next.width &&
        previous.height === next.height &&
        previous.top === next.top &&
        previous.left === next.left &&
        previous.activeRowNumber === next.activeRowNumber &&
        previous.activeColumnNumber === next.activeColumnNumber &&
        areSelectionRangesEqual(previous.selectionRange, next.selectionRange) &&
        areCellPositionsEqual(previous.currentSelection, next.currentSelection) &&
        areEditingCellsEqual(previous.activeEditingCell, next.activeEditingCell) &&
        arePendingEditsEqual(previous.pendingEdit, next.pendingEdit) &&
        areSelectionRangesEqual(previous.fillSourceRange, next.fillSourceRange) &&
        areSelectionRangesEqual(previous.fillPreviewRange, next.fillPreviewRange)
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
    const bodyItems: React.ReactElement[] = [];
    const topItems: React.ReactElement[] = [];
    const leftItems: React.ReactElement[] = [];
    const cornerItems: React.ReactElement[] = [
        <EditorCornerHeader key="corner" rowHeaderWidth={metrics.rowHeaderWidth} />,
    ];
    const createGridCellItem = (
        keyPrefix: "body" | "left" | "top" | "corner",
        rowNumber: number,
        columnNumber: number
    ): React.ReactElement => {
        const top = getEditorGridTop(metrics.rowLayout, rowNumber);
        const left = getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, columnNumber);
        const width = getEditorColumnWidth(metrics.columnLayout, columnNumber);
        const height = getEditorRowHeight(metrics.rowLayout, rowNumber);
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
                width={width}
                height={height}
                top={top}
                left={left}
                selectionRange={selectionRange}
                activeRowNumber={activeRowNumber}
                activeColumnNumber={activeColumnNumber}
                currentSelection={currentSelection}
                activeEditingCell={activeEditingCell}
                pendingEdit={pendingEdit}
                fillSourceRange={fillSourceRange}
                fillPreviewRange={fillPreviewRange}
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
                activeColumnNumber={activeColumnNumber}
                selectionRange={selectionRange}
                rowCount={metrics.rowLayout.totalRowCount}
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
                activeColumnNumber={activeColumnNumber}
                selectionRange={selectionRange}
                rowCount={metrics.rowLayout.totalRowCount}
                top={0}
                left={getEditorGridLeft(metrics.rowHeaderWidth, metrics.columnLayout, columnNumber)}
            />
        );
    }

    for (const rowNumber of metrics.rowNumbers) {
        leftItems.push(
            <EditorRowHeaderCell
                key={`left:row:${rowNumber}`}
                sheetKey={currentModel.activeSheet.key}
                rowNumber={rowNumber}
                height={getEditorRowHeight(metrics.rowLayout, rowNumber)}
                canResize={currentModel.canEdit}
                isResizing={
                    rowResizeState?.sheetKey === currentModel.activeSheet.key &&
                    rowResizeState.rowNumber === rowNumber
                }
                hasPending={pendingSummary.rows.has(rowNumber)}
                activeRowNumber={activeRowNumber}
                selectionRange={selectionRange}
                columnCount={metrics.columnLayout.totalColumnCount}
                top={getEditorGridTop(metrics.rowLayout, rowNumber)}
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
                sheetKey={currentModel.activeSheet.key}
                rowNumber={rowNumber}
                height={getEditorRowHeight(metrics.rowLayout, rowNumber)}
                canResize={currentModel.canEdit}
                isResizing={
                    rowResizeState?.sheetKey === currentModel.activeSheet.key &&
                    rowResizeState.rowNumber === rowNumber
                }
                hasPending={pendingSummary.rows.has(rowNumber)}
                activeRowNumber={activeRowNumber}
                selectionRange={selectionRange}
                columnCount={metrics.columnLayout.totalColumnCount}
                top={getEditorGridTop(metrics.rowLayout, rowNumber)}
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

    const fillHandle =
        fillHandleCell && fillHandleLayer ? (
            <EditorFillHandle
                key={`fill-handle:${fillHandleCell.rowNumber}:${fillHandleCell.columnNumber}`}
                rowNumber={fillHandleCell.rowNumber}
                columnNumber={fillHandleCell.columnNumber}
                rowHeaderWidth={metrics.rowHeaderWidth}
                rowLayout={metrics.rowLayout}
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
    editorMaximumDigitWidth = maximumDigitWidth;
    const pendingSummary = getPendingSummary(currentModel.activeSheet.key);
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

    return (
        <div
            className={classNames([
                "app",
                "app--editor",
                columnResizeState?.isDragging && "app--column-resizing",
                rowResizeState?.isDragging && "app--row-resizing",
            ])}
            data-role="editor-app"
        >
            <EditorToolbar
                strings={STRINGS}
                currentModel={currentModel}
                isSearchPanelOpen={isSearchPanelOpen}
                isSaving={isSaving}
                hasPendingEdits={hasPendingEdits}
                canUndo={canUndo}
                canRedo={canRedo}
                viewLocked={viewLocked}
                getPositionInputValue={getPositionInputValue}
                getCellValueInputValue={getCellValueInputValue}
                getCellValueInputPlaceholder={getCellValueInputPlaceholder}
                canEditSelectedCellValue={canEditSelectedCellValue}
                getActiveCellEditTarget={() => getActiveToolbarCellEditTarget(currentModel)}
                onOpenSearch={openSearchPanel}
                onUndo={undoPendingEdits}
                onRedo={redoPendingEdits}
                onReload={requestReload}
                onToggleViewLock={toggleViewLock}
                onSave={triggerSave}
                onSubmitGoto={submitGotoSelection}
                onCommitCellValue={commitToolbarCellValue}
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
                        : STRINGS.searchScopeWholeSheet
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
        rowResizeState = null;
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
        clearFillDragState();
        columnResizeState = null;
        rowResizeState = null;
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
        clearFillDragState();
        columnResizeState = null;
        rowResizeState = null;
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
    if (columnResizeState?.isDragging && columnResizeState.pointerId === event.pointerId) {
        if ((event.buttons & 1) === 0) {
            stopColumnResize(event.pointerId, { commit: true });
            return;
        }

        event.preventDefault();
        updateColumnResize(event.clientX);
        return;
    }

    if (rowResizeState?.isDragging && rowResizeState.pointerId === event.pointerId) {
        if ((event.buttons & 1) === 0) {
            stopRowResize(event.pointerId, { commit: true });
            return;
        }

        event.preventDefault();
        updateRowResize(event.clientY);
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
    stopRowResize(event.pointerId, { commit: true });
    stopSearchPanelDrag(event.pointerId);
    stopFillDrag(event.pointerId, { commit: true, timeStamp: event.timeStamp });
    stopSelectionDrag(event.pointerId);
    scheduleSuppressedSelectionClickReset();
});

document.addEventListener("pointercancel", (event: PointerEvent) => {
    stopColumnResize(event.pointerId);
    stopRowResize(event.pointerId);
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
