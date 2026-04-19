import type {
    EditorGridCellView,
    EditorRenderModel,
    EditorSheetTabView,
} from "../core/model/types";
import { getColumnLabel } from "../core/model/cells";
import { DEFAULT_PAGE_SIZE } from "../constants";

interface VsCodeApi {
    postMessage(message: OutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

type OutgoingMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "prevPage" }
    | { type: "nextPage" }
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
      };

const vscode = acquireVsCodeApi();

let model: EditorRenderModel | null = null;
let selectedCell: { rowNumber: number; columnNumber: number } | null = null;
let editState: {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    cellEl: HTMLElement;
    input: HTMLInputElement;
} | null = null;
let isSaving = false;
let lastPendingNotification: boolean | null = null;
let pendingSelectionAfterRender: {
    rowNumber: number;
    columnNumber: number;
    reveal: boolean;
} | null = null;
let searchQuery = "";
let gotoReference = "";
let lastPendingEditsSyncKey: string | null = null;

interface SearchOptions {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

let searchOptions: SearchOptions = {
    isRegexp: false,
    matchCase: false,
    wholeWord: false,
};

interface PendingEdit {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    value: string;
}

const pendingEdits = new Map<string, PendingEdit>();

interface PendingEditChange {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    modelValue: string;
    beforeValue: string;
    afterValue: string;
}

interface HistoryEntry {
    changes: PendingEditChange[];
}

const undoStack: HistoryEntry[] = [];
const redoStack: HistoryEntry[] = [];

function getPendingEditKey(sheetKey: string, rowNumber: number, columnNumber: number): string {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
}

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
    prevPage: "Prev Page",
    nextPage: "Next Page",
    size: "Size",
    modified: "Modified",
    sheet: "Sheet",
    rows: "Rows",
    noRows: "No rows",
    page: "Page",
    visibleRows: "Visible rows",
    readOnly: "Read-only",
    save: "Save",
    selectedCell: "Selected cell",
    noCellSelected: "None",
    totalSheets: "Sheets",
    totalRows: "Rows",
    nonEmptyCells: "Non-empty cells",
    mergedRanges: "Merged ranges",
    noRowsAvailable: "No rows available on this page.",
    readOnlyBadge: "Read-only",
};

type Strings = typeof DEFAULT_STRINGS;

const STRINGS: Strings =
    ((globalThis as Record<string, unknown>).__XLSX_EDITOR_STRINGS__ as Strings | undefined) ??
    DEFAULT_STRINGS;

function escapeHtml(value: string): string {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
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

    if (
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target.isContentEditable
    ) {
        return true;
    }

    return Boolean(
        target.closest('input, textarea, [contenteditable="true"], [contenteditable=""]')
    );
}

function clearBrowserTextSelection(): void {
    globalThis.getSelection?.()?.removeAllRanges();
}

function renderLoading(message: string): void {
    document.body.innerHTML = `
		<div id="app" class="loading-shell">
			<div class="loading-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

function renderError(message: string): void {
    document.body.innerHTML = `
		<div id="app" class="empty-shell">
			<div class="empty-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

function updateLabelMarker(labelEl: Element, hasPending: boolean, extraClass?: string): void {
    let markerEl = labelEl.querySelector<HTMLElement>(".diff-marker");

    if (!hasPending) {
        markerEl?.remove();
        return;
    }

    if (!markerEl) {
        markerEl = document.createElement("span");
        markerEl.setAttribute("aria-hidden", "true");
        labelEl.insertBefore(markerEl, labelEl.firstChild);
    }

    markerEl.className = `diff-marker diff-marker--pending${extraClass ? ` ${extraClass}` : ""}`;
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

function renderCellValue(value: string, formula: string | null): string {
    if (!value && !formula) {
        return "";
    }

    return `${value ? `<span class="grid__cell-value">${escapeHtml(value)}</span>` : ""}${
        formula ? `<span class="cell__formula" title="${escapeHtml(formula)}">fx</span>` : ""
    }`;
}

function getCellView(rowNumber: number, columnNumber: number): EditorGridCellView | null {
    const row = model?.page.rows.find((item) => item.rowNumber === rowNumber);
    return row?.cells[columnNumber - 1] ?? null;
}

function isGridCellEditable(cell: EditorGridCellView | null): boolean {
    return Boolean(model?.canEdit && !cell?.formula);
}

function canEditCellAt(rowNumber: number, columnNumber: number): boolean {
    return isGridCellEditable(getCellView(rowNumber, columnNumber));
}

function clearSelectedCells(): void {
    for (const cell of document.querySelectorAll(".grid__cell--selected")) {
        cell.classList.remove("grid__cell--selected");
    }

    for (const cell of document.querySelectorAll('[data-role="grid-cell"][aria-selected="true"]')) {
        cell.setAttribute("aria-selected", "false");
    }
}

function getSelectedCellElements(): HTMLElement[] {
    if (!selectedCell) {
        return [];
    }

    return Array.from(
        document.querySelectorAll<HTMLElement>(
            `[data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"][data-column-number="${selectedCell.columnNumber}"]`
        )
    );
}

function getSelectedCellAddress(): string {
    if (!selectedCell) {
        return STRINGS.noCellSelected;
    }

    const cell = getCellView(selectedCell.rowNumber, selectedCell.columnNumber);
    return cell?.address ?? `${getColumnLabel(selectedCell.columnNumber)}${selectedCell.rowNumber}`;
}

function updateSelectedCellBadge(): void {
    const badge = document.querySelector<HTMLElement>('[data-role="selected-cell-address"]');
    if (!badge) {
        return;
    }

    const address = getSelectedCellAddress();
    const title = `${STRINGS.selectedCell}: ${address}`;
    badge.textContent = address;
    badge.title = title;
    badge.setAttribute("aria-label", title);
}

function clearSelectionContextHighlights(): void {
    for (const element of document.querySelectorAll(
        ".grid__cell--active-row, .grid__cell--active-column, .grid__row-number--active, .grid__column--active"
    )) {
        element.classList.remove(
            "grid__cell--active-row",
            "grid__cell--active-column",
            "grid__row-number--active",
            "grid__column--active"
        );
    }
}

function applySelectionContextHighlights(): void {
    clearSelectionContextHighlights();
    updateSelectedCellBadge();

    if (!selectedCell) {
        return;
    }

    for (const element of document.querySelectorAll<HTMLElement>(
        `[data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"]`
    )) {
        element.classList.add("grid__cell--active-row");
    }

    for (const element of document.querySelectorAll<HTMLElement>(
        `[data-role="grid-cell"][data-column-number="${selectedCell.columnNumber}"]`
    )) {
        element.classList.add("grid__cell--active-column");
    }

    const rowHeader = document.querySelector<HTMLElement>(
        `th[data-role="grid-row-header"][data-row-number="${selectedCell.rowNumber}"]`
    );
    rowHeader?.classList.add("grid__row-number--active");

    const columnHeader = document.querySelector<HTMLElement>(
        `thead th[data-column-number="${selectedCell.columnNumber}"]`
    );
    columnHeader?.classList.add("grid__column--active");
}

function clampScrollPosition(value: number, maxValue: number): number {
    return Math.max(0, Math.min(value, Math.max(maxValue, 0)));
}

function getPaneScrollState(): { top: number; left: number } | null {
    const pane = document.querySelector<HTMLElement>(".pane__table");
    if (!pane) {
        return null;
    }

    return {
        top: pane.scrollTop,
        left: pane.scrollLeft,
    };
}

function restorePaneScrollState(scrollState: { top: number; left: number } | null): void {
    if (!scrollState) {
        return;
    }

    const pane = document.querySelector<HTMLElement>(".pane__table");
    if (!pane) {
        return;
    }

    pane.scrollTop = clampScrollPosition(scrollState.top, pane.scrollHeight - pane.clientHeight);
    pane.scrollLeft = clampScrollPosition(scrollState.left, pane.scrollWidth - pane.clientWidth);
}

function getStickyPaneInsets(pane: HTMLElement): { top: number; left: number } {
    const headerRow = pane.querySelector("thead tr");
    const firstColumn = pane.querySelector("thead th:first-child");

    return {
        top: headerRow?.getBoundingClientRect().height ?? 0,
        left: firstColumn?.getBoundingClientRect().width ?? 0,
    };
}

function revealSelectedCells(elements: HTMLElement[]): void {
    const pane = document.querySelector<HTMLElement>(".pane__table");
    if (!pane || elements.length === 0) {
        return;
    }

    const target = elements[0];
    const paneRect = pane.getBoundingClientRect();
    const elementRect = target.getBoundingClientRect();
    const stickyInsets = getStickyPaneInsets(pane);
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
}

function applySelectedCell({ reveal = false }: { reveal?: boolean } = {}): void {
    clearSelectedCells();
    clearBrowserTextSelection();
    applySelectionContextHighlights();

    if (!selectedCell) {
        return;
    }

    const elements = getSelectedCellElements();
    if (elements.length === 0) {
        return;
    }

    for (const element of elements) {
        element.classList.add("grid__cell--selected");
        element.setAttribute("aria-selected", "true");
    }

    if (reveal) {
        revealSelectedCells(elements);
    }
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

function setSelectedCellLocal(
    nextCell: { rowNumber: number; columnNumber: number } | null,
    { reveal = false, syncHost = true }: { reveal?: boolean; syncHost?: boolean } = {}
): void {
    selectedCell = nextCell;
    applySelectedCell({ reveal });

    if (syncHost) {
        syncSelectedCellToHost();
    }
}

function syncSelectedCellAfterRender({
    reveal = false,
    useModelSelection = false,
}: { reveal?: boolean; useModelSelection?: boolean } = {}): void {
    if (
        useModelSelection &&
        model?.selection &&
        getCellView(model.selection.rowNumber, model.selection.columnNumber)
    ) {
        selectedCell = {
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        };
        pendingSelectionAfterRender = null;
        applySelectedCell({ reveal });
        syncSelectedCellToHost();
        return;
    }

    if (
        pendingSelectionAfterRender &&
        getCellView(pendingSelectionAfterRender.rowNumber, pendingSelectionAfterRender.columnNumber)
    ) {
        selectedCell = {
            rowNumber: pendingSelectionAfterRender.rowNumber,
            columnNumber: pendingSelectionAfterRender.columnNumber,
        };
        pendingSelectionAfterRender = null;
        applySelectedCell({ reveal: true });
        syncSelectedCellToHost();
        return;
    }

    pendingSelectionAfterRender = null;

    if (selectedCell && getCellView(selectedCell.rowNumber, selectedCell.columnNumber)) {
        applySelectedCell({ reveal });
        syncSelectedCellToHost();
        return;
    }

    selectedCell = model?.selection
        ? {
              rowNumber: model.selection.rowNumber,
              columnNumber: model.selection.columnNumber,
          }
        : model?.page.rows[0]
          ? { rowNumber: model.page.rows[0].rowNumber, columnNumber: 1 }
          : null;

    applySelectedCell({ reveal });
    if (selectedCell) {
        syncSelectedCellToHost();
    }
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

function updateSaveButtonState(): void {
    notifyPendingEditState();

    const saveBtn = document.querySelector<HTMLButtonElement>('[data-action="save-edits"]');
    const hasPendingEdits = pendingEdits.size > 0 || Boolean(model?.hasPendingEdits);
    if (saveBtn) {
        saveBtn.disabled = !model?.canEdit || !hasPendingEdits || isSaving;
        saveBtn.classList.toggle("is-dirty", hasPendingEdits);
        if (isSaving) {
            saveBtn.setAttribute("aria-busy", "true");
        } else {
            saveBtn.removeAttribute("aria-busy");
        }
    }

    updateHistoryButtonState();
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

function updateHistoryButtonState(): void {
    const undoButton = document.querySelector<HTMLButtonElement>('[data-action="undo"]');
    if (undoButton) {
        undoButton.disabled = !model?.canEdit || undoStack.length === 0 || isSaving;
    }

    const redoButton = document.querySelector<HTMLButtonElement>('[data-action="redo"]');
    if (redoButton) {
        redoButton.disabled = !model?.canEdit || redoStack.length === 0 || isSaving;
    }
}

function syncCellDisplay(
    cellEl: HTMLElement,
    sheetKey: string,
    rowNumber: number,
    columnNumber: number
): void {
    const content = cellEl.querySelector<HTMLElement>(".grid__cell-content");
    if (!content) {
        return;
    }

    const pendingKey = getPendingEditKey(sheetKey, rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const value = pendingEdit ? pendingEdit.value : getCellModelValue(rowNumber, columnNumber);
    const formula = pendingEdit ? null : getCellFormula(rowNumber, columnNumber);

    content.innerHTML = renderCellValue(value, formula);
    cellEl.classList.toggle("grid__cell--pending", Boolean(pendingEdit));
}

function finishEdit({
    mode,
    clearSelection = false,
}: {
    mode: "commit" | "cancel";
    clearSelection?: boolean;
}): void {
    const session = editState;
    if (!session) {
        return;
    }

    editState = null;
    const { sheetKey, rowNumber, columnNumber, cellEl, input } = session;
    cellEl.classList.remove("grid__cell--editing");

    if (mode === "commit") {
        commitEdit(sheetKey, rowNumber, columnNumber, input.value);
        syncCellDisplay(cellEl, sheetKey, rowNumber, columnNumber);
        if (clearSelection) {
            setSelectedCellLocal(null, { syncHost: false });
        }
        return;
    }

    syncCellDisplay(cellEl, sheetKey, rowNumber, columnNumber);
}

function clearSelectedCellValue(): void {
    if (
        !model ||
        !selectedCell ||
        !canEditCellAt(selectedCell.rowNumber, selectedCell.columnNumber)
    ) {
        return;
    }

    commitEdit(model.activeSheet.key, selectedCell.rowNumber, selectedCell.columnNumber, "");
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

function applyPendingEditStyles(): void {
    const activeSheetKey = model?.activeSheet.key;
    if (!activeSheetKey) {
        return;
    }

    const pendingSheetKeys = new Set<string>();
    const pendingRows = new Set<number>();
    const pendingColumns = new Set<number>();

    for (const pendingEdit of pendingEdits.values()) {
        pendingSheetKeys.add(pendingEdit.sheetKey);
        if (pendingEdit.sheetKey !== activeSheetKey) {
            continue;
        }

        pendingRows.add(pendingEdit.rowNumber);
        pendingColumns.add(pendingEdit.columnNumber);
    }

    for (const cellEl of document.querySelectorAll<HTMLElement>('[data-role="grid-cell"]')) {
        if (cellEl.classList.contains("grid__cell--editing")) {
            continue;
        }

        const rowNumber = Number(cellEl.getAttribute("data-row-number"));
        const columnNumber = Number(cellEl.getAttribute("data-column-number"));
        syncCellDisplay(cellEl, activeSheetKey, rowNumber, columnNumber);
    }

    for (const rowHeader of document.querySelectorAll<HTMLElement>(
        'th[data-role="grid-row-header"]'
    )) {
        const rowNumber = Number(rowHeader.getAttribute("data-row-number"));
        const hasPending = pendingRows.has(rowNumber);
        rowHeader.classList.toggle("grid__row-number--pending", hasPending);
        const labelEl = rowHeader.querySelector(".grid__row-label");
        if (labelEl) {
            updateLabelMarker(labelEl, hasPending);
        }
    }

    for (const colHeader of document.querySelectorAll<HTMLElement>(
        "thead th[data-column-number]"
    )) {
        const columnNumber = Number(colHeader.getAttribute("data-column-number"));
        const hasPending = pendingColumns.has(columnNumber);
        colHeader.classList.toggle("grid__column--diff", hasPending);
        colHeader.classList.toggle("grid__column--pending", hasPending);
        const labelEl = colHeader.querySelector(".grid__column-label");
        if (labelEl) {
            updateLabelMarker(labelEl, hasPending);
        }
    }

    for (const tabEl of document.querySelectorAll<HTMLElement>('[data-action="set-sheet"]')) {
        const sheetKey = tabEl.getAttribute("data-sheet-key");
        const hasPending = sheetKey ? pendingSheetKeys.has(sheetKey) : false;
        updateLabelMarker(tabEl, hasPending, "tab__marker");
    }
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

function applyEditChanges(
    changes: PendingEditChange[],
    { recordHistory = true }: { recordHistory?: boolean } = {}
): void {
    const effectiveChanges = changes.filter((change) => change.beforeValue !== change.afterValue);
    if (effectiveChanges.length === 0) {
        updateSaveButtonState();
        return;
    }

    if (recordHistory) {
        undoStack.push({ changes: effectiveChanges });
        redoStack.length = 0;
    }

    for (const change of effectiveChanges) {
        setPendingCellValue(change, change.afterValue);
    }

    applyPendingEditStyles();
    updateSaveButtonState();
    syncPendingEditsToHost();
}

function applyHistoryEntry(entry: HistoryEntry, direction: "undo" | "redo"): void {
    for (const change of entry.changes) {
        setPendingCellValue(change, direction === "undo" ? change.beforeValue : change.afterValue);
    }

    applyPendingEditStyles();
    updateSaveButtonState();
    syncPendingEditsToHost();
}

function undoPendingEdits(): void {
    const entry = undoStack.pop();
    if (!entry) {
        updateHistoryButtonState();
        return;
    }

    applyHistoryEntry(entry, "undo");
    redoStack.push(entry);
    updateHistoryButtonState();
}

function redoPendingEdits(): void {
    const entry = redoStack.pop();
    if (!entry) {
        updateHistoryButtonState();
        return;
    }

    applyHistoryEntry(entry, "redo");
    undoStack.push(entry);
    updateHistoryButtonState();
}

function commitEdit(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    value: string
): void {
    const modelValue = getCellModelValue(rowNumber, columnNumber);
    const beforeValue =
        pendingEdits.get(getPendingEditKey(sheetKey, rowNumber, columnNumber))?.value ?? modelValue;

    applyEditChanges([
        {
            sheetKey,
            rowNumber,
            columnNumber,
            modelValue,
            beforeValue,
            afterValue: value,
        },
    ]);
}

function triggerSave(): void {
    if (!model || (!model.hasPendingEdits && pendingEdits.size === 0) || isSaving) {
        return;
    }

    finishEdit({ mode: "commit", clearSelection: true });
    isSaving = true;
    updateSaveButtonState();
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
    if (!model || !selectedCell || grid.length === 0 || !model.canEdit) {
        return;
    }

    const maxRow = model.activeSheet.rowCount;
    const maxColumn = model.activeSheet.columnCount;
    const changes: PendingEditChange[] = [];

    for (let rowOffset = 0; rowOffset < grid.length; rowOffset += 1) {
        const targetRow = selectedCell.rowNumber + rowOffset;
        if (targetRow > maxRow) {
            break;
        }

        const values = grid[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
            const targetColumn = selectedCell.columnNumber + columnOffset;
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

    applyEditChanges(changes);

    applySelectedCell({ reveal: true });
}

function enterEditMode(
    cellEl: HTMLElement,
    rowNumber: number,
    columnNumber: number,
    currentValue: string
): void {
    finishEdit({ mode: "commit" });
    clearBrowserTextSelection();

    const capturedSheetKey = model?.activeSheet.key;
    const content = cellEl.querySelector<HTMLElement>(".grid__cell-content");
    if (!content || !capturedSheetKey) {
        return;
    }

    content.innerHTML = "";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "grid__cell-input";
    input.value = currentValue;
    content.appendChild(input);
    cellEl.classList.add("grid__cell--editing");
    editState = { sheetKey: capturedSheetKey, rowNumber, columnNumber, cellEl, input };

    input.focus();
    input.select();

    input.addEventListener("keydown", (event: KeyboardEvent) => {
        if (event.key === "Enter" || event.key === "Tab") {
            event.preventDefault();
            finishEdit({ mode: "commit", clearSelection: true });
        } else if (event.key === "Escape") {
            event.preventDefault();
            finishEdit({ mode: "cancel" });
        }
    });

    input.addEventListener("blur", () => {
        if (editState?.input === input) {
            finishEdit({ mode: "commit", clearSelection: true });
        }
    });
}

function getSelectionBounds(): {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
} | null {
    if (!model || model.page.rows.length === 0) {
        return null;
    }

    return {
        minRow: model.page.rows[0].rowNumber,
        maxRow: model.page.rows[model.page.rows.length - 1].rowNumber,
        minColumn: 1,
        maxColumn: model.page.columns.length,
    };
}

function ensureSelection(): { rowNumber: number; columnNumber: number } | null {
    if (selectedCell) {
        return selectedCell;
    }

    if (model?.selection) {
        selectedCell = {
            rowNumber: model.selection.rowNumber,
            columnNumber: model.selection.columnNumber,
        };
        return selectedCell;
    }

    if (model?.page.rows[0]) {
        selectedCell = {
            rowNumber: model.page.rows[0].rowNumber,
            columnNumber: 1,
        };
        return selectedCell;
    }

    return null;
}

function moveSelection(rowDelta: number, columnDelta: number): void {
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

    setSelectedCellLocal({ rowNumber: nextRow, columnNumber: nextColumn }, { reveal: true });
}

function moveSelectionByPage(direction: -1 | 1): void {
    const selection = ensureSelection();
    if (!model || !selection) {
        return;
    }

    const canMove = direction < 0 ? model.canPrevPage : model.canNextPage;
    if (!canMove) {
        return;
    }

    pendingSelectionAfterRender = {
        rowNumber: Math.max(
            1,
            Math.min(
                model.activeSheet.rowCount,
                selection.rowNumber + direction * DEFAULT_PAGE_SIZE
            )
        ),
        columnNumber: selection.columnNumber,
        reveal: true,
    };

    vscode.postMessage({ type: direction < 0 ? "prevPage" : "nextPage" });
}

function renderTable(): string {
    if (!model || model.page.rows.length === 0) {
        return `<div class="empty-table">${escapeHtml(STRINGS.noRowsAvailable)}</div>`;
    }

    const headerColumns = model.page.columns
        .map(
            (column, index) => `<th class="grid__column" data-column-number="${index + 1}">
				<span class="grid__column-label"><span>${escapeHtml(column)}</span></span>
			</th>`
        )
        .join("");

    const bodyRows = model.page.rows
        .map((row) => {
            const cells = row.cells
                .map((cell, columnIndex) => {
                    const cellClass = [
                        "grid__cell",
                        cell.isSelected ? "grid__cell--selected" : "",
                        isGridCellEditable(cell) ? "" : "grid__cell--locked",
                    ]
                        .filter(Boolean)
                        .join(" ");
                    const editable = isGridCellEditable(cell) ? "true" : "false";

                    return `<td title="${escapeHtml(getCellTooltip(cell.address, cell.value, cell.formula))}" class="${cellClass}" data-role="grid-cell" data-row-number="${row.rowNumber}" data-column-number="${columnIndex + 1}" data-editable="${editable}" aria-selected="${cell.isSelected ? "true" : "false"}">
						<div class="grid__cell-content">${renderCellValue(cell.value, cell.formula)}</div>
					</td>`;
                })
                .join("");

            return `
				<tr data-role="grid-row" data-row-number="${row.rowNumber}">
					<th class="grid__row-number" data-role="grid-row-header" data-row-number="${row.rowNumber}">
						<span class="grid__row-label"><span>${row.rowNumber}</span></span>
					</th>
					${cells}
				</tr>
			`;
        })
        .join("");

    return `
		<table class="grid">
			<thead>
				<tr>
					<th class="grid__row-number">#</th>
					${headerColumns}
				</tr>
			</thead>
			<tbody>${bodyRows}</tbody>
		</table>
	`;
}

function renderToolbarButton({
    action,
    icon,
    label,
    disabled = false,
    isActive = false,
    iconOnly = false,
}: {
    action: string;
    icon: string;
    label: string;
    disabled?: boolean;
    isActive?: boolean;
    iconOnly?: boolean;
}): string {
    return `<button class="toolbar__button${isActive ? " is-active" : ""}${iconOnly ? " toolbar__button--icon" : ""}" data-action="${action}" ${disabled ? "disabled" : ""} title="${escapeHtml(label)}" aria-label="${escapeHtml(label)}" type="button">
		<span class="codicon ${icon} toolbar__button-icon" aria-hidden="true"></span>
		${iconOnly ? "" : `<span>${escapeHtml(label)}</span>`}
	</button>`;
}

function renderEmbeddedSearchToggle({
    action,
    text,
    label,
    isActive,
}: {
    action: string;
    text: string;
    label: string;
    isActive: boolean;
}): string {
    return `<button class="toolbar__toggle${isActive ? " is-active" : ""}" data-action="${action}" title="${escapeHtml(label)}" aria-label="${escapeHtml(label)}" type="button">${escapeHtml(text)}</button>`;
}

function renderToolbar(): string {
    const currentModel = model!;

    return `
		<header class="toolbar toolbar--editor">
			<div class="toolbar__group toolbar__group--grow">
				<label class="toolbar__field">
					<span class="codicon codicon-search toolbar__field-icon" aria-hidden="true"></span>
					<input class="toolbar__input" data-role="search-input" type="text" value="${escapeHtml(searchQuery)}" placeholder="${escapeHtml(STRINGS.searchPlaceholder)}" />
					<span class="toolbar__field-actions">
						${renderEmbeddedSearchToggle({ action: "toggle-search-regex", text: ".*", label: STRINGS.searchRegex, isActive: searchOptions.isRegexp })}
						${renderEmbeddedSearchToggle({ action: "toggle-search-match-case", text: "Aa", label: STRINGS.searchMatchCase, isActive: searchOptions.matchCase })}
						${renderEmbeddedSearchToggle({ action: "toggle-search-whole-word", text: "ab", label: STRINGS.searchWholeWord, isActive: searchOptions.wholeWord })}
					</span>
				</label>
				${renderToolbarButton({ action: "search-prev", icon: "codicon-arrow-up", label: STRINGS.findPrev, iconOnly: true })}
				${renderToolbarButton({ action: "search-next", icon: "codicon-arrow-down", label: STRINGS.findNext, iconOnly: true })}
				<label class="toolbar__field toolbar__field--goto">
					<span class="codicon codicon-target toolbar__field-icon" aria-hidden="true"></span>
					<input class="toolbar__input" data-role="goto-input" type="text" value="${escapeHtml(gotoReference)}" placeholder="${escapeHtml(STRINGS.gotoPlaceholder)}" />
				</label>
				${renderToolbarButton({ action: "goto-cell", icon: "codicon-target", label: STRINGS.goto })}
			</div>
			<div class="toolbar__group">
				${renderToolbarButton({ action: "undo", icon: "codicon-undo", label: STRINGS.undo, disabled: !currentModel.canEdit || undoStack.length === 0, iconOnly: true })}
				${renderToolbarButton({ action: "redo", icon: "codicon-redo", label: STRINGS.redo, disabled: !currentModel.canEdit || redoStack.length === 0, iconOnly: true })}
				${renderToolbarButton({ action: "reload", icon: "codicon-refresh", label: STRINGS.reload })}
				${renderToolbarButton({ action: "save-edits", icon: "codicon-save", label: STRINGS.save, disabled: !currentModel.canSave })}
			</div>
		</header>
	`;
}

function renderPane(): string {
    const selectedAddress = model!.selection?.address ?? STRINGS.noCellSelected;
    const selectedCellTitle = `${STRINGS.selectedCell}: ${selectedAddress}`;

    return `
		<section class="pane pane--single">
			<div class="pane__header">
				<div class="pane__header-group">
					<div class="pane__title">${escapeHtml(model!.activeSheet.label)}</div>
					<span class="badge badge--selection" data-role="selected-cell-address" title="${escapeHtml(selectedCellTitle)}">${escapeHtml(selectedAddress)}</span>
				</div>
				${model!.activeSheet.hasMergedRanges ? `<span class="badge badge--warn">${escapeHtml(STRINGS.mergedRanges)}: ${model!.activeSheet.mergedRangeCount}</span>` : ""}
			</div>
			<div class="pane__table">${renderTable()}</div>
		</section>
	`;
}

function renderTabs(): string {
    return model!.sheets
        .map(
            (sheet: EditorSheetTabView) => `
				<button
					class="tab ${sheet.isActive ? "is-active" : ""}"
					data-action="set-sheet"
					data-sheet-key="${escapeHtml(sheet.key)}"
					title="${escapeHtml(sheet.label)}"
				>
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`
        )
        .join("");
}

function renderStatus(): string {
    const currentModel = model!;
    const rowRangeLabel =
        currentModel.page.visibleRowCount === 0 ? STRINGS.noRows : currentModel.page.rangeLabel;

    return `
		<footer class="footer">
			<div class="tabs">${renderTabs()}</div>
			<div class="status">
				<span><strong>${escapeHtml(STRINGS.sheet)}:</strong> ${escapeHtml(currentModel.activeSheet.label)}</span>
				<span><strong>${escapeHtml(STRINGS.rows)}:</strong> ${escapeHtml(rowRangeLabel)}</span>
				<span><strong>${escapeHtml(STRINGS.page)}:</strong> ${currentModel.page.currentPage} / ${currentModel.page.totalPages}</span>
				<span><strong>${escapeHtml(STRINGS.visibleRows)}:</strong> ${currentModel.page.visibleRowCount}</span>
			</div>
		</footer>
	`;
}

function renderApp({
    revealSelection = false,
    useModelSelection = false,
}: { revealSelection?: boolean; useModelSelection?: boolean } = {}): void {
    if (!model) {
        renderLoading(STRINGS.loading);
        return;
    }

    finishEdit({ mode: "commit" });
    const previousScrollState = getPaneScrollState();

    document.body.innerHTML = `
		<div id="app" class="app app--editor">
			${renderToolbar()}
			<section class="panes panes--single">
				${renderPane()}
			</section>
			${renderStatus()}
		</div>
	`;

    bindToolbarInputs();
    restorePaneScrollState(previousScrollState);
    applyPendingEditStyles();
    updateSaveButtonState();
    syncSelectedCellAfterRender({ reveal: revealSelection, useModelSelection });
}

function bindToolbarInputs(): void {
    const searchInput = document.querySelector<HTMLInputElement>('[data-role="search-input"]');
    if (searchInput) {
        searchInput.addEventListener("input", () => {
            searchQuery = searchInput.value;
        });

        searchInput.addEventListener("keydown", (event: KeyboardEvent) => {
            if (event.key !== "Enter") {
                return;
            }

            event.preventDefault();
            submitSearch(event.shiftKey ? "prev" : "next");
        });
    }

    const gotoInput = document.querySelector<HTMLInputElement>('[data-role="goto-input"]');
    if (gotoInput) {
        gotoInput.addEventListener("input", () => {
            gotoReference = gotoInput.value;
        });

        gotoInput.addEventListener("keydown", (event: KeyboardEvent) => {
            if (event.key !== "Enter") {
                return;
            }

            event.preventDefault();
            submitGoto();
        });
    }
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

window.addEventListener("message", (event: MessageEvent<IncomingMessage>) => {
    const message = event.data;

    if (message.type === "loading") {
        renderLoading(message.message);
        return;
    }

    if (message.type === "error") {
        isSaving = false;
        updateSaveButtonState();
        renderError(message.message);
        return;
    }

    if (message.type === "render") {
        model = message.payload;
        isSaving = false;
        if (message.clearPendingEdits) {
            pendingEdits.clear();
            undoStack.length = 0;
            redoStack.length = 0;
            lastPendingNotification = null;
            lastPendingEditsSyncKey = serializePendingEdits([]);
        }

        renderApp({
            revealSelection: !message.silent,
            useModelSelection: message.useModelSelection,
        });
    }
});

document.addEventListener("pointerdown", (event: PointerEvent) => {
    if (!editState) {
        return;
    }

    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget) {
        return;
    }

    if (eventTarget === editState.input || eventTarget.closest(".grid__cell--editing")) {
        return;
    }

    finishEdit({ mode: "commit" });
});

document.addEventListener("click", (event: MouseEvent) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget) {
        return;
    }

    const cellTarget = eventTarget.closest<HTMLElement>('[data-role="grid-cell"]');
    if (cellTarget) {
        if (editState) {
            finishEdit({ mode: "commit" });
        }

        selectedCell = {
            rowNumber: Number(cellTarget.getAttribute("data-row-number")),
            columnNumber: Number(cellTarget.getAttribute("data-column-number")),
        };
        setSelectedCellLocal(selectedCell);
        return;
    }

    const target = eventTarget.closest<HTMLElement>("[data-action]");
    if (!target) {
        return;
    }

    finishEdit({ mode: "commit" });

    const action = target.getAttribute("data-action");
    switch (action) {
        case "set-sheet":
            vscode.postMessage({
                type: "setSheet",
                sheetKey: target.getAttribute("data-sheet-key")!,
            });
            return;
        case "toggle-search-regex":
            searchOptions = { ...searchOptions, isRegexp: !searchOptions.isRegexp };
            renderApp();
            return;
        case "toggle-search-match-case":
            searchOptions = { ...searchOptions, matchCase: !searchOptions.matchCase };
            renderApp();
            return;
        case "toggle-search-whole-word":
            searchOptions = { ...searchOptions, wholeWord: !searchOptions.wholeWord };
            renderApp();
            return;
        case "search-prev":
            submitSearch("prev");
            return;
        case "search-next":
            submitSearch("next");
            return;
        case "goto-cell":
            submitGoto();
            return;
        case "undo":
            undoPendingEdits();
            return;
        case "redo":
            redoPendingEdits();
            return;
        case "reload":
            vscode.postMessage({ type: "reload" });
            return;
        case "save-edits":
            triggerSave();
            return;
    }
});

document.addEventListener("keydown", (event: KeyboardEvent) => {
    const isTextInputContext = isTextInputTarget(event.target);

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        triggerSave();
        return;
    }

    if (isTextInputContext) {
        return;
    }

    if (!editState && (event.ctrlKey || event.metaKey) && !event.altKey) {
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

    if (editState) {
        return;
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
            moveSelectionByPage(-1);
            return;
        case "PageDown":
            event.preventDefault();
            moveSelectionByPage(1);
            return;
    }
});

document.addEventListener("paste", (event: ClipboardEvent) => {
    if (isTextInputTarget(event.target) || editState || !model || !selectedCell || !model.canEdit) {
        return;
    }

    const text = event.clipboardData?.getData("text/plain");
    if (!text) {
        return;
    }

    event.preventDefault();
    applyPastedGrid(normalizePastedRows(text));
});

document.addEventListener("dblclick", (event: MouseEvent) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget || !model || !model.canEdit) {
        return;
    }

    const cellTarget = eventTarget.closest<HTMLElement>('[data-role="grid-cell"]');
    if (!cellTarget || cellTarget.getAttribute("data-editable") !== "true") {
        return;
    }

    const rowNumber = Number(cellTarget.getAttribute("data-row-number"));
    const columnNumber = Number(cellTarget.getAttribute("data-column-number"));
    const pendingKey = getPendingEditKey(model.activeSheet.key, rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const currentValue = pendingEdit
        ? pendingEdit.value
        : getCellModelValue(rowNumber, columnNumber);

    event.preventDefault();
    setSelectedCellLocal({ rowNumber, columnNumber }, { syncHost: true });
    enterEditMode(cellTarget, rowNumber, columnNumber, currentValue);
});

renderLoading(STRINGS.loading);
vscode.postMessage({ type: "ready" });
