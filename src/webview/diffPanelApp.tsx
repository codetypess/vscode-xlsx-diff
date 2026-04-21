import * as React from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import type {
    CellDiffStatus,
    GridCellView,
    GridRowView,
    RenderModel,
    RowFilterMode,
    SheetTabView,
    WorkbookFileView,
} from "../core/model/types";

interface VsCodeApi {
    postMessage(message: OutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

type Side = "left" | "right";

type OutgoingMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "setFilter"; filter: RowFilterMode }
    | { type: "prevPage" }
    | { type: "nextPage" }
    | { type: "prevDiff" }
    | { type: "nextDiff" }
    | { type: "selectCell"; rowNumber: number; columnNumber: number }
    | {
          type: "saveEdits";
          edits: Array<{
              sheetKey: string;
              side: Side;
              rowNumber: number;
              columnNumber: number;
              value: string;
          }>;
      }
    | { type: "swap" }
    | { type: "reload" };

type IncomingMessage =
    | { type: "loading"; message: string }
    | { type: "error"; message: string }
    | { type: "render"; payload: RenderModel; silent?: boolean; clearPendingEdits?: boolean };

interface CellPosition {
    rowNumber: number;
    columnNumber: number;
    side: Side;
}

interface EditingCell extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingEdit extends CellPosition {
    sheetKey: string;
    value: string;
}

interface PendingSummary {
    sheetKeys: Set<string>;
    rowsBySide: Record<Side, Set<number>>;
    columnsBySide: Record<Side, Set<number>>;
}

interface ScrollState {
    top: number;
    left: number;
}

type ViewState =
    | { kind: "loading"; message: string }
    | { kind: "error"; message: string }
    | {
          kind: "app";
          model: RenderModel;
          revision: number;
          revealSelection: boolean;
          scrollState: ScrollState | null;
      };

const DEFAULT_STRINGS = {
    loading: "Loading XLSX diff...",
    all: "All",
    diffs: "Diffs",
    same: "Same",
    prevDiff: "Prev Diff",
    nextDiff: "Next Diff",
    prevPage: "Prev Page",
    nextPage: "Next Page",
    swap: "Swap",
    reload: "Reload",
    left: "Left",
    right: "Right",
    mergedRangesChanged: "Merged ranges changed",
    noRowsAvailable: "No rows available for this filter.",
    size: "Size",
    modified: "Modified",
    sheet: "Sheet",
    rows: "Rows",
    noRows: "No rows",
    page: "Page",
    filter: "Filter",
    diffCells: "Diff cells",
    diffRows: "Diff rows",
    sameRows: "Same rows",
    visibleRows: "Visible rows",
    readOnly: "Read-only",
    save: "Save",
};

type Strings = typeof DEFAULT_STRINGS;

const STRINGS: Strings =
    ((globalThis as Record<string, unknown>).__XLSX_DIFF_STRINGS__ as Strings | undefined) ??
    DEFAULT_STRINGS;

const vscode = acquireVsCodeApi();
const pendingEdits = new Map<string, PendingEdit>();

let model: RenderModel | null = null;
let selectedCell: CellPosition | null = null;
let editingCell: EditingCell | null = null;
let pendingSelectionReason: "highlighted-diff" | null = null;
let suppressAutoSelection = false;
let isSyncingScroll = false;
let layoutSyncFrame = 0;
let viewRevision = 0;
let setViewState: React.Dispatch<React.SetStateAction<ViewState>> | null = null;

function getPendingEditKey(
    sheetKey: string,
    side: Side,
    rowNumber: number,
    columnNumber: number
): string {
    return `${sheetKey}:${side}:${rowNumber}:${columnNumber}`;
}

function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}

function getColumnDiffTones(rows: GridRowView[]): Map<number, CellDiffStatus> {
    const tones = new Map<number, CellDiffStatus>();

    for (const row of rows) {
        row.cells.forEach((cell, index) => {
            if (cell.status === "equal") {
                return;
            }

            const currentTone = tones.get(index);
            if (currentTone === "modified" || currentTone === "removed") {
                return;
            }

            if (cell.status === "modified" || cell.status === "removed") {
                tones.set(index, cell.status);
                return;
            }

            if (!currentTone) {
                tones.set(index, cell.status);
            }
        });
    }

    return tones;
}

function getDiffToneClass(diffTone: CellDiffStatus | undefined): string {
    return diffTone ? `diff-marker--${diffTone}` : "";
}

function getEffectiveDiffMarkerClass(
    diffTone: CellDiffStatus | "",
    hasPending: boolean
): string | null {
    if (hasPending) {
        return "diff-marker--pending";
    }

    if (diffTone) {
        return getDiffToneClass(diffTone);
    }

    return null;
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

function shouldHighlightCell(cell: GridCellView, side: Side, isHighlighted: boolean): boolean {
    if (!isHighlighted) {
        return false;
    }

    if (cell.status === "modified") {
        return true;
    }

    if (cell.status === "added") {
        return side === "right";
    }

    if (cell.status === "removed") {
        return side === "left";
    }

    return false;
}

function getSideCellClass(cell: GridCellView, side: Side, isHighlighted: boolean): string {
    const classes = ["grid__cell"];

    if (cell.status === "modified") {
        classes.push("grid__cell--modified");
    } else if (cell.status === "added") {
        classes.push(side === "right" ? "grid__cell--added" : "grid__cell--ghost");
    } else if (cell.status === "removed") {
        classes.push(side === "left" ? "grid__cell--removed" : "grid__cell--ghost");
    } else {
        classes.push("grid__cell--equal");
    }

    if (isHighlighted) {
        classes.push("grid__cell--highlighted");
    }

    return classes.join(" ");
}

function getFilterLabel(filter: RowFilterMode): string {
    switch (filter) {
        case "diffs":
            return STRINGS.diffs;
        case "same":
            return STRINGS.same;
        case "all":
        default:
            return STRINGS.all;
    }
}

function getSheetTooltip(sheet: SheetTabView): string {
    return `${sheet.label} · ${sheet.diffCellCount} ${STRINGS.diffCells} · ${sheet.diffRowCount} ${STRINGS.diffRows}`;
}

function getCellView(rowNumber: number, columnNumber: number): GridCellView | null {
    const row = model?.page.rows.find((item) => item.rowNumber === rowNumber);
    return row?.cells[columnNumber - 1] ?? null;
}

function getPreferredSelectionSide(rowNumber: number, columnNumber: number): Side {
    const cell = getCellView(rowNumber, columnNumber);
    return cell?.status === "removed" ? "left" : "right";
}

function getHighlightedDiffSelection(): CellPosition | null {
    if (!model?.page.highlightedDiffCell) {
        return null;
    }

    const { rowNumber, columnNumber } = model.page.highlightedDiffCell;
    return {
        rowNumber,
        columnNumber,
        side: getPreferredSelectionSide(rowNumber, columnNumber),
    };
}

function isCellVisible(cell: CellPosition | null): boolean {
    if (!cell || !model) {
        return false;
    }

    return model.page.rows.some(
        (row) => row.rowNumber === cell.rowNumber && cell.columnNumber <= row.cells.length
    );
}

function prepareSelectionForRender(revealSelection: boolean): boolean {
    let shouldReveal = revealSelection;

    if (pendingSelectionReason === "highlighted-diff") {
        selectedCell = getHighlightedDiffSelection() ?? selectedCell;
        suppressAutoSelection = false;
        shouldReveal = true;
    }

    pendingSelectionReason = null;

    if (!isCellVisible(selectedCell) && !suppressAutoSelection) {
        selectedCell = getHighlightedDiffSelection();
    }

    return shouldReveal;
}

function clampScrollPosition(value: number, maxValue: number): number {
    return Math.max(0, Math.min(value, Math.max(maxValue, 0)));
}

function getPaneScrollState(): ScrollState | null {
    const pane = document.querySelector<HTMLElement>(".pane__table");
    if (!pane) {
        return null;
    }

    return {
        top: pane.scrollTop,
        left: pane.scrollLeft,
    };
}

interface ScrollUpdate {
    pane: HTMLElement;
    top: number;
    left: number;
}

function setPaneScrollPositions(updates: ScrollUpdate[]): void {
    if (updates.length === 0) {
        return;
    }

    isSyncingScroll = true;
    for (const { pane, top, left } of updates) {
        pane.scrollTop = clampScrollPosition(top, pane.scrollHeight - pane.clientHeight);
        pane.scrollLeft = clampScrollPosition(left, pane.scrollWidth - pane.clientWidth);
    }

    requestAnimationFrame(() => {
        isSyncingScroll = false;
    });
}

function restorePaneScrollState(scrollState: ScrollState | null): void {
    if (!scrollState) {
        return;
    }

    setPaneScrollPositions(
        Array.from(document.querySelectorAll<HTMLElement>(".pane__table")).map((pane) => ({
            pane,
            top: scrollState.top,
            left: scrollState.left,
        }))
    );
}

function getStickyPaneInsets(pane: HTMLElement): { top: number; left: number } {
    const headerRow = pane.querySelector("thead tr");
    const firstColumn = pane.querySelector("thead th:first-child");

    return {
        top: headerRow?.getBoundingClientRect().height ?? 0,
        left: firstColumn?.getBoundingClientRect().width ?? 0,
    };
}

function getDesiredPaneScrollPosition(
    pane: HTMLElement,
    element: HTMLElement
): { top: number; left: number } {
    const paneRect = pane.getBoundingClientRect();
    const elementRect = element.getBoundingClientRect();
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

    return { top, left };
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

function revealSelectedCells(): void {
    setPaneScrollPositions(
        getSelectedCellElements()
            .map((element) => {
                const pane = element.closest<HTMLElement>(".pane__table");
                if (!pane) {
                    return null;
                }

                return {
                    pane,
                    ...getDesiredPaneScrollPosition(pane, element),
                };
            })
            .filter((update): update is ScrollUpdate => update !== null)
    );
}

function getGridRows(side: Side): HTMLElement[] {
    return Array.from(
        document.querySelectorAll<HTMLElement>(`.pane[data-side="${side}"] [data-role="grid-row"]`)
    );
}

function syncTableRowHeights(): void {
    const leftRows = getGridRows("left");
    const rightRows = getGridRows("right");
    const rowCount = Math.min(leftRows.length, rightRows.length);

    for (const row of [...leftRows, ...rightRows]) {
        row.style.height = "";
    }

    for (let index = 0; index < rowCount; index += 1) {
        const leftRow = leftRows[index]!;
        const rightRow = rightRows[index]!;
        const syncedHeight = Math.ceil(
            Math.max(
                leftRow.getBoundingClientRect().height,
                rightRow.getBoundingClientRect().height
            )
        );

        if (syncedHeight <= 0) {
            continue;
        }

        leftRow.style.height = `${syncedHeight}px`;
        rightRow.style.height = `${syncedHeight}px`;
    }
}

function scheduleLayoutSync({ revealSelection = false }: { revealSelection?: boolean } = {}): void {
    if (layoutSyncFrame) {
        cancelAnimationFrame(layoutSyncFrame);
    }

    layoutSyncFrame = requestAnimationFrame(() => {
        layoutSyncFrame = 0;
        syncTableRowHeights();

        if (revealSelection) {
            revealSelectedCells();
        }
    });
}

function syncPaneScroll(sourcePane: HTMLElement): void {
    if (isSyncingScroll) {
        return;
    }

    const panes = Array.from(document.querySelectorAll<HTMLElement>(".pane__table"));
    if (panes.length < 2) {
        return;
    }

    isSyncingScroll = true;
    for (const pane of panes) {
        if (pane === sourcePane) {
            continue;
        }

        pane.scrollTop = clampScrollPosition(
            sourcePane.scrollTop,
            pane.scrollHeight - pane.clientHeight
        );
        pane.scrollLeft = clampScrollPosition(
            sourcePane.scrollLeft,
            pane.scrollWidth - pane.clientWidth
        );
    }

    requestAnimationFrame(() => {
        isSyncingScroll = false;
    });
}

function canEditCell(status: CellDiffStatus, side: Side): boolean {
    if (!model) {
        return false;
    }

    const isReadonly = side === "left" ? model.leftFile.isReadonly : model.rightFile.isReadonly;
    if (isReadonly) {
        return false;
    }

    const sheetName = side === "left" ? model.activeSheet.leftName : model.activeSheet.rightName;
    if (!sheetName) {
        return false;
    }

    if (status === "added" && side === "left") {
        return false;
    }

    if (status === "removed" && side === "right") {
        return false;
    }

    return true;
}

function getCellModelValue(rowNumber: number, columnNumber: number, side: Side): string {
    const cell = getCellView(rowNumber, columnNumber);
    if (!cell) {
        return "";
    }

    return side === "left" ? cell.leftValue : cell.rightValue;
}

function getCellFormula(rowNumber: number, columnNumber: number, side: Side): string | null {
    const cell = getCellView(rowNumber, columnNumber);
    if (!cell) {
        return null;
    }

    return side === "left" ? cell.leftFormula : cell.rightFormula;
}

function commitEdit(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    side: Side,
    value: string
): void {
    const key = getPendingEditKey(sheetKey, side, rowNumber, columnNumber);
    const modelValue = getCellModelValue(rowNumber, columnNumber, side);

    if (value === modelValue) {
        pendingEdits.delete(key);
    } else {
        pendingEdits.set(key, { sheetKey, side, rowNumber, columnNumber, value });
    }
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
        commitEdit(
            session.sheetKey,
            session.rowNumber,
            session.columnNumber,
            session.side,
            session.value
        );
    }

    if (clearSelection) {
        selectedCell = null;
        suppressAutoSelection = true;
    }

    if (refresh) {
        renderApp({ commitEditing: false });
    }
}

function clearSelectedCellValue(): void {
    if (!model || !selectedCell) {
        return;
    }

    const { rowNumber, columnNumber, side } = selectedCell;
    const cell = getCellView(rowNumber, columnNumber);
    if (!cell || !canEditCell(cell.status, side)) {
        return;
    }

    commitEdit(model.activeSheet.key, rowNumber, columnNumber, side, "");
    renderApp({ commitEditing: false, revealSelection: true });
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
    if (editingCell) {
        finishEdit({ mode: "commit", clearSelection: true, refresh: false });
    }

    if (pendingEdits.size === 0) {
        renderApp({ commitEditing: false });
        return;
    }

    const edits = Array.from(pendingEdits.values());
    pendingEdits.clear();
    renderApp({ commitEditing: false });
    vscode.postMessage({ type: "saveEdits", edits });
}

function startEditCell(position: CellPosition, value: string): void {
    if (!model) {
        return;
    }

    if (editingCell) {
        finishEdit({ mode: "commit", refresh: false });
    }

    selectedCell = position;
    suppressAutoSelection = false;
    editingCell = {
        ...position,
        sheetKey: model.activeSheet.key,
        value,
    };
    renderApp({ commitEditing: false, revealSelection: true });
}

function getPendingSummary(activeSheetKey: string): PendingSummary {
    const summary: PendingSummary = {
        sheetKeys: new Set<string>(),
        rowsBySide: {
            left: new Set<number>(),
            right: new Set<number>(),
        },
        columnsBySide: {
            left: new Set<number>(),
            right: new Set<number>(),
        },
    };

    for (const pendingEdit of pendingEdits.values()) {
        summary.sheetKeys.add(pendingEdit.sheetKey);

        if (pendingEdit.sheetKey !== activeSheetKey) {
            continue;
        }

        summary.rowsBySide[pendingEdit.side].add(pendingEdit.rowNumber);
        summary.columnsBySide[pendingEdit.side].add(pendingEdit.columnNumber);
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
    updateView({ kind: "loading", message });
}

function renderError(message: string): void {
    updateView({ kind: "error", message });
}

function renderApp({
    commitEditing = true,
    revealSelection = false,
}: { commitEditing?: boolean; revealSelection?: boolean } = {}): void {
    if (!model) {
        renderLoading(STRINGS.loading);
        return;
    }

    if (commitEditing) {
        finishEdit({ mode: "commit", refresh: false });
    }

    const scrollState = getPaneScrollState();
    const shouldRevealSelection = prepareSelectionForRender(revealSelection);
    viewRevision += 1;

    updateView({
        kind: "app",
        model,
        revision: viewRevision,
        revealSelection: shouldRevealSelection,
        scrollState,
    });
}

function DiffMarker({
    markerClass,
    extraClass,
}: {
    markerClass: string | null;
    extraClass?: string;
}): React.ReactElement | null {
    if (!markerClass) {
        return null;
    }

    return <span className={classNames(["diff-marker", markerClass, extraClass])} aria-hidden />;
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
    icon,
    label,
    active = false,
    disabled = false,
    onClick,
}: {
    icon: string;
    label: string;
    active?: boolean;
    disabled?: boolean;
    onClick(): void;
}): React.ReactElement {
    return (
        <button
            className={classNames(["toolbar__button", active && "is-active"])}
            disabled={disabled}
            type="button"
            onClick={onClick}
        >
            <span className={classNames(["codicon", icon, "toolbar__button-icon"])} aria-hidden />
            <span>{label}</span>
        </button>
    );
}

function Toolbar({ currentModel }: { currentModel: RenderModel }): React.ReactElement {
    return (
        <header className="toolbar">
            <div className="toolbar__group">
                <ToolbarButton
                    active={currentModel.filter === "all"}
                    icon="codicon-list-flat"
                    label={STRINGS.all}
                    onClick={() => vscode.postMessage({ type: "setFilter", filter: "all" })}
                />
                <ToolbarButton
                    active={currentModel.filter === "diffs"}
                    icon="codicon-diff-multiple"
                    label={STRINGS.diffs}
                    onClick={() => vscode.postMessage({ type: "setFilter", filter: "diffs" })}
                />
                <ToolbarButton
                    active={currentModel.filter === "same"}
                    icon="codicon-check-all"
                    label={STRINGS.same}
                    onClick={() => vscode.postMessage({ type: "setFilter", filter: "same" })}
                />
            </div>
            <div className="toolbar__group">
                <ToolbarButton
                    disabled={!currentModel.canPrevDiff}
                    icon="codicon-arrow-up"
                    label={STRINGS.prevDiff}
                    onClick={() => {
                        pendingSelectionReason = "highlighted-diff";
                        vscode.postMessage({ type: "prevDiff" });
                    }}
                />
                <ToolbarButton
                    disabled={!currentModel.canNextDiff}
                    icon="codicon-arrow-down"
                    label={STRINGS.nextDiff}
                    onClick={() => {
                        pendingSelectionReason = "highlighted-diff";
                        vscode.postMessage({ type: "nextDiff" });
                    }}
                />
                <ToolbarButton
                    icon="codicon-arrow-swap"
                    label={STRINGS.swap}
                    onClick={() => vscode.postMessage({ type: "swap" })}
                />
                <ToolbarButton
                    icon="codicon-refresh"
                    label={STRINGS.reload}
                    onClick={() => vscode.postMessage({ type: "reload" })}
                />
                <ToolbarButton
                    disabled={pendingEdits.size === 0}
                    icon="codicon-save"
                    label={STRINGS.save}
                    onClick={triggerSave}
                />
            </div>
        </header>
    );
}

function FileCard({ file }: { file: WorkbookFileView }): React.ReactElement {
    return (
        <div className="file-card">
            <div className="file-card__name" title={file.filePath}>
                <span className="file-card__name-text">{file.fileName}</span>
                {file.isReadonly ? (
                    <span
                        className="codicon codicon-lock file-card__lock"
                        title={STRINGS.readOnly}
                        aria-label={STRINGS.readOnly}
                    />
                ) : null}
            </div>
            <div className="file-card__meta">
                <div className="file-card__path" title={file.filePath}>
                    {file.filePath}
                </div>
                <div className="file-card__facts">
                    <span>
                        {STRINGS.size}: {file.fileSizeLabel}
                    </span>
                    {file.detailLabel && file.detailValue ? (
                        <span>
                            {file.detailLabel}: {file.detailValue}
                        </span>
                    ) : null}
                    <span>
                        {STRINGS.modified}: {file.modifiedTimeLabel}
                    </span>
                </div>
            </div>
        </div>
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
    row,
    columnNumber,
    side,
}: {
    cell: GridCellView;
    row: GridRowView;
    columnNumber: number;
    side: Side;
}): React.ReactElement {
    const pendingKey = getPendingEditKey(model!.activeSheet.key, side, row.rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const value = pendingEdit
        ? pendingEdit.value
        : side === "left"
          ? cell.leftValue
          : cell.rightValue;
    const formula = pendingEdit ? null : side === "left" ? cell.leftFormula : cell.rightFormula;
    const highlighted = shouldHighlightCell(cell, side, row.isHighlighted);
    const editable = canEditCell(cell.status, side);
    const isSelected =
        selectedCell?.rowNumber === row.rowNumber && selectedCell.columnNumber === columnNumber;
    const isEditing =
        editingCell?.side === side &&
        editingCell.rowNumber === row.rowNumber &&
        editingCell.columnNumber === columnNumber;
    const cellClass = classNames([
        getSideCellClass(cell, side, highlighted),
        pendingEdit && "grid__cell--pending",
        isSelected && "grid__cell--selected",
        isEditing && "grid__cell--editing",
    ]);

    return (
        <td
            aria-selected={isSelected}
            className={cellClass}
            data-cell-status={cell.status}
            data-column-number={columnNumber}
            data-editable={editable}
            data-role="grid-cell"
            data-row-number={row.rowNumber}
            title={getCellTooltip(cell.address, value, formula)}
            onClick={() => {
                if (editingCell) {
                    finishEdit({ mode: "commit", refresh: false });
                }

                selectedCell = { rowNumber: row.rowNumber, columnNumber, side };
                suppressAutoSelection = false;
                pendingSelectionReason = null;

                if (
                    cell.status !== "equal" &&
                    (model?.page.highlightedDiffCell?.rowNumber !== row.rowNumber ||
                        model?.page.highlightedDiffCell?.columnNumber !== columnNumber)
                ) {
                    vscode.postMessage({
                        type: "selectCell",
                        rowNumber: row.rowNumber,
                        columnNumber,
                    });
                }

                renderApp({ commitEditing: false });
            }}
            onDoubleClick={(event) => {
                if (!editable) {
                    return;
                }

                event.preventDefault();
                startEditCell({ rowNumber: row.rowNumber, columnNumber, side }, value);
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

function DiffTable({
    currentModel,
    side,
    pendingSummary,
}: {
    currentModel: RenderModel;
    side: Side;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    if (currentModel.page.rows.length === 0) {
        return <div className="empty-table">{STRINGS.noRowsAvailable}</div>;
    }

    const diffColumnTones = getColumnDiffTones(currentModel.page.rows);
    const pendingRows = pendingSummary.rowsBySide[side];
    const pendingColumns = pendingSummary.columnsBySide[side];

    return (
        <table className="grid">
            <thead>
                <tr>
                    <th className="grid__row-number">#</th>
                    {currentModel.page.columns.map((column, index) => {
                        const columnNumber = index + 1;
                        const diffTone = diffColumnTones.get(index) ?? "";
                        const hasPending = pendingColumns.has(columnNumber);
                        const markerClass = getEffectiveDiffMarkerClass(diffTone, hasPending);

                        return (
                            <th
                                key={columnNumber}
                                className={classNames([
                                    "grid__column",
                                    (diffTone || hasPending) && "grid__column--diff",
                                    diffTone && `grid__column--${diffTone}`,
                                    hasPending && "grid__column--pending",
                                ])}
                                data-column-number={columnNumber}
                                data-diff-tone={diffTone}
                            >
                                <span className="grid__column-label">
                                    <DiffMarker markerClass={markerClass} />
                                    <span>{column}</span>
                                </span>
                            </th>
                        );
                    })}
                </tr>
            </thead>
            <tbody>
                {currentModel.page.rows.map((row) => {
                    const hasPending = pendingRows.has(row.rowNumber);
                    const diffTone = row.hasDiff ? row.diffTone : "";
                    const markerClass = getEffectiveDiffMarkerClass(diffTone, hasPending);

                    return (
                        <tr
                            key={row.rowNumber}
                            className={classNames([
                                row.hasDiff && "row--diff",
                                row.hasDiff && `row--diff-${row.diffTone}`,
                                row.isHighlighted && "row--highlight",
                            ])}
                            data-role="grid-row"
                            data-row-has-diff={row.hasDiff}
                            data-row-number={row.rowNumber}
                        >
                            <th
                                className={classNames([
                                    "grid__row-number",
                                    row.hasDiff && "grid__row-number--diff",
                                    row.hasDiff && `grid__row-number--${row.diffTone}`,
                                    hasPending && "grid__row-number--pending",
                                ])}
                                data-diff-tone={diffTone}
                            >
                                <span className="grid__row-label">
                                    <DiffMarker markerClass={markerClass} />
                                    <span>{row.rowNumber}</span>
                                </span>
                            </th>
                            {row.cells.map((cell, index) => (
                                <GridCell
                                    key={`${row.rowNumber}:${index + 1}`}
                                    cell={cell}
                                    columnNumber={index + 1}
                                    row={row}
                                    side={side}
                                />
                            ))}
                        </tr>
                    );
                })}
            </tbody>
        </table>
    );
}

function Pane({
    currentModel,
    pendingSummary,
    side,
    title,
}: {
    currentModel: RenderModel;
    pendingSummary: PendingSummary;
    side: Side;
    title: string;
}): React.ReactElement {
    return (
        <section className="pane" data-side={side}>
            <div className="pane__header">
                <div className="pane__title">{title}</div>
                {currentModel.activeSheet.mergedRangesChanged ? (
                    <span className="badge badge--warn">{STRINGS.mergedRangesChanged}</span>
                ) : null}
            </div>
            <div
                className="pane__table"
                data-side={side}
                onScroll={(event) => syncPaneScroll(event.currentTarget)}
            >
                <DiffTable
                    currentModel={currentModel}
                    pendingSummary={pendingSummary}
                    side={side}
                />
            </div>
        </section>
    );
}

function Tabs({
    currentModel,
    pendingSummary,
}: {
    currentModel: RenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    return (
        <div className="tabs">
            {currentModel.sheets.map((sheet) => {
                const hasPending = pendingSummary.sheetKeys.has(sheet.key);
                const diffTone = sheet.hasDiff ? sheet.diffTone : "";
                const markerClass = getEffectiveDiffMarkerClass(diffTone, hasPending);

                return (
                    <button
                        key={sheet.key}
                        className={classNames([
                            "tab",
                            `tab--${sheet.diffTone}`,
                            sheet.isActive && "is-active",
                            sheet.hasDiff && "has-diff",
                        ])}
                        data-diff-tone={diffTone}
                        title={getSheetTooltip(sheet)}
                        type="button"
                        onClick={() =>
                            vscode.postMessage({ type: "setSheet", sheetKey: sheet.key })
                        }
                    >
                        <DiffMarker extraClass="tab__marker" markerClass={markerClass} />
                        <span className="tab__label">{sheet.label}</span>
                    </button>
                );
            })}
        </div>
    );
}

function Status({
    currentModel,
    pendingSummary,
}: {
    currentModel: RenderModel;
    pendingSummary: PendingSummary;
}): React.ReactElement {
    const rowRangeLabel =
        currentModel.page.visibleRowCount === 0 ? STRINGS.noRows : currentModel.page.rangeLabel;

    return (
        <footer className="footer">
            <Tabs currentModel={currentModel} pendingSummary={pendingSummary} />
            <div className="status">
                <span>
                    <strong>{STRINGS.sheet}:</strong> {currentModel.activeSheet.label}
                </span>
                <span>
                    <strong>{STRINGS.rows}:</strong> {rowRangeLabel}
                </span>
                <span>
                    <strong>{STRINGS.page}:</strong> {currentModel.page.currentPage} /{" "}
                    {currentModel.page.totalPages}
                </span>
                <span>
                    <strong>{STRINGS.filter}:</strong> {getFilterLabel(currentModel.filter)}
                </span>
                <span>
                    <strong>{STRINGS.diffRows}:</strong> {currentModel.page.diffRowCount}
                </span>
                <span>
                    <strong>{STRINGS.sameRows}:</strong> {currentModel.page.sameRowCount}
                </span>
                <span>
                    <strong>{STRINGS.visibleRows}:</strong> {currentModel.page.visibleRowCount}
                </span>
            </div>
        </footer>
    );
}

function DiffApp({ view }: { view: Extract<ViewState, { kind: "app" }> }): React.ReactElement {
    const pendingSummary = getPendingSummary(view.model.activeSheet.key);

    React.useLayoutEffect(() => {
        restorePaneScrollState(view.scrollState);
        scheduleLayoutSync({ revealSelection: view.revealSelection });
    }, [view.revision, view.scrollState, view.revealSelection]);

    React.useEffect(() => {
        const handleResize = () => {
            if (model) {
                scheduleLayoutSync();
            }
        };

        window.addEventListener("resize", handleResize);
        return () => window.removeEventListener("resize", handleResize);
    }, []);

    return (
        <div className="app">
            <Toolbar currentModel={view.model} />
            <section className="files">
                <FileCard file={view.model.leftFile} />
                <FileCard file={view.model.rightFile} />
            </section>
            <section className="panes">
                <Pane
                    currentModel={view.model}
                    pendingSummary={pendingSummary}
                    side="left"
                    title={STRINGS.left}
                />
                <Pane
                    currentModel={view.model}
                    pendingSummary={pendingSummary}
                    side="right"
                    title={STRINGS.right}
                />
            </section>
            <Status currentModel={view.model} pendingSummary={pendingSummary} />
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
        return <DiffApp view={view} />;
    }

    return <Shell kind={view.kind} message={view.message} />;
}

window.addEventListener("message", (event: MessageEvent<IncomingMessage>) => {
    const message = event.data;

    if (message.type === "loading") {
        renderLoading(message.message);
        return;
    }

    if (message.type === "error") {
        renderError(message.message);
        return;
    }

    if (message.type === "render") {
        model = message.payload;
        if (message.clearPendingEdits) {
            pendingEdits.clear();
            editingCell = null;
        }

        renderApp({ revealSelection: !message.silent });
    }
});

document.addEventListener("keydown", (event: KeyboardEvent) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        triggerSave();
        return;
    }

    if (isTextInputTarget(event.target)) {
        return;
    }

    if (isClearSelectedCellKey(event) && !editingCell && selectedCell) {
        event.preventDefault();
        clearSelectedCellValue();
    }
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
