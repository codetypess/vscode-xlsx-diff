import * as React from "react";
import { createRoot } from "react-dom/client";
import type { CellDiffStatus } from "../core/model/types";
import type {
    DiffPanelRenderModel,
    DiffPanelRowView,
    DiffPanelSheetTabView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "./diffPanelTypes";

interface VsCodeApi {
    postMessage(message: OutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

type Side = "left" | "right";
type FilterMode = "all" | "diffs" | "same";

type OutgoingMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
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
    | { type: "render"; payload: DiffPanelRenderModel; clearPendingEdits?: boolean };

interface CellSelection {
    rowNumber: number;
    columnNumber: number;
    side: Side;
}

interface EditingCell extends CellSelection {
    sheetKey: string;
    value: string;
    modelValue: string;
}

interface PendingEdit extends CellSelection {
    sheetKey: string;
    value: string;
}

interface PendingSummary {
    sheetKeys: Set<string>;
    rowsBySide: Record<Side, Set<number>>;
    columnsBySide: Record<Side, Set<number>>;
}

interface SheetRuntime {
    rowByNumber: Map<number, DiffPanelRowView>;
    diffRowIndexByNumber: Map<number, number>;
    sameRows: number[];
    sameRowIndexByNumber: Map<number, number>;
    diffCellIndexByKey: Map<string, number>;
    columnDiffTones: Array<CellDiffStatus | null>;
}

type ViewState =
    | { kind: "loading"; message: string }
    | { kind: "error"; message: string }
    | { kind: "app"; model: DiffPanelRenderModel };

type DiffMarkerTone = CellDiffStatus | "pending" | null;

const ROW_HEIGHT = 27;
const ROW_OVERSCAN = 8;
const ROW_HEADER_WIDTH = 56;
const COLUMN_WIDTH = 120;

const DEFAULT_STRINGS = {
    loading: "Loading XLSX diff...",
    reload: "Reload",
    swap: "Swap",
    all: "All",
    diffs: "Diffs",
    same: "Same",
    allRows: "All Rows",
    diffRows: "Diff Rows",
    sameRows: "Same Rows",
    prevDiff: "Prev Diff",
    nextDiff: "Next Diff",
    sheets: "Sheets",
    diffCells: "Diff Cells",
    diffRowsShort: "Diff Rows",
    rows: "Rows",
    columns: "Columns",
    filter: "Filter",
    visibleRows: "Visible Rows",
    currentDiff: "Current Diff",
    selected: "Selected",
    save: "Save",
    none: "-",
    modified: "Modified",
    size: "Size",
    readOnly: "Read-only",
    mergedRangesChanged: "Merged ranges changed",
    noSheet: "No sheet is available.",
    noRows: "No rows are available in this sheet.",
};

type Strings = typeof DEFAULT_STRINGS;

const STRINGS: Strings =
    ((globalThis as Record<string, unknown>).__XLSX_DIFF_STRINGS__ as Strings | undefined) ??
    DEFAULT_STRINGS;

const vscode = acquireVsCodeApi();

function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}

function getPendingEditKey(
    sheetKey: string,
    side: Side,
    rowNumber: number,
    columnNumber: number
): string {
    return `${sheetKey}:${side}:${rowNumber}:${columnNumber}`;
}

function getDiffTonePriority(status: CellDiffStatus | null): number {
    switch (status) {
        case "modified":
            return 3;
        case "removed":
            return 2;
        case "added":
            return 1;
        case "equal":
        case null:
        default:
            return 0;
    }
}

function mergeDiffTone(
    current: CellDiffStatus | null,
    next: CellDiffStatus
): CellDiffStatus | null {
    if (next === "equal") {
        return current;
    }

    return getDiffTonePriority(next) > getDiffTonePriority(current) ? next : current;
}

function createSheetRuntime(sheet: DiffPanelSheetView): SheetRuntime {
    const diffRowsSet = new Set(sheet.diffRows);
    const sameRows: number[] = [];
    const columnDiffTones = Array<CellDiffStatus | null>(sheet.columnCount).fill(null);
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        if (!diffRowsSet.has(rowNumber)) {
            sameRows.push(rowNumber);
        }
    }

    for (const row of sheet.rows) {
        for (const cell of row.cells) {
            const columnIndex = cell.columnNumber - 1;
            columnDiffTones[columnIndex] = mergeDiffTone(columnDiffTones[columnIndex], cell.status);
        }
    }

    return {
        rowByNumber: new Map(sheet.rows.map((row) => [row.rowNumber, row] as const)),
        diffRowIndexByNumber: new Map(
            sheet.diffRows.map((rowNumber, index) => [rowNumber, index] as const)
        ),
        sameRows,
        sameRowIndexByNumber: new Map(
            sameRows.map((rowNumber, index) => [rowNumber, index] as const)
        ),
        diffCellIndexByKey: new Map(
            sheet.diffCells.map((cell, index) => [cell.key, index] as const)
        ),
        columnDiffTones,
    };
}

function getSelectedDiffIndex(
    selection: CellSelection | null,
    row: DiffPanelRowView | null,
    runtime: SheetRuntime | null
): number | null {
    if (!selection || !row || !runtime) {
        return null;
    }

    const cell = row.cells.find((item) => item.columnNumber === selection.columnNumber);
    if (!cell) {
        return null;
    }

    return cell.diffIndex === null
        ? null
        : (runtime.diffCellIndexByKey.get(cell.key) ?? cell.diffIndex);
}

function getPreferredSide(cell: DiffPanelSparseCellView | null): Side {
    return cell?.status === "removed" ? "left" : "right";
}

function getInitialSelection(
    sheet: DiffPanelSheetView | null,
    runtime: SheetRuntime | null
): CellSelection | null {
    if (!sheet || sheet.columnCount === 0 || sheet.rowCount === 0) {
        return null;
    }

    const firstDiffCell = sheet.diffCells[0];
    if (firstDiffCell) {
        const row = runtime?.rowByNumber.get(firstDiffCell.rowNumber) ?? null;
        const cell =
            row?.cells.find((item) => item.columnNumber === firstDiffCell.columnNumber) ?? null;

        return {
            rowNumber: firstDiffCell.rowNumber,
            columnNumber: firstDiffCell.columnNumber,
            side: getPreferredSide(cell),
        };
    }

    return {
        rowNumber: 1,
        columnNumber: 1,
        side: "right",
    };
}

function getFilteredRowCount(sheet: DiffPanelSheetView, filter: FilterMode): number {
    switch (filter) {
        case "diffs":
            return sheet.diffRows.length;
        case "same":
            return sheet.rowCount - sheet.diffRows.length;
        case "all":
        default:
            return sheet.rowCount;
    }
}

function getRowNumberAtIndex(
    sheet: DiffPanelSheetView,
    filter: FilterMode,
    runtime: SheetRuntime | null,
    index: number
): number | null {
    if (filter === "diffs") {
        return sheet.diffRows[index] ?? null;
    }

    if (filter === "same") {
        return runtime?.sameRows[index] ?? null;
    }

    const rowNumber = index + 1;
    return rowNumber <= sheet.rowCount ? rowNumber : null;
}

function getRowIndex(
    sheet: DiffPanelSheetView,
    filter: FilterMode,
    runtime: SheetRuntime | null,
    rowNumber: number
): number {
    if (filter === "diffs") {
        return runtime?.diffRowIndexByNumber.get(rowNumber) ?? -1;
    }

    if (filter === "same") {
        return runtime?.sameRowIndexByNumber.get(rowNumber) ?? -1;
    }

    return rowNumber > 0 && rowNumber <= sheet.rowCount ? rowNumber - 1 : -1;
}

function getDefaultSelectionForFilter(
    sheet: DiffPanelSheetView,
    runtime: SheetRuntime | null,
    filter: FilterMode
): CellSelection | null {
    if (filter === "diffs") {
        const firstDiffCell = sheet.diffCells[0];
        if (!firstDiffCell) {
            return null;
        }

        const row = runtime?.rowByNumber.get(firstDiffCell.rowNumber) ?? null;
        const cell =
            row?.cells.find((item) => item.columnNumber === firstDiffCell.columnNumber) ?? null;

        return {
            rowNumber: firstDiffCell.rowNumber,
            columnNumber: firstDiffCell.columnNumber,
            side: getPreferredSide(cell),
        };
    }

    if (filter === "same") {
        const firstSameRowNumber = runtime?.sameRows[0];
        if (!firstSameRowNumber) {
            return null;
        }

        return {
            rowNumber: firstSameRowNumber,
            columnNumber: 1,
            side: "right",
        };
    }

    return getInitialSelection(sheet, runtime);
}

function getFilterLabel(filter: FilterMode): string {
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

function DiffMarker({
    tone,
    className,
}: {
    tone: DiffMarkerTone;
    className?: string;
}): React.JSX.Element | null {
    if (!tone || tone === "equal") {
        return null;
    }

    return (
        <span
            className={classNames(["diff-diffMarker", `diff-diffMarker--${tone}`, className])}
            aria-hidden="true"
        />
    );
}

function getSheetTooltip(sheet: DiffPanelSheetTabView): string {
    return `${sheet.label} · ${sheet.diffCellCount} ${STRINGS.diffCells} · ${sheet.diffRowCount} ${STRINGS.diffRows}`;
}

function getCellTitle(
    columnLabel: string,
    rowNumber: number,
    value: string,
    formula: string | null
): string {
    const lines = [`${columnLabel}${rowNumber}`];

    if (value) {
        lines.push(value);
    }

    if (formula) {
        lines.push(`fx ${formula}`);
    }

    return lines.join("\n");
}

function getCellDisplay(
    cell: DiffPanelSparseCellView | null,
    side: Side
): { present: boolean; value: string; formula: string | null } {
    if (!cell) {
        return {
            present: false,
            value: "",
            formula: null,
        };
    }

    return side === "left"
        ? {
              present: cell.leftPresent,
              value: cell.leftValue,
              formula: cell.leftFormula,
          }
        : {
              present: cell.rightPresent,
              value: cell.rightValue,
              formula: cell.rightFormula,
          };
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

function canEditCell(
    model: DiffPanelRenderModel,
    sheet: DiffPanelSheetView,
    side: Side,
    status: CellDiffStatus
): boolean {
    const file = side === "left" ? model.leftFile : model.rightFile;
    if (file.isReadonly) {
        return false;
    }

    const sheetName = side === "left" ? sheet.leftName : sheet.rightName;
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

function applyPendingEdit(
    pendingEdits: ReadonlyMap<string, PendingEdit>,
    edit: EditingCell
): Map<string, PendingEdit> {
    const nextPendingEdits = new Map(pendingEdits);
    const key = getPendingEditKey(edit.sheetKey, edit.side, edit.rowNumber, edit.columnNumber);

    if (edit.value === edit.modelValue) {
        nextPendingEdits.delete(key);
    } else {
        nextPendingEdits.set(key, {
            sheetKey: edit.sheetKey,
            side: edit.side,
            rowNumber: edit.rowNumber,
            columnNumber: edit.columnNumber,
            value: edit.value,
        });
    }

    return nextPendingEdits;
}

function createPendingSummary(): PendingSummary {
    return {
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
}

function getPendingSummary(
    activeSheetKey: string | null,
    pendingEdits: ReadonlyMap<string, PendingEdit>
): PendingSummary {
    const summary = createPendingSummary();

    for (const pendingEdit of pendingEdits.values()) {
        summary.sheetKeys.add(pendingEdit.sheetKey);

        if (!activeSheetKey || pendingEdit.sheetKey !== activeSheetKey) {
            continue;
        }

        summary.rowsBySide[pendingEdit.side].add(pendingEdit.rowNumber);
        summary.columnsBySide[pendingEdit.side].add(pendingEdit.columnNumber);
    }

    return summary;
}

function getEffectiveMarkerTone(
    diffTone: CellDiffStatus | null,
    hasPending: boolean
): DiffMarkerTone {
    if (hasPending) {
        return "pending";
    }

    if (!diffTone || diffTone === "equal") {
        return null;
    }

    return diffTone;
}

function PaneHeader({
    columns,
    columnDiffTones,
    pendingColumns,
    scrollLeft,
    selectedColumnNumber,
    viewportRef,
}: {
    columns: string[];
    columnDiffTones: Array<CellDiffStatus | null>;
    pendingColumns: ReadonlySet<number>;
    scrollLeft: number;
    selectedColumnNumber: number | null;
    viewportRef: React.RefObject<HTMLDivElement | null>;
}): React.JSX.Element {
    return (
        <div className="diff-paneHeader">
            <div className="diff-indexCell diff-indexCell--header">
                <span className="diff-indexLabel">
                    <span>#</span>
                </span>
            </div>
            <div ref={viewportRef} className="diff-columnsViewport">
                <div
                    className="diff-columnsTrack"
                    style={{
                        width: columns.length * COLUMN_WIDTH,
                        transform: `translateX(-${scrollLeft}px)`,
                    }}
                >
                    {columns.map((column, index) => {
                        const columnNumber = index + 1;
                        const diffTone = columnDiffTones[index] ?? null;
                        const hasPending = pendingColumns.has(columnNumber);
                        const markerTone = getEffectiveMarkerTone(diffTone, hasPending);

                        return (
                            <div
                                key={column}
                                className={classNames([
                                    "diff-headerCell",
                                    diffTone && "diff-headerCell--diff",
                                    diffTone && `diff-headerCell--${diffTone}`,
                                    hasPending && "diff-headerCell--pending",
                                    selectedColumnNumber === columnNumber &&
                                        "diff-headerCell--active",
                                ])}
                            >
                                <span className="diff-headerLabel">
                                    <DiffMarker tone={markerTone} />
                                    <span>{column}</span>
                                </span>
                            </div>
                        );
                    })}
                </div>
            </div>
        </div>
    );
}

function PaneMeta({ file }: { file: DiffPanelRenderModel["leftFile"] }): React.JSX.Element {
    return (
        <div className="diff-paneMeta">
            <div className="diff-paneMeta__name" title={file.path}>
                <span className="diff-paneMeta__nameText">{file.title}</span>
                {file.isReadonly ? (
                    <span
                        className="codicon codicon-lock diff-paneMeta__lock"
                        title={STRINGS.readOnly}
                        aria-label={STRINGS.readOnly}
                    />
                ) : null}
            </div>
            <div className="diff-paneMeta__meta">
                <div className="diff-paneMeta__path" title={file.path}>
                    {file.path}
                </div>
                <div className="diff-paneMeta__facts">
                    <span>
                        {STRINGS.size}: {file.sizeLabel}
                    </span>
                    {file.detailLabel && file.detailValue ? (
                        <span>
                            {file.detailLabel}: {file.detailValue}
                        </span>
                    ) : null}
                    <span>
                        {STRINGS.modified}: {file.modifiedLabel}
                    </span>
                </div>
            </div>
        </div>
    );
}

function PaneScrollbar({
    scrollbarRef,
    enabled,
    columnTrackWidth,
    onScroll,
}: {
    scrollbarRef: React.RefObject<HTMLDivElement | null>;
    enabled: boolean;
    columnTrackWidth: number;
    onScroll: (event: React.UIEvent<HTMLDivElement>) => void;
}): React.JSX.Element {
    return (
        <div className={classNames(["diff-scrollbarRow", !enabled && "diff-scrollbarRow--inactive"])}>
            <div className="diff-scrollbarSpacer" />
            <div
                ref={enabled ? scrollbarRef : undefined}
                className={classNames(["diff-scrollbar", !enabled && "diff-scrollbar--inactive"])}
                onScroll={enabled ? onScroll : undefined}
            >
                {enabled ? (
                    <div
                        style={{
                            width: columnTrackWidth,
                            height: 1,
                        }}
                    />
                ) : null}
            </div>
        </div>
    );
}

function CellEditor({
    edit,
    onChange,
    onCommit,
    onCancel,
}: {
    edit: EditingCell;
    onChange: (value: string) => void;
    onCommit: () => void;
    onCancel: () => void;
}): React.JSX.Element {
    const inputRef = React.useRef<HTMLInputElement | null>(null);

    React.useLayoutEffect(() => {
        inputRef.current?.focus();
        inputRef.current?.select();
    }, []);

    return (
        <input
            ref={inputRef}
            className="diff-cell__input"
            type="text"
            value={edit.value}
            onBlur={() => {
                setTimeout(() => onCommit(), 0);
            }}
            onChange={(event) => {
                onChange(event.currentTarget.value);
            }}
            onClick={(event) => event.stopPropagation()}
            onDoubleClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === "Tab") {
                    event.preventDefault();
                    onCommit();
                } else if (event.key === "Escape") {
                    event.preventDefault();
                    onCancel();
                }
            }}
        />
    );
}

function SideRow({
    model,
    activeSheet,
    side,
    rowNumber,
    row,
    columns,
    scrollLeft,
    selectedCell,
    pendingEdits,
    editingCell,
    isPendingRow,
    onSelect,
    onStartEdit,
    onEditingChange,
    onCommitEdit,
    onCancelEdit,
}: {
    model: DiffPanelRenderModel;
    activeSheet: DiffPanelSheetView;
    side: Side;
    rowNumber: number;
    row: DiffPanelRowView | null;
    columns: string[];
    scrollLeft: number;
    selectedCell: CellSelection | null;
    pendingEdits: ReadonlyMap<string, PendingEdit>;
    editingCell: EditingCell | null;
    isPendingRow: boolean;
    onSelect: (selection: CellSelection | null) => void;
    onStartEdit: (selection: CellSelection, initialValue: string, modelValue: string) => void;
    onEditingChange: (value: string) => void;
    onCommitEdit: () => void;
    onCancelEdit: () => void;
}): React.JSX.Element {
    const sparseByColumn = new Map(
        row?.cells.map((cell) => [cell.columnNumber, cell] as const) ?? []
    );
    const isRowSelected = selectedCell?.rowNumber === rowNumber;
    const rowDiffTone = row?.hasDiff ? row.diffTone : null;
    const rowMarkerTone = getEffectiveMarkerTone(rowDiffTone, isPendingRow);

    return (
        <div className="diff-sideRow">
            <div
                className={classNames([
                    "diff-indexCell",
                    rowDiffTone && "diff-indexCell--diff",
                    rowDiffTone && `diff-indexCell--${rowDiffTone}`,
                    isPendingRow && "diff-indexCell--pending",
                    isRowSelected && "diff-indexCell--selected",
                ])}
            >
                <span className="diff-indexLabel">
                    <DiffMarker tone={rowMarkerTone} />
                    <span>{rowNumber}</span>
                </span>
            </div>
            <div className="diff-rowCellsViewport">
                <div
                    className="diff-rowCellsTrack"
                    style={{
                        width: columns.length * COLUMN_WIDTH,
                        transform: `translateX(-${scrollLeft}px)`,
                    }}
                >
                    {columns.map((columnLabel, index) => {
                        const columnNumber = index + 1;
                        const cell = sparseByColumn.get(columnNumber) ?? null;
                        const modelDisplay = getCellDisplay(cell, side);
                        const pendingKey = getPendingEditKey(
                            activeSheet.key,
                            side,
                            rowNumber,
                            columnNumber
                        );
                        const pendingEdit = pendingEdits.get(pendingKey);
                        const display = pendingEdit
                            ? {
                                  present: true,
                                  value: pendingEdit.value,
                                  formula: null,
                              }
                            : modelDisplay;
                        const isActive =
                            selectedCell?.rowNumber === rowNumber &&
                            selectedCell.columnNumber === columnNumber &&
                            selectedCell.side === side;
                        const status = cell?.status ?? "equal";
                        const editable = canEditCell(model, activeSheet, side, status);
                        const isEditing =
                            editingCell?.sheetKey === activeSheet.key &&
                            editingCell.side === side &&
                            editingCell.rowNumber === rowNumber &&
                            editingCell.columnNumber === columnNumber;
                        const isGhost =
                            status !== "equal" &&
                            (!display.present || (!display.value && !display.formula));

                        return (
                            <div
                                key={`${side}:${rowNumber}:${columnNumber}`}
                                className={classNames([
                                    "diff-cell",
                                    `diff-cell--${status}`,
                                    isGhost && "diff-cell--ghost",
                                    pendingEdit && "diff-cell--pending",
                                    isActive && "diff-cell--active",
                                    isEditing && "diff-cell--editing",
                                ])}
                                title={getCellTitle(
                                    columnLabel,
                                    rowNumber,
                                    display.value,
                                    display.formula
                                )}
                                onClick={() => {
                                    onSelect({
                                        rowNumber,
                                        columnNumber,
                                        side,
                                    });
                                }}
                                onDoubleClick={(event) => {
                                    if (!editable) {
                                        return;
                                    }

                                    event.preventDefault();
                                    onStartEdit(
                                        {
                                            rowNumber,
                                            columnNumber,
                                            side,
                                        },
                                        display.value,
                                        modelDisplay.value
                                    );
                                }}
                                role="button"
                                tabIndex={0}
                            >
                                {isEditing && editingCell ? (
                                    <CellEditor
                                        edit={editingCell}
                                        onChange={onEditingChange}
                                        onCommit={onCommitEdit}
                                        onCancel={onCancelEdit}
                                    />
                                ) : (
                                    <span className="diff-cell__text">{display.value}</span>
                                )}
                            </div>
                        );
                    })}
                </div>
            </div>
        </div>
    );
}

function EmptyState({ message }: { message: string }): React.JSX.Element {
    return (
        <div className="diff-emptyState">
            <div className="diff-emptyState__message">{message}</div>
        </div>
    );
}

function App(): React.JSX.Element {
    const [viewState, setViewState] = React.useState<ViewState>({
        kind: "loading",
        message: STRINGS.loading,
    });
    const [filter, setFilter] = React.useState<FilterMode>("all");
    const [selectedCell, setSelectedCell] = React.useState<CellSelection | null>(null);
    const [editingCell, setEditingCell] = React.useState<EditingCell | null>(null);
    const [pendingEdits, setPendingEdits] = React.useState<Map<string, PendingEdit>>(
        () => new Map()
    );
    const [activeDiffIndex, setActiveDiffIndex] = React.useState(0);
    const [scrollTop, setScrollTop] = React.useState(0);
    const [viewportHeight, setViewportHeight] = React.useState(480);
    const [horizontalScrollLeft, setHorizontalScrollLeft] = React.useState(0);
    const [leftHeaderViewportWidth, setLeftHeaderViewportWidth] = React.useState(0);
    const [rightHeaderViewportWidth, setRightHeaderViewportWidth] = React.useState(0);

    const viewportRef = React.useRef<HTMLDivElement | null>(null);
    const leftHeaderViewportRef = React.useRef<HTMLDivElement | null>(null);
    const rightHeaderViewportRef = React.useRef<HTMLDivElement | null>(null);
    const leftScrollbarRef = React.useRef<HTMLDivElement | null>(null);
    const rightScrollbarRef = React.useRef<HTMLDivElement | null>(null);
    const scrollFrameRef = React.useRef(0);
    const horizontalScrollSyncFrameRef = React.useRef(0);
    const isSyncingHorizontalScrollRef = React.useRef(false);
    const previousSheetViewKeyRef = React.useRef<string | null>(null);
    React.useEffect(() => {
        const handleMessage = (event: MessageEvent<IncomingMessage>) => {
            const message = event.data;
            if (!message) {
                return;
            }

            if (message.type === "loading") {
                React.startTransition(() => {
                    setViewState({
                        kind: "loading",
                        message: message.message,
                    });
                });
                return;
            }

            if (message.type === "error") {
                React.startTransition(() => {
                    setViewState({
                        kind: "error",
                        message: message.message,
                    });
                });
                return;
            }

            if (message.type === "render") {
                React.startTransition(() => {
                    setViewState({
                        kind: "app",
                        model: message.payload,
                    });
                });

                if (message.clearPendingEdits) {
                    setPendingEdits(new Map());
                    setEditingCell(null);
                }
            }
        };

        window.addEventListener("message", handleMessage);
        vscode.postMessage({ type: "ready" });

        return () => {
            window.removeEventListener("message", handleMessage);
        };
    }, []);

    React.useEffect(() => {
        const element = viewportRef.current;
        if (!element) {
            return;
        }

        const updateHeight = () => {
            setViewportHeight(element.clientHeight);
        };

        updateHeight();

        const observer = new ResizeObserver(() => {
            updateHeight();
        });
        observer.observe(element);

        return () => {
            observer.disconnect();
        };
    }, []);

    React.useEffect(() => {
        return () => {
            if (scrollFrameRef.current) {
                cancelAnimationFrame(scrollFrameRef.current);
            }
            if (horizontalScrollSyncFrameRef.current) {
                cancelAnimationFrame(horizontalScrollSyncFrameRef.current);
            }
        };
    }, []);

    const model = viewState.kind === "app" ? viewState.model : null;
    const activeSheet = model?.activeSheet ?? null;
    const activeSheetViewKey =
        model && activeSheet
            ? `${model.leftFile.path}::${model.rightFile.path}::${activeSheet.key}`
            : null;

    React.useEffect(() => {
        const observers: ResizeObserver[] = [];
        const observeWidth = (
            element: HTMLDivElement | null,
            setWidth: React.Dispatch<React.SetStateAction<number>>
        ) => {
            if (!element) {
                setWidth(0);
                return;
            }

            const updateWidth = () => {
                setWidth(element.clientWidth);
            };

            updateWidth();

            const observer = new ResizeObserver(() => {
                updateWidth();
            });
            observer.observe(element);
            observers.push(observer);
        };

        observeWidth(leftHeaderViewportRef.current, setLeftHeaderViewportWidth);
        observeWidth(rightHeaderViewportRef.current, setRightHeaderViewportWidth);

        return () => {
            observers.forEach((observer) => observer.disconnect());
        };
    }, [activeSheet?.key]);

    const runtime = React.useMemo(
        () => (activeSheet ? createSheetRuntime(activeSheet) : null),
        [activeSheet]
    );
    const totalRowCount = activeSheet ? getFilteredRowCount(activeSheet, filter) : 0;
    const startIndex = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - ROW_OVERSCAN);
    const visibleRowCount =
        Math.ceil(Math.max(viewportHeight, ROW_HEIGHT) / ROW_HEIGHT) + ROW_OVERSCAN * 2;
    const endIndex = Math.min(totalRowCount, startIndex + visibleRowCount);
    const offsetY = startIndex * ROW_HEIGHT;
    const totalHeight = totalRowCount * ROW_HEIGHT;
    const viewportContentHeight = Math.max(totalHeight, viewportHeight);
    const columnTrackWidth = activeSheet ? activeSheet.columnCount * COLUMN_WIDTH : 0;
    const leftHasHorizontalOverflow =
        leftHeaderViewportWidth > 0 && columnTrackWidth > leftHeaderViewportWidth;
    const rightHasHorizontalOverflow =
        rightHeaderViewportWidth > 0 && columnTrackWidth > rightHeaderViewportWidth;
    const hasHorizontalOverflow = leftHasHorizontalOverflow || rightHasHorizontalOverflow;
    const visibleRows: Array<{ rowNumber: number; row: DiffPanelRowView | null }> = [];

    for (let index = startIndex; index < endIndex; index += 1) {
        const rowNumber = activeSheet
            ? getRowNumberAtIndex(activeSheet, filter, runtime, index)
            : null;
        if (rowNumber === null) {
            continue;
        }

        visibleRows.push({
            rowNumber,
            row: runtime?.rowByNumber.get(rowNumber) ?? null,
        });
    }

    React.useEffect(() => {
        if (!activeSheet || !activeSheetViewKey) {
            if (viewState.kind === "app") {
                previousSheetViewKeyRef.current = null;
                setSelectedCell(null);
                setActiveDiffIndex(0);
            }
            return;
        }

        const previousSheetViewKey = previousSheetViewKeyRef.current;
        previousSheetViewKeyRef.current = activeSheetViewKey;

        if (previousSheetViewKey === activeSheetViewKey) {
            return;
        }

        const initialSelection = getInitialSelection(activeSheet, runtime);
        setSelectedCell(initialSelection);
        setEditingCell(null);
        setActiveDiffIndex(0);
        setFilter("all");
        setScrollTop(0);
        setHorizontalScrollLeft(0);

        if (viewportRef.current) {
            viewportRef.current.scrollTop = 0;
        }

        if (leftScrollbarRef.current) {
            leftScrollbarRef.current.scrollLeft = 0;
        }

        if (rightScrollbarRef.current) {
            rightScrollbarRef.current.scrollLeft = 0;
        }
    }, [activeSheet, activeSheetViewKey, runtime, viewState.kind]);

    React.useEffect(() => {
        if (!activeSheet || !selectedCell) {
            return;
        }

        if (
            selectedCell.rowNumber < 1 ||
            selectedCell.rowNumber > activeSheet.rowCount ||
            selectedCell.columnNumber < 1 ||
            selectedCell.columnNumber > activeSheet.columnCount
        ) {
            setSelectedCell(getInitialSelection(activeSheet, runtime));
            return;
        }

        if (getRowIndex(activeSheet, filter, runtime, selectedCell.rowNumber) < 0) {
            setSelectedCell(getDefaultSelectionForFilter(activeSheet, runtime, filter));
        }
    }, [activeSheet, filter, runtime, selectedCell]);

    React.useEffect(() => {
        if (!activeSheet) {
            return;
        }

        const maxScrollTop = Math.max(0, totalRowCount * ROW_HEIGHT - viewportHeight);
        if (scrollTop > maxScrollTop) {
            setScrollTop(maxScrollTop);
        }
    }, [activeSheet, scrollTop, totalRowCount, viewportHeight]);

    React.useEffect(() => {
        if (!hasHorizontalOverflow) {
            if (leftScrollbarRef.current) {
                leftScrollbarRef.current.scrollLeft = 0;
            }
            if (rightScrollbarRef.current) {
                rightScrollbarRef.current.scrollLeft = 0;
            }
            if (horizontalScrollLeft !== 0) {
                setHorizontalScrollLeft(0);
            }
        }
    }, [hasHorizontalOverflow, horizontalScrollLeft]);

    React.useLayoutEffect(() => {
        if (!activeSheet || !viewportRef.current) {
            return;
        }

        if (viewportRef.current.scrollTop !== scrollTop) {
            viewportRef.current.scrollTop = scrollTop;
        }
    }, [activeSheet, scrollTop]);

    React.useLayoutEffect(() => {
        const scrollbars = [leftScrollbarRef.current, rightScrollbarRef.current];
        for (const scrollbar of scrollbars) {
            if (scrollbar && scrollbar.scrollLeft !== horizontalScrollLeft) {
                scrollbar.scrollLeft = horizontalScrollLeft;
            }
        }
    }, [
        activeSheetViewKey,
        horizontalScrollLeft,
        leftHasHorizontalOverflow,
        rightHasHorizontalOverflow,
    ]);

    React.useEffect(() => {
        if (!activeSheet || !runtime) {
            if (activeDiffIndex !== 0) {
                setActiveDiffIndex(0);
            }
            return;
        }

        const row = selectedCell ? (runtime.rowByNumber.get(selectedCell.rowNumber) ?? null) : null;
        const nextDiffIndex = getSelectedDiffIndex(selectedCell, row, runtime);
        if (nextDiffIndex !== null) {
            if (nextDiffIndex !== activeDiffIndex) {
                setActiveDiffIndex(nextDiffIndex);
            }
            return;
        }

        if (activeSheet.diffCells.length === 0) {
            if (activeDiffIndex !== 0) {
                setActiveDiffIndex(0);
            }
            return;
        }

        if (activeDiffIndex >= activeSheet.diffCells.length) {
            setActiveDiffIndex(0);
        }
    }, [activeDiffIndex, activeSheet, runtime, selectedCell]);

    const scrollToRow = React.useEffectEvent((rowNumber: number) => {
        if (!activeSheet || !viewportRef.current) {
            return;
        }

        const rowIndex = getRowIndex(activeSheet, filter, runtime, rowNumber);
        if (rowIndex < 0) {
            return;
        }

        viewportRef.current.scrollTo({
            top: rowIndex * ROW_HEIGHT,
            behavior: "auto",
        });
        setScrollTop(rowIndex * ROW_HEIGHT);
    });

    const commitCurrentEdit = React.useEffectEvent((mode: "commit" | "cancel") => {
        const session = editingCell;
        if (!session) {
            return;
        }

        setEditingCell(null);

        if (mode === "commit") {
            setPendingEdits((currentPendingEdits) =>
                applyPendingEdit(currentPendingEdits, session)
            );
        }
    });

    const updateSelectedCell = React.useEffectEvent((nextSelection: CellSelection | null) => {
        if (editingCell) {
            commitCurrentEdit("commit");
        }

        setSelectedCell(nextSelection);

        if (!nextSelection || !runtime) {
            return;
        }

        const row = runtime.rowByNumber.get(nextSelection.rowNumber) ?? null;
        const nextDiffIndex = getSelectedDiffIndex(nextSelection, row, runtime);
        if (nextDiffIndex !== null) {
            setActiveDiffIndex(nextDiffIndex);
        }
    });

    const handleEditingChange = React.useEffectEvent((value: string) => {
        setEditingCell((currentEditingCell) =>
            currentEditingCell
                ? {
                      ...currentEditingCell,
                      value,
                  }
                : currentEditingCell
        );
    });

    const handleStartEdit = React.useEffectEvent(
        (selection: CellSelection, initialValue: string, modelValue: string) => {
            if (!activeSheet) {
                return;
            }

            if (editingCell) {
                commitCurrentEdit("commit");
            }

            setSelectedCell(selection);
            setEditingCell({
                ...selection,
                sheetKey: activeSheet.key,
                value: initialValue,
                modelValue,
            });
        }
    );

    const handleSave = React.useEffectEvent(() => {
        const nextPendingEdits = editingCell
            ? applyPendingEdit(pendingEdits, editingCell)
            : new Map(pendingEdits);

        setEditingCell(null);

        if (nextPendingEdits.size === 0) {
            setPendingEdits(nextPendingEdits);
            return;
        }

        setPendingEdits(new Map());
        vscode.postMessage({
            type: "saveEdits",
            edits: Array.from(nextPendingEdits.values()),
        });
    });

    const handleClearSelectedCell = React.useEffectEvent(() => {
        if (!model || !activeSheet || !selectedCell || editingCell) {
            return;
        }

        const row = runtime?.rowByNumber.get(selectedCell.rowNumber) ?? null;
        const cell =
            row?.cells.find((item) => item.columnNumber === selectedCell.columnNumber) ?? null;
        const status = cell?.status ?? "equal";
        if (!canEditCell(model, activeSheet, selectedCell.side, status)) {
            return;
        }

        const modelDisplay = getCellDisplay(cell, selectedCell.side);
        const key = getPendingEditKey(
            activeSheet.key,
            selectedCell.side,
            selectedCell.rowNumber,
            selectedCell.columnNumber
        );

        setPendingEdits((currentPendingEdits) => {
            const nextPendingEdits = new Map(currentPendingEdits);

            if (modelDisplay.value === "") {
                nextPendingEdits.delete(key);
            } else {
                nextPendingEdits.set(key, {
                    sheetKey: activeSheet.key,
                    side: selectedCell.side,
                    rowNumber: selectedCell.rowNumber,
                    columnNumber: selectedCell.columnNumber,
                    value: "",
                });
            }

            return nextPendingEdits;
        });
    });

    const handleFilterChange = React.useEffectEvent((nextFilter: FilterMode) => {
        if (editingCell) {
            commitCurrentEdit("commit");
        }

        React.startTransition(() => {
            setFilter(nextFilter);
        });
        setScrollTop(0);

        if (viewportRef.current) {
            viewportRef.current.scrollTop = 0;
        }

        if (!activeSheet) {
            return;
        }

        const defaultSelection = getDefaultSelectionForFilter(activeSheet, runtime, nextFilter);
        updateSelectedCell(defaultSelection);

        if (nextFilter === "diffs") {
            setActiveDiffIndex(0);
        }
    });

    const handleMoveDiff = React.useEffectEvent((direction: -1 | 1) => {
        if (editingCell) {
            commitCurrentEdit("commit");
        }

        if (!activeSheet || activeSheet.diffCells.length === 0) {
            return;
        }

        const totalDiffs = activeSheet.diffCells.length;
        const nextIndex = (activeDiffIndex + direction + totalDiffs) % totalDiffs;
        const nextDiff = activeSheet.diffCells[nextIndex];
        const row = runtime?.rowByNumber.get(nextDiff.rowNumber) ?? null;
        const cell = row?.cells.find((item) => item.columnNumber === nextDiff.columnNumber) ?? null;

        setActiveDiffIndex(nextIndex);
        updateSelectedCell({
            rowNumber: nextDiff.rowNumber,
            columnNumber: nextDiff.columnNumber,
            side: getPreferredSide(cell),
        });
        scrollToRow(nextDiff.rowNumber);
    });

    const syncHorizontalScroll = React.useEffectEvent((nextScrollLeft: number) => {
        const referenceElement = leftScrollbarRef.current ?? rightScrollbarRef.current;
        const maxScrollLeft = referenceElement
            ? Math.max(0, referenceElement.scrollWidth - referenceElement.clientWidth)
            : 0;
        const clampedScrollLeft = Math.max(0, Math.min(nextScrollLeft, maxScrollLeft));

        isSyncingHorizontalScrollRef.current = true;

        if (leftScrollbarRef.current && leftScrollbarRef.current.scrollLeft !== clampedScrollLeft) {
            leftScrollbarRef.current.scrollLeft = clampedScrollLeft;
        }
        if (
            rightScrollbarRef.current &&
            rightScrollbarRef.current.scrollLeft !== clampedScrollLeft
        ) {
            rightScrollbarRef.current.scrollLeft = clampedScrollLeft;
        }

        setHorizontalScrollLeft(clampedScrollLeft);

        if (horizontalScrollSyncFrameRef.current) {
            cancelAnimationFrame(horizontalScrollSyncFrameRef.current);
        }
        horizontalScrollSyncFrameRef.current = requestAnimationFrame(() => {
            horizontalScrollSyncFrameRef.current = 0;
            isSyncingHorizontalScrollRef.current = false;
        });
    });

    const handleHorizontalScrollbarScroll = React.useEffectEvent(
        (event: React.UIEvent<HTMLDivElement>) => {
            if (isSyncingHorizontalScrollRef.current) {
                return;
            }

            syncHorizontalScroll(event.currentTarget.scrollLeft);
        }
    );

    const adjustHorizontalScroll = React.useEffectEvent((delta: number) => {
        const element = leftScrollbarRef.current ?? rightScrollbarRef.current;
        if (!element) {
            return;
        }

        syncHorizontalScroll(element.scrollLeft + delta);
    });

    const handleSetSheet = React.useEffectEvent((sheetKey: string) => {
        if (editingCell) {
            commitCurrentEdit("commit");
        }

        vscode.postMessage({ type: "setSheet", sheetKey });
    });

    React.useEffect(() => {
        const handleKeyDown = (event: KeyboardEvent) => {
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
                event.preventDefault();
                handleSave();
                return;
            }

            if (isTextInputTarget(event.target)) {
                return;
            }

            if (isClearSelectedCellKey(event)) {
                event.preventDefault();
                handleClearSelectedCell();
            }
        };

        document.addEventListener("keydown", handleKeyDown);
        return () => {
            document.removeEventListener("keydown", handleKeyDown);
        };
    }, [handleClearSelectedCell, handleSave]);

    if (viewState.kind === "loading") {
        return <div className="diff-loading">{viewState.message}</div>;
    }

    if (viewState.kind === "error") {
        return (
            <div className="diff-errorState">
                <div className="diff-errorState__message">{viewState.message}</div>
                <button
                    type="button"
                    className="diff-button"
                    onClick={() => vscode.postMessage({ type: "reload" })}
                >
                    {STRINGS.reload}
                </button>
            </div>
        );
    }

    if (!model) {
        return <div className="diff-loading">{STRINGS.loading}</div>;
    }

    const sameRowCount = activeSheet
        ? Math.max(0, activeSheet.rowCount - activeSheet.diffRowCount)
        : 0;
    const effectivePendingEdits = editingCell
        ? applyPendingEdit(pendingEdits, editingCell)
        : pendingEdits;
    const pendingSummary = getPendingSummary(activeSheet?.key ?? null, effectivePendingEdits);
    const hasUnsavedChanges = effectivePendingEdits.size > 0;
    const selectedColumnNumber = selectedCell?.columnNumber ?? null;
    const selectedAddress =
        activeSheet && selectedCell
            ? `${activeSheet.columns[selectedCell.columnNumber - 1] ?? selectedCell.columnNumber}${selectedCell.rowNumber}`
            : STRINGS.none;
    const currentDiffLabel =
        activeSheet && activeSheet.diffCells.length > 0
            ? `${activeDiffIndex + 1}/${activeSheet.diffCells.length}`
            : STRINGS.none;

    return (
        <div className="diff-shell">
            <header className="diff-toolbarBar">
                <div className="diff-toolbarGroup">
                    <div className="diff-chipGroup">
                        <button
                            type="button"
                            className={classNames([
                                "diff-chip",
                                filter === "all" && "diff-chip--active",
                            ])}
                            onClick={() => handleFilterChange("all")}
                        >
                            <span
                                className="codicon codicon-list-flat diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.all}</span>
                        </button>
                        <button
                            type="button"
                            className={classNames([
                                "diff-chip",
                                filter === "diffs" && "diff-chip--active",
                            ])}
                            onClick={() => handleFilterChange("diffs")}
                        >
                            <span
                                className="codicon codicon-diff-multiple diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.diffs}</span>
                        </button>
                        <button
                            type="button"
                            className={classNames([
                                "diff-chip",
                                filter === "same" && "diff-chip--active",
                            ])}
                            onClick={() => handleFilterChange("same")}
                        >
                            <span
                                className="codicon codicon-check-all diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.same}</span>
                        </button>
                    </div>
                </div>
                <div className="diff-toolbarGroup diff-toolbarGroup--actions">
                    <div className="diff-actionGroup">
                        <button
                            type="button"
                            className="diff-button"
                            disabled={!activeSheet || activeSheet.diffCells.length === 0}
                            onClick={() => handleMoveDiff(-1)}
                        >
                            <span
                                className="codicon codicon-arrow-up diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.prevDiff}</span>
                        </button>
                        <button
                            type="button"
                            className="diff-button"
                            disabled={!activeSheet || activeSheet.diffCells.length === 0}
                            onClick={() => handleMoveDiff(1)}
                        >
                            <span
                                className="codicon codicon-arrow-down diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.nextDiff}</span>
                        </button>
                        <button
                            type="button"
                            className="diff-button"
                            onClick={() => vscode.postMessage({ type: "swap" })}
                        >
                            <span
                                className="codicon codicon-arrow-swap diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.swap}</span>
                        </button>
                        <button
                            type="button"
                            className="diff-button"
                            onClick={() => vscode.postMessage({ type: "reload" })}
                        >
                            <span
                                className="codicon codicon-refresh diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.reload}</span>
                        </button>
                        <button
                            type="button"
                            className={classNames([
                                "diff-button",
                                hasUnsavedChanges && "diff-button--active",
                            ])}
                            disabled={!hasUnsavedChanges}
                            onClick={handleSave}
                        >
                            <span
                                className="codicon codicon-save diff-controlIcon"
                                aria-hidden="true"
                            />
                            <span>{STRINGS.save}</span>
                        </button>
                    </div>
                </div>
            </header>

            <section className="diff-gridSection">
                {!activeSheet ? (
                    <EmptyState message={STRINGS.noSheet} />
                ) : (
                    <>
                        <div className="diff-paneMetaRow">
                            <div className="diff-pane">
                                <PaneMeta file={model.leftFile} />
                            </div>
                            <div className="diff-divider" />
                            <div className="diff-pane">
                                <PaneMeta file={model.rightFile} />
                            </div>
                        </div>

                        {totalRowCount === 0 || activeSheet.columnCount === 0 ? (
                            <EmptyState message={STRINGS.noRows} />
                        ) : (
                            <div className="diff-gridShell">
                                <div className="diff-gridHeaderRow">
                                    <div className="diff-pane">
                                        <PaneHeader
                                            columns={activeSheet.columns}
                                            columnDiffTones={runtime?.columnDiffTones ?? []}
                                            pendingColumns={pendingSummary.columnsBySide.left}
                                            scrollLeft={horizontalScrollLeft}
                                            selectedColumnNumber={selectedColumnNumber}
                                            viewportRef={leftHeaderViewportRef}
                                        />
                                    </div>
                                    <div className="diff-divider" />
                                    <div className="diff-pane">
                                        <PaneHeader
                                            columns={activeSheet.columns}
                                            columnDiffTones={runtime?.columnDiffTones ?? []}
                                            pendingColumns={pendingSummary.columnsBySide.right}
                                            scrollLeft={horizontalScrollLeft}
                                            selectedColumnNumber={selectedColumnNumber}
                                            viewportRef={rightHeaderViewportRef}
                                        />
                                    </div>
                                </div>

                                <div
                                    ref={viewportRef}
                                    className="diff-gridViewport"
                                    onScroll={(event) => {
                                        const nextScrollTop = event.currentTarget.scrollTop;
                                        if (scrollFrameRef.current) {
                                            cancelAnimationFrame(scrollFrameRef.current);
                                        }

                                        scrollFrameRef.current = requestAnimationFrame(() => {
                                            scrollFrameRef.current = 0;
                                            setScrollTop(nextScrollTop);
                                        });
                                    }}
                                    onWheel={(event) => {
                                        if (Math.abs(event.deltaX) <= Math.abs(event.deltaY)) {
                                            return;
                                        }

                                        event.preventDefault();
                                        adjustHorizontalScroll(event.deltaX);
                                    }}
                                >
                                    <div
                                        className="diff-gridViewportInner"
                                        style={{ height: viewportContentHeight }}
                                    >
                                        <div className="diff-gridBackdrop" aria-hidden="true">
                                            <div />
                                            <div className="diff-gridBackdrop__divider" />
                                            <div />
                                        </div>
                                        <div
                                            className="diff-visibleRows"
                                            style={{ transform: `translateY(${offsetY}px)` }}
                                        >
                                            {visibleRows.map(({ rowNumber, row }) => (
                                                <div
                                                    key={rowNumber}
                                                    className="diff-pairRow"
                                                    style={{ height: ROW_HEIGHT }}
                                                >
                                                    <div className="diff-pane">
                                                        <SideRow
                                                            model={model}
                                                            activeSheet={activeSheet}
                                                            side="left"
                                                            rowNumber={rowNumber}
                                                            row={row}
                                                            columns={activeSheet.columns}
                                                            scrollLeft={horizontalScrollLeft}
                                                            selectedCell={selectedCell}
                                                            pendingEdits={effectivePendingEdits}
                                                            editingCell={editingCell}
                                                            isPendingRow={pendingSummary.rowsBySide.left.has(
                                                                rowNumber
                                                            )}
                                                            onSelect={updateSelectedCell}
                                                            onStartEdit={handleStartEdit}
                                                            onEditingChange={handleEditingChange}
                                                            onCommitEdit={() =>
                                                                commitCurrentEdit("commit")
                                                            }
                                                            onCancelEdit={() =>
                                                                commitCurrentEdit("cancel")
                                                            }
                                                        />
                                                    </div>
                                                    <div className="diff-divider" />
                                                    <div className="diff-pane">
                                                        <SideRow
                                                            model={model}
                                                            activeSheet={activeSheet}
                                                            side="right"
                                                            rowNumber={rowNumber}
                                                            row={row}
                                                            columns={activeSheet.columns}
                                                            scrollLeft={horizontalScrollLeft}
                                                            selectedCell={selectedCell}
                                                            pendingEdits={effectivePendingEdits}
                                                            editingCell={editingCell}
                                                            isPendingRow={pendingSummary.rowsBySide.right.has(
                                                                rowNumber
                                                            )}
                                                            onSelect={updateSelectedCell}
                                                            onStartEdit={handleStartEdit}
                                                            onEditingChange={handleEditingChange}
                                                            onCommitEdit={() =>
                                                                commitCurrentEdit("commit")
                                                            }
                                                            onCancelEdit={() =>
                                                                commitCurrentEdit("cancel")
                                                            }
                                                        />
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                </div>
                                {hasHorizontalOverflow ? (
                                    <div className="diff-gridScrollbarRow">
                                        <div className="diff-pane">
                                            <PaneScrollbar
                                                scrollbarRef={leftScrollbarRef}
                                                enabled={leftHasHorizontalOverflow}
                                                columnTrackWidth={columnTrackWidth}
                                                onScroll={handleHorizontalScrollbarScroll}
                                            />
                                        </div>
                                        <div className="diff-divider" />
                                        <div className="diff-pane">
                                            <PaneScrollbar
                                                scrollbarRef={rightScrollbarRef}
                                                enabled={rightHasHorizontalOverflow}
                                                columnTrackWidth={columnTrackWidth}
                                                onScroll={handleHorizontalScrollbarScroll}
                                            />
                                        </div>
                                    </div>
                                ) : null}
                            </div>
                        )}
                    </>
                )}
            </section>

            <nav className="diff-sheetTabs diff-sheetTabs--bottom" aria-label={STRINGS.sheets}>
                {model.sheets.map((sheet) => {
                    const hasPending = pendingSummary.sheetKeys.has(sheet.key);
                    const markerTone = getEffectiveMarkerTone(
                        sheet.hasDiff ? sheet.diffTone : null,
                        hasPending
                    );

                    return (
                        <button
                            key={sheet.key}
                            type="button"
                            className={classNames([
                                "diff-sheetTab",
                                sheet.isActive && "diff-sheetTab--active",
                                hasPending && "diff-sheetTab--pending",
                            ])}
                            title={getSheetTooltip(sheet)}
                            onClick={() => handleSetSheet(sheet.key)}
                        >
                            <DiffMarker tone={markerTone} className="diff-sheetTab__marker" />
                            <span className="diff-sheetTab__label">{sheet.label}</span>
                        </button>
                    );
                })}
            </nav>

            <footer className="diff-statusBar">
                <span>
                    {STRINGS.sheets} {activeSheet?.label ?? STRINGS.none}
                </span>
                <span>
                    {STRINGS.rows} {activeSheet?.rowCount ?? 0}
                </span>
                <span>
                    {STRINGS.filter} {getFilterLabel(filter)}
                </span>
                <span>
                    {STRINGS.diffRows} {activeSheet?.diffRowCount ?? 0}
                </span>
                <span>
                    {STRINGS.sameRows} {sameRowCount}
                </span>
                <span>
                    {STRINGS.visibleRows} {totalRowCount}
                </span>
                <span>
                    {STRINGS.currentDiff} {currentDiffLabel}
                </span>
                <span>
                    {STRINGS.selected} {selectedAddress}
                </span>
            </footer>
        </div>
    );
}

const root = createRoot(document.getElementById("app")!);
root.render(<App />);
