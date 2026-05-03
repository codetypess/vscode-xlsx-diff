import { For, Show, createEffect, createMemo, createSignal, onCleanup, onMount } from "solid-js";
import { render } from "solid-js/web";
import type { CellDiffStatus } from "../../core/model/types";
import { RUNTIME_MESSAGES, type DiffPanelStrings } from "../../i18n/catalog";
import {
    DEFAULT_COLUMN_PIXEL_WIDTH,
    DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
    createPixelColumnLayout,
    getFontShorthand,
    getPixelColumnLeft,
    getPixelColumnRight,
    getPixelColumnWidth,
    getPixelColumnWindow,
    measureMaximumDigitWidth,
} from "../../webview/column-layout";
import { getDiffRowHeaderWidth } from "../../webview/diff-panel/diff-grid-layout";
import { getSelectionPreviewInlineDiff } from "../../webview/diff-panel/selection-preview-diff";
import type {
    DiffPanelColumnView,
    DiffPanelRowView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "../../webview/diff-panel/diff-panel-types";
import {
    getMaxVisibleSheetTabsForWidth,
    partitionSheetTabs,
} from "../../webview/editor-sheet-tabs";
import {
    createWebviewReadyMessage,
    isDiffSessionIncomingMessage,
    type DiffSessionIncomingMessage,
    type DiffWebviewOutgoingMessage,
} from "../shared/session-protocol";
import { createInitialDiffSessionState, reduceDiffSessionMessage } from "./session";
import {
    applyPendingDiffEdit,
    beginDiffCellEdit,
    clampDiffHorizontalScroll,
    createSelectedDiffCell,
    createSaveDiffEditsMessage,
    finalizeDiffCellEdit,
    filterDiffRows,
    getPendingDiffEditKey,
    getRenderedDiffCellValue,
    getSelectedDiffCellState,
    getWrappedDiffIndex,
    type DraftEditState,
    type PendingDiffEdit,
    type RowFilterMode,
    type SelectedDiffCell,
} from "./grid-helpers";

interface VsCodeApi {
    postMessage(message: DiffWebviewOutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

const vscode = acquireVsCodeApi();

const DEFAULT_STRINGS = RUNTIME_MESSAGES.en.diffPanel;
const DIFF_SHEET_TAB_ITEM_GAP = 1;
const DIFF_SHEET_TAB_ESTIMATED_WIDTH = 120;
const DIFF_SHEET_TAB_VISIBLE_MAX_WIDTH = 160;
const DIFF_SHEET_TAB_OVERFLOW_TRIGGER_WIDTH = 32;
const ROW_HEIGHT = 27;
const ROW_OVERSCAN = 8;

type DiffMarkerTone = CellDiffStatus | "pending" | null;

interface PendingSummary {
    sheetKeys: Set<string>;
    rowsBySide: Record<"left" | "right", Set<number>>;
    columnsBySide: Record<"left" | "right", Set<number>>;
}

interface SelectionPreviewValue {
    value: string;
}

function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}

function getDiffStrings(): Pick<
    DiffPanelStrings,
    | "all"
    | "close"
    | "currentDiff"
    | "definedNames"
    | "diffCells"
    | "diffRows"
    | "diffs"
    | "emptyValue"
    | "loading"
    | "filter"
    | "freezePanes"
    | "mergedRanges"
    | "modified"
    | "moreSheets"
    | "nextDiff"
    | "noRows"
    | "noSheet"
    | "none"
    | "prevDiff"
    | "readOnly"
    | "reload"
    | "rows"
    | "sameRows"
    | "sheetAdded"
    | "sheetOrder"
    | "sheetRemoved"
    | "sheetRenamed"
    | "sheetVisibility"
    | "sheets"
    | "same"
    | "save"
    | "selected"
    | "size"
    | "structure"
    | "swap"
    | "visibleRows"
    | "viewDetails"
    | "workbookStructure"
> {
    const strings = (globalThis as Record<string, unknown>).__XLSX_DIFF_STRINGS__ as
        | Partial<DiffPanelStrings>
        | undefined;
    return {
        all: strings?.all ?? "All",
        close: strings?.close ?? DEFAULT_STRINGS.close,
        currentDiff: strings?.currentDiff ?? "Current Diff",
        definedNames: strings?.definedNames ?? "Defined Names",
        diffCells: strings?.diffCells ?? "Diff Cells",
        diffRows: strings?.diffRows ?? "Diff Rows",
        diffs: strings?.diffs ?? "Diffs",
        emptyValue: strings?.emptyValue ?? "(empty)",
        filter: strings?.filter ?? DEFAULT_STRINGS.filter,
        freezePanes: strings?.freezePanes ?? DEFAULT_STRINGS.freezePanes,
        loading: strings?.loading ?? "Loading diff...",
        mergedRanges: strings?.mergedRanges ?? DEFAULT_STRINGS.mergedRanges,
        modified: strings?.modified ?? DEFAULT_STRINGS.modified,
        moreSheets: strings?.moreSheets ?? DEFAULT_STRINGS.moreSheets,
        nextDiff: strings?.nextDiff ?? "Next Diff",
        noRows: strings?.noRows ?? DEFAULT_STRINGS.noRows,
        noSheet: strings?.noSheet ?? "No sheet selected.",
        none: strings?.none ?? "None",
        prevDiff: strings?.prevDiff ?? "Prev Diff",
        readOnly: strings?.readOnly ?? "Read-only",
        reload: strings?.reload ?? "Reload",
        rows: strings?.rows ?? DEFAULT_STRINGS.rows,
        sameRows: strings?.sameRows ?? DEFAULT_STRINGS.sameRows,
        sheetAdded: strings?.sheetAdded ?? DEFAULT_STRINGS.sheetAdded,
        sheetOrder: strings?.sheetOrder ?? DEFAULT_STRINGS.sheetOrder,
        sheetRemoved: strings?.sheetRemoved ?? DEFAULT_STRINGS.sheetRemoved,
        sheetRenamed: strings?.sheetRenamed ?? DEFAULT_STRINGS.sheetRenamed,
        sheetVisibility: strings?.sheetVisibility ?? DEFAULT_STRINGS.sheetVisibility,
        sheets: strings?.sheets ?? DEFAULT_STRINGS.sheets,
        same: strings?.same ?? "Same",
        save: strings?.save ?? "Save",
        selected: strings?.selected ?? "Selected",
        size: strings?.size ?? "Size",
        structure: strings?.structure ?? DEFAULT_STRINGS.structure,
        swap: strings?.swap ?? "Swap",
        visibleRows: strings?.visibleRows ?? DEFAULT_STRINGS.visibleRows,
        viewDetails: strings?.viewDetails ?? "View Details",
        workbookStructure: strings?.workbookStructure ?? DEFAULT_STRINGS.workbookStructure,
    };
}

function getColumnHeaderLabel(column: DiffPanelColumnView): string {
    if (column.leftLabel && column.rightLabel && column.leftLabel !== column.rightLabel) {
        return `${column.leftLabel} / ${column.rightLabel}`;
    }

    return column.leftLabel || column.rightLabel || String(column.columnNumber);
}

function getCellToneClass(cell: DiffPanelSparseCellView): string {
    switch (cell.status) {
        case "modified":
            return "diff-cell--modified";
        case "added":
            return "diff-cell--added";
        case "removed":
            return "diff-cell--removed";
        default:
            return cell.leftPresent || cell.rightPresent ? "" : "diff-cell--ghost";
    }
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

function getRowToneClass(row: DiffPanelRowView): string {
    switch (row.diffTone) {
        case "modified":
            return "diff-indexCell--modified";
        case "added":
            return "diff-indexCell--added";
        case "removed":
            return "diff-indexCell--removed";
        default:
            return "";
    }
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

function findCellInSheet(
    sheet: DiffPanelSheetView | null,
    rowNumber: number,
    columnNumber: number
): DiffPanelSparseCellView | null {
    if (!sheet) {
        return null;
    }

    const row = sheet.rows.find((candidate) => candidate.rowNumber === rowNumber);
    return row?.cells.find((candidate) => candidate.columnNumber === columnNumber) ?? null;
}

function getDefaultSideForRow(row: DiffPanelRowView | null): "left" | "right" {
    return row?.rightRowNumber !== null ? "right" : "left";
}

function getPreferredSide(cell: DiffPanelSparseCellView | null): "left" | "right" {
    return cell?.status === "removed" ? "left" : "right";
}

function getSelectionAddress(
    sheet: DiffPanelSheetView | null,
    selection: SelectedDiffCell | null,
    side: "left" | "right"
): string {
    if (!sheet || !selection) {
        return DEFAULT_STRINGS.none;
    }

    const column = sheet.columns[selection.columnNumber - 1];
    const columnLabel = side === "left" ? (column?.leftLabel ?? "") : (column?.rightLabel ?? "");

    if (!columnLabel || selection.sourceRowNumber === null) {
        return DEFAULT_STRINGS.none;
    }

    return `${columnLabel}${selection.sourceRowNumber}`;
}

function getDiffCellTitle(
    columnLabel: string,
    rowNumber: number | null,
    value: string,
    formula: string | null
): string {
    const lines = rowNumber === null || !columnLabel ? [] : [`${columnLabel}${rowNumber}`];

    if (value) {
        lines.push(value);
    }

    if (formula) {
        lines.push(`fx ${formula}`);
    }

    return lines.join("\n");
}

function getPendingSummary(
    activeSheetKey: string | null,
    pendingEdits: Record<string, PendingDiffEdit>
): PendingSummary {
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

    for (const pendingEdit of Object.values(pendingEdits)) {
        summary.sheetKeys.add(pendingEdit.sheetKey);

        if (!activeSheetKey || pendingEdit.sheetKey !== activeSheetKey) {
            continue;
        }

        summary.rowsBySide[pendingEdit.side].add(pendingEdit.rowNumber);
        summary.columnsBySide[pendingEdit.side].add(pendingEdit.columnNumber);
    }

    return summary;
}

function getSelectionPreviewValue(
    sheet: DiffPanelSheetView | null,
    selection: SelectedDiffCell | null,
    editingCell: DraftEditState | null,
    pendingEdits: Record<string, PendingDiffEdit>,
    side: "left" | "right"
): SelectionPreviewValue {
    if (!sheet || !selection) {
        return {
            value: "",
        };
    }

    const row = sheet.rows.find((candidate) => candidate.rowNumber === selection.rowNumber) ?? null;
    const cell =
        row?.cells.find((candidate) => candidate.columnNumber === selection.columnNumber) ?? null;
    if (
        editingCell?.sheetKey === sheet.key &&
        editingCell.side === side &&
        editingCell.rowNumber === selection.rowNumber &&
        editingCell.columnNumber === selection.columnNumber
    ) {
        return {
            value: editingCell.value,
        };
    }

    return {
        value: getRenderedDiffCellValue(
            pendingEdits,
            sheet.key,
            selection.rowNumber,
            selection.columnNumber,
            cell,
            side
        ),
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

function getStructuralChangeLabels(
    strings: ReturnType<typeof getDiffStrings>,
    sheet: {
        kind: DiffPanelSheetView["kind"];
        mergedRangesChanged: boolean;
        freezePaneChanged: boolean;
        visibilityChanged: boolean;
        sheetOrderChanged: boolean;
    }
): string[] {
    const structuralChanges: string[] = [];

    if (sheet.kind === "added") {
        structuralChanges.push(strings.sheetAdded);
    } else if (sheet.kind === "removed") {
        structuralChanges.push(strings.sheetRemoved);
    } else if (sheet.kind === "renamed") {
        structuralChanges.push(strings.sheetRenamed);
    }

    if (sheet.mergedRangesChanged) {
        structuralChanges.push(strings.mergedRanges);
    }

    if (sheet.freezePaneChanged) {
        structuralChanges.push(strings.freezePanes);
    }

    if (sheet.visibilityChanged) {
        structuralChanges.push(strings.sheetVisibility);
    }

    if (sheet.sheetOrderChanged) {
        structuralChanges.push(strings.sheetOrder);
    }

    return structuralChanges;
}

function getWorkbookStructuralChangeLabels(
    strings: ReturnType<typeof getDiffStrings>,
    definedNamesChanged: boolean
): string[] {
    return definedNamesChanged ? [strings.definedNames] : [];
}

function getSheetTooltip(
    strings: ReturnType<typeof getDiffStrings>,
    sheet: {
        key: string;
        label: string;
        diffCellCount: number;
        diffRowCount: number;
        mergedRangesChanged: boolean;
        freezePaneChanged: boolean;
        visibilityChanged: boolean;
        sheetOrderChanged: boolean;
        kind: DiffPanelSheetView["kind"];
    }
): string {
    const structuralChanges = getStructuralChangeLabels(strings, sheet);
    const tooltip = `${sheet.label} · ${sheet.diffCellCount} ${strings.diffCells} · ${sheet.diffRowCount} ${strings.diffRows}`;
    return structuralChanges.length > 0 ? `${tooltip} · ${structuralChanges.join(", ")}` : tooltip;
}

function getMaxVisibleDiffSheetTabs(
    sheets: ReadonlyArray<{ key: string; isActive: boolean }>,
    containerWidth: number,
    measuredTabWidths: Record<string, number>
): number {
    return getMaxVisibleSheetTabsForWidth(sheets, {
        containerWidth,
        getTabWidth: (sheet) => measuredTabWidths[sheet.key] ?? DIFF_SHEET_TAB_ESTIMATED_WIDTH,
        itemGap: DIFF_SHEET_TAB_ITEM_GAP,
        overflowTriggerWidth: DIFF_SHEET_TAB_OVERFLOW_TRIGGER_WIDTH,
    });
}

function DiffMarker(props: { tone: DiffMarkerTone; class?: string }) {
    if (!props.tone || props.tone === "equal") {
        return null;
    }

    return (
        <span
            class={classNames(["diff-diffMarker", `diff-diffMarker--${props.tone}`, props.class])}
            aria-hidden="true"
        />
    );
}

function SelectionValueContent(props: {
    value: string;
    otherValue: string;
    class: string;
    emptyValueLabel: string;
}) {
    const inlineDiff = createMemo(() =>
        getSelectionPreviewInlineDiff(props.value, props.otherValue)
    );

    return (
        <span class={props.class}>
            <Show
                when={props.value.length > 0}
                fallback={
                    <span
                        class={classNames([
                            "diff-selectionValue__empty",
                            props.otherValue.length > 0 && "diff-selectionValue__diff",
                        ])}
                    >
                        {props.emptyValueLabel}
                    </span>
                }
            >
                <>
                    {inlineDiff().before}
                    <Show when={inlineDiff().changed.length > 0}>
                        <span class="diff-selectionValue__diff">{inlineDiff().changed}</span>
                    </Show>
                    {inlineDiff().after}
                </>
            </Show>
        </span>
    );
}

function PaneMeta(props: {
    file: {
        title: string;
        path: string;
        sizeLabel: string;
        detailFacts: Array<{
            label: string;
            value: string;
            title?: string;
        }>;
        modifiedLabel: string;
        isReadonly: boolean;
    } | null;
    strings: ReturnType<typeof getDiffStrings>;
}) {
    return (
        <div class="diff-paneMeta">
            <div class="diff-paneMeta__name" title={props.file?.path ?? props.strings.none}>
                <span class="diff-paneMeta__nameText">
                    {props.file?.title ?? props.strings.none}
                </span>
                <Show when={props.file?.isReadonly}>
                    <span
                        class="codicon codicon-lock diff-paneMeta__lock"
                        title={props.strings.readOnly}
                        aria-label={props.strings.readOnly}
                    />
                </Show>
            </div>
            <div class="diff-paneMeta__meta">
                <div class="diff-paneMeta__path" title={props.file?.path ?? props.strings.none}>
                    {props.file?.path ?? props.strings.none}
                </div>
                <div class="diff-paneMeta__facts">
                    <span>{`${props.strings.size}: ${props.file?.sizeLabel ?? props.strings.none}`}</span>
                    <Show when={props.file?.modifiedLabel}>
                        {(label) => <span>{`${props.strings.modified}: ${label()}`}</span>}
                    </Show>
                    <For each={props.file?.detailFacts ?? []}>
                        {(fact) => (
                            <span
                                title={fact.title ?? fact.value}
                            >{`${fact.label}: ${fact.value}`}</span>
                        )}
                    </For>
                </div>
            </div>
        </div>
    );
}

function EmptyState(props: { message: string }) {
    return (
        <div class="diff-emptyState">
            <div class="diff-emptyState__message">{props.message}</div>
        </div>
    );
}

function CellEditor(props: {
    edit: DraftEditState;
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
            class="diff-cell__input"
            type="text"
            value={props.edit.value}
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

function canEditDiffCell(
    comparison: ReturnType<typeof createInitialDiffSessionState>["document"]["comparison"],
    sheet: DiffPanelSheetView,
    side: "left" | "right",
    sourceRowNumber: number | null,
    sourceColumnNumber: number | null,
    status: CellDiffStatus
): boolean {
    const file = side === "left" ? comparison.leftFile : comparison.rightFile;
    if (file?.isReadonly) {
        return false;
    }

    const sheetName = side === "left" ? sheet.leftName : sheet.rightName;
    if (!sheetName) {
        return false;
    }

    if (sourceRowNumber === null || sourceColumnNumber === null) {
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

function getDefaultSelectionForFilter(
    sheet: DiffPanelSheetView,
    filter: RowFilterMode
): {
    selection: SelectedDiffCell;
    side: "left" | "right";
} | null {
    if (sheet.columnCount <= 0 || sheet.rowCount <= 0) {
        return null;
    }

    if (filter === "diffs") {
        const firstDiffCell = sheet.diffCells[0];
        if (firstDiffCell) {
            const row =
                sheet.rows.find((candidate) => candidate.rowNumber === firstDiffCell.rowNumber) ??
                null;
            const cell =
                row?.cells.find(
                    (candidate) => candidate.columnNumber === firstDiffCell.columnNumber
                ) ?? null;
            const side = getPreferredSide(cell);

            return {
                selection: createSelectedDiffCell(
                    sheet,
                    row,
                    firstDiffCell.rowNumber,
                    firstDiffCell.columnNumber,
                    side
                ),
                side,
            };
        }
    }

    const rows = filterDiffRows(sheet.rows, filter);
    const firstRow =
        rows.find((candidate) => candidate.rowNumber === 1) ?? rows[0] ?? sheet.rows[0] ?? null;
    if (!firstRow) {
        return null;
    }

    const side = getDefaultSideForRow(firstRow);
    return {
        selection: createSelectedDiffCell(sheet, firstRow, firstRow.rowNumber, 1, side),
        side,
    };
}

function DiffBootstrapApp() {
    let gridViewportElement: HTMLDivElement | undefined;
    let leftHeaderViewportElement: HTMLDivElement | undefined;
    let rightHeaderViewportElement: HTMLDivElement | undefined;
    let leftScrollbarElement: HTMLDivElement | undefined;
    let rightScrollbarElement: HTMLDivElement | undefined;
    let sheetTabsViewportElement: HTMLDivElement | undefined;
    let sheetTabsOverflowElement: HTMLDivElement | undefined;
    let sheetTabsMeasureElement: HTMLDivElement | undefined;
    let scrollFrameId = 0;

    const [session, setSession] = createSignal(createInitialDiffSessionState());
    const [rowFilter, setRowFilter] = createSignal<RowFilterMode>("all");
    const [activeDiffIndex, setActiveDiffIndex] = createSignal(0);
    const [previewExpanded, setPreviewExpanded] = createSignal(false);
    const [selectedCell, setSelectedCell] = createSignal<SelectedDiffCell | null>(null);
    const [selectedSide, setSelectedSide] = createSignal<"left" | "right">("right");
    const [editingCell, setEditingCell] = createSignal<DraftEditState | null>(null);
    const [pendingEdits, setPendingEdits] = createSignal<Record<string, PendingDiffEdit>>({});
    const [scrollTop, setScrollTop] = createSignal(0);
    const [viewportHeight, setViewportHeight] = createSignal(480);
    const [horizontalScrollLeft, setHorizontalScrollLeft] = createSignal(0);
    const [horizontalViewportWidth, setHorizontalViewportWidth] = createSignal(0);
    const [viewportScrollbarWidth, setViewportScrollbarWidth] = createSignal(0);
    const [maximumDigitWidth, setMaximumDigitWidth] = createSignal(DEFAULT_MAXIMUM_DIGIT_WIDTH_PX);
    const [sheetTabsViewportWidth, setSheetTabsViewportWidth] = createSignal(0);
    const [measuredSheetTabWidths, setMeasuredSheetTabWidths] = createSignal<
        Record<string, number>
    >({});
    const [isSheetOverflowOpen, setIsSheetOverflowOpen] = createSignal(false);
    const strings = getDiffStrings();
    const comparison = () => session().document.comparison;
    const activeSheet = () => comparison().activeSheet;

    const measureHorizontalViewportWidth = () => {
        const nextWidth =
            leftHeaderViewportElement?.clientWidth ?? rightHeaderViewportElement?.clientWidth ?? 0;
        setHorizontalViewportWidth(nextWidth);
    };

    const measureViewportMetrics = () => {
        measureHorizontalViewportWidth();
        setViewportHeight(gridViewportElement?.clientHeight ?? 480);
        setViewportScrollbarWidth(
            gridViewportElement
                ? Math.max(0, gridViewportElement.offsetWidth - gridViewportElement.clientWidth)
                : 0
        );
    };

    const measureMaximumDigitWidthFromDocument = () => {
        if (globalThis.navigator?.userAgent?.includes("jsdom")) {
            setMaximumDigitWidth(DEFAULT_MAXIMUM_DIGIT_WIDTH_PX);
            return;
        }

        try {
            setMaximumDigitWidth(measureMaximumDigitWidth(getFontShorthand(document.body)));
        } catch {
            setMaximumDigitWidth(DEFAULT_MAXIMUM_DIGIT_WIDTH_PX);
        }
    };

    const measureSheetTabsViewport = () => {
        setSheetTabsViewportWidth(
            Math.ceil(sheetTabsViewportElement?.getBoundingClientRect().width ?? 0)
        );
    };

    const measureSheetTabWidths = () => {
        if (!sheetTabsMeasureElement) {
            return;
        }

        const nextMeasuredTabWidths: Record<string, number> = {};
        for (const element of sheetTabsMeasureElement.querySelectorAll<HTMLElement>(
            '[data-role="diff-sheet-tab-measure"]'
        )) {
            const sheetKey = element.dataset.sheetKey;
            if (!sheetKey) {
                continue;
            }

            nextMeasuredTabWidths[sheetKey] = Math.min(
                DIFF_SHEET_TAB_VISIBLE_MAX_WIDTH,
                Math.ceil(element.getBoundingClientRect().width)
            );
        }

        setMeasuredSheetTabWidths(nextMeasuredTabWidths);
    };

    onMount(() => {
        const handleMessage = (event: MessageEvent<DiffSessionIncomingMessage>) => {
            const payload = event.data;
            if (!isDiffSessionIncomingMessage(payload)) {
                return;
            }

            setSession((current) => reduceDiffSessionMessage(current, payload));
        };
        const handleResize = () => {
            measureViewportMetrics();
            measureMaximumDigitWidthFromDocument();
            measureSheetTabsViewport();
            measureSheetTabWidths();
        };
        const handleKeyDown = (event: KeyboardEvent) => {
            if (previewExpanded() && (event.key === "Escape" || event.code === "Escape")) {
                event.preventDefault();
                setPreviewExpanded(false);
                return;
            }

            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
                event.preventDefault();
                savePendingEdits();
                return;
            }

            if (isTextInputTarget(event.target)) {
                return;
            }

            if (isClearSelectedCellKey(event)) {
                event.preventDefault();
                clearSelectedCell();
            }
        };

        window.addEventListener("message", handleMessage);
        window.addEventListener("resize", handleResize);
        document.addEventListener("keydown", handleKeyDown);
        vscode.postMessage(createWebviewReadyMessage());
        measureViewportMetrics();
        measureMaximumDigitWidthFromDocument();
        measureSheetTabsViewport();
        measureSheetTabWidths();

        const resizeObserver =
            typeof ResizeObserver === "undefined"
                ? null
                : new ResizeObserver(() => {
                      measureViewportMetrics();
                      measureSheetTabsViewport();
                      measureSheetTabWidths();
                  });
        resizeObserver?.observe(document.body);
        if (gridViewportElement) {
            resizeObserver?.observe(gridViewportElement);
        }
        if (leftHeaderViewportElement) {
            resizeObserver?.observe(leftHeaderViewportElement);
        }
        if (rightHeaderViewportElement) {
            resizeObserver?.observe(rightHeaderViewportElement);
        }
        if (sheetTabsViewportElement) {
            resizeObserver?.observe(sheetTabsViewportElement);
        }

        onCleanup(() => {
            if (scrollFrameId) {
                cancelAnimationFrame(scrollFrameId);
            }
            window.removeEventListener("message", handleMessage);
            window.removeEventListener("resize", handleResize);
            document.removeEventListener("keydown", handleKeyDown);
            resizeObserver?.disconnect();
        });
    });

    const columnDiffTones = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return [] as Array<CellDiffStatus | null>;
        }

        const tones = Array<CellDiffStatus | null>(sheet.columnCount).fill(null);
        for (const row of sheet.rows) {
            for (const cell of row.cells) {
                const columnIndex = cell.columnNumber - 1;
                tones[columnIndex] = mergeDiffTone(tones[columnIndex] ?? null, cell.status);
            }
        }

        return tones;
    });

    const filteredRows = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return [] as DiffPanelRowView[];
        }

        return filterDiffRows(sheet.rows, rowFilter());
    });

    const totalRowCount = createMemo(() => filteredRows().length);
    const startRowIndex = createMemo(() =>
        Math.max(0, Math.floor(scrollTop() / ROW_HEIGHT) - ROW_OVERSCAN)
    );
    const visibleRowCount = createMemo(
        () => Math.ceil(Math.max(viewportHeight(), ROW_HEIGHT) / ROW_HEIGHT) + ROW_OVERSCAN * 2
    );
    const endRowIndex = createMemo(() =>
        Math.min(totalRowCount(), startRowIndex() + visibleRowCount())
    );
    const visibleRows = createMemo(() => filteredRows().slice(startRowIndex(), endRowIndex()));
    const visibleRowsOffsetY = createMemo(() => startRowIndex() * ROW_HEIGHT);
    const totalRowsHeight = createMemo(() => totalRowCount() * ROW_HEIGHT);
    const viewportContentHeight = createMemo(() => Math.max(totalRowsHeight(), viewportHeight()));

    const columnLayout = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return null;
        }

        return createPixelColumnLayout({
            columnWidths: sheet.columns.map((column) => column.columnWidth),
            maximumDigitWidth: maximumDigitWidth(),
            fallbackPixelWidth: DEFAULT_COLUMN_PIXEL_WIDTH,
        });
    });

    const visibleColumnWindow = createMemo(() => {
        const sheet = activeSheet();
        const layout = columnLayout();
        if (!sheet || !layout) {
            return {
                columns: [] as DiffPanelColumnView[],
                leadingSpacerWidth: 0,
                trailingSpacerWidth: 0,
                totalWidth: 0,
            };
        }

        const columnWindow = getPixelColumnWindow(
            layout,
            horizontalScrollLeft(),
            horizontalViewportWidth(),
            2
        );

        return {
            columns: sheet.columns.slice(columnWindow.startIndex, columnWindow.endIndex),
            leadingSpacerWidth: columnWindow.leadingSpacerWidth,
            trailingSpacerWidth: columnWindow.trailingSpacerWidth,
            totalWidth: layout.totalWidth,
        };
    });

    const horizontalTrackWidth = createMemo(() => visibleColumnWindow().totalWidth);
    const showHorizontalScrollbar = createMemo(
        () => horizontalTrackWidth() > horizontalViewportWidth() + 1
    );
    const horizontalTrackStyle = createMemo(() => ({
        width: `${horizontalTrackWidth()}px`,
        transform: `translateX(-${horizontalScrollLeft()}px)`,
    }));

    createEffect(() => {
        const nextTrackWidth = horizontalTrackWidth();
        const nextViewportWidth = horizontalViewportWidth();
        setHorizontalScrollLeft((current) =>
            clampDiffHorizontalScroll(current, nextTrackWidth, nextViewportWidth)
        );
    });

    let previousSheetKey: string | null = null;
    createEffect(() => {
        const nextSheetKey = activeSheet()?.key ?? null;
        if (nextSheetKey !== previousSheetKey) {
            previousSheetKey = nextSheetKey;
            setActiveDiffIndex(0);
            setPreviewExpanded(false);
            setEditingCell(null);
            setRowFilter("all");

            const sheet = activeSheet();
            if (sheet && sheet.rowCount > 0 && sheet.columnCount > 0) {
                const defaultSelection = getDefaultSelectionForFilter(sheet, "all");
                setSelectedCell(defaultSelection?.selection ?? null);
                if (defaultSelection) {
                    setSelectedSide(defaultSelection.side);
                }
            } else {
                setSelectedCell(null);
            }

            setHorizontalScrollLeft(0);
            setScrollTop(0);
            queueMicrotask(() => {
                measureViewportMetrics();
                measureSheetTabsViewport();
                measureSheetTabWidths();
                if (gridViewportElement) {
                    gridViewportElement.scrollTop = 0;
                }
                if (leftScrollbarElement) {
                    leftScrollbarElement.scrollLeft = 0;
                }
                if (rightScrollbarElement) {
                    rightScrollbarElement.scrollLeft = 0;
                }
            });
        }

        if (session().ui.pendingEdits.clearRequested) {
            setPendingEdits({});
        }
    });

    createEffect(() => {
        const rows = filteredRows();
        const sheet = activeSheet();
        const currentSelection = selectedCell();
        const currentSide = selectedSide();

        if (!sheet || rows.length === 0 || sheet.columnCount === 0) {
            if (currentSelection) {
                setSelectedCell(null);
            }
            return;
        }

        if (
            !currentSelection ||
            !rows.some((row) => row.rowNumber === currentSelection.rowNumber)
        ) {
            const firstRow = rows[0] ?? null;
            if (!firstRow) {
                return;
            }

            const side = getDefaultSideForRow(firstRow);
            setSelectedCell(createSelectedDiffCell(sheet, firstRow, firstRow.rowNumber, 1, side));
            setSelectedSide(side);
            return;
        }

        const row =
            rows.find((candidate) => candidate.rowNumber === currentSelection.rowNumber) ?? null;
        if (!row) {
            return;
        }

        if (currentSelection.columnNumber > sheet.columnCount) {
            setSelectedCell(
                createSelectedDiffCell(
                    sheet,
                    row,
                    currentSelection.rowNumber,
                    sheet.columnCount,
                    currentSide
                )
            );
            return;
        }

        const normalizedSelection = createSelectedDiffCell(
            sheet,
            row,
            currentSelection.rowNumber,
            currentSelection.columnNumber,
            currentSide
        );
        if (
            normalizedSelection.sourceRowNumber !== currentSelection.sourceRowNumber ||
            normalizedSelection.sourceColumnNumber !== currentSelection.sourceColumnNumber
        ) {
            setSelectedCell(normalizedSelection);
        }
    });

    createEffect(() => {
        const maxScrollTop = Math.max(0, totalRowsHeight() - viewportHeight());
        if (scrollTop() > maxScrollTop) {
            setScrollTop(maxScrollTop);
        }
    });

    createEffect(() => {
        const nextScrollLeft = horizontalScrollLeft();
        if (leftScrollbarElement && leftScrollbarElement.scrollLeft !== nextScrollLeft) {
            leftScrollbarElement.scrollLeft = nextScrollLeft;
        }
        if (rightScrollbarElement && rightScrollbarElement.scrollLeft !== nextScrollLeft) {
            rightScrollbarElement.scrollLeft = nextScrollLeft;
        }
    });

    createEffect(() => {
        const nextScrollTop = scrollTop();
        if (gridViewportElement && gridViewportElement.scrollTop !== nextScrollTop) {
            gridViewportElement.scrollTop = nextScrollTop;
        }
    });

    const ensureColumnVisible = (columnNumber: number) => {
        const layout = columnLayout();
        if (!layout || horizontalViewportWidth() <= 0) {
            return;
        }

        const cellLeft = getPixelColumnLeft(layout, columnNumber);
        const cellRight = getPixelColumnRight(layout, columnNumber);
        const viewportLeft = horizontalScrollLeft();
        const viewportRight = viewportLeft + horizontalViewportWidth();

        if (cellLeft < viewportLeft) {
            syncHorizontalScroll(cellLeft);
            return;
        }

        if (cellRight > viewportRight) {
            syncHorizontalScroll(cellRight - horizontalViewportWidth());
        }
    };

    const scrollToRow = (rowNumber: number) => {
        if (!gridViewportElement) {
            return;
        }

        const rowIndex = filteredRows().findIndex((row) => row.rowNumber === rowNumber);
        if (rowIndex < 0) {
            return;
        }

        const nextScrollTop = rowIndex * ROW_HEIGHT;
        gridViewportElement.scrollTo({
            top: nextScrollTop,
            behavior: "auto",
        });
        setScrollTop(nextScrollTop);
    };

    createEffect(() => {
        const currentSheetKey = activeSheet()?.key ?? null;
        const currentViewportWidth = horizontalViewportWidth();
        const currentSelection = selectedCell();
        if (!currentSheetKey || currentViewportWidth <= 0 || !currentSelection) {
            return;
        }

        ensureColumnVisible(currentSelection.columnNumber);
    });

    createEffect(() => {
        if (!isSheetOverflowOpen()) {
            return;
        }

        const handlePointerDown = (event: PointerEvent) => {
            const target = event.target;
            if (!(target instanceof Node)) {
                return;
            }
            if (sheetTabsOverflowElement?.contains(target)) {
                return;
            }
            setIsSheetOverflowOpen(false);
        };

        const handleKeyDown = (event: KeyboardEvent) => {
            if (event.key === "Escape") {
                setIsSheetOverflowOpen(false);
            }
        };

        document.addEventListener("pointerdown", handlePointerDown);
        document.addEventListener("keydown", handleKeyDown);

        onCleanup(() => {
            document.removeEventListener("pointerdown", handlePointerDown);
            document.removeEventListener("keydown", handleKeyDown);
        });
    });

    const sheetTabLayout = createMemo(() =>
        partitionSheetTabs(
            comparison().sheets,
            getMaxVisibleDiffSheetTabs(
                comparison().sheets,
                sheetTabsViewportWidth(),
                measuredSheetTabWidths()
            )
        )
    );

    const effectivePendingEdits = createMemo(() => {
        const draft = editingCell();
        return draft ? applyPendingDiffEdit(pendingEdits(), draft) : pendingEdits();
    });
    const pendingSummary = createMemo(() =>
        getPendingSummary(activeSheet()?.key ?? null, effectivePendingEdits())
    );
    const hasUnsavedChanges = createMemo(() => Object.keys(effectivePendingEdits()).length > 0);
    const canMoveDiff = createMemo(() => (activeSheet()?.diffCells.length ?? 0) > 0);
    const selectedColumnNumber = createMemo(() => selectedCell()?.columnNumber ?? null);
    const selectedAddress = createMemo(() =>
        getSelectionAddress(activeSheet(), selectedCell(), selectedSide())
    );
    const leftSelectionPreview = createMemo(() =>
        getSelectionPreviewValue(
            activeSheet(),
            selectedCell(),
            editingCell(),
            effectivePendingEdits(),
            "left"
        )
    );
    const rightSelectionPreview = createMemo(() =>
        getSelectionPreviewValue(
            activeSheet(),
            selectedCell(),
            editingCell(),
            effectivePendingEdits(),
            "right"
        )
    );
    const currentDiffLabel = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet || sheet.diffCells.length === 0) {
            return strings.none;
        }

        return `${Math.min(activeDiffIndex() + 1, sheet.diffCells.length)}/${sheet.diffCells.length}`;
    });
    const activeFilterLabel = createMemo(() => {
        switch (rowFilter()) {
            case "diffs":
                return strings.diffs;
            case "same":
                return strings.same;
            case "all":
            default:
                return strings.all;
        }
    });
    const structuralChanges = createMemo(() =>
        activeSheet() ? getStructuralChangeLabels(strings, activeSheet()!) : []
    );
    const workbookStructuralChanges = createMemo(() =>
        getWorkbookStructuralChangeLabels(strings, comparison().definedNamesChanged)
    );
    const shellStyle = createMemo(() => ({
        "--diff-row-header-width": `${getDiffRowHeaderWidth(activeSheet()?.rowCount ?? 0)}px`,
        "--diff-viewport-scrollbar-width": `${viewportScrollbarWidth()}px`,
    }));

    const moveDiff = (offset: number) => {
        if (editingCell()) {
            finalizeEditingCell("commit");
        }

        const sheet = activeSheet();
        if (!sheet || sheet.diffCells.length === 0) {
            return;
        }

        const nextIndex = getWrappedDiffIndex(activeDiffIndex(), sheet.diffCells.length, offset);
        const diffCell = sheet.diffCells[nextIndex];
        const row =
            sheet.rows.find((candidate) => candidate.rowNumber === diffCell.rowNumber) ?? null;
        const cell =
            row?.cells.find((candidate) => candidate.columnNumber === diffCell.columnNumber) ??
            null;
        const side = getPreferredSide(cell);

        setActiveDiffIndex(nextIndex);
        setSelectedCell(
            createSelectedDiffCell(sheet, row, diffCell.rowNumber, diffCell.columnNumber, side)
        );
        setSelectedSide(side);
        scrollToRow(diffCell.rowNumber);
    };

    const syncHorizontalScroll = (nextScrollLeft: number) => {
        setHorizontalScrollLeft(
            clampDiffHorizontalScroll(
                nextScrollLeft,
                horizontalTrackWidth(),
                horizontalViewportWidth()
            )
        );
    };

    const selectCell = (
        row: DiffPanelRowView,
        cell: DiffPanelSparseCellView,
        side: "left" | "right"
    ) => {
        if (editingCell()) {
            finalizeEditingCell("commit");
        }

        const sheet = activeSheet();
        if (!sheet) {
            return;
        }

        const nextSelection = getSelectedDiffCellState(sheet, row, cell, side);
        setSelectedCell(nextSelection.selectedCell);
        setSelectedSide(side);
        if (nextSelection.activeDiffIndex !== null) {
            setActiveDiffIndex(nextSelection.activeDiffIndex);
        }
    };

    const finalizeEditingCell = (disposition: "commit" | "cancel") => {
        const nextState = finalizeDiffCellEdit(pendingEdits(), editingCell(), disposition);
        setPendingEdits(nextState.pendingEdits);
        setEditingCell(nextState.editingCell);
    };

    const savePendingEdits = () => {
        const nextPendingEdits = effectivePendingEdits();
        setEditingCell(null);

        const message = createSaveDiffEditsMessage(nextPendingEdits);
        if (!message) {
            setPendingEdits(nextPendingEdits);
            return;
        }

        setPendingEdits({});
        vscode.postMessage(message);
    };

    const startEditingCell = (
        row: DiffPanelRowView,
        columnNumber: number,
        cell: DiffPanelSparseCellView | null,
        side: "left" | "right"
    ) => {
        if (editingCell()) {
            finalizeEditingCell("commit");
        }

        const sheet = activeSheet();
        if (!sheet) {
            return;
        }

        const selection = createSelectedDiffCell(sheet, row, row.rowNumber, columnNumber, side);
        const status = cell?.status ?? "equal";
        if (
            !canEditDiffCell(
                comparison(),
                sheet,
                side,
                selection.sourceRowNumber,
                selection.sourceColumnNumber,
                status
            )
        ) {
            return;
        }

        const nextEditingState = beginDiffCellEdit({
            activeSheetKey: sheet.key,
            side,
            sideEditable: true,
            pendingEdits: pendingEdits(),
            selection,
            cell,
        });
        if (!nextEditingState) {
            return;
        }

        setEditingCell(nextEditingState.editingCell);
        setSelectedCell(nextEditingState.selectedCell);
        setSelectedSide(side);
    };

    const getRenderedCellValue = (
        sheetKey: string,
        row: DiffPanelRowView,
        columnNumber: number,
        cell: DiffPanelSparseCellView | null,
        side: "left" | "right"
    ): string => {
        return getRenderedDiffCellValue(
            effectivePendingEdits(),
            sheetKey,
            row.rowNumber,
            columnNumber,
            cell,
            side
        );
    };

    const clearSelectedCell = () => {
        const sheet = activeSheet();
        const selection = selectedCell();
        if (!sheet || !selection || editingCell()) {
            return;
        }

        const row = sheet.rows.find((candidate) => candidate.rowNumber === selection.rowNumber) ?? null;
        const cell =
            row?.cells.find((candidate) => candidate.columnNumber === selection.columnNumber) ?? null;
        const status = cell?.status ?? "equal";
        const side = selectedSide();
        if (
            !canEditDiffCell(
                comparison(),
                sheet,
                side,
                selection.sourceRowNumber,
                selection.sourceColumnNumber,
                status
            )
        ) {
            return;
        }

        const modelValue = side === "left" ? (cell?.leftValue ?? "") : (cell?.rightValue ?? "");
        const key = getPendingDiffEditKey(
            sheet.key,
            selection.rowNumber,
            selection.columnNumber,
            side
        );

        setPendingEdits((currentPendingEdits) => {
            if (modelValue === "") {
                if (!Object.hasOwn(currentPendingEdits, key)) {
                    return currentPendingEdits;
                }

                const nextPendingEdits = { ...currentPendingEdits };
                delete nextPendingEdits[key];
                return nextPendingEdits;
            }

            return {
                ...currentPendingEdits,
                [key]: {
                    sheetKey: sheet.key,
                    side,
                    rowNumber: selection.rowNumber,
                    columnNumber: selection.columnNumber,
                    sourceRowNumber: selection.sourceRowNumber,
                    sourceColumnNumber: selection.sourceColumnNumber,
                    value: "",
                },
            };
        });
    };

    const handleFilterChange = (mode: RowFilterMode) => {
        if (editingCell()) {
            finalizeEditingCell("commit");
        }

        setRowFilter(mode);
        setScrollTop(0);
        if (gridViewportElement) {
            gridViewportElement.scrollTop = 0;
        }

        const sheet = activeSheet();
        if (!sheet) {
            return;
        }

        const defaultSelection = getDefaultSelectionForFilter(sheet, mode);
        setSelectedCell(defaultSelection?.selection ?? null);
        if (defaultSelection) {
            setSelectedSide(defaultSelection.side);
        }
        if (mode === "diffs") {
            setActiveDiffIndex(0);
        }
    };

    const handleSetSheet = (sheetKey: string) => {
        if (editingCell()) {
            finalizeEditingCell("commit");
        }

        vscode.postMessage({
            type: "setSheet",
            sheetKey,
        });
    };

    const renderHeaderPane = (side: "left" | "right") => (
        <div class="diff-paneHeader diff-pane">
            <div class="diff-indexCell diff-indexCell--header">#</div>
            <div
                class="diff-columnsViewport"
                ref={(element) => {
                    if (side === "left") {
                        leftHeaderViewportElement = element;
                    } else {
                        rightHeaderViewportElement = element;
                    }
                }}
            >
                <div class="diff-columnsTrack" style={horizontalTrackStyle()}>
                    <Show when={visibleColumnWindow().leadingSpacerWidth > 0}>
                        <div
                            class="diff-columnSpacer"
                            style={{
                                width: `${visibleColumnWindow().leadingSpacerWidth}px`,
                            }}
                            aria-hidden="true"
                        />
                    </Show>
                    <For each={visibleColumnWindow().columns}>
                        {(column) => {
                            const columnNumber = column.columnNumber;
                            const columnWidth = () => {
                                const layout = columnLayout();
                                return layout
                                    ? getPixelColumnWidth(layout, columnNumber)
                                    : DEFAULT_COLUMN_PIXEL_WIDTH;
                            };
                            const diffTone = () => columnDiffTones()[columnNumber - 1] ?? null;
                            const hasPending = () =>
                                pendingSummary().columnsBySide[side].has(columnNumber);
                            const markerTone = () =>
                                getEffectiveMarkerTone(diffTone(), hasPending());
                            const columnLabel =
                                side === "left" ? column.leftLabel : column.rightLabel;

                            return (
                                <div
                                    class={classNames([
                                        "diff-headerCell",
                                        diffTone() && diffTone() !== "equal"
                                            ? `diff-headerCell--${diffTone()}`
                                            : undefined,
                                        hasPending() && "diff-headerCell--pending",
                                        selectedColumnNumber() === columnNumber &&
                                            "diff-headerCell--active",
                                    ])}
                                    style={{
                                        width: `${columnWidth()}px`,
                                        "flex-basis": `${columnWidth()}px`,
                                        "--diff-column-width": `${columnWidth()}px`,
                                    }}
                                >
                                    <span class="diff-headerLabel">
                                        <DiffMarker tone={markerTone()} />
                                        <span>{columnLabel || String(columnNumber)}</span>
                                    </span>
                                </div>
                            );
                        }}
                    </For>
                    <Show when={visibleColumnWindow().trailingSpacerWidth > 0}>
                        <div
                            class="diff-columnSpacer"
                            style={{
                                width: `${visibleColumnWindow().trailingSpacerWidth}px`,
                            }}
                            aria-hidden="true"
                        />
                    </Show>
                </div>
            </div>
        </div>
    );

    const renderSideRow = (row: DiffPanelRowView, side: "left" | "right") => {
        const sheet = activeSheet();
        if (!sheet) {
            return null;
        }

        const rowToneClass = getRowToneClass(row);
        const isPendingRow = () => pendingSummary().rowsBySide[side].has(row.rowNumber);
        const rowMarkerTone = () => getEffectiveMarkerTone(row.diffTone, isPendingRow());
        const sourceRowNumber = side === "left" ? row.leftRowNumber : row.rightRowNumber;

        return (
            <div class="diff-sideRow diff-pane">
                <div
                    class={classNames([
                        "diff-indexCell",
                        rowToneClass,
                        isPendingRow() && "diff-indexCell--pending",
                        selectedCell()?.rowNumber === row.rowNumber && "diff-indexCell--selected",
                    ])}
                >
                    <span class="diff-indexLabel">
                        <DiffMarker tone={rowMarkerTone()} />
                        <span>{sourceRowNumber ?? ""}</span>
                    </span>
                </div>
                <div class="diff-rowCellsViewport">
                    <div class="diff-rowCellsTrack" style={horizontalTrackStyle()}>
                        <Show when={visibleColumnWindow().leadingSpacerWidth > 0}>
                            <div
                                class="diff-columnSpacer"
                                style={{
                                    width: `${visibleColumnWindow().leadingSpacerWidth}px`,
                                }}
                                aria-hidden="true"
                            />
                        </Show>
                        <For each={visibleColumnWindow().columns}>
                            {(column) => {
                                const resolvedCell = () =>
                                    findCellInSheet(
                                        activeSheet(),
                                        row.rowNumber,
                                        column.columnNumber
                                    );
                                const columnWidth = () => {
                                    const layout = columnLayout();
                                    return layout
                                        ? getPixelColumnWidth(layout, column.columnNumber)
                                        : DEFAULT_COLUMN_PIXEL_WIDTH;
                                };
                                const toneClass = () => {
                                    const cell = resolvedCell();
                                    return cell ? getCellToneClass(cell) : "";
                                };
                                const isSelected = () =>
                                    selectedSide() === side &&
                                    selectedCell()?.rowNumber === row.rowNumber &&
                                    selectedCell()?.columnNumber === column.columnNumber;
                                const isEditing = () =>
                                    editingCell()?.rowNumber === row.rowNumber &&
                                    editingCell()?.columnNumber === column.columnNumber &&
                                    editingCell()?.side === side;
                                const pendingKey = getPendingDiffEditKey(
                                    sheet.key,
                                    row.rowNumber,
                                    column.columnNumber,
                                    side
                                );
                                const displayValue = () => {
                                    const cell = resolvedCell();
                                    return getRenderedCellValue(
                                        sheet.key,
                                        row,
                                        column.columnNumber,
                                        cell,
                                        side
                                    );
                                };
                                const title = () => {
                                    const cell = resolvedCell();
                                    const sourceRowNumber =
                                        side === "left" ? row.leftRowNumber : row.rightRowNumber;
                                    const columnLabel =
                                        side === "left"
                                            ? (column.leftLabel ?? "")
                                            : (column.rightLabel ?? "");
                                    const formula = effectivePendingEdits()[pendingKey]
                                        ? null
                                        : side === "left"
                                          ? (cell?.leftFormula ?? null)
                                          : (cell?.rightFormula ?? null);
                                    return getDiffCellTitle(
                                        columnLabel,
                                        sourceRowNumber,
                                        displayValue(),
                                        formula
                                    );
                                };

                                return (
                                    <button
                                        class={classNames([
                                            "diff-cell",
                                            toneClass(),
                                            isSelected() && "diff-cell--active",
                                            Boolean(effectivePendingEdits()[pendingKey]) &&
                                                "diff-cell--pending",
                                            isEditing() && "diff-cell--editing",
                                        ])}
                                        style={{
                                            width: `${columnWidth()}px`,
                                            "flex-basis": `${columnWidth()}px`,
                                            "--diff-column-width": `${columnWidth()}px`,
                                        }}
                                        data-role="diff-cell"
                                        data-side={side}
                                        data-row-number={row.rowNumber}
                                        data-column-number={column.columnNumber}
                                        title={title()}
                                        type="button"
                                        onClick={() => {
                                            const sheet = activeSheet();
                                            if (!sheet) {
                                                return;
                                            }

                                            if (editingCell()) {
                                                finalizeEditingCell("commit");
                                            }

                                            const cell = resolvedCell();
                                            if (!cell) {
                                                setSelectedCell(
                                                    createSelectedDiffCell(
                                                        sheet,
                                                        row,
                                                        row.rowNumber,
                                                        column.columnNumber,
                                                        side
                                                    )
                                                );
                                                setSelectedSide(side);
                                                return;
                                            }

                                            selectCell(row, cell, side);
                                        }}
                                        onDblClick={() => {
                                            startEditingCell(
                                                row,
                                                column.columnNumber,
                                                resolvedCell(),
                                                side
                                            );
                                        }}
                                    >
                                        <Show
                                            when={isEditing()}
                                            fallback={
                                                <span class="diff-cell__text">
                                                    {displayValue()}
                                                </span>
                                            }
                                        >
                                            <CellEditor
                                                edit={editingCell()!}
                                                onUpdateDraft={(value) =>
                                                    setEditingCell((current) =>
                                                        current
                                                            ? {
                                                                  ...current,
                                                                  value,
                                                              }
                                                            : current
                                                    )
                                                }
                                                onCommit={() => finalizeEditingCell("commit")}
                                                onCancel={() => finalizeEditingCell("cancel")}
                                            />
                                        </Show>
                                    </button>
                                );
                            }}
                        </For>
                        <Show when={visibleColumnWindow().trailingSpacerWidth > 0}>
                            <div
                                class="diff-columnSpacer"
                                style={{
                                    width: `${visibleColumnWindow().trailingSpacerWidth}px`,
                                }}
                                aria-hidden="true"
                            />
                        </Show>
                    </div>
                </div>
            </div>
        );
    };

    const renderPreviewPane = (side: "left" | "right") => {
        const previewValue = () =>
            side === "left" ? leftSelectionPreview() : rightSelectionPreview();
        const peerPreview = () =>
            side === "left" ? rightSelectionPreview() : leftSelectionPreview();

        return (
            <div
                class={classNames([
                    "diff-selectionPreviewPane",
                    "diff-pane",
                    selectedSide() === side && "diff-selectionPreviewPane--active",
                    previewExpanded() && "diff-selectionPreviewPane--expanded",
                ])}
                title={previewValue().value || strings.emptyValue}
            >
                <div style={{ flex: "1 1 auto", "min-width": 0 }}>
                    <SelectionValueContent
                        value={previewValue().value}
                        otherValue={peerPreview().value}
                        class="diff-selectionPreviewPane__value"
                        emptyValueLabel={strings.emptyValue}
                    />
                </div>
                <Show when={selectedCell() && selectedSide() === side}>
                    <button
                        type="button"
                        class="diff-button diff-selectionPreviewPane__action"
                        aria-label={previewExpanded() ? strings.close : strings.viewDetails}
                        aria-expanded={previewExpanded()}
                        title={previewExpanded() ? strings.close : strings.viewDetails}
                        onClick={() => setPreviewExpanded((current) => !current)}
                    >
                        <span
                            class={classNames([
                                "codicon",
                                previewExpanded() ? "codicon-screen-normal" : "codicon-screen-full",
                            ])}
                            aria-hidden="true"
                        />
                    </button>
                </Show>
            </div>
        );
    };

    const message = () => {
        const current = session();
        if (current.mode === "error") {
            return current.ui.panel.statusMessage ?? strings.loading;
        }

        return current.ui.panel.statusMessage ?? strings.loading;
    };

    return (
        <Show
            when={session().mode === "ready"}
            fallback={
                session().mode === "error" ? (
                    <div class="diff-errorState">
                        <div class="diff-errorState__message">{message()}</div>
                        <button
                            type="button"
                            class="diff-button"
                            onClick={() => vscode.postMessage({ type: "reload" })}
                        >
                            {strings.reload}
                        </button>
                    </div>
                ) : (
                    <div class="diff-loading">{message()}</div>
                )
            }
        >
            <div class="diff-shell" style={shellStyle()}>
                <header class="diff-toolbarBar">
                    <div class="diff-toolbarGroup">
                        <div class="diff-chipGroup">
                            <For each={["all", "diffs", "same"] as const}>
                                {(mode) => (
                                    <button
                                        type="button"
                                        class={classNames([
                                            "diff-chip",
                                            rowFilter() === mode && "diff-chip--active",
                                        ])}
                                        onClick={() => handleFilterChange(mode)}
                                    >
                                        <span
                                            class={classNames([
                                                "codicon",
                                                mode === "all"
                                                    ? "codicon-list-flat"
                                                    : mode === "diffs"
                                                      ? "codicon-diff-multiple"
                                                      : "codicon-check-all",
                                                "diff-controlIcon",
                                            ])}
                                            aria-hidden="true"
                                        />
                                        <span>
                                            {mode === "all"
                                                ? strings.all
                                                : mode === "diffs"
                                                  ? strings.diffs
                                                  : strings.same}
                                        </span>
                                    </button>
                                )}
                            </For>
                        </div>
                    </div>
                    <div class="diff-toolbarGroup diff-toolbarGroup--actions">
                        <div class="diff-actionGroup">
                            <button
                                type="button"
                                class="diff-button"
                                disabled={!canMoveDiff()}
                                onClick={() => moveDiff(-1)}
                            >
                                <span
                                    class="codicon codicon-arrow-up diff-controlIcon"
                                    aria-hidden="true"
                                />
                                <span>{strings.prevDiff}</span>
                            </button>
                            <button
                                type="button"
                                class="diff-button"
                                disabled={!canMoveDiff()}
                                onClick={() => moveDiff(1)}
                            >
                                <span
                                    class="codicon codicon-arrow-down diff-controlIcon"
                                    aria-hidden="true"
                                />
                                <span>{strings.nextDiff}</span>
                            </button>
                            <button
                                type="button"
                                class="diff-button"
                                onClick={() => vscode.postMessage({ type: "swap" })}
                            >
                                <span
                                    class="codicon codicon-arrow-swap diff-controlIcon"
                                    aria-hidden="true"
                                />
                                <span>{strings.swap}</span>
                            </button>
                            <button
                                type="button"
                                class="diff-button"
                                onClick={() => vscode.postMessage({ type: "reload" })}
                            >
                                <span
                                    class="codicon codicon-refresh diff-controlIcon"
                                    aria-hidden="true"
                                />
                                <span>{strings.reload}</span>
                            </button>
                            <button
                                type="button"
                                class={classNames([
                                    "diff-button",
                                    hasUnsavedChanges() && "diff-button--active",
                                ])}
                                disabled={!hasUnsavedChanges()}
                                onClick={savePendingEdits}
                            >
                                <span
                                    class="codicon codicon-save diff-controlIcon"
                                    aria-hidden="true"
                                />
                                <span>{strings.save}</span>
                            </button>
                        </div>
                    </div>
                </header>

                <section class="diff-gridSection">
                    <div class="diff-paneMetaRow">
                        <div class="diff-pane">
                            <PaneMeta file={comparison().leftFile} strings={strings} />
                        </div>
                        <div class="diff-divider" />
                        <div class="diff-pane">
                            <PaneMeta file={comparison().rightFile} strings={strings} />
                        </div>
                    </div>

                    <Show when={activeSheet()} fallback={<EmptyState message={strings.noSheet} />}>
                        {(sheet) => (
                            <Show
                                when={totalRowCount() > 0 && sheet().columnCount > 0}
                                fallback={<EmptyState message={strings.noRows} />}
                            >
                                <div class="diff-gridShell">
                                    <div class="diff-gridHeaderRow">
                                        {renderHeaderPane("left")}
                                        <div class="diff-divider" />
                                        {renderHeaderPane("right")}
                                    </div>

                                    <div
                                        class="diff-gridViewport"
                                        style={{ "overflow-x": "hidden" }}
                                        ref={(element) => {
                                            gridViewportElement = element;
                                            measureViewportMetrics();
                                        }}
                                        onScroll={(event) => {
                                            const nextScrollTop = event.currentTarget.scrollTop;
                                            if (scrollFrameId) {
                                                cancelAnimationFrame(scrollFrameId);
                                            }

                                            scrollFrameId = requestAnimationFrame(() => {
                                                scrollFrameId = 0;
                                                setScrollTop(nextScrollTop);
                                            });
                                        }}
                                        onWheel={(event) => {
                                            if (Math.abs(event.deltaX) <= Math.abs(event.deltaY)) {
                                                return;
                                            }

                                            event.preventDefault();
                                            syncHorizontalScroll(
                                                horizontalScrollLeft() + event.deltaX
                                            );
                                        }}
                                    >
                                        <div
                                            class="diff-gridViewportInner"
                                            style={{ height: `${viewportContentHeight()}px` }}
                                        >
                                            <div class="diff-gridBackdrop" aria-hidden="true">
                                                <div />
                                                <div class="diff-gridBackdrop__divider" />
                                                <div />
                                            </div>
                                            <div
                                                class="diff-visibleRows"
                                                style={{
                                                    transform: `translateY(${visibleRowsOffsetY()}px)`,
                                                }}
                                            >
                                                <For each={visibleRows()}>
                                                    {(row) => (
                                                        <div
                                                            class="diff-pairRow"
                                                            style={{ height: `${ROW_HEIGHT}px` }}
                                                        >
                                                            {renderSideRow(row, "left")}
                                                            <div class="diff-divider" />
                                                            {renderSideRow(row, "right")}
                                                        </div>
                                                    )}
                                                </For>
                                            </div>
                                        </div>
                                    </div>

                                    <div class="diff-gridScrollbarRow">
                                        <div
                                            class={classNames([
                                                "diff-scrollbarRow",
                                                "diff-pane",
                                                !showHorizontalScrollbar() &&
                                                    "diff-scrollbarRow--inactive",
                                            ])}
                                        >
                                            <div class="diff-scrollbarSpacer" />
                                            <div
                                                class={classNames([
                                                    "diff-scrollbar",
                                                    !showHorizontalScrollbar() &&
                                                        "diff-scrollbar--inactive",
                                                ])}
                                                ref={(element) => {
                                                    leftScrollbarElement = element;
                                                }}
                                                onScroll={(event) =>
                                                    syncHorizontalScroll(
                                                        event.currentTarget.scrollLeft
                                                    )
                                                }
                                            >
                                                <div
                                                    style={{
                                                        width: `${horizontalTrackWidth()}px`,
                                                        height: "1px",
                                                    }}
                                                />
                                            </div>
                                        </div>
                                        <div class="diff-divider" />
                                        <div
                                            class={classNames([
                                                "diff-scrollbarRow",
                                                "diff-pane",
                                                !showHorizontalScrollbar() &&
                                                    "diff-scrollbarRow--inactive",
                                            ])}
                                        >
                                            <div class="diff-scrollbarSpacer" />
                                            <div
                                                class={classNames([
                                                    "diff-scrollbar",
                                                    !showHorizontalScrollbar() &&
                                                        "diff-scrollbar--inactive",
                                                ])}
                                                ref={(element) => {
                                                    rightScrollbarElement = element;
                                                }}
                                                onScroll={(event) =>
                                                    syncHorizontalScroll(
                                                        event.currentTarget.scrollLeft
                                                    )
                                                }
                                            >
                                                <div
                                                    style={{
                                                        width: `${horizontalTrackWidth()}px`,
                                                        height: "1px",
                                                    }}
                                                />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </Show>
                        )}
                    </Show>

                    <Show when={activeSheet()}>
                        <div
                            class={classNames([
                                "diff-selectionPreviewRow",
                                previewExpanded() && "diff-selectionPreviewRow--expanded",
                            ])}
                        >
                            {renderPreviewPane("left")}
                            <div class="diff-divider" />
                            {renderPreviewPane("right")}
                        </div>
                    </Show>
                </section>

                <nav class="diff-sheetTabs diff-sheetTabs--bottom" aria-label={strings.sheets}>
                    <div
                        class="diff-sheetTabs__viewport"
                        ref={(element) => {
                            sheetTabsViewportElement = element;
                        }}
                    >
                        <div class="diff-sheetTabs__content">
                            <div class="diff-sheetTabs__list">
                                <For each={sheetTabLayout().visibleTabs}>
                                    {(sheet) => {
                                        const hasPending = pendingSummary().sheetKeys.has(
                                            sheet.key
                                        );
                                        const markerTone = getEffectiveMarkerTone(
                                            sheet.hasDiff ? sheet.diffTone : null,
                                            hasPending
                                        );

                                        return (
                                            <button
                                                type="button"
                                                class={classNames([
                                                    "diff-sheetTab",
                                                    sheet.isActive && "diff-sheetTab--active",
                                                    hasPending && "diff-sheetTab--pending",
                                                ])}
                                                title={getSheetTooltip(strings, sheet)}
                                                onClick={() => handleSetSheet(sheet.key)}
                                            >
                                                <DiffMarker
                                                    tone={markerTone}
                                                    class="diff-sheetTab__marker"
                                                />
                                                <span class="diff-sheetTab__label">
                                                    {sheet.label}
                                                </span>
                                            </button>
                                        );
                                    }}
                                </For>
                            </div>
                            <Show when={sheetTabLayout().hasOverflow}>
                                <div
                                    class="diff-sheetTabs__overflow"
                                    ref={(element) => {
                                        sheetTabsOverflowElement = element;
                                    }}
                                >
                                    <button
                                        aria-label={strings.moreSheets}
                                        aria-expanded={isSheetOverflowOpen()}
                                        aria-haspopup="menu"
                                        class={classNames([
                                            "diff-sheetTab",
                                            "diff-sheetTab--overflowTrigger",
                                            isSheetOverflowOpen() &&
                                                "diff-sheetTab--overflowActive",
                                        ])}
                                        title={strings.moreSheets}
                                        type="button"
                                        onClick={() =>
                                            setIsSheetOverflowOpen((current) => !current)
                                        }
                                    >
                                        <span
                                            class="codicon codicon-more diff-sheetTab__icon"
                                            aria-hidden="true"
                                        />
                                        <span
                                            class="diff-sheetTabs__overflowCount"
                                            aria-hidden="true"
                                        >
                                            {sheetTabLayout().overflowTabs.length}
                                        </span>
                                    </button>
                                    <Show when={isSheetOverflowOpen()}>
                                        <div class="diff-sheetTabs__overflowMenu" role="menu">
                                            <For each={sheetTabLayout().overflowTabs}>
                                                {(sheet) => {
                                                    const hasPending =
                                                        pendingSummary().sheetKeys.has(sheet.key);
                                                    const markerTone = getEffectiveMarkerTone(
                                                        sheet.hasDiff ? sheet.diffTone : null,
                                                        hasPending
                                                    );

                                                    return (
                                                        <button
                                                            class="diff-sheetTabs__overflowItem"
                                                            role="menuitem"
                                                            title={getSheetTooltip(strings, sheet)}
                                                            type="button"
                                                            onClick={() => {
                                                                setIsSheetOverflowOpen(false);
                                                                handleSetSheet(sheet.key);
                                                            }}
                                                        >
                                                            <DiffMarker
                                                                tone={markerTone}
                                                                class="diff-sheetTab__marker"
                                                            />
                                                            <span class="diff-sheetTabs__overflowLabel">
                                                                {sheet.label}
                                                            </span>
                                                        </button>
                                                    );
                                                }}
                                            </For>
                                        </div>
                                    </Show>
                                </div>
                            </Show>
                        </div>
                    </div>
                    <div
                        class="diff-sheetTabs__measure"
                        aria-hidden="true"
                        ref={(element) => {
                            sheetTabsMeasureElement = element;
                        }}
                    >
                        <For each={comparison().sheets}>
                            {(sheet) => {
                                const hasPending = pendingSummary().sheetKeys.has(sheet.key);
                                const markerTone = getEffectiveMarkerTone(
                                    sheet.hasDiff ? sheet.diffTone : null,
                                    hasPending
                                );

                                return (
                                    <button
                                        class={classNames([
                                            "diff-sheetTab",
                                            "diff-sheetTabs__measureTab",
                                            sheet.isActive && "diff-sheetTab--active",
                                            hasPending && "diff-sheetTab--pending",
                                        ])}
                                        data-role="diff-sheet-tab-measure"
                                        data-sheet-key={sheet.key}
                                        tabIndex={-1}
                                        type="button"
                                    >
                                        <DiffMarker
                                            tone={markerTone}
                                            class="diff-sheetTab__marker"
                                        />
                                        <span class="diff-sheetTab__label">{sheet.label}</span>
                                    </button>
                                );
                            }}
                        </For>
                    </div>
                </nav>

                <footer class="diff-statusBar">
                    <span>{`${strings.sheets} ${activeSheet()?.label ?? strings.none}`}</span>
                    <span>{`${strings.rows} ${activeSheet()?.rowCount ?? 0}`}</span>
                    <span>{`${strings.filter} ${activeFilterLabel()}`}</span>
                    <span>{`${strings.diffRows} ${activeSheet()?.diffRowCount ?? 0}`}</span>
                    <span>{`${strings.visibleRows} ${filteredRows().length}`}</span>
                    <span>{`${strings.currentDiff} ${currentDiffLabel()}`}</span>
                    <span>{`${strings.selected} ${selectedAddress()}`}</span>
                    <span>{`${strings.structure} ${structuralChanges().join(", ") || strings.none}`}</span>
                    <span>
                        {`${strings.workbookStructure} ${workbookStructuralChanges().join(", ") || strings.none}`}
                    </span>
                </footer>
            </div>
        </Show>
    );
}

const rootElement = document.getElementById("app");
if (rootElement) {
    render(() => <DiffBootstrapApp />, rootElement);
}
