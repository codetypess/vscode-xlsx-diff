import { For, Show, createEffect, createMemo, createSignal, onCleanup, onMount } from "solid-js";
import { render } from "solid-js/web";
import type { DiffPanelStrings } from "../../i18n";
import type {
    DiffPanelColumnView,
    DiffPanelRowView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "../../webview/diff-panel/diff-panel-types";
import {
    createWebviewReadyMessage,
    isDiffSessionIncomingMessage,
    type DiffSessionIncomingMessage,
    type DiffWebviewPendingEdit,
    type DiffWebviewOutgoingMessage,
} from "../shared/session-protocol";
import { createInitialDiffSessionState, reduceDiffSessionMessage } from "./session";
import {
    beginDiffCellEdit,
    clampDiffHorizontalScroll,
    createSaveDiffEditsMessage,
    finalizeDiffCellEdit,
    filterDiffRows,
    getDiffPreviewState,
    getPendingDiffEditKey,
    getRenderedDiffCellValue,
    getSelectedDiffCellState,
    getDiffTrackWidth,
    getWrappedDiffIndex,
    type DiffPreviewState,
    type DraftEditState,
    type RowFilterMode,
    type SelectedDiffCell,
} from "./grid-helpers";

interface VsCodeApi {
    postMessage(message: DiffWebviewOutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

const vscode = acquireVsCodeApi();

function getDiffStrings(): Pick<
    DiffPanelStrings,
    | "all"
    | "currentDiff"
    | "definedNames"
    | "diffCells"
    | "diffRows"
    | "diffs"
    | "emptyValue"
    | "loading"
    | "nextDiff"
    | "noSheet"
    | "none"
    | "prevDiff"
    | "readOnly"
    | "reload"
    | "same"
    | "save"
    | "selected"
    | "size"
    | "swap"
    | "viewDetails"
> {
    const strings = (globalThis as Record<string, unknown>).__XLSX_DIFF_STRINGS__ as
        | Partial<DiffPanelStrings>
        | undefined;
    return {
        all: strings?.all ?? "All",
        currentDiff: strings?.currentDiff ?? "Current Diff",
        definedNames: strings?.definedNames ?? "Defined Names",
        diffCells: strings?.diffCells ?? "Diff Cells",
        diffRows: strings?.diffRows ?? "Diff Rows",
        diffs: strings?.diffs ?? "Diffs",
        emptyValue: strings?.emptyValue ?? "(empty)",
        loading: strings?.loading ?? "Loading diff...",
        nextDiff: strings?.nextDiff ?? "Next Diff",
        noSheet: strings?.noSheet ?? "No sheet selected.",
        none: strings?.none ?? "None",
        prevDiff: strings?.prevDiff ?? "Prev Diff",
        readOnly: strings?.readOnly ?? "Read-only",
        reload: strings?.reload ?? "Reload",
        same: strings?.same ?? "Same",
        save: strings?.save ?? "Save",
        selected: strings?.selected ?? "Selected",
        size: strings?.size ?? "Size",
        swap: strings?.swap ?? "Swap",
        viewDetails: strings?.viewDetails ?? "View Details",
    };
}

function getDisplayValue(value: string, emptyValueLabel: string): string {
    return value.length > 0 ? value : emptyValueLabel;
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

function isSideEditable(
    side: "left" | "right",
    comparison: ReturnType<typeof createInitialDiffSessionState>["document"]["comparison"]
): boolean {
    return side === "left" ? !comparison.leftFile?.isReadonly : !comparison.rightFile?.isReadonly;
}

function DiffBootstrapApp() {
    let leftHeaderViewportElement: HTMLDivElement | undefined;
    let rightHeaderViewportElement: HTMLDivElement | undefined;
    let leftScrollbarElement: HTMLDivElement | undefined;
    let rightScrollbarElement: HTMLDivElement | undefined;

    const [session, setSession] = createSignal(createInitialDiffSessionState());
    const [rowFilter, setRowFilter] = createSignal<RowFilterMode>("diffs");
    const [activeDiffIndex, setActiveDiffIndex] = createSignal(0);
    const [previewExpanded, setPreviewExpanded] = createSignal(false);
    const [selectedCell, setSelectedCell] = createSignal<SelectedDiffCell | null>(null);
    const [editingCell, setEditingCell] = createSignal<DraftEditState | null>(null);
    const [pendingEdits, setPendingEdits] = createSignal<Record<string, DiffWebviewPendingEdit>>(
        {}
    );
    const [horizontalScrollLeft, setHorizontalScrollLeft] = createSignal(0);
    const [horizontalViewportWidth, setHorizontalViewportWidth] = createSignal(0);
    const strings = getDiffStrings();

    const measureHorizontalViewportWidth = () => {
        const nextWidth =
            leftHeaderViewportElement?.clientWidth ?? rightHeaderViewportElement?.clientWidth ?? 0;
        setHorizontalViewportWidth(nextWidth);
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
            measureHorizontalViewportWidth();
        };

        window.addEventListener("message", handleMessage);
        window.addEventListener("resize", handleResize);
        vscode.postMessage(createWebviewReadyMessage());
        measureHorizontalViewportWidth();

        onCleanup(() => {
            window.removeEventListener("message", handleMessage);
            window.removeEventListener("resize", handleResize);
        });
    });

    const comparison = () => session().document.comparison;
    const activeSheet = () => comparison().activeSheet;

    createEffect(() => {
        session().ui.navigation.activeSheetKey;
        setActiveDiffIndex(0);
        setPreviewExpanded(false);
        setSelectedCell(null);
        setEditingCell(null);
        queueMicrotask(measureHorizontalViewportWidth);
        if (session().ui.pendingEdits.clearRequested) {
            setPendingEdits({});
        }
    });

    const rowStats = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return {
                all: 0,
                diffs: 0,
                same: 0,
            };
        }

        return {
            all: sheet.rowCount,
            diffs: sheet.diffRowCount,
            same: Math.max(sheet.rowCount - sheet.diffRowCount, 0),
        };
    });

    const preview = createMemo(() => getDiffPreviewState(activeSheet(), activeDiffIndex()));

    createEffect(() => {
        const nextPreview = preview();
        if (nextPreview) {
            setSelectedCell({
                rowNumber: nextPreview.rowNumber,
                columnNumber: nextPreview.columnNumber,
            });
        }
    });

    const filteredRows = createMemo(() => {
        const sheet = activeSheet();
        if (!sheet) {
            return [] as DiffPanelRowView[];
        }

        return filterDiffRows(sheet.rows, rowFilter());
    });

    const visibleColumns = createMemo(() => activeSheet()?.columns ?? []);
    const horizontalTrackWidth = createMemo(() => getDiffTrackWidth(visibleColumns()));
    const showHorizontalScrollbar = createMemo(
        () => horizontalTrackWidth() > horizontalViewportWidth() + 1
    );
    const horizontalTrackStyle = createMemo(() => ({
        transform: `translateX(-${horizontalScrollLeft()}px)`,
    }));

    createEffect(() => {
        const nextTrackWidth = horizontalTrackWidth();
        const nextViewportWidth = horizontalViewportWidth();
        setHorizontalScrollLeft((current) =>
            clampDiffHorizontalScroll(current, nextTrackWidth, nextViewportWidth)
        );
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

    const canMoveDiff = createMemo(() => (activeSheet()?.diffCells.length ?? 0) > 1);

    const moveDiff = (offset: number) => {
        const sheet = activeSheet();
        if (!sheet || sheet.diffCells.length === 0) {
            return;
        }

        setActiveDiffIndex((current) =>
            getWrappedDiffIndex(current, sheet.diffCells.length, offset)
        );
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

    const selectCell = (row: DiffPanelRowView, cell: DiffPanelSparseCellView) => {
        const nextSelection = getSelectedDiffCellState(row, cell);
        setSelectedCell(nextSelection.selectedCell);
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
        const message = createSaveDiffEditsMessage(pendingEdits());
        if (!message) {
            return;
        }

        vscode.postMessage(message);
    };

    const startEditingCell = (
        row: DiffPanelRowView,
        cell: DiffPanelSparseCellView,
        side: "left" | "right"
    ) => {
        const nextEditingState = beginDiffCellEdit({
            activeSheetKey: activeSheet()?.key ?? null,
            side,
            sideEditable: isSideEditable(side, comparison()),
            pendingEdits: pendingEdits(),
            row,
            cell,
        });
        if (!nextEditingState) {
            return;
        }

        setEditingCell(nextEditingState.editingCell);
        setSelectedCell(nextEditingState.selectedCell);
    };

    const getRenderedCellValue = (
        row: DiffPanelRowView,
        cell: DiffPanelSparseCellView,
        side: "left" | "right"
    ): string => {
        return getRenderedDiffCellValue(pendingEdits(), row, cell, side);
    };

    const message = () => {
        const current = session();
        if (current.mode === "ready") {
            const title = current.document.comparison.title ?? "diff";
            return `SolidJS diff session initialized for ${title} (${current.renderCount} render${current.renderCount === 1 ? "" : "s"}).`;
        }

        return current.ui.panel.statusMessage ?? strings.loading;
    };

    return (
        <Show
            when={session().mode === "ready"}
            fallback={<div class="diff-loading">{message()}</div>}
        >
            <div class="diff-shell">
                <div class="diff-toolbarBar">
                    <div class="diff-toolbarGroup">
                        <div class="diff-actionGroup">
                            <button
                                class="diff-button"
                                type="button"
                                onClick={() => vscode.postMessage({ type: "reload" })}
                            >
                                {strings.reload}
                            </button>
                            <button
                                class="diff-button"
                                type="button"
                                onClick={() => vscode.postMessage({ type: "swap" })}
                            >
                                {strings.swap}
                            </button>
                        </div>
                        <span class="diff-separator">|</span>
                        <div class="diff-statusBar">
                            <span>{`${strings.diffCells}: ${activeSheet()?.diffCellCount ?? 0}`}</span>
                            <span>{`${strings.diffRows}: ${activeSheet()?.diffRowCount ?? 0}`}</span>
                            <span>{`${strings.definedNames}: ${comparison().definedNamesChanged ? strings.diffs : strings.same}`}</span>
                        </div>
                    </div>
                    <div class="diff-toolbarGroup diff-toolbarGroup--actions">
                        <div class="diff-chipGroup">
                            <For each={["all", "diffs", "same"] as const}>
                                {(mode) => (
                                    <button
                                        class="diff-chip"
                                        classList={{ "diff-chip--active": rowFilter() === mode }}
                                        type="button"
                                        onClick={() => setRowFilter(mode)}
                                    >
                                        {mode === "all"
                                            ? `${strings.all} ${rowStats().all}`
                                            : mode === "diffs"
                                              ? `${strings.diffs} ${rowStats().diffs}`
                                              : `${strings.same} ${rowStats().same}`}
                                    </button>
                                )}
                            </For>
                        </div>
                        <div class="diff-actionGroup">
                            <button
                                class="diff-button"
                                type="button"
                                disabled={Object.keys(pendingEdits()).length === 0}
                                onClick={savePendingEdits}
                            >
                                {strings.save}
                                {Object.keys(pendingEdits()).length > 0
                                    ? `(${Object.keys(pendingEdits()).length})`
                                    : ""}
                            </button>
                            <button
                                class="diff-button"
                                type="button"
                                disabled={!canMoveDiff()}
                                onClick={() => moveDiff(-1)}
                            >
                                {strings.prevDiff}
                            </button>
                            <button
                                class="diff-button"
                                type="button"
                                disabled={!canMoveDiff()}
                                onClick={() => moveDiff(1)}
                            >
                                {strings.nextDiff}
                            </button>
                        </div>
                    </div>
                </div>

                <div class="diff-sheetTabs">
                    <div class="diff-sheetTabs__viewport">
                        <div class="diff-sheetTabs__content">
                            <div class="diff-sheetTabs__list">
                                <For each={comparison().sheets}>
                                    {(sheet) => (
                                        <button
                                            class="diff-sheetTab"
                                            classList={{
                                                "diff-sheetTab--active":
                                                    session().ui.navigation.activeSheetKey ===
                                                    sheet.key,
                                            }}
                                            type="button"
                                            onClick={() =>
                                                vscode.postMessage({
                                                    type: "setSheet",
                                                    sheetKey: sheet.key,
                                                })
                                            }
                                        >
                                            <span class="diff-sheetTab__label">{sheet.label}</span>
                                        </button>
                                    )}
                                </For>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="diff-gridSection">
                    <div class="diff-paneMetaRow">
                        <div class="diff-paneMeta diff-pane">
                            <div class="diff-paneMeta__name">
                                <span class="diff-paneMeta__nameText">
                                    {comparison().leftFile?.title ?? strings.none}
                                </span>
                            </div>
                            <div class="diff-paneMeta__meta">
                                <div class="diff-paneMeta__path">
                                    {comparison().leftFile?.path ?? strings.none}
                                </div>
                                <div class="diff-paneMeta__facts">
                                    <span>{`${strings.size}: ${comparison().leftFile?.sizeLabel ?? strings.none}`}</span>
                                    <Show when={comparison().leftFile?.isReadonly}>
                                        <span>{strings.readOnly}</span>
                                    </Show>
                                </div>
                            </div>
                        </div>
                        <div class="diff-divider" />
                        <div class="diff-paneMeta diff-pane">
                            <div class="diff-paneMeta__name">
                                <span class="diff-paneMeta__nameText">
                                    {comparison().rightFile?.title ?? strings.none}
                                </span>
                            </div>
                            <div class="diff-paneMeta__meta">
                                <div class="diff-paneMeta__path">
                                    {comparison().rightFile?.path ?? strings.none}
                                </div>
                                <div class="diff-paneMeta__facts">
                                    <span>{`${strings.size}: ${comparison().rightFile?.sizeLabel ?? strings.none}`}</span>
                                    <Show when={comparison().rightFile?.isReadonly}>
                                        <span>{strings.readOnly}</span>
                                    </Show>
                                </div>
                            </div>
                        </div>
                    </div>

                    <Show
                        when={activeSheet()}
                        fallback={
                            <div class="diff-emptyState">
                                <div class="diff-emptyState__message">{strings.noSheet}</div>
                            </div>
                        }
                    >
                        <div
                            class="diff-selectionPreviewRow"
                            classList={{ "diff-selectionPreviewRow--expanded": previewExpanded() }}
                        >
                            <div
                                class="diff-selectionPreviewPane diff-pane"
                                classList={{
                                    "diff-selectionPreviewPane--active": Boolean(preview()),
                                    "diff-selectionPreviewPane--expanded": previewExpanded(),
                                }}
                            >
                                <div class="diff-selectionPreviewPane__value">
                                    <div class="diff-statusBar">
                                        <span>{`${strings.selected}: ${preview()?.address ?? strings.none}`}</span>
                                        <span>
                                            {preview()
                                                ? `${strings.currentDiff}: ${preview()!.index + 1}/${preview()!.total}`
                                                : `${strings.currentDiff}: ${strings.none}`}
                                        </span>
                                    </div>
                                    {preview()
                                        ? getDisplayValue(preview()!.leftValue, strings.emptyValue)
                                        : strings.none}
                                </div>
                                <button
                                    class="diff-selectionPreviewPane__action"
                                    type="button"
                                    title={strings.viewDetails}
                                    onClick={() => setPreviewExpanded((current) => !current)}
                                >
                                    <span
                                        class={`codicon ${previewExpanded() ? "codicon-chevron-up" : "codicon-chevron-down"}`}
                                    />
                                </button>
                            </div>
                            <div class="diff-divider" />
                            <div
                                class="diff-selectionPreviewPane diff-pane"
                                classList={{
                                    "diff-selectionPreviewPane--active": Boolean(preview()),
                                    "diff-selectionPreviewPane--expanded": previewExpanded(),
                                }}
                            >
                                <div class="diff-selectionPreviewPane__value">
                                    <div class="diff-statusBar">
                                        <span>{`${strings.diffs}: ${activeSheet()?.diffCellCount ?? 0}`}</span>
                                        <span>{`${rowFilter()}: ${rowStats()[rowFilter()]}`}</span>
                                    </div>
                                    {preview()
                                        ? getDisplayValue(preview()!.rightValue, strings.emptyValue)
                                        : strings.none}
                                </div>
                                <button
                                    class="diff-selectionPreviewPane__action"
                                    type="button"
                                    title={strings.viewDetails}
                                    onClick={() => setPreviewExpanded((current) => !current)}
                                >
                                    <span
                                        class={`codicon ${previewExpanded() ? "codicon-chevron-up" : "codicon-chevron-down"}`}
                                    />
                                </button>
                            </div>
                        </div>

                        <div class="diff-gridShell">
                            <div class="diff-gridHeaderRow">
                                <div class="diff-paneHeader diff-pane">
                                    <div class="diff-indexCell diff-indexCell--header">#</div>
                                    <div
                                        class="diff-columnsViewport"
                                        ref={(element) => {
                                            leftHeaderViewportElement = element;
                                        }}
                                    >
                                        <div
                                            class="diff-columnsTrack"
                                            style={horizontalTrackStyle()}
                                        >
                                            <For each={visibleColumns()}>
                                                {(column) => (
                                                    <div class="diff-headerCell">
                                                        <span class="diff-headerLabel">
                                                            {getColumnHeaderLabel(column)}
                                                        </span>
                                                    </div>
                                                )}
                                            </For>
                                        </div>
                                    </div>
                                </div>
                                <div class="diff-divider" />
                                <div class="diff-paneHeader diff-pane">
                                    <div class="diff-indexCell diff-indexCell--header">#</div>
                                    <div
                                        class="diff-columnsViewport"
                                        ref={(element) => {
                                            rightHeaderViewportElement = element;
                                        }}
                                    >
                                        <div
                                            class="diff-columnsTrack"
                                            style={horizontalTrackStyle()}
                                        >
                                            <For each={visibleColumns()}>
                                                {(column) => (
                                                    <div class="diff-headerCell">
                                                        <span class="diff-headerLabel">
                                                            {getColumnHeaderLabel(column)}
                                                        </span>
                                                    </div>
                                                )}
                                            </For>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="diff-gridViewport" style={{ "overflow-x": "hidden" }}>
                                <div class="diff-gridViewportInner">
                                    <div class="diff-visibleRows">
                                        <For each={filteredRows()}>
                                            {(row) => (
                                                <div class="diff-pairRow">
                                                    <div class="diff-sideRow diff-pane">
                                                        <div
                                                            class="diff-indexCell"
                                                            classList={{
                                                                [getRowToneClass(row)]: Boolean(
                                                                    getRowToneClass(row)
                                                                ),
                                                                "diff-indexCell--selected":
                                                                    selectedCell()?.rowNumber ===
                                                                    row.rowNumber,
                                                            }}
                                                        >
                                                            <span class="diff-indexLabel">
                                                                {row.leftRowNumber ?? ""}
                                                            </span>
                                                        </div>
                                                        <div class="diff-rowCellsViewport">
                                                            <div
                                                                class="diff-rowCellsTrack"
                                                                style={horizontalTrackStyle()}
                                                            >
                                                                <For each={visibleColumns()}>
                                                                    {(column) => {
                                                                        const cell =
                                                                            row.cells.find(
                                                                                (candidate) =>
                                                                                    candidate.columnNumber ===
                                                                                    column.columnNumber
                                                                            ) ??
                                                                            findCellInSheet(
                                                                                activeSheet(),
                                                                                row.rowNumber,
                                                                                column.columnNumber
                                                                            );
                                                                        return (
                                                                            <Show when={cell}>
                                                                                {(resolvedCell) => {
                                                                                    const currentCell =
                                                                                        resolvedCell();
                                                                                    const isSelected =
                                                                                        selectedCell()
                                                                                            ?.rowNumber ===
                                                                                            row.rowNumber &&
                                                                                        selectedCell()
                                                                                            ?.columnNumber ===
                                                                                            currentCell.columnNumber;
                                                                                    const isEditing =
                                                                                        editingCell()
                                                                                            ?.rowNumber ===
                                                                                            row.rowNumber &&
                                                                                        editingCell()
                                                                                            ?.columnNumber ===
                                                                                            currentCell.columnNumber &&
                                                                                        editingCell()
                                                                                            ?.side ===
                                                                                            "left";
                                                                                    return (
                                                                                        <button
                                                                                            class="diff-cell"
                                                                                            classList={{
                                                                                                [getCellToneClass(
                                                                                                    currentCell
                                                                                                )]:
                                                                                                    Boolean(
                                                                                                        getCellToneClass(
                                                                                                            currentCell
                                                                                                        )
                                                                                                    ),
                                                                                                "diff-cell--active":
                                                                                                    isSelected,
                                                                                                "diff-cell--pending":
                                                                                                    Boolean(
                                                                                                        pendingEdits()[
                                                                                                            getPendingDiffEditKey(
                                                                                                                row.rowNumber,
                                                                                                                currentCell.columnNumber,
                                                                                                                "left"
                                                                                                            )
                                                                                                        ]
                                                                                                    ),
                                                                                                "diff-cell--editing":
                                                                                                    isEditing,
                                                                                            }}
                                                                                            type="button"
                                                                                            onClick={() =>
                                                                                                selectCell(
                                                                                                    row,
                                                                                                    currentCell
                                                                                                )
                                                                                            }
                                                                                            onDblClick={() =>
                                                                                                startEditingCell(
                                                                                                    row,
                                                                                                    currentCell,
                                                                                                    "left"
                                                                                                )
                                                                                            }
                                                                                        >
                                                                                            <Show
                                                                                                when={
                                                                                                    isEditing
                                                                                                }
                                                                                                fallback={
                                                                                                    <span class="diff-cell__text">
                                                                                                        {getDisplayValue(
                                                                                                            getRenderedCellValue(
                                                                                                                row,
                                                                                                                currentCell,
                                                                                                                "left"
                                                                                                            ),
                                                                                                            strings.emptyValue
                                                                                                        )}
                                                                                                    </span>
                                                                                                }
                                                                                            >
                                                                                                <input
                                                                                                    class="diff-cell__input"
                                                                                                    value={
                                                                                                        editingCell()
                                                                                                            ?.value ??
                                                                                                        ""
                                                                                                    }
                                                                                                    onInput={(
                                                                                                        event
                                                                                                    ) =>
                                                                                                        setEditingCell(
                                                                                                            (
                                                                                                                current
                                                                                                            ) =>
                                                                                                                current
                                                                                                                    ? {
                                                                                                                          ...current,
                                                                                                                          value: event
                                                                                                                              .currentTarget
                                                                                                                              .value,
                                                                                                                      }
                                                                                                                    : current
                                                                                                        )
                                                                                                    }
                                                                                                    onBlur={() => {
                                                                                                        const draft =
                                                                                                            editingCell();
                                                                                                        if (
                                                                                                            draft
                                                                                                        ) {
                                                                                                            finalizeEditingCell(
                                                                                                                "commit"
                                                                                                            );
                                                                                                        }
                                                                                                    }}
                                                                                                    onKeyDown={(
                                                                                                        event
                                                                                                    ) => {
                                                                                                        if (
                                                                                                            event.key ===
                                                                                                            "Enter"
                                                                                                        ) {
                                                                                                            event.preventDefault();
                                                                                                            const draft =
                                                                                                                editingCell();
                                                                                                            if (
                                                                                                                draft
                                                                                                            ) {
                                                                                                                finalizeEditingCell(
                                                                                                                    "commit"
                                                                                                                );
                                                                                                            }
                                                                                                        }
                                                                                                        if (
                                                                                                            event.key ===
                                                                                                            "Escape"
                                                                                                        ) {
                                                                                                            event.preventDefault();
                                                                                                            finalizeEditingCell(
                                                                                                                "cancel"
                                                                                                            );
                                                                                                        }
                                                                                                    }}
                                                                                                />
                                                                                            </Show>
                                                                                        </button>
                                                                                    );
                                                                                }}
                                                                            </Show>
                                                                        );
                                                                    }}
                                                                </For>
                                                            </div>
                                                        </div>
                                                    </div>
                                                    <div class="diff-divider" />
                                                    <div class="diff-sideRow diff-pane">
                                                        <div
                                                            class="diff-indexCell"
                                                            classList={{
                                                                [getRowToneClass(row)]: Boolean(
                                                                    getRowToneClass(row)
                                                                ),
                                                                "diff-indexCell--selected":
                                                                    selectedCell()?.rowNumber ===
                                                                    row.rowNumber,
                                                            }}
                                                        >
                                                            <span class="diff-indexLabel">
                                                                {row.rightRowNumber ?? ""}
                                                            </span>
                                                        </div>
                                                        <div class="diff-rowCellsViewport">
                                                            <div
                                                                class="diff-rowCellsTrack"
                                                                style={horizontalTrackStyle()}
                                                            >
                                                                <For each={visibleColumns()}>
                                                                    {(column) => {
                                                                        const cell =
                                                                            row.cells.find(
                                                                                (candidate) =>
                                                                                    candidate.columnNumber ===
                                                                                    column.columnNumber
                                                                            ) ??
                                                                            findCellInSheet(
                                                                                activeSheet(),
                                                                                row.rowNumber,
                                                                                column.columnNumber
                                                                            );
                                                                        return (
                                                                            <Show when={cell}>
                                                                                {(resolvedCell) => {
                                                                                    const currentCell =
                                                                                        resolvedCell();
                                                                                    const isSelected =
                                                                                        selectedCell()
                                                                                            ?.rowNumber ===
                                                                                            row.rowNumber &&
                                                                                        selectedCell()
                                                                                            ?.columnNumber ===
                                                                                            currentCell.columnNumber;
                                                                                    const isEditing =
                                                                                        editingCell()
                                                                                            ?.rowNumber ===
                                                                                            row.rowNumber &&
                                                                                        editingCell()
                                                                                            ?.columnNumber ===
                                                                                            currentCell.columnNumber &&
                                                                                        editingCell()
                                                                                            ?.side ===
                                                                                            "right";
                                                                                    return (
                                                                                        <button
                                                                                            class="diff-cell"
                                                                                            classList={{
                                                                                                [getCellToneClass(
                                                                                                    currentCell
                                                                                                )]:
                                                                                                    Boolean(
                                                                                                        getCellToneClass(
                                                                                                            currentCell
                                                                                                        )
                                                                                                    ),
                                                                                                "diff-cell--active":
                                                                                                    isSelected,
                                                                                                "diff-cell--pending":
                                                                                                    Boolean(
                                                                                                        pendingEdits()[
                                                                                                            getPendingDiffEditKey(
                                                                                                                row.rowNumber,
                                                                                                                currentCell.columnNumber,
                                                                                                                "right"
                                                                                                            )
                                                                                                        ]
                                                                                                    ),
                                                                                                "diff-cell--editing":
                                                                                                    isEditing,
                                                                                            }}
                                                                                            type="button"
                                                                                            onClick={() =>
                                                                                                selectCell(
                                                                                                    row,
                                                                                                    currentCell
                                                                                                )
                                                                                            }
                                                                                            onDblClick={() =>
                                                                                                startEditingCell(
                                                                                                    row,
                                                                                                    currentCell,
                                                                                                    "right"
                                                                                                )
                                                                                            }
                                                                                        >
                                                                                            <Show
                                                                                                when={
                                                                                                    isEditing
                                                                                                }
                                                                                                fallback={
                                                                                                    <span class="diff-cell__text">
                                                                                                        {getDisplayValue(
                                                                                                            getRenderedCellValue(
                                                                                                                row,
                                                                                                                currentCell,
                                                                                                                "right"
                                                                                                            ),
                                                                                                            strings.emptyValue
                                                                                                        )}
                                                                                                    </span>
                                                                                                }
                                                                                            >
                                                                                                <input
                                                                                                    class="diff-cell__input"
                                                                                                    value={
                                                                                                        editingCell()
                                                                                                            ?.value ??
                                                                                                        ""
                                                                                                    }
                                                                                                    onInput={(
                                                                                                        event
                                                                                                    ) =>
                                                                                                        setEditingCell(
                                                                                                            (
                                                                                                                current
                                                                                                            ) =>
                                                                                                                current
                                                                                                                    ? {
                                                                                                                          ...current,
                                                                                                                          value: event
                                                                                                                              .currentTarget
                                                                                                                              .value,
                                                                                                                      }
                                                                                                                    : current
                                                                                                        )
                                                                                                    }
                                                                                                    onBlur={() => {
                                                                                                        const draft =
                                                                                                            editingCell();
                                                                                                        if (
                                                                                                            draft
                                                                                                        ) {
                                                                                                            finalizeEditingCell(
                                                                                                                "commit"
                                                                                                            );
                                                                                                        }
                                                                                                    }}
                                                                                                    onKeyDown={(
                                                                                                        event
                                                                                                    ) => {
                                                                                                        if (
                                                                                                            event.key ===
                                                                                                            "Enter"
                                                                                                        ) {
                                                                                                            event.preventDefault();
                                                                                                            const draft =
                                                                                                                editingCell();
                                                                                                            if (
                                                                                                                draft
                                                                                                            ) {
                                                                                                                finalizeEditingCell(
                                                                                                                    "commit"
                                                                                                                );
                                                                                                            }
                                                                                                        }
                                                                                                        if (
                                                                                                            event.key ===
                                                                                                            "Escape"
                                                                                                        ) {
                                                                                                            event.preventDefault();
                                                                                                            finalizeEditingCell(
                                                                                                                "cancel"
                                                                                                            );
                                                                                                        }
                                                                                                    }}
                                                                                                />
                                                                                            </Show>
                                                                                        </button>
                                                                                    );
                                                                                }}
                                                                            </Show>
                                                                        );
                                                                    }}
                                                                </For>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            )}
                                        </For>
                                    </div>
                                </div>
                            </div>

                            <div class="diff-gridScrollbarRow">
                                <div
                                    class="diff-scrollbarRow diff-pane"
                                    classList={{
                                        "diff-scrollbarRow--inactive": !showHorizontalScrollbar(),
                                    }}
                                >
                                    <div class="diff-scrollbarSpacer" />
                                    <div
                                        class="diff-scrollbar"
                                        classList={{
                                            "diff-scrollbar--inactive": !showHorizontalScrollbar(),
                                        }}
                                        ref={(element) => {
                                            leftScrollbarElement = element;
                                        }}
                                        onScroll={(event) =>
                                            syncHorizontalScroll(event.currentTarget.scrollLeft)
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
                                    class="diff-scrollbarRow diff-pane"
                                    classList={{
                                        "diff-scrollbarRow--inactive": !showHorizontalScrollbar(),
                                    }}
                                >
                                    <div class="diff-scrollbarSpacer" />
                                    <div
                                        class="diff-scrollbar"
                                        classList={{
                                            "diff-scrollbar--inactive": !showHorizontalScrollbar(),
                                        }}
                                        ref={(element) => {
                                            rightScrollbarElement = element;
                                        }}
                                        onScroll={(event) =>
                                            syncHorizontalScroll(event.currentTarget.scrollLeft)
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
                </div>
            </div>
        </Show>
    );
}

const rootElement = document.getElementById("app");
if (rootElement) {
    render(() => <DiffBootstrapApp />, rootElement);
}
