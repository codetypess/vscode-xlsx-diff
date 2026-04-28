import * as React from "react";
import type { EditorRenderModel, EditorSheetTabView } from "../../core/model/types";
import { getMaxVisibleSheetTabsForWidth, partitionSheetTabs } from "../editor-sheet-tabs";
import type { EditorPanelStrings, SearchOptions } from "./editor-panel-types";
import type { SelectionRange as CellRange } from "./editor-selection-range";
import {
    getToolbarCellEditTargetKey,
    shouldResetToolbarCellValueDraft,
    type ToolbarCellEditTarget,
} from "./editor-toolbar-input";
import { getEditorToolbarSyncSnapshot, subscribeEditorToolbarSync } from "./editor-toolbar-sync";
import {
    classNames,
    type ContextMenuState,
    type PendingSummary,
    type SearchPanelFeedback,
    type SearchPanelMode,
    type SearchPanelPosition,
} from "./editor-panel-ui-shared";

const SHEET_TAB_ITEM_GAP = 1;
const SHEET_TAB_ESTIMATED_WIDTH = 120;
const SHEET_TAB_VISIBLE_MAX_WIDTH = 144;
const SHEET_TAB_OVERFLOW_TRIGGER_WIDTH = 32;

export function PendingMarker({ extraClass }: { extraClass?: string }): React.ReactElement {
    return (
        <span
            className={classNames(["diff-marker", "diff-marker--pending", extraClass])}
            aria-hidden
        />
    );
}

export function CellValue({
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
    className,
    disabled = false,
    isActive = false,
    isLoading = false,
    iconOnly = false,
    iconMirrored = false,
    onClick,
}: {
    actionLabel: string;
    icon: string;
    className?: string;
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
                className,
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

export function SearchPanel({
    strings,
    isOpen,
    mode,
    query,
    replaceValue,
    options,
    feedback,
    scopeSummary,
    hasSelectionScope,
    position,
    hasSearchableGrid,
    canReplace,
    isInteractiveTarget,
    onSyncPosition,
    onBeginDrag,
    onClose,
    onModeChange,
    onQueryChange,
    onReplaceValueChange,
    onToggleOption,
    onSubmitSearch,
    onSubmitReplace,
}: {
    strings: EditorPanelStrings;
    isOpen: boolean;
    mode: SearchPanelMode;
    query: string;
    replaceValue: string;
    options: SearchOptions;
    feedback: SearchPanelFeedback | null;
    scopeSummary: string;
    hasSelectionScope: boolean;
    position: SearchPanelPosition | null;
    hasSearchableGrid: boolean;
    canReplace: boolean;
    isInteractiveTarget(target: EventTarget | null): boolean;
    onSyncPosition(shell: HTMLElement): void;
    onBeginDrag(pointerId: number, clientX: number, clientY: number): void;
    onClose(): void;
    onModeChange(mode: SearchPanelMode): void;
    onQueryChange(value: string): void;
    onReplaceValueChange(value: string): void;
    onToggleOption(option: keyof SearchOptions): void;
    onSubmitSearch(direction: "next" | "prev"): void;
    onSubmitReplace(mode: "single" | "all"): void;
}): React.ReactElement | null {
    const shellRef = React.useRef<HTMLElement | null>(null);

    React.useLayoutEffect(() => {
        if (!isOpen || !shellRef.current) {
            return;
        }

        onSyncPosition(shellRef.current);
    }, [isOpen, onSyncPosition, position]);

    if (!isOpen) {
        return null;
    }

    const isReplaceMode = mode === "replace";
    const isQueryEmpty = query.trim().length === 0;
    const feedbackToneClass =
        feedback?.status === "matched" || feedback?.status === "replaced"
            ? "search-strip__feedback--success"
            : feedback?.status === "invalid-pattern"
              ? "search-strip__feedback--error"
              : feedback?.status === "no-match" || feedback?.status === "no-change"
                ? "search-strip__feedback--warn"
                : undefined;
    const panelStyle: React.CSSProperties | undefined = position
        ? {
              left: `${position.left}px`,
              top: `${position.top}px`,
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
                if (event.button !== 0 || isInteractiveTarget(event.target)) {
                    return;
                }

                event.preventDefault();
                onBeginDrag(event.pointerId, event.clientX, event.clientY);
            }}
        >
            <div className="search-strip" data-role="search-panel" role="search">
                <div className="search-strip__header">
                    <div className="search-strip__tabs" role="tablist" aria-label={strings.search}>
                        <button
                            aria-selected={!isReplaceMode}
                            className={classNames([
                                "search-strip__tab",
                                !isReplaceMode && "is-active",
                            ])}
                            role="tab"
                            tabIndex={!isReplaceMode ? 0 : -1}
                            type="button"
                            onClick={() => onModeChange("find")}
                        >
                            <span
                                className="codicon codicon-search search-strip__tab-icon"
                                aria-hidden
                            />
                            <span>{strings.searchFind}</span>
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
                            onClick={() => onModeChange("replace")}
                        >
                            <span
                                className="codicon codicon-replace search-strip__tab-icon"
                                aria-hidden
                            />
                            <span>{strings.searchReplace}</span>
                        </button>
                    </div>
                    <div className="search-strip__header-tools">
                        <ToolbarButton
                            actionLabel={strings.searchClose}
                            className="search-strip__close-button"
                            icon="codicon-close"
                            iconOnly={true}
                            onClick={onClose}
                        />
                    </div>
                </div>
                <div className="search-strip__row search-strip__row--primary">
                    <div
                        className={classNames([
                            "search-strip__input-wrap",
                            feedback?.status === "invalid-pattern" && "is-invalid",
                        ])}
                    >
                        <span
                            className="codicon codicon-search search-strip__input-icon"
                            aria-hidden
                        />
                        <input
                            aria-label={strings.search}
                            className="search-strip__input"
                            data-role="search-input"
                            placeholder={strings.searchPlaceholder}
                            type="text"
                            value={query}
                            onChange={(event) => {
                                onQueryChange(event.currentTarget.value);
                            }}
                            onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                    event.preventDefault();
                                    onSubmitSearch(event.shiftKey ? "prev" : "next");
                                    return;
                                }

                                if (event.key === "Escape") {
                                    event.preventDefault();
                                    event.stopPropagation();
                                    onClose();
                                }
                            }}
                        />
                        <div
                            className="search-strip__input-tools"
                            role="group"
                            aria-label={strings.search}
                        >
                            <button
                                aria-label={strings.searchRegex}
                                aria-pressed={options.isRegexp}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    options.isRegexp && "is-active",
                                ])}
                                title={strings.searchRegex}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => onToggleOption("isRegexp")}
                            >
                                <span className="codicon codicon-regex" aria-hidden />
                            </button>
                            <button
                                aria-label={strings.searchMatchCase}
                                aria-pressed={options.matchCase}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    options.matchCase && "is-active",
                                ])}
                                title={strings.searchMatchCase}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => onToggleOption("matchCase")}
                            >
                                <span className="codicon codicon-case-sensitive" aria-hidden />
                            </button>
                            <button
                                aria-label={strings.searchWholeWord}
                                aria-pressed={options.wholeWord}
                                className={classNames([
                                    "search-strip__icon-toggle",
                                    options.wholeWord && "is-active",
                                ])}
                                title={strings.searchWholeWord}
                                type="button"
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => onToggleOption("wholeWord")}
                            >
                                <span className="codicon codicon-whole-word" aria-hidden />
                            </button>
                        </div>
                    </div>
                    <div className="search-strip__actions">
                        <ToolbarButton
                            actionLabel={strings.findPrev}
                            disabled={isQueryEmpty || !hasSearchableGrid}
                            icon="codicon-arrow-up"
                            iconOnly={true}
                            onClick={() => onSubmitSearch("prev")}
                        />
                        <ToolbarButton
                            actionLabel={strings.findNext}
                            disabled={isQueryEmpty || !hasSearchableGrid}
                            icon="codicon-arrow-down"
                            iconOnly={true}
                            onClick={() => onSubmitSearch("next")}
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
                                aria-label={strings.searchReplace}
                                className="search-strip__input"
                                data-role="replace-input"
                                placeholder={strings.replacePlaceholder}
                                type="text"
                                value={replaceValue}
                                onChange={(event) => {
                                    onReplaceValueChange(event.currentTarget.value);
                                }}
                                onKeyDown={(event) => {
                                    if (event.key === "Enter") {
                                        event.preventDefault();
                                        onSubmitReplace(
                                            event.ctrlKey || event.metaKey ? "all" : "single"
                                        );
                                        return;
                                    }

                                    if (event.key === "Escape") {
                                        event.preventDefault();
                                        event.stopPropagation();
                                        onClose();
                                    }
                                }}
                            />
                        </div>
                        <div className="search-strip__replace-actions">
                            <ToolbarButton
                                actionLabel={strings.searchReplace}
                                disabled={!canReplace}
                                icon="codicon-replace"
                                iconOnly={true}
                                onClick={() => onSubmitReplace("single")}
                            />
                            <ToolbarButton
                                actionLabel={strings.replaceAll}
                                disabled={!canReplace}
                                icon="codicon-replace-all"
                                iconOnly={true}
                                onClick={() => onSubmitReplace("all")}
                            />
                        </div>
                    </div>
                ) : null}
                <div className="search-strip__row search-strip__row--meta">
                    <div className="search-strip__scope-summary">
                        <span className="search-strip__scope-summary-label">
                            {strings.searchScopeLabel}
                        </span>
                        <span
                            className={classNames([
                                "search-strip__scope-summary-value",
                                hasSelectionScope && "is-selection",
                            ])}
                        >
                            {scopeSummary}
                        </span>
                    </div>
                </div>
                {feedback?.message ? (
                    <div className={classNames(["search-strip__feedback", feedbackToneClass])}>
                        {feedback.message}
                    </div>
                ) : null}
            </div>
        </section>
    );
}

export function EditorToolbar({
    strings,
    currentModel,
    isSearchPanelOpen,
    isSaving,
    hasPendingEdits,
    canUndo,
    canRedo,
    viewLocked,
    getPositionInputValue,
    getCellValueInputValue,
    getCellValueInputPlaceholder,
    canEditSelectedCellValue,
    getActiveCellEditTarget,
    onOpenSearch,
    onUndo,
    onRedo,
    onReload,
    onToggleViewLock,
    onSave,
    onSubmitGoto,
    onCommitCellValue,
    onFinishGridEdit,
}: {
    strings: EditorPanelStrings;
    currentModel: EditorRenderModel;
    isSearchPanelOpen: boolean;
    isSaving: boolean;
    hasPendingEdits: boolean;
    canUndo: boolean;
    canRedo: boolean;
    viewLocked: boolean;
    getPositionInputValue(): string;
    getCellValueInputValue(): string;
    getCellValueInputPlaceholder(): string;
    canEditSelectedCellValue(): boolean;
    getActiveCellEditTarget(): ToolbarCellEditTarget | null;
    onOpenSearch(mode: SearchPanelMode): void;
    onUndo(): void;
    onRedo(): void;
    onReload(): void;
    onToggleViewLock(): void;
    onSave(): void;
    onSubmitGoto(reference: string): void;
    onCommitCellValue(target: ToolbarCellEditTarget, value: string): void;
    onFinishGridEdit(): void;
}): React.ReactElement {
    React.useSyncExternalStore(
        subscribeEditorToolbarSync,
        getEditorToolbarSyncSnapshot,
        getEditorToolbarSyncSnapshot
    );

    const viewLockActionLabel = viewLocked ? strings.unlockView : strings.lockView;
    const activeCellAddress = getPositionInputValue();
    const selectedCellValue = getCellValueInputValue();
    const cellValuePlaceholder = getCellValueInputPlaceholder();
    const selectedCellEditable = currentModel.canEdit && canEditSelectedCellValue();
    const activeCellEditTarget = getActiveCellEditTarget();
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
        activeCellEditTarget,
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

        onSubmitGoto(nextReference);
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

        onCommitCellValue(target, cellValueInputValue);
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
                        placeholder={strings.gotoPlaceholder}
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
                            onFinishGridEdit();

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
                                aria-label={strings.cancelInput}
                                title={strings.cancelInput}
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={resetCellValueInput}
                            >
                                <span
                                    className="codicon codicon-close toolbar__toggle-icon"
                                    aria-hidden
                                />
                            </button>
                            <button
                                type="button"
                                className="toolbar__toggle is-active"
                                aria-label={strings.confirmInput}
                                title={strings.confirmInput}
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={commitCellValueInput}
                            >
                                <span
                                    className="codicon codicon-check toolbar__toggle-icon"
                                    aria-hidden
                                />
                            </button>
                        </span>
                    ) : null}
                </label>
            </div>
            <div className="toolbar__group">
                <ToolbarButton
                    actionLabel={strings.search}
                    icon="codicon-search"
                    iconOnly={true}
                    isActive={isSearchPanelOpen}
                    onClick={() => onOpenSearch("find")}
                />
                <ToolbarButton
                    actionLabel={strings.undo}
                    disabled={!currentModel.canEdit || !canUndo || isSaving}
                    icon="codicon-redo"
                    iconMirrored={true}
                    iconOnly={true}
                    onClick={onUndo}
                />
                <ToolbarButton
                    actionLabel={strings.redo}
                    disabled={!currentModel.canEdit || !canRedo || isSaving}
                    icon="codicon-redo"
                    iconOnly={true}
                    onClick={onRedo}
                />
                <ToolbarButton
                    actionLabel={strings.reload}
                    icon="codicon-refresh"
                    iconOnly={true}
                    onClick={onReload}
                />
                <ToolbarButton
                    actionLabel={viewLockActionLabel}
                    disabled={!currentModel.canEdit || isSaving}
                    icon={viewLocked ? "codicon-lock" : "codicon-unlock"}
                    iconOnly={true}
                    isActive={viewLocked}
                    onClick={onToggleViewLock}
                />
                <ToolbarButton
                    actionLabel={strings.save}
                    disabled={!currentModel.canEdit || !hasPendingEdits || isSaving}
                    icon="codicon-save"
                    iconOnly={true}
                    isActive={hasPendingEdits}
                    isLoading={isSaving}
                    onClick={onSave}
                />
            </div>
        </header>
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

export function Tabs({
    strings,
    currentModel,
    pendingSummary,
    onSetSheet,
    onOpenTabContextMenu,
    onCloseContextMenu,
}: {
    strings: EditorPanelStrings;
    currentModel: EditorRenderModel;
    pendingSummary: PendingSummary;
    onSetSheet(sheetKey: string): void;
    onOpenTabContextMenu(sheetKey: string, x: number, y: number): void;
    onCloseContextMenu(): void;
}): React.ReactElement {
    const viewportRef = React.useRef<HTMLDivElement | null>(null);
    const overflowRef = React.useRef<HTMLDivElement | null>(null);
    const measureRef = React.useRef<HTMLDivElement | null>(null);
    const [isOverflowOpen, setIsOverflowOpen] = React.useState(false);
    const [measuredTabWidths, setMeasuredTabWidths] = React.useState<Record<string, number>>({});
    const viewportWidth = useObservedElementWidth(viewportRef);
    const pendingSheetKeySignature = Array.from(pendingSummary.sheetKeys).sort().join("\0");
    const maxVisibleTabs = getMaxVisibleSheetTabs(
        currentModel.sheets,
        viewportWidth,
        measuredTabWidths
    );
    const tabLayout = partitionSheetTabs(currentModel.sheets, maxVisibleTabs);

    React.useLayoutEffect(() => {
        const measureRoot = measureRef.current;
        if (!measureRoot) {
            return;
        }

        const nextMeasuredTabWidths: Record<string, number> = {};
        for (const element of measureRoot.querySelectorAll<HTMLElement>(
            '[data-role="sheet-tab-measure"]'
        )) {
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
    }, [currentModel.sheets, currentModel.activeSheet.key, pendingSheetKeySignature]);

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
        onCloseContextMenu();
        setIsOverflowOpen(false);
        onSetSheet(sheetKey);
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
                onOpenTabContextMenu(currentModel.activeSheet.key, event.clientX, event.clientY);
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
                                        onOpenTabContextMenu(
                                            sheet.key,
                                            event.clientX,
                                            event.clientY
                                        );
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
                                aria-label={strings.moreSheets}
                                aria-expanded={isOverflowOpen}
                                aria-haspopup="menu"
                                className={classNames([
                                    "tab",
                                    "tab--overflowTrigger",
                                    isOverflowOpen && "is-active",
                                ])}
                                title={strings.moreSheets}
                                type="button"
                                onClick={() => {
                                    onCloseContextMenu();
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
                                                    onOpenTabContextMenu(
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

export function TabContextMenu({
    strings,
    currentModel,
    contextMenu,
    onRequestAddSheet,
    onRequestDeleteSheet,
    onRequestRenameSheet,
    onRequestInsertRow,
    onRequestDeleteRow,
    onRequestPromptRowHeight,
    onRequestInsertColumn,
    onRequestDeleteColumn,
    onRequestPromptColumnWidth,
}: {
    strings: EditorPanelStrings;
    currentModel: EditorRenderModel;
    contextMenu: ContextMenuState | null;
    onRequestAddSheet(): void;
    onRequestDeleteSheet(sheetKey: string): void;
    onRequestRenameSheet(sheetKey: string): void;
    onRequestInsertRow(rowNumber: number): void;
    onRequestDeleteRow(rowNumber: number): void;
    onRequestPromptRowHeight(rowNumber: number): void;
    onRequestInsertColumn(columnNumber: number): void;
    onRequestDeleteColumn(columnNumber: number): void;
    onRequestPromptColumnWidth(columnNumber: number): void;
}): React.ReactElement | null {
    if (!contextMenu || !currentModel.canEdit) {
        return null;
    }

    const estimatedMenuHeight =
        contextMenu.kind === "column" || contextMenu.kind === "row" ? 168 : 132;
    const menuStyle: React.CSSProperties = {
        left: Math.max(8, Math.min(contextMenu.x, window.innerWidth - 188)),
        top: Math.max(8, Math.min(contextMenu.y, window.innerHeight - estimatedMenuHeight)),
    };

    if (contextMenu.kind === "tab") {
        const disableDelete = currentModel.sheets.length <= 1;
        return (
            <div className="context-menu" data-role="context-menu" style={menuStyle}>
                <button className="context-menu__item" type="button" onClick={onRequestAddSheet}>
                    <span className="codicon codicon-add context-menu__icon" aria-hidden />
                    <span>{strings.addSheet}</span>
                </button>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => onRequestRenameSheet(contextMenu.sheetKey)}
                >
                    <span className="codicon codicon-edit context-menu__icon" aria-hidden />
                    <span>{strings.renameSheet}</span>
                </button>
                <button
                    className="context-menu__item context-menu__item--danger"
                    disabled={disableDelete}
                    type="button"
                    onClick={() => onRequestDeleteSheet(contextMenu.sheetKey)}
                >
                    <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                    <span>{strings.deleteSheet}</span>
                </button>
            </div>
        );
    }

    if (contextMenu.kind === "row") {
        return (
            <div className="context-menu" data-role="context-menu" style={menuStyle}>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => onRequestPromptRowHeight(contextMenu.rowNumber)}
                >
                    <span className="codicon codicon-symbol-number context-menu__icon" aria-hidden />
                    <span>{strings.setRowHeight}</span>
                </button>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => onRequestInsertRow(contextMenu.rowNumber)}
                >
                    <span className="codicon codicon-add context-menu__icon" aria-hidden />
                    <span>{strings.insertRowAbove}</span>
                </button>
                <button
                    className="context-menu__item"
                    type="button"
                    onClick={() => onRequestInsertRow(contextMenu.rowNumber + 1)}
                >
                    <span className="codicon codicon-add context-menu__icon" aria-hidden />
                    <span>{strings.insertRowBelow}</span>
                </button>
                <button
                    className="context-menu__item context-menu__item--danger"
                    disabled={currentModel.activeSheet.rowCount <= 1}
                    type="button"
                    onClick={() => onRequestDeleteRow(contextMenu.rowNumber)}
                >
                    <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                    <span>{strings.deleteRow}</span>
                </button>
            </div>
        );
    }

    return (
        <div className="context-menu" data-role="context-menu" style={menuStyle}>
            <button
                className="context-menu__item"
                type="button"
                onClick={() => onRequestPromptColumnWidth(contextMenu.columnNumber)}
            >
                <span className="codicon codicon-symbol-number context-menu__icon" aria-hidden />
                <span>{strings.setColumnWidth}</span>
            </button>
            <button
                className="context-menu__item"
                type="button"
                onClick={() => onRequestInsertColumn(contextMenu.columnNumber)}
            >
                <span className="codicon codicon-add context-menu__icon" aria-hidden />
                <span>{strings.insertColumnLeft}</span>
            </button>
            <button
                className="context-menu__item"
                type="button"
                onClick={() => onRequestInsertColumn(contextMenu.columnNumber + 1)}
            >
                <span className="codicon codicon-add context-menu__icon" aria-hidden />
                <span>{strings.insertColumnRight}</span>
            </button>
            <button
                className="context-menu__item context-menu__item--danger"
                disabled={currentModel.activeSheet.columnCount <= 1}
                type="button"
                onClick={() => onRequestDeleteColumn(contextMenu.columnNumber)}
            >
                <span className="codicon codicon-trash context-menu__icon" aria-hidden />
                <span>{strings.deleteColumn}</span>
            </button>
        </div>
    );
}

export function Shell({
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
