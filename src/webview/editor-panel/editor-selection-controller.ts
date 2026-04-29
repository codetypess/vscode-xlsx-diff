import {
    createSelectionRange,
    hasExpandedSelectionRange,
    type SelectionPositionLike,
    type SelectionRange,
} from "./editor-selection-range";

export interface PendingSelection<T extends SelectionPositionLike = SelectionPositionLike>
    extends SelectionPositionLike {
    reveal: boolean;
}

export type SelectionDragState<T extends SelectionPositionLike = SelectionPositionLike> =
    | {
          kind: "cell";
          anchorCell: T;
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

export interface SelectionControllerState<T extends SelectionPositionLike = SelectionPositionLike> {
    selectedCell: T | null;
    selectionAnchorCell: T | null;
    selectionRangeOverride: SelectionRange | null;
    pendingSelectionAfterRender: PendingSelection<T> | null;
    suppressAutoSelection: boolean;
    selectionDragState: SelectionDragState<T> | null;
}

export function createSelectionControllerState<
    T extends SelectionPositionLike = SelectionPositionLike,
>(
    overrides: Partial<SelectionControllerState<T>> = {}
): SelectionControllerState<T> {
    return {
        selectedCell: null,
        selectionAnchorCell: null,
        selectionRangeOverride: null,
        pendingSelectionAfterRender: null,
        suppressAutoSelection: false,
        selectionDragState: null,
        ...overrides,
    };
}

export function getSelectionRange<T extends SelectionPositionLike>(
    state: Pick<
        SelectionControllerState<T>,
        "selectedCell" | "selectionAnchorCell" | "selectionRangeOverride"
    >
): SelectionRange | null {
    return (
        state.selectionRangeOverride ??
        createSelectionRange(state.selectionAnchorCell ?? state.selectedCell, state.selectedCell)
    );
}

export function getExpandedSelectionRange<T extends SelectionPositionLike>(
    state: Pick<
        SelectionControllerState<T>,
        "selectedCell" | "selectionAnchorCell" | "selectionRangeOverride"
    >
): SelectionRange | null {
    const range = getSelectionRange(state);
    return hasExpandedSelectionRange(range) ? range : null;
}

export function getActiveHighlightCell<T extends SelectionPositionLike>(
    state: Pick<SelectionControllerState<T>, "selectedCell" | "selectionAnchorCell">
): T | null {
    return state.selectionAnchorCell ?? state.selectedCell;
}

export function getSelectionExtendAnchorCell<T extends SelectionPositionLike>(
    state: Pick<SelectionControllerState<T>, "selectedCell" | "selectionAnchorCell">
): T | null {
    return state.selectionAnchorCell ?? state.selectedCell;
}

export function isActiveSelectionCell<T extends SelectionPositionLike>(
    state: Pick<SelectionControllerState<T>, "selectedCell">,
    rowNumber: number,
    columnNumber: number
): boolean {
    return Boolean(
        state.selectedCell &&
            state.selectedCell.rowNumber === rowNumber &&
            state.selectedCell.columnNumber === columnNumber
    );
}

export function isSimpleSelectionState<T extends SelectionPositionLike>(
    state: Pick<
        SelectionControllerState<T>,
        "selectedCell" | "selectionAnchorCell" | "selectionRangeOverride"
    >
): boolean {
    if (!state.selectedCell) {
        return state.selectionAnchorCell === null && state.selectionRangeOverride === null;
    }

    return !hasExpandedSelectionRange(getSelectionRange(state));
}

export function setSelectedCell<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    nextCell: T | null,
    {
        anchorCell,
        selectionRangeOverride,
    }: {
        anchorCell?: T | null;
        selectionRangeOverride?: SelectionRange | null;
    } = {}
): SelectionControllerState<T> {
    const nextAnchorCell = nextCell ? (anchorCell ?? nextCell) : null;
    return {
        ...state,
        selectedCell: nextCell,
        selectionAnchorCell: nextAnchorCell,
        selectionRangeOverride: selectionRangeOverride ?? null,
        suppressAutoSelection: nextCell === null,
    };
}

export function clearSelectedCell<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>
): SelectionControllerState<T> {
    return setSelectedCell(state, null);
}

export function setSelectionAnchorCell<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    selectionAnchorCell: T | null
): SelectionControllerState<T> {
    return {
        ...state,
        selectionAnchorCell,
    };
}

export function syncSelectionAnchorToSelectedCell<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>
): SelectionControllerState<T> {
    return {
        ...state,
        selectionAnchorCell: state.selectedCell,
    };
}

export function setSelectionRangeOverride<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    selectionRangeOverride: SelectionRange | null
): SelectionControllerState<T> {
    return {
        ...state,
        selectionRangeOverride,
    };
}

export function setPendingSelectionAfterRender<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    pendingSelectionAfterRender: PendingSelection<T> | null
): SelectionControllerState<T> {
    return {
        ...state,
        pendingSelectionAfterRender,
    };
}

export function clearPendingSelectionAfterRender<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>
): SelectionControllerState<T> {
    return setPendingSelectionAfterRender(state, null);
}

export function setSuppressAutoSelection<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    suppressAutoSelection: boolean
): SelectionControllerState<T> {
    return {
        ...state,
        suppressAutoSelection,
    };
}

export function startCellSelectionDrag<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    pointerId: number,
    anchorCell: T
): SelectionControllerState<T> {
    return {
        ...state,
        selectionDragState: {
            kind: "cell",
            anchorCell,
            pointerId,
        },
    };
}

export function startRowSelectionDrag<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    pointerId: number,
    anchorRowNumber: number
): SelectionControllerState<T> {
    return {
        ...state,
        selectionDragState: {
            kind: "row",
            anchorRowNumber,
            pointerId,
        },
    };
}

export function startColumnSelectionDrag<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    pointerId: number,
    anchorColumnNumber: number
): SelectionControllerState<T> {
    return {
        ...state,
        selectionDragState: {
            kind: "column",
            anchorColumnNumber,
            pointerId,
        },
    };
}

export function stopSelectionDrag<T extends SelectionPositionLike>(
    state: SelectionControllerState<T>,
    pointerId?: number
): SelectionControllerState<T> {
    if (!state.selectionDragState) {
        return state;
    }

    if (pointerId !== undefined && state.selectionDragState.pointerId !== pointerId) {
        return state;
    }

    return {
        ...state,
        selectionDragState: null,
    };
}
