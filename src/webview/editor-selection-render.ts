export interface LocalSelectionUpdateOptions {
    hasNextCell: boolean;
    hasModel: boolean;
    hasEditingCell: boolean;
    isNextCellVisible: boolean;
    hasExpandedSelection: boolean;
    isSimpleSelection: boolean;
    forceRender?: boolean;
}

export interface SelectionPositionLike {
    rowNumber: number;
    columnNumber: number;
}

export function isSelectionFocusCell(
    currentSelection: SelectionPositionLike | null,
    rowNumber: number,
    columnNumber: number
): boolean {
    return Boolean(
        currentSelection &&
            currentSelection.rowNumber === rowNumber &&
            currentSelection.columnNumber === columnNumber
    );
}

export function shouldUseLocalSimpleSelectionUpdate({
    hasNextCell,
    hasModel,
    hasEditingCell,
    isNextCellVisible,
    hasExpandedSelection,
    isSimpleSelection,
    forceRender = false,
}: LocalSelectionUpdateOptions): boolean {
    return Boolean(
        hasNextCell &&
            hasModel &&
            !hasEditingCell &&
            !forceRender &&
            isNextCellVisible &&
            !hasExpandedSelection &&
            isSimpleSelection
    );
}

export function shouldSyncLocalSelectionDomFromModelSelection(
    previousCell: SelectionPositionLike | null,
    previousAnchorCell: SelectionPositionLike | null,
    nextCell: SelectionPositionLike | null
): boolean {
    if (!previousCell || !nextCell) {
        return false;
    }

    const anchor = previousAnchorCell ?? previousCell;
    return (
        anchor.rowNumber === previousCell.rowNumber &&
        anchor.columnNumber === previousCell.columnNumber
    );
}

export function shouldResetInvisibleSelectionAnchor({
    hasSelectionRangeOverride,
    hasExpandedSelection,
    isAnchorVisible,
}: {
    hasSelectionRangeOverride: boolean;
    hasExpandedSelection: boolean;
    isAnchorVisible: boolean;
}): boolean {
    return !hasSelectionRangeOverride && !hasExpandedSelection && !isAnchorVisible;
}
