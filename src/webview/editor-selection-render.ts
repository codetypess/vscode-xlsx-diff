export interface LocalSelectionUpdateOptions {
    hasNextCell: boolean;
    hasModel: boolean;
    hasEditingCell: boolean;
    isNextCellVisible: boolean;
    hasExpandedSelection: boolean;
    isSimpleSelection: boolean;
    forceRender?: boolean;
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
