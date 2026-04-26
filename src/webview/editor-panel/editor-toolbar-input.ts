export interface ToolbarCellEditTarget {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
}

export function getToolbarCellEditTargetKey(target: ToolbarCellEditTarget | null): string | null {
    if (!target) {
        return null;
    }

    return `${target.sheetKey}:${target.rowNumber}:${target.columnNumber}`;
}

export function shouldResetToolbarCellValueDraft(
    editTarget: ToolbarCellEditTarget | null,
    activeTarget: ToolbarCellEditTarget | null,
    isEditable: boolean
): boolean {
    if (!editTarget) {
        return false;
    }

    if (!activeTarget || !isEditable) {
        return true;
    }

    return getToolbarCellEditTargetKey(editTarget) !== getToolbarCellEditTargetKey(activeTarget);
}
