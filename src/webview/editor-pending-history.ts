export interface PendingHistoryEdit {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    value: string;
}

export interface PendingHistoryChange {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    modelValue: string;
    beforeValue: string;
    afterValue: string;
}

export interface PendingHistoryEntry {
    changes: PendingHistoryChange[];
}

function getPendingHistoryKey(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number
): string {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
}

export function rebasePendingHistory(
    undoStack: readonly PendingHistoryEntry[],
    redoStack: readonly PendingHistoryEntry[],
    pendingEdits: readonly PendingHistoryEdit[]
): {
    undoStack: PendingHistoryEntry[];
    redoStack: PendingHistoryEntry[];
} {
    const savedModelValues = new Map(
        pendingEdits.map((edit) => [
            getPendingHistoryKey(edit.sheetKey, edit.rowNumber, edit.columnNumber),
            edit.value,
        ])
    );

    const getSavedModelValue = (change: PendingHistoryChange): string => {
        const key = getPendingHistoryKey(change.sheetKey, change.rowNumber, change.columnNumber);
        const savedModelValue = savedModelValues.get(key);
        if (savedModelValue !== undefined) {
            return savedModelValue;
        }

        savedModelValues.set(key, change.modelValue);
        return change.modelValue;
    };

    const rebaseEntry = (entry: PendingHistoryEntry): PendingHistoryEntry => ({
        changes: entry.changes.map((change) => ({
            ...change,
            modelValue: getSavedModelValue(change),
        })),
    });

    return {
        undoStack: undoStack.map(rebaseEntry),
        redoStack: redoStack.map(rebaseEntry),
    };
}
