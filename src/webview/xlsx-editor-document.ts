import * as vscode from "vscode";
import {
    type CellEdit,
    type SheetEdit,
    type SheetViewEdit,
    type WorkbookEditState,
    writeWorkbookEditsToDestination,
} from "../core/fastxlsx/write-cell-value";

function getCellEditKey(edit: CellEdit): string {
    return `${edit.sheetName}:${edit.rowNumber}:${edit.columnNumber}`;
}

function compareCellEdits(left: CellEdit, right: CellEdit): number {
    if (left.sheetName !== right.sheetName) {
        return left.sheetName.localeCompare(right.sheetName);
    }

    if (left.rowNumber !== right.rowNumber) {
        return left.rowNumber - right.rowNumber;
    }

    if (left.columnNumber !== right.columnNumber) {
        return left.columnNumber - right.columnNumber;
    }

    return left.value.localeCompare(right.value);
}

function areCellEditsEqual(left: readonly CellEdit[], right: readonly CellEdit[]): boolean {
    if (left.length !== right.length) {
        return false;
    }

    return left.every((edit, index) => {
        const other = right[index];
        return (
            edit.sheetName === other.sheetName &&
            edit.rowNumber === other.rowNumber &&
            edit.columnNumber === other.columnNumber &&
            edit.value === other.value
        );
    });
}

function areSheetEditsEqual(left: readonly SheetEdit[], right: readonly SheetEdit[]): boolean {
    if (left.length !== right.length) {
        return false;
    }

    return left.every((edit, index) => {
        const other = right[index];
        if (
            edit.type !== other.type ||
            edit.sheetKey !== other.sheetKey ||
            edit.sheetName !== other.sheetName
        ) {
            return false;
        }

        if (edit.type === "addSheet" && other.type === "addSheet") {
            return edit.targetIndex === other.targetIndex;
        }

        if (edit.type === "deleteSheet" && other.type === "deleteSheet") {
            return edit.targetIndex === other.targetIndex;
        }

        if (edit.type === "renameSheet" && other.type === "renameSheet") {
            return edit.nextSheetName === other.nextSheetName;
        }

        return true;
    });
}

function compareViewEdits(left: SheetViewEdit, right: SheetViewEdit): number {
    if (left.sheetKey !== right.sheetKey) {
        return left.sheetKey.localeCompare(right.sheetKey);
    }

    const leftColumnCount = left.freezePane?.columnCount ?? -1;
    const rightColumnCount = right.freezePane?.columnCount ?? -1;
    if (leftColumnCount !== rightColumnCount) {
        return leftColumnCount - rightColumnCount;
    }

    const leftRowCount = left.freezePane?.rowCount ?? -1;
    const rightRowCount = right.freezePane?.rowCount ?? -1;
    return leftRowCount - rightRowCount;
}

function areViewEditsEqual(
    left: readonly SheetViewEdit[],
    right: readonly SheetViewEdit[]
): boolean {
    if (left.length !== right.length) {
        return false;
    }

    return left.every((edit, index) => {
        const other = right[index];
        return (
            edit.sheetKey === other.sheetKey &&
            edit.sheetName === other.sheetName &&
            (edit.freezePane?.columnCount ?? null) === (other.freezePane?.columnCount ?? null) &&
            (edit.freezePane?.rowCount ?? null) === (other.freezePane?.rowCount ?? null)
        );
    });
}

export class XlsxEditorDocument implements vscode.CustomDocument {
    private readonly pendingEdits = new Map<string, CellEdit>();
    private pendingViewEdits: SheetViewEdit[] = [];
    private pendingSheetEdits: SheetEdit[] = [];
    private backupUri: vscode.Uri | null;
    private shouldMarkDirtyFromBackup: boolean;

    public constructor(
        public readonly uri: vscode.Uri,
        backupUri?: vscode.Uri
    ) {
        this.backupUri = backupUri ?? null;
        this.shouldMarkDirtyFromBackup = Boolean(backupUri);
    }

    public dispose(): void {}

    public getReadUri(): vscode.Uri {
        return this.backupUri ?? this.uri;
    }

    public hasPendingEdits(): boolean {
        return (
            this.pendingEdits.size > 0 ||
            this.pendingSheetEdits.length > 0 ||
            this.pendingViewEdits.length > 0 ||
            this.backupUri !== null
        );
    }

    public consumeInitialDirtyState(): boolean {
        if (!this.shouldMarkDirtyFromBackup) {
            return false;
        }

        this.shouldMarkDirtyFromBackup = false;
        return true;
    }

    public getPendingEdits(): CellEdit[] {
        return [...this.pendingEdits.values()].sort(compareCellEdits);
    }

    public getPendingState(): WorkbookEditState {
        return {
            cellEdits: this.getPendingEdits(),
            sheetEdits: [...this.pendingSheetEdits],
            viewEdits: [...this.pendingViewEdits],
        };
    }

    public replacePendingState(state: Readonly<WorkbookEditState>): boolean {
        const normalizedEdits = [...state.cellEdits].sort(compareCellEdits);
        const normalizedSheetEdits = [...state.sheetEdits];
        const normalizedViewEdits = [...(state.viewEdits ?? [])].sort(compareViewEdits);
        const currentState = this.getPendingState();
        if (
            areCellEditsEqual(currentState.cellEdits, normalizedEdits) &&
            areSheetEditsEqual(currentState.sheetEdits, normalizedSheetEdits) &&
            areViewEditsEqual(currentState.viewEdits ?? [], normalizedViewEdits)
        ) {
            return false;
        }

        this.pendingEdits.clear();
        for (const edit of normalizedEdits) {
            this.pendingEdits.set(getCellEditKey(edit), edit);
        }

        this.pendingSheetEdits = normalizedSheetEdits;
        this.pendingViewEdits = normalizedViewEdits;

        return true;
    }

    public markSaved(): void {
        this.pendingEdits.clear();
        this.pendingSheetEdits = [];
        this.pendingViewEdits = [];
        this.backupUri = null;
        this.shouldMarkDirtyFromBackup = false;
    }

    public markReverted(): void {
        this.markSaved();
    }

    public async saveTo(destination: vscode.Uri): Promise<void> {
        await writeWorkbookEditsToDestination(
            this.getReadUri(),
            destination,
            this.getPendingState()
        );
    }
}
