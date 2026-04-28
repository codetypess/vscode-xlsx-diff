import * as vscode from "vscode";
import {
    type CellEdit,
    type SheetEdit,
    type SheetViewEdit,
    type WorkbookEditState,
} from "../../core/fastxlsx/write-cell-value";
import { areCellAlignmentMapsEquivalent, cloneCellAlignmentMap } from "../../core/model/alignment";

async function writePendingWorkbookEditsToDestination(
    sourceUri: vscode.Uri,
    destinationUri: vscode.Uri,
    edits: WorkbookEditState
): Promise<void> {
    const { writeWorkbookEditsToDestination } = await import(
        "../../core/fastxlsx/write-cell-value"
    );
    await writeWorkbookEditsToDestination(sourceUri, destinationUri, edits);
}

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

        if (
            (edit.type === "insertRow" && other.type === "insertRow") ||
            (edit.type === "deleteRow" && other.type === "deleteRow")
        ) {
            return edit.rowNumber === other.rowNumber && edit.count === other.count;
        }

        if (
            (edit.type === "insertColumn" && other.type === "insertColumn") ||
            (edit.type === "deleteColumn" && other.type === "deleteColumn")
        ) {
            return edit.columnNumber === other.columnNumber && edit.count === other.count;
        }

        return true;
    });
}

function compareViewEdits(left: SheetViewEdit, right: SheetViewEdit): number {
    if (left.sheetKey !== right.sheetKey) {
        return left.sheetKey.localeCompare(right.sheetKey);
    }

    if (left.sheetName !== right.sheetName) {
        return left.sheetName.localeCompare(right.sheetName);
    }

    const leftColumnCount = left.freezePane?.columnCount ?? -1;
    const rightColumnCount = right.freezePane?.columnCount ?? -1;
    if (leftColumnCount !== rightColumnCount) {
        return leftColumnCount - rightColumnCount;
    }

    const leftRowCount = left.freezePane?.rowCount ?? -1;
    const rightRowCount = right.freezePane?.rowCount ?? -1;
    if (leftRowCount !== rightRowCount) {
        return leftRowCount - rightRowCount;
    }

    const maxColumnWidthCount = Math.max(
        left.columnWidths?.length ?? 0,
        right.columnWidths?.length ?? 0
    );
    for (let index = 0; index < maxColumnWidthCount; index += 1) {
        const leftWidth = left.columnWidths?.[index] ?? null;
        const rightWidth = right.columnWidths?.[index] ?? null;
        if (leftWidth === rightWidth) {
            continue;
        }

        return (leftWidth ?? -1) - (rightWidth ?? -1);
    }

    const leftRowHeightKeys = Object.keys(left.rowHeights ?? {});
    const rightRowHeightKeys = Object.keys(right.rowHeights ?? {});
    if (leftRowHeightKeys.length !== rightRowHeightKeys.length) {
        return leftRowHeightKeys.length - rightRowHeightKeys.length;
    }

    for (const rowNumber of leftRowHeightKeys.sort((a, b) => Number(a) - Number(b))) {
        const leftHeight = left.rowHeights?.[rowNumber] ?? null;
        const rightHeight = right.rowHeights?.[rowNumber] ?? null;
        if (leftHeight !== rightHeight) {
            return (leftHeight ?? -1) - (rightHeight ?? -1);
        }
    }

    const leftCellAlignmentKeys = Object.keys(left.cellAlignments ?? {});
    const rightCellAlignmentKeys = Object.keys(right.cellAlignments ?? {});
    if (leftCellAlignmentKeys.length !== rightCellAlignmentKeys.length) {
        return leftCellAlignmentKeys.length - rightCellAlignmentKeys.length;
    }

    const leftRowAlignmentKeys = Object.keys(left.rowAlignments ?? {});
    const rightRowAlignmentKeys = Object.keys(right.rowAlignments ?? {});
    if (leftRowAlignmentKeys.length !== rightRowAlignmentKeys.length) {
        return leftRowAlignmentKeys.length - rightRowAlignmentKeys.length;
    }

    const leftColumnAlignmentKeys = Object.keys(left.columnAlignments ?? {});
    const rightColumnAlignmentKeys = Object.keys(right.columnAlignments ?? {});
    if (leftColumnAlignmentKeys.length !== rightColumnAlignmentKeys.length) {
        return leftColumnAlignmentKeys.length - rightColumnAlignmentKeys.length;
    }

    return 0;
}

function cloneViewEdit(edit: SheetViewEdit): SheetViewEdit {
    return {
        ...edit,
        freezePane: edit.freezePane ? { ...edit.freezePane } : null,
        ...(edit.columnWidths
            ? {
                  columnWidths: edit.columnWidths.map((columnWidth) => columnWidth ?? null),
              }
            : {}),
        ...(edit.rowHeights
            ? {
                  rowHeights: { ...edit.rowHeights },
              }
            : {}),
        ...(edit.cellAlignments
            ? {
                  cellAlignments: cloneCellAlignmentMap(edit.cellAlignments),
              }
            : {}),
        ...(edit.rowAlignments
            ? {
                  rowAlignments: cloneCellAlignmentMap(edit.rowAlignments),
              }
            : {}),
        ...(edit.columnAlignments
            ? {
                  columnAlignments: cloneCellAlignmentMap(edit.columnAlignments),
              }
            : {}),
    };
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
            (edit.freezePane?.rowCount ?? null) === (other.freezePane?.rowCount ?? null) &&
            (edit.columnWidths?.length ?? 0) === (other.columnWidths?.length ?? 0) &&
            (edit.columnWidths ?? []).every(
                (columnWidth, columnIndex) =>
                    columnWidth === (other.columnWidths?.[columnIndex] ?? null)
            ) &&
            Object.keys(edit.rowHeights ?? {}).length ===
                Object.keys(other.rowHeights ?? {}).length &&
            Object.entries(edit.rowHeights ?? {}).every(
                ([rowNumber, rowHeight]) => rowHeight === (other.rowHeights?.[rowNumber] ?? null)
            ) &&
            areCellAlignmentMapsEquivalent(edit.cellAlignments, other.cellAlignments) &&
            areCellAlignmentMapsEquivalent(edit.rowAlignments, other.rowAlignments) &&
            areCellAlignmentMapsEquivalent(edit.columnAlignments, other.columnAlignments)
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
        options?: {
            backupUri?: vscode.Uri;
            backupState?: Readonly<WorkbookEditState>;
        }
    ) {
        this.backupUri = options?.backupUri ?? null;
        this.shouldMarkDirtyFromBackup = Boolean(options?.backupUri || options?.backupState);
        if (options?.backupState) {
            this.replacePendingState(options.backupState);
        }
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
            viewEdits: this.pendingViewEdits.map(cloneViewEdit),
        };
    }

    public replacePendingState(state: Readonly<WorkbookEditState>): boolean {
        const normalizedEdits = [...state.cellEdits].sort(compareCellEdits);
        const normalizedSheetEdits = [...state.sheetEdits];
        const normalizedViewEdits = (state.viewEdits ?? [])
            .map(cloneViewEdit)
            .sort(compareViewEdits);
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
        await writePendingWorkbookEditsToDestination(
            this.getReadUri(),
            destination,
            this.getPendingState()
        );
    }
}
