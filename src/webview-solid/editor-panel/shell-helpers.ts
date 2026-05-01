import type {
    EditorPendingEdit,
    EditorSearchDirection,
    SearchOptions,
} from "../../webview/editor-panel/editor-panel-types";
import type {
    CellSnapshot,
    SheetAutoFilterSnapshot,
    SheetRangeSnapshot,
} from "../../core/model/types";
import { createCellKey } from "../../core/model/cells";
import {
    createEditorSheetFilterSnapshot,
    createEditorSheetFilterStateFromSnapshot,
    resolveEditorFilterRangeFromActiveCell,
    toggleEditorSheetFilterState,
} from "../../webview/editor-panel/editor-panel-filter";
import type { EditorWebviewOutgoingMessage } from "../shared/session-protocol";
import type { SelectionRange } from "../../webview/editor-panel/editor-selection-range";

export interface EditorShellWorkbookState {
    canEdit: boolean;
    hasPendingEdits: boolean;
    canUndoStructuralEdits: boolean;
    canRedoStructuralEdits: boolean;
}

export interface EditorShellCapabilities {
    isReadOnly: boolean;
    canRequestSave: boolean;
    canUndo: boolean;
    canRedo: boolean;
}

export interface EditorSheetContextMenuState {
    canAddSheet: boolean;
    canRenameSheet: boolean;
    canDeleteSheet: boolean;
}

export interface EditorShellActiveSheetState {
    key: string;
    rowCount: number;
    columnCount: number;
    cells?: Record<string, CellSnapshot>;
    autoFilter?: SheetAutoFilterSnapshot | null;
}

export interface EditorShellSelectionState {
    rowNumber: number;
    columnNumber: number;
}

export interface EditorFilterShellState {
    hasActiveFilter: boolean;
    currentRange: SheetRangeSnapshot | null;
    candidateRange: SelectionRange | null;
    effectiveRange: SheetRangeSnapshot | SelectionRange | null;
    nextFilterState: SheetAutoFilterSnapshot | null;
    canToggle: boolean;
}

export function getEditorShellCapabilities(
    workbook: EditorShellWorkbookState
): EditorShellCapabilities {
    return {
        isReadOnly: !workbook.canEdit,
        canRequestSave: workbook.canEdit && workbook.hasPendingEdits,
        canUndo: workbook.canEdit && workbook.canUndoStructuralEdits,
        canRedo: workbook.canEdit && workbook.canRedoStructuralEdits,
    };
}

export function createEditorSearchMessage(
    query: string,
    direction: EditorSearchDirection,
    options: SearchOptions
): EditorWebviewOutgoingMessage | null {
    const normalizedQuery = query.trim();
    if (normalizedQuery.length === 0) {
        return null;
    }

    return {
        type: "search",
        query: normalizedQuery,
        direction,
        options,
        scope: "sheet",
    };
}

export function createEditorGotoMessage(reference: string): EditorWebviewOutgoingMessage | null {
    const normalizedReference = reference.trim();
    if (normalizedReference.length === 0) {
        return null;
    }

    return {
        type: "gotoCell",
        reference: normalizedReference,
    };
}

export function createEditorSetSheetMessage(sheetKey: string): EditorWebviewOutgoingMessage {
    return {
        type: "setSheet",
        sheetKey,
    };
}

export function getEditorSheetContextMenuState({
    canEdit,
    sheetCount,
}: {
    canEdit: boolean;
    sheetCount: number;
}): EditorSheetContextMenuState {
    return {
        canAddSheet: canEdit,
        canRenameSheet: canEdit,
        canDeleteSheet: canEdit && sheetCount > 1,
    };
}

export function createEditorAddSheetMessage(): EditorWebviewOutgoingMessage {
    return {
        type: "addSheet",
    };
}

export function createEditorRenameSheetMessage(sheetKey: string): EditorWebviewOutgoingMessage {
    return {
        type: "renameSheet",
        sheetKey,
    };
}

export function createEditorDeleteSheetMessage(sheetKey: string): EditorWebviewOutgoingMessage {
    return {
        type: "deleteSheet",
        sheetKey,
    };
}

function getPendingEditValue(
    sheetKey: string,
    rowNumber: number,
    columnNumber: number,
    pendingEdits: readonly EditorPendingEdit[]
): string | undefined {
    return pendingEdits.find(
        (edit) =>
            edit.sheetKey === sheetKey &&
            edit.rowNumber === rowNumber &&
            edit.columnNumber === columnNumber
    )?.value;
}

export function getEditorFilterShellState({
    activeSheet,
    selection,
    pendingEdits,
}: {
    activeSheet: EditorShellActiveSheetState | null;
    selection: EditorShellSelectionState | null;
    pendingEdits: readonly EditorPendingEdit[];
}): EditorFilterShellState {
    const currentFilterState = createEditorSheetFilterStateFromSnapshot(
        activeSheet?.autoFilter ?? null
    );

    let candidateRange: SelectionRange | null = null;
    if (activeSheet?.cells && selection) {
        candidateRange = resolveEditorFilterRangeFromActiveCell(
            {
                rowCount: activeSheet.rowCount,
                columnCount: activeSheet.columnCount,
                getCellValue: (rowNumber, columnNumber) =>
                    getPendingEditValue(activeSheet.key, rowNumber, columnNumber, pendingEdits) ??
                    activeSheet.cells?.[createCellKey(rowNumber, columnNumber)]?.displayValue ??
                    "",
            },
            selection
        );
    }

    const nextFilterState = activeSheet
        ? toggleEditorSheetFilterState(
              {
                  rowCount: activeSheet.rowCount,
                  columnCount: activeSheet.columnCount,
              },
              currentFilterState,
              candidateRange
          )
        : null;
    const hasActiveFilter = Boolean(currentFilterState);

    return {
        hasActiveFilter,
        currentRange: currentFilterState?.range ?? null,
        candidateRange,
        effectiveRange: currentFilterState?.range ?? candidateRange,
        nextFilterState: createEditorSheetFilterSnapshot(nextFilterState),
        canToggle: hasActiveFilter || Boolean(nextFilterState),
    };
}

export function createEditorFilterToggleMessage({
    activeSheet,
    selection,
    pendingEdits,
}: {
    activeSheet: EditorShellActiveSheetState | null;
    selection: EditorShellSelectionState | null;
    pendingEdits: readonly EditorPendingEdit[];
}): EditorWebviewOutgoingMessage | null {
    if (!activeSheet) {
        return null;
    }

    const filterState = getEditorFilterShellState({
        activeSheet,
        selection,
        pendingEdits,
    });
    if (!filterState.canToggle) {
        return null;
    }

    return {
        type: "setFilterState",
        sheetKey: activeSheet.key,
        filterState: filterState.nextFilterState,
    };
}
