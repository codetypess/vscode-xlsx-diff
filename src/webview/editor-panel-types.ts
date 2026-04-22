import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../core/fastxlsx/write-cell-value";
import type { EditorPanelState, SheetSnapshot } from "../core/model/types";

export interface SearchOptions {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

export type EditorWebviewMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "addSheet" }
    | { type: "deleteSheet"; sheetKey: string }
    | { type: "renameSheet"; sheetKey: string }
    | { type: "setViewportStartRow"; rowNumber: number }
    | {
          type: "search";
          query: string;
          direction: "next" | "prev";
          options: SearchOptions;
      }
    | { type: "gotoCell"; reference: string }
    | { type: "selectCell"; rowNumber: number; columnNumber: number }
    | {
          type: "setPendingEdits";
          edits: EditorPendingEdit[];
      }
    | { type: "requestSave" }
    | { type: "pendingEditStateChanged"; hasPendingEdits: boolean }
    | { type: "undoSheetEdit" }
    | { type: "redoSheetEdit" }
    | { type: "toggleViewLock"; rowCount: number; columnCount: number }
    | { type: "reload" };

export interface XlsxEditorPanelController {
    onPendingStateChanged(state: WorkbookEditState): Promise<void> | void;
    onRequestSave(): Promise<void> | void;
    onRequestRevert(): Promise<void> | void;
}

export interface EditorPanelStrings {
    loading: string;
    reload: string;
    size: string;
    modified: string;
    sheet: string;
    rows: string;
    noRows: string;
    visibleRows: string;
    readOnly: string;
    save: string;
    lockView: string;
    unlockView: string;
    addSheet: string;
    deleteSheet: string;
    renameSheet: string;
    renameSheetPrompt: string;
    renameSheetTitle: string;
    sheetNameEmpty: string;
    sheetNameDuplicate: string;
    sheetNameTooLong: string;
    sheetNameInvalidChars: string;
    undo: string;
    redo: string;
    searchPlaceholder: string;
    findPrev: string;
    findNext: string;
    gotoPlaceholder: string;
    goto: string;
    totalSheets: string;
    totalRows: string;
    nonEmptyCells: string;
    selectedCell: string;
    noCellSelected: string;
    mergedRanges: string;
    pendingChanges: string;
    noRowsAvailable: string;
    readOnlyBadge: string;
    localChangesBlockedReload: string;
    confirmReloadDiscard: string;
    discardChangesAndReload: string;
    keepEditing: string;
    displayLanguageRefreshBlocked: string;
    noSearchMatches: string;
    invalidCellReference: string;
    invalidSearchPattern: string;
    searchRegex: string;
    searchMatchCase: string;
    searchWholeWord: string;
}

export interface WorkingSheetEntry {
    key: string;
    index: number;
    sheet: SheetSnapshot;
}

export interface StructuralSnapshot {
    state: EditorPanelState;
    sheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
}

export interface StructuralHistoryEntry {
    before: StructuralSnapshot;
    after: StructuralSnapshot;
    resetPendingHistory: boolean;
}

export interface RestoredStructuralState {
    state: EditorPanelState;
    sheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
}

export interface EditorPendingEdit {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
    value: string;
}

export interface EditorSearchMatch {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
}
