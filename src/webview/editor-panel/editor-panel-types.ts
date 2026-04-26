import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../../core/fastxlsx/write-cell-value";
import type { EditorPanelState, SheetSnapshot } from "../../core/model/types";
import type { SelectionRange } from "./editor-selection-range";

export interface SearchOptions {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
}

export type EditorSearchDirection = "next" | "prev";
export type EditorSearchScope = "sheet" | "selection";
export type EditorSearchStatus = "matched" | "no-match" | "invalid-pattern";

export interface EditorSearchMatch {
    sheetKey: string;
    rowNumber: number;
    columnNumber: number;
}

export interface EditorSearchRequest {
    query: string;
    direction: EditorSearchDirection;
    options: SearchOptions;
    scope: EditorSearchScope;
    selectionRange?: SelectionRange;
}

export interface EditorSearchResult {
    status: EditorSearchStatus;
    match?: EditorSearchMatch;
    matchCount?: number;
    matchIndex?: number;
}

export interface EditorSearchResultMessage extends EditorSearchResult {
    type: "searchResult";
    scope: EditorSearchScope;
    message?: string;
}

export type EditorWebviewMessage =
    | { type: "ready" }
    | { type: "setSheet"; sheetKey: string }
    | { type: "addSheet" }
    | { type: "deleteSheet"; sheetKey: string }
    | { type: "renameSheet"; sheetKey: string }
    | { type: "insertRow"; rowNumber: number }
    | { type: "deleteRow"; rowNumber: number }
    | { type: "insertColumn"; columnNumber: number }
    | { type: "deleteColumn"; columnNumber: number }
    | ({ type: "search" } & EditorSearchRequest)
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
    search: string;
    searchFind: string;
    searchReplace: string;
    searchReplaceComingSoon: string;
    searchScopeLabel: string;
    searchScopeSheet: string;
    searchScopeSelection: string;
    searchScopeSelectionDisabled: string;
    searchScopeWholeSheet: string;
    searchClose: string;
    loading: string;
    reload: string;
    save: string;
    lockView: string;
    unlockView: string;
    moreSheets: string;
    addSheet: string;
    deleteSheet: string;
    renameSheet: string;
    insertRowAbove: string;
    insertRowBelow: string;
    deleteRow: string;
    insertColumnLeft: string;
    insertColumnRight: string;
    deleteColumn: string;
    renameSheetPrompt: string;
    renameSheetTitle: string;
    sheetNameEmpty: string;
    sheetNameDuplicate: string;
    sheetNameTooLong: string;
    sheetNameInvalidChars: string;
    undo: string;
    redo: string;
    searchPlaceholder: string;
    replacePlaceholder: string;
    findPrev: string;
    findNext: string;
    replaceAll: string;
    gotoPlaceholder: string;
    goto: string;
    cancelInput: string;
    confirmInput: string;
    selectedCell: string;
    multipleCellsSelected: string;
    noCellSelected: string;
    noRowsAvailable: string;
    localChangesBlockedReload: string;
    externalChangesSavePrompt: string;
    externalChangesSaveAnyway: string;
    externalChangesReload: string;
    displayLanguageRefreshBlocked: string;
    noSearchMatches: string;
    invalidCellReference: string;
    invalidSearchPattern: string;
    searchMatchFound: string;
    searchMatchFoundInSelection: string;
    searchMatchSummary: string;
    replaceCount: string;
    replaceNoEditableMatches: string;
    replaceNoChanges: string;
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
