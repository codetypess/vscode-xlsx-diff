import * as vscode from "vscode";
import { loadWorkbookSnapshot } from "../../core/fastxlsx/load-workbook-snapshot";
import type {
    CellEdit,
    SheetEdit,
    SheetViewEdit,
    WorkbookEditState,
} from "../../core/fastxlsx/write-cell-value";
import type { EditorPanelState, EditorRenderModel, WorkbookSnapshot } from "../../core/model/types";
import { getWorkbookResourceName } from "../../workbook/resource-uri";
import {
    createCommittedWorkbookState,
    createWorkingSheetEntries,
    createWorkingWorkbook,
    restorePendingWorkbookState,
} from "./editor-panel-state";
import type { WorkingSheetEntry } from "./editor-panel-types";
import {
    createEditorRenderModel,
    createInitialEditorPanelState,
    normalizeEditorPanelState,
} from "./editor-render-model";

export interface ReloadedEditorWorkbookSession {
    workbook: WorkbookSnapshot;
    workingSheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
    nextNewSheetId: number;
    panelState: EditorPanelState;
    renderModel: EditorRenderModel;
}

export interface CommittedEditorWorkbookSession {
    workbook: WorkbookSnapshot;
    workingSheetEntries: WorkingSheetEntry[];
    pendingCellEdits: CellEdit[];
    pendingSheetEdits: SheetEdit[];
    pendingViewEdits: SheetViewEdit[];
    nextNewSheetId: number;
    panelState: EditorPanelState;
}

function decorateLoadedWorkbook(workbook: WorkbookSnapshot, workbookUri: vscode.Uri): WorkbookSnapshot {
    return {
        ...workbook,
        filePath: workbookUri.fsPath,
        fileName: getWorkbookResourceName(workbookUri),
    };
}

export function buildReloadedEditorWorkbookSession({
    workbook,
    workbookUri,
    clearPendingEdits,
    currentPanelState,
    pendingCellEdits,
    pendingSheetEdits,
    pendingViewEdits,
    documentPendingState,
    documentHasPendingEdits,
    canUndoStructuralEdits,
    canRedoStructuralEdits,
}: {
    workbook: WorkbookSnapshot;
    workbookUri: vscode.Uri;
    clearPendingEdits: boolean;
    currentPanelState: EditorPanelState;
    pendingCellEdits: readonly CellEdit[];
    pendingSheetEdits: readonly SheetEdit[];
    pendingViewEdits: readonly SheetViewEdit[];
    documentPendingState: WorkbookEditState;
    documentHasPendingEdits: boolean;
    canUndoStructuralEdits: boolean;
    canRedoStructuralEdits: boolean;
}): ReloadedEditorWorkbookSession {
    const nextWorkbook = decorateLoadedWorkbook(workbook, workbookUri);
    let nextWorkingSheetEntries = createWorkingSheetEntries(nextWorkbook);
    let nextPendingCellEdits = clearPendingEdits ? [] : [...pendingCellEdits];
    let nextPendingSheetEdits = clearPendingEdits ? [] : [...pendingSheetEdits];
    let nextPendingViewEdits = clearPendingEdits ? [] : [...pendingViewEdits];
    let nextNewSheetId = 1;

    const shouldRestoreDocumentPendingState =
        !clearPendingEdits &&
        nextPendingCellEdits.length === 0 &&
        nextPendingSheetEdits.length === 0 &&
        nextPendingViewEdits.length === 0 &&
        documentPendingState.cellEdits.length +
            documentPendingState.sheetEdits.length +
            (documentPendingState.viewEdits?.length ?? 0) >
            0;

    if (shouldRestoreDocumentPendingState) {
        const restored = restorePendingWorkbookState(nextWorkbook, documentPendingState);
        nextWorkingSheetEntries = restored.sheetEntries;
        nextPendingCellEdits = restored.pendingCellEdits;
        nextPendingSheetEdits = restored.pendingSheetEdits;
        nextPendingViewEdits = restored.pendingViewEdits;
        nextNewSheetId = restored.nextNewSheetId;
    }

    const workingWorkbook = createWorkingWorkbook(
        nextWorkbook,
        nextWorkingSheetEntries,
        nextPendingCellEdits
    );
    const nextPanelState = nextWorkingSheetEntries.length
        ? normalizeEditorPanelState(
              workingWorkbook,
              currentPanelState.activeSheetKey
                  ? currentPanelState
                  : createInitialEditorPanelState(workingWorkbook, nextWorkingSheetEntries),
              nextWorkingSheetEntries
          )
        : createInitialEditorPanelState(workingWorkbook, nextWorkingSheetEntries);

    return {
        workbook: nextWorkbook,
        workingSheetEntries: nextWorkingSheetEntries,
        pendingCellEdits: nextPendingCellEdits,
        pendingSheetEdits: nextPendingSheetEdits,
        pendingViewEdits: nextPendingViewEdits,
        nextNewSheetId,
        panelState: nextPanelState,
        renderModel: createEditorRenderModel(workingWorkbook, nextPanelState, {
            hasPendingEdits: documentHasPendingEdits,
            sheetEntries: nextWorkingSheetEntries,
            canUndoStructuralEdits,
            canRedoStructuralEdits,
        }),
    };
}

export async function loadEditorWorkbookSession({
    readUri,
    ...options
}: {
    readUri: vscode.Uri;
    workbookUri: vscode.Uri;
    clearPendingEdits: boolean;
    currentPanelState: EditorPanelState;
    pendingCellEdits: readonly CellEdit[];
    pendingSheetEdits: readonly SheetEdit[];
    pendingViewEdits: readonly SheetViewEdit[];
    documentPendingState: WorkbookEditState;
    documentHasPendingEdits: boolean;
    canUndoStructuralEdits: boolean;
    canRedoStructuralEdits: boolean;
}): Promise<ReloadedEditorWorkbookSession> {
    const workbook = await loadWorkbookSnapshot(readUri);
    return buildReloadedEditorWorkbookSession({
        ...options,
        workbook,
    });
}

export function commitEditorWorkbookSession({
    workbook,
    workingSheetEntries,
    pendingCellEdits,
    currentPanelState,
}: {
    workbook: WorkbookSnapshot;
    workingSheetEntries: readonly WorkingSheetEntry[];
    pendingCellEdits: readonly CellEdit[];
    currentPanelState: EditorPanelState;
}): CommittedEditorWorkbookSession {
    const activeSheetName =
        currentPanelState.activeSheetKey
            ? workingSheetEntries.find((entry) => entry.key === currentPanelState.activeSheetKey)?.sheet
                  .name ?? null
            : null;
    const selectedCell = currentPanelState.selectedCell ? { ...currentPanelState.selectedCell } : null;
    const committedState = createCommittedWorkbookState(
        workbook,
        [...workingSheetEntries],
        [...pendingCellEdits]
    );
    const nextWorkbook = committedState.workbook;
    const nextWorkingSheetEntries = committedState.sheetEntries;
    const workingWorkbook = createWorkingWorkbook(nextWorkbook, nextWorkingSheetEntries, []);
    const nextActiveEntry =
        (activeSheetName
            ? nextWorkingSheetEntries.find((entry) => entry.sheet.name === activeSheetName)
            : null) ??
        nextWorkingSheetEntries[0] ??
        null;

    return {
        workbook: nextWorkbook,
        workingSheetEntries: nextWorkingSheetEntries,
        pendingCellEdits: [],
        pendingSheetEdits: [],
        pendingViewEdits: [],
        nextNewSheetId: 1,
        panelState: normalizeEditorPanelState(
            workingWorkbook,
            {
                activeSheetKey: nextActiveEntry?.key ?? null,
                selectedCell,
            },
            nextWorkingSheetEntries
        ),
    };
}
