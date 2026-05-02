import type { EditorRenderPayload } from "../../core/model/types";
import type {
    EditorSessionIncomingMessage,
    EditorSessionPatch,
    EditorSessionRenderOptions,
} from "../shared/session-protocol";

export type EditorSessionMode = "idle" | "loading" | "error" | "ready";

export interface EditorSessionDocumentState {
    workbook: {
        title: string | null;
        sheets: EditorRenderPayload["sheets"];
        activeSheet: EditorRenderPayload["activeSheet"] | null;
        canEdit: boolean;
        hasPendingEdits: boolean;
        canUndoStructuralEdits: boolean;
        canRedoStructuralEdits: boolean;
    };
}

export interface EditorSessionEditingDraftState {
    pendingEdits: NonNullable<EditorSessionRenderOptions["replacePendingEdits"]>;
    clearRequested: boolean;
    preservePendingHistory: boolean;
    resetPendingHistory: boolean;
}

export interface EditorSessionViewportState {
    reuseActiveSheetData: boolean;
    useModelSelection: boolean;
}

export interface EditorSessionPanelState {
    silent: boolean;
    statusMessage: string | null;
    perfTraceId: string | null;
}

export interface EditorSessionUiState {
    selection: EditorRenderPayload["selection"] | null;
    editingDrafts: EditorSessionEditingDraftState;
    viewport: EditorSessionViewportState;
    panel: EditorSessionPanelState;
}

export interface EditorSessionState {
    initialized: boolean;
    mode: EditorSessionMode;
    document: EditorSessionDocumentState;
    ui: EditorSessionUiState;
    renderCount: number;
}

function createInitialEditorDocumentState(): EditorSessionDocumentState {
    return {
        workbook: {
            title: null,
            sheets: [],
            activeSheet: null,
            canEdit: false,
            hasPendingEdits: false,
            canUndoStructuralEdits: false,
            canRedoStructuralEdits: false,
        },
    };
}

function createInitialEditorUiState(): EditorSessionUiState {
    return {
        selection: null,
        editingDrafts: {
            pendingEdits: [],
            clearRequested: false,
            preservePendingHistory: false,
            resetPendingHistory: false,
        },
        viewport: {
            reuseActiveSheetData: false,
            useModelSelection: false,
        },
        panel: {
            silent: false,
            statusMessage: null,
            perfTraceId: null,
        },
    };
}

function hydrateEditorDocumentState(payload: EditorRenderPayload): EditorSessionDocumentState {
    return {
        workbook: {
            title: payload.title,
            sheets: payload.sheets,
            activeSheet: payload.activeSheet,
            canEdit: payload.canEdit,
            hasPendingEdits: payload.hasPendingEdits,
            canUndoStructuralEdits: payload.canUndoStructuralEdits,
            canRedoStructuralEdits: payload.canRedoStructuralEdits,
        },
    };
}

function hydrateEditorUiState(
    previousState: EditorSessionState,
    payload: EditorRenderPayload,
    options: EditorSessionRenderOptions,
    initializeSelection: boolean
): EditorSessionUiState {
    return {
        selection:
            initializeSelection || options.useModelSelection
                ? payload.selection
                : previousState.ui.selection,
        editingDrafts: {
            pendingEdits:
                options.replacePendingEdits ??
                (options.clearPendingEdits ? [] : previousState.ui.editingDrafts.pendingEdits),
            clearRequested: options.clearPendingEdits,
            preservePendingHistory: options.preservePendingHistory,
            resetPendingHistory: options.resetPendingHistory,
        },
        viewport: {
            reuseActiveSheetData: options.reuseActiveSheetData,
            useModelSelection: options.useModelSelection,
        },
        panel: {
            silent: options.silent,
            statusMessage: null,
            perfTraceId: options.perfTraceId,
        },
    };
}

function mergePatchedAlignmentMap(
    previousAlignments:
        | Readonly<
              Record<
                  string,
                  NonNullable<EditorRenderPayload["activeSheet"]["cellAlignments"]>[string]
              >
          >
        | undefined,
    nextAlignments:
        | Readonly<
              Record<
                  string,
                  NonNullable<EditorRenderPayload["activeSheet"]["cellAlignments"]>[string]
              >
          >
        | undefined,
    dirtyKeys: readonly string[] | undefined
):
    | Record<string, NonNullable<EditorRenderPayload["activeSheet"]["cellAlignments"]>[string]>
    | undefined {
    if (!dirtyKeys) {
        return nextAlignments
            ? { ...nextAlignments }
            : previousAlignments
              ? { ...previousAlignments }
              : undefined;
    }

    const merged = { ...(previousAlignments ?? {}) };
    for (const key of dirtyKeys) {
        if (nextAlignments && Object.hasOwn(nextAlignments, key)) {
            const nextAlignment = nextAlignments[key];
            if (nextAlignment) {
                merged[key] = nextAlignment;
                continue;
            }
        }

        delete merged[key];
    }

    return merged;
}

function mergePatchedActiveSheet(
    previousActiveSheet: EditorRenderPayload["activeSheet"],
    nextActiveSheet: EditorRenderPayload["activeSheet"]
): EditorRenderPayload["activeSheet"] {
    const mergedActiveSheet: EditorRenderPayload["activeSheet"] = {
        ...previousActiveSheet,
        ...nextActiveSheet,
    };

    if (
        nextActiveSheet.cellAlignmentDirtyKeys !== undefined ||
        nextActiveSheet.cellAlignments !== undefined
    ) {
        mergedActiveSheet.cellAlignments = mergePatchedAlignmentMap(
            previousActiveSheet.cellAlignments,
            nextActiveSheet.cellAlignments,
            nextActiveSheet.cellAlignmentDirtyKeys
        );
    }

    if (
        nextActiveSheet.rowAlignmentDirtyKeys !== undefined ||
        nextActiveSheet.rowAlignments !== undefined
    ) {
        mergedActiveSheet.rowAlignments = mergePatchedAlignmentMap(
            previousActiveSheet.rowAlignments,
            nextActiveSheet.rowAlignments,
            nextActiveSheet.rowAlignmentDirtyKeys
        );
    }

    if (
        nextActiveSheet.columnAlignmentDirtyKeys !== undefined ||
        nextActiveSheet.columnAlignments !== undefined
    ) {
        mergedActiveSheet.columnAlignments = mergePatchedAlignmentMap(
            previousActiveSheet.columnAlignments,
            nextActiveSheet.columnAlignments,
            nextActiveSheet.columnAlignmentDirtyKeys
        );
    }

    return mergedActiveSheet;
}

export function applyEditorSessionPatch(
    state: EditorSessionState,
    patch: EditorSessionPatch
): EditorSessionState {
    switch (patch.kind) {
        case "document:workbook":
            return {
                ...state,
                document: {
                    workbook: {
                        ...state.document.workbook,
                        ...(patch.title !== undefined ? { title: patch.title } : null),
                        ...(patch.sheets !== undefined ? { sheets: patch.sheets } : null),
                        ...(patch.canEdit !== undefined ? { canEdit: patch.canEdit } : null),
                        ...(patch.hasPendingEdits !== undefined
                            ? { hasPendingEdits: patch.hasPendingEdits }
                            : null),
                        ...(patch.canUndoStructuralEdits !== undefined
                            ? { canUndoStructuralEdits: patch.canUndoStructuralEdits }
                            : null),
                        ...(patch.canRedoStructuralEdits !== undefined
                            ? { canRedoStructuralEdits: patch.canRedoStructuralEdits }
                            : null),
                    },
                },
            };
        case "document:activeSheet":
            return {
                ...state,
                document: {
                    workbook: {
                        ...state.document.workbook,
                        activeSheet:
                            patch.activeSheet &&
                            state.document.workbook.activeSheet &&
                            state.document.workbook.activeSheet.key === patch.activeSheet.key
                                ? mergePatchedActiveSheet(
                                      state.document.workbook.activeSheet,
                                      patch.activeSheet
                                  )
                                : patch.activeSheet,
                    },
                },
            };
        case "ui:selection":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    selection: patch.selection,
                },
            };
        case "ui:editingDrafts":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    editingDrafts: {
                        pendingEdits:
                            patch.pendingEdits ??
                            (patch.clearPendingEdits ? [] : state.ui.editingDrafts.pendingEdits),
                        clearRequested:
                            patch.clearPendingEdits ?? state.ui.editingDrafts.clearRequested,
                        preservePendingHistory:
                            patch.preservePendingHistory ??
                            state.ui.editingDrafts.preservePendingHistory,
                        resetPendingHistory:
                            patch.resetPendingHistory ?? state.ui.editingDrafts.resetPendingHistory,
                    },
                },
            };
        case "ui:viewport":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    viewport: {
                        reuseActiveSheetData:
                            patch.reuseActiveSheetData ?? state.ui.viewport.reuseActiveSheetData,
                        useModelSelection:
                            patch.useModelSelection ?? state.ui.viewport.useModelSelection,
                    },
                },
            };
        case "ui:panel":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    panel: {
                        silent: patch.silent ?? state.ui.panel.silent,
                        statusMessage:
                            patch.statusMessage !== undefined
                                ? patch.statusMessage
                                : state.ui.panel.statusMessage,
                        perfTraceId:
                            patch.perfTraceId !== undefined
                                ? patch.perfTraceId
                                : state.ui.panel.perfTraceId,
                    },
                },
            };
    }
}

export function createInitialEditorSessionState(): EditorSessionState {
    return {
        initialized: false,
        mode: "idle",
        document: createInitialEditorDocumentState(),
        ui: createInitialEditorUiState(),
        renderCount: 0,
    };
}

export function reduceEditorSessionMessage(
    state: EditorSessionState,
    message: EditorSessionIncomingMessage
): EditorSessionState {
    switch (message.type) {
        case "session:status":
            return {
                ...state,
                mode: message.status.kind,
                ui: {
                    ...state.ui,
                    panel: {
                        ...state.ui.panel,
                        statusMessage: message.status.message,
                    },
                },
            };
        case "session:init":
        case "session:update":
            return {
                initialized: true,
                mode: "ready",
                document: hydrateEditorDocumentState(message.payload),
                ui: hydrateEditorUiState(
                    state,
                    message.payload,
                    message.options,
                    message.type === "session:init"
                ),
                renderCount: state.renderCount + 1,
            };
        case "session:patch": {
            const patchedState = message.patches.reduce(applyEditorSessionPatch, state);
            return {
                ...patchedState,
                initialized: true,
                mode: "ready",
                renderCount: state.renderCount + 1,
            };
        }
    }
}
