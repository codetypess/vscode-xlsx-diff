import type { DiffPanelRenderModel } from "./diff-panel-types";
import type {
    DiffSessionIncomingMessage,
    DiffSessionPatch,
    DiffSessionRenderOptions,
} from "../shared/session-protocol";

export type DiffSessionMode = "idle" | "loading" | "error" | "ready";

export interface DiffSessionDocumentState {
    comparison: {
        title: string | null;
        leftFile: DiffPanelRenderModel["leftFile"] | null;
        rightFile: DiffPanelRenderModel["rightFile"] | null;
        definedNamesChanged: boolean;
        sheets: DiffPanelRenderModel["sheets"];
        activeSheet: DiffPanelRenderModel["activeSheet"];
    };
}

export interface DiffSessionUiState {
    navigation: {
        activeSheetKey: string | null;
    };
    pendingEdits: {
        clearRequested: boolean;
    };
    panel: {
        statusMessage: string | null;
    };
}

export interface DiffSessionState {
    initialized: boolean;
    mode: DiffSessionMode;
    document: DiffSessionDocumentState;
    ui: DiffSessionUiState;
    renderCount: number;
}

function createInitialDiffDocumentState(): DiffSessionDocumentState {
    return {
        comparison: {
            title: null,
            leftFile: null,
            rightFile: null,
            definedNamesChanged: false,
            sheets: [],
            activeSheet: null,
        },
    };
}

function createInitialDiffUiState(): DiffSessionUiState {
    return {
        navigation: {
            activeSheetKey: null,
        },
        pendingEdits: {
            clearRequested: false,
        },
        panel: {
            statusMessage: null,
        },
    };
}

function hydrateDiffDocumentState(payload: DiffPanelRenderModel): DiffSessionDocumentState {
    return {
        comparison: {
            title: payload.title,
            leftFile: payload.leftFile,
            rightFile: payload.rightFile,
            definedNamesChanged: payload.definedNamesChanged,
            sheets: payload.sheets,
            activeSheet: payload.activeSheet,
        },
    };
}

function hydrateDiffUiState(
    payload: DiffPanelRenderModel,
    options: DiffSessionRenderOptions
): DiffSessionUiState {
    return {
        navigation: {
            activeSheetKey: payload.activeSheet?.key ?? null,
        },
        pendingEdits: {
            clearRequested: options.clearPendingEdits,
        },
        panel: {
            statusMessage: null,
        },
    };
}

export function applyDiffSessionPatch(
    state: DiffSessionState,
    patch: DiffSessionPatch
): DiffSessionState {
    switch (patch.kind) {
        case "document:comparison":
            return {
                ...state,
                document: {
                    comparison: {
                        ...state.document.comparison,
                        ...(patch.title !== undefined ? { title: patch.title } : null),
                        ...(patch.leftFile !== undefined ? { leftFile: patch.leftFile } : null),
                        ...(patch.rightFile !== undefined ? { rightFile: patch.rightFile } : null),
                        ...(patch.definedNamesChanged !== undefined
                            ? { definedNamesChanged: patch.definedNamesChanged }
                            : null),
                        ...(patch.sheets !== undefined ? { sheets: patch.sheets } : null),
                    },
                },
            };
        case "document:activeSheet":
            return {
                ...state,
                document: {
                    comparison: {
                        ...state.document.comparison,
                        activeSheet: patch.activeSheet,
                    },
                },
            };
        case "ui:navigation":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    navigation: {
                        activeSheetKey:
                            patch.activeSheetKey !== undefined
                                ? patch.activeSheetKey
                                : state.ui.navigation.activeSheetKey,
                    },
                },
            };
        case "ui:pendingEdits":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    pendingEdits: {
                        clearRequested:
                            patch.clearPendingEdits ?? state.ui.pendingEdits.clearRequested,
                    },
                },
            };
        case "ui:panel":
            return {
                ...state,
                ui: {
                    ...state.ui,
                    panel: {
                        statusMessage:
                            patch.statusMessage !== undefined
                                ? patch.statusMessage
                                : state.ui.panel.statusMessage,
                    },
                },
            };
    }
}

export function createInitialDiffSessionState(): DiffSessionState {
    return {
        initialized: false,
        mode: "idle",
        document: createInitialDiffDocumentState(),
        ui: createInitialDiffUiState(),
        renderCount: 0,
    };
}

export function reduceDiffSessionMessage(
    state: DiffSessionState,
    message: DiffSessionIncomingMessage
): DiffSessionState {
    switch (message.type) {
        case "session:status":
            return {
                ...state,
                mode: message.status.kind,
                ui: {
                    ...state.ui,
                    panel: {
                        statusMessage: message.status.message,
                    },
                },
            };
        case "session:init":
        case "session:update":
            return {
                initialized: true,
                mode: "ready",
                document: hydrateDiffDocumentState(message.payload),
                ui: hydrateDiffUiState(message.payload, message.options),
                renderCount: state.renderCount + 1,
            };
        case "session:patch": {
            const patchedState = message.patches.reduce(applyDiffSessionPatch, state);
            return {
                ...patchedState,
                initialized: true,
                mode: "ready",
                renderCount: state.renderCount + 1,
            };
        }
    }
}
