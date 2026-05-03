import type { EditorRenderPayload } from "../../core/model/types";
import type { DiffPanelRenderModel } from "../diff-panel/diff-panel-types";
import type {
    EditorPendingEdit,
    EditorWebviewMessage,
} from "../editor-panel/editor-panel-types";

export type WebviewSessionView = "editor" | "diff";
export type WebviewSessionStatusKind = "loading" | "error";

export interface WebviewSessionStatusPayload {
    kind: WebviewSessionStatusKind;
    message: string;
}

export interface WebviewSessionStatusMessage<
    TView extends WebviewSessionView = WebviewSessionView,
> {
    type: "session:status";
    view: TView;
    status: WebviewSessionStatusPayload;
}

export interface EditorSessionRenderOptions {
    silent: boolean;
    clearPendingEdits: boolean;
    preservePendingHistory: boolean;
    reuseActiveSheetData: boolean;
    useModelSelection: boolean;
    perfTraceId: string | null;
    replacePendingEdits?: EditorPendingEdit[];
    resetPendingHistory: boolean;
}

export interface EditorSessionInitMessage {
    type: "session:init";
    view: "editor";
    payload: EditorRenderPayload;
    options: EditorSessionRenderOptions;
}

export interface EditorSessionUpdateMessage {
    type: "session:update";
    view: "editor";
    payload: EditorRenderPayload;
    options: EditorSessionRenderOptions;
}

export type EditorSessionPatch =
    | {
          kind: "document:workbook";
          title?: string;
          sheets?: EditorRenderPayload["sheets"];
          canEdit?: boolean;
          hasPendingEdits?: boolean;
          canUndoStructuralEdits?: boolean;
          canRedoStructuralEdits?: boolean;
      }
    | {
          kind: "document:activeSheet";
          activeSheet: EditorRenderPayload["activeSheet"] | null;
      }
    | {
          kind: "ui:selection";
          selection: EditorRenderPayload["selection"] | null;
      }
    | {
          kind: "ui:editingDrafts";
          pendingEdits?: EditorPendingEdit[];
          clearPendingEdits?: boolean;
          preservePendingHistory?: boolean;
          resetPendingHistory?: boolean;
      }
    | {
          kind: "ui:viewport";
          reuseActiveSheetData?: boolean;
          useModelSelection?: boolean;
      }
    | {
          kind: "ui:panel";
          silent?: boolean;
          statusMessage?: string | null;
          perfTraceId?: string | null;
      };

export interface EditorSessionPatchMessage {
    type: "session:patch";
    view: "editor";
    patches: EditorSessionPatch[];
}

export interface DiffSessionRenderOptions {
    clearPendingEdits: boolean;
}

export interface DiffSessionInitMessage {
    type: "session:init";
    view: "diff";
    payload: DiffPanelRenderModel;
    options: DiffSessionRenderOptions;
}

export interface DiffSessionUpdateMessage {
    type: "session:update";
    view: "diff";
    payload: DiffPanelRenderModel;
    options: DiffSessionRenderOptions;
}

export type DiffSessionPatch =
    | {
          kind: "document:comparison";
          title?: string;
          leftFile?: DiffPanelRenderModel["leftFile"];
          rightFile?: DiffPanelRenderModel["rightFile"];
          definedNamesChanged?: boolean;
          sheets?: DiffPanelRenderModel["sheets"];
      }
    | {
          kind: "document:activeSheet";
          activeSheet: DiffPanelRenderModel["activeSheet"];
      }
    | {
          kind: "ui:navigation";
          activeSheetKey?: string | null;
      }
    | {
          kind: "ui:pendingEdits";
          clearPendingEdits?: boolean;
      }
    | {
          kind: "ui:panel";
          statusMessage?: string | null;
      };

export interface DiffSessionPatchMessage {
    type: "session:patch";
    view: "diff";
    patches: DiffSessionPatch[];
}

export type EditorSessionIncomingMessage =
    | WebviewSessionStatusMessage<"editor">
    | EditorSessionInitMessage
    | EditorSessionUpdateMessage
    | EditorSessionPatchMessage;

export type DiffSessionIncomingMessage =
    | WebviewSessionStatusMessage<"diff">
    | DiffSessionInitMessage
    | DiffSessionUpdateMessage
    | DiffSessionPatchMessage;

export interface WebviewReadyMessage {
    type: "ready";
}

export type EditorWebviewOutgoingMessage = EditorWebviewMessage;

export interface DiffWebviewPendingEdit {
    sheetKey: string;
    side: "left" | "right";
    rowNumber: number;
    columnNumber: number;
    value: string;
}

export type DiffWebviewOutgoingMessage =
    | WebviewReadyMessage
    | { type: "setSheet"; sheetKey: string }
    | { type: "saveEdits"; edits: DiffWebviewPendingEdit[] }
    | { type: "swap" }
    | { type: "reload" };

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === "object";
}

function isStringOrNull(value: unknown): value is string | null {
    return typeof value === "string" || value === null;
}

function isNumberOrNull(value: unknown): value is number | null {
    return typeof value === "number" || value === null;
}

function isSearchOptions(value: unknown): value is {
    isRegexp: boolean;
    matchCase: boolean;
    wholeWord: boolean;
} {
    return (
        isRecord(value) &&
        typeof value.isRegexp === "boolean" &&
        typeof value.matchCase === "boolean" &&
        typeof value.wholeWord === "boolean"
    );
}

function isDiffWebviewPendingEdit(value: unknown): value is DiffWebviewPendingEdit {
    return (
        isRecord(value) &&
        typeof value.sheetKey === "string" &&
        (value.side === "left" || value.side === "right") &&
        typeof value.rowNumber === "number" &&
        typeof value.columnNumber === "number" &&
        typeof value.value === "string"
    );
}

export function createWebviewReadyMessage(): WebviewReadyMessage {
    return { type: "ready" };
}

export function createSessionStatusMessage<TView extends WebviewSessionView>(
    view: TView,
    kind: WebviewSessionStatusKind,
    message: string
): WebviewSessionStatusMessage<TView> {
    return {
        type: "session:status",
        view,
        status: {
            kind,
            message,
        },
    };
}

export function createEditorSessionInitMessage(
    payload: EditorRenderPayload,
    options: EditorSessionRenderOptions
): EditorSessionInitMessage {
    return {
        type: "session:init",
        view: "editor",
        payload,
        options,
    };
}

export function createEditorSessionUpdateMessage(
    payload: EditorRenderPayload,
    options: EditorSessionRenderOptions
): EditorSessionUpdateMessage {
    return {
        type: "session:update",
        view: "editor",
        payload,
        options,
    };
}

export function createEditorSessionPatchMessage(
    patches: EditorSessionPatch[]
): EditorSessionPatchMessage {
    return {
        type: "session:patch",
        view: "editor",
        patches,
    };
}

export function createDiffSessionInitMessage(
    payload: DiffPanelRenderModel,
    options: DiffSessionRenderOptions
): DiffSessionInitMessage {
    return {
        type: "session:init",
        view: "diff",
        payload,
        options,
    };
}

export function createDiffSessionUpdateMessage(
    payload: DiffPanelRenderModel,
    options: DiffSessionRenderOptions
): DiffSessionUpdateMessage {
    return {
        type: "session:update",
        view: "diff",
        payload,
        options,
    };
}

export function createDiffSessionPatchMessage(
    patches: DiffSessionPatch[]
): DiffSessionPatchMessage {
    return {
        type: "session:patch",
        view: "diff",
        patches,
    };
}

export function isEditorSessionIncomingMessage(
    value: unknown
): value is EditorSessionIncomingMessage {
    if (!isRecord(value) || value.view !== "editor" || typeof value.type !== "string") {
        return false;
    }

    switch (value.type) {
        case "session:status":
            return (
                isRecord(value.status) &&
                (value.status.kind === "loading" || value.status.kind === "error") &&
                typeof value.status.message === "string"
            );
        case "session:init":
        case "session:update":
            return isRecord(value.payload) && isRecord(value.options);
        case "session:patch":
            return Array.isArray(value.patches);
        default:
            return false;
    }
}

export function isDiffSessionIncomingMessage(value: unknown): value is DiffSessionIncomingMessage {
    if (!isRecord(value) || value.view !== "diff" || typeof value.type !== "string") {
        return false;
    }

    switch (value.type) {
        case "session:status":
            return (
                isRecord(value.status) &&
                (value.status.kind === "loading" || value.status.kind === "error") &&
                typeof value.status.message === "string"
            );
        case "session:init":
        case "session:update":
            return isRecord(value.payload) && isRecord(value.options);
        case "session:patch":
            return Array.isArray(value.patches);
        default:
            return false;
    }
}

export function isEditorWebviewOutgoingMessage(
    value: unknown
): value is EditorWebviewOutgoingMessage {
    if (!isRecord(value) || typeof value.type !== "string") {
        return false;
    }

    switch (value.type) {
        case "ready":
        case "addSheet":
        case "requestSave":
        case "undoSheetEdit":
        case "redoSheetEdit":
        case "reload":
            return true;
        case "setSheet":
        case "deleteSheet":
        case "renameSheet":
            return typeof value.sheetKey === "string";
        case "insertRow":
        case "deleteRow":
        case "promptRowHeight":
            return typeof value.rowNumber === "number";
        case "setRowHeight":
            return typeof value.rowNumber === "number" && isNumberOrNull(value.height);
        case "insertColumn":
        case "deleteColumn":
        case "promptColumnWidth":
            return typeof value.columnNumber === "number";
        case "setColumnWidth":
            return typeof value.columnNumber === "number" && isNumberOrNull(value.width);
        case "setAlignment":
            return (
                (value.target === "cell" ||
                    value.target === "row" ||
                    value.target === "column" ||
                    value.target === "range") &&
                isRecord(value.selection) &&
                isRecord(value.alignment) &&
                (value.perfTraceId === undefined || typeof value.perfTraceId === "string")
            );
        case "search":
            return (
                typeof value.query === "string" &&
                (value.direction === "next" || value.direction === "prev") &&
                isSearchOptions(value.options) &&
                (value.scope === "sheet" || value.scope === "selection") &&
                (value.selectionRange === undefined || isRecord(value.selectionRange))
            );
        case "gotoCell":
            return typeof value.reference === "string";
        case "selectCell":
            return typeof value.rowNumber === "number" && typeof value.columnNumber === "number";
        case "setPendingEdits":
            return Array.isArray(value.edits);
        case "setFilterState":
            return (
                typeof value.sheetKey === "string" &&
                (value.filterState === null || isRecord(value.filterState))
            );
        case "pendingEditStateChanged":
            return typeof value.hasPendingEdits === "boolean";
        case "toggleViewLock":
            return typeof value.rowCount === "number" && typeof value.columnCount === "number";
        default:
            return false;
    }
}

export function isDiffWebviewOutgoingMessage(value: unknown): value is DiffWebviewOutgoingMessage {
    if (!isRecord(value) || typeof value.type !== "string") {
        return false;
    }

    switch (value.type) {
        case "ready":
        case "swap":
        case "reload":
            return true;
        case "setSheet":
            return typeof value.sheetKey === "string";
        case "saveEdits":
            return Array.isArray(value.edits) && value.edits.every(isDiffWebviewPendingEdit);
        default:
            return false;
    }
}
