import * as vscode from "vscode";
import type { CellEdit } from "../../core/fastxlsx/write-cell-value";
import type { CellAlignmentSnapshot } from "../../core/model/alignment";
import { createCellKey } from "../../core/model/cells";
import type { EditorRenderModel, EditorRenderPayload } from "../../core/model/types";
import { getHtmlLanguageTag } from "../../display-language";
import {
    createEditorSessionInitMessage,
    createEditorSessionPatchMessage,
    type EditorSessionInitMessage,
    type EditorSessionPatchMessage,
} from "../../webview-solid/shared/session-protocol";
import { createWebviewNonce } from "../webview-utils";
import { mapPendingCellEditsToWebview } from "./editor-panel-state";
import type {
    EditorAlignmentTargetKind,
    EditorPanelStrings,
    WorkingSheetEntry,
} from "./editor-panel-types";
import type { SelectionRange } from "./editor-selection-range";

export interface ActiveSheetAlignmentRenderPatch {
    sheetKey: string;
    cellAlignmentKeys?: string[];
    rowAlignmentKeys?: string[];
    columnAlignmentKeys?: string[];
}

export function createActiveSheetAlignmentRenderPatch(
    sheetKey: string,
    target: EditorAlignmentTargetKind,
    selection: SelectionRange
): ActiveSheetAlignmentRenderPatch {
    if (target === "cell" || target === "range") {
        const cellAlignmentKeys: string[] = [];
        for (let rowNumber = selection.startRow; rowNumber <= selection.endRow; rowNumber += 1) {
            for (
                let columnNumber = selection.startColumn;
                columnNumber <= selection.endColumn;
                columnNumber += 1
            ) {
                cellAlignmentKeys.push(createCellKey(rowNumber, columnNumber));
            }
        }

        return {
            sheetKey,
            cellAlignmentKeys,
        };
    }

    if (target === "row") {
        const rowAlignmentKeys: string[] = [];
        for (let rowNumber = selection.startRow; rowNumber <= selection.endRow; rowNumber += 1) {
            rowAlignmentKeys.push(String(rowNumber));
        }

        return {
            sheetKey,
            rowAlignmentKeys,
        };
    }

    const columnAlignmentKeys: string[] = [];
    for (
        let columnNumber = selection.startColumn;
        columnNumber <= selection.endColumn;
        columnNumber += 1
    ) {
        columnAlignmentKeys.push(String(columnNumber));
    }

    return {
        sheetKey,
        columnAlignmentKeys,
    };
}

function pickAlignmentEntries<T>(
    alignments: Readonly<Record<string, T>> | undefined,
    keys: readonly string[]
): Record<string, T> {
    const nextEntries: Record<string, T> = {};
    for (const key of keys) {
        const value = alignments?.[key];
        if (value !== undefined) {
            nextEntries[key] = value;
        }
    }

    return nextEntries;
}

function canReuseRenderedActiveSheetData(
    previousRenderModel: EditorRenderModel | null,
    renderModel: EditorRenderModel
): previousRenderModel is EditorRenderModel {
    return Boolean(
        previousRenderModel &&
            previousRenderModel.activeSheet.key === renderModel.activeSheet.key &&
            previousRenderModel.activeSheet.rowCount === renderModel.activeSheet.rowCount &&
            previousRenderModel.activeSheet.columnCount === renderModel.activeSheet.columnCount
    );
}

export function createEditorPanelHtml({
    webview,
    extensionUri,
    strings,
    isDebugMode,
}: {
    webview: vscode.Webview;
    extensionUri: vscode.Uri;
    strings: EditorPanelStrings;
    isDebugMode: boolean;
}): string {
    const nonce = createWebviewNonce();
    const serializedStrings = JSON.stringify(strings).replace(/</g, "\\u003c");
    const scriptUri = webview.asWebviewUri(
        vscode.Uri.joinPath(extensionUri, "media", "editor-panel.js")
    );
    const styleUri = webview.asWebviewUri(
        vscode.Uri.joinPath(extensionUri, "media", "editor-panel.css")
    );
    const codiconStyleUri = webview.asWebviewUri(
        vscode.Uri.joinPath(extensionUri, "media", "codicons", "codicon.css")
    );

    return `<!DOCTYPE html>
<html lang="${getHtmlLanguageTag()}">
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource};" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${codiconStyleUri}" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Editor</title>
</head>
<body>
    <div id="app"></div>
	<script nonce="${nonce}">window.__XLSX_EDITOR_STRINGS__ = ${serializedStrings}; window.__XLSX_EDITOR_DEBUG__ = ${JSON.stringify(isDebugMode)};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
}

export function createEditorRenderPayload({
    renderModel,
    previousRenderModel,
    reuseActiveSheetData,
    alignmentRenderPatch,
}: {
    renderModel: EditorRenderModel;
    previousRenderModel: EditorRenderModel | null;
    reuseActiveSheetData: boolean;
    alignmentRenderPatch?: ActiveSheetAlignmentRenderPatch;
}): EditorRenderPayload {
    if (!reuseActiveSheetData || !canReuseRenderedActiveSheetData(previousRenderModel, renderModel)) {
        return renderModel;
    }

    const previousActiveSheet = previousRenderModel.activeSheet;
    const nextActiveSheet = renderModel.activeSheet;
    const payloadActiveSheet: EditorRenderPayload["activeSheet"] = {
        key: nextActiveSheet.key,
        rowCount: nextActiveSheet.rowCount,
        columnCount: nextActiveSheet.columnCount,
        freezePane: nextActiveSheet.freezePane,
        autoFilter: nextActiveSheet.autoFilter,
    };

    if (
        previousActiveSheet.columns.length !== nextActiveSheet.columns.length ||
        previousActiveSheet.columns.some((label, index) => label !== nextActiveSheet.columns[index])
    ) {
        payloadActiveSheet.columns = nextActiveSheet.columns;
    }

    if (previousActiveSheet.columnWidths !== nextActiveSheet.columnWidths) {
        payloadActiveSheet.columnWidths = nextActiveSheet.columnWidths;
    }

    if (previousActiveSheet.rowHeights !== nextActiveSheet.rowHeights) {
        payloadActiveSheet.rowHeights = nextActiveSheet.rowHeights;
    }

    if (previousActiveSheet.cellAlignments !== nextActiveSheet.cellAlignments) {
        if (alignmentRenderPatch?.sheetKey === nextActiveSheet.key && alignmentRenderPatch.cellAlignmentKeys) {
            payloadActiveSheet.cellAlignments = pickAlignmentEntries(
                nextActiveSheet.cellAlignments,
                alignmentRenderPatch.cellAlignmentKeys
            ) as Record<string, CellAlignmentSnapshot>;
            payloadActiveSheet.cellAlignmentDirtyKeys = [...alignmentRenderPatch.cellAlignmentKeys];
        } else {
            payloadActiveSheet.cellAlignments = nextActiveSheet.cellAlignments;
        }
    }

    if (previousActiveSheet.rowAlignments !== nextActiveSheet.rowAlignments) {
        if (alignmentRenderPatch?.sheetKey === nextActiveSheet.key && alignmentRenderPatch.rowAlignmentKeys) {
            payloadActiveSheet.rowAlignments = pickAlignmentEntries(
                nextActiveSheet.rowAlignments,
                alignmentRenderPatch.rowAlignmentKeys
            ) as Record<string, CellAlignmentSnapshot>;
            payloadActiveSheet.rowAlignmentDirtyKeys = [...alignmentRenderPatch.rowAlignmentKeys];
        } else {
            payloadActiveSheet.rowAlignments = nextActiveSheet.rowAlignments;
        }
    }

    if (previousActiveSheet.columnAlignments !== nextActiveSheet.columnAlignments) {
        if (
            alignmentRenderPatch?.sheetKey === nextActiveSheet.key &&
            alignmentRenderPatch.columnAlignmentKeys
        ) {
            payloadActiveSheet.columnAlignments = pickAlignmentEntries(
                nextActiveSheet.columnAlignments,
                alignmentRenderPatch.columnAlignmentKeys
            ) as Record<string, CellAlignmentSnapshot>;
            payloadActiveSheet.columnAlignmentDirtyKeys = [
                ...alignmentRenderPatch.columnAlignmentKeys,
            ];
        } else {
            payloadActiveSheet.columnAlignments = nextActiveSheet.columnAlignments;
        }
    }

    if (previousActiveSheet.cells !== nextActiveSheet.cells) {
        payloadActiveSheet.cells = nextActiveSheet.cells;
    }

    return {
        ...renderModel,
        activeSheet: payloadActiveSheet,
    };
}

export function createEditorSessionMessage({
    hasSentSessionInit,
    renderPayload,
    silent,
    clearPendingEdits,
    preservePendingHistory,
    reuseActiveSheetData,
    useModelSelection,
    replacePendingEdits,
    sheetEntries,
    resetPendingHistory,
    perfTraceId,
}: {
    hasSentSessionInit: boolean;
    renderPayload: EditorRenderPayload;
    silent: boolean;
    clearPendingEdits: boolean;
    preservePendingHistory: boolean;
    reuseActiveSheetData: boolean;
    useModelSelection: boolean;
    replacePendingEdits?: CellEdit[];
    sheetEntries: WorkingSheetEntry[];
    resetPendingHistory: boolean;
    perfTraceId: string | null;
}): EditorSessionInitMessage | EditorSessionPatchMessage {
    const mappedPendingEdits =
        replacePendingEdits !== undefined
            ? mapPendingCellEditsToWebview(replacePendingEdits, sheetEntries)
            : undefined;

    return hasSentSessionInit
        ? createEditorSessionPatchMessage([
              {
                  kind: "document:workbook",
                  title: renderPayload.title,
                  sheets: renderPayload.sheets,
                  canEdit: renderPayload.canEdit,
                  hasPendingEdits: renderPayload.hasPendingEdits,
                  canUndoStructuralEdits: renderPayload.canUndoStructuralEdits,
                  canRedoStructuralEdits: renderPayload.canRedoStructuralEdits,
              },
              {
                  kind: "document:activeSheet",
                  activeSheet: renderPayload.activeSheet,
              },
              ...(useModelSelection
                  ? [
                        {
                            kind: "ui:selection" as const,
                            selection: renderPayload.selection,
                        },
                    ]
                  : []),
              {
                  kind: "ui:editingDrafts",
                  ...(mappedPendingEdits !== undefined
                      ? {
                            pendingEdits: mappedPendingEdits,
                        }
                      : null),
                  clearPendingEdits,
                  preservePendingHistory,
                  resetPendingHistory,
              },
              {
                  kind: "ui:viewport",
                  reuseActiveSheetData,
                  useModelSelection,
              },
              {
                  kind: "ui:panel",
                  silent,
                  statusMessage: null,
                  perfTraceId,
              },
          ])
        : createEditorSessionInitMessage(renderPayload, {
              silent,
              clearPendingEdits,
              preservePendingHistory,
              reuseActiveSheetData,
              useModelSelection,
              perfTraceId,
              replacePendingEdits: mappedPendingEdits,
              resetPendingHistory,
          });
}
