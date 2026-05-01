/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { EditorRenderModel, EditorRenderPayload } from "../core/model/types";
import type { EditorSessionPatchMessage } from "../webview-solid/shared/session-protocol";
import { XlsxEditorPanel } from "../webview/editor-panel/editor-panel";

function createEditorRenderModel(): EditorRenderModel {
    return {
        title: "Editor",
        activeSheet: {
            key: "sheet:1",
            rowCount: 4,
            columnCount: 3,
            columns: ["A", "B", "C"],
            cells: {},
            freezePane: null,
            autoFilter: null,
        },
        selection: null,
        hasPendingEdits: true,
        canEdit: true,
        sheets: [{ key: "sheet:1", label: "Sheet1", isActive: true }],
        canUndoStructuralEdits: true,
        canRedoStructuralEdits: false,
    };
}

function createEditorRenderPayload(): EditorRenderPayload {
    return {
        ...createEditorRenderModel(),
        activeSheet: {
            key: "sheet:1",
            rowCount: 4,
            columnCount: 3,
            columns: ["A", "B", "C"],
            cells: {},
            freezePane: null,
            autoFilter: null,
        },
    };
}

suite("Editor panel session protocol", () => {
    test("rerendering after initialization emits a session patch", async () => {
        const sentMessages: unknown[] = [];
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        panel.workbook = {};
        panel.panel = {
            title: "",
            webview: {
                postMessage: async (message: unknown) => {
                    sentMessages.push(message);
                },
            },
        };
        panel.isWebviewReady = true;
        panel.hasPendingRender = false;
        panel.hasSentSessionInit = true;
        panel.activePerfTraceId = null;
        panel.lastRenderedModel = null;
        panel.getWorkingWorkbook = () => ({});
        panel.getSheetEntries = () => [];
        panel.createRenderPayload = () => createEditorRenderPayload();

        await panel.render(createEditorRenderModel(), {
            silent: true,
            clearPendingEdits: false,
            preservePendingHistory: true,
            reuseActiveSheetData: false,
            useModelSelection: false,
            resetPendingHistory: true,
        });

        assert.strictEqual(sentMessages.length, 1);
        const message = sentMessages[0] as EditorSessionPatchMessage;

        assert.strictEqual(message.type, "session:patch");
        assert.strictEqual(message.view, "editor");
        assert.strictEqual(message.patches.length, 5);
        assert.strictEqual(message.patches[0]?.kind, "document:workbook");
        assert.strictEqual(message.patches[1]?.kind, "document:activeSheet");
        assert.strictEqual(message.patches[2]?.kind, "ui:editingDrafts");
        assert.strictEqual(message.patches[3]?.kind, "ui:viewport");
        assert.strictEqual(message.patches[4]?.kind, "ui:panel");

        const workbookPatch = message.patches[0];
        const editingDraftPatch = message.patches[2];
        const viewportPatch = message.patches[3];
        const panelPatch = message.patches[4];

        assert.ok(workbookPatch && workbookPatch.kind === "document:workbook");
        assert.ok(editingDraftPatch && editingDraftPatch.kind === "ui:editingDrafts");
        assert.ok(viewportPatch && viewportPatch.kind === "ui:viewport");
        assert.ok(panelPatch && panelPatch.kind === "ui:panel");

        assert.strictEqual(workbookPatch.title, "Editor");
        assert.strictEqual(workbookPatch.hasPendingEdits, true);
        assert.strictEqual(editingDraftPatch.clearPendingEdits, false);
        assert.strictEqual(editingDraftPatch.preservePendingHistory, true);
        assert.strictEqual(editingDraftPatch.resetPendingHistory, true);
        assert.strictEqual(viewportPatch.reuseActiveSheetData, false);
        assert.strictEqual(viewportPatch.useModelSelection, false);
        assert.strictEqual(panelPatch.statusMessage, null);
        assert.strictEqual(panelPatch.perfTraceId, null);
    });
});
