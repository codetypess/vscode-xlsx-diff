/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { XlsxEditorPanel } from "../webview/editor-panel";

suite("Editor panel save guards", () => {
    test("ignores the next file watcher refresh after a local save", () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;

        panel.autoRefreshTimer = undefined;
        panel.suppressAutoRefreshUntil = 0;
        panel.ignoredAutoRefreshTriggerCount = 0;
        panel.ignoredAutoRefreshUntil = 0;
        panel.isSavingDocument = false;
        panel.hasWarnedPendingExternalChange = false;
        panel.document = {
            hasPendingEdits: () => false,
        };

        panel.noteLocalSaveCompletion();
        panel.scheduleAutoRefresh("change");

        assert.strictEqual(panel.autoRefreshTimer, undefined);
        assert.strictEqual(panel.ignoredAutoRefreshTriggerCount, 3);
        assert.ok(panel.suppressAutoRefreshUntil > Date.now());
    });

    test("restores auto refresh after save starts with an infinite suppression window", () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;

        panel.autoRefreshTimer = undefined;
        panel.suppressAutoRefreshUntil = Number.POSITIVE_INFINITY;
        panel.ignoredAutoRefreshTriggerCount = 0;
        panel.ignoredAutoRefreshUntil = 0;
        panel.isSavingDocument = false;
        panel.hasWarnedPendingExternalChange = false;
        panel.document = {
            hasPendingEdits: () => false,
        };

        panel.noteLocalSaveCompletion();

        assert.ok(Number.isFinite(panel.suppressAutoRefreshUntil));
        assert.ok(panel.suppressAutoRefreshUntil > Date.now());
    });

    test("reloads from disk after save when an external change was detected", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        let reloaded = false;
        let committed = false;

        panel.hasPendingExternalWorkbookChange = true;
        panel.isSavingDocument = true;
        panel.noteLocalSaveCompletion = () => undefined;
        panel.enqueueReload = async (options: { silent?: boolean; clearPendingEdits?: boolean }) => {
            reloaded = options.silent === true && options.clearPendingEdits === true;
        };
        panel.commitSavedState = async () => {
            committed = true;
        };

        await panel.handleDocumentSave();

        assert.strictEqual(panel.hasPendingExternalWorkbookChange, false);
        assert.strictEqual(reloaded, true);
        assert.strictEqual(committed, false);
    });

    test("prefers the panel with a pending external change when confirming save", async () => {
        const cleanPanel = Object.create(XlsxEditorPanel.prototype) as any;
        const changedPanel = Object.create(XlsxEditorPanel.prototype) as any;
        const document = {};
        const originalPanels = (XlsxEditorPanel as any).panels;
        let cleanPanelCalls = 0;
        let changedPanelCalls = 0;

        cleanPanel.document = document;
        cleanPanel.hasPendingExternalWorkbookChange = false;
        cleanPanel.confirmSaveIfNeeded = async (): Promise<boolean> => {
            cleanPanelCalls += 1;
            return false;
        };

        changedPanel.document = document;
        changedPanel.hasPendingExternalWorkbookChange = true;
        changedPanel.confirmSaveIfNeeded = async (): Promise<boolean> => {
            changedPanelCalls += 1;
            return true;
        };

        try {
            (XlsxEditorPanel as any).panels = new Map([
                [1, cleanPanel],
                [2, changedPanel],
            ]);

            const allowed = await XlsxEditorPanel.confirmDocumentSave(document as any);

            assert.strictEqual(allowed, true);
            assert.strictEqual(cleanPanelCalls, 0);
            assert.strictEqual(changedPanelCalls, 1);
        } finally {
            (XlsxEditorPanel as any).panels = originalPanels;
        }
    });

    test("bypasses the next save confirmation until the save flow finishes", async () => {
        const panel = Object.create(XlsxEditorPanel.prototype) as any;
        const document = {};
        const originalPanels = (XlsxEditorPanel as any).panels;

        panel.document = document;
        panel.hasPendingExternalWorkbookChange = true;
        panel.confirmSaveIfNeeded = async (): Promise<boolean> => false;

        try {
            (XlsxEditorPanel as any).panels = new Map([[1, panel]]);

            (XlsxEditorPanel as any).allowNextConfirmedSave(document);
            const allowedBeforeClear = await XlsxEditorPanel.confirmDocumentSave(document as any);

            (XlsxEditorPanel as any).clearConfirmedSaveBypass(document);
            const allowedAfterClear = await XlsxEditorPanel.confirmDocumentSave(document as any);

            assert.strictEqual(allowedBeforeClear, true);
            assert.strictEqual(allowedAfterClear, false);
        } finally {
            (XlsxEditorPanel as any).panels = originalPanels;
            (XlsxEditorPanel as any).clearConfirmedSaveBypass(document);
        }
    });
});
