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
});
