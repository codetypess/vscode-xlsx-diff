/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    getEditorToolbarSyncSnapshot,
    notifyEditorToolbarSync,
    resetEditorToolbarSyncForTests,
    subscribeEditorToolbarSync,
} from "../webview/editor-toolbar-sync";

suite("Editor toolbar sync", () => {
    setup(() => {
        resetEditorToolbarSyncForTests();
    });

    test("increments the snapshot and notifies subscribers", () => {
        let notifications = 0;
        const unsubscribe = subscribeEditorToolbarSync(() => {
            notifications += 1;
        });

        assert.strictEqual(getEditorToolbarSyncSnapshot(), 0);

        notifyEditorToolbarSync();
        notifyEditorToolbarSync();

        assert.strictEqual(getEditorToolbarSyncSnapshot(), 2);
        assert.strictEqual(notifications, 2);

        unsubscribe();
    });

    test("stops notifying after unsubscribe", () => {
        let notifications = 0;
        const unsubscribe = subscribeEditorToolbarSync(() => {
            notifications += 1;
        });

        unsubscribe();
        notifyEditorToolbarSync();

        assert.strictEqual(notifications, 0);
        assert.strictEqual(getEditorToolbarSyncSnapshot(), 1);
    });
});
