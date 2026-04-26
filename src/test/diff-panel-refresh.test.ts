/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { XlsxDiffPanel } from "../webview/diff-panel";

suite("Diff panel refresh", () => {
    test("queues clearPendingEdits while a reload is already in flight", async () => {
        const panel = Object.create(XlsxDiffPanel.prototype) as any;

        panel.isReloading = true;
        panel.queuedReloadOptions = undefined;

        await panel.enqueueReload({ silent: true, clearPendingEdits: true });

        assert.deepStrictEqual(panel.queuedReloadOptions, {
            silent: true,
            clearPendingEdits: true,
        });
    });

    test("replays queued reloads with their strongest options", async () => {
        const panel = Object.create(XlsxDiffPanel.prototype) as any;
        const reloadCalls: Array<{ silent?: boolean; clearPendingEdits?: boolean }> = [];

        panel.isReloading = false;
        panel.queuedReloadOptions = undefined;
        panel.reloadModel = async (options: { silent?: boolean; clearPendingEdits?: boolean }) => {
            reloadCalls.push(options);
            if (reloadCalls.length === 1) {
                await panel.enqueueReload({ silent: true, clearPendingEdits: true });
                await panel.enqueueReload({ silent: false, clearPendingEdits: false });
            }
        };

        await panel.enqueueReload({ silent: true, clearPendingEdits: false });

        assert.deepStrictEqual(reloadCalls, [
            { silent: true, clearPendingEdits: false },
            { silent: false, clearPendingEdits: true },
        ]);
        assert.strictEqual(panel.queuedReloadOptions, undefined);
        assert.strictEqual(panel.isReloading, false);
    });
});
