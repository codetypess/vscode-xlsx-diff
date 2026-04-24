/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    shouldSyncLocalSelectionDomFromModelSelection,
    shouldUseLocalSimpleSelectionUpdate,
} from "../webview/editor-selection-render";

suite("Editor selection render helpers", () => {
    test("allows local selection updates when no full render is needed", () => {
        assert.strictEqual(
            shouldUseLocalSimpleSelectionUpdate({
                hasNextCell: true,
                hasModel: true,
                hasEditingCell: false,
                isNextCellVisible: true,
                hasExpandedSelection: false,
                isSimpleSelection: true,
            }),
            true
        );
    });

    test("forces a full render after leaving edit mode", () => {
        assert.strictEqual(
            shouldUseLocalSimpleSelectionUpdate({
                hasNextCell: true,
                hasModel: true,
                hasEditingCell: false,
                isNextCellVisible: true,
                hasExpandedSelection: false,
                isSimpleSelection: true,
                forceRender: true,
            }),
            false
        );
    });

    test("reconciles simple local selection DOM for model-driven jumps", () => {
        assert.strictEqual(
            shouldSyncLocalSelectionDomFromModelSelection(
                { rowNumber: 5, columnNumber: 4 },
                { rowNumber: 5, columnNumber: 4 },
                { rowNumber: 10, columnNumber: 6 }
            ),
            true
        );
    });

    test("skips local selection DOM reconciliation for expanded selections", () => {
        assert.strictEqual(
            shouldSyncLocalSelectionDomFromModelSelection(
                { rowNumber: 5, columnNumber: 4 },
                { rowNumber: 7, columnNumber: 6 },
                { rowNumber: 10, columnNumber: 6 }
            ),
            false
        );
    });
});
