/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    isSelectionFocusCell,
    shouldResetInvisibleSelectionAnchor,
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

    test("marks the current selection cell as the focus cell", () => {
        assert.strictEqual(
            isSelectionFocusCell({ rowNumber: 7, columnNumber: 3 }, 7, 3),
            true
        );
        assert.strictEqual(
            isSelectionFocusCell({ rowNumber: 7, columnNumber: 3 }, 7, 4),
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

    test("resets an invisible anchor for simple selections", () => {
        assert.strictEqual(
            shouldResetInvisibleSelectionAnchor({
                hasSelectionRangeOverride: false,
                hasExpandedSelection: false,
                isAnchorVisible: false,
            }),
            true
        );
    });

    test("preserves an invisible anchor for expanded selections", () => {
        assert.strictEqual(
            shouldResetInvisibleSelectionAnchor({
                hasSelectionRangeOverride: false,
                hasExpandedSelection: true,
                isAnchorVisible: false,
            }),
            false
        );
    });
});
