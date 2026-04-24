/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { shouldUseLocalSimpleSelectionUpdate } from "../webview/editor-selection-render";

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
});
