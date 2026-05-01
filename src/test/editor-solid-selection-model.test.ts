/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey } from "../core/model/cells";
import type { EditorActiveSheetView } from "../core/model/types";
import { createOptimisticEditorSelection } from "../webview-solid/editor-panel/selection-model-helpers";

function createActiveSheet(): EditorActiveSheetView {
    return {
        key: "sheet:1",
        rowCount: 10,
        columnCount: 8,
        columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
        cells: {
            [createCellKey(2, 3)]: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                displayValue: "hello",
                formula: null,
                styleId: null,
            },
            [createCellKey(4, 5)]: {
                key: createCellKey(4, 5),
                rowNumber: 4,
                columnNumber: 5,
                address: "E4",
                displayValue: "42",
                formula: "=SUM(A1:A2)",
                styleId: null,
            },
        },
        freezePane: null,
        autoFilter: null,
    };
}

suite("Solid editor selection model", () => {
    test("creates an optimistic selection for an existing cell", () => {
        const selection = createOptimisticEditorSelection({
            activeSheet: createActiveSheet(),
            rowNumber: 2,
            columnNumber: 3,
        });

        assert.ok(selection);
        assert.strictEqual(selection?.address, "C2");
        assert.strictEqual(selection?.value, "hello");
        assert.strictEqual(selection?.formula, null);
        assert.strictEqual(selection?.isPresent, true);
    });

    test("creates an optimistic selection for an empty cell", () => {
        const selection = createOptimisticEditorSelection({
            activeSheet: createActiveSheet(),
            rowNumber: 3,
            columnNumber: 4,
        });

        assert.ok(selection);
        assert.strictEqual(selection?.address, "D3");
        assert.strictEqual(selection?.value, "");
        assert.strictEqual(selection?.formula, null);
        assert.strictEqual(selection?.isPresent, false);
    });

    test("returns null for out-of-bounds coordinates", () => {
        const selection = createOptimisticEditorSelection({
            activeSheet: createActiveSheet(),
            rowNumber: 11,
            columnNumber: 2,
        });

        assert.strictEqual(selection, null);
    });
});
