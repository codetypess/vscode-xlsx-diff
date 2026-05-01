/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { EditorActiveSheetView } from "../core/model/types";
import {
    createEditorGridMetricsInputFromSheet,
    createEditorGridViewportPatchFromElement,
} from "../webview-solid/editor-panel/viewport-services";

suite("Solid editor viewport services", () => {
    test("maps an active sheet into grid metrics input", () => {
        const sheet: EditorActiveSheetView = {
            key: "sheet:1",
            rowCount: 128,
            columnCount: 24,
            columns: ["A", "B", "C"],
            columnWidths: [12, null, 18],
            rowHeights: { "2": 30 },
            cellAlignments: undefined,
            rowAlignments: undefined,
            columnAlignments: undefined,
            cells: {},
            freezePane: {
                rowCount: 1,
                columnCount: 2,
                topLeftCell: "C2",
                activePane: "bottomRight",
            },
            autoFilter: null,
        };

        assert.deepStrictEqual(createEditorGridMetricsInputFromSheet(sheet), {
            rowCount: 128,
            columnCount: 24,
            rowHeaderLabelCount: 128,
            rowHeights: { "2": 30 },
            columnWidths: [12, null, 18],
            freezePane: {
                rowCount: 1,
                columnCount: 2,
                topLeftCell: "C2",
                activePane: "bottomRight",
            },
        });
    });

    test("creates a viewport patch from a scroll container element", () => {
        assert.deepStrictEqual(
            createEditorGridViewportPatchFromElement({
                scrollTop: 420,
                scrollLeft: 96,
                clientHeight: 640,
                clientWidth: 1280,
            }),
            {
                scrollTop: 420,
                scrollLeft: 96,
                viewportHeight: 640,
                viewportWidth: 1280,
            }
        );
    });
});
