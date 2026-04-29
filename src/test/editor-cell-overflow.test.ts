/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getCellOverflowMetrics } from "../webview/editor-panel/editor-cell-overflow";

suite("Editor cell overflow helpers", () => {
    test("extends text into consecutive blank cells on the right", () => {
        const widths = new Map([
            [1, 120],
            [2, 90],
            [3, 110],
            [4, 140],
        ]);

        const metrics = getCellOverflowMetrics({
            value: "hello",
            alignment: null,
            baseColumnWidth: widths.get(1)!,
            visibleColumnNumbers: [1, 2, 3, 4],
            visibleColumnIndex: 0,
            getColumnWidth: (columnNumber) => widths.get(columnNumber) ?? 0,
            getTrailingCellState: (columnNumber) =>
                columnNumber === 4
                    ? { value: "stop", formula: null }
                    : { value: "", formula: null },
        });

        assert.deepStrictEqual(metrics, {
            contentWidthPx: 306,
            spillsIntoNextCells: true,
        });
    });

    test("does not spill when the source cell is wrapped or right-aligned", () => {
        const wrappedMetrics = getCellOverflowMetrics({
            value: "hello",
            alignment: {
                wrapText: true,
            },
            baseColumnWidth: 120,
            visibleColumnNumbers: [1, 2],
            visibleColumnIndex: 0,
            getColumnWidth: () => 120,
            getTrailingCellState: () => ({ value: "", formula: null }),
        });
        const rightAlignedMetrics = getCellOverflowMetrics({
            value: "hello",
            alignment: {
                horizontal: "right",
            },
            baseColumnWidth: 120,
            visibleColumnNumbers: [1, 2],
            visibleColumnIndex: 0,
            getColumnWidth: () => 120,
            getTrailingCellState: () => ({ value: "", formula: null }),
        });

        assert.deepStrictEqual(wrappedMetrics, {
            contentWidthPx: 106,
            spillsIntoNextCells: false,
        });
        assert.deepStrictEqual(rightAlignedMetrics, {
            contentWidthPx: 106,
            spillsIntoNextCells: false,
        });
    });

    test("stops before cells that are visually empty but reserved for other UI", () => {
        const metrics = getCellOverflowMetrics({
            value: "hello",
            alignment: null,
            baseColumnWidth: 120,
            visibleColumnNumbers: [1, 2, 3],
            visibleColumnIndex: 0,
            getColumnWidth: () => 120,
            getTrailingCellState: (columnNumber) =>
                columnNumber === 2
                    ? {
                          value: "",
                          formula: null,
                          blocksOverflow: true,
                      }
                    : {
                          value: "",
                          formula: null,
                      },
        });

        assert.deepStrictEqual(metrics, {
            contentWidthPx: 106,
            spillsIntoNextCells: false,
        });
    });
});
