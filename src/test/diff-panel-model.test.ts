/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { createCellKey, getCellAddress } from "../core/model/cells";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import { createDiffPanelRenderModel } from "../webview/diff-panel/diff-panel-model";

function createWorkbook(overrides: Partial<WorkbookSnapshot>): WorkbookSnapshot {
    const sheet: SheetSnapshot = {
        name: "Sheet1",
        rowCount: 1,
        columnCount: 1,
        visibility: "visible",
        mergedRanges: [],
        freezePane: null,
        cells: {},
        signature: "Sheet1",
    };

    return {
        filePath: "/tmp/item.xlsx",
        fileName: "item.xlsx",
        fileSize: 128,
        modifiedTime: new Date("2026-04-18T06:51:00.000Z").toISOString(),
        sheets: [sheet],
        ...overrides,
    };
}

suite("Diff panel render model", () => {
    test("preserves side-specific column labels after a column insertion", () => {
        const createCell = (
            rowNumber: number,
            columnNumber: number,
            value: string
        ): CellSnapshot => ({
            key: createCellKey(rowNumber, columnNumber),
            rowNumber,
            columnNumber,
            address: getCellAddress(rowNumber, columnNumber),
            displayValue: value,
            formula: null,
            styleId: null,
        });
        const createGridSheet = (name: string, rows: string[][]): SheetSnapshot => {
            const cells: Record<string, CellSnapshot> = {};
            let columnCount = 0;

            rows.forEach((row, rowIndex) => {
                columnCount = Math.max(columnCount, row.length);
                row.forEach((value, columnIndex) => {
                    if (!value) {
                        return;
                    }

                    const cell = createCell(rowIndex + 1, columnIndex + 1, value);
                    cells[cell.key] = cell;
                });
            });

            return {
                name,
                rowCount: rows.length,
                columnCount,
                visibility: "visible",
                mergedRanges: [],
                freezePane: null,
                cells,
                signature: `${name}:${rows.map((row) => row.join("|")).join("/")}`,
            };
        };

        const diff = buildWorkbookDiff(
            createWorkbook({
                sheets: [
                    createGridSheet("Sheet1", [
                        ["ID", "Name", "Score"],
                        ["1", "Alice", "90"],
                    ]),
                ],
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
                sheets: [
                    createGridSheet("Sheet1", [
                        ["ID", "Status", "Name", "Score"],
                        ["1", "New", "Alice", "90"],
                    ]),
                ],
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);
        const activeSheet = renderModel.activeSheet!;

        assert.deepStrictEqual(
            activeSheet.columns.map((column) => [column.leftLabel, column.rightLabel]),
            [
                ["A", "A"],
                ["", "B"],
                ["B", "C"],
                ["C", "D"],
            ]
        );
    });

    test("merges the primary workbook detail fact into the pane title", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                detailFacts: [
                    {
                        label: "Commit",
                        value: "d4ce7e0",
                        titleValue: "d4ce7e0",
                    },
                    {
                        label: "Committer",
                        value: "Alice <alice@example.com>",
                    },
                ],
                titleDetail: "d4ce7e0",
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.leftFile.title, "item.xlsx (d4ce7e0)");
        assert.deepStrictEqual(renderModel.leftFile.detailFacts, [
            {
                label: "Committer",
                value: "Alice <alice@example.com>",
                title: undefined,
            },
        ]);
    });

    test("removes a legacy single detail entry from the facts row after merging it into the title", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                detailLabel: "Commit",
                detailValue: "d4ce7e0",
                titleDetail: "d4ce7e0",
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.leftFile.title, "item.xlsx (d4ce7e0)");
        assert.deepStrictEqual(renderModel.leftFile.detailFacts, []);
    });

    test("uses the full primary detail value when a source fact is present", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                detailFacts: [
                    {
                        label: "Source",
                        value: "Index · base d4ce7e0",
                        titleValue: "d4ce7e0",
                    },
                    {
                        label: "Committer",
                        value: "Alice <alice@example.com>",
                    },
                ],
                titleDetail: "d4ce7e0",
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.leftFile.title, "item.xlsx (Index · base d4ce7e0)");
        assert.deepStrictEqual(renderModel.leftFile.detailFacts, [
            {
                label: "Committer",
                value: "Alice <alice@example.com>",
                title: undefined,
            },
        ]);
    });

    test("surfaces freeze-pane-only sheet changes as structural diffs", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                sheets: [
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:freeze:none",
                    },
                ],
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
                sheets: [
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: {
                            columnCount: 1,
                            rowCount: 1,
                            topLeftCell: "B2",
                            activePane: "bottomRight",
                        },
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:freeze:set",
                    },
                ],
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.sheets[0]?.hasDiff, true);
        assert.strictEqual(renderModel.sheets[0]?.freezePaneChanged, true);
        assert.strictEqual(renderModel.sheets[0]?.mergedRangesChanged, false);
        assert.strictEqual(renderModel.activeSheet?.freezePaneChanged, true);
        assert.strictEqual(renderModel.activeSheet?.diffCellCount, 0);
        assert.strictEqual(renderModel.activeSheet?.diffRowCount, 0);
    });

    test("surfaces sheet-visibility-only changes as structural diffs", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                sheets: [
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:visibility:visible",
                    },
                ],
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
                sheets: [
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "hidden",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:visibility:hidden",
                    },
                ],
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.sheets[0]?.hasDiff, true);
        assert.strictEqual(renderModel.sheets[0]?.visibilityChanged, true);
        assert.strictEqual(renderModel.sheets[0]?.freezePaneChanged, false);
        assert.strictEqual(renderModel.activeSheet?.visibilityChanged, true);
        assert.strictEqual(renderModel.activeSheet?.diffCellCount, 0);
        assert.strictEqual(renderModel.activeSheet?.diffRowCount, 0);
    });

    test("surfaces sheet-order-only changes as structural diffs", () => {
        const diff = buildWorkbookDiff(
            createWorkbook({
                sheets: [
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:order",
                    },
                    {
                        name: "Sheet2",
                        rowCount: 1,
                        columnCount: 1,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {},
                        signature: "Sheet2:order",
                    },
                ],
            }),
            createWorkbook({
                filePath: "/tmp/item-next.xlsx",
                fileName: "item-next.xlsx",
                sheets: [
                    {
                        name: "Sheet2",
                        rowCount: 1,
                        columnCount: 1,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {},
                        signature: "Sheet2:order",
                    },
                    {
                        name: "Sheet1",
                        rowCount: 5,
                        columnCount: 6,
                        visibility: "visible",
                        mergedRanges: [],
                        freezePane: null,
                        cells: {
                            "5:6": {
                                key: "5:6",
                                rowNumber: 5,
                                columnNumber: 6,
                                address: "F5",
                                displayValue: "same",
                                formula: null,
                                styleId: null,
                            },
                        },
                        signature: "Sheet1:order",
                    },
                ],
            })
        );

        const renderModel = createDiffPanelRenderModel(diff, null);

        assert.strictEqual(renderModel.sheets[0]?.hasDiff, true);
        assert.strictEqual(renderModel.sheets[0]?.sheetOrderChanged, true);
        assert.strictEqual(renderModel.sheets[1]?.sheetOrderChanged, true);
        assert.strictEqual(renderModel.activeSheet?.sheetOrderChanged, true);
        assert.strictEqual(renderModel.activeSheet?.diffCellCount, 0);
        assert.strictEqual(renderModel.activeSheet?.diffRowCount, 0);
    });
});
