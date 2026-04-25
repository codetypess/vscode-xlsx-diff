/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import type { SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import { createDiffPanelRenderModel } from "../webview/diff-panel-model";

function createWorkbook(overrides: Partial<WorkbookSnapshot>): WorkbookSnapshot {
    const sheet: SheetSnapshot = {
        name: "Sheet1",
        rowCount: 1,
        columnCount: 1,
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
});
