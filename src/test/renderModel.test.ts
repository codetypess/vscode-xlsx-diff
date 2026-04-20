/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/buildWorkbookDiff";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    createInitialPanelState,
    createRenderModel,
    moveDiffCursor,
    movePageCursor,
    setCurrentPage,
    setHighlightedDiffCell,
    setHighlightedDiffRow,
} from "../webview/renderModel";

function createCell(rowNumber: number, columnNumber: number, displayValue: string): CellSnapshot {
    return {
        key: `${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        address: `R${rowNumber}C${columnNumber}`,
        displayValue,
        formula: null,
        styleId: null,
    };
}

function createSheet(
    name: string,
    cells: CellSnapshot[],
    rowCount = 1,
    columnCount = 1
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        mergedRanges: [],
        rowHeights: {},
        columnWidths: {},
        cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
        signature: `${name}-signature`,
    };
}

function createWorkbookWithSheets(
    overrides: Partial<WorkbookSnapshot>,
    sheets: SheetSnapshot[]
): WorkbookSnapshot {
    return {
        filePath: "/tmp/item.xlsx",
        fileName: "item.xlsx",
        fileSize: 128,
        modifiedTime: new Date("2026-04-18T06:51:00.000Z").toISOString(),
        sheets,
        ...overrides,
    };
}

function createWorkbook(overrides: Partial<WorkbookSnapshot>, cellValue: string): WorkbookSnapshot {
    return {
        filePath: "/tmp/item.xlsx",
        fileName: "item.xlsx",
        fileSize: 128,
        modifiedTime: new Date("2026-04-18T06:51:00.000Z").toISOString(),
        sheets: [createSheet("item", [createCell(1, 1, cellValue)])],
        ...overrides,
    };
}

suite("Render model", () => {
    test("keeps commit detail, title hints, and read-only state aligned", () => {
        const left = createWorkbook(
            {
                detailLabel: "Commit",
                detailValue: "d4ce7e0",
                titleDetail: "d4ce7e0",
                modifiedTimeLabel: "Apr 18, 2026, 6:51 AM",
                isReadonly: true,
            },
            "left"
        );
        const right = createWorkbook({}, "right");

        const diff = buildWorkbookDiff(left, right);
        const renderModel = createRenderModel(diff, createInitialPanelState(diff));

        assert.strictEqual(renderModel.title, "item.xlsx (d4ce7e0) ↔ item.xlsx");
        assert.deepStrictEqual(renderModel.leftFile, {
            fileName: "item.xlsx",
            filePath: "/tmp/item.xlsx",
            fileSizeLabel: "128 B",
            detailLabel: "Commit",
            detailValue: "d4ce7e0",
            modifiedTimeLabel: "Apr 18, 2026, 6:51 AM",
            isReadonly: true,
        });
        assert.strictEqual(renderModel.rightFile.isReadonly, false);
    });

    test("moves diff navigation anchor to a clicked diff row", () => {
        const left = createWorkbookWithSheets({}, [
            createSheet("item", [createCell(2, 1, "old"), createCell(205, 1, "before")], 205, 1),
        ]);
        const right = createWorkbookWithSheets({}, [
            createSheet("item", [createCell(2, 1, "new"), createCell(205, 1, "after")], 205, 1),
        ]);

        const diff = buildWorkbookDiff(left, right);
        const state = setHighlightedDiffRow(diff, createInitialPanelState(diff), 205);
        const renderModel = createRenderModel(diff, state);

        assert.strictEqual(renderModel.page.currentPage, 2);
        assert.strictEqual(renderModel.page.highlightedDiffRow, 205);
        assert.strictEqual(renderModel.canPrevDiff, true);
        assert.strictEqual(renderModel.canNextDiff, false);
    });

    test("navigates diff cells from left to right and then top to bottom", () => {
        const left = createWorkbookWithSheets({}, [
            createSheet(
                "item",
                [createCell(2, 1, "old-a"), createCell(2, 2, "old-b"), createCell(3, 1, "old-c")],
                3,
                2
            ),
        ]);
        const right = createWorkbookWithSheets({}, [
            createSheet(
                "item",
                [createCell(2, 1, "new-a"), createCell(2, 2, "new-b"), createCell(3, 1, "new-c")],
                3,
                2
            ),
        ]);

        const diff = buildWorkbookDiff(left, right);
        const secondDiffState = moveDiffCursor(diff, createInitialPanelState(diff), 1);
        const thirdDiffState = moveDiffCursor(diff, secondDiffState, 1);
        const secondRenderModel = createRenderModel(diff, secondDiffState);
        const thirdRenderModel = createRenderModel(diff, thirdDiffState);

        assert.strictEqual(secondRenderModel.page.highlightedDiffCell?.key, "2:2");
        assert.strictEqual(secondRenderModel.page.highlightedDiffRow, 2);
        assert.strictEqual(thirdRenderModel.page.highlightedDiffCell?.key, "3:1");
        assert.strictEqual(thirdRenderModel.page.highlightedDiffRow, 3);
        assert.strictEqual(thirdRenderModel.canNextDiff, false);
    });

    test("keeps clicked diff cell as the diff navigation anchor", () => {
        const left = createWorkbookWithSheets({}, [
            createSheet("item", [createCell(2, 1, "old-a"), createCell(2, 2, "old-b")], 2, 2),
        ]);
        const right = createWorkbookWithSheets({}, [
            createSheet("item", [createCell(2, 1, "new-a"), createCell(2, 2, "new-b")], 2, 2),
        ]);

        const diff = buildWorkbookDiff(left, right);
        const state = setHighlightedDiffCell(diff, createInitialPanelState(diff), 2, 2);
        const renderModel = createRenderModel(diff, state);

        assert.strictEqual(renderModel.page.highlightedDiffCell?.key, "2:2");
        assert.strictEqual(renderModel.canPrevDiff, true);
        assert.strictEqual(renderModel.canNextDiff, false);
    });

    test("moves page navigation across sheet boundaries", () => {
        const left = createWorkbookWithSheets({}, [
            createSheet("first", [createCell(205, 1, "old-first")], 205, 1),
            createSheet("second", [createCell(1, 1, "old-second")], 1, 1),
        ]);
        const right = createWorkbookWithSheets({}, [
            createSheet("first", [createCell(205, 1, "new-first")], 205, 1),
            createSheet("second", [createCell(1, 1, "new-second")], 1, 1),
        ]);

        const diff = buildWorkbookDiff(left, right);
        const secondPageState = setCurrentPage(diff, createInitialPanelState(diff), 2);
        const crossSheetState = movePageCursor(diff, secondPageState, 1);
        const renderModel = createRenderModel(diff, crossSheetState);

        assert.strictEqual(renderModel.activeSheet.label, "second");
        assert.strictEqual(renderModel.page.currentPage, 1);
        assert.strictEqual(renderModel.canPrevPage, true);
        assert.strictEqual(renderModel.canNextPage, false);
    });
});
