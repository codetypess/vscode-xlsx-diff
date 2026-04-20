import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/buildWorkbookDiff";
import { createPageSlice } from "../core/paging/createPageSlice";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";

function createCell(rowNumber: number, columnNumber: number, displayValue: string): CellSnapshot {
    return {
        key: `${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        address: `C${columnNumber}-${rowNumber}`,
        displayValue,
        formula: null,
        styleId: null,
    };
}

function createSheet(
    name: string,
    rowCount: number,
    columnCount: number,
    cells: CellSnapshot[],
    signature = `${name}-signature`
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        mergedRanges: [],
        cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
        signature,
    };
}

function createWorkbook(sheets: SheetSnapshot[]): WorkbookSnapshot {
    return {
        filePath: "/tmp/test.xlsx",
        fileName: "test.xlsx",
        fileSize: 128,
        modifiedTime: new Date("2026-01-01T00:00:00.000Z").toISOString(),
        sheets,
    };
}

suite("Diff engine", () => {
    test("matches renamed sheets by signature", () => {
        const left = createWorkbook([
            createSheet("Before", 1, 1, [createCell(1, 1, "hello")], "same-signature"),
        ]);
        const right = createWorkbook([
            createSheet("After", 1, 1, [createCell(1, 1, "hello")], "same-signature"),
        ]);

        const diff = buildWorkbookDiff(left, right);

        assert.strictEqual(diff.sheets.length, 1);
        assert.strictEqual(diff.sheets[0].kind, "renamed");
        assert.strictEqual(diff.sheets[0].leftSheetName, "Before");
        assert.strictEqual(diff.sheets[0].rightSheetName, "After");
    });

    test("builds diff rows and paginates them", () => {
        const left = createWorkbook([
            createSheet("Data", 205, 2, [createCell(2, 1, "old"), createCell(100, 2, "same")]),
        ]);
        const right = createWorkbook([
            createSheet("Data", 205, 2, [
                createCell(2, 1, "new"),
                createCell(4, 1, "added"),
                createCell(100, 2, "same"),
            ]),
        ]);

        const diff = buildWorkbookDiff(left, right);
        const sheet = diff.sheets[0];

        assert.deepStrictEqual(sheet.diffRows, [2, 4]);
        assert.deepStrictEqual(
            sheet.diffCells.map((cell) => cell.key),
            ["2:1", "4:1"]
        );
        assert.strictEqual(sheet.diffCellCount, 2);

        const diffPage = createPageSlice(sheet, "diffs", 1, "2:1");
        assert.deepStrictEqual(
            diffPage.rows.map((row) => row.rowNumber),
            [2, 4]
        );
        assert.strictEqual(diffPage.rows[0].diffTone, "modified");
        assert.strictEqual(diffPage.rows[1].diffTone, "added");
        assert.strictEqual(diffPage.sameRowCount, 203);
        assert.strictEqual(diffPage.highlightedDiffCell?.key, "2:1");

        const secondAllPage = createPageSlice(sheet, "all", 2, null);
        assert.strictEqual(secondAllPage.currentPage, 2);
        assert.strictEqual(secondAllPage.rows[0].rowNumber, 201);
        assert.strictEqual(secondAllPage.rows[secondAllPage.rows.length - 1].rowNumber, 205);
    });
});
