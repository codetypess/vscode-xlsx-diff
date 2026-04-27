/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { createCellKey, getCellAddress } from "../core/model/cells";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";

function createCell(rowNumber: number, columnNumber: number, value: string): CellSnapshot {
    return {
        key: createCellKey(rowNumber, columnNumber),
        rowNumber,
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        displayValue: value,
        formula: null,
        styleId: null,
    };
}

function createSheet(name: string, rowValues: string[]): SheetSnapshot {
    const cells: Record<string, CellSnapshot> = {};

    rowValues.forEach((value, index) => {
        if (!value) {
            return;
        }

        const cell = createCell(index + 1, 1, value);
        cells[cell.key] = cell;
    });

    return {
        name,
        rowCount: rowValues.length,
        columnCount: 1,
        mergedRanges: [],
        freezePane: null,
        cells,
        signature: `${name}:${rowValues.join("|")}`,
    };
}

function createGridSheet(name: string, rows: string[][]): SheetSnapshot {
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
        mergedRanges: [],
        freezePane: null,
        cells,
        signature: `${name}:${rows.map((row) => row.join("|")).join("/")}`,
    };
}

function createWorkbook(fileName: string, sheets: SheetSnapshot[]): WorkbookSnapshot {
    return {
        filePath: `/tmp/${fileName}`,
        fileName,
        fileSize: 0,
        modifiedTime: new Date("2026-01-01T00:00:00.000Z").toISOString(),
        sheets,
        isReadonly: false,
    };
}

suite("Workbook diff row alignment", () => {
    test("aligns rows after a right-side row insertion", () => {
        const diff = buildWorkbookDiff(
            createWorkbook("left.xlsx", [createSheet("Sheet1", ["A", "B", "C"])]),
            createWorkbook("right.xlsx", [createSheet("Sheet1", ["A", "X", "B", "C"])])
        );
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(
            sheet.alignedRows.map((row) => [row.rowNumber, row.leftRowNumber, row.rightRowNumber]),
            [
                [1, 1, 1],
                [2, null, 2],
                [3, 2, 3],
                [4, 3, 4],
            ]
        );
        assert.deepStrictEqual(sheet.diffRows, [2]);
        assert.deepStrictEqual(
            sheet.diffCells.map((cell) => [cell.rowNumber, cell.columnNumber]),
            [[2, 1]]
        );
    });

    test("keeps modified rows paired instead of treating them as delete plus insert", () => {
        const diff = buildWorkbookDiff(
            createWorkbook("left.xlsx", [createSheet("Sheet1", ["A", "B", "C"])]),
            createWorkbook("right.xlsx", [createSheet("Sheet1", ["A", "B2", "C"])])
        );
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(
            sheet.alignedRows.map((row) => [row.leftRowNumber, row.rightRowNumber]),
            [
                [1, 1],
                [2, 2],
                [3, 3],
            ]
        );
        assert.deepStrictEqual(sheet.diffRows, [2]);
        assert.deepStrictEqual(
            sheet.diffCells.map((cell) => [cell.rowNumber, cell.columnNumber]),
            [[2, 1]]
        );
    });

    test("ignores newline-style-only cell differences", () => {
        const diff = buildWorkbookDiff(
            createWorkbook("left.xlsx", [
                createSheet("Sheet1", [
                    "row1",
                    "row2",
                    "$&key1=ARMY==#army.id\n$&key1=ASSET==#assets.id",
                ]),
            ]),
            createWorkbook("right.xlsx", [
                createSheet("Sheet1", [
                    "row1",
                    "row2",
                    "$&key1=ARMY==#army.id\r\n$&key1=ASSET==#assets.id",
                ]),
            ])
        );
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(sheet.diffRows, []);
        assert.deepStrictEqual(sheet.diffCells, []);
        assert.strictEqual(diff.totalDiffCells, 0);
        assert.strictEqual(diff.totalDiffRows, 0);
        assert.strictEqual(diff.totalDiffSheets, 0);
    });

    test("keeps later rows aligned when an inserted row is followed by a modified row", () => {
        const diff = buildWorkbookDiff(
            createWorkbook("left.xlsx", [createSheet("Sheet1", ["A", "B", "C"])]),
            createWorkbook("right.xlsx", [createSheet("Sheet1", ["A", "X", "B*", "C"])])
        );
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(
            sheet.alignedRows.map((row) => [row.rowNumber, row.leftRowNumber, row.rightRowNumber]),
            [
                [1, 1, 1],
                [2, null, 2],
                [3, 2, 3],
                [4, 3, 4],
            ]
        );
        assert.deepStrictEqual(sheet.diffRows, [2, 3]);
        assert.deepStrictEqual(
            sheet.diffCells.map((cell) => [cell.rowNumber, cell.columnNumber]),
            [
                [2, 1],
                [3, 1],
            ]
        );
    });

    test("aligns columns after a right-side column insertion", () => {
        const diff = buildWorkbookDiff(
            createWorkbook("left.xlsx", [
                createGridSheet("Sheet1", [
                    ["ID", "Name", "Score"],
                    ["1", "Alice", "90"],
                ]),
            ]),
            createWorkbook("right.xlsx", [
                createGridSheet("Sheet1", [
                    ["ID", "Status", "Name", "Score"],
                    ["1", "New", "Alice", "90"],
                ]),
            ])
        );
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(
            sheet.alignedColumns.map((column) => [
                column.columnNumber,
                column.leftColumnNumber,
                column.rightColumnNumber,
            ]),
            [
                [1, 1, 1],
                [2, null, 2],
                [3, 2, 3],
                [4, 3, 4],
            ]
        );
        assert.deepStrictEqual(sheet.diffRows, [1, 2]);
        assert.deepStrictEqual(
            sheet.diffCells.map((cell) => [cell.rowNumber, cell.columnNumber]),
            [
                [1, 2],
                [2, 2],
            ]
        );
    });
});
