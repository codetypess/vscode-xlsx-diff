/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import {
    canCreateEditorFilterRange,
    clearEditorFilterColumn,
    createEditorSheetFilterState,
    getEditorFilterColumnValues,
    getEditorVisibleRows,
    resolveEditorFilterRangeFromActiveCell,
    resolveEditorFilterRangeFromSelection,
    toggleEditorSheetFilterState,
    updateEditorFilterIncludedValues,
    updateEditorFilterSort,
    type EditorFilterCellSource,
} from "../webview/editor-panel/editor-panel-filter";

function createSource(rows: string[][]): EditorFilterCellSource {
    const cells = new Map<string, string>();
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex] ?? [];
        for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
            cells.set(`${rowIndex + 1}:${columnIndex + 1}`, row[columnIndex] ?? "");
        }
    }

    return {
        rowCount: rows.length,
        columnCount: Math.max(...rows.map((row) => row.length), 1),
        getCellValue: (rowNumber, columnNumber) => cells.get(`${rowNumber}:${columnNumber}`) ?? "",
    };
}

suite("Editor panel filter helpers", () => {
    test("requires a range with a header row and at least one data row", () => {
        const source = createSource([
            ["h1", "h2"],
            ["a", "b"],
        ]);

        assert.strictEqual(
            canCreateEditorFilterRange(source, {
                startRow: 1,
                endRow: 1,
                startColumn: 1,
                endColumn: 2,
            }),
            false
        );
        assert.strictEqual(
            canCreateEditorFilterRange(source, {
                startRow: 1,
                endRow: 2,
                startColumn: 1,
                endColumn: 2,
            }),
            true
        );
    });

    test("expands a single selected header row to the sheet end", () => {
        const source = createSource([
            ["note"],
            ["header-a", "header-b", "header-c"],
            ["1", "2", "3"],
            ["4", "5", "6"],
        ]);

        assert.deepStrictEqual(
            resolveEditorFilterRangeFromSelection(source, {
                startRow: 2,
                endRow: 2,
                startColumn: 1,
                endColumn: 3,
            }),
            {
                startRow: 2,
                endRow: 4,
                startColumn: 1,
                endColumn: 3,
            }
        );
    });

    test("infers a filter range from the active cell inside a data block", () => {
        const source = createSource([
            ["title"],
            ["", "", ""],
            ["header-a", "header-b", "header-c"],
            ["1", "", "3"],
            ["4", "5", "6"],
            ["", "", ""],
        ]);

        assert.deepStrictEqual(
            resolveEditorFilterRangeFromActiveCell(source, {
                rowNumber: 4,
                columnNumber: 2,
            }),
            {
                startRow: 3,
                endRow: 5,
                startColumn: 1,
                endColumn: 3,
            }
        );
    });

    test("does not infer a filter range across a blank boundary column", () => {
        const source = createSource([
            ["header-a", "header-b", "", ""],
            ["1", "2", "", ""],
            ["3", "4", "", ""],
        ]);

        assert.strictEqual(
            resolveEditorFilterRangeFromActiveCell(source, {
                rowNumber: 2,
                columnNumber: 3,
            }),
            null
        );
    });

    test("lists distinct values from data rows in filter order", () => {
        const source = createSource([["header"], [""], ["beta"], ["alpha"], ["beta"]]);
        const filterState = createEditorSheetFilterState(source, {
            startRow: 1,
            endRow: 5,
            startColumn: 1,
            endColumn: 1,
        });

        assert.deepStrictEqual(getEditorFilterColumnValues(source, filterState!, 1), [
            { value: "alpha", count: 1 },
            { value: "beta", count: 2 },
            { value: "", count: 1 },
        ]);
    });

    test("filters data rows inside the active range and keeps surrounding rows intact", () => {
        const source = createSource([
            ["before"],
            ["header"],
            ["red"],
            ["blue"],
            ["red"],
            ["after"],
        ]);
        let filterState = createEditorSheetFilterState(source, {
            startRow: 2,
            endRow: 5,
            startColumn: 1,
            endColumn: 1,
        });
        filterState = updateEditorFilterIncludedValues(filterState!, 1, ["red"]);

        assert.deepStrictEqual(getEditorVisibleRows(source, filterState), {
            visibleRows: [1, 2, 3, 5, 6],
            hiddenRows: [4],
        });
    });

    test("sorts the visible data rows inside the active range", () => {
        const source = createSource([["header"], ["20"], ["3"], ["11"]]);
        let filterState = createEditorSheetFilterState(source, {
            startRow: 1,
            endRow: 4,
            startColumn: 1,
            endColumn: 1,
        });
        filterState = updateEditorFilterSort(filterState!, 1, "asc");

        assert.deepStrictEqual(getEditorVisibleRows(source, filterState).visibleRows, [1, 3, 4, 2]);
    });

    test("clears filter values and sort state for one column", () => {
        const source = createSource([
            ["headerA", "headerB"],
            ["red", "z"],
            ["blue", "a"],
        ]);
        let filterState = createEditorSheetFilterState(source, {
            startRow: 1,
            endRow: 3,
            startColumn: 1,
            endColumn: 2,
        });
        filterState = updateEditorFilterIncludedValues(filterState!, 1, ["red"]);
        filterState = updateEditorFilterSort(filterState, 1, "desc");
        filterState = clearEditorFilterColumn(filterState, 1);

        assert.deepStrictEqual(filterState.includedValuesByColumn, {});
        assert.strictEqual(filterState.sort, null);
    });

    test("toggle clears an active filter for a whole-row selection", () => {
        const source = createSource([
            ["headerA", "headerB"],
            ["red", "z"],
            ["blue", "a"],
            ["green", "m"],
        ]);
        const activeFilterState = createEditorSheetFilterState(source, {
            startRow: 1,
            endRow: 4,
            startColumn: 1,
            endColumn: 2,
        });

        assert.strictEqual(
            toggleEditorSheetFilterState(source, activeFilterState, {
                startRow: 2,
                endRow: 2,
                startColumn: 1,
                endColumn: 2,
            }),
            null
        );
    });
});
