/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import { createCellKey } from "../core/model/cells";
import { getTestFixturePath } from "./fixture-paths";

const newlineOnlyFixtureDirectory = ["xlsx-regressions", "newline-only-cell-diff"];
const lfValue = "$&key1=ARMY==#army.id\n$&key1=ASSET==#assets.id";
const crlfValue = "$&key1=ARMY==#army.id\r\n$&key1=ASSET==#assets.id";

suite("XLSX fixture regressions", () => {
    test("ignores newline-style-only cell differences in real workbooks", async () => {
        const basePath = getTestFixturePath(...newlineOnlyFixtureDirectory, "base.xlsx");
        const headPath = getTestFixturePath(...newlineOnlyFixtureDirectory, "head.xlsx");
        const baseSnapshot = await loadWorkbookSnapshot(basePath);
        const headSnapshot = await loadWorkbookSnapshot(headPath);
        const cellKey = createCellKey(5, 6);

        assert.deepStrictEqual(
            baseSnapshot.sheets.map((sheet) => sheet.name),
            ["define"]
        );
        assert.deepStrictEqual(
            headSnapshot.sheets.map((sheet) => sheet.name),
            ["define"]
        );
        assert.strictEqual(baseSnapshot.sheets[0]?.cells[cellKey]?.displayValue, lfValue);
        assert.strictEqual(headSnapshot.sheets[0]?.cells[cellKey]?.displayValue, crlfValue);

        const diff = buildWorkbookDiff(baseSnapshot, headSnapshot);
        const sheet = diff.sheets[0]!;

        assert.deepStrictEqual(sheet.diffRows, []);
        assert.deepStrictEqual(sheet.diffCells, []);
        assert.strictEqual(diff.totalDiffCells, 0);
        assert.strictEqual(diff.totalDiffRows, 0);
        assert.strictEqual(diff.totalDiffSheets, 0);
    });
});
