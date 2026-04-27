/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import { createCellKey } from "../core/model/cells";
import { fixtureRegressionCases } from "./fixture-regression-cases";
import { getTestFixturePath } from "./fixture-paths";

suite("XLSX fixture regressions", () => {
    for (const fixtureCase of fixtureRegressionCases) {
        test(`ignores ${fixtureCase.name} differences in real workbooks`, async () => {
            const basePath = getTestFixturePath("xlsx-regressions", fixtureCase.name, "base.xlsx");
            const headPath = getTestFixturePath("xlsx-regressions", fixtureCase.name, "head.xlsx");
            const baseSnapshot = await loadWorkbookSnapshot(basePath);
            const headSnapshot = await loadWorkbookSnapshot(headPath);
            const cellKey = createCellKey(
                fixtureCase.focusCellRowNumber,
                fixtureCase.focusCellColumnNumber
            );

            assert.deepStrictEqual(
                baseSnapshot.sheets.map((sheet) => sheet.name),
                ["define"]
            );
            assert.deepStrictEqual(
                headSnapshot.sheets.map((sheet) => sheet.name),
                ["define"]
            );
            assert.strictEqual(
                baseSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                fixtureCase.expectedBaseDisplayValue
            );
            assert.strictEqual(
                headSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                fixtureCase.expectedHeadDisplayValue
            );

            const diff = buildWorkbookDiff(baseSnapshot, headSnapshot);
            const sheet = diff.sheets[0]!;

            assert.deepStrictEqual(sheet.diffRows, []);
            assert.deepStrictEqual(sheet.diffCells, []);
            assert.strictEqual(diff.totalDiffCells, 0);
            assert.strictEqual(diff.totalDiffRows, 0);
            assert.strictEqual(diff.totalDiffSheets, 0);
        });
    }
});
