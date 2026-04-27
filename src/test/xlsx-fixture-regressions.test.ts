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
        test(`handles ${fixtureCase.name} real workbooks`, async () => {
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
            assert.deepStrictEqual(
                baseSnapshot.sheets[0]?.freezePane ?? null,
                fixtureCase.expectedBaseFreezePane ?? null
            );
            assert.deepStrictEqual(
                headSnapshot.sheets[0]?.freezePane ?? null,
                fixtureCase.expectedHeadFreezePane ?? null
            );
            if (fixtureCase.expectStyleDifference) {
                assert.notStrictEqual(
                    baseSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null,
                    headSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null
                );
            }

            const diff = buildWorkbookDiff(baseSnapshot, headSnapshot);
            const sheet = diff.sheets[0]!;

            assert.deepStrictEqual(sheet.diffRows, []);
            assert.deepStrictEqual(sheet.diffCells, []);
            assert.strictEqual(
                sheet.mergedRangesChanged,
                fixtureCase.expectedDiff.mergedRangesChanged
            );
            assert.strictEqual(sheet.freezePaneChanged, fixtureCase.expectedDiff.freezePaneChanged);
            assert.strictEqual(diff.totalDiffCells, fixtureCase.expectedDiff.totalDiffCells);
            assert.strictEqual(diff.totalDiffRows, fixtureCase.expectedDiff.totalDiffRows);
            assert.strictEqual(diff.totalDiffSheets, fixtureCase.expectedDiff.totalDiffSheets);
        });
    }
});
