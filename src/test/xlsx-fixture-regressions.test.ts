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
            const expectedBaseSheetNames = fixtureCase.expectedBaseSheetNames ??
                fixtureCase.expectedSheetNames ?? [fixtureCase.sheetName];
            const expectedHeadSheetNames = fixtureCase.expectedHeadSheetNames ??
                fixtureCase.expectedSheetNames ?? [fixtureCase.sheetName];
            const baseSheet = baseSnapshot.sheets.find(
                (sheet) => sheet.name === fixtureCase.sheetName
            );
            const headSheet = headSnapshot.sheets.find(
                (sheet) => sheet.name === fixtureCase.sheetName
            );

            assert.deepStrictEqual(
                baseSnapshot.sheets.map((sheet) => sheet.name),
                expectedBaseSheetNames
            );
            assert.deepStrictEqual(
                headSnapshot.sheets.map((sheet) => sheet.name),
                expectedHeadSheetNames
            );
            assert.ok(baseSheet);
            assert.ok(headSheet);
            assert.strictEqual(
                baseSheet?.cells[cellKey]?.displayValue,
                fixtureCase.expectedBaseDisplayValue
            );
            assert.strictEqual(
                headSheet?.cells[cellKey]?.displayValue,
                fixtureCase.expectedHeadDisplayValue
            );
            assert.deepStrictEqual(
                baseSnapshot.definedNames,
                fixtureCase.expectedBaseDefinedNames ?? []
            );
            assert.deepStrictEqual(
                headSnapshot.definedNames,
                fixtureCase.expectedHeadDefinedNames ?? []
            );
            assert.strictEqual(
                baseSheet?.visibility,
                fixtureCase.expectedBaseVisibility ?? "visible"
            );
            assert.strictEqual(
                headSheet?.visibility,
                fixtureCase.expectedHeadVisibility ?? "visible"
            );
            assert.deepStrictEqual(
                baseSheet?.freezePane ?? null,
                fixtureCase.expectedBaseFreezePane ?? null
            );
            assert.deepStrictEqual(
                headSheet?.freezePane ?? null,
                fixtureCase.expectedHeadFreezePane ?? null
            );
            if (fixtureCase.expectStyleDifference) {
                assert.notStrictEqual(
                    baseSheet?.cells[cellKey]?.styleId ?? null,
                    headSheet?.cells[cellKey]?.styleId ?? null
                );
            }

            const diff = buildWorkbookDiff(baseSnapshot, headSnapshot);
            const sheet =
                diff.sheets.find(
                    (entry) =>
                        entry.leftSheetName === fixtureCase.sheetName ||
                        entry.rightSheetName === fixtureCase.sheetName
                ) ?? diff.sheets[0]!;

            assert.deepStrictEqual(sheet.diffRows, []);
            assert.deepStrictEqual(sheet.diffCells, []);
            assert.strictEqual(
                sheet.mergedRangesChanged,
                fixtureCase.expectedDiff.mergedRangesChanged
            );
            assert.strictEqual(sheet.freezePaneChanged, fixtureCase.expectedDiff.freezePaneChanged);
            assert.strictEqual(sheet.visibilityChanged, fixtureCase.expectedDiff.visibilityChanged);
            assert.strictEqual(sheet.sheetOrderChanged, fixtureCase.expectedDiff.sheetOrderChanged);
            assert.strictEqual(
                diff.definedNamesChanged,
                fixtureCase.expectedDiff.definedNamesChanged
            );
            assert.strictEqual(diff.totalDiffCells, fixtureCase.expectedDiff.totalDiffCells);
            assert.strictEqual(diff.totalDiffRows, fixtureCase.expectedDiff.totalDiffRows);
            assert.strictEqual(diff.totalDiffSheets, fixtureCase.expectedDiff.totalDiffSheets);
        });
    }
});
