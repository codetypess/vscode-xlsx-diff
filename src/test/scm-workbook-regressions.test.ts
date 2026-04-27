/// <reference types="mocha" />
/// <reference types="node" />

import { execFile as execFileCallback } from "node:child_process";
import { copyFile, mkdir, mkdtemp, rm } from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { promisify } from "node:util";
import * as assert from "assert";
import * as vscode from "vscode";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import { createCellKey } from "../core/model/cells";
import { fixtureRegressionCases } from "./fixture-regression-cases";
import { getTestFixturePath } from "./fixture-paths";

const execFile = promisify(execFileCallback);

function createGitUri(resourcePath: string, ref = "HEAD") {
    return vscode.Uri.from({
        scheme: "git",
        path: resourcePath,
        query: JSON.stringify({
            path: resourcePath,
            ref,
        }),
    });
}

function createSvnShowUri(resourcePath: string, ref: string | number = "BASE") {
    return vscode.Uri.from({
        scheme: "svn",
        path: resourcePath,
        query: JSON.stringify({
            action: "SHOW",
            fsPath: resourcePath,
            extra: {
                ref,
            },
        }),
    });
}

async function createTempGitWorkspace(
    baseWorkbookPath: string,
    headWorkbookPath: string
): Promise<{
    tempDirectory: string;
    workbookPath: string;
}> {
    const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-git-fixture-"));
    const workbookPath = path.join(tempDirectory, "asset.xlsx");

    await copyFile(baseWorkbookPath, workbookPath);
    await execFile("git", ["init"], { cwd: tempDirectory });
    await execFile("git", ["config", "user.name", "Alice Example"], {
        cwd: tempDirectory,
    });
    await execFile("git", ["config", "user.email", "alice@example.com"], {
        cwd: tempDirectory,
    });
    await execFile("git", ["add", "asset.xlsx"], { cwd: tempDirectory });
    await execFile("git", ["commit", "-m", "Add base workbook"], {
        cwd: tempDirectory,
    });
    await copyFile(headWorkbookPath, workbookPath);

    return {
        tempDirectory,
        workbookPath,
    };
}

async function createTempSvnWorkspace(
    baseWorkbookPath: string,
    headWorkbookPath: string
): Promise<{
    tempDirectory: string;
    workbookPath: string;
    repositoryUrl: string;
}> {
    const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-svn-fixture-"));
    const repositoryPath = path.join(tempDirectory, "repo");
    const repositoryUrl = `file://${repositoryPath}`;
    const workingCopyPath = path.join(tempDirectory, "wc");
    const workbookPath = path.join(workingCopyPath, "design", "asset.xlsx");

    await execFile("svnadmin", ["create", repositoryPath]);
    await execFile("svn", ["checkout", repositoryUrl, workingCopyPath]);
    await mkdir(path.dirname(workbookPath), { recursive: true });
    await copyFile(baseWorkbookPath, workbookPath);
    await execFile("svn", ["add", path.join(workingCopyPath, "design")], {
        cwd: workingCopyPath,
    });
    await execFile(
        "svn",
        ["commit", workingCopyPath, "-m", "Add base workbook", "--username", "alice"],
        {
            cwd: tempDirectory,
        }
    );
    await copyFile(headWorkbookPath, workbookPath);

    return {
        tempDirectory,
        workbookPath,
        repositoryUrl,
    };
}

function assertNoWorkbookDiff(diff: ReturnType<typeof buildWorkbookDiff>) {
    const firstSheet = diff.sheets[0]!;
    return firstSheet;
}

suite("SCM workbook regressions", () => {
    for (const fixtureCase of fixtureRegressionCases) {
        test(`loads git HEAD snapshots for ${fixtureCase.name} and keeps expected diff semantics`, async function () {
            this.timeout(10000);

            const baseWorkbookPath = getTestFixturePath(
                "xlsx-regressions",
                fixtureCase.name,
                "base.xlsx"
            );
            const headWorkbookPath = getTestFixturePath(
                "xlsx-regressions",
                fixtureCase.name,
                "head.xlsx"
            );
            const workspace = await createTempGitWorkspace(baseWorkbookPath, headWorkbookPath);

            try {
                const headResourceSnapshot = await loadWorkbookSnapshot(
                    createGitUri(workspace.workbookPath, "HEAD")
                );
                const localSnapshot = await loadWorkbookSnapshot(
                    vscode.Uri.file(workspace.workbookPath)
                );
                const cellKey = createCellKey(
                    fixtureCase.focusCellRowNumber,
                    fixtureCase.focusCellColumnNumber
                );
                const expectedSheetNames = fixtureCase.expectedSheetNames ?? [
                    fixtureCase.sheetName,
                ];
                const baseSheet = headResourceSnapshot.sheets.find(
                    (sheet) => sheet.name === fixtureCase.sheetName
                );
                const localSheet = localSnapshot.sheets.find(
                    (sheet) => sheet.name === fixtureCase.sheetName
                );

                assert.strictEqual(headResourceSnapshot.isReadonly, true);
                assert.strictEqual(localSnapshot.isReadonly, false);
                assert.deepStrictEqual(
                    headResourceSnapshot.sheets.map((sheet) => sheet.name),
                    expectedSheetNames
                );
                assert.deepStrictEqual(
                    localSnapshot.sheets.map((sheet) => sheet.name),
                    expectedSheetNames
                );
                assert.ok(baseSheet);
                assert.ok(localSheet);
                assert.strictEqual(
                    baseSheet?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedBaseDisplayValue
                );
                assert.strictEqual(
                    localSheet?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedHeadDisplayValue
                );
                assert.strictEqual(
                    baseSheet?.visibility,
                    fixtureCase.expectedBaseVisibility ?? "visible"
                );
                assert.strictEqual(
                    localSheet?.visibility,
                    fixtureCase.expectedHeadVisibility ?? "visible"
                );
                assert.deepStrictEqual(
                    baseSheet?.freezePane ?? null,
                    fixtureCase.expectedBaseFreezePane ?? null
                );
                assert.deepStrictEqual(
                    localSheet?.freezePane ?? null,
                    fixtureCase.expectedHeadFreezePane ?? null
                );
                if (fixtureCase.expectStyleDifference) {
                    assert.notStrictEqual(
                        baseSheet?.cells[cellKey]?.styleId ?? null,
                        localSheet?.cells[cellKey]?.styleId ?? null
                    );
                }

                const diff = buildWorkbookDiff(headResourceSnapshot, localSnapshot);
                const sheet =
                    diff.sheets.find(
                        (entry) =>
                            entry.leftSheetName === fixtureCase.sheetName ||
                            entry.rightSheetName === fixtureCase.sheetName
                    ) ?? assertNoWorkbookDiff(diff);
                assert.deepStrictEqual(sheet.diffRows, []);
                assert.deepStrictEqual(sheet.diffCells, []);
                assert.strictEqual(
                    sheet.mergedRangesChanged,
                    fixtureCase.expectedDiff.mergedRangesChanged
                );
                assert.strictEqual(
                    sheet.freezePaneChanged,
                    fixtureCase.expectedDiff.freezePaneChanged
                );
                assert.strictEqual(
                    sheet.visibilityChanged,
                    fixtureCase.expectedDiff.visibilityChanged
                );
                assert.strictEqual(diff.totalDiffCells, fixtureCase.expectedDiff.totalDiffCells);
                assert.strictEqual(diff.totalDiffRows, fixtureCase.expectedDiff.totalDiffRows);
                assert.strictEqual(diff.totalDiffSheets, fixtureCase.expectedDiff.totalDiffSheets);
            } finally {
                await rm(workspace.tempDirectory, { recursive: true, force: true });
            }
        });

        test(`loads svn BASE snapshots for ${fixtureCase.name} and keeps expected diff semantics`, async function () {
            this.timeout(10000);

            const baseWorkbookPath = getTestFixturePath(
                "xlsx-regressions",
                fixtureCase.name,
                "base.xlsx"
            );
            const headWorkbookPath = getTestFixturePath(
                "xlsx-regressions",
                fixtureCase.name,
                "head.xlsx"
            );
            const workspace = await createTempSvnWorkspace(baseWorkbookPath, headWorkbookPath);

            try {
                const baseResourceSnapshot = await loadWorkbookSnapshot(
                    createSvnShowUri(workspace.workbookPath, "BASE")
                );
                const localSnapshot = await loadWorkbookSnapshot(
                    vscode.Uri.file(workspace.workbookPath)
                );
                const cellKey = createCellKey(
                    fixtureCase.focusCellRowNumber,
                    fixtureCase.focusCellColumnNumber
                );
                const expectedSheetNames = fixtureCase.expectedSheetNames ?? [
                    fixtureCase.sheetName,
                ];
                const baseSheet = baseResourceSnapshot.sheets.find(
                    (sheet) => sheet.name === fixtureCase.sheetName
                );
                const localSheet = localSnapshot.sheets.find(
                    (sheet) => sheet.name === fixtureCase.sheetName
                );

                assert.strictEqual(baseResourceSnapshot.isReadonly, true);
                assert.strictEqual(localSnapshot.isReadonly, false);
                assert.deepStrictEqual(
                    baseResourceSnapshot.sheets.map((sheet) => sheet.name),
                    expectedSheetNames
                );
                assert.deepStrictEqual(
                    localSnapshot.sheets.map((sheet) => sheet.name),
                    expectedSheetNames
                );
                assert.ok(baseSheet);
                assert.ok(localSheet);
                assert.strictEqual(
                    baseSheet?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedBaseDisplayValue
                );
                assert.strictEqual(
                    localSheet?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedHeadDisplayValue
                );
                assert.strictEqual(
                    baseSheet?.visibility,
                    fixtureCase.expectedBaseVisibility ?? "visible"
                );
                assert.strictEqual(
                    localSheet?.visibility,
                    fixtureCase.expectedHeadVisibility ?? "visible"
                );
                assert.deepStrictEqual(
                    baseSheet?.freezePane ?? null,
                    fixtureCase.expectedBaseFreezePane ?? null
                );
                assert.deepStrictEqual(
                    localSheet?.freezePane ?? null,
                    fixtureCase.expectedHeadFreezePane ?? null
                );
                if (fixtureCase.expectStyleDifference) {
                    assert.notStrictEqual(
                        baseSheet?.cells[cellKey]?.styleId ?? null,
                        localSheet?.cells[cellKey]?.styleId ?? null
                    );
                }

                const diff = buildWorkbookDiff(baseResourceSnapshot, localSnapshot);
                const sheet =
                    diff.sheets.find(
                        (entry) =>
                            entry.leftSheetName === fixtureCase.sheetName ||
                            entry.rightSheetName === fixtureCase.sheetName
                    ) ?? assertNoWorkbookDiff(diff);
                assert.deepStrictEqual(sheet.diffRows, []);
                assert.deepStrictEqual(sheet.diffCells, []);
                assert.strictEqual(
                    sheet.mergedRangesChanged,
                    fixtureCase.expectedDiff.mergedRangesChanged
                );
                assert.strictEqual(
                    sheet.freezePaneChanged,
                    fixtureCase.expectedDiff.freezePaneChanged
                );
                assert.strictEqual(
                    sheet.visibilityChanged,
                    fixtureCase.expectedDiff.visibilityChanged
                );
                assert.strictEqual(diff.totalDiffCells, fixtureCase.expectedDiff.totalDiffCells);
                assert.strictEqual(diff.totalDiffRows, fixtureCase.expectedDiff.totalDiffRows);
                assert.strictEqual(diff.totalDiffSheets, fixtureCase.expectedDiff.totalDiffSheets);
            } finally {
                await rm(workspace.tempDirectory, { recursive: true, force: true });
            }
        });
    }
});
