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
    const firstSheet = diff.sheets[0];

    assert.deepStrictEqual(firstSheet?.diffRows, []);
    assert.deepStrictEqual(firstSheet?.diffCells, []);
    assert.strictEqual(diff.totalDiffCells, 0);
    assert.strictEqual(diff.totalDiffRows, 0);
    assert.strictEqual(diff.totalDiffSheets, 0);
}

suite("SCM workbook regressions", () => {
    for (const fixtureCase of fixtureRegressionCases) {
        test(`loads git HEAD snapshots for ${fixtureCase.name} and ignores invisible diffs`, async function () {
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

                assert.strictEqual(headResourceSnapshot.isReadonly, true);
                assert.strictEqual(localSnapshot.isReadonly, false);
                assert.strictEqual(
                    headResourceSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedBaseDisplayValue
                );
                assert.strictEqual(
                    localSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedHeadDisplayValue
                );
                if (fixtureCase.expectStyleDifference) {
                    assert.notStrictEqual(
                        headResourceSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null,
                        localSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null
                    );
                }

                assertNoWorkbookDiff(buildWorkbookDiff(headResourceSnapshot, localSnapshot));
            } finally {
                await rm(workspace.tempDirectory, { recursive: true, force: true });
            }
        });

        test(`loads svn BASE snapshots for ${fixtureCase.name} and ignores invisible diffs`, async function () {
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

                assert.strictEqual(baseResourceSnapshot.isReadonly, true);
                assert.strictEqual(localSnapshot.isReadonly, false);
                assert.strictEqual(
                    baseResourceSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedBaseDisplayValue
                );
                assert.strictEqual(
                    localSnapshot.sheets[0]?.cells[cellKey]?.displayValue,
                    fixtureCase.expectedHeadDisplayValue
                );
                if (fixtureCase.expectStyleDifference) {
                    assert.notStrictEqual(
                        baseResourceSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null,
                        localSnapshot.sheets[0]?.cells[cellKey]?.styleId ?? null
                    );
                }

                assertNoWorkbookDiff(buildWorkbookDiff(baseResourceSnapshot, localSnapshot));
            } finally {
                await rm(workspace.tempDirectory, { recursive: true, force: true });
            }
        });
    }
});
