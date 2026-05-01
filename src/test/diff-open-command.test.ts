/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import { runCompareActiveWith } from "../commands/compare-active-with";
import { runCompareTwoFiles } from "../commands/compare-two-files";

suite("Diff open commands", () => {
    test("compareTwoFiles opens the diff panel after both workbooks are selected", async () => {
        const extensionUri = vscode.Uri.file("/tmp/extension");
        const leftUri = vscode.Uri.file("/tmp/left.xlsx");
        const rightUri = vscode.Uri.file("/tmp/right.xlsx");
        const openCalls: Array<{
            extensionUri: vscode.Uri;
            leftUri: vscode.Uri;
            rightUri: vscode.Uri;
        }> = [];
        const pickCalls: Array<{ openLabel: string; seedUri?: vscode.Uri }> = [];

        await runCompareTwoFiles(extensionUri, {
            pickWorkbook: async (openLabel, seedUri) => {
                pickCalls.push({ openLabel, seedUri });
                return pickCalls.length === 1 ? leftUri : rightUri;
            },
            openDiffPanel: async (receivedExtensionUri, receivedLeftUri, receivedRightUri) => {
                openCalls.push({
                    extensionUri: receivedExtensionUri,
                    leftUri: receivedLeftUri,
                    rightUri: receivedRightUri,
                });
            },
        });

        assert.strictEqual(pickCalls.length, 2);
        assert.strictEqual(pickCalls[0]?.seedUri, undefined);
        assert.strictEqual(pickCalls[1]?.seedUri?.fsPath, leftUri.fsPath);
        assert.deepStrictEqual(openCalls, [
            {
                extensionUri,
                leftUri,
                rightUri,
            },
        ]);
    });

    test("compareTwoFiles does not open when the second picker is cancelled", async () => {
        const extensionUri = vscode.Uri.file("/tmp/extension");
        const leftUri = vscode.Uri.file("/tmp/left.xlsx");
        let openCount = 0;

        await runCompareTwoFiles(extensionUri, {
            pickWorkbook: async (_openLabel, seedUri) => (seedUri ? undefined : leftUri),
            openDiffPanel: async () => {
                openCount += 1;
            },
        });

        assert.strictEqual(openCount, 0);
    });

    test("compareActiveWith prefers the command resource and opens the diff panel", async () => {
        const extensionUri = vscode.Uri.file("/tmp/extension");
        const resourceUri = vscode.Uri.file("/tmp/resource.xlsx");
        const activeUri = vscode.Uri.file("/tmp/active.xlsx");
        const targetUri = vscode.Uri.file("/tmp/target.xlsx");
        const openCalls: Array<{ leftUri: vscode.Uri; rightUri: vscode.Uri }> = [];

        await runCompareActiveWith(
            extensionUri,
            { resourceUri },
            {
                getWorkbookUriFromCommandArg: (value) =>
                    value && typeof value === "object" && "resourceUri" in value
                        ? resourceUri
                        : undefined,
                getActiveWorkbookUri: () => activeUri,
                pickWorkbook: async (_openLabel, seedUri) => {
                    assert.strictEqual(seedUri?.fsPath, resourceUri.fsPath);
                    return targetUri;
                },
                showErrorMessage: async () => undefined,
                openDiffPanel: async (_receivedExtensionUri, leftUri, rightUri) => {
                    openCalls.push({ leftUri, rightUri });
                },
            }
        );

        assert.deepStrictEqual(openCalls, [
            {
                leftUri: resourceUri,
                rightUri: targetUri,
            },
        ]);
    });

    test("compareActiveWith shows an error when no workbook is available", async () => {
        const extensionUri = vscode.Uri.file("/tmp/extension");
        const errorMessages: string[] = [];
        let openCount = 0;

        await runCompareActiveWith(extensionUri, undefined, {
            getWorkbookUriFromCommandArg: () => undefined,
            getActiveWorkbookUri: () => undefined,
            pickWorkbook: async () => undefined,
            showErrorMessage: async (message: string) => {
                errorMessages.push(message);
                return undefined;
            },
            openDiffPanel: async () => {
                openCount += 1;
            },
        });

        assert.strictEqual(errorMessages.length, 1);
        assert.strictEqual(openCount, 0);
    });
});
