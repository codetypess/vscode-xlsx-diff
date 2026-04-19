import * as assert from "assert";
import * as vscode from "vscode";
import {
    describeGitResourceRef,
    getScmWorkbookDiffUrisFromTabInput,
    getWorkbookDiffUrisFromTabInput,
    getWorkbookResourcePathLabel,
    getWorkbookResourceTimeLabel,
} from "../workbook/resourceUri";

suite("Workbook resource URIs", () => {
    test("recognizes xlsx diff tab inputs", () => {
        const original = vscode.Uri.file("/tmp/before.xlsx");
        const modified = vscode.Uri.file("/tmp/after.xlsx");
        const input = new vscode.TabInputTextDiff(original, modified);

        assert.deepStrictEqual(getWorkbookDiffUrisFromTabInput(input), {
            original,
            modified,
        });
    });

    test("ignores non-xlsx diff tab inputs", () => {
        const input = new vscode.TabInputTextDiff(
            vscode.Uri.file("/tmp/before.txt"),
            vscode.Uri.file("/tmp/after.txt")
        );

        assert.strictEqual(getWorkbookDiffUrisFromTabInput(input), undefined);
    });

    test("extracts git ref labels for readonly workbook resources", () => {
        const gitUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "HEAD",
            }),
        });

        assert.strictEqual(getWorkbookResourcePathLabel(gitUri), "/tmp/item.xlsx (HEAD)");
        assert.match(getWorkbookResourceTimeLabel(gitUri) ?? "", /HEAD$/);
    });

    test("filters scm workbook diffs to non-file originals", () => {
        const scmInput = new vscode.TabInputTextDiff(
            vscode.Uri.from({
                scheme: "git",
                path: "/tmp/item.xlsx",
                query: JSON.stringify({
                    path: "/tmp/item.xlsx",
                    ref: "HEAD",
                }),
            }),
            vscode.Uri.file("/tmp/item.xlsx")
        );

        const fileDiffInput = new vscode.TabInputTextDiff(
            vscode.Uri.file("/tmp/left.xlsx"),
            vscode.Uri.file("/tmp/right.xlsx")
        );

        assert.ok(getScmWorkbookDiffUrisFromTabInput(scmInput));
        assert.strictEqual(getScmWorkbookDiffUrisFromTabInput(fileDiffInput), undefined);
    });

    test("describes git refs for commit and index-backed resources", () => {
        const commitRef = describeGitResourceRef("HEAD", {
            resolvedCommit: "d44224e",
        });
        assert.strictEqual(commitRef.value, "d44224e");
        assert.ok(["Commit", "提交"].includes(commitRef.label));

        const indexRef = describeGitResourceRef("~", {
            resolvedCommit: "d44224e",
            hasStagedChanges: true,
        });
        assert.match(indexRef.value, /d44224e$/);
        assert.ok(["Source", "来源"].includes(indexRef.label));

        const stageRef = describeGitResourceRef("~2");
        assert.match(stageRef.value, /2$/);
        assert.ok(["Source", "来源"].includes(stageRef.label));
    });
});
