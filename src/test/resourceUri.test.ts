import * as assert from "assert";
import * as vscode from "vscode";
import { getGitWorkbookResourceInfo } from "../git/resourceInfo";
import { getSvnWorkbookResourceInfo } from "../svn/resourceInfo";
import {
    describeGitResourceRef,
    getScmWorkbookDiffUrisFromEditorUris,
    getScmWorkbookDiffUrisFromTabInput,
    getWorkbookDiffUrisFromTabInput,
    getWorkbookResourceDetail,
    getWorkbookResourcePathLabel,
    getWorkbookResourceTimeLabel,
    isWorkbookResourceUri,
} from "../workbook/resourceUri";

function createSvnShowUri(resourcePath = "/tmp/item.xlsx", ref: string | number = "HEAD") {
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

    test("recognizes git virtual workbook resources with rewritten paths", () => {
        const gitUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx.git",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "",
            }),
        });

        assert.strictEqual(isWorkbookResourceUri(gitUri), true);
        assert.strictEqual(getWorkbookResourcePathLabel(gitUri), "/tmp/item.xlsx");
    });

    test("extracts git adapter resource info without changing rewritten paths", () => {
        const gitUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx.git",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "",
            }),
        });

        const info = getGitWorkbookResourceInfo(gitUri);
        assert.strictEqual(info?.provider, "git");
        assert.strictEqual(info?.uri, gitUri);
        assert.strictEqual(info?.resourcePath, "/tmp/item.xlsx");
        assert.strictEqual(info?.ref, "");
    });

    test("recognizes svn workbook resources from svn-scm show uris", async () => {
        const svnUri = createSvnShowUri();

        const info = getSvnWorkbookResourceInfo(svnUri);
        assert.strictEqual(info?.provider, "svn");
        assert.strictEqual(info?.uri, svnUri);
        assert.strictEqual(info?.resourcePath, "/tmp/item.xlsx");
        assert.strictEqual(info?.ref, "HEAD");
        assert.strictEqual(isWorkbookResourceUri(svnUri), true);
        assert.strictEqual(getWorkbookResourcePathLabel(svnUri), "/tmp/item.xlsx (HEAD)");
        assert.match(getWorkbookResourceTimeLabel(svnUri) ?? "", /SVN .*HEAD$/);

        const detail = await getWorkbookResourceDetail(svnUri);
        assert.match(detail?.value ?? "", /SVN .*HEAD$/);
        assert.ok(["Source", "来源"].includes(detail?.label ?? ""));
    });

    test("ignores unsupported svn virtual resources", () => {
        const svnPatchUri = vscode.Uri.from({
            scheme: "svn",
            path: "/tmp/item.xlsx",
            query: JSON.stringify({
                action: "PATCH",
                fsPath: "/tmp/item.xlsx",
                extra: {},
            }),
        });

        assert.strictEqual(getSvnWorkbookResourceInfo(svnPatchUri), undefined);
        assert.strictEqual(isWorkbookResourceUri(svnPatchUri), false);
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

        const svnInput = new vscode.TabInputTextDiff(
            createSvnShowUri(),
            vscode.Uri.file("/tmp/item.xlsx")
        );

        const fileDiffInput = new vscode.TabInputTextDiff(
            vscode.Uri.file("/tmp/left.xlsx"),
            vscode.Uri.file("/tmp/right.xlsx")
        );

        assert.ok(getScmWorkbookDiffUrisFromTabInput(scmInput));
        assert.ok(getScmWorkbookDiffUrisFromTabInput(svnInput));
        assert.strictEqual(getScmWorkbookDiffUrisFromTabInput(fileDiffInput), undefined);
    });

    test("normalizes svn local diffs from HEAD to BASE", () => {
        const svnHeadUri = createSvnShowUri("/tmp/item.xlsx", "HEAD");
        const fileUri = vscode.Uri.file("/tmp/item.xlsx");
        const diffInput = new vscode.TabInputTextDiff(svnHeadUri, fileUri);

        const tabDiffUris = getScmWorkbookDiffUrisFromTabInput(diffInput);
        const editorDiffUris = getScmWorkbookDiffUrisFromEditorUris(fileUri, svnHeadUri);

        assert.strictEqual(getSvnWorkbookResourceInfo(tabDiffUris?.original ?? fileUri)?.ref, "BASE");
        assert.strictEqual(tabDiffUris?.modified.toString(), fileUri.toString());
        assert.strictEqual(
            getSvnWorkbookResourceInfo(editorDiffUris?.original ?? fileUri)?.ref,
            "BASE"
        );
        assert.strictEqual(editorDiffUris?.modified.toString(), fileUri.toString());
    });

    test("pairs scm custom editor resources for the same workbook", () => {
        const gitUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "~",
            }),
        });
        const fileUri = vscode.Uri.file("/tmp/item.xlsx");

        assert.deepStrictEqual(getScmWorkbookDiffUrisFromEditorUris(fileUri, gitUri), {
            original: gitUri,
            modified: fileUri,
        });
    });

    test("pairs svn custom editor resources for the same workbook", () => {
        const svnUri = createSvnShowUri();
        const fileUri = vscode.Uri.file("/tmp/item.xlsx");

        const diffUris = getScmWorkbookDiffUrisFromEditorUris(fileUri, svnUri);
        assert.strictEqual(getSvnWorkbookResourceInfo(diffUris?.original ?? fileUri)?.ref, "BASE");
        assert.strictEqual(diffUris?.modified.toString(), fileUri.toString());
    });

    test("pairs scm custom editor resources for git index diffs", () => {
        const headUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "HEAD",
            }),
        });
        const indexUri = vscode.Uri.from({
            scheme: "git",
            path: "/tmp/item.xlsx.git",
            query: JSON.stringify({
                path: "/tmp/item.xlsx",
                ref: "",
            }),
        });

        assert.deepStrictEqual(getScmWorkbookDiffUrisFromEditorUris(indexUri, headUri), {
            original: headUri,
            modified: indexUri,
        });
    });

    test("ignores non-scm custom editor resource pairs", () => {
        assert.strictEqual(
            getScmWorkbookDiffUrisFromEditorUris(
                vscode.Uri.file("/tmp/item.xlsx"),
                vscode.Uri.file("/tmp/item.xlsx")
            ),
            undefined
        );
        assert.strictEqual(
            getScmWorkbookDiffUrisFromEditorUris(
                vscode.Uri.from({
                    scheme: "git",
                    path: "/tmp/left.xlsx",
                    query: JSON.stringify({
                        path: "/tmp/left.xlsx",
                        ref: "~",
                    }),
                }),
                vscode.Uri.file("/tmp/right.xlsx")
            ),
            undefined
        );
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
