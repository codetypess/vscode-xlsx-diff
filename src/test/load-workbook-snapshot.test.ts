import * as assert from "assert";
import * as vscode from "vscode";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";

function createSvnGraphEmptyUri(label = "repo/item.xlsx (deleted)") {
    const params = new URLSearchParams({
        label,
        source: "empty",
    });

    return vscode.Uri.from({
        scheme: "svn-graph",
        path: label.startsWith("/") ? label : `/${label}`,
        query: params.toString(),
    });
}

suite("Load workbook snapshot", () => {
    test("creates empty snapshots for svn-graph empty workbook resources", async () => {
        const uri = createSvnGraphEmptyUri();

        const snapshot = await loadWorkbookSnapshot(uri);

        assert.strictEqual(snapshot.fileName, "item.xlsx");
        assert.strictEqual(snapshot.filePath, "repo/item.xlsx");
        assert.strictEqual(snapshot.fileSize, 0);
        assert.strictEqual(snapshot.isReadonly, true);
        assert.deepStrictEqual(snapshot.sheets, []);
        assert.ok(["Empty workbook", "空工作簿"].includes(snapshot.modifiedTimeLabel ?? ""));
    });
});
