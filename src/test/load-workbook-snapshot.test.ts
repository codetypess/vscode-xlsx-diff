import * as assert from "assert";
import { mkdtemp, rm } from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import * as vscode from "vscode";
import { Workbook } from "../core/fastxlsx/runtime";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";

function createSvnTreeEmptyUri(label = "repo/item.xlsx (deleted)") {
    const params = new URLSearchParams({
        label,
        source: "empty",
    });

    return vscode.Uri.from({
        scheme: "svn-tree",
        path: label.startsWith("/") ? label : `/${label}`,
        query: params.toString(),
    });
}

suite("Load workbook snapshot", () => {
    test("creates empty snapshots for svn-tree empty workbook resources", async () => {
        const uri = createSvnTreeEmptyUri();

        const snapshot = await loadWorkbookSnapshot(uri);

        assert.strictEqual(snapshot.fileName, "item.xlsx");
        assert.strictEqual(snapshot.filePath, "repo/item.xlsx");
        assert.strictEqual(snapshot.fileSize, 0);
        assert.strictEqual(snapshot.isReadonly, true);
        assert.deepStrictEqual(snapshot.definedNames, []);
        assert.deepStrictEqual(snapshot.sheets, []);
        assert.ok(["Empty workbook", "空工作簿"].includes(snapshot.modifiedTimeLabel ?? ""));
    });

    test("keeps explicit widths and row heights sparse when loading workbook snapshots", async () => {
        const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-load-"));

        try {
            const workbookPath = path.join(tempDirectory, "sparse-dimensions.xlsx");
            const workbook = Workbook.create("Sheet1");
            const sheet = workbook.getSheet("Sheet1");
            sheet.cell("J10").setValue("tail");
            sheet.setColumnWidth(2, 15.125);
            sheet.setRowHeight(3, 18.13);
            await workbook.save(workbookPath);

            const snapshot = await loadWorkbookSnapshot(vscode.Uri.file(workbookPath));
            const activeSheet = snapshot.sheets[0];

            assert.ok(activeSheet);
            assert.strictEqual(activeSheet?.columnCount, 10);
            assert.strictEqual(activeSheet?.rowCount, 10);
            assert.deepStrictEqual(activeSheet?.columnWidths, [null, 15.125]);
            assert.deepStrictEqual(activeSheet?.rowHeights, { "3": 18.13 });
        } finally {
            await rm(tempDirectory, { recursive: true, force: true });
        }
    });
});
