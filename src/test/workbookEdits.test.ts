/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { mkdtemp, rm } from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import * as vscode from "vscode";
import { loadWorkbookSnapshot } from "../core/fastxlsx/loadWorkbookSnapshot";
import {
    writeWorkbookEditsToDestination,
    type WorkbookEditState,
} from "../core/fastxlsx/writeCellValue";
import { XlsxEditorDocument } from "../webview/xlsxEditorDocument";

suite("Workbook edit writer", () => {
    test("applies sheet edits before cell edits", async () => {
        const { Workbook } = await import("fastxlsx");
        const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-"));

        try {
            const sourcePath = path.join(tempDirectory, "source.xlsx");
            const destinationPath = path.join(tempDirectory, "destination.xlsx");
            const workbook = Workbook.create("Base");
            workbook.addSheet("Legacy");
            workbook.getSheet("Base").cell("A1").setValue("base");
            workbook.getSheet("Legacy").cell("A1").setValue("legacy");
            await workbook.save(sourcePath);

            const edits: WorkbookEditState = {
                sheetEdits: [
                    {
                        type: "addSheet",
                        sheetName: "Added",
                        targetIndex: 1,
                    },
                    {
                        type: "deleteSheet",
                        sheetName: "Legacy",
                    },
                ],
                cellEdits: [
                    {
                        sheetName: "Base",
                        rowNumber: 1,
                        columnNumber: 1,
                        value: "updated",
                    },
                    {
                        sheetName: "Added",
                        rowNumber: 2,
                        columnNumber: 2,
                        value: "new",
                    },
                ],
            };

            await writeWorkbookEditsToDestination(
                vscode.Uri.file(sourcePath),
                vscode.Uri.file(destinationPath),
                edits
            );

            const snapshot = await loadWorkbookSnapshot(destinationPath);
            assert.deepStrictEqual(
                snapshot.sheets.map((sheet) => sheet.name),
                ["Base", "Added"]
            );
            assert.strictEqual(snapshot.sheets[0]?.cells["1:1"]?.displayValue, "updated");
            assert.strictEqual(snapshot.sheets[1]?.cells["2:2"]?.displayValue, "new");
        } finally {
            await rm(tempDirectory, { recursive: true, force: true });
        }
    });

    test("tracks sheet edits as pending document state", () => {
        const document = new XlsxEditorDocument(vscode.Uri.file("/tmp/workbook-edits.xlsx"));

        assert.strictEqual(
            document.replacePendingState({
                cellEdits: [],
                sheetEdits: [
                    {
                        type: "addSheet",
                        sheetName: "Sheet2",
                        targetIndex: 1,
                    },
                ],
            }),
            true
        );
        assert.strictEqual(document.hasPendingEdits(), true);
        assert.deepStrictEqual(document.getPendingState().sheetEdits, [
            {
                type: "addSheet",
                sheetName: "Sheet2",
                targetIndex: 1,
            },
        ]);

        document.markSaved();
        assert.strictEqual(document.hasPendingEdits(), false);
    });
});
