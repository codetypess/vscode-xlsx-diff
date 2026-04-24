/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { mkdtemp, rm } from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import * as vscode from "vscode";
import { loadWorkbookSnapshot } from "../core/fastxlsx/load-workbook-snapshot";
import { Workbook } from "../core/fastxlsx/runtime";
import {
    writeWorkbookEditsToDestination,
    type WorkbookEditState,
} from "../core/fastxlsx/write-cell-value";
import { XlsxEditorDocument } from "../webview/xlsx-editor-document";

suite("Workbook edit writer", () => {
    test("applies sheet edits before cell edits", async () => {
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
                        sheetKey: "added-sheet",
                        sheetName: "Added",
                        targetIndex: 1,
                    },
                    {
                        type: "renameSheet",
                        sheetKey: "added-sheet",
                        sheetName: "Added",
                        nextSheetName: "Renamed",
                    },
                    {
                        type: "deleteSheet",
                        sheetKey: "legacy-sheet",
                        sheetName: "Legacy",
                        targetIndex: 2,
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
                        sheetName: "Renamed",
                        rowNumber: 2,
                        columnNumber: 2,
                        value: "new",
                    },
                ],
                viewEdits: [
                    {
                        sheetKey: "base-sheet",
                        sheetName: "Base",
                        freezePane: {
                            columnCount: 1,
                            rowCount: 1,
                        },
                    },
                    {
                        sheetKey: "added-sheet",
                        sheetName: "Renamed",
                        freezePane: {
                            columnCount: 2,
                            rowCount: 0,
                        },
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
                ["Base", "Renamed"]
            );
            assert.strictEqual(snapshot.sheets[0]?.cells["1:1"]?.displayValue, "updated");
            assert.strictEqual(snapshot.sheets[1]?.cells["2:2"]?.displayValue, "new");
            assert.deepStrictEqual(snapshot.sheets[0]?.freezePane, {
                columnCount: 1,
                rowCount: 1,
                topLeftCell: "B2",
                activePane: "bottomRight",
            });
            assert.deepStrictEqual(snapshot.sheets[1]?.freezePane, {
                columnCount: 2,
                rowCount: 0,
                topLeftCell: "C1",
                activePane: "topRight",
            });
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
                        sheetKey: "sheet-2",
                        sheetName: "Sheet2",
                        targetIndex: 1,
                    },
                ],
                viewEdits: [
                    {
                        sheetKey: "sheet-1",
                        sheetName: "Sheet1",
                        freezePane: {
                            columnCount: 1,
                            rowCount: 1,
                        },
                    },
                ],
            }),
            true
        );
        assert.strictEqual(document.hasPendingEdits(), true);
        assert.deepStrictEqual(document.getPendingState().sheetEdits, [
            {
                type: "addSheet",
                sheetKey: "sheet-2",
                sheetName: "Sheet2",
                targetIndex: 1,
            },
        ]);
        assert.deepStrictEqual(document.getPendingState().viewEdits, [
            {
                sheetKey: "sheet-1",
                sheetName: "Sheet1",
                freezePane: {
                    columnCount: 1,
                    rowCount: 1,
                },
            },
        ]);

        document.markSaved();
        assert.strictEqual(document.hasPendingEdits(), false);
    });
});
