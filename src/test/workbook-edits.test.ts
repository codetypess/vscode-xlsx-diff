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
import { XlsxEditorDocument } from "../webview/editor-panel";

suite("Workbook edit writer", () => {
    test("batches workbook mutations before saving", async () => {
        const originalOpen = Workbook.open;
        const saveCalls: string[] = [];
        const mutationContexts: boolean[] = [];
        let inBatch = false;
        let batchCalls = 0;

        interface FakeSheet {
            cell(): {
                setValue(): void;
            };
            insertRow(): void;
            deleteRow(): void;
            insertColumn(): void;
            deleteColumn(): void;
            freezePane(): void;
            unfreezePane(): void;
        }

        interface FakeWorkbook {
            batch<Result>(applyChanges: (workbook: FakeWorkbook) => Result): Result;
            addSheet(): void;
            moveSheet(): void;
            renameSheet(): void;
            deleteSheet(): void;
            getSheet(): FakeSheet;
            save(filePath: string): Promise<void>;
        }

        const fakeWorkbook: FakeWorkbook = {
            batch<Result>(applyChanges: (workbook: FakeWorkbook) => Result): Result {
                batchCalls += 1;
                inBatch = true;

                try {
                    return applyChanges(fakeWorkbook);
                } finally {
                    inBatch = false;
                }
            },
            addSheet(): void {},
            moveSheet(): void {},
            renameSheet(): void {},
            deleteSheet(): void {},
            getSheet() {
                return {
                    cell() {
                        return {
                            setValue() {
                                mutationContexts.push(inBatch);
                            },
                        };
                    },
                    insertRow() {
                        mutationContexts.push(inBatch);
                    },
                    deleteRow() {
                        mutationContexts.push(inBatch);
                    },
                    insertColumn() {
                        mutationContexts.push(inBatch);
                    },
                    deleteColumn() {
                        mutationContexts.push(inBatch);
                    },
                    freezePane() {
                        mutationContexts.push(inBatch);
                    },
                    unfreezePane() {
                        mutationContexts.push(inBatch);
                    },
                };
            },
            async save(filePath: string): Promise<void> {
                saveCalls.push(filePath);
            },
        };

        (
            Workbook as unknown as {
                open: typeof Workbook.open;
            }
        ).open = async () => fakeWorkbook as unknown as Workbook;

        try {
            await writeWorkbookEditsToDestination(
                vscode.Uri.file(path.join(os.tmpdir(), "writer-batch.xlsx")),
                vscode.Uri.file(path.join(os.tmpdir(), "writer-batch.xlsx")),
                {
                    cellEdits: [
                        {
                            sheetName: "Sheet1",
                            rowNumber: 1,
                            columnNumber: 1,
                            value: "updated",
                        },
                    ],
                    sheetEdits: [],
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
                }
            );

            assert.strictEqual(batchCalls, 1);
            assert.deepStrictEqual(mutationContexts, [true, true]);
            assert.strictEqual(saveCalls.length, 1);
        } finally {
            (
                Workbook as unknown as {
                    open: typeof Workbook.open;
                }
            ).open = originalOpen;
        }
    });

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

    test("applies row and column edits before cell edits", async () => {
        const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-"));

        try {
            const sourcePath = path.join(tempDirectory, "source-structure.xlsx");
            const destinationPath = path.join(tempDirectory, "destination-structure.xlsx");
            const workbook = Workbook.create("Sheet1");
            workbook.getSheet("Sheet1").cell("A1").setValue("remove");
            workbook.getSheet("Sheet1").cell("B2").setValue("move");
            await workbook.save(sourcePath);

            await writeWorkbookEditsToDestination(
                vscode.Uri.file(sourcePath),
                vscode.Uri.file(destinationPath),
                {
                    sheetEdits: [
                        {
                            type: "deleteRow",
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            rowNumber: 1,
                            count: 1,
                        },
                        {
                            type: "deleteColumn",
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            columnNumber: 1,
                            count: 1,
                        },
                        {
                            type: "insertRow",
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            rowNumber: 1,
                            count: 1,
                        },
                        {
                            type: "insertColumn",
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            columnNumber: 1,
                            count: 1,
                        },
                    ],
                    cellEdits: [
                        {
                            sheetName: "Sheet1",
                            rowNumber: 2,
                            columnNumber: 2,
                            value: "shifted",
                        },
                    ],
                    viewEdits: [],
                }
            );

            const snapshot = await loadWorkbookSnapshot(destinationPath);
            assert.strictEqual(snapshot.sheets[0]?.rowCount, 2);
            assert.strictEqual(snapshot.sheets[0]?.columnCount, 2);
            assert.strictEqual(snapshot.sheets[0]?.cells["2:2"]?.displayValue, "shifted");
            assert.strictEqual(snapshot.sheets[0]?.cells["1:1"], undefined);
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
