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
            setColumnWidth(): void;
            setColumnStyle(): void;
            setRowHeight(): void;
            setRowStyle(): void;
            setAlignment(): void;
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
                    setColumnWidth() {
                        mutationContexts.push(inBatch);
                    },
                    setColumnStyle() {
                        mutationContexts.push(inBatch);
                    },
                    setRowHeight() {
                        mutationContexts.push(inBatch);
                    },
                    setRowStyle() {
                        mutationContexts.push(inBatch);
                    },
                    setAlignment() {
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
                            columnWidths: [12],
                            rowHeights: {
                                "2": 18.13,
                            },
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
            assert.deepStrictEqual(mutationContexts, [true, true, true, true]);
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

    test("writes and reloads cell, row, and column alignments", async () => {
        const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-"));

        try {
            const sourcePath = path.join(tempDirectory, "source-alignment.xlsx");
            const destinationPath = path.join(tempDirectory, "destination-alignment.xlsx");
            const workbook = Workbook.create("Sheet1");
            workbook.getSheet("Sheet1").cell("C3").setValue("tail");
            await workbook.save(sourcePath);

            await writeWorkbookEditsToDestination(
                vscode.Uri.file(sourcePath),
                vscode.Uri.file(destinationPath),
                {
                    sheetEdits: [],
                    cellEdits: [],
                    viewEdits: [
                        {
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            freezePane: null,
                            cellAlignments: {
                                "2:2": {
                                    horizontal: "right",
                                },
                            },
                            rowAlignments: {
                                "2": {
                                    vertical: "center",
                                },
                            },
                            columnAlignments: {
                                "3": {
                                    horizontal: "center",
                                },
                            },
                        },
                    ],
                }
            );

            const savedWorkbook = await Workbook.open(destinationPath);
            const savedSheet = savedWorkbook.getSheet("Sheet1");
            assert.strictEqual(savedSheet.getAlignment(2, 2)?.horizontal, "right");
            assert.strictEqual(savedSheet.getRowStyle(2)?.alignment?.vertical, "center");
            assert.strictEqual(savedSheet.getColumnStyle(3)?.alignment?.horizontal, "center");

            const snapshot = await loadWorkbookSnapshot(destinationPath);
            assert.strictEqual(snapshot.sheets[0]?.cellAlignments?.["2:2"]?.horizontal, "right");
            assert.strictEqual(snapshot.sheets[0]?.rowAlignments?.["2"]?.vertical, "center");
            assert.strictEqual(snapshot.sheets[0]?.columnAlignments?.["3"]?.horizontal, "center");
        } finally {
            await rm(tempDirectory, { recursive: true, force: true });
        }
    });

    test("writes and reloads worksheet auto-filter range and sort state", async () => {
        const tempDirectory = await mkdtemp(path.join(os.tmpdir(), "xlsx-diff-"));

        try {
            const sourcePath = path.join(tempDirectory, "source-filter.xlsx");
            const destinationPath = path.join(tempDirectory, "destination-filter.xlsx");
            const workbook = Workbook.create("Sheet1");
            const sheet = workbook.getSheet("Sheet1");
            sheet.cell("A1").setValue("Name");
            sheet.cell("B1").setValue("Score");
            sheet.cell("A2").setValue("alpha");
            sheet.cell("B2").setValue("2");
            sheet.cell("A3").setValue("beta");
            sheet.cell("B3").setValue("1");
            await workbook.save(sourcePath);

            await writeWorkbookEditsToDestination(
                vscode.Uri.file(sourcePath),
                vscode.Uri.file(destinationPath),
                {
                    sheetEdits: [],
                    cellEdits: [],
                    viewEdits: [
                        {
                            sheetKey: "sheet-1",
                            sheetName: "Sheet1",
                            freezePane: null,
                            autoFilter: {
                                range: {
                                    startRow: 1,
                                    endRow: 3,
                                    startColumn: 1,
                                    endColumn: 2,
                                },
                                sort: {
                                    columnNumber: 2,
                                    direction: "desc",
                                },
                            },
                        },
                    ],
                }
            );

            const savedWorkbook = await Workbook.open(destinationPath);
            assert.deepStrictEqual(savedWorkbook.getSheet("Sheet1").getAutoFilterDefinition(), {
                range: "A1:B3",
                columns: [],
                sortState: {
                    range: "A1:B3",
                    conditions: [
                        {
                            columnNumber: 2,
                            descending: true,
                        },
                    ],
                },
            });

            const snapshot = await loadWorkbookSnapshot(destinationPath);
            assert.deepStrictEqual(snapshot.sheets[0]?.autoFilter, {
                range: {
                    startRow: 1,
                    endRow: 3,
                    startColumn: 1,
                    endColumn: 2,
                },
                sort: {
                    columnNumber: 2,
                    direction: "desc",
                },
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
                        autoFilter: {
                            range: {
                                startRow: 1,
                                endRow: 4,
                                startColumn: 1,
                                endColumn: 2,
                            },
                            sort: {
                                columnNumber: 2,
                                direction: "asc",
                            },
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
                autoFilter: {
                    range: {
                        startRow: 1,
                        endRow: 4,
                        startColumn: 1,
                        endColumn: 2,
                    },
                    sort: {
                        columnNumber: 2,
                        direction: "asc",
                    },
                },
            },
        ]);

        document.markSaved();
        assert.strictEqual(document.hasPendingEdits(), false);
    });
});
