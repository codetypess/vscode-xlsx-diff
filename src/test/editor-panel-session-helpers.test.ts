/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as vscode from "vscode";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import {
    createEditorSessionMessage,
} from "../webview/editor-panel/editor-panel-protocol-adapter";
import { buildReloadedEditorWorkbookSession } from "../webview/editor-panel/editor-workbook-session";
import type { WorkingSheetEntry } from "../webview/editor-panel/editor-panel-types";

function createCell(rowNumber: number, columnNumber: number, displayValue: string): CellSnapshot {
    return {
        key: `${rowNumber}:${columnNumber}`,
        rowNumber,
        columnNumber,
        address: `R${rowNumber}C${columnNumber}`,
        displayValue,
        formula: null,
        styleId: null,
    };
}

function createSheet(
    name: string,
    cells: CellSnapshot[],
    rowCount = 1,
    columnCount = 1
): SheetSnapshot {
    return {
        name,
        rowCount,
        columnCount,
        visibility: "visible",
        mergedRanges: [],
        cells: Object.fromEntries(cells.map((cell) => [cell.key, cell])),
        signature: `${name}-signature`,
    };
}

function createWorkbook(
    fileName: string,
    sheets: SheetSnapshot[],
    overrides: Partial<WorkbookSnapshot> = {}
): WorkbookSnapshot {
    return {
        filePath: `/tmp/${fileName}`,
        fileName,
        fileSize: 128,
        modifiedTime: new Date("2026-05-01T10:00:00.000Z").toISOString(),
        definedNames: [],
        sheets,
        ...overrides,
    };
}

suite("Editor panel session helpers", () => {
    test("creates init and patch messages through the protocol adapter", () => {
        const activeSheet = {
            key: "sheet:0",
            rowCount: 2,
            columnCount: 2,
            columns: ["A", "B"],
            cells: {
                "1:1": createCell(1, 1, "saved"),
            },
            freezePane: null,
            autoFilter: null,
        };
        const renderPayload = {
            title: "Workbook",
            activeSheet,
            selection: null,
            hasPendingEdits: true,
            canEdit: true,
            sheets: [{ key: "sheet:0", label: "Sheet1", isActive: true }],
            canUndoStructuralEdits: false,
            canRedoStructuralEdits: false,
        };
        const sheetEntries: WorkingSheetEntry[] = [
            {
                key: "sheet:0",
                index: 0,
                sheet: createSheet("Sheet1", [createCell(1, 1, "saved")], 2, 2),
            },
        ];

        const initMessage = createEditorSessionMessage({
            hasSentSessionInit: false,
            renderPayload,
            silent: false,
            clearPendingEdits: false,
            preservePendingHistory: false,
            reuseActiveSheetData: true,
            useModelSelection: true,
            replacePendingEdits: [
                {
                    sheetName: "Sheet1",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "draft",
                },
            ],
            sheetEntries,
            resetPendingHistory: false,
            perfTraceId: "trace-1",
        });

        assert.strictEqual(initMessage.type, "session:init");
        assert.deepStrictEqual(initMessage.options.replacePendingEdits, [
            {
                sheetKey: "sheet:0",
                rowNumber: 1,
                columnNumber: 1,
                value: "draft",
            },
        ]);

        const patchMessage = createEditorSessionMessage({
            hasSentSessionInit: true,
            renderPayload,
            silent: true,
            clearPendingEdits: false,
            preservePendingHistory: true,
            reuseActiveSheetData: false,
            useModelSelection: false,
            replacePendingEdits: undefined,
            sheetEntries,
            resetPendingHistory: true,
            perfTraceId: null,
        });

        assert.strictEqual(patchMessage.type, "session:patch");
        assert.ok(
            !patchMessage.patches.some(
                (patch) => patch.kind === "ui:selection"
            )
        );
    });

    test("rebuilds a reloaded workbook session and restores document pending edits", () => {
        const workbook = createWorkbook("saved.xlsx", [
            createSheet("Sheet1", [createCell(1, 1, "saved")], 1, 1),
        ]);

        const reloadedSession = buildReloadedEditorWorkbookSession({
            workbook,
            workbookUri: vscode.Uri.file("/tmp/reloaded.xlsx"),
            clearPendingEdits: false,
            currentPanelState: {
                activeSheetKey: null,
                selectedCell: null,
            },
            pendingCellEdits: [],
            pendingSheetEdits: [],
            pendingViewEdits: [],
            documentPendingState: {
                cellEdits: [
                    {
                        sheetName: "Sheet1",
                        rowNumber: 3,
                        columnNumber: 2,
                        value: "draft",
                    },
                ],
                sheetEdits: [],
                viewEdits: [],
            },
            documentHasPendingEdits: true,
            canUndoStructuralEdits: false,
            canRedoStructuralEdits: false,
        });

        assert.strictEqual(reloadedSession.workbook.fileName, "reloaded.xlsx");
        assert.deepStrictEqual(reloadedSession.pendingCellEdits, [
            {
                sheetName: "Sheet1",
                rowNumber: 3,
                columnNumber: 2,
                value: "draft",
            },
        ]);
        assert.strictEqual(reloadedSession.renderModel.hasPendingEdits, true);
        assert.strictEqual(reloadedSession.renderModel.activeSheet.key, "sheet:0");
    });
});
