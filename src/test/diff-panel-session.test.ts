/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { buildWorkbookDiff } from "../core/diff/build-workbook-diff";
import type { CellSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import { createCellKey, getCellAddress } from "../core/model/cells";
import type { DiffSessionPatchMessage } from "../webview/shared/session-protocol";
import { XlsxDiffPanel } from "../webview/diff-panel";

function createCell(rowNumber: number, columnNumber: number, value: string): CellSnapshot {
    return {
        key: createCellKey(rowNumber, columnNumber),
        rowNumber,
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        displayValue: value,
        formula: null,
        styleId: null,
    };
}

function createSheet(name: string, rows: string[][]): SheetSnapshot {
    const cells: Record<string, CellSnapshot> = {};
    let columnCount = 0;

    rows.forEach((row, rowIndex) => {
        columnCount = Math.max(columnCount, row.length);
        row.forEach((value, columnIndex) => {
            if (!value) {
                return;
            }

            const cell = createCell(rowIndex + 1, columnIndex + 1, value);
            cells[cell.key] = cell;
        });
    });

    return {
        name,
        rowCount: rows.length,
        columnCount,
        visibility: "visible",
        mergedRanges: [],
        columnWidths: [],
        freezePane: null,
        cells,
        signature: `${name}:${rows.map((row) => row.join("|")).join("/")}`,
    };
}

function createWorkbook(fileName: string, sheets: SheetSnapshot[]): WorkbookSnapshot {
    return {
        filePath: `/tmp/${fileName}`,
        fileName,
        fileSize: 128,
        modifiedTime: new Date("2026-04-18T06:51:00.000Z").toISOString(),
        definedNames: [],
        sheets,
    };
}

suite("Diff panel session protocol", () => {
    test("switching sheets emits a session patch after initialization", async () => {
        const leftWorkbook = createWorkbook("left.xlsx", [
            createSheet("Sheet1", [["alpha"]]),
            createSheet("Sheet2", [["beta"]]),
        ]);
        const rightWorkbook = createWorkbook("right.xlsx", [
            createSheet("Sheet1", [["alpha changed"]]),
            createSheet("Sheet2", [["beta changed"]]),
        ]);
        const diffModel = buildWorkbookDiff(leftWorkbook, rightWorkbook, {
            compareFormula: false,
        });

        const sentMessages: unknown[] = [];
        const panel = Object.create(XlsxDiffPanel.prototype) as any;
        panel.diffModel = diffModel;
        panel.activeSheetKey = diffModel.sheets[0]?.key ?? null;
        panel.isWebviewReady = true;
        panel.hasPendingRender = false;
        panel.hasSentSessionInit = true;
        panel.panel = {
            title: "",
            webview: {
                postMessage: async (message: unknown) => {
                    sentMessages.push(message);
                },
            },
        };

        await panel.handleMessage({
            type: "setSheet",
            sheetKey: diffModel.sheets[1]!.key,
        });

        assert.strictEqual(sentMessages.length, 1);
        const message = sentMessages[0] as DiffSessionPatchMessage;

        assert.strictEqual(message.type, "session:patch");
        assert.strictEqual(message.view, "diff");
        assert.strictEqual(message.patches.length, 4);
        assert.strictEqual(message.patches[0]?.kind, "document:comparison");
        assert.strictEqual(message.patches[1]?.kind, "document:activeSheet");
        assert.strictEqual(message.patches[2]?.kind, "ui:navigation");
        assert.strictEqual(message.patches[3]?.kind, "ui:pendingEdits");

        const comparisonPatch = message.patches[0];
        const activeSheetPatch = message.patches[1];
        const navigationPatch = message.patches[2];
        const pendingPatch = message.patches[3];

        assert.ok(comparisonPatch && comparisonPatch.kind === "document:comparison");
        assert.ok(activeSheetPatch && activeSheetPatch.kind === "document:activeSheet");
        assert.ok(navigationPatch && navigationPatch.kind === "ui:navigation");
        assert.ok(pendingPatch && pendingPatch.kind === "ui:pendingEdits");

        assert.strictEqual(activeSheetPatch.activeSheet?.key, diffModel.sheets[1]!.key);
        assert.strictEqual(navigationPatch.activeSheetKey, diffModel.sheets[1]!.key);
        assert.strictEqual(pendingPatch.clearPendingEdits, false);
        assert.strictEqual(comparisonPatch.sheets?.[1]?.isActive, true);
        assert.strictEqual(comparisonPatch.sheets?.[0]?.isActive, false);
    });

    test("rerendering after initialization emits a full session patch", async () => {
        const leftWorkbook = createWorkbook("left.xlsx", [createSheet("Sheet1", [["alpha"]])]);
        const rightWorkbook = createWorkbook("right.xlsx", [
            createSheet("Sheet1", [["alpha changed"]]),
        ]);
        const diffModel = buildWorkbookDiff(leftWorkbook, rightWorkbook, {
            compareFormula: false,
        });

        const sentMessages: unknown[] = [];
        const panel = Object.create(XlsxDiffPanel.prototype) as any;
        panel.diffModel = diffModel;
        panel.activeSheetKey = diffModel.sheets[0]?.key ?? null;
        panel.isWebviewReady = true;
        panel.hasPendingRender = false;
        panel.hasSentSessionInit = true;
        panel.panel = {
            title: "",
            webview: {
                postMessage: async (message: unknown) => {
                    sentMessages.push(message);
                },
            },
        };

        await panel.render(undefined, { clearPendingEdits: true });

        assert.strictEqual(sentMessages.length, 1);
        const message = sentMessages[0] as DiffSessionPatchMessage;

        assert.strictEqual(message.type, "session:patch");
        assert.strictEqual(message.view, "diff");
        assert.strictEqual(message.patches.length, 5);
        assert.strictEqual(message.patches[0]?.kind, "document:comparison");
        assert.strictEqual(message.patches[1]?.kind, "document:activeSheet");
        assert.strictEqual(message.patches[2]?.kind, "ui:navigation");
        assert.strictEqual(message.patches[3]?.kind, "ui:pendingEdits");
        assert.strictEqual(message.patches[4]?.kind, "ui:panel");

        const comparisonPatch = message.patches[0];
        const pendingPatch = message.patches[3];
        const panelPatch = message.patches[4];

        assert.ok(comparisonPatch && comparisonPatch.kind === "document:comparison");
        assert.ok(pendingPatch && pendingPatch.kind === "ui:pendingEdits");
        assert.ok(panelPatch && panelPatch.kind === "ui:panel");

        assert.strictEqual(typeof comparisonPatch.title, "string");
        assert.ok(comparisonPatch.leftFile);
        assert.ok(comparisonPatch.rightFile);
        assert.strictEqual(pendingPatch.clearPendingEdits, true);
        assert.strictEqual(panelPatch.statusMessage, null);
    });
});
