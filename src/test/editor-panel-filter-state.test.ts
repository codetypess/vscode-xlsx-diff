/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { WorkbookEditState } from "../core/fastxlsx/write-cell-value";
import type { SheetAutoFilterSnapshot, SheetSnapshot, WorkbookSnapshot } from "../core/model/types";
import { XlsxEditorPanel } from "../webview/editor-panel";
import { createWorkingSheetEntries } from "../webview/editor-panel/editor-panel-state";

function createSheet(overrides: Partial<SheetSnapshot> = {}): SheetSnapshot {
    return {
        name: "Sheet1",
        rowCount: 5,
        columnCount: 3,
        visibility: "visible",
        mergedRanges: [],
        freezePane: null,
        autoFilter: null,
        cells: {},
        signature: "sheet-signature",
        ...overrides,
    };
}

function createWorkbook(sheet: SheetSnapshot): WorkbookSnapshot {
    return {
        filePath: "/tmp/editor-panel-filter-state.xlsx",
        fileName: "editor-panel-filter-state.xlsx",
        fileSize: 0,
        modifiedTime: new Date("2026-01-01T00:00:00.000Z").toISOString(),
        definedNames: [],
        sheets: [sheet],
        isReadonly: false,
    };
}

function clonePendingState(state: WorkbookEditState): WorkbookEditState {
    return structuredClone(state);
}

function createFilter(sort: SheetAutoFilterSnapshot["sort"]): SheetAutoFilterSnapshot {
    return {
        range: {
            startRow: 1,
            endRow: 4,
            startColumn: 1,
            endColumn: 2,
        },
        sort,
    };
}

function createPanel(workbook: WorkbookSnapshot) {
    const panel = Object.create(XlsxEditorPanel.prototype) as any;
    const pendingStates: WorkbookEditState[] = [];
    const renderCalls: unknown[] = [];
    let latestPendingState: WorkbookEditState = {
        cellEdits: [],
        sheetEdits: [],
        viewEdits: [],
    };

    panel.workbook = workbook;
    panel.workingSheetEntries = createWorkingSheetEntries(workbook);
    panel.state = {
        activeSheetKey: "sheet:0",
        selectedCell: { rowNumber: 1, columnNumber: 1 },
    };
    panel.pendingCellEdits = [];
    panel.pendingSheetEdits = [];
    panel.pendingViewEdits = [];
    panel.sheetUndoStack = [];
    panel.sheetRedoStack = [];
    panel.hasWarnedPendingExternalChange = false;
    panel.hasPendingExternalWorkbookChange = false;
    panel.document = {
        hasPendingEdits: () =>
            latestPendingState.cellEdits.length > 0 ||
            latestPendingState.sheetEdits.length > 0 ||
            (latestPendingState.viewEdits?.length ?? 0) > 0,
    };
    panel.controller = {
        onPendingStateChanged: async (state: WorkbookEditState) => {
            latestPendingState = clonePendingState(state);
            pendingStates.push(latestPendingState);
        },
        onRequestSave: async () => undefined,
        onRequestRevert: async () => undefined,
    };
    panel.render = async (_reason?: unknown, options?: unknown) => {
        renderCalls.push(options);
    };
    panel.enqueueReload = async () => {
        throw new Error("unexpected reload");
    };
    panel.handleError = async (error: unknown) => {
        throw error;
    };

    return {
        panel,
        pendingStates,
        renderCalls,
        getActiveSheet: () => panel.workingSheetEntries[0]!.sheet as SheetSnapshot,
    };
}

suite("Editor panel filter state", () => {
    test("routes filter sort changes into pending view edits", async () => {
        const initialFilter = createFilter(null);
        const sortedFilter = createFilter({
            columnNumber: 2,
            direction: "desc",
        });
        const { panel, pendingStates, renderCalls, getActiveSheet } = createPanel(
            createWorkbook(createSheet({ autoFilter: initialFilter }))
        );

        await panel.handleMessage({
            type: "setFilterState",
            sheetKey: "sheet:0",
            filterState: sortedFilter,
        });

        assert.deepStrictEqual(getActiveSheet().autoFilter, sortedFilter);
        assert.deepStrictEqual(panel.pendingViewEdits, [
            {
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                freezePane: null,
                autoFilter: sortedFilter,
            },
        ]);
        assert.deepStrictEqual(pendingStates[pendingStates.length - 1], {
            cellEdits: [],
            sheetEdits: [],
            viewEdits: [
                {
                    sheetKey: "sheet:0",
                    sheetName: "Sheet1",
                    freezePane: null,
                    autoFilter: sortedFilter,
                },
            ],
        });
        assert.strictEqual(panel.sheetUndoStack.length, 1);
        assert.strictEqual(panel.sheetRedoStack.length, 0);
        assert.deepStrictEqual(renderCalls[renderCalls.length - 1], {
            useModelSelection: true,
            replacePendingEdits: [],
            resetPendingHistory: false,
        });
    });

    test("undoes and redoes clearing a saved filter through message handling", async () => {
        const savedFilter = createFilter({
            columnNumber: 2,
            direction: "asc",
        });
        const { panel, pendingStates, getActiveSheet } = createPanel(
            createWorkbook(createSheet({ autoFilter: savedFilter }))
        );

        await panel.handleMessage({
            type: "setFilterState",
            sheetKey: "sheet:0",
            filterState: null,
        });

        assert.strictEqual(getActiveSheet().autoFilter, null);
        assert.deepStrictEqual(panel.pendingViewEdits, [
            {
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                freezePane: null,
                autoFilter: null,
            },
        ]);

        await panel.handleMessage({
            type: "undoSheetEdit",
        });

        assert.deepStrictEqual(getActiveSheet().autoFilter, savedFilter);
        assert.deepStrictEqual(panel.pendingViewEdits, []);
        assert.deepStrictEqual(pendingStates[pendingStates.length - 1], {
            cellEdits: [],
            sheetEdits: [],
            viewEdits: [],
        });
        assert.strictEqual(panel.sheetUndoStack.length, 0);
        assert.strictEqual(panel.sheetRedoStack.length, 1);

        await panel.handleMessage({
            type: "redoSheetEdit",
        });

        assert.strictEqual(getActiveSheet().autoFilter, null);
        assert.deepStrictEqual(panel.pendingViewEdits, [
            {
                sheetKey: "sheet:0",
                sheetName: "Sheet1",
                freezePane: null,
                autoFilter: null,
            },
        ]);
        assert.deepStrictEqual(pendingStates[pendingStates.length - 1], {
            cellEdits: [],
            sheetEdits: [],
            viewEdits: [
                {
                    sheetKey: "sheet:0",
                    sheetName: "Sheet1",
                    freezePane: null,
                    autoFilter: null,
                },
            ],
        });
        assert.strictEqual(panel.sheetUndoStack.length, 1);
        assert.strictEqual(panel.sheetRedoStack.length, 0);
    });
});
