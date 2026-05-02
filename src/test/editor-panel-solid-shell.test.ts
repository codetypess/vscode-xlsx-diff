/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey } from "../core/model/cells";
import {
    createEditorAddSheetMessage,
    createEditorDeleteSheetMessage,
    createEditorFilterToggleMessage,
    createEditorGotoMessage,
    createEditorRenameSheetMessage,
    createEditorSearchMessage,
    createEditorSetSheetMessage,
    getEditorFilterShellState,
    getEditorSheetContextMenuState,
    getEditorShellCapabilities,
} from "../webview-solid/editor-panel/shell-helpers";

suite("Solid editor shell helpers", () => {
    test("derives read-only command gating from workbook capabilities", () => {
        assert.deepStrictEqual(
            getEditorShellCapabilities({
                canEdit: false,
                hasPendingEdits: true,
                canUndoStructuralEdits: true,
                canRedoStructuralEdits: true,
            }),
            {
                isReadOnly: true,
                canRequestSave: false,
                canUndo: false,
                canRedo: false,
            }
        );

        assert.deepStrictEqual(
            getEditorShellCapabilities({
                canEdit: true,
                hasPendingEdits: true,
                canUndoStructuralEdits: true,
                canRedoStructuralEdits: false,
            }),
            {
                isReadOnly: false,
                canRequestSave: true,
                canUndo: true,
                canRedo: false,
            }
        );
    });

    test("creates trimmed goto messages and rejects blank references", () => {
        assert.strictEqual(createEditorGotoMessage("   \t  "), null);
        assert.deepStrictEqual(createEditorGotoMessage("  Sheet2!B4  "), {
            type: "gotoCell",
            reference: "Sheet2!B4",
        });
    });

    test("creates trimmed search messages and rejects blank queries", () => {
        assert.strictEqual(
            createEditorSearchMessage("   ", "next", {
                isRegexp: false,
                matchCase: false,
                wholeWord: false,
            }),
            null
        );
        assert.deepStrictEqual(
            createEditorSearchMessage("  needle  ", "prev", {
                isRegexp: true,
                matchCase: true,
                wholeWord: false,
            }),
            {
                type: "search",
                query: "needle",
                direction: "prev",
                options: {
                    isRegexp: true,
                    matchCase: true,
                    wholeWord: false,
                },
                scope: "sheet",
            }
        );
    });

    test("creates selection-scoped search messages when an expanded range is provided", () => {
        assert.deepStrictEqual(
            createEditorSearchMessage(
                "needle",
                "next",
                {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: true,
                },
                {
                    scope: "selection",
                    selectionRange: {
                        startRow: 2,
                        endRow: 4,
                        startColumn: 3,
                        endColumn: 5,
                    },
                }
            ),
            {
                type: "search",
                query: "needle",
                direction: "next",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: true,
                },
                scope: "selection",
                selectionRange: {
                    startRow: 2,
                    endRow: 4,
                    startColumn: 3,
                    endColumn: 5,
                },
            }
        );
    });

    test("creates sheet-switch messages", () => {
        assert.deepStrictEqual(createEditorSetSheetMessage("sheet:2"), {
            type: "setSheet",
            sheetKey: "sheet:2",
        });
    });

    test("derives sheet-context actions from editability and sheet count", () => {
        assert.deepStrictEqual(
            getEditorSheetContextMenuState({
                canEdit: false,
                sheetCount: 2,
            }),
            {
                canAddSheet: false,
                canRenameSheet: false,
                canDeleteSheet: false,
            }
        );

        assert.deepStrictEqual(
            getEditorSheetContextMenuState({
                canEdit: true,
                sheetCount: 1,
            }),
            {
                canAddSheet: true,
                canRenameSheet: true,
                canDeleteSheet: false,
            }
        );
    });

    test("creates sheet context-menu commands", () => {
        assert.deepStrictEqual(createEditorAddSheetMessage(), {
            type: "addSheet",
        });
        assert.deepStrictEqual(createEditorRenameSheetMessage("sheet:2"), {
            type: "renameSheet",
            sheetKey: "sheet:2",
        });
        assert.deepStrictEqual(createEditorDeleteSheetMessage("sheet:2"), {
            type: "deleteSheet",
            sheetKey: "sheet:2",
        });
    });

    test("derives a candidate filter range from the active cell and creates a toggle message", () => {
        const activeSheet = {
            key: "sheet:1",
            rowCount: 4,
            columnCount: 2,
            autoFilter: null,
            cells: {
                [createCellKey(1, 1)]: {
                    key: createCellKey(1, 1),
                    rowNumber: 1,
                    columnNumber: 1,
                    address: "A1",
                    displayValue: "Name",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(1, 2)]: {
                    key: createCellKey(1, 2),
                    rowNumber: 1,
                    columnNumber: 2,
                    address: "B1",
                    displayValue: "City",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(2, 1)]: {
                    key: createCellKey(2, 1),
                    rowNumber: 2,
                    columnNumber: 1,
                    address: "A2",
                    displayValue: "Alice",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(2, 2)]: {
                    key: createCellKey(2, 2),
                    rowNumber: 2,
                    columnNumber: 2,
                    address: "B2",
                    displayValue: "Paris",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(3, 1)]: {
                    key: createCellKey(3, 1),
                    rowNumber: 3,
                    columnNumber: 1,
                    address: "A3",
                    displayValue: "Bob",
                    formula: null,
                    styleId: null,
                },
                [createCellKey(3, 2)]: {
                    key: createCellKey(3, 2),
                    rowNumber: 3,
                    columnNumber: 2,
                    address: "B3",
                    displayValue: "Rome",
                    formula: null,
                    styleId: null,
                },
            },
        };

        const filterState = getEditorFilterShellState({
            activeSheet,
            selection: {
                rowNumber: 2,
                columnNumber: 1,
            },
            pendingEdits: [],
        });

        assert.strictEqual(filterState.hasActiveFilter, false);
        assert.strictEqual(filterState.canToggle, true);
        assert.deepStrictEqual(filterState.candidateRange, {
            startRow: 1,
            endRow: 3,
            startColumn: 1,
            endColumn: 2,
        });
        assert.deepStrictEqual(filterState.nextFilterState, {
            range: {
                startRow: 1,
                endRow: 3,
                startColumn: 1,
                endColumn: 2,
            },
            sort: null,
        });
        assert.deepStrictEqual(
            createEditorFilterToggleMessage({
                activeSheet,
                selection: {
                    rowNumber: 2,
                    columnNumber: 1,
                },
                pendingEdits: [],
            }),
            {
                type: "setFilterState",
                sheetKey: "sheet:1",
                filterState: {
                    range: {
                        startRow: 1,
                        endRow: 3,
                        startColumn: 1,
                        endColumn: 2,
                    },
                    sort: null,
                },
            }
        );
    });

    test("clears an active filter even when there is no candidate range", () => {
        const activeSheet = {
            key: "sheet:1",
            rowCount: 2,
            columnCount: 2,
            autoFilter: {
                range: {
                    startRow: 1,
                    endRow: 2,
                    startColumn: 1,
                    endColumn: 2,
                },
                sort: null,
            },
            cells: {},
        };

        const filterState = getEditorFilterShellState({
            activeSheet,
            selection: null,
            pendingEdits: [],
        });

        assert.strictEqual(filterState.hasActiveFilter, true);
        assert.strictEqual(filterState.canToggle, true);
        assert.strictEqual(filterState.nextFilterState, null);
        assert.deepStrictEqual(
            createEditorFilterToggleMessage({
                activeSheet,
                selection: null,
                pendingEdits: [],
            }),
            {
                type: "setFilterState",
                sheetKey: "sheet:1",
                filterState: null,
            }
        );
    });
});
