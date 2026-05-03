/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { createCellKey } from "../core/model/cells";
import type { EditorActiveSheetView, EditorSelectionView } from "../core/model/types";
import {
    applyCommittedEditorCellEdit,
    createEditorCellEditingState,
    getEditorSelectionDisplayValue,
    isEditorCellEditingActive,
} from "../webview/editor-panel/editing-surface-helpers";

function createActiveSheet(): EditorActiveSheetView {
    return {
        key: "sheet:1",
        rowCount: 10,
        columnCount: 8,
        columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
        cells: {
            [createCellKey(2, 3)]: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                displayValue: "hello",
                formula: null,
                styleId: null,
            },
            [createCellKey(4, 5)]: {
                key: createCellKey(4, 5),
                rowNumber: 4,
                columnNumber: 5,
                address: "E4",
                displayValue: "42",
                formula: "=SUM(A1:A2)",
                styleId: null,
            },
        },
        freezePane: null,
        autoFilter: null,
    };
}

function createSelection(
    rowNumber: number,
    columnNumber: number,
    value: string
): EditorSelectionView {
    return {
        key: createCellKey(rowNumber, columnNumber),
        rowNumber,
        columnNumber,
        address: `${String.fromCharCode(64 + columnNumber)}${rowNumber}`,
        value,
        formula: null,
        isPresent: value.length > 0,
    };
}

suite("Editor editing surface helpers", () => {
    test("creates editing state from the pending edit when present", () => {
        const editingCell = createEditorCellEditingState({
            activeSheet: createActiveSheet(),
            rowNumber: 2,
            columnNumber: 3,
            canEdit: true,
            pendingEdits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "draft",
                },
            ],
        });

        assert.deepStrictEqual(editingCell, {
            sheetKey: "sheet:1",
            rowNumber: 2,
            columnNumber: 3,
            modelValue: "hello",
            draftValue: "draft",
        });
    });

    test("blocks editing for formula cells and read-only sessions", () => {
        assert.strictEqual(
            createEditorCellEditingState({
                activeSheet: createActiveSheet(),
                rowNumber: 4,
                columnNumber: 5,
                canEdit: true,
                pendingEdits: [],
            }),
            null
        );
        assert.strictEqual(
            createEditorCellEditingState({
                activeSheet: createActiveSheet(),
                rowNumber: 2,
                columnNumber: 3,
                canEdit: false,
                pendingEdits: [],
            }),
            null
        );
    });

    test("allows editing synthetic viewport cells beyond the current sheet bounds", () => {
        const editingCell = createEditorCellEditingState({
            activeSheet: createActiveSheet(),
            rowNumber: 12,
            columnNumber: 10,
            canEdit: true,
            pendingEdits: [],
        });

        assert.deepStrictEqual(editingCell, {
            sheetKey: "sheet:1",
            rowNumber: 12,
            columnNumber: 10,
            modelValue: "",
            draftValue: "",
        });
    });

    test("committing a changed draft stages a pending edit", () => {
        const editingCell = createEditorCellEditingState({
            activeSheet: createActiveSheet(),
            rowNumber: 2,
            columnNumber: 3,
            canEdit: true,
            pendingEdits: [],
        });

        assert.ok(editingCell);
        assert.deepStrictEqual(
            applyCommittedEditorCellEdit({
                pendingEdits: [],
                editingCell: editingCell!,
                nextValue: "changed",
            }),
            [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "changed",
                },
            ]
        );
    });

    test("committing the model value clears an existing pending edit", () => {
        const editingCell = createEditorCellEditingState({
            activeSheet: createActiveSheet(),
            rowNumber: 2,
            columnNumber: 3,
            canEdit: true,
            pendingEdits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "draft",
                },
            ],
        });

        assert.ok(editingCell);
        assert.deepStrictEqual(
            applyCommittedEditorCellEdit({
                pendingEdits: [
                    {
                        sheetKey: "sheet:1",
                        rowNumber: 2,
                        columnNumber: 3,
                        value: "draft",
                    },
                ],
                editingCell: editingCell!,
                nextValue: "hello",
            }),
            []
        );
    });

    test("selection display prefers the active draft and pending edits", () => {
        const selection = createSelection(2, 3, "hello");
        const editingCell = createEditorCellEditingState({
            activeSheet: createActiveSheet(),
            rowNumber: 2,
            columnNumber: 3,
            canEdit: true,
            pendingEdits: [],
        });

        assert.strictEqual(
            getEditorSelectionDisplayValue({
                activeSheetKey: "sheet:1",
                selection,
                editingCell: {
                    ...editingCell!,
                    draftValue: "typing",
                },
                pendingEdits: [],
            }),
            "typing"
        );
        assert.strictEqual(
            getEditorSelectionDisplayValue({
                activeSheetKey: "sheet:1",
                selection,
                editingCell: null,
                pendingEdits: [
                    {
                        sheetKey: "sheet:1",
                        rowNumber: 2,
                        columnNumber: 3,
                        value: "pending",
                    },
                ],
            }),
            "pending"
        );
        assert.strictEqual(
            isEditorCellEditingActive({
                editingCell,
                sheetKey: "sheet:1",
                rowNumber: 2,
                columnNumber: 3,
            }),
            true
        );
    });
});
