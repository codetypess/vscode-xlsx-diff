/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import type { EditorRenderPayload } from "../core/model/types";
import type { DiffPanelRenderModel } from "../webview/diff-panel/diff-panel-types";
import {
    createWebviewReadyMessage,
    createDiffSessionPatchMessage,
    createDiffSessionInitMessage,
    createEditorSessionPatchMessage,
    createEditorSessionInitMessage,
    createSessionStatusMessage,
    isDiffWebviewOutgoingMessage,
    isEditorWebviewOutgoingMessage,
} from "../webview-solid/shared/session-protocol";
import {
    createInitialDiffSessionState,
    reduceDiffSessionMessage,
} from "../webview-solid/diff-panel/session";
import {
    createInitialEditorSessionState,
    reduceEditorSessionMessage,
} from "../webview-solid/editor-panel/session";

function createEditorPayload(): EditorRenderPayload {
    return {
        title: "Editor",
        activeSheet: {
            key: "sheet:1",
            rowCount: 1,
            columnCount: 1,
            columns: ["A"],
            cells: {},
            freezePane: null,
            autoFilter: null,
        },
        selection: null,
        hasPendingEdits: false,
        canEdit: true,
        sheets: [{ key: "sheet:1", label: "Sheet1", isActive: true }],
        canUndoStructuralEdits: false,
        canRedoStructuralEdits: false,
    };
}

function createDiffPayload(): DiffPanelRenderModel {
    return {
        title: "Diff",
        leftFile: {
            title: "left",
            path: "/tmp/left.xlsx",
            sizeLabel: "1 KB",
            detailFacts: [],
            modifiedLabel: "today",
            isReadonly: false,
        },
        rightFile: {
            title: "right",
            path: "/tmp/right.xlsx",
            sizeLabel: "1 KB",
            detailFacts: [],
            modifiedLabel: "today",
            isReadonly: false,
        },
        definedNamesChanged: false,
        sheets: [],
        activeSheet: null,
    };
}

suite("Solid webview session reducers", () => {
    test("editor reducer initializes from session:init", () => {
        const nextState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(createEditorPayload(), {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );

        assert.strictEqual(nextState.initialized, true);
        assert.strictEqual(nextState.mode, "ready");
        assert.strictEqual(nextState.renderCount, 1);
        assert.strictEqual(nextState.document.workbook.title, "Editor");
        assert.strictEqual(nextState.ui.selection, null);
        assert.deepStrictEqual(nextState.ui.editingDrafts.pendingEdits, []);
    });

    test("editor reducer applies status updates without dropping payload", () => {
        const initializedState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(createEditorPayload(), {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );

        const nextState = reduceEditorSessionMessage(
            initializedState,
            createSessionStatusMessage("editor", "loading", "Reloading")
        );

        assert.strictEqual(nextState.mode, "loading");
        assert.strictEqual(nextState.ui.panel.statusMessage, "Reloading");
        assert.strictEqual(nextState.document.workbook.title, "Editor");
    });

    test("editor reducer applies session patches to UI and document state", () => {
        const initializedState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(createEditorPayload(), {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );

        const nextState = reduceEditorSessionMessage(
            initializedState,
            createEditorSessionPatchMessage([
                {
                    kind: "ui:panel",
                    statusMessage: "Patched",
                    perfTraceId: "trace-1",
                },
                {
                    kind: "ui:viewport",
                    reuseActiveSheetData: true,
                    useModelSelection: true,
                },
                {
                    kind: "document:workbook",
                    hasPendingEdits: true,
                },
            ])
        );

        assert.strictEqual(nextState.ui.panel.statusMessage, "Patched");
        assert.strictEqual(nextState.ui.panel.perfTraceId, "trace-1");
        assert.strictEqual(nextState.ui.viewport.reuseActiveSheetData, true);
        assert.strictEqual(nextState.ui.viewport.useModelSelection, true);
        assert.strictEqual(nextState.document.workbook.hasPendingEdits, true);
    });

    test("editor reducer treats session patches as ready renders and clears nullable panel fields", () => {
        const initializedState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(createEditorPayload(), {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );
        const loadingState = reduceEditorSessionMessage(
            initializedState,
            createSessionStatusMessage("editor", "loading", "Reloading")
        );

        const nextState = reduceEditorSessionMessage(
            loadingState,
            createEditorSessionPatchMessage([
                {
                    kind: "ui:panel",
                    statusMessage: null,
                    perfTraceId: null,
                },
                {
                    kind: "document:workbook",
                    title: "Editor Reloaded",
                },
            ])
        );

        assert.strictEqual(nextState.initialized, true);
        assert.strictEqual(nextState.mode, "ready");
        assert.strictEqual(nextState.renderCount, 2);
        assert.strictEqual(nextState.document.workbook.title, "Editor Reloaded");
        assert.strictEqual(nextState.ui.panel.statusMessage, null);
        assert.strictEqual(nextState.ui.panel.perfTraceId, null);
    });

    test("editor reducer merges partial active-sheet patches with the existing sheet payload", () => {
        const initializedState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(createEditorPayload(), {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );

        const nextState = reduceEditorSessionMessage(
            initializedState,
            createEditorSessionPatchMessage([
                {
                    kind: "document:activeSheet",
                    activeSheet: {
                        key: "sheet:1",
                        rowCount: 4,
                        columnCount: 2,
                        freezePane: null,
                        autoFilter: {
                            range: {
                                startRow: 1,
                                endRow: 4,
                                startColumn: 1,
                                endColumn: 2,
                            },
                            sort: null,
                        },
                    },
                },
            ])
        );

        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.columns, ["A"]);
        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.cells, {});
        assert.strictEqual(nextState.document.workbook.activeSheet?.columnCount, 2);
        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.autoFilter, {
            range: {
                startRow: 1,
                endRow: 4,
                startColumn: 1,
                endColumn: 2,
            },
            sort: null,
        });
    });

    test("editor reducer preserves untouched alignment entries when active-sheet patches use dirty keys", () => {
        const initialPayload = createEditorPayload();
        initialPayload.activeSheet.cellAlignments = {
            "1:1": {
                horizontal: "center",
            },
            "2:2": {
                vertical: "bottom",
            },
        };
        initialPayload.activeSheet.rowAlignments = {
            "2": {
                vertical: "center",
            },
        };
        initialPayload.activeSheet.columnAlignments = {
            "1": {
                horizontal: "right",
            },
        };

        const initializedState = reduceEditorSessionMessage(
            createInitialEditorSessionState(),
            createEditorSessionInitMessage(initialPayload, {
                silent: false,
                clearPendingEdits: false,
                preservePendingHistory: false,
                reuseActiveSheetData: false,
                useModelSelection: false,
                perfTraceId: null,
                resetPendingHistory: false,
            })
        );

        const nextState = reduceEditorSessionMessage(
            initializedState,
            createEditorSessionPatchMessage([
                {
                    kind: "document:activeSheet",
                    activeSheet: {
                        key: "sheet:1",
                        rowCount: 1,
                        columnCount: 1,
                        freezePane: null,
                        autoFilter: null,
                        cellAlignments: {
                            "1:1": {
                                horizontal: "left",
                            },
                        },
                        cellAlignmentDirtyKeys: ["1:1"],
                        rowAlignments: {},
                        rowAlignmentDirtyKeys: ["2"],
                        columnAlignments: {
                            "1": {
                                horizontal: "center",
                            },
                        },
                        columnAlignmentDirtyKeys: ["1"],
                    },
                },
            ])
        );

        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.cellAlignments, {
            "1:1": {
                horizontal: "left",
            },
            "2:2": {
                vertical: "bottom",
            },
        });
        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.rowAlignments, {});
        assert.deepStrictEqual(nextState.document.workbook.activeSheet?.columnAlignments, {
            "1": {
                horizontal: "center",
            },
        });
    });

    test("diff reducer initializes from session:init", () => {
        const nextState = reduceDiffSessionMessage(
            createInitialDiffSessionState(),
            createDiffSessionInitMessage(createDiffPayload(), {
                clearPendingEdits: true,
            })
        );

        assert.strictEqual(nextState.initialized, true);
        assert.strictEqual(nextState.mode, "ready");
        assert.strictEqual(nextState.renderCount, 1);
        assert.strictEqual(nextState.document.comparison.title, "Diff");
        assert.strictEqual(nextState.ui.navigation.activeSheetKey, null);
    });

    test("diff reducer applies session patches", () => {
        const initializedState = reduceDiffSessionMessage(
            createInitialDiffSessionState(),
            createDiffSessionInitMessage(createDiffPayload(), {
                clearPendingEdits: true,
            })
        );

        const nextState = reduceDiffSessionMessage(
            initializedState,
            createDiffSessionPatchMessage([
                {
                    kind: "ui:navigation",
                    activeSheetKey: "sheet:2",
                },
                {
                    kind: "ui:panel",
                    statusMessage: "Ready",
                },
            ])
        );

        assert.strictEqual(nextState.ui.navigation.activeSheetKey, "sheet:2");
        assert.strictEqual(nextState.ui.panel.statusMessage, "Ready");
    });

    test("diff reducer treats session patches as ready renders", () => {
        const initializedState = reduceDiffSessionMessage(
            createInitialDiffSessionState(),
            createDiffSessionInitMessage(createDiffPayload(), {
                clearPendingEdits: true,
            })
        );
        const loadingState = reduceDiffSessionMessage(
            initializedState,
            createSessionStatusMessage("diff", "loading", "Reloading")
        );

        const nextState = reduceDiffSessionMessage(
            loadingState,
            createDiffSessionPatchMessage([
                {
                    kind: "document:comparison",
                    title: "Diff Reloaded",
                },
                {
                    kind: "ui:panel",
                    statusMessage: null,
                },
            ])
        );

        assert.strictEqual(nextState.initialized, true);
        assert.strictEqual(nextState.mode, "ready");
        assert.strictEqual(nextState.renderCount, 2);
        assert.strictEqual(nextState.document.comparison.title, "Diff Reloaded");
        assert.strictEqual(nextState.ui.panel.statusMessage, null);
    });
});

suite("Solid webview protocol guards", () => {
    test("encodes the shared ready message", () => {
        assert.deepStrictEqual(createWebviewReadyMessage(), { type: "ready" });
    });

    test("accepts valid editor outgoing messages", () => {
        assert.strictEqual(isEditorWebviewOutgoingMessage(createWebviewReadyMessage()), true);
        assert.strictEqual(
            isEditorWebviewOutgoingMessage({
                type: "search",
                query: "total",
                direction: "next",
                options: {
                    isRegexp: false,
                    matchCase: false,
                    wholeWord: false,
                },
                scope: "sheet",
            }),
            true
        );
    });

    test("rejects malformed editor outgoing messages", () => {
        assert.strictEqual(
            isEditorWebviewOutgoingMessage({
                type: "setSheet",
            }),
            false
        );
    });

    test("accepts valid diff outgoing messages", () => {
        assert.strictEqual(isDiffWebviewOutgoingMessage(createWebviewReadyMessage()), true);
        assert.strictEqual(
            isDiffWebviewOutgoingMessage({
                type: "saveEdits",
                edits: [
                    {
                        sheetKey: "sheet:1",
                        side: "left",
                        rowNumber: 2,
                        columnNumber: 3,
                        value: "42",
                    },
                ],
            }),
            true
        );
    });

    test("rejects malformed diff outgoing messages", () => {
        assert.strictEqual(
            isDiffWebviewOutgoingMessage({
                type: "saveEdits",
                edits: [{ side: "left" }],
            }),
            false
        );
    });

    test("accepts valid incoming session patch messages", () => {
        assert.strictEqual(
            reduceEditorSessionMessage(
                createInitialEditorSessionState(),
                createEditorSessionPatchMessage([
                    {
                        kind: "ui:panel",
                        statusMessage: "Booting",
                    },
                ])
            ).ui.panel.statusMessage,
            "Booting"
        );
    });
});
