/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as fs from "node:fs";
import * as path from "node:path";
import { JSDOM } from "jsdom";
import { createCellKey } from "../core/model/cells";
import type { EditorRenderPayload } from "../core/model/types";
import {
    createEditorSessionInitMessage,
    createEditorSessionPatchMessage,
} from "../webview-solid/shared/session-protocol";

const EDITOR_BUNDLE_PATH = path.resolve(__dirname, "../../media/editor-panel.js");

const SESSION_OPTIONS = {
    silent: false,
    clearPendingEdits: false,
    preservePendingHistory: false,
    reuseActiveSheetData: false,
    useModelSelection: true,
    perfTraceId: null,
    resetPendingHistory: false,
};

function createEditorPayload(
    overrides: Partial<EditorRenderPayload> & {
        activeSheet?: Partial<EditorRenderPayload["activeSheet"]>;
    } = {}
): EditorRenderPayload {
    const base: EditorRenderPayload = {
        title: "Editor",
        activeSheet: {
            key: "sheet:1",
            rowCount: 20,
            columnCount: 8,
            columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
            cells: {},
            freezePane: null,
            autoFilter: null,
        },
        selection: {
            key: createCellKey(2, 3),
            rowNumber: 2,
            columnNumber: 3,
            address: "C2",
            value: "hello",
            formula: null,
            isPresent: true,
        },
        hasPendingEdits: false,
        canEdit: true,
        sheets: [
            { key: "sheet:1", label: "Sheet1", isActive: true },
            { key: "sheet:2", label: "Sheet2", isActive: false },
        ],
        canUndoStructuralEdits: false,
        canRedoStructuralEdits: false,
    };

    return {
        ...base,
        ...overrides,
        activeSheet: {
            ...base.activeSheet,
            ...(overrides.activeSheet ?? {}),
        },
    };
}

function createPointerEvent(windowLike: any, type: string, init: Record<string, number> = {}) {
    const event = new windowLike.MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        button: init.button ?? 0,
        buttons: init.buttons ?? (type === "pointerup" || type === "pointercancel" ? 0 : 1),
        clientX: init.clientX ?? 0,
        clientY: init.clientY ?? 0,
    });
    Object.defineProperty(event, "pointerId", {
        configurable: true,
        value: init.pointerId ?? 1,
    });
    if (typeof init.timeStamp === "number") {
        Object.defineProperty(event, "timeStamp", {
            configurable: true,
            value: init.timeStamp,
        });
    }
    return event;
}

function normalizeMessage<T>(value: T): T {
    return JSON.parse(JSON.stringify(value)) as T;
}

suite("Solid editor shell DOM", () => {
    let dom: JSDOM | null = null;
    let documentLike: any;
    let windowLike: any;
    let postedMessages: unknown[];

    const flush = async () => {
        await Promise.resolve();
        await Promise.resolve();
    };

    const mountEditorDom = async ({ debugMode = false }: { debugMode?: boolean } = {}) => {
        dom?.window.close();
        dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
            url: "https://example.test/",
            pretendToBeVisual: true,
            runScripts: "dangerously",
        });
        windowLike = dom.window as any;
        documentLike = windowLike.document;
        postedMessages = [];

        windowLike.ResizeObserver = class {
            observe() {
                return undefined;
            }

            disconnect() {
                return undefined;
            }
        };
        windowLike.HTMLElement.prototype.scrollTo = function scrollTo(position: {
            top?: number;
            left?: number;
        }) {
            if (typeof position.top === "number") {
                this.scrollTop = position.top;
            }
            if (typeof position.left === "number") {
                this.scrollLeft = position.left;
            }
        };
        windowLike.__XLSX_EDITOR_STRINGS__ = {};
        windowLike.__XLSX_EDITOR_DEBUG__ = debugMode;
        windowLike.acquireVsCodeApi = () => ({
            postMessage(message: unknown) {
                postedMessages.push(message);
            },
        });

        const bundle = fs.readFileSync(EDITOR_BUNDLE_PATH, "utf8");
        windowLike.eval(bundle);
        await flush();
    };

    const dispatchSessionInit = async (
        overrides: Partial<EditorRenderPayload> & {
            activeSheet?: Partial<EditorRenderPayload["activeSheet"]>;
        } = {}
    ) => {
        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createEditorSessionInitMessage(
                    createEditorPayload(overrides),
                    SESSION_OPTIONS
                ),
            })
        );
        await flush();
    };

    const query = (selector: string) => {
        const element = documentLike.querySelector(selector);
        assert.ok(element, `Expected element for selector: ${selector}`);
        return element as any;
    };

    const click = async (target: string | any) => {
        const element = typeof target === "string" ? query(target) : target;
        element.dispatchEvent(
            new windowLike.MouseEvent("click", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();
        return element;
    };

    const pointerDown = async (target: string | any, init: Record<string, number> = {}) => {
        const element = typeof target === "string" ? query(target) : target;
        element.dispatchEvent(createPointerEvent(windowLike, "pointerdown", init));
        await flush();
        return element;
    };

    const pointerMove = async (target: string | any, init: Record<string, number> = {}) => {
        const element = typeof target === "string" ? query(target) : target;
        element.dispatchEvent(createPointerEvent(windowLike, "pointermove", init));
        await flush();
        return element;
    };

    const pointerUp = async (target: string | any, init: Record<string, number> = {}) => {
        const element = typeof target === "string" ? query(target) : target;
        element.dispatchEvent(createPointerEvent(windowLike, "pointerup", init));
        await flush();
        return element;
    };

    const contextMenu = async (target: string | any, init: Record<string, number> = {}) => {
        const element = typeof target === "string" ? query(target) : target;
        element.dispatchEvent(
            new windowLike.MouseEvent("contextmenu", {
                bubbles: true,
                cancelable: true,
                button: init.button ?? 2,
                buttons: init.buttons ?? 0,
                clientX: init.clientX ?? 0,
                clientY: init.clientY ?? 0,
            })
        );
        await flush();
        return element;
    };

    const inputText = async (selector: string, value: string) => {
        const input = query(selector);
        input.value = value;
        input.dispatchEvent(
            new windowLike.Event("input", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();
        return input;
    };

    const keyDown = async (
        target: Window | Element | string,
        init: KeyboardEventInit & { key: string }
    ) => {
        const eventTarget = typeof target === "string" ? query(target) : target;
        eventTarget.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                ...init,
            })
        );
        await flush();
        return eventTarget;
    };

    const dispatchClipboardEvent = async (
        type: "copy" | "paste",
        clipboardText = "",
        target: Document | Element = documentLike
    ) => {
        const clipboardStore = new Map<string, string>();
        if (clipboardText) {
            clipboardStore.set("text/plain", clipboardText);
        }

        const event = new windowLike.Event(type, {
            bubbles: true,
            cancelable: true,
        }) as Event & {
            clipboardData?: {
                getData: (format: string) => string;
                setData: (format: string, value: string) => void;
            };
        };
        Object.defineProperty(event, "clipboardData", {
            configurable: true,
            value: {
                getData: (format: string) => clipboardStore.get(format) ?? "",
                setData: (format: string, value: string) => {
                    clipboardStore.set(format, value);
                },
            },
        });
        target.dispatchEvent(event);
        await flush();
        return clipboardStore;
    };

    setup(async () => {
        await mountEditorDom();
    });

    teardown(() => {
        dom?.window.close();
    });

    test("opens, closes, switches replace mode, and drags the search panel", async () => {
        assert.deepStrictEqual(normalizeMessage(postedMessages[0]), { type: "ready" });
        await dispatchSessionInit();

        await click('[data-role="search-toggle"]');

        const appElement = query(".app--editor");
        const searchShell = query('[data-role="search-strip"]');
        const searchTabs = Array.from(documentLike.querySelectorAll(".search-strip__tab"));
        assert.strictEqual(searchTabs.length >= 2, true);

        await click(searchTabs[1]);
        assert.ok(documentLike.querySelector(".search-strip__row--replace"));

        Object.defineProperty(appElement, "clientWidth", {
            configurable: true,
            value: 800,
        });
        Object.defineProperty(appElement, "clientHeight", {
            configurable: true,
            value: 600,
        });
        appElement.getBoundingClientRect = () => ({
            left: 0,
            top: 0,
            right: 800,
            bottom: 600,
            width: 800,
            height: 600,
            x: 0,
            y: 0,
            toJSON() {
                return null;
            },
        });
        Object.defineProperty(searchShell, "offsetWidth", {
            configurable: true,
            value: 300,
        });
        Object.defineProperty(searchShell, "offsetHeight", {
            configurable: true,
            value: 180,
        });
        searchShell.getBoundingClientRect = () => ({
            left: 100,
            top: 66,
            right: 400,
            bottom: 246,
            width: 300,
            height: 180,
            x: 100,
            y: 66,
            toJSON() {
                return null;
            },
        });

        searchShell.dispatchEvent(
            createPointerEvent(windowLike, "pointerdown", {
                pointerId: 7,
                clientX: 180,
                clientY: 100,
            })
        );
        windowLike.dispatchEvent(
            createPointerEvent(windowLike, "pointermove", {
                pointerId: 7,
                clientX: 360,
                clientY: 260,
            })
        );
        windowLike.dispatchEvent(
            createPointerEvent(windowLike, "pointerup", {
                pointerId: 7,
                clientX: 360,
                clientY: 260,
            })
        );
        await flush();

        assert.strictEqual(searchShell.style.left, "280px");
        assert.strictEqual(searchShell.style.top, "226px");
        assert.strictEqual(searchShell.style.right, "auto");

        await click('[data-role="search-close"]');
        assert.strictEqual(documentLike.querySelector('[data-role="search-strip"]'), null);
    });

    test("opens find and replace from keyboard shortcuts and focuses goto", async () => {
        await dispatchSessionInit();

        const selectedCellValueInput = query('[data-role="selected-cell-value"]');
        await keyDown(selectedCellValueInput, {
            key: "f",
            metaKey: true,
        });

        const searchInput = query('[data-role="search-input"]') as HTMLInputElement;
        assert.strictEqual(documentLike.activeElement, searchInput);

        await keyDown(windowLike, {
            key: "h",
            metaKey: true,
        });

        assert.ok(documentLike.querySelector(".search-strip__row--replace"));
        assert.strictEqual(documentLike.activeElement, searchInput);

        await keyDown(windowLike, {
            key: "g",
            metaKey: true,
        });

        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        assert.strictEqual(documentLike.activeElement, gotoInput);
        assert.strictEqual(gotoInput.selectionStart, 0);
        assert.strictEqual(gotoInput.selectionEnd, gotoInput.value.length);
    });

    test("keeps search open and switches back to find when clicking the search button from replace mode", async () => {
        await dispatchSessionInit();

        await click('[data-role="search-toggle"]');

        const searchTabs = Array.from(documentLike.querySelectorAll(".search-strip__tab"));
        await click(searchTabs[1]);
        assert.ok(documentLike.querySelector(".search-strip__row--replace"));

        await click('[data-role="search-toggle"]');

        assert.ok(documentLike.querySelector('[data-role="search-strip"]'));
        assert.strictEqual(documentLike.querySelector(".search-strip__row--replace"), null);
    });

    test("dispatches save from keyboard shortcuts while an editable toolbar field is focused", async () => {
        await dispatchSessionInit({
            hasPendingEdits: true,
        });

        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        await keyDown(selectedCellValueInput, {
            key: "s",
            metaKey: true,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "requestSave",
        });
    });

    test("does not open search shortcuts while a grid cell editor is active", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                },
            },
        });

        const cell = query('[data-cell-address="C2"]');
        cell.dispatchEvent(
            new windowLike.MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cellInput = query('[data-role="grid-cell-input"]') as HTMLInputElement;
        await keyDown(cellInput, {
            key: "f",
            metaKey: true,
        });
        assert.strictEqual(documentLike.querySelector('[data-role="search-strip"]'), null);

        await keyDown(cellInput, {
            key: "h",
            metaKey: true,
        });
        assert.strictEqual(documentLike.querySelector('[data-role="search-strip"]'), null);
    });

    test("closes the search panel from escape when focus is outside editable inputs", async () => {
        await dispatchSessionInit();

        await keyDown(windowLike, {
            key: "f",
            metaKey: true,
        });
        assert.ok(query('[data-role="search-strip"]'));

        await keyDown(windowLike, {
            key: "Escape",
        });

        assert.strictEqual(documentLike.querySelector('[data-role="search-strip"]'), null);
    });

    test("applies replace actions locally and syncs pending edits to the host", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "alpha",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 2,
                columnCount: 1,
                columns: ["A"],
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "alpha",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(2, 1)]: {
                        key: createCellKey(2, 1),
                        rowNumber: 2,
                        columnNumber: 1,
                        address: "A2",
                        displayValue: "alpha",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        await click('[data-role="search-toggle"]');
        await inputText('[data-role="search-input"]', "alpha");

        const searchTabs = Array.from(documentLike.querySelectorAll(".search-strip__tab"));
        await click(searchTabs[1]);
        await inputText('[data-role="replace-input"]', "gamma");
        await click('[data-role="replace-button"]');

        assert.strictEqual(
            query(".search-strip__feedback--success").textContent?.trim(),
            "Replaced 1 matching cells."
        );
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-2)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 1,
                    columnNumber: 1,
                    value: "gamma",
                },
            ],
        });
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 2,
            columnNumber: 1,
        });
        assert.strictEqual((query('[data-role="goto-input"]') as HTMLInputElement).value, "A2");
        assert.strictEqual(
            (query('[data-role="selected-cell-value"]') as HTMLInputElement).value,
            "alpha"
        );
    });

    test("switches the search scope summary to the live selected range", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 5)]: {
                        key: createCellKey(4, 5),
                        rowNumber: 4,
                        columnNumber: 5,
                        address: "E4",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        await click('[data-role="search-toggle"]');
        await inputText('[data-role="search-input"]', "target");
        assert.strictEqual(
            query(".search-strip__scope-summary-value").textContent?.trim(),
            "Whole sheet"
        );

        const anchorCell = query('[data-cell-address="C2"]');
        const targetCell = query('[data-cell-address="E4"]');

        await pointerDown(anchorCell, {
            pointerId: 19,
            clientX: 20,
            clientY: 20,
        });
        await pointerMove(targetCell, {
            pointerId: 19,
            buttons: 1,
            clientX: 120,
            clientY: 84,
        });

        const scopeSummary = query(".search-strip__scope-summary-value");
        assert.strictEqual(scopeSummary.textContent?.trim(), "C2:E4");
        assert.strictEqual(scopeSummary.classList.contains("is-selection"), true);

        await pointerUp(windowLike, {
            pointerId: 19,
            buttons: 0,
            clientX: 120,
            clientY: 84,
        });
        await click('[data-role="search-next"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "search",
            query: "target",
            direction: "next",
            options: {
                isRegexp: false,
                matchCase: false,
                wholeWord: false,
            },
            scope: "selection",
            selectionRange: {
                startRow: 2,
                endRow: 4,
                startColumn: 3,
                endColumn: 5,
            },
        });
    });

    test("shows a stronger primary highlight for matched cells inside a selected range", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 5)]: {
                        key: createCellKey(4, 5),
                        rowNumber: 4,
                        columnNumber: 5,
                        address: "E4",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const anchorCell = query('[data-cell-address="C2"]');
        const targetCell = query('[data-cell-address="E4"]');
        const bodyLayer = query(".editor-grid__layer--body");

        await pointerDown(anchorCell, {
            pointerId: 29,
            clientX: 20,
            clientY: 20,
        });
        await pointerMove(targetCell, {
            pointerId: 29,
            buttons: 1,
            clientX: 120,
            clientY: 84,
        });
        await pointerUp(windowLike, {
            pointerId: 29,
            buttons: 0,
            clientX: 120,
            clientY: 84,
        });

        assert.strictEqual(
            bodyLayer.querySelectorAll(".editor-grid__selection-overlay--primary").length,
            0
        );

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: {
                    type: "searchResult",
                    status: "matched",
                    scope: "selection",
                    match: {
                        rowNumber: 4,
                        columnNumber: 5,
                    },
                    matchCount: 1,
                    matchIndex: 1,
                },
            })
        );
        await flush();

        assert.ok(
            bodyLayer.querySelector(
                ".editor-grid__selection-overlay--primary.editor-grid__selection-overlay--search-focus"
            )
        );
        assert.ok(bodyLayer.querySelector(".editor-grid__selection-overlay--range"));
    });

    test("submits goto requests from the toolbar input", async () => {
        await dispatchSessionInit();

        const gotoInput = await inputText('[data-role="goto-input"]', "  Sheet2!B4  ");
        gotoInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "Enter",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "gotoCell",
            reference: "Sheet2!B4",
        });
    });

    test("toggles filters directly from the toolbar without opening a filter strip", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "Name",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 3,
                columnCount: 2,
                columns: ["A", "B"],
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
                        displayValue: "Team",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(2, 1)]: {
                        key: createCellKey(2, 1),
                        rowNumber: 2,
                        columnNumber: 1,
                        address: "A2",
                        displayValue: "Ada",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(2, 2)]: {
                        key: createCellKey(2, 2),
                        rowNumber: 2,
                        columnNumber: 2,
                        address: "B2",
                        displayValue: "Core",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(3, 1)]: {
                        key: createCellKey(3, 1),
                        rowNumber: 3,
                        columnNumber: 1,
                        address: "A3",
                        displayValue: "Grace",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(3, 2)]: {
                        key: createCellKey(3, 2),
                        rowNumber: 3,
                        columnNumber: 2,
                        address: "B3",
                        displayValue: "Data",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        await click('[data-role="filter-toggle"]');

        assert.strictEqual(documentLike.querySelector('[data-role="filter-strip"]'), null);
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
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
        });
    });

    test("syncs the visible selection when search results move to a matched cell", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 5)]: {
                        key: createCellKey(4, 5),
                        rowNumber: 4,
                        columnNumber: 5,
                        address: "E4",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 640,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        const fillHandle = () => query('[data-role="grid-fill-handle"]') as HTMLSpanElement;

        assert.strictEqual(gotoInput.value, "C2");
        assert.strictEqual(selectedCellValueInput.value, "anchor");
        assert.strictEqual(fillHandle().dataset.rowNumber, "2");
        assert.strictEqual(fillHandle().dataset.columnNumber, "3");

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: {
                    type: "searchResult",
                    status: "matched",
                    scope: "sheet",
                    match: {
                        rowNumber: 4,
                        columnNumber: 5,
                    },
                    matchCount: 1,
                    matchIndex: 1,
                },
            })
        );
        await flush();

        assert.strictEqual(gotoInput.value, "E4");
        assert.strictEqual(selectedCellValueInput.value, "target");
        assert.strictEqual(fillHandle().dataset.rowNumber, "4");
        assert.strictEqual(fillHandle().dataset.columnNumber, "5");
    });

    test("switches sheets from footer tabs", async () => {
        await dispatchSessionInit();

        await click('[data-role="sheet-tab"][data-sheet-key="sheet:2"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setSheet",
            sheetKey: "sheet:2",
        });
    });

    test("keeps the toolbar cell value synchronized with live drafts and committed edits", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                },
            },
        });

        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        assert.strictEqual(selectedCellValueInput.value, "hello");

        const cell = query('[data-cell-address="C2"]');
        cell.dispatchEvent(
            new windowLike.MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cellInput = query('[data-role="grid-cell-input"]') as HTMLInputElement;
        cellInput.value = "draft value";
        cellInput.dispatchEvent(
            new windowLike.Event("input", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        assert.strictEqual(selectedCellValueInput.value, "draft value");

        cellInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "Enter",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "draft value",
                },
            ],
        });
        assert.strictEqual(selectedCellValueInput.value, "draft value");
    });

    test("commits toolbar cell value edits only when pressing enter", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                },
            },
        });

        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        assert.strictEqual(selectedCellValueInput.readOnly, false);

        selectedCellValueInput.focus();
        await flush();
        await inputText('[data-role="selected-cell-value"]', "blur value");

        const messageCountBeforeBlur = postedMessages.length;
        selectedCellValueInput.blur();
        await flush();

        assert.strictEqual(postedMessages.length, messageCountBeforeBlur);
        assert.strictEqual(selectedCellValueInput.value, "hello");

        selectedCellValueInput.focus();
        await flush();
        await inputText('[data-role="selected-cell-value"]', "enter value");
        await keyDown(selectedCellValueInput, {
            key: "Enter",
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "enter value",
                },
            ],
        });
        assert.strictEqual(selectedCellValueInput.value, "enter value");
    });

    test("resets toolbar cell value edits when selecting another cell", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                    [createCellKey(2, 4)]: {
                        key: createCellKey(2, 4),
                        rowNumber: 2,
                        columnNumber: 4,
                        address: "D2",
                        displayValue: "world",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 640,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const targetCell = query('[data-cell-address="D2"]');

        selectedCellValueInput.focus();
        await flush();
        assert.strictEqual(documentLike.activeElement, selectedCellValueInput);
        await inputText('[data-role="selected-cell-value"]', "draft value");

        const messageCountBeforeSelection = postedMessages.length;
        await pointerDown(targetCell, {
            pointerId: 23,
            clientX: 72,
            clientY: 24,
        });
        await pointerUp(windowLike, {
            pointerId: 23,
            buttons: 0,
            clientX: 72,
            clientY: 24,
        });

        assert.deepStrictEqual(
            postedMessages.slice(messageCountBeforeSelection).map(normalizeMessage),
            [
                {
                    type: "selectCell",
                    rowNumber: 2,
                    columnNumber: 4,
                },
            ]
        );
        assert.notStrictEqual(documentLike.activeElement, selectedCellValueInput);
        assert.strictEqual(selectedCellValueInput.value, "world");
    });

    test("copies the current selection to the clipboard", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                },
            },
        });

        const clipboard = await dispatchClipboardEvent("copy");
        assert.strictEqual(clipboard.get("text/plain"), "hello");
    });

    test("pastes clipboard text into editable cells", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 6,
                columnCount: 6,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        await dispatchClipboardEvent("paste", "first\tsecond\nthird\tfourth");

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "first",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 4,
                    value: "second",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 3,
                    columnNumber: 3,
                    value: "third",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 3,
                    columnNumber: 4,
                    value: "fourth",
                },
            ],
        });
    });

    test("enters edit mode for synthetic viewport rows beyond the sheet row count", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "",
                formula: null,
                isPresent: false,
            },
            activeSheet: {
                rowCount: 2,
                columnCount: 2,
                columns: ["A", "B"],
                cells: {},
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.scrollTop = 0;
        viewport.scrollLeft = 0;
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const syntheticCell = query('[data-cell-address="B6"]');
        syntheticCell.dispatchEvent(
            new windowLike.MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cellInput = query('[data-role="grid-cell-input"]') as HTMLInputElement;
        assert.strictEqual(cellInput.value, "");
        assert.strictEqual(documentLike.activeElement, cellInput);

        cellInput.value = "synthetic";
        cellInput.dispatchEvent(
            new windowLike.Event("input", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        cellInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "Enter",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 6,
                    columnNumber: 2,
                    value: "synthetic",
                },
            ],
        });
    });

    test("clears pending save state after a save-complete session patch", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "",
                formula: null,
                isPresent: false,
            },
            activeSheet: {
                rowCount: 2,
                columnCount: 2,
                columns: ["A", "B"],
                cells: {},
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const syntheticCell = query('[data-cell-address="B6"]');
        syntheticCell.dispatchEvent(
            new windowLike.MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cellInput = query('[data-role="grid-cell-input"]') as HTMLInputElement;
        cellInput.value = "synthetic";
        cellInput.dispatchEvent(
            new windowLike.Event("input", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        cellInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "Enter",
            })
        );
        await flush();

        const saveButton = query('[data-role="save-button"]') as HTMLButtonElement;
        assert.strictEqual(saveButton.disabled, false);
        assert.ok(saveButton.classList.contains("is-dirty"));
        assert.ok(query('[data-cell-address="B6"]').classList.contains("grid__cell--pending"));

        await click(saveButton);

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "requestSave",
        });
        assert.strictEqual(saveButton.disabled, true);
        assert.ok(saveButton.classList.contains("is-loading"));

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createEditorSessionPatchMessage([
                    {
                        kind: "document:workbook",
                        hasPendingEdits: false,
                    },
                    {
                        kind: "ui:editingDrafts",
                        clearPendingEdits: true,
                        preservePendingHistory: true,
                    },
                ]),
            })
        );
        await flush();

        assert.strictEqual(saveButton.disabled, true);
        assert.ok(!saveButton.classList.contains("is-dirty"));
        assert.ok(!saveButton.classList.contains("is-loading"));
        assert.ok(!query('[data-cell-address="B6"]').classList.contains("grid__cell--pending"));
    });

    test("preserves viewport scroll after a save-complete patch for the same sheet", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 200,
                columnCount: 8,
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.scrollTop = 960;
        viewport.scrollLeft = 120;
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createEditorSessionPatchMessage([
                    {
                        kind: "document:workbook",
                        hasPendingEdits: false,
                    },
                    {
                        kind: "document:activeSheet",
                        activeSheet: {
                            key: "sheet:1",
                            rowCount: 200,
                            columnCount: 8,
                            freezePane: null,
                            autoFilter: null,
                        },
                    },
                    {
                        kind: "ui:selection",
                        selection: {
                            key: createCellKey(2, 3),
                            rowNumber: 2,
                            columnNumber: 3,
                            address: "C2",
                            value: "hello",
                            formula: null,
                            isPresent: true,
                        },
                    },
                    {
                        kind: "ui:editingDrafts",
                        clearPendingEdits: true,
                        preservePendingHistory: true,
                    },
                    {
                        kind: "ui:viewport",
                        reuseActiveSheetData: true,
                        useModelSelection: true,
                    },
                ]),
            })
        );
        await flush();

        const nextViewport = query('[data-role="editor-grid-viewport"]');
        assert.strictEqual(nextViewport, viewport);
        assert.strictEqual(nextViewport.scrollTop, 960);
        assert.strictEqual(nextViewport.scrollLeft, 120);
    });

    test("moves the active selection with keyboard navigation", async () => {
        await dispatchSessionInit();

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "ArrowRight",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 2,
            columnNumber: 4,
        });
    });

    test("reveals the next page selection when keyboard navigation moves offscreen", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 200,
                columnCount: 8,
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 120,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 240,
        });
        viewport.scrollTop = 0;
        viewport.scrollLeft = 0;
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        assert.strictEqual(gotoInput.value, "C2");

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "PageDown",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 5,
            columnNumber: 3,
        });
        assert.strictEqual(gotoInput.value, "C5");
        assert.ok(viewport.scrollTop > 0);
    });

    test("dispatches undo and redo messages from keyboard shortcuts", async () => {
        await dispatchSessionInit({
            canUndoStructuralEdits: true,
            canRedoStructuralEdits: true,
        });

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "z",
                metaKey: true,
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "undoSheetEdit",
        });

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "z",
                metaKey: true,
                shiftKey: true,
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "redoSheetEdit",
        });
    });

    test("dispatches undo and redo shortcuts when a readonly toolbar field is focused", async () => {
        await dispatchSessionInit({
            canUndoStructuralEdits: true,
            canRedoStructuralEdits: true,
        });

        const selectedCellValueInput = query('[data-role="selected-cell-value"]');
        selectedCellValueInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "z",
                metaKey: true,
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "undoSheetEdit",
        });

        selectedCellValueInput.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "z",
                metaKey: true,
                shiftKey: true,
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "redoSheetEdit",
        });
    });

    test("enables local undo and redo after pending cell edits", async () => {
        await dispatchSessionInit({
            activeSheet: {
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
                },
            },
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "hello",
                formula: null,
                isPresent: true,
            },
            canUndoStructuralEdits: false,
            canRedoStructuralEdits: false,
        });

        const undoButton = query('[data-role="undo-button"]') as HTMLButtonElement;
        const redoButton = query('[data-role="redo-button"]') as HTMLButtonElement;
        assert.strictEqual(undoButton.disabled, true);
        assert.strictEqual(redoButton.disabled, true);

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "Delete",
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "",
                },
            ],
        });
        assert.strictEqual(undoButton.disabled, false);
        assert.strictEqual(redoButton.disabled, true);

        windowLike.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                key: "z",
                metaKey: true,
            })
        );
        await flush();

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [],
        });
        assert.strictEqual(undoButton.disabled, true);
        assert.strictEqual(redoButton.disabled, false);

        await click(redoButton);

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 2,
                    columnNumber: 3,
                    value: "",
                },
            ],
        });
        assert.strictEqual(undoButton.disabled, false);
    });

    test("moves the visible primary highlight when clicking a different cell", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(8, 2)]: {
                        key: createCellKey(8, 2),
                        rowNumber: 8,
                        columnNumber: 2,
                        address: "B8",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 640,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        const fillHandle = () => query('[data-role="grid-fill-handle"]') as HTMLSpanElement;
        const targetCell = query('[data-cell-address="B8"]');
        const bodyLayer = query(".editor-grid__layer--body");

        assert.strictEqual(gotoInput.value, "A1");
        assert.strictEqual(selectedCellValueInput.value, "anchor");
        assert.strictEqual(fillHandle().dataset.rowNumber, "1");
        assert.strictEqual(fillHandle().dataset.columnNumber, "1");
        assert.strictEqual(
            bodyLayer.querySelectorAll(".editor-grid__selection-overlay--primary").length,
            1
        );

        await pointerDown(targetCell, {
            pointerId: 19,
            clientX: 64,
            clientY: 160,
        });
        await pointerUp(windowLike, {
            pointerId: 19,
            buttons: 0,
            clientX: 64,
            clientY: 160,
        });

        assert.strictEqual(gotoInput.value, "B8");
        assert.strictEqual(selectedCellValueInput.value, "target");
        assert.strictEqual(fillHandle().dataset.rowNumber, "8");
        assert.strictEqual(fillHandle().dataset.columnNumber, "2");
        assert.strictEqual(
            bodyLayer.querySelectorAll(".editor-grid__selection-overlay--primary").length,
            1
        );
    });

    test("extends the selection range while dragging across cells", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 5)]: {
                        key: createCellKey(4, 5),
                        rowNumber: 4,
                        columnNumber: 5,
                        address: "E4",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const anchorCell = query('[data-cell-address="C2"]');
        const targetCell = query('[data-cell-address="E4"]');
        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;

        await pointerDown(anchorCell, {
            pointerId: 11,
            clientX: 20,
            clientY: 20,
        });
        await pointerMove(targetCell, {
            pointerId: 11,
            buttons: 1,
            clientX: 120,
            clientY: 84,
        });

        const rangeOverlay = documentLike.querySelector(".editor-grid__selection-overlay--range");
        assert.ok(rangeOverlay);
        assert.strictEqual(gotoInput.value, "C2:E4");
        assert.strictEqual(selectedCellValueInput.value, "hello");

        await pointerUp(windowLike, {
            pointerId: 11,
            buttons: 0,
            clientX: 120,
            clientY: 84,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 2,
            columnNumber: 3,
        });
        assert.strictEqual(gotoInput.value, "C2:E4");
    });

    test("keeps the anchor cell selected and hides the primary box for dragged ranges", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(4, 2),
                rowNumber: 4,
                columnNumber: 2,
                address: "B4",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(4, 2)]: {
                        key: createCellKey(4, 2),
                        rowNumber: 4,
                        columnNumber: 2,
                        address: "B4",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(14, 3)]: {
                        key: createCellKey(14, 3),
                        rowNumber: 14,
                        columnNumber: 3,
                        address: "C14",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 420,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 420,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const anchorCell = query('[data-cell-address="B4"]');
        const targetCell = query('[data-cell-address="C14"]');
        const gotoInput = query('[data-role="goto-input"]') as HTMLInputElement;
        const selectedCellValueInput = query(
            '[data-role="selected-cell-value"]'
        ) as HTMLInputElement;
        const bodyLayer = query(".editor-grid__layer--body");

        await pointerDown(anchorCell, {
            pointerId: 44,
            clientX: 48,
            clientY: 96,
        });
        await pointerMove(targetCell, {
            pointerId: 44,
            buttons: 1,
            clientX: 168,
            clientY: 320,
        });

        assert.strictEqual(gotoInput.value, "B4:C14");
        assert.strictEqual(selectedCellValueInput.value, "anchor");
        assert.ok(bodyLayer.querySelector(".editor-grid__selection-overlay--range"));
        assert.strictEqual(
            bodyLayer.querySelectorAll(".editor-grid__selection-overlay--primary").length,
            0
        );

        await pointerUp(windowLike, {
            pointerId: 44,
            buttons: 0,
            clientX: 168,
            clientY: 320,
        });

        assert.strictEqual(gotoInput.value, "B4:C14");
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 4,
            columnNumber: 2,
        });
    });

    test("fills preview cells and commits pending edits while dragging the fill handle", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "5",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "5",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(3, 3)]: {
                        key: createCellKey(3, 3),
                        rowNumber: 3,
                        columnNumber: 3,
                        address: "C3",
                        displayValue: "",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 3)]: {
                        key: createCellKey(4, 3),
                        rowNumber: 4,
                        columnNumber: 3,
                        address: "C4",
                        displayValue: "",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const fillHandle = query('[data-role="grid-fill-handle"]');
        const targetCell = query('[data-cell-address="C4"]');

        await pointerDown(fillHandle, {
            pointerId: 17,
            clientX: 96,
            clientY: 56,
        });
        await pointerMove(targetCell, {
            pointerId: 17,
            buttons: 1,
            clientX: 96,
            clientY: 128,
        });

        assert.ok(query('[data-cell-address="C3"]').classList.contains("grid__cell--fill-preview"));
        assert.ok(query('[data-cell-address="C4"]').classList.contains("grid__cell--fill-preview"));

        await pointerUp(windowLike, {
            pointerId: 17,
            buttons: 0,
            clientX: 96,
            clientY: 128,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 3,
                    columnNumber: 3,
                    value: "5",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 4,
                    columnNumber: 3,
                    value: "5",
                },
            ],
        });
        assert.strictEqual(query('[data-cell-address="C4"]').textContent?.trim(), "5");
    });

    test("auto-fills down when double-clicking the fill handle", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "5",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 4,
                columnCount: 4,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "5",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(3, 3)]: {
                        key: createCellKey(3, 3),
                        rowNumber: 3,
                        columnNumber: 3,
                        address: "C3",
                        displayValue: "",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 3)]: {
                        key: createCellKey(4, 3),
                        rowNumber: 4,
                        columnNumber: 3,
                        address: "C4",
                        displayValue: "",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const fillHandle = query('[data-role="grid-fill-handle"]');

        await pointerDown(fillHandle, {
            pointerId: 29,
            clientX: 96,
            clientY: 56,
            timeStamp: 100,
        });
        await pointerUp(windowLike, {
            pointerId: 29,
            buttons: 0,
            clientX: 96,
            clientY: 56,
            timeStamp: 100,
        });
        await pointerDown(fillHandle, {
            pointerId: 29,
            clientX: 96,
            clientY: 56,
            timeStamp: 220,
        });
        await pointerUp(windowLike, {
            pointerId: 29,
            buttons: 0,
            clientX: 96,
            clientY: 56,
            timeStamp: 220,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setPendingEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    rowNumber: 3,
                    columnNumber: 3,
                    value: "5",
                },
                {
                    sheetKey: "sheet:1",
                    rowNumber: 4,
                    columnNumber: 3,
                    value: "5",
                },
            ],
        });
        assert.strictEqual(query('[data-cell-address="C4"]').textContent?.trim(), "5");
    });

    test("dispatches single-cell alignment changes and renders updated alignment state", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "aligned",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cellAlignments: {
                    [createCellKey(2, 3)]: {
                        horizontal: "left",
                    },
                },
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "aligned",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        await click('[data-role="align-right"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setAlignment",
            target: "cell",
            selection: {
                startRow: 2,
                endRow: 2,
                startColumn: 3,
                endColumn: 3,
            },
            alignment: {
                horizontal: "right",
            },
        });

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createEditorSessionPatchMessage([
                    {
                        kind: "document:activeSheet",
                        activeSheet: createEditorPayload({
                            activeSheet: {
                                cellAlignments: {
                                    [createCellKey(2, 3)]: {
                                        horizontal: "right",
                                    },
                                },
                                cells: {
                                    [createCellKey(2, 3)]: {
                                        key: createCellKey(2, 3),
                                        rowNumber: 2,
                                        columnNumber: 3,
                                        address: "C2",
                                        displayValue: "aligned",
                                        formula: null,
                                        styleId: null,
                                    },
                                },
                            },
                        }).activeSheet,
                    },
                ]),
            })
        );
        await flush();

        const cellContent = query('[data-cell-address="C2"] .grid__cell-content');
        assert.strictEqual(cellContent.style.textAlign, "right");
        assert.strictEqual(cellContent.style.justifyContent, "flex-end");
        assert.ok(query('[data-role="align-right"]').classList.contains("is-active"));
    });

    test("keeps neighboring cells unchanged after clicking a cell and applying alignment", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(1, 1),
                rowNumber: 1,
                columnNumber: 1,
                address: "A1",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(8, 2)]: {
                        key: createCellKey(8, 2),
                        rowNumber: 8,
                        columnNumber: 2,
                        address: "B8",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 640,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const targetCell = query('[data-cell-address="B8"]');
        await pointerDown(targetCell, {
            pointerId: 31,
            clientX: 48,
            clientY: 200,
        });
        await pointerUp(windowLike, {
            pointerId: 31,
            buttons: 0,
            clientX: 48,
            clientY: 200,
        });

        await click('[data-role="align-right"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setAlignment",
            target: "cell",
            selection: {
                startRow: 8,
                endRow: 8,
                startColumn: 2,
                endColumn: 2,
            },
            alignment: {
                horizontal: "right",
            },
        });

        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createEditorSessionPatchMessage([
                    {
                        kind: "document:activeSheet",
                        activeSheet: createEditorPayload({
                            activeSheet: {
                                cellAlignments: {
                                    [createCellKey(8, 2)]: {
                                        horizontal: "right",
                                    },
                                },
                                cells: {
                                    [createCellKey(1, 1)]: {
                                        key: createCellKey(1, 1),
                                        rowNumber: 1,
                                        columnNumber: 1,
                                        address: "A1",
                                        displayValue: "anchor",
                                        formula: null,
                                        styleId: null,
                                    },
                                    [createCellKey(8, 2)]: {
                                        key: createCellKey(8, 2),
                                        rowNumber: 8,
                                        columnNumber: 2,
                                        address: "B8",
                                        displayValue: "target",
                                        formula: null,
                                        styleId: null,
                                    },
                                },
                            },
                        }).activeSheet,
                    },
                ]),
            })
        );
        await flush();

        const anchorCellContent = query('[data-cell-address="A1"] .grid__cell-content');
        const targetCellContent = query('[data-cell-address="B8"] .grid__cell-content');

        assert.strictEqual(anchorCellContent.style.textAlign, "left");
        assert.strictEqual(anchorCellContent.style.justifyContent, "flex-start");
        assert.strictEqual(targetCellContent.style.textAlign, "right");
        assert.strictEqual(targetCellContent.style.justifyContent, "flex-end");
    });

    test("collapses an expanded selection before applying single-cell alignment", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 20,
                columnCount: 8,
                cells: {
                    [createCellKey(2, 3)]: {
                        key: createCellKey(2, 3),
                        rowNumber: 2,
                        columnNumber: 3,
                        address: "C2",
                        displayValue: "anchor",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(4, 5)]: {
                        key: createCellKey(4, 5),
                        rowNumber: 4,
                        columnNumber: 5,
                        address: "E4",
                        displayValue: "range",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(8, 2)]: {
                        key: createCellKey(8, 2),
                        rowNumber: 8,
                        columnNumber: 2,
                        address: "B8",
                        displayValue: "target",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 640,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const anchorCell = query('[data-cell-address="C2"]');
        const rangeCell = query('[data-cell-address="E4"]');
        const targetCell = query('[data-cell-address="B8"]');

        await pointerDown(anchorCell, {
            pointerId: 23,
            clientX: 24,
            clientY: 24,
        });
        await pointerMove(rangeCell, {
            pointerId: 23,
            buttons: 1,
            clientX: 160,
            clientY: 88,
        });
        await pointerUp(windowLike, {
            pointerId: 23,
            buttons: 0,
            clientX: 160,
            clientY: 88,
        });

        await pointerDown(targetCell, {
            pointerId: 24,
            clientX: 48,
            clientY: 200,
        });
        await pointerUp(windowLike, {
            pointerId: 24,
            buttons: 0,
            clientX: 48,
            clientY: 200,
        });

        await click('[data-role="align-right"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "setAlignment",
            target: "cell",
            selection: {
                startRow: 8,
                endRow: 8,
                startColumn: 2,
                endColumn: 2,
            },
            alignment: {
                horizontal: "right",
            },
        });
    });

    test("renders overflow spill styles and clamped value content for long cells", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 20,
                columnCount: 4,
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "Long text",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(1, 4)]: {
                        key: createCellKey(1, 4),
                        rowNumber: 1,
                        columnNumber: 4,
                        address: "D1",
                        displayValue: "stop",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const cell = query('[data-cell-address="A1"]');
        const value = query('[data-cell-address="A1"] .grid__cell-value');

        assert.ok(cell.classList.contains("grid__cell--overflow-spill"));
        assert.strictEqual(cell.style.getPropertyValue("--grid-column-max-width"), "120px");
        assert.strictEqual(cell.style.getPropertyValue("--grid-cell-content-max-height"), "22px");
        assert.strictEqual(cell.style.getPropertyValue("--grid-cell-line-clamp"), "1");
        assert.strictEqual(cell.style.getPropertyValue("--grid-cell-display-max-width"), "346px");
        assert.strictEqual(value.textContent, "Long text");
    });

    test("renders overflow spill styles across frozen panes", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 20,
                columnCount: 4,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                    topLeftCell: "B2",
                    activePane: "bottomRight",
                },
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "Long text",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(1, 4)]: {
                        key: createCellKey(1, 4),
                        rowNumber: 1,
                        columnNumber: 4,
                        address: "D1",
                        displayValue: "stop",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 320,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cornerOverlay = query('[data-role="editor-grid-corner-overlay"]');
        const cell = query('[data-cell-address="A1"]');

        assert.ok(cornerOverlay.contains(cell));
        assert.ok(cell.classList.contains("grid__cell--overflow-spill"));
        assert.strictEqual(cell.style.getPropertyValue("--grid-cell-display-max-width"), "346px");
    });

    test("updates the visible grid window after viewport scrolling", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 200,
                columnCount: 8,
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
                    [createCellKey(120, 4)]: {
                        key: createCellKey(120, 4),
                        rowNumber: 120,
                        columnNumber: 4,
                        address: "D120",
                        displayValue: "far",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 120,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 240,
        });
        viewport.scrollTop = 0;
        viewport.scrollLeft = 0;

        assert.strictEqual(
            documentLike.querySelector('[data-role="grid-cell"][data-cell-address="D120"]'),
            null
        );

        viewport.scrollTop = 3200;
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        assert.ok(documentLike.querySelector('[data-role="grid-cell"][data-cell-address="D120"]'));
    });

    test("renders frozen panes into the correct overlay layers", async () => {
        await dispatchSessionInit({
            activeSheet: {
                rowCount: 20,
                columnCount: 6,
                freezePane: {
                    rowCount: 1,
                    columnCount: 1,
                    topLeftCell: "B2",
                    activePane: "bottomRight",
                },
                cells: {
                    [createCellKey(1, 1)]: {
                        key: createCellKey(1, 1),
                        rowNumber: 1,
                        columnNumber: 1,
                        address: "A1",
                        displayValue: "corner",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(1, 2)]: {
                        key: createCellKey(1, 2),
                        rowNumber: 1,
                        columnNumber: 2,
                        address: "B1",
                        displayValue: "top",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(2, 1)]: {
                        key: createCellKey(2, 1),
                        rowNumber: 2,
                        columnNumber: 1,
                        address: "A2",
                        displayValue: "left",
                        formula: null,
                        styleId: null,
                    },
                    [createCellKey(2, 2)]: {
                        key: createCellKey(2, 2),
                        rowNumber: 2,
                        columnNumber: 2,
                        address: "B2",
                        displayValue: "body",
                        formula: null,
                        styleId: null,
                    },
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 320,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const cornerOverlay = query('[data-role="editor-grid-corner-overlay"]');
        const topOverlay = query('[data-role="editor-grid-top-overlay"]');
        const leftOverlay = query('[data-role="editor-grid-left-overlay"]');
        const bodyLayer = query(".editor-grid__layer--body");

        assert.ok(cornerOverlay.querySelector('[data-cell-address="A1"]'));
        assert.ok(topOverlay.querySelector('[data-cell-address="B1"]'));
        assert.ok(leftOverlay.querySelector('[data-cell-address="A2"]'));
        assert.ok(bodyLayer.querySelector('[data-cell-address="B2"]'));

        assert.strictEqual(bodyLayer.querySelector('[data-cell-address="A1"]'), null);
        assert.strictEqual(bodyLayer.querySelector('[data-cell-address="B1"]'), null);
        assert.strictEqual(bodyLayer.querySelector('[data-cell-address="A2"]'), null);
    });

    test("opens a cell context menu and dispatches common table commands", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        await contextMenu('[data-cell-address="E4"]', {
            clientX: 220,
            clientY: 120,
        });

        const cellContextMenu = query('[data-role="cell-context-menu"]');
        assert.ok(cellContextMenu);
        assert.strictEqual(
            cellContextMenu.querySelectorAll(".context-menu__separator").length,
            3
        );
        assert.strictEqual(
            query('[data-role="cell-context-find-shortcut"]').textContent?.trim(),
            "Ctrl/Cmd+F"
        );
        assert.strictEqual(
            query('[data-role="cell-context-replace-shortcut"]').textContent?.trim(),
            "Ctrl/Cmd+H"
        );
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 4,
            columnNumber: 5,
        });

        await click('[data-role="cell-context-insert-row-below"]');
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "insertRow",
            rowNumber: 5,
        });
        assert.strictEqual(documentLike.querySelector('[data-role="cell-context-menu"]'), null);

        await contextMenu('[data-cell-address="E4"]', {
            clientX: 220,
            clientY: 120,
        });
        await click('[data-role="cell-context-delete-column"]');
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "deleteColumn",
            columnNumber: 5,
        });

        await contextMenu('[data-cell-address="E4"]', {
            clientX: 220,
            clientY: 120,
        });
        await click('[data-role="cell-context-find"]');
        assert.strictEqual(documentLike.querySelector('[data-role="cell-context-menu"]'), null);
        assert.ok(query('[data-role="search-strip"]'));
    });

    test("keeps an expanded selection when opening find from the cell context menu", async () => {
        await dispatchSessionInit({
            selection: {
                key: createCellKey(2, 3),
                rowNumber: 2,
                columnNumber: 3,
                address: "C2",
                value: "anchor",
                formula: null,
                isPresent: true,
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const anchorCell = query('[data-cell-address="C2"]');
        const targetCell = query('[data-cell-address="E4"]');

        await pointerDown(anchorCell, {
            pointerId: 29,
            clientX: 20,
            clientY: 20,
        });
        await pointerMove(targetCell, {
            pointerId: 29,
            buttons: 1,
            clientX: 120,
            clientY: 84,
        });
        await pointerUp(windowLike, {
            pointerId: 29,
            buttons: 0,
            clientX: 120,
            clientY: 84,
        });

        const messageCountBeforeContextMenu = postedMessages.length;
        await contextMenu('[data-cell-address="D3"]', {
            clientX: 84,
            clientY: 56,
        });

        assert.ok(query('[data-role="cell-context-menu"]'));
        assert.strictEqual(postedMessages.length, messageCountBeforeContextMenu);

        await click('[data-role="cell-context-find"]');
        assert.strictEqual(
            query(".search-strip__scope-summary-value").textContent?.trim(),
            "C2:E4"
        );
    });

    test("opens row header context menu and dispatches row commands", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        await contextMenu('[data-role="grid-row-header"][data-row-number="4"]', {
            clientX: 24,
            clientY: 120,
        });

        assert.ok(query('[data-role="row-context-menu"]'));
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 4,
            columnNumber: 3,
        });

        await click('[data-role="row-context-insert-below"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "insertRow",
            rowNumber: 5,
        });
        assert.strictEqual(documentLike.querySelector('[data-role="row-context-menu"]'), null);

        await contextMenu('[data-role="grid-row-header"][data-row-number="4"]', {
            clientX: 24,
            clientY: 120,
        });
        await click('[data-role="row-context-height"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "promptRowHeight",
            rowNumber: 4,
        });

        await contextMenu('[data-role="grid-row-header"][data-row-number="4"]', {
            clientX: 24,
            clientY: 120,
        });
        await click('[data-role="row-context-delete"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "deleteRow",
            rowNumber: 4,
        });
    });

    test("opens column header context menu and dispatches column commands", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        await contextMenu('[data-role="grid-column-header"][data-column-number="5"]', {
            clientX: 260,
            clientY: 24,
        });

        assert.ok(query('[data-role="column-context-menu"]'));
        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "selectCell",
            rowNumber: 2,
            columnNumber: 5,
        });

        await click('[data-role="column-context-insert-right"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "insertColumn",
            columnNumber: 6,
        });
        assert.strictEqual(documentLike.querySelector('[data-role="column-context-menu"]'), null);

        await contextMenu('[data-role="grid-column-header"][data-column-number="5"]', {
            clientX: 260,
            clientY: 24,
        });
        await click('[data-role="column-context-width"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "promptColumnWidth",
            columnNumber: 5,
        });

        await contextMenu('[data-role="grid-column-header"][data-column-number="5"]', {
            clientX: 260,
            clientY: 24,
        });
        await click('[data-role="column-context-delete"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "deleteColumn",
            columnNumber: 5,
        });
    });

    test("drags a column resize handle and commits the new width", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const columnHeader = query('[data-role="grid-column-header"][data-column-number="2"]');
        const resizeHandle = query(
            '[data-role="grid-column-header"][data-column-number="2"] [data-role="grid-column-resize-handle"]'
        );
        const startWidth = columnHeader.style.width;

        await pointerDown(resizeHandle, {
            pointerId: 41,
            clientX: 260,
        });
        await pointerMove(windowLike, {
            pointerId: 41,
            buttons: 1,
            clientX: 308,
        });

        assert.notStrictEqual(
            query('[data-role="grid-column-header"][data-column-number="2"]').style.width,
            startWidth
        );

        await pointerUp(windowLike, {
            pointerId: 41,
            buttons: 0,
            clientX: 308,
        });

        const lastMessage = normalizeMessage(postedMessages.at(-1)) as {
            type: string;
            columnNumber?: number;
            width?: unknown;
        };
        assert.strictEqual(lastMessage.type, "setColumnWidth");
        assert.strictEqual(lastMessage.columnNumber, 2);
        assert.strictEqual(typeof lastMessage.width, "number");
    });

    test("drags a row resize handle and commits the new height", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const rowHeader = query('[data-role="grid-row-header"][data-row-number="4"]');
        const resizeHandle = query(
            '[data-role="grid-row-header"][data-row-number="4"] [data-role="grid-row-resize-handle"]'
        );
        const startHeight = rowHeader.style.height;

        await pointerDown(resizeHandle, {
            pointerId: 42,
            clientY: 120,
        });
        await pointerMove(windowLike, {
            pointerId: 42,
            buttons: 1,
            clientY: 148,
        });

        assert.notStrictEqual(
            query('[data-role="grid-row-header"][data-row-number="4"]').style.height,
            startHeight
        );

        await pointerUp(windowLike, {
            pointerId: 42,
            buttons: 0,
            clientY: 148,
        });

        const lastMessage = normalizeMessage(postedMessages.at(-1)) as {
            type: string;
            rowNumber?: number;
            height?: unknown;
        };
        assert.strictEqual(lastMessage.type, "setRowHeight");
        assert.strictEqual(lastMessage.rowNumber, 4);
        assert.strictEqual(typeof lastMessage.height, "number");
    });

    test("closes the grid context menu with escape", async () => {
        await dispatchSessionInit();

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 320,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 520,
        });
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        await contextMenu('[data-role="grid-row-header"][data-row-number="4"]', {
            clientX: 24,
            clientY: 120,
        });
        assert.ok(query('[data-role="row-context-menu"]'));

        await keyDown(windowLike, {
            key: "Escape",
        });

        assert.strictEqual(documentLike.querySelector('[data-role="row-context-menu"]'), null);
    });

    test("disables editing commands in read-only sessions", async () => {
        await dispatchSessionInit({
            canEdit: false,
            hasPendingEdits: true,
            canUndoStructuralEdits: true,
            canRedoStructuralEdits: true,
        });

        assert.ok(documentLike.querySelector('[data-role="read-only-badge"]'));
        assert.strictEqual(query('[data-role="save-button"]').disabled, true);
        assert.strictEqual(query('[data-role="undo-button"]').disabled, true);
        assert.strictEqual(query('[data-role="redo-button"]').disabled, true);

        await contextMenu('[data-role="sheet-tab"][data-sheet-key="sheet:1"]', {
            clientX: 32,
            clientY: 32,
        });

        assert.strictEqual(query('[data-role="sheet-context-add"]').disabled, true);
        assert.strictEqual(query('[data-role="sheet-context-rename"]').disabled, true);
        assert.strictEqual(query('[data-role="sheet-context-delete"]').disabled, true);
    });

    test("toggles the view lock from the toolbar", async () => {
        await dispatchSessionInit();

        await click('[data-role="view-lock-button"]');

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "toggleViewLock",
            rowCount: 1,
            columnCount: 2,
        });
    });

    test("renders debug render stats with extended hover details", async () => {
        await mountEditorDom({ debugMode: true });
        await dispatchSessionInit({
            activeSheet: {
                freezePane: {
                    rowCount: 1,
                    columnCount: 2,
                    topLeftCell: "C2",
                    activePane: "bottomRight",
                },
            },
        });

        const viewport = query('[data-role="editor-grid-viewport"]');
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 240,
        });
        Object.defineProperty(viewport, "clientWidth", {
            configurable: true,
            value: 360,
        });
        viewport.scrollTop = 48;
        viewport.scrollLeft = 72;
        viewport.dispatchEvent(
            new windowLike.Event("scroll", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const badge = query('[data-role="debug-render-stats"]');
        const badgeText = badge.textContent?.trim() ?? "";
        assert.strictEqual(documentLike.querySelector('[data-role="debug-render-popover"]'), null);

        badge.focus();
        await flush();

        const hoverPopover = query('[data-role="debug-render-popover"]');
        const hoverText = hoverPopover.textContent ?? "";

        assert.match(badgeText, /^\d+ rows \d+ cols$/);
        assert.match(hoverText, /Rendered \d+ rows and \d+ columns\./);
        assert.ok(hoverText.includes("sheet: sheet:1 (20x8)"));
        assert.ok(hoverText.includes("render rows: frozen 1 + scroll"));
        assert.ok(hoverText.includes("render cols: frozen 2 + scroll"));
        assert.ok(hoverText.includes("viewport: 360x240"));
        assert.ok(hoverText.includes("scroll: top 48, left 72"));
        assert.ok(hoverText.includes("selection: C2"));
        assert.ok(hoverText.includes("pending edits: drafts 0, workbook no"));
    });
});
