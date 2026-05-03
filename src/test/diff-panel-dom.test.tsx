/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import * as fs from "node:fs";
import * as path from "node:path";
import { JSDOM } from "jsdom";
import type { DiffPanelRenderModel } from "../webview/diff-panel/diff-panel-types";
import { createDiffSessionInitMessage } from "../webview/shared/session-protocol";

const DIFF_BUNDLE_PATH = path.resolve(__dirname, "../../media/diff-panel.js");

const SESSION_OPTIONS = {
    clearPendingEdits: false,
};

function createDiffPayload(): DiffPanelRenderModel {
    return {
        title: "Diff",
        leftFile: {
            title: "left.xlsx",
            path: "/tmp/left.xlsx",
            sizeLabel: "1 KB",
            detailFacts: [],
            modifiedLabel: "today",
            isReadonly: false,
        },
        rightFile: {
            title: "right.xlsx",
            path: "/tmp/right.xlsx",
            sizeLabel: "1 KB",
            detailFacts: [],
            modifiedLabel: "today",
            isReadonly: false,
        },
        definedNamesChanged: false,
        sheets: [
            {
                key: "sheet:1",
                kind: "matched",
                label: "Sheet1",
                diffRowCount: 1,
                diffCellCount: 1,
                mergedRangesChanged: false,
                freezePaneChanged: false,
                visibilityChanged: false,
                sheetOrderChanged: false,
                hasDiff: true,
                diffTone: "modified",
                isActive: true,
            },
        ],
        activeSheet: {
            key: "sheet:1",
            kind: "matched",
            label: "Sheet1",
            leftName: "Sheet1",
            rightName: "Sheet1",
            rowCount: 2,
            columnCount: 2,
            columns: [
                {
                    columnNumber: 1,
                    leftColumnNumber: 1,
                    rightColumnNumber: 1,
                    columnWidth: null,
                    leftLabel: "A",
                    rightLabel: "A",
                },
                {
                    columnNumber: 2,
                    leftColumnNumber: 2,
                    rightColumnNumber: 2,
                    columnWidth: null,
                    leftLabel: "B",
                    rightLabel: "B",
                },
            ],
            rows: [
                {
                    rowNumber: 1,
                    leftRowNumber: 1,
                    rightRowNumber: 1,
                    hasDiff: true,
                    diffTone: "modified",
                    cells: [
                        {
                            key: "cell:1:1",
                            columnNumber: 1,
                            address: "A1",
                            status: "equal",
                            diffIndex: null,
                            leftPresent: true,
                            rightPresent: true,
                            leftValue: "alpha",
                            rightValue: "alpha",
                            leftFormula: null,
                            rightFormula: null,
                        },
                        {
                            key: "cell:1:2",
                            columnNumber: 2,
                            address: "B1",
                            status: "modified",
                            diffIndex: 0,
                            leftPresent: true,
                            rightPresent: true,
                            leftValue: "left target",
                            rightValue: "right target",
                            leftFormula: null,
                            rightFormula: null,
                        },
                    ],
                },
                {
                    rowNumber: 2,
                    leftRowNumber: 2,
                    rightRowNumber: 2,
                    hasDiff: false,
                    diffTone: "equal",
                    cells: [],
                },
            ],
            diffRows: [1],
            diffCells: [
                {
                    key: "diff:B1",
                    rowNumber: 1,
                    columnNumber: 2,
                    address: "B1",
                    diffIndex: 0,
                },
            ],
            diffRowCount: 1,
            diffCellCount: 1,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: false,
            sheetOrderChanged: false,
        },
    };
}

function createMappedDiffPayload(): DiffPanelRenderModel {
    return {
        ...createDiffPayload(),
        activeSheet: {
            key: "sheet:1",
            kind: "matched",
            label: "Sheet1",
            leftName: "LeftSheet",
            rightName: "RightSheet",
            rowCount: 1,
            columnCount: 2,
            columns: [
                {
                    columnNumber: 1,
                    leftColumnNumber: 3,
                    rightColumnNumber: 4,
                    columnWidth: null,
                    leftLabel: "C",
                    rightLabel: "D",
                },
                {
                    columnNumber: 2,
                    leftColumnNumber: 5,
                    rightColumnNumber: 6,
                    columnWidth: null,
                    leftLabel: "E",
                    rightLabel: "F",
                },
            ],
            rows: [
                {
                    rowNumber: 1,
                    leftRowNumber: 7,
                    rightRowNumber: 8,
                    hasDiff: true,
                    diffTone: "modified",
                    cells: [
                        {
                            key: "cell:1:1",
                            columnNumber: 1,
                            address: "A1",
                            status: "equal",
                            diffIndex: null,
                            leftPresent: true,
                            rightPresent: true,
                            leftValue: "same",
                            rightValue: "same",
                            leftFormula: null,
                            rightFormula: null,
                        },
                        {
                            key: "cell:1:2",
                            columnNumber: 2,
                            address: "B1",
                            status: "modified",
                            diffIndex: 0,
                            leftPresent: true,
                            rightPresent: true,
                            leftValue: "before",
                            rightValue: "after",
                            leftFormula: null,
                            rightFormula: null,
                        },
                    ],
                },
            ],
            diffRows: [1],
            diffCells: [
                {
                    key: "diff:B1",
                    rowNumber: 1,
                    columnNumber: 2,
                    address: "B1",
                    diffIndex: 0,
                },
            ],
            diffRowCount: 1,
            diffCellCount: 1,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: false,
            sheetOrderChanged: false,
        },
    };
}

function createLargeDiffPayload(): DiffPanelRenderModel {
    const rowCount = 20;
    const diffRowNumbers = [1, 15];
    const rows = Array.from({ length: rowCount }, (_, index) => {
        const rowNumber = index + 1;
        const diffIndex = diffRowNumbers.indexOf(rowNumber);
        const hasDiff = diffIndex >= 0;

        return {
            rowNumber,
            leftRowNumber: rowNumber,
            rightRowNumber: rowNumber,
            hasDiff,
            diffTone: hasDiff ? ("modified" as const) : ("equal" as const),
            cells: hasDiff
                ? [
                      {
                          key: `cell:${rowNumber}:2`,
                          columnNumber: 2,
                          address: `B${rowNumber}`,
                          status: "modified" as const,
                          diffIndex,
                          leftPresent: true,
                          rightPresent: true,
                          leftValue: `left ${rowNumber}`,
                          rightValue: `right ${rowNumber}`,
                          leftFormula: null,
                          rightFormula: null,
                      },
                  ]
                : [],
        };
    });

    return {
        ...createDiffPayload(),
        sheets: [
            {
                key: "sheet:1",
                kind: "matched",
                label: "Sheet1",
                diffRowCount: diffRowNumbers.length,
                diffCellCount: diffRowNumbers.length,
                mergedRangesChanged: false,
                freezePaneChanged: false,
                visibilityChanged: false,
                sheetOrderChanged: false,
                hasDiff: true,
                diffTone: "modified",
                isActive: true,
            },
        ],
        activeSheet: {
            ...createDiffPayload().activeSheet!,
            rowCount,
            rows,
            diffRows: diffRowNumbers,
            diffCells: diffRowNumbers.map((rowNumber, diffIndex) => ({
                key: `diff:B${rowNumber}`,
                rowNumber,
                columnNumber: 2,
                address: `B${rowNumber}`,
                diffIndex,
            })),
            diffRowCount: diffRowNumbers.length,
            diffCellCount: diffRowNumbers.length,
        },
    };
}

suite("Diff panel DOM", () => {
    let dom: JSDOM;
    let documentLike: any;
    let windowLike: any;
    let postedMessages: unknown[];

    const flush = async () => {
        await Promise.resolve();
        await Promise.resolve();
    };

    const dispatchSessionInit = async (payload = createDiffPayload()) => {
        windowLike.dispatchEvent(
            new windowLike.MessageEvent("message", {
                data: createDiffSessionInitMessage(payload, SESSION_OPTIONS),
            })
        );
        await flush();
    };

    const query = (selector: string) => {
        const element = documentLike.querySelector(selector);
        assert.ok(element, `Expected element for selector: ${selector}`);
        return element as HTMLElement;
    };

    const click = async (selector: string) => {
        const element = query(selector);
        element.dispatchEvent(
            new windowLike.MouseEvent("click", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();
        return element;
    };

    const clickElement = async (element: HTMLElement) => {
        element.dispatchEvent(
            new windowLike.MouseEvent("click", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();
        return element;
    };

    const doubleClick = async (selector: string) => {
        const element = query(selector);
        element.dispatchEvent(
            new windowLike.MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();
        return element;
    };

    const inputText = async (selector: string, value: string) => {
        const input = query(selector) as HTMLInputElement;
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
        target: HTMLElement | Document,
        init: KeyboardEventInit & { key: string }
    ) => {
        target.dispatchEvent(
            new windowLike.KeyboardEvent("keydown", {
                bubbles: true,
                cancelable: true,
                ...init,
            })
        );
        await flush();
    };

    const normalizeMessage = <T,>(value: T): T => JSON.parse(JSON.stringify(value)) as T;

    setup(async () => {
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
        windowLike.HTMLCanvasElement.prototype.getContext = function getContext() {
            return {
                font: "",
                measureText: () => ({ width: 7 }),
            };
        };
        (windowLike as Record<string, unknown>).__XLSX_DIFF_STRINGS__ = {};
        windowLike.acquireVsCodeApi = () => ({
            postMessage(message: unknown) {
                postedMessages.push(message);
            },
        });

        const bundle = fs.readFileSync(DIFF_BUNDLE_PATH, "utf8");
        windowLike.eval(bundle);
        await flush();
    });

    teardown(() => {
        dom.window.close();
    });

    test("moves the active cell highlight immediately when clicking another visible cell", async () => {
        assert.deepStrictEqual(normalizeMessage(postedMessages[0]), { type: "ready" });
        await dispatchSessionInit();

        const rightA1 = query(
            'button[data-role="diff-cell"][data-side="right"][data-row-number="1"][data-column-number="1"]'
        );
        const leftB1Selector =
            'button[data-role="diff-cell"][data-side="left"][data-row-number="1"][data-column-number="2"]';

        assert.ok(rightA1.classList.contains("diff-cell--active"));

        const leftB1 = await click(leftB1Selector);

        assert.ok(leftB1.classList.contains("diff-cell--active"));
        assert.ok(!rightA1.classList.contains("diff-cell--active"));
    });

    test("enters edit mode for writable empty cells without sparse diff entries", async () => {
        await dispatchSessionInit();

        await doubleClick(
            'button[data-role="diff-cell"][data-side="right"][data-row-number="2"][data-column-number="1"]'
        );

        const input = query(".diff-cell__input") as HTMLInputElement;
        assert.strictEqual(input.value, "");
        assert.strictEqual(documentLike.activeElement, input);

        await inputText(".diff-cell__input", "new value");
        await keyDown(input, {
            key: "s",
            ctrlKey: true,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "saveEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    side: "right",
                    rowNumber: 2,
                    columnNumber: 1,
                    value: "new value",
                },
            ],
        });
    });

    test("saves edited values using source worksheet coordinates", async () => {
        await dispatchSessionInit(createMappedDiffPayload());

        await doubleClick(
            'button[data-role="diff-cell"][data-side="left"][data-row-number="1"][data-column-number="2"]'
        );
        const input = await inputText(".diff-cell__input", "edited");

        await keyDown(input, {
            key: "s",
            ctrlKey: true,
        });

        assert.deepStrictEqual(normalizeMessage(postedMessages.at(-1)), {
            type: "saveEdits",
            edits: [
                {
                    sheetKey: "sheet:1",
                    side: "left",
                    rowNumber: 7,
                    columnNumber: 5,
                    value: "edited",
                },
            ],
        });
    });

    test("reveals the target diff row when navigating to the next diff", async () => {
        await dispatchSessionInit(createLargeDiffPayload());

        const viewport = query(".diff-gridViewport");
        Object.defineProperty(viewport, "clientHeight", {
            configurable: true,
            value: 54,
        });
        windowLike.dispatchEvent(
            new windowLike.Event("resize", {
                bubbles: true,
                cancelable: true,
            })
        );
        await flush();

        const nextDiffButton = Array.from(
            documentLike.querySelectorAll(".diff-button")
        ).find((element) => element.textContent?.includes("Next Diff"));
        assert.ok(nextDiffButton, "Expected Next Diff button");

        await clickElement(nextDiffButton as HTMLElement);

        assert.ok(viewport.scrollTop > 0, `Expected viewport scrollTop to advance, got ${viewport.scrollTop}`);
    });

    test("renders diff cell tooltips with source addresses and values", async () => {
        await dispatchSessionInit(createMappedDiffPayload());

        const leftCell = query(
            'button[data-role="diff-cell"][data-side="left"][data-row-number="1"][data-column-number="2"]'
        );
        const rightCell = query(
            'button[data-role="diff-cell"][data-side="right"][data-row-number="1"][data-column-number="2"]'
        );

        assert.strictEqual(leftCell.getAttribute("title"), "E7\nbefore");
        assert.strictEqual(rightCell.getAttribute("title"), "F8\nafter");
    });
});
