import type { SheetVisibility } from "../core/model/types.js";

export type FixtureWorkbookOperation =
    | {
          type: "setText";
          cellAddress: string;
          value: string;
      }
    | {
          type: "setBackgroundColor";
          cellAddress: string;
          color: string;
      }
    | {
          type: "setFreezePane";
          columnCount: number;
          rowCount: number;
      }
    | {
          type: "setSheetVisibility";
          visibility: SheetVisibility;
      };

export interface FixtureRegressionCase {
    name: string;
    sheetName: string;
    extraSheetNames?: string[];
    focusCellAddress: string;
    focusCellRowNumber: number;
    focusCellColumnNumber: number;
    expectedSheetNames?: string[];
    expectedBaseDisplayValue: string | undefined;
    expectedHeadDisplayValue: string | undefined;
    expectedBaseVisibility?: SheetVisibility;
    expectedHeadVisibility?: SheetVisibility;
    expectedBaseFreezePane?: {
        columnCount: number;
        rowCount: number;
        topLeftCell: string;
        activePane: "bottomLeft" | "topRight" | "bottomRight" | null;
    } | null;
    expectedHeadFreezePane?: {
        columnCount: number;
        rowCount: number;
        topLeftCell: string;
        activePane: "bottomLeft" | "topRight" | "bottomRight" | null;
    } | null;
    expectStyleDifference?: boolean;
    baseOperations: FixtureWorkbookOperation[];
    headOperations: FixtureWorkbookOperation[];
    expectedDiff: {
        totalDiffSheets: number;
        totalDiffRows: number;
        totalDiffCells: number;
        mergedRangesChanged: boolean;
        freezePaneChanged: boolean;
        visibilityChanged: boolean;
    };
}

const lfValue = "$&key1=ARMY==#army.id\n$&key1=ASSET==#assets.id";
const crlfValue = "$&key1=ARMY==#army.id\r\n$&key1=ASSET==#assets.id";

export const fixtureRegressionCases: FixtureRegressionCase[] = [
    {
        name: "newline-only-cell-diff",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: lfValue,
        expectedHeadDisplayValue: crlfValue,
        expectedBaseVisibility: "visible",
        expectedHeadVisibility: "visible",
        expectedBaseFreezePane: null,
        expectedHeadFreezePane: null,
        baseOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: lfValue,
            },
        ],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: crlfValue,
            },
        ],
        expectedDiff: {
            totalDiffSheets: 0,
            totalDiffRows: 0,
            totalDiffCells: 0,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: false,
        },
    },
    {
        name: "empty-string-vs-blank-cell",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: undefined,
        expectedHeadDisplayValue: "",
        expectedBaseVisibility: "visible",
        expectedHeadVisibility: "visible",
        expectedBaseFreezePane: null,
        expectedHeadFreezePane: null,
        baseOperations: [],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "",
            },
        ],
        expectedDiff: {
            totalDiffSheets: 0,
            totalDiffRows: 0,
            totalDiffCells: 0,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: false,
        },
    },
    {
        name: "style-only-background-color",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: "same",
        expectedHeadDisplayValue: "same",
        expectedBaseVisibility: "visible",
        expectedHeadVisibility: "visible",
        expectedBaseFreezePane: null,
        expectedHeadFreezePane: null,
        expectStyleDifference: true,
        baseOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
        ],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
            {
                type: "setBackgroundColor",
                cellAddress: "F5",
                color: "FFFF0000",
            },
        ],
        expectedDiff: {
            totalDiffSheets: 0,
            totalDiffRows: 0,
            totalDiffCells: 0,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: false,
        },
    },
    {
        name: "freeze-pane-only-view-change",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: "same",
        expectedHeadDisplayValue: "same",
        expectedBaseVisibility: "visible",
        expectedHeadVisibility: "visible",
        expectedBaseFreezePane: null,
        expectedHeadFreezePane: {
            columnCount: 1,
            rowCount: 1,
            topLeftCell: "B2",
            activePane: "bottomRight",
        },
        baseOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
        ],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
            {
                type: "setFreezePane",
                columnCount: 1,
                rowCount: 1,
            },
        ],
        expectedDiff: {
            totalDiffSheets: 1,
            totalDiffRows: 0,
            totalDiffCells: 0,
            mergedRangesChanged: false,
            freezePaneChanged: true,
            visibilityChanged: false,
        },
    },
    {
        name: "sheet-visibility-only-structure-change",
        sheetName: "define",
        extraSheetNames: ["helper"],
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedSheetNames: ["define", "helper"],
        expectedBaseDisplayValue: "same",
        expectedHeadDisplayValue: "same",
        expectedBaseVisibility: "visible",
        expectedHeadVisibility: "hidden",
        expectedBaseFreezePane: null,
        expectedHeadFreezePane: null,
        baseOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
        ],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "same",
            },
            {
                type: "setSheetVisibility",
                visibility: "hidden",
            },
        ],
        expectedDiff: {
            totalDiffSheets: 1,
            totalDiffRows: 0,
            totalDiffCells: 0,
            mergedRangesChanged: false,
            freezePaneChanged: false,
            visibilityChanged: true,
        },
    },
];
