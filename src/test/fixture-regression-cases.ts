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
      };

export interface FixtureRegressionCase {
    name: string;
    sheetName: string;
    focusCellAddress: string;
    focusCellRowNumber: number;
    focusCellColumnNumber: number;
    expectedBaseDisplayValue: string | undefined;
    expectedHeadDisplayValue: string | undefined;
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
        },
    },
];
