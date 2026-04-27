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
      };

export interface FixtureRegressionCase {
    name: string;
    sheetName: string;
    focusCellAddress: string;
    focusCellRowNumber: number;
    focusCellColumnNumber: number;
    expectedBaseDisplayValue: string | undefined;
    expectedHeadDisplayValue: string | undefined;
    expectStyleDifference?: boolean;
    baseOperations: FixtureWorkbookOperation[];
    headOperations: FixtureWorkbookOperation[];
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
    },
    {
        name: "empty-string-vs-blank-cell",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: undefined,
        expectedHeadDisplayValue: "",
        baseOperations: [],
        headOperations: [
            {
                type: "setText",
                cellAddress: "F5",
                value: "",
            },
        ],
    },
    {
        name: "style-only-background-color",
        sheetName: "define",
        focusCellAddress: "F5",
        focusCellRowNumber: 5,
        focusCellColumnNumber: 6,
        expectedBaseDisplayValue: "same",
        expectedHeadDisplayValue: "same",
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
    },
];
