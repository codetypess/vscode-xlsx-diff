export interface FixtureCellEdit {
    cellAddress: string;
    value: string;
}

export interface FixtureRegressionCase {
    name: string;
    sheetName: string;
    focusCellAddress: string;
    focusCellRowNumber: number;
    focusCellColumnNumber: number;
    expectedBaseDisplayValue: string | undefined;
    expectedHeadDisplayValue: string | undefined;
    baseEdits: FixtureCellEdit[];
    headEdits: FixtureCellEdit[];
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
        baseEdits: [
            {
                cellAddress: "F5",
                value: lfValue,
            },
        ],
        headEdits: [
            {
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
        baseEdits: [],
        headEdits: [
            {
                cellAddress: "F5",
                value: "",
            },
        ],
    },
];
