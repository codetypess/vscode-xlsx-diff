import * as vscode from "vscode";

const CONFIGURATION_SECTION = "xlsx-diff";
const COMPARE_FORMULA_SETTING = "compareFormula";

export const COMPARE_FORMULA_CONFIGURATION_KEY =
    `${CONFIGURATION_SECTION}.${COMPARE_FORMULA_SETTING}`;

export function getCompareFormulaEnabled(): boolean {
    return vscode.workspace
        .getConfiguration(CONFIGURATION_SECTION)
        .get<boolean>(COMPARE_FORMULA_SETTING, false);
}

export function affectsCompareFormula(event: vscode.ConfigurationChangeEvent): boolean {
    return event.affectsConfiguration(COMPARE_FORMULA_CONFIGURATION_KEY);
}
