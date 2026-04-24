import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import { getWorkbookResourceName } from "./resource-uri";

function getSaveProgressTitle(workbookUri?: vscode.Uri): string {
    if (workbookUri) {
        const workbookName = getWorkbookResourceName(workbookUri);
        return isChineseDisplayLanguage()
            ? `正在保存 ${workbookName}...`
            : `Saving ${workbookName}...`;
    }

    return isChineseDisplayLanguage() ? "正在保存工作簿修改..." : "Saving workbook changes...";
}

export async function withWorkbookSaveProgress<Result>(
    task: () => Promise<Result>,
    options: { workbookUri?: vscode.Uri } = {}
): Promise<Result> {
    return vscode.window.withProgress(
        {
            location: vscode.ProgressLocation.Window,
            title: getSaveProgressTitle(options.workbookUri),
        },
        () => task()
    );
}
