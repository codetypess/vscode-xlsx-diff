import * as vscode from "vscode";
import { formatI18nMessage, getRuntimeMessages } from "../i18n";
import { getWorkbookResourceName } from "./resource-uri";

function getSaveProgressTitle(workbookUri?: vscode.Uri): string {
    const { workbook } = getRuntimeMessages();
    if (workbookUri) {
        const workbookName = getWorkbookResourceName(workbookUri);
        return formatI18nMessage(workbook.savingWorkbook, { workbookName });
    }

    return workbook.savingChanges;
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
