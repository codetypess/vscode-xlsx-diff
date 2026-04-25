import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
import { XlsxDiffPanel } from "../webview/diff-panel";
import { getActiveWorkbookUri, getWorkbookUriFromCommandArg, pickWorkbook } from "./workbook-picker";

export async function compareActiveWith(
    extensionUri: vscode.Uri,
    resource?: unknown
): Promise<void> {
    const { commands } = getRuntimeMessages();
    const leftUri = getWorkbookUriFromCommandArg(resource) ?? getActiveWorkbookUri();

    if (!leftUri) {
        await vscode.window.showErrorMessage(commands.compareActiveWithOpenWorkbookFirst);
        return;
    }

    const rightUri = await pickWorkbook(commands.compareActiveWithSelectTargetWorkbook, leftUri);
    if (!rightUri) {
        return;
    }

    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
