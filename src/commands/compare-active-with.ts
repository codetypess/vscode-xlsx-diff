import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
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

    const { XlsxDiffPanel } = await import("../webview/diff-panel");
    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
