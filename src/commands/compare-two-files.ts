import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
import { XlsxDiffPanel } from "../webview/diff-panel";
import { pickWorkbook } from "./workbook-picker";

export async function compareTwoFiles(extensionUri: vscode.Uri): Promise<void> {
    const { commands } = getRuntimeMessages();
    const leftUri = await pickWorkbook(commands.compareTwoFilesSelectLeftWorkbook);
    if (!leftUri) {
        return;
    }

    const rightUri = await pickWorkbook(commands.compareTwoFilesSelectRightWorkbook, leftUri);
    if (!rightUri) {
        return;
    }

    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
