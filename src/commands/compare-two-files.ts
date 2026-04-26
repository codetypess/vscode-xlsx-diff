import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
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

    const { XlsxDiffPanel } = await import("../webview/diff-panel");
    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
