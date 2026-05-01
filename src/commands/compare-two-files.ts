import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
import { pickWorkbook } from "./workbook-picker";

async function openDiffPanel(
    extensionUri: vscode.Uri,
    leftUri: vscode.Uri,
    rightUri: vscode.Uri
): Promise<void> {
    const { XlsxDiffPanel } = await import("../webview/diff-panel");
    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}

export interface CompareTwoFilesDependencies {
    pickWorkbook: typeof pickWorkbook;
    openDiffPanel: (
        extensionUri: vscode.Uri,
        leftUri: vscode.Uri,
        rightUri: vscode.Uri
    ) => Promise<void>;
}

export async function runCompareTwoFiles(
    extensionUri: vscode.Uri,
    dependencies: CompareTwoFilesDependencies
): Promise<void> {
    const { commands } = getRuntimeMessages();
    const leftUri = await dependencies.pickWorkbook(commands.compareTwoFilesSelectLeftWorkbook);
    if (!leftUri) {
        return;
    }

    const rightUri = await dependencies.pickWorkbook(
        commands.compareTwoFilesSelectRightWorkbook,
        leftUri
    );
    if (!rightUri) {
        return;
    }

    await dependencies.openDiffPanel(extensionUri, leftUri, rightUri);
}

export async function compareTwoFiles(extensionUri: vscode.Uri): Promise<void> {
    await runCompareTwoFiles(extensionUri, {
        pickWorkbook,
        openDiffPanel,
    });
}
