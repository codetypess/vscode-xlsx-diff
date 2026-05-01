import * as vscode from "vscode";
import { getRuntimeMessages } from "../i18n";
import {
    getActiveWorkbookUri,
    getWorkbookUriFromCommandArg,
    pickWorkbook,
} from "./workbook-picker";

async function openDiffPanel(
    extensionUri: vscode.Uri,
    leftUri: vscode.Uri,
    rightUri: vscode.Uri
): Promise<void> {
    const { XlsxDiffPanel } = await import("../webview/diff-panel");
    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}

export interface CompareActiveWithDependencies {
    getWorkbookUriFromCommandArg: typeof getWorkbookUriFromCommandArg;
    getActiveWorkbookUri: typeof getActiveWorkbookUri;
    pickWorkbook: typeof pickWorkbook;
    showErrorMessage: typeof vscode.window.showErrorMessage;
    openDiffPanel: (
        extensionUri: vscode.Uri,
        leftUri: vscode.Uri,
        rightUri: vscode.Uri
    ) => Promise<void>;
}

export async function runCompareActiveWith(
    extensionUri: vscode.Uri,
    resource: unknown,
    dependencies: CompareActiveWithDependencies
): Promise<void> {
    const { commands } = getRuntimeMessages();
    const leftUri =
        dependencies.getWorkbookUriFromCommandArg(resource) ?? dependencies.getActiveWorkbookUri();

    if (!leftUri) {
        await dependencies.showErrorMessage(commands.compareActiveWithOpenWorkbookFirst);
        return;
    }

    const rightUri = await dependencies.pickWorkbook(
        commands.compareActiveWithSelectTargetWorkbook,
        leftUri
    );
    if (!rightUri) {
        return;
    }

    await dependencies.openDiffPanel(extensionUri, leftUri, rightUri);
}

export async function compareActiveWith(
    extensionUri: vscode.Uri,
    resource?: unknown
): Promise<void> {
    await runCompareActiveWith(extensionUri, resource, {
        getWorkbookUriFromCommandArg,
        getActiveWorkbookUri,
        pickWorkbook,
        showErrorMessage: vscode.window.showErrorMessage,
        openDiffPanel,
    });
}
