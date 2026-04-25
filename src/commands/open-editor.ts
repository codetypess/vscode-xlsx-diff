import * as vscode from "vscode";
import { WEBVIEW_TYPE_EDITOR_PANEL } from "../constants";
import { getRuntimeMessages } from "../i18n";
import { getActiveWorkbookUri, getWorkbookUriFromCommandArg } from "./workbook-picker";

export function resolveWorkbookUriForOpenEditor(
    resource: unknown,
    activeWorkbookUri?: vscode.Uri
): vscode.Uri | undefined {
    return getWorkbookUriFromCommandArg(resource) ?? activeWorkbookUri;
}

export async function openEditor(resource?: unknown): Promise<void> {
    const { commands } = getRuntimeMessages();
    const workbookUri = resolveWorkbookUriForOpenEditor(resource, getActiveWorkbookUri());

    if (!workbookUri) {
        await vscode.window.showErrorMessage(commands.openEditorSelectLocalWorkbook);
        return;
    }

    await vscode.commands.executeCommand("vscode.openWith", workbookUri, WEBVIEW_TYPE_EDITOR_PANEL);
}
