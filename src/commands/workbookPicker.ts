import * as path from "node:path";
import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../displayLanguage";

function getWorkbookFilters(): Record<string, string[]> {
    return {
        [isChineseDisplayLanguage() ? "Excel 工作簿" : "Excel Workbooks"]: ["xlsx"],
    };
}

function getWorkbookUriFromTabInput(
    input:
        | vscode.TabInputText
        | vscode.TabInputTextDiff
        | vscode.TabInputCustom
        | vscode.TabInputNotebook
        | unknown
): vscode.Uri | undefined {
    if (input instanceof vscode.TabInputText) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputTextDiff) {
        return input.modified;
    }

    if (input instanceof vscode.TabInputCustom) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputNotebook) {
        return input.uri;
    }

    return undefined;
}

export function isWorkbookUri(uri: vscode.Uri | undefined): uri is vscode.Uri {
    return Boolean(
        uri && uri.scheme === "file" && path.extname(uri.fsPath).toLowerCase() === ".xlsx"
    );
}

function toDirectoryUri(uri: vscode.Uri | undefined): vscode.Uri | undefined {
    if (!uri) {
        return undefined;
    }

    return vscode.Uri.file(path.dirname(uri.fsPath));
}

export async function pickWorkbook(
    openLabel: string,
    seedUri?: vscode.Uri
): Promise<vscode.Uri | undefined> {
    const selection = await vscode.window.showOpenDialog({
        canSelectFiles: true,
        canSelectFolders: false,
        canSelectMany: false,
        defaultUri: toDirectoryUri(seedUri) ?? seedUri,
        filters: getWorkbookFilters(),
        openLabel,
    });

    return selection?.[0];
}

export function getWorkbookUriFromCommandArg(value: unknown): vscode.Uri | undefined {
    if (value instanceof vscode.Uri) {
        return isWorkbookUri(value) ? value : undefined;
    }

    if (Array.isArray(value)) {
        return value
            .map((item) => getWorkbookUriFromCommandArg(item))
            .find((uri): uri is vscode.Uri => Boolean(uri));
    }

    if (typeof value === "object" && value !== null && "resourceUri" in value) {
        const resourceUri = (value as { resourceUri?: unknown }).resourceUri;
        return getWorkbookUriFromCommandArg(resourceUri);
    }

    return undefined;
}

export function getActiveWorkbookUri(): vscode.Uri | undefined {
    const editorUri = vscode.window.activeTextEditor?.document.uri;
    if (isWorkbookUri(editorUri)) {
        return editorUri;
    }

    const tabUri = getWorkbookUriFromTabInput(
        vscode.window.tabGroups.activeTabGroup.activeTab?.input
    );
    return isWorkbookUri(tabUri) ? tabUri : undefined;
}
