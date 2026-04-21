import * as vscode from "vscode";

export function getTabInputKind(input: vscode.Tab["input"]): string {
    if (input instanceof vscode.TabInputText) {
        return "text";
    }

    if (input instanceof vscode.TabInputTextDiff) {
        return "textDiff";
    }

    if (input instanceof vscode.TabInputCustom) {
        return `custom:${input.viewType}`;
    }

    if (input instanceof vscode.TabInputWebview) {
        return `webview:${input.viewType}`;
    }

    if (input instanceof vscode.TabInputNotebook) {
        return `notebook:${input.notebookType}`;
    }

    if (input instanceof vscode.TabInputNotebookDiff) {
        return `notebookDiff:${input.notebookType}`;
    }

    if (input instanceof vscode.TabInputTerminal) {
        return "terminal";
    }

    return `unknown:${input?.constructor?.name ?? typeof input}`;
}

export function isUnknownTabInput(input: vscode.Tab["input"]): boolean {
    return getTabInputKind(input).startsWith("unknown:");
}

export function getTabResourceUri(input: vscode.Tab["input"]): vscode.Uri | undefined {
    if (input instanceof vscode.TabInputText) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputCustom) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputNotebook) {
        return input.uri;
    }

    return undefined;
}
