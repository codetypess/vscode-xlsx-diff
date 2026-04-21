import * as vscode from "vscode";

export function describeUri(uri: vscode.Uri | undefined): string {
    if (!uri) {
        return "<none>";
    }

    const query = uri.query ? ` query=${uri.query}` : "";
    const fragment = uri.fragment ? ` fragment=${uri.fragment}` : "";
    const fsPath = uri.scheme === "file" ? ` fsPath=${uri.fsPath}` : "";
    return `${uri.toString()} [scheme=${uri.scheme} path=${uri.path}${fsPath}${query}${fragment}]`;
}
