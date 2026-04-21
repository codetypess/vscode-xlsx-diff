import * as path from "node:path";
import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../displayLanguage";
import {
    getScmWorkbookResourceDetail,
    getScmWorkbookResourceInfo,
    getScmWorkbookResourceTimeLabel,
    type WorkbookResourceDetail,
} from "../scm/resourceInfo";

export { describeGitResourceRef } from "../git/resourceInfo";

function getUriPathForExtension(uri: vscode.Uri): string {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (scmInfo) {
        return scmInfo.resourcePath;
    }

    return uri.scheme === "file" ? uri.fsPath : decodeURIComponent(uri.path);
}

export function isWorkbookResourceUri(uri: vscode.Uri | undefined): uri is vscode.Uri {
    if (!uri) {
        return false;
    }

    const resourcePath = getUriPathForExtension(uri);
    const normalizedPath = resourcePath.toLowerCase().endsWith(".git")
        ? resourcePath.slice(0, -".git".length)
        : resourcePath;
    return path.extname(normalizedPath).toLowerCase() === ".xlsx";
}

export function getWorkbookResourceName(uri: vscode.Uri): string {
    return path.basename(getUriPathForExtension(uri));
}

export function getWorkbookResourcePathLabel(uri: vscode.Uri): string {
    const resourcePath = getUriPathForExtension(uri);
    const scmInfo = getScmWorkbookResourceInfo(uri);
    return scmInfo?.ref ? `${resourcePath} (${scmInfo.ref})` : resourcePath;
}

export function getWorkbookResourceTimeLabel(uri: vscode.Uri): string | undefined {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (scmInfo) {
        const scmTimeLabel = getScmWorkbookResourceTimeLabel(scmInfo);
        if (scmTimeLabel) {
            return scmTimeLabel;
        }
    }

    return uri.scheme === "file"
        ? undefined
        : isChineseDisplayLanguage()
          ? `${uri.scheme.toUpperCase()} 资源`
          : `${uri.scheme.toUpperCase()} resource`;
}

export async function getWorkbookResourceDetail(
    uri: vscode.Uri
): Promise<WorkbookResourceDetail | undefined> {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (!scmInfo) {
        return undefined;
    }

    return getScmWorkbookResourceDetail(scmInfo);
}

export function isWorkbookResourceReadOnly(uri: vscode.Uri): boolean {
    const isWritable = vscode.workspace.fs.isWritableFileSystem(uri.scheme);
    return isWritable === false || (isWritable === undefined && uri.scheme !== "file");
}

export function getWorkbookDiffUrisFromTabInput(
    input: unknown
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    if (!(input instanceof vscode.TabInputTextDiff)) {
        return undefined;
    }

    if (!isWorkbookResourceUri(input.original) || !isWorkbookResourceUri(input.modified)) {
        return undefined;
    }

    return {
        original: input.original,
        modified: input.modified,
    };
}

export function getScmWorkbookDiffUrisFromTabInput(
    input: unknown
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    const diffUris = getWorkbookDiffUrisFromTabInput(input);
    if (!diffUris) {
        return undefined;
    }

    if (diffUris.original.scheme === "file" && diffUris.modified.scheme === "file") {
        return undefined;
    }

    return diffUris;
}

function normalizeResourcePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function getWorkbookResourcePathKey(uri: vscode.Uri): string | undefined {
    if (!isWorkbookResourceUri(uri)) {
        return undefined;
    }

    return normalizeResourcePathForComparison(getUriPathForExtension(uri));
}

function getScmWorkbookResourceRef(uri: vscode.Uri): string | undefined {
    return getScmWorkbookResourceInfo(uri)?.ref;
}

export function getScmWorkbookDiffUrisFromEditorUris(
    firstUri: vscode.Uri,
    secondUri: vscode.Uri
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    if (firstUri.toString() === secondUri.toString()) {
        return undefined;
    }

    const firstPathKey = getWorkbookResourcePathKey(firstUri);
    const secondPathKey = getWorkbookResourcePathKey(secondUri);
    if (!firstPathKey || firstPathKey !== secondPathKey) {
        return undefined;
    }

    if (firstUri.scheme === "file" && secondUri.scheme === "file") {
        return undefined;
    }

    if (firstUri.scheme === "file") {
        return { original: secondUri, modified: firstUri };
    }

    if (secondUri.scheme === "file") {
        return { original: firstUri, modified: secondUri };
    }

    const firstRef = getScmWorkbookResourceRef(firstUri);
    const secondRef = getScmWorkbookResourceRef(secondUri);
    if (firstRef === "" && secondRef !== "") {
        return { original: secondUri, modified: firstUri };
    }

    if (secondRef === "" && firstRef !== "") {
        return { original: firstUri, modified: secondUri };
    }

    return { original: firstUri, modified: secondUri };
}
