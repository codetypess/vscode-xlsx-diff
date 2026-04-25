import * as path from "node:path";
import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import {
    getScmWorkbookResourceDetail,
    getScmWorkbookResourceInfo,
    getScmWorkbookResourceTimeLabel,
    hasScmWorkbookResourceProvider,
    isEmptyScmWorkbook,
    normalizeScmWorkbookDiffUris,
    readScmWorkbookArchive,
    type WorkbookResourceDetail,
} from "../scm/resource-info";

export { describeGitResourceRef } from "../git/resource-info";

function getUriDisplayPath(uri: vscode.Uri): string {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (scmInfo) {
        return scmInfo.displayPath ?? scmInfo.resourcePath;
    }

    return uri.scheme === "file" ? uri.fsPath : decodeURIComponent(uri.path);
}

function getUriComparisonPaths(uri: vscode.Uri): string[] {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (scmInfo) {
        return [
            scmInfo.resourcePath,
            ...(scmInfo.comparisonPaths ?? []),
        ];
    }

    return [uri.scheme === "file" ? uri.fsPath : decodeURIComponent(uri.path)];
}

export function isWorkbookResourceUri(uri: vscode.Uri | undefined): uri is vscode.Uri {
    if (!uri) {
        return false;
    }

    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (!scmInfo && hasScmWorkbookResourceProvider(uri.scheme)) {
        return false;
    }

    const resourcePath = scmInfo?.displayPath ?? scmInfo?.resourcePath ?? getUriDisplayPath(uri);
    const normalizedPath = resourcePath.toLowerCase().endsWith(".git")
        ? resourcePath.slice(0, -".git".length)
        : resourcePath;
    return path.extname(normalizedPath).toLowerCase() === ".xlsx";
}

export function getWorkbookResourceName(uri: vscode.Uri): string {
    return path.basename(getUriDisplayPath(uri));
}

export function getWorkbookResourcePathLabel(uri: vscode.Uri): string {
    const resourcePath = getUriDisplayPath(uri);
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

export function isEmptyWorkbookResourceUri(uri: vscode.Uri): boolean {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    return scmInfo ? isEmptyScmWorkbook(scmInfo) : false;
}

export async function readWorkbookResourceArchive(
    uri: vscode.Uri
): Promise<Uint8Array | undefined> {
    const scmInfo = getScmWorkbookResourceInfo(uri);
    if (!scmInfo) {
        return undefined;
    }

    return readScmWorkbookArchive(scmInfo);
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

    return normalizeScmWorkbookDiffUris(diffUris);
}

function normalizeResourcePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function getWorkbookResourceComparisonSet(uri: vscode.Uri): Set<string> | undefined {
    if (!isWorkbookResourceUri(uri)) {
        return undefined;
    }

    return new Set(
        getUriComparisonPaths(uri).map((resourcePath) =>
            normalizeResourcePathForComparison(resourcePath)
        )
    );
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

    const firstPathSet = getWorkbookResourceComparisonSet(firstUri);
    const secondPathSet = getWorkbookResourceComparisonSet(secondUri);
    if (!firstPathSet || !secondPathSet) {
        return undefined;
    }

    const hasSharedPath = [...firstPathSet].some((resourcePath) => secondPathSet.has(resourcePath));
    if (!hasSharedPath) {
        return undefined;
    }

    if (firstUri.scheme === "file" && secondUri.scheme === "file") {
        return undefined;
    }

    if (firstUri.scheme === "file") {
        return normalizeScmWorkbookDiffUris({ original: secondUri, modified: firstUri });
    }

    if (secondUri.scheme === "file") {
        return normalizeScmWorkbookDiffUris({ original: firstUri, modified: secondUri });
    }

    const firstRef = getScmWorkbookResourceRef(firstUri);
    const secondRef = getScmWorkbookResourceRef(secondUri);
    if (firstRef === "" && secondRef !== "") {
        return normalizeScmWorkbookDiffUris({ original: secondUri, modified: firstUri });
    }

    if (secondRef === "" && firstRef !== "") {
        return normalizeScmWorkbookDiffUris({ original: firstUri, modified: secondUri });
    }

    return normalizeScmWorkbookDiffUris({ original: firstUri, modified: secondUri });
}
