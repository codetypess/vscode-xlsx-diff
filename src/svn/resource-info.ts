import * as path from "node:path";
import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import type {
    WorkbookDiffUris,
    ScmWorkbookResourceInfo,
    ScmWorkbookResourceProvider,
    WorkbookResourceDetail,
} from "../scm/resource-info";

interface SvnUriQuery {
    action?: string;
    fsPath?: string;
    extra?: {
        ref?: string | number;
    };
}

function parseSvnUriQuery(uri: vscode.Uri): SvnUriQuery | undefined {
    if (uri.scheme !== "svn" || !uri.query) {
        return undefined;
    }

    try {
        const parsed = JSON.parse(uri.query) as SvnUriQuery;
        return typeof parsed === "object" && parsed !== null ? parsed : undefined;
    } catch {
        return undefined;
    }
}

function getSvnRef(query: SvnUriQuery): string | undefined {
    const ref = query.extra?.ref;
    if (typeof ref === "string") {
        return ref;
    }

    if (typeof ref === "number") {
        return String(ref);
    }

    return undefined;
}

function createSvnRefLabel(ref: string): string {
    return isChineseDisplayLanguage() ? `SVN 引用: ${ref}` : `SVN ref: ${ref}`;
}

function normalizePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function withSvnRef(uri: vscode.Uri, ref: string): vscode.Uri {
    const query = parseSvnUriQuery(uri);
    if (!query) {
        return uri;
    }

    return uri.with({
        query: JSON.stringify({
            ...query,
            extra: {
                ...query.extra,
                ref,
            },
        }),
    });
}

export function getSvnWorkbookResourceInfo(
    uri: vscode.Uri
): ScmWorkbookResourceInfo | undefined {
    const query = parseSvnUriQuery(uri);
    if (query?.action !== "SHOW" || typeof query.fsPath !== "string") {
        return undefined;
    }

    return {
        provider: "svn",
        uri,
        resourcePath: query.fsPath,
        ref: getSvnRef(query),
    };
}

export function getSvnWorkbookResourceTimeLabel(
    info: ScmWorkbookResourceInfo
): string | undefined {
    if (info.provider !== "svn" || info.ref === undefined) {
        return undefined;
    }

    return createSvnRefLabel(info.ref);
}

export function getSvnWorkbookResourceDetail(
    info: ScmWorkbookResourceInfo
): WorkbookResourceDetail | undefined {
    if (info.provider !== "svn" || info.ref === undefined) {
        return undefined;
    }

    return {
        label: isChineseDisplayLanguage() ? "来源" : "Source",
        value: createSvnRefLabel(info.ref),
        titleValue: info.ref,
    };
}

function normalizeSvnDiffUris(diffUris: WorkbookDiffUris): WorkbookDiffUris {
    const originalInfo = getSvnWorkbookResourceInfo(diffUris.original);
    if (!originalInfo || originalInfo.ref?.toUpperCase() !== "HEAD") {
        return diffUris;
    }

    if (diffUris.modified.scheme !== "file") {
        return diffUris;
    }

    if (
        normalizePathForComparison(originalInfo.resourcePath) !==
        normalizePathForComparison(diffUris.modified.fsPath)
    ) {
        return diffUris;
    }

    return {
        original: withSvnRef(diffUris.original, "BASE"),
        modified: diffUris.modified,
    };
}

export const svnWorkbookResourceProvider: ScmWorkbookResourceProvider = {
    scheme: "svn",
    getResourceInfo: getSvnWorkbookResourceInfo,
    getResourceTimeLabel: getSvnWorkbookResourceTimeLabel,
    getResourceDetail: async (info) => getSvnWorkbookResourceDetail(info),
    normalizeDiffUris: normalizeSvnDiffUris,
};
