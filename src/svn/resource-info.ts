import { execFile } from "node:child_process";
import * as path from "node:path";
import * as vscode from "vscode";
import { formatI18nMessage, getRuntimeMessages } from "../i18n";
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

type SvnGraphSource = "empty" | "svn";

interface SvnGraphUriDescriptor {
    label?: string;
    source?: SvnGraphSource;
    target?: string;
    revision?: string;
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

function parseSvnGraphUriDescriptor(uri: vscode.Uri): SvnGraphUriDescriptor | undefined {
    if (uri.scheme !== "svn-graph") {
        return undefined;
    }

    const params = new URLSearchParams(uri.query);
    return {
        label: params.get("label") ?? undefined,
        source: (params.get("source") as SvnGraphSource | null) ?? "svn",
        target: params.get("target") ?? undefined,
        revision: params.get("revision") ?? undefined,
    };
}

function createSourceLabel(): string {
    return getRuntimeMessages().scm.sourceLabel;
}

function createSvnRefLabel(ref: string): string {
    return formatI18nMessage(getRuntimeMessages().scm.svnRefLabel, { ref });
}

function createEmptyWorkbookLabel(): string {
    return getRuntimeMessages().scm.emptyWorkbookLabel;
}

function createSvnGraphRef(revision: string): string {
    return /^\d+$/.test(revision) ? `r${revision}` : revision;
}

function normalizePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function getSvnGraphLabel(descriptor: SvnGraphUriDescriptor, uri: vscode.Uri): string {
    const label = descriptor.label ?? decodeURIComponent(uri.path);
    return label.startsWith("/") ? label.slice(1) : label;
}

function stripWorkbookRefSuffix(label: string): string {
    const match = /^(.+\.xlsx)(?: \([^()]+\))$/i.exec(label);
    return match?.[1] ?? label;
}

function getSvnGraphTargetPath(target: string | undefined): string | undefined {
    if (!target) {
        return undefined;
    }

    if (path.isAbsolute(target)) {
        return target;
    }

    try {
        return decodeURIComponent(new URL(target).pathname);
    } catch {
        return target;
    }
}

function createComparisonPaths(
    resourcePath: string,
    displayPath: string,
    targetPath: string | undefined
): string[] | undefined {
    const comparisonPaths = [displayPath, targetPath]
        .filter((value): value is string => typeof value === "string" && value.length > 0)
        .filter((value) => value !== resourcePath);

    if (comparisonPaths.length === 0) {
        return undefined;
    }

    return [...new Set(comparisonPaths)];
}

function runSvnCat(target: string, revision: string): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
        execFile(
            "svn",
            ["cat", "-r", revision, target],
            {
                cwd: path.isAbsolute(target) ? path.dirname(target) : undefined,
                encoding: "buffer",
                maxBuffer: 64 * 1024 * 1024,
            },
            (error, stdout, stderr) => {
                if (error) {
                    const stderrText =
                        stderr instanceof Buffer ? stderr.toString("utf8").trim() : "";
                    reject(new Error(stderrText || error.message));
                    return;
                }

                const archive =
                    stdout instanceof Buffer ? stdout : Buffer.from(stdout ?? "");
                resolve(new Uint8Array(archive));
            }
        );
    });
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

export function getSvnGraphWorkbookResourceInfo(
    uri: vscode.Uri
): ScmWorkbookResourceInfo | undefined {
    const descriptor = parseSvnGraphUriDescriptor(uri);
    if (!descriptor) {
        return undefined;
    }

    const displayPath = stripWorkbookRefSuffix(getSvnGraphLabel(descriptor, uri));
    const targetPath = getSvnGraphTargetPath(descriptor.target);
    const resourcePath =
        targetPath && path.extname(targetPath).toLowerCase() === ".xlsx" ? targetPath : displayPath;

    if (
        path.extname(displayPath).toLowerCase() !== ".xlsx" &&
        path.extname(resourcePath).toLowerCase() !== ".xlsx"
    ) {
        return undefined;
    }

    return {
        provider: "svn-graph",
        uri,
        resourcePath,
        displayPath,
        comparisonPaths: createComparisonPaths(resourcePath, displayPath, targetPath),
        ref:
            descriptor.source === "svn" && descriptor.revision
                ? createSvnGraphRef(descriptor.revision)
                : undefined,
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
        label: createSourceLabel(),
        value: createSvnRefLabel(info.ref),
        titleValue: info.ref,
    };
}

export function getSvnGraphWorkbookResourceTimeLabel(
    info: ScmWorkbookResourceInfo
): string | undefined {
    if (info.provider !== "svn-graph") {
        return undefined;
    }

    const descriptor = parseSvnGraphUriDescriptor(info.uri);
    if (descriptor?.source === "empty") {
        return createEmptyWorkbookLabel();
    }

    if (info.ref === undefined) {
        return undefined;
    }

    return createSvnRefLabel(info.ref);
}

export function getSvnGraphWorkbookResourceDetail(
    info: ScmWorkbookResourceInfo
): WorkbookResourceDetail | undefined {
    if (info.provider !== "svn-graph") {
        return undefined;
    }

    const descriptor = parseSvnGraphUriDescriptor(info.uri);
    if (descriptor?.source === "empty") {
        return {
            label: createSourceLabel(),
            value: createEmptyWorkbookLabel(),
        };
    }

    if (info.ref === undefined) {
        return undefined;
    }

    return {
        label: createSourceLabel(),
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

export const svnGraphWorkbookResourceProvider: ScmWorkbookResourceProvider = {
    scheme: "svn-graph",
    getResourceInfo: getSvnGraphWorkbookResourceInfo,
    getResourceTimeLabel: getSvnGraphWorkbookResourceTimeLabel,
    getResourceDetail: async (info) => getSvnGraphWorkbookResourceDetail(info),
    readWorkbookArchive: async (info) => {
        if (info.provider !== "svn-graph") {
            return undefined;
        }

        const descriptor = parseSvnGraphUriDescriptor(info.uri);
        if (
            descriptor?.source !== "svn" ||
            !descriptor.target ||
            !descriptor.revision
        ) {
            return undefined;
        }

        return runSvnCat(descriptor.target, descriptor.revision);
    },
    isEmptyWorkbook: (info) => {
        if (info.provider !== "svn-graph") {
            return false;
        }

        return parseSvnGraphUriDescriptor(info.uri)?.source === "empty";
    },
};
