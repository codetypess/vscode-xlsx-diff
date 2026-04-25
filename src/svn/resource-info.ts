import { execFile as execFileCallback } from "node:child_process";
import * as path from "node:path";
import { promisify } from "node:util";
import * as vscode from "vscode";
import { formatI18nMessage, getRuntimeMessages } from "../i18n";
import type {
    WorkbookDiffUris,
    ScmWorkbookResourceInfo,
    ScmWorkbookResourceProvider,
    WorkbookResourceFact,
    WorkbookResourceDetail,
} from "../scm/resource-info";

interface SvnUriQuery {
    action?: string;
    fsPath?: string;
    extra?: {
        ref?: string | number;
    };
}

type SvnTreeSource = "empty" | "svn";

interface SvnTreeUriDescriptor {
    label?: string;
    source?: SvnTreeSource;
    target?: string;
    revision?: string;
}

const execFile = promisify(execFileCallback);

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

function isSvnTreeScheme(scheme: string): scheme is "svn-tree" {
    return scheme === "svn-tree";
}

function parseSvnTreeUriDescriptor(uri: vscode.Uri): SvnTreeUriDescriptor | undefined {
    if (!isSvnTreeScheme(uri.scheme)) {
        return undefined;
    }

    const params = new URLSearchParams(uri.query);
    return {
        label: params.get("label") ?? undefined,
        source: (params.get("source") as SvnTreeSource | null) ?? "svn",
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

function createAuthorLabel(): string {
    return getRuntimeMessages().scm.authorLabel;
}

function createSvnTreeRef(revision: string): string {
    return /^\d+$/.test(revision) ? `r${revision}` : revision;
}

function decodeXmlEntities(value: string): string {
    return value
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&");
}

function extractSvnAuthor(xml: string): string | undefined {
    const match = /<author>([\s\S]*?)<\/author>/.exec(xml);
    const author = match?.[1]?.trim();
    return author ? decodeXmlEntities(author) : undefined;
}

function normalizePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function getSvnTreeLabel(descriptor: SvnTreeUriDescriptor, uri: vscode.Uri): string {
    const label = descriptor.label ?? decodeURIComponent(uri.path);
    return label.startsWith("/") ? label.slice(1) : label;
}

function stripWorkbookRefSuffix(label: string): string {
    const match = /^(.+\.xlsx)(?: \([^()]+\))$/i.exec(label);
    return match?.[1] ?? label;
}

function getSvnTreeTargetPath(target: string | undefined): string | undefined {
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

async function runSvn(args: string[], cwd: string | undefined): Promise<string | undefined> {
    try {
        const { stdout } = await execFile("svn", args, {
            cwd,
            maxBuffer: 16 * 1024 * 1024,
        });
        const trimmed = stdout.trim();
        return trimmed.length > 0 ? trimmed : undefined;
    } catch {
        return undefined;
    }
}

function getSvnTargetCwd(target: string): string | undefined {
    return path.isAbsolute(target) ? path.dirname(target) : undefined;
}

async function resolveSvnAuthor(
    target: string | undefined,
    revision: string | undefined
): Promise<string | undefined> {
    if (!target || !revision) {
        return undefined;
    }

    const infoXml = await runSvn(["info", "--xml", "-r", revision, target], getSvnTargetCwd(target));
    return infoXml ? extractSvnAuthor(infoXml) : undefined;
}

function createSvnExtraFacts(author: string | undefined): WorkbookResourceFact[] | undefined {
    if (!author) {
        return undefined;
    }

    return [
        {
            label: createAuthorLabel(),
            value: author,
        },
    ];
}

function runSvnCat(target: string, revision: string): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
        execFileCallback(
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

export function getSvnTreeWorkbookResourceInfo(
    uri: vscode.Uri
): ScmWorkbookResourceInfo | undefined {
    const descriptor = parseSvnTreeUriDescriptor(uri);
    if (!descriptor) {
        return undefined;
    }

    const displayPath = stripWorkbookRefSuffix(getSvnTreeLabel(descriptor, uri));
    const targetPath = getSvnTreeTargetPath(descriptor.target);
    const resourcePath =
        targetPath && path.extname(targetPath).toLowerCase() === ".xlsx" ? targetPath : displayPath;

    if (
        path.extname(displayPath).toLowerCase() !== ".xlsx" &&
        path.extname(resourcePath).toLowerCase() !== ".xlsx"
    ) {
        return undefined;
    }

    return {
        provider: "svn-tree",
        uri,
        resourcePath,
        displayPath,
        comparisonPaths: createComparisonPaths(resourcePath, displayPath, targetPath),
        ref:
            descriptor.source === "svn" && descriptor.revision
                ? createSvnTreeRef(descriptor.revision)
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
): Promise<WorkbookResourceDetail | undefined> {
    if (info.provider !== "svn" || info.ref === undefined) {
        return Promise.resolve(undefined);
    }

    const ref = info.ref;
    return resolveSvnAuthor(info.resourcePath, ref).then((author) => ({
        label: createSourceLabel(),
        value: createSvnRefLabel(ref),
        titleValue: ref,
        extraFacts: createSvnExtraFacts(author),
    }));
}

export function getSvnTreeWorkbookResourceTimeLabel(
    info: ScmWorkbookResourceInfo
): string | undefined {
    if (info.provider !== "svn-tree") {
        return undefined;
    }

    const descriptor = parseSvnTreeUriDescriptor(info.uri);
    if (descriptor?.source === "empty") {
        return createEmptyWorkbookLabel();
    }

    if (info.ref === undefined) {
        return undefined;
    }

    return createSvnRefLabel(info.ref);
}

export function getSvnTreeWorkbookResourceDetail(
    info: ScmWorkbookResourceInfo
): Promise<WorkbookResourceDetail | undefined> {
    if (info.provider !== "svn-tree") {
        return Promise.resolve(undefined);
    }

    const descriptor = parseSvnTreeUriDescriptor(info.uri);
    if (descriptor?.source === "empty") {
        return Promise.resolve({
            label: createSourceLabel(),
            value: createEmptyWorkbookLabel(),
        });
    }

    if (info.ref === undefined) {
        return Promise.resolve(undefined);
    }

    const ref = info.ref;
    return resolveSvnAuthor(descriptor?.target, descriptor?.revision).then((author) => ({
        label: createSourceLabel(),
        value: createSvnRefLabel(ref),
        titleValue: ref,
        extraFacts: createSvnExtraFacts(author),
    }));
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
    getResourceDetail: getSvnWorkbookResourceDetail,
    normalizeDiffUris: normalizeSvnDiffUris,
};

export const svnTreeWorkbookResourceProvider: ScmWorkbookResourceProvider = {
    scheme: "svn-tree",
    getResourceInfo: getSvnTreeWorkbookResourceInfo,
    getResourceTimeLabel: getSvnTreeWorkbookResourceTimeLabel,
    getResourceDetail: getSvnTreeWorkbookResourceDetail,
    readWorkbookArchive: async (info) => {
        if (info.provider !== "svn-tree") {
            return undefined;
        }

        const descriptor = parseSvnTreeUriDescriptor(info.uri);
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
        if (info.provider !== "svn-tree") {
            return false;
        }

        return parseSvnTreeUriDescriptor(info.uri)?.source === "empty";
    },
};
