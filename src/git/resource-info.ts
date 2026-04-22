import { execFile as execFileCallback } from "node:child_process";
import * as path from "node:path";
import { promisify } from "node:util";
import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import type {
    ScmWorkbookResourceInfo,
    ScmWorkbookResourceProvider,
    WorkbookResourceDetail,
} from "../scm/resource-info";

interface GitUriQuery {
    path?: string;
    ref?: string;
}

const execFile = promisify(execFileCallback);

export function parseGitUriQuery(uri: vscode.Uri): GitUriQuery | undefined {
    if (uri.scheme !== "git" || !uri.query) {
        return undefined;
    }

    try {
        const parsed = JSON.parse(uri.query) as GitUriQuery;
        return typeof parsed === "object" && parsed !== null ? parsed : undefined;
    } catch {
        return undefined;
    }
}

async function runGit(cwd: string, args: string[]): Promise<string | undefined> {
    try {
        const { stdout } = await execFile("git", ["-C", cwd, ...args]);
        const trimmed = stdout.trim();
        return trimmed.length > 0 ? trimmed : undefined;
    } catch {
        return undefined;
    }
}

async function getGitRepositoryRoot(resourcePath: string): Promise<string | undefined> {
    return runGit(path.dirname(resourcePath), ["rev-parse", "--show-toplevel"]);
}

async function hasStagedChanges(repositoryRoot: string, resourcePath: string): Promise<boolean> {
    const relativePath = path.relative(repositoryRoot, resourcePath);
    const changedFiles = await runGit(repositoryRoot, [
        "diff",
        "--cached",
        "--name-only",
        "--",
        relativePath,
    ]);
    return Boolean(changedFiles);
}

async function resolveShortCommit(
    repositoryRoot: string,
    ref: string
): Promise<string | undefined> {
    return runGit(repositoryRoot, ["rev-parse", "--short", ref]);
}

function createGitResourcePresentation(
    ref: string,
    options: {
        resolvedCommit?: string;
        hasStagedChanges?: boolean;
    } = {}
): WorkbookResourceDetail {
    const isChinese = isChineseDisplayLanguage();
    const sourceLabel = isChinese ? "来源" : "Source";
    const commitLabel = isChinese ? "提交" : "Commit";
    const indexLabel = isChinese ? "暂存区" : "Index";

    if (ref === "") {
        return { label: sourceLabel, value: indexLabel };
    }

    if (/^~\d$/.test(ref)) {
        return {
            label: sourceLabel,
            value: isChinese ? `阶段 ${ref[1]}` : `Stage ${ref[1]}`,
        };
    }

    if (ref === "~") {
        if (options.hasStagedChanges) {
            return {
                label: sourceLabel,
                value: options.resolvedCommit
                    ? isChinese
                        ? `暂存区 · 基线 ${options.resolvedCommit}`
                        : `Index · base ${options.resolvedCommit}`
                    : indexLabel,
                titleValue: options.resolvedCommit,
            };
        }

        return {
            label: commitLabel,
            value: options.resolvedCommit ?? "HEAD",
            titleValue: options.resolvedCommit,
        };
    }

    if (options.resolvedCommit) {
        return {
            label: commitLabel,
            value: options.resolvedCommit,
            titleValue: options.resolvedCommit,
        };
    }

    return {
        label: sourceLabel,
        value: isChinese ? `Git 引用: ${ref}` : `Git ref: ${ref}`,
    };
}

export function describeGitResourceRef(
    ref: string,
    options: {
        resolvedCommit?: string;
        hasStagedChanges?: boolean;
    } = {}
): { label: string; value: string } {
    const { label, value } = createGitResourcePresentation(ref, options);
    return { label, value };
}

export function getGitWorkbookResourceInfo(
    uri: vscode.Uri
): ScmWorkbookResourceInfo | undefined {
    if (uri.scheme !== "git") {
        return undefined;
    }

    const gitQuery = parseGitUriQuery(uri);
    return {
        provider: "git",
        uri,
        resourcePath: gitQuery?.path ?? decodeURIComponent(uri.path),
        ref: gitQuery?.ref,
    };
}

export function getGitWorkbookResourceTimeLabel(
    info: ScmWorkbookResourceInfo
): string | undefined {
    if (info.provider !== "git" || info.ref === undefined) {
        return undefined;
    }

    return isChineseDisplayLanguage() ? `Git 引用: ${info.ref}` : `Git ref: ${info.ref}`;
}

export async function getGitWorkbookResourceDetail(
    info: ScmWorkbookResourceInfo
): Promise<WorkbookResourceDetail | undefined> {
    if (info.provider !== "git" || info.ref === undefined) {
        return undefined;
    }

    const repositoryRoot = await getGitRepositoryRoot(info.resourcePath);
    if (!repositoryRoot) {
        return describeGitResourceRef(info.ref);
    }

    if (info.ref === "~") {
        const [resolvedCommit, stagedChanges] = await Promise.all([
            resolveShortCommit(repositoryRoot, "HEAD"),
            hasStagedChanges(repositoryRoot, info.resourcePath),
        ]);
        return createGitResourcePresentation(info.ref, {
            resolvedCommit,
            hasStagedChanges: stagedChanges,
        });
    }

    if (info.ref === "" || /^~\d$/.test(info.ref)) {
        return createGitResourcePresentation(info.ref);
    }

    const resolvedCommit = await resolveShortCommit(repositoryRoot, info.ref);
    return createGitResourcePresentation(info.ref, { resolvedCommit });
}

export const gitWorkbookResourceProvider: ScmWorkbookResourceProvider = {
    scheme: "git",
    getResourceInfo: getGitWorkbookResourceInfo,
    getResourceTimeLabel: getGitWorkbookResourceTimeLabel,
    getResourceDetail: getGitWorkbookResourceDetail,
};
