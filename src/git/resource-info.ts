import { execFile as execFileCallback } from "node:child_process";
import { realpath } from "node:fs/promises";
import * as path from "node:path";
import { promisify } from "node:util";
import * as vscode from "vscode";
import { formatI18nMessage, getRuntimeMessages } from "../i18n";
import type {
    WorkbookResourceFact,
    ScmWorkbookResourceInfo,
    ScmWorkbookResourceProvider,
    WorkbookResourceDetail,
} from "../scm/resource-info";

interface GitUriQuery {
    path?: string;
    ref?: string;
}

interface GitCommitMetadata {
    shortCommit?: string;
    committer?: string;
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

function toGitRelativePath(repositoryRoot: string, resourcePath: string): string {
    return path.relative(repositoryRoot, resourcePath).split(path.sep).join("/");
}

function createGitBlobSpecifier(relativePath: string, ref: string): string {
    if (ref === "") {
        return `:${relativePath}`;
    }

    const stageMatch = /^~(\d)$/.exec(ref);
    if (stageMatch) {
        return `:${stageMatch[1]}:${relativePath}`;
    }

    const normalizedRef = ref === "~" ? "HEAD" : ref;
    return `${normalizedRef}:${relativePath}`;
}

async function readGitBlob(
    repositoryRoot: string,
    resourcePath: string,
    ref: string
): Promise<Uint8Array | undefined> {
    const [normalizedRepositoryRoot, normalizedResourcePath] = await Promise.all([
        realpath(repositoryRoot).catch(() => repositoryRoot),
        realpath(resourcePath).catch(() => resourcePath),
    ]);
    const relativePath = toGitRelativePath(normalizedRepositoryRoot, normalizedResourcePath);
    const blobSpecifier = createGitBlobSpecifier(relativePath, ref);

    try {
        const { stdout } = await execFile("git", ["-C", repositoryRoot, "show", blobSpecifier], {
            encoding: "buffer",
            maxBuffer: 64 * 1024 * 1024,
        });
        const archive = stdout instanceof Buffer ? stdout : Buffer.from(stdout ?? "");
        return new Uint8Array(archive);
    } catch {
        return undefined;
    }
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

function formatGitCommitter(
    name: string | undefined,
    email: string | undefined
): string | undefined {
    const trimmedName = name?.trim();
    const trimmedEmail = email?.trim();
    if (!trimmedName && !trimmedEmail) {
        return undefined;
    }

    if (!trimmedName) {
        return trimmedEmail;
    }

    return trimmedEmail ? `${trimmedName} <${trimmedEmail}>` : trimmedName;
}

async function resolveGitCommitMetadata(
    repositoryRoot: string,
    ref: string
): Promise<GitCommitMetadata | undefined> {
    const output = await runGit(repositoryRoot, ["show", "-s", "--format=%h%x00%cn%x00%ce", ref]);
    if (!output) {
        return undefined;
    }

    const [shortCommit, committerName, committerEmail] = output.split("\u0000");
    const metadata: GitCommitMetadata = {
        shortCommit: shortCommit?.trim() || undefined,
        committer: formatGitCommitter(committerName, committerEmail),
    };

    return metadata.shortCommit || metadata.committer ? metadata : undefined;
}

function createGitExtraFacts(committer: string | undefined): WorkbookResourceFact[] | undefined {
    if (!committer) {
        return undefined;
    }

    return [
        {
            label: getRuntimeMessages().scm.committerLabel,
            value: committer,
        },
    ];
}

function createGitResourcePresentation(
    ref: string,
    options: {
        resolvedCommit?: string;
        hasStagedChanges?: boolean;
        committer?: string;
    } = {}
): WorkbookResourceDetail {
    const { scm } = getRuntimeMessages();
    const extraFacts = createGitExtraFacts(options.committer);

    if (ref === "") {
        return { label: scm.sourceLabel, value: scm.indexLabel };
    }

    if (/^~\d$/.test(ref)) {
        return {
            label: scm.sourceLabel,
            value: formatI18nMessage(scm.stageLabel, { stage: ref[1] }),
        };
    }

    if (ref === "~") {
        if (options.hasStagedChanges) {
            return {
                label: scm.sourceLabel,
                value: options.resolvedCommit
                    ? formatI18nMessage(scm.indexBaseLabel, {
                          commit: options.resolvedCommit,
                      })
                    : scm.indexLabel,
                titleValue: options.resolvedCommit,
                extraFacts,
            };
        }

        return {
            label: scm.commitLabel,
            value: options.resolvedCommit ?? "HEAD",
            titleValue: options.resolvedCommit,
            extraFacts,
        };
    }

    if (options.resolvedCommit) {
        return {
            label: scm.commitLabel,
            value: options.resolvedCommit,
            titleValue: options.resolvedCommit,
            extraFacts,
        };
    }

    return {
        label: scm.sourceLabel,
        value: formatI18nMessage(scm.gitRefLabel, { ref }),
        extraFacts,
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

export function getGitWorkbookResourceInfo(uri: vscode.Uri): ScmWorkbookResourceInfo | undefined {
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

export function getGitWorkbookResourceTimeLabel(info: ScmWorkbookResourceInfo): string | undefined {
    if (info.provider !== "git" || info.ref === undefined) {
        return undefined;
    }

    return formatI18nMessage(getRuntimeMessages().scm.gitRefLabel, { ref: info.ref });
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
        const [commitMetadata, stagedChanges] = await Promise.all([
            resolveGitCommitMetadata(repositoryRoot, "HEAD"),
            hasStagedChanges(repositoryRoot, info.resourcePath),
        ]);
        return createGitResourcePresentation(info.ref, {
            resolvedCommit: commitMetadata?.shortCommit,
            hasStagedChanges: stagedChanges,
            committer: commitMetadata?.committer,
        });
    }

    if (info.ref === "" || /^~\d$/.test(info.ref)) {
        return createGitResourcePresentation(info.ref);
    }

    const commitMetadata = await resolveGitCommitMetadata(repositoryRoot, info.ref);
    return createGitResourcePresentation(info.ref, {
        resolvedCommit: commitMetadata?.shortCommit,
        committer: commitMetadata?.committer,
    });
}

export const gitWorkbookResourceProvider: ScmWorkbookResourceProvider = {
    scheme: "git",
    getResourceInfo: getGitWorkbookResourceInfo,
    getResourceTimeLabel: getGitWorkbookResourceTimeLabel,
    getResourceDetail: getGitWorkbookResourceDetail,
    readWorkbookArchive: async (info) => {
        if (info.provider !== "git" || info.ref === undefined) {
            return undefined;
        }

        const repositoryRoot = await getGitRepositoryRoot(info.resourcePath);
        if (!repositoryRoot) {
            return undefined;
        }

        return readGitBlob(repositoryRoot, info.resourcePath, info.ref);
    },
};
