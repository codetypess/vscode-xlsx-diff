import * as path from "node:path";
import * as vscode from "vscode";
import { getRecentWorkbookResourceEntries } from "../scm/recentWorkbookResourceContext";
import { isUnknownTabInput } from "../scm/tabInput";
import { describeUri } from "../scm/uriDescription";

interface ParsedSvnTabLabel {
    basename: string;
    ref: string;
}

interface ParsedSvnShowUri {
    action?: string;
    fsPath?: string;
    ref?: string;
}

interface SvnResource {
    readonly resourceUri: vscode.Uri;
    readonly renameResourceUri?: vscode.Uri;
    readonly type?: string;
    readonly remote?: boolean;
}

interface SvnResourceGroup {
    readonly resourceStates: readonly SvnResource[];
}

interface SvnRepository {
    readonly root?: string;
    readonly workspaceRoot?: string;
    readonly changes?: SvnResourceGroup;
    readonly conflicts?: SvnResourceGroup;
    readonly remoteChanges?: SvnResourceGroup;
    readonly changelists?: Map<string, SvnResourceGroup>;
}

interface SvnSourceControlManager {
    readonly isInitialized?: Promise<void>;
    readonly repositories?: readonly SvnRepository[];
}

interface SvnDiffCandidate {
    readonly repository: SvnRepository;
    readonly group: string;
    readonly resource: SvnResource;
    readonly diffUris: {
        original: vscode.Uri;
        modified: vscode.Uri;
    };
}

function parseSvnTabLabel(label: string): ParsedSvnTabLabel | undefined {
    const match = /^(.+\.xlsx) \(([^()]+)\)$/i.exec(label);
    if (!match) {
        return undefined;
    }

    return {
        basename: match[1].toLowerCase(),
        ref: match[2].toUpperCase(),
    };
}

function getWorkbookBasename(uri: vscode.Uri): string {
    return path.basename(uri.fsPath || uri.path).toLowerCase();
}

function normalizePathForComparison(resourcePath: string): string {
    const normalizedPath = path.normalize(resourcePath);
    return process.platform === "win32" ? normalizedPath.toLowerCase() : normalizedPath;
}

function createSvnShowUri(uri: vscode.Uri, ref: string): vscode.Uri {
    return uri.with({
        scheme: "svn",
        query: JSON.stringify({
            action: "SHOW",
            fsPath: uri.fsPath,
            extra: {
                ref,
            },
        }),
    });
}

function parseSvnShowUri(uri: vscode.Uri): ParsedSvnShowUri | undefined {
    if (uri.scheme !== "svn" || !uri.query) {
        return undefined;
    }

    try {
        const parsed = JSON.parse(uri.query) as {
            action?: string;
            fsPath?: string;
            extra?: {
                ref?: string | number;
            };
        };
        return {
            action: parsed.action,
            fsPath: parsed.fsPath,
            ref:
                typeof parsed.extra?.ref === "number"
                    ? String(parsed.extra.ref)
                    : parsed.extra?.ref,
        };
    } catch {
        return undefined;
    }
}

async function getSvnSourceControlManager(): Promise<SvnSourceControlManager | undefined> {
    try {
        return await vscode.commands.executeCommand<SvnSourceControlManager>(
            "svn.getSourceControlManager"
        );
    } catch {
        return undefined;
    }
}

function getRepositoryResourceGroups(
    repository: SvnRepository
): Array<{ name: string; group: SvnResourceGroup }> {
    const groups: Array<{ name: string; group: SvnResourceGroup }> = [];

    if (repository.changes) {
        groups.push({ name: "changes", group: repository.changes });
    }

    if (repository.conflicts) {
        groups.push({ name: "conflicts", group: repository.conflicts });
    }

    if (repository.remoteChanges) {
        groups.push({ name: "remoteChanges", group: repository.remoteChanges });
    }

    for (const [name, group] of repository.changelists ?? []) {
        groups.push({ name: `changelist:${name}`, group });
    }

    return groups;
}

function getDiffUrisForSvnResource(
    group: string,
    resource: SvnResource
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    if (resource.resourceUri.scheme !== "file") {
        return undefined;
    }

    if (!getWorkbookBasename(resource.resourceUri).endsWith(".xlsx")) {
        return undefined;
    }

    if (group === "remoteChanges") {
        switch (resource.type) {
            case "modified":
            case "replaced":
            case "conflicted":
            case undefined:
                return {
                    original: resource.resourceUri,
                    modified: createSvnShowUri(resource.resourceUri, "HEAD"),
                };
            default:
                return undefined;
        }
    }

    switch (resource.type) {
        case "modified":
        case "replaced":
        case "conflicted":
        case undefined:
            return {
                original: createSvnShowUri(resource.resourceUri, "BASE"),
                modified: resource.resourceUri,
            };
        default:
            return undefined;
    }
}

function addUniqueUri(target: vscode.Uri[], uri: vscode.Uri | undefined): void {
    if (!uri || !isSupportedWorkbookContextUri(uri)) {
        return;
    }

    if (target.some((existingUri) => existingUri.toString() === uri.toString())) {
        return;
    }

    target.push(uri);
}

function isSupportedWorkbookContextUri(uri: vscode.Uri): boolean {
    if (uri.scheme === "file") {
        return getWorkbookBasename(uri).endsWith(".xlsx");
    }

    const parsedUri = parseSvnShowUri(uri);
    return (
        uri.scheme === "svn" &&
        parsedUri?.action === "SHOW" &&
        typeof parsedUri.fsPath === "string" &&
        path.extname(parsedUri.fsPath).toLowerCase() === ".xlsx"
    );
}

function getRecentWorkbookContextUris(parsedLabel: ParsedSvnTabLabel): vscode.Uri[] {
    const contextUris: vscode.Uri[] = [];

    for (const entry of getRecentWorkbookResourceEntries()) {
        if (getWorkbookBasename(entry.uri) !== parsedLabel.basename) {
            continue;
        }

        if (entry.uri.scheme === "file") {
            addUniqueUri(contextUris, entry.uri);
            continue;
        }

        const parsedUri = parseSvnShowUri(entry.uri);
        if (parsedUri?.ref?.toUpperCase() !== parsedLabel.ref) {
            continue;
        }

        addUniqueUri(contextUris, entry.uri);
    }

    return contextUris;
}

function getUnknownTabContextUris(tab: vscode.Tab): vscode.Uri[] {
    const contextUris: vscode.Uri[] = [];
    addUniqueUri(contextUris, vscode.window.activeTextEditor?.document.uri);

    for (const editor of vscode.window.visibleTextEditors) {
        addUniqueUri(contextUris, editor.document.uri);
    }

    if (typeof tab.input !== "object" || tab.input === null) {
        return contextUris;
    }

    const rawInput = tab.input as Record<string, unknown>;
    for (const key of ["uri", "resource", "resourceUri", "original", "modified"]) {
        const value = rawInput[key];
        if (value instanceof vscode.Uri) {
            addUniqueUri(contextUris, value);
        }
    }

    return contextUris;
}

function matchesSvnResourceContextUri(
    resource: SvnResource,
    ref: string,
    contextUri: vscode.Uri
): boolean {
    if (contextUri.scheme === "file") {
        return (
            normalizePathForComparison(contextUri.fsPath) ===
            normalizePathForComparison(resource.resourceUri.fsPath)
        );
    }

    const parsedContextUri = parseSvnShowUri(contextUri);
    if (!parsedContextUri || parsedContextUri.action !== "SHOW" || !parsedContextUri.fsPath) {
        return false;
    }

    return (
        normalizePathForComparison(parsedContextUri.fsPath) ===
            normalizePathForComparison(resource.resourceUri.fsPath) &&
        parsedContextUri.ref?.toUpperCase() === ref.toUpperCase()
    );
}

function describeCandidate(candidate: SvnDiffCandidate): string {
    return `${candidate.group}:${candidate.resource.type ?? "<unknown>"}:${candidate.resource.resourceUri.fsPath}`;
}

export async function resolveUnknownSvnWorkbookDiff(
    tab: vscode.Tab,
    log: (message: string) => void
): Promise<{ original: vscode.Uri; modified: vscode.Uri } | undefined> {
    if (!isUnknownTabInput(tab.input)) {
        return undefined;
    }

    const parsedLabel = parseSvnTabLabel(tab.label);
    if (!parsedLabel || parsedLabel.ref !== "HEAD") {
        return undefined;
    }

    const sourceControlManager = await getSvnSourceControlManager();
    if (!sourceControlManager) {
        log("svn unknown tab fallback skipped; svn.getSourceControlManager is unavailable");
        return undefined;
    }

    try {
        await sourceControlManager.isInitialized;
    } catch (error) {
        log(`svn unknown tab fallback skipped; svn-scm init failed: ${String(error)}`);
        return undefined;
    }

    const candidates: SvnDiffCandidate[] = [];
    for (const repository of sourceControlManager.repositories ?? []) {
        for (const { name: group, group: resourceGroup } of getRepositoryResourceGroups(
            repository
        )) {
            for (const resource of resourceGroup.resourceStates) {
                if (getWorkbookBasename(resource.resourceUri) !== parsedLabel.basename) {
                    continue;
                }

                const diffUris = getDiffUrisForSvnResource(group, resource);
                if (!diffUris) {
                    continue;
                }

                candidates.push({
                    repository,
                    group,
                    resource,
                    diffUris,
                });
            }
        }
    }

    if (candidates.length === 0) {
        log(`svn unknown tab fallback skipped; no workbook candidate for label="${tab.label}"`);
        return undefined;
    }

    const contextUris = [
        ...getRecentWorkbookContextUris(parsedLabel),
        ...getUnknownTabContextUris(tab),
    ].filter(
        (uri, index, allUris) =>
            allUris.findIndex((otherUri) => otherUri.toString() === uri.toString()) === index
    );
    const contextMatchedCandidates = candidates.filter((candidate) =>
        contextUris.some((contextUri) =>
            matchesSvnResourceContextUri(candidate.resource, parsedLabel.ref, contextUri)
        )
    );
    const resolvedCandidates =
        contextMatchedCandidates.length === 1 ? contextMatchedCandidates : candidates;

    if (resolvedCandidates.length > 1) {
        const contextSummary =
            contextUris.length > 0
                ? contextUris.map((uri) => describeUri(uri)).join(", ")
                : "<none>";
        const candidateSummary = resolvedCandidates
            .slice(0, 4)
            .map((candidate) => describeCandidate(candidate))
            .join(", ");
        log(
            `svn unknown tab fallback skipped; ambiguous workbook candidates=${resolvedCandidates.length} label="${tab.label}" context=${contextSummary} sample=${candidateSummary}`
        );
        return undefined;
    }

    const [{ repository, group, resource, diffUris }] = resolvedCandidates;
    if (contextMatchedCandidates.length === 1) {
        log(
            `svn unknown tab fallback narrowed by context: label="${tab.label}" context=${contextUris.map((uri) => describeUri(uri)).join(", ")}`
        );
    }
    log(
        `svn unknown tab fallback matched: label="${tab.label}" repo=${repository.root ?? repository.workspaceRoot ?? "<unknown>"} group=${group} type=${resource.type ?? "<unknown>"} original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)}`
    );
    return diffUris;
}
