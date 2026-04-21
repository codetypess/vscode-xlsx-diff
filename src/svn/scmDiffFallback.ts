import * as path from "node:path";
import * as vscode from "vscode";
import { isUnknownTabInput } from "../scm/tabInput";
import { describeUri } from "../scm/uriDescription";

interface ParsedSvnTabLabel {
    basename: string;
    ref: string;
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

    for (const [name, group] of repository.changelists ?? []) {
        groups.push({ name: `changelist:${name}`, group });
    }

    return groups;
}

function getDiffUrisForSvnResource(
    resource: SvnResource
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    if (resource.resourceUri.scheme !== "file") {
        return undefined;
    }

    if (!getWorkbookBasename(resource.resourceUri).endsWith(".xlsx")) {
        return undefined;
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

                const diffUris = getDiffUrisForSvnResource(resource);
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

    if (candidates.length > 1) {
        log(
            `svn unknown tab fallback skipped; ambiguous workbook basename candidates=${candidates.length} label="${tab.label}"`
        );
        return undefined;
    }

    const [{ repository, group, resource, diffUris }] = candidates;
    log(
        `svn unknown tab fallback matched: label="${tab.label}" repo=${repository.root ?? repository.workspaceRoot ?? "<unknown>"} group=${group} type=${resource.type ?? "<unknown>"} original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)}`
    );
    return diffUris;
}
