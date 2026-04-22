import * as path from "node:path";
import * as vscode from "vscode";
import { isUnknownTabInput } from "../scm/tab-input";
import { describeUri } from "../scm/uri-description";

type GitChangeGroup = "index" | "workingTree" | "merge" | "untracked";
type ScmTabKind =
    | "index"
    | "workingTree"
    | "deleted"
    | "untracked"
    | "intentToAdd"
    | "typeChanged"
    | "ours"
    | "theirs";

interface ParsedScmTabLabel {
    basename: string;
    tabKind: ScmTabKind;
}

interface GitChange {
    readonly uri: vscode.Uri;
    readonly originalUri: vscode.Uri;
    readonly renameUri?: vscode.Uri;
    readonly status: GitStatus;
}

interface GitRepositoryState {
    readonly indexChanges: readonly GitChange[];
    readonly workingTreeChanges: readonly GitChange[];
    readonly mergeChanges: readonly GitChange[];
    readonly untrackedChanges?: readonly GitChange[];
}

interface GitRepository {
    readonly rootUri: vscode.Uri;
    readonly state: GitRepositoryState;
}

interface GitApi {
    readonly repositories: readonly GitRepository[];
    toGitUri(uri: vscode.Uri, ref: string): vscode.Uri;
}

interface GitExtension {
    getAPI(version: 1): GitApi;
}

interface GitChangeCandidate {
    repository: GitRepository;
    group: GitChangeGroup;
    change: GitChange;
}

const GitStatus = {
    INDEX_MODIFIED: 0,
    INDEX_ADDED: 1,
    INDEX_DELETED: 2,
    INDEX_RENAMED: 3,
    INDEX_COPIED: 4,
    MODIFIED: 5,
    DELETED: 6,
    UNTRACKED: 7,
    IGNORED: 8,
    INTENT_TO_ADD: 9,
    INTENT_TO_RENAME: 10,
    TYPE_CHANGED: 11,
    ADDED_BY_US: 12,
    ADDED_BY_THEM: 13,
    DELETED_BY_US: 14,
    DELETED_BY_THEM: 15,
    BOTH_ADDED: 16,
    BOTH_DELETED: 17,
    BOTH_MODIFIED: 18,
} as const;

type GitStatus = (typeof GitStatus)[keyof typeof GitStatus];

function parseScmTabKind(rawKind: string): ScmTabKind | undefined {
    switch (rawKind.toLowerCase()) {
        case "index":
            return "index";
        case "working tree":
            return "workingTree";
        case "deleted":
            return "deleted";
        case "untracked":
            return "untracked";
        case "intent to add":
            return "intentToAdd";
        case "type changed":
            return "typeChanged";
        case "ours":
            return "ours";
        case "theirs":
            return "theirs";
        default:
            return undefined;
    }
}

function parseScmTabLabel(label: string): ParsedScmTabLabel | undefined {
    const match = /^(.+\.xlsx) \(([^()]+)\)$/i.exec(label);
    if (!match) {
        return undefined;
    }

    const tabKind = parseScmTabKind(match[2]);
    if (!tabKind) {
        return undefined;
    }

    return {
        basename: match[1].toLowerCase(),
        tabKind,
    };
}

function getStatusName(status: GitStatus): string {
    const entry = Object.entries(GitStatus).find(([, value]) => value === status);
    return entry?.[0] ?? `UNKNOWN(${status})`;
}

function getExpectedScmTabKind(group: GitChangeGroup, status: GitStatus): ScmTabKind | undefined {
    if (group === "index") {
        switch (status) {
            case GitStatus.INDEX_MODIFIED:
            case GitStatus.INDEX_ADDED:
            case GitStatus.INDEX_RENAMED:
            case GitStatus.INDEX_COPIED:
                return "index";
            case GitStatus.INDEX_DELETED:
                return "deleted";
            default:
                return undefined;
        }
    }

    if (group === "workingTree") {
        switch (status) {
            case GitStatus.MODIFIED:
            case GitStatus.BOTH_ADDED:
            case GitStatus.BOTH_MODIFIED:
                return "workingTree";
            case GitStatus.DELETED:
                return "deleted";
            case GitStatus.INTENT_TO_ADD:
            case GitStatus.INTENT_TO_RENAME:
                return "intentToAdd";
            case GitStatus.TYPE_CHANGED:
                return "typeChanged";
            default:
                return undefined;
        }
    }

    if (group === "untracked") {
        return "untracked";
    }

    switch (status) {
        case GitStatus.DELETED_BY_US:
            return "theirs";
        case GitStatus.DELETED_BY_THEM:
            return "ours";
        case GitStatus.BOTH_ADDED:
        case GitStatus.BOTH_MODIFIED:
            return "workingTree";
        default:
            return undefined;
    }
}

function getWorkbookBasename(uri: vscode.Uri): string {
    return path.basename(uri.fsPath || uri.path).toLowerCase();
}

function isXlsxGitChange(change: GitChange): boolean {
    return getWorkbookBasename(change.uri).endsWith(".xlsx");
}

async function getGitApi(): Promise<GitApi | undefined> {
    const extension = vscode.extensions.getExtension<GitExtension>("vscode.git");
    if (!extension) {
        return undefined;
    }

    const gitExtension = extension.isActive ? extension.exports : await extension.activate();
    return gitExtension.getAPI(1);
}

function getCandidateGroups(repository: GitRepository): GitChangeCandidate[] {
    return [
        ...repository.state.indexChanges.map((change) => ({
            repository,
            group: "index" as const,
            change,
        })),
        ...repository.state.workingTreeChanges.map((change) => ({
            repository,
            group: "workingTree" as const,
            change,
        })),
        ...repository.state.mergeChanges.map((change) => ({
            repository,
            group: "merge" as const,
            change,
        })),
        ...(repository.state.untrackedChanges ?? []).map((change) => ({
            repository,
            group: "untracked" as const,
            change,
        })),
    ];
}

function getDiffUrisForGitCandidate(
    api: GitApi,
    candidate: GitChangeCandidate
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
    const { group, change } = candidate;

    if (group === "index") {
        switch (change.status) {
            case GitStatus.INDEX_MODIFIED:
            case GitStatus.INDEX_RENAMED:
            case GitStatus.INDEX_COPIED:
                return {
                    original: api.toGitUri(change.originalUri, "HEAD"),
                    modified: api.toGitUri(change.uri, ""),
                };
            default:
                return undefined;
        }
    }

    if (group === "workingTree") {
        switch (change.status) {
            case GitStatus.MODIFIED:
                return {
                    original: api.toGitUri(change.uri, "~"),
                    modified: change.uri,
                };
            case GitStatus.TYPE_CHANGED:
            case GitStatus.INTENT_TO_RENAME:
                return {
                    original: api.toGitUri(change.originalUri, "HEAD"),
                    modified: change.uri,
                };
            default:
                return undefined;
        }
    }

    return undefined;
}

export async function resolveUnknownGitWorkbookDiff(
    tab: vscode.Tab,
    log: (message: string) => void
): Promise<{ original: vscode.Uri; modified: vscode.Uri } | undefined> {
    if (!isUnknownTabInput(tab.input)) {
        return undefined;
    }

    const parsedLabel = parseScmTabLabel(tab.label);
    if (!parsedLabel) {
        return undefined;
    }

    let gitApi: GitApi | undefined;
    try {
        gitApi = await getGitApi();
    } catch (error) {
        log(`unknown tab fallback skipped; failed to activate vscode.git: ${String(error)}`);
        return undefined;
    }

    if (!gitApi) {
        log("unknown tab fallback skipped; vscode.git extension is unavailable");
        return undefined;
    }

    const candidates = gitApi.repositories
        .flatMap((repository) => getCandidateGroups(repository))
        .filter((candidate) => {
            const expectedTabKind = getExpectedScmTabKind(
                candidate.group,
                candidate.change.status
            );
            return (
                expectedTabKind === parsedLabel.tabKind &&
                isXlsxGitChange(candidate.change) &&
                getWorkbookBasename(candidate.change.uri) === parsedLabel.basename
            );
        });

    const diffCandidates = candidates
        .map((candidate) => ({
            candidate,
            diffUris: getDiffUrisForGitCandidate(gitApi, candidate),
        }))
        .filter(
            (
                item
            ): item is {
                candidate: GitChangeCandidate;
                diffUris: { original: vscode.Uri; modified: vscode.Uri };
            } => Boolean(item.diffUris)
        );

    if (diffCandidates.length === 0) {
        log(
            `unknown tab fallback skipped; no two-sided workbook diff candidate for label="${tab.label}" candidates=${candidates.length}`
        );
        return undefined;
    }

    if (diffCandidates.length > 1) {
        log(
            `unknown tab fallback skipped; ambiguous workbook basename candidates=${diffCandidates.length}`
        );
        return undefined;
    }

    const [{ candidate, diffUris }] = diffCandidates;
    log(
        `unknown tab fallback matched: label="${tab.label}" repo=${describeUri(candidate.repository.rootUri)} group=${candidate.group} status=${getStatusName(candidate.change.status)} original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)}`
    );
    return diffUris;
}
