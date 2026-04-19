import * as path from "node:path";
import * as vscode from "vscode";
import { WEBVIEW_TYPE_DIFF_PANEL, WEBVIEW_TYPE_EDITOR_PANEL } from "../constants";
import { XlsxDiffPanel } from "../webview/diffPanel";
import {
    getScmWorkbookDiffUrisFromEditorUris,
    getScmWorkbookDiffUrisFromTabInput,
    isWorkbookResourceUri,
} from "../workbook/resourceUri";

interface CustomWorkbookEditorTab {
    tab: vscode.Tab;
    uri: vscode.Uri;
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

function describeUri(uri: vscode.Uri | undefined): string {
    if (!uri) {
        return "<none>";
    }

    const query = uri.query ? ` query=${uri.query}` : "";
    const fragment = uri.fragment ? ` fragment=${uri.fragment}` : "";
    const fsPath = uri.scheme === "file" ? ` fsPath=${uri.fsPath}` : "";
    return `${uri.toString()} [scheme=${uri.scheme} path=${uri.path}${fsPath}${query}${fragment}]`;
}

function getTabInputKind(input: vscode.Tab["input"]): string {
    if (input instanceof vscode.TabInputText) {
        return "text";
    }

    if (input instanceof vscode.TabInputTextDiff) {
        return "textDiff";
    }

    if (input instanceof vscode.TabInputCustom) {
        return `custom:${input.viewType}`;
    }

    if (input instanceof vscode.TabInputWebview) {
        return `webview:${input.viewType}`;
    }

    if (input instanceof vscode.TabInputNotebook) {
        return `notebook:${input.notebookType}`;
    }

    if (input instanceof vscode.TabInputNotebookDiff) {
        return `notebookDiff:${input.notebookType}`;
    }

    if (input instanceof vscode.TabInputTerminal) {
        return "terminal";
    }

    return `unknown:${input?.constructor?.name ?? typeof input}`;
}

function isUnknownTabInput(input: vscode.Tab["input"]): boolean {
    return getTabInputKind(input).startsWith("unknown:");
}

function getTabResourceUri(input: vscode.Tab["input"]): vscode.Uri | undefined {
    if (input instanceof vscode.TabInputText) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputCustom) {
        return input.uri;
    }

    if (input instanceof vscode.TabInputNotebook) {
        return input.uri;
    }

    return undefined;
}

function getTabKey(tab: vscode.Tab): string {
    if (tab.input instanceof vscode.TabInputTextDiff) {
        return [
            "textDiff",
            tab.input.original.toString(),
            tab.input.modified.toString(),
            tab.group.viewColumn,
        ].join("::");
    }

    const resourceUri = getTabResourceUri(tab.input);
    if (resourceUri) {
        return [getTabInputKind(tab.input), resourceUri.toString(), tab.group.viewColumn].join(
            "::"
        );
    }

    return [getTabInputKind(tab.input), tab.label, tab.group.viewColumn].join("::");
}

function getCustomWorkbookEditorTab(tab: vscode.Tab): CustomWorkbookEditorTab | undefined {
    if (
        !(tab.input instanceof vscode.TabInputCustom) ||
        tab.input.viewType !== WEBVIEW_TYPE_EDITOR_PANEL ||
        !isWorkbookResourceUri(tab.input.uri)
    ) {
        return undefined;
    }

    return { tab, uri: tab.input.uri };
}

function isXlsxDiffPanelTab(tab: vscode.Tab): boolean {
    return (
        tab.input instanceof vscode.TabInputWebview &&
        tab.input.viewType.endsWith(WEBVIEW_TYPE_DIFF_PANEL)
    );
}

function getCustomWorkbookEditorTabs(): CustomWorkbookEditorTab[] {
    return vscode.window.tabGroups.all.flatMap((group) =>
        group.tabs
            .map((tab) => getCustomWorkbookEditorTab(tab))
            .filter((tab): tab is CustomWorkbookEditorTab => Boolean(tab))
    );
}

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

async function closePreviewWorkbookTabs(resourceUri: vscode.Uri): Promise<void> {
    const tabsToClose = vscode.window.tabGroups.all.flatMap((group) =>
        group.tabs.filter((tab) => {
            if (tab.isDirty) {
                return false;
            }

            const tabResourceUri = getTabResourceUri(tab.input);
            if (tabResourceUri?.toString() !== resourceUri.toString()) {
                return false;
            }

            return tab.isPreview || tab.input instanceof vscode.TabInputText;
        })
    );

    if (tabsToClose.length > 0) {
        await vscode.window.tabGroups.close(tabsToClose, true);
    }
}

async function closeTabs(tabs: vscode.Tab[]): Promise<void> {
    const tabsToClose = [...new Set(tabs)].filter((tab) => !tab.isDirty);
    if (tabsToClose.length > 0) {
        await vscode.window.tabGroups.close(tabsToClose, true);
    }
}

export function registerScmWorkbookDiffInterceptor(extensionUri: vscode.Uri): vscode.Disposable {
    const outputChannel = vscode.window.createOutputChannel("XLSX Diff");
    const inFlight = new Set<string>();
    const recentCustomEditorTabs = new Map<string, number>();
    const scheduledCustomEditorPairScans = new Set<NodeJS.Timeout>();
    const tabKeysInProgress = new Set<string>();
    const tabKeysPendingClose = new Map<string, number>();

    const log = (message: string): void => {
        outputChannel.appendLine(`[SCM ${new Date().toISOString()}] ${message}`);
    };

    const prunePendingCloseTabKeys = (): void => {
        const now = Date.now();
        for (const [tabKey, markedAt] of tabKeysPendingClose) {
            if (now - markedAt > 5_000) {
                tabKeysPendingClose.delete(tabKey);
            }
        }
    };

    const markTabsPendingClose = (tabs: readonly vscode.Tab[]): void => {
        const now = Date.now();
        for (const tab of tabs) {
            tabKeysPendingClose.set(getTabKey(tab), now);
        }
    };

    const rememberCustomEditorTabs = (tabs: readonly vscode.Tab[]): boolean => {
        const now = Date.now();
        let sawCustomWorkbookEditorTab = false;
        for (const tab of tabs) {
            const customTab = getCustomWorkbookEditorTab(tab);
            if (customTab) {
                sawCustomWorkbookEditorTab = true;
                recentCustomEditorTabs.set(customTab.uri.toString(), now);
            }
        }

        for (const [uriString, openedAt] of recentCustomEditorTabs) {
            if (now - openedAt > 2_000) {
                recentCustomEditorTabs.delete(uriString);
            }
        }

        return sawCustomWorkbookEditorTab;
    };

    const openDiffPanel = async (
        diffUris: { original: vscode.Uri; modified: vscode.Uri },
        viewColumn: vscode.ViewColumn,
        tabsToClose: vscode.Tab[] = []
    ): Promise<void> => {
        const requestKey = `${diffUris.original.toString()}::${diffUris.modified.toString()}`;
        if (inFlight.has(requestKey)) {
            return;
        }

        inFlight.add(requestKey);
        try {
            const openPanelPromise = XlsxDiffPanel.create(
                extensionUri,
                diffUris.original,
                diffUris.modified,
                viewColumn
            );
            await closeTabs(tabsToClose);
            await closePreviewWorkbookTabs(diffUris.modified);
            await openPanelPromise;
        } finally {
            inFlight.delete(requestKey);
            recentCustomEditorTabs.delete(diffUris.original.toString());
            recentCustomEditorTabs.delete(diffUris.modified.toString());
        }
    };

    const maybeResolveUnknownScmDiff = async (
        tab: vscode.Tab
    ): Promise<{ original: vscode.Uri; modified: vscode.Uri } | undefined> => {
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
    };

    const maybeInterceptTab = async (tab: vscode.Tab | undefined): Promise<void> => {
        prunePendingCloseTabKeys();
        const tabKey = tab ? getTabKey(tab) : undefined;
        if (
            !tab?.isActive ||
            !tabKey ||
            tabKeysInProgress.has(tabKey) ||
            tabKeysPendingClose.has(tabKey) ||
            isXlsxDiffPanelTab(tab)
        ) {
            return;
        }

        tabKeysInProgress.add(tabKey);
        try {
            // Explorer opens .xlsx files as preview custom editors too, so only replace real diff tabs.
            const diffUris = getScmWorkbookDiffUrisFromTabInput(tab.input);

            if (!diffUris) {
                const unknownScmDiffUris = await maybeResolveUnknownScmDiff(tab);
                if (unknownScmDiffUris) {
                    markTabsPendingClose([tab]);
                    await openDiffPanel(unknownScmDiffUris, tab.group.viewColumn, [tab]);
                    return;
                }

                return;
            }

            markTabsPendingClose([tab]);
            await openDiffPanel(diffUris, tab.group.viewColumn, [tab]);
        } finally {
            tabKeysInProgress.delete(tabKey);
        }
    };

    const maybeInterceptCustomEditorPair = async (): Promise<void> => {
        const customTabs = getCustomWorkbookEditorTabs();
        if (customTabs.length < 2) {
            return;
        }

        for (const firstTab of customTabs) {
            for (const secondTab of customTabs) {
                if (firstTab.tab === secondTab.tab) {
                    continue;
                }

                const diffUris = getScmWorkbookDiffUrisFromEditorUris(firstTab.uri, secondTab.uri);
                if (!diffUris) {
                    continue;
                }

                const shouldInterceptPair =
                    recentCustomEditorTabs.has(firstTab.uri.toString()) ||
                    recentCustomEditorTabs.has(secondTab.uri.toString()) ||
                    firstTab.tab.isPreview ||
                    secondTab.tab.isPreview;
                if (!shouldInterceptPair) {
                    continue;
                }

                const modifiedTab =
                    firstTab.uri.toString() === diffUris.modified.toString() ? firstTab : secondTab;
                const tabsToClose = [firstTab, secondTab]
                    .filter(
                        (customTab) =>
                            customTab.tab.isPreview ||
                            recentCustomEditorTabs.has(customTab.uri.toString())
                    )
                    .map((customTab) => customTab.tab);
                log(
                    `custom editor pair intercept matched: original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)} tabsToClose=${tabsToClose.length}`
                );
                markTabsPendingClose(tabsToClose);
                await openDiffPanel(diffUris, modifiedTab.tab.group.viewColumn, tabsToClose);
                return;
            }
        }
    };

    const scheduleCustomEditorPairScan = (): void => {
        for (const delay of [50, 250, 750]) {
            const timeout = setTimeout(() => {
                scheduledCustomEditorPairScans.delete(timeout);
                void maybeInterceptCustomEditorPair();
            }, delay);
            scheduledCustomEditorPairScans.add(timeout);
        }
    };

    const handleTabChange = (event: vscode.TabChangeEvent) => {
        prunePendingCloseTabKeys();
        for (const tab of event.closed) {
            tabKeysPendingClose.delete(getTabKey(tab));
        }

        const changedTabs = [...event.opened, ...event.changed];
        const sawCustomWorkbookEditorTab = rememberCustomEditorTabs(changedTabs);
        const tabsToInspect = new Set<vscode.Tab>();

        for (const tab of changedTabs) {
            if (tab.isActive) {
                tabsToInspect.add(tab);
            }
        }

        for (const group of vscode.window.tabGroups.all) {
            if (group.activeTab) {
                tabsToInspect.add(group.activeTab);
            }
        }

        for (const tab of tabsToInspect) {
            void maybeInterceptTab(tab);
        }

        if (sawCustomWorkbookEditorTab) {
            void maybeInterceptCustomEditorPair();
            scheduleCustomEditorPairScan();
        }
    };

    void maybeInterceptTab(vscode.window.tabGroups.activeTabGroup.activeTab);

    return vscode.Disposable.from(
        outputChannel,
        vscode.window.tabGroups.onDidChangeTabs(handleTabChange),
        vscode.window.tabGroups.onDidChangeTabGroups(() => {
            void maybeInterceptTab(vscode.window.tabGroups.activeTabGroup.activeTab);
        }),
        new vscode.Disposable(() => {
            for (const timeout of scheduledCustomEditorPairScans) {
                clearTimeout(timeout);
            }
            scheduledCustomEditorPairScans.clear();
        })
    );
}
