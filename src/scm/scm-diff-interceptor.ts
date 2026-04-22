import * as vscode from "vscode";
import {
    WEBVIEW_TYPE_DIFF_PANEL,
    WEBVIEW_TYPE_EDITOR_PANEL,
} from "../constants";
import { resolveUnknownGitWorkbookDiff } from "../git/scm-diff-fallback";
import { resolveUnknownSvnWorkbookDiff } from "../svn/scm-diff-fallback";
import { XlsxDiffPanel } from "../webview/diff-panel";
import {
    getScmWorkbookDiffUrisFromEditorUris,
    getScmWorkbookDiffUrisFromTabInput,
    isWorkbookResourceUri,
} from "../workbook/resource-uri";
import { getTabInputKind, getTabResourceUri } from "./tab-input";
import { describeUri } from "./uri-description";

interface CustomWorkbookEditorTab {
    tab: vscode.Tab;
    uri: vscode.Uri;
}

function isWorkbookRelatedTab(tab: vscode.Tab): boolean {
    if (tab.input instanceof vscode.TabInputTextDiff) {
        const textDiffInput = tab.input as vscode.TabInputTextDiff;
        const { original, modified } = textDiffInput;
        const originalIsWorkbook = isWorkbookResourceUri(original);
        const modifiedIsWorkbook = isWorkbookResourceUri(modified);
        return (
            original.scheme === "git" ||
            original.scheme === "svn" ||
            modified.scheme === "git" ||
            modified.scheme === "svn" ||
            originalIsWorkbook ||
            modifiedIsWorkbook
        );
    }

    const resourceUri = getTabResourceUri(tab.input);
    return isWorkbookResourceUri(resourceUri);
}

function isPotentialUnknownWorkbookScmTab(tab: vscode.Tab): boolean {
    return (
        getTabInputKind(tab.input).startsWith("unknown:") &&
        /^.+\.xlsx \(([^()]+)\)$/i.test(tab.label)
    );
}

function describeTab(tab: vscode.Tab): string {
    const tabSummary = `label="${tab.label}" active=${tab.isActive} preview=${tab.isPreview} dirty=${tab.isDirty} kind=${getTabInputKind(tab.input)} column=${tab.group.viewColumn}`;

    if (tab.input instanceof vscode.TabInputTextDiff) {
        return `${tabSummary} original=${describeUri(tab.input.original)} modified=${describeUri(tab.input.modified)}`;
    }

    const resourceUri = getTabResourceUri(tab.input);
    if (resourceUri) {
        return `${tabSummary} resource=${describeUri(resourceUri)}`;
    }

    return tabSummary;
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
    return tab.input instanceof vscode.TabInputWebview && tab.input.viewType === WEBVIEW_TYPE_DIFF_PANEL;
}

function getCustomWorkbookEditorTabs(): CustomWorkbookEditorTab[] {
    return vscode.window.tabGroups.all.flatMap((group) =>
        group.tabs
            .map((tab) => getCustomWorkbookEditorTab(tab))
            .filter((tab): tab is CustomWorkbookEditorTab => Boolean(tab))
    );
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
    const recentlyLoggedCandidateTabs = new Map<string, number>();
    const recentLogMessages = new Map<string, number>();

    const log = (message: string): void => {
        const now = Date.now();
        const lastLoggedAt = recentLogMessages.get(message);
        if (lastLoggedAt && now - lastLoggedAt < 500) {
            return;
        }

        recentLogMessages.set(message, now);
        for (const [loggedMessage, loggedAt] of recentLogMessages) {
            if (now - loggedAt > 5_000) {
                recentLogMessages.delete(loggedMessage);
            }
        }

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

    const pruneRecentlyLoggedCandidateTabs = (): void => {
        const now = Date.now();
        for (const [tabKey, loggedAt] of recentlyLoggedCandidateTabs) {
            if (now - loggedAt > 2_000) {
                recentlyLoggedCandidateTabs.delete(tabKey);
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
            log(
                `skip diff panel open; request already in flight original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)}`
            );
            return;
        }

        inFlight.add(requestKey);
        try {
            log(
                `open diff panel original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)} viewColumn=${viewColumn} tabsToClose=${tabsToClose.length}`
            );
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
        return (
            (await resolveUnknownGitWorkbookDiff(tab, log)) ??
            (await resolveUnknownSvnWorkbookDiff(tab, log))
        );
    };

    const maybeInterceptTab = async (tab: vscode.Tab | undefined): Promise<void> => {
        prunePendingCloseTabKeys();
        pruneRecentlyLoggedCandidateTabs();
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

        if (
            (isWorkbookRelatedTab(tab) || isPotentialUnknownWorkbookScmTab(tab)) &&
            !recentlyLoggedCandidateTabs.has(tabKey)
        ) {
            recentlyLoggedCandidateTabs.set(tabKey, Date.now());
            log(`inspect scm candidate tab: ${describeTab(tab)}`);
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

                if (isWorkbookRelatedTab(tab)) {
                    log(`no scm workbook diff match for tab: ${describeTab(tab)}`);
                }
                return;
            }

            log(
                `direct scm diff match original=${describeUri(diffUris.original)} modified=${describeUri(diffUris.modified)}`
            );
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
                    firstTab.uri.scheme !== "file" ||
                    secondTab.uri.scheme !== "file" ||
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

        if (sawCustomWorkbookEditorTab || getCustomWorkbookEditorTabs().length >= 2) {
            void maybeInterceptCustomEditorPair();
            scheduleCustomEditorPairScan();
        }
    };

    void maybeInterceptTab(vscode.window.tabGroups.activeTabGroup.activeTab);
    void maybeInterceptCustomEditorPair();

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
