import * as vscode from "vscode";

export interface RecentWorkbookResourceEntry {
    readonly uri: vscode.Uri;
    readonly source: string;
    readonly recordedAt: number;
}

const RECENT_RESOURCE_TTL_MS = 10_000;
const MAX_RECENT_RESOURCES = 20;
const recentWorkbookResources: RecentWorkbookResourceEntry[] = [];

function pruneRecentWorkbookResources(): void {
    const cutoff = Date.now() - RECENT_RESOURCE_TTL_MS;
    let writeIndex = 0;

    for (const entry of recentWorkbookResources) {
        if (entry.recordedAt < cutoff) {
            continue;
        }

        recentWorkbookResources[writeIndex] = entry;
        writeIndex += 1;
    }

    recentWorkbookResources.length = writeIndex;
}

export function rememberRecentWorkbookResourceUri(uri: vscode.Uri, source: string): void {
    pruneRecentWorkbookResources();

    const existingEntryIndex = recentWorkbookResources.findIndex(
        (entry) => entry.uri.toString() === uri.toString() && entry.source === source
    );
    if (existingEntryIndex >= 0) {
        recentWorkbookResources.splice(existingEntryIndex, 1);
    }

    recentWorkbookResources.unshift({
        uri,
        source,
        recordedAt: Date.now(),
    });

    if (recentWorkbookResources.length > MAX_RECENT_RESOURCES) {
        recentWorkbookResources.length = MAX_RECENT_RESOURCES;
    }
}

export function getRecentWorkbookResourceEntries(): readonly RecentWorkbookResourceEntry[] {
    pruneRecentWorkbookResources();
    return [...recentWorkbookResources];
}
