export function createWebviewNonce(): string {
    return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

export function escapeWatcherGlobSegment(value: string): string {
    return value.replace(/[{}\[\]*?]/g, "[$&]");
}
