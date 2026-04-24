type ToolbarListener = () => void;

const listeners = new Set<ToolbarListener>();
let revision = 0;

export function subscribeEditorToolbarSync(listener: ToolbarListener): () => void {
    listeners.add(listener);

    return () => {
        listeners.delete(listener);
    };
}

export function getEditorToolbarSyncSnapshot(): number {
    return revision;
}

export function notifyEditorToolbarSync(): void {
    revision += 1;

    for (const listener of listeners) {
        listener();
    }
}

export function resetEditorToolbarSyncForTests(): void {
    revision = 0;
    listeners.clear();
}
