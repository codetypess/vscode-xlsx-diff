import * as vscode from "vscode";
import type { WorkbookEditState } from "../../core/fastxlsx/write-cell-value";

const BACKUP_STATE_TYPE = "xlsx-diff-editor-backup";
const BACKUP_STATE_VERSION = 1;

interface SerializedEditorBackupState {
    type: typeof BACKUP_STATE_TYPE;
    version: typeof BACKUP_STATE_VERSION;
    pendingState: WorkbookEditState;
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizePendingState(value: unknown): WorkbookEditState | null {
    if (!isObjectRecord(value)) {
        return null;
    }

    const cellEdits = Array.isArray(value.cellEdits) ? value.cellEdits : null;
    const sheetEdits = Array.isArray(value.sheetEdits) ? value.sheetEdits : null;
    const viewEdits = Array.isArray(value.viewEdits) ? value.viewEdits : [];
    if (!cellEdits || !sheetEdits) {
        return null;
    }

    return {
        cellEdits: cellEdits as WorkbookEditState["cellEdits"],
        sheetEdits: sheetEdits as WorkbookEditState["sheetEdits"],
        viewEdits: viewEdits as WorkbookEditState["viewEdits"],
    };
}

export async function writeEditorBackupState(
    destination: vscode.Uri,
    pendingState: WorkbookEditState
): Promise<void> {
    const payload: SerializedEditorBackupState = {
        type: BACKUP_STATE_TYPE,
        version: BACKUP_STATE_VERSION,
        pendingState,
    };
    const encoder = new TextEncoder();
    await vscode.workspace.fs.writeFile(destination, encoder.encode(JSON.stringify(payload)));
}

export async function readEditorBackupState(
    backupUri: vscode.Uri
): Promise<WorkbookEditState | null> {
    const decoder = new TextDecoder();
    const raw = decoder.decode(await vscode.workspace.fs.readFile(backupUri));
    let parsed: unknown;

    try {
        parsed = JSON.parse(raw);
    } catch {
        return null;
    }

    if (!isObjectRecord(parsed)) {
        return null;
    }

    if (parsed.type !== BACKUP_STATE_TYPE || parsed.version !== BACKUP_STATE_VERSION) {
        return null;
    }

    return normalizePendingState(parsed.pendingState);
}
