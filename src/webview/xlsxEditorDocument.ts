import * as vscode from 'vscode';
import {
	type CellEdit,
	writeCellValuesToDestination,
} from '../core/fastxlsx/writeCellValue';

function getCellEditKey(edit: CellEdit): string {
	return `${edit.sheetName}:${edit.rowNumber}:${edit.columnNumber}`;
}

function compareCellEdits(left: CellEdit, right: CellEdit): number {
	if (left.sheetName !== right.sheetName) {
		return left.sheetName.localeCompare(right.sheetName);
	}

	if (left.rowNumber !== right.rowNumber) {
		return left.rowNumber - right.rowNumber;
	}

	if (left.columnNumber !== right.columnNumber) {
		return left.columnNumber - right.columnNumber;
	}

	return left.value.localeCompare(right.value);
}

function areCellEditsEqual(left: readonly CellEdit[], right: readonly CellEdit[]): boolean {
	if (left.length !== right.length) {
		return false;
	}

	return left.every((edit, index) => {
		const other = right[index];
		return (
			edit.sheetName === other.sheetName &&
			edit.rowNumber === other.rowNumber &&
			edit.columnNumber === other.columnNumber &&
			edit.value === other.value
		);
	});
}

export class XlsxEditorDocument implements vscode.CustomDocument {
	private readonly pendingEdits = new Map<string, CellEdit>();
	private backupUri: vscode.Uri | null;
	private shouldMarkDirtyFromBackup: boolean;

	public constructor(
		public readonly uri: vscode.Uri,
		backupUri?: vscode.Uri,
	) {
		this.backupUri = backupUri ?? null;
		this.shouldMarkDirtyFromBackup = Boolean(backupUri);
	}

	public dispose(): void {}

	public getReadUri(): vscode.Uri {
		return this.backupUri ?? this.uri;
	}

	public hasPendingEdits(): boolean {
		return this.pendingEdits.size > 0 || this.backupUri !== null;
	}

	public consumeInitialDirtyState(): boolean {
		if (!this.shouldMarkDirtyFromBackup) {
			return false;
		}

		this.shouldMarkDirtyFromBackup = false;
		return true;
	}

	public getPendingEdits(): CellEdit[] {
		return [...this.pendingEdits.values()].sort(compareCellEdits);
	}

	public replacePendingEdits(edits: readonly CellEdit[]): boolean {
		const normalizedEdits = [...edits].sort(compareCellEdits);
		const currentEdits = this.getPendingEdits();
		if (areCellEditsEqual(currentEdits, normalizedEdits)) {
			return false;
		}

		this.pendingEdits.clear();
		for (const edit of normalizedEdits) {
			this.pendingEdits.set(getCellEditKey(edit), edit);
		}

		return true;
	}

	public markSaved(): void {
		this.pendingEdits.clear();
		this.backupUri = null;
		this.shouldMarkDirtyFromBackup = false;
	}

	public markReverted(): void {
		this.markSaved();
	}

	public async saveTo(destination: vscode.Uri): Promise<void> {
		await writeCellValuesToDestination(
			this.getReadUri(),
			destination,
			this.getPendingEdits(),
		);
	}
}