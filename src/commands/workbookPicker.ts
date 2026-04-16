import * as path from 'node:path';
import * as vscode from 'vscode';

const XLSX_FILTER = {
	'Excel Workbooks': ['xlsx'],
};

export function isWorkbookUri(uri: vscode.Uri | undefined): uri is vscode.Uri {
	return Boolean(
		uri &&
			uri.scheme === 'file' &&
			path.extname(uri.fsPath).toLowerCase() === '.xlsx',
	);
}

function toDirectoryUri(uri: vscode.Uri | undefined): vscode.Uri | undefined {
	if (!uri) {
		return undefined;
	}

	return vscode.Uri.file(path.dirname(uri.fsPath));
}

export async function pickWorkbook(
	openLabel: string,
	seedUri?: vscode.Uri,
): Promise<vscode.Uri | undefined> {
	const selection = await vscode.window.showOpenDialog({
		canSelectFiles: true,
		canSelectFolders: false,
		canSelectMany: false,
		defaultUri: toDirectoryUri(seedUri) ?? seedUri,
		filters: XLSX_FILTER,
		openLabel,
	});

	return selection?.[0];
}

export function getActiveWorkbookUri(): vscode.Uri | undefined {
	const candidate = vscode.window.activeTextEditor?.document.uri;
	return isWorkbookUri(candidate) ? candidate : undefined;
}
