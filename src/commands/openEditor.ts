import * as vscode from 'vscode';
import { isChineseDisplayLanguage } from '../displayLanguage';
import { WEBVIEW_TYPE_EDITOR_PANEL } from '../constants';
import {
	getActiveWorkbookUri,
	getWorkbookUriFromCommandArg,
} from './workbookPicker';

export function resolveWorkbookUriForOpenEditor(
	resource: unknown,
	activeWorkbookUri?: vscode.Uri,
): vscode.Uri | undefined {
	return getWorkbookUriFromCommandArg(resource) ?? activeWorkbookUri;
}

export async function openEditor(
	resource?: unknown,
): Promise<void> {
	const isChinese = isChineseDisplayLanguage();
	const workbookUri = resolveWorkbookUriForOpenEditor(
		resource,
		getActiveWorkbookUri(),
	);

	if (!workbookUri) {
		await vscode.window.showErrorMessage(
			isChinese
				? '请先选择或打开一个本地 .xlsx 文件。'
				: 'Select or open a local .xlsx file first.',
		);
		return;
	}

	await vscode.commands.executeCommand(
		'vscode.openWith',
		workbookUri,
		WEBVIEW_TYPE_EDITOR_PANEL,
	);
}