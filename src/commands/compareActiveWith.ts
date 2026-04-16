import * as vscode from 'vscode';
import { XlsxDiffPanel } from '../webview/diffPanel';
import {
	getActiveWorkbookUri,
	isWorkbookUri,
	pickWorkbook,
} from './workbookPicker';

export async function compareActiveWith(
	extensionUri: vscode.Uri,
	resource?: vscode.Uri,
): Promise<void> {
	const leftUri =
		(isWorkbookUri(resource) ? resource : undefined) ?? getActiveWorkbookUri();

	if (!leftUri) {
		await vscode.window.showErrorMessage(
			'Open an .xlsx file first, or run "Compare Two XLSX Files" from the Command Palette.',
		);
		return;
	}

	const rightUri = await pickWorkbook(
		'Select the XLSX workbook to compare against',
		leftUri,
	);
	if (!rightUri) {
		return;
	}

	await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
