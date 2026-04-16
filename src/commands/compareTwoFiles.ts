import * as vscode from 'vscode';
import { XlsxDiffPanel } from '../webview/diffPanel';
import { pickWorkbook } from './workbookPicker';

export async function compareTwoFiles(extensionUri: vscode.Uri): Promise<void> {
	const leftUri = await pickWorkbook('Select the left XLSX workbook');
	if (!leftUri) {
		return;
	}

	const rightUri = await pickWorkbook('Select the right XLSX workbook', leftUri);
	if (!rightUri) {
		return;
	}

	await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
