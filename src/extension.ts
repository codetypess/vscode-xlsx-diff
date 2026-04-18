import * as vscode from 'vscode';
import { compareActiveWith } from './commands/compareActiveWith';
import { compareTwoFiles } from './commands/compareTwoFiles';
import { openEditor } from './commands/openEditor';
import {
	COMMAND_COMPARE_ACTIVE_WITH,
	COMMAND_COMPARE_TWO_FILES,
	COMMAND_OPEN_EDITOR,
} from './constants';
import { affectsDisplayLanguage } from './displayLanguage';
import { XlsxDiffUriHandler } from './git/uriHandler';
import { registerScmWorkbookDiffInterceptor } from './scm/scmDiffInterceptor';
import { XlsxDiffPanel } from './webview/diffPanel';
import { XlsxCustomEditorProvider } from './webview/xlsxCustomEditorProvider';
import { XlsxEditorPanel } from './webview/xlsxEditorPanel';

export function activate(context: vscode.ExtensionContext) {
	context.subscriptions.push(
		XlsxCustomEditorProvider.register(context),
		vscode.window.registerUriHandler(
			new XlsxDiffUriHandler(context.extensionUri),
		),
		registerScmWorkbookDiffInterceptor(context.extensionUri),
		vscode.workspace.onDidChangeConfiguration((event) => {
			if (!affectsDisplayLanguage(event)) {
				return;
			}

			void Promise.all([XlsxDiffPanel.refreshAll(), XlsxEditorPanel.refreshAll()]);
		}),
		vscode.commands.registerCommand(COMMAND_COMPARE_TWO_FILES, async () => {
			await compareTwoFiles(context.extensionUri);
		}),
		vscode.commands.registerCommand(
			COMMAND_OPEN_EDITOR,
			async (resource?: unknown) => {
				await openEditor(resource);
			},
		),
		vscode.commands.registerCommand(
			COMMAND_COMPARE_ACTIVE_WITH,
			async (resource?: unknown) => {
				await compareActiveWith(context.extensionUri, resource);
			},
		),
	);
}

export function deactivate() {}
