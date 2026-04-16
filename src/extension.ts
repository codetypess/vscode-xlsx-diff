import * as vscode from 'vscode';
import { compareActiveWith } from './commands/compareActiveWith';
import { compareTwoFiles } from './commands/compareTwoFiles';
import {
	COMMAND_COMPARE_ACTIVE_WITH,
	COMMAND_COMPARE_TWO_FILES,
} from './constants';
import { XlsxDiffUriHandler } from './git/uriHandler';

export function activate(context: vscode.ExtensionContext) {
	context.subscriptions.push(
		vscode.window.registerUriHandler(
			new XlsxDiffUriHandler(context.extensionUri),
		),
		vscode.commands.registerCommand(COMMAND_COMPARE_TWO_FILES, async () => {
			await compareTwoFiles(context.extensionUri);
		}),
		vscode.commands.registerCommand(
			COMMAND_COMPARE_ACTIVE_WITH,
			async (resource?: vscode.Uri) => {
				await compareActiveWith(context.extensionUri, resource);
			},
		),
	);
}

export function deactivate() {}
