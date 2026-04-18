import * as assert from 'assert';
import * as vscode from 'vscode';
import {
	COMMAND_COMPARE_ACTIVE_WITH,
	COMMAND_COMPARE_TWO_FILES,
	COMMAND_OPEN_EDITOR,
} from '../constants';

suite('Extension Test Suite', () => {
	test('registers compare commands', async () => {
		const extension = vscode.extensions.all.find(
			(candidate) => candidate.packageJSON.name === 'xlsx-diff',
		);
		assert.ok(extension);
		await extension.activate();

		const commands = await vscode.commands.getCommands(true);
		assert.ok(commands.includes(COMMAND_COMPARE_TWO_FILES));
		assert.ok(commands.includes(COMMAND_COMPARE_ACTIVE_WITH));
		assert.ok(commands.includes(COMMAND_OPEN_EDITOR));
	});
});
