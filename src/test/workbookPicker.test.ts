import * as assert from 'assert';
import * as vscode from 'vscode';
import { getWorkbookUriFromCommandArg } from '../commands/workbookPicker';

suite('Workbook picker helpers', () => {
	test('extracts workbook uri from direct command arguments', () => {
		const uri = vscode.Uri.file('/tmp/sample.xlsx');
		assert.deepStrictEqual(getWorkbookUriFromCommandArg(uri), uri);
	});

	test('extracts workbook uri from scm resource arguments', () => {
		const uri = vscode.Uri.file('/tmp/sample.xlsx');
		assert.deepStrictEqual(
			getWorkbookUriFromCommandArg({ resourceUri: uri }),
			uri,
		);
	});

	test('extracts workbook uri from multi-selection arguments', () => {
		const uri = vscode.Uri.file('/tmp/sample.xlsx');
		assert.deepStrictEqual(
			getWorkbookUriFromCommandArg([
				{ resourceUri: vscode.Uri.file('/tmp/ignore.txt') },
				{ resourceUri: uri },
			]),
			uri,
		);
	});

	test('ignores non-workbook command arguments', () => {
		assert.strictEqual(
			getWorkbookUriFromCommandArg(vscode.Uri.file('/tmp/sample.txt')),
			undefined,
		);
	});
});
