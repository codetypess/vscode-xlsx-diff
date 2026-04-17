import * as assert from 'assert';
import * as vscode from 'vscode';
import {
	describeGitResourceRef,
	getScmWorkbookDiffUrisFromTabInput,
	getWorkbookDiffUrisFromTabInput,
	getWorkbookResourcePathLabel,
	getWorkbookResourceTimeLabel,
} from '../workbook/resourceUri';

suite('Workbook resource URIs', () => {
	test('recognizes xlsx diff tab inputs', () => {
		const original = vscode.Uri.file('/tmp/before.xlsx');
		const modified = vscode.Uri.file('/tmp/after.xlsx');
		const input = new vscode.TabInputTextDiff(original, modified);

		assert.deepStrictEqual(getWorkbookDiffUrisFromTabInput(input), {
			original,
			modified,
		});
	});

	test('ignores non-xlsx diff tab inputs', () => {
		const input = new vscode.TabInputTextDiff(
			vscode.Uri.file('/tmp/before.txt'),
			vscode.Uri.file('/tmp/after.txt'),
		);

		assert.strictEqual(getWorkbookDiffUrisFromTabInput(input), undefined);
	});

	test('extracts git ref labels for readonly workbook resources', () => {
		const gitUri = vscode.Uri.from({
			scheme: 'git',
			path: '/tmp/item.xlsx',
			query: JSON.stringify({
				path: '/tmp/item.xlsx',
				ref: 'HEAD',
			}),
		});

		assert.strictEqual(
			getWorkbookResourcePathLabel(gitUri),
			'/tmp/item.xlsx @ HEAD',
		);
		assert.strictEqual(getWorkbookResourceTimeLabel(gitUri), 'Git ref: HEAD');
	});

	test('filters scm workbook diffs to non-file originals', () => {
		const scmInput = new vscode.TabInputTextDiff(
			vscode.Uri.from({
				scheme: 'git',
				path: '/tmp/item.xlsx',
				query: JSON.stringify({
					path: '/tmp/item.xlsx',
					ref: 'HEAD',
				}),
			}),
			vscode.Uri.file('/tmp/item.xlsx'),
		);

		const fileDiffInput = new vscode.TabInputTextDiff(
			vscode.Uri.file('/tmp/left.xlsx'),
			vscode.Uri.file('/tmp/right.xlsx'),
		);

		assert.ok(getScmWorkbookDiffUrisFromTabInput(scmInput));
		assert.strictEqual(getScmWorkbookDiffUrisFromTabInput(fileDiffInput), undefined);
	});

	test('describes git refs for commit and index-backed resources', () => {
		assert.deepStrictEqual(describeGitResourceRef('HEAD', { resolvedCommit: 'd44224e' }), {
			label: 'Commit',
			value: 'd44224e',
		});
		assert.deepStrictEqual(
			describeGitResourceRef('~', {
				resolvedCommit: 'd44224e',
				hasStagedChanges: true,
			}),
			{
				label: 'Source',
				value: 'Index · base d44224e',
			},
		);
		assert.deepStrictEqual(describeGitResourceRef('~2'), {
			label: 'Source',
			value: 'Stage 2',
		});
	});
});
