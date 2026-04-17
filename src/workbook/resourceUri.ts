import { execFile as execFileCallback } from 'node:child_process';
import * as path from 'node:path';
import { promisify } from 'node:util';
import * as vscode from 'vscode';

interface GitUriQuery {
	path?: string;
	ref?: string;
}

const execFile = promisify(execFileCallback);

function isChineseLocale(): boolean {
	return vscode.env.language.toLowerCase().startsWith('zh');
}

function getUriPathForExtension(uri: vscode.Uri): string {
	return uri.scheme === 'file' ? uri.fsPath : decodeURIComponent(uri.path);
}

function parseGitUriQuery(uri: vscode.Uri): GitUriQuery | undefined {
	if (uri.scheme !== 'git' || !uri.query) {
		return undefined;
	}

	try {
		const parsed = JSON.parse(uri.query) as GitUriQuery;
		return typeof parsed === 'object' && parsed !== null ? parsed : undefined;
	} catch {
		return undefined;
	}
}

async function runGit(cwd: string, args: string[]): Promise<string | undefined> {
	try {
		const { stdout } = await execFile('git', ['-C', cwd, ...args]);
		const trimmed = stdout.trim();
		return trimmed.length > 0 ? trimmed : undefined;
	} catch {
		return undefined;
	}
}

async function getGitRepositoryRoot(resourcePath: string): Promise<string | undefined> {
	return runGit(path.dirname(resourcePath), ['rev-parse', '--show-toplevel']);
}

async function hasStagedChanges(
	repositoryRoot: string,
	resourcePath: string,
): Promise<boolean> {
	const relativePath = path.relative(repositoryRoot, resourcePath);
	const changedFiles = await runGit(repositoryRoot, [
		'diff',
		'--cached',
		'--name-only',
		'--',
		relativePath,
	]);
	return Boolean(changedFiles);
}

async function resolveShortCommit(
	repositoryRoot: string,
	ref: string,
): Promise<string | undefined> {
	return runGit(repositoryRoot, ['rev-parse', '--short', ref]);
}

export function describeGitResourceRef(
	ref: string,
	options: {
		resolvedCommit?: string;
		hasStagedChanges?: boolean;
	} = {},
): { label: string; value: string } {
	const isChinese = isChineseLocale();
	const sourceLabel = isChinese ? '来源' : 'Source';
	const commitLabel = isChinese ? '提交' : 'Commit';
	const indexLabel = isChinese ? '暂存区' : 'Index';

	if (ref === '') {
		return { label: sourceLabel, value: indexLabel };
	}

	if (/^~\d$/.test(ref)) {
		return {
			label: sourceLabel,
			value: isChinese ? `阶段 ${ref[1]}` : `Stage ${ref[1]}`,
		};
	}

	if (ref === '~') {
		if (options.hasStagedChanges) {
			return {
				label: sourceLabel,
				value: options.resolvedCommit
					? isChinese
						? `暂存区 · 基线 ${options.resolvedCommit}`
						: `Index · base ${options.resolvedCommit}`
					: indexLabel,
			};
		}

		return {
			label: commitLabel,
			value: options.resolvedCommit ?? 'HEAD',
		};
	}

	if (options.resolvedCommit) {
		return {
			label: commitLabel,
			value: options.resolvedCommit,
		};
	}

	return {
		label: sourceLabel,
		value: isChinese ? `Git 引用: ${ref}` : `Git ref: ${ref}`,
	};
}

export function isWorkbookResourceUri(uri: vscode.Uri | undefined): uri is vscode.Uri {
	return Boolean(uri && path.extname(uri.path).toLowerCase() === '.xlsx');
}

export function getWorkbookResourceName(uri: vscode.Uri): string {
	return path.basename(getUriPathForExtension(uri));
}

export function getWorkbookResourcePathLabel(uri: vscode.Uri): string {
	const resourcePath = getUriPathForExtension(uri);
	const gitQuery = parseGitUriQuery(uri);
	return gitQuery?.ref ? `${resourcePath} @ ${gitQuery.ref}` : resourcePath;
}

export function getWorkbookResourceTimeLabel(uri: vscode.Uri): string | undefined {
	const gitQuery = parseGitUriQuery(uri);
	if (gitQuery && gitQuery.ref !== undefined) {
		return `Git ref: ${gitQuery.ref}`;
	}

	return uri.scheme === 'file' ? undefined : `${uri.scheme.toUpperCase()} resource`;
}

export async function getWorkbookResourceDetail(
	uri: vscode.Uri,
): Promise<{ label: string; value: string } | undefined> {
	const gitQuery = parseGitUriQuery(uri);
	if (!gitQuery || gitQuery.ref === undefined) {
		return undefined;
	}

	const resourcePath = gitQuery.path ?? getUriPathForExtension(uri);
	const repositoryRoot = await getGitRepositoryRoot(resourcePath);
	if (!repositoryRoot) {
		return describeGitResourceRef(gitQuery.ref);
	}

	if (gitQuery.ref === '~') {
		const [resolvedCommit, stagedChanges] = await Promise.all([
			resolveShortCommit(repositoryRoot, 'HEAD'),
			hasStagedChanges(repositoryRoot, resourcePath),
		]);
		return describeGitResourceRef(gitQuery.ref, {
			resolvedCommit,
			hasStagedChanges: stagedChanges,
		});
	}

	if (gitQuery.ref === '' || /^~\d$/.test(gitQuery.ref)) {
		return describeGitResourceRef(gitQuery.ref);
	}

	const resolvedCommit = await resolveShortCommit(repositoryRoot, gitQuery.ref);
	return describeGitResourceRef(gitQuery.ref, { resolvedCommit });
}

export function getWorkbookDiffUrisFromTabInput(
	input: unknown,
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
	if (!(input instanceof vscode.TabInputTextDiff)) {
		return undefined;
	}

	if (
		!isWorkbookResourceUri(input.original) ||
		!isWorkbookResourceUri(input.modified)
	) {
		return undefined;
	}

	return {
		original: input.original,
		modified: input.modified,
	};
}

export function getScmWorkbookDiffUrisFromTabInput(
	input: unknown,
): { original: vscode.Uri; modified: vscode.Uri } | undefined {
	const diffUris = getWorkbookDiffUrisFromTabInput(input);
	if (!diffUris) {
		return undefined;
	}

	if (
		diffUris.original.scheme === 'file' &&
		diffUris.modified.scheme === 'file'
	) {
		return undefined;
	}

	return diffUris;
}
