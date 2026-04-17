import * as path from 'node:path';
import * as vscode from 'vscode';

interface GitUriQuery {
	path?: string;
	ref?: string;
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
	if (gitQuery?.ref) {
		return `Git ref: ${gitQuery.ref}`;
	}

	return uri.scheme === 'file' ? undefined : `${uri.scheme.toUpperCase()} resource`;
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
