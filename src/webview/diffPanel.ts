import * as path from 'node:path';
import * as vscode from 'vscode';
import { WEBVIEW_TYPE_DIFF_PANEL } from '../constants';
import { buildWorkbookDiff } from '../core/diff/buildWorkbookDiff';
import { loadWorkbookSnapshot } from '../core/fastxlsx/loadWorkbookSnapshot';
import type {
	PanelState,
	RowFilterMode,
	WorkbookDiffModel,
} from '../core/model/types';
import {
	createInitialPanelState,
	createRenderModel,
	moveDiffCursor,
	normalizePanelState,
	setActiveSheet,
	setCurrentPage,
	setFilterMode,
} from './renderModel';
import { getWorkbookResourceName } from '../workbook/resourceUri';

type WebviewMessage =
	| { type: 'ready' }
	| { type: 'setSheet'; sheetKey: string }
	| { type: 'setFilter'; filter: RowFilterMode }
	| { type: 'setPage'; page: number }
	| { type: 'prevPage' }
	| { type: 'nextPage' }
	| { type: 'prevDiff' }
	| { type: 'nextDiff' }
	| { type: 'swap' }
	| { type: 'reload' };

interface WebviewStrings {
	loading: string;
	all: string;
	diffs: string;
	same: string;
	prevDiff: string;
	nextDiff: string;
	prevPage: string;
	nextPage: string;
	swap: string;
	reload: string;
	left: string;
	right: string;
	mergedRangesChanged: string;
	noRowsAvailable: string;
	size: string;
	modified: string;
	sheet: string;
	rows: string;
	page: string;
	filter: string;
	diffRows: string;
	sameRows: string;
	visibleRows: string;
}

function getNonce(): string {
	return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function toErrorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

function getWebviewStrings(): WebviewStrings {
	const isChinese = vscode.env.language.toLowerCase().startsWith('zh');
	if (isChinese) {
		return {
			loading: '正在加载 XLSX 对比...',
			all: '全部',
			diffs: '差异',
			same: '相同',
			prevDiff: '上一处差异',
			nextDiff: '下一处差异',
			prevPage: '上一页',
			nextPage: '下一页',
			swap: '交换',
			reload: '刷新',
			left: '左侧',
			right: '右侧',
			mergedRangesChanged: '合并区域已变化',
			noRowsAvailable: '当前筛选条件下没有可显示的行。',
			size: '大小',
			modified: '修改时间',
			sheet: '工作表',
			rows: '行',
			page: '页码',
			filter: '筛选',
			diffRows: '差异行',
			sameRows: '相同行',
			visibleRows: '可见行',
		};
	}

	return {
		loading: 'Loading XLSX diff...',
		all: 'All',
		diffs: 'Diffs',
		same: 'Same',
		prevDiff: 'Prev Diff',
		nextDiff: 'Next Diff',
		prevPage: 'Prev Page',
		nextPage: 'Next Page',
		swap: 'Swap',
		reload: 'Reload',
		left: 'Left',
		right: 'Right',
		mergedRangesChanged: 'Merged ranges changed',
		noRowsAvailable: 'No rows available for this filter.',
		size: 'Size',
		modified: 'Modified',
		sheet: 'Sheet',
		rows: 'Rows',
		page: 'Page',
		filter: 'Filter',
		diffRows: 'Diff rows',
		sameRows: 'Same rows',
		visibleRows: 'Visible rows',
	};
}

export class XlsxDiffPanel {
	private static readonly panels = new Map<string, XlsxDiffPanel>();

	private readonly panel: vscode.WebviewPanel;
	private readonly extensionUri: vscode.Uri;
	private readonly disposables: vscode.Disposable[] = [];
	private readonly panelKey: string;

	private leftFileUri: vscode.Uri;
	private rightFileUri: vscode.Uri;
	private diffModel: WorkbookDiffModel | null = null;
	private state: PanelState = {
		activeSheetKey: null,
		filter: 'all',
		currentPage: 1,
		highlightedDiffRow: null,
	};
	private isWebviewReady = false;
	private hasPendingRender = false;

	private constructor(
		panel: vscode.WebviewPanel,
		extensionUri: vscode.Uri,
		leftFileUri: vscode.Uri,
		rightFileUri: vscode.Uri,
		panelKey: string,
	) {
		this.panel = panel;
		this.extensionUri = extensionUri;
		this.leftFileUri = leftFileUri;
		this.rightFileUri = rightFileUri;
		this.panelKey = panelKey;

		this.panel.webview.html = this.getHtml();
		this.panel.onDidDispose(
			() => {
				XlsxDiffPanel.panels.delete(this.panelKey);
				this.dispose();
			},
			null,
			this.disposables,
		);
		this.panel.webview.onDidReceiveMessage(
			(message: WebviewMessage) => {
				void this.handleMessage(message);
			},
			null,
			this.disposables,
		);
	}

	public static async create(
		extensionUri: vscode.Uri,
		leftFileUri: vscode.Uri,
		rightFileUri: vscode.Uri,
		viewColumn: vscode.ViewColumn = vscode.ViewColumn.Active,
	): Promise<void> {
		const panelKey = XlsxDiffPanel.getPanelKey(leftFileUri, rightFileUri);
		const existingPanel = XlsxDiffPanel.panels.get(panelKey);
		if (existingPanel) {
			existingPanel.leftFileUri = leftFileUri;
			existingPanel.rightFileUri = rightFileUri;
			existingPanel.panel.reveal(viewColumn, true);
			await existingPanel.reloadModel();
			return;
		}

		const panel = vscode.window.createWebviewPanel(
			WEBVIEW_TYPE_DIFF_PANEL,
			`${getWorkbookResourceName(leftFileUri)} ↔ ${getWorkbookResourceName(rightFileUri)}`,
			viewColumn,
			{
				enableScripts: true,
				retainContextWhenHidden: true,
				localResourceRoots: [extensionUri],
			},
		);

		const instance = new XlsxDiffPanel(
			panel,
			extensionUri,
			leftFileUri,
			rightFileUri,
			panelKey,
		);
		XlsxDiffPanel.panels.set(panelKey, instance);
		await instance.reloadModel();
	}

	private static getPanelKey(leftFileUri: vscode.Uri, rightFileUri: vscode.Uri): string {
		return [leftFileUri.toString(), rightFileUri.toString()].sort().join('::');
	}

	private dispose(): void {
		for (const disposable of this.disposables) {
			disposable.dispose();
		}
	}

	private getHtml(): string {
		const webview = this.panel.webview;
		const nonce = getNonce();
		const strings = JSON.stringify(getWebviewStrings()).replace(/</g, '\\u003c');
		const scriptUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.js'),
		);
		const styleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.css'),
		);
		const codiconStyleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(
				this.extensionUri,
				'node_modules',
				'@vscode',
				'codicons',
				'dist',
				'codicon.css',
			),
		);

		return `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource};" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${codiconStyleUri}" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Diff</title>
</head>
<body>
	<div id="app" class="loading-shell">
		<div class="loading-shell__message">${getWebviewStrings().loading}</div>
	</div>
	<script nonce="${nonce}">window.__XLSX_DIFF_STRINGS__ = ${strings};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
	}

	private async handleMessage(message: WebviewMessage): Promise<void> {
		try {
			switch (message.type) {
				case 'ready':
					this.isWebviewReady = true;
					if (this.hasPendingRender) {
						await this.render();
					}
					return;
				case 'setSheet':
					if (!this.diffModel) {
						return;
					}
					this.state = setActiveSheet(this.diffModel, this.state, message.sheetKey);
					await this.render();
					return;
				case 'setFilter':
					if (!this.diffModel) {
						return;
					}
					this.state = setFilterMode(this.diffModel, this.state, message.filter);
					await this.render();
					return;
				case 'setPage':
					if (!this.diffModel) {
						return;
					}
					this.state = setCurrentPage(this.diffModel, this.state, message.page);
					await this.render();
					return;
				case 'prevPage':
					if (!this.diffModel) {
						return;
					}
					this.state = setCurrentPage(
						this.diffModel,
						this.state,
						this.state.currentPage - 1,
					);
					await this.render();
					return;
				case 'nextPage':
					if (!this.diffModel) {
						return;
					}
					this.state = setCurrentPage(
						this.diffModel,
						this.state,
						this.state.currentPage + 1,
					);
					await this.render();
					return;
				case 'prevDiff':
					if (!this.diffModel) {
						return;
					}
					this.state = moveDiffCursor(this.diffModel, this.state, -1);
					await this.render();
					return;
				case 'nextDiff':
					if (!this.diffModel) {
						return;
					}
					this.state = moveDiffCursor(this.diffModel, this.state, 1);
					await this.render();
					return;
				case 'swap': {
					const leftFileUri = this.leftFileUri;
					this.leftFileUri = this.rightFileUri;
					this.rightFileUri = leftFileUri;
					await this.reloadModel();
					return;
				}
				case 'reload':
					await this.reloadModel();
					return;
			}
		} catch (error) {
			const errorMessage = toErrorMessage(error);
			console.error(error);
			await vscode.window.showErrorMessage(errorMessage);
			if (this.isWebviewReady) {
				await this.panel.webview.postMessage({
					type: 'error',
					message: errorMessage,
				});
			}
		}
	}

	private async reloadModel(): Promise<void> {
		this.panel.title = 'Loading XLSX diff...';

		if (this.isWebviewReady) {
			await this.panel.webview.postMessage({
				type: 'loading',
				message: 'Loading XLSX diff...',
			});
		}

		const [leftWorkbook, rightWorkbook] = await Promise.all([
			loadWorkbookSnapshot(this.leftFileUri),
			loadWorkbookSnapshot(this.rightFileUri),
		]);

		this.diffModel = buildWorkbookDiff(leftWorkbook, rightWorkbook);
		this.state = this.diffModel.sheets.length
			? normalizePanelState(
					this.diffModel,
					this.state.activeSheetKey
						? this.state
						: createInitialPanelState(this.diffModel),
				)
			: createInitialPanelState(this.diffModel);

		this.panel.title = `${leftWorkbook.fileName} ↔ ${rightWorkbook.fileName}`;
		await this.render();
	}

	private async render(): Promise<void> {
		if (!this.diffModel || !this.isWebviewReady) {
			this.hasPendingRender = true;
			return;
		}

		this.hasPendingRender = false;
		await this.panel.webview.postMessage({
			type: 'render',
			payload: createRenderModel(this.diffModel, this.state),
		});
	}
}
