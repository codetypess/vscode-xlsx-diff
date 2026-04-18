import * as path from 'node:path';
import * as vscode from 'vscode';
import { WEBVIEW_TYPE_DIFF_PANEL } from '../constants';
import { buildWorkbookDiff } from '../core/diff/buildWorkbookDiff';
import { loadWorkbookSnapshot } from '../core/fastxlsx/loadWorkbookSnapshot';
import { writeCellValue } from '../core/fastxlsx/writeCellValue';
import type {
	PanelState,
	RenderModel,
	RowFilterMode,
	WorkbookDiffModel,
} from '../core/model/types';
import { getHtmlLanguageTag, isChineseDisplayLanguage } from '../displayLanguage';
import {
	createInitialPanelState,
	createRenderModel,
	movePageCursor,
	moveDiffCursor,
	normalizePanelState,
	setActiveSheet,
	setCurrentPage,
	setFilterMode,
	setHighlightedDiffCell,
	setHighlightedDiffRow,
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
	| { type: 'selectCell'; rowNumber: number; columnNumber: number }
	| { type: 'editCell'; side: 'left' | 'right'; rowNumber: number; columnNumber: number; value: string }
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
	noRows: string;
	page: string;
	filter: string;
	diffCells: string;
	diffRows: string;
	sameRows: string;
	visibleRows: string;
	readOnly: string;
}

function getNonce(): string {
	return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function toErrorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

function escapeWatcherGlobSegment(value: string): string {
	return value.replace(/[{}\[\]*?]/g, '[$&]');
}

function getWebviewStrings(): WebviewStrings {
	const isChinese = isChineseDisplayLanguage();
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
			noRows: '无行',
			page: '页码',
			filter: '筛选',
			diffCells: '差异单元格',
			diffRows: '差异行',
			sameRows: '相同行',
			visibleRows: '可见行',
			readOnly: '只读',
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
		noRows: 'No rows',
		page: 'Page',
		filter: 'Filter',
		diffCells: 'Diff cells',
		diffRows: 'Diff rows',
		sameRows: 'Same rows',
		visibleRows: 'Visible rows',
		readOnly: 'Read-only',
	};
}

export class XlsxDiffPanel {
	private static readonly panels = new Map<string, XlsxDiffPanel>();

	private readonly panel: vscode.WebviewPanel;
	private readonly extensionUri: vscode.Uri;
	private readonly disposables: vscode.Disposable[] = [];
	private readonly fileWatchers: vscode.Disposable[] = [];
	private readonly panelKey: string;

	private leftFileUri: vscode.Uri;
	private rightFileUri: vscode.Uri;
	private diffModel: WorkbookDiffModel | null = null;
	private state: PanelState = {
		activeSheetKey: null,
		filter: 'all',
		currentPage: 1,
		highlightedDiffCellKey: null,
	};
	private isWebviewReady = false;
	private hasPendingRender = false;
	private isReloading = false;
	private hasQueuedReload = false;
	private autoRefreshTimer: ReturnType<typeof setTimeout> | undefined;

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
		this.refreshFileWatchers();
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
			existingPanel.setFileUris(leftFileUri, rightFileUri);
			existingPanel.panel.reveal(viewColumn, true);
			await existingPanel.enqueueReload();
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
		await instance.enqueueReload();
	}

	public static async refreshAll(): Promise<void> {
		await Promise.all(
			[...XlsxDiffPanel.panels.values()].map((panel) =>
				panel.refreshForDisplayLanguageChange(),
			),
		);
	}

	private static getPanelKey(leftFileUri: vscode.Uri, rightFileUri: vscode.Uri): string {
		return [leftFileUri.toString(), rightFileUri.toString()].sort().join('::');
	}

	private dispose(): void {
		if (this.autoRefreshTimer) {
			clearTimeout(this.autoRefreshTimer);
			this.autoRefreshTimer = undefined;
		}

		this.disposeFileWatchers();

		for (const disposable of this.disposables) {
			disposable.dispose();
		}
	}

	private disposeFileWatchers(): void {
		for (const disposable of this.fileWatchers) {
			disposable.dispose();
		}

		this.fileWatchers.length = 0;
	}

	private setFileUris(leftFileUri: vscode.Uri, rightFileUri: vscode.Uri): void {
		this.leftFileUri = leftFileUri;
		this.rightFileUri = rightFileUri;
		this.refreshFileWatchers();
	}

	private refreshFileWatchers(): void {
		this.disposeFileWatchers();

		const watchTargets = new Map<string, vscode.Uri>();
		for (const uri of [this.leftFileUri, this.rightFileUri]) {
			if (uri.scheme !== 'file') {
				continue;
			}

			watchTargets.set(uri.toString(), uri);
		}

		for (const uri of watchTargets.values()) {
			const watcher = vscode.workspace.createFileSystemWatcher(
				new vscode.RelativePattern(
					vscode.Uri.file(path.dirname(uri.fsPath)),
					escapeWatcherGlobSegment(path.basename(uri.fsPath)),
				),
			);
			const scheduleRefresh = () => {
				this.scheduleAutoRefresh();
			};

			this.fileWatchers.push(watcher);
			this.fileWatchers.push(watcher.onDidChange(scheduleRefresh));
			this.fileWatchers.push(watcher.onDidCreate(scheduleRefresh));
			this.fileWatchers.push(watcher.onDidDelete(scheduleRefresh));
		}
	}

	private scheduleAutoRefresh(): void {
		if (this.autoRefreshTimer) {
			clearTimeout(this.autoRefreshTimer);
		}

		this.autoRefreshTimer = setTimeout(() => {
			this.autoRefreshTimer = undefined;
			void this.enqueueReload().catch((error) => {
				void this.handleError(error);
			});
		}, 250);
	}

	private async enqueueReload(): Promise<void> {
		if (this.isReloading) {
			this.hasQueuedReload = true;
			return;
		}

		this.isReloading = true;
		let reloadError: unknown;

		try {
			await this.reloadModel();
		} catch (error) {
			reloadError = error;
		} finally {
			this.isReloading = false;

			if (this.hasQueuedReload) {
				this.hasQueuedReload = false;
				await this.enqueueReload();
			}
		}

		if (reloadError) {
			throw reloadError;
		}
	}

	private async handleError(error: unknown): Promise<void> {
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

	private async refreshForDisplayLanguageChange(): Promise<void> {
		this.isWebviewReady = false;
		this.hasPendingRender = Boolean(this.diffModel);
		this.panel.webview.html = this.getHtml();
		await this.enqueueReload();
	}

	private getHtml(): string {
		const webview = this.panel.webview;
		const nonce = getNonce();
		const webviewStrings = getWebviewStrings();
		const strings = JSON.stringify(webviewStrings).replace(/</g, '\\u003c');
		const scriptUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.js'),
		);
		const styleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.css'),
		);
		const codiconStyleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(
				this.extensionUri,
				'media',
				'codicons',
				'codicon.css',
			),
		);

		return `<!DOCTYPE html>
<html lang="${getHtmlLanguageTag()}">
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
		<div class="loading-shell__message">${webviewStrings.loading}</div>
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
					this.state = movePageCursor(this.diffModel, this.state, -1);
					await this.render();
					return;
				case 'nextPage':
					if (!this.diffModel) {
						return;
					}
					this.state = movePageCursor(this.diffModel, this.state, 1);
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
				case 'selectCell':
					if (!this.diffModel) {
						return;
					}
					this.state = setHighlightedDiffCell(
						this.diffModel,
						this.state,
						message.rowNumber,
						message.columnNumber,
					);
					await this.render();
					return;
				case 'editCell': {
					if (!this.diffModel) {
						return;
					}

					const activeSheet = this.diffModel.sheets.find(
						(sheet) => sheet.key === this.state.activeSheetKey,
					);
					if (!activeSheet) {
						return;
					}

					const sheetSnapshot = message.side === 'left'
						? activeSheet.leftSheet
						: activeSheet.rightSheet;
					if (!sheetSnapshot) {
						return;
					}

					const fileUri = message.side === 'left' ? this.leftFileUri : this.rightFileUri;
					await writeCellValue(
						fileUri,
						sheetSnapshot.name,
						message.rowNumber,
						message.columnNumber,
						message.value,
					);
					await this.enqueueReload();
					return;
				}
				case 'swap': {
					this.setFileUris(this.rightFileUri, this.leftFileUri);
					await this.enqueueReload();
					return;
				}
				case 'reload':
					await this.enqueueReload();
					return;
			}
		} catch (error) {
			await this.handleError(error);
		}
	}

	private async reloadModel(): Promise<void> {
		const webviewStrings = getWebviewStrings();
		this.panel.title = webviewStrings.loading;

		if (this.isWebviewReady) {
			await this.panel.webview.postMessage({
				type: 'loading',
				message: webviewStrings.loading,
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

		const renderModel = createRenderModel(this.diffModel, this.state);
		this.panel.title = renderModel.title;
		await this.render(renderModel);
	}

	private async render(renderModel?: RenderModel): Promise<void> {
		if (!this.diffModel) {
			return;
		}

		const payload = renderModel ?? createRenderModel(this.diffModel, this.state);
		this.panel.title = payload.title;

		if (!this.isWebviewReady) {
			this.hasPendingRender = true;
			return;
		}

		this.hasPendingRender = false;
		await this.panel.webview.postMessage({
			type: 'render',
			payload,
		});
	}
}
