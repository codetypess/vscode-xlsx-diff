import * as path from 'node:path';
import * as vscode from 'vscode';
import { getColumnNumber } from '../core/model/cells';
import { loadWorkbookSnapshot } from '../core/fastxlsx/loadWorkbookSnapshot';
import { writeCellValues, type CellEdit } from '../core/fastxlsx/writeCellValue';
import type {
	EditorPanelState,
	EditorRenderModel,
	WorkbookSnapshot,
} from '../core/model/types';
import { getHtmlLanguageTag, isChineseDisplayLanguage } from '../displayLanguage';
import {
	createEditorRenderModel,
	createInitialEditorPanelState,
	getEditorSheetKey,
	moveEditorPageCursor,
	normalizeEditorPanelState,
	setActiveEditorSheet,
	setEditorCurrentPage,
	setSelectedEditorCell,
} from './editorRenderModel';
import { getWorkbookResourceName } from '../workbook/resourceUri';

type WebviewMessage =
	| { type: 'ready' }
	| { type: 'setSheet'; sheetKey: string }
	| { type: 'setPage'; page: number }
	| { type: 'prevPage' }
	| { type: 'nextPage' }
	| { type: 'search'; query: string; direction: 'next' | 'prev' }
	| { type: 'gotoCell'; reference: string }
	| { type: 'selectCell'; rowNumber: number; columnNumber: number }
	| { type: 'saveEdits'; edits: Array<{ sheetKey: string; rowNumber: number; columnNumber: number; value: string }> }
	| { type: 'pendingEditStateChanged'; hasPendingEdits: boolean }
	| { type: 'reload' };

interface WebviewStrings {
	loading: string;
	reload: string;
	prevPage: string;
	nextPage: string;
	size: string;
	modified: string;
	sheet: string;
	rows: string;
	noRows: string;
	page: string;
	visibleRows: string;
	readOnly: string;
	save: string;
	undo: string;
	redo: string;
	searchPlaceholder: string;
	findPrev: string;
	findNext: string;
	gotoPlaceholder: string;
	goto: string;
	totalSheets: string;
	totalRows: string;
	nonEmptyCells: string;
	selectedCell: string;
	noCellSelected: string;
	mergedRanges: string;
	pendingChanges: string;
	noRowsAvailable: string;
	readOnlyBadge: string;
	localChangesBlockedReload: string;
	confirmReloadDiscard: string;
	discardChangesAndReload: string;
	keepEditing: string;
	displayLanguageRefreshBlocked: string;
	noSearchMatches: string;
	invalidCellReference: string;
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
	if (isChineseDisplayLanguage()) {
		return {
			loading: '正在加载 XLSX 编辑器...',
			reload: '刷新',
			undo: '撤销',
			redo: '重做',
			searchPlaceholder: '搜索值或公式',
			findPrev: '上一个',
			findNext: '下一个',
			gotoPlaceholder: 'A1 或 Sheet1!B2',
			goto: '定位',
			prevPage: '上一页',
			nextPage: '下一页',
			size: '大小',
			modified: '修改时间',
			sheet: '工作表',
			rows: '行',
			noRows: '无行',
			page: '页码',
			visibleRows: '可见行',
			readOnly: '只读',
			save: '保存',
			totalSheets: '总工作表',
			totalRows: '总行数',
			nonEmptyCells: '非空单元格',
			selectedCell: '当前单元格',
			noCellSelected: '未选择',
			mergedRanges: '合并区域',
			pendingChanges: '待保存修改',
			noRowsAvailable: '当前页没有可显示的行。',
			readOnlyBadge: '只读模式',
			localChangesBlockedReload: '工作簿文件已在磁盘上变化。请先保存或放弃当前未保存修改，再刷新。',
			confirmReloadDiscard: '刷新会丢弃当前未保存修改，是否继续？',
			discardChangesAndReload: '放弃修改并刷新',
			keepEditing: '继续编辑',
			displayLanguageRefreshBlocked: '当前有未保存修改，语言变更将在保存或手动刷新后生效。',
			noSearchMatches: '没有找到匹配的单元格。',
			invalidCellReference: '无法定位该单元格，请使用 A1 或 Sheet1!B2 格式，并确保目标在当前工作簿范围内。',
		};
	}

	return {
		loading: 'Loading XLSX editor...',
		reload: 'Reload',
		undo: 'Undo',
		redo: 'Redo',
		searchPlaceholder: 'Search values or formulas',
		findPrev: 'Prev Match',
		findNext: 'Next Match',
		gotoPlaceholder: 'A1 or Sheet1!B2',
		goto: 'Go',
		prevPage: 'Prev Page',
		nextPage: 'Next Page',
		size: 'Size',
		modified: 'Modified',
		sheet: 'Sheet',
		rows: 'Rows',
		noRows: 'No rows',
		page: 'Page',
		visibleRows: 'Visible rows',
		readOnly: 'Read-only',
		save: 'Save',
		totalSheets: 'Sheets',
		totalRows: 'Rows',
		nonEmptyCells: 'Non-empty cells',
		selectedCell: 'Selected cell',
		noCellSelected: 'None',
		mergedRanges: 'Merged ranges',
		pendingChanges: 'Pending changes',
		noRowsAvailable: 'No rows available on this page.',
		readOnlyBadge: 'Read-only',
		localChangesBlockedReload: 'The workbook changed on disk. Save or discard your pending edits before reloading.',
		confirmReloadDiscard: 'Reloading will discard your pending edits. Continue?',
		discardChangesAndReload: 'Discard Changes and Reload',
		keepEditing: 'Keep Editing',
		displayLanguageRefreshBlocked: 'Pending edits are open. Display language changes will apply after you save or reload the editor.',
		noSearchMatches: 'No matching cells were found.',
		invalidCellReference: 'Unable to locate that cell. Use A1 or Sheet1!B2 and stay within the workbook range.',
	};
}

export class XlsxEditorPanel {
	private static readonly panels = new Map<number, XlsxEditorPanel>();
	private static nextPanelId = 1;

	private readonly panel: vscode.WebviewPanel;
	private readonly extensionUri: vscode.Uri;
	private readonly disposables: vscode.Disposable[] = [];
	private readonly fileWatchers: vscode.Disposable[] = [];
	private readonly panelId: number;

	private workbookUri: vscode.Uri;
	private workbook: WorkbookSnapshot | null = null;
	private state: EditorPanelState = {
		activeSheetKey: null,
		currentPage: 1,
		selectedCell: null,
	};
	private isWebviewReady = false;
	private hasPendingRender = false;
	private isReloading = false;
	private hasQueuedReload = false;
	private autoRefreshTimer: ReturnType<typeof setTimeout> | undefined;
	private suppressAutoRefreshUntil = 0;
	private hasPendingEdits = false;
	private hasWarnedPendingExternalChange = false;

	private constructor(
		panel: vscode.WebviewPanel,
		extensionUri: vscode.Uri,
		workbookUri: vscode.Uri,
		panelId: number,
	) {
		this.panel = panel;
		this.extensionUri = extensionUri;
		this.workbookUri = workbookUri;
		this.panelId = panelId;
		this.panel.webview.options = {
			enableScripts: true,
			localResourceRoots: [extensionUri],
		};

		this.panel.webview.html = this.getHtml();
		this.panel.onDidDispose(
			() => {
				XlsxEditorPanel.panels.delete(this.panelId);
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

	public static async resolveCustomEditor(
		extensionUri: vscode.Uri,
		workbookUri: vscode.Uri,
		panel: vscode.WebviewPanel,
	): Promise<void> {
		const panelId = XlsxEditorPanel.nextPanelId;
		XlsxEditorPanel.nextPanelId += 1;
		panel.title = getWorkbookResourceName(workbookUri);
		const instance = new XlsxEditorPanel(panel, extensionUri, workbookUri, panelId);
		XlsxEditorPanel.panels.set(panelId, instance);
		await instance.enqueueReload();
	}

	public static async refreshAll(): Promise<void> {
		await Promise.all(
			[...XlsxEditorPanel.panels.values()].map((panel) =>
				panel.refreshForDisplayLanguageChange(),
			),
		);
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

	private refreshFileWatchers(): void {
		this.disposeFileWatchers();

		if (this.workbookUri.scheme !== 'file') {
			return;
		}

		const watcher = vscode.workspace.createFileSystemWatcher(
			new vscode.RelativePattern(
				vscode.Uri.file(path.dirname(this.workbookUri.fsPath)),
				escapeWatcherGlobSegment(path.basename(this.workbookUri.fsPath)),
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

	private scheduleAutoRefresh(): void {
		if (Date.now() < this.suppressAutoRefreshUntil) {
			if (this.autoRefreshTimer) {
				clearTimeout(this.autoRefreshTimer);
				this.autoRefreshTimer = undefined;
			}

			return;
		}

		if (this.hasPendingEdits) {
			if (!this.hasWarnedPendingExternalChange) {
				this.hasWarnedPendingExternalChange = true;
				void vscode.window.showWarningMessage(
					getWebviewStrings().localChangesBlockedReload,
				);
			}

			return;
		}

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

	private async enqueueReload({ silent = false, clearPendingEdits = false }: { silent?: boolean; clearPendingEdits?: boolean } = {}): Promise<void> {
		if (this.isReloading) {
			this.hasQueuedReload = true;
			return;
		}

		this.isReloading = true;
		let reloadError: unknown;

		try {
			await this.reloadModel({ silent, clearPendingEdits });
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
		if (this.hasPendingEdits) {
			void vscode.window.showWarningMessage(
				getWebviewStrings().displayLanguageRefreshBlocked,
			);
			return;
		}

		this.isWebviewReady = false;
		this.hasPendingRender = Boolean(this.workbook);
		this.panel.webview.html = this.getHtml();
		await this.enqueueReload();
	}

	private getHtml(): string {
		const webview = this.panel.webview;
		const nonce = getNonce();
		const webviewStrings = getWebviewStrings();
		const strings = JSON.stringify(webviewStrings).replace(/</g, '\\u003c');
		const scriptUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'editorPanel.js'),
		);
		const styleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.css'),
		);
		const codiconStyleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'codicons', 'codicon.css'),
		);

		return `<!DOCTYPE html>
<html lang="${getHtmlLanguageTag()}">
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource};" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${codiconStyleUri}" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Editor</title>
</head>
<body>
	<div id="app" class="loading-shell">
		<div class="loading-shell__message">${webviewStrings.loading}</div>
	</div>
	<script nonce="${nonce}">window.__XLSX_EDITOR_STRINGS__ = ${strings};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
	}

	private async confirmDiscardPendingEdits(): Promise<boolean> {
		if (!this.hasPendingEdits) {
			return true;
		}

		const strings = getWebviewStrings();
		const choice = await vscode.window.showWarningMessage(
			strings.confirmReloadDiscard,
			{ modal: true },
			strings.discardChangesAndReload,
			strings.keepEditing,
		);

		return choice === strings.discardChangesAndReload;
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
					if (!this.workbook) {
						return;
					}
					this.state = setActiveEditorSheet(this.workbook, this.state, message.sheetKey);
					await this.render();
					return;
				case 'setPage':
					if (!this.workbook) {
						return;
					}
					this.state = setEditorCurrentPage(this.workbook, this.state, message.page);
					await this.render();
					return;
				case 'prevPage':
					if (!this.workbook) {
						return;
					}
					this.state = moveEditorPageCursor(this.workbook, this.state, -1);
					await this.render();
					return;
				case 'nextPage':
					if (!this.workbook) {
						return;
					}
					this.state = moveEditorPageCursor(this.workbook, this.state, 1);
					await this.render();
					return;
				case 'search': {
					if (!this.workbook) {
						return;
					}

					const match = this.findSearchMatch(message.query, message.direction);
					if (!match) {
						await vscode.window.showInformationMessage(
							getWebviewStrings().noSearchMatches,
						);
						return;
					}

					this.revealCell(match.sheetKey, match.rowNumber, match.columnNumber);
					await this.render();
					return;
				}
				case 'gotoCell':
					if (!this.workbook) {
						return;
					}

					if (!this.gotoCellReference(message.reference)) {
						await vscode.window.showInformationMessage(
							getWebviewStrings().invalidCellReference,
						);
						return;
					}

					await this.render();
					return;
				case 'selectCell':
					if (!this.workbook) {
						return;
					}
					this.state = setSelectedEditorCell(
						this.workbook,
						this.state,
						message.rowNumber,
						message.columnNumber,
					);
					return;
				case 'pendingEditStateChanged':
					this.hasPendingEdits = message.hasPendingEdits;
					if (!message.hasPendingEdits) {
						this.hasWarnedPendingExternalChange = false;
					}
					return;
				case 'saveEdits': {
					if (!this.workbook || message.edits.length === 0) {
						return;
					}

					const cellEdits: CellEdit[] = message.edits.flatMap((edit) => {
						const sheet = this.workbook!.sheets.find(
							(candidate, index) => getEditorSheetKey(candidate, index) === edit.sheetKey,
						);

						return sheet
							? [{ sheetName: sheet.name, rowNumber: edit.rowNumber, columnNumber: edit.columnNumber, value: edit.value }]
							: [];
					});

					if (cellEdits.length === 0) {
						return;
					}

					this.suppressAutoRefreshUntil = Date.now() + 2000;
					await writeCellValues(this.workbookUri, cellEdits);
					this.hasPendingEdits = false;
					this.hasWarnedPendingExternalChange = false;
					await this.enqueueReload({ silent: true, clearPendingEdits: true });
					return;
				}
				case 'reload':
					if (!(await this.confirmDiscardPendingEdits())) {
						return;
					}
					await this.enqueueReload({ clearPendingEdits: true });
					return;
			}
		} catch (error) {
			await this.handleError(error);
		}
	}

	private revealCell(
		sheetKey: string,
		rowNumber: number,
		columnNumber: number,
	): void {
		if (!this.workbook) {
			return;
		}

		this.state = setSelectedEditorCell(
			this.workbook,
			setActiveEditorSheet(this.workbook, this.state, sheetKey),
			rowNumber,
			columnNumber,
		);
	}

	private getSheetEntries(): Array<{
		key: string;
		index: number;
		sheet: WorkbookSnapshot['sheets'][number];
	}> {
		if (!this.workbook) {
			return [];
		}

		return this.workbook.sheets.map((sheet, index) => ({
			key: getEditorSheetKey(sheet, index),
			index,
			sheet,
		}));
	}

	private findSearchMatch(
		query: string,
		direction: 'next' | 'prev',
	): { sheetKey: string; rowNumber: number; columnNumber: number } | null {
		const normalizedQuery = query.trim().toLocaleLowerCase();
		if (!this.workbook || !normalizedQuery) {
			return null;
		}

		const sheetEntries = this.getSheetEntries();
		const matches = sheetEntries.flatMap((entry) =>
			Object.values(entry.sheet.cells)
				.filter((cell) => {
					const value = cell.displayValue.toLocaleLowerCase();
					const formula = cell.formula?.toLocaleLowerCase() ?? '';
					return value.includes(normalizedQuery) || formula.includes(normalizedQuery);
				})
				.map((cell) => ({
					sheetKey: entry.key,
					sheetIndex: entry.index,
					rowNumber: cell.rowNumber,
					columnNumber: cell.columnNumber,
				})),
		);

		if (matches.length === 0) {
			return null;
		}

		matches.sort((left, right) => {
			if (left.sheetIndex !== right.sheetIndex) {
				return left.sheetIndex - right.sheetIndex;
			}

			if (left.rowNumber !== right.rowNumber) {
				return left.rowNumber - right.rowNumber;
			}

			return left.columnNumber - right.columnNumber;
		});

		const activeSheetIndex = sheetEntries.findIndex(
			(entry) => entry.key === this.state.activeSheetKey,
		);
		const anchor = {
			sheetIndex: activeSheetIndex < 0 ? 0 : activeSheetIndex,
			rowNumber: this.state.selectedCell?.rowNumber ?? 1,
			columnNumber: this.state.selectedCell?.columnNumber ?? 1,
		};
		const compare = (
			candidate: { sheetIndex: number; rowNumber: number; columnNumber: number },
			current: { sheetIndex: number; rowNumber: number; columnNumber: number },
		): number => {
			if (candidate.sheetIndex !== current.sheetIndex) {
				return candidate.sheetIndex - current.sheetIndex;
			}

			if (candidate.rowNumber !== current.rowNumber) {
				return candidate.rowNumber - current.rowNumber;
			}

			return candidate.columnNumber - current.columnNumber;
		};

		if (direction === 'prev') {
			for (let index = matches.length - 1; index >= 0; index -= 1) {
				if (compare(matches[index], anchor) < 0) {
					return matches[index];
				}
			}

			return matches[matches.length - 1];
		}

		return matches.find((match) => compare(match, anchor) > 0) ?? matches[0];
	}

	private gotoCellReference(reference: string): boolean {
		if (!this.workbook) {
			return false;
		}

		const trimmedReference = reference.trim();
		if (!trimmedReference) {
			return false;
		}

		const separatorIndex = trimmedReference.lastIndexOf('!');
		const sheetName = separatorIndex > 0 ? trimmedReference.slice(0, separatorIndex).trim() : null;
		const address = separatorIndex > 0 ? trimmedReference.slice(separatorIndex + 1).trim() : trimmedReference;
		const addressMatch = /^([A-Za-z]+)(\d+)$/.exec(address);
		if (!addressMatch) {
			return false;
		}

		const columnNumber = getColumnNumber(addressMatch[1]);
		const rowNumber = Number(addressMatch[2]);
		if (!columnNumber || rowNumber < 1) {
			return false;
		}

		const sheetEntries = this.getSheetEntries();
		const targetSheet = sheetName
			? sheetEntries.find(
				(entry) =>
					entry.sheet.name === sheetName ||
					entry.sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase(),
			  )
			: sheetEntries.find((entry) => entry.key === this.state.activeSheetKey) ?? sheetEntries[0];

		if (
			!targetSheet ||
			rowNumber > targetSheet.sheet.rowCount ||
			columnNumber > targetSheet.sheet.columnCount
		) {
			return false;
		}

		this.revealCell(targetSheet.key, rowNumber, columnNumber);
		return true;
	}

	private async reloadModel({ silent = false, clearPendingEdits = false }: { silent?: boolean; clearPendingEdits?: boolean } = {}): Promise<void> {
		const webviewStrings = getWebviewStrings();

		if (!silent) {
			this.panel.title = webviewStrings.loading;

			if (this.isWebviewReady) {
				await this.panel.webview.postMessage({
					type: 'loading',
					message: webviewStrings.loading,
				});
			}
		}

		this.workbook = await loadWorkbookSnapshot(this.workbookUri);
		this.state = this.workbook.sheets.length
			? normalizeEditorPanelState(
					this.workbook,
					this.state.activeSheetKey
						? this.state
						: createInitialEditorPanelState(this.workbook),
				)
			: createInitialEditorPanelState(this.workbook);

		const renderModel = createEditorRenderModel(this.workbook, this.state, {
			hasPendingEdits: clearPendingEdits ? false : this.hasPendingEdits,
		});

		if (clearPendingEdits) {
			this.hasWarnedPendingExternalChange = false;
		}
		this.panel.title = renderModel.title;
		await this.render(renderModel, { silent, clearPendingEdits });
	}

	private async render(renderModel?: EditorRenderModel, { silent = false, clearPendingEdits = false }: { silent?: boolean; clearPendingEdits?: boolean } = {}): Promise<void> {
		if (!this.workbook) {
			return;
		}

		const payload = renderModel ?? createEditorRenderModel(this.workbook, this.state, {
			hasPendingEdits: this.hasPendingEdits,
		});
		this.panel.title = payload.title;

		if (!this.isWebviewReady) {
			this.hasPendingRender = true;
			return;
		}

		this.hasPendingRender = false;
		await this.panel.webview.postMessage({
			type: 'render',
			payload,
			silent,
			clearPendingEdits,
		});
	}
}