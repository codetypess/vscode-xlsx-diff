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

function getNonce(): string {
	return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function toErrorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

export class XlsxDiffPanel {
	private readonly panel: vscode.WebviewPanel;
	private readonly extensionUri: vscode.Uri;
	private readonly disposables: vscode.Disposable[] = [];

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
	) {
		this.panel = panel;
		this.extensionUri = extensionUri;
		this.leftFileUri = leftFileUri;
		this.rightFileUri = rightFileUri;

		this.panel.webview.html = this.getHtml();
		this.panel.onDidDispose(() => this.dispose(), null, this.disposables);
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
	): Promise<void> {
		const panel = vscode.window.createWebviewPanel(
			WEBVIEW_TYPE_DIFF_PANEL,
			`${path.basename(leftFileUri.fsPath)} ↔ ${path.basename(rightFileUri.fsPath)}`,
			vscode.ViewColumn.Active,
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
		);
		await instance.reloadModel();
	}

	private dispose(): void {
		for (const disposable of this.disposables) {
			disposable.dispose();
		}
	}

	private getHtml(): string {
		const webview = this.panel.webview;
		const nonce = getNonce();
		const scriptUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.js'),
		);
		const styleUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.extensionUri, 'media', 'panel.css'),
		);

		return `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} https: data:; script-src 'nonce-${nonce}'; style-src ${webview.cspSource};" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="${styleUri}" />
	<title>XLSX Diff</title>
</head>
<body>
	<div id="app" class="loading-shell">
		<div class="loading-shell__message">Loading XLSX diff...</div>
	</div>
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
			loadWorkbookSnapshot(this.leftFileUri.fsPath),
			loadWorkbookSnapshot(this.rightFileUri.fsPath),
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
