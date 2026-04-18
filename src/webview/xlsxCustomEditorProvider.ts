import * as vscode from 'vscode';
import { WEBVIEW_TYPE_EDITOR_PANEL } from '../constants';
import { XlsxEditorPanel } from './xlsxEditorPanel';

class XlsxEditorDocument implements vscode.CustomDocument {
	public constructor(public readonly uri: vscode.Uri) {}

	public dispose(): void {}
}

export class XlsxCustomEditorProvider
	implements vscode.CustomReadonlyEditorProvider<XlsxEditorDocument>
{
	public static register(context: vscode.ExtensionContext): vscode.Disposable {
		return vscode.window.registerCustomEditorProvider(
			WEBVIEW_TYPE_EDITOR_PANEL,
			new XlsxCustomEditorProvider(context.extensionUri),
			{
				webviewOptions: {
					retainContextWhenHidden: true,
				},
			},
		);
	}

	private constructor(private readonly extensionUri: vscode.Uri) {}

	public openCustomDocument(uri: vscode.Uri): XlsxEditorDocument {
		return new XlsxEditorDocument(uri);
	}

	public async resolveCustomEditor(
		document: XlsxEditorDocument,
		webviewPanel: vscode.WebviewPanel,
		_token: vscode.CancellationToken,
	): Promise<void> {
		await XlsxEditorPanel.resolveCustomEditor(
			this.extensionUri,
			document.uri,
			webviewPanel,
		);
	}
}