import * as vscode from 'vscode';
import { XlsxDiffPanel } from '../webview/diffPanel';
import { getScmWorkbookDiffUrisFromTabInput } from '../workbook/resourceUri';

export function registerScmWorkbookDiffInterceptor(
	extensionUri: vscode.Uri,
): vscode.Disposable {
	const inFlight = new Set<string>();

	const maybeInterceptTab = async (tab: vscode.Tab | undefined): Promise<void> => {
		if (!tab?.isActive) {
			return;
		}

		const diffUris = getScmWorkbookDiffUrisFromTabInput(tab.input);
		if (!diffUris) {
			return;
		}

		const requestKey = `${diffUris.original.toString()}::${diffUris.modified.toString()}`;
		if (inFlight.has(requestKey)) {
			return;
		}

		inFlight.add(requestKey);
		try {
			await XlsxDiffPanel.create(
				extensionUri,
				diffUris.original,
				diffUris.modified,
				tab.group.viewColumn,
			);
			await vscode.window.tabGroups.close(tab, true);
		} finally {
			inFlight.delete(requestKey);
		}
	};

	const handleTabChange = (event: vscode.TabChangeEvent) => {
		for (const tab of [...event.opened, ...event.changed]) {
			void maybeInterceptTab(tab);
		}

		for (const group of vscode.window.tabGroups.all) {
			void maybeInterceptTab(group.activeTab);
		}
	};

	void maybeInterceptTab(vscode.window.tabGroups.activeTabGroup.activeTab);
	return vscode.window.tabGroups.onDidChangeTabs(handleTabChange);
}
