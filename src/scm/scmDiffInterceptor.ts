import * as vscode from 'vscode';
import { XlsxDiffPanel } from '../webview/diffPanel';
import { getScmWorkbookDiffUrisFromTabInput } from '../workbook/resourceUri';

function getTabResourceUri(input: vscode.Tab['input']): vscode.Uri | undefined {
	if (input instanceof vscode.TabInputText) {
		return input.uri;
	}

	if (input instanceof vscode.TabInputCustom) {
		return input.uri;
	}

	if (input instanceof vscode.TabInputNotebook) {
		return input.uri;
	}

	return undefined;
}

async function closePreviewWorkbookTabs(resourceUri: vscode.Uri): Promise<void> {
	const tabsToClose = vscode.window.tabGroups.all.flatMap((group) =>
		group.tabs.filter((tab) => {
			if (tab.isDirty) {
				return false;
			}

			const tabResourceUri = getTabResourceUri(tab.input);
			if (tabResourceUri?.toString() !== resourceUri.toString()) {
				return false;
			}

			return tab.isPreview || tab.input instanceof vscode.TabInputText;
		}),
	);

	if (tabsToClose.length > 0) {
		await vscode.window.tabGroups.close(tabsToClose, true);
	}
}

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
			const openPanelPromise = XlsxDiffPanel.create(
				extensionUri,
				diffUris.original,
				diffUris.modified,
				tab.group.viewColumn,
			);
			await vscode.window.tabGroups.close(tab, true);
			await closePreviewWorkbookTabs(diffUris.modified);
			await openPanelPromise;
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
