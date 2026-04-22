import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import { XlsxDiffPanel } from "../webview/diff-panel";
import { getActiveWorkbookUri, getWorkbookUriFromCommandArg, pickWorkbook } from "./workbook-picker";

export async function compareActiveWith(
    extensionUri: vscode.Uri,
    resource?: unknown
): Promise<void> {
    const isChinese = isChineseDisplayLanguage();
    const leftUri = getWorkbookUriFromCommandArg(resource) ?? getActiveWorkbookUri();

    if (!leftUri) {
        await vscode.window.showErrorMessage(
            isChinese
                ? "请先打开一个 .xlsx 文件，或从命令面板运行“比较两个 XLSX 文件”。"
                : 'Open an .xlsx file first, or run "Compare Two XLSX Files" from the Command Palette.'
        );
        return;
    }

    const rightUri = await pickWorkbook(
        isChinese
            ? "选择要与当前文件比较的 XLSX 工作簿"
            : "Select the XLSX workbook to compare against",
        leftUri
    );
    if (!rightUri) {
        return;
    }

    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
