import * as vscode from "vscode";
import { isChineseDisplayLanguage } from "../display-language";
import { XlsxDiffPanel } from "../webview/diff-panel";
import { pickWorkbook } from "./workbook-picker";

export async function compareTwoFiles(extensionUri: vscode.Uri): Promise<void> {
    const isChinese = isChineseDisplayLanguage();
    const leftUri = await pickWorkbook(
        isChinese ? "选择左侧 XLSX 工作簿" : "Select the left XLSX workbook"
    );
    if (!leftUri) {
        return;
    }

    const rightUri = await pickWorkbook(
        isChinese ? "选择右侧 XLSX 工作簿" : "Select the right XLSX workbook",
        leftUri
    );
    if (!rightUri) {
        return;
    }

    await XlsxDiffPanel.create(extensionUri, leftUri, rightUri);
}
