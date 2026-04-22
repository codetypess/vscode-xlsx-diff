import * as vscode from "vscode";
import { compareActiveWith } from "./commands/compare-active-with";
import { compareTwoFiles } from "./commands/compare-two-files";
import { openEditor } from "./commands/open-editor";
import {
    COMMAND_COMPARE_ACTIVE_WITH,
    COMMAND_COMPARE_TWO_FILES,
    COMMAND_OPEN_EDITOR,
} from "./constants";
import { affectsDisplayLanguage } from "./display-language";
import { XlsxDiffUriHandler } from "./git/uri-handler";
import { registerScmWorkbookDiffInterceptor } from "./scm/scm-diff-interceptor";
import { XlsxDiffPanel } from "./webview/diff-panel";
import { XlsxCustomEditorProvider } from "./webview/xlsx-custom-editor-provider";
import { XlsxEditorPanel } from "./webview/xlsx-editor-panel";

export function activate(context: vscode.ExtensionContext) {
    context.subscriptions.push(
        XlsxCustomEditorProvider.register(context),
        vscode.window.registerUriHandler(new XlsxDiffUriHandler(context.extensionUri)),
        registerScmWorkbookDiffInterceptor(context.extensionUri),
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (!affectsDisplayLanguage(event)) {
                return;
            }

            void Promise.all([
                XlsxDiffPanel.refreshAll(),
                XlsxEditorPanel.refreshAll(),
            ]);
        }),
        vscode.commands.registerCommand(COMMAND_COMPARE_TWO_FILES, async () => {
            await compareTwoFiles(context.extensionUri);
        }),
        vscode.commands.registerCommand(COMMAND_OPEN_EDITOR, async (resource?: unknown) => {
            await openEditor(resource);
        }),
        vscode.commands.registerCommand(COMMAND_COMPARE_ACTIVE_WITH, async (resource?: unknown) => {
            await compareActiveWith(context.extensionUri, resource);
        })
    );
}

export function deactivate() {}
