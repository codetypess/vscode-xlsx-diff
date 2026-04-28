import * as vscode from "vscode";
import { affectsCompareFormula } from "./compare-formula";
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
import { XlsxCustomEditorProvider } from "./webview/editor-panel";

export function activate(context: vscode.ExtensionContext) {
    context.subscriptions.push(
        XlsxCustomEditorProvider.register(context),
        vscode.window.registerUriHandler(new XlsxDiffUriHandler(context.extensionUri)),
        registerScmWorkbookDiffInterceptor(context.extensionUri),
        vscode.workspace.onDidChangeConfiguration((event) => {
            const displayLanguageChanged = affectsDisplayLanguage(event);
            const compareFormulaChanged = affectsCompareFormula(event);

            if (!displayLanguageChanged && !compareFormulaChanged) {
                return;
            }

            const refreshes = [
                import("./webview/diff-panel").then(({ XlsxDiffPanel }) => XlsxDiffPanel.refreshAll()),
            ];

            if (displayLanguageChanged) {
                refreshes.push(
                    import("./webview/editor-panel").then(({ XlsxEditorPanel }) =>
                        XlsxEditorPanel.refreshAll()
                    )
                );
            }

            void Promise.all(refreshes);
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
