import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import { XlsxDiffPanel } from "../webview/diff-panel";

function toErrorMessage(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
}

async function ensureWorkbookFile(filePath: string, label: string): Promise<vscode.Uri> {
    const normalizedPath = path.resolve(filePath);

    if (path.extname(normalizedPath).toLowerCase() !== ".xlsx") {
        throw new Error(`${label} must point to an .xlsx file.`);
    }

    await fs.access(normalizedPath);
    return vscode.Uri.file(normalizedPath);
}

export class XlsxDiffUriHandler implements vscode.UriHandler {
    public constructor(private readonly extensionUri: vscode.Uri) {}

    public async handleUri(uri: vscode.Uri): Promise<void> {
        try {
            if (uri.path !== "/compare") {
                throw new Error(`Unsupported XLSX diff URI path: ${uri.path || "/"}`);
            }

            const params = new URLSearchParams(uri.query);
            const leftPath = params.get("left") ?? params.get("local");
            const rightPath = params.get("right") ?? params.get("remote");

            if (!leftPath || !rightPath) {
                throw new Error("The compare URI must include both left and right workbook paths.");
            }

            const [leftUri, rightUri] = await Promise.all([
                ensureWorkbookFile(leftPath, "Left workbook"),
                ensureWorkbookFile(rightPath, "Right workbook"),
            ]);

            await XlsxDiffPanel.create(this.extensionUri, leftUri, rightUri);
        } catch (error) {
            const errorMessage = toErrorMessage(error);
            console.error(error);
            await vscode.window.showErrorMessage(errorMessage);
        }
    }
}
