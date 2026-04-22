import * as path from "node:path";
import { spawn } from "node:child_process";

function resolveExtensionId(extensionId) {
    return extensionId ?? process.env.VSCODE_XLSX_DIFF_EXTENSION_ID ?? "codetypess.xlsx-diff";
}

export function buildCompareUri(leftPath, rightPath, extensionId) {
    const compareUri = new URL(`vscode://${resolveExtensionId(extensionId)}/compare`);
    compareUri.searchParams.set("left", path.resolve(leftPath));
    compareUri.searchParams.set("right", path.resolve(rightPath));
    return compareUri.toString();
}

export function openUri(uri) {
    if (process.platform === "darwin") {
        return spawn("open", [uri], {
            detached: true,
            stdio: "ignore",
        });
    }

    if (process.platform === "win32") {
        return spawn("cmd", ["/c", "start", "", uri], {
            detached: true,
            stdio: "ignore",
            windowsHide: true,
        });
    }

    return spawn("xdg-open", [uri], {
        detached: true,
        stdio: "ignore",
    });
}

export function launchCompare(leftPath, rightPath, options = {}) {
    const uri = buildCompareUri(leftPath, rightPath, options.extensionId);
    const child = openUri(uri);
    child.unref();
}
