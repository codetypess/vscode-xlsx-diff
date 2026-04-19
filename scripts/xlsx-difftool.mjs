#!/usr/bin/env node

import * as path from "node:path";
import { spawn } from "node:child_process";

function printUsage() {
    process.stderr.write(
        [
            "Usage: xlsx-difftool <local.xlsx> <remote.xlsx> [merged.xlsx]",
            "",
            "This helper opens the VS Code XLSX Diff extension via a vscode:// URI.",
            "It is intended for use from git difftool with $LOCAL and $REMOTE.",
            "",
        ].join("\n")
    );
}

function openUri(uri) {
    const platform = process.platform;

    if (platform === "darwin") {
        return spawn("open", [uri], {
            detached: true,
            stdio: "ignore",
        });
    }

    if (platform === "win32") {
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

const [, , localPath, remotePath] = process.argv;

if (!localPath || !remotePath) {
    printUsage();
    process.exit(1);
}

const extensionId = process.env.VSCODE_XLSX_DIFF_EXTENSION_ID ?? "codetypess.xlsx-diff";
const compareUri = new URL(`vscode://${extensionId}/compare`);

compareUri.searchParams.set("left", path.resolve(localPath));
compareUri.searchParams.set("right", path.resolve(remotePath));

const child = openUri(compareUri.toString());
child.unref();
