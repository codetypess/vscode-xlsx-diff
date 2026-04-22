#!/usr/bin/env node

import { launchCompare } from "./lib/open-compare-uri.mjs";

function printUsage() {
    process.stderr.write(
        [
            "Usage:",
            "  xlsx-difftool <left.xlsx> <right.xlsx> [merged.xlsx]",
            "  xlsx-difftool --left <left.xlsx> --right <right.xlsx> [--extension-id <id>]",
            "",
            "This helper opens the VS Code XLSX Diff extension via a vscode:// URI.",
            "It is intended for use from Git difftool and other external compare tools.",
            "The optional merged.xlsx argument is accepted for Git-style integrations and ignored.",
            "",
        ].join("\n")
    );
}

function parseArgs(argv) {
    const positionals = [];
    let extensionId;

    for (let index = 0; index < argv.length; index += 1) {
        const value = argv[index];

        if (value === "--help" || value === "-h") {
            return { help: true };
        }

        if (value === "--extension-id") {
            extensionId = argv[index + 1];
            index += 1;
            continue;
        }

        if (value === "--left") {
            positionals[0] = argv[index + 1];
            index += 1;
            continue;
        }

        if (value === "--right") {
            positionals[1] = argv[index + 1];
            index += 1;
            continue;
        }

        positionals.push(value);
    }

    return {
        extensionId,
        leftPath: positionals[0],
        rightPath: positionals[1],
    };
}

const { help, extensionId, leftPath, rightPath } = parseArgs(process.argv.slice(2));

if (help) {
    printUsage();
    process.exit(0);
}

if (!leftPath || !rightPath) {
    printUsage();
    process.exit(1);
}

launchCompare(leftPath, rightPath, { extensionId });
