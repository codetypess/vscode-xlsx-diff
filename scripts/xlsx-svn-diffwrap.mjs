#!/usr/bin/env node

import { launchCompare } from "./lib/openCompareUri.mjs";

function printUsage() {
    process.stderr.write(
        [
            "Usage:",
            "  xlsx-svn-diffwrap [--extension-id <id>] <ignored> <ignored> <ignored> <ignored> <left.xlsx> <right.xlsx>",
            "",
            "This helper adapts Subversion's external diff arguments to xlsx-difftool.",
            "It opens the VS Code XLSX Diff extension via a vscode:// URI and exits with code 1,",
            "which matches the conventional diff exit code for 'differences found'.",
            "",
        ].join("\n")
    );
}

function parseArgs(argv) {
    let extensionId;
    const rawArgs = [];

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

        rawArgs.push(value);
    }

    return { extensionId, rawArgs };
}

const { help, extensionId, rawArgs = [] } = parseArgs(process.argv.slice(2));

if (help) {
    printUsage();
    process.exit(0);
}

if (rawArgs.length < 2) {
    printUsage();
    process.exit(2);
}

const leftPath = rawArgs[rawArgs.length - 2];
const rightPath = rawArgs[rawArgs.length - 1];
launchCompare(leftPath, rightPath, { extensionId });
process.exit(1);
