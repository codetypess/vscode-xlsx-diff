import { existsSync } from "node:fs";
import * as path from "node:path";

function findRepositoryRoot(startDirectory: string): string {
    let currentDirectory = startDirectory;

    while (true) {
        if (existsSync(path.join(currentDirectory, "package.json"))) {
            return currentDirectory;
        }

        const parentDirectory = path.dirname(currentDirectory);
        if (parentDirectory === currentDirectory) {
            throw new Error(`Could not locate repository root from ${startDirectory}`);
        }

        currentDirectory = parentDirectory;
    }
}

const repositoryRoot = findRepositoryRoot(__dirname);

export function getTestFixturePath(...segments: string[]): string {
    return path.join(repositoryRoot, "src", "test", "fixtures", ...segments);
}
