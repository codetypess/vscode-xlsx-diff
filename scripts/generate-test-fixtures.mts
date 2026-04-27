import { access, mkdir } from "node:fs/promises";
import { constants as fsConstants } from "node:fs";
import { execFile } from "node:child_process";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.resolve(scriptDirectory, "..");
const fixtureRoot = path.join(
    repositoryRoot,
    "src",
    "test",
    "fixtures",
    "xlsx-regressions",
    "newline-only-cell-diff"
);
const sheetName = "define";
const cellAddress = "F5";
const lfValue = "$&key1=ARMY==#army.id\n$&key1=ASSET==#assets.id";
const crlfValue = "$&key1=ARMY==#army.id\r\n$&key1=ASSET==#assets.id";

interface FastxlsxCommand {
    command: string;
    argsPrefix: string[];
}

async function pathExists(targetPath: string): Promise<boolean> {
    try {
        await access(targetPath, fsConstants.F_OK);
        return true;
    } catch {
        return false;
    }
}

async function resolveFastxlsxCommand(): Promise<FastxlsxCommand> {
    const localCommand = path.join(
        repositoryRoot,
        "node_modules",
        ".bin",
        process.platform === "win32" ? "fastxlsx.cmd" : "fastxlsx"
    );

    if (await pathExists(localCommand)) {
        return {
            command: localCommand,
            argsPrefix: [],
        };
    }

    return {
        command: process.platform === "win32" ? "npx.cmd" : "npx",
        argsPrefix: ["fastxlsx"],
    };
}

async function runFastxlsx(
    fastxlsxCommand: FastxlsxCommand,
    args: string[],
    options: {
        quiet?: boolean;
    } = {}
): Promise<void> {
    const commandArgs = [...fastxlsxCommand.argsPrefix, ...args];
    const renderedCommand = [fastxlsxCommand.command, ...commandArgs]
        .map((value) => JSON.stringify(value))
        .join(" ");

    if (!options.quiet) {
        console.log(`$ ${renderedCommand}`);
    }

    const { stdout, stderr } = await execFileAsync(fastxlsxCommand.command, commandArgs, {
        cwd: repositoryRoot,
    });

    if (stdout.trim()) {
        process.stdout.write(stdout);
        if (!stdout.endsWith("\n")) {
            process.stdout.write("\n");
        }
    }

    if (stderr.trim()) {
        process.stderr.write(stderr);
        if (!stderr.endsWith("\n")) {
            process.stderr.write("\n");
        }
    }
}

async function createFixtureWorkbook(
    fastxlsxCommand: FastxlsxCommand,
    workbookPath: string,
    cellValue: string
): Promise<void> {
    await runFastxlsx(fastxlsxCommand, ["create", workbookPath, "--sheet", sheetName]);
    await runFastxlsx(
        fastxlsxCommand,
        [
            "set",
            workbookPath,
            "--sheet",
            sheetName,
            "--cell",
            cellAddress,
            "--text",
            cellValue,
            "--in-place",
        ],
        { quiet: true }
    );
    await runFastxlsx(fastxlsxCommand, ["validate", workbookPath]);
    await runFastxlsx(
        fastxlsxCommand,
        ["get", workbookPath, "--sheet", sheetName, "--cell", cellAddress],
        { quiet: true }
    );
}

async function main(): Promise<void> {
    const fastxlsxCommand = await resolveFastxlsxCommand();
    await mkdir(fixtureRoot, { recursive: true });

    await createFixtureWorkbook(fastxlsxCommand, path.join(fixtureRoot, "base.xlsx"), lfValue);
    await createFixtureWorkbook(fastxlsxCommand, path.join(fixtureRoot, "head.xlsx"), crlfValue);

    console.log(`Fixtures generated in ${fixtureRoot}`);
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
