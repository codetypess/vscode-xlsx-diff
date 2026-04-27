import { access, mkdir } from "node:fs/promises";
import { constants as fsConstants } from "node:fs";
import { execFile } from "node:child_process";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.resolve(scriptDirectory, "..");
const fixtureRoot = path.join(repositoryRoot, "src", "test", "fixtures", "xlsx-regressions");
const lfValue = "$&key1=ARMY==#army.id\n$&key1=ASSET==#assets.id";
const crlfValue = "$&key1=ARMY==#army.id\r\n$&key1=ASSET==#assets.id";

interface FastxlsxCommand {
    command: string;
    argsPrefix: string[];
}

interface FixtureCellEdit {
    cellAddress: string;
    value: string;
}

interface FixtureCase {
    name: string;
    sheetName: string;
    baseEdits: FixtureCellEdit[];
    headEdits: FixtureCellEdit[];
}

const fixtureCases: FixtureCase[] = [
    {
        name: "newline-only-cell-diff",
        sheetName: "define",
        baseEdits: [
            {
                cellAddress: "F5",
                value: lfValue,
            },
        ],
        headEdits: [
            {
                cellAddress: "F5",
                value: crlfValue,
            },
        ],
    },
    {
        name: "empty-string-vs-blank-cell",
        sheetName: "define",
        baseEdits: [],
        headEdits: [
            {
                cellAddress: "F5",
                value: "",
            },
        ],
    },
];

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
    sheetName: string,
    cellEdits: FixtureCellEdit[]
): Promise<void> {
    await runFastxlsx(fastxlsxCommand, ["create", workbookPath, "--sheet", sheetName]);

    for (const cellEdit of cellEdits) {
        await runFastxlsx(
            fastxlsxCommand,
            [
                "set",
                workbookPath,
                "--sheet",
                sheetName,
                "--cell",
                cellEdit.cellAddress,
                "--text",
                cellEdit.value,
                "--in-place",
            ],
            { quiet: true }
        );
    }

    await runFastxlsx(fastxlsxCommand, ["validate", workbookPath]);

    for (const cellEdit of cellEdits) {
        await runFastxlsx(
            fastxlsxCommand,
            ["get", workbookPath, "--sheet", sheetName, "--cell", cellEdit.cellAddress],
            { quiet: true }
        );
    }
}

async function main(): Promise<void> {
    const fastxlsxCommand = await resolveFastxlsxCommand();
    await mkdir(fixtureRoot, { recursive: true });

    for (const fixtureCase of fixtureCases) {
        const caseDirectory = path.join(fixtureRoot, fixtureCase.name);
        await mkdir(caseDirectory, { recursive: true });
        await createFixtureWorkbook(
            fastxlsxCommand,
            path.join(caseDirectory, "base.xlsx"),
            fixtureCase.sheetName,
            fixtureCase.baseEdits
        );
        await createFixtureWorkbook(
            fastxlsxCommand,
            path.join(caseDirectory, "head.xlsx"),
            fixtureCase.sheetName,
            fixtureCase.headEdits
        );

        console.log(`Fixtures generated in ${caseDirectory}`);
    }
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
