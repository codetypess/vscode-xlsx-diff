import { glob } from "node:fs/promises";
import { context, type BuildResult, type Plugin, type PluginBuild } from "esbuild";

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

const esbuildProblemMatcherPlugin: Plugin = {
    name: "esbuild-problem-matcher",

    setup(build: PluginBuild) {
        build.onStart(() => {
            console.log("[watch] build started");
        });
        build.onEnd((result: BuildResult) => {
            for (const { text, location } of result.errors) {
                console.error(`✘ [ERROR] ${text}`);
                if (location) {
                    console.error(`    ${location.file}:${location.line}:${location.column}:`);
                }
            }
            console.log("[watch] build finished");
        });
    },
};

async function collectTestEntryPoints(): Promise<string[]> {
    const entryPoints: string[] = [];

    for await (const filePath of glob("src/test/**/*.test.ts")) {
        entryPoints.push(filePath);
    }

    for await (const filePath of glob("src/test/**/*.test.tsx")) {
        entryPoints.push(filePath);
    }

    return entryPoints.sort((left, right) => left.localeCompare(right));
}

async function main() {
    const entryPoints = await collectTestEntryPoints();
    const testsCtx = await context({
        entryPoints,
        outbase: "src/test",
        outdir: "out/test",
        bundle: true,
        format: "cjs",
        outExtension: {
            ".js": ".cjs",
        },
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "node",
        external: ["vscode", "jsdom"],
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    if (watch) {
        await testsCtx.watch();
        return;
    }

    await testsCtx.rebuild();
    await testsCtx.dispose();
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
