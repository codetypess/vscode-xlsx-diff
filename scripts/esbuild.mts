import { rm } from "node:fs/promises";
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

async function main() {
    await rm("dist", { recursive: true, force: true });

    const extensionCtx = await context({
        entryPoints: ["src/extension.ts"],
        outbase: "src",
        bundle: true,
        format: "esm",
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "node",
        splitting: true,
        outdir: "dist",
        entryNames: "[name]",
        chunkNames: "chunks/[name]-[hash]",
        external: ["vscode"],
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    if (watch) {
        await extensionCtx.watch();
        return;
    }

    await extensionCtx.rebuild();
    await extensionCtx.dispose();
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
