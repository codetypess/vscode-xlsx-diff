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
    const extensionCtx = await context({
        entryPoints: ["src/extension.ts"],
        bundle: true,
        format: "cjs",
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "node",
        outfile: "dist/extension.js",
        external: ["vscode"],
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    const editorWebviewCtx = await context({
        entryPoints: ["src/webview/editor-webview.ts"],
        bundle: true,
        format: "iife",
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "browser",
        outfile: "media/editor-panel.js",
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    const diffPanelWebviewCtx = await context({
        entryPoints: ["src/webview/panel.ts"],
        bundle: true,
        format: "iife",
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "browser",
        outfile: "media/panel.js",
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    if (watch) {
        await extensionCtx.watch();
        await editorWebviewCtx.watch();
        await diffPanelWebviewCtx.watch();
        return;
    }

    await extensionCtx.rebuild();
    await extensionCtx.dispose();
    await editorWebviewCtx.rebuild();
    await editorWebviewCtx.dispose();
    await diffPanelWebviewCtx.rebuild();
    await diffPanelWebviewCtx.dispose();
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
