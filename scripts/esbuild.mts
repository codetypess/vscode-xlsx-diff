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

    const webviewCtx = await context({
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

    const editorWebviewCtx = await context({
        entryPoints: ["src/webview/editorPanel.ts"],
        bundle: true,
        format: "iife",
        minify: production,
        sourcemap: !production,
        sourcesContent: false,
        platform: "browser",
        outfile: "media/editorPanel.js",
        logLevel: "silent",
        plugins: [esbuildProblemMatcherPlugin],
    });

    if (watch) {
        await extensionCtx.watch();
        await webviewCtx.watch();
        await editorWebviewCtx.watch();
        return;
    }

    await extensionCtx.rebuild();
    await extensionCtx.dispose();
    await webviewCtx.rebuild();
    await webviewCtx.dispose();
    await editorWebviewCtx.rebuild();
    await editorWebviewCtx.dispose();
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
