import { rm } from "node:fs/promises";
import { resolve } from "node:path";
import { build, type InlineConfig, type PluginOption } from "vite";
import solidPlugin from "vite-plugin-solid";

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

function watchProblemMatcherPlugin(): PluginOption {
    return {
        name: "vite-webview-problem-matcher",
        buildStart() {
            console.log("[watch] build started");
        },
        buildEnd(error) {
            if (error) {
                console.error(`✘ [ERROR] ${error.message}`);
            }
            console.log("[watch] build finished");
        },
    };
}

function createWebviewBuildConfig({
    entryFile,
    outputFile,
    globalName,
}: {
    entryFile: string;
    outputFile: string;
    globalName: string;
}): InlineConfig {
    return {
        configFile: false,
        publicDir: false,
        logLevel: "info",
        plugins: [
            solidPlugin({
                dev: !production,
            }),
            watchProblemMatcherPlugin(),
        ],
        define: {
            "process.env.NODE_ENV": JSON.stringify(production ? "production" : "development"),
        },
        build: {
            target: "es2022",
            outDir: resolve(process.cwd(), "media"),
            emptyOutDir: false,
            minify: production,
            sourcemap: !production,
            watch: watch ? {} : null,
            lib: {
                entry: resolve(process.cwd(), entryFile),
                formats: ["iife"],
                name: globalName,
                fileName: () => outputFile,
            },
            rollupOptions: {
                output: {
                    inlineDynamicImports: true,
                },
            },
        },
    };
}

async function main() {
    if (!watch) {
        await Promise.all([
            rm(resolve(process.cwd(), "media", "editor-panel.js"), { force: true }),
            rm(resolve(process.cwd(), "media", "editor-panel.js.map"), { force: true }),
            rm(resolve(process.cwd(), "media", "diff-panel.js"), { force: true }),
            rm(resolve(process.cwd(), "media", "diff-panel.js.map"), { force: true }),
        ]);
    }

    await build(
        createWebviewBuildConfig({
            entryFile: "src/webview-solid/editor-panel/main.tsx",
            outputFile: "editor-panel.js",
            globalName: "XlsxEditorPanelBootstrap",
        })
    );

    await build(
        createWebviewBuildConfig({
            entryFile: "src/webview-solid/diff-panel/main.tsx",
            outputFile: "diff-panel.js",
            globalName: "XlsxDiffPanelBootstrap",
        })
    );
}

main().catch((error) => {
    if (error instanceof Error) {
        console.error(`✘ [ERROR] ${error.message}`);
    } else {
        console.error("✘ [ERROR] Unknown Vite build failure");
    }
    process.exit(1);
});
