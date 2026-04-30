function formatPerfLogValue(value: unknown): string {
    if (value === undefined) {
        return "undefined";
    }

    if (
        value === null ||
        typeof value === "number" ||
        typeof value === "boolean" ||
        typeof value === "bigint"
    ) {
        return String(value);
    }

    if (typeof value === "string") {
        return JSON.stringify(value);
    }

    try {
        return JSON.stringify(value) ?? String(value);
    } catch {
        return String(value);
    }
}

function getPerfNowMs(): number | null {
    const perf = globalThis.performance;
    if (typeof perf?.now !== "function") {
        return null;
    }

    return Number(perf.now().toFixed(2));
}

export function formatPerfLog(
    scope: "host" | "provider" | "webview",
    event: string,
    details: Readonly<Record<string, unknown>> = {}
): string {
    const now = new Date();
    const serializedDetails = Object.entries({
        time: now.toISOString(),
        epochMs: now.getTime(),
        perfNowMs: getPerfNowMs(),
        ...details,
    }).map(([key, value]) => `${key}=${formatPerfLogValue(value)}`);

    return `[xlsx-editor][${scope}] ${event} ${serializedDetails.join(" ")}`;
}
