import { getFontShorthand, measureMaximumDigitWidth } from "./column-layout";

export interface VsCodeApi<TMessage> {
    postMessage(message: TMessage): void;
}

export function getVsCodeApi<TMessage>(): VsCodeApi<TMessage> {
    const candidate = (globalThis as Record<string, unknown>).acquireVsCodeApi;
    if (typeof candidate === "function") {
        return (candidate as () => VsCodeApi<TMessage>)();
    }

    return {
        postMessage: () => undefined,
    };
}

export function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}

export function isTextInputTarget(target: EventTarget | null): boolean {
    if (!(target instanceof HTMLElement)) {
        return false;
    }

    return Boolean(
        target instanceof HTMLInputElement ||
            target instanceof HTMLTextAreaElement ||
            target.isContentEditable ||
            target.closest('input, textarea, [contenteditable="true"], [contenteditable=""]')
    );
}

export function areMeasuredItemWidthsEqual(
    left: Record<string, number>,
    right: Record<string, number>
): boolean {
    const leftKeys = Object.keys(left);
    const rightKeys = Object.keys(right);
    if (leftKeys.length !== rightKeys.length) {
        return false;
    }

    return leftKeys.every((key) => left[key] === right[key]);
}

export function measureElementPixelWidth(element: Element | null | undefined): number {
    return Math.ceil(element?.getBoundingClientRect().width ?? 0);
}

export function measureMaximumDigitWidthFromDocument(defaultWidth: number): number {
    if (globalThis.navigator?.userAgent?.includes("jsdom")) {
        return defaultWidth;
    }

    try {
        return measureMaximumDigitWidth(getFontShorthand(document.body));
    } catch {
        return defaultWidth;
    }
}

export function measureSheetTabWidths(
    container: ParentNode | null | undefined,
    {
        itemSelector,
        maxWidth,
    }: {
        itemSelector: string;
        maxWidth: number;
    }
): Record<string, number> {
    if (!container) {
        return {};
    }

    const nextMeasuredTabWidths: Record<string, number> = {};
    for (const element of container.querySelectorAll<HTMLElement>(itemSelector)) {
        const sheetKey = element.dataset.sheetKey;
        if (!sheetKey) {
            continue;
        }

        nextMeasuredTabWidths[sheetKey] = Math.min(
            maxWidth,
            measureElementPixelWidth(element)
        );
    }

    return nextMeasuredTabWidths;
}
