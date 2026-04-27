export const DEFAULT_COLUMN_PIXEL_WIDTH = 120;
export const MIN_COLUMN_PIXEL_WIDTH = 40;
export const MAX_COLUMN_PIXEL_WIDTH = 1024;
export const DEFAULT_MAXIMUM_DIGIT_WIDTH_PX = 7;

const DIGITS = "0123456789";

interface FontStyleSnapshot {
    font: string;
    fontWeight: string;
    fontSize: string;
    fontFamily: string;
}

interface MeasureContext {
    font: string;
    measureText(text: string): { width: number };
}

interface ElementFactory {
    getContext(type: "2d"): MeasureContext | null;
}

interface DomGlobals {
    window?: {
        getComputedStyle(target: object): FontStyleSnapshot;
    };
    document?: {
        body?: object | null;
        documentElement?: object | null;
        createElement(tagName: "canvas"): ElementFactory;
    };
}

export interface PixelColumnLayout {
    pixelWidths: number[];
    prefixOffsets: number[];
    actualColumnCount: number;
    totalColumnCount: number;
    fallbackPixelWidth: number;
    actualColumnsWidth: number;
    totalWidth: number;
}

export function convertWorkbookColumnWidthToPixels(
    workbookWidth: number,
    maximumDigitWidth: number
): number {
    const mdw = Math.max(1, Math.floor(maximumDigitWidth));
    return Math.trunc(((256 * workbookWidth + Math.trunc(128 / mdw)) / 256) * mdw);
}

export function convertPixelsToWorkbookColumnWidth(
    pixels: number,
    maximumDigitWidth: number
): number {
    const mdw = Math.max(1, Math.floor(maximumDigitWidth));
    const normalizedPixels = Math.max(1, Math.round(pixels));
    return Math.round((normalizedPixels / mdw - Math.trunc(128 / mdw) / 256) * 256) / 256;
}

export function resolveColumnPixelWidth(
    workbookWidth: number | null | undefined,
    {
        maximumDigitWidth = DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
        fallbackPixelWidth = DEFAULT_COLUMN_PIXEL_WIDTH,
        minPixelWidth = MIN_COLUMN_PIXEL_WIDTH,
        maxPixelWidth = MAX_COLUMN_PIXEL_WIDTH,
    }: {
        maximumDigitWidth?: number;
        fallbackPixelWidth?: number;
        minPixelWidth?: number;
        maxPixelWidth?: number;
    } = {}
): number {
    const resolvedWidth =
        workbookWidth === null || workbookWidth === undefined
            ? fallbackPixelWidth
            : convertWorkbookColumnWidthToPixels(workbookWidth, maximumDigitWidth);
    return Math.max(minPixelWidth, Math.min(maxPixelWidth, Math.round(resolvedWidth)));
}

export function createPixelColumnLayout({
    columnWidths,
    totalColumnCount = columnWidths.length,
    maximumDigitWidth = DEFAULT_MAXIMUM_DIGIT_WIDTH_PX,
    fallbackPixelWidth = DEFAULT_COLUMN_PIXEL_WIDTH,
    minPixelWidth = MIN_COLUMN_PIXEL_WIDTH,
    maxPixelWidth = MAX_COLUMN_PIXEL_WIDTH,
}: {
    columnWidths: readonly (number | null | undefined)[];
    totalColumnCount?: number;
    maximumDigitWidth?: number;
    fallbackPixelWidth?: number;
    minPixelWidth?: number;
    maxPixelWidth?: number;
}): PixelColumnLayout {
    const pixelWidths = columnWidths.map((columnWidth) =>
        resolveColumnPixelWidth(columnWidth, {
            maximumDigitWidth,
            fallbackPixelWidth,
            minPixelWidth,
            maxPixelWidth,
        })
    );
    const prefixOffsets = [0];
    for (const pixelWidth of pixelWidths) {
        prefixOffsets.push(prefixOffsets[prefixOffsets.length - 1]! + pixelWidth);
    }

    const actualColumnCount = pixelWidths.length;
    const actualColumnsWidth = prefixOffsets[actualColumnCount] ?? 0;
    const normalizedTotalColumnCount = Math.max(totalColumnCount, actualColumnCount);
    const totalWidth =
        actualColumnsWidth +
        Math.max(0, normalizedTotalColumnCount - actualColumnCount) * fallbackPixelWidth;

    return {
        pixelWidths,
        prefixOffsets,
        actualColumnCount,
        totalColumnCount: normalizedTotalColumnCount,
        fallbackPixelWidth,
        actualColumnsWidth,
        totalWidth,
    };
}

export function extendPixelColumnLayout(
    layout: PixelColumnLayout,
    totalColumnCount: number
): PixelColumnLayout {
    const normalizedTotalColumnCount = Math.max(totalColumnCount, layout.actualColumnCount);
    if (normalizedTotalColumnCount === layout.totalColumnCount) {
        return layout;
    }

    return {
        ...layout,
        totalColumnCount: normalizedTotalColumnCount,
        totalWidth:
            layout.actualColumnsWidth +
            Math.max(0, normalizedTotalColumnCount - layout.actualColumnCount) *
                layout.fallbackPixelWidth,
    };
}

export function getPixelColumnOffset(layout: PixelColumnLayout, index: number): number {
    const normalizedIndex = Math.max(0, Math.min(index, layout.totalColumnCount));
    if (normalizedIndex <= layout.actualColumnCount) {
        return layout.prefixOffsets[normalizedIndex] ?? layout.actualColumnsWidth;
    }

    return (
        layout.actualColumnsWidth +
        (normalizedIndex - layout.actualColumnCount) * layout.fallbackPixelWidth
    );
}

export function getPixelColumnLeft(layout: PixelColumnLayout, columnNumber: number): number {
    return getPixelColumnOffset(layout, columnNumber - 1);
}

export function getPixelColumnWidth(layout: PixelColumnLayout, columnNumber: number): number {
    if (columnNumber < 1 || columnNumber > layout.totalColumnCount) {
        return layout.fallbackPixelWidth;
    }

    return columnNumber <= layout.actualColumnCount
        ? (layout.pixelWidths[columnNumber - 1] ?? layout.fallbackPixelWidth)
        : layout.fallbackPixelWidth;
}

export function getPixelColumnRight(layout: PixelColumnLayout, columnNumber: number): number {
    return getPixelColumnLeft(layout, columnNumber) + getPixelColumnWidth(layout, columnNumber);
}

export function findColumnIndexForOffset(layout: PixelColumnLayout, offset: number): number {
    if (layout.totalColumnCount <= 0) {
        return 0;
    }

    const normalizedOffset = Math.max(0, Math.min(offset, Math.max(0, layout.totalWidth - 1)));
    if (normalizedOffset >= layout.actualColumnsWidth) {
        return Math.min(
            layout.totalColumnCount - 1,
            layout.actualColumnCount +
                Math.floor((normalizedOffset - layout.actualColumnsWidth) / layout.fallbackPixelWidth)
        );
    }

    let low = 0;
    let high = layout.actualColumnCount;
    while (low < high) {
        const mid = Math.floor((low + high + 1) / 2);
        if ((layout.prefixOffsets[mid] ?? layout.actualColumnsWidth) <= normalizedOffset) {
            low = mid;
        } else {
            high = mid - 1;
        }
    }

    return Math.min(layout.totalColumnCount - 1, low);
}

export function getPixelColumnWindow(
    layout: PixelColumnLayout,
    scrollLeft: number,
    viewportWidth: number,
    overscan: number
): {
    startIndex: number;
    endIndex: number;
    leadingSpacerWidth: number;
    trailingSpacerWidth: number;
} {
    if (layout.totalColumnCount <= 0) {
        return {
            startIndex: 0,
            endIndex: 0,
            leadingSpacerWidth: 0,
            trailingSpacerWidth: 0,
        };
    }

    const normalizedScrollLeft = Math.max(0, scrollLeft);
    const effectiveViewportWidth = Math.max(
        viewportWidth,
        getPixelColumnWidth(layout, Math.min(layout.totalColumnCount, 1))
    );
    const startVisibleIndex = findColumnIndexForOffset(layout, normalizedScrollLeft);
    const endVisibleIndex = findColumnIndexForOffset(
        layout,
        Math.max(normalizedScrollLeft, normalizedScrollLeft + effectiveViewportWidth - 1)
    );
    const startIndex = Math.max(0, startVisibleIndex - overscan);
    const endIndex = Math.min(layout.totalColumnCount, endVisibleIndex + overscan + 1);

    return {
        startIndex,
        endIndex,
        leadingSpacerWidth: getPixelColumnOffset(layout, startIndex),
        trailingSpacerWidth: Math.max(0, layout.totalWidth - getPixelColumnOffset(layout, endIndex)),
    };
}

export function getFontShorthand(target: object | null): string | null {
    const dom = globalThis as DomGlobals;
    if (!dom.window) {
        return null;
    }

    const element = target ?? dom.document?.body ?? dom.document?.documentElement ?? null;
    if (!element) {
        return null;
    }

    const style = dom.window.getComputedStyle(element);
    return style.font || `${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
}

let measureContext: MeasureContext | null | undefined;

function getMeasureContext(): MeasureContext | null {
    if (measureContext !== undefined) {
        return measureContext;
    }

    const dom = globalThis as DomGlobals;
    if (!dom.document) {
        measureContext = null;
        return measureContext;
    }

    measureContext = dom.document.createElement("canvas").getContext("2d");
    return measureContext;
}

export function measureMaximumDigitWidth(font: string | null | undefined): number {
    const context = getMeasureContext();
    if (!context || !font) {
        return DEFAULT_MAXIMUM_DIGIT_WIDTH_PX;
    }

    context.font = font;
    let maximumWidth = 0;
    for (const digit of DIGITS) {
        maximumWidth = Math.max(maximumWidth, context.measureText(digit).width);
    }

    return Math.max(1, Math.ceil(maximumWidth)) || DEFAULT_MAXIMUM_DIGIT_WIDTH_PX;
}
