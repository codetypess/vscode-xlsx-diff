export const DEFAULT_WORKBOOK_ROW_HEIGHT = 15;
export const DEFAULT_ROW_PIXEL_HEIGHT = 28;
export const MIN_ROW_PIXEL_HEIGHT = 14;
export const MAX_ROW_PIXEL_HEIGHT = 512;

export interface PixelRowLayout {
    overriddenPixelHeights: Record<string, number>;
    overriddenRowNumbers: number[];
    prefixAdjustments: number[];
    totalRowCount: number;
    fallbackPixelHeight: number;
    totalHeight: number;
}

export function convertWorkbookRowHeightToPixels(workbookHeight: number): number {
    return Math.round(
        (workbookHeight / DEFAULT_WORKBOOK_ROW_HEIGHT) * DEFAULT_ROW_PIXEL_HEIGHT
    );
}

export function convertPixelsToWorkbookRowHeight(pixels: number): number {
    const normalizedPixels = Math.max(1, Math.round(pixels));
    return (
        Math.round(
            (normalizedPixels / DEFAULT_ROW_PIXEL_HEIGHT) *
                DEFAULT_WORKBOOK_ROW_HEIGHT *
                100
        ) / 100
    );
}

export function stabilizeRowPixelHeight(pixels: number): number {
    return convertWorkbookRowHeightToPixels(convertPixelsToWorkbookRowHeight(pixels));
}

export function resolveRowPixelHeight(
    workbookHeight: number | null | undefined,
    {
        fallbackPixelHeight = DEFAULT_ROW_PIXEL_HEIGHT,
        minPixelHeight = MIN_ROW_PIXEL_HEIGHT,
        maxPixelHeight = MAX_ROW_PIXEL_HEIGHT,
    }: {
        fallbackPixelHeight?: number;
        minPixelHeight?: number;
        maxPixelHeight?: number;
    } = {}
): number {
    const resolvedHeight =
        workbookHeight === null || workbookHeight === undefined
            ? fallbackPixelHeight
            : convertWorkbookRowHeightToPixels(workbookHeight);
    return Math.max(minPixelHeight, Math.min(maxPixelHeight, Math.round(resolvedHeight)));
}

export function createPixelRowLayout({
    rowHeights,
    totalRowCount,
    fallbackPixelHeight = DEFAULT_ROW_PIXEL_HEIGHT,
    minPixelHeight = MIN_ROW_PIXEL_HEIGHT,
    maxPixelHeight = MAX_ROW_PIXEL_HEIGHT,
}: {
    rowHeights?: Readonly<Record<string, number | null>>;
    totalRowCount: number;
    fallbackPixelHeight?: number;
    minPixelHeight?: number;
    maxPixelHeight?: number;
}): PixelRowLayout {
    const normalizedEntries = Object.entries(rowHeights ?? {})
        .map(([rowNumberText, rowHeight]) => ({
            rowNumber: Number(rowNumberText),
            pixelHeight: resolveRowPixelHeight(rowHeight, {
                fallbackPixelHeight,
                minPixelHeight,
                maxPixelHeight,
            }),
        }))
        .filter(
            ({ rowNumber }) =>
                Number.isInteger(rowNumber) && rowNumber >= 1 && rowNumber <= totalRowCount
        )
        .sort((left, right) => left.rowNumber - right.rowNumber);
    const overriddenPixelHeights = Object.fromEntries(
        normalizedEntries.map(({ rowNumber, pixelHeight }) => [String(rowNumber), pixelHeight])
    );
    const overriddenRowNumbers = normalizedEntries.map(({ rowNumber }) => rowNumber);
    const prefixAdjustments: number[] = [];
    let cumulativeAdjustment = 0;
    for (const { pixelHeight } of normalizedEntries) {
        cumulativeAdjustment += pixelHeight - fallbackPixelHeight;
        prefixAdjustments.push(cumulativeAdjustment);
    }

    return {
        overriddenPixelHeights,
        overriddenRowNumbers,
        prefixAdjustments,
        totalRowCount,
        fallbackPixelHeight,
        totalHeight: totalRowCount * fallbackPixelHeight + cumulativeAdjustment,
    };
}

export function extendPixelRowLayout(
    layout: PixelRowLayout,
    totalRowCount: number
): PixelRowLayout {
    const normalizedTotalRowCount = Math.max(totalRowCount, layout.totalRowCount);
    if (normalizedTotalRowCount === layout.totalRowCount) {
        return layout;
    }

    return {
        ...layout,
        totalRowCount: normalizedTotalRowCount,
        totalHeight:
            normalizedTotalRowCount * layout.fallbackPixelHeight +
            (layout.prefixAdjustments[layout.prefixAdjustments.length - 1] ?? 0),
    };
}

function findLastOverriddenRowIndex(
    layout: PixelRowLayout,
    rowNumber: number
): number {
    let low = 0;
    let high = layout.overriddenRowNumbers.length - 1;
    let result = -1;

    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        const candidate = layout.overriddenRowNumbers[mid]!;
        if (candidate <= rowNumber) {
            result = mid;
            low = mid + 1;
        } else {
            high = mid - 1;
        }
    }

    return result;
}

export function getPixelRowOffset(layout: PixelRowLayout, index: number): number {
    const normalizedIndex = Math.max(0, Math.min(index, layout.totalRowCount));
    const baseHeight = normalizedIndex * layout.fallbackPixelHeight;
    const adjustmentIndex = findLastOverriddenRowIndex(layout, normalizedIndex);
    if (adjustmentIndex < 0) {
        return baseHeight;
    }

    return baseHeight + (layout.prefixAdjustments[adjustmentIndex] ?? 0);
}

export function getPixelRowTop(layout: PixelRowLayout, rowNumber: number): number {
    return getPixelRowOffset(layout, rowNumber - 1);
}

export function getPixelRowHeight(layout: PixelRowLayout, rowNumber: number): number {
    if (rowNumber < 1 || rowNumber > layout.totalRowCount) {
        return layout.fallbackPixelHeight;
    }

    return (
        layout.overriddenPixelHeights[String(rowNumber)] ?? layout.fallbackPixelHeight
    );
}

export function getPixelRowBottom(layout: PixelRowLayout, rowNumber: number): number {
    return getPixelRowTop(layout, rowNumber) + getPixelRowHeight(layout, rowNumber);
}

export function findRowIndexForOffset(layout: PixelRowLayout, offset: number): number {
    if (layout.totalRowCount <= 0) {
        return 0;
    }

    const normalizedOffset = Math.max(0, Math.min(offset, Math.max(0, layout.totalHeight - 1)));
    let low = 0;
    let high = layout.totalRowCount - 1;

    while (low < high) {
        const mid = Math.floor((low + high) / 2);
        if (getPixelRowOffset(layout, mid + 1) <= normalizedOffset) {
            low = mid + 1;
        } else {
            high = mid;
        }
    }

    return low;
}

export function getPixelRowWindow(
    layout: PixelRowLayout,
    scrollTop: number,
    viewportHeight: number,
    overscan: number
): {
    startIndex: number;
    endIndex: number;
    leadingSpacerHeight: number;
    trailingSpacerHeight: number;
} {
    if (layout.totalRowCount <= 0) {
        return {
            startIndex: 0,
            endIndex: 0,
            leadingSpacerHeight: 0,
            trailingSpacerHeight: 0,
        };
    }

    const normalizedScrollTop = Math.max(0, scrollTop);
    const effectiveViewportHeight = Math.max(
        viewportHeight,
        getPixelRowHeight(layout, Math.min(layout.totalRowCount, 1))
    );
    const startVisibleIndex = findRowIndexForOffset(layout, normalizedScrollTop);
    const endVisibleIndex = findRowIndexForOffset(
        layout,
        Math.max(normalizedScrollTop, normalizedScrollTop + effectiveViewportHeight - 1)
    );
    const startIndex = Math.max(0, startVisibleIndex - overscan);
    const endIndex = Math.min(layout.totalRowCount, endVisibleIndex + overscan + 1);

    return {
        startIndex,
        endIndex,
        leadingSpacerHeight: getPixelRowOffset(layout, startIndex),
        trailingSpacerHeight: Math.max(0, layout.totalHeight - getPixelRowOffset(layout, endIndex)),
    };
}
