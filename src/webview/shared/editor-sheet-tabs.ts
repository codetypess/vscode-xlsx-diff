export interface SheetTabLike {
    isActive: boolean;
}

export interface SheetTabOverflowLayout<TTab> {
    activeTab: TTab | null;
    activeTabVisibleIndex: number;
    hasOverflow: boolean;
    overflowTabs: TTab[];
    visibleTabs: TTab[];
}

export interface SheetTabWidthLayoutOptions<TTab> {
    containerWidth: number;
    getTabWidth(tab: TTab): number;
    itemGap?: number;
    overflowTriggerWidth: number;
}

function normalizeMaxVisibleTabs(maxVisibleTabs: number): number {
    if (!Number.isFinite(maxVisibleTabs)) {
        return 1;
    }

    return Math.max(1, Math.floor(maxVisibleTabs));
}

function getActiveTabIndex<TTab extends SheetTabLike>(tabs: readonly TTab[]): number {
    const activeIndex = tabs.findIndex((tab) => tab.isActive);
    return activeIndex >= 0 ? activeIndex : 0;
}

function normalizeWidth(width: number): number {
    if (!Number.isFinite(width)) {
        return 0;
    }

    return Math.max(0, width);
}

function getTabsWidth<TTab>(
    tabs: readonly TTab[],
    getTabWidth: (tab: TTab) => number,
    itemGap: number
): number {
    if (tabs.length === 0) {
        return 0;
    }

    const tabWidths = tabs.reduce(
        (totalWidth, tab) => totalWidth + normalizeWidth(getTabWidth(tab)),
        0
    );
    return tabWidths + itemGap * Math.max(0, tabs.length - 1);
}

export function partitionSheetTabs<TTab extends SheetTabLike>(
    tabs: readonly TTab[],
    maxVisibleTabs: number
): SheetTabOverflowLayout<TTab> {
    if (tabs.length === 0) {
        return {
            activeTab: null,
            activeTabVisibleIndex: -1,
            hasOverflow: false,
            overflowTabs: [],
            visibleTabs: [],
        };
    }

    const visibleCount = normalizeMaxVisibleTabs(maxVisibleTabs);
    const activeTabIndex = getActiveTabIndex(tabs);
    const activeTab = tabs[activeTabIndex] ?? null;

    if (tabs.length <= visibleCount) {
        return {
            activeTab,
            activeTabVisibleIndex: activeTabIndex,
            hasOverflow: false,
            overflowTabs: [],
            visibleTabs: [...tabs],
        };
    }

    // Keep the active tab visible and try to preserve nearby sheets on both sides.
    let startIndex = Math.max(0, activeTabIndex - Math.floor((visibleCount - 1) / 2));
    const maxStartIndex = tabs.length - visibleCount;
    startIndex = Math.min(startIndex, maxStartIndex);

    const endIndex = startIndex + visibleCount;
    const visibleTabs = tabs.slice(startIndex, endIndex);
    const overflowTabs = [...tabs.slice(0, startIndex), ...tabs.slice(endIndex)];

    return {
        activeTab,
        activeTabVisibleIndex: activeTabIndex - startIndex,
        hasOverflow: true,
        overflowTabs,
        visibleTabs,
    };
}

export function getMaxVisibleSheetTabsForWidth<TTab extends SheetTabLike>(
    tabs: readonly TTab[],
    {
        containerWidth,
        getTabWidth,
        itemGap = 0,
        overflowTriggerWidth,
    }: SheetTabWidthLayoutOptions<TTab>
): number {
    if (tabs.length === 0) {
        return 0;
    }

    const normalizedContainerWidth = normalizeWidth(containerWidth);
    const normalizedOverflowTriggerWidth = normalizeWidth(overflowTriggerWidth);
    const normalizedItemGap = normalizeWidth(itemGap);

    if (normalizedContainerWidth === 0) {
        return 1;
    }

    const allTabsWidth = getTabsWidth(tabs, getTabWidth, normalizedItemGap);
    if (allTabsWidth <= normalizedContainerWidth) {
        return tabs.length;
    }

    const availableVisibleTabsWidth = Math.max(
        0,
        normalizedContainerWidth - normalizedOverflowTriggerWidth - normalizedItemGap
    );

    for (let visibleCount = tabs.length - 1; visibleCount >= 1; visibleCount -= 1) {
        const layout = partitionSheetTabs(tabs, visibleCount);
        if (
            getTabsWidth(layout.visibleTabs, getTabWidth, normalizedItemGap) <=
            availableVisibleTabsWidth
        ) {
            return visibleCount;
        }
    }

    return 1;
}
