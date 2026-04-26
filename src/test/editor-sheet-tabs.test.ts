/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { getMaxVisibleSheetTabsForWidth, partitionSheetTabs } from "../webview/editor-sheet-tabs";

interface TestSheetTab {
    key: string;
    label: string;
    isActive: boolean;
}

function createSheetTabs(activeKey: string, count: number): TestSheetTab[] {
    return Array.from({ length: count }, (_, index) => {
        const key = `Sheet${index + 1}`;
        return {
            key,
            label: key,
            isActive: key === activeKey,
        };
    });
}

function getKeys(tabs: readonly TestSheetTab[]): string[] {
    return tabs.map((tab) => tab.key);
}

suite("Editor sheet tab helpers", () => {
    test("keeps every tab visible when the list fits", () => {
        const layout = partitionSheetTabs(createSheetTabs("Sheet2", 3), 4);

        assert.strictEqual(layout.hasOverflow, false);
        assert.strictEqual(layout.activeTab?.key, "Sheet2");
        assert.strictEqual(layout.activeTabVisibleIndex, 1);
        assert.deepStrictEqual(getKeys(layout.visibleTabs), ["Sheet1", "Sheet2", "Sheet3"]);
        assert.deepStrictEqual(getKeys(layout.overflowTabs), []);
    });

    test("anchors the visible window near the start when the active sheet is early", () => {
        const layout = partitionSheetTabs(createSheetTabs("Sheet1", 8), 4);

        assert.strictEqual(layout.hasOverflow, true);
        assert.strictEqual(layout.activeTab?.key, "Sheet1");
        assert.strictEqual(layout.activeTabVisibleIndex, 0);
        assert.deepStrictEqual(getKeys(layout.visibleTabs), [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet4",
        ]);
        assert.deepStrictEqual(getKeys(layout.overflowTabs), [
            "Sheet5",
            "Sheet6",
            "Sheet7",
            "Sheet8",
        ]);
    });

    test("surfaces a middle active tab with nearby sheets kept visible", () => {
        const layout = partitionSheetTabs(createSheetTabs("Sheet5", 8), 4);

        assert.strictEqual(layout.activeTab?.key, "Sheet5");
        assert.strictEqual(layout.activeTabVisibleIndex, 1);
        assert.deepStrictEqual(getKeys(layout.visibleTabs), [
            "Sheet4",
            "Sheet5",
            "Sheet6",
            "Sheet7",
        ]);
        assert.deepStrictEqual(getKeys(layout.overflowTabs), [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet8",
        ]);
    });

    test("brings a sheet chosen from overflow into the visible tab row", () => {
        const initialLayout = partitionSheetTabs(createSheetTabs("Sheet1", 8), 4);
        const overflowSelectionLayout = partitionSheetTabs(createSheetTabs("Sheet7", 8), 4);

        assert.deepStrictEqual(getKeys(initialLayout.visibleTabs), [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet4",
        ]);
        assert.deepStrictEqual(getKeys(initialLayout.overflowTabs), [
            "Sheet5",
            "Sheet6",
            "Sheet7",
            "Sheet8",
        ]);

        assert.strictEqual(overflowSelectionLayout.activeTab?.key, "Sheet7");
        assert.ok(getKeys(overflowSelectionLayout.visibleTabs).includes("Sheet7"));
        assert.ok(getKeys(overflowSelectionLayout.overflowTabs).includes("Sheet1"));
        assert.deepStrictEqual(getKeys(overflowSelectionLayout.visibleTabs), [
            "Sheet5",
            "Sheet6",
            "Sheet7",
            "Sheet8",
        ]);
    });

    test("falls back to a single visible tab when the limit is invalid", () => {
        const layout = partitionSheetTabs(createSheetTabs("Sheet3", 5), 0);

        assert.strictEqual(layout.activeTab?.key, "Sheet3");
        assert.strictEqual(layout.activeTabVisibleIndex, 0);
        assert.deepStrictEqual(getKeys(layout.visibleTabs), ["Sheet3"]);
        assert.deepStrictEqual(getKeys(layout.overflowTabs), [
            "Sheet1",
            "Sheet2",
            "Sheet4",
            "Sheet5",
        ]);
    });

    test("fits more visible tabs when their measured widths are narrow", () => {
        const tabs = createSheetTabs("Sheet5", 12);
        const visibleCount = getMaxVisibleSheetTabsForWidth(tabs, {
            containerWidth: 470,
            getTabWidth: () => 60,
            itemGap: 1,
            overflowTriggerWidth: 72,
        });

        assert.strictEqual(visibleCount, 6);
        assert.deepStrictEqual(getKeys(partitionSheetTabs(tabs, visibleCount).visibleTabs), [
            "Sheet3",
            "Sheet4",
            "Sheet5",
            "Sheet6",
            "Sheet7",
            "Sheet8",
        ]);
    });

    test("keeps every tab visible when the measured widths fit the container", () => {
        const tabs = createSheetTabs("Sheet2", 4);
        const visibleCount = getMaxVisibleSheetTabsForWidth(tabs, {
            containerWidth: 360,
            getTabWidth: () => 80,
            itemGap: 1,
            overflowTriggerWidth: 72,
        });

        assert.strictEqual(visibleCount, 4);
    });
});
