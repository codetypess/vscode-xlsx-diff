/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { resolveDisplayLanguage } from "../display-language";

suite("Display language", () => {
    test("honors explicit language overrides", () => {
        assert.strictEqual(resolveDisplayLanguage("en", "zh-CN"), "en");
        assert.strictEqual(resolveDisplayLanguage("zh-CN", "en"), "zh-CN");
    });

    test("falls back to the VS Code language in auto mode", () => {
        assert.strictEqual(resolveDisplayLanguage("auto", "zh-CN"), "zh-CN");
        assert.strictEqual(resolveDisplayLanguage(undefined, "en-us"), "en");
    });
});
