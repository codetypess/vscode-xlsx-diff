/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { resolveDisplayLanguage } from "../display-language";

suite("Display language", () => {
    test("honors explicit language overrides", () => {
        assert.strictEqual(resolveDisplayLanguage("en", "zh-cn"), "en");
        assert.strictEqual(resolveDisplayLanguage("zh-cn", "en"), "zh-cn");
    });

    test("falls back to the VS Code language in auto mode", () => {
        assert.strictEqual(resolveDisplayLanguage("auto", "zh-cn"), "zh-cn");
        assert.strictEqual(resolveDisplayLanguage(undefined, "en-us"), "en");
    });
});
