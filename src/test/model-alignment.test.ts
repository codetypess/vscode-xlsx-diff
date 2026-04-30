/// <reference types="mocha" />
/// <reference types="node" />

import * as assert from "assert";
import { areCellAlignmentsEqual, cloneCellAlignment } from "../core/model/alignment";

suite("Model alignment helpers", () => {
    test("drops default-only alignment values", () => {
        assert.strictEqual(
            cloneCellAlignment({
                horizontal: "general",
                textRotation: 0,
                wrapText: false,
                shrinkToFit: false,
                indent: 0,
                relativeIndent: 0,
                justifyLastLine: false,
                readingOrder: 0,
            }),
            null
        );
    });

    test("keeps non-default alignment values", () => {
        assert.deepStrictEqual(
            cloneCellAlignment({
                horizontal: "right",
                vertical: "bottom",
                wrapText: true,
            }),
            {
                horizontal: "right",
                vertical: "bottom",
                wrapText: true,
            }
        );
    });

    test("treats default-only and empty alignments as equivalent", () => {
        assert.strictEqual(
            areCellAlignmentsEqual(
                {
                    horizontal: "general",
                    textRotation: 0,
                    wrapText: false,
                },
                null
            ),
            true
        );
    });
});
