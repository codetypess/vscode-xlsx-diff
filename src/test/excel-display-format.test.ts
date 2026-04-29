import * as assert from "assert";
import { formatExcelLikeDisplayValue } from "../core/fastxlsx/excel-display-format";

suite("Excel display format", () => {
    test("normalizes general numeric display values like Excel", () => {
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 0.1 + 0.2,
                displayValue: "0.30000000000000004",
                numberFormatCode: null,
            }),
            "0.3"
        );
    });

    test("applies fixed decimal number formats", () => {
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 0.1 + 0.2,
                displayValue: "0.30000000000000004",
                numberFormatCode: "0.00",
            }),
            "0.30"
        );
    });

    test("applies grouping and percent formats", () => {
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 1234.5,
                displayValue: "1234.5",
                numberFormatCode: "#,##0.00",
            }),
            "1,234.50"
        );
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 0.1234,
                displayValue: "0.1234",
                numberFormatCode: "0.00%",
            }),
            "12.34%"
        );
    });

    test("falls back to workbook display text for unsupported date-like formats", () => {
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 45292,
                displayValue: "2024-01-01",
                numberFormatCode: "yyyy-mm-dd",
            }),
            "2024-01-01"
        );
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 0,
                displayValue: "1899-12-30",
                numberFormatCode: "yyyy-mm-dd",
            }),
            "1899-12-30"
        );
    });

    test("supports literal zero sections", () => {
        assert.strictEqual(
            formatExcelLikeDisplayValue({
                rawValue: 0,
                displayValue: "0",
                numberFormatCode: '0.00;-0.00;"zero"',
            }),
            "zero"
        );
    });
});
