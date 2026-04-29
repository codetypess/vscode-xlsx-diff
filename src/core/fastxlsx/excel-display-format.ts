import type { CellValue } from "fastxlsx";

const GENERAL_NUMBER_PRECISION = 15;
const MAX_FORMATTED_FRACTION_DIGITS = 15;

interface ResolvedFormatSection {
    section: string;
    implicitNegative: boolean;
}

function normalizeGeneralNumber(value: number): string {
    if (Object.is(value, -0)) {
        return "0";
    }

    const normalizedValue = Number(value.toPrecision(GENERAL_NUMBER_PRECISION));
    return Object.is(normalizedValue, -0) ? "0" : String(normalizedValue);
}

function splitFormatSections(formatCode: string): string[] {
    const sections: string[] = [];
    let currentSection = "";
    let inQuotedLiteral = false;

    for (let index = 0; index < formatCode.length; index += 1) {
        const character = formatCode[index]!;

        if (character === '"') {
            inQuotedLiteral = !inQuotedLiteral;
            currentSection += character;
            continue;
        }

        if (character === "\\" && index + 1 < formatCode.length) {
            currentSection += character;
            currentSection += formatCode[index + 1]!;
            index += 1;
            continue;
        }

        if (character === ";" && !inQuotedLiteral) {
            sections.push(currentSection);
            currentSection = "";
            continue;
        }

        currentSection += character;
    }

    sections.push(currentSection);
    return sections;
}

function extractBracketLiteral(token: string): string | null {
    const currencyTokenMatch = token.match(/^\[\$([^\]-]+)(?:-[^\]]+)?\]$/);
    if (!currencyTokenMatch) {
        return null;
    }

    return currencyTokenMatch[1] ?? null;
}

function sanitizeFormatSection(section: string): {
    sanitizedSection: string;
    unsupported: boolean;
} {
    let sanitizedSection = "";
    let unsupported = false;

    for (let index = 0; index < section.length; index += 1) {
        const character = section[index]!;

        if (character === '"') {
            let literal = "";
            index += 1;
            while (index < section.length && section[index] !== '"') {
                literal += section[index]!;
                index += 1;
            }
            sanitizedSection += literal;
            continue;
        }

        if (character === "\\") {
            if (index + 1 < section.length) {
                sanitizedSection += section[index + 1]!;
                index += 1;
            }
            continue;
        }

        if (character === "_") {
            index += 1;
            continue;
        }

        if (character === "*") {
            index += 1;
            continue;
        }

        if (character === "[") {
            const endIndex = section.indexOf("]", index);
            if (endIndex === -1) {
                unsupported = true;
                break;
            }

            const token = section.slice(index, endIndex + 1);
            const bracketLiteral = extractBracketLiteral(token);
            if (bracketLiteral !== null) {
                sanitizedSection += bracketLiteral;
                index = endIndex;
                continue;
            }

            if (/^\[(<=|>=|<>|=|<|>).+\]$/i.test(token) || /^\[[hms]+\]$/i.test(token)) {
                unsupported = true;
                break;
            }

            index = endIndex;
            continue;
        }

        sanitizedSection += character;
    }

    return {
        sanitizedSection,
        unsupported,
    };
}

function resolveFormatSection(value: number, formatCode: string | null): ResolvedFormatSection | null {
    if (!formatCode) {
        return null;
    }

    const normalizedCode = formatCode.trim();
    if (!normalizedCode || /^general$/i.test(normalizedCode)) {
        return null;
    }

    const sections = splitFormatSections(normalizedCode);
    if (sections.length === 0) {
        return null;
    }

    if (value > 0) {
        return {
            section: sections[0] ?? "",
            implicitNegative: false,
        };
    }

    if (value < 0) {
        return {
            section: sections[1] ?? sections[0] ?? "",
            implicitNegative: sections[1] === undefined,
        };
    }

    return {
        section: sections[2] ?? sections[0] ?? "",
        implicitNegative: false,
    };
}

function containsUnsupportedNumericTokens(section: string): boolean {
    const normalizedSection = section.replace(/E[+-]0+/gi, "");
    return (
        /@/.test(normalizedSection) ||
        /[ymdhsa]/i.test(normalizedSection) ||
        /[#0?]+\/[#0?]+/.test(normalizedSection)
    );
}

function formatScientificNumber(
    value: number,
    pattern: string
): string | null {
    const scientificMatch = pattern.match(/^(.*?)([0#?]+(?:\.[0#?]+)?)E([+-])([0#?]+)(.*)$/i);
    if (!scientificMatch) {
        return null;
    }

    const prefix = scientificMatch[1] ?? "";
    const mantissaPattern = scientificMatch[2] ?? "";
    const exponentSignMode = scientificMatch[3] ?? "+";
    const exponentDigitsPattern = scientificMatch[4] ?? "";
    const suffix = scientificMatch[5] ?? "";
    const fractionPattern = mantissaPattern.split(".")[1] ?? "";
    const minimumFractionDigits = Math.min(
        fractionPattern.replaceAll("#", "").replaceAll("?", "").length,
        MAX_FORMATTED_FRACTION_DIGITS
    );
    const maximumFractionDigits = Math.min(
        fractionPattern.length,
        MAX_FORMATTED_FRACTION_DIGITS
    );
    const normalizedValue = Number(value.toPrecision(GENERAL_NUMBER_PRECISION));
    const [rawMantissa, rawExponent] = normalizedValue
        .toExponential(maximumFractionDigits)
        .split("e");
    if (!rawMantissa || !rawExponent) {
        return null;
    }

    let mantissa = rawMantissa;
    if (maximumFractionDigits > minimumFractionDigits && mantissa.includes(".")) {
        const [integerPart, fractionPart = ""] = mantissa.split(".");
        let trimmedFractionPart = fractionPart;
        while (
            trimmedFractionPart.length > minimumFractionDigits &&
            trimmedFractionPart.endsWith("0")
        ) {
            trimmedFractionPart = trimmedFractionPart.slice(0, -1);
        }
        mantissa = trimmedFractionPart ? `${integerPart}.${trimmedFractionPart}` : integerPart;
    }

    const exponentValue = Number(rawExponent);
    const absoluteExponent = String(Math.abs(exponentValue)).padStart(
        exponentDigitsPattern.length,
        "0"
    );
    const exponentSign =
        exponentValue < 0 ? "-" : exponentSignMode === "+" ? "+" : "";
    return `${prefix}${mantissa.toUpperCase()}E${exponentSign}${absoluteExponent}${suffix}`;
}

function formatPatternNumber(value: number, patternSection: string): string | null {
    const { sanitizedSection, unsupported } = sanitizeFormatSection(patternSection);
    if (unsupported || !sanitizedSection.trim()) {
        return null;
    }

    if (containsUnsupportedNumericTokens(sanitizedSection)) {
        return null;
    }

    const firstPlaceholderIndex = sanitizedSection.search(/[0#?]/);
    if (firstPlaceholderIndex === -1) {
        return sanitizedSection;
    }

    const lastPlaceholderIndex = Math.max(
        sanitizedSection.lastIndexOf("0"),
        sanitizedSection.lastIndexOf("#"),
        sanitizedSection.lastIndexOf("?")
    );
    if (lastPlaceholderIndex < firstPlaceholderIndex) {
        return null;
    }

    let prefix = sanitizedSection.slice(0, firstPlaceholderIndex);
    let numericPattern = sanitizedSection.slice(firstPlaceholderIndex, lastPlaceholderIndex + 1);
    let suffix = sanitizedSection.slice(lastPlaceholderIndex + 1);

    if (/E[+-][0#?]+/i.test(numericPattern)) {
        return formatScientificNumber(value, `${prefix}${numericPattern}${suffix}`);
    }

    const percentCount = (prefix + numericPattern + suffix).match(/%/g)?.length ?? 0;
    const scaleByPercent = 100 ** percentCount;
    prefix = prefix.replaceAll("%", "");
    numericPattern = numericPattern.replaceAll("%", "");
    suffix = suffix.replaceAll("%", "") + "%".repeat(percentCount);

    const decimalIndex = numericPattern.indexOf(".");
    const integerPattern =
        decimalIndex === -1 ? numericPattern : numericPattern.slice(0, decimalIndex);
    const fractionPattern =
        decimalIndex === -1 ? "" : numericPattern.slice(decimalIndex + 1);
    const scalingCommaMatch = integerPattern.match(/,+$/);
    const scaleByThousands = scalingCommaMatch ? scalingCommaMatch[0].length : 0;
    const normalizedIntegerPattern = integerPattern
        .replace(/,+$/, "")
        .replaceAll(",", "");
    const minimumIntegerDigits = Math.max(
        1,
        (normalizedIntegerPattern.match(/0/g)?.length ?? 0)
    );
    const minimumFractionDigits = Math.min(
        fractionPattern.replaceAll("#", "").replaceAll("?", "").length,
        MAX_FORMATTED_FRACTION_DIGITS
    );
    const maximumFractionDigits = Math.min(
        fractionPattern.length,
        MAX_FORMATTED_FRACTION_DIGITS
    );
    const normalizedValue =
        Number(value.toPrecision(GENERAL_NUMBER_PRECISION)) *
        scaleByPercent /
        1000 ** scaleByThousands;
    const formatter = new Intl.NumberFormat("en-US", {
        minimumIntegerDigits,
        minimumFractionDigits,
        maximumFractionDigits,
        useGrouping: integerPattern.includes(","),
    });

    return `${prefix}${formatter.format(normalizedValue)}${suffix}`;
}

export function formatExcelLikeDisplayValue({
    rawValue,
    displayValue,
    numberFormatCode,
}: {
    rawValue: CellValue;
    displayValue: string | null;
    numberFormatCode: string | null;
}): string | null {
    if (rawValue === null) {
        return displayValue;
    }

    if (typeof rawValue !== "number" || !Number.isFinite(rawValue)) {
        return displayValue ?? String(rawValue);
    }

    const resolvedSection = resolveFormatSection(rawValue, numberFormatCode);
    if (!resolvedSection) {
        return normalizeGeneralNumber(rawValue);
    }

    const formattedValue = formatPatternNumber(
        Math.abs(rawValue),
        resolvedSection.section
    );
    if (formattedValue === null) {
        return displayValue ?? normalizeGeneralNumber(rawValue);
    }

    if (rawValue < 0 && resolvedSection.implicitNegative) {
        return formattedValue.startsWith("-") || formattedValue.startsWith("(")
            ? formattedValue
            : `-${formattedValue}`;
    }

    return formattedValue;
}
