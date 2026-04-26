export interface SelectionPreviewInlineDiff {
    before: string;
    changed: string;
    after: string;
    hasDifference: boolean;
}

export function getSelectionPreviewInlineDiff(
    value: string,
    otherValue: string
): SelectionPreviewInlineDiff {
    if (value === otherValue) {
        return {
            before: value,
            changed: "",
            after: "",
            hasDifference: false,
        };
    }

    const maxPrefixLength = Math.min(value.length, otherValue.length);
    let prefixLength = 0;
    while (
        prefixLength < maxPrefixLength &&
        value.charCodeAt(prefixLength) === otherValue.charCodeAt(prefixLength)
    ) {
        prefixLength += 1;
    }

    let suffixLength = 0;
    const maxSuffixLength = Math.min(value.length - prefixLength, otherValue.length - prefixLength);
    while (
        suffixLength < maxSuffixLength &&
        value.charCodeAt(value.length - suffixLength - 1) ===
            otherValue.charCodeAt(otherValue.length - suffixLength - 1)
    ) {
        suffixLength += 1;
    }

    return {
        before: value.slice(0, prefixLength),
        changed: value.slice(prefixLength, value.length - suffixLength),
        after: suffixLength > 0 ? value.slice(value.length - suffixLength) : "",
        hasDifference: true,
    };
}
