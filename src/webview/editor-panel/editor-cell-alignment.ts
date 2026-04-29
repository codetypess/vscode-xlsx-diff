import type { CellAlignmentSnapshot } from "../../core/model/alignment";

type CellJustifyContent = "flex-start" | "center" | "flex-end";
type CellAlignItems = "flex-start" | "center" | "flex-end";
type CellTextAlign = "left" | "center" | "right";

export interface CellContentAlignmentStyle {
    justifyContent: CellJustifyContent;
    alignItems: CellAlignItems;
    textAlign: CellTextAlign;
    height: "100%";
    maxHeight: "100%";
}

function getNormalizedCellHorizontalDisplayAlignment(
    alignment: CellAlignmentSnapshot | null
): CellTextAlign {
    switch (alignment?.horizontal) {
        case "center":
        case "centerContinuous":
            return "center";
        case "right":
            return "right";
        default:
            return "left";
    }
}

export function getCellHorizontalJustifyContent(
    alignment: CellAlignmentSnapshot | null
): CellContentAlignmentStyle["justifyContent"] {
    switch (getNormalizedCellHorizontalDisplayAlignment(alignment)) {
        case "center":
            return "center";
        case "right":
            return "flex-end";
        default:
            return "flex-start";
    }
}

export function getCellHorizontalTextAlign(
    alignment: CellAlignmentSnapshot | null
): CellContentAlignmentStyle["textAlign"] {
    return getNormalizedCellHorizontalDisplayAlignment(alignment);
}

export function getCellContentAlignmentStyle(
    alignment: CellAlignmentSnapshot | null
): CellContentAlignmentStyle {
    return {
        justifyContent: getCellHorizontalJustifyContent(alignment),
        alignItems:
            alignment?.vertical === "center"
                ? "center"
                : alignment?.vertical === "bottom"
                  ? "flex-end"
                  : "flex-start",
        textAlign: getCellHorizontalTextAlign(alignment),
        height: "100%",
        maxHeight: "100%",
    };
}

export function getToolbarHorizontalAlignment(
    alignment: CellAlignmentSnapshot | null
): "left" | "center" | "right" | undefined {
    switch (alignment?.horizontal) {
        case "left":
            return "left";
        case "center":
        case "centerContinuous":
            return "center";
        case "right":
            return "right";
        default:
            return undefined;
    }
}

export function getToolbarVerticalAlignment(
    alignment: CellAlignmentSnapshot | null
): "top" | "center" | "bottom" | undefined {
    switch (alignment?.vertical) {
        case "top":
            return "top";
        case "center":
            return "center";
        case "bottom":
            return "bottom";
        default:
            return undefined;
    }
}
