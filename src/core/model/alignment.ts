import type {
    CellStyleAlignment,
    CellStyleHorizontalAlignment,
    CellStyleVerticalAlignment,
} from "fastxlsx";

export type CellAlignmentSnapshot = CellStyleAlignment;

export type SheetCellAlignmentsSnapshot = Record<string, CellAlignmentSnapshot>;
export type SheetRowAlignmentsSnapshot = Record<string, CellAlignmentSnapshot>;
export type SheetColumnAlignmentsSnapshot = Record<string, CellAlignmentSnapshot>;

export type EditorHorizontalAlignment = Extract<
    CellStyleHorizontalAlignment,
    "left" | "center" | "right"
>;
export type EditorVerticalAlignment = Extract<
    CellStyleVerticalAlignment,
    "top" | "center" | "bottom"
>;

export interface EditorAlignmentPatch {
    horizontal?: EditorHorizontalAlignment;
    vertical?: EditorVerticalAlignment;
}

const ALIGNMENT_KEYS = [
    "horizontal",
    "vertical",
    "textRotation",
    "wrapText",
    "shrinkToFit",
    "indent",
    "relativeIndent",
    "justifyLastLine",
    "readingOrder",
] as const;

type AlignmentKey = (typeof ALIGNMENT_KEYS)[number];

export function cloneCellAlignment(
    alignment: CellAlignmentSnapshot | null | undefined
): CellAlignmentSnapshot | null {
    if (!alignment) {
        return null;
    }

    const nextAlignment = {} as CellAlignmentSnapshot;
    const nextAlignmentRecord = nextAlignment as Record<
        AlignmentKey,
        CellAlignmentSnapshot[AlignmentKey] | undefined
    >;
    for (const key of ALIGNMENT_KEYS) {
        const value = alignment[key];
        if (value !== undefined) {
            nextAlignmentRecord[key] = value;
        }
    }

    return Object.keys(nextAlignment).length > 0 ? nextAlignment : null;
}

export function cloneCellAlignmentMap(
    alignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): Record<string, CellAlignmentSnapshot> {
    return Object.fromEntries(
        Object.entries(alignments ?? {})
            .flatMap(([key, alignment]) => {
                const nextAlignment = cloneCellAlignment(alignment);
                return nextAlignment ? [[key, nextAlignment] as const] : [];
            })
            .sort(([leftKey], [rightKey]) => leftKey.localeCompare(rightKey))
    );
}

export function areCellAlignmentsEqual(
    left: CellAlignmentSnapshot | null | undefined,
    right: CellAlignmentSnapshot | null | undefined
): boolean {
    const normalizedLeft = cloneCellAlignment(left);
    const normalizedRight = cloneCellAlignment(right);
    if (!normalizedLeft || !normalizedRight) {
        return normalizedLeft === normalizedRight;
    }

    return ALIGNMENT_KEYS.every((key) => normalizedLeft[key] === normalizedRight[key]);
}

export function areCellAlignmentMapsEquivalent(
    left: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    right: Readonly<Record<string, CellAlignmentSnapshot>> | undefined
): boolean {
    const normalizedLeft = cloneCellAlignmentMap(left);
    const normalizedRight = cloneCellAlignmentMap(right);
    const leftKeys = Object.keys(normalizedLeft);
    const rightKeys = Object.keys(normalizedRight);
    if (leftKeys.length !== rightKeys.length) {
        return false;
    }

    return leftKeys.every((key) => areCellAlignmentsEqual(normalizedLeft[key], normalizedRight[key]));
}

export function mergeCellAlignments(
    ...alignments: Array<CellAlignmentSnapshot | null | undefined>
): CellAlignmentSnapshot | null {
    const nextAlignment = {} as CellAlignmentSnapshot;
    const nextAlignmentRecord = nextAlignment as Record<
        AlignmentKey,
        CellAlignmentSnapshot[AlignmentKey] | undefined
    >;

    for (const alignment of alignments) {
        if (!alignment) {
            continue;
        }

        for (const key of ALIGNMENT_KEYS) {
            const value = alignment[key];
            if (value !== undefined) {
                nextAlignmentRecord[key] = value;
            }
        }
    }

    return Object.keys(nextAlignment).length > 0 ? nextAlignment : null;
}

export function applyCellAlignmentPatch(
    alignment: CellAlignmentSnapshot | null | undefined,
    patch: Partial<CellAlignmentSnapshot>
): CellAlignmentSnapshot | null {
    return mergeCellAlignments(alignment, patch as CellAlignmentSnapshot);
}
