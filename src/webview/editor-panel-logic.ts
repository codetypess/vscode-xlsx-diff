import { getColumnNumber } from "../core/model/cells";
import type { EditorPanelState } from "../core/model/types";
import type {
    EditorPanelStrings,
    EditorSearchMatch,
    SearchOptions,
    WorkingSheetEntry,
} from "./editor-panel-types";

function escapeRegex(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function createEditorSearchPattern(query: string, options: SearchOptions): RegExp {
    const source = options.isRegexp ? query.trim() : escapeRegex(query.trim());
    const wrappedSource = options.wholeWord ? `\\b(?:${source})\\b` : source;
    return new RegExp(wrappedSource, options.matchCase ? "" : "i");
}

export function findEditorSearchMatch(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState,
    query: string,
    direction: "next" | "prev",
    options: SearchOptions
): EditorSearchMatch | null {
    const normalizedQuery = query.trim();
    if (!normalizedQuery) {
        return null;
    }

    const pattern = createEditorSearchPattern(normalizedQuery, options);
    const activeSheetEntry =
        sheetEntries.find((entry) => entry.key === state.activeSheetKey) ?? sheetEntries[0];
    if (!activeSheetEntry) {
        return null;
    }

    const matches = Object.values(activeSheetEntry.sheet.cells)
        .filter((cell) => {
            const value = cell.displayValue;
            const formula = cell.formula ?? "";
            return pattern.test(value) || pattern.test(formula);
        })
        .map((cell) => ({
            sheetKey: activeSheetEntry.key,
            rowNumber: cell.rowNumber,
            columnNumber: cell.columnNumber,
        }));

    if (matches.length === 0) {
        return null;
    }

    matches.sort((left, right) => {
        if (left.rowNumber !== right.rowNumber) {
            return left.rowNumber - right.rowNumber;
        }

        return left.columnNumber - right.columnNumber;
    });

    const selectionOnCurrentSheet =
        state.selectedCell &&
        state.selectedCell.rowNumber >= 1 &&
        state.selectedCell.rowNumber <= activeSheetEntry.sheet.rowCount;
    const anchor = {
        rowNumber: selectionOnCurrentSheet ? state.selectedCell!.rowNumber : 1,
        columnNumber: selectionOnCurrentSheet ? state.selectedCell!.columnNumber : 0,
    };
    const compare = (
        candidate: { rowNumber: number; columnNumber: number },
        current: { rowNumber: number; columnNumber: number }
    ): number => {
        if (candidate.rowNumber !== current.rowNumber) {
            return candidate.rowNumber - current.rowNumber;
        }

        return candidate.columnNumber - current.columnNumber;
    };

    if (direction === "prev") {
        for (let index = matches.length - 1; index >= 0; index -= 1) {
            if (compare(matches[index]!, anchor) < 0) {
                return matches[index]!;
            }
        }

        return matches[matches.length - 1]!;
    }

    return matches.find((match) => compare(match, anchor) > 0) ?? matches[0]!;
}

export function resolveEditorCellReference(
    sheetEntries: WorkingSheetEntry[],
    activeSheetKey: string | null,
    reference: string
): EditorSearchMatch | null {
    const trimmedReference = reference.trim();
    if (!trimmedReference) {
        return null;
    }

    const separatorIndex = trimmedReference.lastIndexOf("!");
    const sheetName = separatorIndex > 0 ? trimmedReference.slice(0, separatorIndex).trim() : null;
    const address =
        separatorIndex > 0 ? trimmedReference.slice(separatorIndex + 1).trim() : trimmedReference;
    const addressMatch = /^([A-Za-z]+)(\d+)$/.exec(address);
    if (!addressMatch) {
        return null;
    }

    const columnNumber = getColumnNumber(addressMatch[1]!);
    const rowNumber = Number(addressMatch[2]!);
    if (!columnNumber || rowNumber < 1) {
        return null;
    }

    const targetSheet = sheetName
        ? sheetEntries.find(
              (entry) =>
                  entry.sheet.name === sheetName ||
                  entry.sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()
          )
        : (sheetEntries.find((entry) => entry.key === activeSheetKey) ?? sheetEntries[0]);

    if (
        !targetSheet ||
        rowNumber > targetSheet.sheet.rowCount ||
        columnNumber > targetSheet.sheet.columnCount
    ) {
        return null;
    }

    return {
        sheetKey: targetSheet.key,
        rowNumber,
        columnNumber,
    };
}

export function validateEditorSheetName(
    value: string,
    sheetEntries: WorkingSheetEntry[],
    strings: Pick<
        EditorPanelStrings,
        "sheetNameDuplicate" | "sheetNameEmpty" | "sheetNameInvalidChars" | "sheetNameTooLong"
    >,
    currentSheetKey?: string
): string | undefined {
    const trimmed = value.trim();

    if (!trimmed) {
        return strings.sheetNameEmpty;
    }

    if (trimmed.length > 31) {
        return strings.sheetNameTooLong;
    }

    if (/[\\/:?*\[\]]/.test(trimmed)) {
        return strings.sheetNameInvalidChars;
    }

    if (
        sheetEntries.some(
            (entry) =>
                entry.key !== currentSheetKey &&
                entry.sheet.name.toLocaleLowerCase() === trimmed.toLocaleLowerCase()
        )
    ) {
        return strings.sheetNameDuplicate;
    }

    return undefined;
}

export function getNewEditorSheetName(
    sheetEntries: WorkingSheetEntry[],
    baseName: string
): string {
    const existingNames = new Set(sheetEntries.map((entry) => entry.sheet.name));

    let suffix = 1;
    while (existingNames.has(`${baseName}${suffix}`)) {
        suffix += 1;
    }

    return `${baseName}${suffix}`;
}

export function getInsertEditorSheetIndex(
    sheetEntries: WorkingSheetEntry[],
    activeSheetKey: string | null
): number {
    if (sheetEntries.length === 0) {
        return 0;
    }

    const activeSheetIndex = sheetEntries.findIndex((sheet) => sheet.key === activeSheetKey);
    return activeSheetIndex >= 0 ? activeSheetIndex + 1 : sheetEntries.length;
}
