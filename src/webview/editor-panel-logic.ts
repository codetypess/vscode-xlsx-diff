import type { CellEdit } from "../core/fastxlsx/write-cell-value";
import { createCellKey, getColumnNumber } from "../core/model/cells";
import type { EditorPanelState } from "../core/model/types";
import { isCellWithinSelectionRange } from "./editor-selection-range";
import type {
    EditorSearchRequest,
    EditorPanelStrings,
    EditorSearchResult,
    EditorSearchMatch,
    EditorSearchScope,
    SearchOptions,
    WorkingSheetEntry,
} from "./editor-panel-types";

function escapeRegex(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

interface SearchableEditorCell {
    rowNumber: number;
    columnNumber: number;
    value: string;
    formula: string | null;
}

interface SearchMatchContext {
    effectiveScope: EditorSearchScope;
    matches: EditorSearchMatch[];
}

export function createEditorSearchPattern(query: string, options: SearchOptions): RegExp {
    const source = options.isRegexp ? query.trim() : escapeRegex(query.trim());
    const wrappedSource = options.wholeWord ? `\\b(?:${source})\\b` : source;
    return new RegExp(wrappedSource, options.matchCase ? "" : "i");
}

function getActiveSearchSheetEntry(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState
): WorkingSheetEntry | undefined {
    return sheetEntries.find((entry) => entry.key === state.activeSheetKey) ?? sheetEntries[0];
}

function getEffectiveSearchScope(
    scope: EditorSearchScope | undefined,
    hasSelectionRange: boolean
): EditorSearchScope {
    return scope === "selection" && hasSelectionRange ? "selection" : "sheet";
}

function getSearchableEditorCells(
    sheetEntry: WorkingSheetEntry,
    pendingEdits: CellEdit[] | undefined
): SearchableEditorCell[] {
    const cells = new Map<string, SearchableEditorCell>(
        Object.values(sheetEntry.sheet.cells).map((cell) => [
            cell.key,
            {
                rowNumber: cell.rowNumber,
                columnNumber: cell.columnNumber,
                value: cell.displayValue,
                formula: cell.formula ?? null,
            },
        ])
    );

    for (const edit of pendingEdits ?? []) {
        if (edit.sheetName !== sheetEntry.sheet.name) {
            continue;
        }

        cells.set(createCellKey(edit.rowNumber, edit.columnNumber), {
            rowNumber: edit.rowNumber,
            columnNumber: edit.columnNumber,
            value: edit.value,
            formula: null,
        });
    }

    return [...cells.values()];
}

function getEditorSearchMatchContext(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState,
    query: string,
    options: SearchOptions,
    searchContext: {
        pendingEdits?: CellEdit[];
        scope?: EditorSearchScope;
        selectionRange?: {
            startRow: number;
            endRow: number;
            startColumn: number;
            endColumn: number;
        };
    } = {}
): SearchMatchContext | null {
    const normalizedQuery = query.trim();
    if (!normalizedQuery) {
        return null;
    }

    const pattern = createEditorSearchPattern(normalizedQuery, options);
    const activeSheetEntry = getActiveSearchSheetEntry(sheetEntries, state);
    if (!activeSheetEntry) {
        return null;
    }

    const effectiveScope = getEffectiveSearchScope(
        searchContext.scope,
        Boolean(searchContext.selectionRange)
    );
    const matches = getSearchableEditorCells(activeSheetEntry, searchContext.pendingEdits)
        .filter((cell) => {
            if (
                effectiveScope === "selection" &&
                !isCellWithinSelectionRange(
                    searchContext.selectionRange ?? null,
                    cell.rowNumber,
                    cell.columnNumber
                )
            ) {
                return false;
            }

            const value = cell.value;
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

    return {
        effectiveScope,
        matches,
    };
}

export function findEditorSearchMatch(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState,
    query: string,
    direction: "next" | "prev",
    options: SearchOptions,
    searchContext: {
        pendingEdits?: CellEdit[];
        scope?: EditorSearchScope;
        selectionRange?: {
            startRow: number;
            endRow: number;
            startColumn: number;
            endColumn: number;
        };
    } = {}
): EditorSearchMatch | null {
    const matchContext = getEditorSearchMatchContext(sheetEntries, state, query, options, searchContext);
    if (!matchContext) {
        return null;
    }
    const { effectiveScope, matches } = matchContext;
    const activeSheetEntry = getActiveSearchSheetEntry(sheetEntries, state);
    if (!activeSheetEntry) {
        return null;
    }

    const selectionOnCurrentSheet =
        state.selectedCell &&
        state.selectedCell.rowNumber >= 1 &&
        state.selectedCell.rowNumber <= activeSheetEntry.sheet.rowCount &&
        state.selectedCell.columnNumber >= 1 &&
        state.selectedCell.columnNumber <= activeSheetEntry.sheet.columnCount;
    const selectionWithinScope =
        selectionOnCurrentSheet &&
        (effectiveScope !== "selection" ||
            isCellWithinSelectionRange(
                searchContext.selectionRange ?? null,
                state.selectedCell!.rowNumber,
                state.selectedCell!.columnNumber
            ));
    const anchor =
        selectionWithinScope
            ? {
                  rowNumber: state.selectedCell!.rowNumber,
                  columnNumber: state.selectedCell!.columnNumber,
              }
            : effectiveScope === "selection" && searchContext.selectionRange
              ? {
                    rowNumber: searchContext.selectionRange.startRow,
                    columnNumber: searchContext.selectionRange.startColumn - 1,
                }
              : {
                    rowNumber: 1,
                    columnNumber: 0,
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

export function resolveEditorSearchResult(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState,
    request: EditorSearchRequest,
    searchContext: {
        pendingEdits?: CellEdit[];
    } = {}
): EditorSearchResult {
    try {
        const matchContext = getEditorSearchMatchContext(
            sheetEntries,
            state,
            request.query,
            request.options,
            {
                pendingEdits: searchContext.pendingEdits,
                scope: request.scope,
                selectionRange: request.selectionRange,
            }
        );
        if (!matchContext) {
            return { status: "no-match" };
        }

        const match = findEditorSearchMatch(
            sheetEntries,
            state,
            request.query,
            request.direction,
            request.options,
            {
                pendingEdits: searchContext.pendingEdits,
                scope: request.scope,
                selectionRange: request.selectionRange,
            }
        );

        if (!match) {
            return { status: "no-match" };
        }

        const matchIndex = matchContext.matches.findIndex(
            (candidate) =>
                candidate.sheetKey === match.sheetKey &&
                candidate.rowNumber === match.rowNumber &&
                candidate.columnNumber === match.columnNumber
        );

        return {
            status: "matched",
            match,
            matchCount: matchContext.matches.length,
            matchIndex: matchIndex >= 0 ? matchIndex + 1 : undefined,
        };
    } catch {
        return { status: "invalid-pattern" };
    }
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
