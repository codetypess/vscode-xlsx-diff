import type { CellEdit } from "../../core/fastxlsx/write-cell-value";
import { createCellKey, getColumnNumber } from "../../core/model/cells";
import type { EditorPanelState } from "../../core/model/types";
import { isCellWithinSelectionRange, type SelectionRange } from "./editor-selection-range";
import type {
    EditorSearchDirection,
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
    match: EditorSearchMatch;
}

interface SearchMatchContext {
    effectiveScope: EditorSearchScope;
    matches: SearchableEditorCell[];
}

export interface EditorSearchSheetSource {
    key: string;
    rowCount: number;
    columnCount: number;
    cells: Record<
        string,
        {
            key: string;
            rowNumber: number;
            columnNumber: number;
            displayValue: string;
            formula: string | null;
        }
    >;
}

export interface EditorSearchPendingEditLike {
    rowNumber: number;
    columnNumber: number;
    value: string;
}

export interface EditorReplaceChange {
    rowNumber: number;
    columnNumber: number;
    beforeValue: string;
    afterValue: string;
}

export interface EditorReplaceResult {
    status: "replaced" | "no-match" | "invalid-pattern" | "no-change";
    changes?: EditorReplaceChange[];
    replacedCellCount?: number;
    match?: EditorSearchMatch;
    nextMatch?: EditorSearchMatch | null;
}

export interface EditorSheetSearchContext {
    pendingEdits?: EditorSearchPendingEditLike[];
    scope?: EditorSearchScope;
    selectionRange?: SelectionRange;
    includeFormulaMatches?: boolean;
    editableOnly?: boolean;
}

export function createEditorSearchPattern(
    query: string,
    options: SearchOptions,
    { global = false }: { global?: boolean } = {}
): RegExp {
    const source = options.isRegexp ? query.trim() : escapeRegex(query.trim());
    const wrappedSource = options.wholeWord ? `\\b(?:${source})\\b` : source;
    const flags = `${global ? "g" : ""}${options.matchCase ? "" : "i"}`;
    return new RegExp(wrappedSource, flags);
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
    sheet: EditorSearchSheetSource,
    pendingEdits: EditorSearchPendingEditLike[] | undefined
): SearchableEditorCell[] {
    const cells = new Map<string, SearchableEditorCell>(
        Object.values(sheet.cells).map((cell) => [
            cell.key,
            {
                rowNumber: cell.rowNumber,
                columnNumber: cell.columnNumber,
                value: cell.displayValue,
                formula: cell.formula ?? null,
                match: {
                    sheetKey: sheet.key,
                    rowNumber: cell.rowNumber,
                    columnNumber: cell.columnNumber,
                },
            },
        ])
    );

    for (const edit of pendingEdits ?? []) {
        cells.set(createCellKey(edit.rowNumber, edit.columnNumber), {
            rowNumber: edit.rowNumber,
            columnNumber: edit.columnNumber,
            value: edit.value,
            formula: null,
            match: {
                sheetKey: sheet.key,
                rowNumber: edit.rowNumber,
                columnNumber: edit.columnNumber,
            },
        });
    }

    return [...cells.values()];
}

function createEditorSearchSheetSource(sheetEntry: WorkingSheetEntry): EditorSearchSheetSource {
    return {
        key: sheetEntry.key,
        rowCount: sheetEntry.sheet.rowCount,
        columnCount: sheetEntry.sheet.columnCount,
        cells: sheetEntry.sheet.cells,
    };
}

function getEditorSearchPendingEditsForSheet(
    sheetEntry: WorkingSheetEntry,
    pendingEdits: CellEdit[] | undefined
): EditorSearchPendingEditLike[] {
    return (pendingEdits ?? [])
        .filter((edit) => edit.sheetName === sheetEntry.sheet.name)
        .map((edit) => ({
            rowNumber: edit.rowNumber,
            columnNumber: edit.columnNumber,
            value: edit.value,
        }));
}

function compareEditorCellPosition(
    candidate: { rowNumber: number; columnNumber: number },
    current: { rowNumber: number; columnNumber: number }
): number {
    if (candidate.rowNumber !== current.rowNumber) {
        return candidate.rowNumber - current.rowNumber;
    }

    return candidate.columnNumber - current.columnNumber;
}

function getEditorSearchAnchor(
    sheet: EditorSearchSheetSource,
    selectedCell: EditorPanelState["selectedCell"],
    effectiveScope: EditorSearchScope,
    selectionRange: SelectionRange | undefined
): { rowNumber: number; columnNumber: number } {
    const selectionOnCurrentSheet =
        selectedCell &&
        selectedCell.rowNumber >= 1 &&
        selectedCell.rowNumber <= sheet.rowCount &&
        selectedCell.columnNumber >= 1 &&
        selectedCell.columnNumber <= sheet.columnCount;
    const selectionWithinScope =
        selectionOnCurrentSheet &&
        (effectiveScope !== "selection" ||
            isCellWithinSelectionRange(
                selectionRange ?? null,
                selectedCell!.rowNumber,
                selectedCell!.columnNumber
            ));

    if (selectionWithinScope) {
        return {
            rowNumber: selectedCell!.rowNumber,
            columnNumber: selectedCell!.columnNumber,
        };
    }

    if (effectiveScope === "selection" && selectionRange) {
        return {
            rowNumber: selectionRange.startRow,
            columnNumber: selectionRange.startColumn - 1,
        };
    }

    return {
        rowNumber: 1,
        columnNumber: 0,
    };
}

function getEditorSearchMatchContextForSheet(
    sheet: EditorSearchSheetSource,
    query: string,
    options: SearchOptions,
    searchContext: EditorSheetSearchContext = {}
): SearchMatchContext | null {
    const normalizedQuery = query.trim();
    if (!normalizedQuery) {
        return null;
    }

    const pattern = createEditorSearchPattern(normalizedQuery, options);
    const effectiveScope = getEffectiveSearchScope(
        searchContext.scope,
        Boolean(searchContext.selectionRange)
    );
    const matches = getSearchableEditorCells(sheet, searchContext.pendingEdits)
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

            if (searchContext.editableOnly && cell.formula) {
                return false;
            }

            const value = cell.value;
            const formula = cell.formula ?? "";
            return (
                pattern.test(value) ||
                (searchContext.includeFormulaMatches !== false && pattern.test(formula))
            );
        });

    if (matches.length === 0) {
        return null;
    }

    matches.sort((left, right) => {
        return compareEditorCellPosition(left, right);
    });

    return {
        effectiveScope,
        matches,
    };
}

export function findEditorSearchMatchInSheet(
    sheet: EditorSearchSheetSource,
    selectedCell: EditorPanelState["selectedCell"],
    query: string,
    direction: EditorSearchDirection,
    options: SearchOptions,
    searchContext: EditorSheetSearchContext = {}
): EditorSearchMatch | null {
    const matchContext = getEditorSearchMatchContextForSheet(sheet, query, options, searchContext);
    if (!matchContext) {
        return null;
    }

    const { effectiveScope, matches } = matchContext;
    const anchor = getEditorSearchAnchor(
        sheet,
        selectedCell,
        effectiveScope,
        searchContext.selectionRange
    );

    if (direction === "prev") {
        for (let index = matches.length - 1; index >= 0; index -= 1) {
            if (compareEditorCellPosition(matches[index]!, anchor) < 0) {
                return matches[index]!.match;
            }
        }

        return matches[matches.length - 1]!.match;
    }

    return (
        matches.find((match) => compareEditorCellPosition(match, anchor) > 0)?.match ??
        matches[0]!.match
    );
}

function resolveEditorSearchResultInSheet(
    sheet: EditorSearchSheetSource,
    selectedCell: EditorPanelState["selectedCell"],
    request: EditorSearchRequest,
    searchContext: EditorSheetSearchContext = {}
): EditorSearchResult {
    const matchContext = getEditorSearchMatchContextForSheet(sheet, request.query, request.options, {
        ...searchContext,
        scope: request.scope,
        selectionRange: request.selectionRange,
    });
    if (!matchContext) {
        return { status: "no-match" };
    }

    const match = findEditorSearchMatchInSheet(
        sheet,
        selectedCell,
        request.query,
        request.direction,
        request.options,
        {
            ...searchContext,
            scope: request.scope,
            selectionRange: request.selectionRange,
        }
    );
    if (!match) {
        return { status: "no-match" };
    }

    const matchIndex = matchContext.matches.findIndex(
        (candidate) =>
            candidate.match.sheetKey === match.sheetKey &&
            candidate.match.rowNumber === match.rowNumber &&
            candidate.match.columnNumber === match.columnNumber
    );

    return {
        status: "matched",
        match,
        matchCount: matchContext.matches.length,
        matchIndex: matchIndex >= 0 ? matchIndex + 1 : undefined,
    };
}

function applyEditorReplacement(
    value: string,
    query: string,
    replacement: string,
    options: SearchOptions,
    mode: "single" | "all"
): string {
    return value.replace(
        createEditorSearchPattern(query, options, { global: mode === "all" }),
        replacement
    );
}

function createEditorReplaceChange(
    cell: SearchableEditorCell,
    query: string,
    replacement: string,
    options: SearchOptions,
    mode: "single" | "all"
): EditorReplaceChange | null {
    const afterValue = applyEditorReplacement(cell.value, query, replacement, options, mode);
    if (afterValue === cell.value) {
        return null;
    }

    return {
        rowNumber: cell.rowNumber,
        columnNumber: cell.columnNumber,
        beforeValue: cell.value,
        afterValue,
    };
}

function applyEditorReplaceChangesToPendingEdits(
    sheet: EditorSearchSheetSource,
    pendingEdits: EditorSearchPendingEditLike[],
    changes: EditorReplaceChange[]
): EditorSearchPendingEditLike[] {
    const nextPendingEdits = new Map<string, EditorSearchPendingEditLike>(
        pendingEdits.map((edit) => [createCellKey(edit.rowNumber, edit.columnNumber), { ...edit }])
    );

    for (const change of changes) {
        const key = createCellKey(change.rowNumber, change.columnNumber);
        const modelValue = sheet.cells[key]?.displayValue ?? "";
        if (change.afterValue === modelValue) {
            nextPendingEdits.delete(key);
            continue;
        }

        nextPendingEdits.set(key, {
            rowNumber: change.rowNumber,
            columnNumber: change.columnNumber,
            value: change.afterValue,
        });
    }

    return [...nextPendingEdits.values()];
}

export function resolveEditorReplaceResultInSheet(
    sheet: EditorSearchSheetSource,
    selectedCell: EditorPanelState["selectedCell"],
    {
        query,
        replacement,
        options,
        scope,
        selectionRange,
        pendingEdits,
        mode,
    }: {
        query: string;
        replacement: string;
        options: SearchOptions;
        scope?: EditorSearchScope;
        selectionRange?: SelectionRange;
        pendingEdits?: EditorSearchPendingEditLike[];
        mode: "single" | "all";
    }
): EditorReplaceResult {
    try {
        const matchContext = getEditorSearchMatchContextForSheet(sheet, query, options, {
            pendingEdits,
            scope,
            selectionRange,
            includeFormulaMatches: false,
            editableOnly: true,
        });
        if (!matchContext) {
            return { status: "no-match" };
        }

        if (mode === "all") {
            const changes = matchContext.matches
                .map((cell) => createEditorReplaceChange(cell, query, replacement, options, "all"))
                .filter((change): change is EditorReplaceChange => Boolean(change));
            if (changes.length === 0) {
                return { status: "no-change" };
            }

            return {
                status: "replaced",
                changes,
                replacedCellCount: changes.length,
                match: matchContext.matches[0]!.match,
            };
        }

        const selectedMatch =
            selectedCell &&
            matchContext.matches.find(
                (candidate) =>
                    candidate.rowNumber === selectedCell.rowNumber &&
                    candidate.columnNumber === selectedCell.columnNumber
            );
        const nextTargetMatch =
            selectedMatch?.match ??
            findEditorSearchMatchInSheet(
                sheet,
                selectedCell,
                query,
                "next",
                options,
                {
                    pendingEdits,
                    scope,
                    selectionRange,
                    includeFormulaMatches: false,
                    editableOnly: true,
                }
            );
        const targetMatch =
            selectedMatch ??
            matchContext.matches.find(
                (candidate) =>
                    candidate.match.sheetKey === nextTargetMatch?.sheetKey &&
                    candidate.rowNumber === nextTargetMatch?.rowNumber &&
                    candidate.columnNumber === nextTargetMatch?.columnNumber
            );
        if (!targetMatch) {
            return { status: "no-match" };
        }

        const change = createEditorReplaceChange(
            targetMatch,
            query,
            replacement,
            options,
            "single"
        );
        if (!change) {
            return {
                status: "no-change",
                match: targetMatch.match,
            };
        }

        const nextMatch = findEditorSearchMatchInSheet(
            sheet,
            {
                rowNumber: targetMatch.rowNumber,
                columnNumber: targetMatch.columnNumber,
            },
            query,
            "next",
            options,
            {
                pendingEdits: applyEditorReplaceChangesToPendingEdits(
                    sheet,
                    pendingEdits ?? [],
                    [change]
                ),
                scope,
                selectionRange,
                includeFormulaMatches: false,
                editableOnly: true,
            }
        );

        return {
            status: "replaced",
            changes: [change],
            replacedCellCount: 1,
            match: targetMatch.match,
            nextMatch,
        };
    } catch {
        return { status: "invalid-pattern" };
    }
}

export function findEditorSearchMatch(
    sheetEntries: WorkingSheetEntry[],
    state: EditorPanelState,
    query: string,
    direction: EditorSearchDirection,
    options: SearchOptions,
    searchContext: {
        pendingEdits?: CellEdit[];
        scope?: EditorSearchScope;
        selectionRange?: SelectionRange;
    } = {}
): EditorSearchMatch | null {
    const activeSheetEntry = getActiveSearchSheetEntry(sheetEntries, state);
    if (!activeSheetEntry) {
        return null;
    }

    return findEditorSearchMatchInSheet(
        createEditorSearchSheetSource(activeSheetEntry),
        state.selectedCell,
        query,
        direction,
        options,
        {
            pendingEdits: getEditorSearchPendingEditsForSheet(
                activeSheetEntry,
                searchContext.pendingEdits
            ),
            scope: searchContext.scope,
            selectionRange: searchContext.selectionRange,
        }
    );
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
        const activeSheetEntry = getActiveSearchSheetEntry(sheetEntries, state);
        if (!activeSheetEntry) {
            return { status: "no-match" };
        }

        return resolveEditorSearchResultInSheet(
            createEditorSearchSheetSource(activeSheetEntry),
            state.selectedCell,
            request,
            {
                pendingEdits: getEditorSearchPendingEditsForSheet(
                    activeSheetEntry,
                    searchContext.pendingEdits
                ),
            }
        );
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
