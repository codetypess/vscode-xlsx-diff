import type { EditorRenderModel, EditorRenderPayload } from "../../core/model/types";
import {
    areCellAlignmentMapsEquivalent,
    type CellAlignmentSnapshot,
} from "../../core/model/alignment";

const EMPTY_COLUMNS: string[] = [];
const EMPTY_COLUMN_WIDTHS: Array<number | null> = [];
const EMPTY_ROW_HEIGHTS = {};
const EMPTY_ALIGNMENTS: Record<string, CellAlignmentSnapshot> = {};
const EMPTY_CELLS = {};

function canReusePreviousActiveSheetData(
    previousModel: EditorRenderModel | null,
    nextModel: EditorRenderPayload
): previousModel is EditorRenderModel {
    return Boolean(
        previousModel &&
            previousModel.activeSheet.key === nextModel.activeSheet.key &&
            previousModel.activeSheet.rowCount === nextModel.activeSheet.rowCount &&
            previousModel.activeSheet.columnCount === nextModel.activeSheet.columnCount
    );
}

function areColumnsEquivalent(left: readonly string[], right: readonly string[]): boolean {
    return left.length === right.length && left.every((label, index) => label === right[index]);
}

function applyAlignmentMapPatch(
    previousAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    nextAlignments: Readonly<Record<string, CellAlignmentSnapshot>> | undefined,
    dirtyKeys: readonly string[] | undefined
): Readonly<Record<string, CellAlignmentSnapshot>> {
    if (!dirtyKeys || dirtyKeys.length === 0) {
        return nextAlignments ?? previousAlignments ?? EMPTY_ALIGNMENTS;
    }

    const patchedAlignments = Object.create(previousAlignments ?? null) as Record<
        string,
        CellAlignmentSnapshot | undefined
    >;
    for (const key of dirtyKeys) {
        if (nextAlignments && Object.hasOwn(nextAlignments, key)) {
            const nextAlignment = nextAlignments[key];
            if (nextAlignment) {
                patchedAlignments[key] = nextAlignment;
                continue;
            }
        }

        patchedAlignments[key] = undefined;
    }

    return patchedAlignments as Record<string, CellAlignmentSnapshot>;
}

function normalizeIncomingRenderModel(nextModel: EditorRenderPayload): EditorRenderModel {
    const {
        cellAlignmentDirtyKeys: _cellAlignmentDirtyKeys,
        rowAlignmentDirtyKeys: _rowAlignmentDirtyKeys,
        columnAlignmentDirtyKeys: _columnAlignmentDirtyKeys,
        ...nextActiveSheet
    } = nextModel.activeSheet;

    return {
        ...nextModel,
        activeSheet: {
            ...nextActiveSheet,
            columns: nextActiveSheet.columns ?? EMPTY_COLUMNS,
            columnWidths: nextActiveSheet.columnWidths ?? EMPTY_COLUMN_WIDTHS,
            rowHeights: nextActiveSheet.rowHeights ?? EMPTY_ROW_HEIGHTS,
            cellAlignments: nextActiveSheet.cellAlignments ?? EMPTY_ALIGNMENTS,
            rowAlignments: nextActiveSheet.rowAlignments ?? EMPTY_ALIGNMENTS,
            columnAlignments: nextActiveSheet.columnAlignments ?? EMPTY_ALIGNMENTS,
            cells: nextActiveSheet.cells ?? EMPTY_CELLS,
            freezePane: nextActiveSheet.freezePane ?? null,
            autoFilter: nextActiveSheet.autoFilter ?? null,
        },
    };
}

export function stabilizeIncomingRenderModel(
    previousModel: EditorRenderModel | null,
    nextModel: EditorRenderPayload,
    {
        canReuseActiveSheetData,
    }: {
        canReuseActiveSheetData: boolean;
    }
): EditorRenderModel {
    if (!canReuseActiveSheetData || !canReusePreviousActiveSheetData(previousModel, nextModel)) {
        return normalizeIncomingRenderModel(nextModel);
    }

    const {
        cellAlignmentDirtyKeys,
        rowAlignmentDirtyKeys,
        columnAlignmentDirtyKeys,
        ...nextActiveSheet
    } = nextModel.activeSheet;
    const previousColumnWidths = previousModel.activeSheet.columnWidths ?? [];
    const nextColumnWidths = nextActiveSheet.columnWidths ?? previousColumnWidths;
    const reusedActiveSheetColumnWidths =
        previousColumnWidths.length === nextColumnWidths.length &&
        previousColumnWidths.every((width, index) => width === nextColumnWidths[index]);
    const previousRowHeights = previousModel.activeSheet.rowHeights ?? {};
    const nextRowHeights = nextActiveSheet.rowHeights ?? previousRowHeights;
    const previousRowHeightKeys = Object.keys(previousRowHeights);
    const nextRowHeightKeys = Object.keys(nextRowHeights);
    const reusedActiveSheetRowHeights =
        previousRowHeightKeys.length === nextRowHeightKeys.length &&
        previousRowHeightKeys.every(
            (rowNumber) => previousRowHeights[rowNumber] === nextRowHeights[rowNumber]
        );
    const nextCellAlignments = applyAlignmentMapPatch(
        previousModel.activeSheet.cellAlignments,
        nextActiveSheet.cellAlignments,
        cellAlignmentDirtyKeys
    );
    const nextRowAlignments = applyAlignmentMapPatch(
        previousModel.activeSheet.rowAlignments,
        nextActiveSheet.rowAlignments,
        rowAlignmentDirtyKeys
    );
    const nextColumnAlignments = applyAlignmentMapPatch(
        previousModel.activeSheet.columnAlignments,
        nextActiveSheet.columnAlignments,
        columnAlignmentDirtyKeys
    );
    const reusedActiveSheetCellAlignments =
        !cellAlignmentDirtyKeys?.length &&
        areCellAlignmentMapsEquivalent(previousModel.activeSheet.cellAlignments, nextCellAlignments);
    const reusedActiveSheetRowAlignments =
        !rowAlignmentDirtyKeys?.length &&
        areCellAlignmentMapsEquivalent(previousModel.activeSheet.rowAlignments, nextRowAlignments);
    const reusedActiveSheetColumnAlignments =
        !columnAlignmentDirtyKeys?.length &&
        areCellAlignmentMapsEquivalent(
            previousModel.activeSheet.columnAlignments,
            nextColumnAlignments
        );
    const nextColumns = nextActiveSheet.columns ?? previousModel.activeSheet.columns;
    const nextCells = nextActiveSheet.cells ?? previousModel.activeSheet.cells;

    return {
        ...nextModel,
        activeSheet: {
            ...nextActiveSheet,
            cells:
                nextCells === previousModel.activeSheet.cells
                    ? previousModel.activeSheet.cells
                    : nextCells,
            columns: areColumnsEquivalent(previousModel.activeSheet.columns, nextColumns)
                ? previousModel.activeSheet.columns
                : nextColumns,
            columnWidths: reusedActiveSheetColumnWidths
                ? previousModel.activeSheet.columnWidths
                : nextColumnWidths,
            rowHeights: reusedActiveSheetRowHeights
                ? previousModel.activeSheet.rowHeights
                : nextRowHeights,
            cellAlignments: reusedActiveSheetCellAlignments
                ? previousModel.activeSheet.cellAlignments
                : nextCellAlignments,
            rowAlignments: reusedActiveSheetRowAlignments
                ? previousModel.activeSheet.rowAlignments
                : nextRowAlignments,
            columnAlignments: reusedActiveSheetColumnAlignments
                ? previousModel.activeSheet.columnAlignments
                : nextColumnAlignments,
            freezePane: nextActiveSheet.freezePane ?? null,
            autoFilter: nextActiveSheet.autoFilter ?? null,
        },
    };
}
