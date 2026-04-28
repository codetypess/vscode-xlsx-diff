import type { EditorRenderModel } from "../../core/model/types";
import { areCellAlignmentMapsEquivalent } from "../../core/model/alignment";

export function stabilizeIncomingRenderModel(
    previousModel: EditorRenderModel | null,
    nextModel: EditorRenderModel,
    {
        canReuseActiveSheetData,
    }: {
        canReuseActiveSheetData: boolean;
    }
): EditorRenderModel {
    if (
        !canReuseActiveSheetData ||
        !previousModel ||
        previousModel.activeSheet.key !== nextModel.activeSheet.key ||
        previousModel.activeSheet.rowCount !== nextModel.activeSheet.rowCount ||
        previousModel.activeSheet.columnCount !== nextModel.activeSheet.columnCount
    ) {
        return nextModel;
    }

    const reusedActiveSheetColumns =
        previousModel.activeSheet.columns.length === nextModel.activeSheet.columns.length &&
        previousModel.activeSheet.columns.every(
            (label, index) => label === nextModel.activeSheet.columns[index]
        );
    const previousColumnWidths = previousModel.activeSheet.columnWidths ?? [];
    const nextColumnWidths = nextModel.activeSheet.columnWidths ?? [];
    const reusedActiveSheetColumnWidths =
        previousColumnWidths.length === nextColumnWidths.length &&
        previousColumnWidths.every((width, index) => width === nextColumnWidths[index]);
    const previousRowHeights = previousModel.activeSheet.rowHeights ?? {};
    const nextRowHeights = nextModel.activeSheet.rowHeights ?? {};
    const previousRowHeightKeys = Object.keys(previousRowHeights);
    const nextRowHeightKeys = Object.keys(nextRowHeights);
    const reusedActiveSheetRowHeights =
        previousRowHeightKeys.length === nextRowHeightKeys.length &&
        previousRowHeightKeys.every(
            (rowNumber) => previousRowHeights[rowNumber] === nextRowHeights[rowNumber]
        );
    const reusedActiveSheetCellAlignments = areCellAlignmentMapsEquivalent(
        previousModel.activeSheet.cellAlignments,
        nextModel.activeSheet.cellAlignments
    );
    const reusedActiveSheetRowAlignments = areCellAlignmentMapsEquivalent(
        previousModel.activeSheet.rowAlignments,
        nextModel.activeSheet.rowAlignments
    );
    const reusedActiveSheetColumnAlignments = areCellAlignmentMapsEquivalent(
        previousModel.activeSheet.columnAlignments,
        nextModel.activeSheet.columnAlignments
    );

    return {
        ...nextModel,
        activeSheet: {
            ...nextModel.activeSheet,
            cells: previousModel.activeSheet.cells,
            columns: reusedActiveSheetColumns
                ? previousModel.activeSheet.columns
                : nextModel.activeSheet.columns,
            columnWidths: reusedActiveSheetColumnWidths
                ? previousModel.activeSheet.columnWidths
                : nextModel.activeSheet.columnWidths,
            rowHeights: reusedActiveSheetRowHeights
                ? previousModel.activeSheet.rowHeights
                : nextModel.activeSheet.rowHeights,
            cellAlignments: reusedActiveSheetCellAlignments
                ? previousModel.activeSheet.cellAlignments
                : nextModel.activeSheet.cellAlignments,
            rowAlignments: reusedActiveSheetRowAlignments
                ? previousModel.activeSheet.rowAlignments
                : nextModel.activeSheet.rowAlignments,
            columnAlignments: reusedActiveSheetColumnAlignments
                ? previousModel.activeSheet.columnAlignments
                : nextModel.activeSheet.columnAlignments,
        },
    };
}
