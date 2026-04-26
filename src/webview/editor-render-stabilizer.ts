import type { EditorRenderModel } from "../core/model/types";

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

    return {
        ...nextModel,
        activeSheet: {
            ...nextModel.activeSheet,
            cells: previousModel.activeSheet.cells,
            columns: reusedActiveSheetColumns
                ? previousModel.activeSheet.columns
                : nextModel.activeSheet.columns,
        },
    };
}
