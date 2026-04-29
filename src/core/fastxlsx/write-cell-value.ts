import { copyFile, mkdir } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import { getCellAddress, getRangeAddress } from "../model/cells";
import type {
    SheetCellAlignmentsSnapshot,
    SheetColumnAlignmentsSnapshot,
    SheetRowAlignmentsSnapshot,
} from "../model/alignment";
import type { SheetAutoFilterSnapshot } from "../model/types";
import type { AutoFilterDefinition } from "fastxlsx";
import { Workbook } from "./runtime";

export interface CellEdit {
    sheetName: string;
    rowNumber: number;
    columnNumber: number;
    value: string;
}

export interface AddSheetEdit {
    type: "addSheet";
    sheetKey: string;
    sheetName: string;
    targetIndex: number;
}

export interface DeleteSheetEdit {
    type: "deleteSheet";
    sheetKey: string;
    sheetName: string;
    targetIndex: number;
}

export interface RenameSheetEdit {
    type: "renameSheet";
    sheetKey: string;
    sheetName: string;
    nextSheetName: string;
}

export interface InsertRowEdit {
    type: "insertRow";
    sheetKey: string;
    sheetName: string;
    rowNumber: number;
    count: number;
}

export interface DeleteRowEdit {
    type: "deleteRow";
    sheetKey: string;
    sheetName: string;
    rowNumber: number;
    count: number;
}

export interface InsertColumnEdit {
    type: "insertColumn";
    sheetKey: string;
    sheetName: string;
    columnNumber: number;
    count: number;
}

export interface DeleteColumnEdit {
    type: "deleteColumn";
    sheetKey: string;
    sheetName: string;
    columnNumber: number;
    count: number;
}

export type SheetEdit =
    | AddSheetEdit
    | DeleteSheetEdit
    | RenameSheetEdit
    | InsertRowEdit
    | DeleteRowEdit
    | InsertColumnEdit
    | DeleteColumnEdit;

export interface SheetViewEdit {
    sheetKey: string;
    sheetName: string;
    freezePane: {
        columnCount: number;
        rowCount: number;
    } | null;
    autoFilter?: SheetAutoFilterSnapshot | null;
    columnWidths?: Array<number | null>;
    rowHeights?: Record<string, number | null>;
    cellAlignments?: SheetCellAlignmentsSnapshot;
    rowAlignments?: SheetRowAlignmentsSnapshot;
    columnAlignments?: SheetColumnAlignmentsSnapshot;
}

export interface WorkbookEditState {
    cellEdits: CellEdit[];
    sheetEdits: SheetEdit[];
    viewEdits?: SheetViewEdit[];
}

function createAutoFilterDefinition(
    currentDefinition: AutoFilterDefinition | null,
    autoFilter: SheetAutoFilterSnapshot
): AutoFilterDefinition {
    const range = getRangeAddress(autoFilter.range);
    const preservedColumns = currentDefinition?.range === range ? currentDefinition.columns : [];

    return {
        range,
        columns: preservedColumns,
        sortState: autoFilter.sort
            ? {
                  range,
                  conditions: [
                      {
                          columnNumber: autoFilter.sort.columnNumber,
                          descending: autoFilter.sort.direction === "desc",
                      },
                  ],
              }
            : null,
    };
}

/**
 * Writes a new display value to a specific cell in a local .xlsx file.
 * Only local `file://` URIs are supported; read-only/git URIs must be rejected before calling.
 */
export async function writeCellValue(
    fileUri: vscode.Uri,
    sheetName: string,
    rowNumber: number,
    columnNumber: number,
    value: string
): Promise<void> {
    await writeCellValues(fileUri, [{ sheetName, rowNumber, columnNumber, value }]);
}

/**
 * Writes multiple cell values to a local .xlsx file in a single open/save cycle.
 * Only local `file://` URIs are supported.
 */
export async function writeCellValues(fileUri: vscode.Uri, edits: CellEdit[]): Promise<void> {
    await writeCellValuesToDestination(fileUri, fileUri, edits);
}

export async function writeCellValuesToDestination(
    sourceUri: vscode.Uri,
    destinationUri: vscode.Uri,
    edits: CellEdit[]
): Promise<void> {
    await writeWorkbookEditsToDestination(sourceUri, destinationUri, {
        cellEdits: edits,
        sheetEdits: [],
        viewEdits: [],
    });
}

export async function writeWorkbookEditsToDestination(
    sourceUri: vscode.Uri,
    destinationUri: vscode.Uri,
    edits: WorkbookEditState
): Promise<void> {
    if (sourceUri.scheme !== "file" || destinationUri.scheme !== "file") {
        throw new Error("Cell editing is only supported for local files.");
    }

    await mkdir(path.dirname(destinationUri.fsPath), { recursive: true });

    if (sourceUri.fsPath !== destinationUri.fsPath) {
        await copyFile(sourceUri.fsPath, destinationUri.fsPath);
    }

    if (
        edits.cellEdits.length === 0 &&
        edits.sheetEdits.length === 0 &&
        (edits.viewEdits?.length ?? 0) === 0
    ) {
        return;
    }

    const workbook = await Workbook.open(destinationUri.fsPath);

    workbook.batch((currentWorkbook) => {
        for (const edit of edits.sheetEdits) {
            if (edit.type === "addSheet") {
                currentWorkbook.addSheet(edit.sheetName);
                currentWorkbook.moveSheet(edit.sheetName, edit.targetIndex);
                continue;
            }

            if (edit.type === "renameSheet") {
                currentWorkbook.renameSheet(edit.sheetName, edit.nextSheetName);
                continue;
            }

            if (edit.type === "deleteSheet") {
                currentWorkbook.deleteSheet(edit.sheetName);
            }
        }

        const sheets = new Map<string, ReturnType<typeof currentWorkbook.getSheet>>();
        const getSheet = (sheetName: string) => {
            const cachedSheet = sheets.get(sheetName);
            if (cachedSheet) {
                return cachedSheet;
            }

            const sheet = currentWorkbook.getSheet(sheetName);
            sheets.set(sheetName, sheet);
            return sheet;
        };

        for (const edit of edits.sheetEdits) {
            if (edit.type === "insertRow") {
                getSheet(edit.sheetName).insertRow(edit.rowNumber, edit.count);
                continue;
            }

            if (edit.type === "deleteRow") {
                getSheet(edit.sheetName).deleteRow(edit.rowNumber, edit.count);
                continue;
            }

            if (edit.type === "insertColumn") {
                getSheet(edit.sheetName).insertColumn(edit.columnNumber, edit.count);
                continue;
            }

            if (edit.type === "deleteColumn") {
                getSheet(edit.sheetName).deleteColumn(edit.columnNumber, edit.count);
            }
        }

        for (const edit of edits.cellEdits) {
            const sheet = getSheet(edit.sheetName);
            const address = getCellAddress(edit.rowNumber, edit.columnNumber);
            sheet.cell(address).setValue(edit.value);
        }

        for (const edit of edits.viewEdits ?? []) {
            const sheet = getSheet(edit.sheetName);
            if (edit.autoFilter === null) {
                sheet.removeAutoFilter();
            } else if (edit.autoFilter) {
                sheet.setAutoFilterDefinition(
                    createAutoFilterDefinition(sheet.getAutoFilterDefinition(), edit.autoFilter)
                );
            }

            if (
                !edit.freezePane ||
                (edit.freezePane.columnCount === 0 && edit.freezePane.rowCount === 0)
            ) {
                sheet.unfreezePane();
            } else {
                sheet.freezePane(edit.freezePane.columnCount, edit.freezePane.rowCount);
            }

            for (
                let columnIndex = 0;
                columnIndex < (edit.columnWidths?.length ?? 0);
                columnIndex += 1
            ) {
                sheet.setColumnWidth(columnIndex + 1, edit.columnWidths?.[columnIndex] ?? null);
            }

            for (const [rowNumberText, rowHeight] of Object.entries(edit.rowHeights ?? {})) {
                sheet.setRowHeight(Number(rowNumberText), rowHeight ?? null);
            }

            for (const [columnNumberText, alignment] of Object.entries(
                edit.columnAlignments ?? {}
            )) {
                sheet.setColumnStyle(Number(columnNumberText), {
                    applyAlignment: true,
                    alignment,
                });
            }

            for (const [rowNumberText, alignment] of Object.entries(edit.rowAlignments ?? {})) {
                sheet.setRowStyle(Number(rowNumberText), {
                    applyAlignment: true,
                    alignment,
                });
            }

            for (const [cellKey, alignment] of Object.entries(edit.cellAlignments ?? {})) {
                const [rowNumberText, columnNumberText] = cellKey.split(":");
                const rowNumber = Number(rowNumberText);
                const columnNumber = Number(columnNumberText);
                if (!Number.isInteger(rowNumber) || !Number.isInteger(columnNumber)) {
                    continue;
                }

                sheet.setAlignment(rowNumber, columnNumber, alignment);
            }
        }
    });

    await workbook.save(destinationUri.fsPath);
}
