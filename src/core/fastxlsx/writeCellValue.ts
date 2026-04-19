import { copyFile, mkdir } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import { getCellAddress } from "../model/cells";

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

export type SheetEdit = AddSheetEdit | DeleteSheetEdit | RenameSheetEdit;

export interface WorkbookEditState {
    cellEdits: CellEdit[];
    sheetEdits: SheetEdit[];
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

    if (edits.cellEdits.length === 0 && edits.sheetEdits.length === 0) {
        return;
    }

    const { Workbook } = await import("fastxlsx");
    const workbook = await Workbook.open(destinationUri.fsPath);

    for (const edit of edits.sheetEdits) {
        if (edit.type === "addSheet") {
            workbook.addSheet(edit.sheetName);
            workbook.moveSheet(edit.sheetName, edit.targetIndex);
            continue;
        }

        if (edit.type === "renameSheet") {
            workbook.renameSheet(edit.sheetName, edit.nextSheetName);
            continue;
        }

        workbook.deleteSheet(edit.sheetName);
    }

    for (const edit of edits.cellEdits) {
        const sheet = workbook.getSheet(edit.sheetName);
        const address = getCellAddress(edit.rowNumber, edit.columnNumber);
        sheet.cell(address).setValue(edit.value);
    }

    await workbook.save(destinationUri.fsPath);
}
