import { createHash } from "node:crypto";
import { stat } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import { createCellKey, getCellAddress } from "../model/cells";
import {
    type CellSnapshot,
    type SheetFreezePaneSnapshot,
    type SheetSnapshot,
    type WorkbookSnapshot,
} from "../model/types";
import {
    getWorkbookResourceDetail,
    getWorkbookResourceName,
    getWorkbookResourcePathLabel,
    getWorkbookResourceTimeLabel,
    isWorkbookResourceReadOnly,
} from "../../workbook/resourceUri";

interface SheetReader {
    name: string;
    rowCount: number;
    columnCount: number;
    getDisplayValue(rowNumber: number, columnNumber: number): string | null;
    getFormula(rowNumber: number, columnNumber: number): string | null;
    getStyleId(rowNumber: number, columnNumber: number): number | null;
    getRowHeight(rowNumber: number): number | null;
    getColumnWidth(columnNumber: number): number | null;
    getMergedRanges(): string[];
    getFreezePane(): SheetFreezePaneSnapshot | null;
}

interface WorkbookReader {
    getSheet(sheetName: string): SheetReader;
    getSheetNames(): string[];
}

interface WorkbookSnapshotMetadata {
    filePath: string;
    fileName: string;
    fileSize: number;
    modifiedTime: string;
    modifiedTimeLabel?: string;
    detailLabel?: string;
    detailValue?: string;
    titleDetail?: string;
    isReadonly?: boolean;
}

function createSheetSignature(sheet: SheetSnapshot): string {
    const hash = createHash("sha1");
    hash.update(`${sheet.name}\n`);
    hash.update(`${sheet.rowCount}:${sheet.columnCount}\n`);

    for (const [rowNumber, rowHeight] of Object.entries(sheet.rowHeights).sort(
        ([left], [right]) => Number(left) - Number(right)
    )) {
        hash.update(`row:${rowNumber}:${rowHeight}\n`);
    }

    for (const [columnNumber, columnWidth] of Object.entries(sheet.columnWidths).sort(
        ([left], [right]) => Number(left) - Number(right)
    )) {
        hash.update(`col:${columnNumber}:${columnWidth}\n`);
    }

    for (const mergedRange of sheet.mergedRanges) {
        hash.update(`merge:${mergedRange}\n`);
    }

    for (const cell of Object.values(sheet.cells).sort((left, right) =>
        left.key.localeCompare(right.key)
    )) {
        hash.update(`${cell.address}\u0000${cell.displayValue}\u0000${cell.formula ?? ""}\n`);
    }

    return hash.digest("hex");
}

function readExplicitRowHeights(sheet: SheetReader): Record<number, number> {
    const rowHeights: Record<number, number> = {};

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const rowHeight = sheet.getRowHeight(rowNumber);
        if (rowHeight === null) {
            continue;
        }

        rowHeights[rowNumber] = rowHeight;
    }

    return rowHeights;
}

function readExplicitColumnWidths(sheet: SheetReader): Record<number, number> {
    const columnWidths: Record<number, number> = {};

    for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
        const columnWidth = sheet.getColumnWidth(columnNumber);
        if (columnWidth === null) {
            continue;
        }

        columnWidths[columnNumber] = columnWidth;
    }

    return columnWidths;
}

function loadSheetSnapshot(workbook: WorkbookReader, sheetName: string): SheetSnapshot {
    const sheet = workbook.getSheet(sheetName);
    const cells: Record<string, CellSnapshot> = {};
    const freezePane = sheet.getFreezePane();
    const rowHeights = readExplicitRowHeights(sheet);
    const columnWidths = readExplicitColumnWidths(sheet);

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
            const displayValue = sheet.getDisplayValue(rowNumber, columnNumber);
            const formula = sheet.getFormula(rowNumber, columnNumber);

            if (displayValue === null && formula === null) {
                continue;
            }

            const key = createCellKey(rowNumber, columnNumber);
            cells[key] = {
                key,
                rowNumber,
                columnNumber,
                address: getCellAddress(rowNumber, columnNumber),
                displayValue: displayValue ?? "",
                formula,
                styleId: sheet.getStyleId(rowNumber, columnNumber),
            };
        }
    }

    const snapshot: SheetSnapshot = {
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        mergedRanges: [...sheet.getMergedRanges()].sort((left, right) => left.localeCompare(right)),
        freezePane: freezePane ? { ...freezePane } : null,
        rowHeights,
        columnWidths,
        cells,
        signature: "",
    };

    snapshot.signature = createSheetSignature(snapshot);
    return snapshot;
}

function createWorkbookSnapshot(
    workbook: WorkbookReader,
    metadata: WorkbookSnapshotMetadata
): WorkbookSnapshot {
    const sheets = workbook
        .getSheetNames()
        .map((sheetName) => loadSheetSnapshot(workbook, sheetName));

    return {
        ...metadata,
        sheets,
    };
}

export async function loadWorkbookSnapshot(
    filePathOrUri: string | vscode.Uri
): Promise<WorkbookSnapshot> {
    const { Workbook } = await import("fastxlsx");

    if (typeof filePathOrUri === "string") {
        const [workbook, fileStats] = await Promise.all([
            Workbook.open(filePathOrUri),
            stat(filePathOrUri),
        ]);

        return createWorkbookSnapshot(workbook, {
            filePath: filePathOrUri,
            fileName: path.basename(filePathOrUri),
            fileSize: fileStats.size,
            modifiedTime: fileStats.mtime.toISOString(),
            isReadonly: false,
        });
    }

    const archiveData = await vscode.workspace.fs.readFile(filePathOrUri);
    const workbook = Workbook.fromUint8Array(archiveData);
    const resourceName = getWorkbookResourceName(filePathOrUri);
    const resourcePath = getWorkbookResourcePathLabel(filePathOrUri);
    const resourceDetail = await getWorkbookResourceDetail(filePathOrUri);
    const resourceFilePath =
        filePathOrUri.scheme === "git" ? decodeURIComponent(filePathOrUri.path) : undefined;

    if (filePathOrUri.scheme === "file") {
        const fileStats = await stat(filePathOrUri.fsPath);
        return createWorkbookSnapshot(workbook, {
            filePath: resourcePath,
            fileName: resourceName,
            fileSize: fileStats.size,
            modifiedTime: fileStats.mtime.toISOString(),
            isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
        });
    }

    const resourceStats = resourceFilePath
        ? await stat(resourceFilePath).catch(() => undefined)
        : undefined;

    return createWorkbookSnapshot(workbook, {
        filePath: resourcePath,
        fileName: resourceName,
        fileSize: archiveData.byteLength,
        modifiedTime: resourceStats?.mtime.toISOString() ?? new Date().toISOString(),
        modifiedTimeLabel: resourceStats
            ? undefined
            : (getWorkbookResourceTimeLabel(filePathOrUri) ??
              `${filePathOrUri.scheme.toUpperCase()} resource`),
        detailLabel: resourceDetail?.label,
        detailValue: resourceDetail?.value,
        titleDetail: resourceDetail?.titleValue,
        isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
    });
}
