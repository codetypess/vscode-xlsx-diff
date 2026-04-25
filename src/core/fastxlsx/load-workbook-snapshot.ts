import { createHash } from "node:crypto";
import { stat } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import { createCellKey, getCellAddress } from "../model/cells";
import { Workbook } from "./runtime";
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
    isEmptyWorkbookResourceUri,
    isWorkbookResourceReadOnly,
    readWorkbookResourceArchive,
} from "../../workbook/resource-uri";

interface SheetReader {
    name: string;
    rowCount: number;
    columnCount: number;
    getDisplayValue(rowNumber: number, columnNumber: number): string | null;
    getFormula(rowNumber: number, columnNumber: number): string | null;
    getStyleId(rowNumber: number, columnNumber: number): number | null;
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

function loadSheetSnapshot(workbook: WorkbookReader, sheetName: string): SheetSnapshot {
    const sheet = workbook.getSheet(sheetName);
    const cells: Record<string, CellSnapshot> = {};
    const freezePane = sheet.getFreezePane();

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
            const displayValue = sheet.getDisplayValue(rowNumber, columnNumber);
            const formula = sheet.getFormula(rowNumber, columnNumber);

            if (displayValue === null && formula === null) {
                continue;
            }

            const key = createCellKey(rowNumber, columnNumber);
            const styleId = sheet.getStyleId(rowNumber, columnNumber);
            cells[key] = {
                key,
                rowNumber,
                columnNumber,
                address: getCellAddress(rowNumber, columnNumber),
                displayValue: displayValue ?? "",
                formula,
                styleId,
            };
        }
    }

    const snapshot: SheetSnapshot = {
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        mergedRanges: [...sheet.getMergedRanges()].sort((left, right) => left.localeCompare(right)),
        freezePane: freezePane ? { ...freezePane } : null,
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

function createEmptyWorkbookSnapshot(metadata: WorkbookSnapshotMetadata): WorkbookSnapshot {
    return {
        ...metadata,
        sheets: [],
    };
}

export async function loadWorkbookSnapshot(
    filePathOrUri: string | vscode.Uri
): Promise<WorkbookSnapshot> {
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

    if (filePathOrUri.scheme === "file") {
        const [workbook, fileStats] = await Promise.all([
            Workbook.open(filePathOrUri.fsPath),
            stat(filePathOrUri.fsPath),
        ]);

        return createWorkbookSnapshot(workbook, {
            filePath: getWorkbookResourcePathLabel(filePathOrUri),
            fileName: getWorkbookResourceName(filePathOrUri),
            fileSize: fileStats.size,
            modifiedTime: fileStats.mtime.toISOString(),
            isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
        });
    }

    const resourceName = getWorkbookResourceName(filePathOrUri);
    const resourcePath = getWorkbookResourcePathLabel(filePathOrUri);
    const resourceDetail = await getWorkbookResourceDetail(filePathOrUri);
    const resourceFilePath =
        filePathOrUri.scheme === "git" ? decodeURIComponent(filePathOrUri.path) : undefined;

    const resourceStats = resourceFilePath
        ? await stat(resourceFilePath).catch(() => undefined)
        : undefined;
    const modifiedTimeLabel =
        resourceStats
            ? undefined
            : (getWorkbookResourceTimeLabel(filePathOrUri) ??
              `${filePathOrUri.scheme.toUpperCase()} resource`);

    if (isEmptyWorkbookResourceUri(filePathOrUri)) {
        return createEmptyWorkbookSnapshot({
            filePath: resourcePath,
            fileName: resourceName,
            fileSize: 0,
            modifiedTime: new Date(0).toISOString(),
            modifiedTimeLabel,
            detailLabel: resourceDetail?.label,
            detailValue: resourceDetail?.value,
            titleDetail: resourceDetail?.titleValue,
            isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
        });
    }

    const archiveData =
        (await readWorkbookResourceArchive(filePathOrUri)) ??
        (await vscode.workspace.fs.readFile(filePathOrUri));
    const workbook = Workbook.fromUint8Array(archiveData);

    return createWorkbookSnapshot(workbook, {
        filePath: resourcePath,
        fileName: resourceName,
        fileSize: archiveData.byteLength,
        modifiedTime: resourceStats?.mtime.toISOString() ?? new Date().toISOString(),
        modifiedTimeLabel,
        detailLabel: resourceDetail?.label,
        detailValue: resourceDetail?.value,
        titleDetail: resourceDetail?.titleValue,
        isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
    });
}
